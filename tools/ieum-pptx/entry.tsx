/**
 * Browser entry that exposes IeumOffice's PowerPoint engine to the Slide Remote page.
 * Parsing (pptx-engine), layout (pptx-render) and drawing (the slides app's Konva
 * adapter) are IeumOffice / GenOffice code (Apache-2.0); this file only wires them
 * together for a read-only viewer.
 */
import React from 'react'
import { createRoot } from 'react-dom/client'
import type Konva from 'konva'
import { openPptx, getSlideNotes, type SlideElement } from '@genoffice/pptx-engine'
import { buildRenderSlide, type RenderSlide, type FontMetricsProvider, type RunStyle } from '@genoffice/pptx-render'
import { displayFontFamily } from '@ieum-slides/renderer/konva-adapter'
import { SlideThumb } from '@ieum-slides/renderer/SlideThumb'
import { createImageLoader } from '@ieum-slides/renderer/image-loader'
import { displayMime } from '@ieum-slides/main/media-mime'

const MEDIA_MIME: Record<string, string> = {
  mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime', webm: 'video/webm', ogv: 'video/ogg',
  wmv: 'video/x-ms-wmv', avi: 'video/x-msvideo', mpg: 'video/mpeg', mpeg: 'video/mpeg',
  mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', ogg: 'audio/ogg', wma: 'audio/x-ms-wma', aac: 'audio/aac',
}
const BASE_W = 1280

/**
 * Text metrics from the browser's own canvas, using the same font stack the Konva
 * adapter draws with, so line breaks and centering match what is painted.
 */
class CanvasMetrics implements FontMetricsProvider {
  private ctx = document.createElement('canvas').getContext('2d')!
  private widths = new Map<string, number>()
  private font(s: RunStyle): string {
    return `${s.italic ? 'italic ' : ''}${s.bold ? 'bold ' : ''}${s.fontSizePx}px ${displayFontFamily(s.fontFamily)}`
  }
  metrics(s: RunStyle) {
    this.ctx.font = this.font(s)
    const m = this.ctx.measureText('Hg가')
    const a = m.fontBoundingBoxAscent, d = m.fontBoundingBoxDescent
    if (!(a > 0) || !(d >= 0)) return { ascent: s.fontSizePx * 0.8, descent: s.fontSizePx * 0.2, lineHeight: s.fontSizePx * 1.2 }
    // keep the engine's 1.2em line model unless the face is taller
    return { ascent: a, descent: d, lineHeight: Math.max(a + d, s.fontSizePx * 1.2) }
  }
  measure(text: string, s: RunStyle): number {
    const font = this.font(s), key = font + '|' + text
    let w = this.widths.get(key)
    if (w === undefined) {
      this.ctx.font = font
      w = this.ctx.measureText(text).width
      if (this.widths.size > 20000) this.widths.clear()
      this.widths.set(key, w)
    }
    return w
  }
}

function toDataUrl(bytes: Uint8Array, mime: string): string {
  let s = ''
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  return `data:${mime};base64,${btoa(s)}`
}

export async function open(data: ArrayBuffer | Uint8Array) {
  const opened = await openPptx(data instanceof Uint8Array ? data : new Uint8Array(data))
  const { deck, archive } = opened
  const urlByRef = new Map<string, string>()
  const blobUrls: string[] = []
  const images = new Map<string, HTMLImageElement>()
  const loader = createImageLoader((entries) => { for (const [k, v] of entries) images.set(k, v) }, 1, 0)
  const blobUrl = (bytes: Uint8Array, mime: string) => {
    const u = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime }))
    blobUrls.push(u)
    return u
  }
  // every image the render tree asks for goes through here, so this is also the load list
  const media = (ref: string): string | undefined => {
    let u = urlByRef.get(ref)
    if (u) return u
    const bytes = archive.readBytes(ref)
    if (!bytes) return undefined
    const mime = displayMime(ref, bytes)
    // metafiles are rasterized by the image loader, which wants them as data URLs
    u = /emf|wmf/.test(mime) ? toDataUrl(bytes, mime) : blobUrl(bytes, mime)
    urlByRef.set(ref, u)
    return u
  }
  const metrics = new CanvasMetrics()
  const built = new Map<number, RenderSlide>()
  const build = (i: number) => {
    let r = built.get(i)
    if (!r) {
      r = buildRenderSlide(deck.slides[i]!, deck.size, { fitWidthPx: BASE_W, media, metrics, slideNo: i + 1 })
      built.set(i, r)
      loader.load(urlByRef.values())
    }
    return r
  }
  const settle = async () => {
    const until = Date.now() + 15000
    while (loader.pending() > 0 && Date.now() < until) await new Promise((r) => setTimeout(r, 30))
    await new Promise((r) => setTimeout(r, 0))
    try { await document.fonts.ready } catch {}
  }

  const host = document.createElement('div')
  host.style.cssText = 'position:fixed;left:-100000px;top:0;pointer-events:none;'
  document.body.appendChild(host)
  const root = createRoot(host)
  let queue: Promise<unknown> = Promise.resolve()

  /** Draw slide n (1-based) into a new canvas that fits inside maxW x maxH. */
  function render(n: number, maxW: number, maxH: number): Promise<HTMLCanvasElement> {
    const job = queue.then(async () => {
      const slide = build(n - 1)
      await settle()
      const k = Math.min(maxW / slide.widthPx, maxH / slide.heightPx)
      const stage = await new Promise<Konva.Stage>((resolve) => {
        root.render(<SlideThumb key={`${n}:${k}`} slide={slide} images={images} width={slide.widthPx} stageRef={(s) => s && resolve(s)} />)
      })
      // two frames: react-konva commits, then Konva's batched draw lands
      await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))
      const src = stage.toCanvas({ pixelRatio: k })
      const out = document.createElement('canvas')
      out.width = src.width
      out.height = src.height
      const ctx = out.getContext('2d')!
      ctx.fillStyle = '#fff'
      ctx.fillRect(0, 0, out.width, out.height)
      ctx.drawImage(src, 0, 0)
      return out
    })
    queue = job.catch(() => {})
    return job
  }

  /** Videos / audio clips on slide n, positioned as fractions of the slide. */
  function mediaOf(n: number) {
    const slide = build(n - 1)
    const targets = new Map<string, { kind: 'video' | 'audio'; target: string }>()
    const walk = (els: SlideElement[]) => {
      for (const el of els) {
        const m = (el as { media?: { kind: 'video' | 'audio'; target?: string; external?: boolean } }).media
        if (m?.target && !m.external) targets.set(el.id, { kind: m.kind, target: m.target })
        const kids = (el as { children?: SlideElement[] }).children
        if (kids) walk(kids)
      }
    }
    walk(deck.slides[n - 1]!.elements)
    const out: Array<{ id: string; kind: 'video' | 'audio'; url: string; mime: string; x: number; y: number; w: number; h: number; rotation: number }> = []
    const visit = (nodes: RenderSlide['nodes'], ox: number, oy: number) => {
      for (const node of nodes) {
        const t = targets.get(node.sourceId)
        if (t) {
          const bytes = archive.readBytes(t.target)
          if (bytes) {
            const mime = MEDIA_MIME[t.target.split('.').pop()!.toLowerCase()] ?? (t.kind === 'video' ? 'video/mp4' : 'audio/mpeg')
            let url = urlByRef.get('media:' + t.target)
            if (!url) { url = blobUrl(bytes, mime); urlByRef.set('media:' + t.target, url) }
            out.push({
              id: node.sourceId, kind: t.kind, url, mime,
              x: (ox + node.box.x) / slide.widthPx, y: (oy + node.box.y) / slide.heightPx,
              w: node.box.w / slide.widthPx, h: node.box.h / slide.heightPx,
              rotation: (node.box as { rotation?: number }).rotation ?? 0,
            })
          }
        }
        const kids = (node as { children?: RenderSlide['nodes'] }).children
        if (node.type === 'group' && kids) visit(kids, ox + node.box.x, oy + node.box.y)
      }
    }
    visit(slide.nodes, 0, 0)
    return out
  }

  return {
    count: deck.slides.length,
    render,
    media: mediaOf,
    notes: (n: number) => { try { return getSlideNotes(archive, deck.slides[n - 1]!.path) } catch { return '' } },
    close() { root.unmount(); host.remove(); loader.dispose(); for (const u of blobUrls) URL.revokeObjectURL(u) },
  }
}
