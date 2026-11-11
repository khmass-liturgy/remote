/**
 * Browser entry that exposes IeumOffice's PowerPoint engine to the Slide Remote page.
 * Parsing (pptx-engine), layout (pptx-render) and drawing (the slides app's Konva
 * adapter) are IeumOffice / GenOffice code (Apache-2.0); this file only wires them
 * together for a read-only viewer.
 */
import React, { useEffect, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import type Konva from 'konva'
import { openPptx, getSlideNotes, getSlideAnimations, getSlideTransition, elementSpid, type SlideElement } from '@genoffice/pptx-engine'
import { AnimatedSlideStage, useAnimPlayer } from '@ieum-slides/renderer/components/AnimatedSlide'
import type { AnimationItem } from '@ieum-slides/shared/ipc'
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

/** Page-transition approximations, the same keyframes IeumOffice's show view uses (styles.css ss-anim-*). */
const TRANSITION_CSS = `
.ieum-tr{line-height:0}
.ieum-tr-fade,.ieum-tr-morph{animation:ieum-fade .5s ease both}
.ieum-tr-push{animation:ieum-push .5s ease-out both}
.ieum-tr-wipe{animation:ieum-wipe .6s ease-out both}
.ieum-tr-split{animation:ieum-split .55s ease-out both}
.ieum-tr-circle{animation:ieum-circle .6s ease-out both}
.ieum-tr-cover{animation:ieum-cover .5s ease both}
.ieum-tr-pull{animation:ieum-pull .5s ease both}
.ieum-tr-dissolve{animation:ieum-fade .6s steps(12) both}
.ieum-tr-zoom{animation:ieum-zoomin .5s ease both}
@keyframes ieum-fade{from{opacity:0}to{opacity:1}}
@keyframes ieum-push{from{transform:translateY(100%)}to{transform:translateY(0)}}
@keyframes ieum-wipe{from{clip-path:inset(0 100% 0 0)}to{clip-path:inset(0 0 0 0)}}
@keyframes ieum-split{from{clip-path:inset(50% 0 50% 0)}to{clip-path:inset(0 0 0 0)}}
@keyframes ieum-circle{from{clip-path:circle(0% at 50% 50%)}to{clip-path:circle(75% at 50% 50%)}}
@keyframes ieum-cover{from{transform:translateX(100%)}to{transform:translateX(0)}}
@keyframes ieum-pull{from{clip-path:inset(0 0 0 100%)}to{clip-path:inset(0 0 0 0)}}
@keyframes ieum-zoomin{from{transform:scale(.2);opacity:0}to{transform:scale(1);opacity:1}}
`
const ANIMATED_TRANSITIONS = ['fade', 'push', 'wipe', 'split', 'circle', 'cover', 'pull', 'dissolve', 'zoom']

interface ShowHandle {
  go(n: number, mode: 'fresh' | 'all', width: number): void
  advance(): boolean
  resize(width: number): void
}
interface ShowDeck {
  slide(n: number): RenderSlide
  anims(n: number): AnimationItem[]
  transition(n: number): string
  images: Map<string, HTMLImageElement>
  onMedia?: (sourceId: string, effect: string) => void
}

/** Live slide: IeumOffice's animation player (useAnimPlayer) driving its animated Konva stage. */
function ShowView({ deck, bind }: { deck: ShowDeck; bind: (h: ShowHandle) => void }) {
  const [st, setSt] = useState({ n: 0, width: 0, nonce: 0, trans: 'none' })
  const slide = st.n ? deck.slide(st.n) : null
  const player = useAnimPlayer(slide?.heightPx ?? 720, slide?.widthPx ?? 1280)
  const ref = useRef(player)
  ref.current = player
  useEffect(() => {
    bind({
      go(n, mode, width) {
        let trans = mode === 'fresh' ? deck.transition(n) : 'none'
        if (trans === 'random') trans = ANIMATED_TRANSITIONS[Math.floor(Math.random() * ANIMATED_TRANSITIONS.length)]!
        setSt((s) => ({ n, width, nonce: s.nonce + 1, trans }))
        ref.current.load(deck.anims(n), mode)
      },
      advance: () => ref.current.advance(),
      resize: (width) => setSt((s) => ({ ...s, width })),
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])
  // media animations (play / pause / stop in the click sequence) are handed to the page's players
  const applied = useRef(0)
  const seenEpoch = useRef(-1)
  useEffect(() => {
    if (seenEpoch.current !== player.epoch) {
      seenEpoch.current = player.epoch
      applied.current = player.mediaBase
    }
    while (applied.current < player.mediaCmds.length) {
      const c = player.mediaCmds[applied.current++]!
      deck.onMedia?.(c.sourceId, c.effect)
    }
  }, [player.mediaCmds, player.epoch, player.mediaBase, deck])
  if (!slide || !st.width) return null
  return (
    <div key={st.nonce} className={`ieum-tr${st.trans !== 'none' ? ` ieum-tr-${st.trans}` : ''}`}>
      <AnimatedSlideStage slide={slide} images={deck.images} width={st.width} states={player.states} />
    </div>
  )
}

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

  /** This slide's animations, keyed to render nodes the way IeumOffice's main process does. */
  function animsOf(n: number): AnimationItem[] {
    const slide = deck.slides[n - 1]!
    const bySpid = new Map<number, SlideElement>()
    const walk = (els: SlideElement[]) => {
      for (const el of els) {
        const id = elementSpid(el)
        if (id != null) bySpid.set(id, el)
        const kids = (el as { children?: SlideElement[] }).children
        if (kids) walk(kids)
      }
    }
    walk(slide.elements)
    const out: AnimationItem[] = []
    try {
      for (const a of getSlideAnimations(slide)) {
        const el = bySpid.get(a.spid)
        if (!el) continue
        out.push({
          sourceId: el.id, targetName: el.name ?? '', effect: a.effect, trigger: a.trigger,
          durationMs: a.durationMs, delayMs: a.delayMs,
          ...(a.direction != null ? { direction: a.direction } : {}),
          ...(a.motionPath != null ? { motionPath: a.motionPath } : {}),
          ...(a.paragraph != null ? { paragraph: a.paragraph } : {}),
        })
      }
    } catch (e) { console.warn('animations unreadable on slide', n, e) }
    return out
  }

  /**
   * Live show inside `container`: go(n, 'fresh') enters a slide with its transition and
   * animations armed, go(n, 'all') shows it fully built; advance() plays the next
   * animation step and returns false once the slide has nothing left to play.
   */
  function createShow(container: HTMLElement, opts: { onMedia?: (sourceId: string, effect: string) => void } = {}) {
    if (!document.getElementById('ieum-tr-css')) {
      const st = document.createElement('style')
      st.id = 'ieum-tr-css'
      st.textContent = TRANSITION_CSS
      document.head.appendChild(st)
    }
    let handle: ShowHandle | null = null
    let token = 0
    const showRoot = createRoot(container)
    const showDeck: ShowDeck = {
      slide: (n) => build(n - 1),
      anims: animsOf,
      transition: (n) => { try { return getSlideTransition(deck.slides[n - 1]!) } catch { return 'none' } },
      images,
      onMedia: opts.onMedia,
    }
    const ready = new Promise<void>((resolve) => {
      showRoot.render(<ShowView deck={showDeck} bind={(h) => { handle = h; resolve() }} />)
    })
    return {
      async go(n: number, mode: 'fresh' | 'all', width: number) {
        const mine = ++token
        build(n - 1)
        await ready
        await settle()
        if (mine !== token) return
        handle!.go(n, mode, width)
        if (n < deck.slides.length) build(n) // start decoding the next slide's pictures
      },
      advance: () => handle?.advance() ?? false,
      resize: (width: number) => handle?.resize(width),
      steps: (n: number) => animsOf(n).filter((a) => a.trigger === 'onClick').length,
      destroy() { token++; showRoot.unmount() },
    }
  }

  return {
    count: deck.slides.length,
    aspect: deck.size.cx / deck.size.cy,
    createShow,
    render,
    media: mediaOf,
    notes: (n: number) => { try { return getSlideNotes(archive, deck.slides[n - 1]!.path) } catch { return '' } },
    close() { root.unmount(); host.remove(); loader.dispose(); for (const u of blobUrls) URL.revokeObjectURL(u) },
  }
}
