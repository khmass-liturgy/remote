@echo off
title Slide Remote - PowerPoint helper
set "SR_SELF=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((Get-Content -LiteralPath $env:SR_SELF -Raw -Encoding UTF8) -replace '(?s)^.*?\n#PSSTART', '')"
echo.
pause
exit /b
#PSSTART
# ---------------------------------------------------------------------------
# Slide Remote - PowerPoint helper (Windows PowerShell 5.1, no install needed)
#
# Lets the Slide Remote web page (running in a browser on this PC) control the
# real PowerPoint program: next / previous / go to slide / black screen /
# laser pointer, and read slide thumbnails and speaker notes.
#
# It only listens on this PC (http://localhost:8765) and only answers the
# web origins listed in $Allowed. Close this window to stop it.
# ---------------------------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Port = 8765
$Allowed = @('https://khmass-liturgy.github.io')
$Cache = Join-Path $env:TEMP 'slide-remote-thumbs'
$script:CacheKey = ''

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class SRWin {
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
}
"@
[void][SRWin]::SetProcessDPIAware()

function Get-Ctx {
  $app = $null
  try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { return $null }
  $c = @{ app = $app; win = $null; view = $null; pres = $null }
  try {
    if ($app.SlideShowWindows.Count -ge 1) {
      $c.win = $app.SlideShowWindows.Item(1)
      $c.view = $c.win.View
      $c.pres = $c.win.Presentation
    }
  } catch { $c.win = $null; $c.view = $null; $c.pres = $null }
  if (-not $c.pres) { try { $c.pres = $app.ActivePresentation } catch {} }
  if (-not $c.pres) { return $null }
  return $c
}

function Get-State($c) {
  $total = $c.pres.Slides.Count
  $cur = 1; $black = $false; $running = $false
  if ($c.view) {
    $running = $true
    try { $cur = $c.view.Slide.SlideIndex } catch { $cur = $total }
    try { $black = ($c.view.State -eq 3) } catch {}
  } else {
    try { $cur = $c.app.ActiveWindow.View.Slide.SlideIndex } catch {}
  }
  return @{ ok = $true; name = [string]$c.pres.Name; total = [int]$total; cur = [int]$cur; black = $black; running = $running }
}

function Get-ThumbPath($c, [int]$n) {
  $key = [string]$c.pres.FullName + '|' + $c.pres.Slides.Count
  if ($key -ne $script:CacheKey) {
    if (Test-Path $Cache) { Remove-Item (Join-Path $Cache '*.jpg') -Force -ErrorAction SilentlyContinue }
    else { [void](New-Item -ItemType Directory -Path $Cache) }
    $script:CacheKey = $key
  }
  $path = Join-Path $Cache ("s$n.jpg")
  if (-not (Test-Path $path)) {
    $w = 960
    $h = [int](960 * $c.pres.PageSetup.SlideHeight / $c.pres.PageSetup.SlideWidth)
    $c.pres.Slides.Item($n).Export($path, 'JPG', $w, $h)
  }
  return $path
}

function Get-Notes($c) {
  $r = @{}
  $count = $c.pres.Slides.Count
  for ($i = 1; $i -le $count; $i++) {
    try {
      $t = [string]$c.pres.Slides.Item($i).NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
      if ($t.Trim()) { $r["$i"] = ($t -replace '[\r\v]', "`n") }
    } catch {}
  }
  return $r
}

function Set-Laser($c, $q) {
  if (-not $c.view) { return }
  if ($q['on'] -ne '1') { try { $c.view.LaserPointerEnabled = $false } catch {}; return }
  try { $c.view.LaserPointerEnabled = $true } catch {}
  $x = [double]$q['x']; $y = [double]$q['y']
  $r = New-Object SRWin+RECT
  if (-not [SRWin]::GetWindowRect([IntPtr][int64]$c.win.HWND, [ref]$r)) { return }
  $W = $r.R - $r.L; $H = $r.B - $r.T
  if ($W -le 0 -or $H -le 0) { return }
  $ar = $c.pres.PageSetup.SlideWidth / $c.pres.PageSetup.SlideHeight
  if (($W / $H) -gt $ar) { $sh = $H; $sw = $H * $ar } else { $sw = $W; $sh = $W / $ar }
  $px = $r.L + ($W - $sw) / 2 + $x * $sw
  $py = $r.T + ($H - $sh) / 2 + $y * $sh
  [void][SRWin]::SetCursorPos([int]$px, [int]$py)
}

function Invoke-Cmd($c, $q) {
  switch ($q['c']) {
    'next'  { if ($c.view) { $c.view.Next() } else { [void]$c.pres.SlideShowSettings.Run() } }
    'prev'  { if ($c.view) { $c.view.Previous() } }
    'go'    {
      $n = [int]$q['n']
      if ($n -ge 1 -and $n -le $c.pres.Slides.Count) {
        if ($c.view) { $c.view.GotoSlide($n) } else { $c.app.ActiveWindow.View.GotoSlide($n) }
      }
    }
    'black' { if ($c.view) { if ($q['on'] -eq '1') { $c.view.State = 3 } else { $c.view.State = 1 } } }
    'start' { if (-not $c.view) { [void]$c.pres.SlideShowSettings.Run() } }
    'end'   { if ($c.view) { $c.view.Exit() } }
    'laser' { Set-Laser $c $q }
  }
}

function Send-Bytes($res, [int]$status, [string]$type, [byte[]]$bytes) {
  $res.StatusCode = $status
  $res.ContentType = $type
  $res.ContentLength64 = $bytes.Length
  $res.OutputStream.Write($bytes, 0, $bytes.Length)
}
function Send-Json($res, [int]$status, $obj) {
  $json = ConvertTo-Json $obj -Compress -Depth 4
  Send-Bytes $res $status 'application/json; charset=utf-8' ([Text.Encoding]::UTF8.GetBytes($json))
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
try { $listener.Start() } catch {
  Write-Host "Could not start on port $Port. Is the helper already running?" -ForegroundColor Red
  Write-Host $_.Exception.Message
  return
}
Write-Host ''
Write-Host '  Slide Remote - PowerPoint helper is running.' -ForegroundColor Green
Write-Host "  Listening on http://localhost:$Port"
Write-Host ''
Write-Host '  1. Open your presentation in PowerPoint.'
Write-Host '  2. On the Slide Remote web page (PC), click the PowerPoint connect button.'
Write-Host '  3. Keep this window open. Close it to stop.'
Write-Host ''

while ($listener.IsListening) {
  $hc = $listener.GetContext()
  $req = $hc.Request; $res = $hc.Response
  try {
    $origin = $req.Headers['Origin']
    $okOrigin = (-not $origin) -or ($Allowed -contains $origin) -or ($origin -match '^http://(localhost|127\.0\.0\.1)(:\d+)?$')
    if ($origin -and $okOrigin) {
      $res.Headers.Add('Access-Control-Allow-Origin', $origin)
      $res.Headers.Add('Vary', 'Origin')
      $res.Headers.Add('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
      $res.Headers.Add('Access-Control-Allow-Headers', 'Content-Type')
      $res.Headers.Add('Access-Control-Allow-Private-Network', 'true')
      $res.Headers.Add('Access-Control-Max-Age', '600')
    }
    $res.Headers.Add('Cache-Control', 'no-store')
    $path = $req.Url.AbsolutePath
    $q = $req.QueryString

    if (-not $okOrigin) { Send-Json $res 403 @{ ok = $false; error = 'origin not allowed' } }
    elseif ($req.HttpMethod -eq 'OPTIONS') { $res.StatusCode = 204 }
    else {
      $c = Get-Ctx
      if (-not $c) { Send-Json $res 200 @{ ok = $false; error = 'no-presentation' } }
      elseif ($path -eq '/state') { Send-Json $res 200 (Get-State $c) }
      elseif ($path -eq '/notes') { Send-Json $res 200 @{ ok = $true; notes = (Get-Notes $c) } }
      elseif ($path -eq '/thumb') {
        $n = [int]$q['n']
        if ($n -lt 1 -or $n -gt $c.pres.Slides.Count) { Send-Json $res 404 @{ ok = $false; error = 'no such slide' } }
        else { Send-Bytes $res 200 'image/jpeg' ([IO.File]::ReadAllBytes((Get-ThumbPath $c $n))) }
      }
      elseif ($path -eq '/cmd') {
        if ($req.HttpMethod -ne 'POST') { Send-Json $res 405 @{ ok = $false; error = 'use POST' } }
        else { Invoke-Cmd $c $q; Send-Json $res 200 (Get-State $c) }
      }
      else { Send-Json $res 404 @{ ok = $false; error = 'not found' } }
    }
  } catch {
    Write-Host ('  error: ' + $_.Exception.Message) -ForegroundColor Yellow
    try { Send-Json $res 500 @{ ok = $false; error = [string]$_.Exception.Message } } catch {}
  } finally {
    try { $res.Close() } catch {}
  }
}
