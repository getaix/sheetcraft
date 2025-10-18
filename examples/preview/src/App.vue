<script lang="ts" setup>
// 预览页面（TypeScript + Vue3 Composition API）
// 功能：
// 1) 仅通过本地选择/拖拽 .xlsx 文件进行预览；
// 2) 自动修复部分工作簿图片关系路径（/xl/ 前缀缺失），提升前端预览兼容性；
// 3) 统一中文提示与错误信息；
// 4) 提升 UI 与交互体验（文件信息、清空、拖拽高亮）。

import { ref, computed, onMounted, onBeforeUnmount } from 'vue'
import type { Ref } from 'vue'
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate'
import VueOfficeExcel from '@vue-office/excel'
import '@vue-office/excel/lib/index.css'
import html2canvas from 'html2canvas'
import { jsPDF } from 'jspdf'

// 受支持的 MIME 类型
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

// 组件状态
const excel: Ref<Blob | File | null> = ref(null)
const filename = ref<string>('')
const message = ref<string>('请选择或拖拽 .xlsx 文件进行预览')
const messageType = ref<'info' | 'error' | 'success'>('info')
const dragging = ref<boolean>(false)

// 预览区是否可见
const hasFile = computed<boolean>(() => !!excel.value)
const viewerRef = ref<HTMLDivElement | null>(null)
let lastUnzippedFiles: Record<string, Uint8Array> | null = null


// —— 事件处理 ——
function setMessage(text: string, type: 'info' | 'error' | 'success' = 'info') {
  message.value = text
  messageType.value = type
}

function onFileChange(e: Event) {
  const input = e.target as HTMLInputElement
  const file = input.files?.[0]
  if (!file) return
  handleFile(file)
}

function onDrop(e: DragEvent) {
  e.preventDefault()
  dragging.value = false
  const file = e.dataTransfer?.files?.[0]
  if (file) handleFile(file)
}

function onDragOver(e: DragEvent) {
  e.preventDefault()
  dragging.value = true
}

function onDragLeave(e: DragEvent) {
  e.preventDefault()
  dragging.value = false
}

function clearFile() {
  excel.value = null
  filename.value = ''
  lastUnzippedFiles = null
  setMessage('请选择或拖拽 .xlsx 文件进行预览', 'info')
}

async function handleFile(file: File) {
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    return setMessage('仅支持 .xlsx 文件', 'error')
  }
  filename.value = file.name
  setMessage('正在解析并加载文件…', 'info')

  try {
    const patched = await patchXlsxImagePaths(file)
    excel.value = patched
    if (patched === file) {
      setMessage('文件已加载（无需图片兼容修复）', 'success')
    } else {
      setMessage('文件加载完成', 'success')
    }
    console.log('选择文件', file.name, file.size, '已完成图片关系兼容处理')
  } catch (err) {
    // 修复异常时不阻断预览，直接使用原文件
    console.warn('图片关系修复异常，将直接预览原文件', err)
    excel.value = file
    lastUnzippedFiles = null
    setMessage('文件已加载（兼容修复跳过）', 'info')
  }
}

// —— 图片关系修复 ——
// 某些工具导出的 xlsx 在关系文件中使用了相对路径（如 ../media 或 media），
// 前端解析库通常要求以 /xl/ 为根的绝对路径。此处对相关 rels 进行轻量修复。
async function patchXlsxImagePaths(src: Blob): Promise<Blob> {
   const buf = new Uint8Array(await src.arrayBuffer())
   const files = unzipSync(buf)
  lastUnzippedFiles = files

  let touched = false

  const patchRelsText = (text: string): string => {
    // 修复 Target 指向 media 的关系，统一为 /xl/media/xxx
    const replaced = text
      .replace(/Target="\.??\/media\//g, 'Target="/xl/media/')
      .replace(/Target="media\//g, 'Target="/xl/media/')
    if (replaced !== text) touched = true
    return replaced
  }

  const patchDrawingTarget = (text: string): string => {
    // 修复指向 drawings 的关系，统一为 /xl/drawings/xxx
    const replaced = text
      .replace(/Target="\.??\/drawings\//g, 'Target="/xl/drawings/')
      .replace(/Target="drawings\//g, 'Target="/xl/drawings/')
    if (replaced !== text) touched = true
    return replaced
  }

  Object.keys(files).forEach((name) => {
    if (name.endsWith('.rels')) {
      const contentU8 = files[name]
      const xml = strFromU8(contentU8)
      const patched = patchDrawingTarget(patchRelsText(xml))
      if (patched !== xml) {
        files[name] = strToU8(patched)
      }
    }
  })

  if (!touched) {
    // 无改动：直接返回原文件，以避免误报失败
    return src
  }

  const out = zipSync(files, { level: 0 })
  const arrayBuffer = out.buffer.slice(0) as ArrayBuffer
  return new Blob([arrayBuffer], { type: XLSX_MIME })
}

// —— 预览事件回调 ——
 function renderedHandler() {
   console.log('渲染完成')
   fitToWidth()
   injectPlaceholderImages()
   injectDrawingImages()
 }
 
 function errorHandler(e: unknown) {
   console.error('渲染失败', e)
   setMessage('预览失败，请尝试下载后用 Excel 打开', 'error')
 }
// —— 全宽自适应缩放 ——
function fitToWidth() {
  const container = viewerRef.value
  if (!container) return
  const content = container.firstElementChild as HTMLElement | null
  if (!content) return
  const cw = container.clientWidth
  const sw = content.scrollWidth || content.clientWidth || cw
  if (!sw) return
  const scale = cw / sw
  content.style.transformOrigin = 'left top'
  content.style.transform = `scale(${scale})`
}

function onResize() { if (hasFile.value) fitToWidth() }
onMounted(() => { window.addEventListener('resize', onResize) })
onBeforeUnmount(() => { window.removeEventListener('resize', onResize) })

// —— 占位图片解析与渲染 ——
function injectPlaceholderImages() {
  const container = viewerRef.value
  if (!container) return
  const cells = container.querySelectorAll('td')
  const marker = /__SHEETCRAFT_IMG__\s*\{[^}]+\}/
  for (const td of Array.from(cells)) {
    const text = td.textContent || ''
    if (!marker.test(text)) continue
    const jsonMatch = text.match(/\{.*\}/)
    let path: string | null = null
    try { path = jsonMatch ? JSON.parse(jsonMatch[0]).path : null } catch { path = null }
    if (!path) continue
    td.style.position = 'relative'
    td.style.color = 'transparent' // 隐藏占位文本
    const img = document.createElement('img')
    img.style.position = 'absolute'
    img.style.left = '2px'
    img.style.top = '2px'
    img.style.right = '2px'
    img.style.bottom = '2px'
    img.style.maxWidth = '100%'
    img.style.maxHeight = '100%'
    img.style.objectFit = 'contain'
    // 从解压缓存查找同名资源
    let srcUrl: string | null = null
    const candidates = [`xl/media/${path}`, `xl/media/${path.replace(/^.*\//, '')}`]
    if (lastUnzippedFiles) {
      for (const key of Object.keys(lastUnzippedFiles)) {
        if (candidates.some(c => key.endsWith(c))) {
          const u8 = lastUnzippedFiles[key]
          const ab = u8.buffer.slice(0) as ArrayBuffer
          const blob = new Blob([ab])
          srcUrl = URL.createObjectURL(blob)
          break
        }
      }
    }
    // 兜底：尝试从 public 路径获取
    if (!srcUrl) srcUrl = `/${path.replace(/^.*\//, '')}`
    img.src = srcUrl
    td.appendChild(img)
  }
}

// —— drawings 图片解析与叠加展示（无需占位符） ——
function injectDrawingImages() {
  try {
    if (!lastUnzippedFiles) return
    const container = viewerRef.value
    if (!container) return
    const content = container.firstElementChild as HTMLElement | null
    if (!content) return
    content.style.position = 'relative'

    // 清理旧的覆盖元素，避免重复添加
    for (const old of Array.from(content.querySelectorAll('.sheetcraft-img-overlay'))) {
      old.remove()
    }

    // 简化：当前仅处理第一个工作表
    const sheetXmlKey = Object.keys(lastUnzippedFiles).find(k => /xl\/worksheets\/sheet1\.xml$/.test(k))
    if (!sheetXmlKey) return
    const sheetXml = strFromU8(lastUnzippedFiles[sheetXmlKey])
    const ridMatch = sheetXml.match(/<drawing[^>]*r:id="([^"]+)"/)
    if (!ridMatch) return
    const sheetRelsKey = Object.keys(lastUnzippedFiles).find(k => /xl\/worksheets\/_rels\/sheet1\.xml\.rels$/.test(k))
    if (!sheetRelsKey) return
    const sheetRels = strFromU8(lastUnzippedFiles[sheetRelsKey])
    const drawingTargetMatch = sheetRels.match(new RegExp(`<Relationship[^>]*Id="${ridMatch[1]}"[^>]*Target="([^"]+)"`))
    if (!drawingTargetMatch) return
    // 关系可能是相对路径，统一去掉前缀
    const drawingPath = drawingTargetMatch[1].replace(/^\/?/, 'xl/')
    const drawingXmlKey = Object.keys(lastUnzippedFiles).find(k => k.endsWith(drawingPath))
    if (!drawingXmlKey) return
    const drawingXml = strFromU8(lastUnzippedFiles[drawingXmlKey])

    // 解析 anchors 与 embed rId
    const anchors: Array<{ c1: number, r1: number, c2: number, r2: number, rid: string }> = []
    const re = /<(?:xdr:)?twoCellAnchor[\s\S]*?<\s*(?:xdr:)?from[\s\S]*?<\s*(?:xdr:)?col\s*>\s*(\d+)\s*<\/(?:xdr:)?col\s*>[\s\S]*?<\s*(?:xdr:)?row\s*>\s*(\d+)\s*<\/(?:xdr:)?row\s*>[\s\S]*?<\s*(?:xdr:)?to[\s\S]*?<\s*(?:xdr:)?col\s*>\s*(\d+)\s*<\/(?:xdr:)?col\s*>[\s\S]*?<\s*(?:xdr:)?row\s*>\s*(\d+)\s*<\/(?:xdr:)?row\s*>[\s\S]*?<a:blip[^>]*r:embed="([^"]+)"[\s\S]*?<\/(?:xdr:)?twoCellAnchor>/g
    let m: RegExpExecArray | null
    while ((m = re.exec(drawingXml)) !== null) {
      anchors.push({ c1: parseInt(m[1], 10), r1: parseInt(m[2], 10), c2: parseInt(m[3], 10), r2: parseInt(m[4], 10), rid: m[5] })
    }
    if (anchors.length === 0) return

    // drawing rels：rId -> 媒体路径
    const drawingRelsKey = drawingXmlKey.replace(/drawing(\d+)\.xml$/, '_rels/drawing$1.xml.rels')
    const relsU8 = lastUnzippedFiles[drawingRelsKey]
    if (!relsU8) return
    const relsXml = strFromU8(relsU8)

    const mediaPathOf = (rid: string): string | null => {
      const mm = relsXml.match(new RegExp(`<Relationship[^>]*Id="${rid}"[^>]*Target="([^"]+)"`))
      if (!mm) return null
      const t = mm[1]
      return t.startsWith('/') ? t.slice(1) : `xl/${t}`
    }

    // DOM: 查找 table 以定位单元格
    const table = content.querySelector('table') as HTMLTableElement | null
    if (!table) return
    const trs = table.querySelectorAll('tr')

    // 辅助：获取某个单元格的矩形（基于 0 基索引）
    const cellRect = (r0: number, c0: number): DOMRect | null => {
      const tr = trs[r0]
      if (!tr) return null
      const tds = tr.querySelectorAll('td')
      const td = tds[c0]
      if (!td) return null
      return td.getBoundingClientRect()
    }

    const contentRect = content.getBoundingClientRect()

    for (const a of anchors) {
      const r0 = a.r1 // 0 基
      const c0 = a.c1 // 0 基
      const r1 = a.r2 // 0 基（包含端点）
      const c1 = a.c2 // 0 基（包含端点）
      const p0 = cellRect(r0, c0)
      const p1 = cellRect(r1, c1)
      if (!p0 || !p1) continue
      // 计算覆盖区域
      const left = p0.left - contentRect.left
      const top = p0.top - contentRect.top
      const width = (p1.right - p0.left)
      const height = (p1.bottom - p0.top)

      const media = mediaPathOf(a.rid)
      if (!media) continue
      const mediaKey = Object.keys(lastUnzippedFiles).find(k => k.endsWith(media))
      if (!mediaKey) continue
      const u8 = lastUnzippedFiles[mediaKey]
      const ab = u8.buffer.slice(0) as ArrayBuffer
      const blob = new Blob([ab])
      const url = URL.createObjectURL(blob)

      const img = document.createElement('img')
      img.className = 'sheetcraft-img-overlay'
      img.style.position = 'absolute'
      img.style.left = `${left}px`
      img.style.top = `${top}px`
      img.style.width = `${width}px`
      img.style.height = `${height}px`
      img.style.objectFit = 'contain'
      img.style.pointerEvents = 'none'
      img.src = url
      content.appendChild(img)
    }
  } catch (e) {
    console.warn('绘图图片解析失败或不支持，跳过前端覆盖显示', e)
  }
}

// —— 导出 PNG / PDF ——
function getPreviewContentElement(): HTMLElement | null {
  const container = viewerRef.value
  if (!container) return null
  const el = container.firstElementChild as HTMLElement | null
  return el
}

async function exportPNG() {
  const el = getPreviewContentElement()
  if (!el) {
    setMessage('无法导出：内容未渲染', 'error')
    return
  }
  const canvas = await html2canvas(el, { scale: 2, useCORS: true, backgroundColor: '#ffffff' })
  canvas.toBlob((blob: Blob | null) => {
    if (!blob) return
    const a = document.createElement('a')
    a.href = URL.createObjectURL(blob)
    a.download = `${filename.value || 'export'}.png`
    a.click()
    URL.revokeObjectURL(a.href)
  }, 'image/png', 1.0)
}

async function exportPDF() {
  const el = getPreviewContentElement()
  if (!el) {
    setMessage('无法导出：内容未渲染', 'error')
    return
  }
  const canvas = await html2canvas(el, { scale: 2, useCORS: true, backgroundColor: '#ffffff' })
  const imgData = canvas.toDataURL('image/png')
  const pdf = new jsPDF({ unit: 'pt', format: 'a4' })
  const pageWidth = pdf.internal.pageSize.getWidth()
  const pageHeight = pdf.internal.pageSize.getHeight()
  const contentWidth = canvas.width
  const contentHeight = canvas.height
  const ratio = Math.min(pageWidth / contentWidth, pageHeight / contentHeight)
  const width = contentWidth * ratio
  const height = contentHeight * ratio
  const x = (pageWidth - width) / 2
  const y = 20
  pdf.addImage(imgData, 'PNG', x, y, width, height)
  pdf.save(`${filename.value || 'export'}.pdf`)
}
</script>

<template>
  <div class="page">
    <header class="toolbar">
      <div class="left">
        <label class="btn primary">
          选择文件
          <input type="file" accept=".xlsx" @change="onFileChange" />
        </label>
        <button class="btn" @click="clearFile">清空</button>
        <button v-if="hasFile" class="btn" @click="exportPNG">导出PNG</button>
        <button v-if="hasFile" class="btn" @click="exportPDF">导出PDF</button>
      </div>
      <div class="right">
        <span :class="['msg', messageType]">{{ message }}</span>
        <span v-if="filename" class="filename">{{ filename }}</span>
      </div>
    </header>

    <section class="content" @drop="onDrop" @dragover="onDragOver" @dragleave="onDragLeave">
      <div v-if="!hasFile" :class="['dropzone', { dragging }]">
        <p class="title">拖拽 .xlsx 文件到此处或点击“选择文件”</p>
        <p class="desc">支持本地工作簿预览，并自动进行图片关系路径兼容修复</p>
      </div>
      <div v-else class="preview" ref="viewerRef">
        <VueOfficeExcel
          :src="excel"
          style="height: 100%; width: 100%;"
          @rendered="renderedHandler"
          @error="errorHandler"
        />
      </div>
    </section>
  </div>
</template>

<style scoped>
/* —— 页面布局与配色（浅色主题） —— */
.page { height: 100vh; display: flex; flex-direction: column; background: #f6f8fb; }
.toolbar { display: flex; align-items: center; justify-content: space-between; padding: 10px 12px; background: #ffffff; border-bottom: 1px solid #e5e8eb; }
.left { display:flex; align-items:center; gap:8px; }
.right { display:flex; align-items:center; gap:12px; }

.btn { padding: 6px 12px; border: 1px solid #cbd5e1; background: #fff; color:#334155; border-radius: 6px; cursor: pointer; transition: all .2s ease; }
.btn:hover { background: #f1f5f9; }
.btn.primary { background: #2563eb; border-color: #2563eb; color: #fff; }
.btn.primary:hover { background: #1d4ed8; }

.btn input[type="file"] { display: none; }

.msg { font-size: 13px; }
.msg.info { color: #64748b; }
.msg.success { color: #16a34a; }
.msg.error { color: #ef4444; }

.filename { font-size: 12px; color: #475569; background: #e2e8f0; padding: 2px 6px; border-radius: 4px; }

.content { flex: 1; min-height: 0; padding: 12px; }
.dropzone { height: 100%; border: 2px dashed #cbd5e1; border-radius: 10px; display:flex; align-items:center; justify-content:center; flex-direction:column; color:#64748b; background: #f8fafc; transition: border-color .2s ease, background .2s ease; }
.dropzone.dragging { border-color: #2563eb; background: #eef2ff; }
.dropzone .title { font-size: 16px; margin-bottom: 6px; }
.dropzone .desc { font-size: 13px; }

.preview { height: 100%; background: #fff; border: 1px solid #e5e8eb; border-radius: 8px; overflow: hidden; }
</style>
