'use client'
// dashboard/src/app/page.tsx

import { useEffect, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import { supabase } from '../lib/supabase'

type Tab          = 'dashboard' | 'egresados' | 'secretarios' | 'validaciones'
type FiltroEstado = 'todos' | 'activo' | 'vencido' | 'inactivo'
type FiltroFecha  = 'hoy' | 'semana' | 'mes' | 'todo'

function pad(n: number) { return n.toString().padStart(2, '0') }
function fechaHoy() {
  const d = new Date()
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`
}
function sumar365(base: string) {
  const d = new Date(base + 'T00:00:00')
  d.setDate(d.getDate() + 365)
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`
}
function diasRestantes(fechaVenc: string): number {
  const hoy  = new Date(fechaHoy() + 'T00:00:00')
  const venc = new Date(fechaVenc  + 'T00:00:00')
  return Math.ceil((venc.getTime() - hoy.getTime()) / 86400000)
}
function fmtFechaLegible(str: string | null) {
  if (!str) return '—'
  const [y, m, d] = str.split('-')
  const meses = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic']
  return `${d} ${meses[parseInt(m)-1]} ${y}`
}

// ── Exportar Excel COMPLETO ───────────────────────────────────
function exportarExcel(datos: any[], nombre: string) {
  const cols = [
    'cedula','nombre_completo','email','telefono','estado',
    'fecha_expedicion','fecha_vencimiento','created_at',
    'ciudad_nacimiento','direccion','fecha_nacimiento',
    'titulo_pregrado','institucion_pregrado','fecha_grado_pregrado',
    'titulo_posgrado','institucion_posgrado','fecha_grado_posgrado',
    'empresa','cargo','hobbies',
  ]
  const headers = [
    'Cédula','Nombre completo','Email','Teléfono','Estado',
    'Inicio membresía','Vencimiento membresía','Fecha registro',
    'Ciudad nacimiento','Dirección','Fecha nacimiento',
    'Título pregrado','Institución pregrado','Fecha grado pregrado',
    'Título posgrado','Institución posgrado','Fecha grado posgrado',
    'Empresa','Cargo','Hobbies',
  ]
  // Crear workbook con xlsx
  const ws_data = [headers, ...datos.map(row => cols.map(c => row[c] ?? ''))]
  const ws = XLSX.utils.aoa_to_sheet(ws_data)
  // Ancho columnas
  ws['!cols'] = cols.map((c, i) => ({ wch: i < 4 ? 18 : i < 8 ? 20 : 24 }))
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Egresados')
  // Hoja resumen
  const resumen = [
    ['REPORTE ASEDUIS', ''],
    ['Fecha', fechaHoy()],
    ['Total', datos.length],
    ['Activos', datos.filter(e => e.estado === 'activo').length],
    ['Vencidos', datos.filter(e => e.estado === 'vencido').length],
    ['Inactivos', datos.filter(e => e.estado === 'inactivo').length],
  ]
  const ws2 = XLSX.utils.aoa_to_sheet(resumen)
  XLSX.utils.book_append_sheet(wb, ws2, 'Resumen')
  XLSX.writeFile(wb, `${nombre}.xlsx`)
}

// ── Exportar PDF ──────────────────────────────────────────────
function exportarPDF(datos: any[], nombre: string) {
  const hoy = new Date().toLocaleDateString('es-CO', { day:'2-digit', month:'long', year:'numeric' })
  const filas = datos.map(eg => `
    <tr>
      <td>${eg.cedula??''}</td>
      <td><strong>${eg.nombre_completo??''}</strong></td>
      <td>${eg.email??''}</td>
      <td>${eg.telefono??'—'}</td>
      <td>${eg.empresa??'—'}</td>
      <td>${eg.cargo??'—'}</td>
      <td><span class="badge ${eg.estado}">${eg.estado??''}</span></td>
      <td style="color:${eg.estado==='vencido'?'#D97706':eg.estado==='inactivo'?'#BE1522':'#16A34A'};font-weight:700">
        ${eg.fecha_vencimiento??'—'}
      </td>
    </tr>`).join('')
  const html = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>${nombre}</title>
  <style>
    html, body { margin:0; padding:0; width:100%; height:100%; }html, body { margin:0; padding:0; width:100%; height:100%; }
    *{margin:0;padding:0;box-sizing:border-box}
    body{font-family:"Segoe UI",sans-serif;color:#111;font-size:10.5px}
    .header{background:#BE1522;color:#fff;padding:18px 24px;display:flex;justify-content:space-between;align-items:center;margin-bottom:16px}
    .header h1{font-size:20px;font-weight:800;letter-spacing:1px}.header p{font-size:11px;opacity:.8;margin-top:3px}
    .meta{padding:0 24px 14px;display:flex;gap:20px;flex-wrap:wrap}
    .meta-item{background:#FDE8EA;border-radius:8px;padding:8px 14px;text-align:center}
    .meta-num{font-size:20px;font-weight:800;color:#BE1522}
    .meta-lbl{font-size:10px;color:#6B7280;margin-top:2px}
    table{width:100%;border-collapse:collapse;font-size:10px}
    th{background:#FDE8EA;color:#BE1522;font-weight:700;text-transform:uppercase;font-size:9px;letter-spacing:.4px;padding:8px 10px;text-align:left;border-bottom:2px solid #F0D4D6}
    td{padding:7px 10px;border-bottom:1px solid #F5F5F5;vertical-align:middle}
    tr:nth-child(even) td{background:#FDF8F8}
    .badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:9px;font-weight:700}
    .badge.activo{background:#DCFCE7;color:#16A34A}
    .badge.vencido{background:#FEF3C7;color:#D97706}
    .badge.inactivo{background:#FDE8EA;color:#BE1522}
    .footer{margin-top:16px;padding:10px 24px;border-top:1px solid #F0D4D6;display:flex;justify-content:space-between;font-size:10px;color:#94A3B8}
    @media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
  </style></head><body>
  <div class="header">
    <div><h1>ASEDUIS</h1><p>Reporte de Egresados — Membresías</p></div>
    <div style="text-align:right">
      <div style="font-size:15px;font-weight:700">${datos.length} egresados</div>
      <div style="font-size:10px;opacity:.8">${hoy}</div>
    </div>
  </div>
  <div class="meta">
    <div class="meta-item"><div class="meta-num">${datos.length}</div><div class="meta-lbl">Total</div></div>
    <div class="meta-item"><div class="meta-num" style="color:#16A34A">${datos.filter(e=>e.estado==='activo').length}</div><div class="meta-lbl">Activos</div></div>
    <div class="meta-item"><div class="meta-num" style="color:#D97706">${datos.filter(e=>e.estado==='vencido').length}</div><div class="meta-lbl">Vencidos</div></div>
    <div class="meta-item"><div class="meta-num" style="color:#94A3B8">${datos.filter(e=>e.estado==='inactivo').length}</div><div class="meta-lbl">Inactivos</div></div>
  </div>
  <table>
    <thead><tr><th>Cédula</th><th>Nombre</th><th>Email</th><th>Teléfono</th><th>Empresa</th><th>Cargo</th><th>Estado</th><th>Vencimiento</th></tr></thead>
    <tbody>${filas}</tbody>
  </table>
  <div class="footer"><span>Generado: ${hoy} · Sistema ASEDUIS</span><span>Confidencial — uso interno</span></div>
  <script>window.onload=()=>window.print()</script>
  </body></html>`
  const blob = new Blob([html], { type: 'text/html;charset=utf-8;' })
  window.open(URL.createObjectURL(blob), '_blank')
}

// ── Mini gráfica de barras SVG ────────────────────────────────
function BarChart({ data }: { data: { label: string; value: number; color: string }[] }) {
  const max = Math.max(...data.map(d => d.value), 1)
  const H = 80, W = 200, barW = 36, gap = 14
  return (
    <svg width={W} height={H + 24} viewBox={`0 0 ${W} ${H + 24}`}>
      {data.map((d, i) => {
        const x = i * (barW + gap) + gap
        const h = (d.value / max) * H
        const y = H - h
        return (
          <g key={i}>
            <rect x={x} y={y} width={barW} height={h} rx={4} fill={d.color} opacity={0.85} />
            <text x={x + barW/2} y={H + 14} textAnchor="middle" fontSize={9} fill="#94A3B8" fontFamily="sans-serif">{d.label}</text>
            <text x={x + barW/2} y={y - 4} textAnchor="middle" fontSize={10} fill={d.color} fontWeight="700" fontFamily="sans-serif">{d.value}</text>
          </g>
        )
      })}
    </svg>
  )
}

// ── Gráfica de dona SVG ───────────────────────────────────────
function DonutChart({ activos, vencidos, inactivos, total }: { activos:number; vencidos:number; inactivos:number; total:number }) {
  if (total === 0) return <div style={{ width:100, height:100, borderRadius:'50%', background:'#F5F5F5', display:'flex', alignItems:'center', justifyContent:'center', fontSize:12, color:'#94A3B8' }}>Sin datos</div>
  const R = 36, cx = 50, cy = 50, stroke = 14
  const circ = 2 * Math.PI * R
  const pA = (activos  / total) * circ
  const pV = (vencidos / total) * circ
  const pI = (inactivos/ total) * circ
  return (
    <div style={{ position:'relative', width:100, height:100 }}>
      <svg width={100} height={100} viewBox="0 0 100 100">
        <circle cx={cx} cy={cy} r={R} fill="none" stroke="#F0D4D6" strokeWidth={stroke} />
        <circle cx={cx} cy={cy} r={R} fill="none" stroke="#16A34A" strokeWidth={stroke}
          strokeDasharray={`${pA} ${circ}`} strokeDashoffset={circ * 0.25} strokeLinecap="butt" />
        <circle cx={cx} cy={cy} r={R} fill="none" stroke="#D97706" strokeWidth={stroke}
          strokeDasharray={`${pV} ${circ}`} strokeDashoffset={circ * 0.25 - pA} strokeLinecap="butt" />
        <circle cx={cx} cy={cy} r={R} fill="none" stroke="#94A3B8" strokeWidth={stroke}
          strokeDasharray={`${pI} ${circ}`} strokeDashoffset={circ * 0.25 - pA - pV} strokeLinecap="butt" />
      </svg>
      <div style={{ position:'absolute', inset:0, display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center' }}>
        <div style={{ fontSize:18, fontWeight:800, color:'#BE1522', lineHeight:1 }}>{total}</div>
        <div style={{ fontSize:9, color:'#94A3B8', marginTop:2 }}>total</div>
      </div>
    </div>
  )
}

const ESTADO_BADGE: Record<string, string> = {
  activo: 'badge active', vencido: 'badge expired', inactivo: 'badge inactive',
}
const RESULTADO_EMOJI: Record<string, string> = {
  activo: '✅', vencido: '⚠️', inactivo: '❌', no_encontrado: '❓',
}
const RESULTADO_BADGE: Record<string, string> = {
  activo: 'badge active', vencido: 'badge expired',
  inactivo: 'badge inactive', no_encontrado: 'badge inactive',
}

function normalizarHeader(h: string): string {
  return h.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, ' ').trim().replace(/\s+/g, '_')
}
const HEADER_MAP: Record<string, string> = {
  cedula: 'cedula', numero_de_cedula: 'cedula', documento: 'cedula',
  nombre_completo: 'nombre', nombre: 'nombre', nombres: 'nombre',
  correo_electronico: 'email', correo: 'email', email: 'email',
  telefono: 'telefono', celular: 'telefono',
}

const CSS = `
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:"Segoe UI",system-ui,sans-serif;background:#F9F1F1;color:#111827}
  .layout{display:flex;min-height:100vh}
  .sidebar{width:240px;background:#BE1522;min-height:100vh;display:flex;flex-direction:column;position:sticky;top:0;height:100vh;overflow-y:auto;flex-shrink:0}
  .sidebar-backdrop{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:199}
  @media(max-width:768px){
    .sidebar{position:fixed;top:0;left:0;height:100vh;z-index:200;transform:translateX(-100%);transition:transform .25s}
    .sidebar.open{transform:translateX(0)}
    .sidebar-backdrop.show{display:block}
  }
  .main{flex:1;display:flex;flex-direction:column;min-width:0;overflow:hidden}
  .nav-item{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:10px;cursor:pointer;margin-bottom:2px;transition:background .15s;user-select:none}
  .nav-item:hover{background:rgba(255,255,255,.12)}
  .nav-item.active{background:rgba(255,255,255,.22)}
  .nav-lbl{font-size:13px;color:rgba(255,255,255,.7);font-weight:500}
  .nav-item.active .nav-lbl{color:#fff;font-weight:700}
  .topbar{background:#fff;border-bottom:1px solid #E5E7EB;padding:14px 28px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;gap:10px;flex-wrap:wrap}
  .topbar-actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
  .menu-btn{display:none;background:none;border:none;cursor:pointer;font-size:22px;padding:4px 8px;border-radius:8px;color:#BE1522;flex-shrink:0}
  @media(max-width:768px){.menu-btn{display:flex;align-items:center}.topbar{padding:12px 16px}}
  .content{flex:1;padding:24px 28px;overflow-y:auto}
  @media(max-width:768px){.content{padding:16px}}

  /* Stats */
  .stats-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:22px}
  @media(max-width:900px){.stats-grid{grid-template-columns:repeat(2,1fr)}}
  @media(max-width:480px){.stats-grid{grid-template-columns:1fr 1fr}}
  .stat-card{background:#fff;border-radius:14px;border:1px solid #F0D4D6;padding:18px 20px;border-left-width:4px}
  .stat-num{font-size:32px;font-weight:800;line-height:1}
  .stat-lbl{font-size:11px;color:#94A3B8;font-weight:700;text-transform:uppercase;margin-bottom:6px}
  .stat-sub{font-size:11px;color:#94A3B8;margin-top:6px}

  /* Cards */
  .card{background:#fff;border-radius:14px;border:1px solid #F0D4D6;overflow:hidden;margin-bottom:16px}
  .card-hdr{padding:14px 20px;border-bottom:1px solid #F0D4D6;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}
  .card-hdr-title{font-weight:700;font-size:14px;color:#111827}
  .card-hdr-sub{font-size:12px;color:#94A3B8}

  /* Grid 2 col para dashboard */
  .dash-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
  @media(max-width:768px){.dash-grid{grid-template-columns:1fr}}

  /* Tabla */
  table{width:100%;border-collapse:collapse}
  th{background:#FDE8EA;font-size:11px;font-weight:700;color:#94A3B8;text-transform:uppercase;padding:11px 16px;text-align:left;white-space:nowrap}
  td{font-size:13px;color:#111827;padding:12px 16px;border-top:1px solid #FDF0F0;vertical-align:middle}
  tr:hover td{background:#FDF8F8}

  /* Badges */
  .badge{display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700}
  .badge.active{background:#DCFCE7;color:#16A34A}
  .badge.expired{background:#FEF3C7;color:#D97706}
  .badge.inactive{background:#F3F4F6;color:#6B7280}
  .badge-dot{width:5px;height:5px;border-radius:50%;background:currentColor;flex-shrink:0}
  .dias-badge{display:inline-flex;align-items:center;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;margin-left:6px}
  .dias-ok{background:#DCFCE7;color:#16A34A}
  .dias-warn{background:#FEF3C7;color:#D97706}
  .dias-venc{background:#FEE2E2;color:#DC2626}

  /* Botones */
  .btn{border-radius:10px;padding:8px 16px;font-size:12px;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;gap:6px;border:none;transition:opacity .15s;font-family:inherit}
  .btn:hover{opacity:.88}.btn:disabled{opacity:.45;cursor:not-allowed}
  .btn-primary{background:#BE1522;color:#fff}
  .btn-outline{background:#fff;color:#BE1522;border:1.5px solid #BE1522}
  .btn-success{background:#DCFCE7;color:#16A34A;border:1px solid #BBF7D0}
  .btn-danger{background:#FEE2E2;color:#DC2626;border:1px solid #FECACA}
  .btn-green{background:#DCFCE7;color:#16A34A;border:1.5px solid #BBF7D0}
  .btn-gray{background:#F3F4F6;color:#6B7280;border:1px solid #E5E7EB}
  .btn-sm{padding:5px 12px;font-size:11px;border-radius:8px}
  .tab-btn{padding:6px 14px;border-radius:20px;font-size:12px;font-weight:700;cursor:pointer;border:none;background:#F3F4F6;color:#94A3B8;transition:.15s;font-family:inherit}
  .tab-btn.on{background:#BE1522;color:#fff}
  .action-btn{background:none;border:none;cursor:pointer;padding:5px 7px;border-radius:6px;font-size:16px;transition:.15s}
  .action-btn:hover{background:#FDE8EA}
  .toggle{width:40px;height:22px;background:#E5E7EB;border-radius:11px;position:relative;cursor:pointer;border:none;transition:.2s;flex-shrink:0}
  .toggle.on{background:#BE1522}
  .toggle::after{content:'';position:absolute;width:16px;height:16px;background:#fff;border-radius:50%;top:3px;left:3px;transition:.2s;box-shadow:0 1px 3px rgba(0,0,0,.2)}
  .toggle.on::after{left:21px}

  /* Search */
  .search-row{display:flex;gap:10px;margin-bottom:14px;align-items:center;flex-wrap:wrap}
  .search-box{display:flex;align-items:center;gap:8px;background:#F9FAFB;border:1.5px solid #E5E7EB;border-radius:10px;padding:8px 14px;flex:1;min-width:180px}
  .search-box input{border:none;outline:none;background:transparent;font-size:13px;color:#111827;flex:1;font-family:inherit}

  /* Modal */
  .overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:300;align-items:center;justify-content:center;padding:16px}
  .overlay.show{display:flex}
  .modal{background:#fff;border-radius:20px;width:100%;max-width:560px;padding:28px;max-height:92vh;overflow-y:auto}
  .modal-wide{max-width:720px}
  @media(max-width:540px){.modal{padding:20px 16px;border-radius:14px}}
  .modal-title{font-size:18px;font-weight:700;color:#111827;margin-bottom:16px;display:flex;align-items:center;gap:8px}

  /* Form */
  .form-lbl{font-size:11px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.4px;margin-bottom:5px;margin-top:14px;display:block}
  .form-inp{width:100%;background:#F9FAFB;border:1.5px solid #E5E7EB;border-radius:10px;padding:11px 14px;font-size:14px;outline:none;color:#111827;font-family:inherit;transition:border-color .15s}
  .form-inp:focus{border-color:#BE1522}
  .form-inp:disabled{opacity:.6;background:#F5F5F5;cursor:not-allowed}
  .form-row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  .form-row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}
  @media(max-width:480px){.form-row2,.form-row3{grid-template-columns:1fr}}
  .form-section{font-size:10px;font-weight:700;color:#BE1522;text-transform:uppercase;letter-spacing:.6px;margin:18px 0 4px;padding-bottom:6px;border-bottom:1px solid #F0D4D6}

  /* Info boxes */
  .auto-box{background:#DCFCE7;border:1px solid #BBF7D0;border-radius:10px;padding:12px;font-size:12px;color:#15803D;line-height:1.6;margin-bottom:4px}
  .warn-box{background:#FEF3C7;border:1px solid #FDE68A;border-radius:10px;padding:12px;font-size:12px;color:#92400E;line-height:1.6;margin-bottom:4px}
  .err-box{background:#FEE2E2;border:1px solid #FECACA;border-radius:10px;padding:12px;font-size:12px;color:#DC2626;line-height:1.6;margin-bottom:4px}
  .info-box{background:#EFF6FF;border:1px solid #BFDBFE;border-radius:10px;padding:12px;font-size:12px;color:#1D4ED8;line-height:1.6;margin-bottom:4px}

  /* Detalle egresado */
  .detail-tabs{display:flex;gap:6px;margin-bottom:14px;flex-wrap:wrap}
  .detail-tab{padding:5px 14px;border-radius:20px;font-size:12px;font-weight:600;cursor:pointer;border:none;background:#F3F4F6;color:#6B7280;font-family:inherit;transition:.15s}
  .detail-tab.on{background:#BE1522;color:#fff}
  .detail-grid{background:#F9F1F1;border-radius:12px;padding:14px;margin:10px 0}
  .detail-row{display:flex;justify-content:space-between;align-items:flex-start;padding:7px 0;border-bottom:1px solid #F0D4D6;gap:12px}
  .detail-row:last-child{border-bottom:none}
  .detail-key{font-size:11px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.3px;flex-shrink:0}
  .detail-val{font-size:13px;font-weight:500;color:#111827;text-align:right;word-break:break-all}
  .actions-row{display:flex;gap:8px;flex-wrap:wrap;margin:14px 0}

  /* Avatar */
  .avatar{border-radius:50%;object-fit:cover;flex-shrink:0}
  .avatar-initials{border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:700;flex-shrink:0;background:#FDE8EA;color:#BE1522}

  /* Alertas */
  .alerta-box{background:#FEF3C7;border:1px solid #FDE68A;border-radius:12px;padding:14px 18px;margin-bottom:16px;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
  .val-hoy-box{background:#FDE8EA;border:1px solid #F0D4D6;border-radius:12px;padding:14px 18px;margin-bottom:16px;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
  .export-bar{display:flex;gap:8px;align-items:center;background:#fff;border:1px solid #F0D4D6;border-radius:12px;padding:10px 14px;margin-bottom:14px;flex-wrap:wrap}
  .export-bar span{font-size:12px;color:#94A3B8;font-weight:600}

  /* Import */
  .import-table{font-size:12px;width:100%;border-collapse:collapse;margin-top:10px}
  .import-table th{background:#F9FAFB;font-weight:700;padding:6px 10px;text-align:left;border-bottom:1px solid #E5E7EB}
  .import-table td{padding:5px 10px;border-bottom:1px solid #F5F5F5}
  .drop-zone{border:2px dashed #E5E7EB;border-radius:12px;padding:28px;text-align:center;cursor:pointer;transition:.2s;background:#FAFAFA;position:relative}
  .drop-zone:hover,.drop-zone.drag{border-color:#BE1522;background:#FDE8EA10}
  .drop-zone input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
  .import-scroll{max-height:200px;overflow-y:auto;border:1px solid #E5E7EB;border-radius:10px}

  /* Barra progreso membresía */
  .prog-bar{height:6px;background:#F0D4D6;border-radius:3px;overflow:hidden;margin-top:4px}
  .prog-fill{height:100%;border-radius:3px;transition:width .3s}

  /* Timeline validaciones */
  .val-item{display:flex;gap:12px;padding:10px 0;border-bottom:1px solid #FDF0F0}
  .val-item:last-child{border-bottom:none}
  .val-dot{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0}
`

export default function DashboardPage() {
  const [sesion,       setSesion]       = useState<any>(null)
  const [emailLogin,   setEmailLogin]   = useState('')
  const [claveLogin,   setClaveLogin]   = useState('')
  const [cargandoAuth, setCargandoAuth] = useState(true)
  const [tab,          setTab]          = useState<Tab>('dashboard')
  const [egresados,    setEgresados]    = useState<any[]>([])
  const [secretarios,  setSecretarios]  = useState<any[]>([])
  const [validaciones, setValidaciones] = useState<any[]>([])
  const [valHoy,       setValHoy]       = useState<any[]>([])
  const [stats, setStats] = useState({ activos:0, vencidos:0, inactivos:0, total:0, proxVencer:0 })
  const [busqueda,      setBusqueda]     = useState('')
  const [busquedaSec,   setBusquedaSec]  = useState('')
  const [filtroEstado,  setFiltroEstado] = useState<FiltroEstado>('todos')
  const [filtroFecha,   setFiltroFecha]  = useState<FiltroFecha>('hoy')
  const [modalNuevoEg,  setModalNuevoEg]  = useState(false)
  const [modalNuevoSec, setModalNuevoSec] = useState(false)
  const [modalDetalle,  setModalDetalle]  = useState(false)
  const [modalImport,   setModalImport]   = useState(false)
  const [egSel,         setEgSel]         = useState<any>(null)
  const [detTab,        setDetTab]        = useState<'info'|'membresia'|'estudios'|'laboral'>('info')
  const [sidebarOpen,   setSidebarOpen]   = useState(false)
  const [verificando,   setVerificando]   = useState(false)

  // Import
  const [importRows,   setImportRows]   = useState<any[]>([])
  const [importStatus, setImportStatus] = useState<('idle'|'ok'|'err'|'dup')[]>([])
  const [importMsg,    setImportMsg]    = useState<string[]>([])
  const [importando,   setImportando]   = useState(false)
  const [importDone,   setImportDone]   = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)

  // Form egresado completo
  const [formEg, setFormEg] = useState({
    cedula:'', nombre:'', email:'', telefono:'',
    inicioMembresia: fechaHoy(),
    vencimiento: sumar365(fechaHoy()),
    estado:'activo',
    ciudad_nacimiento:'', direccion:'', fecha_nacimiento:'',
    titulo_pregrado:'', institucion_pregrado:'', fecha_grado_pregrado:'',
    titulo_posgrado:'', institucion_posgrado:'', fecha_grado_posgrado:'',
    empresa:'', cargo:'', hobbies:'',
  })
  const [formSec,      setFormSec]      = useState({ cedula:'', nombre:'', email:'' })
  const [saving,       setSaving]       = useState(false)
  const [editandoEmail,setEditandoEmail]= useState(false)
  const [nuevoEmail,   setNuevoEmail]   = useState('')
  const [savingEmail,  setSavingEmail]  = useState(false)

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => { setSesion(session); setCargandoAuth(false) })
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, s) => setSesion(s))
    return () => subscription.unsubscribe()
  }, [])

  useEffect(() => { if (sesion) verificarVencidosYCargar() }, [sesion])
  useEffect(() => { if (sesion && tab === 'validaciones') cargarValidaciones() }, [tab, filtroFecha])

  async function verificarVencidosYCargar() {
    setVerificando(true)
    const hoy = fechaHoy()
    const { data: aVencer } = await supabase.from('egresados').select('id').eq('estado', 'activo').lt('fecha_vencimiento', hoy)
    if (aVencer?.length) await supabase.from('egresados').update({ estado:'vencido' }).in('id', aVencer.map((e:any)=>e.id))
    setVerificando(false)
    await cargarTodo()
  }

  async function cargarTodo() {
    const [{ data: eg }, { data: sec }, { data: val }] = await Promise.all([
      supabase.from('egresados').select('*').order('nombre_completo'),
      supabase.from('secretarios').select('*').order('nombre_completo'),
      supabase.from('validaciones_detalle').select('*').eq('fecha', fechaHoy()).order('hora_validacion', { ascending:false }).limit(5),
    ])
    const lista = eg || []
    setEgresados(lista); setSecretarios(sec||[]); setValHoy(val||[])
    setStats({
      activos:    lista.filter((e:any)=>e.estado==='activo').length,
      vencidos:   lista.filter((e:any)=>e.estado==='vencido').length,
      inactivos:  lista.filter((e:any)=>e.estado==='inactivo').length,
      total:      lista.length,
      proxVencer: lista.filter((e:any)=>{
        if (e.estado!=='activo'||!e.fecha_vencimiento) return false
        const d=diasRestantes(e.fecha_vencimiento); return d>=0&&d<=30
      }).length,
    })
  }

  async function cargarValidaciones() {
    let q = supabase.from('validaciones_detalle').select('*').order('hora_validacion',{ascending:false}).limit(200)
    const hoy = new Date()
    if (filtroFecha==='hoy') q=q.eq('fecha',fechaHoy())
    else if (filtroFecha==='semana') { const d=new Date(hoy);d.setDate(hoy.getDate()-7);q=q.gte('hora_validacion',d.toISOString()) }
    else if (filtroFecha==='mes')    { const d=new Date(hoy);d.setDate(hoy.getDate()-30);q=q.gte('hora_validacion',d.toISOString()) }
    const { data } = await q
    setValidaciones(data||[])
  }

  function leerArchivoExcel(file: File) {
    const reader = new FileReader()
    reader.onload = (e) => {
      const data = new Uint8Array(e.target!.result as ArrayBuffer)
      const wb   = XLSX.read(data,{type:'array'})
      const ws   = wb.Sheets[wb.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json<any>(ws,{defval:''})
      if (!rows.length) { alert('El archivo está vacío.'); return }
      const normalized = rows.map((row:any) => {
        const out: Record<string,string> = {}
        for (const key of Object.keys(row)) {
          const norm=normalizarHeader(String(key)); const campo=HEADER_MAP[norm]
          if (campo) out[campo]=String(row[key]).trim()
        }
        return out
      }).filter(r=>r.cedula&&r.nombre&&r.email)
      if (!normalized.length) { alert('Sin filas válidas. Columnas requeridas: Cédula, Nombre, Email.'); return }
      setImportRows(normalized); setImportStatus(normalized.map(()=>'idle')); setImportMsg(normalized.map(()=>'')); setImportDone(false); setModalImport(true)
    }
    reader.readAsArrayBuffer(file)
  }

  async function ejecutarImportacion() {
    setImportando(true)
    const status=[...importStatus]; const msgs=[...importMsg]
    const inputInicio=document.getElementById('importFechaInicio') as HTMLInputElement
    const inicio=inputInicio?.value||fechaHoy()
    const venc=sumar365(inicio)
    for (let i=0;i<importRows.length;i++) {
      const row=importRows[i]; const cedula=row.cedula.replace(/[.\s]/g,'')
      const {data:existing}=await supabase.from('egresados').select('id').eq('cedula',cedula).maybeSingle()
      if (existing) { status[i]='dup';msgs[i]='Ya existe';setImportStatus([...status]);setImportMsg([...msgs]);continue }
      const {error}=await supabase.from('egresados').insert({
        cedula, nombre_completo:row.nombre, email:row.email.toLowerCase(),
        telefono:row.telefono||null, fecha_expedicion:inicio, fecha_vencimiento:venc,
        estado:'activo', requiere_cambio_clave:true,
      })
      if (error) { status[i]='err';msgs[i]=error.message } else { status[i]='ok';msgs[i]=`Vence: ${venc}` }
      setImportStatus([...status]);setImportMsg([...msgs])
    }
    setImportando(false);setImportDone(true);cargarTodo()
  }

  const egresadosFiltrados = egresados.filter(eg => {
    const matchEstado=filtroEstado==='todos'||eg.estado===filtroEstado
    const q=busqueda.toLowerCase()
    const matchBusqueda=!q||eg.nombre_completo?.toLowerCase().includes(q)||eg.cedula?.includes(q)||(eg.email??'').toLowerCase().includes(q)||(eg.empresa??'').toLowerCase().includes(q)
    return matchEstado&&matchBusqueda
  })

  const secretariosFiltrados = secretarios.filter(sec => {
    const q=busquedaSec.toLowerCase()
    return !q||sec.nombre_completo?.toLowerCase().includes(q)||sec.cedula?.includes(q)||(sec.email??'').toLowerCase().includes(q)
  })

  function onChangeInicio(val:string) { setFormEg(p=>({...p,inicioMembresia:val,vencimiento:sumar365(val)})) }
  function updFormEg(k:string,v:string) { setFormEg(p=>({...p,[k]:v})) }

  async function manejarLogin(e:React.FormEvent) {
    e.preventDefault()
    const {error}=await supabase.auth.signInWithPassword({email:emailLogin.trim(),password:claveLogin})
    if (error) alert('Error: '+error.message)
  }
  async function manejarLogout() { await supabase.auth.signOut();setSesion(null) }

  async function crearEgresado() {
    if (!formEg.cedula||!formEg.nombre||!formEg.email) return alert('Cédula, nombre y email son obligatorios.')
    setSaving(true)
    const {error}=await supabase.from('egresados').insert({
      cedula:formEg.cedula.replace(/[.\s]/g,''), nombre_completo:formEg.nombre,
      email:formEg.email.toLowerCase(), telefono:formEg.telefono||null,
      fecha_expedicion:formEg.inicioMembresia, fecha_vencimiento:formEg.vencimiento,
      estado:formEg.estado, requiere_cambio_clave:true,
      ciudad_nacimiento:formEg.ciudad_nacimiento||null, direccion:formEg.direccion||null,
      fecha_nacimiento:formEg.fecha_nacimiento||null,
      titulo_pregrado:formEg.titulo_pregrado||null, institucion_pregrado:formEg.institucion_pregrado||null,
      fecha_grado_pregrado:formEg.fecha_grado_pregrado||null,
      titulo_posgrado:formEg.titulo_posgrado||null, institucion_posgrado:formEg.institucion_posgrado||null,
      fecha_grado_posgrado:formEg.fecha_grado_posgrado||null,
      empresa:formEg.empresa||null, cargo:formEg.cargo||null, hobbies:formEg.hobbies||null,
    })
    setSaving(false)
    if (error) return alert('Error: '+error.message)
    alert(`✓ Egresado creado\nMembresía: ${formEg.inicioMembresia} → ${formEg.vencimiento}\nContraseña inicial: ${formEg.cedula}`)
    setFormEg({cedula:'',nombre:'',email:'',telefono:'',inicioMembresia:fechaHoy(),vencimiento:sumar365(fechaHoy()),estado:'activo',ciudad_nacimiento:'',direccion:'',fecha_nacimiento:'',titulo_pregrado:'',institucion_pregrado:'',fecha_grado_pregrado:'',titulo_posgrado:'',institucion_posgrado:'',fecha_grado_posgrado:'',empresa:'',cargo:'',hobbies:''})
    setModalNuevoEg(false);cargarTodo()
  }

  async function crearSecretario() {
    if (!formSec.cedula||!formSec.nombre||!formSec.email) return alert('Todos los campos son obligatorios.')
    setSaving(true)
    const {error}=await supabase.from('secretarios').insert({cedula:formSec.cedula.replace(/[.\s]/g,''),nombre_completo:formSec.nombre,email:formSec.email.toLowerCase(),activo:true,requiere_cambio_clave:true})
    setSaving(false)
    if (error) return alert('Error: '+error.message)
    alert(`✓ Secretario creado\nContraseña inicial: ${formSec.cedula}`)
    setFormSec({cedula:'',nombre:'',email:''});setModalNuevoSec(false);cargarTodo()
  }

  async function renovarCarnet(eg:any, fechaInicio:string) {
    const venc=sumar365(fechaInicio)
    if (!confirm(`¿Renovar membresía de ${eg.nombre_completo}?\n\nInicio: ${fechaInicio}\nVence: ${venc}`)) return
    const {error}=await supabase.from('egresados').update({estado:'activo',fecha_expedicion:fechaInicio,fecha_vencimiento:venc}).eq('id',eg.id)
    if (error) return alert('Error: '+error.message)
    alert(`✓ Membresía renovada hasta ${venc}`)
    setModalDetalle(false);cargarTodo()
  }

  async function cambiarEstado(eg:any, estado:string) {
    if (!confirm(`¿${estado==='activo'?'Activar':'Cambiar estado de'} ${eg.nombre_completo} a "${estado}"?`)) return
    await supabase.from('egresados').update({estado}).eq('id',eg.id)
    setModalDetalle(false);cargarTodo()
  }

  async function eliminarEgresado(eg:any) {
    if (!confirm(`⚠️ ¿Eliminar permanentemente a ${eg.nombre_completo}?`)) return
    if (!confirm(`Confirma de nuevo: ¿eliminar a ${eg.nombre_completo}? Esta acción es IRREVERSIBLE.`)) return
    const {error}=await supabase.from('egresados').delete().eq('id',eg.id)
    if (error) return alert('Error: '+error.message)
    alert('✓ Egresado eliminado.')
    setModalDetalle(false);cargarTodo()
  }

  async function actualizarEmail(eg:any,email:string) {
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { alert('Email inválido.'); return }
    if (!confirm(`¿Cambiar email de ${eg.nombre_completo} a "${email}"?`)) return
    setSavingEmail(true)
    const {error}=await supabase.from('egresados').update({email:email.toLowerCase()}).eq('id',eg.id)
    if (error) { alert('Error: '+error.message);setSavingEmail(false);return }
    setSavingEmail(false);alert('✓ Email actualizado.');setEditandoEmail(false);setNuevoEmail('');cargarTodo()
  }

  async function toggleSecretario(sec:any) {
    await supabase.from('secretarios').update({activo:!sec.activo}).eq('id',sec.id);cargarTodo()
  }

  function abrirDetalle(eg:any) { setEgSel(eg);setModalDetalle(true);setDetTab('info');setEditandoEmail(false);setNuevoEmail(eg.email??'') }
  function initials(n:string) { return n.split(' ').slice(0,2).map((x:string)=>x[0]).join('').toUpperCase() }
  function nombreArchivo(p:string) { return `ASEDUIS-${p}-${fechaHoy()}` }

  function badgeDias(fechaVenc:string,estado:string) {
    if (estado!=='activo'||!fechaVenc) return null
    const dias=diasRestantes(fechaVenc)
    if (dias<0) return <span className="dias-badge dias-venc">Vencido</span>
    if (dias<=30) return <span className="dias-badge dias-warn">{dias}d</span>
    return <span className="dias-badge dias-ok">{dias}d</span>
  }

  function porcentajeMembresia(inicio:string|null, venc:string|null): number {
    if (!inicio||!venc) return 0
    const total=diasRestantes(inicio)-diasRestantes(venc)
    const transcurrido=total-diasRestantes(venc)
    return Math.max(0,Math.min(100,Math.round((transcurrido/total)*100)))
  }

  const importOk=importStatus.filter(s=>s==='ok').length
  const importErr=importStatus.filter(s=>s==='err').length
  const importDup=importStatus.filter(s=>s==='dup').length

  function navTo(t:Tab) { setTab(t);setSidebarOpen(false) }

  if (cargandoAuth) return (
    <div style={{display:'flex',height:'100vh',alignItems:'center',justifyContent:'center',fontFamily:'sans-serif',background:'#F9F1F1',color:'#BE1522',fontWeight:600,fontSize:16}}>
      Cargando ASEDUIS...
    </div>
  )

  if (!sesion) return (
    <>
      <style>{`
        *{box-sizing:border-box}
        .login-page{
            display:flex;
            min-height:100vh;
            width:100vw;
            align-items:center;
            justify-content:center;
            background: linear-gradient(135deg, #BE1522 0%, #7B0D15 50%, #1a1a2e 100%);
            font-family:"Segoe UI",system-ui,sans-serif;
            padding:16px;
            margin:0;
            box-sizing:border-box;
          }
          .login-card{
            background:#fff;
            padding:40px;
            border-radius:20px;
            box-shadow:0 20px 60px rgba(0,0,0,.35);
            width:100%;
            max-width:400px
          }
        .login-logo{width:110px;height:110px;background:#BE1522;border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 14px;font-size:28px;font-weight:900;color:#fff;overflow:hidden}
        .login-logo img{width:100%;height:100%;object-fit:contain;padding:10px}
        .login-title{color:#BE1522;font-weight:800;font-size:22px;letter-spacing:1px;margin:0;text-align:center}
        .login-sub{color:#94A3B8;font-size:13px;margin-top:4px;text-align:center;margin-bottom:28px}
        .login-label{font-size:11px;font-weight:700;color:#94A3B8;text-transform:uppercase;display:block;margin-bottom:6px;letter-spacing:.5px}
        .login-input{width:100%;padding:12px 14px;border-radius:12px;border:1.5px solid #E5E7EB;margin-bottom:16px;outline:none;font-size:14px;font-family:inherit;box-sizing:border-box;color:#111827;transition:border-color .15s}
        .login-input:focus{border-color:#BE1522}
        .login-input-last{margin-bottom:24px}
        .login-btn{width:100%;background:#BE1522;color:#fff;padding:14px;border-radius:12px;border:none;font-weight:700;cursor:pointer;font-size:15px;font-family:inherit;transition:opacity .15s}
        .login-btn:hover{opacity:.88}
      `}</style>
      <div className="login-page">
        <form onSubmit={manejarLogin} className="login-card">
          <div className="login-logo">
            <img src="/logo.png" alt="ASEDUIS" onError={(e)=>{const t=e.target as HTMLImageElement;t.style.display='none';t.parentElement!.textContent='A'}} />
          </div>
          <h2 className="login-title">ASEDUIS</h2>
          <p className="login-sub">Panel Administrativo</p>
          <label className="login-label">Correo electrónico</label>
          <input type="email" required className="login-input" placeholder="admin@aseduis.com" value={emailLogin} onChange={e=>setEmailLogin(e.target.value)} />
          <label className="login-label">Contraseña</label>
          <input type="password" required className="login-input login-input-last" placeholder="••••••••" value={claveLogin} onChange={e=>setClaveLogin(e.target.value)} />
          <button type="submit" className="login-btn">Ingresar al sistema</button>
        </form>
      </div>
    </>
  )

  return (
    <div className="layout">
      <style>{CSS}</style>
      <div className={`sidebar-backdrop ${sidebarOpen?'show':''}`} onClick={()=>setSidebarOpen(false)} />

      {/* ── SIDEBAR ── */}
      <aside className={`sidebar ${sidebarOpen?'open':''}`}>
        <div style={{display:'flex',alignItems:'center',gap:12,padding:'22px 20px',borderBottom:'1px solid rgba(255,255,255,.15)'}}>
          <div style={{width:40,height:40,background:'#fff',borderRadius:'50%',display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0,overflow:'hidden'}}>
            <img src="/logo_black.png" alt="ASEDUIS" style={{width:'80%',height:'80%',objectFit:'contain'}} />
          </div>
          <div>
            <div style={{color:'#fff',fontWeight:800,fontSize:15,letterSpacing:1}}>ASEDUIS</div>
            <div style={{color:'rgba(255,255,255,.55)',fontSize:11}}>Panel administrativo</div>
          </div>
        </div>
        <div style={{padding:'14px 12px 8px'}}>
          <div style={{fontSize:10,fontWeight:700,color:'rgba(255,255,255,.4)',textTransform:'uppercase',letterSpacing:.6,padding:'0 8px',marginBottom:8}}>Principal</div>
          {([
            {k:'dashboard',    e:'📊', l:'Dashboard',    badge:null},
            {k:'egresados',    e:'👥', l:'Egresados',    badge:stats.total},
            {k:'secretarios',  e:'🔑', l:'Secretarios',  badge:secretarios.length},
            {k:'validaciones', e:'📷', l:'Validaciones', badge:valHoy.length||null},
          ] as {k:Tab;e:string;l:string;badge:number|null}[]).map(n=>(
            <div key={n.k} className={`nav-item ${tab===n.k?'active':''}`} onClick={()=>navTo(n.k)}>
              <span style={{fontSize:18}}>{n.e}</span>
              <span className="nav-lbl">{n.l}</span>
              {n.badge!==null&&<span style={{marginLeft:'auto',background:'rgba(255,255,255,.2)',color:'#fff',fontSize:10,fontWeight:700,padding:'1px 7px',borderRadius:10}}>{n.badge}</span>}
            </div>
          ))}
        </div>
        {/* Mini stats en sidebar */}
        <div style={{margin:'12px',background:'rgba(255,255,255,.1)',borderRadius:12,padding:'12px 14px'}}>
          <div style={{fontSize:10,color:'rgba(255,255,255,.5)',fontWeight:700,textTransform:'uppercase',letterSpacing:.5,marginBottom:8}}>Estado membresías</div>
          {[{l:'Activos',c:'#4ADE80',v:stats.activos},{l:'Vencidos',c:'#FCD34D',v:stats.vencidos},{l:'Inactivos',c:'#94A3B8',v:stats.inactivos}].map(s=>(
            <div key={s.l} style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6}}>
              <div style={{display:'flex',alignItems:'center',gap:6}}>
                <div style={{width:6,height:6,borderRadius:'50%',background:s.c}}/>
                <span style={{fontSize:11,color:'rgba(255,255,255,.7)'}}>{s.l}</span>
              </div>
              <span style={{fontSize:12,fontWeight:700,color:'#fff'}}>{s.v}</span>
            </div>
          ))}
        </div>
        <div style={{marginTop:'auto',padding:16,borderTop:'1px solid rgba(255,255,255,.15)'}}>
          <div style={{color:'#fff',fontSize:13,fontWeight:600,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',marginBottom:2}}>{sesion.user.email}</div>
          <div style={{color:'rgba(255,255,255,.5)',fontSize:11,marginBottom:10}}>Administrador</div>
          <button onClick={manejarLogout} style={{display:'flex',alignItems:'center',gap:8,padding:'8px 12px',borderRadius:8,background:'rgba(255,255,255,.1)',border:'none',color:'#fff',cursor:'pointer',width:'100%',fontSize:13,fontFamily:'inherit'}}>
            <span>⬅</span><span>Cerrar sesión</span>
          </button>
        </div>
      </aside>

      {/* ── MAIN ── */}
      <div className="main">
        <div className="topbar">
          <div style={{display:'flex',alignItems:'center',gap:10}}>
            <button className="menu-btn" onClick={()=>setSidebarOpen(o=>!o)}>☰</button>
            <div>
              <div style={{fontSize:18,fontWeight:700,color:'#BE1522',lineHeight:1.2}}>
                {{dashboard:'Dashboard',egresados:'Egresados',secretarios:'Secretarios',validaciones:'Validaciones'}[tab]}
              </div>
              <div style={{fontSize:11,color:'#94A3B8',marginTop:2}}>
                {verificando?'⏳ Verificando membresías...':`Actualizado · ${fechaHoy()}`}
              </div>
            </div>
          </div>
          <div className="topbar-actions">
            {tab==='egresados'&&<>
              <button className="btn btn-green btn-sm" onClick={()=>exportarExcel(egresadosFiltrados,nombreArchivo('egresados'))}>📊 Excel</button>
              <button className="btn btn-outline btn-sm" onClick={()=>exportarPDF(egresadosFiltrados,nombreArchivo('egresados'))}>📄 PDF</button>
              <button className="btn btn-gray btn-sm" onClick={()=>{setModalImport(true);setImportRows([]);setImportStatus([]);setImportMsg([]);setImportDone(false)}}>📥 Importar</button>
              <button className="btn btn-primary btn-sm" onClick={()=>setModalNuevoEg(true)}>+ Nuevo egresado</button>
            </>}
            {tab==='secretarios'&&<button className="btn btn-primary btn-sm" onClick={()=>setModalNuevoSec(true)}>+ Nuevo secretario</button>}
            {tab==='dashboard'&&<>
              <button className="btn btn-gray btn-sm" onClick={()=>{setModalImport(true);setImportRows([]);setImportStatus([]);setImportMsg([]);setImportDone(false)}}>📥 Importar Excel</button>
              <button className="btn btn-primary btn-sm" onClick={()=>setModalNuevoEg(true)}>+ Nuevo egresado</button>
            </>}
            {tab==='validaciones'&&<button className="btn btn-green btn-sm" onClick={()=>exportarExcel(validaciones,nombreArchivo('validaciones'))}>📊 Excel</button>}
            <button className="btn btn-gray btn-sm" onClick={verificarVencidosYCargar} title="Actualizar datos">🔄</button>
          </div>
        </div>

        <div className="content">

          {/* ══════════ DASHBOARD ══════════ */}
          {tab==='dashboard'&&<>
            {/* Stats */}
            <div className="stats-grid">
              {[
                {lbl:'Activos',num:stats.activos,sub:'membresías vigentes',color:'#16A34A',border:'#16A34A',icon:'✅'},
                {lbl:'Vencidos',num:stats.vencidos,sub:'membresía expirada',color:'#D97706',border:'#D97706',icon:'⚠️'},
                {lbl:'Inactivos',num:stats.inactivos,sub:'suspendidos',color:'#6B7280',border:'#94A3B8',icon:'⏸'},
                {lbl:'Total egresados',num:stats.total,sub:`${secretarios.length} secretarios activos`,color:'#BE1522',border:'#BE1522',icon:'👥'},
              ].map(s=>(
                <div key={s.lbl} className="stat-card" style={{borderLeftColor:s.border}}>
                  <div className="stat-lbl">{s.icon} {s.lbl}</div>
                  <div className="stat-num" style={{color:s.color}}>{s.num}</div>
                  <div className="stat-sub">{s.sub}</div>
                </div>
              ))}
            </div>

            {/* Alertas */}
            {stats.proxVencer>0&&(
              <div className="alerta-box">
                <span style={{fontSize:20}}>⏰</span>
                <div style={{flex:1}}>
                  <div style={{fontWeight:700,fontSize:13,color:'#92400E'}}>{stats.proxVencer} membresía{stats.proxVencer!==1?'s':''} vence{stats.proxVencer===1?'':'n'} en los próximos 30 días</div>
                  <div style={{fontSize:12,color:'#B45309',marginTop:2}}>Contacta a estos egresados para renovación anticipada.</div>
                </div>
                <button className="btn btn-sm" style={{background:'#FEF3C7',color:'#92400E',border:'1px solid #FDE68A'}} onClick={()=>{setTab('egresados');setFiltroEstado('activo')}}>Ver →</button>
              </div>
            )}
            {valHoy.length>0&&(
              <div className="val-hoy-box">
                <div style={{fontSize:13,fontWeight:700,color:'#BE1522',flexShrink:0}}>📷 Hoy: {valHoy.length} validacion{valHoy.length!==1?'es':''}</div>
                {valHoy.slice(0,3).map((v:any)=>(
                  <div key={v.id} style={{display:'flex',alignItems:'center',gap:6,background:'#fff',borderRadius:8,padding:'5px 10px',fontSize:12}}>
                    <span>{RESULTADO_EMOJI[v.resultado]??'?'}</span>
                    <span style={{fontWeight:600}}>{v.egresado_nombre??'Desconocido'}</span>
                    <span style={{color:'#94A3B8'}}>{v.hora?.slice(0,5)}</span>
                  </div>
                ))}
                <button className="btn btn-outline btn-sm" style={{marginLeft:'auto'}} onClick={()=>setTab('validaciones')}>Ver todas →</button>
              </div>
            )}

            {/* Gráficas + tabla reciente */}
            <div className="dash-grid">
              {/* Gráfica distribución */}
              <div className="card">
                <div className="card-hdr">
                  <span className="card-hdr-title">📊 Distribución de membresías</span>
                </div>
                <div style={{padding:'20px',display:'flex',gap:24,alignItems:'center',flexWrap:'wrap'}}>
                  <DonutChart activos={stats.activos} vencidos={stats.vencidos} inactivos={stats.inactivos} total={stats.total} />
                  <div style={{flex:1,minWidth:120}}>
                    {[
                      {l:'Activos',c:'#16A34A',v:stats.activos},
                      {l:'Vencidos',c:'#D97706',v:stats.vencidos},
                      {l:'Inactivos',c:'#6B7280',v:stats.inactivos},
                    ].map(item=>(
                      <div key={item.l} style={{marginBottom:10}}>
                        <div style={{display:'flex',justifyContent:'space-between',marginBottom:3}}>
                          <div style={{display:'flex',alignItems:'center',gap:6}}>
                            <div style={{width:8,height:8,borderRadius:'50%',background:item.c}}/>
                            <span style={{fontSize:12,color:'#6B7280'}}>{item.l}</span>
                          </div>
                          <span style={{fontSize:12,fontWeight:700,color:item.c}}>{stats.total?Math.round((item.v/stats.total)*100):0}%</span>
                        </div>
                        <div className="prog-bar">
                          <div className="prog-fill" style={{width:`${stats.total?Math.round((item.v/stats.total)*100):0}%`,background:item.c}}/>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Gráfica barras */}
              <div className="card">
                <div className="card-hdr">
                  <span className="card-hdr-title">📈 Resumen visual</span>
                </div>
                <div style={{padding:'20px',display:'flex',flexDirection:'column',alignItems:'center',gap:8}}>
                  <BarChart data={[
                    {label:'Activos',  value:stats.activos,  color:'#16A34A'},
                    {label:'Vencidos', value:stats.vencidos, color:'#D97706'},
                    {label:'Inactivos',value:stats.inactivos,color:'#94A3B8'},
                    {label:'Total',    value:stats.total,    color:'#BE1522'},
                  ]}/>
                  <div style={{fontSize:11,color:'#94A3B8',marginTop:4}}>Total de egresados registrados: <strong style={{color:'#BE1522'}}>{stats.total}</strong></div>
                </div>
              </div>
            </div>

            {/* Exportar */}
            <div className="export-bar">
              <span>📤 Exportar:</span>
              <button className="btn btn-green btn-sm" onClick={()=>exportarExcel(egresados,nombreArchivo('todos'))}>📊 Todos</button>
              <button className="btn btn-green btn-sm" onClick={()=>exportarExcel(egresados.filter(e=>e.estado==='activo'),nombreArchivo('activos'))}>✅ Activos</button>
              <button className="btn btn-green btn-sm" onClick={()=>exportarExcel(egresados.filter(e=>e.estado==='vencido'),nombreArchivo('vencidos'))}>⚠️ Vencidos</button>
              <button className="btn btn-outline btn-sm" onClick={()=>exportarPDF(egresados,nombreArchivo('todos'))}>📄 PDF completo</button>
            </div>

            {/* Tabla reciente */}
            <div className="card">
              <div className="card-hdr">
                <span className="card-hdr-title">👥 Egresados recientes</span>
                <button className="btn btn-outline btn-sm" onClick={()=>setTab('egresados')}>Ver todos ({stats.total}) →</button>
              </div>
              <div style={{overflowX:'auto'}}>
                <table>
                  <thead><tr><th>Nombre</th><th>Cédula</th><th>Estado</th><th>Empresa / Cargo</th><th>Vencimiento</th><th></th></tr></thead>
                  <tbody>
                    {egresados.slice(0,8).map(eg=>(
                      <tr key={eg.id}>
                        <td>
                          <div style={{display:'flex',alignItems:'center',gap:10}}>
                            <div style={{width:34,height:34,borderRadius:'50%',overflow:'hidden',background:'#FDE8EA',display:'flex',alignItems:'center',justifyContent:'center',fontWeight:700,fontSize:12,color:'#BE1522',flexShrink:0}}>
                              {eg.foto_perfil?<img src={eg.foto_perfil} style={{width:'100%',height:'100%',objectFit:'cover'}} alt=""/>:initials(eg.nombre_completo)}
                            </div>
                            <div>
                              <div style={{fontWeight:600,fontSize:13}}>{eg.nombre_completo}</div>
                              <div style={{fontSize:11,color:'#94A3B8'}}>{eg.email}</div>
                            </div>
                          </div>
                        </td>
                        <td style={{fontSize:12,color:'#6B7280'}}>{eg.cedula}</td>
                        <td><span className={ESTADO_BADGE[eg.estado]??'badge inactive'}><span className="badge-dot"></span>{eg.estado}</span></td>
                        <td style={{fontSize:12,color:'#6B7280'}}>{eg.empresa?`${eg.empresa}${eg.cargo?` · ${eg.cargo}`:'`'}`:'—'}</td>
                        <td>
                          <span style={{color:eg.estado==='vencido'?'#D97706':eg.estado==='inactivo'?'#94A3B8':'#16A34A',fontWeight:600,fontSize:12}}>{eg.fecha_vencimiento||'—'}</span>
                          {eg.fecha_vencimiento&&badgeDias(eg.fecha_vencimiento,eg.estado)}
                        </td>
                        <td><button className="action-btn" onClick={()=>abrirDetalle(eg)}>👁</button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </>}

          {/* ══════════ EGRESADOS ══════════ */}
          {tab==='egresados'&&<>
            <div className="search-row">
              <div className="search-box">
                <span>🔍</span>
                <input placeholder="Buscar por nombre, cédula, email o empresa..." value={busqueda} onChange={e=>setBusqueda(e.target.value)} />
                {busqueda&&<button onClick={()=>setBusqueda('')} style={{background:'none',border:'none',cursor:'pointer',color:'#94A3B8',fontSize:16}}>✕</button>}
              </div>
              {(['todos','activo','vencido','inactivo'] as FiltroEstado[]).map(f=>(
                <button key={f} className={`tab-btn ${filtroEstado===f?'on':''}`} onClick={()=>setFiltroEstado(f)}>
                  {f==='todos'?`Todos (${egresados.length})`:f==='activo'?`✅ Activos (${stats.activos})`:f==='vencido'?`⚠️ Vencidos (${stats.vencidos})`:`⏸ Inactivos (${stats.inactivos})`}
                </button>
              ))}
            </div>
            <div className="card">
              <div className="card-hdr">
                <span className="card-hdr-title">Base de datos · {egresadosFiltrados.length} egresados</span>
                <div style={{display:'flex',gap:8}}>
                  <button className="btn btn-green btn-sm" onClick={()=>exportarExcel(egresadosFiltrados,nombreArchivo('filtrados'))}>📊 Exportar Excel</button>
                  <button className="btn btn-outline btn-sm" onClick={()=>exportarPDF(egresadosFiltrados,nombreArchivo('filtrados'))}>📄 PDF</button>
                </div>
              </div>
              <div style={{overflowX:'auto'}}>
                <table>
                  <thead>
                    <tr><th>Nombre</th><th>Cédula</th><th>Email</th><th>Empresa / Cargo</th><th>Estado</th><th>Vencimiento</th><th>Días</th><th>Acciones</th></tr>
                  </thead>
                  <tbody>
                    {egresadosFiltrados.length===0
                      ?<tr><td colSpan={8} style={{textAlign:'center',padding:32,color:'#94A3B8'}}>Sin resultados para esta búsqueda</td></tr>
                      :egresadosFiltrados.map(eg=>{
                        const dias=eg.fecha_vencimiento?diasRestantes(eg.fecha_vencimiento):null
                        return(
                        <tr key={eg.id}>
                          <td>
                            <div style={{display:'flex',alignItems:'center',gap:10}}>
                              <div style={{width:34,height:34,borderRadius:'50%',overflow:'hidden',background:'#FDE8EA',display:'flex',alignItems:'center',justifyContent:'center',fontWeight:700,fontSize:12,color:'#BE1522',flexShrink:0}}>
                                {eg.foto_perfil?<img src={eg.foto_perfil} style={{width:'100%',height:'100%',objectFit:'cover'}} alt=""/>:initials(eg.nombre_completo)}
                              </div>
                              <div>
                                <div style={{fontWeight:600,fontSize:13}}>{eg.nombre_completo}</div>
                                {eg.titulo_pregrado&&<div style={{fontSize:10,color:'#94A3B8'}}>{eg.titulo_pregrado}</div>}
                              </div>
                            </div>
                          </td>
                          <td style={{fontSize:12,color:'#6B7280'}}>{eg.cedula}</td>
                          <td style={{fontSize:12,color:'#6B7280'}}>{eg.email}</td>
                          <td style={{fontSize:12,color:'#6B7280'}}>{eg.empresa?<><div style={{fontWeight:600,color:'#374151'}}>{eg.empresa}</div><div>{eg.cargo}</div></>:'—'}</td>
                          <td><span className={ESTADO_BADGE[eg.estado]??'badge inactive'}><span className="badge-dot"></span>{eg.estado}</span></td>
                          <td style={{color:eg.estado==='vencido'?'#D97706':eg.estado==='inactivo'?'#94A3B8':'#16A34A',fontWeight:600,fontSize:12}}>{eg.fecha_vencimiento||'—'}</td>
                          <td>
                            {dias!==null&&eg.estado==='activo'&&<span className={`dias-badge ${dias<=0?'dias-venc':dias<=30?'dias-warn':'dias-ok'}`}>{dias<=0?'Vencido':`${dias}d`}</span>}
                            {eg.estado==='vencido'&&<span className="dias-badge dias-venc">Expirado</span>}
                            {eg.estado==='inactivo'&&<span style={{fontSize:11,color:'#94A3B8'}}>—</span>}
                          </td>
                          <td>
                            <div style={{display:'flex',gap:2}}>
                              <button className="action-btn" title="Ver detalle" onClick={()=>abrirDetalle(eg)}>👁</button>
                              {eg.estado!=='inactivo'&&<button className="action-btn" title="Desactivar" onClick={()=>cambiarEstado(eg,'inactivo')}>⏸</button>}
                              {eg.estado==='inactivo'&&<button className="action-btn" title="Activar" onClick={()=>cambiarEstado(eg,'activo')}>▶</button>}
                            </div>
                          </td>
                        </tr>
                      )})}
                  </tbody>
                </table>
              </div>
            </div>
          </>}

          {/* ══════════ SECRETARIOS ══════════ */}
          {tab==='secretarios'&&<>
            <div className="search-row">
              <div className="search-box"><span>🔍</span><input placeholder="Buscar secretario..." value={busquedaSec} onChange={e=>setBusquedaSec(e.target.value)} /></div>
            </div>
            <div className="card" style={{maxWidth:760}}>
              <div className="card-hdr">
                <span className="card-hdr-title">Secretarios ({secretariosFiltrados.length})</span>
                <span className="card-hdr-sub">{secretarios.filter(s=>s.activo).length} activos · {secretarios.filter(s=>!s.activo).length} inactivos</span>
              </div>
              <div style={{overflowX:'auto'}}>
                <table>
                  <thead><tr><th>Nombre</th><th>Cédula</th><th>Email</th><th>Registro</th><th>Estado</th></tr></thead>
                  <tbody>
                    {secretariosFiltrados.length===0
                      ?<tr><td colSpan={5} style={{textAlign:'center',padding:32,color:'#94A3B8'}}>No hay secretarios</td></tr>
                      :secretariosFiltrados.map(sec=>(
                      <tr key={sec.id}>
                        <td>
                          <div style={{display:'flex',alignItems:'center',gap:10}}>
                            <div style={{width:34,height:34,borderRadius:'50%',background:sec.activo?'#FDE8EA':'#F3F4F6',display:'flex',alignItems:'center',justifyContent:'center',fontWeight:700,fontSize:12,color:sec.activo?'#BE1522':'#9CA3AF',flexShrink:0}}>{initials(sec.nombre_completo)}</div>
                            <strong>{sec.nombre_completo}</strong>
                          </div>
                        </td>
                        <td>{sec.cedula}</td>
                        <td style={{fontSize:12,color:'#6B7280'}}>{sec.email}</td>
                        <td style={{fontSize:12,color:'#94A3B8'}}>{sec.created_at?.split('T')[0]}</td>
                        <td>
                          <div style={{display:'flex',alignItems:'center',gap:8}}>
                            <button className={`toggle ${sec.activo?'on':''}`} onClick={()=>toggleSecretario(sec)} />
                            <span style={{fontSize:12,fontWeight:600,color:sec.activo?'#16A34A':'#94A3B8'}}>{sec.activo?'Activo':'Inactivo'}</span>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </>}

          {/* ══════════ VALIDACIONES ══════════ */}
          {tab==='validaciones'&&<>
            <div style={{display:'flex',gap:8,marginBottom:14,flexWrap:'wrap',alignItems:'center'}}>
              {(['hoy','semana','mes','todo'] as FiltroFecha[]).map(f=>(
                <button key={f} className={`tab-btn ${filtroFecha===f?'on':''}`} onClick={()=>setFiltroFecha(f)}>
                  {f==='hoy'?'Hoy':f==='semana'?'Últimos 7 días':f==='mes'?'Últimos 30 días':'Todo'}
                </button>
              ))}
              <span style={{marginLeft:'auto',fontSize:12,color:'#94A3B8'}}>{validaciones.length} registros</span>
            </div>
            {/* Stats validaciones */}
            <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:12,marginBottom:16}}>
              {[
                {l:'Válidos',    v:validaciones.filter(v=>v.resultado==='activo').length,  c:'#16A34A',e:'✅'},
                {l:'Vencidos',   v:validaciones.filter(v=>v.resultado==='vencido').length,  c:'#D97706',e:'⚠️'},
                {l:'Inactivos',  v:validaciones.filter(v=>v.resultado==='inactivo').length, c:'#DC2626',e:'❌'},
                {l:'No encontrados',v:validaciones.filter(v=>v.resultado==='no_encontrado').length,c:'#94A3B8',e:'❓'},
              ].map(s=>(
                <div key={s.l} style={{background:'#fff',borderRadius:12,border:'1px solid #F0D4D6',padding:'12px 14px',borderLeftWidth:3,borderLeftColor:s.c}}>
                  <div style={{fontSize:10,color:'#94A3B8',fontWeight:700,textTransform:'uppercase',marginBottom:4}}>{s.e} {s.l}</div>
                  <div style={{fontSize:24,fontWeight:800,color:s.c}}>{s.v}</div>
                </div>
              ))}
            </div>
            <div className="card">
              <div style={{overflowX:'auto'}}>
                <table>
                  <thead><tr><th>Egresado</th><th>Cédula</th><th>Secretario</th><th>Resultado</th><th>Fecha</th><th>Hora</th></tr></thead>
                  <tbody>
                    {validaciones.length===0
                      ?<tr><td colSpan={6} style={{textAlign:'center',padding:32,color:'#94A3B8'}}>Sin validaciones en este período</td></tr>
                      :validaciones.map((v:any)=>(
                      <tr key={v.id}>
                        <td><div style={{display:'flex',alignItems:'center',gap:8}}><span style={{fontSize:18}}>{RESULTADO_EMOJI[v.resultado]??'?'}</span><span style={{fontWeight:600}}>{v.egresado_nombre??'Desconocido'}</span></div></td>
                        <td style={{fontSize:12,color:'#6B7280'}}>{v.egresado_cedula??'—'}</td>
                        <td style={{fontSize:12,color:'#6B7280'}}>{v.secretario_nombre??'Admin'}</td>
                        <td><span className={RESULTADO_BADGE[v.resultado]??'badge inactive'}><span className="badge-dot"></span>{v.resultado}</span></td>
                        <td style={{fontSize:12,color:'#94A3B8'}}>{v.fecha}</td>
                        <td style={{fontWeight:700}}>{v.hora?.slice(0,8)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </>}

        </div>
      </div>

      {/* ══════════ MODAL IMPORTAR ══════════ */}
      <div className={`overlay ${modalImport?'show':''}`} onClick={e=>{if((e.target as any).classList?.contains('overlay')&&!importando)setModalImport(false)}}>
        <div className="modal" style={{maxWidth:600}}>
          <div className="modal-title">📥 Importar egresados desde Excel</div>
          {importRows.length===0?<>
            <div className="warn-box" style={{marginBottom:12}}>Columnas requeridas: <strong>Cédula</strong>, <strong>Nombre completo</strong>, <strong>Correo electrónico</strong>. Teléfono es opcional.</div>
            <div className="drop-zone"
              onDragOver={e=>{e.preventDefault();(e.currentTarget as HTMLElement).classList.add('drag')}}
              onDragLeave={e=>(e.currentTarget as HTMLElement).classList.remove('drag')}
              onDrop={e=>{e.preventDefault();(e.currentTarget as HTMLElement).classList.remove('drag');const f=e.dataTransfer.files[0];if(f)leerArchivoExcel(f)}}>
              <input type="file" accept=".xlsx,.xls,.csv" ref={fileInputRef} onChange={e=>{const f=e.target.files?.[0];if(f)leerArchivoExcel(f)}} />
              <div style={{fontSize:36,marginBottom:8}}>📂</div>
              <div style={{fontWeight:700,fontSize:14,color:'#111827',marginBottom:4}}>Arrastra el archivo aquí</div>
              <div style={{fontSize:12,color:'#94A3B8'}}>o haz clic para seleccionar · .xlsx .xls .csv</div>
            </div>
          </>:<>
            {!importDone&&<>
              <div className="auto-box" style={{marginBottom:10}}>Se encontraron <strong>{importRows.length}</strong> filas válidas. Define la fecha de inicio de membresía.</div>
              <label className="form-lbl">Fecha de inicio de membresía</label>
              <input type="date" id="importFechaInicio" defaultValue={fechaHoy()} className="form-inp" style={{marginTop:4}} />
              <div className="info-box" style={{marginTop:8}}>📅 Vencimiento = fecha inicio + 365 días (automático)</div>
            </>}
            {importDone&&(
              <div style={{display:'flex',gap:8,marginBottom:10,flexWrap:'wrap'}}>
                {importOk>0&&<span className="badge active">✅ {importOk} creados</span>}
                {importDup>0&&<span className="badge expired">⚠️ {importDup} duplicados</span>}
                {importErr>0&&<span className="badge inactive">❌ {importErr} errores</span>}
              </div>
            )}
            <div className="import-scroll">
              <table className="import-table">
                <thead><tr><th>#</th><th>Cédula</th><th>Nombre</th><th>Email</th>{importStatus.some(s=>s!=='idle')&&<th>Estado</th>}</tr></thead>
                <tbody>
                  {importRows.map((r,i)=>(
                    <tr key={i} style={{background:importStatus[i]==='ok'?'#F0FDF4':importStatus[i]==='err'?'#FEF2F2':importStatus[i]==='dup'?'#FFFBEB':''}}>
                      <td style={{color:'#94A3B8'}}>{i+1}</td>
                      <td>{r.cedula}</td>
                      <td style={{fontWeight:600}}>{r.nombre}</td>
                      <td style={{fontSize:11,color:'#6B7280'}}>{r.email}</td>
                      {importStatus.some(s=>s!=='idle')&&<td>
                        {importStatus[i]==='idle'&&<span style={{color:'#94A3B8',fontSize:11}}>Pendiente</span>}
                        {importStatus[i]==='ok'&&<span style={{color:'#16A34A',fontWeight:700,fontSize:11}}>✓ {importMsg[i]}</span>}
                        {importStatus[i]==='dup'&&<span style={{color:'#D97706',fontWeight:700,fontSize:11}}>⚠ Duplicado</span>}
                        {importStatus[i]==='err'&&<span style={{color:'#DC2626',fontWeight:700,fontSize:11}}>✕ Error</span>}
                      </td>}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>}
          <div style={{display:'flex',gap:10,marginTop:16}}>
            <button className="btn btn-outline" style={{flex:1,justifyContent:'center'}} disabled={importando}
              onClick={()=>{setImportRows([]);setImportStatus([]);setImportMsg([]);setImportDone(false);setModalImport(false);if(fileInputRef.current)fileInputRef.current.value=''}}>
              {importDone?'Cerrar':importRows.length?'Cambiar archivo':'Cancelar'}
            </button>
            {importRows.length>0&&!importDone&&(
              <button className="btn btn-primary" style={{flex:1,justifyContent:'center'}} onClick={ejecutarImportacion} disabled={importando}>
                {importando?`Importando ${importStatus.filter(s=>s!=='idle').length}/${importRows.length}...`:`✓ Importar ${importRows.length} egresados`}
              </button>
            )}
          </div>
        </div>
      </div>

      {/* ══════════ MODAL DETALLE EGRESADO (con tabs) ══════════ */}
      <div className={`overlay ${modalDetalle?'show':''}`}>
        <div className="modal modal-wide">
          {egSel&&<>
            {/* Header */}
            <div style={{display:'flex',alignItems:'center',gap:14,marginBottom:16}}>
              <div style={{width:64,height:64,borderRadius:'50%',background:'#FDE8EA',display:'flex',alignItems:'center',justifyContent:'center',fontSize:24,fontWeight:700,color:'#BE1522',overflow:'hidden',flexShrink:0}}>
                {egSel.foto_perfil?<img src={egSel.foto_perfil} style={{width:'100%',height:'100%',objectFit:'cover'}} alt=""/>:initials(egSel.nombre_completo)}
              </div>
              <div style={{flex:1}}>
                <div style={{fontSize:17,fontWeight:700}}>{egSel.nombre_completo}</div>
                <div style={{fontSize:13,color:'#94A3B8',marginTop:2}}>CC {egSel.cedula}</div>
                <div style={{display:'flex',gap:8,marginTop:6,flexWrap:'wrap'}}>
                  <span className={ESTADO_BADGE[egSel.estado]??'badge inactive'}><span className="badge-dot"></span>{egSel.estado}</span>
                  {egSel.empresa&&<span style={{fontSize:11,background:'#EFF6FF',color:'#1D4ED8',padding:'3px 10px',borderRadius:20,fontWeight:600}}>🏢 {egSel.empresa}</span>}
                </div>
              </div>
            </div>

            {/* Tabs detalle */}
            <div className="detail-tabs">
              {([{k:'info',l:'👤 Datos'},{k:'membresia',l:'🎫 Membresía'},{k:'estudios',l:'🎓 Estudios'},{k:'laboral',l:'💼 Laboral'}] as {k:typeof detTab;l:string}[]).map(t=>(
                <button key={t.k} className={`detail-tab ${detTab===t.k?'on':''}`} onClick={()=>setDetTab(t.k)}>{t.l}</button>
              ))}
            </div>

            {/* Tab: Datos personales */}
            {detTab==='info'&&<>
              <div className="detail-grid">
                {[
                  {k:'Email',v:egSel.email??'—'},
                  {k:'Teléfono',v:egSel.telefono??'—'},
                  {k:'Fecha nacimiento',v:fmtFechaLegible(egSel.fecha_nacimiento)},
                  {k:'Ciudad nacimiento',v:egSel.ciudad_nacimiento??'—'},
                  {k:'Dirección',v:egSel.direccion??'—'},
                  {k:'Registro',v:fmtFechaLegible(egSel.created_at?.split('T')[0])},
                  {k:'Hobbies',v:egSel.hobbies??'—'},
                ].map(r=>(
                  <div key={r.k} className="detail-row">
                    <span className="detail-key">{r.k}</span>
                    <span className="detail-val">{r.v}</span>
                  </div>
                ))}
              </div>
              {/* Editar email */}
              <div style={{background:'#F9F1F1',borderRadius:10,padding:12,marginTop:8}}>
                <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:8}}>
                  <span style={{fontSize:11,fontWeight:700,color:'#BE1522',textTransform:'uppercase',letterSpacing:.5}}>✏️ Cambiar email</span>
                  {!editandoEmail
                    ?<button className="btn btn-outline btn-sm" onClick={()=>{setEditandoEmail(true);setNuevoEmail(egSel.email??'')}}>Editar</button>
                    :<button className="btn btn-gray btn-sm" onClick={()=>setEditandoEmail(false)}>Cancelar</button>
                  }
                </div>
                {editandoEmail&&<div style={{display:'flex',gap:8}}>
                  <input className="form-inp" type="email" value={nuevoEmail} onChange={e=>setNuevoEmail(e.target.value.toLowerCase())} placeholder="nuevo@correo.com" style={{flex:1}} />
                  <button className="btn btn-primary btn-sm" onClick={()=>actualizarEmail(egSel,nuevoEmail)} disabled={savingEmail||nuevoEmail===egSel.email} style={{flexShrink:0}}>
                    {savingEmail?'...':'✓ Guardar'}
                  </button>
                </div>}
              </div>
            </>}

            {/* Tab: Membresía */}
            {detTab==='membresia'&&<>
              <div className="detail-grid">
                {[
                  {k:'Estado',v:<span className={ESTADO_BADGE[egSel.estado]??'badge inactive'}><span className="badge-dot"></span>{egSel.estado}</span>},
                  {k:'Inicio membresía',v:fmtFechaLegible(egSel.fecha_expedicion)},
                  {k:'Vencimiento',v:<span style={{color:'#BE1522',fontWeight:700}}>{fmtFechaLegible(egSel.fecha_vencimiento)}</span>},
                  {k:'Días restantes',v:egSel.fecha_vencimiento?(()=>{const d=diasRestantes(egSel.fecha_vencimiento);return d<0?<span style={{color:'#DC2626',fontWeight:700}}>Expirado hace {Math.abs(d)} días</span>:d<=30?<span style={{color:'#D97706',fontWeight:700}}>{d} días</span>:<span style={{color:'#16A34A',fontWeight:700}}>{d} días</span>})():'—'},
                ].map((r,i)=>(
                  <div key={i} className="detail-row">
                    <span className="detail-key">{r.k}</span>
                    <span className="detail-val">{r.v}</span>
                  </div>
                ))}
              </div>
              {/* Barra progreso */}
              {egSel.fecha_expedicion&&egSel.fecha_vencimiento&&(
                <div style={{margin:'8px 0 14px'}}>
                  <div style={{display:'flex',justifyContent:'space-between',fontSize:11,color:'#94A3B8',marginBottom:4}}>
                    <span>{egSel.fecha_expedicion}</span>
                    <span style={{fontWeight:700,color:egSel.estado==='activo'?'#16A34A':'#D97706'}}>{Math.min(100,porcentajeMembresia(egSel.fecha_expedicion,egSel.fecha_vencimiento))}% transcurrido</span>
                    <span>{egSel.fecha_vencimiento}</span>
                  </div>
                  <div className="prog-bar" style={{height:10,borderRadius:5}}>
                    <div className="prog-fill" style={{width:`${Math.min(100,porcentajeMembresia(egSel.fecha_expedicion,egSel.fecha_vencimiento))}%`,background:egSel.estado==='activo'?'#16A34A':'#D97706',borderRadius:5}}/>
                  </div>
                </div>
              )}
              {/* Renovar */}
              <div style={{background:'#F0FDF4',border:'1px solid #BBF7D0',borderRadius:12,padding:14}}>
                <div style={{fontSize:12,fontWeight:700,color:'#15803D',marginBottom:10}}>🔄 Renovar membresía (+365 días)</div>
                <label className="form-lbl" style={{marginTop:0}}>Nueva fecha de inicio</label>
                <input type="date" id="fechaRenovar" defaultValue={fechaHoy()} className="form-inp" style={{marginTop:4}} />
                <div style={{fontSize:11,color:'#94A3B8',marginTop:4}}>El vencimiento se calcula automáticamente: inicio + 365 días.</div>
                <button className="btn btn-success" style={{marginTop:10,width:'100%',justifyContent:'center'}}
                  onClick={()=>{const i=document.getElementById('fechaRenovar') as HTMLInputElement;renovarCarnet(egSel,i?.value||fechaHoy())}}>
                  🔄 Renovar membresía
                </button>
              </div>
            </>}

            {/* Tab: Estudios */}
            {detTab==='estudios'&&<>
              <div style={{fontSize:11,fontWeight:700,color:'#BE1522',textTransform:'uppercase',letterSpacing:.5,marginBottom:8}}>Pregrado</div>
              <div className="detail-grid">
                {[
                  {k:'Título',v:egSel.titulo_pregrado??'—'},
                  {k:'Institución',v:egSel.institucion_pregrado??'—'},
                  {k:'Fecha de grado',v:fmtFechaLegible(egSel.fecha_grado_pregrado)},
                ].map(r=>(
                  <div key={r.k} className="detail-row"><span className="detail-key">{r.k}</span><span className="detail-val">{r.v}</span></div>
                ))}
              </div>
              <div style={{fontSize:11,fontWeight:700,color:'#BE1522',textTransform:'uppercase',letterSpacing:.5,marginBottom:8,marginTop:14}}>Posgrado</div>
              <div className="detail-grid">
                {[
                  {k:'Título',v:egSel.titulo_posgrado??'—'},
                  {k:'Institución',v:egSel.institucion_posgrado??'—'},
                  {k:'Fecha de grado',v:fmtFechaLegible(egSel.fecha_grado_posgrado)},
                ].map(r=>(
                  <div key={r.k} className="detail-row"><span className="detail-key">{r.k}</span><span className="detail-val">{r.v}</span></div>
                ))}
              </div>
            </>}

            {/* Tab: Laboral */}
            {detTab==='laboral'&&<>
              <div className="detail-grid">
                {[
                  {k:'Empresa',v:egSel.empresa??'—'},
                  {k:'Cargo',v:egSel.cargo??'—'},
                  {k:'Hobbies',v:egSel.hobbies??'—'},
                ].map(r=>(
                  <div key={r.k} className="detail-row"><span className="detail-key">{r.k}</span><span className="detail-val">{r.v}</span></div>
                ))}
              </div>
            </>}

            {/* Acciones */}
            <div className="actions-row" style={{marginTop:16}}>
              {egSel.estado!=='activo'  &&<button className="btn btn-success" onClick={()=>cambiarEstado(egSel,'activo')}>▶ Activar</button>}
              {egSel.estado!=='inactivo'&&<button className="btn btn-outline"  onClick={()=>cambiarEstado(egSel,'inactivo')}>⏸ Desactivar</button>}
              {egSel.estado!=='vencido' &&<button className="btn btn-sm" style={{background:'#FEF3C7',color:'#92400E',border:'1px solid #FDE68A'}} onClick={()=>cambiarEstado(egSel,'vencido')}>⚠ Marcar vencido</button>}
              <button className="btn btn-danger" onClick={()=>eliminarEgresado(egSel)}>🗑️ Eliminar</button>
              <button className="btn btn-gray" style={{marginLeft:'auto'}} onClick={()=>setModalDetalle(false)}>Cerrar</button>
            </div>
          </>}
        </div>
      </div>

      {/* ══════════ MODAL NUEVO EGRESADO (completo) ══════════ */}
      <div className={`overlay ${modalNuevoEg?'show':''}`} onClick={e=>{if((e.target as any).className?.includes?.('overlay'))setModalNuevoEg(false)}}>
        <div className="modal modal-wide">
          <div className="modal-title">👤 Nuevo egresado</div>
          <div className="auto-box">Contraseña inicial = cédula. Vencimiento = inicio membresía + 365 días.</div>

          <div className="form-section">Datos básicos *</div>
          <div className="form-row2">
            <div><label className="form-lbl">Cédula *</label><input className="form-inp" placeholder="Sin puntos ni espacios" value={formEg.cedula} onChange={e=>updFormEg('cedula',e.target.value)} /></div>
            <div><label className="form-lbl">Teléfono</label><input className="form-inp" placeholder="+57 300..." value={formEg.telefono} onChange={e=>updFormEg('telefono',e.target.value)} /></div>
          </div>
          <label className="form-lbl">Nombre completo *</label>
          <input className="form-inp" placeholder="Nombres y apellidos" value={formEg.nombre} onChange={e=>updFormEg('nombre',e.target.value)} />
          <label className="form-lbl">Email *</label>
          <input className="form-inp" type="email" placeholder="correo@dominio.com" value={formEg.email} onChange={e=>updFormEg('email',e.target.value)} />
          <div className="form-row2">
            <div><label className="form-lbl">Inicio membresía *</label><input className="form-inp" type="date" value={formEg.inicioMembresia} onChange={e=>onChangeInicio(e.target.value)} /></div>
            <div><label className="form-lbl">Vencimiento (auto +365d)</label><input className="form-inp" value={formEg.vencimiento} disabled /></div>
          </div>

          <div className="form-section">Datos personales</div>
          <div className="form-row3">
            <div><label className="form-lbl">Fecha nacimiento</label><input className="form-inp" type="date" value={formEg.fecha_nacimiento} onChange={e=>updFormEg('fecha_nacimiento',e.target.value)} /></div>
            <div><label className="form-lbl">Ciudad nacimiento</label><input className="form-inp" placeholder="Bucaramanga" value={formEg.ciudad_nacimiento} onChange={e=>updFormEg('ciudad_nacimiento',e.target.value)} /></div>
            <div><label className="form-lbl">Dirección</label><input className="form-inp" placeholder="Dirección actual" value={formEg.direccion} onChange={e=>updFormEg('direccion',e.target.value)} /></div>
          </div>

          <div className="form-section">Estudios de pregrado</div>
          <div className="form-row3">
            <div><label className="form-lbl">Título</label><input className="form-inp" placeholder="Ing. de Sistemas" value={formEg.titulo_pregrado} onChange={e=>updFormEg('titulo_pregrado',e.target.value)} /></div>
            <div><label className="form-lbl">Institución</label><input className="form-inp" placeholder="UIS" value={formEg.institucion_pregrado} onChange={e=>updFormEg('institucion_pregrado',e.target.value)} /></div>
            <div><label className="form-lbl">Fecha de grado</label><input className="form-inp" type="date" value={formEg.fecha_grado_pregrado} onChange={e=>updFormEg('fecha_grado_pregrado',e.target.value)} /></div>
          </div>

          <div className="form-section">Estudios de posgrado (opcional)</div>
          <div className="form-row3">
            <div><label className="form-lbl">Título</label><input className="form-inp" placeholder="Magíster en..." value={formEg.titulo_posgrado} onChange={e=>updFormEg('titulo_posgrado',e.target.value)} /></div>
            <div><label className="form-lbl">Institución</label><input className="form-inp" placeholder="UIS" value={formEg.institucion_posgrado} onChange={e=>updFormEg('institucion_posgrado',e.target.value)} /></div>
            <div><label className="form-lbl">Fecha de grado</label><input className="form-inp" type="date" value={formEg.fecha_grado_posgrado} onChange={e=>updFormEg('fecha_grado_posgrado',e.target.value)} /></div>
          </div>

          <div className="form-section">Información laboral (opcional)</div>
          <div className="form-row2">
            <div><label className="form-lbl">Empresa</label><input className="form-inp" placeholder="Bancolombia" value={formEg.empresa} onChange={e=>updFormEg('empresa',e.target.value)} /></div>
            <div><label className="form-lbl">Cargo</label><input className="form-inp" placeholder="Analista senior" value={formEg.cargo} onChange={e=>updFormEg('cargo',e.target.value)} /></div>
          </div>

          <div style={{display:'flex',gap:10,marginTop:20}}>
            <button className="btn btn-outline" style={{flex:1,justifyContent:'center'}} onClick={()=>setModalNuevoEg(false)}>Cancelar</button>
            <button className="btn btn-primary" style={{flex:1,justifyContent:'center'}} onClick={crearEgresado} disabled={saving}>{saving?'Creando...':'✓ Crear egresado'}</button>
          </div>
        </div>
      </div>

      {/* ══════════ MODAL NUEVO SECRETARIO ══════════ */}
      <div className={`overlay ${modalNuevoSec?'show':''}`} onClick={e=>{if((e.target as any).className?.includes?.('overlay'))setModalNuevoSec(false)}}>
        <div className="modal">
          <div className="modal-title">🔑 Nuevo secretario</div>
          <div className="auto-box">Contraseña inicial = cédula del secretario.</div>
          <label className="form-lbl">Cédula *</label>
          <input className="form-inp" placeholder="Sin puntos" value={formSec.cedula} onChange={e=>setFormSec(p=>({...p,cedula:e.target.value}))} />
          <label className="form-lbl">Nombre completo *</label>
          <input className="form-inp" placeholder="Nombre del funcionario" value={formSec.nombre} onChange={e=>setFormSec(p=>({...p,nombre:e.target.value}))} />
          <label className="form-lbl">Email institucional *</label>
          <input className="form-inp" type="email" placeholder="secretaria@aseduis.com" value={formSec.email} onChange={e=>setFormSec(p=>({...p,email:e.target.value}))} />
          <div style={{display:'flex',gap:10,marginTop:20}}>
            <button className="btn btn-outline" style={{flex:1,justifyContent:'center'}} onClick={()=>setModalNuevoSec(false)}>Cancelar</button>
            <button className="btn btn-primary" style={{flex:1,justifyContent:'center'}} onClick={crearSecretario} disabled={saving}>{saving?'Creando...':'✓ Crear secretario'}</button>
          </div>
        </div>
      </div>

    </div>
  )
}