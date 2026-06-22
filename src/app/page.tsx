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

// ── Exportar Excel ────────────────────────────────────────────
function exportarExcel(datos: any[], nombre: string) {
  const cols = [
    'cedula','nombre_completo','email','telefono','estado',
    'fecha_nacimiento','empresa','cargo','ciudad_nacimiento','direccion',
    'fecha_vencimiento','fecha_expedicion',
    'titulo_pregrado','institucion_pregrado','fecha_grado_pregrado',
    'titulo_posgrado','institucion_posgrado','fecha_grado_posgrado',
    'created_at',
  ]
  const headers = [
    'Cédula','Nombre completo','Email','Teléfono','Estado',
    'Fecha nacimiento','Empresa','Cargo','Ciudad','Dirección',
    'Vencimiento','Fecha expedición',
    'Título pregrado','Institución pregrado','Fecha grado pregrado',
    'Título posgrado','Institución posgrado','Fecha grado posgrado',
    'Fecha registro',
  ]
  const esc = (v: any) => {
    const s = (v ?? '').toString().replace(/"/g, '""')
    return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s}"` : s
  }
  const filas = [headers.join(','), ...datos.map(row => cols.map(c => esc(row[c])).join(','))].join('\n')
  const blob = new Blob(['\uFEFF' + filas], { type: 'text/csv;charset=utf-8;' })
  const url  = URL.createObjectURL(blob)
  const a    = document.createElement('a')
  a.href = url; a.download = `${nombre}.csv`; a.click()
  URL.revokeObjectURL(url)
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
      <td><span class="badge ${eg.estado}">${eg.estado??''}</span></td>
      <td>${eg.empresa??'—'}</td>
      <td>${eg.cargo??'—'}</td>
      <td style="color:${eg.estado==='vencido'?'#D97706':'inherit'};font-weight:${eg.estado==='vencido'?700:400}">${eg.fecha_vencimiento??'—'}</td>
    </tr>`).join('')
  const html = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>${nombre}</title>
  <style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:"Segoe UI",sans-serif;color:#111;font-size:11px}
  .header{background:#BE1522;color:#fff;padding:18px 24px;display:flex;justify-content:space-between;align-items:center;margin-bottom:20px}
  .header h1{font-size:18px;font-weight:800;letter-spacing:1px}.header p{font-size:11px;opacity:.8;margin-top:3px}
  .meta{padding:0 24px 14px;display:flex;gap:24px}.meta span{font-size:11px;color:#6B7280}.meta strong{color:#111}
  table{width:100%;border-collapse:collapse;font-size:10.5px}
  th{background:#FDE8EA;color:#BE1522;font-weight:700;text-transform:uppercase;font-size:9px;letter-spacing:.4px;padding:8px 12px;text-align:left;border-bottom:2px solid #F0D4D6}
  td{padding:8px 12px;border-bottom:1px solid #F5F5F5;vertical-align:middle}tr:nth-child(even) td{background:#FDF8F8}
  .badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:9px;font-weight:700}
  .badge.activo{background:#DCFCE7;color:#16A34A}.badge.vencido{background:#FEF3C7;color:#D97706}.badge.inactivo{background:#FDE8EA;color:#BE1522}
  .footer{margin-top:20px;padding:12px 24px;border-top:1px solid #F0D4D6;display:flex;justify-content:space-between;font-size:10px;color:#94A3B8}
  @media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}</style></head><body>
  <div class="header"><div><h1>ASEDUIS</h1><p>Reporte de Egresados</p></div>
  <div style="text-align:right"><div style="font-size:13px;font-weight:700">${datos.length} egresados</div><div style="font-size:10px;opacity:.8">${hoy}</div></div></div>
  <div class="meta">
    <span>📊 Total: <strong>${datos.length}</strong></span>
    <span>✅ Activos: <strong>${datos.filter((e:any)=>e.estado==='activo').length}</strong></span>
    <span>⚠️ Vencidos: <strong>${datos.filter((e:any)=>e.estado==='vencido').length}</strong></span>
    <span>❌ Inactivos: <strong>${datos.filter((e:any)=>e.estado==='inactivo').length}</strong></span>
  </div>
  <table><thead><tr><th>Cédula</th><th>Nombre</th><th>Email</th><th>Teléfono</th><th>Estado</th><th>Empresa</th><th>Cargo</th><th>Vencimiento</th></tr></thead>
  <tbody>${filas}</tbody></table>
  <div class="footer"><span>Generado: ${hoy} · Sistema ASEDUIS</span><span>Página 1</span></div>
  <script>window.onload=()=>{window.print()}</script></body></html>`
  const blob = new Blob([html], { type: 'text/html;charset=utf-8;' })
  const url  = URL.createObjectURL(blob)
  window.open(url, '_blank')
  setTimeout(() => URL.revokeObjectURL(url), 60000)
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

// ── Normalizar encabezados del Excel ─────────────────────────
function normalizarHeader(h: string): string {
  return h.toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')   // quitar tildes
    .replace(/[^a-z0-9]/g, ' ')
    .trim()
    .replace(/\s+/g, '_')
}
// Mapeo de variantes de encabezado → campo interno
const HEADER_MAP: Record<string, string> = {
  cedula: 'cedula',
  numero_de_cedula: 'cedula',
  documento: 'cedula',
  nombre_completo: 'nombre',
  nombre: 'nombre',
  nombres: 'nombre',
  correo_electronico: 'email',
  correo: 'email',
  email: 'email',
  correo_institucional: 'email',
  telefono: 'telefono',
  celular: 'telefono',
}

const CSS = `
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:"Segoe UI",system-ui,sans-serif;background:#F9F1F1;color:#111827}

  /* ── Layout ── */
  .layout{display:flex;min-height:100vh}
  .sidebar{width:240px;background:#BE1522;min-height:100vh;display:flex;flex-direction:column;position:sticky;top:0;height:100vh;overflow-y:auto;flex-shrink:0;transition:transform .25s}
  .main{flex:1;display:flex;flex-direction:column;min-width:0;overflow:hidden}

  /* ── Mobile sidebar overlay ── */
  .sidebar-backdrop{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:199}
  @media(max-width:768px){
    .sidebar{position:fixed;top:0;left:0;height:100vh;z-index:200;transform:translateX(-100%)}
    .sidebar.open{transform:translateX(0)}
    .sidebar-backdrop.show{display:block}
  }

  /* ── Nav ── */
  .nav-item{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:10px;cursor:pointer;margin-bottom:2px;transition:background .15s;user-select:none}
  .nav-item:hover{background:rgba(255,255,255,.12)}
  .nav-item.active{background:rgba(255,255,255,.22)}
  .nav-lbl{font-size:13px;color:rgba(255,255,255,.7);font-weight:500}
  .nav-item.active .nav-lbl{color:#fff;font-weight:700}

  /* ── Topbar ── */
  .topbar{background:#fff;border-bottom:1px solid #E5E7EB;padding:14px 28px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;gap:10px;flex-wrap:wrap}
  .topbar-actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
  .menu-btn{display:none;background:none;border:none;cursor:pointer;font-size:22px;padding:4px 8px;border-radius:8px;color:#BE1522;flex-shrink:0}
  @media(max-width:768px){
    .menu-btn{display:flex;align-items:center}
    .topbar{padding:12px 16px}
  }

  /* ── Content ── */
  .content{flex:1;padding:24px 28px;overflow-y:auto}
  @media(max-width:768px){.content{padding:16px}}

  /* ── Stats grid ── */
  .stats-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:22px}
  @media(max-width:900px){.stats-grid{grid-template-columns:repeat(2,1fr)}}
  @media(max-width:480px){.stats-grid{grid-template-columns:1fr 1fr}}

  /* ── Cards ── */
  .stat-card{background:#fff;border-radius:14px;border:1px solid #F0D4D6;padding:18px 20px;border-left-width:4px}
  .card{background:#fff;border-radius:14px;border:1px solid #F0D4D6;overflow:hidden;margin-bottom:16px}
  .card-hdr{padding:14px 20px;border-bottom:1px solid #F0D4D6;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}

  /* ── Table ── */
  table{width:100%;border-collapse:collapse}
  th{background:#FDE8EA;font-size:11px;font-weight:700;color:#94A3B8;text-transform:uppercase;padding:11px 16px;text-align:left;white-space:nowrap}
  td{font-size:13px;color:#111827;padding:12px 16px;border-top:1px solid #FDF0F0;vertical-align:middle}
  tr:hover td{background:#FDF8F8}

  /* ── Badges & Buttons ── */
  .badge{display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700}
  .badge.active{background:#DCFCE7;color:#16A34A}
  .badge.expired{background:#FEF3C7;color:#D97706}
  .badge.inactive{background:#FDE8EA;color:#BE1522}
  .badge.warn{background:#FEF3C7;color:#D97706}
  .badge-dot{width:5px;height:5px;border-radius:50%;background:currentColor;flex-shrink:0}
  .btn{border-radius:10px;padding:8px 16px;font-size:12px;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;gap:6px;border:none;transition:opacity .15s;font-family:inherit}
  .btn:hover{opacity:.88}
  .btn:disabled{opacity:.45;cursor:not-allowed}
  .btn-primary{background:#BE1522;color:#fff}
  .btn-outline{background:#fff;color:#BE1522;border:1.5px solid #BE1522}
  .btn-success{background:#DCFCE7;color:#16A34A;border:1px solid #BBF7D0}
  .btn-danger{background:#FEE2E2;color:#DC2626;border:1px solid #FECACA}
  .btn-green{background:#DCFCE7;color:#16A34A;border:1.5px solid #BBF7D0}
  .btn-sm{padding:5px 12px;font-size:11px;border-radius:8px}
  .tab-btn{padding:6px 14px;border-radius:20px;font-size:12px;font-weight:700;cursor:pointer;border:none;background:#F9FAFB;color:#94A3B8;transition:.15s;font-family:inherit}
  .tab-btn.on{background:#BE1522;color:#fff}
  .action-btn{background:none;border:none;cursor:pointer;padding:5px 7px;border-radius:6px;font-size:16px;transition:.15s}
  .action-btn:hover{background:#FDE8EA}
  .toggle{width:40px;height:22px;background:#E5E7EB;border-radius:11px;position:relative;cursor:pointer;border:none;transition:.2s;flex-shrink:0}
  .toggle.on{background:#BE1522}
  .toggle::after{content:'';position:absolute;width:16px;height:16px;background:#fff;border-radius:50%;top:3px;left:3px;transition:.2s;box-shadow:0 1px 3px rgba(0,0,0,.2)}
  .toggle.on::after{left:21px}

  /* ── Search & filters ── */
  .search-row{display:flex;gap:10px;margin-bottom:14px;align-items:center;flex-wrap:wrap}
  .search-box{display:flex;align-items:center;gap:8px;background:#F9FAFB;border:1.5px solid #E5E7EB;border-radius:10px;padding:8px 14px;flex:1;min-width:180px}
  .search-box input{border:none;outline:none;background:transparent;font-size:13px;color:#111827;flex:1;font-family:inherit}

  /* ── Modals ── */
  .overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:300;align-items:center;justify-content:center;padding:16px}
  .overlay.show{display:flex}
  .modal{background:#fff;border-radius:20px;width:100%;max-width:520px;padding:28px;max-height:90vh;overflow-y:auto}
  @media(max-width:540px){.modal{padding:20px 16px;border-radius:14px}}
  .modal-title{font-size:18px;font-weight:700;color:#111827;margin-bottom:16px}
  .form-lbl{font-size:11px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.4px;margin-bottom:5px;margin-top:14px;display:block}
  .form-inp{width:100%;background:#F9FAFB;border:1.5px solid #E5E7EB;border-radius:10px;padding:11px 14px;font-size:14px;outline:none;color:#111827;font-family:inherit}
  .form-inp:focus{border-color:#BE1522}
  .form-inp:disabled{opacity:.6;background:#F5F5F5}
  .form-row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  @media(max-width:480px){.form-row2{grid-template-columns:1fr}}
  .auto-box{background:#DCFCE7;border:1px solid #BBF7D0;border-radius:10px;padding:12px;font-size:12px;color:#15803D;line-height:1.6;margin-bottom:4px}
  .warn-box{background:#FEF3C7;border:1px solid #FDE68A;border-radius:10px;padding:12px;font-size:12px;color:#92400E;line-height:1.6;margin-bottom:4px}
  .err-box{background:#FEE2E2;border:1px solid #FECACA;border-radius:10px;padding:12px;font-size:12px;color:#DC2626;line-height:1.6;margin-bottom:4px}
  .detail-grid{background:#F9F1F1;border-radius:12px;padding:14px;margin:14px 0}
  .detail-section{font-size:10px;font-weight:700;color:#BE1522;text-transform:uppercase;letter-spacing:.5px;margin:10px 0 6px}
  .detail-row{display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #F0D4D6}
  .detail-row:last-child{border-bottom:none}
  .detail-key{font-size:11px;font-weight:700;color:#94A3B8;text-transform:uppercase;letter-spacing:.3px}
  .detail-val{font-size:13px;font-weight:500;color:#111827;text-align:right;max-width:65%;word-break:break-all}
  .actions-row{display:flex;gap:8px;flex-wrap:wrap;margin:14px 0}
  .export-bar{display:flex;gap:8px;align-items:center;background:#fff;border:1px solid #F0D4D6;border-radius:12px;padding:10px 14px;margin-bottom:14px;flex-wrap:wrap}
  .export-bar span{font-size:12px;color:#94A3B8;font-weight:600;margin-right:4px}
  .val-hoy-box{background:#FDE8EA;border:1px solid #F0D4D6;border-radius:12px;padding:14px 18px;margin-bottom:16px;display:flex;align-items:center;gap:16px;flex-wrap:wrap}

  /* ── Import progress table ── */
  .import-table{font-size:12px;width:100%;border-collapse:collapse;margin-top:10px;max-height:180px;overflow-y:auto;display:block}
  .import-table th{background:#F9FAFB;font-weight:700;padding:6px 10px;text-align:left;border-bottom:1px solid #E5E7EB;position:sticky;top:0}
  .import-table td{padding:5px 10px;border-bottom:1px solid #F5F5F5}

  /* ── File drop zone ── */
  .drop-zone{border:2px dashed #E5E7EB;border-radius:12px;padding:28px;text-align:center;cursor:pointer;transition:.2s;background:#FAFAFA;position:relative}
  .drop-zone:hover,.drop-zone.drag{border-color:#BE1522;background:#FDE8EA10}
  .drop-zone input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
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
  const [stats, setStats] = useState({ activos:0, vencidos:0, inactivos:0, total:0 })
  const [busqueda,      setBusqueda]     = useState('')
  const [busquedaSec,   setBusquedaSec]  = useState('')
  const [filtroEstado,  setFiltroEstado] = useState<FiltroEstado>('todos')
  const [filtroFecha,   setFiltroFecha]  = useState<FiltroFecha>('hoy')
  const [modalNuevoEg,  setModalNuevoEg]  = useState(false)
  const [modalNuevoSec, setModalNuevoSec] = useState(false)
  const [modalDetalle,  setModalDetalle]  = useState(false)
  const [modalImport,   setModalImport]   = useState(false)
  const [egSel,         setEgSel]         = useState<any>(null)
  const [sidebarOpen,   setSidebarOpen]   = useState(false)

  // ── Import state ──────────────────────────────────────────
  const [importRows,    setImportRows]    = useState<any[]>([])
  const [importStatus,  setImportStatus]  = useState<('idle'|'ok'|'err'|'dup')[]>([])
  const [importMsg,     setImportMsg]     = useState<string[]>([])
  const [importando,    setImportando]    = useState(false)
  const [importDone,    setImportDone]    = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const [formEg, setFormEg] = useState({
    cedula:'', nombre:'', email:'', telefono:'',
    expedicion: fechaHoy(), vencimiento: sumar365(fechaHoy()), estado:'activo',
  })
  const [formSec, setFormSec] = useState({ cedula:'', nombre:'', email:'' })
  const [saving,  setSaving]  = useState(false)
  const [editandoEmail, setEditandoEmail] = useState(false)
  const [nuevoEmail,    setNuevoEmail]    = useState('')
  const [savingEmail,   setSavingEmail]   = useState(false)

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => { setSesion(session); setCargandoAuth(false) })
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, s) => setSesion(s))
    return () => subscription.unsubscribe()
  }, [])

  useEffect(() => { if (sesion) cargarTodo() }, [sesion])
  useEffect(() => { if (sesion && tab === 'validaciones') cargarValidaciones() }, [tab, filtroFecha])

  async function cargarTodo() {
    const [{ data: eg }, { data: sec }, { data: val }] = await Promise.all([
      supabase.from('egresados').select('*').order('nombre_completo'),
      supabase.from('secretarios').select('*').order('nombre_completo'),
      supabase.from('validaciones_detalle').select('*').eq('fecha', fechaHoy()).order('hora_validacion', { ascending: false }).limit(5),
    ])
    const lista = eg || []
    setEgresados(lista)
    setSecretarios(sec || [])
    setValHoy(val || [])
    setStats({
      activos:   lista.filter((e:any) => e.estado==='activo').length,
      vencidos:  lista.filter((e:any) => e.estado==='vencido').length,
      inactivos: lista.filter((e:any) => e.estado==='inactivo').length,
      total:     lista.length,
    })
  }

  async function cargarValidaciones() {
    let q = supabase.from('validaciones_detalle').select('*')
      .order('hora_validacion', { ascending:false }).limit(100)
    const hoy = new Date()
    if (filtroFecha==='hoy') q = q.eq('fecha', fechaHoy())
    else if (filtroFecha==='semana') { const d=new Date(hoy); d.setDate(hoy.getDate()-7); q = q.gte('hora_validacion', d.toISOString()) }
    else if (filtroFecha==='mes')    { const d=new Date(hoy); d.setDate(hoy.getDate()-30); q = q.gte('hora_validacion', d.toISOString()) }
    const { data } = await q
    setValidaciones(data || [])
  }

  // ── IMPORTAR EXCEL ────────────────────────────────────────
  function leerArchivoExcel(file: File) {
    const reader = new FileReader()
    reader.onload = (e) => {
      const data = new Uint8Array(e.target!.result as ArrayBuffer)
      const wb   = XLSX.read(data, { type: 'array' })
      const ws   = wb.Sheets[wb.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json<any>(ws, { defval: '' })

      if (!rows.length) { alert('El archivo está vacío.'); return }

      // Normalizar las keys de cada fila
      const normalized = rows.map((row: any) => {
        const out: Record<string, string> = {}
        for (const key of Object.keys(row)) {
          const norm = normalizarHeader(String(key))
          const campo = HEADER_MAP[norm]
          if (campo) out[campo] = String(row[key]).trim()
        }
        return out
      }).filter(r => r.cedula && r.nombre && r.email)   // omitir filas incompletas

      if (!normalized.length) {
        alert('No se encontraron filas válidas. Asegúrate que el Excel tenga columnas: Cédula, Nombre completo, Correo electrónico.')
        return
      }

      setImportRows(normalized)
      setImportStatus(normalized.map(() => 'idle'))
      setImportMsg(normalized.map(() => ''))
      setImportDone(false)
      setModalImport(true)
    }
    reader.readAsArrayBuffer(file)
  }

  async function ejecutarImportacion() {
    setImportando(true)
    const status = [...importStatus]
    const msgs   = [...importMsg]
    const hoy    = fechaHoy()
    const venc   = sumar365(hoy)

    for (let i = 0; i < importRows.length; i++) {
      const row = importRows[i]
      const cedula = row.cedula.replace(/[.\s]/g, '')

      // Verificar duplicado
      const { data: existing } = await supabase
        .from('egresados').select('id').eq('cedula', cedula).maybeSingle()

      if (existing) {
        status[i] = 'dup'
        msgs[i]   = 'Ya existe'
        setImportStatus([...status])
        setImportMsg([...msgs])
        continue
      }

      const { error } = await supabase.from('egresados').insert({
        cedula,
        nombre_completo:     row.nombre,
        email:               row.email.toLowerCase(),
        telefono:            row.telefono || null,
        fecha_vencimiento:   venc,
        fecha_expedicion:    hoy,
        estado:              'activo',
        requiere_cambio_clave: true,
      })

      if (error) {
        status[i] = 'err'
        msgs[i]   = error.message
      } else {
        status[i] = 'ok'
        msgs[i]   = 'Creado ✓'
      }
      setImportStatus([...status])
      setImportMsg([...msgs])
    }

    setImportando(false)
    setImportDone(true)
    cargarTodo()
  }

  const egresadosFiltrados = egresados.filter(eg => {
    const matchEstado = filtroEstado==='todos' || eg.estado===filtroEstado
    const q = busqueda.toLowerCase()
    const matchBusqueda = !q ||
      eg.nombre_completo?.toLowerCase().includes(q) ||
      eg.cedula?.includes(q) ||
      (eg.email??'').toLowerCase().includes(q) ||
      (eg.empresa??'').toLowerCase().includes(q)
    return matchEstado && matchBusqueda
  })

  const secretariosFiltrados = secretarios.filter(sec => {
    const q = busquedaSec.toLowerCase()
    return !q ||
      sec.nombre_completo?.toLowerCase().includes(q) ||
      sec.cedula?.includes(q) ||
      (sec.email??'').toLowerCase().includes(q)
  })

  function onChangeExpedicion(val: string) { setFormEg(p => ({ ...p, expedicion:val, vencimiento:sumar365(val) })) }

  async function manejarLogin(e: React.FormEvent) {
    e.preventDefault()
    const { error } = await supabase.auth.signInWithPassword({ email:emailLogin.trim(), password:claveLogin })
    if (error) alert('Error: ' + error.message)
  }

  async function manejarLogout() { await supabase.auth.signOut(); setSesion(null) }

  async function crearEgresado() {
    if (!formEg.cedula || !formEg.nombre || !formEg.email) return alert('Cédula, nombre y email son obligatorios.')
    setSaving(true)
    const { error } = await supabase.from('egresados').insert({
      cedula: formEg.cedula.replace(/[.\s]/g,''), nombre_completo:formEg.nombre,
      email:formEg.email.toLowerCase(), telefono:formEg.telefono||null,
      fecha_vencimiento:formEg.vencimiento, estado:formEg.estado, requiere_cambio_clave:true,
    })
    setSaving(false)
    if (error) return alert('Error: ' + error.message)
    alert(`✓ Egresado creado\nContraseña inicial: ${formEg.cedula}`)
    setFormEg({ cedula:'', nombre:'', email:'', telefono:'', expedicion:fechaHoy(), vencimiento:sumar365(fechaHoy()), estado:'activo' })
    setModalNuevoEg(false); cargarTodo()
  }

  async function crearSecretario() {
    if (!formSec.cedula || !formSec.nombre || !formSec.email) return alert('Cédula, nombre y email son obligatorios.')
    setSaving(true)
    const { error } = await supabase.from('secretarios').insert({
      cedula:formSec.cedula.replace(/[.\s]/g,''), nombre_completo:formSec.nombre,
      email:formSec.email.toLowerCase(), activo:true, requiere_cambio_clave:true,
    })
    setSaving(false)
    if (error) return alert('Error: ' + error.message)
    alert(`✓ Secretario creado\nContraseña inicial: ${formSec.cedula}`)
    setFormSec({ cedula:'', nombre:'', email:'' }); setModalNuevoSec(false); cargarTodo()
  }

  async function renovarCarnet(eg: any, fechaBase: string) {
    const venc = sumar365(fechaBase)
    if (!confirm(`¿Renovar carnet de ${eg.nombre_completo}?\n\nInicio: ${fechaBase}\nVence: ${venc}`)) return
    const { error } = await supabase.from('egresados').update({ estado:'activo', fecha_vencimiento:venc }).eq('id',eg.id)
    if (error) return alert('Error: ' + error.message)
    alert(`✓ Carnet renovado hasta ${venc}`)
    setModalDetalle(false); cargarTodo()
  }

  async function cambiarEstado(eg: any, estado: string) {
    if (!confirm(`¿${estado==='activo'?'Activar':'Desactivar'} a ${eg.nombre_completo}?`)) return
    await supabase.from('egresados').update({ estado }).eq('id',eg.id)
    setModalDetalle(false); cargarTodo()
  }

  async function eliminarEgresado(eg: any) {
    if (!confirm(`⚠️ ELIMINAR PERMANENTE\n\n¿Eliminar a ${eg.nombre_completo}?\n\nEsta acción es IRREVERSIBLE.`)) return
    if (!confirm(`Confirma una vez más: ¿eliminar a ${eg.nombre_completo}?`)) return
    const { error } = await supabase.from('egresados').delete().eq('id',eg.id)
    if (error) return alert('Error: ' + error.message)
    alert('✓ Egresado eliminado.')
    setModalDetalle(false); cargarTodo()
  }

  async function actualizarEmail(eg: any, email: string) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/
    if (!emailRegex.test(email)) { alert('Email inválido.'); return }
    if (!confirm(`¿Cambiar email de ${eg.nombre_completo}?\n\nNuevo email: ${email}`)) return
    setSavingEmail(true)
    const { error } = await supabase
      .from('egresados').update({ email: email.toLowerCase() }).eq('id', eg.id)
    if (error) { alert('Error: ' + error.message); setSavingEmail(false); return }
    const res = await fetch('/api/update-email', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ user_id: eg.user_id, email: email.toLowerCase() }),
    })
    const result = await res.json()
    if (!res.ok) { alert('Email actualizado en BD pero no en auth: ' + result.error); setSavingEmail(false); return }
    setSavingEmail(false)
    alert('✓ Email actualizado correctamente.')
    setEditandoEmail(false); setNuevoEmail(''); cargarTodo()
  }

  async function toggleSecretario(sec: any) {
    await supabase.from('secretarios').update({ activo:!sec.activo }).eq('id',sec.id)
    cargarTodo()
  }

  function abrirDetalle(eg: any) {
    setEgSel(eg); setModalDetalle(true); setEditandoEmail(false); setNuevoEmail(eg.email ?? '')
  }
  function initials(n: string) { return n.split(' ').slice(0,2).map((x:string)=>x[0]).join('').toUpperCase() }
  function nombreArchivo(p: string) { return `${p}-${fechaHoy()}` }

  // Conteos de importación
  const importOk  = importStatus.filter(s=>s==='ok').length
  const importErr = importStatus.filter(s=>s==='err').length
  const importDup = importStatus.filter(s=>s==='dup').length

  function navTo(t: Tab) { setTab(t); setSidebarOpen(false) }

  // ── Login screen ──────────────────────────────────────────
  if (cargandoAuth) return (
    <div style={{ display:'flex', height:'100vh', alignItems:'center', justifyContent:'center', fontFamily:'sans-serif', background:'#F9F1F1', color:'#BE1522', fontWeight:600, fontSize:16 }}>
      Cargando ASEDUIS...
    </div>
  )

  if (!sesion) return (
    <>
      <style>{`
        *{box-sizing:border-box}
        .login-page{display:flex;min-height:100vh;align-items:center;justify-content:center;background:#BE1522;font-family:"Segoe UI",system-ui,sans-serif;padding:16px}
        .login-card{background:#fff;padding:40px;border-radius:20px;box-shadow:0 8px 40px rgba(0,0,0,.18);width:100%;max-width:400px}
        .login-logo{width:72px;height:72px;background:#BE1522;border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 14px;font-size:28px;font-weight:900;color:#fff;overflow:hidden}
        .login-logo img{width:100%;height:100%;object-fit:contain;padding:8px}
        .login-title{color:#BE1522;font-weight:800;font-size:22px;letter-spacing:1px;margin:0;text-align:center}
        .login-sub{color:#94A3B8;font-size:13px;margin-top:4px;text-align:center;margin-bottom:28px}
        .login-label{font-size:11px;font-weight:700;color:#94A3B8;text-transform:uppercase;display:block;margin-bottom:6px;letter-spacing:.5px}
        .login-input{width:100%;padding:12px 14px;border-radius:12px;border:1.5px solid #E5E7EB;margin-bottom:16px;outline:none;font-size:14px;font-family:inherit;box-sizing:border-box;color:#111827;transition:border-color .15s}
        .login-input:focus{border-color:#BE1522}
        .login-input-last{margin-bottom:24px}
        .login-btn{width:100%;background:#BE1522;color:#fff;padding:14px;border-radius:12px;border:none;font-weight:700;cursor:pointer;font-size:15px;font-family:inherit;transition:opacity .15s}
        .login-btn:hover{opacity:.88}
        @media(max-width:480px){.login-card{padding:28px 20px;border-radius:16px}.login-logo{width:60px;height:60px;font-size:22px}.login-title{font-size:20px}.login-btn{font-size:14px;padding:13px}}
      `}</style>
      <div className="login-page">
        <form onSubmit={manejarLogin} className="login-card">
          <div className="login-logo">
            <img src="/logo.png" alt="ASEDUIS" onError={(e)=>{ const t=e.target as HTMLImageElement; t.style.display='none'; t.parentElement!.textContent='A' }} />
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

  // ── Main dashboard ────────────────────────────────────────
  return (
    <div className="layout">
      <style>{CSS}</style>

      {/* Mobile backdrop */}
      <div className={`sidebar-backdrop ${sidebarOpen?'show':''}`} onClick={()=>setSidebarOpen(false)} />

      {/* ── SIDEBAR ── */}
      <aside className={`sidebar ${sidebarOpen?'open':''}`}>
        <div style={{ display:'flex', alignItems:'center', gap:12, padding:'22px 20px', borderBottom:'1px solid rgba(255,255,255,.15)' }}>
          <div style={{ width:40, height:40, background:'#fff', borderRadius:'50%', display:'flex', alignItems:'center', justifyContent:'center', flexShrink:0, overflow:'hidden' }}>
            <img src="/logo_black.png" alt="Logo ASEDUIS" style={{ width:'80%', height:'80%', objectFit:'contain' }} />
          </div>
          <div>
            <div style={{ color:'#fff', fontWeight:800, fontSize:15, letterSpacing:1 }}>ASEDUIS</div>
            <div style={{ color:'rgba(255,255,255,.55)', fontSize:11 }}>Panel administrativo</div>
          </div>
        </div>
        <div style={{ padding:'14px 12px 8px' }}>
          <div style={{ fontSize:10, fontWeight:700, color:'rgba(255,255,255,.4)', textTransform:'uppercase', letterSpacing:.6, padding:'0 8px', marginBottom:8 }}>Principal</div>
          {([
            { k:'dashboard',    e:'📊', l:'Dashboard'    },
            { k:'egresados',    e:'👥', l:'Egresados'    },
            { k:'secretarios',  e:'🔑', l:'Secretarios'  },
            { k:'validaciones', e:'📷', l:'Validaciones' },
          ] as {k:Tab;e:string;l:string}[]).map(n=>(
            <div key={n.k} className={`nav-item ${tab===n.k?'active':''}`} onClick={()=>navTo(n.k)}>
              <span style={{ fontSize:18 }}>{n.e}</span>
              <span className="nav-lbl">{n.l}</span>
            </div>
          ))}
        </div>
        <div style={{ marginTop:'auto', padding:16, borderTop:'1px solid rgba(255,255,255,.15)' }}>
          <div style={{ color:'#fff', fontSize:13, fontWeight:600, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap', marginBottom:2 }}>{sesion.user.email}</div>
          <div style={{ color:'rgba(255,255,255,.5)', fontSize:11, marginBottom:10 }}>Administrador</div>
          <button onClick={manejarLogout} style={{ display:'flex', alignItems:'center', gap:8, padding:'8px 12px', borderRadius:8, background:'rgba(255,255,255,.1)', border:'none', color:'#fff', cursor:'pointer', width:'100%', fontSize:13, fontFamily:'inherit' }}>
            <span>⬅</span><span>Cerrar sesión</span>
          </button>
        </div>
      </aside>

      {/* ── MAIN ── */}
      <div className="main">
        <div className="topbar">
          <div style={{ display:'flex', alignItems:'center', gap:10 }}>
            <button className="menu-btn" onClick={()=>setSidebarOpen(o=>!o)}>☰</button>
            <div>
              <div style={{ fontSize:18, fontWeight:700, color:'#BE1522', lineHeight:1.2 }}>
                {{ dashboard:'Dashboard', egresados:'Egresados', secretarios:'Secretarios', validaciones:'Validaciones' }[tab]}
              </div>
              <div style={{ fontSize:11, color:'#94A3B8', marginTop:2 }}>Sincronizado · {fechaHoy()}</div>
            </div>
          </div>
          <div className="topbar-actions">
            {tab==='egresados' && <>
              <button className="btn btn-green btn-sm" onClick={()=>exportarExcel(egresadosFiltrados, nombreArchivo('egresados'))}>📊 Excel</button>
              <button className="btn btn-outline btn-sm" onClick={()=>exportarPDF(egresadosFiltrados, nombreArchivo('egresados'))}>📄 PDF</button>
              <button className="btn btn-outline btn-sm" onClick={()=>{ setModalImport(true); setImportRows([]); setImportStatus([]); setImportMsg([]); setImportDone(false) }}>📥 Importar</button>
              <button className="btn btn-primary btn-sm" onClick={()=>setModalNuevoEg(true)}>+ Nuevo</button>
            </>}
            {tab==='secretarios'  && <button className="btn btn-primary btn-sm" onClick={()=>setModalNuevoSec(true)}>+ Nuevo secretario</button>}
            {tab==='dashboard'    && <>
              <button className="btn btn-outline btn-sm" onClick={()=>{ setModalImport(true); setImportRows([]); setImportStatus([]); setImportMsg([]); setImportDone(false) }}>📥 Importar Excel</button>
              <button className="btn btn-primary btn-sm" onClick={()=>setModalNuevoEg(true)}>+ Nuevo egresado</button>
            </>}
            {tab==='validaciones' && <button className="btn btn-green btn-sm" onClick={()=>exportarExcel(validaciones, nombreArchivo('validaciones'))}>📊 Excel</button>}
            <button className="btn btn-outline btn-sm" onClick={cargarTodo} title="Actualizar">🔄</button>
          </div>
        </div>

        <div className="content">

          {/* ═══ DASHBOARD ═══ */}
          {tab==='dashboard' && <>
            <div className="stats-grid">
              <div className="stat-card" style={{ borderLeftColor:'#16A34A' }}>
                <div style={{ fontSize:11, color:'#94A3B8', fontWeight:700, textTransform:'uppercase', marginBottom:6 }}>Activos</div>
                <div style={{ fontSize:32, fontWeight:800, color:'#16A34A' }}>{stats.activos}</div>
                <div style={{ fontSize:11, color:'#94A3B8', marginTop:4 }}>carnets vigentes</div>
              </div>
              <div className="stat-card" style={{ borderLeftColor:'#D97706' }}>
                <div style={{ fontSize:11, color:'#94A3B8', fontWeight:700, textTransform:'uppercase', marginBottom:6 }}>Vencidos</div>
                <div style={{ fontSize:32, fontWeight:800, color:'#D97706' }}>{stats.vencidos}</div>
                <div style={{ fontSize:11, color:'#94A3B8', marginTop:4 }}>requieren renovación</div>
              </div>
              <div className="stat-card" style={{ borderLeftColor:'#94A3B8' }}>
                <div style={{ fontSize:11, color:'#94A3B8', fontWeight:700, textTransform:'uppercase', marginBottom:6 }}>Inactivos</div>
                <div style={{ fontSize:32, fontWeight:800, color:'#94A3B8' }}>{stats.inactivos}</div>
                <div style={{ fontSize:11, color:'#94A3B8', marginTop:4 }}>suspendidos</div>
              </div>
              <div className="stat-card" style={{ borderLeftColor:'#BE1522' }}>
                <div style={{ fontSize:11, color:'#94A3B8', fontWeight:700, textTransform:'uppercase', marginBottom:6 }}>Total egresados</div>
                <div style={{ fontSize:32, fontWeight:800, color:'#BE1522' }}>{stats.total}</div>
                <div style={{ fontSize:11, color:'#94A3B8', marginTop:4 }}>{secretarios.length} secretarios</div>
              </div>
            </div>

            {valHoy.length > 0 && (
              <div className="val-hoy-box">
                <div style={{ fontSize:13, fontWeight:700, color:'#BE1522', flexShrink:0 }}>📷 Hoy: {valHoy.length} validacion{valHoy.length!==1?'es':''}</div>
                {valHoy.slice(0,3).map((v:any) => (
                  <div key={v.id} style={{ display:'flex', alignItems:'center', gap:6, background:'#fff', borderRadius:8, padding:'6px 10px', fontSize:12 }}>
                    <span>{RESULTADO_EMOJI[v.resultado]??'?'}</span>
                    <span style={{ fontWeight:600 }}>{v.egresado_nombre??'Desconocido'}</span>
                    <span style={{ color:'#94A3B8' }}>{v.hora?.slice(0,5)}</span>
                  </div>
                ))}
                <button className="btn btn-outline btn-sm" style={{ marginLeft:'auto' }} onClick={()=>setTab('validaciones')}>Ver todas →</button>
              </div>
            )}

            <div className="export-bar">
              <span>Exportar:</span>
              <button className="btn btn-green btn-sm" onClick={()=>exportarExcel(egresados, nombreArchivo('todos'))}>📊 Todos</button>
              <button className="btn btn-green btn-sm" onClick={()=>exportarExcel(egresados.filter(e=>e.estado==='activo'), nombreArchivo('activos'))}>📊 Activos</button>
              <button className="btn btn-green btn-sm" onClick={()=>exportarExcel(egresados.filter(e=>e.estado==='vencido'), nombreArchivo('vencidos'))}>📊 Vencidos</button>
              <button className="btn btn-outline btn-sm" onClick={()=>exportarPDF(egresados, nombreArchivo('todos'))}>📄 PDF completo</button>
            </div>

            <div className="card">
              <div className="card-hdr">
                <span style={{ fontWeight:700, fontSize:14 }}>Egresados recientes</span>
                <button className="btn btn-outline btn-sm" onClick={()=>setTab('egresados')}>Ver todos →</button>
              </div>
              <div style={{ overflowX:'auto' }}>
                <table>
                  <thead><tr><th>Nombre</th><th>Cédula</th><th>Estado</th><th>Vencimiento</th><th></th></tr></thead>
                  <tbody>
                    {egresados.slice(0,6).map(eg=>(
                      <tr key={eg.id}>
                        <td><strong>{eg.nombre_completo}</strong><br/><span style={{fontSize:11,color:'#94A3B8'}}>{eg.email}</span></td>
                        <td>{eg.cedula}</td>
                        <td><span className={ESTADO_BADGE[eg.estado]??'badge inactive'}><span className="badge-dot"></span>{eg.estado}</span></td>
                        <td style={{color:eg.estado==='vencido'?'#D97706':'inherit',fontWeight:eg.estado==='vencido'?700:400}}>{eg.fecha_vencimiento||'—'}</td>
                        <td><button className="action-btn" onClick={()=>abrirDetalle(eg)}>👁</button></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </>}

          {/* ═══ EGRESADOS ═══ */}
          {tab==='egresados' && <>
            <div className="search-row">
              <div className="search-box">
                <span>🔍</span>
                <input placeholder="Buscar por nombre, cédula, email o empresa..." value={busqueda} onChange={e=>setBusqueda(e.target.value)} />
                {busqueda && <button onClick={()=>setBusqueda('')} style={{background:'none',border:'none',cursor:'pointer',color:'#94A3B8',fontSize:16}}>✕</button>}
              </div>
              {(['todos','activo','vencido','inactivo'] as FiltroEstado[]).map(f=>(
                <button key={f} className={`tab-btn ${filtroEstado===f?'on':''}`} onClick={()=>setFiltroEstado(f)}>
                  {f==='todos'?`Todos (${egresados.length})`:f==='activo'?`Activos (${stats.activos})`:f==='vencido'?`Vencidos (${stats.vencidos})`:`Inactivos (${stats.inactivos})`}
                </button>
              ))}
            </div>
            <div className="card">
              <div className="card-hdr">
                <span style={{fontWeight:700,fontSize:14}}>Base de datos ({egresadosFiltrados.length})</span>
                <span style={{fontSize:12,color:'#94A3B8'}}>{stats.total} total</span>
              </div>
              <div style={{overflowX:'auto'}}>
                <table>
                  <thead><tr><th>Nombre</th><th>Cédula</th><th>Email</th><th>Empresa</th><th>Estado</th><th>Vencimiento</th><th>Acciones</th></tr></thead>
                  <tbody>
                    {egresadosFiltrados.length===0
                      ? <tr><td colSpan={7} style={{textAlign:'center',padding:32,color:'#94A3B8'}}>No se encontraron egresados</td></tr>
                      : egresadosFiltrados.map(eg=>(
                        <tr key={eg.id}>
                          <td>
                            <div style={{display:'flex',alignItems:'center',gap:10}}>
                              <div style={{width:34,height:34,borderRadius:'50%',background:'#FDE8EA',display:'flex',alignItems:'center',justifyContent:'center',fontWeight:700,fontSize:12,color:'#BE1522',flexShrink:0,overflow:'hidden'}}>
                                {eg.foto_perfil?<img src={eg.foto_perfil} style={{width:'100%',height:'100%',objectFit:'cover'}} alt=""/>:initials(eg.nombre_completo)}
                              </div>
                              <div><strong>{eg.nombre_completo}</strong><br/><span style={{fontSize:11,color:'#94A3B8'}}>{eg.cedula}</span></div>
                            </div>
                          </td>
                          <td style={{fontSize:12,color:'#6B7280'}}>{eg.cedula}</td>
                          <td style={{fontSize:12,color:'#6B7280'}}>{eg.email}</td>
                          <td style={{fontSize:12}}>{eg.empresa||'—'}</td>
                          <td><span className={ESTADO_BADGE[eg.estado]??'badge inactive'}><span className="badge-dot"></span>{eg.estado}</span></td>
                          <td style={{color:eg.estado==='vencido'?'#D97706':'inherit',fontWeight:eg.estado==='vencido'?700:400}}>{eg.fecha_vencimiento||'—'}</td>
                          <td>
                            <div style={{display:'flex',gap:2}}>
                              <button className="action-btn" title="Ver detalle" onClick={()=>abrirDetalle(eg)}>👁</button>
                              {eg.estado!=='inactivo'&&<button className="action-btn" title="Desactivar" onClick={()=>cambiarEstado(eg,'inactivo')}>⏸</button>}
                              {eg.estado==='inactivo'&&<button className="action-btn" title="Activar" onClick={()=>cambiarEstado(eg,'activo')}>▶</button>}
                            </div>
                          </td>
                        </tr>
                      ))
                    }
                  </tbody>
                </table>
              </div>
            </div>
          </>}

          {/* ═══ SECRETARIOS ═══ */}
          {tab==='secretarios' && <>
            <div className="search-row">
              <div className="search-box">
                <span>🔍</span>
                <input placeholder="Buscar secretario..." value={busquedaSec} onChange={e=>setBusquedaSec(e.target.value)} />
                {busquedaSec && <button onClick={()=>setBusquedaSec('')} style={{background:'none',border:'none',cursor:'pointer',color:'#94A3B8',fontSize:16}}>✕</button>}
              </div>
            </div>
            <div className="card" style={{maxWidth:760}}>
              <div className="card-hdr">
                <span style={{fontWeight:700,fontSize:14}}>Secretarios ({secretariosFiltrados.length})</span>
                <span style={{fontSize:12,color:'#94A3B8'}}>{secretarios.filter(s=>s.activo).length} activos · {secretarios.filter(s=>!s.activo).length} inactivos</span>
              </div>
              <div style={{overflowX:'auto'}}>
                <table>
                  <thead><tr><th>Nombre</th><th>Cédula</th><th>Email</th><th>Registro</th><th>Estado</th></tr></thead>
                  <tbody>
                    {secretariosFiltrados.length===0
                      ? <tr><td colSpan={5} style={{textAlign:'center',padding:32,color:'#94A3B8'}}>No hay secretarios registrados</td></tr>
                      : secretariosFiltrados.map(sec=>(
                        <tr key={sec.id}>
                          <td>
                            <div style={{display:'flex',alignItems:'center',gap:10}}>
                              <div style={{width:34,height:34,borderRadius:'50%',background:sec.activo?'#FDE8EA':'#F3F4F6',display:'flex',alignItems:'center',justifyContent:'center',fontWeight:700,fontSize:12,color:sec.activo?'#BE1522':'#9CA3AF',flexShrink:0}}>
                                {initials(sec.nombre_completo)}
                              </div>
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
                      ))
                    }
                  </tbody>
                </table>
              </div>
            </div>
          </>}

          {/* ═══ VALIDACIONES ═══ */}
          {tab==='validaciones' && <>
            <div style={{display:'flex',gap:8,marginBottom:14,flexWrap:'wrap',alignItems:'center'}}>
              {(['hoy','semana','mes','todo'] as FiltroFecha[]).map(f=>(
                <button key={f} className={`tab-btn ${filtroFecha===f?'on':''}`} onClick={()=>setFiltroFecha(f)}>
                  {f==='hoy'?'Hoy':f==='semana'?'Semana':f==='mes'?'Mes':'Todo'}
                </button>
              ))}
              <span style={{marginLeft:'auto',fontSize:12,color:'#94A3B8'}}>{validaciones.length} registros</span>
            </div>
            <div className="card">
              <div style={{overflowX:'auto'}}>
                <table>
                  <thead><tr><th>Egresado</th><th>Cédula</th><th>Secretario</th><th>Resultado</th><th>Fecha</th><th>Hora</th></tr></thead>
                  <tbody>
                    {validaciones.length===0
                      ? <tr><td colSpan={6} style={{textAlign:'center',padding:32,color:'#94A3B8'}}>Sin validaciones en este período</td></tr>
                      : validaciones.map((v:any)=>(
                        <tr key={v.id}>
                          <td><div style={{display:'flex',alignItems:'center',gap:8}}><span style={{fontSize:18}}>{RESULTADO_EMOJI[v.resultado]??'?'}</span><span style={{fontWeight:600}}>{v.egresado_nombre??'Desconocido'}</span></div></td>
                          <td style={{fontSize:12,color:'#6B7280'}}>{v.egresado_cedula??'—'}</td>
                          <td style={{fontSize:12,color:'#6B7280'}}>{v.secretario_nombre??'Admin'}</td>
                          <td><span className={RESULTADO_BADGE[v.resultado]??'badge inactive'}><span className="badge-dot"></span>{v.resultado}</span></td>
                          <td style={{fontSize:12,color:'#94A3B8'}}>{v.fecha}</td>
                          <td style={{fontWeight:700}}>{v.hora?.slice(0,8)}</td>
                        </tr>
                      ))
                    }
                  </tbody>
                </table>
              </div>
            </div>
          </>}

        </div>
      </div>

      {/* ═══ MODAL IMPORTAR EXCEL ═══ */}
      <div className={`overlay ${modalImport?'show':''}`} onClick={e=>{if((e.target as any).classList?.contains('overlay')&&!importando)setModalImport(false)}}>
        <div className="modal" style={{maxWidth:560}}>
          <div className="modal-title">📥 Importar egresados desde Excel</div>

          {importRows.length === 0 ? <>
            {/* Drop zone */}
            <div className="warn-box" style={{marginBottom:12}}>
              El archivo debe tener columnas: <strong>Cédula</strong>, <strong>Nombre completo</strong>, <strong>Correo electrónico</strong>. La columna de teléfono es opcional.
            </div>
            <div
              className="drop-zone"
              onDragOver={e=>{e.preventDefault();(e.currentTarget as HTMLElement).classList.add('drag')}}
              onDragLeave={e=>(e.currentTarget as HTMLElement).classList.remove('drag')}
              onDrop={e=>{
                e.preventDefault();(e.currentTarget as HTMLElement).classList.remove('drag')
                const f=e.dataTransfer.files[0]
                if(f) leerArchivoExcel(f)
              }}
            >
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                ref={fileInputRef}
                onChange={e=>{ const f=e.target.files?.[0]; if(f) leerArchivoExcel(f) }}
              />
              <div style={{fontSize:32,marginBottom:8}}>📂</div>
              <div style={{fontWeight:700,fontSize:14,color:'#111827',marginBottom:4}}>Arrastra el archivo aquí</div>
              <div style={{fontSize:12,color:'#94A3B8'}}>o haz clic para seleccionar · .xlsx .xls .csv</div>
            </div>
          </> : <>
            {/* Preview & results */}
            {!importDone && (
              <div className="auto-box" style={{marginBottom:10}}>
                Se encontraron <strong>{importRows.length}</strong> filas válidas. Revisa y luego haz clic en <strong>Importar</strong>.
              </div>
            )}
            {importDone && (
              <div style={{display:'flex',gap:8,marginBottom:10,flexWrap:'wrap'}}>
                {importOk>0 && <span className="badge active">✅ {importOk} creados</span>}
                {importDup>0 && <span className="badge warn">⚠️ {importDup} duplicados</span>}
                {importErr>0 && <span className="badge inactive">❌ {importErr} errores</span>}
              </div>
            )}
            <div style={{overflowX:'auto',border:'1px solid #E5E7EB',borderRadius:10}}>
              <table className="import-table">
                <thead>
                  <tr>
                    <th>#</th><th>Cédula</th><th>Nombre</th><th>Email</th>
                    {importRows.some(r=>r.telefono) && <th>Teléfono</th>}
                    {importStatus.some(s=>s!=='idle') && <th>Estado</th>}
                  </tr>
                </thead>
                <tbody>
                  {importRows.map((r,i)=>(
                    <tr key={i} style={{background:importStatus[i]==='ok'?'#F0FDF4':importStatus[i]==='err'?'#FEF2F2':importStatus[i]==='dup'?'#FFFBEB':''}}>
                      <td style={{color:'#94A3B8'}}>{i+1}</td>
                      <td>{r.cedula}</td>
                      <td style={{fontWeight:600}}>{r.nombre}</td>
                      <td style={{fontSize:11,color:'#6B7280'}}>{r.email}</td>
                      {importRows.some(r=>r.telefono) && <td style={{fontSize:11}}>{r.telefono||'—'}</td>}
                      {importStatus.some(s=>s!=='idle') && (
                        <td>
                          {importStatus[i]==='idle' && <span style={{color:'#94A3B8',fontSize:11}}>Pendiente</span>}
                          {importStatus[i]==='ok'   && <span style={{color:'#16A34A',fontWeight:700,fontSize:11}}>✓ Creado</span>}
                          {importStatus[i]==='dup'  && <span style={{color:'#D97706',fontWeight:700,fontSize:11}}>⚠ Duplicado</span>}
                          {importStatus[i]==='err'  && <span style={{color:'#DC2626',fontWeight:700,fontSize:11}} title={importMsg[i]}>✕ Error</span>}
                        </td>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>}

          <div style={{display:'flex',gap:10,marginTop:16,flexWrap:'wrap'}}>
            <button
              className="btn btn-outline"
              style={{flex:1,justifyContent:'center'}}
              disabled={importando}
              onClick={()=>{
                if(!importando){ setImportRows([]); setImportStatus([]); setImportMsg([]); setImportDone(false)
                  if(importDone) setModalImport(false) }
              }}
            >
              {importDone ? 'Cerrar' : importRows.length ? 'Cambiar archivo' : 'Cancelar'}
            </button>
            {importRows.length > 0 && !importDone && (
              <button
                className="btn btn-primary"
                style={{flex:1,justifyContent:'center'}}
                onClick={ejecutarImportacion}
                disabled={importando}
              >
                {importando ? `Importando ${importStatus.filter(s=>s!=='idle').length}/${importRows.length}...` : `✓ Importar ${importRows.length} egresados`}
              </button>
            )}
          </div>
        </div>
      </div>

      {/* ═══ MODAL DETALLE EGRESADO ═══ */}
      <div className={`overlay ${modalDetalle?'show':''}`}>
        <div className="modal" style={{maxWidth:520}}>
          {egSel && <>
            <div style={{display:'flex',alignItems:'center',gap:14,marginBottom:16}}>
              <div style={{width:60,height:60,borderRadius:'50%',background:'#FDE8EA',display:'flex',alignItems:'center',justifyContent:'center',fontSize:22,fontWeight:700,color:'#BE1522',overflow:'hidden',flexShrink:0}}>
                {egSel.foto_perfil?<img src={egSel.foto_perfil} style={{width:'100%',height:'100%',objectFit:'cover'}} alt=""/>:initials(egSel.nombre_completo)}
              </div>
              <div>
                <div style={{fontSize:17,fontWeight:700}}>{egSel.nombre_completo}</div>
                <div style={{fontSize:13,color:'#94A3B8',marginTop:2}}>CC {egSel.cedula}</div>
                <span className={ESTADO_BADGE[egSel.estado]??'badge inactive'} style={{marginTop:6,display:'inline-flex'}}><span className="badge-dot"></span>{egSel.estado}</span>
              </div>
            </div>

            <div className="detail-grid">
              <div className="detail-section">Datos personales</div>
              {[
                {k:'Teléfono',         v:egSel.telefono??'—'},
                {k:'Fecha nacimiento', v:egSel.fecha_nacimiento??'—'},
                {k:'Ciudad',           v:egSel.ciudad_nacimiento??'—'},
                {k:'Dirección',        v:egSel.direccion??'—'},
              ].map(r=>(
                <div key={r.k} className="detail-row">
                  <span className="detail-key">{r.k}</span>
                  <span className="detail-val">{r.v}</span>
                </div>
              ))}

              <div style={{marginTop:10,marginBottom:4}}>
                <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:6}}>
                  <span style={{fontSize:10,fontWeight:700,color:'#BE1522',textTransform:'uppercase',letterSpacing:.5}}>Email</span>
                  {!editandoEmail
                    ? <button className="btn btn-outline btn-sm" onClick={()=>{setEditandoEmail(true);setNuevoEmail(egSel.email??'')}}>✏️ Editar</button>
                    : <button className="btn btn-sm" style={{background:'#F9FAFB',color:'#6B7280',border:'1px solid #E5E7EB'}} onClick={()=>setEditandoEmail(false)}>Cancelar</button>
                  }
                </div>
                {!editandoEmail
                  ? <div style={{fontSize:13,color:'#111827',padding:'8px 12px',background:'#F9F1F1',borderRadius:8}}>{egSel.email??'—'}</div>
                  : <div style={{display:'flex',gap:8,alignItems:'center'}}>
                      <input className="form-inp" type="email" value={nuevoEmail} onChange={e=>setNuevoEmail(e.target.value.toLowerCase())} placeholder="nuevo@correo.com" style={{flex:1,marginTop:0}} />
                      <button className="btn btn-primary btn-sm" onClick={()=>actualizarEmail(egSel, nuevoEmail)} disabled={savingEmail || nuevoEmail===egSel.email} style={{flexShrink:0}}>
                        {savingEmail ? '...' : '✓ Guardar'}
                      </button>
                    </div>
                }
              </div>

              <div className="detail-section">Carnet</div>
              {[
                {k:'Fecha expedición', v:egSel.fecha_expedicion??'—'},
                {k:'Vencimiento',      v:egSel.fecha_vencimiento??'—', highlight:true},
              ].map(r=>(
                <div key={r.k} className="detail-row">
                  <span className="detail-key">{r.k}</span>
                  <span className="detail-val" style={(r as any).highlight?{color:'#BE1522',fontWeight:700}:{}}>{r.v}</span>
                </div>
              ))}

              {(egSel.empresa || egSel.cargo) && <>
                <div className="detail-section">Laboral</div>
                {egSel.empresa&&<div className="detail-row"><span className="detail-key">Empresa</span><span className="detail-val">{egSel.empresa}</span></div>}
                {egSel.cargo  &&<div className="detail-row"><span className="detail-key">Cargo</span><span className="detail-val">{egSel.cargo}</span></div>}
              </>}

              {(egSel.titulo_pregrado || egSel.titulo_posgrado) && <>
                <div className="detail-section">Académico</div>
                {egSel.titulo_pregrado&&<div className="detail-row"><span className="detail-key">Pregrado</span><span className="detail-val">{egSel.titulo_pregrado}</span></div>}
                {egSel.institucion_pregrado&&<div className="detail-row"><span className="detail-key">Institución</span><span className="detail-val">{egSel.institucion_pregrado}</span></div>}
                {egSel.fecha_grado_pregrado&&<div className="detail-row"><span className="detail-key">Fecha grado</span><span className="detail-val">{egSel.fecha_grado_pregrado}</span></div>}
                {egSel.titulo_posgrado&&<div className="detail-row"><span className="detail-key">Posgrado</span><span className="detail-val">{egSel.titulo_posgrado}</span></div>}
                {egSel.institucion_posgrado&&<div className="detail-row"><span className="detail-key">Institución</span><span className="detail-val">{egSel.institucion_posgrado}</span></div>}
              </>}
            </div>

            <div style={{marginBottom:12}}>
              <label className="form-lbl">Fecha de expedición para renovar</label>
              <input type="date" id="fechaRenovar" defaultValue={fechaHoy()} className="form-inp" style={{marginTop:4}} />
              <div style={{fontSize:11,color:'#94A3B8',marginTop:4}}>Se calculan 365 días automáticamente</div>
            </div>

            <div className="actions-row">
              <button className="btn btn-primary" onClick={()=>{const i=document.getElementById('fechaRenovar') as HTMLInputElement;renovarCarnet(egSel,i?.value||fechaHoy())}}>🔄 Renovar (365d)</button>
              {egSel.estado!=='activo'&&<button className="btn btn-success" onClick={()=>cambiarEstado(egSel,'activo')}>▶ Activar</button>}
              {egSel.estado!=='inactivo'&&<button className="btn btn-outline" onClick={()=>cambiarEstado(egSel,'inactivo')}>⏸ Desactivar</button>}
              <button className="btn btn-danger" onClick={()=>eliminarEgresado(egSel)}>🗑️ Eliminar</button>
            </div>
            <button className="btn" style={{width:'100%',justifyContent:'center',background:'#F9FAFB',color:'#6B7280',border:'1px solid #E5E7EB',marginTop:4}} onClick={()=>setModalDetalle(false)}>Cerrar</button>
          </>}
        </div>
      </div>

      {/* ═══ MODAL NUEVO EGRESADO ═══ */}
      <div className={`overlay ${modalNuevoEg?'show':''}`} onClick={e=>{if((e.target as any).className?.includes?.('overlay'))setModalNuevoEg(false)}}>
        <div className="modal">
          <div className="modal-title">👤 Nuevo egresado</div>
          <div className="auto-box">✅ El acceso se crea automáticamente. Contraseña inicial = cédula.</div>
          <div className="form-row2">
            <div><label className="form-lbl">Cédula *</label><input className="form-inp" placeholder="Sin puntos" value={formEg.cedula} onChange={e=>setFormEg(p=>({...p,cedula:e.target.value}))} /></div>
            <div><label className="form-lbl">Teléfono</label><input className="form-inp" placeholder="+57 300..." value={formEg.telefono} onChange={e=>setFormEg(p=>({...p,telefono:e.target.value}))} /></div>
          </div>
          <label className="form-lbl">Nombre completo *</label>
          <input className="form-inp" placeholder="Nombre y apellidos" value={formEg.nombre} onChange={e=>setFormEg(p=>({...p,nombre:e.target.value}))} />
          <label className="form-lbl">Email *</label>
          <input className="form-inp" type="email" placeholder="correo@dominio.com" value={formEg.email} onChange={e=>setFormEg(p=>({...p,email:e.target.value}))} />
          <div className="form-row2">
            <div><label className="form-lbl">Fecha de expedición</label><input className="form-inp" type="date" value={formEg.expedicion} onChange={e=>onChangeExpedicion(e.target.value)} /></div>
            <div><label className="form-lbl">Vencimiento (365d auto)</label><input className="form-inp" value={formEg.vencimiento} disabled /></div>
          </div>
          <div className="form-row2" style={{marginTop:4}}>
            <div><label className="form-lbl">Estado inicial</label>
              <select className="form-inp" value={formEg.estado} onChange={e=>setFormEg(p=>({...p,estado:e.target.value}))}>
                <option value="activo">Activo</option><option value="inactivo">Inactivo</option>
              </select>
            </div>
          </div>
          <div style={{display:'flex',gap:10,marginTop:20}}>
            <button className="btn btn-outline" style={{flex:1,justifyContent:'center'}} onClick={()=>setModalNuevoEg(false)}>Cancelar</button>
            <button className="btn btn-primary" style={{flex:1,justifyContent:'center'}} onClick={crearEgresado} disabled={saving}>{saving?'Creando...':'✓ Crear egresado'}</button>
          </div>
        </div>
      </div>

      {/* ═══ MODAL NUEVO SECRETARIO ═══ */}
      <div className={`overlay ${modalNuevoSec?'show':''}`} onClick={e=>{if((e.target as any).className?.includes?.('overlay'))setModalNuevoSec(false)}}>
        <div className="modal">
          <div className="modal-title">🔑 Nuevo secretario</div>
          <div className="auto-box">✅ El acceso se crea automáticamente. Contraseña inicial = cédula.</div>
          <label className="form-lbl">Cédula *</label>
          <input className="form-inp" placeholder="Sin puntos ni espacios" value={formSec.cedula} onChange={e=>setFormSec(p=>({...p,cedula:e.target.value}))} />
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