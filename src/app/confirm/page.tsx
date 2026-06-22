'use client'
import { useEffect, useState } from 'react'
import { supabase } from '../../lib/supabase'

export default function ConfirmPage() {
  const [estado, setEstado] = useState<'verificando' | 'ok' | 'error'>('verificando')

  useEffect(() => {
    async function procesar() {
      const hash   = window.location.hash
      const params = new URLSearchParams(hash.replace('#', ''))
      const type   = params.get('type')
      const token  = params.get('access_token')
      const refresh = params.get('refresh_token')

      if (!token) { setEstado('error'); return }

      const { error } = await supabase.auth.setSession({
        access_token:  token,
        refresh_token: refresh ?? '',
      })

      setEstado(error ? 'error' : 'ok')
      window.history.replaceState(null, '', window.location.pathname)
    }
    procesar()
  }, [])

  return (
    <div style={{ display:'flex', minHeight:'100vh', alignItems:'center', justifyContent:'center', background:'#BE1522', fontFamily:'"Segoe UI",sans-serif', padding:16 }}>
      <div style={{ background:'#fff', borderRadius:20, padding:40, width:'100%', maxWidth:400, textAlign:'center', boxShadow:'0 8px 40px rgba(0,0,0,.18)' }}>
        <div style={{ width:72, height:72, background:'#BE1522', borderRadius:'50%', display:'flex', alignItems:'center', justifyContent:'center', margin:'0 auto 16px', fontSize:28, fontWeight:900, color:'#fff' }}>A</div>
        <h2 style={{ color:'#BE1522', fontWeight:800, fontSize:20, marginBottom:8 }}>ASEDUIS</h2>

        {estado === 'verificando' && (
          <p style={{ color:'#94A3B8', fontSize:14 }}>⏳ Verificando tu correo...</p>
        )}
        {estado === 'ok' && (
          <div style={{ background:'#DCFCE7', border:'1px solid #BBF7D0', borderRadius:12, padding:16, color:'#16A34A', fontWeight:600, fontSize:14 }}>
            ✓ Correo confirmado correctamente.<br/>
            <span style={{ fontSize:12, fontWeight:400, display:'block', marginTop:6 }}>
              Ya puedes cerrar esta ventana y volver a la app.
            </span>
          </div>
        )}
        {estado === 'error' && (
          <div style={{ background:'#FEE2E2', border:'1px solid #FECACA', borderRadius:12, padding:16, color:'#DC2626', fontSize:13 }}>
            ❌ Enlace inválido o expirado.<br/>
            <span style={{ fontSize:12, display:'block', marginTop:6 }}>
              Intenta actualizar el email de nuevo desde la app.
            </span>
          </div>
        )}

        <p style={{ fontSize:11, color:'#94A3B8', marginTop:20 }}>App Egresados ASEDUIS · Plataforma institucional</p>
      </div>
    </div>
  )
}