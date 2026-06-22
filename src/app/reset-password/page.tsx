'use client'
// dashboard/src/app/reset-password/page.tsx
import { useState, useEffect } from 'react'
import { useRouter } from 'next/navigation'
import { supabase } from '../../lib/supabase'
 
export default function ResetPasswordPage() {
  const [nueva,       setNueva]       = useState('')
  const [confirmar,   setConfirmar]   = useState('')
  const [loading,     setLoading]     = useState(false)
  const [listo,       setListo]       = useState(false)
  const [error,       setError]       = useState('')
  const [sesionLista, setSesionLista] = useState(false)
  const [verificando, setVerificando] = useState(true)
  const router = useRouter()
 
  useEffect(() => {
    async function procesarHash() {
      const hash = window.location.hash
      if (!hash) {
        setVerificando(false)
        setError('No se encontró el token. Solicita un nuevo enlace desde la app.')
        return
      }
      const params      = new URLSearchParams(hash.replace('#', ''))
      const accessToken  = params.get('access_token')
      const refreshToken = params.get('refresh_token')
      const type         = params.get('type')
      if (type !== 'recovery' || !accessToken) {
        setVerificando(false)
        setError('Enlace inválido o expirado. Solicita uno nuevo desde la app.')
        return
      }
      const { error: sessionError } = await supabase.auth.setSession({
        access_token:  accessToken,
        refresh_token: refreshToken ?? '',
      })
      setVerificando(false)
      if (sessionError) {
        setError('El enlace expiró. Solicita uno nuevo desde la app.')
        return
      }
      setSesionLista(true)
      window.history.replaceState(null, '', window.location.pathname)
    }
    procesarHash()
  }, [])
 
  async function handleReset(e: React.FormEvent) {
    e.preventDefault()
    setError('')
    if (nueva !== confirmar) { setError('Las contraseñas no coinciden.'); return }
    if (nueva.length < 6)    { setError('Mínimo 6 caracteres.'); return }
    setLoading(true)
    const { error: err } = await supabase.auth.updateUser({ password: nueva })
    setLoading(false)
    if (err) { setError('Error: ' + err.message); return }
    setListo(true)
    setTimeout(() => router.push('/'), 3000)
  }
 
  return (
    <>
      <style>{`
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: "Segoe UI", system-ui, sans-serif; }
 
        .page {
          min-height: 100vh;
          display: flex;
          align-items: center;
          justify-content: center;
          background: #BE1522;
          padding: 16px;
        }
 
        .card {
          background: #fff;
          border-radius: 20px;
          padding: 40px;
          width: 100%;
          max-width: 420px;
          box-shadow: 0 8px 40px rgba(0,0,0,.18);
        }
 
        .logo-wrap {
          text-align: center;
          margin-bottom: 28px;
        }
 
        .logo-circle {
          width: 72px;
          height: 72px;
          background: #BE1522;
          border-radius: 50%;
          display: flex;
          align-items: center;
          justify-content: center;
          margin: 0 auto 14px;
          font-size: 28px;
          font-weight: 900;
          color: #fff;
        }
 
        .logo-title {
          color: #BE1522;
          font-weight: 800;
          font-size: 22px;
          letter-spacing: 1px;
        }
 
        .logo-sub {
          color: #94A3B8;
          font-size: 13px;
          margin-top: 4px;
        }
 
        .alert {
          border-radius: 12px;
          padding: 16px;
          text-align: center;
          font-size: 13px;
          line-height: 1.5;
        }
        .alert-warning { background: #FEF3C7; border: 1px solid #FDE68A; color: #D97706; }
        .alert-error   { background: #FEE2E2; border: 1px solid #FECACA; color: #DC2626; }
        .alert-success { background: #DCFCE7; border: 1px solid #BBF7D0; color: #16A34A; font-weight: 600; }
        .alert-success span { font-size: 12px; font-weight: 400; display: block; margin-top: 4px; }
 
        .inline-error {
          background: #FEE2E2;
          border: 1px solid #FECACA;
          border-radius: 10px;
          padding: 10px 14px;
          color: #DC2626;
          font-size: 13px;
          margin: 14px 0;
        }
 
        .form-group { margin-bottom: 16px; }
        .form-group:last-of-type { margin-bottom: 0; }
 
        .form-label {
          font-size: 11px;
          font-weight: 700;
          color: #94A3B8;
          text-transform: uppercase;
          display: block;
          margin-bottom: 6px;
          letter-spacing: .5px;
        }
 
        .form-input {
          width: 100%;
          padding: 12px 14px;
          border-radius: 12px;
          border: 1.5px solid #E5E7EB;
          outline: none;
          font-size: 14px;
          font-family: inherit;
          transition: border-color .15s;
          background: #F9FAFB;
          color: #111827;
        }
        .form-input:focus { border-color: #BE1522; background: #fff; }
 
        .submit-btn {
          width: 100%;
          background: #BE1522;
          color: #fff;
          padding: 14px;
          border-radius: 12px;
          border: none;
          font-weight: 700;
          font-size: 15px;
          font-family: inherit;
          cursor: pointer;
          margin-top: 20px;
          transition: opacity .15s;
        }
        .submit-btn:hover   { opacity: .88; }
        .submit-btn:disabled { opacity: .6; cursor: not-allowed; }
 
        .footer-txt {
          text-align: center;
          font-size: 11px;
          color: #94A3B8;
          margin-top: 20px;
        }
 
        /* ── Responsive móvil ── */
        @media (max-width: 480px) {
          .card {
            padding: 28px 20px;
            border-radius: 16px;
          }
          .logo-circle {
            width: 60px;
            height: 60px;
            font-size: 22px;
          }
          .logo-title { font-size: 20px; }
          .submit-btn { font-size: 14px; padding: 13px; }
        }
 
        @media (max-width: 360px) {
          .card { padding: 24px 16px; }
          .logo-title { font-size: 18px; }
        }
      `}</style>
 
      <div className="page">
        <div className="card">
 
          {/* Logo */}
          <div className="logo-wrap">
            <div className="logo-circle">A</div>
            <div className="logo-title">ASEDUIS</div>
            <div className="logo-sub">Restablecer contraseña</div>
          </div>
 
          {/* Verificando */}
          {verificando && (
            <div className="alert alert-warning">
              ⏳ Verificando enlace...
            </div>
          )}
 
          {/* Error general */}
          {!verificando && error && !listo && (
            <div className="alert alert-error">
              ❌ {error}
            </div>
          )}
 
          {/* Éxito */}
          {listo && (
            <div className="alert alert-success">
              ✓ Contraseña actualizada correctamente.
              <span>Vuelve a la app e ingresa con tu nueva contraseña.</span>
            </div>
          )}
 
          {/* Formulario */}
          {!verificando && sesionLista && !listo && (
            <form onSubmit={handleReset}>
              <div className="form-group">
                <label className="form-label">Nueva contraseña</label>
                <input
                  className="form-input"
                  type="password"
                  required
                  minLength={6}
                  value={nueva}
                  onChange={e => setNueva(e.target.value)}
                  placeholder="Mínimo 6 caracteres"
                />
              </div>
 
              <div className="form-group">
                <label className="form-label">Confirmar contraseña</label>
                <input
                  className="form-input"
                  type="password"
                  required
                  minLength={6}
                  value={confirmar}
                  onChange={e => setConfirmar(e.target.value)}
                  placeholder="Repite la contraseña"
                />
              </div>
 
              {error && (
                <div className="inline-error">{error}</div>
              )}
 
              <button
                type="submit"
                className="submit-btn"
                disabled={loading}
              >
                {loading ? 'Actualizando...' : 'Cambiar contraseña'}
              </button>
            </form>
          )}
 
          <p className="footer-txt">App Egresados ASEDUIS · Plataforma institucional</p>
        </div>
      </div>
    </>
  )
}