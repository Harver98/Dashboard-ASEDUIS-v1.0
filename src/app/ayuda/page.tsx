'use client'
// dashboard/src/app/ayuda/page.tsx

export default function AyudaPage() {
  const FAQS = [
    {
      cat: '🪪 Carnet digital',
      items: [
        {
          q: '¿Cómo funciona el carnet digital?',
          a: 'Tu carnet digital es una identificación virtual que muestra tus datos como egresado de ASEDUIS. Contiene un código QR único que los secretarios escanean para verificar tu vigencia en tiempo real.',
        },
        {
          q: '¿Qué significa que mi carnet está vencido?',
          a: 'El carnet tiene una vigencia de 365 días desde la fecha de expedición. Cuando vence, debes contactar a la institución para renovarlo. Tu información sigue activa pero no podrás ser validado hasta la renovación.',
        },
        {
          q: '¿Puedo usar el carnet sin conexión a internet?',
          a: 'El carnet digital se puede visualizar sin internet, pero la validación por parte del secretario requiere conexión para consultar la base de datos en tiempo real.',
        },
        {
          q: '¿El código QR cambia?',
          a: 'No. Tu código QR es único y permanente. Está vinculado exclusivamente a tu registro como egresado y no cambia aunque actualices tu información de perfil.',
        },
      ],
    },
    {
      cat: '🔒 Cuenta y contraseña',
      items: [
        {
          q: '¿Cuál es mi contraseña inicial?',
          a: 'Al momento de tu registro, tu contraseña inicial es tu número de cédula sin puntos ni espacios. La app te pedirá cambiarla la primera vez que ingreses.',
        },
        {
          q: '¿Cómo cambio mi contraseña?',
          a: 'Ve a Configuración → Actualizar contraseña. Necesitas ingresar tu contraseña actual y luego la nueva. Si olvidaste tu contraseña, usa la opción "¿Olvidaste tu contraseña?" en la pantalla de inicio.',
        },
        {
          q: '¿Cómo recupero mi contraseña si la olvidé?',
          a: 'En la pantalla de inicio, toca "¿Olvidaste tu contraseña?" e ingresa tu número de cédula. Te enviaremos un enlace al correo registrado para restablecerla.',
        },
        {
          q: '¿Con qué datos inicio sesión?',
          a: 'Ingresas con tu número de cédula (sin puntos ni espacios) y tu contraseña. No necesitas recordar tu correo electrónico.',
        },
      ],
    },
    {
      cat: '👤 Perfil',
      items: [
        {
          q: '¿Puedo cambiar mi foto de perfil?',
          a: 'Sí. Ve a la pestaña Perfil en la app y toca tu foto o el ícono de cámara. Puedes tomar una foto nueva o elegir una de tu galería.',
        },
        {
          q: '¿Puedo editar mi cédula o correo?',
          a: 'No. La cédula y el correo son datos institucionales que solo el administrador puede modificar. Si necesitas actualizarlos, contacta a soporte.',
        },
        {
          q: '¿Mis datos son visibles para todos?',
          a: 'Tu información básica (nombre y estado del carnet) es visible para los secretarios al momento de validar tu QR. El resto de tu perfil es privado.',
        },
      ],
    },
    {
      cat: '📷 Validación QR',
      items: [
        {
          q: '¿Cómo me validan el carnet?',
          a: 'Muestra tu código QR desde la pestaña "Mi QR" en la app. El secretario lo escanea y verá en tiempo real si tu carnet es válido, está vencido o inactivo.',
        },
        {
          q: '¿Qué pasa si el secretario dice que mi carnet no es válido?',
          a: 'Verifica el estado de tu carnet en la app. Si aparece como "Vencido" o "Inactivo", contacta a la institución para gestionar la renovación o reactivación.',
        },
      ],
    },
    {
      cat: '📱 Problemas técnicos',
      items: [
        {
          q: '¿La app no abre o se cierra inesperadamente?',
          a: 'Cierra la app completamente y vuelve a abrirla. Si el problema persiste, verifica que tengas la última versión instalada desde la Play Store o App Store.',
        },
        {
          q: '¿No puedo iniciar sesión aunque sé mi contraseña?',
          a: 'Verifica que estés ingresando la cédula sin puntos, comas ni espacios. Si olvidaste la contraseña, usa la opción de recuperación en la pantalla de inicio.',
        },
        {
          q: '¿La foto de perfil no se actualiza?',
          a: 'Asegúrate de tener conexión a internet estable al subir la foto. Si el problema persiste, cierra sesión, vuelve a ingresar e intenta de nuevo.',
        },
      ],
    },
  ]

  return (
    <>
      <style>{`
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: "Segoe UI", system-ui, sans-serif; background: #F9F1F1; color: #111827; }
        .page { min-height: 100vh; }
        .header { background: #BE1522; padding: 32px 24px 40px; }
        .header-inner { max-width: 720px; margin: 0 auto; }
        .logo { width: 52px; height: 52px; background: rgba(255,255,255,.2); border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 22px; font-weight: 900; color: #fff; margin-bottom: 16px; }
        .header h1 { color: #fff; font-size: 26px; font-weight: 800; margin-bottom: 6px; }
        .header p { color: rgba(255,255,255,.75); font-size: 14px; line-height: 1.5; }
        .content { max-width: 720px; margin: 0 auto; padding: 32px 24px 48px; }
        .search-wrap { background: #fff; border-radius: 14px; border: 1px solid #F0D4D6; padding: 14px 18px; display: flex; align-items: center; gap: 10px; margin-bottom: 28px; box-shadow: 0 2px 8px rgba(190,21,34,.06); }
        .search-wrap span { font-size: 18px; }
        .search-wrap input { border: none; outline: none; font-size: 14px; color: #111; flex: 1; font-family: inherit; background: transparent; }
        .cat-section { margin-bottom: 28px; }
        .cat-title { font-size: 13px; font-weight: 700; color: #BE1522; text-transform: uppercase; letter-spacing: .5px; margin-bottom: 10px; display: flex; align-items: center; gap: 8px; }
        .faq-card { background: #fff; border-radius: 14px; border: 1px solid #F0D4D6; overflow: hidden; }
        .faq-item { border-bottom: 1px solid #FDF0F0; }
        .faq-item:last-child { border-bottom: none; }
        .faq-btn { width: 100%; text-align: left; background: none; border: none; padding: 16px 20px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; gap: 12px; font-family: inherit; }
        .faq-btn:hover { background: #FDF8F8; }
        .faq-q { font-size: 14px; font-weight: 600; color: #111827; line-height: 1.4; }
        .faq-icon { font-size: 18px; color: #BE1522; flex-shrink: 0; transition: transform .2s; }
        .faq-icon.open { transform: rotate(45deg); }
        .faq-answer { padding: 0 20px 16px; font-size: 13px; color: #6B7280; line-height: 1.6; display: none; }
        .faq-answer.show { display: block; }
        .contact-box { background: #BE1522; border-radius: 16px; padding: 24px; text-align: center; margin-top: 32px; }
        .contact-box h3 { color: #fff; font-size: 16px; font-weight: 700; margin-bottom: 6px; }
        .contact-box p { color: rgba(255,255,255,.75); font-size: 13px; margin-bottom: 16px; }
        .contact-btns { display: flex; gap: 10px; justify-content: center; flex-wrap: wrap; }
        .contact-btn { padding: 10px 20px; border-radius: 10px; font-size: 13px; font-weight: 700; cursor: pointer; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; border: none; font-family: inherit; }
        .contact-btn-white { background: #fff; color: #BE1522; }
        .contact-btn-outline { background: rgba(255,255,255,.15); color: #fff; border: 1.5px solid rgba(255,255,255,.4); }
        .contact-btn:hover { opacity: .88; }
        @media(max-width:480px) { .header { padding: 24px 16px 32px; } .content { padding: 20px 16px 40px; } .contact-btns { flex-direction: column; align-items: stretch; } }
      `}</style>
      <div className="page">
        <div className="header">
          <div className="header-inner">
            <div className="logo">A</div>
            <h1>Centro de ayuda</h1>
            <p>Encuentra respuestas a las preguntas más frecuentes sobre la App Egresados ASEDUIS.</p>
          </div>
        </div>

        <div className="content">
          <div className="search-wrap">
            <span>🔍</span>
            <input placeholder="Buscar en el centro de ayuda..." id="searchInput" onInput={(e: any) => {
              const q = e.target.value.toLowerCase()
              document.querySelectorAll('.faq-item').forEach((el: any) => {
                const txt = el.textContent.toLowerCase()
                el.style.display = txt.includes(q) ? '' : 'none'
              })
            }} />
          </div>

          {FAQS.map((cat, ci) => (
            <div key={ci} className="cat-section">
              <div className="cat-title">{cat.cat}</div>
              <div className="faq-card">
                {cat.items.map((item, ii) => (
                  <div key={ii} className="faq-item">
                    <button className="faq-btn" onClick={(e) => {
                      const btn = e.currentTarget
                      const answer = btn.nextElementSibling as HTMLElement
                      const icon   = btn.querySelector('.faq-icon') as HTMLElement
                      const isOpen = answer.classList.contains('show')
                      answer.classList.toggle('show', !isOpen)
                      icon.classList.toggle('open', !isOpen)
                    }}>
                      <span className="faq-q">{item.q}</span>
                      <span className="faq-icon">+</span>
                    </button>
                    <div className="faq-answer">{item.a}</div>
                  </div>
                ))}
              </div>
            </div>
          ))}

          <div className="contact-box">
            <h3>¿No encontraste lo que buscabas?</h3>
            <p>Nuestro equipo de soporte está listo para ayudarte.</p>
            <div className="contact-btns">
              <a href="https://wa.me/573242606004" target="_blank" className="contact-btn contact-btn-white">
                💬 WhatsApp
              </a>
              <a href="mailto:soporte@aseduis.com" className="contact-btn contact-btn-outline">
                ✉️ soporte@aseduis.com
              </a>
            </div>
          </div>
        </div>
      </div>
    </>
  )
}