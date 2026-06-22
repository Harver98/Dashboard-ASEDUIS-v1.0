'use client'
// dashboard/src/app/privacidad/page.tsx

export default function PrivacidadPage() {
  const fecha = '9 de junio de 2026'

  const SECCIONES = [
    {
      titulo: '1. Responsable del tratamiento',
      contenido: `ASEDUIS (Asociación de Egresados) es el responsable del tratamiento de los datos personales recopilados a través de la aplicación móvil "Egresados ASEDUIS" y su plataforma web administrativa. Para cualquier consulta relacionada con sus datos, puede contactarnos en soporte@aseduis.com.`,
    },
    {
      titulo: '2. Datos que recopilamos',
      contenido: `Recopilamos los siguientes datos personales:\n\n• Datos de identificación: nombre completo, número de cédula de ciudadanía.\n• Datos de contacto: correo electrónico, número de teléfono, dirección.\n• Datos académicos: título de pregrado y posgrado, institución, fecha de grado.\n• Datos laborales: empresa y cargo actual.\n• Datos biométricos: fotografía de perfil (opcional).\n• Datos de uso: historial de validaciones de carnet, fecha y hora de acceso.`,
    },
    {
      titulo: '3. Finalidad del tratamiento',
      contenido: `Los datos recopilados se utilizan exclusivamente para:\n\n• Emitir y gestionar el carnet digital de egresado.\n• Verificar la identidad del egresado mediante código QR.\n• Enviar comunicaciones institucionales y notificaciones de vencimiento del carnet.\n• Gestionar el acceso a convenios y beneficios institucionales.\n• Cumplir con obligaciones legales aplicables.`,
    },
    {
      titulo: '4. Base legal del tratamiento',
      contenido: `El tratamiento de sus datos se realiza con base en:\n\n• Su consentimiento expreso al registrarse en la plataforma.\n• La ejecución de la relación institucional entre el egresado y ASEDUIS.\n• El cumplimiento de obligaciones legales establecidas en la Ley 1581 de 2012 (Colombia) y sus decretos reglamentarios.`,
    },
    {
      titulo: '5. Conservación de los datos',
      contenido: `Sus datos personales se conservarán mientras mantenga su vínculo como egresado registrado en ASEDUIS. Una vez solicite la eliminación de su cuenta, sus datos serán eliminados en un plazo máximo de 30 días hábiles, salvo que exista obligación legal de conservarlos.`,
    },
    {
      titulo: '6. Compartición de datos',
      contenido: `ASEDUIS no vende, alquila ni comparte sus datos personales con terceros con fines comerciales. Sus datos pueden ser compartidos únicamente con:\n\n• Personal autorizado de ASEDUIS (secretarios y administradores) para la verificación del carnet.\n• Proveedores de servicios tecnológicos que actúan como encargados del tratamiento (Supabase como proveedor de base de datos y autenticación), bajo acuerdos de confidencialidad.\n• Autoridades competentes cuando exista obligación legal.`,
    },
    {
      titulo: '7. Seguridad de los datos',
      contenido: `Implementamos medidas técnicas y organizativas para proteger sus datos personales, incluyendo:\n\n• Cifrado de contraseñas mediante algoritmos seguros (bcrypt).\n• Comunicaciones cifradas mediante HTTPS/TLS.\n• Control de acceso mediante políticas de seguridad a nivel de base de datos (Row Level Security).\n• Autenticación segura gestionada por Supabase Auth.`,
    },
    {
      titulo: '8. Sus derechos',
      contenido: `De acuerdo con la Ley 1581 de 2012, usted tiene derecho a:\n\n• Conocer, actualizar y rectificar sus datos personales.\n• Solicitar la eliminación de sus datos cuando no exista obligación legal de conservarlos.\n• Revocar el consentimiento otorgado para el tratamiento de sus datos.\n• Presentar quejas ante la Superintendencia de Industria y Comercio (SIC).\n\nPara ejercer estos derechos, puede contactarnos en soporte@aseduis.com.`,
    },
    {
      titulo: '9. Cookies y tecnologías similares',
      contenido: `La aplicación móvil no utiliza cookies. La plataforma web administrativa puede utilizar almacenamiento local del navegador (localStorage) únicamente para mantener la sesión activa del administrador. No se utilizan cookies de rastreo ni publicidad.`,
    },
    {
      titulo: '10. Cambios en esta política',
      contenido: `ASEDUIS se reserva el derecho de actualizar esta política de privacidad. Cualquier cambio significativo será notificado a los usuarios registrados a través de la aplicación o por correo electrónico con al menos 15 días de anticipación.`,
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
        .header p { color: rgba(255,255,255,.75); font-size: 13px; }
        .content { max-width: 720px; margin: 0 auto; padding: 32px 24px 60px; }
        .fecha { font-size: 12px; color: #94A3B8; margin-bottom: 24px; }
        .intro { background: #fff; border-radius: 14px; border: 1px solid #F0D4D6; padding: 20px 24px; margin-bottom: 24px; font-size: 14px; color: #374151; line-height: 1.7; }
        .seccion { background: #fff; border-radius: 14px; border: 1px solid #F0D4D6; padding: 22px 24px; margin-bottom: 14px; }
        .seccion h2 { font-size: 14px; font-weight: 700; color: #BE1522; margin-bottom: 12px; }
        .seccion p { font-size: 13px; color: #374151; line-height: 1.75; white-space: pre-line; }
        .contact-box { background: #BE1522; border-radius: 16px; padding: 24px; text-align: center; margin-top: 32px; }
        .contact-box h3 { color: #fff; font-size: 16px; font-weight: 700; margin-bottom: 6px; }
        .contact-box p { color: rgba(255,255,255,.75); font-size: 13px; margin-bottom: 16px; }
        .contact-btn { display: inline-flex; align-items: center; gap: 6px; padding: 10px 20px; border-radius: 10px; font-size: 13px; font-weight: 700; cursor: pointer; text-decoration: none; background: #fff; color: #BE1522; border: none; }
        .contact-btn:hover { opacity: .88; }
        @media(max-width:480px) { .header { padding: 24px 16px 32px; } .content { padding: 20px 16px 40px; } .seccion { padding: 18px 16px; } }
      `}</style>

      <div className="page">
        <div className="header">
          <div className="header-inner">
            <div className="logo">A</div>
            <h1>Política de Privacidad</h1>
            <p>Términos y condiciones de tratamiento de datos personales · App Egresados ASEDUIS</p>
          </div>
        </div>

        <div className="content">
          <p className="fecha">Última actualización: {fecha}</p>

          <div className="intro">
            Esta Política de Privacidad describe cómo ASEDUIS recopila, usa y protege la información personal de los egresados registrados en la aplicación móvil y la plataforma web. Al usar la aplicación, usted acepta los términos descritos en este documento.
          </div>

          {SECCIONES.map((sec, i) => (
            <div key={i} className="seccion">
              <h2>{sec.titulo}</h2>
              <p>{sec.contenido}</p>
            </div>
          ))}

          <div className="contact-box">
            <h3>¿Tienes preguntas sobre tu privacidad?</h3>
            <p>Contáctanos directamente y te respondemos en menos de 48 horas.</p>
            <a href="mailto:soporte@aseduis.com" className="contact-btn">
              ✉️ soporte@aseduis.com
            </a>
          </div>
        </div>
      </div>
    </>
  )
}