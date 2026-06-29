// dashboard/src/app/privacidad/page.tsx
export default function PrivacidadPage() {
  return (
    <>
      <style>{`
        *{box-sizing:border-box;margin:0;padding:0}
        body{font-family:"Segoe UI",system-ui,sans-serif;background:#F9F1F1;color:#111827}
        .wrap{max-width:760px;margin:0 auto;padding:40px 24px 80px}
        .header{background:#BE1522;color:#fff;borderRadius:16px;padding:32px;margin-bottom:32px;text-align:center}
        h1{font-size:28px;font-weight:800;margin-bottom:6px}
        .sub{font-size:13px;opacity:.8}
        .card{background:#fff;border-radius:14px;border:1px solid #F0D4D6;padding:24px;margin-bottom:16px}
        h2{font-size:15px;font-weight:700;color:#BE1522;margin-bottom:12px;text-transform:uppercase;letter-spacing:.5px}
        p{font-size:13px;color:#374151;line-height:1.8;margin-bottom:10px}
        ul{padding-left:18px;margin-bottom:10px}
        li{font-size:13px;color:#374151;line-height:1.8;margin-bottom:4px}
        .footer{text-align:center;font-size:12px;color:#94A3B8;margin-top:32px}
        a{color:#BE1522}
      `}</style>
      <div className="wrap">
        <div className="header">
          <h1>Política de Privacidad</h1>
          <div className="sub">ASEDUIS — Aplicación móvil para egresados</div>
          <div className="sub" style={{marginTop:6}}>Última actualización: junio 2026</div>
        </div>

        <div className="card">
          <h2>1. Información que recopilamos</h2>
          <p>La aplicación ASEDUIS recopila la siguiente información de los egresados registrados:</p>
          <ul>
            <li>Nombre completo y número de cédula</li>
            <li>Correo electrónico y número de teléfono</li>
            <li>Foto de perfil (opcional, subida por el usuario)</li>
            <li>Información académica: título, institución y fecha de grado</li>
            <li>Información laboral: empresa y cargo actual</li>
            <li>Fecha de nacimiento, ciudad y dirección</li>
            <li>Estado y fechas de membresía</li>
          </ul>
        </div>

        <div className="card">
          <h2>2. Cómo usamos la información</h2>
          <p>La información recopilada se usa exclusivamente para:</p>
          <ul>
            <li>Gestionar el carnet digital de egresado</li>
            <li>Verificar el estado de membresía mediante código QR</li>
            <li>Facilitar la comunicación institucional entre ASEDUIS y sus egresados</li>
            <li>Generar reportes internos de la asociación</li>
          </ul>
          <p>Los datos <strong>no se comparten con terceros</strong> ni se usan con fines comerciales.</p>
        </div>

        <div className="card">
          <h2>3. Almacenamiento y seguridad</h2>
          <p>Todos los datos se almacenan de forma segura en servidores de <strong>Supabase</strong>, con cifrado en tránsito (HTTPS) y en reposo. El acceso está restringido únicamente al personal autorizado de ASEDUIS.</p>
        </div>

        <div className="card">
          <h2>4. Permisos de la aplicación</h2>
          <p>La app solicita los siguientes permisos en tu dispositivo:</p>
          <ul>
            <li><strong>Cámara:</strong> para escanear códigos QR de validación de carnets</li>
            <li><strong>Fotos / Galería:</strong> para que el egresado pueda subir su foto de perfil</li>
          </ul>
          <p>Estos permisos son opcionales y puedes revocarlos desde la configuración de tu dispositivo.</p>
        </div>

        <div className="card">
          <h2>5. Derechos del usuario</h2>
          <p>Como egresado registrado tienes derecho a:</p>
          <ul>
            <li>Acceder y actualizar tu información personal desde la app</li>
            <li>Solicitar la eliminación de tus datos contactando a soporte</li>
            <li>Revocar permisos de cámara y fotos en cualquier momento</li>
          </ul>
        </div>

        <div className="card">
          <h2>6. Contacto</h2>
          <p>Para consultas sobre privacidad o solicitudes de eliminación de datos:</p>
          <ul>
          <li>Email: <a href="mailto:soporte@aseduis.com">soporte@aseduis.com</a></li>
          <li>WhatsApp: <a href="https://wa.me/573242606004?text=Hola%2C%20tengo%20una%20consulta%20sobre%20privacidad%20en%20la%20app%20ASEDUIS">+57 324 260 6004</a></li>
          <li>Sitio web: <a href="https://www.aseduis.com">www.aseduis.com</a></li>
          </ul>
        </div>

        <div className="footer">
          © 2026 ASEDUIS — Asociación de Egresados UIS · <a href="https://www.aseduis.com">www.aseduis.com</a>
        </div>
      </div>
    </>
  )
}