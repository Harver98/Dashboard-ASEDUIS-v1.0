export const metadata = {
  title: 'ASEDUIS - Panel Administrativo',
  description: 'Dashboard administrativo de ASEDUIS - Asociación de Egresados UIS',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="es">
      <body>{children}</body>
    </html>
  )
}