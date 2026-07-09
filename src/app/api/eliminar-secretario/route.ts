// app/api/eliminar-secretario/route.ts
import { createClient } from '@supabase/supabase-js'
import { NextResponse } from 'next/server'

const supabaseAdmin = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL!,
  process.env.SUPABASE_SERVICE_ROLE_KEY!
)

export async function POST(req: Request) {
  const { secretarioId, userId } = await req.json()

  if (!secretarioId) {
    return NextResponse.json({ error: 'Falta secretarioId' }, { status: 400 })
  }

  // 1. Borrar la cuenta de auth primero (si existe user_id)
  if (userId) {
    const { error: authError } = await supabaseAdmin.auth.admin.deleteUser(userId)
    // No bloqueamos si el usuario de auth ya no existe (por ejemplo, huérfano ya limpiado antes)
    if (authError && authError.message && !authError.message.toLowerCase().includes('not found')) {
      return NextResponse.json({ error: authError.message }, { status: 400 })
    }
  }

  // 2. Borrar la fila de secretarios
  const { error: dbError } = await supabaseAdmin.from('secretarios').delete().eq('id', secretarioId)
  if (dbError) {
    return NextResponse.json({ error: dbError.message }, { status: 400 })
  }

  return NextResponse.json({ success: true })
}