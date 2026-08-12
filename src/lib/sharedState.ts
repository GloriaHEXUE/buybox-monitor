import { createClient } from '@supabase/supabase-js'

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY
const sharedStateTable = import.meta.env.VITE_SUPABASE_STATE_TABLE || 'app_state'
const sharedStateKey = import.meta.env.VITE_SUPABASE_STATE_KEY || 'global'

export const isSharedStateConfigured = Boolean(supabaseUrl && supabaseAnonKey)

const supabase = isSharedStateConfigured
  ? createClient(supabaseUrl!, supabaseAnonKey!, {
      auth: {
        persistSession: false,
        autoRefreshToken: false,
      },
    })
  : null

export type SharedStateRecord = {
  payload: unknown
  updated_at?: string | null
}

export const getSharedStateConfig = () => ({
  table: sharedStateTable,
  key: sharedStateKey,
})

export const readSharedState = async (): Promise<SharedStateRecord | null> => {
  if (!supabase) return null
  const { data, error } = await supabase
    .from(sharedStateTable)
    .select('payload, updated_at')
    .eq('state_key', sharedStateKey)
    .maybeSingle()

  if (error) throw error
  if (!data) return null
  return data as SharedStateRecord
}

export const writeSharedState = async (payload: unknown) => {
  if (!supabase) return
  const { error } = await supabase
    .from(sharedStateTable)
    .upsert(
      {
        state_key: sharedStateKey,
        payload,
        updated_at: new Date().toISOString(),
      },
      { onConflict: 'state_key' },
    )

  if (error) throw error
}
