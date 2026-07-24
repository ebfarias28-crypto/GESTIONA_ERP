import { createClient } from '@supabase/supabase-js';

// Credenciales del proyecto de BUILDERTECH en Supabase.
// La "publishable key" es segura para usarse en el navegador — está diseñada para eso
// (la seguridad real la da RLS en la base de datos, no el secreto de esta llave).
const SUPABASE_URL = 'https://hpkjmhyxfxsuvbdgkrkz.supabase.co';
const SUPABASE_PUBLISHABLE_KEY = 'sb_publishable_k9pw5-Frx0al1VxOH-_geA_FvWCnsst';

export const supabase = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY);
