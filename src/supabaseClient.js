import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  console.warn(
    "Faltan las variables VITE_SUPABASE_URL y/o VITE_SUPABASE_ANON_KEY. " +
    "Creá un archivo .env en la raíz del proyecto (podés copiar .env.example) " +
    "con los valores de tu proyecto de Supabase."
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
