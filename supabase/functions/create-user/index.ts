// Edge Function: create-user
// Permite que un usuario admin cree usuarios nuevos (cuenta + perfil con permisos)
// desde el dashboard, sin exponer la service_role key en el navegador.

import { createClient } from "npm:@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const authHeader = req.headers.get("Authorization");
    if (!authHeader) {
      return new Response(JSON.stringify({ error: "Falta autenticación." }), { status: 401, headers: corsHeaders });
    }

    // Cliente "de servicio" con permisos totales (usa la service_role key,
    // que solo vive acá del lado del servidor, nunca en el navegador).
    const supabaseAdmin = createClient(
      Deno.env.get("SUPABASE_URL"),
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")
    );

    // Cliente "del usuario que llama", para verificar quién está pidiendo esto.
    const supabaseCaller = createClient(
      Deno.env.get("SUPABASE_URL"),
      Deno.env.get("SUPABASE_ANON_KEY"),
      { global: { headers: { Authorization: authHeader } } }
    );

    const { data: { user: caller }, error: callerError } = await supabaseCaller.auth.getUser();
    if (callerError || !caller) {
      return new Response(JSON.stringify({ error: "No se pudo verificar quién hace el pedido." }), { status: 401, headers: corsHeaders });
    }

    // Solo un admin puede crear usuarios
    const { data: perfilCaller } = await supabaseAdmin.from("perfiles").select("role").eq("id", caller.id).single();
    if (!perfilCaller || perfilCaller.role !== "admin") {
      return new Response(JSON.stringify({ error: "Solo un administrador puede crear usuarios." }), { status: 403, headers: corsHeaders });
    }

    const { email, password, username, permisos } = await req.json();
    if (!email || !password || !username) {
      return new Response(JSON.stringify({ error: "Faltan datos (email, contraseña o nombre de usuario)." }), { status: 400, headers: corsHeaders });
    }

    // Crea la cuenta de acceso (ya confirmada, sin necesidad de mail de verificación)
    const { data: nuevoUsuario, error: createError } = await supabaseAdmin.auth.admin.createUser({
      email, password, email_confirm: true,
    });
    if (createError) {
      return new Response(JSON.stringify({ error: createError.message }), { status: 400, headers: corsHeaders });
    }

    // Crea su perfil con los permisos elegidos
    const { data: perfil, error: perfilError } = await supabaseAdmin.from("perfiles").insert({
      id: nuevoUsuario.user.id,
      username,
      role: "usuario",
      permiso_camiones: !!permisos?.camiones,
      permiso_acoplados: !!permisos?.acoplados,
      permiso_choferes: !!permisos?.choferes,
      permiso_neumaticos: !!permisos?.neumaticos,
      permiso_rutas: !!permisos?.rutas,
    }).select().single();

    if (perfilError) {
      // Si falló crear el perfil, deshacemos la creación del usuario para no dejarlo huérfano
      await supabaseAdmin.auth.admin.deleteUser(nuevoUsuario.user.id);
      return new Response(JSON.stringify({ error: perfilError.message }), { status: 400, headers: corsHeaders });
    }

    return new Response(JSON.stringify({ perfil }), { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message || "Error inesperado." }), { status: 500, headers: corsHeaders });
  }
});
