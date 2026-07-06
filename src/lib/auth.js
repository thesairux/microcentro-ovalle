// ============================================================================
//  Conexión a Supabase + funciones de autenticación
//  Instala:  npm install @supabase/supabase-js
//  Variables (.env):  VITE_SUPABASE_URL, VITE_SUPABASE_ANON_KEY
// ============================================================================
import { createClient } from "@supabase/supabase-js";

export const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

// -------- Iniciar sesión -----------------------------------------------------
export async function login(correo, password) {
  const { data, error } = await supabase.auth.signInWithPassword({ email: correo, password });
  if (error) throw error;
  return data;
}

export async function logout() {
  await supabase.auth.signOut();
}

// -------- Perfil del usuario autenticado (con su rol y microcentro) ----------
export async function getPerfil() {
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return null;
  const { data, error } = await supabase
    .from("usuario")
    .select("*, membresia(microcentro_id, microcentro(nombre))")
    .eq("auth_id", user.id)
    .single();
  if (error) throw error;
  return data;
}

// -------- RECLAMAR ficha existente (evita duplicados) ------------------------
export async function reclamarCuenta({ correo, password }) {
  const { data: signUp, error: e1 } = await supabase.auth.signUp({ email: correo, password });
  if (e1) throw e1;

  const { data, error: e2 } = await supabase
    .from("usuario")
    .update({ auth_id: signUp.user.id, estado: "activo" })
    .eq("correo", correo)
    .eq("estado", "pendiente_activacion")
    .select()
    .single();
  if (e2) throw e2;
  if (!data) throw new Error("No se encontró una ficha pendiente con ese correo. Verifica con tu administrador.");
  return data;
}

// -------- Registrar persona NUEVA (no está en el listado) --------------------
export async function registrarNuevo({ nombre, correo, password, rut, telefono, rol, microcentroId, establecimiento }) {
  const { data: signUp, error: e1 } = await supabase.auth.signUp({ email: correo, password });
  if (e1) throw e1;

  const { data: nuevo, error: e2 } = await supabase
    .from("usuario")
    .insert({ auth_id: signUp.user.id, nombre, correo, rut, telefono, rol, establecimiento, estado: "pendiente_aprobacion" })
    .select()
    .single();
  if (e2) throw e2;

  if (microcentroId) {
    await supabase.from("membresia").insert({ usuario_id: nuevo.id, microcentro_id: microcentroId });
  }
  return nuevo;
}

// -------- Aprobar / rechazar (solo admin) ------------------------------------
export async function aprobarUsuario(usuarioId) {
  const { error } = await supabase.from("usuario").update({ estado: "activo" }).eq("id", usuarioId);
  if (error) throw error;
}
export async function rechazarUsuario(usuarioId) {
  const { error } = await supabase.from("usuario").update({ estado: "rechazado" }).eq("id", usuarioId);
  if (error) throw error;
}