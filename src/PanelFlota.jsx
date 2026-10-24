import React, { useState, useMemo, useEffect } from "react";
import {
  Truck, Container, User, Gauge, Disc3, Wrench, CalendarClock,
  AlertTriangle, Plus, X, Route, CheckCircle2, ChevronRight, ChevronLeft, Trash2, Clock, Search, UserCog,
  FileText, Download, Upload, Menu, LayoutDashboard
} from "lucide-react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";

const SERVICE_INTERVAL_DEFAULT = 50000;
const SERVICE_WARNING_MARGEN = 5000;
const VTV_WARNING_DAYS = 30;
const VTV_VIGENCIA_ANIOS = 1;

const todayISO = () => new Date().toISOString().slice(0, 10);
const daysUntil = (dateStr) => {
  const diff = new Date(dateStr) - new Date(todayISO());
  return Math.ceil(diff / (1000 * 60 * 60 * 24));
};
const addYears = (dateStr, years) => {
  const d = new Date(dateStr);
  d.setFullYear(d.getFullYear() + years);
  return d.toISOString().slice(0, 10);
};
const vencimientoVTV = (vtvUltima) => addYears(vtvUltima, VTV_VIGENCIA_ANIOS);
const fmtKm = (n) => n.toLocaleString("es-AR");

const seedCamiones = [
  { id: "c1", patente: "AC142XY", modelo: "Scania R450", km: 47200, ultimoServiceKm: 40000, intervaloService: 50000, vtvUltima: "2025-09-15",
    servicios: [{ id: "sv1", fecha: "2026-01-10", km: 40000, detalle: "Cambio de aceite y filtros, revisión de frenos", editado: false }] },
  { id: "c2", patente: "AB980LM", modelo: "Volvo FH", km: 31850, ultimoServiceKm: 0, intervaloService: 50000, vtvUltima: "2025-12-01", servicios: [] },
  { id: "c3", patente: "AD330PZ", modelo: "Mercedes Actros", km: 58900, ultimoServiceKm: 50000, intervaloService: 50000, vtvUltima: "2025-10-05",
    servicios: [{ id: "sv2", fecha: "2026-03-02", km: 50000, detalle: "Service completo 50.000km", editado: false }] },
];

const seedAcoplados = [
  { id: "a1", patente: "TR501AA", tipo: "Semirremolque", km: 44100, ultimoServiceKm: 40000, intervaloService: 50000, vtvUltima: "2025-09-20", servicios: [] },
  { id: "a2", patente: "TR622BB", tipo: "Batea", km: 31850, ultimoServiceKm: 0, intervaloService: 50000, vtvUltima: "2026-02-10", servicios: [] },
];

const seedChoferes = [
  { id: "d1", nombre: "Raúl Pereyra", km: 132400, camionId: "c1", acopladoId: "a1" },
  { id: "d2", nombre: "Marta Godoy", km: 98750, camionId: "c2", acopladoId: "a2" },
];

const seedNeumaticos = [
  { id: "n1", etiqueta: "N-0231", estado: "Nuevo", km: 47200, camionId: "c1", acopladoId: null, posicion: "DI",
    historial: [{ fecha: "2026-01-10", tipo: "alta", detalle: "Alta como nuevo, montado en AC142XY · DI" }] },
  { id: "n2", etiqueta: "N-0198", estado: "Recapado", km: 800, camionId: null, acopladoId: "a1", posicion: "1EI1",
    historial: [
      { fecha: "2025-06-02", tipo: "alta", detalle: "Alta como nuevo" },
      { fecha: "2026-07-20", tipo: "recapado", detalle: "Enviado a recapar" },
      { fecha: "2026-07-25", tipo: "cambio_vehiculo", detalle: "Montado en TR501AA · 1EI1" },
    ] },
  { id: "n3", etiqueta: "N-0177", estado: "Nuevo", km: 31850, camionId: "c2", acopladoId: null, posicion: "TRI1",
    historial: [{ fecha: "2025-11-05", tipo: "alta", detalle: "Alta como nuevo, montado en AB980LM · TR I1" }] },
  { id: "n4", etiqueta: "N-0312", estado: "Nuevo", km: 0, camionId: null, acopladoId: null, posicion: null,
    historial: [{ fecha: "2026-08-01", tipo: "alta", detalle: "Alta como nuevo, sin asignar a ningún camión/acoplado" }] },
  { id: "n5", etiqueta: "N-0290", estado: "Usado", km: 12500, camionId: null, acopladoId: null, posicion: null,
    historial: [{ fecha: "2026-05-14", tipo: "alta", detalle: "Alta como usado, sin asignar a ningún camión/acoplado" }] },
];

function estadoNeumatico(n) {
  if (!n.camionId && !n.acopladoId) return "Sin asignar";
  return n.estado === "Recapado" ? "Recapado" : "Sin recapar";
}

function colorEstadoNeumatico(label) {
  if (label === "Sin asignar") return "var(--text-muted)";
  if (label === "Recapado") return "var(--accent-blue)";
  return "var(--accent-green)";
}

const MESES_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
function mesKey(fecha) { return fecha.slice(0, 7); } // "YYYY-MM"

// --- Km real: siempre se calcula sumando los viajes cargados, nunca se confía
// en un contador guardado aparte, para que nunca pueda desincronizarse. ---
function sumaViajes(historialViajes, filtroFn) {
  return historialViajes.filter(filtroFn).reduce((acc, v) => acc + v.km, 0);
}
function kmRealCamion(camion, historialViajes) {
  return (camion.kmInicial || 0) + sumaViajes(historialViajes, (v) => v.camionId === camion.id);
}
function kmRealAcoplado(acoplado, historialViajes) {
  return (acoplado.kmInicial || 0) + sumaViajes(historialViajes, (v) => v.acopladoId === acoplado.id);
}
function kmRealChofer(chofer, historialViajes) {
  return sumaViajes(historialViajes, (v) => v.choferId === chofer.id);
}

function mesLabel(key) {
  const [y, m] = key.split("-");
  return `${MESES_ES[parseInt(m, 10) - 1]} ${y}`;
}
function descargarExcelMes(key, viajesDelMes, choferes, camiones, acoplados) {
  const filas = viajesDelMes.map((v) => {
    const chofer = choferes.find((d) => d.id === v.choferId);
    const camion = camiones.find((c) => c.id === v.camionId);
    const acoplado = acoplados.find((a) => a.id === v.acopladoId);
    return {
      Fecha: v.fecha,
      Código: v.codigo || "",
      Remito: v.remito || "",
      Ruta: v.ruta,
      Chofer: chofer?.nombre || "",
      Camión: camion?.patente || "",
      Acoplado: acoplado?.patente || "",
      Km: v.km,
    };
  });
  const ws = XLSX.utils.json_to_sheet(filas);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Viajes");
  XLSX.writeFile(wb, `viajes-${key}.xlsx`);
}

function colorVencimiento(fecha) {
  if (!fecha) return "var(--text-muted)";
  const d = daysUntil(fecha);
  if (d < 0) return "var(--accent-red)";
  if (d <= VTV_WARNING_DAYS) return "var(--accent-yellow)";
  return "var(--text)";
}

// --- Posiciones de neumáticos ---
const CAMION_EJES = [
  { nombre: "Dirección", tipo: "simple", codigos: ["DI", "DD"] },
  { nombre: "Tracción", tipo: "doble", codigos: ["TRI1", "TRI2", "TRD2", "TRD1"] },
  { nombre: "Neumático", tipo: "doble", codigos: ["NI1", "NI2", "ND2", "ND1"] },
];

const ACOPLADO_EJES = [
  { nombre: "1er eje", tipo: "doble", codigos: ["1EI1", "1EI2", "1ED2", "1ED1"] },
  { nombre: "2do eje", tipo: "doble", codigos: ["2EI1", "2EI2", "2ED2", "2ED1"] },
  { nombre: "3er eje", tipo: "doble", codigos: ["3EI1", "3EI2", "3ED2", "3ED1"] },
];

const POSICION_LABELS = {
  DI: "DI", DD: "DD",
  TRI1: "TR I1", TRI2: "TR I2", TRD1: "TR D1", TRD2: "TR D2",
  NI1: "NI1", NI2: "NI2", ND1: "ND1", ND2: "ND2",
  "1EI1": "1EI1", "1EI2": "1EI2", "1ED1": "1ED1", "1ED2": "1ED2",
  "2EI1": "2EI1", "2EI2": "2EI2", "2ED1": "2ED1", "2ED2": "2ED2",
  "3EI1": "3EI1", "3EI2": "3EI2", "3ED1": "3ED1", "3ED2": "3ED2",
};

const seedUsers = [
  { id: "u1", username: "admin", password: "admin123", role: "admin", permisos: { camiones: true, acoplados: true, choferes: true, neumaticos: true, rutas: true } },
];

const PERMISO_LABELS = { camiones: "Camiones", acoplados: "Acoplados", choferes: "Choferes", neumaticos: "Neumáticos", rutas: "Rutas" };

// --- Mapeo entre las columnas de Supabase (snake_case) y el modelo de la app (camelCase) ---
function mapCamionDb(row, servicios = []) {
  return {
    id: row.id, patente: row.patente, modelo: row.modelo, km: row.km, kmInicial: row.km_inicial,
    ultimoServiceKm: row.ultimo_service_km, intervaloService: row.intervalo_service,
    vtvUltima: row.vtv_ultima, polizaUrl: row.poliza_url, polizaNombre: row.poliza_nombre,
    vencimientoIntercement: row.vencimiento_intercement, servicios,
  };
}
function mapAcopladoDb(row, servicios = []) {
  return {
    id: row.id, patente: row.patente, tipo: row.tipo, km: row.km, kmInicial: row.km_inicial,
    ultimoServiceKm: row.ultimo_service_km, intervaloService: row.intervalo_service,
    vtvUltima: row.vtv_ultima, vencValvula1: row.venc_valvula1, vencValvula2: row.venc_valvula2,
    vencimientoIntercement: row.vencimiento_intercement, servicios,
  };
}
function mapChoferDb(row) {
  return {
    id: row.id, nombre: row.nombre, km: row.km, camionId: row.camion_id, acopladoId: row.acoplado_id,
    vencimientoDni: row.vencimiento_dni, vencimientoLicencia: row.vencimiento_licencia,
    fechaCursoDefensivo: row.fecha_curso_defensivo, fechaEstudiosMedicos: row.fecha_estudios_medicos,
    fechaSvo: row.fecha_svo, fechaArt: row.fecha_art, fechaEepp: row.fecha_eepp,
  };
}
function mapNeumaticoDb(row, historial = []) {
  return { id: row.id, etiqueta: row.etiqueta, estado: row.estado, km: row.km, camionId: row.camion_id, acopladoId: row.acoplado_id, posicion: row.posicion, marca: row.marca, tipo: row.tipo, historial };
}
function mapServicioDb(row) {
  return { id: row.id, fecha: row.fecha, km: row.km, detalle: row.detalle, editado: row.editado, fechaEdicion: row.fecha_edicion };
}
function mapHistorialDb(row) {
  return { fecha: row.fecha, tipo: row.tipo, detalle: row.detalle };
}
function mapViajeDb(row) {
  return { id: row.id, fecha: row.fecha, ruta: row.ruta_nombre, rutaId: row.ruta_id, km: row.km, choferId: row.chofer_id, camionId: row.camion_id, acopladoId: row.acoplado_id, distinto: row.combo_distinto, codigo: row.codigo, remito: row.remito };
}
function mapPerfilDb(row) {
  return {
    id: row.id, username: row.username, role: row.role,
    permisos: {
      camiones: row.permiso_camiones, acoplados: row.permiso_acoplados,
      choferes: row.permiso_choferes, neumaticos: row.permiso_neumaticos, rutas: row.permiso_rutas,
    },
  };
}


const seedRutas = [
  { id: "r1", nombre: "Buenos Aires → Rosario", km: 320 },
  { id: "r2", nombre: "Rosario → Córdoba", km: 400 },
  { id: "r3", nombre: "Buenos Aires → Mar del Plata", km: 410 },
];

const SECCIONES = [
  { key: "general", label: "General", icon: LayoutDashboard },
  { key: "rutas", label: "Rutas", icon: Route, permiso: "rutas" },
  { key: "camiones", label: "Camiones", icon: Truck, permiso: "camiones" },
  { key: "acoplados", label: "Acoplados", icon: Container, permiso: "acoplados" },
  { key: "choferes", label: "Choferes", icon: User, permiso: "choferes" },
  { key: "neumaticos", label: "Neumáticos", icon: Disc3, permiso: "neumaticos" },
];

const QUICK_ADD_TYPES = [
  { key: "camion", label: "Camión", icon: Truck, permiso: "camiones" },
  { key: "acoplado", label: "Acoplado", icon: Container, permiso: "acoplados" },
  { key: "chofer", label: "Chofer", icon: User, permiso: "choferes" },
  { key: "neumatico", label: "Neumático", icon: Disc3, permiso: "neumaticos" },
  { key: "ruta", label: "Ruta", icon: Route, permiso: "rutas" },
];

function Field({ label, children }) {
  return (
    <label className="flex flex-col gap-1.5 text-sm">
      <span className="uppercase tracking-wide text-[11px] font-semibold" style={{ color: "var(--text-muted)" }}>
        {label}
      </span>
      {children}
    </label>
  );
}

const inputCls =
  "bg-white border border-[var(--border)] rounded-md px-3 py-2 text-[#0A1128] text-sm focus:outline-none focus:ring-2 focus:ring-[var(--accent-yellow)]/60 focus:border-transparent";

function LoginScreen({ onLogin, error, loading }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");

  return (
    <div
      style={{
        "--bg": "#0A1128", "--panel": "#101B36", "--panel-2": "#16264A", "--border": "#233258",
        "--text": "#FFFFFF", "--text-muted": "#93A7CC", "--accent-yellow": "#F2B705",
        "--accent-red": "#F87171", "--accent-green": "#34D399", "--accent-blue": "#F2B705",
        background: "var(--bg)", color: "var(--text)", minHeight: "100vh",
        display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', system-ui, sans-serif",
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@600;700&family=Inter:wght@400;500;600&display=swap');
        .font-display { font-family: 'Oswald', sans-serif; letter-spacing: 0.02em; }
        input { color: #0A1128 !important; color-scheme: light; }
      `}</style>
      <div className="w-full max-w-sm p-6 rounded-lg mx-4" style={{ background: "var(--panel)", border: "1px solid var(--border)" }}>
        <div className="text-xs uppercase tracking-[0.2em] mb-1" style={{ color: "var(--accent-yellow)" }}>Control de flota</div>
        <h1 className="font-display text-2xl font-semibold uppercase mb-5">Iniciar sesión</h1>
        <div className="grid gap-3">
          <Field label="Email">
            <input type="email" className={inputCls} value={email} onChange={(e) => setEmail(e.target.value)} placeholder="admin@tuempresa.com" />
          </Field>
          <Field label="Contraseña">
            <input type="password" className={inputCls} value={password} onChange={(e) => setPassword(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && onLogin(email, password)} placeholder="••••••••" />
          </Field>
          {error && <p className="text-sm" style={{ color: "var(--accent-red)" }}>{error}</p>}
          <button
            onClick={() => onLogin(email, password)}
            disabled={loading}
            className="mt-1 px-4 py-2 rounded-md text-sm font-semibold disabled:opacity-50"
            style={{ background: "var(--accent-yellow)", color: "#0A1128" }}
          >
            {loading ? "Ingresando…" : "Ingresar"}
          </button>
        </div>
      </div>
    </div>
  );
}

export default function PanelFlota() {
  const [users, setUsers] = useState([]); // lista de perfiles (para la vista de admin)
  const [currentUser, setCurrentUser] = useState(null);
  const [loginError, setLoginError] = useState("");
  const [loginLoading, setLoginLoading] = useState(false);
  const [dataLoading, setDataLoading] = useState(false);

  const [camiones, setCamiones] = useState([]);
  const [acoplados, setAcoplados] = useState([]);
  const [choferes, setChoferes] = useState([]);
  const [neumaticos, setNeumaticos] = useState([]);
  const [rutas, setRutas] = useState([]);
  const [historialViajes, setHistorialViajes] = useState([]);
  const [selected, setSelected] = useState(null); // { type, id }
  const [verTodos, setVerTodos] = useState(null); // "camiones" | "acoplados" | "choferes" | "neumaticos" | null
  const [verTodasAlertas, setVerTodasAlertas] = useState(false);
  const [busquedaCodigo, setBusquedaCodigo] = useState("");
  const [mesSeleccionado, setMesSeleccionado] = useState("");
  const [choferSeleccionado, setChoferSeleccionado] = useState("");

  const [activeQuickAdd, setActiveQuickAdd] = useState(null);
  const [seccionActiva, setSeccionActiva] = useState("general");
  const [menuAbierto, setMenuAbierto] = useState(false);
  const [toast, setToast] = useState(null);

  const [fCamion, setFCamion] = useState({ patente: "", modelo: "", km: "", vtvUltima: "", ultimoServiceKm: "", intervaloService: "", vencimientoIntercement: "" });
  const [fAcoplado, setFAcoplado] = useState({ patente: "", tipo: "", km: "", vtvUltima: "", ultimoServiceKm: "", intervaloService: "", vencValvula1: "", vencValvula2: "", vencimientoIntercement: "" });
  const [fChofer, setFChofer] = useState({
    nombre: "", camionId: "", acopladoId: "",
    vencimientoDni: "", vencimientoLicencia: "", fechaCursoDefensivo: "", fechaEstudiosMedicos: "", fechaSvo: "", fechaArt: "", fechaEepp: "",
  });
  const [fNeumatico, setFNeumatico] = useState({ etiqueta: "", estado: "Nuevo", marca: "", tipo: "", km: "", montaje: "", vehiculoId: "", posicion: "" });
  const [fRuta, setFRuta] = useState({ nombre: "", km: "" });
  const [busqueda, setBusqueda] = useState({ camiones: "", acoplados: "", choferes: "", neumaticos: "", rutas: "" });
  const [fUsuario, setFUsuario] = useState({ email: "", password: "", username: "", permisos: { camiones: true, acoplados: true, choferes: true, neumaticos: true, rutas: true } });

  const showToast = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 3500);
  };

  // --- Carga inicial de todos los datos desde Supabase, una vez logueado ---
  async function cargarDatos() {
    setDataLoading(true);
    try {
      const [
        { data: camionesRows, error: e1 },
        { data: acopladosRows, error: e2 },
        { data: choferesRows, error: e3 },
        { data: neumaticosRows, error: e4 },
        { data: rutasRows, error: e5 },
        { data: viajesRows, error: e6 },
        { data: serviciosRows, error: e7 },
        { data: historialRows, error: e8 },
        { data: perfilesRows, error: e9 },
      ] = await Promise.all([
        supabase.from("camiones").select("*").order("patente"),
        supabase.from("acoplados").select("*").order("patente"),
        supabase.from("choferes").select("*").order("nombre"),
        supabase.from("neumaticos").select("*").order("etiqueta"),
        supabase.from("rutas").select("*").order("nombre"),
        supabase.from("viajes").select("*").order("created_at", { ascending: false }),
        supabase.from("servicios").select("*").order("fecha", { ascending: false }),
        supabase.from("neumatico_historial").select("*").order("fecha", { ascending: true }),
        supabase.from("perfiles").select("*").order("username"),
      ]);
      const err = e1 || e2 || e3 || e4 || e5 || e6 || e7 || e8 || e9;
      if (err) { showToast("Error cargando datos: " + err.message); return; }

      const serviciosPorCamion = {}, serviciosPorAcoplado = {};
      (serviciosRows || []).forEach((s) => {
        const target = s.camion_id ? serviciosPorCamion : serviciosPorAcoplado;
        const key = s.camion_id || s.acoplado_id;
        (target[key] ||= []).push(mapServicioDb(s));
      });
      const historialPorNeumatico = {};
      (historialRows || []).forEach((h) => {
        (historialPorNeumatico[h.neumatico_id] ||= []).push(mapHistorialDb(h));
      });

      setCamiones((camionesRows || []).map((r) => mapCamionDb(r, serviciosPorCamion[r.id] || [])));
      setAcoplados((acopladosRows || []).map((r) => mapAcopladoDb(r, serviciosPorAcoplado[r.id] || [])));
      setChoferes((choferesRows || []).map(mapChoferDb));
      setNeumaticos((neumaticosRows || []).map((r) => mapNeumaticoDb(r, historialPorNeumatico[r.id] || [])));
      setRutas(rutasRows || []);
      setHistorialViajes((viajesRows || []).map(mapViajeDb));
      setUsers((perfilesRows || []).map(mapPerfilDb));
    } finally {
      setDataLoading(false);
    }
  }

  useEffect(() => {
    if (currentUser) cargarDatos();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentUser?.id]);

  async function handleLogin(email, password) {
    setLoginLoading(true);
    setLoginError("");
    const { data, error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) { setLoginError(`Error de Supabase: ${error.message}`); setLoginLoading(false); return; }
    const { data: perfil, error: perfilError } = await supabase.from("perfiles").select("*").eq("id", data.user.id).single();
    setLoginLoading(false);
    if (perfilError || !perfil) { setLoginError("Este usuario no tiene un perfil asignado en el sistema."); await supabase.auth.signOut(); return; }
    setCurrentUser(mapPerfilDb(perfil));
  }

  async function handleLogout() {
    await supabase.auth.signOut();
    setCurrentUser(null);
    setSelected(null);
    setActiveQuickAdd(null);
    setVerTodos(null);
    setCamiones([]); setAcoplados([]); setChoferes([]); setNeumaticos([]); setRutas([]); setHistorialViajes([]); setUsers([]);
  }

  // Vincula un perfil de permisos a un usuario que ya fue creado en Supabase Auth
  // (Authentication → Users). Crear la cuenta en sí requiere el panel de Supabase
  // o una Edge Function con la service_role key, no se puede hacer de forma segura
  // solo con la anon key desde el navegador.
  async function saveUsuario() {
    if (!fUsuario.email.trim() || !fUsuario.password.trim() || !fUsuario.username.trim()) {
      return showToast("Falta el email, la contraseña o el nombre de usuario.");
    }
    const { data: sessionData } = await supabase.auth.getSession();
    const token = sessionData?.session?.access_token;
    if (!token) return showToast("Sesión inválida, volvé a iniciar sesión.");

    try {
      const res = await fetch(`${import.meta.env.VITE_SUPABASE_URL}/functions/v1/create-user`, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify({
          email: fUsuario.email, password: fUsuario.password, username: fUsuario.username, permisos: fUsuario.permisos,
        }),
      });
      const result = await res.json();
      if (!res.ok) return showToast("Error: " + (result.error || "no se pudo crear el usuario."));
      setUsers([...users, mapPerfilDb(result.perfil)]);
      setFUsuario({ email: "", password: "", username: "", permisos: { camiones: true, acoplados: true, choferes: true, neumaticos: true, rutas: true } });
      setActiveQuickAdd(null);
      showToast(`Usuario ${fUsuario.username} creado.`);
    } catch (err) {
      showToast("Error de conexión al crear el usuario.");
    }
  }

  const alerts = useMemo(() => {
    const list = [];
    const revisarFechaDirecta = (fecha, entidadLabel, campoLabel) => {
      if (!fecha) return;
      const d = daysUntil(fecha);
      if (d <= VTV_WARNING_DAYS) {
        list.push({
          severity: d < 0 ? "danger" : "warning",
          icon: CalendarClock,
          title: `${entidadLabel} — ${campoLabel} ${d < 0 ? "vencido" : "por vencer"}`,
          detail: d < 0 ? `Venció hace ${Math.abs(d)} días` : `Vence en ${d} días (${fecha})`,
        });
      }
    };
    const revisar = (v, tipoLabel, kmActual) => {
      const intervalo = v.intervaloService || SERVICE_INTERVAL_DEFAULT;
      const nextService = v.ultimoServiceKm + intervalo;
      const kmSinceService = kmActual - v.ultimoServiceKm;
      if (kmSinceService >= intervalo - SERVICE_WARNING_MARGEN) {
        const overdue = kmSinceService >= intervalo;
        list.push({
          severity: overdue ? "danger" : "warning",
          icon: Wrench,
          title: `${v.patente} (${tipoLabel}) — ${overdue ? "service vencido" : "service próximo"}`,
          detail: `${fmtKm(kmActual)} km · próximo a los ${fmtKm(nextService)} km`,
        });
      }
      if (v.vtvUltima) {
        const vencimiento = vencimientoVTV(v.vtvUltima);
        const d = daysUntil(vencimiento);
        if (d <= VTV_WARNING_DAYS) {
          list.push({
            severity: d < 0 ? "danger" : "warning",
            icon: CalendarClock,
            title: `${v.patente} (${tipoLabel}) — VTV ${d < 0 ? "vencida" : "por vencer"}`,
            detail: d < 0 ? `Venció hace ${Math.abs(d)} días` : `Vence en ${d} días (${vencimiento})`,
          });
        }
      }
    };
    camiones.forEach((c) => {
      revisar(c, "camión", kmRealCamion(c, historialViajes));
      revisarFechaDirecta(c.vencimientoIntercement, `${c.patente} (camión)`, "Verificación Intercement");
    });
    acoplados.forEach((a) => {
      revisar(a, "acoplado", kmRealAcoplado(a, historialViajes));
      revisarFechaDirecta(a.vencimientoIntercement, `${a.patente} (acoplado)`, "Verificación Intercement");
      if (a.tipo === "TOLVA") {
        revisarFechaDirecta(a.vencValvula1, `${a.patente} (acoplado)`, "Válvula 1");
        revisarFechaDirecta(a.vencValvula2, `${a.patente} (acoplado)`, "Válvula 2");
      }
    });
    choferes.forEach((d) => {
      revisarFechaDirecta(d.vencimientoDni, d.nombre, "DNI");
      revisarFechaDirecta(d.vencimientoLicencia, d.nombre, "Licencia");
      revisarFechaDirecta(d.fechaCursoDefensivo, d.nombre, "Curso defensivo");
      revisarFechaDirecta(d.fechaEstudiosMedicos, d.nombre, "Estudios médicos");
      revisarFechaDirecta(d.fechaSvo, d.nombre, "SVO");
      revisarFechaDirecta(d.fechaArt, d.nombre, "ART");
      revisarFechaDirecta(d.fechaEepp, d.nombre, "EEPP");
    });
    return list;
  }, [camiones, acoplados, choferes, historialViajes]);

  const kpis = [
    { label: "Camiones", value: camiones.length, icon: Truck },
    { label: "Acoplados", value: acoplados.length, icon: Container },
    { label: "Choferes", value: choferes.length, icon: User },
    { label: "Neumáticos", value: neumaticos.length, icon: Disc3 },
  ];

  function openQuickAdd(key) {
    setActiveQuickAdd(activeQuickAdd === key ? null : key);
  }

  async function saveCamion() {
    if (!fCamion.patente.trim()) return showToast("Falta la patente del camión.");
    const kmActual = Math.max(0, parseInt(fCamion.km, 10) || 0);
    const { data, error } = await supabase.from("camiones").insert({
      patente: fCamion.patente.toUpperCase(), modelo: fCamion.modelo || "—",
      km: kmActual, km_inicial: kmActual,
      ultimo_service_km: Math.max(0, parseInt(fCamion.ultimoServiceKm, 10) || 0),
      intervalo_service: Math.max(1, parseInt(fCamion.intervaloService, 10) || SERVICE_INTERVAL_DEFAULT),
      vtv_ultima: fCamion.vtvUltima || null,
      vencimiento_intercement: fCamion.vencimientoIntercement || null,
    }).select().single();
    if (error) return showToast("Error al guardar: " + error.message);
    setCamiones([...camiones, mapCamionDb(data, [])]);
    setFCamion({ patente: "", modelo: "", km: "", vtvUltima: "", ultimoServiceKm: "", intervaloService: "", vencimientoIntercement: "" });
    setActiveQuickAdd(null);
    showToast(`Camión ${data.patente} cargado.`);
  }

  async function saveAcoplado() {
    if (!fAcoplado.patente.trim()) return showToast("Falta la patente del acoplado.");
    const kmActual = Math.max(0, parseInt(fAcoplado.km, 10) || 0);
    const { data, error } = await supabase.from("acoplados").insert({
      patente: fAcoplado.patente.toUpperCase(), tipo: fAcoplado.tipo || "—",
      km: kmActual, km_inicial: kmActual,
      ultimo_service_km: Math.max(0, parseInt(fAcoplado.ultimoServiceKm, 10) || 0),
      intervalo_service: Math.max(1, parseInt(fAcoplado.intervaloService, 10) || SERVICE_INTERVAL_DEFAULT),
      vtv_ultima: fAcoplado.vtvUltima || null,
      venc_valvula1: fAcoplado.tipo === "TOLVA" ? fAcoplado.vencValvula1 || null : null,
      venc_valvula2: fAcoplado.tipo === "TOLVA" ? fAcoplado.vencValvula2 || null : null,
      vencimiento_intercement: fAcoplado.vencimientoIntercement || null,
    }).select().single();
    if (error) return showToast("Error al guardar: " + error.message);
    setAcoplados([...acoplados, mapAcopladoDb(data, [])]);
    setFAcoplado({ patente: "", tipo: "", km: "", vtvUltima: "", ultimoServiceKm: "", intervaloService: "", vencValvula1: "", vencValvula2: "", vencimientoIntercement: "" });
    setActiveQuickAdd(null);
    showToast(`Acoplado ${data.patente} cargado.`);
  }

  async function saveChofer() {
    if (!fChofer.nombre.trim()) return showToast("Falta el nombre del chofer.");
    const { data, error } = await supabase.from("choferes").insert({
      nombre: fChofer.nombre, km: 0,
      camion_id: fChofer.camionId || null, acoplado_id: fChofer.acopladoId || null,
      vencimiento_dni: fChofer.vencimientoDni || null,
      vencimiento_licencia: fChofer.vencimientoLicencia || null,
      fecha_curso_defensivo: fChofer.fechaCursoDefensivo || null,
      fecha_estudios_medicos: fChofer.fechaEstudiosMedicos || null,
      fecha_svo: fChofer.fechaSvo || null,
      fecha_art: fChofer.fechaArt || null,
      fecha_eepp: fChofer.fechaEepp || null,
    }).select().single();
    if (error) return showToast("Error al guardar: " + error.message);
    setChoferes([...choferes, mapChoferDb(data)]);
    setFChofer({ nombre: "", camionId: "", acopladoId: "", vencimientoDni: "", vencimientoLicencia: "", fechaCursoDefensivo: "", fechaEstudiosMedicos: "", fechaSvo: "", fechaArt: "", fechaEepp: "" });
    setActiveQuickAdd(null);
    showToast(`Chofer ${data.nombre} cargado.`);
  }

  async function saveNeumatico() {
    if (!fNeumatico.etiqueta.trim()) return showToast("Falta la etiqueta del neumático.");
    const kmInicial = Math.max(0, parseInt(fNeumatico.km, 10) || 0);
    const camionId = fNeumatico.montaje === "camion" ? fNeumatico.vehiculoId || null : null;
    const acopladoId = fNeumatico.montaje === "acoplado" ? fNeumatico.vehiculoId || null : null;
    const posicion = camionId || acopladoId ? fNeumatico.posicion || null : null;
    const montadoEn = camiones.find((c) => c.id === camionId)?.patente || acoplados.find((a) => a.id === acopladoId)?.patente;
    const posicionLabel = posicion ? POSICION_LABELS[posicion] || posicion : null;

    // si esa posición ya estaba ocupada por otro neumático, lo libera
    const ocupante = posicion
      ? neumaticos.find((n) => n.posicion === posicion && ((camionId && n.camionId === camionId) || (acopladoId && n.acopladoId === acopladoId)))
      : null;
    if (ocupante) {
      await supabase.from("neumaticos").update({ camion_id: null, acoplado_id: null, posicion: null }).eq("id", ocupante.id);
      await supabase.from("neumatico_historial").insert({
        neumatico_id: ocupante.id, tipo: "cambio_vehiculo",
        detalle: `Removido de ${montadoEn} · ${posicionLabel} (reemplazado por otro neumático)`,
      });
    }

    const { data, error } = await supabase.from("neumaticos").insert({
      etiqueta: fNeumatico.etiqueta, estado: fNeumatico.estado, km: kmInicial,
      marca: fNeumatico.marca || null, tipo: fNeumatico.tipo || null,
      camion_id: camionId, acoplado_id: acopladoId, posicion,
    }).select().single();
    if (error) return showToast("Error al guardar: " + error.message);

    const detalleAlta = `Alta como ${fNeumatico.estado.toLowerCase()}${montadoEn ? `, montado en ${montadoEn}${posicionLabel ? ` · ${posicionLabel}` : ""}` : ", sin asignar a ningún camión/acoplado"}`;
    await supabase.from("neumatico_historial").insert({ neumatico_id: data.id, tipo: "alta", detalle: detalleAlta });

    setNeumaticos([
      ...neumaticos.map((n) => (n.id === ocupante?.id ? { ...n, camionId: null, acopladoId: null, posicion: null, historial: [...n.historial, { fecha: todayISO(), tipo: "cambio_vehiculo", detalle: `Removido de ${montadoEn} · ${posicionLabel} (reemplazado por otro neumático)` }] } : n)),
      mapNeumaticoDb(data, [{ fecha: todayISO(), tipo: "alta", detalle: detalleAlta }]),
    ]);
    setFNeumatico({ etiqueta: "", estado: "Nuevo", marca: "", tipo: "", km: "", montaje: "", vehiculoId: "", posicion: "" });
    setActiveQuickAdd(null);
    showToast(`Neumático ${data.etiqueta} cargado con ${fmtKm(kmInicial)} km${montadoEn ? ` en ${montadoEn}${posicionLabel ? ` · ${posicionLabel}` : ""}` : ""}.`);
  }

  async function marcarRecapado(id) {
    const { error } = await supabase.from("neumaticos").update({ estado: "Recapado" }).eq("id", id);
    if (error) return showToast("Error: " + error.message);
    await supabase.from("neumatico_historial").insert({ neumatico_id: id, tipo: "recapado", detalle: "Enviado a recapar" });
    setNeumaticos(neumaticos.map((n) =>
      n.id === id
        ? { ...n, estado: "Recapado", historial: [...n.historial, { fecha: todayISO(), tipo: "recapado", detalle: "Enviado a recapar" }] }
        : n
    ));
    showToast("Neumático marcado como enviado a recapar.");
  }

  async function cambiarVehiculoNeumatico(id, tipo, vehiculoId, posicion) {
    const camionId = tipo === "camion" ? vehiculoId || null : null;
    const acopladoId = tipo === "acoplado" ? vehiculoId || null : null;
    const destino = tipo === "camion" ? camiones.find((c) => c.id === vehiculoId)?.patente : acoplados.find((a) => a.id === vehiculoId)?.patente;
    const label = posicion ? (POSICION_LABELS[posicion] || posicion) : null;
    const posicionFinal = vehiculoId ? posicion || null : null;

    const ocupante = posicionFinal
      ? neumaticos.find((n) => n.id !== id && n.posicion === posicionFinal &&
          ((tipo === "camion" && n.camionId === vehiculoId) || (tipo === "acoplado" && n.acopladoId === vehiculoId)))
      : null;
    if (ocupante) {
      await supabase.from("neumaticos").update({ camion_id: null, acoplado_id: null, posicion: null }).eq("id", ocupante.id);
      await supabase.from("neumatico_historial").insert({ neumatico_id: ocupante.id, tipo: "cambio_vehiculo", detalle: `Removido de ${destino} · ${label} (reemplazado por otro neumático)` });
    }

    const { error } = await supabase.from("neumaticos").update({ camion_id: camionId, acoplado_id: acopladoId, posicion: posicionFinal }).eq("id", id);
    if (error) return showToast("Error: " + error.message);
    const detalle = destino ? `Montado en ${destino}${label ? ` · ${label}` : " (sin posición específica)"}` : "Desmontado, sin asignar";
    await supabase.from("neumatico_historial").insert({ neumatico_id: id, tipo: "cambio_vehiculo", detalle });

    setNeumaticos(neumaticos.map((n) => {
      if (n.id === id) return { ...n, camionId, acopladoId, posicion: posicionFinal, historial: [...n.historial, { fecha: todayISO(), tipo: "cambio_vehiculo", detalle }] };
      if (n.id === ocupante?.id) return { ...n, camionId: null, acopladoId: null, posicion: null, historial: [...n.historial, { fecha: todayISO(), tipo: "cambio_vehiculo", detalle: `Removido de ${destino} · ${label} (reemplazado por otro neumático)` }] };
      return n;
    }));
    showToast(destino ? `Neumático movido a ${destino}${label ? ` · ${label}` : ""}.` : "Neumático desmontado, quedó sin asignar.");
  }

  async function asignarAPosicion(tipo, vehiculoId, codigo, neumaticoId) {
    const camionId = tipo === "camion" ? vehiculoId : null;
    const acopladoId = tipo === "acoplado" ? vehiculoId : null;
    const patente = tipo === "camion" ? camiones.find((c) => c.id === vehiculoId)?.patente : acoplados.find((a) => a.id === vehiculoId)?.patente;
    const label = POSICION_LABELS[codigo] || codigo;

    const ocupante = neumaticos.find((n) => n.id !== neumaticoId && n.posicion === codigo &&
      ((tipo === "camion" && n.camionId === vehiculoId) || (tipo === "acoplado" && n.acopladoId === vehiculoId)));
    if (ocupante) {
      await supabase.from("neumaticos").update({ camion_id: null, acoplado_id: null, posicion: null }).eq("id", ocupante.id);
      await supabase.from("neumatico_historial").insert({ neumatico_id: ocupante.id, tipo: "cambio_vehiculo", detalle: `Removido de ${patente} · ${label} (reemplazado por otro neumático)` });
    }

    const { error } = await supabase.from("neumaticos").update({ camion_id: camionId, acoplado_id: acopladoId, posicion: codigo }).eq("id", neumaticoId);
    if (error) return showToast("Error: " + error.message);
    await supabase.from("neumatico_historial").insert({ neumatico_id: neumaticoId, tipo: "cambio_vehiculo", detalle: `Montado en ${patente} · ${label}` });

    setNeumaticos(neumaticos.map((n) => {
      if (n.id === neumaticoId) return { ...n, camionId, acopladoId, posicion: codigo, historial: [...n.historial, { fecha: todayISO(), tipo: "cambio_vehiculo", detalle: `Montado en ${patente} · ${label}` }] };
      if (n.id === ocupante?.id) return { ...n, camionId: null, acopladoId: null, posicion: null, historial: [...n.historial, { fecha: todayISO(), tipo: "cambio_vehiculo", detalle: `Removido de ${patente} · ${label} (reemplazado por otro neumático)` }] };
      return n;
    }));
    showToast(`Neumático asignado a ${patente} · ${label}.`);
  }

  async function quitarDePosicion(neumaticoId) {
    const { error } = await supabase.from("neumaticos").update({ camion_id: null, acoplado_id: null, posicion: null }).eq("id", neumaticoId);
    if (error) return showToast("Error: " + error.message);
    await supabase.from("neumatico_historial").insert({ neumatico_id: neumaticoId, tipo: "cambio_vehiculo", detalle: "Desmontado de su posición, sin asignar" });
    setNeumaticos(neumaticos.map((n) =>
      n.id === neumaticoId
        ? { ...n, camionId: null, acopladoId: null, posicion: null, historial: [...n.historial, { fecha: todayISO(), tipo: "cambio_vehiculo", detalle: "Desmontado de su posición, sin asignar" }] }
        : n
    ));
    showToast("Neumático desmontado de esa posición.");
  }

  async function agregarService(tipo, vehiculoId, km, detalle, fecha) {
    const kmService = Math.max(0, parseInt(km, 10) || 0);
    const detalleFinal = detalle || "Sin detalle";
    const fechaFinal = fecha || todayISO();
    const { data, error } = await supabase.from("servicios").insert({
      [tipo === "camion" ? "camion_id" : "acoplado_id"]: vehiculoId,
      km: kmService, detalle: detalleFinal, editado: false, fecha: fechaFinal,
    }).select().single();
    if (error) return showToast("Error: " + error.message);
    const { error: e2 } = await supabase.from(tipo === "camion" ? "camiones" : "acoplados").update({ ultimo_service_km: kmService }).eq("id", vehiculoId);
    if (e2) return showToast("Error: " + e2.message);
    const entrada = mapServicioDb(data);
    if (tipo === "camion") {
      setCamiones(camiones.map((c) => (c.id === vehiculoId ? { ...c, ultimoServiceKm: kmService, servicios: [entrada, ...c.servicios] } : c)));
    } else {
      setAcoplados(acoplados.map((a) => (a.id === vehiculoId ? { ...a, ultimoServiceKm: kmService, servicios: [entrada, ...a.servicios] } : a)));
    }
    showToast("Service registrado.");
  }

  async function editarServicio(tipo, vehiculoId, servicioId, cambios) {
    const { error } = await supabase.from("servicios").update({
      detalle: cambios.detalle, km: cambios.km, fecha: cambios.fecha, editado: true, fecha_edicion: todayISO(),
    }).eq("id", servicioId);
    if (error) return showToast("Error: " + error.message);
    const patch = (list) => list.map((s) => (s.id === servicioId ? { ...s, ...cambios, editado: true, fechaEdicion: todayISO() } : s));
    if (tipo === "camion") {
      setCamiones(camiones.map((c) => (c.id === vehiculoId ? { ...c, servicios: patch(c.servicios) } : c)));
    } else {
      setAcoplados(acoplados.map((a) => (a.id === vehiculoId ? { ...a, servicios: patch(a.servicios) } : a)));
    }
    showToast("Registro de service editado.");
  }

  async function editarVehiculo(tipo, vehiculoId, cambios) {
    // Si se corrige el "km actuales" a mano, recalculamos el punto de partida
    // (km_inicial) para que el km real (inicial + viajes) siga dando ese valor.
    const sumaDeViajes = sumaViajes(historialViajes, (v) => (tipo === "camion" ? v.camionId : v.acopladoId) === vehiculoId);
    const kmInicialNuevo = cambios.km - sumaDeViajes;
    const patch = tipo === "camion"
      ? { patente: cambios.patente, modelo: cambios.modelo, km: cambios.km, km_inicial: kmInicialNuevo }
      : { patente: cambios.patente, tipo: cambios.tipo, km: cambios.km, km_inicial: kmInicialNuevo };
    const { error } = await supabase.from(tipo === "camion" ? "camiones" : "acoplados").update(patch).eq("id", vehiculoId);
    if (error) return showToast("Error: " + error.message);
    if (tipo === "camion") {
      setCamiones(camiones.map((c) => (c.id === vehiculoId ? { ...c, ...cambios, kmInicial: kmInicialNuevo } : c)));
    } else {
      setAcoplados(acoplados.map((a) => (a.id === vehiculoId ? { ...a, ...cambios, kmInicial: kmInicialNuevo } : a)));
    }
    showToast("Datos actualizados.");
  }

  async function subirPoliza(camionId, file) {
    const path = `${camionId}/${Date.now()}-${file.name}`;
    const { error: uploadError } = await supabase.storage.from("polizas").upload(path, file, { upsert: true });
    if (uploadError) return showToast("Error al subir el archivo: " + uploadError.message);
    const { error } = await supabase.from("camiones").update({ poliza_url: path, poliza_nombre: file.name }).eq("id", camionId);
    if (error) return showToast("Error: " + error.message);
    setCamiones(camiones.map((c) => (c.id === camionId ? { ...c, polizaUrl: path, polizaNombre: file.name } : c)));
    showToast("Póliza subida.");
  }

  async function descargarPoliza(path, nombre) {
    const { data, error } = await supabase.storage.from("polizas").createSignedUrl(path, 60);
    if (error || !data) return showToast("Error al generar el link de descarga.");
    const a = document.createElement("a");
    a.href = data.signedUrl;
    a.download = nombre || "poliza";
    a.target = "_blank";
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  async function eliminarPoliza(camionId, path) {
    await supabase.storage.from("polizas").remove([path]);
    const { error } = await supabase.from("camiones").update({ poliza_url: null, poliza_nombre: null }).eq("id", camionId);
    if (error) return showToast("Error: " + error.message);
    setCamiones(camiones.map((c) => (c.id === camionId ? { ...c, polizaUrl: null, polizaNombre: null } : c)));
    showToast("Póliza eliminada.");
  }

  async function editarChofer(choferId, cambios) {
    const { error } = await supabase.from("choferes").update({
      nombre: cambios.nombre, camion_id: cambios.camionId, acoplado_id: cambios.acopladoId,
      vencimiento_dni: cambios.vencimientoDni || null,
      vencimiento_licencia: cambios.vencimientoLicencia || null,
      fecha_curso_defensivo: cambios.fechaCursoDefensivo || null,
      fecha_estudios_medicos: cambios.fechaEstudiosMedicos || null,
      fecha_svo: cambios.fechaSvo || null,
      fecha_art: cambios.fechaArt || null,
      fecha_eepp: cambios.fechaEepp || null,
    }).eq("id", choferId);
    if (error) return showToast("Error: " + error.message);
    setChoferes(choferes.map((d) => (d.id === choferId ? { ...d, ...cambios } : d)));
    showToast("Datos del chofer actualizados.");
  }

  async function editarNeumatico(neumaticoId, cambios) {
    const { error } = await supabase.from("neumaticos").update({ etiqueta: cambios.etiqueta, estado: cambios.estado, km: cambios.km, marca: cambios.marca || null, tipo: cambios.tipo || null }).eq("id", neumaticoId);
    if (error) return showToast("Error: " + error.message);
    setNeumaticos(neumaticos.map((n) => (n.id === neumaticoId ? { ...n, ...cambios } : n)));
    showToast("Datos del neumático actualizados.");
  }

  async function eliminarCamion(id) {
    const c = camiones.find((x) => x.id === id);
    const { error } = await supabase.from("camiones").delete().eq("id", id);
    if (error) return showToast("Error: " + error.message);
    setCamiones(camiones.filter((x) => x.id !== id));
    setChoferes(choferes.map((d) => (d.camionId === id ? { ...d, camionId: null } : d)));
    setNeumaticos(neumaticos.map((n) =>
      n.camionId === id
        ? { ...n, camionId: null, posicion: null, historial: [...n.historial, { fecha: todayISO(), tipo: "cambio_vehiculo", detalle: `Desmontado: se eliminó el camión ${c?.patente || ""}` }] }
        : n
    ));
    setSelected(null);
    showToast(`Camión ${c?.patente || ""} eliminado.`);
  }

  async function eliminarAcoplado(id) {
    const a = acoplados.find((x) => x.id === id);
    const { error } = await supabase.from("acoplados").delete().eq("id", id);
    if (error) return showToast("Error: " + error.message);
    setAcoplados(acoplados.filter((x) => x.id !== id));
    setChoferes(choferes.map((d) => (d.acopladoId === id ? { ...d, acopladoId: null } : d)));
    setNeumaticos(neumaticos.map((n) =>
      n.acopladoId === id
        ? { ...n, acopladoId: null, posicion: null, historial: [...n.historial, { fecha: todayISO(), tipo: "cambio_vehiculo", detalle: `Desmontado: se eliminó el acoplado ${a?.patente || ""}` }] }
        : n
    ));
    setSelected(null);
    showToast(`Acoplado ${a?.patente || ""} eliminado.`);
  }

  async function eliminarChofer(id) {
    const d = choferes.find((x) => x.id === id);
    const { error } = await supabase.from("choferes").delete().eq("id", id);
    if (error) return showToast("Error: " + error.message);
    setChoferes(choferes.filter((x) => x.id !== id));
    setSelected(null);
    showToast(`Chofer ${d?.nombre || ""} eliminado.`);
  }

  async function eliminarNeumatico(id) {
    const n = neumaticos.find((x) => x.id === id);
    const { error } = await supabase.from("neumaticos").delete().eq("id", id);
    if (error) return showToast("Error: " + error.message);
    setNeumaticos(neumaticos.filter((x) => x.id !== id));
    setSelected(null);
    showToast(`Neumático ${n?.etiqueta || ""} eliminado.`);
  }

  async function actualizarVtv(tipo, vehiculoId, fecha) {
    const { error } = await supabase.from(tipo === "camion" ? "camiones" : "acoplados").update({ vtv_ultima: fecha }).eq("id", vehiculoId);
    if (error) return showToast("Error: " + error.message);
    if (tipo === "camion") {
      setCamiones(camiones.map((c) => (c.id === vehiculoId ? { ...c, vtvUltima: fecha } : c)));
    } else {
      setAcoplados(acoplados.map((a) => (a.id === vehiculoId ? { ...a, vtvUltima: fecha } : a)));
    }
    showToast("Fecha de VTV actualizada.");
  }

  async function actualizarIntervaloService(tipo, vehiculoId, intervalo) {
    const valor = Math.max(1, parseInt(intervalo, 10) || SERVICE_INTERVAL_DEFAULT);
    const { error } = await supabase.from(tipo === "camion" ? "camiones" : "acoplados").update({ intervalo_service: valor }).eq("id", vehiculoId);
    if (error) return showToast("Error: " + error.message);
    if (tipo === "camion") {
      setCamiones(camiones.map((c) => (c.id === vehiculoId ? { ...c, intervaloService: valor } : c)));
    } else {
      setAcoplados(acoplados.map((a) => (a.id === vehiculoId ? { ...a, intervaloService: valor } : a)));
    }
    showToast("Intervalo de service actualizado.");
  }

  async function actualizarValvulasTolva(acopladoId, valvula1, valvula2) {
    const { error } = await supabase.from("acoplados").update({ venc_valvula1: valvula1 || null, venc_valvula2: valvula2 || null }).eq("id", acopladoId);
    if (error) return showToast("Error: " + error.message);
    setAcoplados(acoplados.map((a) => (a.id === acopladoId ? { ...a, vencValvula1: valvula1 || null, vencValvula2: valvula2 || null } : a)));
    showToast("Vencimientos de válvulas actualizados.");
  }

  async function actualizarIntercement(tipo, vehiculoId, fecha) {
    const { error } = await supabase.from(tipo === "camion" ? "camiones" : "acoplados").update({ vencimiento_intercement: fecha || null }).eq("id", vehiculoId);
    if (error) return showToast("Error: " + error.message);
    if (tipo === "camion") {
      setCamiones(camiones.map((c) => (c.id === vehiculoId ? { ...c, vencimientoIntercement: fecha || null } : c)));
    } else {
      setAcoplados(acoplados.map((a) => (a.id === vehiculoId ? { ...a, vencimientoIntercement: fecha || null } : a)));
    }
    showToast("Vencimiento de Intercement actualizado.");
  }

  async function saveRuta() {
    if (!fRuta.nombre.trim()) return showToast("Falta el nombre de la ruta.");
    const km = Math.max(0, parseInt(fRuta.km, 10) || 0);
    if (!km) return showToast("Cargá los km de la ruta.");
    const { data, error } = await supabase.from("rutas").insert({ nombre: fRuta.nombre, km }).select().single();
    if (error) return showToast("Error: " + error.message);
    setRutas([...rutas, data]);
    setFRuta({ nombre: "", km: "" });
    showToast(`Ruta "${fRuta.nombre}" guardada.`);
  }

  async function deleteRuta(id) {
    const { error } = await supabase.from("rutas").delete().eq("id", id);
    if (error) return showToast("Error: " + error.message);
    setRutas(rutas.filter((r) => r.id !== id));
  }

  async function registrarViaje(choferId, rutaId, overrideCamionId, overrideAcopladoId, usarOtro, codigo, fecha, remito) {
    const ruta = rutas.find((r) => r.id === rutaId);
    const chofer = choferes.find((d) => d.id === choferId);
    if (!ruta) return showToast("Elegí una ruta.");
    if (!chofer) return showToast("Chofer inválido.");
    const camionDefault = camiones.find((c) => c.id === chofer.camionId);
    const camion = usarOtro && overrideCamionId ? camiones.find((c) => c.id === overrideCamionId) : camionDefault;
    const acopladoDefault = acoplados.find((a) => a.id === chofer.acopladoId);
    const acoplado = usarOtro && overrideAcopladoId ? acoplados.find((a) => a.id === overrideAcopladoId) : acopladoDefault;

    const { data: viajeRow, error } = await supabase.from("viajes").insert({
      ruta_id: ruta.id, ruta_nombre: ruta.nombre, chofer_id: chofer.id,
      camion_id: camion?.id || null, acoplado_id: acoplado?.id || null,
      km: ruta.km, combo_distinto: usarOtro, codigo: codigo || null, fecha: fecha || todayISO(), remito: remito || null,
    }).select().single();
    if (error) return showToast("Error: " + error.message);

    // El km de chofer/camión/acoplado ya no se guarda aparte: se calcula siempre
    // sumando los viajes reales (ver kmRealChofer/kmRealCamion/kmRealAcoplado),
    // así nunca puede desincronizarse. Los neumáticos sí siguen con su contador propio.
    const neumaticosAfectados = neumaticos.filter((n) => (camion && n.camionId === camion.id) || (acoplado && n.acopladoId === acoplado.id));
    await Promise.all(neumaticosAfectados.map((n) => supabase.from("neumaticos").update({ km: n.km + ruta.km }).eq("id", n.id)));

    setNeumaticos(neumaticos.map((n) =>
      (camion && n.camionId === camion.id) || (acoplado && n.acopladoId === acoplado.id)
        ? { ...n, km: n.km + ruta.km }
        : n
    ));

    const afectados = [chofer.nombre, camion?.patente, acoplado?.patente].filter(Boolean).join(", ");
    setHistorialViajes([mapViajeDb(viajeRow), ...historialViajes]);
    showToast(`Viaje de ${ruta.km} km cargado. Se sumaron km a: ${afectados}.`);
  }

  async function editarViaje(viajeId, nuevoKm, nuevaFecha, nuevaRutaId, nuevaRutaNombre, nuevoRemito) {
    const viaje = historialViajes.find((v) => v.id === viajeId);
    if (!viaje) return;
    const kmNuevo = Math.max(0, parseInt(nuevoKm, 10) || 0);
    const delta = kmNuevo - viaje.km;
    const fechaNueva = nuevaFecha || viaje.fecha;
    const rutaIdNueva = nuevaRutaId || viaje.rutaId;
    const rutaNombreNueva = nuevaRutaNombre || viaje.ruta;
    const remitoNuevo = nuevoRemito !== undefined ? nuevoRemito : viaje.remito;
    if (delta === 0 && fechaNueva === viaje.fecha && rutaIdNueva === viaje.rutaId && remitoNuevo === viaje.remito) return;

    const { error } = await supabase.from("viajes").update({ km: kmNuevo, fecha: fechaNueva, ruta_id: rutaIdNueva, ruta_nombre: rutaNombreNueva, remito: remitoNuevo || null }).eq("id", viajeId);
    if (error) return showToast("Error: " + error.message);

    if (delta !== 0) {
      // Chofer/camión/acoplado: nada que actualizar, su km real se recalcula
      // solo a partir de los viajes. Solo los neumáticos siguen con su contador propio.
      const neumaticosAfectados = neumaticos.filter((n) => (viaje.camionId && n.camionId === viaje.camionId) || (viaje.acopladoId && n.acopladoId === viaje.acopladoId));
      await Promise.all(neumaticosAfectados.map((n) => supabase.from("neumaticos").update({ km: n.km + delta }).eq("id", n.id)));

      setNeumaticos(neumaticos.map((n) =>
        (viaje.camionId && n.camionId === viaje.camionId) || (viaje.acopladoId && n.acopladoId === viaje.acopladoId)
          ? { ...n, km: n.km + delta }
          : n
      ));
    }
    setHistorialViajes(historialViajes.map((v) => (v.id === viajeId ? { ...v, km: kmNuevo, fecha: fechaNueva, rutaId: rutaIdNueva, ruta: rutaNombreNueva, remito: remitoNuevo || null } : v)));
    showToast("Viaje corregido.");
  }

  async function eliminarViaje(viajeId) {
    const viaje = historialViajes.find((v) => v.id === viajeId);
    if (!viaje) return;
    const { error } = await supabase.from("viajes").delete().eq("id", viajeId);
    if (error) return showToast("Error: " + error.message);

    // Chofer/camión/acoplado: su km real se recalcula solo, no hay nada que tocar.
    // Los neumáticos siguen con su contador propio.
    const neumaticosAfectados = neumaticos.filter((n) => (viaje.camionId && n.camionId === viaje.camionId) || (viaje.acopladoId && n.acopladoId === viaje.acopladoId));
    await Promise.all(neumaticosAfectados.map((n) => supabase.from("neumaticos").update({ km: n.km - viaje.km }).eq("id", n.id)));

    setNeumaticos(neumaticos.map((n) =>
      (viaje.camionId && n.camionId === viaje.camionId) || (viaje.acopladoId && n.acopladoId === viaje.acopladoId)
        ? { ...n, km: n.km - viaje.km }
        : n
    ));
    setHistorialViajes(historialViajes.filter((v) => v.id !== viajeId));
    showToast("Viaje eliminado, se restaron sus km.");
  }

  if (!currentUser) {
    return <LoginScreen onLogin={handleLogin} error={loginError} loading={loginLoading} />;
  }

  const puede = (clave) => currentUser.role === "admin" || currentUser.permisos[clave];

  const camionesFiltrados = camiones
    .filter((c) => `${c.patente} ${c.modelo}`.toLowerCase().includes(busqueda.camiones.toLowerCase()))
    .sort((a, b) => a.patente.localeCompare(b.patente));
  const acopladosFiltrados = acoplados
    .filter((a) => `${a.patente} ${a.tipo}`.toLowerCase().includes(busqueda.acoplados.toLowerCase()))
    .sort((a, b) => a.patente.localeCompare(b.patente));
  const choferesFiltrados = choferes
    .filter((d) => `${d.nombre}`.toLowerCase().includes(busqueda.choferes.toLowerCase()))
    .sort((a, b) => a.nombre.localeCompare(b.nombre));
  const neumaticosFiltrados = neumaticos.filter((n) =>
    `${n.etiqueta} ${n.estado}`.toLowerCase().includes(busqueda.neumaticos.toLowerCase())
  );
  const rutasFiltradas = [...rutas]
    .filter((r) => r.nombre.toLowerCase().includes(busqueda.rutas ? busqueda.rutas.toLowerCase() : ""))
    .sort((a, b) => a.nombre.localeCompare(b.nombre));

  const renderCamion = (c) => (
    <RowItem key={c.id} onClick={() => setSelected({ type: "camion", id: c.id })} title={c.patente} subtitle={c.modelo}
      right={<span className="odometer font-mono text-sm">{fmtKm(kmRealCamion(c, historialViajes))} km</span>} />
  );
  const renderAcoplado = (a) => (
    <RowItem key={a.id} onClick={() => setSelected({ type: "acoplado", id: a.id })} title={a.patente} subtitle={a.tipo}
      right={<span className="odometer font-mono text-sm">{fmtKm(kmRealAcoplado(a, historialViajes))} km</span>} />
  );
  const renderChofer = (d) => (
    <RowItem key={d.id} onClick={() => setSelected({ type: "chofer", id: d.id })} title={d.nombre} subtitle="Chofer"
      right={<span className="odometer font-mono text-sm">{fmtKm(kmRealChofer(d, historialViajes))} km</span>} />
  );
  const renderNeumatico = (n) => {
    const label = estadoNeumatico(n);
    return (
      <RowItem
        key={n.id} onClick={() => setSelected({ type: "neumatico", id: n.id })}
        title={n.etiqueta}
        subtitle={<span style={{ color: colorEstadoNeumatico(label) }}>{label}</span>}
        right={<span className="odometer font-mono text-sm">{fmtKm(n.km)} km</span>}
      />
    );
  };

  return (
    <div
      style={{
        "--bg": "#0A1128",
        "--panel": "#101B36",
        "--panel-2": "#16264A",
        "--border": "#233258",
        "--text": "#FFFFFF",
        "--text-muted": "#93A7CC",
        "--accent-yellow": "#F2B705",
        "--accent-red": "#F87171",
        "--accent-green": "#34D399",
        "--accent-blue": "#F2B705",
        background: "var(--bg)",
        color: "var(--text)",
        minHeight: "100vh",
        fontFamily: "'Inter', system-ui, sans-serif",
      }}
      className="w-full"
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=JetBrains+Mono:wght@500;700&family=Inter:wght@400;500;600&display=swap');
        .font-display { font-family: 'Oswald', sans-serif; letter-spacing: 0.02em; }
        .font-mono { font-family: 'JetBrains Mono', monospace; }
        .odometer {
          background: #0E1013;
          border: 1px solid var(--border);
          border-radius: 4px;
          padding: 2px 8px;
          font-variant-numeric: tabular-nums;
          letter-spacing: 0.05em;
        }
        input, select, textarea {
          color: #0A1128 !important;
          color-scheme: light;
        }
        input::placeholder { color: #6B7690; opacity: 1; }
        select option { color: #0A1128; background: #FFFFFF; }
      `}</style>

      <div className="md:flex">
      {/* Mobile top bar */}
      <div className="md:hidden flex items-center justify-between px-4 py-3 fixed top-0 left-0 right-0 z-40" style={{ background: "var(--panel)", borderBottom: "1px solid var(--border)" }}>
        <button onClick={() => setMenuAbierto(true)} className="p-1">
          <Menu size={22} />
        </button>
        <span className="font-display uppercase text-sm tracking-wide">Control de flota</span>
        <div style={{ width: 30 }} />
      </div>

      {menuAbierto && (
        <div className="fixed inset-0 z-40 md:hidden" style={{ background: "rgba(0,0,0,0.6)" }} onClick={() => setMenuAbierto(false)} />
      )}

      {/* Sidebar */}
      <aside
        className={`fixed md:sticky top-0 left-0 h-screen z-50 md:z-0 w-64 flex-shrink-0 flex flex-col transition-transform duration-200 ${menuAbierto ? "translate-x-0" : "-translate-x-full"} md:translate-x-0`}
        style={{ background: "var(--panel)", borderRight: "1px solid var(--border)" }}
      >
        <div className="p-5 flex-1 overflow-y-auto">
          <div className="text-xs uppercase tracking-[0.2em] mb-1" style={{ color: "var(--accent-yellow)" }}>
            Control de flota
          </div>
          <h1 className="font-display text-xl font-semibold uppercase mb-6">Panel</h1>
          <nav className="flex flex-col gap-1">
            {SECCIONES.filter((s) => !s.permiso || puede(s.permiso)).map((s) => (
              <button
                key={s.key}
                onClick={() => { setSeccionActiva(s.key); setMenuAbierto(false); }}
                className="flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium text-left transition-colors"
                style={{
                  background: seccionActiva === s.key ? "var(--accent-yellow)" : "transparent",
                  color: seccionActiva === s.key ? "#0A1128" : "var(--text)",
                }}
              >
                <s.icon size={16} /> {s.label}
              </button>
            ))}
          </nav>
        </div>
        <div className="p-5" style={{ borderTop: "1px solid var(--border)" }}>
          {dataLoading && <div className="text-xs mb-2" style={{ color: "var(--text-muted)" }}>Cargando datos…</div>}
          <div className="text-xs font-mono odometer inline-block mb-2" style={{ color: "var(--accent-green)" }}>
            {new Date().toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric" })}
          </div>
          <div className="text-xs" style={{ color: "var(--text-muted)" }}>
            {currentUser.username} {currentUser.role === "admin" ? "(admin)" : ""}
          </div>
          <button onClick={handleLogout} className="text-xs underline" style={{ color: "var(--accent-yellow)" }}>
            Cerrar sesión
          </button>
        </div>
      </aside>

      {/* Main content */}
      <div className="flex-1 min-w-0 max-w-6xl mx-auto px-5 pt-20 pb-8 md:px-8 md:pt-8">
        <div className="mb-6">
          <div className="text-xs uppercase tracking-[0.2em] mb-1" style={{ color: "var(--accent-yellow)" }}>
            Control de flota
          </div>
          <h1 className="font-display text-3xl md:text-4xl font-semibold uppercase">
            {SECCIONES.find((s) => s.key === seccionActiva)?.label}
          </h1>
        </div>

        {seccionActiva === "general" && (
          <>
        {/* KPIs */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
          {kpis.map((k) => (
            <div
              key={k.label}
              className="rounded-lg p-4 flex items-center gap-3"
              style={{ background: "var(--panel)", border: "1px solid var(--border)" }}
            >
              <div className="p-2 rounded-md" style={{ background: "var(--panel-2)" }}>
                <k.icon size={20} style={{ color: "var(--accent-yellow)" }} />
              </div>
              <div>
                <div className="font-mono text-2xl font-semibold leading-none">{k.value}</div>
                <div className="text-xs mt-1" style={{ color: "var(--text-muted)" }}>{k.label}</div>
              </div>
            </div>
          ))}
        </div>

        {/* Alerts */}
        <div className="rounded-lg mb-6" style={{ background: "var(--panel)", border: "1px solid var(--border)" }}>
          <div className="flex items-center gap-2 px-4 py-3" style={{ borderBottom: "1px solid var(--border)" }}>
            <AlertTriangle size={16} style={{ color: "var(--accent-red)" }} />
            <h2 className="font-display text-sm uppercase tracking-wide">Alertas</h2>
            <span className="ml-auto text-xs font-mono" style={{ color: "var(--text-muted)" }}>{alerts.length}</span>
          </div>
          <div className="p-2">
            {alerts.length === 0 ? (
              <div className="flex items-center gap-2 px-3 py-4 text-sm" style={{ color: "var(--text-muted)" }}>
                <CheckCircle2 size={16} style={{ color: "var(--accent-green)" }} />
                Sin alertas pendientes.
              </div>
            ) : (
              <>
                {alerts.slice(0, 3).map((a, i) => (
                  <div key={i} className="flex items-center gap-3 px-3 py-2.5 rounded-md" style={{ background: i % 2 ? "transparent" : "var(--panel-2)" }}>
                    <a.icon size={16} style={{ color: a.severity === "danger" ? "var(--accent-red)" : "var(--accent-yellow)", flexShrink: 0 }} />
                    <div className="min-w-0">
                      <div className="text-sm font-medium truncate">{a.title}</div>
                      <div className="text-xs" style={{ color: "var(--text-muted)" }}>{a.detail}</div>
                    </div>
                  </div>
                ))}
                {alerts.length > 3 && (
                  <button
                    onClick={() => setVerTodasAlertas(true)}
                    className="w-full mt-1 px-3 py-2 rounded-md text-xs font-semibold text-center"
                    style={{ background: "var(--panel-2)", border: "1px solid var(--border)", color: "var(--accent-yellow)" }}
                  >
                    Ver todas ({alerts.length})
                  </button>
                )}
              </>
            )}
          </div>
        </div>

        {verTodasAlertas && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(0,0,0,0.6)" }} onClick={() => setVerTodasAlertas(false)}>
            <div
              onClick={(e) => e.stopPropagation()}
              className="w-full max-w-lg rounded-lg flex flex-col"
              style={{ background: "var(--panel)", border: "1px solid var(--border)", maxHeight: "80vh" }}
            >
              <div className="flex items-center gap-2 px-4 py-3" style={{ borderBottom: "1px solid var(--border)" }}>
                <AlertTriangle size={16} style={{ color: "var(--accent-red)" }} />
                <h3 className="font-display text-sm uppercase tracking-wide">Todas las alertas</h3>
                <span className="ml-auto text-xs font-mono mr-2" style={{ color: "var(--text-muted)" }}>{alerts.length}</span>
                <button onClick={() => setVerTodasAlertas(false)} className="p-1 rounded hover:bg-[var(--panel-2)]">
                  <X size={16} style={{ color: "var(--text-muted)" }} />
                </button>
              </div>
              <div className="p-2 overflow-y-auto">
                {alerts.map((a, i) => (
                  <div key={i} className="flex items-center gap-3 px-3 py-2.5 rounded-md" style={{ background: i % 2 ? "transparent" : "var(--panel-2)" }}>
                    <a.icon size={16} style={{ color: a.severity === "danger" ? "var(--accent-red)" : "var(--accent-yellow)", flexShrink: 0 }} />
                    <div className="min-w-0">
                      <div className="text-sm font-medium truncate">{a.title}</div>
                      <div className="text-xs" style={{ color: "var(--text-muted)" }}>{a.detail}</div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* Buscador global de viajes por código */}
        {puede("choferes") && (
          <div className="rounded-lg mb-6" style={{ background: "var(--panel)", border: "1px solid var(--border)" }}>
            <div className="flex items-center gap-2 px-4 py-3" style={{ borderBottom: "1px solid var(--border)" }}>
              <Search size={15} style={{ color: "var(--accent-blue)" }} />
              <h2 className="font-display text-sm uppercase tracking-wide">Buscar viaje por código</h2>
            </div>
            <div className="p-3">
              <input
                className={inputCls + " w-full"}
                placeholder="Ej: V-00123"
                value={busquedaCodigo}
                onChange={(e) => setBusquedaCodigo(e.target.value)}
              />
              {busquedaCodigo.trim() && (() => {
                const resultados = historialViajes.filter((v) =>
                  v.codigo && v.codigo.toLowerCase().includes(busquedaCodigo.trim().toLowerCase())
                );
                if (resultados.length === 0) {
                  return <p className="text-sm mt-3" style={{ color: "var(--text-muted)" }}>Sin resultados para ese código.</p>;
                }
                return (
                  <div className="mt-3">
                    {resultados.map((v) => {
                      const chofer = choferes.find((d) => d.id === v.choferId);
                      const camion = camiones.find((c) => c.id === v.camionId);
                      const acoplado = acoplados.find((a) => a.id === v.acopladoId);
                      return (
                        <div
                          key={v.id}
                          onClick={() => chofer && setSelected({ type: "chofer", id: chofer.id })}
                          className="flex items-center gap-3 py-2.5 cursor-pointer hover:bg-[var(--panel-2)] rounded-md px-2"
                          style={{ borderBottom: "1px solid var(--border)" }}
                        >
                          <div className="min-w-0 flex-1">
                            <div className="text-sm font-semibold">{v.codigo} · {v.ruta}</div>
                            <div className="text-xs" style={{ color: "var(--text-muted)" }}>
                              {chofer?.nombre || "Chofer eliminado"} · {camion?.patente || "sin camión"} · {acoplado?.patente || "sin acoplado"} · {v.fecha}
                            </div>
                          </div>
                          <span className="odometer font-mono text-sm">{fmtKm(v.km)} km</span>
                        </div>
                      );
                    })}
                  </div>
                );
              })()}
            </div>
          </div>
        )}

        {/* Viajes: filtro por mes y/o chofer, con total de km y descarga en Excel */}
        {puede("choferes") && (
          <div className="rounded-lg mb-6" style={{ background: "var(--panel)", border: "1px solid var(--border)" }}>
            <div className="flex items-center gap-2 px-4 py-3" style={{ borderBottom: "1px solid var(--border)" }}>
              <CalendarClock size={15} style={{ color: "var(--accent-blue)" }} />
              <h2 className="font-display text-sm uppercase tracking-wide">Viajes</h2>
            </div>
            <div className="p-3">
              <div className="grid sm:grid-cols-2 gap-3 mb-3">
                <Field label="Mes">
                  <select className={inputCls} value={mesSeleccionado} onChange={(e) => setMesSeleccionado(e.target.value)}>
                    <option value="">Todos los meses</option>
                    {[...new Set(historialViajes.map((v) => mesKey(v.fecha)))].sort().reverse().map((key) => (
                      <option key={key} value={key}>{mesLabel(key)}</option>
                    ))}
                  </select>
                </Field>
                <Field label="Chofer">
                  <select className={inputCls} value={choferSeleccionado} onChange={(e) => setChoferSeleccionado(e.target.value)}>
                    <option value="">Todos los choferes</option>
                    {[...choferes].sort((a, b) => a.nombre.localeCompare(b.nombre)).map((d) => (
                      <option key={d.id} value={d.id}>{d.nombre}</option>
                    ))}
                  </select>
                </Field>
              </div>

              {(mesSeleccionado || choferSeleccionado) && (() => {
                const viajesFiltrados = historialViajes
                  .filter((v) => !mesSeleccionado || mesKey(v.fecha) === mesSeleccionado)
                  .filter((v) => !choferSeleccionado || v.choferId === choferSeleccionado)
                  .sort((a, b) => a.fecha.localeCompare(b.fecha));
                const totalKm = viajesFiltrados.reduce((acc, v) => acc + v.km, 0);
                const nombreArchivo = [mesSeleccionado, choferSeleccionado ? choferes.find((d) => d.id === choferSeleccionado)?.nombre.replace(/\s+/g, "-") : null].filter(Boolean).join("-") || "todos";

                return (
                  <div>
                    <div className="flex items-center justify-between flex-wrap gap-2 mb-3">
                      <span className="text-sm" style={{ color: "var(--text-muted)" }}>
                        {viajesFiltrados.length} viajes · <span className="font-mono odometer" style={{ color: "var(--accent-green)" }}>{fmtKm(totalKm)} km</span> acumulados
                      </span>
                      <button
                        onClick={() => descargarExcelMes(nombreArchivo, viajesFiltrados, choferes, camiones, acoplados)}
                        className="px-3 py-2 rounded-md text-xs font-semibold flex items-center gap-1.5"
                        style={{ background: "var(--accent-yellow)", color: "#0A1128" }}
                      >
                        <Download size={13} /> Descargar Excel
                      </button>
                    </div>
                    {viajesFiltrados.length === 0 ? (
                      <p className="text-sm" style={{ color: "var(--text-muted)" }}>Sin viajes con ese filtro.</p>
                    ) : viajesFiltrados.map((v) => {
                      const chofer = choferes.find((d) => d.id === v.choferId);
                      const camion = camiones.find((c) => c.id === v.camionId);
                      const acoplado = acoplados.find((a) => a.id === v.acopladoId);
                      return (
                        <div
                          key={v.id}
                          onClick={() => chofer && setSelected({ type: "chofer", id: chofer.id })}
                          className="flex items-center gap-3 py-2.5 cursor-pointer hover:bg-[var(--panel-2)] rounded-md px-2"
                          style={{ borderBottom: "1px solid var(--border)" }}
                        >
                          <div className="min-w-0 flex-1">
                            <div className="text-sm font-semibold">
                              {v.ruta}{v.codigo ? ` · ${v.codigo}` : ""}{v.remito ? ` · Rem. ${v.remito}` : ""}
                            </div>
                            <div className="text-xs" style={{ color: "var(--text-muted)" }}>
                              {chofer?.nombre || "Chofer eliminado"} · {camion?.patente || "sin camión"} · {acoplado?.patente || "sin acoplado"} · {v.fecha}
                            </div>
                          </div>
                          <span className="odometer font-mono text-sm">{fmtKm(v.km)} km</span>
                        </div>
                      );
                    })}
                  </div>
                );
              })()}
            </div>
          </div>
        )}

        {puede("choferes") && (
          <GraficoKmChoferes choferes={choferes} historialViajes={historialViajes} onVerChofer={(id) => setSelected({ type: "chofer", id })} />
        )}
          </>
        )}


        {seccionActiva === "rutas" && (
          <div>
            {puede("rutas") && (
              <div className="flex justify-end mb-4">
                <button onClick={() => openQuickAdd(activeQuickAdd === "ruta" ? null : "ruta")} className="px-4 py-2 rounded-md text-sm font-semibold flex items-center gap-1.5" style={{ background: "var(--accent-yellow)", color: "#0A1128" }}>
                  {activeQuickAdd === "ruta" ? <X size={16} /> : <Plus size={16} />} {activeQuickAdd === "ruta" ? "Cerrar" : "Cargar ruta nueva"}
                </button>
              </div>
            )}
            {activeQuickAdd === "ruta" && (
              <div className="rounded-lg p-5 mb-6" style={{ background: "var(--panel)", border: "1px solid var(--accent-yellow)" }}>
                <div className="grid gap-4">
                  <div className="grid sm:grid-cols-[1fr_140px] gap-4">
                    <Field label="Nombre de la ruta">
                      <input className={inputCls} value={fRuta.nombre} onChange={(e) => setFRuta({ ...fRuta, nombre: e.target.value })} placeholder="Buenos Aires → Rosario" />
                    </Field>
                    <Field label="Km">
                      <input className={inputCls} type="number" min="0" value={fRuta.km} onChange={(e) => setFRuta({ ...fRuta, km: e.target.value })} placeholder="320" />
                    </Field>
                  </div>
                  <button onClick={saveRuta} className="justify-self-start px-4 py-2 rounded-md text-sm font-semibold flex items-center gap-1" style={{ background: "var(--accent-yellow)", color: "#14161A" }}>
                    <Plus size={15} /> Guardar ruta
                  </button>
                </div>
              </div>
            )}
            <div className="rounded-lg" style={{ background: "var(--panel)", border: "1px solid var(--border)" }}>
              <div className="p-3">
                <div className="relative">
                  <Search size={14} style={{ color: "var(--text-muted)", position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
                  <input className={inputCls + " w-full pl-8"} placeholder="Buscar ruta…" value={busqueda.rutas} onChange={(e) => setBusqueda({ ...busqueda, rutas: e.target.value })} />
                </div>
              </div>
              <div className="p-2">
                {rutasFiltradas.length === 0 ? (
                  <p className="text-sm px-3 py-4" style={{ color: "var(--text-muted)" }}>Sin resultados.</p>
                ) : rutasFiltradas.map((r) => (
                  <div key={r.id} className="flex items-center gap-3 px-3 py-3" style={{ borderBottom: "1px solid var(--border)" }}>
                    <Route size={14} style={{ color: "var(--accent-blue)" }} />
                    <span className="flex-1 text-sm">{r.nombre}</span>
                    <span className="odometer font-mono text-sm">{fmtKm(r.km)} km</span>
                    {currentUser.role === "admin" && (
                      <button onClick={() => deleteRuta(r.id)} className="p-1.5 rounded hover:bg-[var(--panel-2)]">
                        <Trash2 size={14} style={{ color: "var(--accent-red)" }} />
                      </button>
                    )}
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {seccionActiva === "camiones" && (
          <div>
            {puede("camiones") && (
              <div className="flex justify-end mb-4">
                <button onClick={() => openQuickAdd(activeQuickAdd === "camion" ? null : "camion")} className="px-4 py-2 rounded-md text-sm font-semibold flex items-center gap-1.5" style={{ background: "var(--accent-yellow)", color: "#0A1128" }}>
                  {activeQuickAdd === "camion" ? <X size={16} /> : <Plus size={16} />} {activeQuickAdd === "camion" ? "Cerrar" : "Agregar camión"}
                </button>
              </div>
            )}
            {activeQuickAdd === "camion" && (
              <div className="rounded-lg p-5 mb-6" style={{ background: "var(--panel)", border: "1px solid var(--accent-yellow)" }}>
                <div className="grid gap-4">
                  <div className="grid sm:grid-cols-2 gap-4">
                    <Field label="Patente">
                      <input className={inputCls} value={fCamion.patente} onChange={(e) => setFCamion({ ...fCamion, patente: e.target.value })} placeholder="AB123CD" />
                    </Field>
                    <Field label="Modelo">
                      <input className={inputCls} value={fCamion.modelo} onChange={(e) => setFCamion({ ...fCamion, modelo: e.target.value })} placeholder="Scania R450" />
                    </Field>
                    <Field label="Km actuales">
                      <input type="number" min="0" className={inputCls} value={fCamion.km} onChange={(e) => setFCamion({ ...fCamion, km: e.target.value })} placeholder="0" />
                    </Field>
                    <Field label="Fecha de la última VTV">
                      <input type="date" className={inputCls} value={fCamion.vtvUltima} onChange={(e) => setFCamion({ ...fCamion, vtvUltima: e.target.value })} />
                    </Field>
                    <Field label="Km del último service">
                      <input type="number" min="0" className={inputCls} value={fCamion.ultimoServiceKm} onChange={(e) => setFCamion({ ...fCamion, ultimoServiceKm: e.target.value })} placeholder="0" />
                    </Field>
                    <Field label="Cada cuántos km hacer service">
                      <input type="number" min="1" className={inputCls} value={fCamion.intervaloService} onChange={(e) => setFCamion({ ...fCamion, intervaloService: e.target.value })} placeholder="50000" />
                    </Field>
                    <Field label="Verificación Adicional Intercement">
                      <input type="date" className={inputCls} value={fCamion.vencimientoIntercement} onChange={(e) => setFCamion({ ...fCamion, vencimientoIntercement: e.target.value })} />
                    </Field>
                  </div>
                  <p className="text-xs" style={{ color: "var(--text-muted)" }}>
                    La VTV vence al año de la fecha cargada (avisa un mes antes). El aviso de service aparece cuando falten 5.000 km para el intervalo elegido (por defecto 50.000 km).
                  </p>
                  <button onClick={saveCamion} className="justify-self-start px-4 py-2 rounded-md text-sm font-semibold" style={{ background: "var(--accent-yellow)", color: "#14161A" }}>
                    Guardar camión
                  </button>
                </div>
              </div>
            )}
            <div className="rounded-lg" style={{ background: "var(--panel)", border: "1px solid var(--border)" }}>
              <div className="p-3">
                <div className="relative">
                  <Search size={14} style={{ color: "var(--text-muted)", position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
                  <input className={inputCls + " w-full pl-8"} placeholder="Buscar por patente o modelo…" value={busqueda.camiones} onChange={(e) => setBusqueda({ ...busqueda, camiones: e.target.value })} />
                </div>
              </div>
              <div className="p-2">
                {camionesFiltrados.length === 0 ? (
                  <p className="text-sm px-3 py-4" style={{ color: "var(--text-muted)" }}>Sin resultados.</p>
                ) : camionesFiltrados.map(renderCamion)}
              </div>
            </div>
          </div>
        )}

        {seccionActiva === "acoplados" && (
          <div>
            {puede("acoplados") && (
              <div className="flex justify-end mb-4">
                <button onClick={() => openQuickAdd(activeQuickAdd === "acoplado" ? null : "acoplado")} className="px-4 py-2 rounded-md text-sm font-semibold flex items-center gap-1.5" style={{ background: "var(--accent-yellow)", color: "#0A1128" }}>
                  {activeQuickAdd === "acoplado" ? <X size={16} /> : <Plus size={16} />} {activeQuickAdd === "acoplado" ? "Cerrar" : "Agregar acoplado"}
                </button>
              </div>
            )}
            {activeQuickAdd === "acoplado" && (
              <div className="rounded-lg p-5 mb-6" style={{ background: "var(--panel)", border: "1px solid var(--accent-yellow)" }}>
                <div className="grid gap-4">
                  <div className="grid sm:grid-cols-2 gap-4">
                    <Field label="Patente">
                      <input className={inputCls} value={fAcoplado.patente} onChange={(e) => setFAcoplado({ ...fAcoplado, patente: e.target.value })} placeholder="TR501AA" />
                    </Field>
                    <Field label="Tipo">
                      <select className={inputCls} value={fAcoplado.tipo} onChange={(e) => setFAcoplado({ ...fAcoplado, tipo: e.target.value })}>
                        <option value="">Elegir tipo</option>
                        <option value="TOLVA">TOLVA</option>
                        <option value="SEMIRREMOLQUE">SEMIRREMOLQUE</option>
                        <option value="BATEA">BATEA</option>
                      </select>
                    </Field>
                    <Field label="Km actuales">
                      <input type="number" min="0" className={inputCls} value={fAcoplado.km} onChange={(e) => setFAcoplado({ ...fAcoplado, km: e.target.value })} placeholder="0" />
                    </Field>
                    <Field label="Fecha de la última VTV">
                      <input type="date" className={inputCls} value={fAcoplado.vtvUltima} onChange={(e) => setFAcoplado({ ...fAcoplado, vtvUltima: e.target.value })} />
                    </Field>
                    <Field label="Km del último service">
                      <input type="number" min="0" className={inputCls} value={fAcoplado.ultimoServiceKm} onChange={(e) => setFAcoplado({ ...fAcoplado, ultimoServiceKm: e.target.value })} placeholder="0" />
                    </Field>
                    <Field label="Cada cuántos km hacer service">
                      <input type="number" min="1" className={inputCls} value={fAcoplado.intervaloService} onChange={(e) => setFAcoplado({ ...fAcoplado, intervaloService: e.target.value })} placeholder="50000" />
                    </Field>
                    <Field label="Verificación Adicional Intercement">
                      <input type="date" className={inputCls} value={fAcoplado.vencimientoIntercement} onChange={(e) => setFAcoplado({ ...fAcoplado, vencimientoIntercement: e.target.value })} />
                    </Field>
                    {fAcoplado.tipo === "TOLVA" && (
                      <>
                        <Field label="Vencimiento válvula 1">
                          <input type="date" className={inputCls} value={fAcoplado.vencValvula1} onChange={(e) => setFAcoplado({ ...fAcoplado, vencValvula1: e.target.value })} />
                        </Field>
                        <Field label="Vencimiento válvula 2">
                          <input type="date" className={inputCls} value={fAcoplado.vencValvula2} onChange={(e) => setFAcoplado({ ...fAcoplado, vencValvula2: e.target.value })} />
                        </Field>
                      </>
                    )}
                  </div>
                  <p className="text-xs" style={{ color: "var(--text-muted)" }}>
                    La VTV vence al año de la fecha cargada (avisa un mes antes). El aviso de service aparece cuando falten 5.000 km para el intervalo elegido (por defecto 50.000 km).
                  </p>
                  <button onClick={saveAcoplado} className="justify-self-start px-4 py-2 rounded-md text-sm font-semibold" style={{ background: "var(--accent-yellow)", color: "#14161A" }}>
                    Guardar acoplado
                  </button>
                </div>
              </div>
            )}
            <div className="rounded-lg" style={{ background: "var(--panel)", border: "1px solid var(--border)" }}>
              <div className="p-3">
                <div className="relative">
                  <Search size={14} style={{ color: "var(--text-muted)", position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
                  <input className={inputCls + " w-full pl-8"} placeholder="Buscar por patente o tipo…" value={busqueda.acoplados} onChange={(e) => setBusqueda({ ...busqueda, acoplados: e.target.value })} />
                </div>
              </div>
              <div className="p-2">
                {acopladosFiltrados.length === 0 ? (
                  <p className="text-sm px-3 py-4" style={{ color: "var(--text-muted)" }}>Sin resultados.</p>
                ) : acopladosFiltrados.map(renderAcoplado)}
              </div>
            </div>
          </div>
        )}

        {seccionActiva === "choferes" && (
          <div>
            {puede("choferes") && (
              <div className="flex justify-end mb-4">
                <button onClick={() => openQuickAdd(activeQuickAdd === "chofer" ? null : "chofer")} className="px-4 py-2 rounded-md text-sm font-semibold flex items-center gap-1.5" style={{ background: "var(--accent-yellow)", color: "#0A1128" }}>
                  {activeQuickAdd === "chofer" ? <X size={16} /> : <Plus size={16} />} {activeQuickAdd === "chofer" ? "Cerrar" : "Agregar chofer"}
                </button>
              </div>
            )}
            {activeQuickAdd === "chofer" && (
              <div className="rounded-lg p-5 mb-6" style={{ background: "var(--panel)", border: "1px solid var(--accent-yellow)" }}>
                <div className="grid gap-4">
                  <div className="grid sm:grid-cols-2 gap-4">
                    <Field label="Nombre">
                      <input className={inputCls} value={fChofer.nombre} onChange={(e) => setFChofer({ ...fChofer, nombre: e.target.value })} placeholder="Nombre y apellido" />
                    </Field>
                    <Field label="Camión predeterminado">
                      <select className={inputCls} value={fChofer.camionId} onChange={(e) => setFChofer({ ...fChofer, camionId: e.target.value })}>
                        <option value="">Sin asignar</option>
                        {[...camiones].sort((a, b) => a.patente.localeCompare(b.patente)).map((c) => <option key={c.id} value={c.id}>{c.patente}</option>)}
                      </select>
                    </Field>
                    <Field label="Acoplado predeterminado">
                      <select className={inputCls} value={fChofer.acopladoId} onChange={(e) => setFChofer({ ...fChofer, acopladoId: e.target.value })}>
                        <option value="">Sin asignar</option>
                        {[...acoplados].sort((a, b) => a.patente.localeCompare(b.patente)).map((a) => <option key={a.id} value={a.id}>{a.patente}</option>)}
                      </select>
                    </Field>
                  </div>
                  <p className="text-xs" style={{ color: "var(--text-muted)" }}>
                    Este combo se usa como base al cargar viajes de este chofer; se puede cambiar puntualmente en cada viaje sin afectar el predeterminado.
                  </p>
                  <div className="text-xs uppercase tracking-wide mt-2" style={{ color: "var(--text-muted)" }}>
                    Documentación (opcional)
                  </div>
                  <div className="grid sm:grid-cols-2 gap-4">
                    <Field label="Vencimiento DNI">
                      <input type="date" className={inputCls} value={fChofer.vencimientoDni} onChange={(e) => setFChofer({ ...fChofer, vencimientoDni: e.target.value })} />
                    </Field>
                    <Field label="Vencimiento licencia">
                      <input type="date" className={inputCls} value={fChofer.vencimientoLicencia} onChange={(e) => setFChofer({ ...fChofer, vencimientoLicencia: e.target.value })} />
                    </Field>
                    <Field label="Curso defensivo">
                      <input type="date" className={inputCls} value={fChofer.fechaCursoDefensivo} onChange={(e) => setFChofer({ ...fChofer, fechaCursoDefensivo: e.target.value })} />
                    </Field>
                    <Field label="Estudios médicos">
                      <input type="date" className={inputCls} value={fChofer.fechaEstudiosMedicos} onChange={(e) => setFChofer({ ...fChofer, fechaEstudiosMedicos: e.target.value })} />
                    </Field>
                    <Field label="SVO">
                      <input type="date" className={inputCls} value={fChofer.fechaSvo} onChange={(e) => setFChofer({ ...fChofer, fechaSvo: e.target.value })} />
                    </Field>
                    <Field label="ART">
                      <input type="date" className={inputCls} value={fChofer.fechaArt} onChange={(e) => setFChofer({ ...fChofer, fechaArt: e.target.value })} />
                    </Field>
                    <Field label="EEPP">
                      <input type="date" className={inputCls} value={fChofer.fechaEepp} onChange={(e) => setFChofer({ ...fChofer, fechaEepp: e.target.value })} />
                    </Field>
                  </div>
                  <button onClick={saveChofer} className="justify-self-start px-4 py-2 rounded-md text-sm font-semibold" style={{ background: "var(--accent-yellow)", color: "#14161A" }}>
                    Guardar chofer
                  </button>
                </div>
              </div>
            )}
            <div className="rounded-lg" style={{ background: "var(--panel)", border: "1px solid var(--border)" }}>
              <div className="p-3">
                <div className="relative">
                  <Search size={14} style={{ color: "var(--text-muted)", position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
                  <input className={inputCls + " w-full pl-8"} placeholder="Buscar por nombre…" value={busqueda.choferes} onChange={(e) => setBusqueda({ ...busqueda, choferes: e.target.value })} />
                </div>
              </div>
              <div className="p-2">
                {choferesFiltrados.length === 0 ? (
                  <p className="text-sm px-3 py-4" style={{ color: "var(--text-muted)" }}>Sin resultados.</p>
                ) : choferesFiltrados.map(renderChofer)}
              </div>
            </div>
          </div>
        )}

        {seccionActiva === "neumaticos" && (
          <div>
            {puede("neumaticos") && (
              <div className="flex justify-end mb-4">
                <button onClick={() => openQuickAdd(activeQuickAdd === "neumatico" ? null : "neumatico")} className="px-4 py-2 rounded-md text-sm font-semibold flex items-center gap-1.5" style={{ background: "var(--accent-yellow)", color: "#0A1128" }}>
                  {activeQuickAdd === "neumatico" ? <X size={16} /> : <Plus size={16} />} {activeQuickAdd === "neumatico" ? "Cerrar" : "Agregar neumático"}
                </button>
              </div>
            )}
            {activeQuickAdd === "neumatico" && (
              <div className="rounded-lg p-5 mb-6" style={{ background: "var(--panel)", border: "1px solid var(--accent-yellow)" }}>
                <div className="grid gap-4">
                  <div className="grid sm:grid-cols-2 gap-4">
                    <Field label="Etiqueta">
                      <input className={inputCls} value={fNeumatico.etiqueta} onChange={(e) => setFNeumatico({ ...fNeumatico, etiqueta: e.target.value })} placeholder="N-0245" />
                    </Field>
                    <Field label="Estado">
                      <select className={inputCls} value={fNeumatico.estado} onChange={(e) => setFNeumatico({ ...fNeumatico, estado: e.target.value })}>
                        <option>Nuevo</option>
                        <option>Usado</option>
                        <option>Recapado</option>
                      </select>
                    </Field>
                    <Field label="Marca">
                      <input className={inputCls} value={fNeumatico.marca} onChange={(e) => setFNeumatico({ ...fNeumatico, marca: e.target.value })} placeholder="Ej: Fate, Bridgestone" />
                    </Field>
                    <Field label="Tipo">
                      <select className={inputCls} value={fNeumatico.tipo} onChange={(e) => setFNeumatico({ ...fNeumatico, tipo: e.target.value })}>
                        <option value="">Sin especificar</option>
                        <option value="DE TACO">DE TACO</option>
                        <option value="LISAS">LISAS</option>
                      </select>
                    </Field>
                  </div>
                  <Field label="Kilometraje inicial">
                    <input
                      type="number" min="0" placeholder="0"
                      className={inputCls}
                      value={fNeumatico.km}
                      onChange={(e) => setFNeumatico({ ...fNeumatico, km: e.target.value })}
                    />
                  </Field>
                  <p className="text-xs" style={{ color: "var(--text-muted)" }}>
                    Dejalo en blanco o en 0 si el neumático es 0km. Si ya tiene uso, escribí los km que trae acumulados.
                  </p>
                  <div className="grid sm:grid-cols-2 gap-4">
                    <Field label="Dónde va montado">
                      <select className={inputCls} value={fNeumatico.montaje} onChange={(e) => setFNeumatico({ ...fNeumatico, montaje: e.target.value, vehiculoId: "", posicion: "" })}>
                        <option value="">Sin montar todavía</option>
                        <option value="camion">En un camión</option>
                        <option value="acoplado">En un acoplado</option>
                      </select>
                    </Field>
                    {fNeumatico.montaje && (
                      <Field label={fNeumatico.montaje === "camion" ? "Camión" : "Acoplado"}>
                        <select className={inputCls} value={fNeumatico.vehiculoId} onChange={(e) => setFNeumatico({ ...fNeumatico, vehiculoId: e.target.value, posicion: "" })}>
                          <option value="">Elegir</option>
                          {(fNeumatico.montaje === "camion" ? camiones : acoplados).map((v) => (
                            <option key={v.id} value={v.id}>{v.patente}</option>
                          ))}
                        </select>
                      </Field>
                    )}
                  </div>
                  {fNeumatico.montaje && fNeumatico.vehiculoId && (() => {
                    const ejes = fNeumatico.montaje === "camion" ? CAMION_EJES : ACOPLADO_EJES;
                    return (
                      <div>
                        <div className="text-xs uppercase tracking-wide mb-2" style={{ color: "var(--text-muted)" }}>Posición en el vehículo</div>
                        <DiagramaSeleccionarPosicion
                          ejes={ejes} neumaticos={neumaticos} tipo={fNeumatico.montaje} vehiculoId={fNeumatico.vehiculoId}
                          neumaticoActualId={null} posicionSeleccionada={fNeumatico.posicion}
                          onSeleccionar={(codigo) => setFNeumatico({ ...fNeumatico, posicion: codigo })}
                        />
                      </div>
                    );
                  })()}
                  <button onClick={saveNeumatico} className="justify-self-start px-4 py-2 rounded-md text-sm font-semibold" style={{ background: "var(--accent-yellow)", color: "#14161A" }}>
                    Guardar neumático
                  </button>
                </div>
              </div>
            )}
            <div className="rounded-lg" style={{ background: "var(--panel)", border: "1px solid var(--border)" }}>
              <div className="p-3">
                <div className="relative">
                  <Search size={14} style={{ color: "var(--text-muted)", position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
                  <input className={inputCls + " w-full pl-8"} placeholder="Buscar por etiqueta o estado…" value={busqueda.neumaticos} onChange={(e) => setBusqueda({ ...busqueda, neumaticos: e.target.value })} />
                </div>
              </div>
              <div className="p-2">
                {neumaticosFiltrados.length === 0 ? (
                  <p className="text-sm px-3 py-4" style={{ color: "var(--text-muted)" }}>Sin resultados.</p>
                ) : neumaticosFiltrados.map(renderNeumatico)}
              </div>
            </div>
          </div>
        )}
      </div>
      </div>

      {selected && (
        <DetailPanel
          key={`${selected.type}-${selected.id}`}
          selected={selected}
          onClose={() => setSelected(null)}
          camiones={camiones} acoplados={acoplados} choferes={choferes}
          neumaticos={neumaticos} historialViajes={historialViajes}
          rutas={rutas} onRegistrarViaje={registrarViaje} onEditarViaje={editarViaje} onEliminarViaje={eliminarViaje}
          onMarcarRecapado={marcarRecapado} onCambiarVehiculo={cambiarVehiculoNeumatico}
          onAsignarPosicion={asignarAPosicion} onQuitarPosicion={quitarDePosicion}
          onVerNeumatico={(neumaticoId) => setSelected({ type: "neumatico", id: neumaticoId })}
          onAgregarService={agregarService} onActualizarVtv={actualizarVtv} onActualizarIntervalo={actualizarIntervaloService}
          onActualizarValvulas={actualizarValvulasTolva} onActualizarIntercement={actualizarIntercement}
          onEditarServicio={editarServicio} onEditarVehiculo={editarVehiculo}
          onEditarChofer={editarChofer} onEditarNeumatico={editarNeumatico}
          onEliminarCamion={eliminarCamion} onEliminarAcoplado={eliminarAcoplado}
          onEliminarChofer={eliminarChofer} onEliminarNeumatico={eliminarNeumatico}
          onSubirPoliza={subirPoliza} onDescargarPoliza={descargarPoliza} onEliminarPoliza={eliminarPoliza}
          esAdmin={currentUser.role === "admin"}
        />
      )}

      {toast && (
        <div
          className="fixed bottom-5 left-1/2 -translate-x-1/2 px-4 py-3 rounded-md text-sm font-medium shadow-lg flex items-center gap-2"
          style={{ background: "var(--accent-green)", color: "#0E1013" }}
        >
          <CheckCircle2 size={16} />
          {toast}
        </div>
      )}
    </div>
  );
}

function EntityPanel({ title, icon: Icon, items, renderItem, query, onQuery, placeholder, onVerTodos }) {
  const LIMITE = 3;
  const visibles = items.slice(0, LIMITE);
  return (
    <div className="rounded-lg flex flex-col" style={{ background: "var(--panel)", border: "1px solid var(--border)" }}>
      <div className="flex items-center gap-2 px-4 py-3" style={{ borderBottom: "1px solid var(--border)" }}>
        <Icon size={15} style={{ color: "var(--accent-blue)" }} />
        <h3 className="font-display text-sm uppercase tracking-wide">{title}</h3>
        <span className="ml-auto text-xs font-mono" style={{ color: "var(--text-muted)" }}>{items.length}</span>
      </div>
      <div className="px-3 pt-3">
        <div className="relative">
          <Search size={14} style={{ color: "var(--text-muted)", position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
          <input
            className={inputCls + " w-full pl-8"}
            placeholder={placeholder}
            value={query}
            onChange={(e) => onQuery(e.target.value)}
          />
        </div>
      </div>
      <div className="p-2 flex-1">
        {items.length === 0 ? (
          <p className="text-sm px-3 py-4" style={{ color: "var(--text-muted)" }}>Sin resultados.</p>
        ) : visibles.map(renderItem)}
      </div>
      {items.length > LIMITE && (
        <button
          onClick={onVerTodos}
          className="mx-3 mb-3 px-3 py-2 rounded-md text-xs font-semibold text-center"
          style={{ background: "var(--panel-2)", border: "1px solid var(--border)", color: "var(--accent-yellow)" }}
        >
          Ver todos ({items.length})
        </button>
      )}
    </div>
  );
}

function VerTodosModal({ title, icon: Icon, items, renderItem, query, onQuery, placeholder, onClose }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(0,0,0,0.6)" }} onClick={onClose}>
      <div
        onClick={(e) => e.stopPropagation()}
        className="w-full max-w-lg rounded-lg flex flex-col"
        style={{ background: "var(--panel)", border: "1px solid var(--border)", maxHeight: "80vh" }}
      >
        <div className="flex items-center gap-2 px-4 py-3" style={{ borderBottom: "1px solid var(--border)" }}>
          <Icon size={16} style={{ color: "var(--accent-blue)" }} />
          <h3 className="font-display text-sm uppercase tracking-wide">{title}</h3>
          <span className="ml-auto text-xs font-mono mr-2" style={{ color: "var(--text-muted)" }}>{items.length}</span>
          <button onClick={onClose} className="p-1 rounded hover:bg-[var(--panel-2)]">
            <X size={16} style={{ color: "var(--text-muted)" }} />
          </button>
        </div>
        <div className="px-3 pt-3">
          <div className="relative">
            <Search size={14} style={{ color: "var(--text-muted)", position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input
              className={inputCls + " w-full pl-8"}
              placeholder={placeholder}
              value={query}
              onChange={(e) => onQuery(e.target.value)}
              autoFocus
            />
          </div>
        </div>
        <div className="p-2 overflow-y-auto flex-1">
          {items.length === 0 ? (
            <p className="text-sm px-3 py-4" style={{ color: "var(--text-muted)" }}>Sin resultados.</p>
          ) : items.map(renderItem)}
        </div>
      </div>
    </div>
  );
}

function GraficoKmChoferes({ choferes, historialViajes, onVerChofer }) {
  const [mesSeleccionado, setMesSeleccionado] = useState("");

  if (choferes.length === 0) return null;

  const mesesDisponibles = [...new Set(historialViajes.map((v) => mesKey(v.fecha)))].sort().reverse();

  const datos = choferes.map((d) => {
    const viajesDeEste = historialViajes.filter((v) =>
      v.choferId === d.id && (!mesSeleccionado || mesKey(v.fecha) === mesSeleccionado)
    );
    const km = viajesDeEste.reduce((acc, v) => acc + v.km, 0);
    return { id: d.id, nombre: d.nombre, km };
  });
  const ordenados = [...datos].sort((a, b) => b.km - a.km);
  const maxKm = Math.max(...ordenados.map((d) => d.km), 1);

  return (
    <div className="rounded-lg mb-6" style={{ background: "var(--panel)", border: "1px solid var(--border)" }}>
      <div className="flex items-center gap-2 px-4 py-3 flex-wrap" style={{ borderBottom: "1px solid var(--border)" }}>
        <User size={15} style={{ color: "var(--accent-blue)" }} />
        <h2 className="font-display text-sm uppercase tracking-wide">Km por chofer</h2>
        <span className="text-xs font-mono" style={{ color: "var(--text-muted)" }}>{ordenados.length}</span>
        <select
          className={inputCls + " ml-auto text-xs py-1"}
          value={mesSeleccionado}
          onChange={(e) => setMesSeleccionado(e.target.value)}
        >
          <option value="">Acumulado total</option>
          {mesesDisponibles.map((key) => <option key={key} value={key}>{mesLabel(key)}</option>)}
        </select>
      </div>
      <div className="p-4 flex flex-col gap-3">
        {ordenados.map((d) => (
          <div
            key={d.id}
            onClick={() => onVerChofer(d.id)}
            className="cursor-pointer group"
          >
            <div className="flex justify-between items-baseline text-xs mb-1">
              <span className="font-medium group-hover:underline">{d.nombre}</span>
              <span className="font-mono odometer">{fmtKm(d.km)} km</span>
            </div>
            <div className="w-full rounded-full overflow-hidden" style={{ background: "var(--panel-2)", height: 10 }}>
              <div
                className="h-full rounded-full transition-all"
                style={{ width: `${Math.max(2, (d.km / maxKm) * 100)}%`, background: "var(--accent-yellow)" }}
              />
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function RowItem({ title, subtitle, right, meta, onClick }) {
  return (
    <div
      onClick={onClick}
      className="flex items-center gap-3 px-3 py-3 rounded-md cursor-pointer hover:bg-[var(--panel-2)] transition-colors"
      style={{ borderBottom: "1px solid var(--border)" }}
    >
      <div className="min-w-0 flex-1">
        <div className="text-sm font-semibold font-mono">{title}</div>
        <div className="text-xs" style={{ color: "var(--text-muted)" }}>
          {subtitle}{meta ? ` · ${meta.join(" · ")}` : ""}
        </div>
      </div>
      {right}
      <ChevronRight size={16} style={{ color: "var(--text-muted)" }} />
    </div>
  );
}

function FechaEditableRow({ label, fecha, onGuardar }) {
  const [editando, setEditando] = useState(false);
  const [valor, setValor] = useState(fecha || "");

  if (editando) {
    return (
      <div className="flex items-center gap-2 py-2" style={{ borderBottom: "1px solid var(--border)" }}>
        <span className="text-xs uppercase tracking-wide flex-shrink-0" style={{ color: "var(--text-muted)" }}>{label}</span>
        <input type="date" className={inputCls + " ml-auto"} value={valor} onChange={(e) => setValor(e.target.value)} />
        <button
          onClick={() => { onGuardar(valor); setEditando(false); }}
          className="px-2 py-1.5 rounded-md text-xs font-semibold"
          style={{ background: "var(--accent-green)", color: "#0A1128" }}
        >
          Guardar
        </button>
      </div>
    );
  }

  return (
    <div className="flex items-center justify-between py-2" style={{ borderBottom: "1px solid var(--border)" }}>
      <span className="text-xs uppercase tracking-wide" style={{ color: "var(--text-muted)" }}>{label}</span>
      <div className="flex items-center gap-2">
        <span className="text-sm font-mono" style={{ color: colorVencimiento(fecha) }}>{fecha || "Sin cargar"}</span>
        <button onClick={() => setEditando(true)} className="text-xs underline" style={{ color: "var(--accent-yellow)" }}>
          {fecha ? "editar" : "cargar"}
        </button>
      </div>
    </div>
  );
}

function ValvulasTolvaRow({ vencValvula1, vencValvula2, onGuardar }) {
  const [editando, setEditando] = useState(false);
  const [v1, setV1] = useState(vencValvula1 || "");
  const [v2, setV2] = useState(vencValvula2 || "");

  if (editando) {
    return (
      <div className="py-2" style={{ borderBottom: "1px solid var(--border)" }}>
        <span className="text-xs uppercase tracking-wide" style={{ color: "var(--text-muted)" }}>Válvulas (tolva)</span>
        <div className="grid sm:grid-cols-2 gap-3 mt-2">
          <Field label="Vencimiento válvula 1">
            <input type="date" className={inputCls} value={v1} onChange={(e) => setV1(e.target.value)} />
          </Field>
          <Field label="Vencimiento válvula 2">
            <input type="date" className={inputCls} value={v2} onChange={(e) => setV2(e.target.value)} />
          </Field>
        </div>
        <button
          onClick={() => { onGuardar(v1, v2); setEditando(false); }}
          className="mt-2 px-3 py-1.5 rounded-md text-xs font-semibold"
          style={{ background: "var(--accent-green)", color: "#0A1128" }}
        >
          Guardar
        </button>
      </div>
    );
  }

  return (
    <div className="py-2" style={{ borderBottom: "1px solid var(--border)" }}>
      <div className="flex items-center justify-between">
        <span className="text-xs uppercase tracking-wide" style={{ color: "var(--text-muted)" }}>Válvulas (tolva)</span>
        <button onClick={() => setEditando(true)} className="text-xs underline" style={{ color: "var(--accent-yellow)" }}>
          {vencValvula1 || vencValvula2 ? "editar" : "cargar"}
        </button>
      </div>
      <div className="text-sm font-mono mt-1">
        <span style={{ color: colorVencimiento(vencValvula1) }}>Válvula 1: {vencValvula1 || "sin cargar"}</span>
        {" · "}
        <span style={{ color: colorVencimiento(vencValvula2) }}>Válvula 2: {vencValvula2 || "sin cargar"}</span>
      </div>
    </div>
  );
}

function IntervaloServiceRow({ intervalo, onGuardar }) {
  const [editando, setEditando] = useState(false);
  const [valor, setValor] = useState(String(intervalo));

  if (editando) {
    return (
      <div className="flex items-center gap-2 py-2" style={{ borderBottom: "1px solid var(--border)" }}>
        <span className="text-xs uppercase tracking-wide flex-shrink-0" style={{ color: "var(--text-muted)" }}>Intervalo de service</span>
        <input type="number" min="1" className={inputCls + " ml-auto w-32"} value={valor} onChange={(e) => setValor(e.target.value)} />
        <button
          onClick={() => { onGuardar(valor); setEditando(false); }}
          className="px-2 py-1.5 rounded-md text-xs font-semibold"
          style={{ background: "var(--accent-green)", color: "#0A1128" }}
        >
          Guardar
        </button>
      </div>
    );
  }

  return (
    <div className="flex items-center justify-between py-2" style={{ borderBottom: "1px solid var(--border)" }}>
      <span className="text-xs uppercase tracking-wide" style={{ color: "var(--text-muted)" }}>Intervalo de service</span>
      <div className="flex items-center gap-2">
        <span className="text-sm font-mono">cada {fmtKm(intervalo)} km</span>
        <button onClick={() => setEditando(true)} className="text-xs underline" style={{ color: "var(--accent-yellow)" }}>
          editar
        </button>
      </div>
    </div>
  );
}

function VtvRow({ vtvUltima, dVtv, vencVtv, onGuardar }) {
  const [editando, setEditando] = useState(false);
  const [fecha, setFecha] = useState(vtvUltima || "");

  if (editando) {
    return (
      <div className="flex items-center gap-2 py-2" style={{ borderBottom: "1px solid var(--border)" }}>
        <span className="text-xs uppercase tracking-wide flex-shrink-0" style={{ color: "var(--text-muted)" }}>Última VTV</span>
        <input type="date" className={inputCls + " ml-auto"} value={fecha} onChange={(e) => setFecha(e.target.value)} />
        <button
          onClick={() => { onGuardar(fecha); setEditando(false); }}
          className="px-2 py-1.5 rounded-md text-xs font-semibold"
          style={{ background: "var(--accent-green)", color: "#0A1128" }}
        >
          Guardar
        </button>
      </div>
    );
  }

  return (
    <div className="flex items-center justify-between py-2" style={{ borderBottom: "1px solid var(--border)" }}>
      <span className="text-xs uppercase tracking-wide" style={{ color: "var(--text-muted)" }}>VTV</span>
      <div className="flex items-center gap-2">
        {vtvUltima ? (
          <span className="text-sm font-mono" style={{ color: dVtv < 0 ? "var(--accent-red)" : dVtv <= VTV_WARNING_DAYS ? "var(--accent-yellow)" : "var(--text)" }}>
            Última: {vtvUltima} · vence {vencVtv} {dVtv < 0 ? `(vencida hace ${Math.abs(dVtv)}d)` : `(en ${dVtv}d)`}
          </span>
        ) : (
          <span className="text-sm" style={{ color: "var(--text-muted)" }}>Sin fecha cargada</span>
        )}
        <button onClick={() => setEditando(true)} className="text-xs underline" style={{ color: "var(--accent-yellow)" }}>
          {vtvUltima ? "editar" : "cargar"}
        </button>
      </div>
    </div>
  );
}

function PolizaSeguro({ camion, onSubir, onDescargar, onEliminar }) {
  const [subiendo, setSubiendo] = useState(false);
  const [confirmandoBorrado, setConfirmandoBorrado] = useState(false);

  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setSubiendo(true);
    await onSubir(camion.id, file);
    setSubiendo(false);
    e.target.value = "";
  }

  return (
    <div className="mt-4 rounded-md p-3" style={{ background: "var(--panel-2)", border: "1px solid var(--border)" }}>
      <div className="text-xs uppercase tracking-wide mb-3 flex items-center gap-2" style={{ color: "var(--text-muted)" }}>
        <FileText size={13} /> Póliza de seguro
      </div>

      {camion.polizaUrl ? (
        <div className="flex items-center gap-3 flex-wrap">
          <span className="text-sm flex-1 min-w-[140px]">{camion.polizaNombre || "Documento cargado"}</span>
          <button
            onClick={() => onDescargar(camion.polizaUrl, camion.polizaNombre)}
            className="px-3 py-1.5 rounded-md text-xs font-semibold flex items-center gap-1.5"
            style={{ background: "var(--accent-yellow)", color: "#0A1128" }}
          >
            <Download size={13} /> Descargar
          </button>
          {confirmandoBorrado ? (
            <span className="text-xs flex items-center gap-2">
              <span style={{ color: "var(--accent-red)" }}>¿Eliminar?</span>
              <button onClick={() => { onEliminar(camion.id, camion.polizaUrl); setConfirmandoBorrado(false); }} className="underline" style={{ color: "var(--accent-red)" }}>Sí</button>
              <button onClick={() => setConfirmandoBorrado(false)} className="underline" style={{ color: "var(--text-muted)" }}>No</button>
            </span>
          ) : (
            <button onClick={() => setConfirmandoBorrado(true)} className="text-xs underline" style={{ color: "var(--accent-red)" }}>
              Eliminar
            </button>
          )}
          <label className="text-xs underline cursor-pointer" style={{ color: "var(--text-muted)" }}>
            {subiendo ? "Subiendo…" : "Reemplazar archivo"}
            <input type="file" className="hidden" onChange={handleFile} disabled={subiendo} />
          </label>
        </div>
      ) : (
        <label
          className="flex items-center justify-center gap-2 px-4 py-3 rounded-md text-sm font-semibold cursor-pointer"
          style={{ background: "var(--accent-yellow)", color: "#0A1128", opacity: subiendo ? 0.6 : 1 }}
        >
          <Upload size={14} /> {subiendo ? "Subiendo…" : "Subir póliza (PDF u otro documento)"}
          <input type="file" className="hidden" onChange={handleFile} disabled={subiendo} />
        </label>
      )}
    </div>
  );
}

function ServiceSection({ vehiculo, kmReal, tipo, onAgregarService, onEditarServicio }) {
  const [km, setKm] = useState(String(kmReal));
  const [detalle, setDetalle] = useState("");
  const [fecha, setFecha] = useState(todayISO());

  function handleAgregar() {
    onAgregarService(tipo, vehiculo.id, km, detalle, fecha);
    setDetalle("");
    setFecha(todayISO());
  }

  return (
    <div className="mt-4">
      <div className="text-xs uppercase tracking-wide mb-2 flex items-center gap-2" style={{ color: "var(--text-muted)" }}>
        <Wrench size={13} /> Service
      </div>
      <div className="rounded-md p-3 mb-3" style={{ background: "var(--panel-2)", border: "1px solid var(--border)" }}>
        <div className="grid sm:grid-cols-2 gap-3 mb-3">
          <Field label="Km al momento del service">
            <input type="number" min="0" className={inputCls} value={km} onChange={(e) => setKm(e.target.value)} />
          </Field>
          <Field label="Fecha">
            <input type="date" className={inputCls} value={fecha} onChange={(e) => setFecha(e.target.value)} />
          </Field>
        </div>
        <Field label="Qué se le hizo">
          <textarea
            className={inputCls + " w-full"} rows={2}
            placeholder="Ej: cambio de aceite, filtros y revisión de frenos"
            value={detalle} onChange={(e) => setDetalle(e.target.value)}
          />
        </Field>
        <button
          onClick={handleAgregar}
          className="mt-3 px-4 py-2 rounded-md text-sm font-semibold"
          style={{ background: "var(--accent-green)", color: "#0A1128" }}
        >
          Registrar service
        </button>
      </div>

      {vehiculo.servicios.length === 0 ? (
        <p className="text-sm" style={{ color: "var(--text-muted)" }}>Todavía no hay services registrados.</p>
      ) : vehiculo.servicios.map((s) => (
        <ServicioItem key={s.id} servicio={s} onEditar={(cambios) => onEditarServicio(tipo, vehiculo.id, s.id, cambios)} />
      ))}
    </div>
  );
}

function ViajeItem({ viaje, camion, acoplado, rutas, onEditar, onEliminar }) {
  const [editando, setEditando] = useState(false);
  const [rutaId, setRutaId] = useState(viaje.rutaId || "");
  const [km, setKm] = useState(String(viaje.km));
  const [fecha, setFecha] = useState(viaje.fecha);
  const [remito, setRemito] = useState(viaje.remito || "");
  const [confirmandoBorrado, setConfirmandoBorrado] = useState(false);

  function handleCambiarRuta(nuevaRutaId) {
    setRutaId(nuevaRutaId);
    const ruta = rutas.find((r) => r.id === nuevaRutaId);
    if (ruta) setKm(String(ruta.km));
  }

  function handleGuardar() {
    const ruta = rutas.find((r) => r.id === rutaId);
    onEditar(km, fecha, rutaId, ruta?.nombre, remito);
    setEditando(false);
  }

  if (editando) {
    return (
      <div className="py-2" style={{ borderBottom: "1px solid var(--border)" }}>
        <div className="flex items-center gap-2 text-sm mb-2">
          <Clock size={13} style={{ color: "var(--text-muted)" }} />
          <span className="flex-1">{viaje.codigo ? `Código: ${viaje.codigo}` : "Editando viaje"}</span>
        </div>
        <div className="grid sm:grid-cols-2 gap-2 mb-2">
          <Field label="Ruta">
            <select className={inputCls} value={rutaId} onChange={(e) => handleCambiarRuta(e.target.value)}>
              <option value="">{viaje.ruta} (actual)</option>
              {[...rutas].sort((a, b) => a.nombre.localeCompare(b.nombre)).map((r) => <option key={r.id} value={r.id}>{r.nombre} — {r.km} km</option>)}
            </select>
          </Field>
          <Field label="Fecha">
            <input type="date" className={inputCls} value={fecha} onChange={(e) => setFecha(e.target.value)} />
          </Field>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          <Field label="Km">
            <input type="number" min="0" className={inputCls + " w-32"} value={km} onChange={(e) => setKm(e.target.value)} />
          </Field>
          <Field label="Remito">
            <input className={inputCls + " w-32"} value={remito} onChange={(e) => setRemito(e.target.value)} placeholder="Ej: R-00456" />
          </Field>
          <button
            onClick={handleGuardar}
            className="px-3 py-1.5 rounded-md text-xs font-semibold mt-4" style={{ background: "var(--accent-green)", color: "#0A1128" }}
          >
            Guardar
          </button>
          <button
            onClick={() => { setKm(String(viaje.km)); setFecha(viaje.fecha); setRutaId(viaje.rutaId || ""); setRemito(viaje.remito || ""); setEditando(false); }}
            className="px-3 py-1.5 rounded-md text-xs font-semibold mt-4" style={{ background: "var(--panel)", border: "1px solid var(--border)", color: "var(--text-muted)" }}
          >
            Cancelar
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="py-1.5" style={{ borderBottom: "1px solid var(--border)" }}>
      <div className="flex items-center gap-2 text-sm">
        <Clock size={13} style={{ color: "var(--text-muted)" }} />
        <span className="flex-1">{viaje.ruta}{viaje.codigo ? ` · ${viaje.codigo}` : ""}{viaje.remito ? ` · Rem. ${viaje.remito}` : ""}</span>
        <span className="font-mono text-xs" style={{ color: "var(--text-muted)" }}>{viaje.fecha}</span>
        <span className="font-mono">{fmtKm(viaje.km)} km</span>
      </div>
      {viaje.distinto && (
        <div className="text-xs mt-1 ml-5" style={{ color: "var(--accent-yellow)" }}>
          Combo distinto al predeterminado — usó {camion?.patente || "sin camión"} + {acoplado?.patente || "sin acoplado"}
        </div>
      )}
      <div className="flex items-center gap-3 mt-1 ml-5">
        <button onClick={() => setEditando(true)} className="text-xs underline" style={{ color: "var(--accent-yellow)" }}>editar</button>
        {confirmandoBorrado ? (
          <span className="text-xs flex items-center gap-2">
            <span style={{ color: "var(--accent-red)" }}>¿Eliminar?</span>
            <button onClick={onEliminar} className="underline" style={{ color: "var(--accent-red)" }}>Sí</button>
            <button onClick={() => setConfirmandoBorrado(false)} className="underline" style={{ color: "var(--text-muted)" }}>No</button>
          </span>
        ) : (
          <button onClick={() => setConfirmandoBorrado(true)} className="text-xs underline" style={{ color: "var(--accent-red)" }}>eliminar</button>
        )}
      </div>
    </div>
  );
}

function ServicioItem({ servicio, onEditar }) {
  const [editando, setEditando] = useState(false);
  const [texto, setTexto] = useState(servicio.detalle);
  const [km, setKm] = useState(String(servicio.km));
  const [fecha, setFecha] = useState(servicio.fecha);

  if (editando) {
    return (
      <div className="py-2" style={{ borderBottom: "1px solid var(--border)" }}>
        <div className="grid sm:grid-cols-2 gap-3 mb-2">
          <Field label="Km">
            <input type="number" min="0" className={inputCls} value={km} onChange={(e) => setKm(e.target.value)} />
          </Field>
          <Field label="Fecha">
            <input type="date" className={inputCls} value={fecha} onChange={(e) => setFecha(e.target.value)} />
          </Field>
        </div>
        <textarea className={inputCls + " w-full"} rows={2} value={texto} onChange={(e) => setTexto(e.target.value)} />
        <div className="flex gap-2 mt-1.5">
          <button
            onClick={() => { onEditar({ detalle: texto, km: Math.max(0, parseInt(km, 10) || 0), fecha }); setEditando(false); }}
            className="px-3 py-1 rounded-md text-xs font-semibold"
            style={{ background: "var(--accent-green)", color: "#0A1128" }}
          >
            Guardar
          </button>
          <button
            onClick={() => { setTexto(servicio.detalle); setKm(String(servicio.km)); setFecha(servicio.fecha); setEditando(false); }}
            className="px-3 py-1 rounded-md text-xs font-semibold"
            style={{ background: "var(--panel-2)", border: "1px solid var(--border)", color: "var(--text-muted)" }}
          >
            Cancelar
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="py-2" style={{ borderBottom: "1px solid var(--border)" }}>
      <div className="flex items-center justify-between text-sm">
        <span className="font-mono" style={{ color: "var(--text-muted)" }}>{servicio.fecha} · {fmtKm(servicio.km)} km</span>
        <button onClick={() => setEditando(true)} className="text-xs underline" style={{ color: "var(--accent-yellow)" }}>editar</button>
      </div>
      <div className="text-sm mt-0.5">{servicio.detalle}</div>
      {servicio.editado && (
        <div className="text-xs mt-1" style={{ color: "var(--accent-yellow)" }}>
          Editado{servicio.fechaEdicion ? ` el ${servicio.fechaEdicion}` : ""}
        </div>
      )}
    </div>
  );
}

function PosicionesNeumaticos({ tipo, vehiculo, ejes, neumaticos, neumaticosDelVehiculo, onAsignar, onQuitar, onVerNeumatico }) {
  const [selCodigo, setSelCodigo] = useState(null);

  const porCodigo = {};
  neumaticosDelVehiculo.forEach((n) => { if (n.posicion) porCodigo[n.posicion] = n; });

  const sinPosicion = neumaticosDelVehiculo.filter((n) => !n.posicion);
  const disponibles = neumaticos.filter((n) => !n.camionId && !n.acopladoId);

  const rowGap = 74;
  const topPad = 34;
  const width = 260;
  const height = topPad + ejes.length * rowGap + 30;
  const cx = width / 2;

  function Wheel({ codigo, x, y }) {
    const n = porCodigo[codigo];
    const label = POSICION_LABELS[codigo] || codigo;
    const estadoColor = n ? colorEstadoNeumatico(estadoNeumatico(n)) : "var(--border)";
    const isSel = selCodigo === codigo;
    return (
      <g style={{ cursor: "pointer" }} onClick={() => setSelCodigo(isSel ? null : codigo)}>
        <circle cx={x} cy={y} r={17} fill={n ? "var(--panel-2)" : "transparent"}
          stroke={isSel ? "var(--accent-yellow)" : estadoColor} strokeWidth={isSel ? 3 : 2}
          strokeDasharray={n ? "0" : "4 3"} />
        <text x={x} y={y + 3} textAnchor="middle" fontSize="8" fontFamily="'JetBrains Mono', monospace"
          fill={n ? "var(--text)" : "var(--text-muted)"}>
          {label}
        </text>
      </g>
    );
  }

  const filas = ejes.map((eje, i) => {
    const y = topPad + i * rowGap;
    if (eje.tipo === "simple") {
      return (
        <g key={i}>
          <Wheel codigo={eje.codigos[0]} x={cx - 60} y={y} />
          <Wheel codigo={eje.codigos[1]} x={cx + 60} y={y} />
        </g>
      );
    }
    return (
      <g key={i}>
        <Wheel codigo={eje.codigos[0]} x={cx - 80} y={y} />
        <Wheel codigo={eje.codigos[1]} x={cx - 32} y={y} />
        <Wheel codigo={eje.codigos[2]} x={cx + 32} y={y} />
        <Wheel codigo={eje.codigos[3]} x={cx + 80} y={y} />
      </g>
    );
  });

  const seleccionado = selCodigo ? porCodigo[selCodigo] : null;

  return (
    <div className="mt-4">
      <div className="text-xs uppercase tracking-wide mb-2" style={{ color: "var(--text-muted)" }}>
        Posiciones de neumáticos
      </div>
      <div className="rounded-md p-2" style={{ background: "var(--panel-2)", border: "1px solid var(--border)" }}>
        <svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height}>
          <rect x={cx - 14} y={topPad - 24} width="28" height="14" rx="3" fill="var(--border)" />
          <rect x={cx - 10} y={topPad - 10} width="20" height={height - topPad - 20} fill="var(--border)" opacity="0.4" />
          {filas}
        </svg>
      </div>

      {selCodigo && (
        <div className="mt-3 rounded-md p-3" style={{ background: "var(--panel-2)", border: "1px solid var(--border)" }}>
          <div className="text-xs uppercase tracking-wide mb-2" style={{ color: "var(--accent-yellow)" }}>
            Posición {POSICION_LABELS[selCodigo] || selCodigo}
          </div>
          {seleccionado ? (
            <>
              <div className="text-sm font-mono">{seleccionado.etiqueta}</div>
              <div className="text-xs mb-3" style={{ color: colorEstadoNeumatico(estadoNeumatico(seleccionado)) }}>
                {estadoNeumatico(seleccionado)} · {fmtKm(seleccionado.km)} km
              </div>
              <div className="flex flex-wrap gap-2">
                <button
                  onClick={() => onVerNeumatico(seleccionado.id)}
                  className="px-3 py-1.5 rounded-md text-xs font-semibold"
                  style={{ background: "var(--accent-yellow)", color: "#0A1128" }}
                >
                  Ver ficha completa
                </button>
                <button
                  onClick={() => onQuitar(seleccionado.id)}
                  className="px-3 py-1.5 rounded-md text-xs font-semibold"
                  style={{ background: "var(--accent-red)", color: "#0A1128" }}
                >
                  Quitar de esta posición
                </button>
              </div>
            </>
          ) : (
            <>
              <div className="text-xs mb-2" style={{ color: "var(--text-muted)" }}>Posición vacía</div>
              <AsignarPosicionForm
                disponibles={disponibles}
                onAsignar={(neumaticoId) => { onAsignar(tipo, vehiculo.id, selCodigo, neumaticoId); }}
              />
            </>
          )}
        </div>
      )}

      {sinPosicion.length > 0 && (
        <div className="mt-3">
          <div className="text-xs" style={{ color: "var(--text-muted)" }}>
            Montados sin posición específica todavía:
          </div>
          {sinPosicion.map((n) => (
            <div key={n.id} className="flex justify-between text-sm py-1.5" style={{ borderBottom: "1px solid var(--border)" }}>
              <span>{n.etiqueta}</span>
              <span className="font-mono">{fmtKm(n.km)} km</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function AsignarPosicionForm({ disponibles, onAsignar }) {
  const [query, setQuery] = useState("");
  const [id, setId] = useState("");

  if (disponibles.length === 0) {
    return (
      <p className="text-sm" style={{ color: "var(--text-muted)" }}>
        No hay neumáticos disponibles sin asignar. Cargá uno nuevo desde "Carga rápida" para poder agregarlo acá.
      </p>
    );
  }

  const filtrados = disponibles.filter((n) =>
    `${n.etiqueta} ${n.estado}`.toLowerCase().includes(query.toLowerCase())
  );
  const elegido = disponibles.find((n) => n.id === id);

  return (
    <div>
      <div className="relative">
        <Search size={14} style={{ color: "var(--text-muted)", position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
        <input
          className={inputCls + " w-full pl-8"}
          placeholder={`Buscar entre ${disponibles.length} neumáticos disponibles…`}
          value={elegido ? `${elegido.etiqueta} · ${elegido.estado}` : query}
          onChange={(e) => { setQuery(e.target.value); setId(""); }}
          onFocus={() => setId("")}
        />
      </div>

      {!elegido && query && (
        <div className="mt-1 rounded-md overflow-hidden max-h-40 overflow-y-auto" style={{ border: "1px solid var(--border)" }}>
          {filtrados.length === 0 ? (
            <div className="px-3 py-2 text-sm bg-white" style={{ color: "#6B7690" }}>Sin resultados.</div>
          ) : filtrados.map((n) => (
            <button
              key={n.id}
              onClick={() => { setId(n.id); setQuery(""); }}
              className="w-full text-left px-3 py-2 text-sm bg-white hover:bg-gray-100 flex justify-between"
              style={{ color: "#0A1128", borderBottom: "1px solid #E5E7EB" }}
            >
              <span>{n.etiqueta}</span>
              <span style={{ color: colorEstadoNeumatico(estadoNeumatico(n)) }}>{estadoNeumatico(n)}</span>
            </button>
          ))}
        </div>
      )}

      <button
        disabled={!id}
        onClick={() => onAsignar(id)}
        className="mt-2 px-3 py-2 rounded-md text-xs font-semibold disabled:opacity-40 flex items-center gap-1.5"
        style={{ background: "var(--accent-green)", color: "#0A1128" }}
      >
        <Plus size={13} /> Agregar neumático
      </button>
    </div>
  );
}

function DiagramaSeleccionarPosicion({ ejes, neumaticos, tipo, vehiculoId, neumaticoActualId, posicionSeleccionada, onSeleccionar }) {
  const porCodigo = {};
  neumaticos.forEach((n) => {
    if (!n.posicion || n.id === neumaticoActualId) return;
    const mismoVehiculo = (tipo === "camion" && n.camionId === vehiculoId) || (tipo === "acoplado" && n.acopladoId === vehiculoId);
    if (mismoVehiculo) porCodigo[n.posicion] = n;
  });

  const rowGap = 70;
  const topPad = 30;
  const width = 240;
  const height = topPad + ejes.length * rowGap + 20;
  const cx = width / 2;

  function Wheel({ codigo, x, y }) {
    const ocupante = porCodigo[codigo];
    const label = POSICION_LABELS[codigo] || codigo;
    const isSel = posicionSeleccionada === codigo;
    const estadoColor = ocupante ? colorEstadoNeumatico(estadoNeumatico(ocupante)) : "var(--border)";
    return (
      <g style={{ cursor: "pointer" }} onClick={() => onSeleccionar(isSel ? "" : codigo)}>
        <circle cx={x} cy={y} r={16} fill={ocupante ? "var(--panel)" : "transparent"}
          stroke={isSel ? "var(--accent-yellow)" : estadoColor} strokeWidth={isSel ? 3 : 2}
          strokeDasharray={ocupante ? "0" : "4 3"} />
        <text x={x} y={y + 3} textAnchor="middle" fontSize="7.5" fontFamily="'JetBrains Mono', monospace"
          fill={ocupante ? "var(--text)" : "var(--text-muted)"}>
          {label}
        </text>
      </g>
    );
  }

  const filas = ejes.map((eje, i) => {
    const y = topPad + i * rowGap;
    if (eje.tipo === "simple") {
      return (
        <g key={i}>
          <Wheel codigo={eje.codigos[0]} x={cx - 55} y={y} />
          <Wheel codigo={eje.codigos[1]} x={cx + 55} y={y} />
        </g>
      );
    }
    return (
      <g key={i}>
        <Wheel codigo={eje.codigos[0]} x={cx - 75} y={y} />
        <Wheel codigo={eje.codigos[1]} x={cx - 30} y={y} />
        <Wheel codigo={eje.codigos[2]} x={cx + 30} y={y} />
        <Wheel codigo={eje.codigos[3]} x={cx + 75} y={y} />
      </g>
    );
  });

  const ocupanteSeleccionado = posicionSeleccionada ? porCodigo[posicionSeleccionada] : null;

  return (
    <div>
      <div className="rounded-md p-2" style={{ background: "var(--panel)", border: "1px solid var(--border)" }}>
        <svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height}>
          <rect x={cx - 13} y={topPad - 22} width="26" height="13" rx="3" fill="var(--border)" />
          <rect x={cx - 9} y={topPad - 9} width="18" height={height - topPad - 18} fill="var(--border)" opacity="0.4" />
          {filas}
        </svg>
      </div>
      <div className="flex items-center justify-between mt-2">
        <span className="text-xs" style={{ color: "var(--text-muted)" }}>
          {posicionSeleccionada
            ? `${POSICION_LABELS[posicionSeleccionada] || posicionSeleccionada}${ocupanteSeleccionado ? ` · ocupada por ${ocupanteSeleccionado.etiqueta}` : ""}`
            : "Sin posición específica"}
        </span>
        {posicionSeleccionada && (
          <button onClick={() => onSeleccionar("")} className="text-xs underline" style={{ color: "var(--text-muted)" }}>
            quitar selección
          </button>
        )}
      </div>
      {ocupanteSeleccionado && (
        <p className="text-xs mt-1" style={{ color: "var(--accent-yellow)" }}>
          Esa posición ya tiene a {ocupanteSeleccionado.etiqueta} — al aplicar, ese neumático queda sin asignar.
        </p>
      )}
    </div>
  );
}

function NeumaticoAcciones({ neumatico, camiones, acoplados, neumaticos, onRecapar, onCambiarVehiculo }) {
  const [tipo, setTipo] = useState(neumatico.camionId ? "camion" : neumatico.acopladoId ? "acoplado" : "camion");
  const [vehiculoId, setVehiculoId] = useState(neumatico.camionId || neumatico.acopladoId || "");
  const [posicion, setPosicion] = useState(neumatico.posicion || "");

  const yaRecapado = neumatico.estado === "Recapado";
  const ejes = tipo === "camion" ? CAMION_EJES : ACOPLADO_EJES;

  function handleTipoChange(v) {
    setTipo(v);
    setVehiculoId("");
    setPosicion("");
  }

  function handleVehiculoChange(v) {
    setVehiculoId(v);
    setPosicion("");
  }

  return (
    <div className="mt-4 rounded-md p-3" style={{ background: "var(--panel-2)", border: "1px solid var(--border)" }}>
      <div className="text-xs uppercase tracking-wide mb-3" style={{ color: "var(--text-muted)" }}>Acciones</div>

      <div className="flex flex-wrap gap-2 mb-3">
        <button
          disabled={yaRecapado}
          onClick={() => onRecapar(neumatico.id)}
          className="px-3 py-2 rounded-md text-xs font-semibold disabled:opacity-40 flex items-center gap-1.5"
          style={{ background: "var(--accent-blue)", color: "#0A1128" }}
        >
          <Wrench size={13} /> {yaRecapado ? "Ya está recapado" : "Mandar a recapar"}
        </button>
      </div>

      <div className="text-xs uppercase tracking-wide mb-2" style={{ color: "var(--text-muted)" }}>Cambiar de camión / acoplado</div>
      <div className="grid sm:grid-cols-2 gap-3">
        <Field label="Tipo">
          <select className={inputCls} value={tipo} onChange={(e) => handleTipoChange(e.target.value)}>
            <option value="camion">Camión</option>
            <option value="acoplado">Acoplado</option>
          </select>
        </Field>
        <Field label={tipo === "camion" ? "Camión" : "Acoplado"}>
          <select className={inputCls} value={vehiculoId} onChange={(e) => handleVehiculoChange(e.target.value)}>
            <option value="">Sin asignar</option>
            {(tipo === "camion" ? camiones : acoplados).map((v) => (
              <option key={v.id} value={v.id}>{v.patente}</option>
            ))}
          </select>
        </Field>
      </div>

      {vehiculoId && (
        <div className="mt-3">
          <div className="text-xs uppercase tracking-wide mb-2" style={{ color: "var(--text-muted)" }}>Posición en el vehículo</div>
          <DiagramaSeleccionarPosicion
            ejes={ejes} neumaticos={neumaticos} tipo={tipo} vehiculoId={vehiculoId}
            neumaticoActualId={neumatico.id} posicionSeleccionada={posicion} onSeleccionar={setPosicion}
          />
        </div>
      )}

      <button
        onClick={() => onCambiarVehiculo(neumatico.id, tipo, vehiculoId, posicion)}
        className="mt-3 px-4 py-2 rounded-md text-sm font-semibold"
        style={{ background: "var(--accent-green)", color: "#0A1128" }}
      >
        Aplicar cambio
      </button>
    </div>
  );
}

function RutaBuscador({ rutas, rutaId, onElegir, label = "Ruta que va a hacer" }) {
  const [query, setQuery] = useState("");
  const [abierto, setAbierto] = useState(false);
  const rutaElegida = rutas.find((r) => r.id === rutaId);
  const rutasOrdenadas = [...rutas].sort((a, b) => a.nombre.localeCompare(b.nombre));
  const filtradas = rutasOrdenadas.filter((r) => r.nombre.toLowerCase().includes(query.toLowerCase()));

  return (
    <Field label={label}>
      <div className="relative">
        <input
          className={inputCls + " w-full"}
          placeholder="Buscar ruta…"
          value={abierto ? query : (rutaElegida ? `${rutaElegida.nombre} — ${rutaElegida.km} km` : "")}
          onFocus={() => { setAbierto(true); setQuery(""); }}
          onBlur={() => setTimeout(() => setAbierto(false), 150)}
          onChange={(e) => setQuery(e.target.value)}
        />
        {abierto && (
          <div
            className="absolute z-10 mt-1 w-full rounded-md overflow-hidden max-h-48 overflow-y-auto"
            style={{ border: "1px solid var(--border)" }}
          >
            {filtradas.length === 0 ? (
              <div className="px-3 py-2 text-sm bg-white" style={{ color: "#6B7690" }}>Sin resultados.</div>
            ) : filtradas.map((r) => (
              <button
                key={r.id}
                onClick={() => { onElegir(r.id); setAbierto(false); setQuery(""); }}
                className="w-full text-left px-3 py-2 text-sm bg-white hover:bg-gray-100 flex justify-between"
                style={{ color: "#0A1128", borderBottom: "1px solid #E5E7EB" }}
              >
                <span>{r.nombre}</span>
                <span style={{ color: "#6B7690" }}>{r.km} km</span>
              </button>
            ))}
          </div>
        )}
      </div>
    </Field>
  );
}

function ViajeChoferForm({ chofer, camiones, acoplados, rutas, onRegistrarViaje }) {
  const camionDefault = camiones.find((c) => c.id === chofer.camionId);
  const acopladoDefault = acoplados.find((a) => a.id === chofer.acopladoId);

  const [rutaId, setRutaId] = useState("");
  const [camionId, setCamionId] = useState(camionDefault?.id || "");
  const [acopladoId, setAcopladoId] = useState(acopladoDefault?.id || "");
  const [codigo, setCodigo] = useState("");
  const [remito, setRemito] = useState("");
  const [fecha, setFecha] = useState(todayISO());

  const usarOtro = camionId !== (camionDefault?.id || "") || acopladoId !== (acopladoDefault?.id || "");

  function handleRegistrar() {
    if (!rutaId) return;
    onRegistrarViaje(chofer.id, rutaId, camionId, acopladoId, usarOtro, codigo, fecha, remito);
    setRutaId("");
    setCamionId(camionDefault?.id || "");
    setAcopladoId(acopladoDefault?.id || "");
    setCodigo("");
    setRemito("");
    setFecha(todayISO());
  }

  return (
    <div className="mt-4 rounded-md p-3" style={{ background: "var(--panel-2)", border: "1px solid var(--border)" }}>
      <div className="text-xs uppercase tracking-wide mb-3 flex items-center gap-2" style={{ color: "var(--text-muted)" }}>
        <Route size={13} /> Cargar viaje de este chofer
      </div>
      <div className="grid gap-3">
        <div className="grid sm:grid-cols-2 gap-3">
          <RutaBuscador rutas={rutas} rutaId={rutaId} onElegir={setRutaId} />
          <Field label="Código">
            <input className={inputCls} value={codigo} onChange={(e) => setCodigo(e.target.value)} placeholder="Ej: V-00123" />
          </Field>
          <Field label="Remito">
            <input className={inputCls} value={remito} onChange={(e) => setRemito(e.target.value)} placeholder="Ej: R-00456" />
          </Field>
          <Field label="Fecha del viaje">
            <input type="date" className={inputCls} value={fecha} onChange={(e) => setFecha(e.target.value)} />
          </Field>
        </div>
        <div className="grid sm:grid-cols-2 gap-3">
          <Field label={`Camión${camionId === (camionDefault?.id || "") ? " (predeterminado)" : ""}`}>
            <select className={inputCls} value={camionId} onChange={(e) => setCamionId(e.target.value)}>
              <option value="">Sin camión</option>
              {[...camiones].sort((a, b) => a.patente.localeCompare(b.patente)).map((c) => <option key={c.id} value={c.id}>{c.patente}{c.id === camionDefault?.id ? " · predeterminado" : ""}</option>)}
            </select>
          </Field>
          <Field label={`Acoplado${acopladoId === (acopladoDefault?.id || "") ? " (predeterminado)" : ""}`}>
            <select className={inputCls} value={acopladoId} onChange={(e) => setAcopladoId(e.target.value)}>
              <option value="">Sin acoplado</option>
              {[...acoplados].sort((a, b) => a.patente.localeCompare(b.patente)).map((a) => <option key={a.id} value={a.id}>{a.patente}{a.id === acopladoDefault?.id ? " · predeterminado" : ""}</option>)}
            </select>
          </Field>
        </div>
        {usarOtro && (
          <p className="text-xs" style={{ color: "var(--accent-yellow)" }}>
            Este viaje se va a registrar con un combo distinto al predeterminado del chofer. El predeterminado no cambia.
          </p>
        )}
        <button
          onClick={handleRegistrar}
          disabled={!rutaId}
          className="justify-self-start px-4 py-2 rounded-md text-sm font-semibold disabled:opacity-40"
          style={{ background: "var(--accent-yellow)", color: "#14161A" }}
        >
          Registrar viaje
        </button>
      </div>
    </div>
  );
}

function ConfirmDeleteButton({ label, onConfirm }) {
  const [confirmando, setConfirmando] = useState(false);

  if (confirmando) {
    return (
      <div className="flex items-center gap-2 text-xs">
        <span style={{ color: "var(--accent-red)" }}>¿Seguro que querés eliminarlo?</span>
        <button onClick={onConfirm} className="px-2 py-1 rounded-md font-semibold" style={{ background: "var(--accent-red)", color: "#0A1128" }}>
          Sí, eliminar
        </button>
        <button onClick={() => setConfirmando(false)} className="px-2 py-1 rounded-md" style={{ background: "var(--panel)", border: "1px solid var(--border)", color: "var(--text-muted)" }}>
          Cancelar
        </button>
      </div>
    );
  }

  return (
    <button onClick={() => setConfirmando(true)} className="text-xs underline" style={{ color: "var(--accent-red)" }}>
      {label}
    </button>
  );
}

function EditarChoferForm({ chofer, camiones, acoplados, onGuardar, onCerrar }) {
  const [nombre, setNombre] = useState(chofer.nombre);
  const [camionId, setCamionId] = useState(chofer.camionId || "");
  const [acopladoId, setAcopladoId] = useState(chofer.acopladoId || "");
  const [vencimientoDni, setVencimientoDni] = useState(chofer.vencimientoDni || "");
  const [vencimientoLicencia, setVencimientoLicencia] = useState(chofer.vencimientoLicencia || "");
  const [fechaCursoDefensivo, setFechaCursoDefensivo] = useState(chofer.fechaCursoDefensivo || "");
  const [fechaEstudiosMedicos, setFechaEstudiosMedicos] = useState(chofer.fechaEstudiosMedicos || "");
  const [fechaSvo, setFechaSvo] = useState(chofer.fechaSvo || "");
  const [fechaArt, setFechaArt] = useState(chofer.fechaArt || "");
  const [fechaEepp, setFechaEepp] = useState(chofer.fechaEepp || "");

  function handleGuardar() {
    onGuardar({
      nombre, camionId: camionId || null, acopladoId: acopladoId || null,
      vencimientoDni: vencimientoDni || null, vencimientoLicencia: vencimientoLicencia || null,
      fechaCursoDefensivo: fechaCursoDefensivo || null, fechaEstudiosMedicos: fechaEstudiosMedicos || null,
      fechaSvo: fechaSvo || null, fechaArt: fechaArt || null, fechaEepp: fechaEepp || null,
    });
    onCerrar();
  }

  return (
    <div className="mb-4 rounded-md p-3" style={{ background: "var(--panel-2)", border: "1px solid var(--accent-yellow)" }}>
      <div className="text-xs uppercase tracking-wide mb-3" style={{ color: "var(--text-muted)" }}>Editar datos</div>
      <div className="grid sm:grid-cols-2 gap-3">
        <Field label="Nombre">
          <input className={inputCls} value={nombre} onChange={(e) => setNombre(e.target.value)} />
        </Field>
        <Field label="Camión predeterminado">
          <select className={inputCls} value={camionId} onChange={(e) => setCamionId(e.target.value)}>
            <option value="">Sin asignar</option>
            {[...camiones].sort((a, b) => a.patente.localeCompare(b.patente)).map((c) => <option key={c.id} value={c.id}>{c.patente}</option>)}
          </select>
        </Field>
        <Field label="Acoplado predeterminado">
          <select className={inputCls} value={acopladoId} onChange={(e) => setAcopladoId(e.target.value)}>
            <option value="">Sin asignar</option>
            {[...acoplados].sort((a, b) => a.patente.localeCompare(b.patente)).map((a) => <option key={a.id} value={a.id}>{a.patente}</option>)}
          </select>
        </Field>
      </div>
      <div className="text-xs uppercase tracking-wide mt-3 mb-2" style={{ color: "var(--text-muted)" }}>Documentación (opcional)</div>
      <div className="grid sm:grid-cols-2 gap-3">
        <Field label="Vencimiento DNI">
          <input type="date" className={inputCls} value={vencimientoDni} onChange={(e) => setVencimientoDni(e.target.value)} />
        </Field>
        <Field label="Vencimiento licencia">
          <input type="date" className={inputCls} value={vencimientoLicencia} onChange={(e) => setVencimientoLicencia(e.target.value)} />
        </Field>
        <Field label="Curso defensivo">
          <input type="date" className={inputCls} value={fechaCursoDefensivo} onChange={(e) => setFechaCursoDefensivo(e.target.value)} />
        </Field>
        <Field label="Estudios médicos">
          <input type="date" className={inputCls} value={fechaEstudiosMedicos} onChange={(e) => setFechaEstudiosMedicos(e.target.value)} />
        </Field>
        <Field label="SVO">
          <input type="date" className={inputCls} value={fechaSvo} onChange={(e) => setFechaSvo(e.target.value)} />
        </Field>
        <Field label="ART">
          <input type="date" className={inputCls} value={fechaArt} onChange={(e) => setFechaArt(e.target.value)} />
        </Field>
        <Field label="EEPP">
          <input type="date" className={inputCls} value={fechaEepp} onChange={(e) => setFechaEepp(e.target.value)} />
        </Field>
      </div>
      <div className="flex gap-2 mt-3">
        <button
          onClick={handleGuardar}
          className="px-4 py-2 rounded-md text-sm font-semibold" style={{ background: "var(--accent-green)", color: "#0A1128" }}
        >
          Guardar cambios
        </button>
        <button onClick={onCerrar} className="px-4 py-2 rounded-md text-sm font-semibold" style={{ background: "var(--panel)", border: "1px solid var(--border)", color: "var(--text-muted)" }}>
          Cancelar
        </button>
      </div>
    </div>
  );
}

function EditarNeumaticoForm({ neumatico, onGuardar, onCerrar }) {
  const [etiqueta, setEtiqueta] = useState(neumatico.etiqueta);
  const [estado, setEstado] = useState(neumatico.estado);
  const [km, setKm] = useState(String(neumatico.km));
  const [marca, setMarca] = useState(neumatico.marca || "");
  const [tipo, setTipo] = useState(neumatico.tipo || "");

  return (
    <div className="mb-4 rounded-md p-3" style={{ background: "var(--panel-2)", border: "1px solid var(--accent-yellow)" }}>
      <div className="text-xs uppercase tracking-wide mb-3" style={{ color: "var(--text-muted)" }}>Editar datos</div>
      <div className="grid sm:grid-cols-3 gap-3">
        <Field label="Etiqueta">
          <input className={inputCls} value={etiqueta} onChange={(e) => setEtiqueta(e.target.value)} />
        </Field>
        <Field label="Estado">
          <select className={inputCls} value={estado} onChange={(e) => setEstado(e.target.value)}>
            <option>Nuevo</option>
            <option>Usado</option>
            <option>Recapado</option>
          </select>
        </Field>
        <Field label="Km acumulados">
          <input type="number" min="0" className={inputCls} value={km} onChange={(e) => setKm(e.target.value)} />
        </Field>
        <Field label="Marca">
          <input className={inputCls} value={marca} onChange={(e) => setMarca(e.target.value)} placeholder="Ej: Fate" />
        </Field>
        <Field label="Tipo">
          <select className={inputCls} value={tipo} onChange={(e) => setTipo(e.target.value)}>
            <option value="">Sin especificar</option>
            <option value="DE TACO">DE TACO</option>
            <option value="LISAS">LISAS</option>
          </select>
        </Field>
      </div>
      <div className="flex gap-2 mt-3">
        <button
          onClick={() => { onGuardar({ etiqueta, estado, km: Math.max(0, parseInt(km, 10) || 0), marca, tipo }); onCerrar(); }}
          className="px-4 py-2 rounded-md text-sm font-semibold" style={{ background: "var(--accent-green)", color: "#0A1128" }}
        >
          Guardar cambios
        </button>
        <button onClick={onCerrar} className="px-4 py-2 rounded-md text-sm font-semibold" style={{ background: "var(--panel)", border: "1px solid var(--border)", color: "var(--text-muted)" }}>
          Cancelar
        </button>
      </div>
    </div>
  );
}

function EditarVehiculoForm({ vehiculo, kmReal, tipo, onGuardar, onCerrar }) {
  const [patente, setPatente] = useState(vehiculo.patente);
  const [modeloTipo, setModeloTipo] = useState(tipo === "camion" ? vehiculo.modelo : vehiculo.tipo);
  const [km, setKm] = useState(String(kmReal));

  const OPCIONES_TIPO_ACOPLADO = ["TOLVA", "SEMIRREMOLQUE", "BATEA"];
  const valorActualNoEstandar = tipo === "acoplado" && vehiculo.tipo && !OPCIONES_TIPO_ACOPLADO.includes(vehiculo.tipo);

  function handleGuardar() {
    const cambios = {
      patente: patente.toUpperCase(),
      km: Math.max(0, parseInt(km, 10) || 0),
    };
    if (tipo === "camion") cambios.modelo = modeloTipo;
    else cambios.tipo = modeloTipo;
    onGuardar(cambios);
    onCerrar();
  }

  return (
    <div className="mb-4 rounded-md p-3" style={{ background: "var(--panel-2)", border: "1px solid var(--accent-yellow)" }}>
      <div className="text-xs uppercase tracking-wide mb-3" style={{ color: "var(--text-muted)" }}>Editar datos</div>
      <div className="grid sm:grid-cols-3 gap-3">
        <Field label="Patente">
          <input className={inputCls} value={patente} onChange={(e) => setPatente(e.target.value)} />
        </Field>
        <Field label={tipo === "camion" ? "Modelo" : "Tipo"}>
          {tipo === "camion" ? (
            <input className={inputCls} value={modeloTipo} onChange={(e) => setModeloTipo(e.target.value)} />
          ) : (
            <select className={inputCls} value={modeloTipo} onChange={(e) => setModeloTipo(e.target.value)}>
              {valorActualNoEstandar && <option value={vehiculo.tipo}>{vehiculo.tipo} (valor actual)</option>}
              {!vehiculo.tipo && <option value="">Elegir tipo</option>}
              <option value="TOLVA">TOLVA</option>
              <option value="SEMIRREMOLQUE">SEMIRREMOLQUE</option>
              <option value="BATEA">BATEA</option>
            </select>
          )}
        </Field>
        <Field label="Km actuales">
          <input type="number" min="0" className={inputCls} value={km} onChange={(e) => setKm(e.target.value)} />
        </Field>
      </div>
      <div className="flex gap-2 mt-3">
        <button onClick={handleGuardar} className="px-4 py-2 rounded-md text-sm font-semibold" style={{ background: "var(--accent-green)", color: "#0A1128" }}>
          Guardar cambios
        </button>
        <button onClick={onCerrar} className="px-4 py-2 rounded-md text-sm font-semibold" style={{ background: "var(--panel)", border: "1px solid var(--border)", color: "var(--text-muted)" }}>
          Cancelar
        </button>
      </div>
    </div>
  );
}

function InfoRow({ label, value }) {
  return (
    <div className="flex items-center justify-between py-2" style={{ borderBottom: "1px solid var(--border)" }}>
      <span className="text-xs uppercase tracking-wide" style={{ color: "var(--text-muted)" }}>{label}</span>
      <span className="text-sm font-mono">{value}</span>
    </div>
  );
}

function DetailPanel({ selected, onClose, camiones, acoplados, choferes, neumaticos, historialViajes, rutas, onRegistrarViaje, onEditarViaje, onEliminarViaje, onMarcarRecapado, onCambiarVehiculo, onAsignarPosicion, onQuitarPosicion, onVerNeumatico, onAgregarService, onEditarServicio, onActualizarVtv, onActualizarIntervalo, onActualizarValvulas, onActualizarIntercement, onEditarVehiculo, onEditarChofer, onEditarNeumatico, onEliminarCamion, onEliminarAcoplado, onEliminarChofer, onEliminarNeumatico, onSubirPoliza, onDescargarPoliza, onEliminarPoliza, esAdmin }) {
  const [editandoVehiculo, setEditandoVehiculo] = useState(false);
  const { type, id } = selected;
  let title = "", subtitle = "", body = null;

  if (type === "camion") {
    const c = camiones.find((x) => x.id === id);
    if (!c) return null;
    const kmReal = kmRealCamion(c, historialViajes);
    const neus = neumaticos.filter((n) => n.camionId === c.id);
    const intervalo = c.intervaloService || SERVICE_INTERVAL_DEFAULT;
    const nextService = c.ultimoServiceKm + intervalo;
    const kmParaService = Math.max(0, nextService - kmReal);
    const vencVtv = c.vtvUltima ? vencimientoVTV(c.vtvUltima) : null;
    const dVtv = vencVtv ? daysUntil(vencVtv) : null;
    title = c.patente; subtitle = c.modelo;
    body = (
      <>
        {editandoVehiculo ? (
          <EditarVehiculoForm vehiculo={c} kmReal={kmReal} tipo="camion" onGuardar={(cambios) => onEditarVehiculo("camion", c.id, cambios)} onCerrar={() => setEditandoVehiculo(false)} />
        ) : (
          <button onClick={() => setEditandoVehiculo(true)} className="text-xs underline mb-3" style={{ color: "var(--accent-yellow)" }}>
            Editar datos del camión
          </button>
        )}
        <InfoRow label="Km total" value={`${fmtKm(kmReal)} km`} />
        <InfoRow label="Último service" value={`${fmtKm(c.ultimoServiceKm)} km`} />
        <IntervaloServiceRow intervalo={intervalo} onGuardar={(v) => onActualizarIntervalo("camion", c.id, v)} />
        <InfoRow
          label="Próximo service"
          value={<span style={{ color: kmParaService <= 5000 ? (kmParaService === 0 ? "var(--accent-red)" : "var(--accent-yellow)") : "var(--text)" }}>
            {fmtKm(nextService)} km (faltan {fmtKm(kmParaService)})
          </span>}
        />
        <VtvRow vtvUltima={c.vtvUltima} dVtv={dVtv} vencVtv={vencVtv} onGuardar={(f) => onActualizarVtv("camion", c.id, f)} />
        <FechaEditableRow label="Verificación Adicional Intercement" fecha={c.vencimientoIntercement} onGuardar={(f) => onActualizarIntercement("camion", c.id, f)} />
        <PolizaSeguro camion={c} onSubir={onSubirPoliza} onDescargar={onDescargarPoliza} onEliminar={onEliminarPoliza} />
        <ServiceSection vehiculo={c} kmReal={kmReal} tipo="camion" onAgregarService={onAgregarService} onEditarServicio={onEditarServicio} />
        <PosicionesNeumaticos
          tipo="camion" vehiculo={c} ejes={CAMION_EJES}
          neumaticos={neumaticos} neumaticosDelVehiculo={neus}
          onAsignar={onAsignarPosicion} onQuitar={onQuitarPosicion} onVerNeumatico={onVerNeumatico}
        />
      </>
    );
  }

  if (type === "acoplado") {
    const a = acoplados.find((x) => x.id === id);
    if (!a) return null;
    const kmReal = kmRealAcoplado(a, historialViajes);
    const neus = neumaticos.filter((n) => n.acopladoId === a.id);
    const intervalo = a.intervaloService || SERVICE_INTERVAL_DEFAULT;
    const nextService = a.ultimoServiceKm + intervalo;
    const kmParaService = Math.max(0, nextService - kmReal);
    const vencVtv = a.vtvUltima ? vencimientoVTV(a.vtvUltima) : null;
    const dVtv = vencVtv ? daysUntil(vencVtv) : null;
    title = a.patente; subtitle = a.tipo;
    body = (
      <>
        {editandoVehiculo ? (
          <EditarVehiculoForm vehiculo={a} kmReal={kmReal} tipo="acoplado" onGuardar={(cambios) => onEditarVehiculo("acoplado", a.id, cambios)} onCerrar={() => setEditandoVehiculo(false)} />
        ) : (
          <button onClick={() => setEditandoVehiculo(true)} className="text-xs underline mb-3" style={{ color: "var(--accent-yellow)" }}>
            Editar datos del acoplado
          </button>
        )}
        <InfoRow label="Km total" value={`${fmtKm(kmReal)} km`} />
        <InfoRow label="Último service" value={`${fmtKm(a.ultimoServiceKm)} km`} />
        <IntervaloServiceRow intervalo={intervalo} onGuardar={(v) => onActualizarIntervalo("acoplado", a.id, v)} />
        <InfoRow
          label="Próximo service"
          value={<span style={{ color: kmParaService <= 5000 ? (kmParaService === 0 ? "var(--accent-red)" : "var(--accent-yellow)") : "var(--text)" }}>
            {fmtKm(nextService)} km (faltan {fmtKm(kmParaService)})
          </span>}
        />
        <VtvRow vtvUltima={a.vtvUltima} dVtv={dVtv} vencVtv={vencVtv} onGuardar={(f) => onActualizarVtv("acoplado", a.id, f)} />
        <FechaEditableRow label="Verificación Adicional Intercement" fecha={a.vencimientoIntercement} onGuardar={(f) => onActualizarIntercement("acoplado", a.id, f)} />
        {a.tipo === "TOLVA" && (
          <ValvulasTolvaRow vencValvula1={a.vencValvula1} vencValvula2={a.vencValvula2} onGuardar={(v1, v2) => onActualizarValvulas(a.id, v1, v2)} />
        )}
        <ServiceSection vehiculo={a} kmReal={kmReal} tipo="acoplado" onAgregarService={onAgregarService} onEditarServicio={onEditarServicio} />
        <PosicionesNeumaticos
          tipo="acoplado" vehiculo={a} ejes={ACOPLADO_EJES}
          neumaticos={neumaticos} neumaticosDelVehiculo={neus}
          onAsignar={onAsignarPosicion} onQuitar={onQuitarPosicion} onVerNeumatico={onVerNeumatico}
        />
      </>
    );
  }

  if (type === "chofer") {
    const d = choferes.find((x) => x.id === id);
    if (!d) return null;
    const camion = camiones.find((c) => c.id === d.camionId);
    const acoplado = acoplados.find((a) => a.id === d.acopladoId);
    const viajes = historialViajes.filter((v) => v.choferId === d.id).sort((a, b) => b.fecha.localeCompare(a.fecha));
    title = d.nombre; subtitle = "Chofer";
    body = (
      <>
        {editandoVehiculo ? (
          <EditarChoferForm chofer={d} camiones={camiones} acoplados={acoplados} onGuardar={(cambios) => onEditarChofer(d.id, cambios)} onCerrar={() => setEditandoVehiculo(false)} />
        ) : (
          <button onClick={() => setEditandoVehiculo(true)} className="text-xs underline mb-3" style={{ color: "var(--accent-yellow)" }}>
            Editar datos del chofer
          </button>
        )}
        <InfoRow label="Km total" value={`${fmtKm(kmRealChofer(d, historialViajes))} km`} />
        <InfoRow label="Camión predeterminado" value={camion?.patente || "Sin asignar"} />
        <InfoRow label="Acoplado predeterminado" value={acoplado?.patente || "Sin asignar"} />
        {(d.vencimientoDni || d.vencimientoLicencia || d.fechaCursoDefensivo || d.fechaEstudiosMedicos || d.fechaSvo || d.fechaArt || d.fechaEepp) && (
          <div className="mt-2">
            <div className="text-xs uppercase tracking-wide mb-1" style={{ color: "var(--text-muted)" }}>Documentación</div>
            {d.vencimientoDni && <InfoRow label="Vencimiento DNI" value={<span style={{ color: colorVencimiento(d.vencimientoDni) }}>{d.vencimientoDni}</span>} />}
            {d.vencimientoLicencia && <InfoRow label="Vencimiento licencia" value={<span style={{ color: colorVencimiento(d.vencimientoLicencia) }}>{d.vencimientoLicencia}</span>} />}
            {d.fechaCursoDefensivo && <InfoRow label="Curso defensivo" value={<span style={{ color: colorVencimiento(d.fechaCursoDefensivo) }}>{d.fechaCursoDefensivo}</span>} />}
            {d.fechaEstudiosMedicos && <InfoRow label="Estudios médicos" value={<span style={{ color: colorVencimiento(d.fechaEstudiosMedicos) }}>{d.fechaEstudiosMedicos}</span>} />}
            {d.fechaSvo && <InfoRow label="SVO" value={<span style={{ color: colorVencimiento(d.fechaSvo) }}>{d.fechaSvo}</span>} />}
            {d.fechaArt && <InfoRow label="ART" value={<span style={{ color: colorVencimiento(d.fechaArt) }}>{d.fechaArt}</span>} />}
            {d.fechaEepp && <InfoRow label="EEPP" value={<span style={{ color: colorVencimiento(d.fechaEepp) }}>{d.fechaEepp}</span>} />}
          </div>
        )}
        <ViajeChoferForm chofer={d} camiones={camiones} acoplados={acoplados} rutas={rutas} onRegistrarViaje={onRegistrarViaje} />
        <div className="mt-4">
          <div className="text-xs uppercase tracking-wide mb-2" style={{ color: "var(--text-muted)" }}>Historial de viajes</div>
          {viajes.length === 0 ? (
            <p className="text-sm" style={{ color: "var(--text-muted)" }}>Todavía no registró viajes.</p>
          ) : viajes.map((v) => (
            <ViajeItem
              key={v.id} viaje={v} rutas={rutas}
              camion={camiones.find((c) => c.id === v.camionId)}
              acoplado={acoplados.find((a) => a.id === v.acopladoId)}
              onEditar={(nuevoKm, nuevaFecha, nuevaRutaId, nuevaRutaNombre, nuevoRemito) => onEditarViaje(v.id, nuevoKm, nuevaFecha, nuevaRutaId, nuevaRutaNombre, nuevoRemito)}
              onEliminar={() => onEliminarViaje(v.id)}
            />
          ))}
        </div>
      </>
    );
  }

  if (type === "neumatico") {
    const n = neumaticos.find((x) => x.id === id);
    if (!n) return null;
    const camion = camiones.find((c) => c.id === n.camionId);
    const acoplado = acoplados.find((a) => a.id === n.acopladoId);
    const label = estadoNeumatico(n);
    title = n.etiqueta; subtitle = label;
    body = (
      <>
        {editandoVehiculo ? (
          <EditarNeumaticoForm neumatico={n} onGuardar={(cambios) => onEditarNeumatico(n.id, cambios)} onCerrar={() => setEditandoVehiculo(false)} />
        ) : (
          <button onClick={() => setEditandoVehiculo(true)} className="text-xs underline mb-3" style={{ color: "var(--accent-yellow)" }}>
            Editar datos del neumático
          </button>
        )}
        <InfoRow label="Km acumulados" value={`${fmtKm(n.km)} km`} />
        <InfoRow label="Estado" value={<span style={{ color: colorEstadoNeumatico(label) }}>{label}</span>} />
        <InfoRow label="Montado en" value={camion?.patente || acoplado?.patente || "Sin montar"} />

        <NeumaticoAcciones
          neumatico={n} camiones={camiones} acoplados={acoplados} neumaticos={neumaticos}
          onRecapar={onMarcarRecapado} onCambiarVehiculo={onCambiarVehiculo}
        />

        <div className="mt-4">
          <div className="text-xs uppercase tracking-wide mb-2" style={{ color: "var(--text-muted)" }}>Historial</div>
          {n.historial.map((h, i) => (
            <div key={i} className="flex items-start gap-2 text-sm py-1.5" style={{ borderBottom: "1px solid var(--border)" }}>
              <Clock size={13} className="mt-0.5" style={{ color: "var(--text-muted)", flexShrink: 0 }} />
              <span className="flex-1">{h.detalle}</span>
              <span className="font-mono text-xs" style={{ color: "var(--text-muted)" }}>{h.fecha}</span>
            </div>
          ))}
        </div>
      </>
    );
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(0,0,0,0.6)" }} onClick={onClose}>
      <div
        onClick={(e) => e.stopPropagation()}
        className="w-full max-w-4xl rounded-lg flex flex-col"
        style={{ background: "var(--panel)", border: "1px solid var(--border)", maxHeight: "85vh" }}
      >
        <div className="flex items-start justify-between gap-2 px-6 pt-6 pb-4" style={{ borderBottom: "1px solid var(--border)" }}>
          <div>
            <h2 className="font-display text-2xl font-semibold uppercase mb-0.5">{title}</h2>
            <p className="text-sm" style={{ color: "var(--text-muted)" }}>{subtitle}</p>
          </div>
          <button onClick={onClose} className="p-1 rounded hover:bg-[var(--panel-2)] flex-shrink-0">
            <X size={18} style={{ color: "var(--text-muted)" }} />
          </button>
        </div>
        <div className="px-6 py-4 overflow-y-auto">
          {body}
          {esAdmin && (
            <div className="mt-6 pt-4" style={{ borderTop: "1px solid var(--border)" }}>
              <div className="text-xs uppercase tracking-wide mb-2" style={{ color: "var(--text-muted)" }}>Zona de peligro</div>
              <ConfirmDeleteButton
                label={`Eliminar ${title}`}
                onConfirm={() => {
                  if (type === "camion") onEliminarCamion(id);
                  if (type === "acoplado") onEliminarAcoplado(id);
                  if (type === "chofer") onEliminarChofer(id);
                  if (type === "neumatico") onEliminarNeumatico(id);
                }}
              />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
