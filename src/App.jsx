import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import { login, reclamarCuenta, registrarNuevo, aprobarUsuario, rechazarUsuario, getPerfil, logout, supabase } from "./lib/auth";
import {
  Home, Users, MessageSquare, Calendar, FolderOpen, ClipboardCheck, FileText,
  Bell, Search, Menu, Moon, Sun, Settings, Plus, ChevronRight, ChevronLeft,
  MapPin, Heart, MessageCircle, Paperclip, TrendingUp, CheckCircle2, XCircle,
  Clock, Download, Building2, School, BarChart3, Megaphone, Pin, Send, X,
  Shield, UserCog, GraduationCap, Award, Cake, QrCode, Sparkles, Image as ImageIcon,
  FileSpreadsheet, ShieldCheck, LogOut, Printer, PenLine, Trash2, Circle, Save, Lock,
} from "lucide-react";

/* ---------- helpers ---------- */
const cx = (...c) => c.filter(Boolean).join(" ");
const rgba = (hex, a) => {
  const n = parseInt(hex.slice(1), 16);
  return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
};
const initials = (name) => name.split(" ").slice(0, 2).map((w) => w[0]).join("").toUpperCase();
const firstName = (name) => name.split(" ")[0];

/* ---------- cumpleaños helpers ---------- */
const MONTHS_ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
const parseNac = (nac) => {
  if (!nac) return null;
  const p = nac.split("-").map(Number);
  const [m, d] = p.length === 3 ? [p[1], p[2]] : [p[0], p[1]];
  return m && d ? { m, d } : null;
};
const daysUntilNac = (nac, today = new Date()) => {
  const b = parseNac(nac);
  if (b == null) return Infinity;
  const t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  let next = new Date(t0.getFullYear(), b.m - 1, b.d);
  if (next < t0) next = new Date(t0.getFullYear() + 1, b.m - 1, b.d);
  return Math.round((next - t0) / 86400000);
};
const nacToMMDD = (fecha) => {
  if (!fecha) return "";
  const p = fecha.split("-");
  return p.length === 3 ? `${p[1]}-${p[2]}` : fecha;
};

/* ---------- static data ---------- */
const ROLES = {
  admin_general: { label: "Administrador General", short: "Admin. General", icon: Shield, color: "#1d4ed8", user: "Gloria González", microId: null },
  admin_micro: { label: "Admin. de Microcentro", short: "Admin. Microcentro", icon: UserCog, color: "#059669", user: "Pamela Vega", microId: 1 },
  presidente: { label: "Presidente de Microcentro", short: "Presidente", icon: Award, color: "#d97706", user: "Sin asignar", microId: 1 },
  docente: { label: "Docente", short: "Docente", icon: GraduationCap, color: "#0891b2", user: "Antonio Godoy", microId: 1 },
};

const NAV = [
  { id: "dashboard", label: "Inicio", icon: Home, roles: ["admin_general", "admin_micro", "presidente", "docente"] },
  { id: "microcentros", label: "Microcentros", icon: Building2, roles: ["admin_general", "admin_micro", "presidente", "docente"] },
  { id: "grupo", label: "Grupo", icon: Users, roles: ["admin_general", "admin_micro", "presidente", "docente"] },
  { id: "foro", label: "Foro", icon: MessageSquare, roles: ["admin_general", "admin_micro", "presidente", "docente"] },
  { id: "asistencia", label: "Asistencia", icon: ClipboardCheck, roles: ["admin_general", "admin_micro", "presidente", "docente"] },
  { id: "actas", label: "Actas", icon: FileText, roles: ["admin_general", "admin_micro", "presidente", "docente"] },
  { id: "biblioteca", label: "Biblioteca", icon: FolderOpen, roles: ["admin_general", "admin_micro", "presidente", "docente"] },
  { id: "calendario", label: "Calendario", icon: Calendar, roles: ["admin_general", "admin_micro", "presidente", "docente"] },
  { id: "admin", label: "Panel Admin", icon: Settings, roles: ["admin_general"] },
];

const MICROCENTROS = [
  { id: 1, nombre: "Caminante del Saber", presidente: "Por definir", comuna: "Ovalle", integrantes: 7, establecimientos: 7, color: "#1d4ed8" },
  { id: 2, nombre: "DIAM", presidente: "Por definir", comuna: "Ovalle", integrantes: 8, establecimientos: 6, color: "#059669" },
  { id: 3, nombre: "Amanecer Alegre", presidente: "Por definir", comuna: "Ovalle", integrantes: 6, establecimientos: 5, color: "#0891b2" },
  { id: 4, nombre: "Brisa Fresca", presidente: "Por definir", comuna: "Ovalle", integrantes: 5, establecimientos: 5, color: "#7c3aed" },
  { id: 5, nombre: "Sembradores del Futuro", presidente: "Por definir", comuna: "Ovalle", integrantes: 6, establecimientos: 6, color: "#ea580c" },
];
const microName = (id) => (id === 0 ? "Común" : MICROCENTROS.find((m) => m.id === id)?.nombre || "—");
const microColor = (id) => (id === 0 ? "#64748b" : MICROCENTROS.find((m) => m.id === id)?.color || "#64748b");

const ESTAB_BY_MICRO = {
  1: ["Esc. La Quiroga", "Esc. Los Ciénagos", "Esc. La Cebada", "Esc. La Calera", "Esc. Los Azahares de Cerro Blanco", "Esc. Los Canelos", "Esc. Los Trigos"],
  2: ["Esc. La Araucana", "Esc. Lomas de Punilla", "Esc. Ecológica Pejerreyes", "Esc. El Sauce", "Esc. Oruro", "Esc. Las Sossas"],
  3: ["Esc. Algarrobo", "Esc. Salala", "Esc. La Paloma", "Esc. Cordillera", "Esc. Elena Caffarena"],
  4: ["Esc. Caleta El Maitén", "Esc. Padre Pablo Diehl", "Esc. Delia Andrade de Tello", "Esc. El Teniente", "Esc. Caleta El Toro"],
  5: ["Esc. El Espinal", "Esc. Santa Catalina", "Esc. Valle Hermoso", "Esc. Canelilla", "Esc. Villorrio el Talhuén", "Esc. Tranque Recoleta"],
};

const MEETINGS = [
  { id: "m1", titulo: "Reunión mensual de coordinación", fecha: "8 Jul 2025", hora: "15:00", lugar: "Esc. La Quiroga", micro: "Caminante del Saber" },
  { id: "m2", titulo: "Taller de evaluación formativa", fecha: "10 Jul 2025", hora: "16:30", lugar: "DAEM Ovalle", micro: "Caminante del Saber" },
];

const NOTICIAS = [
  { titulo: "Nuevas orientaciones curriculares 2025", fecha: "2 Jul", pinned: true },
  { titulo: "Capacitación en tecnología educativa", fecha: "30 Jun", pinned: false },
  { titulo: "Postulación a fondos de innovación docente", fecha: "27 Jun", pinned: false },
];

const CATEGORIAS = [
  { name: "Consultas pedagógicas", count: 0, color: "#1d4ed8" },
  { name: "Material educativo", count: 0, color: "#059669" },
  { name: "Tecnología", count: 0, color: "#0891b2" },
  { name: "Inclusión", count: 0, color: "#7c3aed" },
  { name: "Convivencia escolar", count: 0, color: "#db2777" },
  { name: "Proyectos", count: 0, color: "#ea580c" },
];

const THREADS = [];

const FOLDERS = [
  { name: "Actas", count: 0, color: "#1d4ed8" }, { name: "Oficios", count: 0, color: "#0891b2" },
  { name: "Decretos", count: 0, color: "#7c3aed" }, { name: "Bases curriculares", count: 0, color: "#059669" },
  { name: "Material pedagógico", count: 0, color: "#ea580c" }, { name: "Presentaciones", count: 0, color: "#db2777" },
  { name: "Protocolos", count: 0, color: "#0d9488" }, { name: "Formularios", count: 0, color: "#6366f1" },
];

const FILES = [];

const PALETTE = ["#1d4ed8", "#059669", "#0891b2", "#7c3aed", "#db2777", "#ea580c", "#0d9488", "#6366f1"];
const MEMBERS_BY_MICRO = {
  1: [
    { name: "Pamela Vega", rol: "Docente · Esc. La Quiroga" },
    { name: "Antonio Godoy", rol: "Docente · Esc. Los Ciénagos" },
    { name: "Daniela Lobos", rol: "Docente · Esc. La Cebada" },
    { name: "Eugenio Sierra", rol: "Docente · Esc. La Calera" },
    { name: "Michel Díaz", rol: "Docente · Esc. Los Azahares" },
    { name: "Miguel Zamora", rol: "Docente · Esc. Los Canelos" },
    { name: "Claudia Saavedra", rol: "Docente · Esc. Los Trigos" },
  ],
  2: [
    { name: "Elizabeth Carrera", rol: "Docente · Esc. La Araucana" },
    { name: "Julián Díaz", rol: "Docente · Esc. Lomas de Punilla" },
    { name: "Karen Díaz", rol: "Docente · Esc. Ecológica Pejerreyes" },
    { name: "Julio Araya", rol: "Docente · Esc. El Sauce" },
    { name: "Janeth Bravo", rol: "Docente · Esc. Ecológica Pejerreyes" },
    { name: "Sandy Araya", rol: "Docente · Esc. La Araucana" },
    { name: "Macarena Marín", rol: "Docente · Esc. Oruro" },
    { name: "Sandra Castro", rol: "Docente · Esc. Las Sossas" },
  ],
  3: [
    { name: "Amerito Ardiles", rol: "Docente · Esc. Algarrobo" },
    { name: "Cristina Tatán", rol: "Docente · Esc. Salala" },
    { name: "Gloria González", rol: "Docente · Esc. La Paloma" },
    { name: "Sara Castro", rol: "Docente · Esc. Cordillera" },
    { name: "Giovanni Araya", rol: "Docente · Esc. Salala" },
    { name: "Jocelyn Chacana", rol: "Docente · Esc. Elena Caffarena" },
  ],
  4: [
    { name: "Ruth Amaya", rol: "Docente · Esc. Caleta El Maitén" },
    { name: "Cristina Grandón", rol: "Docente · Esc. Padre Pablo Diehl" },
    { name: "Bárbara Zegers", rol: "Docente · Esc. Delia Andrade de Tello" },
    { name: "Maleny Robles", rol: "Docente · Esc. El Teniente" },
    { name: "Lilian López", rol: "Docente · Esc. Caleta El Toro" },
  ],
  5: [
    { name: "Ana Elia Cortés", rol: "Docente · Esc. El Espinal" },
    { name: "Guacolda Villalobos", rol: "Docente · Esc. Santa Catalina" },
    { name: "Edith Tania", rol: "Docente · Esc. Valle Hermoso" },
    { name: "Clara Torrejón", rol: "Docente · Esc. Canelilla" },
    { name: "Máximo Cortés", rol: "Docente · Esc. Villorrio el Talhuén" },
    { name: "Arnoldo Lanas", rol: "Docente · Esc. Tranque Recoleta" },
  ],
};

const BIRTHDAYS = {
  "Pamela Vega": "05-20", "Antonio Godoy": "07-08", "Daniela Lobos": "09-14", "Eugenio Sierra": "07-22",
  "Michel Díaz": "11-03", "Miguel Zamora": "07-07", "Claudia Saavedra": "02-11",
  "Elizabeth Carrera": "03-05", "Julián Díaz": "07-15", "Karen Díaz": "06-28", "Julio Araya": "10-19",
  "Janeth Bravo": "01-30", "Sandy Araya": "08-02", "Macarena Marín": "07-25", "Sandra Castro": "04-12",
  "Amerito Ardiles": "12-01", "Cristina Tatán": "07-19", "Gloria González": "09-09", "Sara Castro": "03-23",
  "Giovanni Araya": "07-11", "Jocelyn Chacana": "05-06",
  "Ruth Amaya": "07-30", "Cristina Grandón": "02-17", "Bárbara Zegers": "08-21", "Maleny Robles": "07-13", "Lilian López": "11-28",
  "Ana Elia Cortés": "07-04", "Guacolda Villalobos": "06-15", "Edith Tania": "09-27", "Clara Torrejón": "07-17",
  "Máximo Cortés": "01-08", "Arnoldo Lanas": "07-06",
};
Object.values(MEMBERS_BY_MICRO).forEach((list) => list.forEach((m, i) => {
  m.color = PALETTE[i % PALETTE.length];
  m.nac = BIRTHDAYS[m.name] || "";
}));
const MEMBERS = MEMBERS_BY_MICRO[1];

const AUDIT = [
  { icon: LogOut, txt: "Inicio de sesión — Antonio Godoy", t: "Hoy 09:12" },
  { icon: FileText, txt: "Subida de documento — Acta N°5", t: "Ayer 17:40" },
  { icon: XCircle, txt: "Eliminación — borrador de acta", t: "Ayer 11:02" },
  { icon: UserCog, txt: "Permisos actualizados — Karen Díaz", t: "1 Jul 15:20" },
];

const EVENT_DAYS = {
  8: { label: "Reunión coordinación", color: "#1d4ed8" },
  10: { label: "Taller ev. formativa", color: "#059669" },
  15: { label: "Visita técnica DAEM", color: "#ea580c" },
  22: { label: "Capacitación TIC", color: "#0891b2" },
  28: { label: "Consejo de microcentro", color: "#7c3aed" },
};

const INITIAL_ACTAS = [
  { id: "a5", titulo: "Acta N°5 — Planificación anual", fecha: "28 Jun 2025", lugar: "Esc. La Quiroga", autor: "Pamela Vega", participantes: 7, v: 3,
    temas: ["Revisión de metas del primer semestre", "Organización del taller de evaluación formativa"],
    acuerdos: [{ desc: "Enviar planificación de Ciencias", resp: "Antonio Godoy", fecha: "10 Jul", done: false }, { desc: "Reservar sala en DAEM", resp: "Eugenio Sierra", fecha: "8 Jul", done: true }] },
  { id: "a4", titulo: "Acta N°4 — Convivencia escolar", fecha: "14 Jun 2025", lugar: "DAEM Ovalle", autor: "Eugenio Sierra", participantes: 6, v: 2,
    temas: ["Protocolo de convivencia", "Casos de mediación"],
    acuerdos: [{ desc: "Difundir protocolo actualizado", resp: "Daniela Lobos", fecha: "20 Jun", done: false }] },
  { id: "a3", titulo: "Acta N°3 — Uso de recursos TIC", fecha: "31 May 2025", lugar: "Esc. La Cebada", autor: "Daniela Lobos", participantes: 7, v: 1,
    temas: ["Inventario de tablets", "Plan de uso en aula"],
    acuerdos: [{ desc: "Capacitación básica de tablets", resp: "Miguel Zamora", fecha: "12 Jun", done: true }] },
];

const BLANK_REG = { nombre: "", correo: "", rut: "", telefono: "", rol: "docente", micro: "Caminante del Saber", establecimiento: "", password: "", confirm: "", terms: false };

/* ================= APP ================= */
export default function App() {
  const [authed, setAuthed] = useState(false);
  const [authView, setAuthView] = useState("login");
  const [reg, setReg] = useState(BLANK_REG);
  const [regError, setRegError] = useState("");
  const [regMode, setRegMode] = useState("claim");
  const [claimed, setClaimed] = useState(["Pamela Vega", "Antonio Godoy", "Gloria González"]);
  const [successKind, setSuccessKind] = useState("nuevo");
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPass, setLoginPass] = useState("");
  const [loginError, setLoginError] = useState("");
  const [authLoading, setAuthLoading] = useState(false);
  const [myMicroId, setMyMicroId] = useState(null);
  const [pendingRequests, setPendingRequests] = useState([]);
  const [dark, setDark] = useState(false);
  const [role, setRole] = useState("admin_general");
  const [screen, setScreen] = useState("dashboard");
  const [drawer, setDrawer] = useState(false);
  const [micro, setMicro] = useState(null);
  const [microTab, setMicroTab] = useState("info");
  const [microcentros, setMicrocentros] = useState(() => MICROCENTROS.map((m) => ({ ...m })));
  const [membersByMicro, setMembersByMicro] = useState(() => {
    const clone = {};
    Object.entries(MEMBERS_BY_MICRO).forEach(([k, v]) => { clone[k] = v.map((m) => ({ ...m })); });
    return clone;
  });
  const [manageMember, setManageMember] = useState(null); // { microId, index }
  const [thread, setThread] = useState(null);
  const [composer, setComposer] = useState("");
  const [search, setSearch] = useState("");
  const [toast, setToast] = useState(null);
  const [editingPostId, setEditingPostId] = useState(null);
  const [editPostText, setEditPostText] = useState("");

  const [profiles, setProfiles] = useState({
    admin_general: { nombre: "Gloria González", correo: "gloriagonro@gmail.com", telefono: "+56 9 5555 1234", rut: "14.567.890-1", fechaNacimiento: "1975-09-09", establecimiento: "Esc. La Paloma", cargo: "Administradora General", notif: { correo: true, reuniones: true, resumen: true } },
    admin_micro: { nombre: "Pamela Vega", correo: "pamela.vega@daemovalle.cl", telefono: "+56 9 8765 4321", rut: "15.234.567-8", fechaNacimiento: "1982-05-20", establecimiento: "Esc. La Quiroga", cargo: "Administradora de Microcentro", notif: { correo: true, reuniones: true, resumen: true } },
    presidente: { nombre: "Sin asignar", correo: "", telefono: "", rut: "", fechaNacimiento: "", establecimiento: "", cargo: "Presidente de Microcentro", notif: { correo: true, reuniones: true, resumen: true } },
    docente: { nombre: "Antonio Godoy", correo: "antonio.godoy@daemovalle.cl", telefono: "+56 9 1234 5678", rut: "16.789.012-3", fechaNacimiento: "1988-07-08", establecimiento: "Esc. Los Ciénagos", cargo: "Docente", notif: { correo: true, reuniones: true, resumen: false } },
  });

  const [posts, setPosts] = useState([
    { id: 2, autor: "Antonio Godoy", rol: "Docente · Esc. Los Ciénagos", tiempo: "hace 2 h", texto: "Comparto la planificación de la unidad de Ciencias que trabajamos en el taller. ¡Espero les sirva!", likes: 12, comentarios: 4, liked: false, attach: "Planificacion_Ciencias_U3.pdf", color: "#059669" },
    { id: 1, autor: "Pamela Vega", rol: "Docente · Esc. La Quiroga", tiempo: "ayer", texto: "Recordatorio: la próxima reunión será en la Escuela La Quiroga. Confirmen su asistencia en el módulo correspondiente.", likes: 8, comentarios: 6, liked: true, color: "#1d4ed8" },
  ]);

  // Asistencia
  const [meetings, setMeetings] = useState(() => MEETINGS.map((m) => ({ ...m })));
  const [meetingId, setMeetingId] = useState("m1");
  const [att, setAtt] = useState({});
  const [horas, setHoras] = useState({});
  const [obs, setObs] = useState({});
  const [firmado, setFirmado] = useState(null);
  const [printOpen, setPrintOpen] = useState(false);
  const [newMeetingOpen, setNewMeetingOpen] = useState(false);
  const [meetDraft, setMeetDraft] = useState({ titulo: "", fecha: "", hora: "", lugar: "" });

  // Foro / Biblioteca scoping
  const [foroScope, setForoScope] = useState("comun");
  const [bibMicroFilter, setBibMicroFilter] = useState("all");
  const [bibCat, setBibCat] = useState("all");

  // Actas
  const [actas, setActas] = useState(INITIAL_ACTAS);
  const [openActa, setOpenActa] = useState(null);
  const [editing, setEditing] = useState(false);
  const blankDraft = { id: null, titulo: "", fecha: "", lugar: "", participants: [], temas: [], acuerdos: [], v: 0 };
  const [draft, setDraft] = useState(blankDraft);
  const [temaInput, setTemaInput] = useState("");
  const [aDesc, setADesc] = useState("");
  const [aResp, setAResp] = useState(MEMBERS[0].name);
  const [aFecha, setAFecha] = useState("");

  const t = dark
    ? { page: "bg-slate-950", surface: "bg-slate-900", alt: "bg-slate-800", border: "border-slate-800", text: "text-slate-100", muted: "text-slate-400", hover: "hover:bg-slate-800", input: "bg-slate-800 border-slate-700 text-slate-100" }
    : { page: "bg-slate-50", surface: "bg-white", alt: "bg-slate-50", border: "border-slate-200", text: "text-slate-900", muted: "text-slate-500", hover: "hover:bg-slate-100", input: "bg-white border-slate-300 text-slate-900" };

  const roleCfg = ROLES[role];
  const visibleNav = NAV.filter((n) => n.roles.includes(role));
  const isAdmin = role === "admin_general" || role === "admin_micro";
  // admin_general SIEMPRE ve todos los microcentros, sin importar si su ficha
  // quedó con una membresía vieja (ej: si antes era docente y fue promovido).
  const assignedMicroId = role === "admin_general" ? null : (myMicroId ?? roleCfg.microId);
  const myMicroName = assignedMicroId ? microName(assignedMicroId) : "Todos los microcentros";
  const font = { fontFamily: 'ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif' };

  // Trae del backend real la lista de personas que esperan aprobación de un admin
  const fetchPendingRequests = async () => {
    const { data, error } = await supabase
      .from("usuario")
      .select("id, nombre, correo, rol, establecimiento, created_at, membresia(microcentro(nombre))")
      .eq("estado", "pendiente_aprobacion")
      .order("created_at", { ascending: false });
    if (error) { showToast(`No se pudieron cargar las solicitudes: ${error.message}`); return; }
    setPendingRequests((data || []).map((u) => ({
      id: u.id,
      nombre: u.nombre,
      correo: u.correo,
      rol: u.rol === "admin_micro" ? "Admin. Microcentro" : u.rol === "presidente" ? "Presidente" : "Docente",
      micro: u.membresia?.[0]?.microcentro?.nombre || "—",
      establecimiento: u.establecimiento || "—",
      fecha: u.created_at ? new Date(u.created_at).toLocaleDateString("es-CL") : "",
    })));
  };

  useEffect(() => { if (authed && isAdmin) fetchPendingRequests(); }, [authed, isAdmin]);

  const showToast = (msg) => { setToast(msg); setTimeout(() => setToast(null), 2400); };
  const go = (id) => { setScreen(id); setDrawer(false); setThread(null); setEditing(false); if (id !== "microcentros") setMicro(null); };
  const switchRole = (r) => {
    setRole(r); setMicro(null); setThread(null); setEditing(false); setForoScope("comun"); setBibMicroFilter("all");
    if (screen !== "perfil" && !NAV.find((n) => n.id === screen)?.roles.includes(r)) setScreen("dashboard");
  };

  // ---- Gestionar integrante (modal en Microcentros › Integrantes) ----
  const ROLE_LABELS = ["Docente", "Presidente de Microcentro", "Admin. de Microcentro"];
  const splitRol = (rolStr) => {
    const idx = rolStr.indexOf(" · ");
    return idx === -1 ? { roleLabel: rolStr, estab: "" } : { roleLabel: rolStr.slice(0, idx), estab: rolStr.slice(idx + 3) };
  };
  const openManage = (microId, index) => setManageMember({ microId, index });
  const closeManage = () => setManageMember(null);
  const setManagedRole = (roleLabel) => {
    setMembersByMicro((prev) => {
      const list = [...prev[manageMember.microId]];
      const current = list[manageMember.index];
      const { estab } = splitRol(current.rol);
      list[manageMember.index] = { ...current, rol: estab ? `${roleLabel} · ${estab}` : roleLabel };
      return { ...prev, [manageMember.microId]: list };
    });
    showToast(`Rol actualizado a ${roleLabel}`);
  };
  const removeManagedMember = () => {
    const { microId, index } = manageMember;
    const nombre = membersByMicro[microId][index].name;
    setMembersByMicro((prev) => ({ ...prev, [microId]: prev[microId].filter((_, i) => i !== index) }));
    setMicrocentros((prev) => prev.map((m) => (m.id === microId ? { ...m, integrantes: Math.max(0, m.integrantes - 1) } : m)));
    showToast(`${firstName(nombre)} fue removido del microcentro`);
    setManageMember(null);
  };

  /* ---------- shared components ---------- */
  const Avatar = ({ name, color, size = 38 }) => (
    <div className="flex items-center justify-center rounded-full shrink-0 font-semibold text-white" style={{ width: size, height: size, backgroundColor: color, fontSize: size * 0.36 }}>{initials(name)}</div>
  );
  const Chip = ({ color, children }) => (
    <span className="inline-flex items-center rounded-full px-2.5 py-0.5 text-xs font-medium" style={{ backgroundColor: rgba(color, dark ? 0.22 : 0.12), color: dark ? "#fff" : color }}>{children}</span>
  );
  const Card = ({ children, className }) => <div className={cx("rounded-2xl border", t.border, t.surface, className)}>{children}</div>;
  const Section = ({ title, action, children }) => (
    <Card className="p-4 sm:p-5">
      <div className="flex items-center justify-between mb-3"><h3 className="font-semibold text-sm sm:text-base">{title}</h3>{action}</div>
      {children}
    </Card>
  );
  const Btn = ({ children, icon: Ic, variant = "primary", onClick }) => {
    const styles = { primary: "bg-blue-600 text-white hover:bg-blue-700", emerald: "bg-emerald-600 text-white hover:bg-emerald-700", ghost: cx(t.text, t.hover, "border", t.border) }[variant];
    return <button onClick={onClick} className={cx("inline-flex items-center gap-1.5 rounded-xl px-3 py-2 text-sm font-medium transition", styles)}>{Ic && <Ic size={16} />} {children}</button>;
  };
  const Field = ({ label, ...p }) => (
    <label className="block"><span className={cx("text-xs font-medium", t.muted)}>{label}</span>
      <input {...p} className={cx("mt-1 w-full rounded-xl border px-3 py-2 text-sm outline-none focus:border-blue-500", t.input)} /></label>
  );
  const Toggle = ({ on, onClick }) => (
    <button onClick={onClick} className="relative h-6 w-11 rounded-full transition shrink-0" style={{ backgroundColor: on ? "#059669" : dark ? "#334155" : "#cbd5e1" }}>
      <span className="absolute top-0.5 h-5 w-5 rounded-full bg-white shadow transition-all" style={{ left: on ? "22px" : "2px" }} />
    </button>
  );

  /* ================= LOGIN ================= */
  if (!authed) {
    const inp = "w-full rounded-xl border border-slate-300 px-3 py-2 text-sm outline-none focus:border-blue-500 text-slate-900";
    const lbl = "text-xs font-medium text-slate-500";
    const setR = (k, v) => setReg({ ...reg, [k]: v });
    const rosterFor = (nombre) => { const m = MICROCENTROS.find((x) => x.nombre === nombre); return m ? MEMBERS_BY_MICRO[m.id] || [] : []; };
    const unclaimed = rosterFor(reg.micro).filter((mem) => !claimed.includes(mem.name));
    const selMember = rosterFor(reg.micro).find((mem) => mem.name === reg.nombre);
    const selEscuela = selMember ? (selMember.rol.split("·")[1] || "").trim() : "";
    const gotoLogin = () => { setAuthView("login"); setRegError(""); setRegMode("claim"); };
    const doLogin = async () => {
      setLoginError(""); setAuthLoading(true);
      try {
        await login(loginEmail.trim(), loginPass);
        const perfil = await getPerfil();
        if (!perfil) throw new Error("No se encontró tu ficha de usuario. Contacta a un administrador.");
        if (perfil.estado !== "activo") {
          await logout();
          throw new Error("Tu cuenta aún no está activa. Espera la aprobación de un administrador.");
        }
        setProfiles((prev) => ({
          ...prev,
          [perfil.rol]: {
            ...prev[perfil.rol],
            nombre: perfil.nombre,
            correo: perfil.correo,
            telefono: perfil.telefono || "",
            rut: perfil.rut || "",
            fechaNacimiento: perfil.fecha_nacimiento || prev[perfil.rol]?.fechaNacimiento || "",
            establecimiento: perfil.establecimiento || "",
          },
        }));
        setMyMicroId(perfil.membresia?.[0]?.microcentro_id ?? null);
        setRole(perfil.rol);
        setAuthed(true);
      } catch (err) {
        setLoginError(err.message === "Invalid login credentials" ? "Correo o contraseña incorrectos." : err.message);
      } finally {
        setAuthLoading(false);
      }
    };
    const validCreds = () => {
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(reg.correo)) { setRegError("El correo no tiene un formato válido."); return false; }
      if (reg.password.length < 8) { setRegError("La contraseña debe tener al menos 8 caracteres."); return false; }
      if (reg.password !== reg.confirm) { setRegError("Las contraseñas no coinciden."); return false; }
      return true;
    };
    const submitClaim = async () => {
      if (!reg.nombre) return setRegError("Selecciona tu nombre en la lista.");
      if (!reg.correo.trim() || !reg.password) return setRegError("Ingresa tu correo y una contraseña.");
      if (!validCreds()) return;
      setAuthLoading(true);
      try {
        await reclamarCuenta({ correo: reg.correo.trim(), password: reg.password });
        setClaimed([...claimed, reg.nombre]);
        setRegError(""); setSuccessKind("claim"); setAuthView("success"); setReg(BLANK_REG);
      } catch (err) {
        setRegError(err.message);
      } finally {
        setAuthLoading(false);
      }
    };
    const submitNew = async () => {
      if (!reg.nombre.trim() || !reg.correo.trim() || !reg.password) return setRegError("Completa al menos nombre, correo y contraseña.");
      if (!validCreds()) return;
      if (!reg.terms) return setRegError("Debes aceptar los términos y condiciones.");
      setAuthLoading(true);
      try {
        const microcentroId = MICROCENTROS.find((m) => m.nombre === reg.micro)?.id ?? null;
        await registrarNuevo({
          nombre: reg.nombre.trim(),
          correo: reg.correo.trim(),
          password: reg.password,
          rut: reg.rut.trim(),
          telefono: reg.telefono.trim(),
          rol: reg.rol,
          microcentroId,
          establecimiento: reg.establecimiento.trim(),
        });
        setRegError(""); setSuccessKind("nuevo"); setAuthView("success"); setReg(BLANK_REG);
      } catch (err) {
        setRegError(err.message);
      } finally {
        setAuthLoading(false);
      }
    };
    return (
      <div className="min-h-screen flex items-start sm:items-center justify-center p-4" style={{ ...font, background: "linear-gradient(135deg,#0b2a4a 0%,#0e355c 55%,#0d5c47 100%)" }}>
        <div className={cx("w-full rounded-3xl bg-white p-7 shadow-2xl my-6", authView === "register" ? "max-w-md" : "max-w-sm")}>
          <div className="flex items-center gap-2.5 mb-6">
            <div className="h-11 w-11 rounded-2xl flex items-center justify-center" style={{ backgroundColor: "#10b981" }}><School size={22} color="#fff" /></div>
            <div><div className="font-bold text-slate-900 leading-tight">Microcentros Rurales</div><div className="text-emerald-600 text-xs font-medium">Comuna de Ovalle</div></div>
          </div>

          {authView === "login" && (
            <div className="space-y-3">
              <div><label className={lbl}>Correo o usuario</label>
                <input value={loginEmail} onChange={(e) => { setLoginEmail(e.target.value); setLoginError(""); }} className={cx(inp, "mt-1 py-2.5")} /></div>
              <div><label className={lbl}>Contraseña</label>
                <input type="password" value={loginPass} onChange={(e) => { setLoginPass(e.target.value); setLoginError(""); }} onKeyDown={(e) => e.key === "Enter" && doLogin()} className={cx(inp, "mt-1 py-2.5")} /></div>
              <div className="flex items-center justify-between text-xs">
                <label className="flex items-center gap-1.5 text-slate-600"><input type="checkbox" defaultChecked className="accent-blue-600" /> Mantener sesión iniciada</label>
                <button className="text-blue-600 font-medium">¿Olvidaste tu contraseña?</button>
              </div>
              {loginError && <div className="flex items-center gap-1.5 text-xs text-red-600 bg-red-50 rounded-lg px-3 py-2"><XCircle size={14} />{loginError}</div>}
              <button onClick={doLogin} disabled={authLoading} className="w-full rounded-xl bg-blue-600 text-white font-semibold py-2.5 hover:bg-blue-700 transition disabled:opacity-60">{authLoading ? "Ingresando…" : "Ingresar"}</button>
              <div className="pt-1 text-center text-sm text-slate-500">¿No tienes cuenta?{" "}
                <button onClick={() => { setAuthView("register"); setRegError(""); }} className="text-blue-600 font-semibold">Crear cuenta</button></div>
            </div>
          )}

          {authView === "register" && (
            <div className="space-y-3">
              <div><h2 className="font-bold text-slate-900">Crear cuenta</h2>
                <p className="text-xs text-slate-500">{regMode === "claim" ? "Elige tu microcentro y tu nombre para vincular tu cuenta a la ficha que el DAEM ya tiene registrada." : "Regístrate si aún no apareces en el listado oficial del DAEM."}</p></div>

              {regMode === "claim" ? (
                <>
                  <div><label className={lbl}>Microcentro</label>
                    <select value={reg.micro} onChange={(e) => setReg({ ...reg, micro: e.target.value, nombre: "" })} className={cx(inp, "mt-1")}>{MICROCENTROS.map((m) => <option key={m.id} value={m.nombre}>{m.nombre}</option>)}</select></div>
                  <div><label className={lbl}>Tu nombre</label>
                    {unclaimed.length === 0 ? (
                      <div className="mt-1 text-xs text-slate-500 rounded-lg bg-slate-50 px-3 py-2">Todos los docentes de este microcentro ya tienen cuenta activada.</div>
                    ) : (
                      <select value={reg.nombre} onChange={(e) => setR("nombre", e.target.value)} className={cx(inp, "mt-1")}>
                        <option value="">Selecciona tu nombre…</option>
                        {unclaimed.map((mem) => <option key={mem.name} value={mem.name}>{mem.name}</option>)}
                      </select>
                    )}
                  </div>

                  {reg.nombre && (
                    <div className="rounded-xl border-2 border-emerald-500 bg-emerald-50 p-3">
                      <div className="text-xs text-emerald-700 font-medium">Vas a activar la cuenta de:</div>
                      <div className="text-slate-900 font-bold text-lg leading-tight">{reg.nombre}</div>
                      <div className="text-sm text-slate-600 flex items-center gap-1.5 mt-0.5"><School size={14} className="text-emerald-600" /> {selEscuela} · {reg.micro}</div>
                      <div className="text-[11px] text-slate-500 mt-1.5">Se vinculará tu acceso a esta ficha existente. No se creará un registro duplicado.</div>
                    </div>
                  )}

                  <div><label className={lbl}>Correo electrónico *</label><input value={reg.correo} onChange={(e) => setR("correo", e.target.value)} placeholder="tu correo institucional" className={cx(inp, "mt-1")} /></div>
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className={lbl}>Contraseña *</label><input type="password" value={reg.password} onChange={(e) => setR("password", e.target.value)} placeholder="Mín. 8" className={cx(inp, "mt-1")} /></div>
                    <div><label className={lbl}>Confirmar *</label><input type="password" value={reg.confirm} onChange={(e) => setR("confirm", e.target.value)} placeholder="Repite" className={cx(inp, "mt-1")} /></div>
                  </div>
                  {regError && <div className="flex items-center gap-1.5 text-xs text-red-600 bg-red-50 rounded-lg px-3 py-2"><XCircle size={14} />{regError}</div>}
                  <button onClick={submitClaim} disabled={authLoading} className="w-full rounded-xl bg-emerald-600 text-white font-semibold py-2.5 hover:bg-emerald-700 transition disabled:opacity-60">{authLoading ? "Activando…" : "Activar mi cuenta"}</button>
                  <div className="text-center text-sm text-slate-500"><button onClick={() => { setRegMode("nuevo"); setRegError(""); setReg({ ...BLANK_REG, micro: reg.micro }); }} className="text-blue-600 font-semibold">No encuentro mi nombre en la lista</button></div>
                </>
              ) : (
                <>
                  <div><label className={lbl}>Nombre completo *</label><input value={reg.nombre} onChange={(e) => setR("nombre", e.target.value)} placeholder="Ej: María Pérez" className={cx(inp, "mt-1")} /></div>
                  <div><label className={lbl}>Correo electrónico *</label><input value={reg.correo} onChange={(e) => setR("correo", e.target.value)} placeholder="nombre@daemovalle.cl" className={cx(inp, "mt-1")} /></div>
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className={lbl}>RUT</label><input value={reg.rut} onChange={(e) => setR("rut", e.target.value)} placeholder="12.345.678-9" className={cx(inp, "mt-1")} /></div>
                    <div><label className={lbl}>Teléfono</label><input value={reg.telefono} onChange={(e) => setR("telefono", e.target.value)} placeholder="+56 9 …" className={cx(inp, "mt-1")} /></div>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className={lbl}>Rol solicitado</label>
                      <select value={reg.rol} onChange={(e) => setR("rol", e.target.value)} className={cx(inp, "mt-1")}><option value="docente">Docente</option><option value="presidente">Presidente de Microcentro</option><option value="admin_micro">Admin. de Microcentro</option></select></div>
                    <div><label className={lbl}>Microcentro</label>
                      <select value={reg.micro} onChange={(e) => setR("micro", e.target.value)} className={cx(inp, "mt-1")}>{MICROCENTROS.map((m) => <option key={m.id} value={m.nombre}>{m.nombre}</option>)}</select></div>
                  </div>
                  <div><label className={lbl}>Establecimiento</label><input value={reg.establecimiento} onChange={(e) => setR("establecimiento", e.target.value)} placeholder="Ej: Esc. La Quiroga" className={cx(inp, "mt-1")} /></div>
                  <div className="grid grid-cols-2 gap-2">
                    <div><label className={lbl}>Contraseña *</label><input type="password" value={reg.password} onChange={(e) => setR("password", e.target.value)} placeholder="Mín. 8 caracteres" className={cx(inp, "mt-1")} /></div>
                    <div><label className={lbl}>Confirmar *</label><input type="password" value={reg.confirm} onChange={(e) => setR("confirm", e.target.value)} placeholder="Repite la clave" className={cx(inp, "mt-1")} /></div>
                  </div>
                  <label className="flex items-start gap-2 text-xs text-slate-600"><input type="checkbox" checked={reg.terms} onChange={(e) => setR("terms", e.target.checked)} className="accent-blue-600 mt-0.5" /> Acepto los términos y el tratamiento de mis datos según la normativa vigente.</label>
                  {regError && <div className="flex items-center gap-1.5 text-xs text-red-600 bg-red-50 rounded-lg px-3 py-2"><XCircle size={14} />{regError}</div>}
                  <button onClick={submitNew} disabled={authLoading} className="w-full rounded-xl bg-emerald-600 text-white font-semibold py-2.5 hover:bg-emerald-700 transition disabled:opacity-60">{authLoading ? "Enviando…" : "Enviar registro"}</button>
                  <div className="text-center text-sm text-slate-500"><button onClick={() => { setRegMode("claim"); setRegError(""); }} className="text-blue-600 font-semibold">Volver a la lista de mi microcentro</button></div>
                </>
              )}

              <div className="text-center text-sm text-slate-500 pt-1 border-t border-slate-100">¿Ya tienes cuenta?{" "}<button onClick={gotoLogin} className="text-blue-600 font-semibold">Inicia sesión</button></div>
            </div>
          )}

          {authView === "success" && (
            <div className="text-center py-4">
              <div className="h-14 w-14 rounded-full bg-emerald-100 flex items-center justify-center mx-auto"><CheckCircle2 size={30} className="text-emerald-600" /></div>
              <h2 className="font-bold text-slate-900 mt-3">{successKind === "claim" ? "¡Cuenta activada!" : "¡Registro enviado!"}</h2>
              <p className="text-sm text-slate-500 mt-1.5">{successKind === "claim" ? "Tu cuenta quedó vinculada a tu ficha del microcentro. Ya puedes iniciar sesión con tu correo y contraseña." : "Un administrador de tu microcentro debe aprobar tu solicitud antes de que puedas ingresar. Te avisaremos por correo cuando esté lista."}</p>
              <button onClick={gotoLogin} className="mt-5 w-full rounded-xl bg-blue-600 text-white font-semibold py-2.5 hover:bg-blue-700 transition">Volver a iniciar sesión</button>
            </div>
          )}

          <p className="mt-5 text-center text-[11px] text-slate-400">Departamento de Educación de Ovalle</p>
        </div>
      </div>
    );
  }

  /* ---------- nav ---------- */
  const renderNav = () => (
    <nav className="px-3 py-2 space-y-1">
      {visibleNav.map((item) => {
        const Ic = item.icon; const active = screen === item.id;
        return (
          <button key={item.id} onClick={() => go(item.id)}
            className={cx("w-full flex items-center gap-3 rounded-xl px-3 py-2.5 text-sm font-medium transition", active ? "bg-blue-900 text-white" : "text-blue-100 hover:bg-blue-900")}
            style={active ? { boxShadow: "inset 3px 0 0 #34d399" } : undefined}><Ic size={18} /> {item.label}</button>
        );
      })}
    </nav>
  );
  const brand = (
    <div className="px-5 pt-5 pb-3"><div className="flex items-center gap-2.5">
      <div className="h-9 w-9 rounded-xl flex items-center justify-center" style={{ backgroundColor: "#10b981" }}><School size={19} color="#fff" /></div>
      <div><div className="text-white font-bold text-sm leading-tight">Microcentros</div><div className="text-emerald-300 text-xs">Rurales de Ovalle</div></div>
    </div></div>
  );
  const contextChip = (
    <div className="px-3 pb-2">
      <div className="rounded-xl bg-blue-900 px-3 py-2 flex items-center gap-2">
        <Building2 size={15} className="text-emerald-300 shrink-0" />
        <div className="text-xs text-blue-100 leading-tight min-w-0">
          {assignedMicroId ? "Microcentro" : "Vista global"}<br /><span className="text-white font-medium truncate block">{myMicroName}</span>
        </div>
        {assignedMicroId === null ? <ChevronRight size={14} className="text-blue-300 ml-auto shrink-0" /> : <Lock size={13} className="text-blue-300 ml-auto shrink-0" />}
      </div>
    </div>
  );
  const sidebarUser = (
    <div className="mt-auto p-3">
      <div className="flex items-center gap-2.5 rounded-xl bg-blue-900 p-2.5">
        <button onClick={() => go("perfil")} className="flex items-center gap-2.5 min-w-0 flex-1 text-left">
          <Avatar name={profiles[role].nombre} color={roleCfg.color} size={34} />
          <div className="min-w-0 flex-1"><div className="text-white text-sm font-medium truncate">{profiles[role].nombre}</div><div className="text-blue-200 text-xs truncate">{roleCfg.short}</div></div>
        </button>
        <button onClick={async () => { await logout(); setAuthed(false); setMyMicroId(null); }} title="Cerrar sesión" className="text-blue-200 hover:text-white shrink-0"><LogOut size={16} /></button>
      </div>
    </div>
  );

  const screenTitle = { dashboard: "Inicio", microcentros: "Microcentros", grupo: `Grupo · ${myMicroName === "Todos los microcentros" ? "Caminante del Saber" : myMicroName}`, foro: "Foro", asistencia: "Asistencia", actas: "Actas", biblioteca: "Biblioteca digital", calendario: "Calendario", admin: "Panel administrativo", perfil: "Mi perfil" }[screen];

  const StatCard = ({ icon: Ic, label, value, delta, color }) => (
    <Card className="p-4">
      <div className="flex items-center justify-between">
        <div className="h-9 w-9 rounded-xl flex items-center justify-center" style={{ backgroundColor: rgba(color, 0.14) }}><Ic size={18} color={color} /></div>
        {delta && <span className="text-xs font-medium text-emerald-600 flex items-center gap-0.5"><TrendingUp size={12} />{delta}</span>}
      </div>
      <div className="mt-2.5 text-2xl font-bold">{value}</div><div className={cx("text-xs", t.muted)}>{label}</div>
    </Card>
  );

  /* ================= DASHBOARD ================= */
  const renderDashboard = () => {
    const reuniones = assignedMicroId ? meetings.filter((m) => m.micro === myMicroName) : meetings;
    const bdayPool = assignedMicroId ? (membersByMicro[assignedMicroId] || []) : Object.values(membersByMicro).flat();
    const upcomingBdays = bdayPool
      .filter((m) => m.nac)
      .map((m) => ({ ...m, dias: daysUntilNac(m.nac) }))
      .sort((a, b) => a.dias - b.dias)
      .slice(0, 3);
    return (
      <div className="space-y-5">
        <div><h1 className="text-xl sm:text-2xl font-bold">Hola, {firstName(profiles[role].nombre)} 👋</h1><p className={cx("text-sm", t.muted)}>{assignedMicroId ? `Microcentro ${myMicroName}` : "Resumen de todos los microcentros"}.</p></div>
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
          {role === "admin_general" ? (<>
            <StatCard icon={Building2} label="Microcentros" value="5" color="#1d4ed8" />
            <StatCard icon={Users} label="Usuarios activos" value="32" delta="+3" color="#059669" />
            <StatCard icon={ClipboardCheck} label="Reuniones (mes)" value="7" color="#0891b2" />
            <StatCard icon={BarChart3} label="Asistencia prom." value="88%" delta="+4%" color="#7c3aed" />
          </>) : (<>
            <StatCard icon={Users} label="Integrantes" value="7" color="#1d4ed8" />
            <StatCard icon={ClipboardCheck} label="Reuniones" value="5" color="#059669" />
            <StatCard icon={FileText} label="Documentos" value="18" delta="+3" color="#0891b2" />
            <StatCard icon={BarChart3} label="Mi asistencia" value="92%" color="#7c3aed" />
          </>)}
        </div>
        <div className="grid lg:grid-cols-3 gap-5">
          <div className="lg:col-span-2 space-y-5">
            <Section title="Próximas reuniones" action={isAdmin && <Btn icon={Plus} variant="ghost" onClick={() => go("calendario")}>Agendar</Btn>}>
              <div className="space-y-2.5">
                {reuniones.map((r, i) => (
                  <div key={i} className={cx("flex items-center gap-3 rounded-xl border p-3", t.border, t.alt)}>
                    <div className="flex flex-col items-center justify-center rounded-lg bg-blue-600 text-white px-2.5 py-1.5 shrink-0"><span className="text-[10px] uppercase leading-none">Jul</span><span className="text-lg font-bold leading-none">{r.fecha.split(" ")[0]}</span></div>
                    <div className="min-w-0"><div className="font-medium text-sm truncate">{r.titulo}</div>
                      <div className={cx("text-xs flex flex-wrap gap-x-3", t.muted)}><span className="flex items-center gap-1"><Clock size={12} />{r.hora}</span><span className="flex items-center gap-1"><MapPin size={12} />{r.lugar}</span></div>
                    </div>
                    <ChevronRight size={16} className={cx("ml-auto shrink-0", t.muted)} />
                  </div>
                ))}
              </div>
            </Section>
            <Section title="Actividad reciente">
              <div className="space-y-3">
                {[{ a: "Antonio Godoy", c: "#059669", txt: "subió un documento a Biblioteca › Actas", t: "hace 40 min" }, { a: "Daniela Lobos", c: "#db2777", txt: "abrió un tema en el Foro común", t: "hace 1 h" }, { a: "Pamela Vega", c: "#0891b2", txt: `publicó en el grupo ${myMicroName === "Todos los microcentros" ? "Caminante del Saber" : myMicroName}`, t: "hace 2 h" }].map((x, i) => (
                  <div key={i} className="flex items-start gap-3"><Avatar name={x.a} color={x.c} size={32} /><div className="text-sm"><span className="font-medium">{x.a}</span> <span className={t.muted}>{x.txt}</span><div className={cx("text-xs", t.muted)}>{x.t}</div></div></div>
                ))}
              </div>
            </Section>
          </div>
          <div className="space-y-5">
            <Section title="Noticias" action={role === "admin_general" && <Btn icon={Plus} variant="ghost">Nueva</Btn>}>
              <div className="space-y-2.5">{NOTICIAS.map((n, i) => (
                <div key={i} className="flex items-start gap-2">{n.pinned ? <Pin size={14} className="text-amber-500 mt-0.5 shrink-0" /> : <Megaphone size={14} className={cx("mt-0.5 shrink-0", t.muted)} />}
                  <div><div className="text-sm font-medium leading-snug">{n.titulo}</div><div className={cx("text-xs", t.muted)}>{n.fecha}</div></div></div>
              ))}</div>
            </Section>
            <Section title="Últimas actas">
              <div className="space-y-2">{actas.slice(0, 2).map((a, i) => (
                <div key={i} className="flex items-center gap-2"><FileText size={15} className="text-blue-600 shrink-0" /><span className="text-sm truncate flex-1">{a.titulo}</span><Chip color="#059669">v{a.v}</Chip></div>
              ))}</div>
            </Section>
            <Section title="Cumpleaños">
              {upcomingBdays.length === 0 ? (
                <p className={cx("text-sm", t.muted)}>No hay cumpleaños registrados.</p>
              ) : (
                <div className="space-y-2.5">
                  {upcomingBdays.map((m, i) => {
                    const b = parseNac(m.nac);
                    const cuando = m.dias === 0 ? "hoy 🎂" : m.dias === 1 ? "mañana 🎂" : `el ${b.d} ${MONTHS_ES[b.m - 1]}`;
                    return (
                      <div key={i} className="flex items-center gap-3">
                        <div className="h-9 w-9 rounded-xl flex items-center justify-center shrink-0" style={{ backgroundColor: rgba("#db2777", 0.14) }}><Cake size={16} className="text-pink-500" /></div>
                        <div className="text-sm min-w-0"><span className="font-medium">{firstName(m.name)}</span> cumple años {cuando}
                          {assignedMicroId === null && <div className={cx("text-xs truncate", t.muted)}>{m.rol}</div>}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </Section>
          </div>
        </div>
      </div>
    );
  };

  /* ================= MICROCENTROS (scoped) ================= */
  const renderMicrocentros = () => {
    const visibleMicros = assignedMicroId ? microcentros.filter((m) => m.id === assignedMicroId) : microcentros;
    if (!micro) {
      return (
        <div className="space-y-4">
          <div className="flex items-center justify-between">
            <div><h1 className="text-xl sm:text-2xl font-bold">{assignedMicroId ? "Mi microcentro" : "Microcentros"}</h1>
              <p className={cx("text-sm", t.muted)}>{assignedMicroId ? "Solo tienes acceso a tu microcentro asignado" : `${microcentros.length} microcentros en la comuna de Ovalle`}</p></div>
            {role === "admin_general" && <Btn icon={Plus}>Crear microcentro</Btn>}
          </div>
          <div className="grid sm:grid-cols-2 gap-4">
            {visibleMicros.map((m) => (
              <Card key={m.id} className="overflow-hidden">
                <div className="h-2" style={{ backgroundColor: m.color }} />
                <div className="p-4">
                  <div className="flex items-start gap-3">
                    <div className="h-11 w-11 rounded-xl flex items-center justify-center shrink-0" style={{ backgroundColor: rgba(m.color, 0.14) }}><Building2 size={20} color={m.color} /></div>
                    <div className="min-w-0"><div className="font-semibold truncate">Microcentro {m.nombre}</div><div className={cx("text-xs", t.muted)}>Presidente: {m.presidente}</div></div>
                  </div>
                  <div className="mt-3 flex flex-wrap gap-2"><Chip color={m.color}>{m.integrantes} integrantes</Chip><Chip color={m.color}>{m.establecimientos} establecimientos</Chip><Chip color={m.color}>{m.comuna}</Chip></div>
                  <button onClick={() => { setMicro(m); setMicroTab("info"); }} className="mt-3 w-full text-sm font-medium text-blue-600 hover:underline flex items-center justify-center gap-1">Ver microcentro <ChevronRight size={15} /></button>
                </div>
              </Card>
            ))}
          </div>
          {assignedMicroId && <p className={cx("text-xs flex items-center gap-1.5", t.muted)}><Lock size={13} />Los demás microcentros de la comuna no son visibles para tu rol.</p>}
        </div>
      );
    }
    const cm = microcentros.find((x) => x.id === micro.id) || micro;
    const tabs = [{ id: "info", label: "Información" }, { id: "integrantes", label: "Integrantes" }, { id: "establecimientos", label: "Establecimientos" }, { id: "documentos", label: "Documentos" }, { id: "fotos", label: "Fotos" }, { id: "calendario", label: "Calendario" }];
    return (
      <div className="space-y-4">
        <button onClick={() => setMicro(null)} className="flex items-center gap-1 text-sm text-blue-600 font-medium"><ChevronLeft size={16} />Volver</button>
        <Card className="overflow-hidden"><div className="h-2" style={{ backgroundColor: cm.color }} />
          <div className="p-4 flex items-center gap-3">
            <div className="h-12 w-12 rounded-xl flex items-center justify-center" style={{ backgroundColor: rgba(cm.color, 0.14) }}><Building2 size={22} color={cm.color} /></div>
            <div><h1 className="text-lg font-bold">Microcentro {cm.nombre}</h1><div className={cx("text-xs", t.muted)}>Presidente: {cm.presidente} · {cm.comuna}</div></div>
          </div>
        </Card>
        <div className="flex gap-1 overflow-x-auto pb-1">{tabs.map((tb) => (
          <button key={tb.id} onClick={() => setMicroTab(tb.id)} className={cx("whitespace-nowrap rounded-lg px-3 py-1.5 text-sm font-medium transition", microTab === tb.id ? "bg-blue-600 text-white" : cx(t.surface, "border", t.border, t.hover))}>{tb.label}</button>
        ))}</div>
        {microTab === "info" && (
          <Card className="p-5 space-y-3"><p className={cx("text-sm", t.muted)}>Agrupación de escuelas rurales del sector {cm.nombre}. Reúne a docentes y directivos para el trabajo técnico-pedagógico colaborativo.</p>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">{[["Integrantes", cm.integrantes], ["Establecimientos", cm.establecimientos], ["Comuna", cm.comuna], ["Reuniones", "5"]].map(([k, v], i) => (
              <div key={i} className={cx("rounded-xl border p-3", t.border, t.alt)}><div className="text-lg font-bold">{v}</div><div className={cx("text-xs", t.muted)}>{k}</div></div>))}</div>
          </Card>
        )}
        {microTab === "integrantes" && (<Card className="p-3">{(membersByMicro[cm.id] || []).map((m, i) => (
          <div key={i} className={cx("flex items-center gap-3 p-2 rounded-xl", t.hover)}><Avatar name={m.name} color={m.color} size={36} />
            <div className="flex-1 min-w-0"><div className="text-sm font-medium truncate">{m.name}</div><div className={cx("text-xs truncate", t.muted)}>{m.rol}</div></div>
            {isAdmin && <button onClick={() => openManage(cm.id, i)} className={cx("text-xs px-2 py-1 rounded-lg", t.hover, t.muted)}>Gestionar</button>}</div>))}
        </Card>)}
        {microTab === "establecimientos" && (<Card className="p-4 space-y-2">{(ESTAB_BY_MICRO[cm.id] || []).map((e, i) => (
          <div key={i} className={cx("flex items-center gap-3 rounded-xl border p-3", t.border)}><School size={18} className="text-emerald-600 shrink-0" /><span className="text-sm flex-1">{e}</span><MapPin size={15} className={t.muted} /></div>))}
          <p className={cx("text-xs pt-1", t.muted)}>📍 Mapa interactivo disponible con la ubicación de cada establecimiento.</p></Card>)}
        {microTab === "documentos" && (<Card className="p-3">{FILES.filter((f) => f.micro === cm.id || f.micro === 0).slice(0, 4).map((f, i) => (
          <div key={i} className={cx("flex items-center gap-3 p-2.5 rounded-xl", t.hover)}><FileText size={18} style={{ color: microColor(f.micro) }} />
            <div className="flex-1 min-w-0"><div className="text-sm truncate">{f.name}</div><div className={cx("text-xs", t.muted)}>{f.cat} · {f.size}</div></div><Download size={16} className={t.muted} /></div>))}
        </Card>)}
        {microTab === "fotos" && (<div className="grid grid-cols-3 sm:grid-cols-4 gap-2">{["#1d4ed8", "#059669", "#0891b2", "#7c3aed", "#db2777", "#ea580c", "#0d9488", "#6366f1"].map((c, i) => (
          <div key={i} className="aspect-square rounded-xl flex items-center justify-center" style={{ backgroundColor: rgba(c, 0.15) }}><ImageIcon size={22} color={c} /></div>))}</div>)}
        {microTab === "calendario" && (<Card className="p-4 space-y-2">{Object.entries(EVENT_DAYS).slice(0, 3).map(([d, e], i) => (
          <div key={i} className={cx("flex items-center gap-3 rounded-xl border p-3", t.border)}><div className="text-sm font-bold w-8" style={{ color: e.color }}>{d} Jul</div><span className="text-sm flex-1">{e.label}</span></div>))}</Card>)}

        {manageMember && manageMember.microId === cm.id && (() => {
          const person = (membersByMicro[manageMember.microId] || [])[manageMember.index];
          if (!person) return null;
          const { roleLabel, estab } = splitRol(person.rol);
          return (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ backgroundColor: "rgba(0,0,0,0.55)" }}>
              <div className={cx("w-full max-w-sm rounded-2xl p-5", t.surface)}>
                <div className="flex items-center justify-between mb-4">
                  <h3 className="font-bold text-base">Gestionar integrante</h3>
                  <button onClick={closeManage} className={t.muted}><X size={18} /></button>
                </div>
                <div className="flex items-center gap-3 mb-4">
                  <Avatar name={person.name} color={person.color} size={40} />
                  <div className="min-w-0"><div className="font-semibold text-sm truncate">{person.name}</div><div className={cx("text-xs truncate", t.muted)}>{estab}</div></div>
                </div>
                <label className="block mb-4">
                  <span className={cx("text-xs font-medium", t.muted)}>Rol en el microcentro</span>
                  <select value={roleLabel} onChange={(e) => setManagedRole(e.target.value)} className={cx("mt-1 w-full rounded-xl border px-3 py-2 text-sm", t.input)}>
                    {ROLE_LABELS.map((r) => <option key={r} value={r}>{r}</option>)}
                  </select>
                </label>
                <div className="flex flex-col gap-2">
                  <button onClick={removeManagedMember} className="w-full rounded-xl border border-red-200 text-red-600 text-sm font-medium py-2 hover:bg-red-50 transition">Quitar del microcentro</button>
                  <Btn variant="ghost" onClick={closeManage}>Cerrar</Btn>
                </div>
              </div>
            </div>
          );
        })()}
      </div>
    );
  };

  /* ================= GRUPO ================= */
  const renderGrupo = () => {
    const addPost = () => { if (!composer.trim()) return; setPosts([{ id: Date.now(), autor: roleCfg.user, rol: roleCfg.short, tiempo: "ahora", texto: composer, likes: 0, comentarios: 0, liked: false, color: roleCfg.color }, ...posts]); setComposer(""); showToast(role === "docente" ? "Publicación enviada a revisión" : "Publicado"); };
    const toggleLike = (id) => setPosts(posts.map((p) => p.id === id ? { ...p, liked: !p.liked, likes: p.likes + (p.liked ? -1 : 1) } : p));

    // Solo Gloria González, en su rol de Administradora General, puede borrar
    // publicaciones de cualquier persona. El resto solo gestiona lo propio.
    const isGloriaAdminGeneral = role === "admin_general" && profiles.admin_general.nombre === "Gloria González";

    const startEditPost = (p) => { setEditingPostId(p.id); setEditPostText(p.texto); };
    const cancelEditPost = () => { setEditingPostId(null); setEditPostText(""); };
    const saveEditPost = (id) => {
      if (!editPostText.trim()) return;
      setPosts(posts.map((p) => (p.id === id ? { ...p, texto: editPostText.trim(), editado: true } : p)));
      setEditingPostId(null); setEditPostText("");
      showToast("Publicación actualizada");
    };
    const deletePost = (id) => {
      setPosts(posts.filter((p) => p.id !== id));
      if (editingPostId === id) cancelEditPost();
      showToast("Publicación eliminada");
    };

    return (
      <div className="max-w-2xl mx-auto space-y-4">
        {role === "admin_micro" && (<Card className="p-3.5 flex items-center gap-3"><div className="h-9 w-9 rounded-xl bg-amber-100 flex items-center justify-center shrink-0"><ShieldCheck size={18} className="text-amber-600" /></div><div className="text-sm flex-1"><span className="font-medium">2 publicaciones</span> pendientes de aprobación</div><Btn variant="emerald" onClick={() => showToast("Publicaciones aprobadas")}>Revisar</Btn></Card>)}
        <Card className="p-4">
          <div className="flex gap-3"><Avatar name={roleCfg.user} color={roleCfg.color} /><textarea value={composer} onChange={(e) => setComposer(e.target.value)} rows={2} placeholder="Comparte algo con tu microcentro…" className={cx("flex-1 resize-none rounded-xl border px-3 py-2 text-sm outline-none focus:border-blue-500", t.input)} /></div>
          <div className="flex items-center justify-between mt-3"><div className={cx("flex items-center gap-1", t.muted)}>{[Paperclip, ImageIcon].map((Ic, i) => <button key={i} className={cx("p-1.5 rounded-lg", t.hover)}><Ic size={18} /></button>)}</div><Btn icon={Send} onClick={addPost}>Publicar</Btn></div>
          {role === "docente" && <p className={cx("text-xs mt-2", t.muted)}>Tus publicaciones serán revisadas por un administrador antes de aparecer.</p>}
        </Card>
        {posts.map((p) => {
          const isOwn = p.autor === roleCfg.user;
          const canEdit = isOwn;
          const canDelete = isOwn || isGloriaAdminGeneral;
          const isEditingThis = editingPostId === p.id;
          return (
            <Card key={p.id} className="p-4">
              <div className="flex items-center gap-3">
                <Avatar name={p.autor} color={p.color} />
                <div className="min-w-0 flex-1"><div className="text-sm font-semibold">{p.autor}</div><div className={cx("text-xs", t.muted)}>{p.rol} · {p.tiempo}{p.editado ? " · editado" : ""}</div></div>
                {!isEditingThis && (canEdit || canDelete) && (
                  <div className="flex items-center gap-1 shrink-0">
                    {canEdit && (
                      <button onClick={() => startEditPost(p)} title="Editar publicación" className={cx("p-1.5 rounded-lg", t.hover, t.muted)}>
                        <PenLine size={15} />
                      </button>
                    )}
                    {canDelete && (
                      <button onClick={() => deletePost(p.id)} title={isOwn ? "Eliminar publicación" : "Eliminar publicación (administradora)"} className={cx("p-1.5 rounded-lg hover:bg-red-50 hover:text-red-600 transition", t.muted)}>
                        <Trash2 size={15} />
                      </button>
                    )}
                  </div>
                )}
              </div>

              {isEditingThis ? (
                <div className="mt-3 space-y-2">
                  <textarea value={editPostText} onChange={(e) => setEditPostText(e.target.value)} rows={3} autoFocus className={cx("w-full resize-none rounded-xl border px-3 py-2 text-sm outline-none focus:border-blue-500", t.input)} />
                  <div className="flex gap-2"><Btn icon={Save} onClick={() => saveEditPost(p.id)}>Guardar</Btn><Btn variant="ghost" onClick={cancelEditPost}>Cancelar</Btn></div>
                </div>
              ) : (
                <p className="text-sm mt-3 leading-relaxed">{p.texto}</p>
              )}

              {p.attach && (<div className={cx("mt-3 flex items-center gap-2.5 rounded-xl border p-2.5", t.border, t.alt)}><div className="h-9 w-9 rounded-lg bg-red-100 flex items-center justify-center"><FileText size={18} className="text-red-600" /></div><div className="text-sm flex-1 truncate">{p.attach}</div><Download size={16} className={t.muted} /></div>)}
              <div className={cx("flex items-center gap-5 mt-3 pt-3 border-t text-sm", t.border, t.muted)}><button onClick={() => toggleLike(p.id)} className={cx("flex items-center gap-1.5", p.liked && "text-red-500 font-medium")}><Heart size={16} fill={p.liked ? "currentColor" : "none"} /> {p.likes}</button><span className="flex items-center gap-1.5"><MessageCircle size={16} /> {p.comentarios}</span></div>
            </Card>
          );
        })}
      </div>
    );
  };

  /* ================= FORO (scoped) ================= */
  const renderForo = () => {
    if (thread) {
      const th = THREADS.find((x) => x.id === thread); const cat = CATEGORIAS.find((c) => c.name === th.cat);
      return (
        <div className="max-w-2xl mx-auto space-y-4">
          <button onClick={() => setThread(null)} className="flex items-center gap-1 text-sm text-blue-600 font-medium"><ChevronLeft size={16} />Volver al foro</button>
          <Card className="p-5"><div className="flex gap-2 mb-2"><Chip color={cat.color}>{th.cat}</Chip><Chip color={microColor(th.micro)}>{microName(th.micro)}</Chip></div>
            <h1 className="text-lg font-bold">{th.titulo}</h1>
            <div className={cx("flex items-center gap-2 mt-2 text-xs", t.muted)}><Avatar name={th.autor} color="#7c3aed" size={22} /> {th.autor} · {th.tiempo}</div>
            <p className="text-sm mt-3 leading-relaxed">Estamos con cursos combinados de 1° a 4° y me interesa saber cómo organizan los grupos de lectura. ¿Trabajan por niveles o de forma integrada?</p></Card>
          {[{ a: "Eugenio Sierra", c: "#059669", txt: "Nosotros armamos parejas de tutoría: los de 3°/4° apoyan a los más pequeños. Ha funcionado muy bien." }, { a: "Miguel Zamora", c: "#ea580c", txt: "Yo uso estaciones de trabajo rotativas, así atiendo un grupo mientras los demás avanzan de forma autónoma." }].map((r, i) => (
            <Card key={i} className="p-4"><div className="flex items-center gap-2.5"><Avatar name={r.a} color={r.c} size={30} /><span className="text-sm font-semibold">{r.a}</span></div><p className="text-sm mt-2 leading-relaxed">{r.txt}</p></Card>))}
          <Card className="p-3 flex gap-2 items-center"><input placeholder="Escribe una respuesta…" className={cx("flex-1 rounded-xl border px-3 py-2 text-sm outline-none focus:border-blue-500", t.input)} /><Btn icon={Send} onClick={() => showToast("Respuesta publicada")}>Responder</Btn></Card>
        </div>
      );
    }
    const scopes = [{ id: "comun", label: "Común", color: "#64748b" }];
    if (assignedMicroId) scopes.push({ id: String(assignedMicroId), label: myMicroName, color: microColor(assignedMicroId) });
    else MICROCENTROS.forEach((m) => scopes.push({ id: String(m.id), label: m.nombre, color: m.color }));
    const visibleThreads = THREADS.filter((th) => (foroScope === "comun" ? th.micro === 0 : th.micro === Number(foroScope)));
    return (
      <div className="space-y-4">
        <div className="flex items-center justify-between flex-wrap gap-2"><div><h1 className="text-xl sm:text-2xl font-bold">Foro</h1><p className={cx("text-sm", t.muted)}>Espacio de intercambio pedagógico</p></div><Btn icon={Plus} onClick={() => showToast("Nuevo tema (demo)")}>Nuevo tema</Btn></div>
        <div className="flex gap-1.5 overflow-x-auto pb-1">{scopes.map((s) => (
          <button key={s.id} onClick={() => setForoScope(s.id)} className={cx("whitespace-nowrap rounded-lg px-3 py-1.5 text-sm font-medium border transition")} style={foroScope === s.id ? { backgroundColor: s.color, color: "#fff", borderColor: s.color } : { borderColor: dark ? "#334155" : "#e2e8f0", color: dark ? "#cbd5e1" : "#475569" }}>{s.id === "comun" ? "🌐 " : ""}{s.label}</button>
        ))}</div>
        <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5">{CATEGORIAS.map((c) => (
          <button key={c.name} className={cx("rounded-xl border p-3 text-left transition", t.border, t.surface, t.hover)}><div className="h-8 w-8 rounded-lg flex items-center justify-center mb-2" style={{ backgroundColor: rgba(c.color, 0.14) }}><MessageSquare size={16} color={c.color} /></div><div className="text-sm font-medium leading-snug">{c.name}</div><div className={cx("text-xs", t.muted)}>{c.count} temas</div></button>))}</div>
        <Section title={foroScope === "comun" ? "Temas del foro común" : `Temas de ${scopes.find((s) => s.id === foroScope)?.label}`}>
          <div className="space-y-1">
            {visibleThreads.length === 0 && <p className={cx("text-sm py-2", t.muted)}>No hay temas en este foro todavía.</p>}
            {visibleThreads.map((th) => { const cat = CATEGORIAS.find((c) => c.name === th.cat); return (
              <button key={th.id} onClick={() => setThread(th.id)} className={cx("w-full flex items-center gap-3 p-2.5 rounded-xl text-left", t.hover)}>
                <Avatar name={th.autor} color={cat.color} size={36} /><div className="flex-1 min-w-0"><div className="text-sm font-medium truncate">{th.titulo}</div><div className={cx("text-xs flex items-center gap-2 flex-wrap", t.muted)}><Chip color={cat.color}>{th.cat}</Chip> {th.tiempo}</div></div>
                <div className={cx("text-xs flex items-center gap-1 shrink-0", t.muted)}><MessageCircle size={13} />{th.respuestas}</div>
              </button>); })}
          </div>
        </Section>
      </div>
    );
  };

  /* ================= ASISTENCIA (functional) ================= */
  const renderAsistencia = () => {
    const canManageAtt = isAdmin || role === "presidente";
    const resolvedMicroName = myMicroName === "Todos los microcentros" ? "Caminante del Saber" : myMicroName;
    const scopedMeetings = meetings.filter((m) => m.micro === resolvedMicroName);
    const meeting = scopedMeetings.find((m) => m.id === meetingId) || scopedMeetings[0];
    const openNewMeeting = () => { setMeetDraft({ titulo: "", fecha: "", hora: "", lugar: "" }); setNewMeetingOpen(true); };
    const createMeeting = () => {
      const id = "m" + Date.now();
      setMeetings((prev) => [...prev, {
        id,
        titulo: meetDraft.titulo.trim() || "Reunión sin título",
        fecha: meetDraft.fecha.trim() || "—",
        hora: meetDraft.hora.trim() || "—",
        lugar: meetDraft.lugar.trim() || "—",
        micro: resolvedMicroName,
      }]);
      setMeetingId(id);
      setAtt({}); setHoras({}); setObs({}); setFirmado(null);
      setNewMeetingOpen(false);
      showToast("Nueva asistencia creada");
    };
    const newMeetingModal = newMeetingOpen && (
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ backgroundColor: "rgba(0,0,0,0.55)" }}>
        <div className={cx("w-full max-w-sm rounded-2xl p-5", t.surface)}>
          <div className="flex items-center justify-between mb-1"><h3 className="font-bold text-base">Nueva asistencia</h3><button onClick={() => setNewMeetingOpen(false)} className={t.muted}><X size={18} /></button></div>
          <p className={cx("text-xs mb-4", t.muted)}>Microcentro {resolvedMicroName}</p>
          <div className="space-y-3">
            <Field label="Título de la reunión" value={meetDraft.titulo} onChange={(e) => setMeetDraft({ ...meetDraft, titulo: e.target.value })} placeholder="Ej: Reunión mensual de coordinación" />
            <div className="grid grid-cols-2 gap-2">
              <Field label="Fecha" value={meetDraft.fecha} onChange={(e) => setMeetDraft({ ...meetDraft, fecha: e.target.value })} placeholder="Ej: 15 Jul 2025" />
              <Field label="Hora" value={meetDraft.hora} onChange={(e) => setMeetDraft({ ...meetDraft, hora: e.target.value })} placeholder="Ej: 15:00" />
            </div>
            <Field label="Lugar" value={meetDraft.lugar} onChange={(e) => setMeetDraft({ ...meetDraft, lugar: e.target.value })} placeholder="Ej: Esc. La Quiroga" />
          </div>
          <div className="flex gap-2 mt-4"><Btn icon={Save} onClick={createMeeting}>Crear asistencia</Btn><Btn variant="ghost" onClick={() => setNewMeetingOpen(false)}>Cancelar</Btn></div>
        </div>
      </div>
    );

    if (!meeting) {
      return (
        <>
          <div className="space-y-4">
            <div className="flex items-center justify-between flex-wrap gap-2">
              <div><h1 className="text-xl sm:text-2xl font-bold">Registro de asistencia</h1><p className={cx("text-sm", t.muted)}>Microcentro {resolvedMicroName}</p></div>
              {canManageAtt && <Btn icon={Plus} onClick={openNewMeeting}>Nueva asistencia</Btn>}
            </div>
            <Card className="p-8 text-center space-y-3">
              <div className="mx-auto h-12 w-12 rounded-2xl flex items-center justify-center" style={{ backgroundColor: rgba("#0891b2", 0.14) }}><ClipboardCheck size={24} color="#0891b2" /></div>
              <p className={cx("text-sm", t.muted)}>Aún no hay reuniones registradas para este microcentro.</p>
              {canManageAtt
                ? <div className="flex justify-center"><Btn icon={Plus} onClick={openNewMeeting}>Crear primera asistencia</Btn></div>
                : <p className={cx("text-xs", t.muted)}>Un presidente o administrador debe crear la primera reunión.</p>}
            </Card>
          </div>
          {newMeetingModal}
        </>
      );
    }

    const states = ["presente", "ausente", "justificado"];
    const cfg = { presente: { c: "#059669", l: "Presente", i: CheckCircle2 }, ausente: { c: "#e11d48", l: "Ausente", i: XCircle }, justificado: { c: "#d97706", l: "Justificado", i: Clock } };
    const total = MEMBERS.length;
    const counts = { presente: 0, ausente: 0, justificado: 0 };
    Object.values(att).forEach((s) => counts[s]++);
    const pct = Math.round((counts.presente / total) * 100);
    const mark = (i, s) => { setAtt({ ...att, [i]: s }); if (s === "presente" && !horas[i]) setHoras((h) => ({ ...h, [i]: "15:00" })); };
    const exportExcel = () => {
      try {
        const rows = MEMBERS.map((m, i) => ({ Nombre: m.name, Rol: m.rol, Estado: cfg[att[i]]?.l || "Sin registrar", "Hora de llegada": att[i] === "presente" ? horas[i] || "15:00" : "—", Observación: obs[i] || "" }));
        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Asistencia");
        XLSX.writeFile(wb, `Asistencia_${meeting.fecha.replace(/ /g, "_")}.xlsx`);
        showToast("Planilla Excel generada");
      } catch (e) { showToast("No se pudo generar el archivo"); }
    };
    return (
      <>
        <div className="space-y-4">
          <div className="flex items-center justify-between flex-wrap gap-2">
            <div><h1 className="text-xl sm:text-2xl font-bold">Registro de asistencia</h1><p className={cx("text-sm", t.muted)}>Microcentro {resolvedMicroName}</p></div>
            <div className="flex gap-2">
              {canManageAtt && <Btn icon={Plus} onClick={openNewMeeting}>Nueva asistencia</Btn>}
              <Btn icon={QrCode} variant="ghost" onClick={() => showToast("Mostrando código QR de asistencia")}>Registrar por QR</Btn>
            </div>
          </div>
          <Card className="p-3 flex items-center gap-2 flex-wrap">
            <ClipboardCheck size={16} className={t.muted} />
            <select value={meeting.id} onChange={(e) => setMeetingId(e.target.value)} className={cx("rounded-lg border px-2 py-1.5 text-sm flex-1 min-w-0", t.input)}>
              {scopedMeetings.map((m) => <option key={m.id} value={m.id}>{m.titulo} — {m.fecha}</option>)}
            </select>
            <span className={cx("text-xs", t.muted)}><MapPin size={12} className="inline" /> {meeting.lugar}</span>
          </Card>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <Card className="p-4"><div className="text-2xl font-bold text-emerald-600">{pct}%</div><div className={cx("text-xs", t.muted)}>Asistencia</div></Card>
            {states.map((s) => (<Card key={s} className="p-4"><div className="text-2xl font-bold" style={{ color: cfg[s].c }}>{counts[s]}</div><div className={cx("text-xs", t.muted)}>{cfg[s].l}</div></Card>))}
          </div>
          <div className="flex h-2.5 rounded-full overflow-hidden">
            <div style={{ width: `${(counts.presente / total) * 100}%`, backgroundColor: "#059669" }} />
            <div style={{ width: `${(counts.justificado / total) * 100}%`, backgroundColor: "#d97706" }} />
            <div style={{ width: `${(counts.ausente / total) * 100}%`, backgroundColor: "#e11d48" }} />
          </div>
          <Card className="p-3">
            {MEMBERS.map((m, i) => (
              <div key={i} className={cx("p-2 rounded-xl", t.hover)}>
                <div className="flex items-center gap-3 flex-wrap">
                  <Avatar name={m.name} color={m.color} size={36} />
                  <div className="flex-1 min-w-0"><div className="text-sm font-medium truncate">{m.name}</div><div className={cx("text-xs truncate", t.muted)}>{m.rol}</div></div>
                  <div className="flex gap-1">{states.map((s) => { const active = att[i] === s; const Ic = cfg[s].i; return (
                    <button key={s} onClick={() => mark(i, s)} className="flex items-center gap-1 rounded-lg px-2 py-1 text-xs font-medium border transition" style={active ? { backgroundColor: cfg[s].c, color: "#fff", borderColor: cfg[s].c } : { borderColor: dark ? "#334155" : "#e2e8f0", color: dark ? "#94a3b8" : "#64748b" }}><Ic size={13} /><span className="hidden sm:inline">{cfg[s].l}</span></button>); })}</div>
                </div>
                {att[i] === "presente" && (<div className="mt-2 flex items-center gap-2 sm:pl-12 text-xs"><Clock size={13} className={t.muted} /><span className={t.muted}>Hora de llegada</span><input type="time" value={horas[i] || "15:00"} onChange={(e) => setHoras({ ...horas, [i]: e.target.value })} className={cx("rounded-lg border px-2 py-1 text-xs outline-none", t.input)} /></div>)}
                {(att[i] === "ausente" || att[i] === "justificado") && (<div className="mt-2 sm:pl-12"><input value={obs[i] || ""} onChange={(e) => setObs({ ...obs, [i]: e.target.value })} placeholder="Observación / motivo…" className={cx("w-full rounded-lg border px-2.5 py-1.5 text-xs outline-none focus:border-blue-500", t.input)} /></div>)}
              </div>
            ))}
          </Card>
          <Card className="p-4 flex items-center justify-between flex-wrap gap-2">
            <div className="text-sm"><div className="font-medium">Firma del presidente</div>{firmado ? <div className="text-xs flex items-center gap-1 text-emerald-600"><CheckCircle2 size={14} /> Firmado por {firmado.by} · {firmado.at}</div> : <div className={cx("text-xs", t.muted)}>Registro sin firmar</div>}</div>
            {!firmado ? <Btn icon={PenLine} variant="emerald" onClick={() => setFirmado({ by: roleCfg.user, at: "hoy 15:20" })}>Firmar registro</Btn> : <Btn icon={PenLine} variant="ghost" onClick={() => setFirmado(null)}>Quitar firma</Btn>}
          </Card>
          <div className="flex gap-2 flex-wrap">
            <Btn icon={Save} onClick={() => showToast("Registro de asistencia guardado")}>Guardar</Btn>
            <Btn icon={FileSpreadsheet} variant="emerald" onClick={exportExcel}>Exportar Excel</Btn>
            <Btn icon={Printer} variant="ghost" onClick={() => setPrintOpen(true)}>Vista para imprimir</Btn>
          </div>
          <Section title="Historial de asistencia">
            {counts.presente + counts.ausente + counts.justificado === 0 ? (
              <p className={cx("text-sm py-1", t.muted)}>Aún no hay registros de asistencia guardados para este microcentro.</p>
            ) : (
              <div className="flex items-end gap-2 h-24">
                <div className="flex-1 flex flex-col items-center gap-1"><div className="w-full rounded-t-md" style={{ height: `${pct}%`, backgroundColor: "#059669" }} /><span className={cx("text-[10px]", t.muted)}>Hoy</span></div>
              </div>
            )}
          </Section>
        </div>

        {printOpen && (
          <div className="fixed inset-0 z-50 flex items-start justify-center p-4 overflow-auto" style={{ backgroundColor: "rgba(0,0,0,0.55)" }}>
            <div className="w-full max-w-2xl rounded-2xl bg-white p-6 my-6" style={{ color: "#0f172a" }}>
              <div className="flex items-start justify-between mb-4">
                <div><div className="font-bold text-lg">Registro de asistencia</div><div className="text-sm" style={{ color: "#475569" }}>Microcentro {resolvedMicroName}</div><div className="text-sm" style={{ color: "#475569" }}>{meeting.titulo} · {meeting.fecha} · {meeting.lugar}</div></div>
                <button onClick={() => setPrintOpen(false)}><X size={20} /></button>
              </div>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "13px" }}>
                <thead><tr style={{ borderBottom: "2px solid #e2e8f0", textAlign: "left" }}><th style={{ padding: "6px" }}>Nombre</th><th>Estado</th><th>Hora</th><th>Observación</th></tr></thead>
                <tbody>{MEMBERS.map((m, i) => (<tr key={i} style={{ borderBottom: "1px solid #f1f5f9" }}><td style={{ padding: "6px" }}>{m.name}</td><td>{cfg[att[i]]?.l || "Sin registrar"}</td><td>{att[i] === "presente" ? horas[i] || "15:00" : "—"}</td><td>{obs[i] || "—"}</td></tr>))}</tbody>
              </table>
              <div className="mt-4 text-sm" style={{ color: "#475569" }}>Asistencia: {pct}% · Presentes {counts.presente} · Ausentes {counts.ausente} · Justificados {counts.justificado}</div>
              <div className="mt-10" style={{ width: "220px" }}><div style={{ borderTop: "1px solid #0f172a", paddingTop: "4px", fontSize: "12px", color: "#475569" }}>{firmado ? firmado.by : "Firma presidente"}</div></div>
              <div className="flex gap-2 justify-end mt-6"><button onClick={() => window.print()} className="rounded-xl bg-blue-600 text-white px-4 py-2 text-sm font-medium">Imprimir / Guardar PDF</button><button onClick={() => setPrintOpen(false)} className="rounded-xl border px-4 py-2 text-sm" style={{ borderColor: "#e2e8f0" }}>Cerrar</button></div>
            </div>
          </div>
        )}

        {newMeetingModal}
      </>
    );
  };

  /* ================= ACTAS (functional) ================= */
  const renderActas = () => {
    const openNew = () => { setDraft({ ...blankDraft }); setEditing(true); };
    const openAI = () => { setDraft({ id: null, titulo: "Acta N°6 — Planificación 2° semestre", fecha: "", lugar: "Esc. La Quiroga", participants: MEMBERS.slice(0, 5).map((m) => m.name), temas: ["Balance del primer semestre", "Metas y focos para el segundo semestre", "Calendario de reuniones"], acuerdos: [{ desc: "Definir focos por nivel", resp: "Pamela Vega", fecha: "15 Jul", done: false }], v: 0 }); setEditing(true); showToast("Borrador generado con IA — revísalo"); };
    const openEdit = (a) => { setDraft({ id: a.id, titulo: a.titulo, fecha: a.fecha, lugar: a.lugar, participants: MEMBERS.slice(0, a.participantes).map((m) => m.name), temas: [...a.temas], acuerdos: a.acuerdos.map((x) => ({ ...x })), v: a.v }); setEditing(true); };
    const addTema = () => { if (!temaInput.trim()) return; setDraft((d) => ({ ...d, temas: [...d.temas, temaInput.trim()] })); setTemaInput(""); };
    const addAcuerdo = () => { if (!aDesc.trim()) return; setDraft((d) => ({ ...d, acuerdos: [...d.acuerdos, { desc: aDesc.trim(), resp: aResp, fecha: aFecha || "—", done: false }] })); setADesc(""); setAFecha(""); };
    const saveActa = () => {
      const participantes = draft.participants.length;
      if (draft.id) { setActas(actas.map((a) => a.id === draft.id ? { ...a, titulo: draft.titulo, fecha: draft.fecha, lugar: draft.lugar, participantes, temas: draft.temas, acuerdos: draft.acuerdos, v: a.v + 1 } : a)); showToast(`Nueva versión (v${draft.v + 1}) guardada`); }
      else { setActas([{ id: "a" + Date.now(), titulo: draft.titulo || "Acta sin título", fecha: draft.fecha || "—", lugar: draft.lugar || "—", autor: roleCfg.user, participantes, temas: draft.temas, acuerdos: draft.acuerdos, v: 1 }, ...actas]); showToast("Acta creada"); }
      setEditing(false);
    };
    const toggleAcuerdo = (actaId, idx) => setActas(actas.map((a) => a.id === actaId ? { ...a, acuerdos: a.acuerdos.map((x, i) => i === idx ? { ...x, done: !x.done } : x) } : a));

    if (editing) {
      return (
        <div className="max-w-2xl mx-auto space-y-4">
          <button onClick={() => setEditing(false)} className="flex items-center gap-1 text-sm text-blue-600 font-medium"><ChevronLeft size={16} />Volver</button>
          <Card className="p-5 space-y-4">
            <div className="flex items-center gap-2"><h2 className="text-lg font-bold">{draft.id ? "Editar acta" : "Nueva acta"}</h2>{draft.id && <Chip color="#d97706">creará v{draft.v + 1}</Chip>}</div>
            <div className="grid sm:grid-cols-2 gap-3">
              <div className="sm:col-span-2"><Field label="Título del acta" value={draft.titulo} onChange={(e) => setDraft({ ...draft, titulo: e.target.value })} placeholder="Ej: Acta N°6 — …" /></div>
              <Field label="Fecha" value={draft.fecha} onChange={(e) => setDraft({ ...draft, fecha: e.target.value })} placeholder="Ej: 15 Jul 2025" />
              <Field label="Lugar" value={draft.lugar} onChange={(e) => setDraft({ ...draft, lugar: e.target.value })} placeholder="Ej: Esc. La Quiroga" />
            </div>
            <div><span className={cx("text-xs font-medium", t.muted)}>Participantes ({draft.participants.length})</span>
              <div className="mt-1.5 flex flex-wrap gap-1.5">{MEMBERS.map((m) => { const on = draft.participants.includes(m.name); return (
                <button key={m.name} onClick={() => setDraft((d) => ({ ...d, participants: on ? d.participants.filter((x) => x !== m.name) : [...d.participants, m.name] }))} className="rounded-full px-2.5 py-1 text-xs border transition" style={on ? { backgroundColor: m.color, color: "#fff", borderColor: m.color } : { borderColor: dark ? "#334155" : "#e2e8f0", color: dark ? "#cbd5e1" : "#475569" }}>{firstName(m.name)}</button>); })}</div>
            </div>
            <div><span className={cx("text-xs font-medium", t.muted)}>Temas tratados</span>
              <div className="mt-1.5 space-y-1.5">{draft.temas.map((tm, i) => (<div key={i} className={cx("flex items-center gap-2 rounded-lg border p-2 text-sm", t.border)}><span className="flex-1">{tm}</span><button onClick={() => setDraft((d) => ({ ...d, temas: d.temas.filter((_, x) => x !== i) }))}><Trash2 size={14} className={t.muted} /></button></div>))}</div>
              <div className="flex gap-2 mt-2"><input value={temaInput} onChange={(e) => setTemaInput(e.target.value)} placeholder="Agregar tema…" className={cx("flex-1 rounded-xl border px-3 py-2 text-sm outline-none focus:border-blue-500", t.input)} /><Btn icon={Plus} variant="ghost" onClick={addTema}>Agregar</Btn></div>
            </div>
            <div><span className={cx("text-xs font-medium", t.muted)}>Acuerdos y compromisos</span>
              <div className="mt-1.5 space-y-1.5">{draft.acuerdos.map((ac, i) => (<div key={i} className={cx("flex items-center gap-2 rounded-lg border p-2", t.border)}><CheckCircle2 size={15} className="text-emerald-600 shrink-0" /><span className="text-sm flex-1">{ac.desc}</span><Chip color="#1d4ed8">{firstName(ac.resp)}</Chip><span className={cx("text-xs", t.muted)}>{ac.fecha}</span><button onClick={() => setDraft((d) => ({ ...d, acuerdos: d.acuerdos.filter((_, x) => x !== i) }))}><Trash2 size={14} className={t.muted} /></button></div>))}</div>
              <div className="grid sm:grid-cols-3 gap-2 mt-2">
                <input value={aDesc} onChange={(e) => setADesc(e.target.value)} placeholder="Acuerdo / compromiso" className={cx("rounded-xl border px-3 py-2 text-sm outline-none focus:border-blue-500", t.input)} />
                <select value={aResp} onChange={(e) => setAResp(e.target.value)} className={cx("rounded-xl border px-3 py-2 text-sm", t.input)}>{MEMBERS.map((m) => <option key={m.name}>{m.name}</option>)}</select>
                <input value={aFecha} onChange={(e) => setAFecha(e.target.value)} placeholder="Fecha límite" className={cx("rounded-xl border px-3 py-2 text-sm outline-none focus:border-blue-500", t.input)} />
              </div>
              <div className="mt-2"><Btn icon={Plus} variant="ghost" onClick={addAcuerdo}>Agregar acuerdo</Btn></div>
            </div>
            <div className="flex gap-2 pt-2 border-t" style={{ borderColor: dark ? "#1e293b" : "#e2e8f0" }}><div className="pt-2 flex gap-2"><Btn icon={Save} onClick={saveActa}>{draft.id ? "Guardar nueva versión" : "Crear acta"}</Btn><Btn variant="ghost" onClick={() => setEditing(false)}>Cancelar</Btn></div></div>
          </Card>
        </div>
      );
    }

    return (
      <div className="space-y-4">
        <div className="flex items-center justify-between flex-wrap gap-2">
          <div><h1 className="text-xl sm:text-2xl font-bold">Actas de reunión</h1><p className={cx("text-sm", t.muted)}>Con versionado e historial de cambios</p></div>
          {isAdmin && <div className="flex gap-2"><Btn icon={Sparkles} variant="emerald" onClick={openAI}>Generar con IA</Btn><Btn icon={Plus} onClick={openNew}>Nueva acta</Btn></div>}
        </div>
        {actas.map((a) => (
          <Card key={a.id} className="p-4">
            <div className="flex items-start justify-between flex-wrap gap-2">
              <div className="min-w-0"><div className="flex items-center gap-2 flex-wrap"><h3 className="font-semibold">{a.titulo}</h3><Chip color="#059669">v{a.v}</Chip></div>
                <div className={cx("text-xs mt-1 flex flex-wrap gap-x-3", t.muted)}><span>📅 {a.fecha}</span><span>📍 {a.lugar}</span><span>👥 {a.participantes}</span><span>✍️ {a.autor}</span></div></div>
              <div className="flex gap-1 shrink-0">
                <Btn variant="ghost" onClick={() => setOpenActa(openActa === a.id ? null : a.id)}>{openActa === a.id ? "Ocultar" : "Ver"}</Btn>
                {isAdmin && <Btn icon={PenLine} variant="ghost" onClick={() => openEdit(a)}>Editar</Btn>}
                <button onClick={() => showToast("Descargando acta…")} className={cx("p-2 rounded-xl border", t.border, t.hover)}><Download size={16} /></button>
              </div>
            </div>
            {openActa === a.id && (
              <div className="mt-3 pt-3 border-t space-y-3 text-sm" style={{ borderColor: dark ? "#1e293b" : "#e2e8f0" }}>
                <div><div className="font-medium mb-1">Temas tratados</div><ul className={cx("list-disc pl-5 space-y-0.5", t.muted)}>{a.temas.map((tm, i) => <li key={i}>{tm}</li>)}</ul></div>
                <div><div className="font-medium mb-1.5">Acuerdos y compromisos</div><div className="space-y-1.5">{a.acuerdos.map((ac, i) => (
                  <div key={i} className={cx("flex items-center gap-2 rounded-xl border p-2.5", t.border, t.alt)}>
                    <button onClick={() => toggleAcuerdo(a.id, i)}>{ac.done ? <CheckCircle2 size={16} className="text-emerald-600" /> : <Circle size={16} className={t.muted} />}</button>
                    <span className={cx("flex-1", ac.done && cx("line-through", t.muted))}>{ac.desc}</span>
                    <Chip color="#1d4ed8">{firstName(ac.resp)}</Chip><span className={cx("text-xs", t.muted)}>{ac.fecha}</span>
                  </div>))}</div></div>
              </div>
            )}
          </Card>
        ))}
      </div>
    );
  };

  /* ================= BIBLIOTECA (scoped) ================= */
  const renderBiblioteca = () => {
    const scopedFiles = FILES.filter((f) => assignedMicroId === null || f.micro === 0 || f.micro === assignedMicroId);
    const filtered = scopedFiles.filter((f) =>
      f.name.toLowerCase().includes(search.toLowerCase()) &&
      (bibCat === "all" || f.cat === bibCat) &&
      (bibMicroFilter === "all" || String(f.micro) === bibMicroFilter || (bibMicroFilter !== "all" && f.micro === 0)));
    return (
      <div className="space-y-4">
        <div><h1 className="text-xl sm:text-2xl font-bold">Biblioteca digital</h1><p className={cx("text-sm", t.muted)}>{assignedMicroId ? `Documentos de ${myMicroName} y recursos comunes` : "Documentos de todos los microcentros"}</p></div>
        <div className="flex gap-2 flex-wrap">
          <div className={cx("flex items-center gap-2 rounded-xl border px-3 flex-1", t.border, t.surface)}><Search size={16} className={t.muted} /><input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar documentos…" className={cx("bg-transparent py-2 text-sm outline-none flex-1", t.text)} /></div>
          {assignedMicroId === null && (<select value={bibMicroFilter} onChange={(e) => setBibMicroFilter(e.target.value)} className={cx("rounded-xl border px-3 py-2 text-sm", t.input)}><option value="all">Todos los microcentros</option>{MICROCENTROS.map((m) => <option key={m.id} value={String(m.id)}>{m.nombre}</option>)}</select>)}
          <select value={bibCat} onChange={(e) => setBibCat(e.target.value)} className={cx("rounded-xl border px-3 py-2 text-sm", t.input)}><option value="all">Todas las categorías</option>{FOLDERS.map((f) => <option key={f.name}>{f.name}</option>)}</select>
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">{FOLDERS.map((f) => (
          <button key={f.name} onClick={() => setBibCat(f.name)} className={cx("rounded-2xl border p-4 text-left transition", t.border, t.surface, t.hover, bibCat === f.name && "ring-2 ring-blue-500")}><FolderOpen size={24} style={{ color: f.color }} /><div className="text-sm font-medium mt-2 leading-snug">{f.name}</div><div className={cx("text-xs", t.muted)}>{f.count} archivos</div></button>))}</div>
        <Section title="Archivos" action={bibCat !== "all" && <button onClick={() => setBibCat("all")} className="text-xs text-blue-600 font-medium">Quitar filtro</button>}>
          <div className="space-y-1">
            {filtered.length === 0 && <p className={cx("text-sm py-2", t.muted)}>No se encontraron documentos con estos filtros.</p>}
            {filtered.map((f, i) => (
              <div key={i} className={cx("flex items-center gap-3 p-2.5 rounded-xl", t.hover)}>
                <div className="h-9 w-9 rounded-lg flex items-center justify-center shrink-0" style={{ backgroundColor: rgba(microColor(f.micro), 0.14) }}><FileText size={17} style={{ color: microColor(f.micro) }} /></div>
                <div className="flex-1 min-w-0"><div className="text-sm truncate">{f.name}</div><div className={cx("text-xs flex items-center gap-2 flex-wrap", t.muted)}><Chip color={microColor(f.micro)}>{microName(f.micro)}</Chip>{f.cat} · {f.year} · {f.size}</div></div>
                <button onClick={() => showToast("Descargando archivo…")}><Download size={16} className={t.muted} /></button>
              </div>))}
          </div>
        </Section>
        {assignedMicroId && <p className={cx("text-xs flex items-center gap-1.5", t.muted)}><Lock size={13} />No ves documentos de otros microcentros; sí los recursos marcados como “Común”.</p>}
      </div>
    );
  };

  /* ================= CALENDARIO ================= */
  const renderCalendario = () => {
    const first = 1;
    const days = Array.from({ length: 35 }, (_, i) => { const d = i - first + 1; return d >= 1 && d <= 31 ? d : null; });
    return (
      <div className="space-y-4">
        <div className="flex items-center justify-between flex-wrap gap-2"><div><h1 className="text-xl sm:text-2xl font-bold">Calendario</h1><p className={cx("text-sm", t.muted)}>Sincronizado con Google Calendar</p></div>{isAdmin && <Btn icon={Plus} onClick={() => showToast("Nuevo evento (demo)")}>Nuevo evento</Btn>}</div>
        <div className="grid lg:grid-cols-3 gap-4">
          <Card className="lg:col-span-2 p-4">
            <div className="flex items-center justify-between mb-3"><button className={cx("p-1.5 rounded-lg", t.hover)}><ChevronLeft size={18} /></button><h3 className="font-semibold">Julio 2025</h3><button className={cx("p-1.5 rounded-lg", t.hover)}><ChevronRight size={18} /></button></div>
            <div className="grid grid-cols-7 gap-1 text-center">
              {["L", "M", "M", "J", "V", "S", "D"].map((d, i) => <div key={i} className={cx("text-xs font-medium py-1", t.muted)}>{d}</div>)}
              {days.map((d, i) => (<div key={i} className={cx("aspect-square rounded-lg text-sm flex flex-col items-center justify-start pt-1.5", d && "border", d && t.border, d && EVENT_DAYS[d] && "font-semibold")}>{d && <span>{d}</span>}{d && EVENT_DAYS[d] && <span className="mt-0.5 h-1.5 w-1.5 rounded-full" style={{ backgroundColor: EVENT_DAYS[d].color }} />}</div>))}
            </div>
          </Card>
          <Section title="Eventos del mes"><div className="space-y-2.5">{Object.entries(EVENT_DAYS).map(([d, e], i) => (
            <div key={i} className="flex items-start gap-2.5"><div className="flex flex-col items-center justify-center rounded-lg text-white px-2 py-1 shrink-0" style={{ backgroundColor: e.color }}><span className="text-[9px] leading-none">JUL</span><span className="text-sm font-bold leading-none">{d}</span></div><div className="text-sm">{e.label}</div></div>))}</div></Section>
        </div>
      </div>
    );
  };

  /* ================= ADMIN ================= */
  const renderAdmin = () => (
    <div className="space-y-4">
      <div><h1 className="text-xl sm:text-2xl font-bold">Panel administrativo</h1><p className={cx("text-sm", t.muted)}>Gestión global de la plataforma</p></div>
      <div className="grid sm:grid-cols-3 gap-3">{[[Users, "Usuarios", "32 registrados", "#1d4ed8"], [Building2, "Microcentros", "5 activos", "#059669"], [ShieldCheck, "Respaldos", "Auto · diario", "#0891b2"]].map(([Ic, l, s, c], i) => (
        <Card key={i} className="p-4"><div className="h-9 w-9 rounded-xl flex items-center justify-center mb-2" style={{ backgroundColor: rgba(c, 0.14) }}><Ic size={18} color={c} /></div><div className="font-semibold text-sm">{l}</div><div className={cx("text-xs", t.muted)}>{s}</div></Card>))}</div>
      <Section title={`Solicitudes de registro (${pendingRequests.length})`} action={pendingRequests.length > 0 && <Chip color="#d97706">Pendientes</Chip>}>
        {pendingRequests.length === 0 ? <p className={cx("text-sm py-1", t.muted)}>No hay solicitudes pendientes por revisar.</p> : (
          <div className="space-y-2">
            {pendingRequests.map((r) => (
              <div key={r.id} className={cx("flex items-center gap-3 rounded-xl border p-3 flex-wrap", t.border, t.alt)}>
                <Avatar name={r.nombre} color="#7c3aed" size={38} />
                <div className="min-w-0 flex-1"><div className="text-sm font-medium truncate">{r.nombre} <span className={cx("text-xs font-normal", t.muted)}>· {r.rol}</span></div><div className={cx("text-xs truncate", t.muted)}>{r.correo} · {r.micro} · {r.establecimiento} · {r.fecha}</div></div>
                <div className="flex gap-1.5">
                  <button onClick={async () => { try { await aprobarUsuario(r.id); setPendingRequests(pendingRequests.filter((x) => x.id !== r.id)); showToast(`Cuenta de ${firstName(r.nombre)} aprobada`); } catch (err) { showToast(`Error: ${err.message}`); } }} className="inline-flex items-center gap-1 rounded-lg bg-emerald-600 text-white px-2.5 py-1.5 text-xs font-medium hover:bg-emerald-700"><CheckCircle2 size={13} />Aprobar</button>
                  <button onClick={async () => { try { await rechazarUsuario(r.id); setPendingRequests(pendingRequests.filter((x) => x.id !== r.id)); showToast(`Solicitud de ${firstName(r.nombre)} rechazada`); } catch (err) { showToast(`Error: ${err.message}`); } }} className={cx("inline-flex items-center gap-1 rounded-lg border px-2.5 py-1.5 text-xs font-medium", t.border, t.hover)}><XCircle size={13} />Rechazar</button>
                </div>
              </div>
            ))}
          </div>
        )}
      </Section>
      <div className="grid lg:grid-cols-2 gap-4">
        <Section title="Usuarios" action={<Btn icon={Plus} variant="ghost">Crear</Btn>}><div className="space-y-1">{MEMBERS.slice(0, 4).map((m, i) => (
          <div key={i} className={cx("flex items-center gap-3 p-2 rounded-xl", t.hover)}><Avatar name={m.name} color={m.color} size={32} /><div className="flex-1 min-w-0"><div className="text-sm font-medium truncate">{m.name}</div><div className={cx("text-xs truncate", t.muted)}>{m.rol}</div></div><UserCog size={16} className={t.muted} /></div>))}</div></Section>
        <Section title="Auditoría reciente"><div className="space-y-2.5">{AUDIT.map((a, i) => { const Ic = a.icon; return (
          <div key={i} className="flex items-center gap-3"><div className={cx("h-8 w-8 rounded-lg flex items-center justify-center shrink-0", t.alt)}><Ic size={15} className={t.muted} /></div><div className="flex-1 min-w-0 text-sm truncate">{a.txt}</div><div className={cx("text-xs shrink-0", t.muted)}>{a.t}</div></div>); })}</div></Section>
      </div>
    </div>
  );

  /* ================= PERFIL ================= */
  const renderPerfil = () => {
    const p = profiles[role];
    const set = (k, v) => setProfiles({ ...profiles, [role]: { ...p, [k]: v } });
    const setN = (k, v) => setProfiles({ ...profiles, [role]: { ...p, notif: { ...p.notif, [k]: v } } });
    const notifRows = [["correo", "Notificaciones por correo", "Recibe avisos importantes en tu email"], ["reuniones", "Recordatorios de reuniones", "Aviso el día previo a cada reunión"], ["resumen", "Resumen semanal", "Un correo con la actividad de tu microcentro"]];
    const guardarPerfil = () => {
      const mmdd = nacToMMDD(p.fechaNacimiento);
      if (mmdd) {
        setMembersByMicro((prev) => {
          const next = {}; let changed = false;
          Object.entries(prev).forEach(([k, list]) => {
            next[k] = list.map((mem) => {
              if (mem.name === p.nombre) { changed = true; return { ...mem, nac: mmdd }; }
              return mem;
            });
          });
          return changed ? next : prev;
        });
      }
      showToast("Perfil actualizado");
    };
    return (
      <div className="max-w-2xl mx-auto space-y-4">
        <div><h1 className="text-xl sm:text-2xl font-bold">Mi perfil</h1><p className={cx("text-sm", t.muted)}>Administra tu información personal y preferencias</p></div>

        <Card className="p-5 flex items-center gap-4 flex-wrap">
          <div className="relative">
            <Avatar name={p.nombre} color={roleCfg.color} size={64} />
            <button onClick={() => showToast("Selecciona una imagen (demo)")} className="absolute -bottom-1 -right-1 h-6 w-6 rounded-full bg-blue-600 text-white flex items-center justify-center border-2" style={{ borderColor: dark ? "#0f172a" : "#fff" }}><ImageIcon size={12} /></button>
          </div>
          <div className="min-w-0 flex-1"><div className="font-bold text-lg truncate">{p.nombre}</div><div className={cx("text-sm truncate", t.muted)}>{p.cargo}</div>
            <div className="mt-1"><Chip color={roleCfg.color}>{roleCfg.label}</Chip></div></div>
        </Card>

        <Section title="Datos personales">
          <div className="grid sm:grid-cols-2 gap-3">
            <div className="sm:col-span-2"><Field label="Nombre completo" value={p.nombre} onChange={(e) => set("nombre", e.target.value)} /></div>
            <Field label="Correo electrónico" type="email" value={p.correo} onChange={(e) => set("correo", e.target.value)} />
            <Field label="Teléfono" value={p.telefono} onChange={(e) => set("telefono", e.target.value)} />
            <Field label="RUT" value={p.rut} onChange={(e) => set("rut", e.target.value)} />
            <Field label="Fecha de nacimiento" type="date" value={p.fechaNacimiento || ""} onChange={(e) => set("fechaNacimiento", e.target.value)} />
            <Field label="Establecimiento" value={p.establecimiento} onChange={(e) => set("establecimiento", e.target.value)} />
          </div>
        </Section>

        <Section title="Rol y microcentro">
          <div className="space-y-2">
            <div className={cx("flex items-center justify-between rounded-xl border p-3", t.border, t.alt)}>
              <div className="flex items-center gap-2.5"><Shield size={17} style={{ color: roleCfg.color }} /><div><div className="text-sm font-medium">{p.cargo}</div><div className={cx("text-xs", t.muted)}>Rol en la plataforma</div></div></div>
              <Lock size={14} className={t.muted} />
            </div>
            <div className={cx("flex items-center justify-between rounded-xl border p-3", t.border, t.alt)}>
              <div className="flex items-center gap-2.5"><Building2 size={17} className="text-emerald-600" /><div><div className="text-sm font-medium">{myMicroName}</div><div className={cx("text-xs", t.muted)}>Microcentro asignado</div></div></div>
              <Lock size={14} className={t.muted} />
            </div>
            <p className={cx("text-xs flex items-center gap-1.5", t.muted)}><Lock size={12} />El rol y el microcentro los asigna la administración; no son editables por el usuario.</p>
          </div>
        </Section>

        <Section title="Seguridad">
          <div className="space-y-3">
            <Field label="Contraseña actual" type="password" placeholder="••••••••" />
            <div className="grid sm:grid-cols-2 gap-3"><Field label="Nueva contraseña" type="password" placeholder="Mínimo 8 caracteres" /><Field label="Confirmar nueva contraseña" type="password" placeholder="Repite la contraseña" /></div>
            <Btn icon={ShieldCheck} variant="ghost" onClick={() => showToast("Contraseña actualizada")}>Actualizar contraseña</Btn>
          </div>
        </Section>

        <Section title="Notificaciones">
          <div className="space-y-1">
            {notifRows.map(([k, title, desc]) => (
              <div key={k} className="flex items-center justify-between gap-3 py-2">
                <div className="min-w-0"><div className="text-sm font-medium">{title}</div><div className={cx("text-xs", t.muted)}>{desc}</div></div>
                <Toggle on={p.notif[k]} onClick={() => setN(k, !p.notif[k])} />
              </div>
            ))}
          </div>
        </Section>

        <div className="flex gap-2 flex-wrap">
          <Btn icon={Save} onClick={guardarPerfil}>Guardar cambios</Btn>
          <Btn variant="ghost" onClick={() => go("dashboard")}>Cancelar</Btn>
        </div>
      </div>
    );
  };

  const renderScreen = () => ({ dashboard: renderDashboard, microcentros: renderMicrocentros, grupo: renderGrupo, foro: renderForo, asistencia: renderAsistencia, actas: renderActas, biblioteca: renderBiblioteca, calendario: renderCalendario, admin: renderAdmin, perfil: renderPerfil }[screen]());

  /* ================= SHELL ================= */
  return (
    <div className={cx("min-h-screen", t.page, t.text)} style={font}>
      <header className={cx("sticky top-0 z-30 flex items-center gap-2 sm:gap-3 px-3 sm:px-4 h-14 border-b", t.border, t.surface)}>
        <button onClick={() => setDrawer(true)} className={cx("lg:hidden p-2 rounded-lg", t.hover)}><Menu size={20} /></button>
        <div className="font-semibold text-sm sm:text-base truncate">{screenTitle}</div>
          <div className="ml-auto flex items-center gap-1 sm:gap-2">
          <div className={cx("hidden md:flex items-center gap-2 rounded-xl border px-3", t.border, t.alt)}><Search size={15} className={t.muted} /><input placeholder="Buscar…" className={cx("bg-transparent py-1.5 text-sm outline-none w-36", t.text)} /></div>
          <button onClick={() => setDark(!dark)} className={cx("p-2 rounded-lg", t.hover)}>{dark ? <Sun size={18} /> : <Moon size={18} />}</button>
          <button className={cx("relative p-2 rounded-lg", t.hover)}><Bell size={18} /><span className="absolute top-1 right-1 h-2 w-2 rounded-full bg-red-500" /></button>
          <button onClick={() => go("perfil")} title="Mi perfil" className="ml-0.5"><Avatar name={profiles[role].nombre} color={roleCfg.color} size={30} /></button>
        </div>
      </header>

      <div className="flex">
        <aside className="hidden lg:flex flex-col w-64 shrink-0" style={{ minHeight: "calc(100vh - 56px)", backgroundColor: "#0b2540" }}>
          {brand}{contextChip}{renderNav()}{sidebarUser}
        </aside>

        {drawer && (
          <div className="lg:hidden fixed inset-0 z-40 flex">
            <div className="flex-1" style={{ backgroundColor: "rgba(0,0,0,0.5)" }} onClick={() => setDrawer(false)} />
            <div className="w-64 flex flex-col" style={{ backgroundColor: "#0b2540" }}>
              <div className="flex items-center justify-between pr-3">{brand}<button onClick={() => setDrawer(false)} className="text-blue-200 p-2"><X size={20} /></button></div>
              {contextChip}{renderNav()}{sidebarUser}
            </div>
          </div>
        )}

        <main className="flex-1 min-w-0 p-4 sm:p-6 pb-16">{renderScreen()}</main>
      </div>

      {toast && <div className="fixed bottom-5 left-1/2 -translate-x-1/2 z-50 rounded-xl bg-slate-900 text-white text-sm px-4 py-2.5 shadow-xl flex items-center gap-2"><CheckCircle2 size={16} className="text-emerald-400" />{toast}</div>}
    </div>
  );
}
