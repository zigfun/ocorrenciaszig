import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── CREDENCIAIS (fixas para todos os usuários) ───────────────────
const FIXED_KEY = "$2a$10$qNAPL3qG9Q6ecS4yKo7Sm.IFi2M/meue8astJrVZZOD.rm.iGBUm.";
const FIXED_BIN = "6a160e16f47d5c455c3aac9c";
const JB = "https://api.jsonbin.io/v3";
async function jbFetch(path, method = "GET", body = null, key) {
  const r = await fetch(`${JB}${path}`, {
    method,
    headers: { "Content-Type": "application/json", "X-Master-Key": key },
    body: body ? JSON.stringify(body) : null,
  });
  if (!r.ok) { let m = `HTTP ${r.status}`; try { const j = await r.json(); m = j.message || m; } catch {} throw new Error(m); }
  return r.json();
}
async function createBin(key, data) { const r = await jbFetch("/b", "POST", data, key); return r.metadata.id; }
async function readBin(key, id)     { const r = await jbFetch(`/b/${id}/latest`, "GET", null, key); return r.record; }
async function writeBin(key, id, d) { await jbFetch(`/b/${id}`, "PUT", d, key); }

// ─── GITHUB (para ocorrências) ────────────────────────────────────
const GH_RAW = "https://raw.githubusercontent.com/zigfun/ocorrenciaszig/main/data.json";
async function fetchGitHub() {
  const r = await fetch(GH_RAW + "?t=" + Date.now());
  if (!r.ok) throw new Error("Erro ao buscar dados do GitHub");
  return r.json();
}

// ─── MAPEAMENTO DOS CAMPOS DO EXCEL ──────────────────────────────
function excelDateToISO(val) {
  if (!val) return "";
  const s = String(val).trim();
  if (s.includes("/")) {
    const [d, m, y] = s.split("/");
    return `${y}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;
  }
  if (!isNaN(s)) {
    const d = new Date(Math.round((parseFloat(s) - 25569) * 86400 * 1000));
    return d.toISOString().split("T")[0];
  }
  return s;
}

function mapRow(row, i) {
  const statusRaw = (row["Status"] || "").toLowerCase();
  const status = statusRaw.includes("conclu") ? "resolvido" : "andamento";

  // Parse multiple attachment URLs (separated by newline or semicolon)
  const attRaw = row["Prints/Audios/PDFs"] || row["Prints/Áudios/PDFs"] || "";
  const attachments = attRaw
    ? attRaw.split(/;\s+|\n+/).map(s => s.trim()).filter(s => s.startsWith("http")).map(url => {
        const name = decodeURIComponent(url.split("/").pop().split("?")[0]) || "Arquivo";
        const ext = name.split(".").pop().toLowerCase();
        const type = ["jpg","jpeg","png","gif","webp"].includes(ext) ? "image"
          : ["mp3","ogg","opus","m4a","aac","wav"].includes(ext) ? "audio"
          : ext === "pdf" ? "pdf" : "file";
        return { url, name, type };
      })
    : [];

  return {
    id: row["Id"] || row["ID"] || i,
    client:   row["Cliente"]           || "",
    date:     excelDateToISO(row["Data da ocorrência"] || row["Data da ocorrencia"] || ""),
    am:       row["AM responsável"]    || row["AM responsavel"] || "",
    coord:    row["Time"]              || "",
    category: row["Área"]              || row["Area"]           || "",
    status,
    desc:     row["Descrição"]         || row["Descricao"]      || "",
    registeredBy: row["Nome"] || "Forms",
    attachments,
  };
}

// ─── HELPERS ─────────────────────────────────────────────────────
function simpleHash(s) { let h = 0; for (let i = 0; i < s.length; i++) { h = ((h << 5) - h) + s.charCodeAt(i); h |= 0; } return h.toString(36); }
function fmt(d) {
  if (!d) return "";
  try {
    const s = d.includes("/") ? d.split("/").reverse().join("-") : d;
    return new Date(s + "T12:00:00").toLocaleDateString("pt-BR");
  } catch { return d; }
}

const AREA_COLORS = {
  "CSI":                    { c: "#1a5fd4", bg: "rgba(26,95,212,.1)"   },
  "CSM":                    { c: "#7c3aed", bg: "rgba(124,58,237,.1)"  },
  "Logística/ Estoque":     { c: "#0d8a60", bg: "rgba(13,138,96,.1)"   },
  "BKO":                    { c: "#c97a00", bg: "rgba(201,122,0,.1)"   },
  "Jurídico":               { c: "#d63b3b", bg: "rgba(214,59,59,.1)"   },
  "Negociação/ Renegociação":{ c:"#0891b2", bg: "rgba(8,145,178,.1)"   },
  "Tech/ Produto":          { c: "#7c3aed", bg: "rgba(124,58,237,.1)"  },
};
function aColor(area) { return AREA_COLORS[area] || { c: "#6b7280", bg: "rgba(107,114,128,.1)" }; }

const MONTHS = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

// ─── CSS ─────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Fraunces:opsz,wght@9..144,700&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
:root{
  --bg:#f4f6ff;--sf:#fff;--s2:#eef1fb;--s3:#e3e8f7;
  --bd:#d0d9f0;--ac:#2952cc;--ac2:#1a3fa3;--tx:#0d1b40;--mu:#6b7a9f;
  --err:#d63b3b;--ok:#0a8a5c;
  --fh:'Fraunces',serif;--fb:'Plus Jakarta Sans',sans-serif;
}
body{background:var(--bg);color:var(--tx);font-family:var(--fb);-webkit-font-smoothing:antialiased;}
.auth{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;background:linear-gradient(135deg,#eef1fb,#f4f6ff 60%,#ede8ff);}
.auth-card{background:var(--sf);border:1px solid var(--bd);border-radius:20px;padding:40px 36px;width:100%;max-width:420px;box-shadow:0 8px 40px rgba(41,82,204,.1);}
.auth-brand{display:flex;align-items:center;gap:10px;margin-bottom:28px;}
.auth-ic{width:42px;height:42px;background:var(--ac);border-radius:11px;display:flex;align-items:center;justify-content:center;font-size:20px;}
.auth-name{font-family:var(--fh);font-size:18px;font-weight:700;}
.auth-sub{font-size:11px;color:var(--mu);margin-top:1px;}
.auth-title{font-family:var(--fh);font-size:22px;margin-bottom:6px;}
.auth-desc{font-size:13px;color:var(--mu);margin-bottom:24px;line-height:1.5;}
.fl{font-size:11px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.09em;display:block;margin-bottom:5px;margin-top:14px;}
.fi{width:100%;background:var(--s2);border:1.5px solid var(--bd);color:var(--tx);font-family:var(--fb);font-size:13px;padding:11px 13px;border-radius:9px;outline:none;transition:border-color .15s;}
.fi:focus{border-color:var(--ac);}
.ferr{font-size:12px;color:var(--err);margin-top:8px;}
.btn-main{width:100%;background:var(--ac);color:#fff;font-family:var(--fb);font-weight:700;font-size:14px;padding:13px;border:none;border-radius:10px;cursor:pointer;margin-top:16px;transition:background .15s;}
.btn-main:hover{background:var(--ac2);}
.btn-main:disabled{background:var(--mu);cursor:not-allowed;}
.auth-note{font-size:11px;color:var(--mu);text-align:center;margin-top:10px;}
.setup-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
.steps{display:flex;flex-direction:column;gap:9px;margin-bottom:22px;}
.sst{display:flex;gap:11px;padding:11px 13px;border-radius:10px;background:var(--s2);border:1px solid var(--bd);}
.sst-n{width:22px;height:22px;min-width:22px;border-radius:50%;background:var(--ac);color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;margin-top:1px;}
.sst-t{font-size:12px;color:var(--tx);line-height:1.55;}
.sst-t a{color:var(--ac);font-weight:600;text-decoration:none;}
.sst-t strong{color:var(--ac);}
.app{display:flex;height:100vh;overflow:hidden;}
.sb{width:230px;min-width:230px;background:var(--sf);border-right:1px solid var(--bd);display:flex;flex-direction:column;}
.sb-top{padding:18px 16px 14px;border-bottom:1px solid var(--bd);}
.sb-brand{display:flex;align-items:center;gap:9px;}
.sb-ic{width:32px;height:32px;background:var(--ac);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:15px;flex-shrink:0;}
.sb-name{font-family:var(--fh);font-size:14px;font-weight:700;}
.sb-sub{font-size:10px;color:var(--mu);}
.sb-user{margin:10px 12px;padding:9px 12px;background:var(--s2);border-radius:9px;border:1px solid var(--bd);}
.sb-uname{font-size:12px;font-weight:700;}
.sb-uemail{font-size:10px;color:var(--mu);margin-top:1px;}
.sb-urole{font-size:9px;font-weight:700;color:var(--ac);text-transform:uppercase;background:rgba(41,82,204,.1);padding:1px 6px;border-radius:4px;display:inline-block;margin-top:3px;}
.nav-sec{padding:10px 11px 4px;}
.nav-lbl{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.12em;padding:0 7px;margin-bottom:4px;}
.ni{display:flex;align-items:center;gap:8px;padding:8px 11px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:500;color:var(--mu);transition:all .15s;position:relative;}
.ni:hover{background:var(--s2);color:var(--tx);}
.ni.on{background:var(--s2);color:var(--tx);font-weight:600;}
.ni.on::before{content:'';position:absolute;left:0;top:5px;bottom:5px;width:3px;border-radius:0 3px 3px 0;background:var(--ac);}
.ni-ic{font-size:13px;width:15px;text-align:center;}
.ni-badge{margin-left:auto;background:var(--sf);border:1px solid var(--bd);font-size:10px;font-weight:600;color:var(--mu);padding:1px 6px;border-radius:20px;}
.ni.on .ni-badge{background:var(--ac);color:#fff;border-color:var(--ac);}
.sb-foot{margin-top:auto;padding:12px 14px;border-top:1px solid var(--bd);display:flex;justify-content:space-between;align-items:center;}
.sb-logout{font-size:12px;color:var(--mu);cursor:pointer;}
.sb-logout:hover{color:var(--err);}
.sb-sync{font-size:10px;padding:3px 8px;border-radius:20px;border:1px solid;}
.sb-sync.ok{color:var(--ok);border-color:rgba(10,138,92,.25);}
.sb-sync.load{color:var(--mu);border-color:var(--bd);}
.sb-sync.err{color:var(--err);border-color:rgba(214,59,59,.25);}
.main{flex:1;display:flex;flex-direction:column;overflow:hidden;}
.topbar{background:var(--sf);border-bottom:1px solid var(--bd);padding:14px 24px;display:flex;align-items:center;gap:10px;}
.page-title{font-family:var(--fh);font-size:20px;font-weight:700;flex:1;}
.page-title span{color:var(--ac);}
.btn{font-family:var(--fb);font-weight:600;font-size:13px;padding:8px 14px;border-radius:8px;cursor:pointer;display:flex;align-items:center;gap:5px;transition:all .15s;white-space:nowrap;border:none;}
.btn-p{background:var(--ac);color:#fff;}.btn-p:hover{background:var(--ac2);}
.btn-o{background:transparent;color:var(--mu);border:1px solid var(--bd)!important;}.btn-o:hover{border-color:var(--ok)!important;color:var(--ok);}
.btn-g{background:transparent;color:var(--mu);border:1px solid var(--bd)!important;}.btn-g:hover{color:var(--tx);}
.filters{padding:10px 24px;background:var(--sf);border-bottom:1px solid var(--bd);display:flex;gap:8px;align-items:center;flex-wrap:wrap;}
.f-lbl{font-size:11px;color:var(--mu);font-weight:600;}
.fsel{background:var(--s2);border:1px solid var(--bd);color:var(--tx);font-family:var(--fb);font-size:12px;padding:6px 10px;border-radius:7px;outline:none;}
.f-count{margin-left:auto;font-size:11px;color:var(--mu);}
.f-clear{font-size:11px;color:var(--ac);cursor:pointer;font-weight:600;}
.content{flex:1;overflow-y:auto;padding:22px 24px;}
.stats-row{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:11px;margin-bottom:22px;}
.stat{background:var(--sf);border:1px solid var(--bd);border-radius:11px;padding:14px 16px;cursor:pointer;transition:all .15s;position:relative;overflow:hidden;}
.stat:hover{border-color:var(--ac);transform:translateY(-1px);}
.stat.on{border-color:var(--ac);background:var(--s2);}
.stat::after{content:'';position:absolute;right:-8px;bottom:-8px;width:50px;height:50px;border-radius:50%;background:var(--sc,#2952cc);opacity:.07;}
.stat-lbl{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.08em;margin-bottom:3px;}
.stat-num{font-family:var(--fh);font-size:26px;font-weight:700;margin-bottom:1px;}
.stat-sub{font-size:10px;color:var(--mu);}
.sec-hd{display:flex;align-items:center;gap:9px;margin-bottom:12px;}
.sec-title{font-family:var(--fh);font-size:14px;font-weight:700;}
.sec-line{flex:1;height:1px;background:var(--bd);}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:11px;margin-bottom:24px;}
.card{background:var(--sf);border:1px solid var(--bd);border-radius:11px;padding:15px;cursor:pointer;transition:all .15s;}
.card:hover{border-color:var(--ac);box-shadow:0 3px 14px rgba(41,82,204,.08);transform:translateY(-1px);}
.card-top{display:flex;align-items:flex-start;justify-content:space-between;gap:8px;margin-bottom:8px;}
.card-client{font-weight:700;font-size:13.5px;}
.card-tags{display:flex;gap:4px;flex-wrap:wrap;justify-content:flex-end;flex-shrink:0;}
.tag{font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;text-transform:uppercase;letter-spacing:.05em;white-space:nowrap;}
.tag-a{background:rgba(201,122,0,.12);color:#a35e00;border:1px solid rgba(201,122,0,.25);}
.tag-r{background:rgba(10,138,92,.12);color:#0a7a52;border:1px solid rgba(10,138,92,.25);}
.card-meta{display:flex;gap:9px;flex-wrap:wrap;margin-bottom:6px;}
.card-meta span{font-size:11px;color:var(--mu);}
.card-desc{font-size:12px;color:var(--mu);line-height:1.5;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;}
.kanban{display:flex;gap:14px;padding:20px 24px;overflow-x:auto;overflow-y:hidden;align-items:flex-start;flex:1;}
.kb-col{flex:0 0 300px;background:var(--s2);border:1px solid var(--bd);border-radius:11px;display:flex;flex-direction:column;max-height:calc(100vh - 170px);}
.kb-head{padding:12px 14px;border-bottom:1px solid var(--bd);display:flex;align-items:center;gap:8px;border-radius:11px 11px 0 0;}
.kb-title{font-family:var(--fh);font-size:13px;font-weight:700;flex:1;}
.kb-count{font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px;}
.kb-body{overflow-y:auto;padding:9px;display:flex;flex-direction:column;gap:8px;}
.kb-card{background:var(--sf);border:1px solid var(--bd);border-radius:8px;padding:12px;cursor:pointer;transition:all .15s;}
.kb-card:hover{border-color:var(--ac);transform:translateY(-1px);}
.kb-client{font-weight:700;font-size:12px;margin-bottom:5px;}
.kb-meta{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:5px;}
.kb-meta span{font-size:11px;color:var(--mu);}
.kb-desc{font-size:11px;color:var(--mu);line-height:1.5;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;}
.ov{position:fixed;inset:0;background:rgba(13,27,64,.65);z-index:100;display:flex;align-items:center;justify-content:center;padding:20px;backdrop-filter:blur(6px);}
.modal{background:var(--sf);border:1px solid var(--bd);border-radius:16px;width:100%;max-width:600px;max-height:92vh;overflow-y:auto;padding:26px;}
.modal-title{font-family:var(--fh);font-size:17px;font-weight:700;margin-bottom:18px;}
.modal-title span{color:var(--ac);}
.det-hd{display:flex;align-items:flex-start;gap:11px;margin-bottom:16px;}
.det-client{font-family:var(--fh);font-size:21px;font-weight:700;flex:1;}
.det-x{width:30px;height:30px;border-radius:7px;border:1px solid var(--bd);background:var(--s2);cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:14px;color:var(--mu);flex-shrink:0;}
.det-x:hover{background:var(--err);border-color:var(--err);color:white;}
.det-pills{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:16px;}
.det-pill{background:var(--s2);border:1px solid var(--bd);border-radius:20px;padding:4px 11px;font-size:12px;color:var(--mu);}
.det-sec{margin-bottom:16px;}
.det-sec-lbl{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.1em;margin-bottom:7px;}
.det-desc{background:var(--s2);border:1px solid var(--bd);border-radius:9px;padding:13px 15px;font-size:13px;line-height:1.7;color:var(--tx);white-space:pre-wrap;}
.det-grid{display:grid;grid-template-columns:1fr 1fr;gap:7px;}
.det-item{background:var(--s2);border:1px solid var(--bd);border-radius:8px;padding:9px 11px;}
.det-item-lbl{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.07em;margin-bottom:2px;}
.det-item-val{font-size:13px;font-weight:600;color:var(--tx);}
.users-wrap{flex:1;overflow-y:auto;padding:20px 24px;}
.users-card{background:var(--sf);border:1px solid var(--bd);border-radius:11px;overflow:hidden;margin-bottom:16px;}
.ut-head{display:grid;grid-template-columns:1fr 1fr 80px 40px;background:var(--s2);border-bottom:1px solid var(--bd);padding:10px 16px;}
.ut-head span{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.09em;}
.ut-row{display:grid;grid-template-columns:1fr 1fr 80px 40px;padding:12px 16px;border-bottom:1px solid var(--bd);align-items:center;}
.ut-row:last-child{border-bottom:none;}
.ut-row:hover{background:var(--s2);}
.ut-name{font-weight:600;font-size:13px;}
.ut-email{font-size:12px;color:var(--mu);}
.ut-role{font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;display:inline-block;}
.ut-role.admin{background:rgba(41,82,204,.12);color:var(--ac);}
.ut-role.user{background:var(--s2);color:var(--mu);}
.del-ic{width:26px;height:26px;border-radius:6px;border:1px solid var(--bd);background:transparent;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:12px;color:var(--mu);}
.del-ic:hover{background:var(--err);border-color:var(--err);color:#fff;}
.add-form{background:var(--sf);border:1px solid var(--bd);border-radius:11px;padding:20px;}
.add-title{font-size:13px;font-weight:700;margin-bottom:14px;}
.add-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px;}
.tog-row{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--mu);cursor:pointer;margin:10px 0 14px;user-select:none;}
.tog{width:28px;height:16px;border-radius:9px;background:var(--bd);position:relative;transition:background .15s;flex-shrink:0;}
.tog.on{background:var(--ac);}
.tog::after{content:'';position:absolute;width:10px;height:10px;border-radius:50%;background:white;top:3px;left:3px;transition:left .15s;}
.tog.on::after{left:15px;}
.att-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:8px;}
.att-item{display:flex;align-items:center;gap:8px;background:var(--s2);border:1px solid var(--bd);border-radius:9px;padding:10px 12px;text-decoration:none;color:var(--tx);transition:all .15s;cursor:pointer;}
.att-item:hover{border-color:var(--ac);background:var(--sf);}
.att-item-ic{font-size:20px;flex-shrink:0;}
.att-item-name{font-size:11px;font-weight:600;color:var(--tx);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.att-item-type{font-size:10px;color:var(--mu);margin-top:1px;}
.empty-ic{font-size:36px;margin-bottom:10px;}
.empty-t{font-size:13px;line-height:1.6;}
::-webkit-scrollbar{width:5px;}
::-webkit-scrollbar-track{background:transparent;}
::-webkit-scrollbar-thumb{background:var(--bd);border-radius:3px;}
`;

export default function App() {
  const [apiKey]  = useState(FIXED_KEY);
  const [binId]   = useState(FIXED_BIN);
  const [session, setSession] = useState(() => { try { return JSON.parse(localStorage.getItem("zg_sess") || "null"); } catch { return null; } });

  const [incidents, setIncidents] = useState([]);
  const [users,     setUsers]     = useState([]);
  const [sync,      setSync]      = useState("ok");

  const [view,   setView]   = useState("dash");
  const [active, setActive] = useState(null);

  const [filterArea,   setFilterArea]   = useState("");
  const [filterAM,     setFilterAM]     = useState("");
  const [filterStatus, setFilterStatus] = useState("");
  const [filterMonth,  setFilterMonth]  = useState("");

  const [lEmail, setLEmail] = useState("");
  const [lPass,  setLPass]  = useState("");
  const [lErr,   setLErr]   = useState("");

  const [nu,   setNu]   = useState({ name:"", email:"", password:"", isAdmin:false });
  const [nuErr,setNuErr]= useState("");

  const loggedIn = !!session;

  const load = useCallback(async () => {
    if (!apiKey || !binId) return;
    setSync("load");
    try {
      const ghData = await fetchGitHub();
      const raw = ghData.incidents || [];
      setIncidents(raw.map((r, i) => mapRow(r, i)));
      const jbData = await readBin(apiKey, binId);
      setUsers(jbData.users || []);
      setSync("ok");
    } catch { setSync("err"); }
  }, [apiKey, binId]);

  // Carrega usuários antes do login
  useEffect(() => {
    (async () => {
      try {
        const jbData = await readBin(apiKey, binId);
        setUsers(jbData.users || []);
      } catch {}
    })();
  }, []);

  useEffect(() => {
    if (!loggedIn) return;
    load();
    const iv = setInterval(load, 30000);
    return () => clearInterval(iv);
  }, [loggedIn, load]);

  const persistUsers = async (nextUsers) => {
    setSync("load");
    try { await writeBin(apiKey, binId, { users: nextUsers }); setSync("ok"); }
    catch { setSync("err"); }
  };

  const handleLogin = () => {
    const u = users.find(u => u.email === lEmail.toLowerCase().trim() && u.passwordHash === simpleHash(lPass));
    if (!u) { setLErr("Email ou senha incorretos."); return; }
    const sess = { id:u.id, name:u.name, email:u.email, isAdmin:u.isAdmin };
    localStorage.setItem("zg_sess", JSON.stringify(sess));
    setSession(sess); setLErr("");
  };

  const handleLogout = () => { localStorage.removeItem("zg_sess"); setSession(null); setLEmail(""); setLPass(""); };

  const addUser = async () => {
    if (!nu.name.trim()||!nu.email.trim()||!nu.password.trim()) { setNuErr("Preencha todos os campos."); return; }
    if (nu.password.length < 6) { setNuErr("Senha mínima de 6 caracteres."); return; }
    if (users.find(u => u.email === nu.email.toLowerCase().trim())) { setNuErr("Email já cadastrado."); return; }
    const user = { id:"u_"+Date.now(), name:nu.name.trim(), email:nu.email.toLowerCase().trim(), passwordHash:simpleHash(nu.password), isAdmin:nu.isAdmin };
    const next = [...users, user];
    setUsers(next); await persistUsers(next);
    setNu({ name:"", email:"", password:"", isAdmin:false }); setNuErr("");
  };

  const removeUser = async (id) => {
    if (id === session?.id) { alert("Você não pode remover sua própria conta."); return; }
    if (!confirm("Remover usuário?")) return;
    const next = users.filter(u => u.id !== id);
    setUsers(next); await persistUsers(next);
  };

  const exportXlsx = () => {
    const rows = filtered.map(i => ({
      "Cliente":i.client,"Data":fmt(i.date),"AM":i.am,
      "Time":i.coord,"Área":i.category,
      "Status":i.status==="resolvido"?"Concluído":"Em andamento",
      "Descrição":i.desc
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = Object.keys(rows[0]||{}).map(()=>({wch:22}));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Ocorrências");
    XLSX.writeFile(wb, `ocorrencias_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  const allAreas = [...new Set(incidents.map(i=>i.category).filter(Boolean))].sort();
  const allAMs   = [...new Set(incidents.map(i=>i.am).filter(Boolean))].sort();
  const anyFilter = filterArea||filterAM||filterStatus||filterMonth;

  const filtered = incidents.filter(i => {
    if (filterArea   && i.category !== filterArea) return false;
    if (filterAM     && i.am !== filterAM)         return false;
    if (filterStatus && i.status !== filterStatus) return false;
    if (filterMonth !== "") {
      const m = i.date ? (() => { try { const s=i.date.includes("/")?i.date.split("/").reverse().join("-"):i.date; return new Date(s+"T12:00:00").getMonth(); } catch { return -1; } })() : -1;
      if (m !== parseInt(filterMonth)) return false;
    }
    return true;
  });

  // LOGIN
  if (!loggedIn) return (
    <>
      <style>{CSS}</style>
      <div className="auth">
        <div className="auth-card">
          <div className="auth-brand"><div className="auth-ic">📋</div><div><div className="auth-name">Ocorrências ZIG</div><div className="auth-sub">Central de registros</div></div></div>
          <div className="auth-title">Entrar</div>
          <div className="auth-desc">Use seu email e senha cadastrados pelo administrador.</div>
          <label className="fl">Email</label>
          <input className="fi" type="email" placeholder="seu@email.com" value={lEmail} onChange={e=>setLEmail(e.target.value)}/>
          <label className="fl">Senha</label>
          <input className="fi" type="password" placeholder="••••••" value={lPass} onChange={e=>setLPass(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()}/>
          {lErr && <div className="ferr">⚠ {lErr}</div>}
          <button className="btn-main" onClick={handleLogin}>Entrar →</button>
          <div className="auth-note">Não tem acesso? Fale com o administrador.</div>
        </div>
      </div>
    </>
  );

  return (
    <>
      <style>{CSS}</style>
      <div className="app">
        <aside className="sb">
          <div className="sb-top">
            <div className="sb-brand"><div className="sb-ic">📋</div><div><div className="sb-name">Ocorrências</div><div className="sb-sub">ZIG Funtech</div></div></div>
          </div>
          <div className="sb-user">
            <div className="sb-uname">{session.name}</div>
            <div className="sb-uemail">{session.email}</div>
            <span className="sb-urole">{session.isAdmin?"Admin":"Usuário"}</span>
          </div>
          <div className="nav-sec">
            <div className="nav-lbl">Visualização</div>
            <div className={`ni ${view==="dash"?"on":""}`} onClick={()=>setView("dash")}>
              <span className="ni-ic">⊞</span> Dashboard <span className="ni-badge">{incidents.length}</span>
            </div>
            <div className={`ni ${view==="kanban"?"on":""}`} onClick={()=>setView("kanban")}>
              <span className="ni-ic">🗂</span> Kanban
              <span className="ni-badge">{incidents.filter(i=>i.status==="andamento").length}</span>
            </div>
            {session.isAdmin && (
              <div className={`ni ${view==="users"?"on":""}`} onClick={()=>setView("users")}>
                <span className="ni-ic">👥</span> Usuários <span className="ni-badge">{users.length}</span>
              </div>
            )}
          </div>
          {allAreas.length > 0 && (
            <div className="nav-sec">
              <div className="nav-lbl">Áreas</div>
              {allAreas.map(a => {
                const ac = aColor(a);
                return (
                  <div key={a} className={`ni ${filterArea===a?"on":""}`}
                    onClick={()=>{setView("dash");setFilterArea(filterArea===a?"":a);}}>
                    <span style={{width:"8px",height:"8px",borderRadius:"50%",background:ac.c,flexShrink:0,display:"inline-block"}}></span>
                    <span style={{fontSize:"12px",flex:1}}>{a}</span>
                    <span className="ni-badge">{incidents.filter(i=>i.category===a).length}</span>
                  </div>
                );
              })}
            </div>
          )}
          <div className="sb-foot">
            <span className="sb-logout" onClick={handleLogout}>🚪 Sair</span>
            <span className={`sb-sync ${sync}`}>
              {sync==="ok"&&"● Sync"}{sync==="load"&&"○ ..."}{sync==="err"&&"● Erro"}
            </span>
          </div>
        </aside>

        <main className="main">
          <div className="topbar">
            <div className="page-title">
              {view==="kanban"?<>Kanban de <span>Status</span></>
              :view==="users"?<><span>Usuários</span></>
              :<>Dashboard de <span>Ocorrências</span></>}
            </div>
            {view!=="users"&&incidents.length>0&&<button className="btn btn-o" onClick={exportXlsx}>↓ Excel</button>}
            {view!=="users"&&<button className="btn btn-o" onClick={load}>↺ Atualizar</button>}
          </div>

          {view!=="users"&&(
            <div className="filters">
              <span className="f-lbl">Filtrar:</span>
              <select className="fsel" value={filterArea} onChange={e=>setFilterArea(e.target.value)}>
                <option value="">Todas as áreas</option>
                {allAreas.map(a=><option key={a}>{a}</option>)}
              </select>
              <select className="fsel" value={filterAM} onChange={e=>setFilterAM(e.target.value)}>
                <option value="">Todos os AMs</option>
                {allAMs.map(a=><option key={a}>{a}</option>)}
              </select>
              <select className="fsel" value={filterMonth} onChange={e=>setFilterMonth(e.target.value)}>
                <option value="">Todos os meses</option>
                {MONTHS.map((m,i)=><option key={i} value={i}>{m}</option>)}
              </select>
              <select className="fsel" value={filterStatus} onChange={e=>setFilterStatus(e.target.value)}>
                <option value="">Todos os status</option>
                <option value="andamento">Em andamento</option>
                <option value="resolvido">Concluído</option>
              </select>
              {anyFilter&&<span className="f-clear" onClick={()=>{setFilterArea("");setFilterAM("");setFilterStatus("");setFilterMonth("");}}>✕ Limpar</span>}
              <span className="f-count">{filtered.length} registro{filtered.length!==1?"s":""}</span>
            </div>
          )}

          {view==="dash"&&(
            <div className="content">
              {!anyFilter&&allAreas.length>0&&(
                <div className="stats-row">
                  {allAreas.map(a=>{
                    const ac=aColor(a); const count=incidents.filter(i=>i.category===a).length;
                    return(
                      <div key={a} className={`stat ${filterArea===a?"on":""}`} style={{"--sc":ac.c}} onClick={()=>setFilterArea(filterArea===a?"":a)}>
                        <div className="stat-lbl">{a}</div>
                        <div className="stat-num" style={{color:ac.c}}>{count}</div>
                        <div className="stat-sub">ocorrência{count!==1?"s":""}</div>
                      </div>
                    );
                  })}
                </div>
              )}
              {filtered.length===0?(
                <div className="empty">
                  <div className="empty-ic">{incidents.length===0?"📋":"🔍"}</div>
                  <div className="empty-t">{incidents.length===0?"Aguardando dados do Forms.\nAs respostas aparecerão aqui automaticamente.":"Nenhum resultado para os filtros aplicados."}</div>
                </div>
              ):(
                (filterArea?[filterArea]:allAreas).map(area=>{
                  const aIncs=filtered.filter(i=>i.category===area);
                  if(!aIncs.length) return null;
                  const ac=aColor(area);
                  return(
                    <div key={area}>
                      <div className="sec-hd">
                        <span className="sec-title" style={{color:ac.c}}>● {area}</span>
                        <div className="sec-line"/>
                        <span style={{fontSize:"11px",color:"var(--mu)"}}>{aIncs.length}</span>
                      </div>
                      <div className="grid">{aIncs.map((inc,i)=><IncCard key={i} inc={inc} onClick={()=>setActive(inc)}/>)}</div>
                    </div>
                  );
                })
              )}
            </div>
          )}

          {view==="kanban"&&(
            <div className="kanban">
              {[
                {key:"andamento",label:"Em andamento 🔄",color:"#a35e00",bg:"rgba(201,122,0,.12)",bd:"rgba(201,122,0,.28)"},
                {key:"resolvido",label:"Concluído ✅",color:"#0a7a52",bg:"rgba(10,138,92,.12)",bd:"rgba(10,138,92,.28)"},
              ].map(col=>{
                const colIncs=filtered.filter(i=>i.status===col.key);
                return(
                  <div className="kb-col" key={col.key}>
                    <div className="kb-head">
                      <div className="kb-title" style={{color:col.color}}>{col.label}</div>
                      <span className="kb-count" style={{background:col.bg,color:col.color,border:`1px solid ${col.bd}`}}>{colIncs.length}</span>
                    </div>
                    <div className="kb-body">
                      {colIncs.length===0&&<div style={{textAlign:"center",padding:"20px 0",fontSize:"12px",color:"var(--mu)"}}>Nenhuma ocorrência</div>}
                      {colIncs.map((inc,i)=>{
                        const ac=aColor(inc.category);
                        return(
                          <div className="kb-card" key={i} onClick={()=>setActive(inc)}>
                            <div className="kb-client">{inc.client}</div>
                            <div className="kb-meta">
                              <span className="tag" style={{background:ac.bg,color:ac.c}}>{inc.category}</span>
                              {inc.date&&<span>📅 {fmt(inc.date)}</span>}
                              {inc.am&&<span>👤 {inc.am}</span>}
                            </div>
                            {inc.desc&&<div className="kb-desc">{inc.desc}</div>}
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {view==="users"&&(
            <div className="users-wrap">
              <div className="users-card">
                <div className="ut-head"><span>Nome</span><span>Email</span><span>Perfil</span><span></span></div>
                {users.map(u=>(
                  <div key={u.id} className="ut-row">
                    <div className="ut-name">{u.name}{u.id===session.id?" (você)":""}</div>
                    <div className="ut-email">{u.email}</div>
                    <span className={`ut-role ${u.isAdmin?"admin":"user"}`}>{u.isAdmin?"Admin":"Usuário"}</span>
                    <div>{u.id!==session.id&&<div className="del-ic" onClick={()=>removeUser(u.id)}>🗑</div>}</div>
                  </div>
                ))}
              </div>
              <div className="add-form">
                <div className="add-title">＋ Adicionar usuário</div>
                <div className="add-grid">
                  <div><label className="fl">Nome</label><input className="fi" placeholder="Nome completo" value={nu.name} onChange={e=>setNu(p=>({...p,name:e.target.value}))}/></div>
                  <div><label className="fl">Email</label><input className="fi" placeholder="email@zig.fun" value={nu.email} onChange={e=>setNu(p=>({...p,email:e.target.value}))}/></div>
                </div>
                <div><label className="fl">Senha inicial</label><input className="fi" type="password" placeholder="Mínimo 6 caracteres" value={nu.password} onChange={e=>setNu(p=>({...p,password:e.target.value}))}/></div>
                <div className="tog-row" onClick={()=>setNu(p=>({...p,isAdmin:!p.isAdmin}))}>
                  <div className={`tog ${nu.isAdmin?"on":""}`}/> Perfil administrador
                </div>
                {nuErr&&<div className="ferr" style={{marginBottom:"10px"}}>⚠ {nuErr}</div>}
                <button className="btn btn-p" onClick={addUser}>Adicionar</button>
              </div>
            </div>
          )}
        </main>
      </div>
      {active&&<DetailModal inc={active} onClose={()=>setActive(null)}/>}
    </>
  );
}

function IncCard({ inc, onClick }) {
  const ac=aColor(inc.category); const status=inc.status||"andamento";
  return(
    <div className="card" onClick={onClick}>
      <div className="card-top">
        <div className="card-client">{inc.client||"(sem cliente)"}</div>
        <div className="card-tags">
          <span className={`tag ${status==="resolvido"?"tag-r":"tag-a"}`}>{status==="resolvido"?"✅ Concluído":"🔄 Andamento"}</span>
          <span className="tag" style={{background:ac.bg,color:ac.c}}>{inc.category}</span>
        </div>
      </div>
      <div className="card-meta">
        {inc.date&&<span>📅 {fmt(inc.date)}</span>}
        {inc.am&&<span>👤 {inc.am}</span>}
        {inc.coord&&<span>👥 {inc.coord}</span>}
      </div>
      {inc.desc&&<div className="card-desc">{inc.desc}</div>}
    </div>
  );
}

function DetailModal({ inc, onClose }) {
  const ac=aColor(inc.category); const status=inc.status||"andamento";

  function attIcon(type) {
    if (type==="image") return "🖼️";
    if (type==="audio") return "🎵";
    if (type==="pdf")   return "📄";
    return "📎";
  }
  function attLabel(type) {
    if (type==="image") return "Imagem";
    if (type==="audio") return "Áudio";
    if (type==="pdf")   return "PDF";
    return "Arquivo";
  }

  return(
    <div className="ov" onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div className="modal">
        <div className="det-hd">
          <div style={{flex:1}}>
            <div className="det-client">{inc.client||"(sem cliente)"}</div>
            <div style={{display:"flex",gap:"6px",marginTop:"7px",flexWrap:"wrap"}}>
              <span className="tag" style={{background:ac.bg,color:ac.c}}>{inc.category}</span>
              <span className={`tag ${status==="resolvido"?"tag-r":"tag-a"}`}>{status==="resolvido"?"✅ Concluído":"🔄 Em andamento"}</span>
            </div>
          </div>
          <div className="det-x" onClick={onClose}>×</div>
        </div>
        <div className="det-pills">
          {inc.date&&<span className="det-pill">📅 {fmt(inc.date)}</span>}
          {inc.am&&<span className="det-pill">👤 {inc.am}</span>}
          {inc.coord&&<span className="det-pill">👥 {inc.coord}</span>}
          {inc.registeredBy&&<span className="det-pill">✍️ {inc.registeredBy}</span>}
        </div>
        {inc.desc&&(
          <div className="det-sec">
            <div className="det-sec-lbl">Descrição</div>
            <div className="det-desc">{inc.desc}</div>
          </div>
        )}
        {inc.attachments?.length>0&&(
          <div className="det-sec">
            <div className="det-sec-lbl">Anexos ({inc.attachments.length})</div>
            <div className="att-grid">
              {inc.attachments.map((att,i)=>(
                <a key={i} className="att-item" href={att.url} target="_blank" rel="noopener noreferrer" title={att.name}>
                  <div className="att-item-ic">{attIcon(att.type)}</div>
                  <div>
                    <div className="att-item-name">{att.name.length>22?att.name.slice(0,20)+"…":att.name}</div>
                    <div className="att-item-type">{attLabel(att.type)} · Abrir ↗</div>
                  </div>
                </a>
              ))}
            </div>
          </div>
        )}
        <div style={{display:"flex",justifyContent:"flex-end",paddingTop:"14px",borderTop:"1px solid var(--bd)"}}>
          <button className="btn btn-g" onClick={onClose}>Fechar</button>
        </div>
      </div>
    </div>
  );
}
