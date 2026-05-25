import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── JSONBIN ──────────────────────────────────────────────────────
const JB = "https://api.jsonbin.io/v3";
async function jbFetch(path, method="GET", body=null, key) {
  const r = await fetch(`${JB}${path}`, {
    method, headers:{"Content-Type":"application/json","X-Master-Key":key},
    body: body ? JSON.stringify(body) : null
  });
  if (!r.ok) { let m=`HTTP ${r.status}`; try{const j=await r.json();m=j.message||m;}catch{} throw new Error(m); }
  return r.json();
}
async function createBin(key, data) { const r=await jbFetch("/b","POST",data,key); return r.metadata.id; }
async function readBin(key, id)     { const r=await jbFetch(`/b/${id}/latest`,"GET",null,key); return r.record; }
async function writeBin(key, id, d) { await jbFetch(`/b/${id}`,"PUT",d,key); }

// ─── CONSTANTS ────────────────────────────────────────────────────
const DEFAULT_CATS = ["Logística / Estoque","CS","CSM"];
const MONTHS_FULL  = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
const CAT_COL = {
  "Logística / Estoque":{ c:"#1a5fd4", bg:"rgba(26,95,212,.12)" },
  "CS":                 { c:"#0d8a60", bg:"rgba(13,138,96,.12)"  },
  "CSM":                { c:"#7c3aed", bg:"rgba(124,58,237,.12)" },
};
const CAT_IC = {"Logística / Estoque":"📦","CS":"🎧","CSM":"🤝"};
const FIELD_TYPES = [{id:"text",l:"Texto"},{id:"number",l:"Número"},{id:"date",l:"Data"},{id:"select",l:"Lista"}];

function cc(cat)  { return CAT_COL[cat]||{c:"#e05cb8",bg:"rgba(224,92,184,.12)"}; }
function cic(cat) { return CAT_IC[cat]||"🏷️"; }
function fmt(d)   { return d?new Date(d+"T12:00:00").toLocaleDateString("pt-BR"):""; }
function toB64(f) { return new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result);r.onerror=rej;r.readAsDataURL(f);}); }
function simpleHash(s) { let h=0; for(let i=0;i<s.length;i++){h=((h<<5)-h)+s.charCodeAt(i);h|=0;} return h.toString(36); }

// ─── CSS ──────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Fraunces:opsz,wght@9..144,300;9..144,700&display=swap');
*{box-sizing:border-box;margin:0;padding:0;}
:root{
  --bg:#f5f7ff;--sf:#fff;--s2:#eef1fb;--s3:#e3e8f7;
  --bd:#d0d9f0;--ac:#2952cc;--ac2:#1a3fa3;--tx:#0d1b40;--mu:#7080a8;
  --err:#d63b3b;--ok:#0a8a5c;--warn:#c97a00;
  --fh:'Fraunces',serif;--fb:'Plus Jakarta Sans',sans-serif;
}
body{background:var(--bg);color:var(--tx);font-family:var(--fb);-webkit-font-smoothing:antialiased;}

/* ── SCREENS ── */
.screen{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;background:linear-gradient(135deg,#eef1fb 0%,#f5f7ff 60%,#ede8ff 100%);}
.card{background:var(--sf);border:1px solid var(--bd);border-radius:20px;padding:40px 36px;width:100%;max-width:440px;box-shadow:0 8px 40px rgba(41,82,204,.10);}
.card-logo{display:flex;align-items:center;gap:10px;margin-bottom:24px;}
.card-logo-ic{width:40px;height:40px;background:var(--ac);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:20px;}
.card-logo-txt{font-family:var(--fh);font-size:18px;font-weight:700;color:var(--tx);}
.card-logo-sub{font-size:11px;color:var(--mu);margin-top:1px;}
.card-title{font-family:var(--fh);font-size:22px;font-weight:700;margin-bottom:6px;}
.card-sub{font-size:13px;color:var(--mu);line-height:1.55;margin-bottom:24px;}
.steps{display:flex;flex-direction:column;gap:9px;margin-bottom:22px;}
.sst{display:flex;gap:11px;align-items:flex-start;padding:11px 13px;border-radius:10px;background:var(--s2);border:1px solid var(--bd);}
.sst-n{width:22px;height:22px;min-width:22px;border-radius:50%;background:var(--ac);color:#fff;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;margin-top:1px;}
.sst-t{font-size:12px;color:var(--tx);line-height:1.55;}
.sst-t a{color:var(--ac);font-weight:600;text-decoration:none;}
.sst-t strong{color:var(--ac);}
.fl{font-size:11px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.09em;display:block;margin-bottom:5px;margin-top:12px;}
.fi,.fs,.ft{width:100%;background:var(--s2);border:1.5px solid var(--bd);color:var(--tx);font-family:var(--fb);font-size:13px;padding:11px 13px;border-radius:9px;outline:none;transition:border-color .15s;}
.fi:focus,.fs:focus,.ft:focus{border-color:var(--ac);background:#fff;}
.ft{resize:vertical;min-height:80px;}
.fs option{background:var(--s2);}
.ferr{font-size:12px;color:var(--err);margin-top:8px;display:flex;align-items:center;gap:5px;}
.btn-full{width:100%;background:var(--ac);color:#fff;font-family:var(--fb);font-weight:700;font-size:14px;padding:13px;border:none;border-radius:10px;cursor:pointer;transition:background .15s;margin-top:14px;}
.btn-full:hover{background:var(--ac2);}
.btn-full:disabled{background:var(--mu);cursor:not-allowed;}
.note{font-size:11px;color:var(--mu);text-align:center;margin-top:10px;line-height:1.5;}
.divider-line{display:flex;align-items:center;gap:10px;margin:16px 0;color:var(--mu);font-size:11px;}
.divider-line::before,.divider-line::after{content:'';flex:1;height:1px;background:var(--bd);}

/* ── APP LAYOUT ── */
.app{display:flex;height:100vh;overflow:hidden;}
.sb{width:232px;min-width:232px;background:var(--sf);border-right:1px solid var(--bd);display:flex;flex-direction:column;}
.sb-brand{padding:20px 18px 16px;border-bottom:1px solid var(--bd);}
.sb-brand-row{display:flex;align-items:center;gap:9px;}
.sb-brand-ic{width:34px;height:34px;background:var(--ac);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:17px;flex-shrink:0;}
.sb-brand-name{font-family:var(--fh);font-size:15px;font-weight:700;color:var(--tx);}
.sb-brand-sub{font-size:10px;color:var(--mu);margin-top:1px;}
.sb-user{padding:10px 14px;background:var(--s2);margin:10px 12px;border-radius:9px;border:1px solid var(--bd);}
.sb-user-name{font-size:12px;font-weight:700;color:var(--tx);}
.sb-user-email{font-size:10px;color:var(--mu);margin-top:1px;}
.sb-user-role{font-size:9px;font-weight:700;color:var(--ac);text-transform:uppercase;letter-spacing:.08em;background:rgba(41,82,204,.1);padding:1px 6px;border-radius:4px;display:inline-block;margin-top:3px;}
.nav-sec{padding:10px 11px 4px;}
.nav-lbl{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.12em;padding:0 7px;margin-bottom:4px;}
.ni{display:flex;align-items:center;gap:8px;padding:8px 11px;border-radius:8px;cursor:pointer;font-size:13px;font-weight:500;color:var(--mu);transition:all .15s;position:relative;user-select:none;}
.ni:hover{background:var(--s2);color:var(--tx);}
.ni.on{background:var(--s2);color:var(--tx);font-weight:600;}
.ni.on::before{content:'';position:absolute;left:0;top:5px;bottom:5px;width:3px;border-radius:0 3px 3px 0;background:var(--ac);}
.ni-ic{font-size:13px;width:15px;text-align:center;}
.ni-badge{margin-left:auto;background:var(--sf);border:1px solid var(--bd);font-size:10px;font-weight:600;color:var(--mu);padding:1px 6px;border-radius:20px;}
.ni.on .ni-badge{background:var(--ac);color:#fff;border-color:var(--ac);}
.add-area{display:flex;align-items:center;gap:7px;padding:7px 11px;margin:2px 11px;border-radius:8px;font-size:12px;font-weight:500;color:var(--mu);cursor:pointer;border:1px dashed var(--bd);background:transparent;transition:all .15s;}
.add-area:hover{border-color:var(--ac);color:var(--ac);}
.sb-foot{margin-top:auto;padding:12px 14px;border-top:1px solid var(--bd);}
.sb-logout{font-size:12px;color:var(--mu);cursor:pointer;display:flex;align-items:center;gap:6px;}
.sb-logout:hover{color:var(--err);}

/* ── MAIN ── */
.main{flex:1;display:flex;flex-direction:column;overflow:hidden;}
.topbar{background:var(--sf);border-bottom:1px solid var(--bd);padding:13px 24px;display:flex;align-items:center;gap:9px;}
.topbar-title{font-family:var(--fh);font-size:18px;font-weight:700;flex:1;}
.topbar-title span{color:var(--ac);}
.sync{font-size:11px;padding:3px 10px;border-radius:20px;border:1px solid;display:flex;align-items:center;gap:4px;white-space:nowrap;}
.sync.ok{color:var(--ok);border-color:rgba(10,138,92,.25);background:rgba(10,138,92,.05);}
.sync.load{color:var(--mu);border-color:var(--bd);}
.sync.err{color:var(--err);border-color:rgba(214,59,59,.25);}
.btn{font-family:var(--fb);font-weight:600;font-size:13px;padding:8px 14px;border-radius:8px;cursor:pointer;display:flex;align-items:center;gap:5px;transition:all .15s;white-space:nowrap;border:none;}
.btn-p{background:var(--ac);color:#fff;}.btn-p:hover{background:var(--ac2);}
.btn-o{background:transparent;color:var(--mu);border:1px solid var(--bd)!important;}.btn-o:hover{border-color:var(--ok)!important;color:var(--ok);}
.btn-g{background:transparent;color:var(--mu);border:1px solid var(--bd)!important;}.btn-g:hover{color:var(--tx);border-color:var(--tx)!important;}
.btn-danger{background:transparent;color:var(--err);border:1px solid rgba(214,59,59,.3)!important;}.btn-danger:hover{background:var(--err);color:#fff;border-color:var(--err)!important;}

/* ── FILTERS ── */
.filters{padding:9px 24px;background:var(--sf);border-bottom:1px solid var(--bd);display:flex;gap:7px;align-items:center;flex-wrap:wrap;}
.f-lbl{font-size:11px;color:var(--mu);font-weight:600;}
.fsel{background:var(--s2);border:1px solid var(--bd);color:var(--tx);font-family:var(--fb);font-size:12px;padding:5px 9px;border-radius:7px;outline:none;}

/* ── CONTENT ── */
.content{flex:1;overflow-y:auto;padding:20px 24px;}
.stats-row{display:grid;grid-template-columns:repeat(3,1fr);gap:11px;margin-bottom:20px;}
.stat{background:var(--sf);border:1px solid var(--bd);border-radius:11px;padding:15px 17px;position:relative;overflow:hidden;}
.stat::after{content:'';position:absolute;right:-10px;bottom:-10px;width:55px;height:55px;border-radius:50%;background:var(--sc,var(--ac));opacity:.08;}
.stat-lbl{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.09em;}
.stat-num{font-family:var(--fh);font-size:28px;font-weight:700;margin:3px 0 2px;}
.stat-sub{font-size:11px;color:var(--mu);}
.sec-hd{display:flex;align-items:center;gap:9px;margin-bottom:11px;}
.sec-title{font-family:var(--fh);font-size:13px;font-weight:700;}
.sec-line{flex:1;height:1px;background:var(--bd);}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:11px;margin-bottom:22px;}

/* ── INCIDENT CARD ── */
.icard{background:var(--sf);border:1px solid var(--bd);border-radius:11px;padding:15px;transition:all .15s;position:relative;cursor:pointer;}
.icard:hover{border-color:var(--ac);box-shadow:0 3px 14px rgba(41,82,204,.08);transform:translateY(-1px);}
.icard-top{display:flex;align-items:flex-start;justify-content:space-between;gap:7px;margin-bottom:8px;}
.icard-name{font-weight:700;font-size:13px;}
.tags{display:flex;gap:4px;flex-wrap:wrap;justify-content:flex-end;flex-shrink:0;}
.tag{font-size:10px;font-weight:700;padding:2px 7px;border-radius:20px;text-transform:uppercase;letter-spacing:.05em;white-space:nowrap;}
.tag-a{background:rgba(201,122,0,.12);color:#a35e00;border:1px solid rgba(201,122,0,.25);}
.tag-r{background:rgba(10,138,92,.12);color:#0a7a52;border:1px solid rgba(10,138,92,.25);}
.icard-meta{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:6px;}
.icard-meta span{font-size:11px;color:var(--mu);}
.icard-desc{font-size:12px;color:var(--mu);line-height:1.5;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;}
.icard-chips{margin-top:6px;display:flex;flex-wrap:wrap;gap:4px;}
.chip{font-size:10px;background:var(--s2);border:1px solid var(--bd);border-radius:5px;padding:2px 6px;color:var(--mu);}
.chip strong{color:var(--tx);}
.icard-atts{display:flex;gap:4px;margin-top:8px;flex-wrap:wrap;}
.ath{width:30px;height:30px;border-radius:5px;background:var(--s2);border:1px solid var(--bd);display:flex;align-items:center;justify-content:center;font-size:12px;overflow:hidden;}
.ath img{width:100%;height:100%;object-fit:cover;border-radius:4px;}
.del-btn{position:absolute;top:10px;right:10px;display:none;width:24px;height:24px;border-radius:5px;border:1px solid var(--bd);background:var(--s2);cursor:pointer;align-items:center;justify-content:center;font-size:11px;}
.icard:hover .del-btn{display:flex;}
.del-btn:hover{background:var(--err);border-color:var(--err);color:#fff;}

/* ── KANBAN ── */
.kanban{display:flex;gap:13px;padding:20px 24px;overflow-x:auto;overflow-y:hidden;align-items:flex-start;flex:1;}
.kb-col{flex:0 0 300px;background:var(--s2);border:1px solid var(--bd);border-radius:11px;display:flex;flex-direction:column;max-height:calc(100vh - 172px);}
.kb-head{padding:12px 14px;border-bottom:1px solid var(--bd);display:flex;align-items:center;gap:7px;border-radius:11px 11px 0 0;}
.kb-title{font-family:var(--fh);font-size:13px;font-weight:700;flex:1;}
.kb-count{font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px;}
.kb-body{overflow-y:auto;padding:9px;display:flex;flex-direction:column;gap:7px;}
.kb-card{background:var(--sf);border:1px solid var(--bd);border-radius:8px;padding:12px;cursor:pointer;transition:all .15s;}
.kb-card:hover{border-color:var(--ac);transform:translateY(-1px);}
.kb-name{font-weight:700;font-size:12px;margin-bottom:5px;}
.kb-meta{display:flex;gap:5px;flex-wrap:wrap;margin-bottom:5px;}
.kb-meta span{font-size:11px;color:var(--mu);}
.kb-desc{font-size:11px;color:var(--mu);line-height:1.5;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;margin-bottom:6px;}
.kb-foot{display:flex;align-items:center;justify-content:space-between;}
.kb-mv{font-size:11px;font-weight:600;padding:3px 9px;border-radius:5px;border:1px solid var(--bd);background:var(--s2);color:var(--mu);cursor:pointer;transition:all .12s;}
.kb-mv:hover{border-color:var(--ac);color:var(--ac);}

/* ── EMPTY ── */
.empty{text-align:center;padding:44px 20px;color:var(--mu);}
.empty-ic{font-size:34px;margin-bottom:9px;}
.empty-t{font-size:13px;line-height:1.6;}

/* ── FORM VIEW ── */
.fv{flex:1;overflow-y:auto;display:flex;align-items:flex-start;justify-content:center;padding:28px 20px;}
.fv-card{background:var(--sf);border:1px solid var(--bd);border-radius:18px;width:100%;max-width:540px;padding:32px 28px;box-shadow:0 4px 20px rgba(41,82,204,.07);}
.fv-eyebrow{font-size:10px;font-weight:700;color:var(--ac);letter-spacing:.12em;text-transform:uppercase;margin-bottom:5px;}
.fv-title{font-family:var(--fh);font-size:22px;font-weight:700;margin-bottom:5px;}
.fv-sub{font-size:13px;color:var(--mu);line-height:1.5;margin-bottom:22px;}
.fg{margin-bottom:13px;}
.fg-lbl{font-size:11px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.09em;display:block;margin-bottom:5px;}
.fr2{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.status-pick{display:flex;gap:9px;}
.sp-opt{flex:1;padding:9px 11px;border-radius:8px;border:1.5px solid var(--bd);cursor:pointer;font-size:13px;font-weight:500;color:var(--mu);text-align:center;transition:all .15s;background:transparent;}
.sp-opt.on{border-color:var(--ac);background:rgba(41,82,204,.06);color:var(--ac);font-weight:700;}
.dropz{border:2px dashed var(--bd);border-radius:10px;padding:16px;text-align:center;cursor:pointer;transition:all .15s;color:var(--mu);font-size:13px;}
.dropz:hover,.dropz.drag{border-color:var(--ac);color:var(--ac);}
.dropz-sub{font-size:11px;margin-top:3px;}
.att-list{display:flex;gap:6px;flex-wrap:wrap;margin-top:9px;}
.att-item{display:flex;align-items:center;gap:5px;background:var(--s2);border:1px solid var(--bd);border-radius:6px;padding:4px 9px;font-size:11px;color:var(--mu);}
.att-rm{cursor:pointer;color:var(--err);}
.cblock{background:var(--s2);border:1px solid var(--bd);border-radius:9px;padding:12px 14px;margin-bottom:13px;}
.cb-title{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.1em;margin-bottom:10px;}
.fv-submit{width:100%;background:var(--ac);color:#fff;font-family:var(--fb);font-weight:700;font-size:14px;padding:13px;border:none;border-radius:10px;cursor:pointer;transition:background .15s;margin-top:6px;}
.fv-submit:hover{background:var(--ac2);}
.fv-submit:disabled{background:var(--mu);cursor:not-allowed;}
.success{display:flex;flex-direction:column;align-items:center;justify-content:center;padding:48px 20px;text-align:center;flex:1;}
.success-ic{font-size:52px;margin-bottom:14px;}
.success-title{font-family:var(--fh);font-size:24px;font-weight:700;margin-bottom:7px;}
.success-sub{font-size:13px;color:var(--mu);line-height:1.6;margin-bottom:24px;max-width:300px;}
.success-btns{display:flex;gap:9px;flex-wrap:wrap;justify-content:center;}

/* ── USERS PANEL ── */
.users-panel{flex:1;overflow-y:auto;padding:20px 24px;}
.users-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;}
.users-title{font-family:var(--fh);font-size:20px;font-weight:700;}
.users-table{background:var(--sf);border:1px solid var(--bd);border-radius:11px;overflow:hidden;}
.ut-head{display:grid;grid-template-columns:1fr 1fr 80px 80px;gap:0;background:var(--s2);border-bottom:1px solid var(--bd);padding:10px 16px;}
.ut-head span{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.09em;}
.ut-row{display:grid;grid-template-columns:1fr 1fr 80px 80px;gap:0;padding:13px 16px;border-bottom:1px solid var(--bd);align-items:center;transition:background .1s;}
.ut-row:last-child{border-bottom:none;}
.ut-row:hover{background:var(--s2);}
.ut-name{font-weight:600;font-size:13px;}
.ut-email{font-size:12px;color:var(--mu);}
.ut-role{font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;text-align:center;}
.ut-role.admin{background:rgba(41,82,204,.12);color:var(--ac);}
.ut-role.user{background:var(--s2);color:var(--mu);}
.ut-actions{display:flex;gap:6px;justify-content:flex-end;}
.icon-btn{width:28px;height:28px;border-radius:6px;border:1px solid var(--bd);background:transparent;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:13px;transition:all .12s;color:var(--mu);}
.icon-btn:hover{background:var(--err);border-color:var(--err);color:#fff;}
.add-user-form{background:var(--sf);border:1px solid var(--bd);border-radius:11px;padding:20px;margin-top:16px;}
.add-user-title{font-size:13px;font-weight:700;margin-bottom:14px;color:var(--tx);}
.add-user-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px;}
.add-user-grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:14px;}

/* ── MODAL ── */
.ov{position:fixed;inset:0;background:rgba(13,27,64,.65);z-index:100;display:flex;align-items:center;justify-content:center;padding:20px;backdrop-filter:blur(6px);}
.modal{background:var(--sf);border:1px solid var(--bd);border-radius:15px;width:100%;max-width:560px;max-height:92vh;overflow-y:auto;padding:24px;}
.modal.lg{max-width:620px;}
.modal-title{font-family:var(--fh);font-size:17px;font-weight:700;margin-bottom:18px;}
.modal-title span{color:var(--ac);}
.mdivider{height:1px;background:var(--bd);margin:14px 0;}
.field-list{display:flex;flex-direction:column;gap:7px;margin-bottom:12px;}
.fl-row{display:flex;align-items:center;gap:8px;background:var(--s2);border:1px solid var(--bd);border-radius:8px;padding:9px 12px;}
.fl-name{flex:1;font-size:13px;font-weight:500;}
.badge{font-size:10px;padding:2px 7px;border-radius:5px;background:var(--s3);color:var(--mu);}
.badge.red{color:var(--err);}
.fdel{cursor:pointer;color:var(--mu);font-size:16px;}
.fdel:hover{color:var(--err);}
.tpills{display:flex;gap:5px;flex-wrap:wrap;}
.tpill{font-size:12px;font-weight:600;padding:5px 13px;border-radius:20px;cursor:pointer;border:1.5px solid var(--bd);color:var(--mu);background:transparent;transition:all .15s;}
.tpill.on{background:var(--ac);color:#fff;border-color:var(--ac);}
.tog-row{display:flex;align-items:center;gap:7px;font-size:12px;color:var(--mu);cursor:pointer;margin-top:8px;user-select:none;}
.tog{width:28px;height:16px;border-radius:9px;background:var(--bd);position:relative;transition:background .15s;flex-shrink:0;}
.tog.on{background:var(--ac);}
.tog::after{content:'';position:absolute;width:10px;height:10px;border-radius:50%;background:white;top:3px;left:3px;transition:left .15s;}
.tog.on::after{left:15px;}
.det-hd{display:flex;align-items:flex-start;gap:10px;margin-bottom:16px;}
.det-name{font-family:var(--fh);font-size:20px;font-weight:700;flex:1;}
.det-x{width:28px;height:28px;border-radius:7px;border:1px solid var(--bd);background:var(--s2);cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:14px;color:var(--mu);flex-shrink:0;}
.det-x:hover{background:var(--err);border-color:var(--err);color:white;}
.det-pills{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:14px;}
.det-pill{display:flex;align-items:center;gap:4px;background:var(--s2);border:1px solid var(--bd);border-radius:20px;padding:4px 10px;font-size:12px;color:var(--mu);}
.det-sec{margin-bottom:14px;}
.det-sec-lbl{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.1em;margin-bottom:6px;}
.det-desc{background:var(--s2);border:1px solid var(--bd);border-radius:8px;padding:12px 14px;font-size:13px;line-height:1.7;color:var(--tx);white-space:pre-wrap;}
.det-att-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(100px,1fr));gap:7px;}
.det-att{border:1px solid var(--bd);border-radius:8px;overflow:hidden;background:var(--s2);cursor:pointer;transition:border-color .15s;}
.det-att:hover{border-color:var(--ac);}
.det-att img{width:100%;aspect-ratio:1;object-fit:cover;display:block;}
.det-att-f{padding:11px 8px;text-align:center;}
.det-att-ic{font-size:24px;margin-bottom:4px;}
.det-att-nm{font-size:10px;color:var(--mu);word-break:break-all;line-height:1.3;}
.det-cg{display:grid;grid-template-columns:1fr 1fr;gap:6px;}
.det-ci{background:var(--s2);border:1px solid var(--bd);border-radius:7px;padding:8px 10px;}
.det-ci-l{font-size:10px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:.07em;margin-bottom:2px;}
.det-ci-v{font-size:13px;font-weight:600;color:var(--tx);}
.lightbox{position:fixed;inset:0;background:rgba(0,0,0,.93);z-index:200;display:flex;align-items:center;justify-content:center;cursor:zoom-out;}
.lightbox img{max-width:90vw;max-height:90vh;border-radius:8px;object-fit:contain;}
::-webkit-scrollbar{width:5px;}
::-webkit-scrollbar-track{background:transparent;}
::-webkit-scrollbar-thumb{background:var(--bd);border-radius:3px;}
`;

// ─── INITIAL DATA ─────────────────────────────────────────────────
const INITIAL_DATA = (adminName, adminEmail, adminPass) => ({
  incidents: [], cats: [...DEFAULT_CATS], cFields: [],
  users: [{
    id: "u_" + Date.now(),
    name: adminName,
    email: adminEmail.toLowerCase(),
    passwordHash: simpleHash(adminPass),
    isAdmin: true,
    createdAt: new Date().toISOString()
  }]
});

// ─── APP ──────────────────────────────────────────────────────────
export default function App() {
  const [apiKey,  setApiKey]  = useState(() => localStorage.getItem("jb_key")||"");
  const [binId,   setBinId]   = useState(() => localStorage.getItem("jb_bin")||"");
  const [session, setSession] = useState(() => { try { return JSON.parse(localStorage.getItem("jb_session")||"null"); } catch { return null; } });

  const [data, setData] = useState({ incidents:[], cats:[...DEFAULT_CATS], cFields:[], users:[] });
  const [sync, setSync] = useState("ok");

  // Setup form
  const [setupKey,   setSetupKey]   = useState("");
  const [setupName,  setSetupName]  = useState("");
  const [setupEmail, setSetupEmail] = useState("");
  const [setupPass,  setSetupPass]  = useState("");
  const [setupErr,   setSetupErr]   = useState("");
  const [setupLoad,  setSetupLoad]  = useState(false);

  // Login form
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPass,  setLoginPass]  = useState("");
  const [loginErr,   setLoginErr]   = useState("");

  // App state
  const [view,   setView]   = useState("form");
  const [tab,    setTab]    = useState("all");
  const [modal,  setModal]  = useState(null);
  const [active, setActive] = useState(null);
  const [filterAM,     setFilterAM]     = useState("");
  const [filterCat,    setFilterCat]    = useState("");
  const [filterCoord,  setFilterCoord]  = useState("");
  const [filterMonth,  setFilterMonth]  = useState("");
  const [filterStatus, setFilterStatus] = useState("");
  const [submitted, setSubmitted] = useState(false);
  const [formDrag,  setFormDrag]  = useState(false);
  const [newCat,    setNewCat]    = useState("");
  const [nf, setNf] = useState({ label:"", type:"text", options:"", required:false });

  // Add user form
  const [nu, setNu] = useState({ name:"", email:"", password:"", isAdmin:false });
  const [nuErr, setNuErr] = useState("");

  const fileRef = useRef();

  const blankForm = useCallback(() => ({
    client:"", date:new Date().toISOString().split("T")[0],
    am: session?.name||"", coord:"", category:data.cats[0]||DEFAULT_CATS[0],
    status:"andamento", desc:"", attachments:[], custom:{}
  }), [data.cats, session]);
  const [form, setForm] = useState(blankForm());

  const connected = !!(apiKey && binId);
  const loggedIn  = !!(connected && session);

  // ── Load data ──
  const load = useCallback(async () => {
    if (!apiKey || !binId) return;
    setSync("load");
    try {
      const d = await readBin(apiKey, binId);
      setData({ incidents:d.incidents||[], cats:d.cats||[...DEFAULT_CATS], cFields:d.cFields||[], users:d.users||[] });
      setSync("ok");
    } catch { setSync("err"); }
  }, [apiKey, binId]);

  useEffect(() => {
    if (!loggedIn) return;
    load();
    const iv = setInterval(load, 20000);
    return () => clearInterval(iv);
  }, [loggedIn, load]);

  const persist = async (next) => {
    setSync("load");
    try { await writeBin(apiKey, binId, next); setSync("ok"); }
    catch { setSync("err"); }
  };

  // ── Setup (first time) ──
  const handleSetup = async () => {
    if (!setupKey.trim() || !setupName.trim() || !setupEmail.trim() || !setupPass.trim()) {
      setSetupErr("Preencha todos os campos."); return;
    }
    if (setupPass.length < 6) { setSetupErr("A senha deve ter ao menos 6 caracteres."); return; }
    setSetupLoad(true); setSetupErr("");
    try {
      const init = INITIAL_DATA(setupName.trim(), setupEmail.trim(), setupPass.trim());
      const id = await createBin(setupKey.trim(), init);
      localStorage.setItem("jb_key", setupKey.trim());
      localStorage.setItem("jb_bin", id);
      const sess = { id: init.users[0].id, name: init.users[0].name, email: init.users[0].email, isAdmin: true };
      localStorage.setItem("jb_session", JSON.stringify(sess));
      setApiKey(setupKey.trim()); setBinId(id); setSession(sess); setData(init);
    } catch (e) { setSetupErr("Erro: " + (e.message||"desconhecido")); }
    setSetupLoad(false);
  };

  // ── Login ──
  const handleLogin = () => {
    const user = data.users.find(u => u.email === loginEmail.toLowerCase().trim() && u.passwordHash === simpleHash(loginPass));
    if (!user) { setLoginErr("Email ou senha incorretos."); return; }
    const sess = { id:user.id, name:user.name, email:user.email, isAdmin:user.isAdmin };
    localStorage.setItem("jb_session", JSON.stringify(sess));
    setSession(sess); setLoginErr("");
  };

  const handleLogout = () => {
    localStorage.removeItem("jb_session");
    setSession(null); setLoginEmail(""); setLoginPass(""); setLoginErr("");
  };

  // ── CRUD ──
  const addFiles = async (files) => {
    const arr = [];
    for (const f of files) arr.push({ name:f.name, type:f.type, data:await toB64(f) });
    setForm(p=>({...p, attachments:[...p.attachments,...arr]}));
  };

  const submitForm = async () => {
    if (!form.client.trim()) return;
    const inc = { id:Date.now(), ...form, registeredBy: session?.name||"" };
    const next = { ...data, incidents:[inc,...data.incidents] };
    setData(next); await persist(next); setSubmitted(true);
  };

  const del = async (id) => {
    const next = { ...data, incidents:data.incidents.filter(i=>i.id!==id) };
    setData(next); await persist(next);
  };

  const setStatus = async (id, status) => {
    const next = { ...data, incidents:data.incidents.map(i=>i.id===id?{...i,status}:i) };
    setData(next); await persist(next);
    if (active?.id===id) setActive(p=>({...p,status}));
  };

  const addCat = async () => {
    if (!newCat.trim()) return;
    const next = { ...data, cats:[...data.cats,newCat.trim()] };
    setData(next); await persist(next); setNewCat(""); setModal(null);
  };

  const addField = async () => {
    if (!nf.label.trim()) return;
    const field = { id:"cf_"+Date.now(), label:nf.label.trim(), type:nf.type,
      options:nf.type==="select"?nf.options.split(",").map(s=>s.trim()).filter(Boolean):[],
      required:nf.required };
    const next = { ...data, cFields:[...data.cFields,field] };
    setData(next); await persist(next);
    setNf({ label:"", type:"text", options:"", required:false });
  };

  const delField = async (id) => {
    const next = { ...data, cFields:data.cFields.filter(f=>f.id!==id) };
    setData(next); await persist(next);
  };

  // ── User management ──
  const addUser = async () => {
    if (!nu.name.trim()||!nu.email.trim()||!nu.password.trim()) { setNuErr("Preencha todos os campos."); return; }
    if (nu.password.length < 6) { setNuErr("Senha mínima de 6 caracteres."); return; }
    if (data.users.find(u=>u.email===nu.email.toLowerCase().trim())) { setNuErr("Email já cadastrado."); return; }
    const user = { id:"u_"+Date.now(), name:nu.name.trim(), email:nu.email.toLowerCase().trim(),
      passwordHash:simpleHash(nu.password), isAdmin:nu.isAdmin, createdAt:new Date().toISOString() };
    const next = { ...data, users:[...data.users,user] };
    setData(next); await persist(next);
    setNu({ name:"", email:"", password:"", isAdmin:false }); setNuErr("");
  };

  const removeUser = async (id) => {
    if (id===session?.id) { alert("Você não pode remover sua própria conta."); return; }
    if (!confirm("Remover este usuário?")) return;
    const next = { ...data, users:data.users.filter(u=>u.id!==id) };
    setData(next); await persist(next);
  };

  const resetPassword = async (id, newPass) => {
    if (!newPass || newPass.length < 6) return;
    const next = { ...data, users:data.users.map(u=>u.id===id?{...u,passwordHash:simpleHash(newPass)}:u) };
    setData(next); await persist(next);
  };

  const exportXlsx = () => {
    const rows = filtered.map(i=>({
      "Cliente":i.client,"Data":i.date,"AM":i.am,"Coordenação":i.coord,"Área":i.category,
      "Status":i.status==="resolvido"?"Resolvido":"Em andamento","Descrição":i.desc,
      "Registrado por":i.registeredBy||"","Anexos":(i.attachments||[]).map(a=>a.name).join(", ")
    }));
    data.cFields.forEach((cf,k)=>rows.forEach((r,j)=>{ r[cf.label]=filtered[j].custom?.[cf.id]??""; }));
    const ws=XLSX.utils.json_to_sheet(rows);
    ws["!cols"]=Object.keys(rows[0]||{}).map(()=>({wch:20}));
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Ocorrências");
    XLSX.writeFile(wb,`ocorrencias_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  const allAMs    = [...new Set(data.incidents.map(i=>i.am).filter(Boolean))];
  const allCoords = [...new Set(data.incidents.map(i=>i.coord).filter(Boolean))];
  const anyFilter = filterAM||filterCat||filterCoord||filterMonth||filterStatus;

  const filtered = data.incidents.filter(i=>{
    if (tab!=="all" && i.category!==tab) return false;
    if (filterAM     && i.am!==filterAM)                            return false;
    if (filterCat    && i.category!==filterCat)                     return false;
    if (filterCoord  && i.coord!==filterCoord)                      return false;
    if (filterStatus && (i.status||"andamento")!==filterStatus)     return false;
    if (filterMonth!==""){const m=i.date?new Date(i.date+"T12:00:00").getMonth():-1;if(m!==parseInt(filterMonth))return false;}
    return true;
  });

  // ── SETUP SCREEN ──
  if (!connected) return (
    <>
      <style>{CSS}</style>
      <div className="screen">
        <div className="card">
          <div className="card-logo">
            <div className="card-logo-ic">📋</div>
            <div><div className="card-logo-txt">Ocorrências ZIG</div><div className="card-logo-sub">Configuração inicial</div></div>
          </div>
          <div className="card-title">Criar banco de dados</div>
          <div className="card-sub">Configure o sistema uma vez. Depois compartilhe o link com o time — cada pessoa faz seu próprio login.</div>
          <div className="steps">
            <div className="sst"><div className="sst-n">1</div><div className="sst-t">Acesse <a href="https://jsonbin.io" target="_blank">jsonbin.io</a>, crie conta gratuita e vá em <strong>API Keys → Create API Key</strong>.</div></div>
            <div className="sst"><div className="sst-n">2</div><div className="sst-t">Preencha os campos abaixo com a API Key e seus dados de admin. Você poderá cadastrar o time depois.</div></div>
          </div>
          <label className="fl">API Key do JSONBin</label>
          <input className="fi" placeholder="$2a$10$..." value={setupKey} onChange={e=>setSetupKey(e.target.value)}/>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"10px",marginTop:"12px"}}>
            <div><label className="fl">Seu nome</label><input className="fi" placeholder="Ex: Talita" value={setupName} onChange={e=>setSetupName(e.target.value)}/></div>
            <div><label className="fl">Seu email</label><input className="fi" placeholder="email@zig.fun" value={setupEmail} onChange={e=>setSetupEmail(e.target.value)}/></div>
          </div>
          <label className="fl">Sua senha (mínimo 6 caracteres)</label>
          <input className="fi" type="password" placeholder="••••••" value={setupPass} onChange={e=>setSetupPass(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleSetup()}/>
          {setupErr && <div className="ferr">⚠ {setupErr}</div>}
          <button className="btn-full" onClick={handleSetup} disabled={setupLoad}>{setupLoad?"Criando...":"Criar e entrar →"}</button>
          <div className="note">Você será o administrador. Poderá cadastrar o time depois.</div>
        </div>
      </div>
    </>
  );

  // ── LOGIN SCREEN ──
  if (!loggedIn) return (
    <>
      <style>{CSS}</style>
      <div className="screen">
        <div className="card">
          <div className="card-logo">
            <div className="card-logo-ic">📋</div>
            <div><div className="card-logo-txt">Ocorrências ZIG</div><div className="card-logo-sub">Central de registros</div></div>
          </div>
          <div className="card-title">Entrar</div>
          <div className="card-sub">Use seu email e senha cadastrados pelo administrador.</div>
          <label className="fl">Email</label>
          <input className="fi" type="email" placeholder="seu@email.com" value={loginEmail} onChange={e=>setLoginEmail(e.target.value)}/>
          <label className="fl">Senha</label>
          <input className="fi" type="password" placeholder="••••••" value={loginPass} onChange={e=>setLoginPass(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()}/>
          {loginErr && <div className="ferr">⚠ {loginErr}</div>}
          <button className="btn-full" onClick={handleLogin}>Entrar →</button>
          <div className="note">Não tem acesso? Fale com o administrador do sistema.</div>
        </div>
      </div>
    </>
  );

  // ── MAIN APP ──
  return (
    <>
      <style>{CSS}</style>
      <div className="app">
        <aside className="sb">
          <div className="sb-brand">
            <div className="sb-brand-row">
              <div className="sb-brand-ic">📋</div>
              <div><div className="sb-brand-name">Ocorrências</div><div className="sb-brand-sub">ZIG Funtech</div></div>
            </div>
          </div>
          <div style={{padding:"10px 12px 0"}}>
            <div className="sb-user">
              <div className="sb-user-name">{session.name}</div>
              <div className="sb-user-email">{session.email}</div>
              <span className="sb-user-role">{session.isAdmin?"Admin":"Usuário"}</span>
            </div>
          </div>
          <div className="nav-sec">
            <div className="nav-lbl">Menu</div>
            <div className={`ni ${view==="form"?"on":""}`} onClick={()=>{setView("form");setSubmitted(false);}}>
              <span className="ni-ic">📝</span> Registrar
            </div>
            <div className={`ni ${view==="list"&&tab==="all"?"on":""}`} onClick={()=>{setView("list");setTab("all");}}>
              <span className="ni-ic">⊞</span> Dashboard
              <span className="ni-badge">{data.incidents.length}</span>
            </div>
            <div className={`ni ${view==="kanban"?"on":""}`} onClick={()=>setView("kanban")}>
              <span className="ni-ic">🗂</span> Kanban
              <span className="ni-badge">{data.incidents.filter(i=>(i.status||"andamento")==="andamento").length}</span>
            </div>
            {session.isAdmin && (
              <div className={`ni ${view==="users"?"on":""}`} onClick={()=>setView("users")}>
                <span className="ni-ic">👥</span> Usuários
                <span className="ni-badge">{data.users.length}</span>
              </div>
            )}
          </div>
          {data.cats.length>0 && (
            <div className="nav-sec">
              <div className="nav-lbl">Áreas</div>
              {data.cats.map(c=>(
                <div key={c} className={`ni ${view==="list"&&tab===c?"on":""}`} onClick={()=>{setView("list");setTab(c);}}>
                  <span className="ni-ic">{cic(c)}</span>
                  <span style={{fontSize:"12px",flex:1}}>{c}</span>
                  <span className="ni-badge">{data.incidents.filter(i=>i.category===c).length}</span>
                </div>
              ))}
              {session.isAdmin && <div className="add-area" onClick={()=>setModal("cat")}>＋ Nova área</div>}
            </div>
          )}
          <div className="sb-foot">
            <div className="sb-logout" onClick={handleLogout}>🚪 Sair</div>
          </div>
        </aside>

        <main className="main">
          <div className="topbar">
            <div className="topbar-title">
              {view==="form"   ? <>Nova <span>Ocorrência</span></>
              :view==="kanban" ? <>Kanban de <span>Status</span></>
              :view==="users"  ? <><span>Usuários</span> do sistema</>
              :tab==="all"     ? <>Todas as <span>Ocorrências</span></>
              :                  <>{cic(tab)} {tab}</>}
            </div>
            <div className={`sync ${sync}`}>
              {sync==="ok"&&"● Sincronizado"}{sync==="load"&&"○ Atualizando..."}{sync==="err"&&"● Erro"}
            </div>
            {view!=="form"&&view!=="users"&&session.isAdmin&&<>
              <button className="btn btn-o" onClick={()=>setModal("fields")}>⚙ Campos</button>
              {data.incidents.length>0&&<button className="btn btn-o" onClick={exportXlsx}>↓ Excel</button>}
            </>}
          </div>

          {view!=="form"&&view!=="users"&&(
            <div className="filters">
              <span className="f-lbl">Filtrar:</span>
              <select className="fsel" value={filterAM} onChange={e=>setFilterAM(e.target.value)}>
                <option value="">Todos os AMs</option>{allAMs.map(a=><option key={a}>{a}</option>)}
              </select>
              <select className="fsel" value={filterCoord} onChange={e=>setFilterCoord(e.target.value)}>
                <option value="">Todas as coordenações</option>{allCoords.map(c=><option key={c}>{c}</option>)}
              </select>
              {tab==="all"&&view==="list"&&<select className="fsel" value={filterCat} onChange={e=>setFilterCat(e.target.value)}>
                <option value="">Todas as áreas</option>{data.cats.map(c=><option key={c}>{c}</option>)}
              </select>}
              <select className="fsel" value={filterMonth} onChange={e=>setFilterMonth(e.target.value)}>
                <option value="">Todos os meses</option>{MONTHS_FULL.map((m,i)=><option key={i} value={i}>{m}</option>)}
              </select>
              <select className="fsel" value={filterStatus} onChange={e=>setFilterStatus(e.target.value)}>
                <option value="">Todos os status</option>
                <option value="andamento">Em andamento</option>
                <option value="resolvido">Resolvido</option>
              </select>
              {anyFilter&&<span style={{fontSize:"11px",color:"var(--ac)",cursor:"pointer",fontWeight:600}} onClick={()=>{setFilterAM("");setFilterCat("");setFilterCoord("");setFilterMonth("");setFilterStatus("");}}>✕ Limpar</span>}
              <span style={{marginLeft:"auto",fontSize:"11px",color:"var(--mu)"}}>{filtered.length} registro{filtered.length!==1?"s":""}</span>
            </div>
          )}

          {/* ── FORM VIEW ── */}
          {view==="form" && (submitted ? (
            <div className="success">
              <div className="success-ic">✅</div>
              <div className="success-title">Registrado!</div>
              <div className="success-sub">A ocorrência foi salva e já está visível para todo o time.</div>
              <div className="success-btns">
                <button className="btn btn-p" onClick={()=>{setForm(blankForm());setSubmitted(false);}}>＋ Nova ocorrência</button>
                <button className="btn btn-g" style={{border:"1px solid var(--bd)"}} onClick={()=>{setView("list");setTab("all");}}>Ver dashboard →</button>
              </div>
            </div>
          ) : (
            <div className="fv">
              <div className="fv-card">
                <div className="fv-eyebrow">Nova ocorrência</div>
                <div className="fv-title">Registrar problema</div>
                <div className="fv-sub">Preencha os campos e clique em registrar.</div>
                <div className="fr2">
                  <div className="fg"><label className="fg-lbl">Cliente *</label><input className="fi" placeholder="Nome do cliente" value={form.client} onChange={e=>setForm(p=>({...p,client:e.target.value}))}/></div>
                  <div className="fg"><label className="fg-lbl">Data *</label><input className="fi" type="date" value={form.date} onChange={e=>setForm(p=>({...p,date:e.target.value}))}/></div>
                </div>
                <div className="fr2">
                  <div className="fg"><label className="fg-lbl">AM Responsável</label><input className="fi" placeholder="Nome do AM" value={form.am} onChange={e=>setForm(p=>({...p,am:e.target.value}))}/></div>
                  <div className="fg"><label className="fg-lbl">Coordenação</label><input className="fi" placeholder="Coordenador" value={form.coord} onChange={e=>setForm(p=>({...p,coord:e.target.value}))}/></div>
                </div>
                <div className="fg"><label className="fg-lbl">Área *</label>
                  <select className="fs" value={form.category} onChange={e=>setForm(p=>({...p,category:e.target.value}))}>
                    {data.cats.map(c=><option key={c}>{c}</option>)}
                  </select>
                </div>
                <div className="fg"><label className="fg-lbl">Status</label>
                  <div className="status-pick">
                    {[{v:"andamento",l:"🔄 Em andamento"},{v:"resolvido",l:"✅ Resolvido"}].map(o=>(
                      <div key={o.v} className={`sp-opt ${form.status===o.v?"on":""}`} onClick={()=>setForm(p=>({...p,status:o.v}))}>{o.l}</div>
                    ))}
                  </div>
                </div>
                <div className="fg"><label className="fg-lbl">Descrição</label><textarea className="ft" placeholder="Descreva o que aconteceu..." value={form.desc} onChange={e=>setForm(p=>({...p,desc:e.target.value}))}/></div>
                {data.cFields.length>0&&(
                  <div className="cblock"><div className="cb-title">Campos adicionais</div>
                    {data.cFields.map(cf=>(
                      <div className="fg" key={cf.id} style={{marginBottom:"9px"}}>
                        <label className="fg-lbl">{cf.label}{cf.required?" *":""}</label>
                        {cf.type==="text"&&<input className="fi" value={form.custom?.[cf.id]||""} onChange={e=>setForm(p=>({...p,custom:{...p.custom,[cf.id]:e.target.value}}))}/>}
                        {cf.type==="number"&&<input className="fi" type="number" value={form.custom?.[cf.id]||""} onChange={e=>setForm(p=>({...p,custom:{...p.custom,[cf.id]:e.target.value}}))}/>}
                        {cf.type==="date"&&<input className="fi" type="date" value={form.custom?.[cf.id]||""} onChange={e=>setForm(p=>({...p,custom:{...p.custom,[cf.id]:e.target.value}}))}/>}
                        {cf.type==="select"&&<select className="fs" value={form.custom?.[cf.id]||""} onChange={e=>setForm(p=>({...p,custom:{...p.custom,[cf.id]:e.target.value}}))}>
                          <option value="">Selecione...</option>{cf.options.map(o=><option key={o}>{o}</option>)}
                        </select>}
                      </div>
                    ))}
                  </div>
                )}
                <div className="fg"><label className="fg-lbl">Anexos</label>
                  <div className={`dropz ${formDrag?"drag":""}`} onClick={()=>fileRef.current.click()}
                    onDragOver={e=>{e.preventDefault();setFormDrag(true);}} onDragLeave={()=>setFormDrag(false)}
                    onDrop={e=>{e.preventDefault();setFormDrag(false);addFiles([...e.dataTransfer.files]);}}>
                    📎 Clique ou arraste arquivos<div className="dropz-sub">Imagens, PDFs, áudios de WhatsApp</div>
                  </div>
                  <input ref={fileRef} type="file" multiple style={{display:"none"}} accept="image/*,.pdf,.opus,.ogg,.mp3,.m4a,.aac,.wav" onChange={e=>addFiles([...e.target.files])}/>
                  {form.attachments.length>0&&<div className="att-list">{form.attachments.map((a,i)=>(
                    <div key={i} className="att-item">{a.name.length>18?a.name.slice(0,16)+"…":a.name}
                      <span className="att-rm" onClick={()=>setForm(p=>({...p,attachments:p.attachments.filter((_,j)=>j!==i)}))}>×</span>
                    </div>
                  ))}</div>}
                </div>
                <button className="fv-submit" onClick={submitForm} disabled={!form.client.trim()}>Registrar ocorrência →</button>
              </div>
            </div>
          ))}

          {/* ── USERS PANEL (admin only) ── */}
          {view==="users" && (
            <div className="users-panel">
              <div className="users-header">
                <div className="users-title">Gerenciar usuários</div>
              </div>
              <div className="users-table">
                <div className="ut-head">
                  <span>Nome</span><span>Email</span><span>Perfil</span><span></span>
                </div>
                {data.users.map(u=>(
                  <div key={u.id} className="ut-row">
                    <div><div className="ut-name">{u.name}{u.id===session.id?" (você)":""}</div></div>
                    <div className="ut-email">{u.email}</div>
                    <div><span className={`ut-role ${u.isAdmin?"admin":"user"}`}>{u.isAdmin?"Admin":"Usuário"}</span></div>
                    <div className="ut-actions">
                      {u.id!==session.id && (
                        <div className="icon-btn" title="Remover usuário" onClick={()=>removeUser(u.id)}>🗑</div>
                      )}
                    </div>
                  </div>
                ))}
              </div>

              <div className="add-user-form">
                <div className="add-user-title">＋ Adicionar usuário</div>
                <div className="add-user-grid">
                  <div><label className="fg-lbl">Nome</label><input className="fi" placeholder="Nome completo" value={nu.name} onChange={e=>setNu(p=>({...p,name:e.target.value}))}/></div>
                  <div><label className="fg-lbl">Email</label><input className="fi" placeholder="email@zig.fun" value={nu.email} onChange={e=>setNu(p=>({...p,email:e.target.value}))}/></div>
                </div>
                <div className="add-user-grid">
                  <div><label className="fg-lbl">Senha inicial</label><input className="fi" type="password" placeholder="Mínimo 6 caracteres" value={nu.password} onChange={e=>setNu(p=>({...p,password:e.target.value}))}/></div>
                  <div style={{display:"flex",alignItems:"flex-end",paddingBottom:"0"}}>
                    <div className="tog-row" style={{marginTop:"22px"}} onClick={()=>setNu(p=>({...p,isAdmin:!p.isAdmin}))}>
                      <div className={`tog ${nu.isAdmin?"on":""}`}/> Perfil administrador
                    </div>
                  </div>
                </div>
                {nuErr&&<div className="ferr" style={{marginBottom:"10px"}}>⚠ {nuErr}</div>}
                <button className="btn btn-p" onClick={addUser}>Adicionar usuário</button>
              </div>
            </div>
          )}

          {/* ── KANBAN ── */}
          {view==="kanban" && (
            <div className="kanban">
              {[
                {key:"andamento",label:"Em andamento 🔄",color:"#a35e00",bg:"rgba(201,122,0,.12)",bd:"rgba(201,122,0,.28)"},
                {key:"resolvido",label:"Resolvido ✅",color:"#0a7a52",bg:"rgba(10,138,92,.12)",bd:"rgba(10,138,92,.28)"},
              ].map(col=>{
                const ci2=filtered.filter(i=>(i.status||"andamento")===col.key);
                return (
                  <div className="kb-col" key={col.key}>
                    <div className="kb-head">
                      <div className="kb-title" style={{color:col.color}}>{col.label}</div>
                      <span className="kb-count" style={{background:col.bg,color:col.color,border:`1px solid ${col.bd}`}}>{ci2.length}</span>
                    </div>
                    <div className="kb-body">
                      {ci2.length===0&&<div style={{textAlign:"center",padding:"20px 0",fontSize:"12px",color:"var(--mu)"}}>Nenhuma ocorrência</div>}
                      {ci2.map(inc=>{
                        const cst=cc(inc.category);
                        const next=col.key==="andamento"?"resolvido":"andamento";
                        return (
                          <div className="kb-card" key={inc.id} onClick={()=>{setActive(inc);setModal("detail");}}>
                            <div className="kb-name">{inc.client}</div>
                            <div className="kb-meta">
                              <span className="tag" style={{background:cst.bg,color:cst.c}}>{inc.category}</span>
                              {inc.date&&<span>📅 {fmt(inc.date)}</span>}
                              {inc.am&&<span>👤 {inc.am}</span>}
                            </div>
                            {inc.desc&&<div className="kb-desc">{inc.desc}</div>}
                            <div className="kb-foot">
                              <button className="kb-mv" onClick={e=>{e.stopPropagation();setStatus(inc.id,next);}}>
                                {col.key==="andamento"?"✅ Marcar resolvido":"🔄 Reabrir"}
                              </button>
                              {inc.attachments?.length>0&&<span style={{fontSize:"11px",color:"var(--mu)"}}>📎{inc.attachments.length}</span>}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* ── LIST ── */}
          {view==="list" && (
            <div className="content">
              {tab==="all"&&data.cats.length>0&&(
                <div className="stats-row">
                  {data.cats.slice(0,3).map(c=>{const cst=cc(c);return(
                    <div className="stat" key={c} style={{"--sc":cst.c}}>
                      <div className="stat-lbl">{cic(c)} {c}</div>
                      <div className="stat-num" style={{color:cst.c}}>{data.incidents.filter(i=>i.category===c).length}</div>
                      <div className="stat-sub">ocorrências</div>
                    </div>
                  );})}
                </div>
              )}
              {tab==="all"?data.cats.map(c=>{
                const ci2=filtered.filter(i=>i.category===c);
                if(!ci2.length&&!anyFilter) return null;
                return(
                  <div key={c}>
                    <div className="sec-hd"><span className="sec-title">{cic(c)} {c}</span><div className="sec-line"/><span style={{fontSize:"11px",color:"var(--mu)"}}>{ci2.length}</span></div>
                    {ci2.length===0?<p style={{fontSize:"12px",color:"var(--mu)",marginBottom:"20px"}}>Nenhuma ocorrência.</p>
                    :<div className="grid">{ci2.map(i=><ICard key={i.id} inc={i} cFields={data.cFields} onDel={session.isAdmin?del:null} onView={inc=>{setActive(inc);setModal("detail");}}/>)}</div>}
                  </div>
                );
              }):(
                filtered.length===0
                  ?<div className="empty"><div className="empty-ic">{cic(tab)}</div><div className="empty-t">Nenhuma ocorrência em <strong>{tab}</strong>.</div></div>
                  :<div className="grid">{filtered.map(i=><ICard key={i.id} inc={i} cFields={data.cFields} onDel={session.isAdmin?del:null} onView={inc=>{setActive(inc);setModal("detail");}}/>)}</div>
              )}
              {data.incidents.length===0&&tab==="all"&&(
                <div className="empty"><div className="empty-ic">📋</div><div className="empty-t">Nenhuma ocorrência ainda.<br/>Clique em <strong>📝 Registrar</strong> para começar.</div></div>
              )}
            </div>
          )}
        </main>
      </div>

      {/* DETAIL */}
      {modal==="detail"&&active&&<DetailModal inc={active} cFields={data.cFields} isAdmin={session.isAdmin}
        onClose={()=>setModal(null)} onDel={id=>{del(id);setModal(null);}} onStatus={setStatus}/>}

      {/* FIELDS */}
      {modal==="fields"&&(
        <div className="ov" onClick={e=>e.target===e.currentTarget&&setModal(null)}>
          <div className="modal">
            <div className="modal-title">Gerenciar <span>Campos</span></div>
            <div className="field-list">
              {["Cliente","Data","AM","Coordenação","Área","Status","Descrição","Anexos"].map(f=>(
                <div key={f} className="fl-row"><span className="fl-name">{f}</span><span className="badge">fixo</span></div>
              ))}
              {data.cFields.map(f=>(
                <div key={f.id} className="fl-row">
                  <span className="fl-name">{f.label}</span>
                  <span className="badge">{FIELD_TYPES.find(t=>t.id===f.type)?.l}</span>
                  {f.required&&<span className="badge red">obrigatório</span>}
                  <span className="fdel" onClick={()=>delField(f.id)}>×</span>
                </div>
              ))}
            </div>
            <div className="mdivider"/>
            <p style={{fontSize:"11px",fontWeight:700,color:"var(--mu)",textTransform:"uppercase",letterSpacing:".09em",marginBottom:"11px"}}>Novo campo</p>
            <div className="fg"><label className="fg-lbl">Nome</label><input className="fi" placeholder="ex: Protocolo..." value={nf.label} onChange={e=>setNf(p=>({...p,label:e.target.value}))}/></div>
            <div className="fg"><label className="fg-lbl">Tipo</label><div className="tpills">{FIELD_TYPES.map(t=><div key={t.id} className={`tpill ${nf.type===t.id?"on":""}`} onClick={()=>setNf(p=>({...p,type:t.id}))}>{t.l}</div>)}</div></div>
            {nf.type==="select"&&<div className="fg"><label className="fg-lbl">Opções (vírgula)</label><input className="fi" placeholder="Urgente, Normal, Baixa" value={nf.options} onChange={e=>setNf(p=>({...p,options:e.target.value}))}/></div>}
            <div className="tog-row" onClick={()=>setNf(p=>({...p,required:!p.required}))}><div className={`tog ${nf.required?"on":""}`}/> Campo obrigatório</div>
            <div style={{display:"flex",gap:"8px",justifyContent:"flex-end",marginTop:"16px"}}>
              <button className="btn btn-g" style={{border:"1px solid var(--bd)"}} onClick={()=>setModal(null)}>Fechar</button>
              <button className="btn btn-p" onClick={addField}>Adicionar</button>
            </div>
          </div>
        </div>
      )}

      {/* CAT */}
      {modal==="cat"&&(
        <div className="ov" onClick={e=>e.target===e.currentTarget&&setModal(null)}>
          <div className="modal" style={{maxWidth:"360px"}}>
            <div className="modal-title">Nova <span>Área</span></div>
            <div className="fg"><label className="fg-lbl">Nome</label>
              <input className="fi" placeholder="ex: Financeiro..." value={newCat} onChange={e=>setNewCat(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addCat()} autoFocus/>
            </div>
            <div style={{display:"flex",gap:"8px",justifyContent:"flex-end",marginTop:"14px"}}>
              <button className="btn btn-g" style={{border:"1px solid var(--bd)"}} onClick={()=>setModal(null)}>Cancelar</button>
              <button className="btn btn-p" onClick={addCat}>Adicionar</button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

function ICard({ inc, cFields, onDel, onView }) {
  const cst=cc(inc.category);
  const filled=cFields.filter(cf=>inc.custom?.[cf.id]);
  const status=inc.status||"andamento";
  return (
    <div className="icard" onClick={()=>onView(inc)}>
      <div className="icard-top">
        <div className="icard-name">{inc.client}</div>
        <div className="tags">
          <span className={`tag ${status==="resolvido"?"tag-r":"tag-a"}`}>{status==="resolvido"?"✅ Resolvido":"🔄 Andamento"}</span>
          <span className="tag" style={{background:cst.bg,color:cst.c}}>{inc.category}</span>
        </div>
      </div>
      <div className="icard-meta">
        {inc.date&&<span>📅 {fmt(inc.date)}</span>}
        {inc.am&&<span>👤 {inc.am}</span>}
        {inc.coord&&<span>🏢 {inc.coord}</span>}
        {inc.registeredBy&&<span>✍️ {inc.registeredBy}</span>}
      </div>
      {inc.desc&&<div className="icard-desc">{inc.desc}</div>}
      {filled.length>0&&<div className="icard-chips">{filled.map(cf=><span key={cf.id} className="chip"><strong>{cf.label}:</strong> {inc.custom[cf.id]}</span>)}</div>}
      {inc.attachments?.length>0&&(
        <div className="icard-atts">
          {inc.attachments.slice(0,4).map((a,i)=>(
            <div key={i} className="ath" title={a.name}>
              {a.type?.startsWith("image/")?<img src={a.data} alt={a.name}/>:
              (a.type?.includes("audio")||a.name?.includes(".opus")||a.name?.includes(".ogg"))?"🎵":"📄"}
            </div>
          ))}
          {inc.attachments.length>4&&<div className="ath" style={{fontSize:"10px",color:"var(--mu)"}}>+{inc.attachments.length-4}</div>}
        </div>
      )}
      {onDel&&<div className="del-btn" onClick={e=>{e.stopPropagation();onDel(inc.id);}}>🗑</div>}
    </div>
  );
}

function DetailModal({ inc, cFields, isAdmin, onClose, onDel, onStatus }) {
  const [lb,setLb]=useState(null);
  const cst=cc(inc.category);
  const filled=cFields.filter(cf=>inc.custom?.[cf.id]);
  const status=inc.status||"andamento";
  const isImg=a=>a.type?.startsWith("image/");
  const isAud=a=>a.type?.includes("audio")||a.name?.includes(".opus")||a.name?.includes(".ogg");
  const open=a=>{if(isImg(a)){setLb(a.data);return;}const l=document.createElement("a");l.href=a.data;l.download=a.name;l.click();};
  return (
    <>
      <div className="ov" onClick={e=>e.target===e.currentTarget&&onClose()}>
        <div className="modal lg">
          <div className="det-hd">
            <div style={{flex:1}}>
              <div className="det-name">{inc.client}</div>
              <div style={{display:"flex",gap:"6px",marginTop:"6px",flexWrap:"wrap"}}>
                <span className="tag" style={{background:cst.bg,color:cst.c}}>{cic(inc.category)} {inc.category}</span>
                <span className={`tag ${status==="resolvido"?"tag-r":"tag-a"}`}>{status==="resolvido"?"✅ Resolvido":"🔄 Em andamento"}</span>
              </div>
            </div>
            <div className="det-x" onClick={onClose}>×</div>
          </div>
          <div className="det-pills">
            {inc.date&&<span className="det-pill">📅 {fmt(inc.date)}</span>}
            {inc.am&&<span className="det-pill">👤 {inc.am}</span>}
            {inc.coord&&<span className="det-pill">🏢 {inc.coord}</span>}
            {inc.registeredBy&&<span className="det-pill">✍️ {inc.registeredBy}</span>}
          </div>
          {inc.desc&&<div className="det-sec"><div className="det-sec-lbl">Descrição</div><div className="det-desc">{inc.desc}</div></div>}
          {filled.length>0&&<div className="det-sec"><div className="det-sec-lbl">Campos adicionais</div>
            <div className="det-cg">{filled.map(cf=><div key={cf.id} className="det-ci"><div className="det-ci-l">{cf.label}</div><div className="det-ci-v">{inc.custom[cf.id]}</div></div>)}</div>
          </div>}
          {inc.attachments?.length>0&&<div className="det-sec"><div className="det-sec-lbl">Anexos ({inc.attachments.length})</div>
            <div className="det-att-grid">{inc.attachments.map((a,i)=>(
              <div key={i} className="det-att" onClick={()=>open(a)}>
                {isImg(a)?<img src={a.data} alt={a.name}/>:<div className="det-att-f"><div className="det-att-ic">{isAud(a)?"🎵":"📄"}</div><div className="det-att-nm">{a.name}</div></div>}
              </div>
            ))}</div>
          </div>}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:"13px",borderTop:"1px solid var(--bd)",gap:"8px",flexWrap:"wrap"}}>
            {isAdmin
              ?<button style={{background:"transparent",border:"1px solid var(--err)",color:"var(--err)",fontFamily:"var(--fb)",fontSize:"12px",fontWeight:600,padding:"6px 13px",borderRadius:"7px",cursor:"pointer"}} onClick={()=>onDel(inc.id)}>🗑 Excluir</button>
              :<div/>}
            <div style={{display:"flex",gap:"7px"}}>
              <button className="btn btn-g" style={{fontSize:"12px",padding:"6px 13px",border:"1px solid var(--bd)"}}
                onClick={()=>onStatus(inc.id,status==="resolvido"?"andamento":"resolvido")}>
                {status==="resolvido"?"🔄 Reabrir":"✅ Marcar resolvido"}
              </button>
              <button className="btn btn-g" style={{border:"1px solid var(--bd)"}} onClick={onClose}>Fechar</button>
            </div>
          </div>
        </div>
      </div>
      {lb&&<div className="lightbox" onClick={()=>setLb(null)}><img src={lb} alt="anexo"/></div>}
    </>
  );
}
