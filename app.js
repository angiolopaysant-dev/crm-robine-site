(function(){
"use strict";
const C = self.CRMC;
const $ = id => document.getElementById(id);
const view = $("view");
const PDFV = "3.11.174", PDFLIB = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/"+PDFV+"/pdf.min.js", PDFWK = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/"+PDFV+"/pdf.worker.min.js";
const XLSX_JS="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js", MAMMOTH_JS="https://cdn.jsdelivr.net/npm/mammoth@1.6.0/mammoth.browser.min.js";
const STAGES = [["etude","En étude"],["attente","Attente retour client"],["accepte","Accepté"],["chantier","En chantier"]];
const STAGE_L = {etude:"En étude",attente:"Attente client",accepte:"Accepté",chantier:"En chantier",termine:"Terminé"};
const CATS = {preprojet:"Pré-projet",devis:"Devis",admin:"Administratif",technique:"Technique",photo:"Photo",chantier:"Dossier chantier"};
const TYPES = {rappel:"Rappel",historique:"Note",devis:"Devis",installation:"Matériel",fiche:"Fiche",fournisseur:"Fournisseur","dossier-chantier":"Dossier chantier"};
const ROLES = {admin:"Administrateur",edition:"Modification",lecture:"Lecture seule"};
const SOURCE = "site CRM";
const S = { fiches:[], rappels:[], envois:[], dossiers:{}, docs:[], recents:[], names:{}, agenda:[], rappelsRaw:[], over:{}, agendaDb:[], mailsDb:[], vus:{}, prevSeen:{}, blobs:{},
  cfg:null, key:null, rawKey:null, me:null, canEdit:false, genere:"",
  filt:"tous", q:"", dtab:"client", form:{type:"rappel"}, ready:false, seen:new Set() };

/* ---------- utilitaires ---------- */
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const slug = s => String(s||"").normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase().replace(/[^a-z0-9]+/g,"-").replace(/^-|-$/g,"");
const pad = n => String(n).padStart(2,"0");
const isoDay = d => d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate());
const today = () => isoDay(new Date());
const addDays = (iso,n) => { const d=new Date(iso+"T12:00:00"); d.setDate(d.getDate()+n); return isoDay(d); };
const monday = () => { const d=new Date(); const k=(d.getDay()+6)%7; d.setDate(d.getDate()-k); return isoDay(d); };
const frDate = iso => { if(!/^\d{4}-\d{2}-\d{2}/.test(iso||"")) return iso||""; const [y,m,d]=iso.slice(0,10).split("-"); return d+"/"+m+"/"+y; };
const frShort = iso => { if(!/^\d{4}-\d{2}-\d{2}/.test(iso||"")) return iso||""; const [y,m,d]=iso.slice(0,10).split("-"); return d+"/"+m; };
const frLong = iso => { try{ return new Date(iso.slice(0,10)+"T12:00:00").toLocaleDateString("fr-FR",{weekday:"short",day:"numeric",month:"short"}); }catch(e){ return iso; } };
const frStamp = iso => { try{ return new Date(iso).toLocaleString("fr-FR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"}); }catch(e){ return iso||""; } };
const relDay = iso => { const n=Math.round((new Date(iso+"T12:00:00")-new Date(today()+"T12:00:00"))/864e5);
  if(n===0) return "aujourd’hui"; if(n===1) return "demain"; if(n===-1) return "hier"; if(n<0) return "il y a "+(-n)+" j"; if(n<7) return "dans "+n+" j"; return frLong(iso); };
const euro = n => (n==null||n===""||isNaN(n)) ? "" : Number(n).toLocaleString("fr-FR",{style:"currency",currency:"EUR"});
const num = s => { if(typeof s==="number") return s; const m=String(s||"").replace(/\s| | /g,"").replace(/€/g,"").match(/-?\d+(?:[.,]\d+)?/); return m?Number(m[0].replace(",",".")):null; };
const rid = () => Math.random().toString(36).slice(2,6);
function nextSync(){ const n=new Date(), m=n.getHours()*60+n.getMinutes(); return m<445?"7h25":m<745?"12h25":"demain 7h25"; }
function toast(msg){ const t=document.createElement("div"); t.className="toast"; t.textContent=msg; document.body.append(t); setTimeout(()=>t.remove(),2800); }
const initials = s => String(s||"?").replace(/^(m\.|mme|cabinet)\s+/i,"").split(/\s+/).slice(0,2).map(w=>w[0]||"").join("").toUpperCase();
const ICON_OK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12l5 5L19 7"/></svg>';
const REF_RE = /\b((?:WH|PUZ|SUZ|PUMY|PUHZ|MXZ|MSZ|SLZ|PEAD|ERST|EHST|EHSD|ERSD|EHPT|WH-?SXC|WH-?SDC|WH-?MXC|WH-?UDZ)[- ]?[A-Z0-9]{2,}[A-Z0-9-]*)\b/g;
function typeProjet(s){ s=String(s||""); if(/air\s*[\/-]\s*eau|hydro|ecodan|aquarea/i.test(s)) return "PAC air/eau"; if(/gainable/i.test(s)) return "PAC air/air gainable"; if(/air\s*[\/-]\s*air|multi\s*split|mono\s*split|clim|drv|mural/i.test(s)) return "PAC air/air"; if(/chaudi/i.test(s)) return "Chaudière"; if(/ballon|thermodyn/i.test(s)) return "Chauffe-eau"; return ""; }
function refsIn(s){ const out=[]; String(s||"").replace(REF_RE,(m)=>{ if(!out.includes(m)) out.push(m); }); return out; }
const lsGet=(k,d)=>{ try{ const v=localStorage.getItem(k); return v==null?d:JSON.parse(v); }catch(e){ return d; } };
const lsSet=(k,v)=>{ try{ localStorage.setItem(k,JSON.stringify(v)); }catch(e){} };

/* ---------- fichiers chiffrés ---------- */
const EXT_MIME={pdf:"application/pdf",jpg:"image/jpeg",jpeg:"image/jpeg",png:"image/png",webp:"image/webp",gif:"image/gif",heic:"image/heic",svg:"image/svg+xml",txt:"text/plain",csv:"text/csv",md:"text/markdown",json:"application/json",mp4:"video/mp4",mov:"video/quicktime",webm:"video/webm",
  doc:"application/msword",docx:"application/vnd.openxmlformats-officedocument.wordprocessingml.document",xls:"application/vnd.ms-excel",xlsx:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",xlsm:"application/vnd.ms-excel.sheet.macroEnabled.12",odt:"application/vnd.oasis.opendocument.text",ods:"application/vnd.oasis.opendocument.spreadsheet"};
const extOf = n => (String(n||"").split(".").pop()||"").toLowerCase();
function mimeOf(file){ if(file.type && !/octet-stream/.test(file.type)) return file.type.split(";")[0]; return EXT_MIME[extOf(file.name)]||"application/octet-stream"; }
const docMime = d => d.mime || EXT_MIME[extOf(d.nom||d.fichier)] || "application/octet-stream";
const LOCAL = new Map();   // fichier → Blob ajouté pendant cette session (pas encore publié)
const URLS = new Map();    // fichier → URL objet déchiffrée
async function getBlob(path,mime){
  if(!path) return null;
  if(LOCAL.has(path)) return LOCAL.get(path);
  const info=S.blobs[path]; if(!info) return null;
  const r=await fetch("data/f/"+info.b); if(!r.ok) throw new Error("fichier "+r.status);
  const plain=await C.decFile(S.key,await r.arrayBuffer());
  return new Blob([plain],{type:mime||EXT_MIME[extOf(path)]||"application/octet-stream"});
}
async function blobUrl(path,mime){
  if(URLS.has(path)) return URLS.get(path);
  const p=getBlob(path,mime).then(b=>b?URL.createObjectURL(b):"").catch(()=>"");
  URLS.set(path,p); const u=await p; if(!u) URLS.delete(path); return u;
}
const blobAttr = (path,bg) => path ? (bg?' data-bg="':' data-src="')+esc(path)+'"' : "";
let hydrIO=null;
function hydrate(root){
  const els=(root||document).querySelectorAll("[data-bg],[data-src]");
  if(!hydrIO) hydrIO=new IntersectionObserver(es=>es.forEach(e=>{ if(!e.isIntersecting) return; hydrIO.unobserve(e.target); fill(e.target); }),{rootMargin:"200px"});
  els.forEach(el=>{ if(el.dataset.h) return; el.dataset.h="1"; hydrIO.observe(el); });
}
async function fill(el){
  const p=el.dataset.bg||el.dataset.src; const u=await blobUrl(p); if(!u) return;
  if(el.dataset.bg){ el.style.backgroundImage="url('"+u+"')"; el.querySelectorAll(".init").forEach(x=>x.remove()); if(el.classList.contains("av")) el.textContent=""; }
  else el.src=u;
}

/* ---------- API GitHub (écritures dans crm-robine) ---------- */
const utf8b64 = s => C.b64e(C.te.encode(s));
const GH = {
  async req(method,url,body){
    if(!S.cfg||!S.cfg.token) { const e=new Error("Pas de jeton d’écriture"); e.status=0; throw e; }
    const r=await fetch("https://api.github.com"+url,{method,cache:"no-store",headers:Object.assign({Authorization:"Bearer "+S.cfg.token,Accept:"application/vnd.github+json","X-GitHub-Api-Version":"2022-11-28"},body?{"Content-Type":"application/json"}:{}),body:body?JSON.stringify(body):undefined});
    if(!r.ok){ const e=new Error("GitHub "+r.status); e.status=r.status; try{ e.info=await r.json(); }catch(x){} throw e; }
    return r.status===204?null:r.json();
  },
  p(path){ return "/repos/"+S.cfg.repo+"/contents/"+path.split("/").map(encodeURIComponent).join("/"); },
  async get(path){
    try{ const o=await GH.req("GET",GH.p(path)+"?ref="+encodeURIComponent(S.cfg.branch));
      let b64=o.content; if(!b64 || o.encoding==="none"){ const bl=await GH.req("GET","/repos/"+S.cfg.repo+"/git/blobs/"+o.sha); b64=bl.content; }
      return {sha:o.sha,text:C.td.decode(C.b64d(b64))};
    }catch(e){ if(e.status===404) return null; throw e; }
  },
  put(path,b64,message,sha){ return GH.req("PUT",GH.p(path),Object.assign({message,content:b64,branch:S.cfg.branch},sha?{sha}:{})); },
  del(path,sha,message){ return GH.req("DELETE",GH.p(path),{message,sha,branch:S.cfg.branch}); }
};
let QUEUE=Promise.resolve(), busy=0;
function enqueue(fn){ busy++; syncLabel(); const p=QUEUE.then(fn); QUEUE=p.catch(()=>{}).then(()=>{ busy--; syncLabel(); }); return p; }
function syncLabel(){ const el=$("sync"); if(!el) return; el.classList.toggle("busy",busy>0); el.textContent=busy>0?"enregistrement…":(S.genere?"données du "+frStamp(S.genere):""); }
const who = () => S.me ? (S.me.nom||S.me.login) : "";
const msg = t => t+(who()?" (par "+who()+")":"");
/* Lit le JSON à jour dans le dépôt, applique fn, réécrit. fn renvoie undefined = rien à changer. */
function updateJson(path,def,fn,message){
  return enqueue(async()=>{
    for(let i=0;i<4;i++){
      const cur=await GH.get(path); let val=def; if(cur){ try{ val=JSON.parse(cur.text); }catch(e){ throw new Error(path+" est illisible dans le dépôt"); } }
      val=JSON.parse(JSON.stringify(val)); const nv=fn(val); if(nv===undefined) return val;
      try{ await GH.put(path,utf8b64(JSON.stringify(nv,null,1)+"\n"),msg(message),cur&&cur.sha); return nv; }
      catch(e){ if((e.status===409||e.status===422) && i<3){ await new Promise(r=>setTimeout(r,400+Math.random()*600)); continue; } throw e; }
    }
  });
}
function requireEdit(){ if(!S.canEdit){ toast("Ton compte est en lecture seule"); throw Object.assign(new Error("lecture"),{user:true}); } }

/* ---------- fiches markdown ---------- */
function parseFiche(md){
  md=String(md||"").replace(/\r/g,""); const fm={}; let body=md;
  if(md.startsWith("---")){ const end=md.indexOf("\n---",3); if(end>0){ md.slice(3,end).split("\n").forEach(l=>{ const i=l.indexOf(":"); if(i>0) fm[l.slice(0,i).trim()]=l.slice(i+1).trim().replace(/^"(.*)"$/,"$1"); }); body=md.slice(end+4); } }
  const sections=[]; let cur={titre:"",lignes:[]};
  body.split("\n").forEach(l=>{ const m=/^##\s+(.*)$/.exec(l); if(m){ sections.push(cur); cur={titre:m[1].trim(),lignes:[]}; } else cur.lignes.push(l); });
  sections.push(cur);
  return { fm, sections: sections.filter(s=>s.titre || s.lignes.some(x=>x.trim() && !/^#\s/.test(x))) };
}
const PCACHE=new Map(); function P(f){ const k=f.id+"|"+(f.maj||"")+"|"+f.contenu.length; if(!PCACHE.has(k)) PCACHE.set(k,parseFiche(f.contenu)); return PCACHE.get(k); }
function parseTable(lines){
  const rows=lines.filter(l=>/^\s*\|/.test(l)).map(l=>l.trim().replace(/^\||\|$/g,"").split("|").map(c=>c.trim()));
  if(rows.length<2) return null; const head=rows[0];
  const data=rows.slice(1).filter(r=>!r.every(c=>/^:?-{2,}:?$/.test(c)||c===""));
  return { head, rows:data.map(r=>{ const o={}; head.forEach((h,i)=>o[h]=r[i]||""); return o; }) };
}
const col=(row,...names)=>{ for(const k of Object.keys(row)){ const kk=slug(k); if(names.some(n=>kk.startsWith(n))) return row[k]; } return ""; };
const sec=(f,re)=>P(f).sections.find(s=>re.test(slug(s.titre)));
function inline(s){
  let h=esc(s);
  h=h.replace(/\[\[([^\]|]+)(?:\|([^\]]+))?\]\]/g,(m,t,a)=>{ const name=t.split("/").pop().replace(/\.md$/,""); const f=findByName(name); const label=a||name;
    return f?'<button class="linkbtn" data-open="'+esc(f.id)+'">'+label+'</button>':'<span>'+label+'</span>'; });
  return h.replace(/\*\*([^*]+)\*\*/g,"<strong>$1</strong>").replace(/(^|[^*])\*([^*\s][^*]*)\*/g,"$1<em>$2</em>");
}
function mdBlock(lines){
  let out="",list=false,tbl=[];
  const flushT=()=>{ if(!tbl.length) return; const t=parseTable(tbl); tbl=[]; if(!t) return;
    out+='<div class="tbl"><table><thead><tr>'+t.head.map(h=>"<th>"+inline(h)+"</th>").join("")+'</tr></thead><tbody>'+t.rows.map(r=>"<tr>"+t.head.map(h=>"<td>"+inline(r[h])+"</td>").join("")+"</tr>").join("")+"</tbody></table></div>"; };
  const closeL=()=>{ if(list){ out+="</ul>"; list=false; } };
  for(const l of lines){
    if(/^\s*<!--.*-->\s*$/.test(l)) continue;
    if(/^\s*\|/.test(l)){ closeL(); tbl.push(l); continue; } flushT();
    const h=/^(#{1,6})\s+(.*)$/.exec(l); if(h){ closeL(); if(h[1].length>1) out+="<h4>"+inline(h[2])+"</h4>"; continue; }
    const li=/^\s*[-*]\s+(.*)$/.exec(l); if(li){ if(!list){ out+="<ul>"; list=true; } out+="<li>"+inline(li[1])+"</li>"; continue; }
    closeL(); if(l.trim()) out+="<p>"+inline(l)+"</p>";
  }
  flushT(); closeL(); return out;
}

/* ---------- modèle de dossier ---------- */
const clients=()=>S.fiches.filter(f=>f.type==="client");
const fournisseurs=()=>S.fiches.filter(f=>f.type==="fournisseur");
function findByName(name){ const s=slug(name); return S.fiches.find(f=>slug(f.nom)===s && (f.type==="client"||f.type==="fournisseur")) || S.fiches.find(f=>slug(f.nom)===s); }
const D=id=>S.dossiers[id]||{};
const docsOf=(id,cat)=>S.docs.filter(d=>d.dossier===id && (!cat || d.cat===cat)).sort((a,b)=>(b.date||b.ajoute_le||"").localeCompare(a.date||a.ajoute_le||""));
function ficheDevis(f){ const s=sec(f,/^devis/); const t=s&&parseTable(s.lignes); return t?t.rows.filter(r=>col(r,"objet")||col(r,"statut")):[]; }
function linkedNotes(f){ return S.fiches.filter(d=>d.type!=="client"&&d.type!=="fournisseur"&&d.id!=="tableau-de-bord"&&slug((P(d).fm.client||"").replace(/[\[\]]/g,""))===slug(f.nom)); }
function projet(f){
  const d=D(f.id), dv=docsOf(f.id,"devis");
  const lines=ficheDevis(f).filter(r=>!/abandonn|refus|annul/i.test(col(r,"statut")));
  const withRef=lines.find(r=>refsIn(col(r,"objet")).length) || lines[lines.length-1];
  let intitule=d.intitule || (dv[0]&&dv[0].intitule) || (withRef?col(withRef,"objet").split(/[;:—]/)[0].replace(/^Devis Codial n°\s*\d+\s*/i,"").trim():"");
  let refs=d.ref_pac || (dv[0]&&dv[0].ref_pac) || (withRef?refsIn(col(withRef,"objet")).slice(0,3).join(" + "):"");
  let type=d.type_projet || (dv[0]&&dv[0].type_projet) || typeProjet(intitule+" "+refs);
  if(intitule.length>90) intitule=intitule.slice(0,88)+"…";
  return {intitule, refs, type};
}
const chantierFini=f=>{ const d=D(f.id); return !!(d.chantier_fin && d.chantier_fin<today()); };
function stageOf(f){
  const d=D(f.id); if(chantierFini(f)) return "termine"; if(d.etape) return d.etape;
  const st=slug(f.statut||"");
  if(/chantier/.test(st)) return "chantier"; if(/termin|entretien/.test(st)) return "termine";
  const dv=docsOf(f.id,"devis"); if(dv.some(x=>/sign|accept/i.test(x.statut||""))) return "accepte";
  if(dv.length || ficheDevis(f).some(r=>/envoy/i.test(col(r,"statut")))) return "attente";
  return "etude";
}
function dateCreation(f){ const d=D(f.id); if(d.date_creation) return d.date_creation; const all=(f.contenu.match(/\b20\d{2}-\d{2}-\d{2}\b/g)||[]).sort(); return all[0]||""; }
function dateSignature(f){ const d=D(f.id); if(d.date_signature) return d.date_signature; const s=docsOf(f.id,"devis").find(x=>/sign|accept/i.test(x.statut||"")); return s?(s.signe_le||""):""; }
const VUES=()=>S.me?lsGet("crm-vues-"+S.me.id,{}):{};
function score(f){ const v=VUES()[f.id]||{}; let s=(v.n||0); if(v.t){ const age=(Date.now()-new Date(v.t))/864e5; s+=Math.max(0,30-age); } if(f.maj){ const a=(Date.now()-new Date(f.maj))/864e5; s+=Math.max(0,20-a); }
  s+=S.envois.filter(e=>e.statut!=="rangé"&&slug(e.client)===slug(f.nom)).length*3; if(stageOf(f)==="chantier") s+=15; if(stageOf(f)==="termine") s-=40; return s; }
function pendingFor(f){ return S.envois.filter(e=>e.statut!=="rangé"&&(e.fiche===f.chemin||slug(e.client)===slug(f.nom))); }
function doneMarks(){ const s=new Set(); S.envois.forEach(e=>{ if(e.statut!=="rangé"&&e.rappel) s.add(e.rappel); }); return s; }
function matchClient(text){ const t=slug(text); if(!t) return null;
  return clients().find(f=>{ const n=slug(f.nom); if(n.length>=3 && t.includes(n)) return true; const c=slug((P(f).fm.contact||"").replace(/^(m|mme|m-et-mme)-/,"")); const last=c.split("-").find(w=>w.length>=4 && !["mme","monsieur","madame"].includes(w)); return last && t.includes(last); }); }
function evDate(e){ const s=(e.debut||"")+" "+(e.jours||""); let m=s.match(/(\d{4}-\d{2}-\d{2})/); if(m) return m[1]; m=s.match(/(\d{2})\/(\d{2})\/(\d{4})/); return m?m[3]+"-"+m[2]+"-"+m[1]:""; }
function evEnd(e){ const m=(e.jours||"").match(/->\s*(\d{4}-\d{2}-\d{2})/); return m?m[1]:evDate(e); }
function evTime(e){ const m=(e.debut||"").match(/(\d{1,2})[:h](\d{2})/); return m?m[1]+"h"+m[2]:""; }
function evClient(e){ if(e.dossier){ const f=S.fiches.find(x=>x.id===e.dossier); if(f) return f; } return matchClient((e.sujet||"")+" "+(e.lieu||"")); }
const WEEK_AGO=()=>new Date(Date.now()-7*864e5).toISOString();
const mine=x=>!!(S.me && x.par && (x.par===S.me.login||x.par===S.me.id));
function newsFor(f,since){
  since=since||S.vus[f.id]||WEEK_AGO(); const out=[];
  S.docs.forEach(d=>{ if(d.dossier===f.id && (d.ajoute_le||"")>since && !mine(d)) out.push({t:d.ajoute_le,l:(CATS[d.cat]||"Document")+" : "+(d.titre||d.nom)}); });
  S.mailsDb.forEach(m=>{ if(m.dossier===f.id && (m.ajoute_le||"")>since && !mine(m)) out.push({t:m.ajoute_le,l:"Mail : "+(m.sujet||"(sans objet)")}); });
  S.agendaDb.forEach(e=>{ if(e.dossier===f.id && (e.ajoute_le||"")>since && !mine(e)) out.push({t:e.ajoute_le,l:"Agenda : "+(e.sujet||"")+(evDate(e)?" ("+frShort(evDate(e))+")":"")}); });
  S.envois.forEach(e=>{ if((e.fiche===f.chemin||slug(e.client)===slug(f.nom)) && (e.cree||"")>since && !mine(e) && e.source!=="app CRM" && e.source!==SOURCE) out.push({t:e.cree,l:(TYPES[e.type]||"Info")+" : "+String(e.texte||"").split("\n")[0].slice(0,140)}); });
  return out.sort((a,b)=>(b.t||"").localeCompare(a.t||""));
}
const bubble=f=>{ if(chantierFini(f)) return ""; const n=newsFor(f).length; return n?'<span class="notif" aria-label="'+n+' nouveauté'+(n>1?"s":"")+'">'+(n>99?"99+":n)+'</span>':""; };
function coordOf(f){ const fm=P(f).fm, d=D(f.id), o=Object.assign({},fm); ["contact","telephone","email","adresse"].forEach(k=>{ if(d[k]) o[k]=d[k]; }); return o; }
function applyOver(){ S.rappels=S.rappelsRaw.map(r=>{ const o=S.over[r.id]; if(o && (o.old_date||"")===(r.date||"") && (o.old_action||"")===(r.action||"")) return Object.assign({},r,{date:o.date,action:o.action,modif:true}); return r; }); }
function mergeAgenda(){ S.agenda=S.agendaDb.slice(); }
const agKey=e=>"db:"+e._id;
const nameOf = par => par ? (S.names[par] || (/^u_/.test(par)?"":par)) : "";

/* ---------- écritures ---------- */
async function envoyer(doc){
  requireEdit();
  const id=today()+"-"+(slug(doc.client||"divers").slice(0,24)||"divers")+"-"+doc.type+"-"+rid();
  const row=Object.assign({cree:new Date().toISOString(),source:SOURCE,statut:"en attente"},doc,{par:S.me.login});
  await enqueue(()=>GH.put("_inbox/"+id+".json",utf8b64(JSON.stringify(row,null,1)+"\n"),msg("Envoi CRM : "+(TYPES[doc.type]||doc.type)+" pour "+(doc.client||"divers"))));
  S.envois.unshift(Object.assign({_id:id,_fichier:"_inbox/"+id+".json"},row)); badge();
  return id;
}
async function setDossier(id,patch){
  requireEdit(); patch=Object.assign({},patch,{maj:new Date().toISOString()});
  S.dossiers[id]=Object.assign({},S.dossiers[id],patch);
  const f=S.fiches.find(x=>x.id===id);
  const nv=await updateJson("_app/dossiers.json",{},o=>{ o[id]=Object.assign({},o[id],patch); return o; },"Dossier "+(f?f.nom:id)+" : "+Object.keys(patch).filter(k=>k!=="maj").join(", "));
  if(nv) S.dossiers=nv;
}
async function setDocs(fn,message){ requireEdit(); const nv=await updateJson("_app/documents.json",[],fn,message); if(nv) S.docs=nv; return nv; }
async function setRows(file,fn,message){ requireEdit(); const nv=await updateJson("_app/"+file,[],fn,message); if(nv) S[file==="mails.json"?"mailsDb":"agendaDb"]=withIds(nv); mergeAgenda(); return nv; }
const withIds = arr => (Array.isArray(arr)?arr:[]).map(x=>Object.assign({},x,{_id:x.doc_id||x.id}));
const cleanName = n => String(n||"fichier").normalize("NFC").replace(/[\\/:*?"<>|\u0000-\u001f]+/g,"_").replace(/^\.+/,"").slice(0,120) || "fichier";
async function shrinkPhoto(file){
  if(!/^image\/(jpeg|png|webp|heic|heif)$/.test(mimeOf(file)) || file.size<1.2*1024*1024) return file;
  try{ const bmp=await createImageBitmap(file); const M=2400, r=Math.min(1,M/Math.max(bmp.width,bmp.height));
    const c=document.createElement("canvas"); c.width=Math.round(bmp.width*r); c.height=Math.round(bmp.height*r); c.getContext("2d").drawImage(bmp,0,0,c.width,c.height);
    const b=await new Promise(ok=>c.toBlob(ok,"image/jpeg",0.85)); if(!b||b.size>=file.size) return file;
    return new File([b],file.name.replace(/\.[^.]+$/,"")+".jpg",{type:"image/jpeg"});
  }catch(e){ return file; }
}
async function uploadFile(file,dossierId,cat){
  requireEdit();
  if(cat==="photo"||/^image\//.test(mimeOf(file))) file=await shrinkPhoto(file);
  if(file.size>40*1024*1024){ const e=new Error("« "+file.name+" » dépasse 40 Mo : allège-le ou envoie-le en PDF."); e.user=true; throw e; }
  const b64=C.b64e(new Uint8Array(await file.arrayBuffer()));
  let nom=cleanName(file.name); const base=nom.replace(/(\.[^.]*)?$/,""), ext=(nom.match(/\.[^.]*$/)||[""])[0];
  for(let i=1;i<20;i++){
    const fichier="fichiers/"+dossierId+"/"+cat+"/"+nom;
    if(!S.docs.some(d=>d.fichier===fichier) && !S.blobs[fichier] && !LOCAL.has(fichier)){
      try{ await enqueue(()=>GH.put("_app/"+fichier,b64,msg("Fichier ajouté : "+fichier)));
        LOCAL.set(fichier,new Blob([file],{type:mimeOf(file)})); return {fichier,nom,mime:mimeOf(file),taille:file.size}; }
      catch(e){ if(e.status!==422) throw e; }
    }
    nom=base+"-"+(i+1)+ext;
  }
  throw new Error("nom de fichier");
}
async function addDocs(dossierId,cat,files,extra){
  const done=[];
  for(const file of files){
    try{ const r=await uploadFile(file,dossierId,cat);
      const row=Object.assign({ajoute_le:new Date().toISOString(),cat,date:today(),dossier:dossierId,mime:r.mime,nom:r.nom,source:SOURCE,taille:r.taille,id:"d"+Date.now().toString(36)+rid(),fichier:r.fichier,par:S.me.login},extra||{});
      await setDocs(arr=>{ arr.push(row); return arr; },"Document ajouté au dossier "+dossierId+" : "+r.nom); done.push(row);
    }catch(e){ toast(e.user?e.message:"« "+file.name+" » n’a pas pu être ajouté. Réessaie."); }
  }
  if(done.length){ toast(done.length>1?done.length+" fichiers ajoutés":"Fichier ajouté"); rerender(); }
  return done;
}
async function removeDoc(docId){
  const d=S.docs.find(x=>x.id===docId); if(!d) return;
  await setDocs(arr=>arr.filter(x=>x.id!==docId),"Document retiré : "+(d.titre||d.nom));
  if(d.fichier && !S.docs.some(x=>x.fichier===d.fichier)){
    enqueue(async()=>{ const cur=await GH.req("GET",GH.p("_app/"+d.fichier)+"?ref="+encodeURIComponent(S.cfg.branch)).catch(()=>null); if(cur&&cur.sha) await GH.del("_app/"+d.fichier,cur.sha,msg("Fichier retiré : "+d.fichier)); }).catch(()=>{});
  }
}

/* ---------- lecture automatique d'un devis PDF ---------- */
function loadScript(src){ return new Promise((ok,ko)=>{ if(document.querySelector('script[src="'+src+'"]')) return ok(); const s=document.createElement("script"); s.src=src; s.crossOrigin="anonymous"; s.onload=ok; s.onerror=()=>ko(new Error("script")); document.head.append(s); }); }
async function pdfLib(){ await loadScript(PDFLIB); const lib=window.pdfjsLib||window["pdfjs-dist/build/pdf"]; lib.GlobalWorkerOptions.workerSrc=PDFWK; return lib; }
async function pdfText(file){
  const lib=await pdfLib();
  const pdf=await lib.getDocument({data:new Uint8Array(await file.arrayBuffer()),isEvalSupported:false}).promise;
  let txt=""; for(let i=1;i<=Math.min(pdf.numPages,5);i++){ const p=await pdf.getPage(i); const c=await p.getTextContent(); txt+=c.items.map(x=>x.str+(x.hasEOL?"\n":" ")).join("")+"\n"; }
  return txt;
}
function regexDevis(t){
  const g=re=>{ const m=t.match(re); return m?m[1]:""; };
  const ht=num(g(/total\s*(?:net\s*)?h\.?\s*t\.?\s*[:\s]*([\d\s.,  ]+)/i)), ttc=num(g(/(?:total|net\s*à\s*payer)\s*t\.?\s*t\.?\s*c\.?\s*[:\s]*([\d\s.,  ]+)/i));
  let date=g(/(\d{2}\/\d{2}\/\d{4})/); if(date){ const [d,m,y]=date.split("/"); date=y+"-"+m+"-"+d; }
  const refs=refsIn(t).slice(0,3);
  const tp=typeProjet(t);
  return {numero:g(/devis\s*(?:n[°o]\.?|num[ée]ro)\s*:?\s*([A-Z0-9-]+)/i), date, ht, ttc, ref_pac:refs.join(" + "), type_projet:tp, intitule:[tp,refs[0]].filter(Boolean).join(" ")};
}
async function lireDevis(file,onStep){
  let t=""; try{ onStep("Lecture du PDF…"); t=await pdfText(file); }catch(e){ return {err:"Je n’arrive pas à lire ce PDF ici : remplis les champs à la main."}; }
  return regexDevis(t);
}

/* ---------- rendu commun ---------- */
function setHead(title,back){ $("title").textContent=title; $("back").hidden=!back; document.title=title==="CRM Robine"?"CRM Robine":title+" · CRM Robine"; }
function setTab(t){ document.querySelectorAll("[data-tab]").forEach(a=>{ if(a.dataset.tab===t) a.setAttribute("aria-current","page"); else a.removeAttribute("aria-current"); }); }
function badge(){ const n=S.envois.filter(e=>e.statut!=="rangé").length; document.querySelectorAll(".badge").forEach(b=>{ b.hidden=!n; b.textContent=n; }); }
function stageSteps(st){ const i=STAGES.findIndex(s=>s[0]===st); return '<div class="steps" aria-hidden="true">'+STAGES.map((s,k)=>'<i class="'+(k<=i||st==="termine"?"on":"")+'"></i>').join("")+'</div>'; }
function card(f){
  const fm=coordOf(f), d=D(f.id), pj=projet(f), st=stageOf(f), cr=dateCreation(f), sg=dateSignature(f);
  return '<button class="card" data-open="'+esc(f.id)+'"><div class="card-ph"'+blobAttr(d.photo,true)+'><span class="init">'+esc(initials(f.nom))+'</span><span class="stage st-'+st+'">'+esc(STAGE_L[st])+'</span></div>'+
    '<div class="card-b"><span class="card-n">'+esc(f.nom)+'</span><span class="card-p">'+esc([pj.type!==pj.intitule&&!slug(pj.intitule).includes(slug(pj.type))?pj.type:"",pj.intitule].filter(Boolean).join(" — ")||"Projet à préciser")+(pj.refs?'<br><small>'+esc(pj.refs)+'</small>':"")+'</span>'+
    (fm.adresse?'<span class="card-l">'+esc(fm.adresse)+'</span>':"")+
    ((fm.telephone||fm.email)?'<span class="card-l">'+esc([fm.telephone,fm.email].filter(Boolean).join("  ·  "))+'</span>':"")+
    '<span class="card-d"><span>Créé <b>'+esc(cr?frDate(cr):"—")+'</b></span><span>Signé <b>'+esc(sg?frDate(sg):"—")+'</b></span></span>'+stageSteps(st)+'</div></button>';
}

/* ---------- ACCUEIL ---------- */
function weekItems(){
  const m=monday(), sun=addDays(m,6), out=[];
  clients().forEach(f=>{ const d=D(f.id); if(d.chantier_debut){ const fin=d.chantier_fin||d.chantier_debut; if(d.chantier_debut<=sun && fin>=m) out.push({edit:"chantier:"+f.id,date:d.chantier_debut<m?m:d.chantier_debut,f,txt:"Chantier "+(d.chantier_fin?"du "+frShort(d.chantier_debut)+" au "+frShort(d.chantier_fin):"le "+frShort(d.chantier_debut))}); } });
  S.agenda.forEach(e=>{ const dt=evDate(e); if(!dt||dt>sun||evEnd(e)<m) return; const f=evClient(e); if(!f) return; if(out.some(o=>o.f===f&&o.date===dt)) return; out.push({edit:"agenda:"+agKey(e),date:dt<m?m:dt,f,txt:(evTime(e)?evTime(e)+" — ":"")+e.sujet}); });
  clients().forEach(f=>{ if(stageOf(f)==="chantier" && !out.some(o=>o.f===f)) out.push({edit:"chantier:"+f.id,date:"",f,txt:"En chantier, dates à fixer"}); });
  return out.sort((a,b)=>(a.date||"9").localeCompare(b.date||"9"));
}
function vAccueil(){
  setHead("CRM Robine"); setTab("accueil");
  const wk=weekItems(), m=monday();
  let h='<section class="week"><h2>Chantiers de la semaine <small>du '+frShort(m)+' au '+frShort(addDays(m,6))+'</small></h2>'+
    (wk.length?'<ul>'+wk.map(w=>'<li><button data-open="'+esc(w.f.id)+'" data-edit="'+esc(w.edit||"dossier:"+w.f.id)+'"><span class="wd">'+esc(w.date?new Date(w.date+"T12:00").toLocaleDateString("fr-FR",{weekday:"long",day:"numeric"}):"À caler")+'</span><span class="wn">'+esc(w.f.nom)+'</span><span class="wt">'+esc(w.txt)+'</span></button></li>').join("")+'</ul>'
      :'<p class="empty">Aucun chantier prévu cette semaine. Indique les dates de chantier dans chaque dossier.</p>')+'</section>';
  const rec=S.recents.map(id=>S.fiches.find(f=>f.id===id)).filter(Boolean).slice(0,10);
  if(rec.length) h+='<section><h2>Consultés récemment</h2><div class="recent">'+rec.map(f=>{ const d=D(f.id); return '<button data-open="'+esc(f.id)+'"><span class="av"'+blobAttr(d.photo,true)+'>'+esc(initials(f.nom))+'</span>'+esc(f.nom)+bubble(f)+'</button>'; }).join("")+'</div></section>';
  const counts={}; clients().forEach(f=>{ const s=stageOf(f); counts[s]=(counts[s]||0)+1; });
  h+='<section><h2>Dossiers <small>les plus actifs en premier, clic droit (ou appui long) pour modifier</small></h2><div style="display:flex;flex-direction:column;gap:10px;margin-bottom:12px"><div class="search"><input type="search" id="q" placeholder="Chercher : nom, ville, téléphone, référence PAC…" value="'+esc(S.q)+'" aria-label="Chercher un dossier"></div>'+
    '<div class="filters" role="group" aria-label="Filtrer par étape"><button data-filt="tous" aria-pressed="'+(S.filt==="tous")+'">Tous</button>'+STAGES.concat([["termine","Terminé"]]).map(([k,l])=>'<button data-filt="'+k+'" aria-pressed="'+(S.filt===k)+'">'+l+' '+(counts[k]||0)+'</button>').join("")+'</div></div><div class="grid" id="grid"></div></section>';
  const fo=fournisseurs(); if(fo.length) h+='<section><h2>Fournisseurs</h2><div class="recent">'+fo.map(f=>'<button data-open="'+esc(f.id)+'"><span class="av">'+esc(initials(f.nom))+'</span>'+esc(f.nom)+'</button>').join("")+'</div></section>';
  view.innerHTML=h; drawGrid(); hydrate(view);
  $("q").addEventListener("input",e=>{ S.q=e.target.value; drawGrid(); });
}
function drawGrid(){
  const g=$("grid"); if(!g) return; const q=slug(S.q);
  const l=clients().filter(f=>(S.filt==="tous"?stageOf(f)!=="termine"||!!q:stageOf(f)===S.filt) && (!q || slug(f.nom+" "+f.contenu+" "+JSON.stringify(D(f.id))).includes(q) || f.contenu.replace(/\s/g,"").includes(S.q.replace(/\s/g,""))))
    .sort((a,b)=>score(b)-score(a));
  g.innerHTML=l.length?l.map(f=>'<div class="cardw" data-edit="dossier:'+esc(f.id)+'">'+bubble(f)+card(f)+'</div>').join(""):'<p class="empty">Aucun dossier ici. Crée-le depuis « Noter » → Nouveau dossier.</p>';
  hydrate(g);
}

/* ---------- ÉCHÉANCES (30 jours) ---------- */
function vEcheances(){
  setHead("Échéances à 1 mois"); setTab("echeances");
  const t=today(), lim=addDays(t,31), marks=doneMarks(), items=[];
  S.rappels.filter(r=>!r.fait).forEach(r=>{ const f=findByName(r.client); const dt=/^\d{4}/.test(r.date||"")?r.date:""; if(dt && dt>lim) return; items.push({edit:"rappel:"+r.id,date:dt,f,nom:r.client,txt:r.action,kind:"Rappel",rem:r}); });
  clients().forEach(f=>{ const d=D(f.id);
    if(d.chantier_debut && d.chantier_debut<=lim && (d.chantier_fin||d.chantier_debut)>=t) items.push({edit:"chantier:"+f.id,date:d.chantier_debut<t?t:d.chantier_debut,f,nom:f.nom,txt:"Chantier"+(d.chantier_fin?" jusqu’au "+frDate(d.chantier_fin):""),kind:"Chantier"});
    if(d.chantier_fin && d.chantier_fin>=t && d.chantier_fin<=lim) items.push({edit:"chantier:"+f.id,date:d.chantier_fin,f,nom:f.nom,txt:"Fin de chantier prévue",kind:"Chantier"});
  });
  S.agenda.forEach(e=>{ const dt=evDate(e); if(!dt||dt<t||dt>lim) return; const f=evClient(e); if(f) items.push({edit:"agenda:"+agKey(e),date:dt,f,nom:f.nom,txt:(evTime(e)?evTime(e)+" — ":"")+e.sujet,kind:"Agenda"}); });
  const late=items.filter(i=>i.date&&i.date<t).length, tod=items.filter(i=>i.date===t).length, up=items.filter(i=>i.date>t).length;
  let h='<div class="gauge"><div class="g-hot"><b>'+late+'</b><span>en retard</span></div><div class="g-warm"><b>'+tod+'</b><span>aujourd’hui</span></div><div class="g-cold"><b>'+up+'</b><span>d’ici le '+frShort(lim)+'</span></div></div>';
  const by={}; items.forEach(i=>{ const k=i.f?i.f.id:"_"+slug(i.nom); (by[k]=by[k]||{f:i.f,nom:i.nom,list:[]}).list.push(i); });
  const groups=Object.values(by).map(g=>{ g.list.sort((a,b)=>(a.date||"9").localeCompare(b.date||"9")); g.first=g.list[0].date||"9"; return g; }).sort((a,b)=>a.first.localeCompare(b.first));
  if(!groups.length) h+='<p class="empty">Rien d’ici un mois. Ajoute des dates de chantier dans les dossiers, ou des rappels depuis « Noter ».</p>';
  groups.forEach(g=>{
    h+='<section class="month"><h2>'+(g.f?'<button class="linkbtn" data-open="'+esc(g.f.id)+'" style="font:inherit">'+esc(g.nom)+'</button>':esc(g.nom))+(g.f?' <small>'+esc(projet(g.f).intitule||"")+'</small>':"")+'</h2><ul class="list">'+
      g.list.map(i=>{ const st=!i.date?"":i.date<t?"hot":i.date===t?"warm":""; const done=i.rem&&marks.has(i.rem.id);
        return '<li class="due '+st+(done?" done":"")+'" data-edit="'+esc(i.edit||"")+'"><span class="due-d">'+esc(i.date?relDay(i.date):"sans date")+'</span><div class="due-b"><span class="due-k">'+esc(i.kind)+'</span><br>'+inline(i.txt)+'</div>'+
          (i.rem?'<button class="tick" aria-pressed="'+!!done+'" aria-label="Marquer comme fait" data-tick="'+esc(i.rem.id)+'" '+(done||!S.canEdit?"disabled":"")+'>'+(done?ICON_OK:"")+'</button>':"")+'</li>'; }).join("")+'</ul></section>';
  });
  view.innerHTML=h;
}

/* ---------- DOSSIER ---------- */
const DTABS=[["client","Client"],["preprojet","Pré-projet"],["devis","Devis & équipement"],["chantier","Dossier chantier"],["admin","Admin & technique"],["photos","Photos"]];
function vDossier(id){
  const f=S.fiches.find(x=>x.id===id);
  if(!f){ setHead("Dossier introuvable",true); view.innerHTML='<p class="empty">Ce dossier n’existe pas encore. S’il vient d’être créé, il apparaîtra après la synchro de '+nextSync()+'.</p>'; return; }
  if(f.type!=="client"&&f.type!=="fournisseur"){ setHead(f.nom,true); view.innerHTML='<article class="box md">'+mdBlock(P(f).sections.flatMap(s=>(s.titre?["## "+s.titre]:[]).concat(s.lignes)))+'</article>'; return; }
  setHead(f.nom,true); setTab("");
  if(!S.seen.has(id)){ S.seen.add(id); S.prevSeen[id]=S.vus[id]||WEEK_AGO(); bumpView(id); }
  const fm=coordOf(f), d=D(id), pj=projet(f), st=stageOf(f), isC=f.type==="client";
  const tel=(fm.telephone||"").trim(), mail=(fm.email||"").trim(), adr=(fm.adresse||"").trim();
  const ic={tel:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.9v3a2 2 0 01-2.2 2 19.8 19.8 0 01-8.6-3.1 19.5 19.5 0 01-6-6A19.8 19.8 0 012.1 4.2 2 2 0 014.1 2h3a2 2 0 012 1.7c.1.9.4 1.8.7 2.7a2 2 0 01-.5 2.1L8 9.8a16 16 0 006 6l1.3-1.3a2 2 0 012.1-.4c.9.3 1.8.6 2.7.7a2 2 0 011.7 2z"/></svg>',
    mail:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="M22 6l-10 7L2 6"/></svg>',
    map:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s7-6.2 7-12a7 7 0 10-14 0c0 5.8 7 12 7 12z"/><circle cx="12" cy="10" r="2.5"/></svg>'};
  const ro=S.canEdit?"":" disabled";
  let h='<section class="dhead" data-edit="dossier:'+esc(id)+'"><div class="cover"'+blobAttr(d.photo,true)+'><span class="init">'+esc(initials(f.nom))+'</span>'+
    (S.canEdit?'<label>'+(d.photo?"Changer la photo":"Ajouter une photo")+'<input type="file" accept="image/*" data-cover="'+esc(id)+'"></label>':"")+'</div>'+
    '<div class="dinfo"><div><h2 class="dname">'+esc(f.nom)+'</h2>'+(isC?'<p class="dproj"><span>'+esc(pj.type||"")+'</span><span>'+esc(pj.intitule&&pj.intitule!==pj.type?pj.intitule:"")+'</span>'+(pj.refs?'<small>'+esc(pj.refs)+'</small>':"")+(S.canEdit?' <button class="linkbtn" data-editproj="'+esc(id)+'" style="font-size:.85rem">modifier</button>':"")+'</p>':'<p class="hint">Fournisseur</p>')+
    '<p class="hint">'+esc([fm.contact,adr].filter(Boolean).join(", "))+'</p></div>'+
    (isC?'<div class="stagebar" role="group" aria-label="Étape du dossier">'+STAGES.map(([k,l])=>'<button class="st-'+k+'" data-stage="'+k+'" aria-pressed="'+(st===k)+'"'+ro+'>'+l+'</button>').join("")+'</div>'+
    '<div class="dates"><label>Création du dossier<input type="date" data-date="date_creation" value="'+esc(dateCreation(f))+'"'+ro+'></label><label>Signature du devis<input type="date" data-date="date_signature" value="'+esc(dateSignature(f))+'"'+ro+'></label><label>Début de chantier<input type="date" data-date="chantier_debut" value="'+esc(d.chantier_debut||"")+'"'+ro+'></label><label>Fin de chantier<input type="date" data-date="chantier_fin" value="'+esc(d.chantier_fin||"")+'"'+ro+'></label></div>':"")+
    '<div class="actions">'+(tel?'<a href="tel:'+esc(tel.replace(/[^\d+]/g,""))+'">'+ic.tel+esc(tel)+'</a>':'<span>'+ic.tel+'Pas de n°</span>')+
    (mail?'<a href="mailto:'+esc(mail)+'">'+ic.mail+'Écrire</a>':'<span>'+ic.mail+'Pas d’email</span>')+
    (adr?'<a href="https://www.google.com/maps/dir/?api=1&destination='+encodeURIComponent(adr)+'" target="_blank" rel="noopener noreferrer">'+ic.map+'Itinéraire</a>':'<span>'+ic.map+'Pas d’adresse</span>')+'</div></div></section>';
  const nw=S.prevSeen[id]?newsFor(f,S.prevSeen[id]):[];
  if(nw.length) h+='<div class="news"><strong>'+nw.length+' nouveauté'+(nw.length>1?"s":"")+' depuis ta dernière visite</strong>'+nw.slice(0,8).map(x=>'<p>'+esc(x.l)+' <small class="hint">'+esc(frStamp(x.t))+'</small></p>').join("")+'</div>';
  const pend=pendingFor(f);
  if(pend.length) h+='<div class="pending"><strong>'+pend.length+' info'+(pend.length>1?"s":"")+' en attente de rangement dans la fiche ('+nextSync()+')</strong>'+pend.slice(0,4).map(e=>'<p>'+esc(TYPES[e.type]||e.type)+' : '+esc(e.texte)+'</p>').join("")+'</div>';
  const tabs=isC?DTABS:[["client","Fiche"]];
  if(!tabs.some(t=>t[0]===S.dtab)) S.dtab="client";
  const n={preprojet:docsOf(id,"preprojet").length+linkedNotes(f).length, devis:docsOf(id,"devis").length, chantier:docsOf(id,"chantier").length, admin:docsOf(id,"admin").length+docsOf(id,"technique").length, photos:docsOf(id,"photo").length};
  h+='<div class="dtabs" role="tablist">'+tabs.map(([k,l])=>'<button role="tab" data-dtab="'+k+'" aria-selected="'+(S.dtab===k)+'">'+l+(n[k]?'<span class="n">'+n[k]+'</span>':"")+'</button>').join("")+'</div><div id="dpane"></div>';
  view.innerHTML=h; drawPane(f); hydrate(view);
}
function drawPane(f){
  const el=$("dpane"); if(!el) return; const t=S.dtab;
  el.innerHTML = t==="preprojet"?paneDocs(f,["preprojet"],"Plans scannés, étude thermique, note de dimensionnement, relevés, notes de visite…",true)
    : t==="devis"?paneDevis(f) : t==="chantier"?paneChantier(f) : t==="admin"?paneAdmin(f) : t==="photos"?panePhotos(f) : paneClient(f);
  hydrate(el);
}
function docRow(d){
  const img=/^image\//.test(docMime(d)); const w=nameOf(d.par); const who2=w?" par "+w:"";
  return '<li class="doc" data-edit="doc:'+esc(d.id)+'"><span class="doc-ic'+(img?" img":"")+'"'+(img?blobAttr(d.fichier,true):"")+'>'+(img?"":esc(((d.nom||"").split(".").pop()||"DOC").toUpperCase().slice(0,4)))+'</span>'+
    '<div class="doc-b"><a class="doc-n" href="#" data-viewfile="'+esc(d.id)+'">'+esc(d.titre||d.nom)+'</a><div class="doc-m">'+esc(CATS[d.cat]||"")+", "+esc(frDate(d.date||(d.ajoute_le||"").slice(0,10)))+esc(who2)+(d.source==="mail"?esc(", reçu par mail"):"")+(d.taille?", "+Math.max(1,Math.round(d.taille/1024))+" Ko":"")+(!S.blobs[d.fichier]&&!LOCAL.has(d.fichier)?esc(", en cours de publication"):"")+'</div></div>'+
    (S.canEdit?'<button class="doc-x" data-deldoc="'+esc(d.id)+'" aria-label="Retirer du dossier">×</button>':"")+'</li>';
}
function dropZone(id,cat,label,accept){
  if(!S.canEdit) return '<p class="conn">L’ajout de fichiers n’est pas disponible pour ton accès. Demande à Angi de te donner le droit de modifier.</p>';
  return '<div class="drop" data-drop="'+esc(cat)+'" data-for="'+esc(id)+'"><span>'+esc(label)+'</span><label class="btn">Ajouter des fichiers<input type="file" multiple '+(accept?'accept="'+accept+'"':"")+' data-up="'+esc(cat)+'" data-for="'+esc(id)+'"></label><small class="hint">Tous formats : PDF, photos, vidéos, Word, Excel… Tout s’ouvre d’un clic dans le site.</small></div>';
}
function paneDocs(f,cats,label,withNotes){
  const list=cats.flatMap(c=>docsOf(f.id,c));
  let h='<div class="cols" style="margin-top:14px"><section style="display:flex;flex-direction:column;gap:10px">'+dropZone(f.id,cats[0],label)+(list.length?'<ul class="list">'+list.map(docRow).join("")+'</ul>':'<p class="empty">Aucun document pour l’instant.</p>')+'</section>';
  if(withNotes){
    const std=/^(coordonnees|contact|installations|devis|rappels|historique|fournisseurs-lies|dossiers-clients-lies|ce-qu-on-y-commande)$/;
    const notes=P(f).sections.filter(s=>s.titre && !std.test(slug(s.titre)));
    const devSec=sec(f,/^devis/); const devExtra=devSec?devSec.lignes.filter(l=>!/^\s*\|/.test(l)&&!/^\s*<!--/.test(l)):[];
    const ln=linkedNotes(f);
    h+='<section style="display:flex;flex-direction:column;gap:10px"><h3>Notes du dossier</h3>';
    if(!notes.length && !devExtra.join("").trim() && !ln.length) h+='<p class="empty">Pas de notes d’étude dans la fiche. Ajoute-les depuis « Noter » → Note.</p>';
    if(devExtra.join("").trim()) h+='<details class="box md" open><summary>Étude et métré</summary>'+mdBlock(devExtra)+'</details>';
    notes.forEach(s=>{ h+='<details class="box md"><summary>'+esc(s.titre)+'</summary>'+mdBlock(s.lignes)+'</details>'; });
    if(ln.length) h+='<ul class="list">'+ln.map(d=>'<li><button class="doc" data-open="'+esc(d.id)+'" style="width:100%;text-align:left;cursor:pointer"><span class="doc-ic">NOTE</span><div class="doc-b"><span class="doc-n">'+esc(d.nom)+'</span><div class="doc-m">'+esc(d.statut||"")+'</div></div></button></li>').join("")+'</ul>';
    h+='</section>';
  }
  return h+'</div>';
}
function paneClient(f){
  const fm=coordOf(f), id=f.id;
  const coord=(sec(f,/^coordonnees|^contact$/)||{lignes:[]}).lignes.filter(l=>/^\s*-\s/.test(l));
  const kv=[["Contact",fm.contact],["Téléphone",fm.telephone],["Email",fm.email],["Adresse",fm.adresse]].filter(x=>x[1]);
  const extra=coord.filter(l=>!/^\s*-\s*(contact|t[ée]l[ée]phone|email|e-mail|adresse)\s*:/i.test(l));
  let left='<section><h3>Coordonnées</h3><div class="box"><dl class="kv">'+kv.map(([k,v])=>'<dt>'+k+'</dt><dd>'+esc(v)+'</dd>').join("")+'</dl>'+(extra.length?'<div class="md" style="margin-top:8px">'+mdBlock(extra)+'</div>':"")+'</div></section>';
  const inst=(()=>{ const s=sec(f,/^installations/); const t=s&&parseTable(s.lignes); return t?t.rows.filter(r=>Object.values(r).some(Boolean)):[]; })();
  const eqIc='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="14" cy="12" r="4"/><path d="M14 8v8M10 12h8M6 9h2M6 13h2"/></svg>';
  left+='<section><h3>Parc d’équipement</h3>'+(inst.length?'<ul class="list">'+inst.map(r=>{ const pose=col(r,"mise-en-service","pose"); const rest=Object.entries(r).filter(([k,v])=>v && !/equipement|marque|mise-en-service|pose/.test(slug(k)));
      return '<li class="box equip" data-edit="ligne:'+esc(id)+'" data-txt="'+esc(Object.entries(r).filter(([k,v])=>v).map(([k,v])=>k+" : "+v).join(" ; "))+'"><span class="equip-ic">'+eqIc+'</span><div><b>'+inline(col(r,"equipement")||"Équipement")+'</b><span>'+inline(col(r,"marque"))+'</span><br><span>Posé le <strong>'+esc(pose?frDate(pose):"—")+'</strong>'+(rest.length?", "+rest.map(([k,v])=>esc(k)+" : "+inline(v)).join(", "):"")+'</span></div></li>'; }).join("")+'</ul>'
      :'<p class="empty">Aucun équipement posé enregistré. Ajoute-le depuis « + Matériel » une fois la pose faite.</p>')+(S.canEdit?'<div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap"><button class="btn ghost" data-new="installation" data-for="'+esc(id)+'">+ Matériel posé</button><button class="btn ghost" data-new="historique" data-for="'+esc(id)+'">+ Note</button><button class="btn ghost" data-new="rappel" data-for="'+esc(id)+'">+ Rappel</button></div>':"")+'</section>';
  const hist=((sec(f,/^historique/)||{lignes:[]}).lignes).map(l=>/^\s*-\s*(\d{4}-\d{2}-\d{2})?\s*[—-]?\s*(.*)$/.exec(l)).filter(m=>m&&m[2]).map(m=>({d:m[1]||"",t:m[2]}));
  S.agenda.forEach(e=>{ const dt=evDate(e); if(dt && dt<=today() && evClient(e)===f) hist.push({d:dt,t:"Agenda : "+e.sujet}); });
  hist.sort((a,b)=>(b.d||"").localeCompare(a.d||""));
  const marks=doneMarks();
  const rem=S.rappels.filter(r=>slug(r.client)===slug(f.nom)&&!r.fait).sort((a,b)=>(a.date||"9").localeCompare(b.date||"9"));
  let right='';
  if(rem.length) right+='<section><h3>Rappels</h3><ul class="list">'+rem.map(r=>{ const dn=marks.has(r.id); const s=!/^\d{4}/.test(r.date||"")?"":r.date<today()?"hot":r.date===today()?"warm":"";
    return '<li class="due '+s+(dn?" done":"")+'" data-edit="rappel:'+esc(r.id)+'"><span class="due-d">'+esc(/^\d{4}/.test(r.date||"")?relDay(r.date):"sans date")+'</span><div class="due-b">'+inline(r.action)+'</div><button class="tick" aria-pressed="'+dn+'" data-tick="'+esc(r.id)+'" aria-label="Marquer comme fait" '+(dn||!S.canEdit?"disabled":"")+'>'+(dn?ICON_OK:"")+'</button></li>'; }).join("")+'</ul></section>';
  const tl=i=>'<li data-edit="ligne:'+esc(id)+'" data-txt="'+esc((i.d?i.d+" — ":"")+i.t)+'"><time>'+esc(i.d?frLong(i.d):"")+'</time><p>'+inline(i.t)+'</p></li>';
  right+='<section><h3>Dernières interventions</h3>'+(hist.length?'<div class="box"><ul class="tl">'+hist.slice(0,8).map(tl).join("")+'</ul>'+(hist.length>8?'<details style="margin-top:10px"><summary class="linkbtn">Tout l’historique ('+hist.length+')</summary><ul class="tl" style="margin-top:10px">'+hist.slice(8).map(tl).join("")+'</ul></details>':"")+'</div>':'<p class="empty">Pas encore d’intervention notée.</p>')+'</section>';
  const ml=S.mailsDb.filter(m=>m.dossier===f.id).sort((a,b)=>mailKey(b).localeCompare(mailKey(a)));
  right+='<section><h3>Mails pro</h3><div id="mails">'+(ml.length?'<ul class="list">'+ml.slice(0,10).map(m=>'<li><button class="mail" data-mail="'+esc(m._id)+'" data-edit="mail:'+esc(m._id)+'" style="width:100%;text-align:left;cursor:pointer"><span class="mail-t"><span>'+esc((m.de||"").replace(/<.*>/,"").trim())+'</span><span>'+esc(m.recu||"")+'</span></span><b>'+esc(m.sujet||"(sans objet)")+'</b><p>'+esc((m.apercu||"").replace(/\s+/g," ").slice(0,220))+'</p></button></li>').join("")+'</ul>'+(ml.length>10?'<p class="hint">'+(ml.length-10)+' mail(s) plus ancien(s)</p>':""):'<p class="empty">Aucun mail rattaché à ce dossier. Ta synchro Claude les rattache depuis ta boîte OVH.</p>')+'</div></section>';
  return '<div class="cols" style="margin-top:14px"><div style="display:flex;flex-direction:column;gap:18px">'+left+'</div><div style="display:flex;flex-direction:column;gap:18px">'+right+'</div></div>';
}
const mailKey=m=>{ const r=(m.recu||"").match(/(\d{2})\/(\d{2})\/(\d{4})\s*(\d{2}:\d{2})?/); return r?r[3]+r[2]+r[1]+(r[4]||""):(m.ajoute_le||""); };
function paneDevis(f){
  const id=f.id, dv=docsOf(id,"devis"), lines=ficheDevis(f), pj=projet(f);
  let h='<div class="cols" style="margin-top:14px"><section style="display:flex;flex-direction:column;gap:10px"><div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap"><h3 style="margin:0">Devis</h3>'+(S.canEdit?'<button class="btn" data-devis="'+esc(id)+'">+ Ajouter un devis PDF</button>':"")+'</div>';
  h+=dv.length?'<ul class="list">'+dv.map(d=>{ const sg=/sign|accept/i.test(d.statut||"");
    return '<li class="devis'+(sg?" signe":"")+'" data-edit="devis:'+esc(d.id)+'"><span class="devis-d">'+esc(frDate(d.date)||"—")+'</span><a class="devis-t" href="#" data-viewfile="'+esc(d.id)+'">'+esc(d.intitule||d.titre||d.nom)+'</a><span class="devis-p"><b>'+esc(euro(d.ht))+(d.ht!=null&&d.ht!==""?" HT":"")+'</b><span>'+esc(euro(d.ttc))+(d.ttc?" TTC":"")+'</span></span>'+
      '<span class="devis-s">'+esc(["N° "+(d.numero||"—"),d.statut||"",d.ref_pac||""].filter(Boolean).join(", "))+(S.canEdit?' <button class="linkbtn" data-editdevis="'+esc(d.id)+'">modifier</button>':"")+'</span></li>'; }).join("")+'</ul>'
    :'<p class="empty">Aucun devis PDF. Ajoute le PDF sorti de Codial : je lis la référence PAC et les montants HT/TTC automatiquement.</p>';
  if(lines.length) h+='<details class="box"><summary>Suivi noté dans la fiche</summary><div class="md" style="margin-top:8px"><p><strong>'+esc(pj.intitule||"Projet")+'</strong></p><ul>'+lines.map(r=>'<li>'+inline([col(r,"date")&&frDate(col(r,"date")),col(r,"objet"),col(r,"montant"),col(r,"statut"),col(r,"fournisseur")].filter(Boolean).join(" — "))+'</li>').join("")+'</ul></div></details>';
  h+='</section><section style="display:flex;flex-direction:column;gap:10px"><h3 style="margin:0">Équipement chiffré</h3>';
  const refs=[...new Set(dv.flatMap(d=>String(d.ref_pac||"").split(/\s*\+\s*/)).concat(String(D(id).ref_pac||pj.refs||"").split(/\s*\+\s*/)).filter(Boolean))];
  h+=refs.length?'<ul class="list">'+refs.map(r=>'<li class="box equip"><span class="equip-ic"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="14" cy="12" r="4"/></svg></span><div><b>'+esc(r)+'</b><span>'+esc(pj.type||"")+'</span></div></li>').join("")+'</ul>':'<p class="empty">Les références apparaîtront ici dès qu’un devis est ajouté.</p>';
  return h+'</section></div>';
}
function paneChantier(f){
  const id=f.id, list=docsOf(id,"chantier"), dv=docsOf(id,"devis"), pj=projet(f), asked=S.envois.some(e=>e.type==="dossier-chantier"&&e.statut!=="rangé"&&slug(e.client)===slug(f.nom));
  let h='<div class="cols" style="margin-top:14px"><section style="display:flex;flex-direction:column;gap:10px"><h3 style="margin:0">Dossier chantier</h3><p class="hint">Fiche de préparation à cocher, alimentations électriques constructeur (câbles, disjoncteurs) et schéma de principe, établis à partir du devis.</p>';
  h+=list.length?'<ul class="list">'+list.map(docRow).join("")+'</ul>':'<p class="empty">Pas encore de dossier chantier.</p>';
  if(S.canEdit) h+=asked?'<p class="pending"><strong>Demande envoyée.</strong> Le dossier sera généré à partir de « '+esc(pj.refs||pj.intitule)+' » et déposé ici.</p>'
     :'<button class="primary" data-askchantier="'+esc(id)+'"'+(dv.length||pj.refs?"":" disabled")+'>Générer le dossier chantier</button>'+(dv.length||pj.refs?"":'<p class="hint">Ajoute d’abord le devis (onglet Devis & équipement) pour que les références PAC soient connues.</p>');
  h+='</section><section style="display:flex;flex-direction:column;gap:10px"><h3 style="margin:0">Ajouter un document de chantier</h3>'+dropZone(id,"chantier","Fiche signée, PV de réception, plan d’implantation…")+'</section></div>';
  return h;
}
function paneAdmin(f){
  const id=f.id;
  const g=(cat,t,l)=>'<section style="display:flex;flex-direction:column;gap:10px"><h3 style="margin:0">'+t+'</h3>'+dropZone(id,cat,l)+(docsOf(id,cat).length?'<ul class="list">'+docsOf(id,cat).map(docRow).join("")+'</ul>':'<p class="empty">Aucun document.</p>')+'</section>';
  return '<div class="cols" style="margin-top:14px">'+g("admin","Administratif","Devis signé, attestation TVA, facture, aides (MaPrimeRénov’, CEE), PV de réception…")+g("technique","Technique","Notices constructeur, fiches techniques, certificats, attestation de mise en service…")+'</div>';
}
function panePhotos(f){
  const id=f.id, ph=docsOf(id,"photo"), cover=D(id).photo;
  return '<div style="display:flex;flex-direction:column;gap:12px;margin-top:14px">'+dropZone(id,"photo","Photos de visite, de chantier, de la plaque signalétique…","image/*,video/*")+
    (ph.length?'<div class="photos">'+ph.map(d=>/^video/.test(docMime(d))?'<figure><video'+blobAttr(d.fichier)+' controls playsinline style="width:100%;height:100%;object-fit:cover"></video>'+(S.canEdit?'<figcaption><button data-deldoc="'+esc(d.id)+'">Retirer</button></figcaption>':"")+'</figure>':
      '<figure data-edit="doc:'+esc(d.id)+'"><a href="#" data-viewfile="'+esc(d.id)+'"><img'+blobAttr(d.fichier)+' alt="'+esc(d.nom)+'"></a><figcaption>'+(cover===d.fichier?'<button disabled>Photo du dossier</button>':(S.canEdit?'<button data-setcover="'+esc(d.fichier)+'" data-for="'+esc(id)+'">En couverture</button>':""))+(S.canEdit?'<button data-deldoc="'+esc(d.id)+'">Retirer</button>':"")+'</figcaption></figure>').join("")+'</div>':'<p class="empty">Aucune photo.</p>')+'</div>';
}
function openMail(id){
  const m=findMail(id);
  sheet('<h3>'+esc(m?m.sujet:"Mail")+'</h3><p class="hint">'+esc(m?[m.de,m.recu].filter(Boolean).join(", "):"")+'</p><div class="box" style="white-space:pre-wrap;overflow-wrap:anywhere;max-height:60vh;overflow:auto">'+esc(m?(m.contenu||m.apercu||""):"")+'</div><p class="hint">Résumé relevé par ta synchro Claude. Le mail complet est dans ta boîte OVH.</p><button class="btn ghost" data-close>Fermer</button>');
}

/* ---------- ouverture des fichiers (aperçu PDF, image, Word, Excel) ---------- */
function download(file,name){ const u=URL.createObjectURL(file); const a=document.createElement("a"); a.href=u; a.download=name||"fichier"; document.body.append(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(u),60000); }
async function openFile(docId){
  const d=S.docs.find(x=>x.id===docId); if(!d) return;
  const ext=extOf(d.nom||d.fichier), mime=docMime(d);
  const sh=sheet('<h3>'+esc(d.titre||d.nom)+'</h3><p class="hint"><span class="spin"></span> Ouverture…</p>'); sh.style.maxWidth="980px";
  let file; try{ file=await getBlob(d.fichier,mime); if(!file) throw new Error("absent"); }catch(e){ sh.innerHTML='<h3>'+esc(d.titre||d.nom)+'</h3><p class="conn">'+(S.blobs[d.fichier]?"Impossible de récupérer ce fichier. Réessaie dans un instant.":"Ce fichier vient d’être ajouté et sera disponible dans une minute environ (publication en cours).")+'</p><button class="btn ghost" data-close>Fermer</button>'; return; }
  const app=/xls|csv|ods/.test(ext)?"Excel":/docx?|odt|rtf/.test(ext)?"Word":"l’application";
  const foot='<div style="display:flex;gap:8px;flex-wrap:wrap"><button class="btn" id="fv-dl">'+(/^(pdf|jpe?g|png|webp|gif|mp4|mov|webm)$/.test(ext)?"Télécharger":"Ouvrir dans "+app)+'</button><button class="btn ghost" data-close>Fermer</button></div>';
  const head='<h3>'+esc(d.titre||d.nom)+'</h3><p class="hint">'+esc(CATS[d.cat]||"")+", ajouté le "+esc(frDate((d.ajoute_le||"").slice(0,10)))+'</p>';
  let body='<p class="conn">Pas d’aperçu possible pour ce format ici : ouvre-le avec le bouton ci-dessous.</p>';
  const bind=()=>{ const b=sh.querySelector("#fv-dl"); if(b) b.addEventListener("click",()=>download(file,d.nom)); };
  try{
    if(/^image\//.test(mime) && !/heic|heif/.test(mime)){ const u=await blobUrl(d.fichier,mime); sh.innerHTML=head+'<div class="viewer"><img src="'+esc(u)+'" alt="'+esc(d.nom)+'"></div>'+foot; }
    else if(/^video\//.test(mime)){ const u=await blobUrl(d.fichier,mime); sh.innerHTML=head+'<div class="viewer"><video src="'+esc(u)+'" controls playsinline></video></div>'+foot; }
    else if(ext==="pdf"){ sh.innerHTML=head+'<div class="pdfv" id="fv-pdf" style="max-height:66vh;overflow:auto"><p class="hint"><span class="spin"></span> Mise en page…</p></div>'+foot; bind(); await renderPdf(file,$("fv-pdf")); return; }
    else if(/^(xlsx|xlsm|xls|ods|csv)$/.test(ext)){ await loadScript(XLSX_JS); const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
      const names=wb.SheetNames; const html=n=>XLSX.utils.sheet_to_html(wb.Sheets[n],{header:"",footer:""}).replace(/^[\s\S]*?<table/i,"<table").replace(/<\/table>[\s\S]*$/i,"</table>");
      body=(names.length>1?'<div class="filters" role="tablist">'+names.map((n,i)=>'<button data-xsheet="'+i+'" aria-pressed="'+(i===0)+'">'+esc(n)+'</button>').join("")+'</div>':"")+'<div class="tbl xl" id="fv-x" style="max-height:62vh;overflow:auto">'+sanitize(html(names[0]))+'</div>';
      sh.innerHTML=head+body+foot;
      sh.querySelectorAll("[data-xsheet]").forEach(b=>b.addEventListener("click",()=>{ sh.querySelectorAll("[data-xsheet]").forEach(x=>x.setAttribute("aria-pressed",x===b)); $("fv-x").innerHTML=sanitize(html(names[+b.dataset.xsheet])); }));
    } else if(ext==="docx"){ await loadScript(MAMMOTH_JS); const r=await mammoth.convertToHtml({arrayBuffer:await file.arrayBuffer()});
      body='<div class="box md word" style="max-height:62vh;overflow:auto">'+sanitize(r.value)+'</div>'; sh.innerHTML=head+body+foot;
    } else sh.innerHTML=head+body+foot;
  }catch(e){ sh.innerHTML=head+'<p class="conn">L’aperçu n’a pas pu s’afficher : ouvre le fichier avec le bouton ci-dessous.</p>'+foot; }
  bind();
}
function sanitize(html){ const t=document.createElement("template"); t.innerHTML=html; t.content.querySelectorAll("script,iframe,object,embed,link,meta,style,base,form").forEach(x=>x.remove());
  t.content.querySelectorAll("*").forEach(el=>[...el.attributes].forEach(a=>{ if(/^on/i.test(a.name)||/^\s*(javascript|vbscript):/i.test(a.value)||(a.name==="src"&&!/^data:image\//i.test(a.value))) el.removeAttribute(a.name); }));
  const d=document.createElement("div"); d.append(t.content); return d.innerHTML; }
async function renderPdf(file,el){
  try{ const lib=await pdfLib(); const pdf=await lib.getDocument({data:new Uint8Array(await file.arrayBuffer()),isEvalSupported:false}).promise; el.innerHTML="";
    const w=Math.min(el.clientWidth||800,1400), dpr=Math.min(window.devicePixelRatio||1,2);
    for(let i=1;i<=Math.min(pdf.numPages,60);i++){ if(!el.isConnected) return; const p=await pdf.getPage(i); const v1=p.getViewport({scale:1}); const vp=p.getViewport({scale:(w/v1.width)*dpr});
      const c=document.createElement("canvas"); c.width=vp.width; c.height=vp.height; el.append(c); await p.render({canvasContext:c.getContext("2d"),viewport:vp}).promise; }
    if(pdf.numPages>60) el.insertAdjacentHTML("beforeend",'<p class="hint">Aperçu limité aux 60 premières pages : télécharge le PDF pour la suite.</p>');
  }catch(e){ el.innerHTML='<p class="conn">L’aperçu PDF n’a pas pu s’afficher : télécharge le fichier.</p>'; }
}

/* ---------- clic droit / appui long : modifier ---------- */
const STAGE_OPTS=[["etude","En étude"],["attente","Attente retour client"],["accepte","Accepté"],["chantier","En chantier"],["termine","Terminé"]];
const dossierOpts=sel=>clients().concat(fournisseurs()).sort((a,b)=>a.nom.localeCompare(b.nom,"fr")).map(f=>'<option value="'+esc(f.id)+'"'+(f.id===sel?" selected":"")+'>'+esc(f.nom)+'</option>').join("");
let ctxEl=null, suppressClick=false;
function closeCtx(){ if(ctxEl){ ctxEl.remove(); ctxEl=null; } }
function openCtx(key,x,y,el){
  closeCtx(); if(!key||!S.canEdit) return; const kind=key.split(":")[0];
  const fid=kind==="dossier"||kind==="chantier"||kind==="ligne"?key.split(":")[1]:((el.closest("[data-open]")||{}).dataset||{}).open;
  const m=document.createElement("div"); m.className="ctx"; m.setAttribute("role","menu");
  m.innerHTML='<button role="menuitem" data-ctx="edit">Modifier</button>'+(kind==="chantier"?'<button role="menuitem" data-ctx="dossier">Modifier tout le dossier</button>':"")+(fid&&!lastRoute.startsWith("c-")?'<button role="menuitem" data-ctx="open">Ouvrir le dossier</button>':"")+(["doc","mail","agenda"].includes(kind)?'<button role="menuitem" class="danger" data-ctx="del">Retirer</button>':"");
  document.body.append(m); ctxEl=m;
  const r=m.getBoundingClientRect(); m.style.left=Math.max(8,Math.min(x,innerWidth-r.width-8))+"px"; m.style.top=Math.max(8,Math.min(y,innerHeight-r.height-8))+"px";
  m.addEventListener("click",ev=>{ const b=ev.target.closest("[data-ctx]"); if(!b) return; ev.stopPropagation(); closeCtx(); const a=b.dataset.ctx;
    if(a==="edit") openEdit(key,el); else if(a==="dossier") openEdit("dossier:"+key.split(":")[1],el); else if(a==="open") location.hash="#c-"+fid; else if(a==="del") removeItem(key); });
  m.querySelector("button").focus();
}
document.addEventListener("contextmenu",e=>{ const el=e.target.closest("[data-edit]"); if(!el||!el.dataset.edit||!S.canEdit) return; e.preventDefault(); openCtx(el.dataset.edit,e.clientX,e.clientY,el); });
let lp=null;
document.addEventListener("touchstart",e=>{ const el=e.target.closest("[data-edit]"); if(!el||!el.dataset.edit||!S.canEdit||e.touches.length>1) return; const t=e.touches[0];
  lp=setTimeout(()=>{ lp=null; if(ctxEl) return; suppressClick=true; openCtx(el.dataset.edit,t.clientX,t.clientY,el); if(navigator.vibrate) navigator.vibrate(15); },550); },{passive:true});
["touchmove","touchend","touchcancel"].forEach(ev=>document.addEventListener(ev,()=>{ if(lp){ clearTimeout(lp); lp=null; } },{passive:true}));
document.addEventListener("click",e=>{ if(suppressClick){ suppressClick=false; e.preventDefault(); e.stopPropagation(); return; } if(ctxEl && !ctxEl.contains(e.target)) closeCtx(); },true);
document.addEventListener("keydown",e=>{ if(e.key==="Escape"){ closeCtx(); closeSheet(); } });
window.addEventListener("scroll",closeCtx,{passive:true});

function findAgenda(ref){ const k=ref.replace(/^db:/,""); return S.agendaDb.find(e=>e._id===k); }
function findMail(id){ return S.mailsDb.find(m=>m._id===id||m.id===id); }
const toTime=s=>{ const m=String(s||"").match(/(\d{1,2})[:h](\d{2})/); return m?pad(m[1])+":"+m[2]:""; };
const failToast=e=>toast(e&&e.user?e.message:e&&e.status===401?"Jeton GitHub refusé : préviens Angi":e&&e.status===403?"Écriture refusée par GitHub (droits du jeton)":"Pas enregistré : vérifie ta connexion et réessaie");
function openEdit(key,el){
  const [kind,...rest]=key.split(":"), ref=rest.join(":");
  if(kind==="devis") return openEditDevis(ref);
  if(kind==="dossier"||kind==="chantier"){
    const f=S.fiches.find(x=>x.id===ref); if(!f) return; const d=D(f.id), c=coordOf(f), pj=projet(f), only=kind==="chantier";
    const sh=sheet('<h3>'+esc(f.nom)+(only?" — dates de chantier":"")+'</h3>'+
      (only?"":'<label class="f">Étape<select id="e-etape">'+STAGE_OPTS.map(([k,l])=>'<option value="'+k+'"'+(stageOf(f)===k?" selected":"")+'>'+l+'</option>').join("")+'</select></label><label class="f">Intitulé du projet<input id="e-int" value="'+esc(pj.intitule)+'"></label><div class="row2"><label class="f">Type<select id="e-type">'+["","PAC air/eau","PAC air/air","PAC air/air gainable","Chaudière","Chauffe-eau","Autre"].map(t=>'<option'+(t===pj.type?" selected":"")+'>'+t+'</option>').join("")+'</select></label><label class="f">Réf. PAC<input id="e-ref" value="'+esc(pj.refs)+'"></label></div>')+
      '<div class="row2"><label class="f">Début de chantier<input type="date" id="e-cd" value="'+esc(d.chantier_debut||"")+'"></label><label class="f">Fin de chantier<input type="date" id="e-cf" value="'+esc(d.chantier_fin||"")+'"></label></div>'+
      (only?"":'<div class="row2"><label class="f">Création du dossier<input type="date" id="e-cr" value="'+esc(dateCreation(f))+'"></label><label class="f">Signature<input type="date" id="e-sg" value="'+esc(dateSignature(f))+'"></label></div><label class="f">Contact<input id="e-ct" value="'+esc(c.contact||"")+'"></label><div class="row2"><label class="f">Téléphone<input type="tel" id="e-tel" value="'+esc(c.telephone||"")+'"></label><label class="f">Email<input type="email" id="e-em" value="'+esc(c.email||"")+'"></label></div><label class="f">Adresse<input id="e-ad" value="'+esc(c.adresse||"")+'"></label>')+
      '<button class="primary" id="e-save">Enregistrer</button><button class="btn ghost" data-close>Annuler</button>');
    sh.querySelector("#e-save").addEventListener("click",async ev=>{
      const g=k=>{ const x=$("e-"+k); return x?x.value.trim():undefined; }; const patch={}, notes=[];
      const cd=g("cd"), cf=g("cf"); if(cf && cd && cf<cd){ toast("La fin est avant le début"); return; }
      if(cd!==(d.chantier_debut||"")||cf!==(d.chantier_fin||"")){ patch.chantier_debut=cd; patch.chantier_fin=cf; notes.push(cd?"Chantier prévu "+(cf&&cf!==cd?"du "+cd+" au "+cf:"le "+cd):"Dates de chantier retirées"); }
      let coordMsg="";
      if(!only){
        const et=g("etape"); if(et!==stageOf(f)){ patch.etape=et; if(et==="accepte"&&!g("sg")) patch.date_signature=today(); notes.push("Étape du dossier : "+STAGE_L[et]); }
        if(g("int")!==pj.intitule) patch.intitule=g("int"); if(g("type")!==pj.type) patch.type_projet=g("type"); if(g("ref")!==pj.refs) patch.ref_pac=g("ref");
        if(g("cr")!==dateCreation(f)) patch.date_creation=g("cr"); if(g("sg")!==dateSignature(f)){ patch.date_signature=g("sg"); if(g("sg")) notes.push("Devis signé le "+g("sg")); }
        const cc=[]; [["ct","contact","Contact"],["tel","telephone","Téléphone"],["em","email","Email"],["ad","adresse","Adresse"]].forEach(([k,f2,l])=>{ if(g(k)!==(c[f2]||"")){ patch[f2]=g(k); cc.push(l+" : "+(g(k)||"(vide)")); } });
        if(cc.length) coordMsg="Coordonnées corrigées dans le site : "+cc.join(" ; ");
      }
      if(!Object.keys(patch).length){ closeSheet(); return; }
      ev.target.disabled=true; ev.target.innerHTML='<span class="spin"></span> Enregistrement…';
      try{ await setDossier(f.id,patch);
        if(coordMsg) await envoyer({client:f.nom,fiche:f.chemin,type:"fiche",texte:coordMsg});
        if(notes.length) await envoyer({client:f.nom,fiche:f.chemin,type:"historique",texte:"- "+today()+" — "+notes.join(" ; ")});
        closeSheet(); toast("Modifié"); rerender(); }catch(e){ ev.target.disabled=false; ev.target.textContent="Enregistrer"; failToast(e); } });
    return;
  }
  if(kind==="agenda"){
    const e=findAgenda(ref); if(!e) return; const dt=evDate(e), allDay=/journ/i.test(e.debut||""), f=evClient(e);
    const sh=sheet('<h3>Rendez-vous</h3><label class="f">Objet<input id="e-suj" value="'+esc(e.sujet||"")+'"></label><div class="row2"><label class="f">Date<input type="date" id="e-dt" value="'+esc(dt)+'"></label><label class="f">Jusqu’au<input type="date" id="e-df" value="'+esc(allDay?evEnd(e):"")+'"><small>Pour plusieurs jours</small></label></div><div class="row2"><label class="f">Début<input type="time" id="e-h1" value="'+esc(allDay?"":toTime(e.debut))+'"></label><label class="f">Fin<input type="time" id="e-h2" value="'+esc(allDay?"":toTime(e.fin))+'"></label></div><label class="f">Lieu<input id="e-lieu" value="'+esc(e.lieu||"")+'"></label><label class="f">Dossier<select id="e-dos"><option value="">Aucun</option>'+dossierOpts(f&&f.id)+'</select></label><p class="hint">La modification est faite dans le site. Ton agenda Outlook garde l’ancienne version : pense à le mettre à jour aussi.</p><button class="primary" id="e-save">Enregistrer</button><button class="btn ghost" data-close>Annuler</button>');
    sh.querySelector("#e-save").addEventListener("click",async()=>{
      const d2=$("e-dt").value; if(!d2){ toast("Indique la date"); return; } const h1=$("e-h1").value, h2=$("e-h2").value, df=$("e-df").value; const fr=x=>x.split("-").reverse().join("/");
      const row={id:e.id||("app-"+Date.now()),sujet:$("e-suj").value.trim(),lieu:$("e-lieu").value.trim(),dossier:$("e-dos").value,ajoute_le:new Date().toISOString(),source:"modifié dans le site",
        debut:h1?fr(d2)+" "+h1:"journée entière", fin:h2?fr(d2)+" "+h2:"", jours:h1?"":d2+" -> "+(df&&df>=d2?df:d2),doc_id:e.doc_id||e._id,par:S.me.login};
      try{ await setRows("agenda.json",arr=>{ const i=arr.findIndex(x=>(x.doc_id||x.id)===e._id); if(i>=0) arr[i]=Object.assign({},arr[i],row); else arr.push(row); return arr; },"Agenda : "+row.sujet);
        closeSheet(); toast("Rendez-vous modifié"); rerender(); }catch(err){ failToast(err); } });
    return;
  }
  if(kind==="rappel"){
    const raw=S.rappelsRaw.find(r=>r.id===ref), cur=S.rappels.find(r=>r.id===ref); if(!raw||!cur) return;
    const sh=sheet('<h3>Rappel — '+esc(cur.client||"")+'</h3><label class="f">Date<input type="date" id="e-dt" value="'+esc(/^\d{4}/.test(cur.date||"")?cur.date:"")+'"><small>Vide = sans date</small></label><label class="f">Quoi faire<textarea id="e-act">'+esc(cur.action||"")+'</textarea></label><button class="primary" id="e-save">Enregistrer</button><button class="btn ghost" data-close>Annuler</button>');
    sh.querySelector("#e-save").addEventListener("click",async()=>{ const nd=$("e-dt").value, na=$("e-act").value.trim(); if(!na){ toast("Le rappel est vide"); return; }
      try{ requireEdit(); const nv=await updateJson("_app/surcharges.json",{},o=>{ o[ref]={date:nd,action:na,old_date:raw.date||"",old_action:raw.action||"",maj:new Date().toISOString()}; return o; },"Rappel modifié : "+raw.client);
        if(nv){ S.over=nv; applyOver(); }
        await envoyer({client:raw.client,fiche:raw.fiche,type:"rappel",rappel_modifie:ref,texte:"Modifier le rappel « "+(raw.date||raw.date_texte||"(sans date)")+" — "+raw.action+" » en : - [ ] "+(nd||"(sans date)")+" — "+na});
        closeSheet(); toast("Rappel modifié"); rerender(); }catch(err){ failToast(err); } });
    return;
  }
  if(kind==="doc"){
    const d=S.docs.find(x=>x.id===ref); if(!d) return;
    const sh=sheet('<h3>Document</h3><label class="f">Nom affiché<input id="e-nom" value="'+esc(d.titre||d.nom)+'"></label><div class="row2"><label class="f">Rubrique<select id="e-cat">'+Object.entries(CATS).map(([k,l])=>'<option value="'+k+'"'+(k===d.cat?" selected":"")+'>'+l+'</option>').join("")+'</select></label><label class="f">Date<input type="date" id="e-dt" value="'+esc(d.date||"")+'"></label></div><label class="f">Dossier<select id="e-dos">'+dossierOpts(d.dossier)+'</select></label><button class="primary" id="e-save">Enregistrer</button><button class="btn ghost" data-close>Annuler</button>');
    sh.querySelector("#e-save").addEventListener("click",async()=>{ const p={titre:$("e-nom").value.trim()||d.nom,cat:$("e-cat").value,date:$("e-dt").value,dossier:$("e-dos").value};
      try{ await setDocs(arr=>{ const x=arr.find(y=>y.id===ref); if(!x) return undefined; Object.assign(x,p); return arr; },"Document modifié : "+p.titre); closeSheet(); toast("Document modifié"); rerender(); }catch(err){ failToast(err); } });
    return;
  }
  if(kind==="mail"){
    const m=findMail(ref); if(!m) return; const cur=m.dossier||(lastRoute.startsWith("c-")?lastRoute.slice(2):"");
    const sh=sheet('<h3>'+esc(m.sujet||"Mail")+'</h3><p class="hint">'+esc([m.de,m.recu].filter(Boolean).join(", "))+'</p><label class="f">Rattaché au dossier<select id="e-dos">'+dossierOpts(cur)+'</select></label><label class="f">Résumé<textarea id="e-ap">'+esc(m.apercu||"")+'</textarea></label><button class="primary" id="e-save">Enregistrer</button><button class="btn ghost" data-close>Annuler</button>');
    sh.querySelector("#e-save").addEventListener("click",async()=>{ const row={dossier:$("e-dos").value,apercu:$("e-ap").value.trim()};
      try{ await setRows("mails.json",arr=>{ const x=arr.find(y=>(y.doc_id||y.id)===m._id); if(!x) return undefined; Object.assign(x,row); return arr; },"Mail rattaché : "+(m.sujet||""));
        closeSheet(); toast("Mail rattaché"); rerender(); }catch(err){ failToast(err); } });
    return;
  }
  if(kind==="ligne"){
    const f=S.fiches.find(x=>x.id===ref); if(!f) return; const old=el.dataset.txt||"";
    const sh=sheet('<h3>Corriger — '+esc(f.nom)+'</h3><label class="f">Texte<textarea id="e-txt">'+esc(old)+'</textarea></label><p class="hint">La correction est rangée dans la fiche Obsidian à la synchro de '+nextSync()+'.</p><button class="primary" id="e-save">Envoyer la correction</button><button class="btn ghost" data-close>Annuler</button>');
    sh.querySelector("#e-save").addEventListener("click",async()=>{ const nv=$("e-txt").value.trim(); if(nv===old){ closeSheet(); return; }
      try{ await envoyer({client:f.nom,fiche:f.chemin,type:"fiche",texte:"Corriger dans la fiche : « "+old+" » → « "+(nv||"(supprimer la ligne)")+" »"}); closeSheet(); toast("Correction envoyée"); rerender(); }catch(err){ failToast(err); } });
  }
}
async function removeItem(key){
  const [kind,...rest]=key.split(":"), ref=rest.join(":");
  if(!confirm(kind==="mail"?"Retirer ce mail du dossier ? (il reste dans ta boîte)":"Retirer cet élément du site ?")) return;
  try{
    if(kind==="doc") await removeDoc(ref);
    else if(kind==="mail"){ const m=findMail(ref); if(m) await setRows("mails.json",arr=>arr.filter(x=>(x.doc_id||x.id)!==m._id),"Mail retiré du dossier : "+(m.sujet||"")); }
    else if(kind==="agenda"){ const e=findAgenda(ref); if(e) await setRows("agenda.json",arr=>arr.filter(x=>(x.doc_id||x.id)!==e._id),"Rendez-vous retiré : "+(e.sujet||"")); }
    toast("Retiré"); rerender();
  }catch(e){ failToast(e); }
}

/* ---------- feuilles modales ---------- */
function sheet(html){ closeSheet(); const bg=document.createElement("div"); bg.className="sheet-bg"; bg.id="sheetbg"; bg.innerHTML='<div class="sheet" role="dialog" aria-modal="true">'+html+'</div>';
  bg.addEventListener("click",e=>{ if(e.target===bg||e.target.closest("[data-close]")) closeSheet(); }); document.body.append(bg); return bg.firstChild; }
function closeSheet(){ const b=$("sheetbg"); if(b){ b.remove(); if(pendingRerender){ pendingRerender=false; rerender(); } } }
function devisForm(v){ v=v||{};
  return '<label class="f">Intitulé du projet<input id="dv-int" value="'+esc(v.intitule||"")+'" placeholder="PAC air/eau Panasonic 9 kW + ballon ECS"></label>'+
    '<div class="row2"><label class="f">Type<select id="dv-type">'+["","PAC air/eau","PAC air/air","PAC air/air gainable","Chaudière","Chauffe-eau","Autre"].map(t=>'<option'+(t===(v.type_projet||"")?" selected":"")+'>'+t+'</option>').join("")+'</select></label><label class="f">Réf. PAC chiffrée<input id="dv-ref" value="'+esc(v.ref_pac||"")+'"></label></div>'+
    '<div class="row2"><label class="f">N° de devis<input id="dv-num" value="'+esc(v.numero||"")+'"></label><label class="f">Date<input type="date" id="dv-date" value="'+esc(v.date||today())+'"></label></div>'+
    '<div class="row2"><label class="f">Total HT (€)<input id="dv-ht" inputmode="decimal" value="'+esc(v.ht??"")+'"></label><label class="f">Total TTC (€)<input id="dv-ttc" inputmode="decimal" value="'+esc(v.ttc??"")+'"></label></div>'+
    '<label class="f">Statut<select id="dv-st">'+["envoyé","en préparation","signé","refusé"].map(s=>'<option'+(s===(v.statut||"envoyé")?" selected":"")+'>'+s+'</option>').join("")+'</select></label>';
}
function readDevisForm(){ const g=k=>($("dv-"+k)||{}).value||""; return {intitule:g("int").trim(),type_projet:g("type"),ref_pac:g("ref").trim(),numero:g("num").trim(),date:g("date"),ht:num(g("ht")),ttc:num(g("ttc")),statut:g("st")}; }
function openDevisSheet(id){
  const f=S.fiches.find(x=>x.id===id); let file=null;
  const sh=sheet('<h3>Ajouter un devis</h3><p class="hint">Choisis le PDF du devis (Codial). Je remplis les champs, tu vérifies.</p><label class="btn" style="align-self:flex-start">Choisir le PDF<input type="file" accept="application/pdf" id="dv-file" hidden></label><p class="hint" id="dv-step"></p><div id="dv-fields">'+devisForm({})+'</div><button class="primary" id="dv-save" disabled>Enregistrer le devis</button><button class="btn ghost" data-close>Annuler</button>');
  sh.querySelector("#dv-file").addEventListener("change",async e=>{ file=e.target.files[0]; if(!file) return; $("dv-save").disabled=true;
    const step=$("dv-step"); step.innerHTML='<span class="spin"></span> '+esc(file.name);
    const v=await lireDevis(file,s=>{ step.innerHTML='<span class="spin"></span> '+esc(s); });
    $("dv-fields").innerHTML=devisForm(v); step.textContent=v.err?v.err:"Vérifie les informations lues dans « "+file.name+" »."; $("dv-save").disabled=false; });
  sh.querySelector("#dv-save").addEventListener("click",async()=>{ if(!file) return; const v=readDevisForm(); const b=$("dv-save"); b.disabled=true; b.innerHTML='<span class="spin"></span> Enregistrement…';
    if(v.statut==="signé") v.signe_le=today();
    const out=await addDocs(id,"devis",[file],Object.assign({titre:v.intitule||file.name},v));
    if(!out.length){ b.disabled=false; b.textContent="Enregistrer le devis"; return; }
    try{ const d=D(id), patch={}; if(!d.intitule&&v.intitule) patch.intitule=v.intitule; if(!d.ref_pac&&v.ref_pac) patch.ref_pac=v.ref_pac; if(!d.type_projet&&v.type_projet) patch.type_projet=v.type_projet;
      if(v.statut==="signé"){ patch.etape="accepte"; patch.date_signature=d.date_signature||today(); } else if(!d.etape||d.etape==="etude") patch.etape="attente";
      await setDossier(id,patch);
      await envoyer({client:f.nom,fiche:f.chemin,type:"devis",texte:"| "+(v.date||today())+" | "+(v.numero?"Devis Codial n°"+v.numero+" — ":"")+(v.intitule||"")+(v.ref_pac?" ("+v.ref_pac+")":"")+" | "+(v.ht!=null?euro(v.ht)+" HT":"")+(v.ttc!=null?" ("+euro(v.ttc)+" TTC)":"")+" | "+v.statut+" | |\nPDF joint dans le site CRM."});
    }catch(e){ failToast(e); }
    closeSheet(); rerender(); });
}
function openEditDevis(docId){
  const d=S.docs.find(x=>x.id===docId); if(!d) return;
  const sh=sheet('<h3>Modifier le devis</h3>'+devisForm(d)+'<button class="primary" id="dv-save">Enregistrer</button><button class="btn ghost" data-close>Annuler</button>');
  sh.querySelector("#dv-save").addEventListener("click",async()=>{ const v=readDevisForm(); if(v.statut==="signé"&&!d.signe_le) v.signe_le=today();
    try{ await setDocs(arr=>{ const x=arr.find(y=>y.id===docId); if(!x) return undefined; Object.assign(x,{titre:v.intitule||d.nom},v); return arr; },"Devis modifié : "+(v.numero||d.nom));
      if(v.statut==="signé"&&!/sign/.test(d.statut||"")){ const f=S.fiches.find(x=>x.id===d.dossier); await setDossier(d.dossier,{etape:"accepte",date_signature:D(d.dossier).date_signature||today()}); if(f) await envoyer({client:f.nom,fiche:f.chemin,type:"devis",texte:"Devis "+(v.numero?"n°"+v.numero+" ":"")+"signé le "+today()+"."}); }
      closeSheet(); toast("Devis mis à jour"); rerender(); }catch(e){ failToast(e); } });
}
function openEditProj(id){
  const pj=projet(S.fiches.find(x=>x.id===id));
  const sh=sheet('<h3>Projet</h3><label class="f">Intitulé<input id="pj-int" value="'+esc(pj.intitule)+'"></label><label class="f">Type<select id="pj-type">'+["","PAC air/eau","PAC air/air","PAC air/air gainable","Chaudière","Chauffe-eau","Autre"].map(t=>'<option'+(t===pj.type?" selected":"")+'>'+t+'</option>').join("")+'</select></label><label class="f">Réf. PAC chiffrée<input id="pj-ref" value="'+esc(pj.refs)+'"></label><button class="primary" id="pj-save">Enregistrer</button><button class="btn ghost" data-close>Annuler</button>');
  sh.querySelector("#pj-save").addEventListener("click",async()=>{ try{ await setDossier(id,{intitule:$("pj-int").value.trim(),type_projet:$("pj-type").value,ref_pac:$("pj-ref").value.trim()}); closeSheet(); rerender(); }catch(e){ failToast(e); } });
}

/* ---------- NOTER ---------- */
function vAjouter(pre){
  setHead("Noter quelque chose"); setTab("ajouter");
  if(!S.canEdit){ view.innerHTML='<p class="empty">Ton compte est en lecture seule : demande à Angi de te donner le droit de modifier.</p>'; return; }
  if(pre) S.form=Object.assign({},S.form,pre);
  const F=S.form, list=clients().concat(fournisseurs()).sort((a,b)=>a.nom.localeCompare(b.nom,"fr")), isNew=F.type==="fiche";
  let h='<form id="form" novalidate style="max-width:640px"><div class="seg" role="group" aria-label="Type">'+["rappel","historique","installation","fiche"].map(t=>'<button type="button" data-type="'+t+'" aria-pressed="'+(F.type===t)+'">'+({rappel:"Rappel",historique:"Note",installation:"Matériel posé",fiche:"Nouveau dossier"})[t]+'</button>').join("")+'</div>';
  if(!isNew) h+='<label class="f">Dossier<select id="f-cli"><option value="">Choisir…</option>'+list.map(f=>'<option value="'+esc(f.id)+'"'+(F.cli===f.id?" selected":"")+'>'+esc(f.nom)+'</option>').join("")+'</select></label>';
  if(F.type==="rappel"){
    h+='<label class="f">Quand<input type="date" id="f-date" value="'+esc(F.date||"")+'"><small>Vide = sans date.</small></label><div class="seg">'+[["Demain",1],["Dans 3 j",3],["Dans 1 sem.",7]].map(([l,n])=>'<button type="button" data-plus="'+n+'">'+l+'</button>').join("")+'</div>';
    h+='<label class="f">Quoi faire<textarea id="f-txt" placeholder="Ex. : relancer sur le devis n°…">'+esc(F.txt||"")+'</textarea></label>';
  } else if(F.type==="installation"){
    h+='<label class="f">Équipement<input id="f-eq" placeholder="PAC air/eau" value="'+esc(F.eq||"")+'"></label><label class="f">Marque / modèle<input id="f-mo" placeholder="Panasonic WH-UDZ09KE5" value="'+esc(F.mo||"")+'"></label>';
    h+='<div class="row2"><label class="f">N° de série<input id="f-sn" value="'+esc(F.sn||"")+'"></label><label class="f">Date de pose<input type="date" id="f-ms" value="'+esc(F.ms||today())+'"></label></div><label class="f">Autres infos<textarea id="f-txt" placeholder="Puissance, fluide, garantie…">'+esc(F.txt||"")+'</textarea></label>';
  } else if(isNew){
    h+='<div class="seg"><button type="button" data-kind="client" aria-pressed="'+(F.kind!=="fournisseur")+'">Client</button><button type="button" data-kind="fournisseur" aria-pressed="'+(F.kind==="fournisseur")+'">Fournisseur</button></div>';
    h+='<label class="f">Nom du dossier<input id="f-nom" placeholder="DUPONT ou Cabinet XYZ" value="'+esc(F.nom||"")+'"></label><label class="f">Contact<input id="f-ct" value="'+esc(F.ct||"")+'"></label>';
    h+='<div class="row2"><label class="f">Téléphone<input type="tel" id="f-tel" value="'+esc(F.tel||"")+'"></label><label class="f">Email<input type="email" id="f-em" value="'+esc(F.em||"")+'"></label></div><label class="f">Adresse<input id="f-ad" value="'+esc(F.ad||"")+'"></label><label class="f">Projet<textarea id="f-txt" placeholder="Ex. : remplacement chaudière fioul par PAC air/eau">'+esc(F.txt||"")+'</textarea></label>';
  } else {
    h+='<label class="f">Date<input type="date" id="f-date" value="'+esc(F.date||today())+'"></label><label class="f">Ce qui s’est passé<textarea id="f-txt" placeholder="Visite technique, relevés, accord du client…">'+esc(F.txt||"")+'</textarea></label>';
  }
  h+='<button class="primary" type="submit" id="go">Envoyer au CRM</button><p class="hint">Visible tout de suite dans le dossier, rangé dans la fiche Obsidian à '+nextSync()+'. Les devis et documents s’ajoutent directement dans le dossier.</p></form>';
  view.innerHTML=h;
}
function readForm(){ ["cli","date","txt","eq","mo","sn","ms","nom","ct","tel","em","ad"].forEach(k=>{ const el=$("f-"+k); if(el) S.form[k]=el.value; }); }
async function submitForm(){
  readForm(); const F=S.form, txt=(F.txt||"").trim(); let doc;
  if(F.type==="fiche"){ const nom=(F.nom||"").trim(); if(!nom){ toast("Donne un nom au dossier"); return; } const kind=F.kind==="fournisseur"?"fournisseur":"client";
    const l=[["Contact",F.ct],["Téléphone",F.tel],["Email",F.em],["Adresse",F.ad]].filter(([k,v])=>(v||"").trim()).map(([k,v])=>"- "+k+" : "+v.trim());
    doc={client:nom,fiche:(kind==="client"?"Clients/":"Fournisseurs/")+nom,type:kind==="client"?"fiche":"fournisseur",texte:"Nouvelle fiche "+kind+"\n"+l.join("\n")+(txt?"\n- Projet : "+txt:"")};
  } else { const f=S.fiches.find(x=>x.id===F.cli); if(!f){ toast("Choisis le dossier"); return; } const base={client:f.nom,fiche:f.chemin};
    if(F.type==="rappel"){ if(!txt){ toast("Écris le rappel"); return; } doc=Object.assign(base,{type:"rappel",date:F.date||"",texte:"- [ ] "+(F.date||"(sans date)")+" — "+txt}); }
    else if(F.type==="installation"){ if(!F.eq&&!F.mo){ toast("Indique l’équipement"); return; } doc=Object.assign(base,{type:"installation",texte:[["Équipement",F.eq],["Marque / modèle",F.mo],["N° de série",F.sn],["Mise en service",F.ms],["Autres",txt]].filter(([k,v])=>(v||"").trim()).map(([k,v])=>k+" : "+v.trim()).join(" ; ")}); }
    else { if(!txt){ toast("Écris la note"); return; } doc=Object.assign(base,{type:"historique",texte:"- "+(F.date||today())+" — "+txt}); }
  }
  const go=$("go"); go.disabled=true; go.textContent="Envoi…";
  try{ await envoyer(doc); toast("Envoyé, rangé à "+nextSync()); const back=F.cli; S.form={type:F.type}; location.hash=(doc.type!=="fiche"&&doc.type!=="fournisseur"&&back)?"#c-"+back:"#envois"; }
  catch(e){ go.disabled=false; go.textContent="Envoyer au CRM"; failToast(e); }
}

/* ---------- ENVOIS ---------- */
function vEnvois(){
  setHead("Boîte d’envoi"); setTab("envois");
  const pend=S.envois.filter(e=>e.statut!=="rangé");
  const it=e=>'<li class="env"><div class="env-top"><b>'+esc(e.client||"Sans client")+'</b><span>'+esc(TYPES[e.type]||e.type||"")+'</span><span>'+esc(frStamp(e.cree))+(e.source?", "+esc(e.source):"")+(nameOf(e.par)?", "+esc(nameOf(e.par)):"")+'</span></div><p>'+esc(e.texte)+'</p>'+(S.canEdit&&e.source===SOURCE&&mine(e)?'<button class="linkbtn" data-undo="'+esc(e._id)+'">Annuler cet envoi</button>':"")+'</li>';
  view.innerHTML='<p class="hint">Ce que l’équipe note ici, ou que Claude note pour vous, passe par cette boîte (dossier _inbox du coffre) avant d’être rangé dans les fiches Obsidian par ta synchro Claude, qui le retire ensuite de la boîte. Prochaine synchro : '+nextSync()+'.</p><section><h2>À ranger <small>'+pend.length+'</small></h2>'+(pend.length?'<ul class="list">'+pend.map(it).join("")+'</ul>':'<p class="empty">Rien en attente.</p>')+'</section>';
}

/* ---------- ADMINISTRATION ---------- */
async function fetchUsers(){ const r=await fetch("users.json?t="+Date.now(),{cache:"no-store"}); if(!r.ok) throw new Error("Comptes indisponibles ("+r.status+")"); return r.json(); }
async function readAnnuaire(users,key){ if(!users||!users.annuaire) return []; try{ return await C.decJson(key,users.annuaire,C.AAD.annuaire); }catch(e){ return []; } }
const passMeter = id => '<div class="meter" aria-hidden="true"><i id="'+id+'-m" style="width:0"></i></div><small id="'+id+'-h">Au moins '+C.MIN_PASS+' caractères. Une phrase de plusieurs mots est idéale.</small>';
function bindMeter(id){ const el=$(id); if(!el) return; el.addEventListener("input",()=>{ const n=[...el.value].length, ok=n>=C.MIN_PASS; const m=$(id+"-m"); m.style.width=Math.min(100,n/C.MIN_PASS*100)+"%"; m.classList.toggle("ok",ok); $(id+"-h").textContent=ok?n+" caractères : OK":n+" / "+C.MIN_PASS+" caractères minimum"; }); }
async function ensureRawKey(){
  if(S.rawKey) return true;
  return new Promise(res=>{ const sh=sheet('<h3>Confirme ton mot de passe</h3><p class="hint">Pour créer ou modifier des comptes, la clé de données doit être déverrouillée avec ton mot de passe.</p><label class="f">Mot de passe<input type="password" id="rk-p" autocomplete="current-password"></label><p class="err" id="rk-e" hidden></p><button class="primary" id="rk-go">Déverrouiller</button><button class="btn ghost" data-close>Annuler</button>');
    let done=false; const obs=new MutationObserver(()=>{ if(!sh.isConnected&&!done){ done=true; obs.disconnect(); res(false); } }); obs.observe(document.body,{childList:true});
    $("rk-go").addEventListener("click",async()=>{ const b=$("rk-go"); b.disabled=true; b.innerHTML='<span class="spin"></span> Vérification…';
      try{ const u=await fetchUsers(); const r=await C.unwrap(u,S.me.login,$("rk-p").value); S.rawKey=r.raw; done=true; obs.disconnect(); closeSheet(); res(true); }
      catch(e){ $("rk-e").hidden=false; $("rk-e").textContent=e.message; b.disabled=false; b.textContent="Déverrouiller"; } }); });
}
async function vAdmin(){
  setHead("Administration",true); setTab("");
  if(!S.me||S.me.role!=="admin"){ view.innerHTML='<p class="empty">Cette page est réservée à l’administrateur.</p>'; return; }
  view.innerHTML='<p class="status"><span class="spin"></span> Lecture des comptes…</p>';
  let users, ann;
  try{ const cur=await GH.get("_site/users.json"); users=cur?JSON.parse(cur.text):await fetchUsers(); ann=await readAnnuaire(users,S.key); }
  catch(e){ try{ users=await fetchUsers(); ann=await readAnnuaire(users,S.key); }catch(x){ view.innerHTML='<p class="err">Impossible de lire les comptes : '+esc(e.message)+'</p>'; return; } }
  S.names={}; ann.forEach(a=>{ S.names[a.login]=a.nom||a.login; S.names[a.id]=a.nom||a.login; });
  const ids=new Set((users.comptes||[]).map(c=>c.id));
  const rows=ann.slice().sort((a,b)=>(a.nom||a.login).localeCompare(b.nom||b.login,"fr"));
  let h='<section><h2>Comptes <small>'+rows.length+'</small></h2><p class="hint">Les mots de passe ne sont jamais stockés : chaque compte contient la clé de données, chiffrée avec son mot de passe (PBKDF2-SHA256, '+C.ITER.toLocaleString("fr-FR")+' itérations, sel unique). Les changements sont publiés en une minute environ.</p><ul class="list" style="margin-top:10px">'+
    rows.map(a=>'<li class="user"><span class="av">'+esc(initials(a.nom||a.login))+'</span><div class="user-b"><b>'+esc(a.nom||a.login)+'</b><span>'+esc(a.login)+(a.maj?" · modifié le "+esc(frDate(a.maj)):"")+(ids.has(a.id)?"":" · clé absente")+'</span></div><span class="role '+esc(a.role)+'">'+esc(ROLES[a.role]||a.role)+'</span><button class="btn ghost" data-uedit="'+esc(a.id)+'">Modifier</button></li>').join("")+'</ul></section>'+
    '<section><button class="primary" data-unew="1" style="width:100%;max-width:420px">+ Créer un compte</button></section>'+
    '<section class="box"><h3>Révoquer vraiment un accès</h3><p class="hint">Supprimer un compte l’empêche de se connecter. Mais une personne qui a déjà eu accès a pu garder la clé de données ou le jeton GitHub : en cas de départ ou de doute, change aussi le jeton (CRM_SITE_TOKEN) et la clé (DATA_KEY), puis recrée les comptes (voir INSTALLATION.md).</p></section>';
  view.innerHTML=h;
  view.querySelectorAll("[data-uedit]").forEach(b=>b.addEventListener("click",async()=>{ if(await ensureRawKey()) userSheet(ann.find(a=>a.id===b.dataset.uedit)); }));
  view.querySelector("[data-unew]").addEventListener("click",async()=>{ if(await ensureRawKey()) userSheet(null); });
}
function userSheet(a){
  const isNew=!a;
  const sh=sheet('<h3>'+(isNew?"Nouveau compte":"Modifier "+esc(a.nom||a.login))+'</h3>'+
    '<label class="f">Identifiant<input id="u-login" autocomplete="off" autocapitalize="none" spellcheck="false" value="'+esc(a?a.login:"")+'"><small>Sert à se connecter (sans majuscules ni espaces de préférence).</small></label>'+
    '<label class="f">Nom affiché<input id="u-nom" value="'+esc(a?a.nom||"":"")+'"></label>'+
    '<label class="f">Rôle<select id="u-role">'+Object.entries(ROLES).map(([k,l])=>'<option value="'+k+'"'+((a?a.role:"edition")===k?" selected":"")+'>'+l+'</option>').join("")+'</select><small>Lecture seule : voit tout, ne modifie rien. Modification : peut tout modifier sauf les comptes.</small></label>'+
    '<label class="f">'+(isNew?"Mot de passe":"Nouveau mot de passe")+'<input type="password" id="u-p1" autocomplete="new-password">'+(isNew?"":'<small>Laisse vide pour le garder (obligatoire si tu changes l’identifiant).</small>')+passMeter("u-p1")+'</label>'+
    '<label class="f">Confirmation<input type="password" id="u-p2" autocomplete="new-password"></label>'+
    '<p class="err" id="u-e" hidden></p><button class="primary" id="u-save">'+(isNew?"Créer le compte":"Enregistrer")+'</button>'+
    (isNew?"":'<button class="btn ghost" id="u-del" style="color:var(--hot);border-color:var(--hot)">Supprimer ce compte</button>')+'<button class="btn ghost" data-close>Annuler</button>');
  bindMeter("u-p1");
  const err=m=>{ const e=$("u-e"); e.hidden=!m; e.textContent=m||""; };
  $("u-save").addEventListener("click",async()=>{
    const login=$("u-login").value.trim(), nom=$("u-nom").value.trim(), role=$("u-role").value, p1=$("u-p1").value, p2=$("u-p2").value;
    if(!login||/\s/.test(login)){ err("Identifiant obligatoire, sans espace."); return; }
    const loginChanged=!isNew && login.toLowerCase()!==a.login.toLowerCase();
    if(isNew||p1||loginChanged){ const e=C.checkPass(p1); if(e){ err(e); return; } if(p1!==p2){ err("Les deux mots de passe ne sont pas identiques."); return; } }
    if(!S.rawKey){ err("Clé de données verrouillée : rouvre cette fenêtre."); return; }
    const b=$("u-save"); b.disabled=true; b.innerHTML='<span class="spin"></span> Chiffrement…';
    try{
      const newId=await C.userId(login); const wrapped=(isNew||p1||loginChanged)?await C.wrap(login,p1,S.rawKey):null;
      await saveUsers(async(users,ann)=>{
        if(ann.some(x=>x.id===newId && (isNew||x.id!==a.id))) throw Object.assign(new Error("Cet identifiant existe déjà."),{user:true});
        if(!isNew && a.role==="admin" && role!=="admin" && ann.filter(x=>x.role==="admin").length<2) throw Object.assign(new Error("Il faut garder au moins un administrateur."),{user:true});
        const now=new Date().toISOString();
        if(!isNew){ users.comptes=users.comptes.filter(c=>c.id!==a.id||(!wrapped&&c.id===a.id)); ann=ann.filter(x=>x.id!==a.id); }
        if(wrapped){ users.comptes=users.comptes.filter(c=>c.id!==wrapped.id); users.comptes.push(wrapped); }
        ann.push({id:newId,login,nom:nom||login,role,cree:a?a.cree:now,maj:now});
        return ann;
      },(isNew?"Compte créé : ":"Compte modifié : ")+login);
      if(!isNew && a.id===S.me.id){ S.me.role=role; S.me.nom=nom||login; if(loginChanged){ S.me.login=login; S.me.id=newId; } }
      closeSheet(); toast(isNew?"Compte créé":"Compte modifié"); vAdmin();
    }catch(e){ b.disabled=false; b.textContent=isNew?"Créer le compte":"Enregistrer"; err(e.user?e.message:"Pas enregistré : "+(e.message||"erreur")); }
  });
  const del=$("u-del"); if(del) del.addEventListener("click",async()=>{
    if(a.id===S.me.id){ err("Tu ne peux pas supprimer ton propre compte."); return; }
    if(!confirm("Supprimer le compte « "+(a.nom||a.login)+" » ? La personne ne pourra plus se connecter.")) return;
    try{ await saveUsers(async(users,ann)=>{ users.comptes=users.comptes.filter(c=>c.id!==a.id); return ann.filter(x=>x.id!==a.id); },"Compte supprimé : "+a.login); closeSheet(); toast("Compte supprimé"); vAdmin(); }
    catch(e){ err(e.user?e.message:"Pas supprimé : "+(e.message||"erreur")); }
  });
}
/* Lit _site/users.json à jour, déchiffre l’annuaire, applique fn, rechiffre et réécrit. */
function saveUsers(fn,message){
  requireEdit();
  return enqueue(async()=>{
    for(let i=0;i<4;i++){
      const cur=await GH.get("_site/users.json"); if(!cur) throw new Error("_site/users.json introuvable dans le dépôt");
      const users=JSON.parse(cur.text); users.comptes=users.comptes||[];
      const ann=await readAnnuaire(users,S.key);
      const nann=await fn(users,ann);
      users.annuaire=await C.encJson(S.key,nann,C.AAD.annuaire);
      const out={v:1,kdf:"PBKDF2-SHA256",iter:C.ITER,comptes:users.comptes.map(c=>({id:c.id,sel:c.sel,iv:c.iv,cle:c.cle})),annuaire:users.annuaire};
      try{ await GH.put("_site/users.json",utf8b64(JSON.stringify(out,null,1)+"\n"),msg(message),cur.sha); return; }
      catch(e){ if((e.status===409||e.status===422)&&i<3){ await new Promise(r=>setTimeout(r,500)); continue; } throw e; }
    }
  });
}

/* ---------- vues, récents ---------- */
function bumpView(id){
  if(!S.me) return; const k="crm-vues-"+S.me.id, v=lsGet(k,{}); v[id]={n:((v[id]||{}).n||0)+1,t:new Date().toISOString()}; lsSet(k,v);
  S.recents=[id].concat(S.recents.filter(x=>x!==id)).slice(0,12);
  S.vus=Object.assign({},S.vus,{[id]:new Date().toISOString()});
  lsSet("crm-recents-"+S.me.id,S.recents); lsSet("crm-vus-"+S.me.id,S.vus);
}

/* ---------- menu compte ---------- */
$("acct").addEventListener("click",()=>{
  const sh=sheet('<h3>'+esc(S.me?(S.me.nom||S.me.login):"")+'</h3><p class="hint">'+esc(S.me?S.me.login+" · "+(ROLES[S.me.role]||""):"")+'</p><p class="hint">'+(S.genere?"Données publiées le "+esc(frStamp(S.genere))+".":"")+'</p>'+
    (S.me&&S.me.role==="admin"?'<button class="btn" id="m-admin">Administration des comptes</button>':"")+'<button class="btn ghost" id="m-reload">Recharger les données</button><button class="btn ghost" id="m-out">Se déconnecter</button><button class="btn ghost" data-close>Fermer</button>');
  const a=$("m-admin"); if(a) a.addEventListener("click",()=>{ closeSheet(); location.hash="#admin"; });
  $("m-reload").addEventListener("click",async()=>{ closeSheet(); toast("Rechargement…"); try{ await loadData(); rerender(); toast("Données à jour"); }catch(e){ toast("Rechargement impossible"); } });
  $("m-out").addEventListener("click",logout);
});

/* ---------- routeur ---------- */
let lastRoute="", pendingRerender=false;
function route(force){
  const h=(location.hash||"#accueil").slice(1);
  if(!S.ready) return;
  if(!force && h===lastRoute) return; const changed=h!==lastRoute;
  if(h.startsWith("c-") && S.dtabFor!==h){ S.dtab="client"; S.dtabFor=h; }
  lastRoute=h; closeCtx();
  if(h.startsWith("c-")) vDossier(decodeURIComponent(h.slice(2))); else if(h==="echeances") vEcheances(); else if(h==="ajouter") vAjouter(); else if(h==="envois") vEnvois(); else if(h==="admin") vAdmin(); else vAccueil();
  if(changed) window.scrollTo(0,0);
}
let rr=0; function rerender(){ cancelAnimationFrame(rr); rr=requestAnimationFrame(()=>{ if(!S.ready) return; if(lastRoute==="ajouter"&&$("form")) return; if(lastRoute==="admin") return; if($("sheetbg")){ pendingRerender=true; return; } const ae=document.activeElement; if(ae&&ae.id==="q") { drawGrid(); return; } if(ae&&ae.matches&&ae.matches("input[type=date]")) return; const y=window.scrollY; const keep=lastRoute; lastRoute=""; route(true); if(lastRoute===keep) window.scrollTo(0,y); }); }
window.addEventListener("hashchange",()=>{ closeSheet(); route(); });
$("back").addEventListener("click",()=>{ if(history.length>1) history.back(); else location.hash="#accueil"; });

/* ---------- interactions ---------- */
document.addEventListener("click",async e=>{
  const fv=e.target.closest("a[data-viewfile]"); if(fv){ e.preventDefault(); openFile(fv.dataset.viewfile); return; }
  const t=e.target.closest("button"); if(!t||t.closest(".tabs,.deskmenu")||!S.ready) return;
  if(t.dataset.mail){ openMail(t.dataset.mail); return; }
  if(t.dataset.open){ closeSheet(); location.hash="#c-"+t.dataset.open; return; }
  if(t.dataset.filt){ S.filt=t.dataset.filt; document.querySelectorAll("[data-filt]").forEach(b=>b.setAttribute("aria-pressed",b===t)); drawGrid(); return; }
  if(t.dataset.dtab){ S.dtab=t.dataset.dtab; document.querySelectorAll("[data-dtab]").forEach(b=>b.setAttribute("aria-selected",b===t)); const f=S.fiches.find(x=>"c-"+x.id===lastRoute); if(f) drawPane(f); return; }
  if(t.dataset.stage){ const id=lastRoute.slice(2), f=S.fiches.find(x=>x.id===id); const k=t.dataset.stage; const patch={etape:k}; if(k==="accepte"&&!D(id).date_signature) patch.date_signature=today();
    document.querySelectorAll("[data-stage]").forEach(b=>b.setAttribute("aria-pressed",b===t));
    try{ await setDossier(id,patch); rerender(); toast("Étape : "+STAGE_L[k]); if(f) await envoyer({client:f.nom,fiche:f.chemin,type:"fiche",texte:"Étape du dossier : "+STAGE_L[k]+(patch.date_signature?" (devis signé le "+patch.date_signature+")":"")}); }catch(err){ failToast(err); rerender(); } return; }
  if(t.dataset.new){ S.form={type:t.dataset.new,cli:t.dataset.for}; lastRoute="ajouter"; location.hash="#ajouter"; vAjouter(); return; }
  if(t.dataset.type){ readForm(); S.form.type=t.dataset.type; vAjouter(); return; }
  if(t.dataset.kind){ readForm(); S.form.kind=t.dataset.kind; vAjouter(); return; }
  if(t.dataset.plus){ $("f-date").value=addDays(today(),Number(t.dataset.plus)); return; }
  if(t.dataset.devis){ openDevisSheet(t.dataset.devis); return; }
  if(t.dataset.editdevis){ openEditDevis(t.dataset.editdevis); return; }
  if(t.dataset.editproj){ openEditProj(t.dataset.editproj); return; }
  if(t.dataset.setcover){ try{ await setDossier(t.dataset.for,{photo:t.dataset.setcover}); toast("Photo du dossier changée"); rerender(); }catch(err){ failToast(err); } return; }
  if(t.dataset.deldoc){ if(!confirm("Retirer ce fichier du dossier ?")) return; try{ await removeDoc(t.dataset.deldoc); toast("Fichier retiré"); rerender(); }catch(err){ failToast(err); } return; }
  if(t.dataset.askchantier){ const f=S.fiches.find(x=>x.id===t.dataset.askchantier); const pj=projet(f), dv=docsOf(f.id,"devis")[0]; t.disabled=true;
    try{ await envoyer({client:f.nom,fiche:f.chemin,type:"dossier-chantier",texte:"Générer le dossier chantier (fiche de préparation, alimentations électriques, schéma de principe) pour : "+[pj.type,pj.intitule,pj.refs].filter(Boolean).join(" — ")+(dv?" — devis n°"+(dv.numero||"?")+" du "+frDate(dv.date)+" (PDF dans le site, document "+dv.id+", fichier _app/"+dv.fichier+")":"")}); toast("Demande envoyée"); rerender(); }catch(err){ t.disabled=false; failToast(err); } return; }
  if(t.dataset.tick){ const r=S.rappels.find(x=>x.id===t.dataset.tick); if(!r) return; t.disabled=true; t.setAttribute("aria-pressed","true"); t.innerHTML=ICON_OK;
    try{ await envoyer({client:r.client,fiche:r.fiche,type:"rappel",rappel:r.id,texte:"Marquer comme fait le "+today()+" : "+(r.date||r.date_texte||"(sans date)")+" — "+r.action}); toast("Noté comme fait"); }catch(err){ t.disabled=false; t.setAttribute("aria-pressed","false"); t.innerHTML=""; failToast(err); } return; }
  if(t.dataset.undo){ if(!confirm("Annuler cet envoi ?")) return; const ev=S.envois.find(x=>x._id===t.dataset.undo); if(!ev) return;
    try{ await enqueue(async()=>{ const cur=await GH.req("GET",GH.p(ev._fichier)+"?ref="+encodeURIComponent(S.cfg.branch)); await GH.del(ev._fichier,cur.sha,msg("Envoi annulé : "+ev._id)); });
      S.envois=S.envois.filter(x=>x!==ev); badge(); toast("Envoi annulé"); rerender(); }catch(err){ toast(err.status===404?"Il a déjà été rangé":"Impossible d’annuler pour le moment"); } return; }
});
document.addEventListener("change",async e=>{
  const el=e.target; if(!S.ready) return;
  if(el.dataset.date){ const id=lastRoute.slice(2), f=S.fiches.find(x=>x.id===id); const k=el.dataset.date;
    if(k==="chantier_fin" && el.value && D(id).chantier_debut && el.value<D(id).chantier_debut){ toast("La fin est avant le début"); el.value=D(id).chantier_fin||""; return; }
    try{ await setDossier(id,{[k]:el.value}); toast("Date enregistrée");
      if(f && el.value && /chantier/.test(k)) await envoyer({client:f.nom,fiche:f.chemin,type:"historique",texte:"- "+today()+" — "+(k==="chantier_debut"?"Début de chantier prévu le ":"Fin de chantier prévue le ")+el.value});
      rerender(); }catch(err){ failToast(err); } return; }
  if(el.dataset.cover && el.files[0]){ const id=el.dataset.cover; toast("Envoi de la photo…"); const out=await addDocs(id,"photo",[el.files[0]]); if(out[0]){ try{ await setDossier(id,{photo:out[0].fichier}); }catch(err){ failToast(err); } rerender(); } return; }
  if(el.dataset.up && el.files.length){ const files=[...el.files]; el.value=""; toast("Envoi de "+files.length+" fichier"+(files.length>1?"s":"")+"…"); await addDocs(el.dataset.for,el.dataset.up,files); return; }
});
["dragover","dragleave","drop"].forEach(ev=>document.addEventListener(ev,e=>{ const z=e.target.closest&&e.target.closest("[data-drop]"); if(!z) return; e.preventDefault();
  if(ev==="dragover") z.classList.add("over"); else z.classList.remove("over");
  if(ev==="drop" && e.dataTransfer.files.length) addDocs(z.dataset.for,z.dataset.drop,[...e.dataTransfer.files]); }));
view.addEventListener("submit",e=>{ e.preventDefault(); if(e.target.id==="form") submitForm(); });
document.querySelectorAll('[data-tab="ajouter"]').forEach(a=>a.addEventListener("click",ev=>{ S.form={type:S.form.type||"rappel"}; if(lastRoute==="ajouter"){ ev.preventDefault(); vAjouter(); } }));

/* ---------- session (clé non exportable gardée sur l’appareil si demandé) ---------- */
const IDB={ open(){ return new Promise((ok,ko)=>{ const r=indexedDB.open("crm-robine",1); r.onupgradeneeded=()=>r.result.createObjectStore("s"); r.onsuccess=()=>ok(r.result); r.onerror=()=>ko(r.error); }); },
  async get(k){ const db=await IDB.open(); return new Promise(ok=>{ const r=db.transaction("s").objectStore("s").get(k); r.onsuccess=()=>ok(r.result); r.onerror=()=>ok(null); }); },
  async set(k,v){ const db=await IDB.open(); return new Promise(ok=>{ const t=db.transaction("s","readwrite"); t.objectStore("s").put(v,k); t.oncomplete=()=>ok(); t.onerror=()=>ok(); }); },
  async del(k){ const db=await IDB.open(); return new Promise(ok=>{ const t=db.transaction("s","readwrite"); t.objectStore("s").delete(k); t.oncomplete=()=>ok(); t.onerror=()=>ok(); }); } };
const SESSION_DAYS=30;
async function logout(){ try{ await IDB.del("session"); }catch(e){} location.hash=""; location.reload(); }

async function loadData(){
  const r=await fetch("data/data.json?t="+Date.now(),{cache:"no-store"});
  if(!r.ok) throw new Error(r.status===404?"Les données ne sont pas encore publiées (l’Action GitHub n’a pas encore tourné).":"Données indisponibles ("+r.status+")");
  let data; try{ data=await C.decJson(S.key,await r.json(),C.AAD.data); }catch(e){ throw new Error("Impossible de déchiffrer les données : la clé ne correspond pas (DATA_KEY changée ?)."); }
  S.cfg=data.config||{}; S.genere=data.genere||"";
  S.fiches=(data.fiches||[]).map(f=>Object.assign(f,{nom:f.nom||f.id,contenu:f.contenu||""}));
  S.rappelsRaw=data.rappels||[]; S.over=data.surcharges||{}; applyOver();
  S.envois=(data.inbox||[]); S.dossiers=data.dossiers||{};
  S.docs=(data.documents||[]).map((d,i)=>Object.assign({},d,{id:d.id||("doc-"+slug(d.fichier||d.nom||String(i)))}));
  S.mailsDb=withIds(data.mails); S.agendaDb=withIds(data.agenda); mergeAgenda();
  S.blobs=data.blobs||{};
  S.canEdit=!!(S.cfg.token && S.me && S.me.role!=="lecture");
  badge(); syncLabel();
  freshOverlay();
}
/* Lit directement dans le dépôt les fichiers que le site modifie (_app, _inbox), plus frais que data.json pendant la minute de publication. */
async function freshOverlay(){ try{ await overlay(); } finally { autoTermine(); } }
async function overlay(){
  if(!S.cfg||!S.cfg.token||!S.cfg.repo) return;
  const [owner,name]=S.cfg.repo.split("/"), br=S.cfg.branch||"main";
  const q='query($o:String!,$n:String!,$d:String!,$c:String!,$m:String!,$a:String!,$s:String!,$i:String!){repository(owner:$o,name:$n){d:object(expression:$d){...on Blob{text}} c:object(expression:$c){...on Blob{text}} m:object(expression:$m){...on Blob{text}} a:object(expression:$a){...on Blob{text}} s:object(expression:$s){...on Blob{text}} i:object(expression:$i){...on Tree{entries{name object{...on Blob{text}}}}}}}';
  const v={o:owner,n:name,d:br+":_app/dossiers.json",c:br+":_app/documents.json",m:br+":_app/mails.json",a:br+":_app/agenda.json",s:br+":_app/surcharges.json",i:br+":_inbox"};
  try{
    const r=await fetch("https://api.github.com/graphql",{method:"POST",cache:"no-store",headers:{Authorization:"Bearer "+S.cfg.token,"Content-Type":"application/json"},body:JSON.stringify({query:q,variables:v})});
    if(!r.ok) return; const j=await r.json(); const R=j&&j.data&&j.data.repository; if(!R) return;
    const P2=(o,def)=>{ try{ return o&&typeof o.text==="string"?JSON.parse(o.text):def; }catch(e){ return def; } };
    const dd=P2(R.d,null); if(dd) S.dossiers=dd;
    const dc=P2(R.c,null); if(Array.isArray(dc)) S.docs=dc.map((d,i)=>Object.assign({},d,{id:d.id||("doc-"+slug(d.fichier||d.nom||String(i)))}));
    const mm=P2(R.m,null); if(Array.isArray(mm)) S.mailsDb=withIds(mm);
    const aa=P2(R.a,null); if(Array.isArray(aa)){ S.agendaDb=withIds(aa); mergeAgenda(); }
    const ss=P2(R.s,null); if(ss){ S.over=ss; applyOver(); }
    if(R.i===null) S.envois=[];
    else if(R.i&&Array.isArray(R.i.entries)){ S.envois=R.i.entries.filter(e=>/\.json$/.test(e.name)).map(e=>{ const o=P2(e.object,null); return o?Object.assign({_id:e.name.replace(/\.json$/,""),_fichier:"_inbox/"+e.name},o):null; }).filter(Boolean).sort((a,b)=>String(b.cree||"").localeCompare(String(a.cree||""))); }
    badge(); rerender();
  }catch(e){}
}
/* Dossiers dont la fin de chantier est dépassée : étape « Terminé » enregistrée une fois. */
let autoDone=false;
async function autoTermine(){
  if(autoDone||!S.canEdit) return; autoDone=true;
  const list=clients().filter(f=>chantierFini(f) && D(f.id).etape!=="termine");
  if(!list.length) return;
  try{
    const now=new Date().toISOString();
    const nv=await updateJson("_app/dossiers.json",{},o=>{ let ch=false; list.forEach(f=>{ const d=o[f.id]||{}; if(d.chantier_fin && d.chantier_fin<today() && d.etape!=="termine"){ o[f.id]=Object.assign({},d,{etape:"termine",maj:now}); ch=true; } }); return ch?o:undefined; },"Dossiers passés en Terminé (fin de chantier dépassée) : "+list.map(f=>f.nom).join(", "));
    if(nv) S.dossiers=nv;
    for(const f of list){ if(pendingFor(f).some(e=>/termin/i.test(e.texte||""))) continue;
      await envoyer({client:f.nom,fiche:f.chemin,type:"fiche",texte:"Étape du dossier : Terminé (fin de chantier le "+D(f.id).chantier_fin+" dépassée) : passer le statut de la fiche à « terminé »."}); }
    rerender();
  }catch(e){}
}

async function afterLogin(){
  S.me.role="lecture";
  try{ const users=await fetchUsers(); const ann=await readAnnuaire(users,S.key); ann.forEach(a=>{ S.names[a.login]=a.nom||a.login; S.names[a.id]=a.nom||a.login; });
    const me=ann.find(a=>a.id===S.me.id); S.me.role=me?me.role:"lecture"; S.me.nom=me?(me.nom||me.login):S.me.login; if(me) S.me.login=me.login;
    if(!(users.comptes||[]).some(c=>c.id===S.me.id)){ await IDB.del("session"); throw new Error("Ce compte a été supprimé."); }
  }catch(e){ if(/supprimé/.test(e.message)) throw e; }
  view.innerHTML='<p class="status"><span class="spin"></span> Déchiffrement des données…</p>';
  await loadData();
  S.recents=lsGet("crm-recents-"+S.me.id,[]); S.vus=lsGet("crm-vus-"+S.me.id,{});
  $("tabs").hidden=false; $("deskmenu").hidden=false; $("acct").hidden=false;
  S.ready=true; lastRoute=""; route(true);
}
function vLogin(errMsg){
  setHead("CRM Robine"); $("tabs").hidden=true; $("deskmenu").hidden=true; $("acct").hidden=true;
  const last=lsGet("crm-dernier-login","");
  view.innerHTML='<form class="login" id="login" autocomplete="on"><div class="logo"><img src="icons/icon-192.png" alt=""><div><b>CRM Robine</b><span>Accès réservé à l’équipe</span></div></div>'+
    (errMsg?'<p class="err">'+esc(errMsg)+'</p>':"")+
    '<label class="f">Identifiant<input id="l-u" name="username" autocomplete="username" autocapitalize="none" spellcheck="false" required value="'+esc(last)+'"></label>'+
    '<label class="f">Mot de passe<input id="l-p" name="password" type="password" autocomplete="current-password" required></label>'+
    '<label class="chk"><input type="checkbox" id="l-r"> Rester connecté sur cet appareil ('+SESSION_DAYS+' jours)</label>'+
    '<button class="primary" id="l-go" type="submit">Se connecter</button><p class="hint">Les données sont chiffrées : rien n’est lisible sans ton mot de passe. Coche « rester connecté » uniquement sur ton propre téléphone ou PC.</p></form>';
  $(last?"l-p":"l-u").focus();
  $("login").addEventListener("submit",async e=>{
    e.preventDefault(); const b=$("l-go"); b.disabled=true; b.innerHTML='<span class="spin"></span> Vérification…';
    const login=$("l-u").value.trim(), pass=$("l-p").value, remember=$("l-r").checked;
    try{
      const users=await fetchUsers(); const r=await C.unwrap(users,login,pass);
      S.rawKey=r.raw; S.key=await C.importData(r.raw,false); S.me={id:r.id,login};
      lsSet("crm-dernier-login",login);
      if(remember) await IDB.set("session",{id:r.id,login,key:await C.importData(r.raw,false),exp:Date.now()+SESSION_DAYS*864e5}).catch(()=>{});
      else await IDB.del("session").catch(()=>{});
      await afterLogin();
    }catch(err){ S.key=null; S.rawKey=null; vLogin(err.message||"Connexion impossible"); }
  });
}
(async()=>{
  if(!window.crypto||!crypto.subtle){ view.innerHTML='<p class="err">Ce navigateur ne permet pas le chiffrement : ouvre le site en https dans un navigateur à jour.</p>'; return; }
  if("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(()=>{});
  try{ const s=await IDB.get("session");
    if(s && s.key && s.exp>Date.now()){ S.key=s.key; S.me={id:s.id,login:s.login}; await afterLogin(); return; }
    if(s) await IDB.del("session");
  }catch(e){ S.key=null; S.me=null; return vLogin(e.message); }
  vLogin();
})();
})();
