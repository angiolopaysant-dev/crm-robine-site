/* Chiffrement du CRM : AES-256-GCM pour les données, PBKDF2-SHA256 (600 000 itérations, sel unique) pour envelopper la clé de données. */
(function(g){
"use strict";
const te=new TextEncoder(), td=new TextDecoder();
const ITER=600000, MIN_PASS=16;
const AAD={data:te.encode("crm-robine:data:v1"), fichier:te.encode("crm-robine:fichier:v1"), annuaire:te.encode("crm-robine:annuaire:v1")};
const b64e=u8=>{ u8=u8 instanceof Uint8Array?u8:new Uint8Array(u8); let s=""; for(let i=0;i<u8.length;i+=0x8000) s+=String.fromCharCode.apply(null,u8.subarray(i,i+0x8000)); return btoa(s); };
const b64d=b=>{ const s=atob(String(b||"").replace(/\s+/g,"")); const u=new Uint8Array(s.length); for(let i=0;i<s.length;i++) u[i]=s.charCodeAt(i); return u; };
const rnd=n=>crypto.getRandomValues(new Uint8Array(n));
const hex=u8=>[...new Uint8Array(u8)].map(b=>b.toString(16).padStart(2,"0")).join("");
async function userId(login){ return hex(await crypto.subtle.digest("SHA-256",te.encode("crm-robine-site:v1:"+String(login||"").normalize("NFC").trim().toLowerCase()))); }
async function kek(pass,salt,iter){
  const base=await crypto.subtle.importKey("raw",te.encode(String(pass).normalize("NFC")),"PBKDF2",false,["deriveKey"]);
  return crypto.subtle.deriveKey({name:"PBKDF2",hash:"SHA-256",salt,iterations:iter},base,{name:"AES-GCM",length:256},false,["encrypt","decrypt"]);
}
async function enc(key,u8,aad){ const iv=rnd(12); const ct=new Uint8Array(await crypto.subtle.encrypt(Object.assign({name:"AES-GCM",iv},aad?{additionalData:aad}:{}),key,u8)); return {iv,ct}; }
async function dec(key,iv,ct,aad){ return new Uint8Array(await crypto.subtle.decrypt(Object.assign({name:"AES-GCM",iv},aad?{additionalData:aad}:{}),key,ct)); }
const importData=(raw,extractable)=>crypto.subtle.importKey("raw",raw,"AES-GCM",!!extractable,["encrypt","decrypt"]);
function checkPass(p){ p=String(p||""); if([...p].length<MIN_PASS) return "Le mot de passe doit faire au moins "+MIN_PASS+" caractères."; return ""; }
/* Enveloppe la clé de données (raw, 32 octets) pour un identifiant + mot de passe. */
async function wrap(login,pass,raw){
  const e=checkPass(pass); if(e) throw new Error(e);
  const id=await userId(login), sel=rnd(16), k=await kek(pass,sel,ITER);
  const {iv,ct}=await enc(k,te.encode(JSON.stringify({k:b64e(raw),id})),te.encode(id));
  return {id,sel:b64e(sel),iv:b64e(iv),cle:b64e(ct)};
}
/* Retrouve la clé de données à partir de users.json, identifiant et mot de passe. */
async function unwrap(users,login,pass){
  if(!users||users.kdf!=="PBKDF2-SHA256"||!(users.iter>=ITER)) throw new Error("Fichier des comptes invalide.");
  const id=await userId(login), c=(users.comptes||[]).find(x=>x.id===id);
  const k=await kek(pass,c?b64d(c.sel):rnd(16),users.iter); // calcul fait même si l'identifiant est inconnu
  if(!c) throw new Error("Identifiant ou mot de passe incorrect.");
  let p; try{ p=await dec(k,b64d(c.iv),b64d(c.cle),te.encode(id)); }catch(e){ throw new Error("Identifiant ou mot de passe incorrect."); }
  const o=JSON.parse(td.decode(p)); if(o.id!==id) throw new Error("Compte invalide.");
  return {id,raw:b64d(o.k)};
}
async function encJson(key,obj,aad){ const {iv,ct}=await enc(key,te.encode(JSON.stringify(obj)),aad); return {iv:b64e(iv),ct:b64e(ct)}; }
async function decJson(key,o,aad){ return JSON.parse(td.decode(await dec(key,b64d(o.iv),b64d(o.ct),aad))); }
async function decFile(key,buf){ const u=new Uint8Array(buf); return dec(key,u.subarray(0,12),u.subarray(12),AAD.fichier); }
g.CRMC={ITER,MIN_PASS,AAD,b64e,b64d,rnd,hex,userId,kek,enc,dec,importData,checkPass,wrap,unwrap,encJson,decJson,decFile,te,td};
})(self);
