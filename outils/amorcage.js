(function(){
"use strict";
const C=self.CRMC, $=id=>document.getElementById(id);
$("p1").addEventListener("input",()=>{ const n=[...$("p1").value].length, ok=n>=C.MIN_PASS; $("m").style.width=Math.min(100,n/C.MIN_PASS*100)+"%"; $("m").classList.toggle("ok",ok); $("mh").textContent=ok?n+" caractères : OK":n+" / "+C.MIN_PASS+" caractères minimum"; });
const err=m=>{ $("e").hidden=!m; $("e").textContent=m||""; };
$("f").addEventListener("submit",async e=>{
  e.preventDefault(); err("");
  const login=$("u").value.trim(), nom=$("n").value.trim()||login, p1=$("p1").value, p2=$("p2").value;
  if(!login||/\s/.test(login)) return err("Identifiant obligatoire, sans espace.");
  const pe=C.checkPass(p1); if(pe) return err(pe); if(p1!==p2) return err("Les deux mots de passe ne sont pas identiques.");
  let raw; const kin=$("k").value.trim();
  if(kin){ try{ raw=C.b64d(kin); }catch(x){ raw=null; } if(!raw||raw.length!==32) return err("La clé collée n’est pas une clé base64 de 32 octets."); }
  else raw=C.rnd(32);
  $("go").disabled=true; $("go").textContent="Chiffrement (quelques secondes)…";
  try{
    const key=await C.importData(raw,false);
    const compte=await C.wrap(login,p1,raw);
    const now=new Date().toISOString();
    const annuaire=await C.encJson(key,[{id:compte.id,login,nom,role:"admin",cree:now,maj:now}],C.AAD.annuaire);
    const users={v:1,kdf:"PBKDF2-SHA256",iter:C.ITER,comptes:[compte],annuaire};
    const check=await C.unwrap(users,login,p1); if(C.b64e(check.raw)!==C.b64e(raw)) throw new Error("vérification");
    $("ok").value=C.b64e(raw); $("ou").value=JSON.stringify(users,null,1)+"\n";
    $("f").hidden=true; $("out").hidden=false; $("p1").value=$("p2").value="";
  }catch(x){ err("Erreur : "+x.message); $("go").disabled=false; $("go").textContent="Générer"; }
});
document.addEventListener("click",async e=>{ const b=e.target.closest("[data-copy]"); if(!b) return; const t=$(b.dataset.copy); t.select();
  try{ await navigator.clipboard.writeText(t.value); b.textContent="Copié"; }catch(x){ document.execCommand("copy"); b.textContent="Copié"; } setTimeout(()=>b.textContent=b.dataset.copy==="ok"?"Copier la clé":"Copier users.json",1500); });
})();
