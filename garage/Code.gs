// ── CONFIG ───────────────────────────────────────────────────────────────────
var CFG={
  token:  "DCL2026secret", // changer avant déploiement
  waNum:  "33668163064",
  sumup:  "VOTRE_LIEN_SUMUP",  // lien SumUp du cousin
  acompte:"50"
};

// ── ROUTING ──────────────────────────────────────────────────────────────────
function doGet(e){return route(e,null);}
function doPost(e){
  var b={};try{b=JSON.parse(e.postData.contents);}catch(x){}
  return route(e,b);
}
function route(e,body){
  var a=((body&&body.action)||e.parameter.action||"").toLowerCase();
  var r;
  try{
    if     (a==="slots")      r=getSlots();
    else if(a==="submit")     r=submitDemande(body);
    else if(a==="requests")   r=auth(e,body)?getRequests(e.parameter.statut||""):{ok:false,error:"Non autorisé"};
    else if(a==="accept")     r=auth(e,body)?acceptDemande(body.id):{ok:false,error:"Non autorisé"};
    else if(a==="refuse")     r=auth(e,body)?refuseDemande(body.id):{ok:false,error:"Non autorisé"};
    else if(a==="addslot")    r=auth(e,body)?addCreneau(body):{ok:false,error:"Non autorisé"};
    else if(a==="removeslot") r=auth(e,body)?removeCreneau(body.id):{ok:false,error:"Non autorisé"};
    else r={ok:false,error:"Action inconnue: "+a};
  }catch(err){r={ok:false,error:err.message};}
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}
function auth(e,body){
  var t=(body&&body.token)||e.parameter.token||"";
  return t===CFG.token;
}

// ── UTILS ────────────────────────────────────────────────────────────────────
function sh(name){return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);}
function genId(p){return p+Date.now().toString(36).toUpperCase();}
function toWa(tel){var c=String(tel).replace(/\D/g,"");return c.startsWith("0")?"33"+c.slice(1):c;}
function dateFR(d){
  if(!d)return"";
  var dt=new Date(d);
  var days=["dimanche","lundi","mardi","mercredi","jeudi","vendredi","samedi"];
  var months=["janvier","février","mars","avril","mai","juin","juillet","août","septembre","octobre","novembre","décembre"];
  return days[dt.getDay()]+" "+dt.getDate()+" "+months[dt.getMonth()];
}

// ── CRÉNEAUX ─────────────────────────────────────────────────────────────────
function getSlots(){
  var s=sh("Creneaux");if(!s)return{ok:true,slots:[]};
  var rows=s.getDataRange().getValues();
  var today=new Date();today.setHours(0,0,0,0);
  var slots=[];
  for(var i=1;i<rows.length;i++){
    var r=rows[i];if(!r[0])continue;
    var d=new Date(r[1]);if(d<today)continue;
    if(r[4]!=="libre")continue;
    slots.push({id:String(r[0]),date:r[1],heure:String(r[2]),duree:Number(r[3])});
  }
  return{ok:true,slots:slots};
}
function addCreneau(data){
  var s=sh("Creneaux"),id=genId("C");
  s.appendRow([id,data.date,data.heure,data.duree||2,"libre",""]);
  return{ok:true,id:id};
}
function removeCreneau(id){
  var s=sh("Creneaux"),rows=s.getDataRange().getValues();
  for(var i=1;i<rows.length;i++){
    if(String(rows[i][0])===String(id)){s.deleteRow(i+1);return{ok:true};}
  }
  return{ok:false,error:"Introuvable"};
}

// ── DEMANDES ─────────────────────────────────────────────────────────────────
function submitDemande(data){
  var cs=sh("Creneaux"),crow=cs.getDataRange().getValues(),ci=-1;
  for(var i=1;i<crow.length;i++){
    if(String(crow[i][0])===String(data.creneauId)){
      if(crow[i][4]!=="libre")return{ok:false,error:"Créneau déjà réservé"};
      ci=i+1;break;
    }
  }
  if(ci<0)return{ok:false,error:"Créneau introuvable"};
  var id=genId("D");
  var now=Utilities.formatDate(new Date(),"Europe/Paris","dd/MM/yyyy HH:mm");
  sh("Demandes").appendRow([id,now,data.prenom,data.tel,data.cat,data.serviceId,data.serviceNom,data.prix,data.marque||"",data.modele||"",data.creneauId,data.date,data.heure,data.notes||"","en_attente"]);
  cs.getRange(ci,5).setValue("prereserve");
  cs.getRange(ci,6).setValue(id);
  return{ok:true,id:id};
}
function getRequests(statut){
  var s=sh("Demandes");if(!s)return{ok:true,requests:[]};
  var rows=s.getDataRange().getValues(),list=[];
  for(var i=1;i<rows.length;i++){
    var r=rows[i];if(!r[0])continue;
    if(statut&&r[14]!==statut)continue;
    list.push({id:String(r[0]),date:String(r[1]),prenom:String(r[2]),tel:String(r[3]),cat:String(r[4]),serviceId:String(r[5]),serviceNom:String(r[6]),prix:String(r[7]),marque:String(r[8]),modele:String(r[9]),creneauId:String(r[10]),rdvDate:String(r[11]),rdvHeure:String(r[12]),notes:String(r[13]),statut:String(r[14])});
  }
  list.sort(function(a,b){return a.date<b.date?1:-1;});
  return{ok:true,requests:list};
}
function acceptDemande(id){
  var ds=sh("Demandes"),cs=sh("Creneaux");
  var drows=ds.getDataRange().getValues();
  for(var i=1;i<drows.length;i++){
    if(String(drows[i][0])!==String(id))continue;
    ds.getRange(i+1,15).setValue("accepte");
    var r=drows[i];
    var crows=cs.getDataRange().getValues();
    for(var j=1;j<crows.length;j++){
      if(String(crows[j][0])===String(r[10])){cs.getRange(j+1,5).setValue("pris");break;}
    }
    var msg="✅ Bonjour "+r[2]+" !\n\nVotre RDV est *confirmé* chez Detailing Center Lyon.\n\n📅 "+dateFR(r[11])+" à "+r[12]+"\n🔧 "+r[6]+" ("+r[7]+")\n";
    if(r[8]||r[9])msg+="🚘 "+[r[8],r[9]].filter(Boolean).join(" ")+"\n";
    msg+="\nPour finaliser, merci de régler l'acompte de "+CFG.acompte+" € :\n"+CFG.sumup+"\n\nÀ très bientôt ! 🏁";
    return{ok:true,waUrl:"https://wa.me/"+toWa(r[3])+"?text="+encodeURIComponent(msg)};
  }
  return{ok:false,error:"Demande introuvable"};
}
function refuseDemande(id){
  var ds=sh("Demandes"),cs=sh("Creneaux");
  var drows=ds.getDataRange().getValues();
  for(var i=1;i<drows.length;i++){
    if(String(drows[i][0])!==String(id))continue;
    ds.getRange(i+1,15).setValue("refuse");
    var r=drows[i];
    var crows=cs.getDataRange().getValues();
    for(var j=1;j<crows.length;j++){
      if(String(crows[j][0])===String(r[10])){cs.getRange(j+1,5).setValue("libre");cs.getRange(j+1,6).setValue("");break;}
    }
    var msg="Bonjour "+r[2]+",\n\nNous ne pouvons malheureusement pas confirmer votre RDV pour le moment.\n\nN'hésitez pas à nous recontacter pour trouver un autre créneau.\n\nDetailing Center Lyon · 06 68 16 30 64";
    return{ok:true,waUrl:"https://wa.me/"+toWa(r[3])+"?text="+encodeURIComponent(msg)};
  }
  return{ok:false,error:"Demande introuvable"};
}

// ── SETUP (lancer une seule fois) ────────────────────────────────────────────
function setup(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var cr=ss.getSheetByName("Creneaux")||ss.insertSheet("Creneaux");
  if(cr.getLastRow()===0){cr.appendRow(["ID","Date","Heure","Duree(h)","Statut","DemandID"]);cr.getRange(1,1,1,6).setFontWeight("bold");}
  var dm=ss.getSheetByName("Demandes")||ss.insertSheet("Demandes");
  if(dm.getLastRow()===0){dm.appendRow(["ID","DateSoumission","Prenom","Tel","Cat","ServiceID","ServiceNom","Prix","Marque","Modele","CreneauID","DateRDV","HeureRDV","Notes","Statut"]);dm.getRange(1,1,1,15).setFontWeight("bold");}
  Logger.log("Setup OK ✓");
}
