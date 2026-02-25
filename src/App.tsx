import { useState } from "react";

const STAFF = ["Claudia","Consuelo","Diane","Giorgia","Giulia","Mary"];
const DAYS  = ["Lunedì","Martedì","Mercoledì","Giovedì","Venerdì","Sabato"];
const DAYS_SHORT = ["Lun","Mar","Mer","Gio","Ven","Sab"];
const LUN=0,MAR=1,MER=2,GIO=3,VEN=4,SAB=5;
const MAX=3;
const WDAYS=[LUN,MAR,MER,GIO,VEN];
const N_ITER=20000;

const CONSTRAINT_OPTS=[
  {val:"",    label:"—"},
  {val:"abs", label:"🏠 Assente"},
  {val:"noM", label:"🚫 No Mattina"},
  {val:"noP", label:"🚫 No Pomeriggio"},
  {val:"onM", label:"☀️ Solo Mattina"},
  {val:"onP", label:"🌙 Solo Pomeriggio"},
];

const CONSTRAINT_OPTS_SAB=[
  {val:"",    label:"—"},
  {val:"abs", label:"🏠 Assente"},
];

const CASSA_SLOTS=new Set(["2-P","3-M","4-P"]);
function isCassa(di,t){ return CASSA_SLOTS.has(`${di}-${t}`); }
function isCassaPerson(n,di,t){ return n==="Consuelo"&&isCassa(di,t); }
const MIN_COV={"1-M":2,"3-M":2};
function minCov(di,t){ return MIN_COV[`${di}-${t}`]??MAX; }

function shuffle(a,rng){
  const b=[...a];
  for(let i=b.length-1;i>0;i--){ const j=Math.floor(rng()*(i+1)); [b[i],b[j]]=[b[j],b[i]]; }
  return b;
}
function makeRng(seed){
  let s=seed>>>0;
  return ()=>{ s+=0x6D2B79F5; let t=Math.imul(s^s>>>15,1|s); t^=t+Math.imul(t^t>>>7,61|t); return ((t^t>>>14)>>>0)/4294967296; };
}

function constraintBlocks(c, t){
  if(!c) return false;
  if(c==="abs") return true;
  if(c==="noM" && t==="M") return true;
  if(c==="noP" && (t==="P"||t==="PS")) return true;
  if(c==="onM" && (t==="P"||t==="PS")) return true;
  if(c==="onP" && t==="M") return true;
  return false;
}

function simulate(weekStart, seed, sabato, constraints){
  const rng=makeRng(seed);
  const sh=a=>shuffle(a,rng);
  const weekIndex=Math.floor(weekStart.getTime()/(7*24*3600*1000));

  let consueloSab;
  if(sabato==="Consuelo") consueloSab=true;
  else if(sabato==="Claudia") consueloSab=false;
  else consueloSab=(weekIndex%2===0);

  const warnings=[];
  const S=Array.from({length:6},()=>({M:[],P:[],PS:[],U:[]}));

  function cof(name,di){ return constraints?.[name]?.[di]||""; }

  function covM(di){ return S[di].M.filter(s=>!isCassaPerson(s,di,"M")).length; }
  function covPnorm(di){ return S[di].P.filter(s=>!isCassaPerson(s,di,"P")).length; }
  function countPS(di){ return S[di].PS.length; }
  function covP(di){ return covPnorm(di)+countPS(di); }
  function covU(){ return S[SAB].U.length; }
  function turniInGiorno(name,di){
    return ["M","P","PS"].reduce((n,t)=>n+(S[di][t].includes(name)&&!isCassaPerson(name,di,t)?1:0),0);
  }
  // constraintForces: se il vincolo IMPONE un turno specifico (onM→M obbligatorio, onP→P/PS obbligatorio)
  function constraintForces(c, t){
    if(!c) return false;
    if(c==="onM" && t==="M") return true;
    if(c==="onP" && (t==="P"||t==="PS")) return true;
    return false;
  }
  function tryAdd(name,di,t,maxPerDay=1){
    if(S[di][t].includes(name)) return false;
    if(constraintBlocks(cof(name,di),t)) return false;
    const forced = constraintForces(cof(name,di),t);
    if(!isCassaPerson(name,di,t)){
      const cur=t==="M"?covM(di):t==="U"?covU():covP(di);
      if(cur>=MAX) return false;
      // I vincoli espliciti (onM/onP) saltano i limiti di turni-per-giorno e giorni-per-settimana
      if(!forced){
        if(turniInGiorno(name,di)>=maxPerDay) return false;
      }
      if(t==="PS"&&countPS(di)>=1) return false;
    }
    S[di][t].push(name); return true;
  }
  function assignN(name,pool,t,n,maxPerDay=1){
    let p;
    if(t==="M"){
      const prio=sh(pool.filter(d=>d!==MAR&&d!==GIO));
      const depr=sh(pool.filter(d=>d===MAR||d===GIO));
      p=[...prio,...depr];
    } else { p=sh(pool); }
    // Giorni con vincolo forzante vanno in testa, quelli bloccati vengono esclusi
    const forced  = p.filter(di=> constraintForces(cof(name,di),t));
    const normal  = p.filter(di=>!constraintForces(cof(name,di),t) && !constraintBlocks(cof(name,di),t));
    p=[...forced,...normal];
    let done=0;
    for(const di of p){ if(done>=n) break; if(tryAdd(name,di,t,maxPerDay)) done++; }
    if(done<n) warnings.push(`${name}: ${done}/${n} turni ${t}`);
  }

  // ── PRE-ASSEGNAZIONE VINCOLI FORZANTI (onM / onP) ───────────────
  // Priorità ASSOLUTA: vengono assegnati prima di tutto il resto,
  // ignorando i limiti di copertura giornaliera (MAX).
  for(const name of STAFF){
    for(let di=0;di<5;di++){
      const c=cof(name,di);
      if(c==="onM" && !S[di].M.includes(name)){
        S[di].M.push(name);
      }
      if(c==="onP" && !S[di].P.includes(name)){
        S[di].P.push(name);
      }
    }
  }

  // ── FISSI ────────────────────────────────────────────────────────
  if(!constraintBlocks(cof("Giulia",LUN),"P")) tryAdd("Giulia",LUN,"P");
  if(!constraintBlocks(cof("Giulia",VEN),"P")) tryAdd("Giulia",VEN,"P");

  const consueloRiposo=consueloSab?((weekIndex+seed)%2===0?LUN:MAR):null;
  if(consueloRiposo!==LUN && !constraintBlocks(cof("Consuelo",LUN),"P")) tryAdd("Consuelo",LUN,"P");
  if(consueloRiposo!==MAR && !constraintBlocks(cof("Consuelo",MAR),"M")) tryAdd("Consuelo",MAR,"M");
  if(!constraintBlocks(cof("Consuelo",MER),"P")) S[MER].P.push("Consuelo");
  if(!constraintBlocks(cof("Consuelo",GIO),"M")) S[GIO].M.push("Consuelo");
  if(!constraintBlocks(cof("Consuelo",VEN),"P")) S[VEN].P.push("Consuelo");
  if(!constraintBlocks(cof("Consuelo",MER),"M")) tryAdd("Consuelo",MER,"M",2);
  if(consueloSab && cof("Consuelo",SAB)!=="abs") S[SAB].U.push("Consuelo");

  if(cof("Giorgia",SAB)!=="abs") S[SAB].U.push("Giorgia");
  if(cof("Mary",SAB)!=="abs")    S[SAB].U.push("Mary");
  if(!consueloSab && cof("Claudia",SAB)!=="abs") S[SAB].U.push("Claudia");

  // ── HELPERS PS ───────────────────────────────────────────────────
  const psUsed=new Set();
  const MAX_P={Diane:2};
  const MAX_DAYS={Giorgia:6,Diane:4,Consuelo:5};
  function workingDays(name){
    return WDAYS.filter(di=>S[di].M.includes(name)||S[di].P.includes(name)||S[di].PS.includes(name)).length;
  }
  function countPomeriggi(name){
    return WDAYS.reduce((acc,di)=>acc+(S[di].P.includes(name)?1:0)+(S[di].PS.includes(name)?1:0),0);
  }

  // ── CONSUELO M+PS PROATTIVO ───────────────────────────────────────
  if(rng()>0.4){
    const consDays=[LUN,MAR].filter(d=>(consueloRiposo===null||d!==consueloRiposo)&&!constraintBlocks(cof("Consuelo",d),"PS"));
    for(const di of consDays){
      if(countPS(di)>=1) continue;
      if(S[di].P.includes("Consuelo")&&!S[di].M.includes("Consuelo")&&covM(di)<MAX&&!constraintBlocks(cof("Consuelo",di),"M")){
        S[di].P.splice(S[di].P.indexOf("Consuelo"),1);
        S[di].M.push("Consuelo"); S[di].PS.push("Consuelo");
        psUsed.add("Consuelo"); break;
      }
      if(S[di].M.includes("Consuelo")&&!S[di].PS.includes("Consuelo")&&!S[di].P.includes("Consuelo")){
        S[di].PS.push("Consuelo");
        psUsed.add("Consuelo"); break;
      }
    }
  }

  // ── CLAUDIA ───────────────────────────────────────────────────────
  const claudiaMaxWdays=consueloSab?5:4;
  const claudiaMcount=consueloSab?2:1;
  const claudiaPool=sh(WDAYS.filter(di=>cof("Claudia",di)!=="abs")).slice(0,claudiaMaxWdays);
  // Conta M già assegnati da vincolo forzante
  const claudiaForcedM=claudiaPool.filter(di=>S[di].M.includes("Claudia"));
  const claudiaNeedM=Math.max(0, claudiaMcount-claudiaForcedM.length);
  if(claudiaNeedM>0){
    const clMpool=[...sh(claudiaPool.filter(d=>d!==MAR&&d!==GIO&&!S[d].M.includes("Claudia"))),...sh(claudiaPool.filter(d=>(d===MAR||d===GIO)&&!S[d].M.includes("Claudia")))];
    let clM=0;
    for(const di of clMpool){ if(clM>=claudiaNeedM) break; if(tryAdd("Claudia",di,"M",1)) clM++; }
    if(claudiaForcedM.length+clM<claudiaMcount) warnings.push(`Claudia: ${claudiaForcedM.length+clM}/${claudiaMcount} turni M`);
  }
  const clMdays=claudiaPool.filter(di=>S[di].M.includes("Claudia"));
  const clPpool=sh(claudiaPool.filter(di=>!clMdays.includes(di)&&!S[di].P.includes("Claudia")));
  const claudiaForcedP=claudiaPool.filter(di=>S[di].P.includes("Claudia")).length;
  let clP=claudiaForcedP;
  for(const di of clPpool){ if(clP>=3) break; if(tryAdd("Claudia",di,"P",1)) clP++; }
  if(clP<3) warnings.push(`Claudia: ${clP}/3 turni P`);

  // ── MARY ──────────────────────────────────────────────────────────
  function maryScore(di){ return Math.max(0,minCov(di,"M")-covM(di))+Math.max(0,minCov(di,"P")-covP(di)); }
  const mSorted=WDAYS.filter(di=>cof("Mary",di)!=="abs").sort((a,b)=>maryScore(b)-maryScore(a)||(rng()-0.5));
  let mD=0,mOM=false,mOP=false; const mAss=[];
  // Conta giorni M già assegnati da pre-vincolo
  for(const di of mSorted){
    if(S[di].M.includes("Mary")&&S[di].PS.includes("Mary")){ mD++; mAss.push(di); }
    else if(S[di].M.includes("Mary")&&!S[di].PS.includes("Mary")){ mOM=true; mAss.push(di); }
    else if(S[di].P.includes("Mary")&&!S[di].M.includes("Mary")){ mOP=true; mAss.push(di); }
  }
  for(const di of mSorted){
    if(mD>=3&&mOM&&mOP) break;
    if(mD<3&&!mAss.includes(di)&&!constraintBlocks(cof("Mary",di),"M")&&!constraintBlocks(cof("Mary",di),"PS")
      &&covM(di)<MAX&&!S[di].M.includes("Mary")&&covP(di)<MAX&&countPS(di)<1){
      S[di].M.push("Mary"); S[di].PS.push("Mary"); mD++; mAss.push(di);
    }
  }
  for(const di of mSorted){ if(mOM) break; if(!mAss.includes(di)&&!constraintBlocks(cof("Mary",di),"M")&&covM(di)<MAX&&!S[di].M.includes("Mary")){ S[di].M.push("Mary"); mOM=true; mAss.push(di); } }
  for(const di of mSorted){ if(mOP) break; if(!mAss.includes(di)&&!constraintBlocks(cof("Mary",di),"P")&&covPnorm(di)<MAX&&!S[di].P.includes("Mary")){ S[di].P.push("Mary"); mOP=true; mAss.push(di); } }
  if(mD<3) warnings.push(`Mary: ${mD}/3 giorni M+PS`);
  if(!mOM) warnings.push("Mary: giorno solo M mancante");
  if(!mOP) warnings.push("Mary: giorno solo P mancante");

  // ── GIORGIA ───────────────────────────────────────────────────────
  // Conta i giorni M già assegnati (da vincolo forzante o da logica precedente)
  {
    const forcedM=WDAYS.filter(di=>S[di].M.includes("Giorgia"));
    const neededM=Math.max(0, 2-forcedM.length);
    if(neededM>0) assignN("Giorgia", WDAYS.filter(di=>!S[di].M.includes("Giorgia")), "M", neededM, 1);
    const giorgiaMD=WDAYS.filter(di=>S[di].M.includes("Giorgia"));
    const giorgiaPPool=sh(WDAYS.filter(di=>!giorgiaMD.includes(di) && !S[di].P.includes("Giorgia")));
    const forcedP=WDAYS.filter(di=>S[di].P.includes("Giorgia")).length;
    let giorgiaP=forcedP;
    for(const di of giorgiaPPool){ if(giorgiaP>=3) break; if(tryAdd("Giorgia",di,"P",1)) giorgiaP++; }
    if(giorgiaP<3) warnings.push(`Giorgia: ${giorgiaP}/3 turni P`);
  }

  // ── DIANE ─────────────────────────────────────────────────────────
  {
    // Ricontrollato ad ogni assegnazione, non pre-calcolato
    function dianeAdjacentPom(di){
      const hasPom=(d)=>d>=LUN&&d<=VEN&&(S[d].P.includes("Diane")||S[d].PS.includes("Diane"));
      return hasPom(di-1)||hasPom(di+1);
    }
    const forcedM=WDAYS.filter(di=>S[di].M.includes("Diane"));
    const neededM=Math.max(0, 2-forcedM.length);
    if(neededM>0) assignN("Diane", WDAYS.filter(di=>!S[di].M.includes("Diane")), "M", neededM, 1);
    const diMD=WDAYS.filter(di=>S[di].M.includes("Diane"));
    const diPPool=sh(WDAYS.filter(di=>!diMD.includes(di) && !S[di].P.includes("Diane")));
    const forcedP=WDAYS.filter(di=>S[di].P.includes("Diane")).length;
    let diP=forcedP;
    for(const di of diPPool){
      if(diP>=2) break;
      if(dianeAdjacentPom(di)) continue; // rivaluta in tempo reale dopo ogni assegnazione
      if(tryAdd("Diane",di,"P",1)) diP++;
    }
    if(diP<2) warnings.push(`Diane: ${diP}/2 turni P`);
  }

  // ── REPERIBILITÀ PS ───────────────────────────────────────────────
  // Helper riutilizzabile per check consecutivi Diane
  function dianeConsecutivePom(di){
    const hasPom=(d)=>d>=0&&d<=4&&(S[d].P.includes("Diane")||S[d].PS.includes("Diane"));
    return hasPom(di-1)||hasPom(di+1);
  }
  const psElig=[
    {name:"Giorgia", days:[...WDAYS]},
    {name:"Diane",   days:[...WDAYS].sort((a,b)=>{ const aP=S[a].P.includes("Diane"),bP=S[b].P.includes("Diane"); return aP&&!bP?-1:!aP&&bP?1:0; })},
    {name:"Consuelo",days:[LUN,MAR].filter(d=>consueloRiposo===null||d!==consueloRiposo)},
  ];
  for(let di=0;di<5;di++){
    const pShort=()=>covP(di)<minCov(di,"P");
    const mShort=()=>covM(di)<minCov(di,"M");
    if(!pShort()&&!mShort()) continue;
    for(const {name,days} of psElig){
      if(psUsed.has(name)) continue;
      if(!days.includes(di)) continue;
      if(constraintBlocks(cof(name,di),"M")||constraintBlocks(cof(name,di),"PS")) continue;
      if(name==="Diane" && dianeConsecutivePom(di)) continue;
      if(S[di].P.includes(name)&&!isCassaPerson(name,di,"P")&&!S[di].M.includes(name)&&covM(di)<MAX&&countPS(di)<1){
        S[di].P.splice(S[di].P.indexOf(name),1);
        S[di].M.push(name); S[di].PS.push(name); psUsed.add(name);
      }
    }
    if(pShort()){
      for(const {name,days} of psElig){
        if(covP(di)>=minCov(di,"P")) break;
        if(psUsed.has(name)) continue;
        if(!days.includes(di)) continue;
        if(constraintBlocks(cof(name,di),"PS")) continue;
        if(name==="Diane" && dianeConsecutivePom(di)) continue;
        if(MAX_P[name]&&countPomeriggi(name)>=MAX_P[name]) continue;
        if(S[di].M.includes(name)&&!S[di].PS.includes(name)&&!S[di].P.includes(name)&&countPS(di)<1){
          S[di].PS.push(name); psUsed.add(name);
        }
      }
    }
    if(pShort()){
      for(const {name,days} of psElig){
        if(covP(di)>=minCov(di,"P")) break;
        if(psUsed.has(name)) continue;
        if(!days.includes(di)) continue;
        if(constraintBlocks(cof(name,di),"M")||constraintBlocks(cof(name,di),"PS")) continue;
        if(name==="Diane" && dianeConsecutivePom(di)) continue;
        if(S[di].M.includes(name)||covM(di)>=MAX||turniInGiorno(name,di)>0) continue;
        if(MAX_DAYS[name]&&workingDays(name)>=MAX_DAYS[name]) continue;
        if(MAX_P[name]&&countPomeriggi(name)>=MAX_P[name]) continue;
        if(countPS(di)>=1) continue;
        S[di].M.push(name); S[di].PS.push(name); psUsed.add(name);
      }
    }
  }

  // ── SCORE ─────────────────────────────────────────────────────────
  let score=0,maxScore=0;
  for(let di=0;di<6;di++){
    if(di===SAB){ maxScore+=MAX; score+=Math.min(covU(),MAX); }
    else {
      ["M","P"].forEach(t=>{
        const req=minCov(di,t),got=t==="M"?covM(di):covP(di);
        maxScore+=req; score+=Math.min(got,req);
      });
      if(covPnorm(di)<2&&covP(di)>0) score-=1;
      if(countPS(di)>1) score-=1;
    }
  }

  const turni={};
  DAYS.forEach((day,di)=>{
    turni[day]={};
    if(di===SAB){ turni[day]["Turno Unico"]=[...S[di].U].sort(); }
    else {
      turni[day]["Mattina"]=[...S[di].M].sort();
      if(S[di].P.length>0)  turni[day]["Pomeriggio"]=[...S[di].P].sort();
      if(S[di].PS.length>0) turni[day]["PS"]=[...S[di].PS].sort();
    }
  });
  return {turni,warnings,consueloSab,consueloRiposo,score,maxScore};
}

function generateBest(weekStart,sabato,constraints){
  let best=null;
  for(let i=0;i<N_ITER;i++){
    const res=simulate(weekStart,Math.floor(Math.random()*2**32),sabato,constraints);
    if(!best||res.score>best.score) best=res;
    if(best.score===best.maxScore) break;
  }
  return best;
}

function getMonday(date){
  const d=new Date(date),day=d.getDay();
  d.setDate(d.getDate()+(day===0?1:day===1?0:-(day-1)));
  d.setHours(0,0,0,0); return d;
}
function formatDate(d){ return d.toLocaleDateString("it-IT",{day:"2-digit",month:"2-digit",year:"numeric"}); }

function exportXLS(turni,weekStart,showOre,constraints,overrides){
  // Carica SheetJS se non già presente
  function doExport(){
    const XLSX=window.XLSX;
    const label=(p,day,shift)=>{ const k=`${p}|${day}|${shift}`; return overrides[k]!==undefined?overrides[k]:showOre?cellLabelOre(p,day,shift):cellLabel(p,day,shift); };

    // Costruisci array di array per il foglio
    const dateRow=["Personale",...DAYS.map((d,i)=>{
      const dd=new Date(weekStart); dd.setDate(dd.getDate()+i);
      return `${d} ${dd.getDate()}`;
    }),"Tot"];

    const dataRows=STAFF.map(person=>{
      const cells=DAYS.map((day,di)=>{
        const shifts=Object.entries(turni[day]||{}).filter(([,arr])=>arr.includes(person)).map(([s])=>s);
        const cv=constraints?.[person]?.[di];
        if(cv==="abs") return "assente";
        if(!shifts.length) return "—";
        return shifts.map(shift=>label(person,day,shift)).join(" | ");
      });
      const tot=DAYS.reduce((acc,day)=>{
        const sh=Object.entries(turni[day]||{}).filter(([,arr])=>arr.includes(person)).map(([s])=>s);
        if(!sh.length) return acc;
        if(sh.includes("Turno Unico")) return acc+1;
        return sh.some(s=>["Mattina","PS","Pomeriggio"].includes(s))?acc+1:acc;
      },0);
      return [person,...cells,tot];
    });

    const ws=XLSX.utils.aoa_to_sheet([dateRow,...dataRows]);

    // Larghezze colonne
    ws["!cols"]=[{wch:14},...DAYS.map(()=>({wch:22})),{wch:6}];

    // Stili header (sfondo indigo, testo bianco)
    const range=XLSX.utils.decode_range(ws["!ref"]);
    for(let C=range.s.c;C<=range.e.c;C++){
      const addr=XLSX.utils.encode_cell({r:0,c:C});
      if(!ws[addr]) continue;
      ws[addr].s={
        font:{bold:true,color:{rgb:"FFFFFF"},sz:11},
        fill:{fgColor:{rgb:"4F46E5"}},
        alignment:{horizontal:"center",vertical:"center",wrapText:true},
        border:{bottom:{style:"thin",color:{rgb:"3730A3"}}},
      };
    }
    // Stili righe dati
    for(let R=1;R<=dataRows.length;R++){
      const bg=R%2===0?"F9FAFB":"FFFFFF";
      for(let C=range.s.c;C<=range.e.c;C++){
        const addr=XLSX.utils.encode_cell({r:R,c:C});
        if(!ws[addr]) ws[addr]={t:"s",v:""};
        ws[addr].s={
          fill:{fgColor:{rgb:bg}},
          font:{sz:10,bold:C===0},
          alignment:{horizontal:C===0?"left":"center",vertical:"center",wrapText:true},
          border:{
            top:{style:"thin",color:{rgb:"E5E7EB"}},
            bottom:{style:"thin",color:{rgb:"E5E7EB"}},
            left:{style:"thin",color:{rgb:"E5E7EB"}},
            right:{style:"thin",color:{rgb:"E5E7EB"}},
          },
        };
      }
    }

    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Turni");
    XLSX.writeFile(wb,`turni_${formatDate(weekStart).replace(/\//g,"-")}.xlsx`);
  }

  if(window.XLSX){ doExport(); return; }
  const s=document.createElement("script");
  s.src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
  s.onload=doExport; document.head.appendChild(s);
}
function exportJSON(turni,weekStart){
  const a=document.createElement("a");
  a.href=URL.createObjectURL(new Blob([JSON.stringify({weekStart:formatDate(weekStart),turni},null,2)],{type:"application/json"}));
  a.download=`turni_${formatDate(weekStart).replace(/\//g,"-")}.json`; a.click();
}

function buildTableHTML(turni,weekStart,weekEnd,colDates,showOre,constraints,overrides={}){
  const label=(p,day,shift)=>{
    const key=`${p}|${day}|${shift}`;
    if(overrides[key]!==undefined) return overrides[key];
    return showOre?cellLabelOre(p,day,shift):cellLabel(p,day,shift);
  };
  const shiftColors={
    Mattina:      {bg:"#FEF3C7",fg:"#92400E",border:"#F59E0B"},
    Pomeriggio:   {bg:"#DBEAFE",fg:"#1E40AF",border:"#60A5FA"},
    PS:           {bg:"#EDE9FE",fg:"#5B21B6",border:"#A78BFA"},
    "Turno Unico":{bg:"#DCFCE7",fg:"#166534",border:"#4ADE80"},
  };
  // A4 portrait: 794px wide, colonne strette, testo a capo
  const colW=82; // px per ogni giorno
  const nameW=90;
  const totW=36;
  const totalW=nameW + colW*6 + totW; // 794px circa

  const thBase=`font-family:'Segoe UI',Arial,sans-serif;font-weight:700;color:white;font-size:11px;text-align:center;padding:8px 3px;border:1px solid #3730A3;background:#4F46E5;`;
  const headers=DAYS.map((d,i)=>`<th style="${thBase}width:${colW}px;">${d}<br/><span style="font-size:9px;font-weight:400;opacity:.8;">${colDates[i]}</span></th>`).join("");

  const rows=STAFF.map((person,pi)=>{
    const bg=pi%2===0?"#FFFFFF":"#F1F5F9";
    const cells=DAYS.map((day,di)=>{
      const shifts=Object.entries(turni[day]||{}).filter(([,arr])=>arr.includes(person)).map(([s])=>s);
      const cv=constraints?.[person]?.[di];
      const tdBase=`background:${bg};text-align:center;vertical-align:top;padding:5px 2px;border:1px solid #CBD5E1;width:${colW}px;`;
      if(cv==="abs") return `<td style="${tdBase}background:#FEE2E2;"><span style="color:#EF4444;font-size:9px;font-weight:700;">assente</span></td>`;
      if(!shifts.length) return `<td style="${tdBase}"><span style="color:#CBD5E1;font-size:11px;">—</span></td>`;
      const badges=shifts.map(shift=>{
        const isCassa=isCassaCell(person,day,shift);
        const c=isCassa?{bg:"#E2E8F0",fg:"#475569",border:"#94A3B8"}:(shiftColors[shift]||{bg:"#F1F5F9",fg:"#334155",border:"#94A3B8"});
        const lbl=label(person,day,shift);
        const lblWrapped=lbl.replace(/–|-/g,"‑<wbr/>").replace(/\s+/g,"<br/>");
        return `<table style="display:inline-table;background:${c.bg};border:1px solid ${c.border};border-radius:5px;margin:2px auto;width:${colW-12}px;min-height:32px;"><tbody><tr><td style="color:${c.fg};font-size:9px;font-weight:700;text-align:center;vertical-align:middle;padding:4px 4px;line-height:1.35;word-break:break-word;">${lblWrapped}</td></tr></tbody></table>`;
      }).join("");
      return `<td style="${tdBase}">${badges}</td>`;
    }).join("");

    const tot=DAYS.reduce((acc,day)=>{
      const sh=Object.entries(turni[day]||{}).filter(([,arr])=>arr.includes(person)).map(([s])=>s);
      if(!sh.length) return acc;
      if(sh.includes("Turno Unico")) return acc+1;
      return sh.some(s=>["Mattina","PS","Pomeriggio"].includes(s))?acc+1:acc;
    },0);

    return `<tr>
      <td style="background:${bg};font-family:'Segoe UI',Arial,sans-serif;font-weight:700;color:#1E293B;padding:6px 8px;border:1px solid #CBD5E1;width:${nameW}px;font-size:11px;white-space:nowrap;">${person}</td>
      ${cells}
      <td style="background:${bg};text-align:center;padding:6px 3px;border:1px solid #CBD5E1;width:${totW}px;">
        <span style="background:#EEF2FF;color:#4338CA;font-weight:700;padding:2px 5px;border-radius:10px;font-size:10px;font-family:'Segoe UI',Arial,sans-serif;">${tot}</span>
      </td>
    </tr>`;
  }).join("");

  const hdrNameStyle=`font-family:'Segoe UI',Arial,sans-serif;font-weight:700;color:white;font-size:11px;text-align:left;padding:8px;border:1px solid #3730A3;background:#4F46E5;width:${nameW}px;`;
  const hdrTotStyle=`font-family:'Segoe UI',Arial,sans-serif;font-weight:700;color:white;font-size:10px;text-align:center;padding:8px 3px;border:1px solid #3730A3;background:#4F46E5;width:${totW}px;`;

  return `<div style="width:${totalW}px;margin:0 auto;">
  <div style="background:#4F46E5;border-radius:8px 8px 0 0;padding:10px 14px;display:flex;justify-content:space-between;align-items:center;">
    <span style="font-family:'Segoe UI',Arial,sans-serif;font-weight:700;color:white;font-size:13px;">📅 Turni Ambulatorio</span>
    <span style="font-family:'Segoe UI',Arial,sans-serif;color:rgba(255,255,255,.8);font-size:10px;">${formatDate(weekStart)} — ${formatDate(weekEnd)}</span>
  </div>
  <table style="width:${totalW}px;border-collapse:collapse;table-layout:fixed;">
    <thead><tr>
      <th style="${hdrNameStyle}">Personale</th>
      ${headers}
      <th style="${hdrTotStyle}">Tot</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table>
  </div>`;
}

async function exportJPG(turni,weekStart,weekEnd,colDates,showOre,constraints,overrides={}){
  const tableHTML=buildTableHTML(turni,weekStart,weekEnd,colDates,showOre,constraints,overrides);
  // Larghezza canvas = A4 portrait ~794px + padding
  const canvasW=842;
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/>
  <style>*{box-sizing:border-box;margin:0;padding:0;}
  body{background:#F8FAFC;padding:24px;width:${canvasW}px;font-family:'Segoe UI',Arial,sans-serif;}
  </style></head><body>${tableHTML}</body></html>`;

  if(!window._h2c){
    await new Promise((res,rej)=>{
      const s=document.createElement("script");
      s.src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
      s.onload=res; s.onerror=rej; document.head.appendChild(s);
    });
    window._h2c=true;
  }
  const iframe=document.createElement("iframe");
  iframe.style.cssText=`position:fixed;left:-9999px;top:0;width:${canvasW}px;height:1200px;border:none;`;
  document.body.appendChild(iframe);
  iframe.contentDocument.open();
  iframe.contentDocument.write(html);
  iframe.contentDocument.close();
  await new Promise(r=>setTimeout(r,400));
  const body=iframe.contentDocument.body;
  const canvas=await window.html2canvas(body,{
    scale:2.5,
    backgroundColor:"#F8FAFC",
    useCORS:true,
    width:canvasW,
    height:body.scrollHeight+48,
    windowWidth:canvasW,
  });
  document.body.removeChild(iframe);
  const a=document.createElement("a");
  a.href=canvas.toDataURL("image/jpeg",0.96);
  a.download=`turni_${formatDate(weekStart).replace(/\//g,"-")}.jpg`;
  a.click();
}

const SS={
  Mattina:       {cell:"bg-amber-100 text-amber-800",   badge:"bg-amber-200 text-amber-900"},
  Pomeriggio:    {cell:"bg-blue-100 text-blue-800",     badge:"bg-blue-200 text-blue-900"},
  PS:            {cell:"bg-purple-100 text-purple-800", badge:"bg-purple-200 text-purple-900"},
  "Turno Unico": {cell:"bg-green-100 text-green-800",   badge:"bg-green-200 text-green-900"},
};
const CASSA_MAP={Consuelo:{Mercoledì:{Pomeriggio:true},Giovedì:{Mattina:true},Venerdì:{Pomeriggio:true}}};
function isCassaCell(p,day,shift){ return !!CASSA_MAP[p]?.[day]?.[shift]; }
function cellLabel(p,day,shift){
  if(isCassaCell(p,day,shift)) return shift==="Mattina"?"M cassa":"P cassa";
  return shift==="Turno Unico"?"Unico":shift==="PS"?"PS":shift==="Mattina"?"M":"P";
}

const CONSTRAINT_COLORS={"abs":"bg-red-100 text-red-600","noM":"bg-orange-100 text-orange-600","noP":"bg-orange-100 text-orange-600","onM":"bg-sky-100 text-sky-600","onP":"bg-sky-100 text-sky-600"};

function cellLabelOre(p,day,shift){
  if(isCassaCell(p,day,shift)) return shift==="Mattina"?"(7-14)":"(14-CH)";
  if(shift==="Mattina") return "8.00-14.30";
  if(shift==="Pomeriggio") return "14.30-21.00";
  if(shift==="PS") return "(18-21 REP)";
  if(shift==="Turno Unico") return "8-CH";
  return shift;
}

export default function App(){
  const [weekOffset,setWeekOffset]=useState(1);
  const [result,setResult]=useState(null);
  const [computing,setComputing]=useState(false);
  const [sabScelto,setSabScelto]=useState("auto");
  const [constraints,setConstraints]=useState({});
  const [showConstraints,setShowConstraints]=useState(false);
  const [showOre,setShowOre]=useState(true);
  const [dirty,setDirty]=useState(false);
  const [overrides,setOverrides]=useState<Record<string,string>>({});
  const [editingCell,setEditingCell]=useState<string|null>(null);

  function getLabel(person:string,day:string,shift:string):string{
    const key=`${person}|${day}|${shift}`;
    if(overrides[key]!==undefined) return overrides[key];
    return showOre ? cellLabelOre(person,day,shift) : cellLabel(person,day,shift);
  }
  function startEdit(key:string){ setEditingCell(key); }
  function commitEdit(key:string,val:string){
    const trimmed=val.slice(0,15);
    const defaultVal=showOre ? cellLabelOre(...(key.split("|") as [string,string,string])) : cellLabel(...(key.split("|") as [string,string,string]));
    if(trimmed===defaultVal||trimmed==="") setOverrides(p=>{const n={...p}; delete n[key]; return n;});
    else setOverrides(p=>({...p,[key]:trimmed}));
    setEditingCell(null);
  }

  const baseMonday=getMonday(new Date());
  const weekStart=new Date(baseMonday);
  weekStart.setDate(weekStart.getDate()+weekOffset*7);
  const weekEnd=new Date(weekStart); weekEnd.setDate(weekEnd.getDate()+5);
  const turni=result?.turni||null;

  const activeConstraints=Object.values(constraints).reduce((a,days)=>a+Object.values(days).filter(v=>v).length,0);

  function setConstraint(name,di,val){
    setConstraints(prev=>{
      const next={...prev,[name]:{...prev[name],[di]:val}};
      if(!val) delete next[name][di];
      if(!Object.keys(next[name]||{}).length) delete next[name];
      return next;
    });
    if(result) setDirty(true);
  }
  function clearConstraints(){ setConstraints({}); if(result) setDirty(true); }

  function handleGenerate(){
    setComputing(true);
    setDirty(false);
    setOverrides({});
    setTimeout(()=>{
      const res=generateBest(weekStart,sabScelto,constraints);
      setResult(res); setComputing(false);
    },20);
  }

  function getShifts(p,day){
    if(!turni) return [];
    return Object.entries(turni[day]||{}).filter(([,arr])=>arr.includes(p)).map(([s])=>s);
  }
  function countShifts(p){
    if(!turni) return 0;
    return DAYS.reduce((acc,day)=>{
      const shifts=getShifts(p,day);
      if(shifts.length===0) return acc;
      // Sabato: Turno Unico = 1 gg
      if(shifts.includes("Turno Unico")) return acc+1;
      const hasM   = shifts.includes("Mattina");
      const hasPS  = shifts.includes("PS");
      const hasPnorm = shifts.includes("Pomeriggio") && !isCassaCell(p,day,"Pomeriggio");
      const hasMcassa = shifts.includes("Mattina") && isCassaCell(p,day,"Mattina");
      const hasPcassa = shifts.includes("Pomeriggio") && isCassaCell(p,day,"Pomeriggio");
      // M+PS = 1 gg; M+P cassa = 1 gg; solo cassa = 1 gg; solo M o P normale = 1 gg
      // In ogni caso un giorno in cui si lavora conta 1
      if(hasM||hasPS||hasPnorm||hasMcassa||hasPcassa) return acc+1;
      return acc;
    },0);
  }
  function covDisplay(day){
    if(!turni) return 0;
    return (turni[day]?.["Pomeriggio"]||[]).filter(s=>!isCassaCell(s,day,"Pomeriggio")).length
          +(turni[day]?.["PS"]||[]).length;
  }
  function handleImport(e){
    const f=e.target.files[0]; if(!f) return;
    const r=new FileReader();
    r.onload=ev=>{ try{ const d=JSON.parse(ev.target.result); if(d.turni) setResult({turni:d.turni,warnings:[],consueloSab:false,score:0,maxScore:0}); }catch{alert("JSON non valido");} };
    r.readAsText(f); e.target.value="";
  }

  function printPDF(){
    const tableHTML=buildTableHTML(turni,weekStart,weekEnd,colDates,showOre,constraints,overrides);
    const html=`<!DOCTYPE html><html><head><meta charset="utf-8"/><title>Turni ${formatDate(weekStart)}</title>
    <style>@page{size:A4 landscape;margin:10mm} body{font-family:'Segoe UI',Arial,sans-serif;background:#F8FAFC;padding:16px;}</style>
    </head><body>${tableHTML}<script>window.onload=()=>window.print();<\/script></body></html>`;
    const w=window.open("","_blank");
    w.document.write(html);
    w.document.close();
  }

  const pct=result&&result.maxScore>0?Math.round(result.score/result.maxScore*100):0;
  const colDates=DAYS.map((_,i)=>{ const d=new Date(weekStart); d.setDate(d.getDate()+i); return d.getDate(); });

  return (
    <div className="min-h-screen bg-gray-50 p-4">
      <div className="max-w-5xl mx-auto space-y-4">

        {/* HEADER */}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-200 p-4">
          <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
            <div>
              <h1 className="text-xl font-bold text-gray-800">📅 Turni Ambulatorio</h1>
              <p className="text-sm text-gray-500 mt-0.5">Settimana {formatDate(weekStart)} — {formatDate(weekEnd)}</p>
            </div>
            <div className="flex gap-2">
              <button onClick={()=>{setWeekOffset(w=>w-1);setResult(null);}} className="px-3 py-1.5 rounded-lg border border-gray-200 text-sm hover:bg-gray-50">← Prec</button>
              <button onClick={()=>{setWeekOffset(1);setResult(null);}} className="px-3 py-1.5 rounded-lg border border-gray-200 text-sm hover:bg-gray-50">Prossima</button>
              <button onClick={()=>{setWeekOffset(w=>w+1);setResult(null);}} className="px-3 py-1.5 rounded-lg border border-gray-200 text-sm hover:bg-gray-50">Succ →</button>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 mb-3 p-3 bg-gray-50 rounded-xl border border-gray-100">
            <span className="text-xs font-semibold text-gray-600 mr-1">🗓 Sabato:</span>
            {[{val:"auto",label:"🔁 Alterne (auto)"},{val:"Claudia",label:"Claudia"},{val:"Consuelo",label:"Consuelo"}].map(({val,label})=>(
              <button key={val} onClick={()=>setSabScelto(val)}
                className={`px-3 py-1.5 rounded-lg text-xs font-semibold border transition-all ${sabScelto===val?"bg-indigo-600 text-white border-indigo-600":"bg-white text-gray-600 border-gray-200 hover:border-indigo-300"}`}>
                {label}
              </button>
            ))}
          </div>

          <div className="flex flex-wrap gap-2">
            <button onClick={handleGenerate} disabled={computing}
              className={`px-4 py-2 rounded-lg text-sm font-medium shadow-sm text-white ${computing?"bg-indigo-400 cursor-wait":dirty?"bg-orange-500 hover:bg-orange-600":"bg-indigo-600 hover:bg-indigo-700"}`}>
              {computing?"⏳ Ottimizzazione…":dirty?"⚠️ Rielabora Turni":"✨ Elabora Turni"}
            </button>
            <button onClick={()=>setShowOre(v=>!v)}
              className={`px-4 py-2 rounded-lg text-sm font-medium border shadow-sm transition-all ${showOre?"bg-indigo-50 border-indigo-400 text-indigo-700":"bg-white border-gray-200 text-gray-600 hover:border-indigo-300"}`}>
              🕐 {showOre?"Mostra sigle":"Mostra orari"}
            </button>
            {turni && <>
              <button onClick={()=>exportXLS(turni,weekStart,showOre,constraints,overrides)} className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-sm font-medium shadow-sm">⬇ XLSX</button>
              <button onClick={printPDF} className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-lg text-sm font-medium shadow-sm">🖨 PDF</button>
              <button onClick={()=>exportJPG(turni,weekStart,weekEnd,colDates,showOre,constraints,overrides)} className="px-4 py-2 bg-violet-600 hover:bg-violet-700 text-white rounded-lg text-sm font-medium shadow-sm">🖼 JPG</button>
            </>}
          </div>

          {turni && (
            <div className="mt-3 flex flex-wrap items-center gap-4">
              <p className="text-xs text-gray-500">
                🗓 Sabato: <strong>{result.consueloSab?"Consuelo":"Claudia"}</strong>
                {result.consueloRiposo!==null&&<>&nbsp;·&nbsp;🏠 Riposo Consuelo: <strong>{DAYS[result.consueloRiposo]}</strong></>}
              </p>
              <div className="flex items-center gap-2">
                <div className="w-32 h-2 bg-gray-200 rounded-full overflow-hidden">
                  <div className={`h-full rounded-full ${pct===100?"bg-green-500":pct>=80?"bg-yellow-400":"bg-red-400"}`} style={{width:`${pct}%`}}/>
                </div>
                <span className={`text-xs font-bold ${pct===100?"text-green-600":pct>=80?"text-yellow-600":"text-red-600"}`}>{pct}% ({result.score}/{result.maxScore})</span>
              </div>
            </div>
          )}
        </div>

        {/* PANNELLO VINCOLI */}
        <div className="bg-white rounded-2xl shadow-sm border border-gray-200 overflow-hidden">
          <button onClick={()=>setShowConstraints(v=>!v)}
            className="w-full flex items-center justify-between px-4 py-3 hover:bg-gray-50 transition-colors">
            <div className="flex items-center gap-2">
              <span className="text-sm font-semibold text-gray-700">⚙️ Vincoli settimanali</span>
              {activeConstraints>0&&(
                <span className="text-xs bg-orange-100 text-orange-700 font-bold px-2 py-0.5 rounded-full">{activeConstraints} attivi</span>
              )}
            </div>
            <span className="text-gray-400 text-xs">{showConstraints?"▲ chiudi":"▼ apri"}</span>
          </button>

          {showConstraints&&(
            <div className="px-4 pb-4 border-t border-gray-100">
              <p className="text-xs text-gray-400 mt-3 mb-3">Imposta eccezioni per questa settimana. L'algoritmo le rispetterà in tutte le simulazioni.</p>
              <div className="flex flex-wrap gap-2 mb-3">
                {CONSTRAINT_OPTS.filter(o=>o.val).map(o=>(
                  <span key={o.val} className={`text-xs px-2 py-0.5 rounded-full font-medium ${CONSTRAINT_COLORS[o.val]||"bg-gray-100 text-gray-600"}`}>{o.label}</span>
                ))}
              </div>
              <div className="overflow-auto">
                <table className="w-full text-xs border-collapse">
                  <thead>
                    <tr className="bg-gray-50">
                      <th className="text-left px-3 py-2 font-semibold text-gray-600 w-24">Persona</th>
                      {DAYS.map((d,i)=>(
                        <th key={d} className="px-2 py-2 text-center font-semibold text-gray-600">
                          <div>{DAYS_SHORT[i]}</div>
                          <div className="font-normal text-gray-400">{colDates[i]}</div>
                        </th>
                      ))}
                      <th className="px-2 py-2"/>
                    </tr>
                  </thead>
                  <tbody>
                    {STAFF.map((name,pi)=>(
                      <tr key={name} className={pi%2===0?"bg-white":"bg-gray-50"}>
                        <td className="px-3 py-1.5 font-semibold text-gray-700">{name}</td>
                        {DAYS.map((_,di)=>{
                          const val=constraints?.[name]?.[di]||"";
                          const opts = di===SAB ? CONSTRAINT_OPTS_SAB : CONSTRAINT_OPTS;
                          return (
                            <td key={di} className="px-1 py-1 text-center">
                              <select value={val} onChange={e=>setConstraint(name,di,e.target.value)}
                                className={`text-xs rounded border px-1 py-0.5 w-full max-w-[90px] ${val?CONSTRAINT_COLORS[val]+" border-transparent":"border-gray-200 text-gray-500"}`}>
                                {opts.map(o=><option key={o.val} value={o.val}>{o.label}</option>)}
                              </select>
                            </td>
                          );
                        })}
                        <td className="px-2 py-1">
                          {constraints?.[name]&&Object.keys(constraints[name]).length>0&&(
                            <button onClick={()=>setConstraints(p=>{const n={...p}; delete n[name]; return n;})}
                              className="text-xs text-red-400 hover:text-red-600">✕</button>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {activeConstraints>0&&(
                <button onClick={clearConstraints} className="mt-3 text-xs text-red-500 hover:text-red-700 font-medium">🗑 Cancella tutti i vincoli</button>
              )}
            </div>
          )}
        </div>

        {/* RISULTATI */}
        {turni&&(
          <div className="space-y-3">
            {(()=>{
              const issues=[];
              DAYS.forEach((day,di)=>{
                if(di===SAB){
                  const c=(turni[day]?.["Turno Unico"]||[]).length;
                  if(c<MAX) issues.push(`${day} Turno Unico: ${c}/${MAX} operatori`);
                } else {
                  const mCov=(turni[day]?.["Mattina"]||[]).filter(s=>!isCassaCell(s,day,"Mattina")).length;
                  if(mCov<minCov(di,"M")) issues.push(`${day} Mattina: ${mCov}/${minCov(di,"M")} operatori`);
                  const pNorm=(turni[day]?.["Pomeriggio"]||[]).filter(s=>!isCassaCell(s,day,"Pomeriggio")).length;
                  const psCount=(turni[day]?.["PS"]||[]).length;
                  const pCov=pNorm+psCount;
                  if(pCov<minCov(di,"P")) issues.push(`${day} Pomeriggio: ${pCov}/${minCov(di,"P")} operatori`);
                  if(pNorm<2&&pCov>0) issues.push(`${day} Pomeriggio: solo ${pNorm} P normali (min 2)`);
                  if(psCount>1) issues.push(`${day} Pomeriggio: ${psCount} PS (max 1)`);
                }
              });
              result.warnings.forEach(w=>{ if(!issues.includes(w)) issues.push(w); });
              return issues.length>0?(
                <div className="bg-red-50 border border-red-200 rounded-xl p-4">
                  <h2 className="text-sm font-bold text-red-700 mb-2">❌ Regole non rispettate ({issues.length})</h2>
                  {issues.map((w,i)=><p key={i} className="text-xs text-red-600 mb-0.5">• {w}</p>)}
                  <p className="text-xs text-red-400 mt-2">Migliore su {N_ITER} simulazioni.</p>
                </div>
              ):(
                <div className="bg-green-50 border border-green-200 rounded-xl p-3 flex items-center gap-2">
                  <span className="text-green-600 text-lg">✅</span>
                  <p className="text-xs font-semibold text-green-700">Tutte le regole rispettate — copertura 100%</p>
                </div>
              );
            })()}

            <div className="flex gap-2 flex-wrap items-center">
              {Object.entries(SS).map(([s,st])=>(
                <span key={s} className={`text-xs font-semibold px-3 py-1 rounded-full ${st.badge}`}>{s==="PS"?"PS — serale":s}</span>
              ))}
              <span className="text-xs bg-gray-200 text-gray-600 px-3 py-1 rounded-full font-semibold">cassa</span>
            </div>

            <div className="bg-white rounded-2xl shadow-sm border border-gray-200 overflow-auto">
              <table className="w-full text-sm border-collapse">
                <thead>
                  <tr className="bg-indigo-600 text-white">
                    <th className="text-left px-4 py-3 font-semibold w-32 rounded-tl-2xl">Personale</th>
                    {DAYS.map((d,i)=>(
                      <th key={d} className="px-3 py-3 font-semibold text-center">
                        <div>{d}</div>
                        <div className="text-xs font-normal opacity-75">{colDates[i]}</div>
                      </th>
                    ))}
                    <th className="px-3 py-3 font-semibold text-center rounded-tr-2xl">Tot</th>
                  </tr>
                </thead>
                <tbody>
                  {STAFF.map((person,pi)=>(
                    <tr key={person} className={pi%2===0?"bg-white":"bg-gray-50"}>
                      <td className="px-4 py-3 font-semibold text-gray-700 border-r border-gray-100 whitespace-nowrap">{person}</td>
                      {DAYS.map((day,di)=>{
                        const shifts=getShifts(person,day);
                        const cv=constraints?.[person]?.[di];
                        return (
                          <td key={day} className={`px-2 py-2 text-center border-r border-gray-100 min-w-[100px] ${cv==="abs"?"bg-red-50":""}`}>
                            {cv==="abs"
                              ?<span className="text-xs text-red-300 font-medium">assente</span>
                              :shifts.length===0
                                ?<span className="text-gray-200">—</span>
                                :<div className="flex flex-col gap-1 items-center">
                                  {shifts.map(shift=>{
                                    const key=`${person}|${day}|${shift}`;
                                    const lbl=getLabel(person,day,shift);
                                    const isEditing=editingCell===key;
                                    const hasOverride=overrides[key]!==undefined;
                                    const cellCls=`text-xs font-semibold px-1.5 py-0.5 rounded ${isCassaCell(person,day,shift)?"bg-gray-200 text-gray-600":SS[shift]?.cell||"bg-gray-100"}`;
                                    return isEditing
                                      ? <input key={shift} autoFocus
                                          defaultValue={lbl} maxLength={15}
                                          className={`${cellCls} w-24 text-center outline outline-2 outline-indigo-400`}
                                          onBlur={e=>commitEdit(key,e.target.value)}
                                          onKeyDown={e=>{ if(e.key==="Enter") commitEdit(key,(e.target as HTMLInputElement).value); if(e.key==="Escape"){setEditingCell(null);} }}
                                        />
                                      : <span key={shift} title="Clicca per modificare"
                                          onClick={()=>startEdit(key)}
                                          className={`${cellCls} cursor-pointer hover:ring-2 hover:ring-indigo-300 transition-all ${hasOverride?"ring-1 ring-amber-400":""}`}>
                                          {lbl}
                                        </span>;
                                  })}
                                </div>
                            }
                          </td>
                        );
                      })}
                      <td className="px-3 py-3 text-center">
                        <span className="text-xs font-bold bg-indigo-100 text-indigo-700 rounded-full px-2 py-0.5">{countShifts(person)}</span>
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  {[
                    {label:"n° Mattina",    fn:(day,di)=>(turni[day]?.["Mattina"]||[]).filter(s=>!isCassaCell(s,day,"Mattina")).length, min:di=>minCov(di,"M"), skip:di=>di===SAB},
                    {label:"n° Pomeriggio", fn:(day)=>covDisplay(day), min:di=>minCov(di,"P"), skip:di=>di===SAB},
                  ].map(({label,fn,min,skip})=>(
                    <tr key={label} className="bg-gray-100 border-t border-gray-200">
                      <td className="px-4 py-2 text-xs text-gray-500 font-medium">{label}</td>
                      {DAYS.map((day,di)=>{
                        if(di===SAB) return <td key={day} className="border-r border-gray-200"/>;
                        if(skip(di)) return <td key={day} className="border-r border-gray-200"/>;
                        const count=fn(day,di),m=min(di);
                        return (
                          <td key={day} className="px-2 py-2 text-center text-xs border-r border-gray-200">
                            <span className={`font-bold px-1.5 py-0.5 rounded ${count<m?"bg-red-100 text-red-700":count===MAX?"bg-green-100 text-green-700":"bg-yellow-100 text-yellow-700"}`}>{count}</span>
                          </td>
                        );
                      })}
                      <td/>
                    </tr>
                  ))}
                </tfoot>
              </table>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-4">
              <h2 className="text-sm font-semibold text-gray-700 mb-3">Regole applicate</h2>
              <div className="grid grid-cols-2 md:grid-cols-3 gap-2 text-xs text-gray-600">
                {[
                  ["Claudia","Con sab: 1M+3P lun-ven + sab TU · Senza sab: 2M+3P lun-ven"],
                  ["Consuelo","lun P · mar M · mer M+P-cassa · gio M-cassa · ven P-cassa · riposo solo sett. con sab · lun/mar → M+PS"],
                  ["Diane","4gg lun-ven · 2M · max 2 pomeriggi · PS reperibilità max 1/sett"],
                  ["Giorgia","6gg · sab TU · 2M+3P · PS reperibilità max 1/sett"],
                  ["Giulia","solo lun P e ven P"],
                  ["Mary","6gg · sab TU · 3gg M+PS · 1gg M · 1gg P"],
                ].map(([name,rule])=>(
                  <div key={name} className="bg-gray-50 rounded-lg p-2 border border-gray-100">
                    <span className="font-semibold text-indigo-700">{name}</span>
                    <p className="mt-0.5 leading-relaxed">{rule}</p>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {!turni&&(
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-12 text-center text-gray-400">
            <div className="text-4xl mb-3">📋</div>
            <p className="text-sm">Clicca <strong>Elabora Turni</strong> per avviare l'ottimizzazione</p>
            <p className="text-xs mt-1 text-gray-300">{N_ITER} simulazioni · restituisce la copertura migliore</p>
          </div>
        )}
      </div>
    </div>
  );
}
