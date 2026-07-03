import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import jsQR from "jsqr";
import QRCode from "qrcode";

// ── Google Fonts ────────────────────────────────────────────────────────
const _fl = document.createElement("link"); _fl.rel="stylesheet";
_fl.href="https://fonts.googleapis.com/css2?family=Noto+Serif+Ethiopic:wght@600;700&family=DM+Sans:wght@300;400;500;600&display=swap";
document.head.appendChild(_fl);

// ══════════════════════════════════════════════════════════════════════
// ⚙️  SUPABASE CONFIG — paste your values here after creating project
// ══════════════════════════════════════════════════════════════════════
const SUPABASE_URL  = "https://gknyuurcbiwivbjvscee.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imdrbnl1dXJjYml3aXZianZzY2VlIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODA1OTExMjIsImV4cCI6MjA5NjE2NzEyMn0.XPamJb0EtrDftt6t1oO_oeDTLBxjmIGMJVUzJQhXATY";
const AUDIO_BUCKET  = "church-audio";

// ── Supabase helpers ────────────────────────────────────────────────────
const sb = {
  headers: { "Content-Type":"application/json", "apikey": SUPABASE_ANON, "Authorization": `Bearer ${SUPABASE_ANON}` },

  async get(table, params="") {
    const r = await fetch(`${SUPABASE_URL}/rest/v1/${table}?${params}`, { headers: sb.headers });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  async insert(table, body) {
    const r = await fetch(`${SUPABASE_URL}/rest/v1/${table}`, {
      method:"POST", headers:{...sb.headers, "Prefer":"return=representation"}, body: JSON.stringify(body)
    });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  async update(table, match, body) {
    const params = Object.entries(match).map(([k,v])=>`${k}=eq.${encodeURIComponent(v)}`).join("&");
    const r = await fetch(`${SUPABASE_URL}/rest/v1/${table}?${params}`, {
      method:"PATCH", headers:{...sb.headers,"Prefer":"return=representation"}, body: JSON.stringify(body)
    });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  async delete(table, match) {
    const params = Object.entries(match).map(([k,v])=>`${k}=eq.${encodeURIComponent(v)}`).join("&");
    const r = await fetch(`${SUPABASE_URL}/rest/v1/${table}?${params}`, { method:"DELETE", headers: sb.headers });
    if (!r.ok) throw new Error(await r.text());
  },
  async uploadAudio(file, fileName) {
    const ext = fileName.split('.').pop().toLowerCase();
    const mimeMap = { mp3:"audio/mpeg", ogg:"audio/ogg", wav:"audio/wav", m4a:"audio/mp4", aac:"audio/aac", flac:"audio/flac" };
    const contentType = mimeMap[ext] || file.type || "audio/mpeg";
    const safeName = fileName.replace(/[^a-zA-Z0-9._-]/g, '_');
    const path = `${Date.now()}_${safeName}`;
    const r = await fetch(`${SUPABASE_URL}/storage/v1/object/${AUDIO_BUCKET}/${path}`, {
      method:"POST",
      headers:{ "apikey": SUPABASE_ANON, "Authorization":`Bearer ${SUPABASE_ANON}`, "Content-Type": contentType, "x-upsert": "true" },
      body: file
    });
    if (!r.ok) {
      const errText = await r.text();
      throw new Error(`Storage error ${r.status}: ${errText}`);
    }
    return `${SUPABASE_URL}/storage/v1/object/public/${AUDIO_BUCKET}/${path}`;
  },
  async deleteAudio(url) {
    const path = url.split(`${AUDIO_BUCKET}/`)[1];
    if (!path) return;
    await fetch(`${SUPABASE_URL}/storage/v1/object/${AUDIO_BUCKET}/${path}`, {
      method:"DELETE", headers:{ "apikey": SUPABASE_ANON, "Authorization":`Bearer ${SUPABASE_ANON}` }
    });
  },
  configured() { return SUPABASE_URL !== "YOUR_SUPABASE_URL"; }
};

// ── Brand colors ────────────────────────────────────────────────────────
const C = {
  teal:"#0e4f4f", tealMid:"#155c5c", tealLight:"#1d6e6e", tealBorder:"#2a7a7a",
  gold:"#f0b429", goldLight:"#f5c84a", goldDim:"#f0b42933", goldFaint:"#f0b42911",
  text:"#fdf6e3", textMuted:"#a8c4c4", textDim:"#6a9a9a", bg:"#0a3a3a",
  success:"#4ecca3", successBg:"#0a3a2a", danger:"#e07070", dangerBg:"#3a1515", warn:"#f0b429", warnBg:"#3a2a08",
};

// ── Constants ───────────────────────────────────────────────────────────
const GROUPS      = ["General","Youth","Choir","Women's Fellowship","Men's Group","Children","Sunday School"];
const EVENT_TYPE  = "መዝሙር ጥናት";
const EVENT_ID    = "EVT-001";
const ADMIN_PIN   = "1234";
const DEFAULT_CATS= ["ምስጋና","አምልኮ","ወጣቶች","ልጆች","ትንሳኤ","ሌላ"];

// ── Helpers ─────────────────────────────────────────────────────────────
const today   = () => new Date().toISOString().split("T")[0];
const fmtDate = d => new Date(d+"T00:00:00").toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"});
const initials= n => n.split(" ").map(p=>p[0]).join("").slice(0,2).toUpperCase();
const uid     = () => "M" + Date.now().toString(36).toUpperCase();
const sid     = () => "S" + Date.now().toString(36).toUpperCase();
function missedWeeks(memberId, records) {
  const now = new Date();
  for (let i=0; i<4; i++) {
    const d = new Date(now); d.setDate(d.getDate()-i*7);
    const wk = d.toISOString().split("T")[0].slice(0,7);
    if (records.some(r=>r.member_id===memberId&&r.date.startsWith(wk))) return i;
  }
  return 4;
}

// ── CSS ─────────────────────────────────────────────────────────────────
const _s=document.createElement("style");
_s.textContent=`
*{box-sizing:border-box;margin:0;padding:0;}
body{background:#0a3a3a;}
@keyframes fadeUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
@keyframes scanLine{0%{top:8%}50%{top:86%}100%{top:8%}}
@keyframes spin{to{transform:rotate(360deg)}}
::-webkit-scrollbar{width:4px}::-webkit-scrollbar-track{background:#0e4f4f}::-webkit-scrollbar-thumb{background:#f0b42955;border-radius:2px}
.row-hover:hover{background:#1d6e6e !important;cursor:pointer;}
.btn-gold{background:#f0b429;color:#0a3a3a;border:none;cursor:pointer;font-family:'DM Sans',sans-serif;font-weight:600;border-radius:8px;transition:all .18s;}
.btn-gold:hover{background:#f5c84a;}
.btn-outline{background:transparent;border:1px solid #f0b429;color:#f0b429;cursor:pointer;font-family:'DM Sans',sans-serif;border-radius:8px;transition:all .18s;}
.btn-outline:hover{background:#f0b42922;}
.btn-danger{background:transparent;border:1px solid #e07070;color:#e07070;cursor:pointer;font-family:'DM Sans',sans-serif;border-radius:8px;transition:all .18s;}
.btn-danger:hover{background:#e0707022;}
input,select,textarea{font-family:'DM Sans',sans-serif;outline:none;}
input:focus,select:focus,textarea:focus{border-color:#f0b429 !important;box-shadow:0 0 0 2px #f0b42922;}
`;
document.head.appendChild(_s);

// ── Shared UI components ────────────────────────────────────────────────
function Avatar({name,size=40,color}){
  const col=color||C.gold;
  return <div style={{width:size,height:size,borderRadius:"50%",background:`${col}22`,border:`1.5px solid ${col}55`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:size*.36,fontWeight:600,color:col,fontFamily:"'DM Sans',sans-serif",flexShrink:0}}>{initials(name)}</div>;
}
function Toast({msg,type,onDone}){
  useEffect(()=>{const t=setTimeout(onDone,3000);return()=>clearTimeout(t);},[]);
  const col=type==="success"?C.success:type==="warn"?C.warn:C.danger;
  const bg =type==="success"?C.successBg:type==="warn"?C.warnBg:C.dangerBg;
  return <div style={{position:"fixed",bottom:32,left:"50%",transform:"translateX(-50%)",background:bg,border:`1px solid ${col}44`,color:col,padding:"12px 24px",borderRadius:10,fontSize:14,fontFamily:"'DM Sans',sans-serif",zIndex:9999,boxShadow:"0 8px 32px #0008",whiteSpace:"nowrap",animation:"fadeUp .3s ease",maxWidth:"90vw",textOverflow:"ellipsis",overflow:"hidden"}}>{msg}</div>;
}
function Spinner({size=20,color}){ return <div style={{width:size,height:size,border:`2px solid ${(color||C.gold)}33`,borderTop:`2px solid ${color||C.gold}`,borderRadius:"50%",animation:"spin .7s linear infinite",flexShrink:0}}/>; }
function Card({children,style:s={}}){ return <div style={{background:C.teal,border:`1px solid ${C.tealBorder}55`,borderRadius:14,overflow:"hidden",...s}}>{children}</div>; }
function Section({label,children,right}){
  return <div style={{marginBottom:20}}>
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
      <div style={{fontSize:11,color:C.textDim,letterSpacing:".1em",textTransform:"uppercase",paddingLeft:2}}>{label}</div>{right}
    </div>{children}
  </div>;
}
function StatCard({label,val,color,icon}){
  const c=color||C.gold;
  return <div style={{background:C.teal,border:`1px solid ${c}33`,borderRadius:12,padding:"18px 14px",textAlign:"center"}}>
    <div style={{fontSize:22,marginBottom:4}}>{icon}</div>
    <div style={{fontSize:26,fontFamily:"'Noto Serif Ethiopic',serif",color:c,fontWeight:700,lineHeight:1.1}}>{val}</div>
    <div style={{fontSize:11,color:C.textDim,marginTop:4,textTransform:"uppercase",letterSpacing:".07em"}}>{label}</div>
  </div>;
}
function FieldInput({value,onChange,placeholder,type="text",style:s={}}){
  return <input type={type} value={value} onChange={onChange} placeholder={placeholder}
    style={{width:"100%",padding:"10px 14px",borderRadius:8,border:`1px solid ${C.tealBorder}`,background:C.bg,color:C.text,fontSize:13,...s}}/>;
}
function FieldSelect({value,onChange,children,full,style:s={}}){
  return <select value={value} onChange={onChange}
    style={{width:full?"100%":"auto",padding:"10px 14px",borderRadius:8,border:`1px solid ${C.tealBorder}`,background:C.bg,color:C.text,fontSize:13,...s}}>
    {children}
  </select>;
}
function ChurchLogo({size=48}){
  return <svg width={size} height={size} viewBox="0 0 100 100" fill="none">
    <path d="M10 62 Q50 48 90 62 L85 72 Q50 58 15 72 Z" fill={C.gold}/>
    <path d="M12 60 Q50 46 88 60" stroke={C.teal} strokeWidth="1.5" fill="none"/>
    <path d="M10 62 Q20 45 50 50 L50 58 Q25 54 15 68 Z" fill={C.goldLight}/>
    <path d="M90 62 Q80 45 50 50 L50 58 Q75 54 85 68 Z" fill={C.goldLight}/>
    <ellipse cx="50" cy="48" rx="16" ry="12" fill={C.gold}/>
    <path d="M34 50 Q30 38 36 34 Q40 44 50 42 Q60 44 64 34 Q70 38 66 50" fill={C.gold}/>
    <path d="M38 46 Q35 36 40 32 Q42 40 50 39 Q58 40 60 32 Q65 36 62 46" fill={C.goldLight}/>
    <circle cx="50" cy="28" r="4" fill={C.gold}/>
    <rect x="48.5" y="14" width="3" height="16" rx="1.5" fill={C.gold}/>
    <rect x="43" y="18" width="14" height="3" rx="1.5" fill={C.gold}/>
    <rect x="46" y="14" width="8" height="2" rx="1" fill={C.gold}/>
  </svg>;
}
// QR prefix so the scanner can tell a Finot Tiguhan member badge apart from
// any other QR code someone might point the camera at.
const QR_PREFIX = "FT-MEMBER:";

// Renders a REAL, camera-scannable QR code (previous version was a fake
// hash-based pattern that could never actually be decoded by anything).
function QRCodeSVG({value,size=130}){
  const canvasRef=useRef(null);
  useEffect(()=>{
    const canvas=canvasRef.current;
    if(!canvas)return;
    QRCode.toCanvas(canvas, QR_PREFIX+value, {
      width:size, margin:1,
      color:{ dark:C.bg, light:"#ffffff" },
      errorCorrectionLevel:"M",
    }, err=>{ if(err) console.error("QR generation failed:", err); });
  },[value,size]);
  return <canvas ref={canvasRef} width={size} height={size} style={{display:"block",width:size,height:size}}/>;
}

// ── Excel export ────────────────────────────────────────────────────────
function exportExcel(members,attendance,filterMemberId){
  const wb=XLSX.utils.book_new();
  const rows=attendance.filter(r=>!filterMemberId||r.member_id===filterMemberId)
    .map(r=>{const m=members.find(x=>x.id===r.member_id)||{};
      return{"Date":r.date,"Time":r.time,"Member ID":r.member_id,"Name":m.name||"","Group":m.group||"","Role":m.role||"","Service":r.event_type};})
    .sort((a,b)=>b.Date.localeCompare(a.Date));
  const ws1=XLSX.utils.json_to_sheet(rows);
  ws1["!cols"]=[{wch:12},{wch:8},{wch:10},{wch:22},{wch:20},{wch:12},{wch:16}];
  XLSX.utils.book_append_sheet(wb,ws1,filterMemberId?"Member History":"Attendance Log");
  if(!filterMemberId){
    const summary=members.map(m=>{
      const total=attendance.filter(r=>r.member_id===m.id).length;
      const missed=missedWeeks(m.id,attendance);
      return{"Member ID":m.id,"Name":m.name,"Phone":m.phone,"Group":m.group,"Role":m.role,"Joined":m.joined,"Total Sessions":total,"Weeks Missed":missed};
    });
    const ws2=XLSX.utils.json_to_sheet(summary);
    ws2["!cols"]=[{wch:10},{wch:22},{wch:14},{wch:20},{wch:12},{wch:12},{wch:14},{wch:14}];
    XLSX.utils.book_append_sheet(wb,ws2,"Member Summary");
  }
  XLSX.writeFile(wb,filterMemberId?`member_history_${filterMemberId}.xlsx`:`ፍኖት_ትጉሃን_${today()}.xlsx`);
}

// ── Not configured banner ───────────────────────────────────────────────
function SetupBanner(){
  return <div style={{background:"#1a2a1a",border:`1px solid ${C.gold}55`,borderRadius:12,padding:20,marginBottom:20}}>
    <div style={{fontSize:15,fontWeight:600,color:C.gold,marginBottom:10}}>⚙️ Supabase Setup Required</div>
    <div style={{fontSize:13,color:C.textMuted,lineHeight:1.8,marginBottom:12}}>
      To use this app with real storage and audio, complete these steps:
    </div>
    {[
      ["1","Go to supabase.com → New project (free)"],
      ["2","Open SQL Editor → paste & run the supabase_setup.sql file"],
      ["3","Go to Settings → API → copy Project URL and anon key"],
      ["4","Replace YOUR_SUPABASE_URL and YOUR_SUPABASE_ANON_KEY at the top of this file"],
    ].map(([n,t])=><div key={n} style={{display:"flex",gap:10,marginBottom:8,alignItems:"flex-start"}}>
      <div style={{width:22,height:22,borderRadius:"50%",background:C.goldDim,border:`1px solid ${C.gold}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:C.gold,flexShrink:0,fontWeight:700}}>{n}</div>
      <div style={{fontSize:12,color:C.textMuted,lineHeight:1.6}}>{t}</div>
    </div>)}
  </div>;
}

// ══════════════════════════════════════════════════════════════════════════
// APP
// ══════════════════════════════════════════════════════════════════════════
export default function App(){
  const [members,   setMembers]   = useState([]);
  const [attendance,setAttendance]= useState([]);
  const [songs,     setSongs]     = useState([]);
  const [songCats,  setSongCats]  = useState(DEFAULT_CATS);
  const [tab,       setTab]       = useState("scanner");
  const [toast,     setToast]     = useState(null);
  const [loading,   setLoading]   = useState(true);
  const [configured]= useState(sb.configured());

  const showToast=(msg,type="success")=>setToast({msg,type});

  // ── Load all data from Supabase on mount ───────────────────────────────
  useEffect(()=>{
    if(!configured){setLoading(false);return;}
    (async()=>{
      try {
        const [m,a,s,cats]=await Promise.all([
          sb.get("members","order=name"),
          sb.get("attendance","order=date.desc"),
          sb.get("songs","order=added_at.desc"),
          sb.get("song_cats","order=name"),
        ]);
        setMembers(m||[]);
        setAttendance(a||[]);
        setSongs(s||[]);
        if(cats?.length) setSongCats(cats.map(c=>c.name));
      } catch(e){ showToast("Failed to load data: "+e.message,"error"); }
      finally{ setLoading(false); }
    })();
  },[]);

  const markAttendance=useCallback(async(memberId)=>{
    const dup=attendance.some(r=>r.member_id===memberId&&r.date===today()&&r.event_id===EVENT_ID);
    if(dup){showToast("Already checked in today!","warn");return false;}
    const rec={member_id:memberId,date:today(),event_id:EVENT_ID,event_type:EVENT_TYPE,time:new Date().toLocaleTimeString("en-US",{hour:"2-digit",minute:"2-digit"})};
    try{
      const inserted=await sb.insert("attendance",rec);
      setAttendance(p=>[inserted[0]||{...rec,id:Date.now()},...p]);
      return true;
    }catch(e){showToast("Check-in failed: "+e.message,"error");return false;}
  },[attendance]);

  const tabs=[
    {id:"scanner",   label:"Scanner",   icon:"⬛"},
    {id:"attendance",label:"Attendance",icon:"📋"},
    {id:"members",   label:"Members",   icon:"👥"},
    {id:"muziq",     label:"መዝሙር",     icon:"🎵"},
    {id:"dashboard", label:"Dashboard", icon:"📊"},
  ];

  if(loading) return <div style={{minHeight:"100vh",background:C.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:16}}>
    <ChurchLogo size={56}/>
    <Spinner size={28}/>
    <div style={{fontSize:13,color:C.textDim,fontFamily:"'DM Sans',sans-serif"}}>Loading ፍኖት ትጉሃን…</div>
  </div>;

  return <div style={{minHeight:"100vh",background:C.bg,fontFamily:"'DM Sans',sans-serif",color:C.text}}>
    {/* Header */}
    <div style={{background:C.teal,borderBottom:`1px solid ${C.gold}44`,position:"sticky",top:0,zIndex:100,boxShadow:"0 4px 24px #00000044"}}>
      <div style={{maxWidth:860,margin:"0 auto",padding:"14px 20px 0"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <ChurchLogo size={46}/>
            <div>
              <div style={{fontSize:20,fontFamily:"'Noto Serif Ethiopic',serif",color:C.gold,fontWeight:700,lineHeight:1.1}}>ፍኖት ትጉሃን</div>
              <div style={{fontSize:10,color:C.textDim,letterSpacing:".12em",textTransform:"uppercase",marginTop:2}}>ሰንበት ትምህርት ቤት</div>
            </div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:12,color:C.gold}}>{fmtDate(today())}</div>
            <div style={{fontSize:11,color:C.textDim,display:"flex",alignItems:"center",gap:4,justifyContent:"flex-end"}}>
              <div style={{width:6,height:6,borderRadius:"50%",background:configured?C.success:C.danger}}/>
              {configured?"Connected":"Not configured"}
            </div>
          </div>
        </div>
        <div style={{display:"flex"}}>
          {tabs.map(t=><button key={t.id} onClick={()=>setTab(t.id)} style={{flex:1,padding:"9px 4px 10px",border:"none",background:"transparent",color:tab===t.id?C.gold:C.textDim,borderBottom:tab===t.id?`2px solid ${C.gold}`:"2px solid transparent",cursor:"pointer",fontSize:11,fontFamily:"'DM Sans',sans-serif",fontWeight:tab===t.id?600:400,transition:"all .2s",textTransform:"uppercase",letterSpacing:".05em"}}>
            <div style={{fontSize:15,marginBottom:2}}>{t.icon}</div>{t.label}
          </button>)}
        </div>
      </div>
    </div>

    {/* Content */}
    <div style={{maxWidth:860,margin:"0 auto",padding:"22px 16px 80px",animation:"fadeUp .35s ease"}}>
      {!configured && <SetupBanner/>}
      {tab==="scanner"   &&<ScannerTab    members={members} attendance={attendance} markAttendance={markAttendance} showToast={showToast} configured={configured}/>}
      {tab==="attendance"&&<AttendanceTab members={members} attendance={attendance} exportExcel={exportExcel}/>}
      {tab==="members"   &&<MembersTab    members={members} setMembers={setMembers} attendance={attendance} showToast={showToast} configured={configured}/>}
      {tab==="muziq"     &&<MuziqTab      songs={songs} setSongs={setSongs} songCats={songCats} setSongCats={setSongCats} showToast={showToast} configured={configured}/>}
      {tab==="dashboard" &&<DashboardTab  members={members} attendance={attendance} exportExcel={exportExcel}/>}
    </div>
    {toast&&<Toast msg={toast.msg} type={toast.type} onDone={()=>setToast(null)}/>}
  </div>;
}

// ══════════════════════════════════════════════════════════════════════════
// SCANNER TAB
// ══════════════════════════════════════════════════════════════════════════
function ScannerTab({members,attendance,markAttendance,showToast,configured}){
  const [mode,setMode]=useState("camera"); // "camera" | "manual"
  const [query,setQuery]=useState("");
  const [lastScanned,setLastScanned]=useState(null);
  const [camStatus,setCamStatus]=useState("starting"); // starting|active|denied|unsupported|error
  const [paused,setPaused]=useState(false);
  const inputRef=useRef();
  const videoRef=useRef(null);
  const canvasRef=useRef(null);
  const streamRef=useRef(null);
  const rafRef=useRef(null);
  const pausedRef=useRef(false);
  const useBarcodeDetectorRef=useRef(false);
  const detectorRef=useRef(null);

  const todayCount=attendance.filter(r=>r.date===today()&&r.event_id===EVENT_ID).length;

  const doCheckin=useCallback(async(member)=>{
    const ok=await markAttendance(member.id);
    if(ok){
      setLastScanned(member);
      showToast("🙏 "+member.name+" checked in","success");
      // Pause scanning briefly so the same badge isn't re-scanned instantly.
      pausedRef.current=true;setPaused(true);
      setTimeout(()=>{pausedRef.current=false;setPaused(false);},2500);
    } else {
      // markAttendance already toasts "already checked in" — still pause a beat.
      pausedRef.current=true;setPaused(true);
      setTimeout(()=>{pausedRef.current=false;setPaused(false);},1500);
    }
    setQuery("");
  },[markAttendance,showToast]);

  const handleDecoded=useCallback((text)=>{
    if(!text)return;
    let memberId=null;
    if(text.startsWith(QR_PREFIX)) memberId=text.slice(QR_PREFIX.length);
    else memberId=text.trim(); // tolerate old/plain-ID badges too
    const member=members.find(m=>m.id===memberId);
    if(member) doCheckin(member);
    else showToast("QR code not recognized — not a member badge","error");
  },[members,doCheckin,showToast]);

  // ── Camera lifecycle ────────────────────────────────────────────────
  useEffect(()=>{
    if(mode!=="camera")return;
    let cancelled=false;

    if(!navigator.mediaDevices?.getUserMedia){
      setCamStatus("unsupported");
      return;
    }
    // BarcodeDetector = fast native path (Chrome/Edge/Android). Falls back to jsQR otherwise (needed for iOS Safari).
    if("BarcodeDetector" in window){
      try{ detectorRef.current=new window.BarcodeDetector({formats:["qr_code"]}); useBarcodeDetectorRef.current=true; }
      catch{ useBarcodeDetectorRef.current=false; }
    }

    (async()=>{
      try{
        const stream=await navigator.mediaDevices.getUserMedia({
          video:{ facingMode:{ideal:"environment"}, width:{ideal:640}, height:{ideal:480} },
          audio:false
        });
        if(cancelled){ stream.getTracks().forEach(t=>t.stop()); return; }
        streamRef.current=stream;
        const v=videoRef.current;
        if(!v)return;
        v.srcObject=stream;
        v.setAttribute("playsinline","true"); // required for iOS Safari — without this it forces fullscreen and blocks the loop
        await v.play();
        setCamStatus("active");
        scanLoop();
      }catch(e){
        if(cancelled)return;
        setCamStatus(e?.name==="NotAllowedError"?"denied":"error");
      }
    })();

    function scanLoop(){
      rafRef.current=requestAnimationFrame(scanLoop);
      const v=videoRef.current,c=canvasRef.current;
      if(!v||!c||v.readyState!==v.HAVE_ENOUGH_DATA)return;
      if(pausedRef.current)return;

      const w=320,h=Math.round(w*(v.videoHeight/v.videoWidth||0.75));
      if(c.width!==w)c.width=w;
      if(c.height!==h)c.height=h;
      const ctx=c.getContext("2d",{willReadFrequently:true});
      ctx.drawImage(v,0,0,w,h);

      if(useBarcodeDetectorRef.current&&detectorRef.current){
        detectorRef.current.detect(c).then(codes=>{
          if(codes?.[0]?.rawValue)handleDecoded(codes[0].rawValue);
        }).catch(()=>{});
      } else {
        const imgData=ctx.getImageData(0,0,w,h);
        const result=jsQR(imgData.data,w,h,{inversionAttempts:"dontInvert"});
        if(result?.data)handleDecoded(result.data);
      }
    }

    return()=>{
      cancelled=true;
      if(rafRef.current)cancelAnimationFrame(rafRef.current);
      streamRef.current?.getTracks().forEach(t=>t.stop());
      streamRef.current=null;
    };
  },[mode,handleDecoded]);

  useEffect(()=>{ if(mode==="manual") inputRef.current?.focus(); },[mode]);

  const filtered=query.length>1?members.filter(m=>m.name.toLowerCase().includes(query.toLowerCase())||m.id.toLowerCase().includes(query.toLowerCase())):[];

  return <div>
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:18,flexWrap:"wrap",gap:10}}>
      <div style={{display:"inline-flex",alignItems:"center",gap:10,background:C.goldDim,border:`1.5px solid ${C.gold}66`,borderRadius:12,padding:"10px 20px"}}>
        <div style={{width:8,height:8,borderRadius:"50%",background:C.gold,boxShadow:`0 0 6px ${C.gold}`}}/>
        <span style={{fontFamily:"'Noto Serif Ethiopic',serif",fontSize:17,color:C.gold,fontWeight:700}}>{"መዝሙር ጥናት"}</span>
      </div>
      <button onClick={()=>setMode(m=>m==="camera"?"manual":"camera")} className="btn-outline" style={{padding:"8px 14px",fontSize:12}}>
        {mode==="camera"?"⌨️ Manual check-in":"📷 Camera scan"}
      </button>
    </div>

    {mode==="camera"?<>
      <div style={{background:`linear-gradient(160deg,${C.teal},${C.tealMid})`,border:`1px solid ${C.gold}44`,borderRadius:18,padding:24,marginBottom:20,textAlign:"center",position:"relative",overflow:"hidden"}}>
        <div style={{position:"relative",width:240,height:240,margin:"0 auto 18px",borderRadius:14,overflow:"hidden",border:`2px solid ${C.gold}55`,background:C.bg}}>
          <video ref={videoRef} muted playsInline autoPlay style={{width:"100%",height:"100%",objectFit:"cover",display:camStatus==="active"?"block":"none"}}/>
          <canvas ref={canvasRef} style={{display:"none"}}/>
          {camStatus!=="active"&&<div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:10,padding:16,textAlign:"center"}}>
            {camStatus==="starting"&&<><Spinner size={26}/><div style={{fontSize:11,color:C.textDim}}>Starting camera…</div></>}
            {camStatus==="denied"&&<><div style={{fontSize:26}}>🚫</div><div style={{fontSize:11,color:C.danger,lineHeight:1.5}}>Camera access denied.<br/>Allow camera permission for this site in your browser settings, then reload.</div></>}
            {camStatus==="unsupported"&&<><div style={{fontSize:26}}>⚠️</div><div style={{fontSize:11,color:C.warn,lineHeight:1.5}}>Camera scanning isn't supported in this browser. Use manual check-in instead.</div></>}
            {camStatus==="error"&&<><div style={{fontSize:26}}>⚠️</div><div style={{fontSize:11,color:C.danger,lineHeight:1.5}}>Couldn't access the camera. Make sure no other app is using it, then reload.</div></>}
          </div>}
          {camStatus==="active"&&<>
            <div style={{position:"absolute",left:0,right:0,height:2,background:`linear-gradient(90deg,transparent,${C.gold},transparent)`,animation:"scanLine 2.2s ease-in-out infinite",opacity:paused?0:.8}}/>
            {[[0,"top:8px;left:8px","2px 0 0 2px"],[1,"top:8px;right:8px","2px 2px 0 0"],[2,"bottom:8px;left:8px","0 0 2px 2px"],[3,"bottom:8px;right:8px","0 2px 2px 0"]].map(([i,pos,bw])=>(
              <div key={i} style={{position:"absolute",width:18,height:18,borderColor:C.gold,borderStyle:"solid",borderWidth:bw,...Object.fromEntries(pos.split(";").map(s=>s.split(":")))}}/>
            ))}
          </>}
        </div>
        <div style={{fontSize:44,fontFamily:"'Noto Serif Ethiopic',serif",color:C.gold,fontWeight:700,lineHeight:1}}>{todayCount}</div>
        <div style={{fontSize:12,color:C.textMuted,marginTop:4}}>checked in today</div>
        {camStatus==="active"&&<div style={{fontSize:11,color:C.textDim,marginTop:8}}>{paused?"✓ Scanned — hold for next person…":"Point the camera at a member's QR badge"}</div>}
      </div>
    </>:<>
      <div style={{position:"relative",marginBottom:8}}>
        <div style={{position:"absolute",left:13,top:"50%",transform:"translateY(-50%)",color:C.textDim,fontSize:15,pointerEvents:"none"}}>🔍</div>
        <input ref={inputRef} value={query} onChange={e=>setQuery(e.target.value)}
          onKeyDown={e=>{if(e.key==="Enter"&&filtered.length===1)doCheckin(filtered[0]);}}
          placeholder="Search member name or ID…"
          style={{width:"100%",padding:"13px 16px 13px 42px",borderRadius:10,border:`1px solid ${C.tealBorder}`,background:C.teal,color:C.text,fontSize:14,transition:"border-color .2s"}}
          onFocus={e=>e.target.style.borderColor=C.gold} onBlur={e=>e.target.style.borderColor=C.tealBorder}/>
      </div>
      {filtered.length>0&&<Card style={{marginBottom:16}}>
        {filtered.map(m=>{
          const checked=attendance.some(r=>r.member_id===m.id&&r.date===today()&&r.event_id===EVENT_ID);
          return <div key={m.id} className={checked?"":"row-hover"} onClick={()=>!checked&&doCheckin(m)}
            style={{display:"flex",alignItems:"center",gap:12,padding:"12px 16px",borderBottom:`1px solid ${C.tealBorder}22`,opacity:checked?.65:1,transition:"background .15s"}}>
            <Avatar name={m.name} size={38}/>
            <div style={{flex:1}}><div style={{fontSize:14,fontWeight:500}}>{m.name}</div><div style={{fontSize:12,color:C.textDim}}>{m.group} · {m.id}</div></div>
            {checked?<span style={{fontSize:12,color:C.success,background:C.successBg,padding:"4px 12px",borderRadius:20}}>✓ In</span>
              :<span style={{fontSize:12,color:C.gold,background:C.goldFaint,border:`1px solid ${C.gold}44`,padding:"4px 12px",borderRadius:20}}>Check in</span>}
          </div>;
        })}
      </Card>}
    </>}

    {lastScanned&&<div style={{background:C.successBg,border:`1px solid ${C.success}44`,borderRadius:12,padding:"16px 20px",display:"flex",alignItems:"center",gap:14}}>
      <div style={{fontSize:28}}>🙏</div>
      <div><div style={{fontSize:15,fontWeight:600,color:C.success}}>Welcome, {lastScanned.name.split(" ")[0]}!</div>
      <div style={{fontSize:12,color:"#4a9a6a"}}>{lastScanned.group} · {new Date().toLocaleTimeString()}</div></div>
    </div>}
  </div>;
}

// ══════════════════════════════════════════════════════════════════════════
// ATTENDANCE TAB
// ══════════════════════════════════════════════════════════════════════════
function AttendanceTab({members,attendance,exportExcel}){
  const [filterDate,setFilterDate]=useState(today());
  const filtered=attendance.filter(r=>r.date===filterDate);
  const presentIds=new Set(filtered.map(r=>r.member_id));
  const absent=members.filter(m=>!presentIds.has(m.id));
  const rate=members.length?Math.round(filtered.length/members.length*100):0;
  return <div>
    <div style={{display:"flex",gap:10,marginBottom:18,flexWrap:"wrap",alignItems:"center"}}>
      <input type="date" value={filterDate} onChange={e=>setFilterDate(e.target.value)}
        style={{flex:1,padding:"10px 14px",borderRadius:8,border:`1px solid ${C.tealBorder}`,background:C.bg,color:C.text,fontSize:13}}/>
      <button onClick={()=>exportExcel(members,attendance,null)} className="btn-outline" style={{padding:"10px 16px",fontSize:13,whiteSpace:"nowrap"}}>📥 Export Excel</button>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:20}}>
      <StatCard label="Present" val={filtered.length} color={C.success} icon="✅"/>
      <StatCard label="Absent"  val={absent.length}   color={C.danger}  icon="⭕"/>
      <StatCard label="Rate"    val={rate+"%"}         color={C.gold}    icon="📊"/>
    </div>
    <Section label={`Present (${filtered.length})`}>
      <Card>
        {filtered.length===0?<div style={{padding:24,textAlign:"center",color:C.textDim,fontSize:14}}>No records for this date</div>
          :filtered.map(r=>{const m=members.find(x=>x.id===r.member_id);if(!m)return null;
            return <div key={r.id||r.member_id} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 16px",borderBottom:`1px solid ${C.tealBorder}22`}}>
              <Avatar name={m.name} size={36}/>
              <div style={{flex:1}}><div style={{fontSize:14,fontWeight:500}}>{m.name}</div><div style={{fontSize:12,color:C.textDim}}>{m.group}</div></div>
              <div style={{textAlign:"right"}}><div style={{fontSize:12,color:C.success}}>✓ {r.time}</div></div>
            </div>;})}
      </Card>
    </Section>
    {absent.length>0&&<Section label={`Absent (${absent.length})`}>
      <Card>
        {absent.map(m=><div key={m.id} style={{display:"flex",alignItems:"center",gap:12,padding:"11px 16px",borderBottom:`1px solid ${C.tealBorder}22`,opacity:.6}}>
          <Avatar name={m.name} size={36} color={C.danger}/>
          <div style={{flex:1}}><div style={{fontSize:14}}>{m.name}</div><div style={{fontSize:12,color:C.textDim}}>{m.group}</div></div>
          <span style={{fontSize:11,color:C.danger}}>Absent</span>
        </div>)}
      </Card>
    </Section>}
  </div>;
}

// ══════════════════════════════════════════════════════════════════════════
// MEMBERS TAB
// ══════════════════════════════════════════════════════════════════════════
function MembersTab({members,setMembers,attendance,showToast,configured}){
  const [search,setSearch]=useState("");
  const [grp,setGrp]=useState("All");
  const [selected,setSelected]=useState(null);
  const [showAdd,setShowAdd]=useState(false);
  const [saving,setSaving]=useState(false);
  const [form,setForm]=useState({name:"",phone:"",group:"General",role:"Member"});
  const filtered=members.filter(m=>(grp==="All"||m.group===grp)&&(m.name.toLowerCase().includes(search.toLowerCase())||m.id.toLowerCase().includes(search.toLowerCase())));

  const addMember=async()=>{
    if(!form.name.trim()){showToast("Name is required","error");return;}
    setSaving(true);
    const newMember={id:uid(),name:form.name.trim(),phone:form.phone,group:form.group,role:form.role,joined:today()};
    try{
      await sb.insert("members",newMember);
      setMembers(p=>[...p,newMember]);
      showToast(form.name+" added!","success");
      setForm({name:"",phone:"",group:"General",role:"Member"});
      setShowAdd(false);
    }catch(e){showToast("Failed: "+e.message,"error");}
    setSaving(false);
  };

  const deleteMember=async(id)=>{
    try{
      await sb.delete("members",{id});
      setMembers(p=>p.filter(m=>m.id!==id));
      setSelected(null);
      showToast("Member removed","warn");
    }catch(e){showToast("Failed: "+e.message,"error");}
  };

  if(selected){
    const m=members.find(x=>x.id===selected);if(!m){setSelected(null);return null;}
    const total=attendance.filter(r=>r.member_id===m.id).length;
    const missed=missedWeeks(m.id,attendance);
    return <div>
      <button onClick={()=>setSelected(null)} style={{background:"none",border:"none",color:C.gold,cursor:"pointer",fontSize:13,marginBottom:16,fontFamily:"'DM Sans',sans-serif"}}>← Back</button>
      <div style={{background:`linear-gradient(135deg,${C.teal},${C.tealMid})`,border:`1px solid ${C.gold}33`,borderRadius:16,padding:24,marginBottom:16}}>
        <div style={{display:"flex",gap:16,alignItems:"flex-start",marginBottom:20}}>
          <Avatar name={m.name} size={60}/>
          <div style={{flex:1}}>
            <div style={{fontSize:20,fontFamily:"'Noto Serif Ethiopic',serif",color:C.text,fontWeight:600}}>{m.name}</div>
            <div style={{fontSize:13,color:C.gold,marginTop:3}}>{m.role} · {m.group}</div>
            <div style={{fontSize:12,color:C.textDim,marginTop:4}}>📞 {m.phone}</div>
            <div style={{fontSize:12,color:C.textDim}}>🗓 Joined {fmtDate(m.joined)}</div>
          </div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
          {[{l:"Sessions",v:total,c:C.gold},{l:"Weeks missed",v:missed,c:missed>=3?C.danger:C.success},{l:"ID",v:m.id,c:C.textMuted}].map(s=>(
            <div key={s.l} style={{textAlign:"center",background:C.bg,borderRadius:8,padding:"12px 6px",border:`1px solid ${s.c}22`}}>
              <div style={{fontSize:20,fontFamily:"'Noto Serif Ethiopic',serif",color:s.c}}>{s.v}</div>
              <div style={{fontSize:10,color:C.textDim,marginTop:2}}>{s.l}</div>
            </div>
          ))}
        </div>
      </div>
      <Card style={{padding:22,textAlign:"center",marginBottom:14}}>
        <div style={{fontSize:11,color:C.textDim,letterSpacing:".1em",textTransform:"uppercase",marginBottom:14}}>Member QR Code</div>
        <div style={{display:"inline-block",padding:14,background:"white",borderRadius:10,boxShadow:`0 0 0 4px ${C.gold}33`}}><QRCodeSVG value={m.id} size={140}/></div>
        <div style={{marginTop:10,fontSize:12,color:C.textDim}}>ID: {m.id}</div>
        <div style={{marginTop:4,fontSize:11,color:C.tealBorder}}>ፍኖት ትጉሃን ሰንበት ትምህርት ቤት</div>
      </Card>
      <button onClick={()=>deleteMember(m.id)} className="btn-danger" style={{width:"100%",padding:11,fontSize:13}}>Remove member</button>
    </div>;
  }

  return <div>
    <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap"}}>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search members…"
        style={{flex:1,minWidth:140,padding:"10px 14px",borderRadius:8,border:`1px solid ${C.tealBorder}`,background:C.teal,color:C.text,fontSize:13}}
        onFocus={e=>e.target.style.borderColor=C.gold} onBlur={e=>e.target.style.borderColor=C.tealBorder}/>
      <FieldSelect value={grp} onChange={e=>setGrp(e.target.value)}>
        {["All",...GROUPS].map(g=><option key={g}>{g}</option>)}
      </FieldSelect>
      <button onClick={()=>setShowAdd(v=>!v)} className={showAdd?"btn-gold":"btn-outline"} style={{padding:"10px 18px",fontSize:13,whiteSpace:"nowrap"}}>+ Add</button>
    </div>
    {showAdd&&<div style={{background:C.teal,border:`1px solid ${C.gold}33`,borderRadius:12,padding:20,marginBottom:16}}>
      <div style={{fontSize:14,fontWeight:600,color:C.gold,marginBottom:14}}>New Member</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
        <div><div style={{fontSize:11,color:C.textDim,marginBottom:4}}>Full Name</div><FieldInput value={form.name} onChange={e=>setForm(p=>({...p,name:e.target.value}))} placeholder="e.g. Abebe Girma"/></div>
        <div><div style={{fontSize:11,color:C.textDim,marginBottom:4}}>Phone</div><FieldInput value={form.phone} onChange={e=>setForm(p=>({...p,phone:e.target.value}))} placeholder="0911-XXXXXX"/></div>
        <div><div style={{fontSize:11,color:C.textDim,marginBottom:4}}>Group</div><FieldSelect value={form.group} onChange={e=>setForm(p=>({...p,group:e.target.value}))} full>{GROUPS.map(g=><option key={g}>{g}</option>)}</FieldSelect></div>
        <div><div style={{fontSize:11,color:C.textDim,marginBottom:4}}>Role</div><FieldSelect value={form.role} onChange={e=>setForm(p=>({...p,role:e.target.value}))} full>{["Member","Leader","Elder","Pastor","Deacon","Visitor"].map(r=><option key={r}>{r}</option>)}</FieldSelect></div>
      </div>
      <button onClick={addMember} className="btn-gold" style={{width:"100%",padding:11,fontSize:14,display:"flex",alignItems:"center",justifyContent:"center",gap:8}} disabled={saving}>
        {saving&&<Spinner size={14} color={C.bg}/>}{saving?"Saving…":"Add Member"}
      </button>
    </div>}
    <Card>
      {filtered.length===0?<div style={{padding:24,textAlign:"center",color:C.textDim,fontSize:14}}>No members found</div>
        :filtered.map(m=>{const missed=missedWeeks(m.id,attendance);
          return <div key={m.id} className="row-hover" onClick={()=>setSelected(m.id)}
            style={{display:"flex",alignItems:"center",gap:12,padding:"12px 16px",borderBottom:`1px solid ${C.tealBorder}22`,transition:"background .15s"}}>
            <Avatar name={m.name} size={38}/>
            <div style={{flex:1}}><div style={{fontSize:14,fontWeight:500}}>{m.name}</div><div style={{fontSize:12,color:C.textDim}}>{m.role} · {m.group}</div></div>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              {missed>=3&&<span style={{fontSize:10,color:C.danger,background:C.dangerBg,padding:"3px 8px",borderRadius:10,border:`1px solid ${C.danger}33`}}>⚠ {missed}w</span>}
              <span style={{color:C.textDim}}>›</span>
            </div>
          </div>;})}
    </Card>
    <div style={{fontSize:11,color:C.textDim,textAlign:"center",marginTop:10}}>{filtered.length} of {members.length} members</div>
  </div>;
}

// ══════════════════════════════════════════════════════════════════════════
// SPOTIFY SONG VIEW
// ══════════════════════════════════════════════════════════════════════════
function SpotifySongView({song,onBack,isAdmin,onEdit,onRequestDelete}){
  const audioRef=useRef(null);
  const [playing,setPlaying]=useState(false);
  const [current,setCurrent]=useState(0);
  const [duration,setDuration]=useState(0);
  const [audioError,setAudioError]=useState(null);
  const [loadingAudio,setLoadingAudio]=useState(false);
  const [ready,setReady]=useState(false);
  const hasAudio=!!song.audio_url;
  const fmt=s=>{if(!isFinite(s)||s<=0)return"0:00";const m=Math.floor(s/60);return`${m}:${Math.floor(s%60).toString().padStart(2,"0")}`;};
  const pct=duration>0?(current/duration)*100:0;

  useEffect(()=>{
    if(!hasAudio)return;
    const a=document.createElement("audio");
    audioRef.current=a;
    a.preload="auto";
    // NOTE: no crossOrigin="anonymous" here on purpose. It's only needed if you
    // pipe the audio through Web Audio API / canvas; for plain playback it forces
    // a strict CORS handshake on every Range request the browser makes while
    // seeking, and if Supabase's storage CORS config doesn't cover that exactly,
    // playback fails silently — which is what was happening.
    a.addEventListener("timeupdate",()=>setCurrent(a.currentTime));
    a.addEventListener("loadedmetadata",()=>{setDuration(a.duration||0);setReady(true);setLoadingAudio(false);});
    a.addEventListener("canplaythrough",()=>{setReady(true);setLoadingAudio(false);});
    a.addEventListener("ended",()=>setPlaying(false));
    a.addEventListener("waiting",()=>setLoadingAudio(true));
    a.addEventListener("playing",()=>setLoadingAudio(false));
    a.addEventListener("error",()=>{setLoadingAudio(false);setPlaying(false);setReady(false);setAudioError("Playback failed — make sure the '"+AUDIO_BUCKET+"' storage bucket in Supabase is set to Public.");});
    a.src=song.audio_url;
    a.load();
    return()=>{a.pause();a.src="";audioRef.current=null;};
  },[song.audio_url]);

  const togglePlay=()=>{
    const a=audioRef.current;if(!a)return;
    setAudioError(null);
    if(playing){a.pause();setPlaying(false);}
    else{setLoadingAudio(true);a.play().then(()=>{setPlaying(true);setLoadingAudio(false);}).catch(e=>{setLoadingAudio(false);setPlaying(false);setAudioError("Could not play: "+e.message);});}
  };
  const skip=sec=>{const a=audioRef.current;if(!a)return;a.currentTime=Math.max(0,Math.min(duration,a.currentTime+sec));};
  const seekClick=e=>{const a=audioRef.current;if(!a||!duration)return;const rect=e.currentTarget.getBoundingClientRect();a.currentTime=Math.max(0,Math.min(1,(e.clientX-rect.left)/rect.width))*duration;};
  const lines=song.lyrics.split("\n");

  return <div style={{position:"fixed",inset:0,background:C.bg,zIndex:150,display:"flex",flexDirection:"column",fontFamily:"'DM Sans',sans-serif"}}>
    <div style={{background:C.teal,borderBottom:`1px solid ${C.gold}33`,padding:"14px 20px",flexShrink:0}}>
      <div style={{maxWidth:860,margin:"0 auto",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <button onClick={onBack} style={{background:"none",border:"none",cursor:"pointer",color:C.textMuted,fontSize:26,lineHeight:1,padding:4}}>‹</button>
        <div style={{textAlign:"center",flex:1}}>
          <div style={{fontSize:15,fontFamily:"'Noto Serif Ethiopic',serif",color:C.text,fontWeight:700}}>{song.title}</div>
          <div style={{fontSize:12,color:C.textDim,marginTop:1}}>{song.category}</div>
        </div>
        {isAdmin?<div style={{display:"flex",gap:6}}>
          <button onClick={()=>onEdit(song)} style={{background:"none",border:"none",cursor:"pointer",color:C.gold,fontSize:14,padding:4}}>✏️</button>
          <button onClick={()=>onRequestDelete(song.id)} style={{background:"none",border:"none",cursor:"pointer",color:C.danger,fontSize:14,padding:4}}>🗑</button>
        </div>:<div style={{width:52}}/>}
      </div>
    </div>
    <div style={{flex:1,overflowY:"auto",padding:`24px 24px ${hasAudio?150:24}px`,maxWidth:860,width:"100%",margin:"0 auto",boxSizing:"border-box"}}>
      {lines.map((line,i)=>{
        const isHdr=line.startsWith("[")&&line.endsWith("]"),isEmpty=line.trim()==="";
        if(isEmpty)return <div key={i} style={{height:"1em"}}/>;
        return <div key={i} style={{fontFamily:"'Noto Serif Ethiopic',serif",fontSize:isHdr?11:18,fontWeight:isHdr?700:500,color:isHdr?C.gold:C.text,letterSpacing:isHdr?".12em":"normal",textTransform:isHdr?"uppercase":"none",marginBottom:isHdr?6:0,marginTop:isHdr?20:0,lineHeight:1.9}}>{line}</div>;
      })}
      <div style={{height:24}}/>
    </div>
    {hasAudio&&<div style={{position:"absolute",bottom:0,left:0,right:0,background:`linear-gradient(180deg,transparent 0%,${C.bg}cc 20%,${C.bg} 35%)`,backdropFilter:"blur(12px)",borderTop:`1px solid ${C.gold}22`,padding:"12px 24px 24px"}}>
      <div style={{maxWidth:860,margin:"0 auto"}}>
        {audioError&&<div style={{fontSize:11,color:C.danger,textAlign:"center",marginBottom:8,background:C.dangerBg,padding:"6px 12px",borderRadius:6}}>⚠ {audioError}</div>}
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16}}>
          <span style={{fontSize:11,color:C.textDim,minWidth:34,textAlign:"right"}}>{fmt(current)}</span>
          <div onClick={seekClick} style={{flex:1,height:4,background:`${C.tealBorder}66`,borderRadius:2,cursor:"pointer",position:"relative"}}>
            <div style={{position:"absolute",left:0,top:0,height:"100%",borderRadius:2,background:C.gold,width:`${pct}%`,transition:"width .1s linear"}}/>
            <div style={{position:"absolute",top:"50%",left:`${pct}%`,transform:"translate(-50%,-50%)",width:13,height:13,borderRadius:"50%",background:C.gold,boxShadow:`0 0 8px ${C.gold}88`}}/>
          </div>
          <span style={{fontSize:11,color:C.textDim,minWidth:34}}>{fmt(duration)}</span>
        </div>
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:32}}>
          <button onClick={()=>skip(-5)} style={{background:"none",border:"none",cursor:"pointer",color:ready?C.textMuted:C.tealBorder,display:"flex",alignItems:"center"}}>
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none"><path d="M12 5V1L7 6l5 5V7c3.31 0 6 2.69 6 6s-2.69 6-6 6-6-2.69-6-6H4c0 4.42 3.58 8 8 8s8-3.58 8-8-3.58-8-8-8z" fill="currentColor"/><text x="12" y="15.5" textAnchor="middle" fontSize="7" fill="currentColor" fontFamily="sans-serif" fontWeight="bold">5</text></svg>
          </button>
          <button onClick={togglePlay} style={{width:62,height:62,borderRadius:"50%",background:loadingAudio?`${C.gold}55`:"white",border:"none",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",boxShadow:"0 4px 20px #0008",transition:"all .15s"}}
            onMouseDown={e=>e.currentTarget.style.transform="scale(.92)"} onMouseUp={e=>e.currentTarget.style.transform="scale(1)"} onMouseLeave={e=>e.currentTarget.style.transform="scale(1)"}>
            {loadingAudio
              ?<svg width="26" height="26" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="9" stroke={C.teal} strokeWidth="3" strokeDasharray="28 14" strokeLinecap="round"><animateTransform attributeName="transform" type="rotate" from="0 12 12" to="360 12 12" dur="0.7s" repeatCount="indefinite"/></circle></svg>
              :playing
                ?<svg width="24" height="24" viewBox="0 0 24 24" fill={C.bg}><rect x="6" y="4" width="4" height="16" rx="1"/><rect x="14" y="4" width="4" height="16" rx="1"/></svg>
                :<svg width="24" height="24" viewBox="0 0 24 24" fill={C.bg}><path d="M8 5v14l11-7z"/></svg>}
          </button>
          <button onClick={()=>skip(5)} style={{background:"none",border:"none",cursor:"pointer",color:ready?C.textMuted:C.tealBorder,display:"flex",alignItems:"center"}}>
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none"><path d="M12 5V1l5 5-5 5V7c-3.31 0-6 2.69-6 6s2.69 6 6 6 6-2.69 6-6h2c0 4.42-3.58 8-8 8s-8-3.58-8-8 3.58-8 8-8z" fill="currentColor"/><text x="12" y="15.5" textAnchor="middle" fontSize="7" fill="currentColor" fontFamily="sans-serif" fontWeight="bold">5</text></svg>
          </button>
        </div>
      </div>
    </div>}
    {!hasAudio&&<div style={{padding:"14px 24px",borderTop:`1px solid ${C.tealBorder}33`,background:C.teal,textAlign:"center",flexShrink:0}}>
      <div style={{fontSize:12,color:C.textDim}}>No audio attached to this song</div>
    </div>}
  </div>;
}

// ══════════════════════════════════════════════════════════════════════════
// PIN MODAL
// ══════════════════════════════════════════════════════════════════════════
function PinModal({pinModal,pinInput,setPinInput,pinError,setPinError,submitPin,onClose}){
  return <div style={{position:"fixed",inset:0,background:"#000a",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
    <div style={{background:C.tealMid,border:`1px solid ${C.gold}44`,borderRadius:16,padding:28,width:"100%",maxWidth:320,textAlign:"center"}}>
      <div style={{fontSize:32,marginBottom:8}}>{pinModal==="confirm_delete"?"🗑":"🔒"}</div>
      <div style={{fontSize:16,fontFamily:"'Noto Serif Ethiopic',serif",color:C.text,fontWeight:600,marginBottom:4}}>{pinModal==="confirm_delete"?"Confirm Delete":"Admin Login"}</div>
      <div style={{fontSize:12,color:C.textDim,marginBottom:20}}>{pinModal==="confirm_delete"?"Enter PIN to confirm deletion":"Enter admin PIN to unlock editing"}</div>
      <input type="password" value={pinInput} onChange={e=>{setPinInput(e.target.value);setPinError(false);}} onKeyDown={e=>e.key==="Enter"&&submitPin()} placeholder="Enter PIN" autoFocus
        style={{width:"100%",padding:"12px 16px",borderRadius:8,border:`1px solid ${pinError?C.danger:C.tealBorder}`,background:C.bg,color:C.text,fontSize:16,textAlign:"center",letterSpacing:"0.3em",marginBottom:8}}/>
      {pinError&&<div style={{fontSize:12,color:C.danger,marginBottom:8}}>Incorrect PIN</div>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginTop:10}}>
        <button onClick={onClose} className="btn-outline" style={{padding:10,fontSize:13}}>Cancel</button>
        <button onClick={submitPin} className="btn-gold" style={{padding:10,fontSize:13}}>Confirm</button>
      </div>
      {pinModal==="unlock"&&<div style={{marginTop:12,fontSize:11,color:C.textDim}}>Default PIN: 1234</div>}
    </div>
  </div>;
}

// ══════════════════════════════════════════════════════════════════════════
// MUZIQ TAB
// ══════════════════════════════════════════════════════════════════════════
function MuziqTab({songs,setSongs,songCats,setSongCats,showToast,configured}){
  const [search,setSearch]=useState("");
  const [catFilter,setCatFilter]=useState("All");
  const [selected,setSelected]=useState(null);
  const [showForm,setShowForm]=useState(false);
  const [editId,setEditId]=useState(null);
  const [isAdmin,setIsAdmin]=useState(false);
  const [pinModal,setPinModal]=useState(false);
  const [pinInput,setPinInput]=useState("");
  const [pinError,setPinError]=useState(false);
  const [pendingDelete,setPendingDelete]=useState(null);
  const [showCatMgr,setShowCatMgr]=useState(false);
  const [newCatInput,setNewCatInput]=useState("");
  const [form,setForm]=useState({title:"",category:"",lyrics:"",audio_url:"",audio_name:""});
  const [audioFile,setAudioFile]=useState(null); // raw File object
  const [uploading,setUploading]=useState(false);
  const [saving,setSaving]=useState(false);
  const audioInputRef=useRef();

  const filtered=songs.filter(s=>(catFilter==="All"||s.category===catFilter)&&(s.title.toLowerCase().includes(search.toLowerCase())||s.lyrics.toLowerCase().includes(search.toLowerCase())));

  const openPinModal=purpose=>{setPinModal(purpose);setPinInput("");setPinError(false);};
  const submitPin=()=>{
    if(pinInput===ADMIN_PIN){
      if(pinModal==="unlock"){setIsAdmin(true);showToast("Admin unlocked 🔓","success");}
      if(pinModal==="confirm_delete"&&pendingDelete){doDelete(pendingDelete);}
      setPinModal(false);setPinInput("");
    }else setPinError(true);
  };

  const addCat=()=>{const v=newCatInput.trim();if(!v)return;if(songCats.includes(v)){showToast("Already exists","warn");return;}
    const updated=[...songCats,v];
    setSongCats(updated);
    if(configured)sb.insert("song_cats",{name:v}).catch(()=>{});
    setNewCatInput("");showToast(`"${v}" added`,"success");
  };
  const removeCat=cat=>{
    if(songs.some(s=>s.category===cat)){showToast(`"${cat}" is used by songs — reassign first`,"warn");return;}
    setSongCats(p=>p.filter(c=>c!==cat));
    if(configured)sb.delete("song_cats",{name:cat}).catch(()=>{});
    if(catFilter===cat)setCatFilter("All");
  };
  const renameCat=(old,newName)=>{
    const v=newName.trim();if(!v||v===old)return;if(songCats.includes(v)){showToast("Name in use","warn");return;}
    setSongCats(p=>p.map(c=>c===old?v:c));
    setSongs(p=>p.map(s=>s.category===old?{...s,category:v}:s));
    if(catFilter===old)setCatFilter(v);
    if(configured){
      sb.delete("song_cats",{name:old}).then(()=>sb.insert("song_cats",{name:v})).catch(()=>{});
      songs.filter(s=>s.category===old).forEach(s=>sb.update("songs",{id:s.id},{category:v}).catch(()=>{}));
    }
    showToast(`Renamed to "${v}"`,"success");
  };

  const openAdd=()=>{
    if(!isAdmin){openPinModal("unlock");return;}
    setEditId(null);setForm({title:"",category:songCats[0]||"",lyrics:"",audio_url:"",audio_name:""});
    setAudioFile(null);setShowForm(true);
  };
  const openEdit=song=>{
    setEditId(song.id);
    setForm({title:song.title,category:song.category,lyrics:song.lyrics,audio_url:song.audio_url||"",audio_name:song.audio_name||""});
    setAudioFile(null);setShowForm(true);setSelected(null);
  };

  const saveForm=async()=>{
    if(!form.title.trim()){showToast("Title is required","error");return;}
    if(!form.lyrics.trim()){showToast("Lyrics are required","error");return;}
    setSaving(true);
    let audio_url=form.audio_url, audio_name=form.audio_name;
    // Upload audio file to Supabase Storage if a new file was selected
    if(audioFile&&configured){
      setUploading(true);
      try{
        audio_url=await sb.uploadAudio(audioFile,audioFile.name);
        audio_name=audioFile.name;
        showToast("Audio uploaded ✓","success");
      }catch(e){
        showToast("Audio upload failed: "+e.message,"error");
        console.error("Upload error:",e);
        setSaving(false);setUploading(false);return;
      }
      setUploading(false);
    }
    const songData={title:form.title.trim(),category:form.category,lyrics:form.lyrics,audio_url,audio_name,added_by:"Admin",added_at:today()};
    try{
      if(editId){
        await sb.update("songs",{id:editId},songData);
        setSongs(p=>p.map(s=>s.id===editId?{...s,...songData}:s));
        showToast("Song updated ✓","success");
      }else{
        const newId=sid();
        const full={id:newId,...songData};
        await sb.insert("songs",full);
        setSongs(p=>[full,...p]);
        showToast("Song added 🎵","success");
      }
      setShowForm(false);setEditId(null);setAudioFile(null);
    }catch(e){showToast("Save failed: "+e.message,"error");}
    setSaving(false);
  };

  const requestDelete=id=>{setPendingDelete(id);openPinModal("confirm_delete");};
  const doDelete=async id=>{
    const song=songs.find(s=>s.id===id);
    try{
      if(song?.audio_url&&configured)await sb.deleteAudio(song.audio_url);
      await sb.delete("songs",{id});
      setSongs(p=>p.filter(s=>s.id!==id));
      setSelected(null);showToast("Song deleted","warn");
    }catch(e){showToast("Delete failed: "+e.message,"error");}
  };

  const handleAudioFile=e=>{
    const file=e.target.files?.[0];if(!file)return;
    if(file.size>50*1024*1024){showToast("Max file size is 50 MB","error");return;}
    setAudioFile(file);
    setForm(p=>({...p,audio_name:file.name}));
    showToast("File selected: "+file.name,"success");
  };

  if(showCatMgr)return <CategoryManager cats={songCats} onAddCat={addCat} onRemoveCat={removeCat} onRenameCat={renameCat} newCatInput={newCatInput} setNewCatInput={setNewCatInput} onBack={()=>setShowCatMgr(false)}/>;

  if(selected){
    const song=songs.find(s=>s.id===selected);if(!song){setSelected(null);return null;}
    return <>
      <SpotifySongView song={song} onBack={()=>setSelected(null)} isAdmin={isAdmin} onEdit={openEdit} onRequestDelete={requestDelete}/>
      {pinModal&&<PinModal pinModal={pinModal} pinInput={pinInput} setPinInput={setPinInput} pinError={pinError} setPinError={setPinError} submitPin={submitPin} onClose={()=>setPinModal(false)}/>}
    </>;
  }

  if(showForm)return <div>
    <button onClick={()=>{setShowForm(false);setEditId(null);setAudioFile(null);}} style={{background:"none",border:"none",color:C.gold,cursor:"pointer",fontSize:13,marginBottom:18,fontFamily:"'DM Sans',sans-serif"}}>← Cancel</button>
    <div style={{fontSize:16,fontFamily:"'Noto Serif Ethiopic',serif",color:C.gold,fontWeight:700,marginBottom:18}}>{editId?"✏️ Edit Song":"🎵 Add New Song"}</div>
    <Card style={{padding:20}}>
      <div style={{marginBottom:14}}>
        <div style={{fontSize:11,color:C.textDim,marginBottom:5}}>ርዕስ (Title)</div>
        <input value={form.title} onChange={e=>setForm(p=>({...p,title:e.target.value}))} placeholder="e.g. ትልቅ አምላክ ነህ"
          style={{width:"100%",padding:"11px 14px",borderRadius:8,border:`1px solid ${C.tealBorder}`,background:C.bg,color:C.text,fontSize:14,fontFamily:"'Noto Serif Ethiopic',serif"}}
          onFocus={e=>e.target.style.borderColor=C.gold} onBlur={e=>e.target.style.borderColor=C.tealBorder}/>
      </div>
      <div style={{marginBottom:14}}>
        <div style={{fontSize:11,color:C.textDim,marginBottom:5}}>ምድብ (Category)</div>
        <FieldSelect value={form.category} onChange={e=>setForm(p=>({...p,category:e.target.value}))} full>
          {songCats.map(c=><option key={c}>{c}</option>)}
        </FieldSelect>
      </div>
      {/* Audio upload */}
      <div style={{marginBottom:14}}>
        <div style={{fontSize:11,color:C.textDim,marginBottom:5}}>ኦዲዮ (Audio — any format, up to 50 MB)</div>
        <input ref={audioInputRef} type="file" accept="audio/*" onChange={handleAudioFile} style={{display:"none"}}/>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <button type="button" onClick={()=>audioInputRef.current?.click()} className="btn-outline" style={{padding:"9px 16px",fontSize:12,flexShrink:0}}>📂 Choose Audio</button>
          {audioFile
            ?<span style={{fontSize:12,color:C.success,flex:1}}>✓ {audioFile.name} ({(audioFile.size/1024/1024).toFixed(1)} MB)</span>
            :form.audio_name
              ?<span style={{fontSize:12,color:C.textMuted,flex:1}}>Current: {form.audio_name}</span>
              :<span style={{fontSize:12,color:C.textDim}}>No file selected</span>}
          {(audioFile||form.audio_url)&&<button type="button" onClick={()=>{setAudioFile(null);setForm(p=>({...p,audio_url:"",audio_name:""}));}} style={{background:"none",border:"none",color:C.danger,cursor:"pointer",fontSize:18,flexShrink:0}}>✕</button>}
        </div>
        {!configured&&<div style={{marginTop:6,fontSize:11,color:C.warn}}>⚠ Connect Supabase first to enable audio upload</div>}
      </div>
      <div style={{marginBottom:18}}>
        <div style={{fontSize:11,color:C.textDim,marginBottom:5}}>ግጥም (Lyrics) <span style={{marginLeft:6,color:C.tealBorder}}>— [ቁ. 1] for section headers</span></div>
        <textarea value={form.lyrics} onChange={e=>setForm(p=>({...p,lyrics:e.target.value}))}
          placeholder={"[chorus]\nቅዱስ ቅዱስ ቅዱስ\n...\n\n[ቁ. 1]\nፊትህን ሸፍነው\n..."}
          rows={13}
          style={{width:"100%",padding:"12px 14px",borderRadius:8,border:`1px solid ${C.tealBorder}`,background:C.bg,color:C.text,fontSize:14,fontFamily:"'Noto Serif Ethiopic',serif",lineHeight:1.9,resize:"vertical"}}
          onFocus={e=>e.target.style.borderColor=C.gold} onBlur={e=>e.target.style.borderColor=C.tealBorder}/>
      </div>
      <button onClick={saveForm} className="btn-gold" style={{width:"100%",padding:12,fontSize:14,display:"flex",alignItems:"center",justifyContent:"center",gap:8}} disabled={saving||uploading}>
        {(saving||uploading)&&<Spinner size={14} color={C.bg}/>}
        {uploading?"Uploading audio…":saving?"Saving…":editId?"Save Changes":"Add Song 🎵"}
      </button>
    </Card>
  </div>;

  return <div>
    <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
      <div style={{position:"relative",flex:1,minWidth:150}}>
        <div style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)",color:C.textDim,pointerEvents:"none",fontSize:14}}>🔍</div>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search songs or lyrics…"
          style={{width:"100%",padding:"10px 14px 10px 34px",borderRadius:8,border:`1px solid ${C.tealBorder}`,background:C.teal,color:C.text,fontSize:13}}
          onFocus={e=>e.target.style.borderColor=C.gold} onBlur={e=>e.target.style.borderColor=C.tealBorder}/>
      </div>
      <button onClick={openAdd} className="btn-gold" style={{padding:"10px 16px",fontSize:13,whiteSpace:"nowrap"}}>+ Add Song</button>
      {!isAdmin
        ?<button onClick={()=>openPinModal("unlock")} className="btn-outline" style={{padding:"10px 13px",fontSize:12}}>🔒 Admin</button>
        :<button onClick={()=>{setIsAdmin(false);showToast("Logged out","warn");}} className="btn-outline" style={{padding:"10px 13px",fontSize:12,borderColor:C.success,color:C.success}}>🔓 Admin</button>}
      {isAdmin&&<button onClick={()=>setShowCatMgr(true)} className="btn-outline" style={{padding:"10px 13px",fontSize:12,borderColor:C.textMuted,color:C.textMuted}}>⚙️ Categories</button>}
    </div>
    <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:16}}>
      {["All",...songCats].map(c=><button key={c} onClick={()=>setCatFilter(c)} style={{padding:"5px 13px",borderRadius:20,fontSize:12,cursor:"pointer",fontFamily:"'Noto Serif Ethiopic',serif",background:catFilter===c?C.goldDim:"transparent",border:`1px solid ${catFilter===c?C.gold:C.tealBorder}`,color:catFilter===c?C.gold:C.textDim,transition:"all .15s"}}>{c}</button>)}
    </div>
    <div style={{fontSize:11,color:C.textDim,marginBottom:10}}>{filtered.length} songs</div>
    {filtered.length===0
      ?<div style={{textAlign:"center",padding:"40px 20px",color:C.textDim}}><div style={{fontSize:40,marginBottom:12}}>🎵</div><div style={{fontSize:14}}>No songs found</div></div>
      :<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(250px,1fr))",gap:12}}>
        {filtered.map(song=>{
          const preview=song.lyrics.split("\n").find(l=>l.trim()&&!l.startsWith("["))||"";
          return <div key={song.id} className="row-hover" onClick={()=>setSelected(song.id)}
            style={{background:C.teal,border:`1px solid ${C.tealBorder}55`,borderRadius:14,padding:18,cursor:"pointer",transition:"all .18s",position:"relative",overflow:"hidden"}}>
            <div style={{position:"absolute",top:-14,right:-8,fontSize:52,opacity:.06,pointerEvents:"none"}}>🎵</div>
            <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:8}}>
              <div style={{fontSize:10,color:C.gold,letterSpacing:".1em",textTransform:"uppercase",marginBottom:6}}>{song.category}</div>
              {song.audio_url&&<span style={{fontSize:14}}>🎵</span>}
            </div>
            <div style={{fontSize:15,fontFamily:"'Noto Serif Ethiopic',serif",fontWeight:700,color:C.text,marginBottom:8,lineHeight:1.4}}>{song.title}</div>
            <div style={{fontSize:12,color:C.textDim,lineHeight:1.6,overflow:"hidden",display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical"}}>{preview}</div>
            <div style={{marginTop:10,fontSize:11,color:C.tealBorder}}>{fmtDate(song.added_at)}</div>
          </div>;
        })}
      </div>}
    {pinModal&&<PinModal pinModal={pinModal} pinInput={pinInput} setPinInput={setPinInput} pinError={pinError} setPinError={setPinError} submitPin={submitPin} onClose={()=>setPinModal(false)}/>}
  </div>;
}

// ══════════════════════════════════════════════════════════════════════════
// CATEGORY MANAGER
// ══════════════════════════════════════════════════════════════════════════
function CategoryManager({cats,onAddCat,onRemoveCat,onRenameCat,newCatInput,setNewCatInput,onBack}){
  const [renaming,setRenaming]=useState(null);
  const [renameVal,setRenameVal]=useState("");
  return <div>
    <button onClick={onBack} style={{background:"none",border:"none",color:C.gold,cursor:"pointer",fontSize:13,marginBottom:18,fontFamily:"'DM Sans',sans-serif"}}>← Back</button>
    <div style={{fontSize:16,fontFamily:"'Noto Serif Ethiopic',serif",color:C.gold,fontWeight:700,marginBottom:18}}>⚙️ Manage Categories</div>
    <Card style={{padding:18,marginBottom:20}}>
      <div style={{fontSize:11,color:C.textDim,marginBottom:10,textTransform:"uppercase",letterSpacing:".08em"}}>Add New Category</div>
      <div style={{display:"flex",gap:8}}>
        <input value={newCatInput} onChange={e=>setNewCatInput(e.target.value)} placeholder="Category name in Amharic…" onKeyDown={e=>e.key==="Enter"&&onAddCat()}
          style={{flex:1,padding:"10px 14px",borderRadius:8,border:`1px solid ${C.tealBorder}`,background:C.bg,color:C.text,fontSize:14,fontFamily:"'Noto Serif Ethiopic',serif"}}
          onFocus={e=>e.target.style.borderColor=C.gold} onBlur={e=>e.target.style.borderColor=C.tealBorder}/>
        <button onClick={onAddCat} className="btn-gold" style={{padding:"10px 18px",fontSize:13,whiteSpace:"nowrap"}}>+ Add</button>
      </div>
    </Card>
    <Card>
      {cats.length===0?<div style={{padding:24,textAlign:"center",color:C.textDim}}>No categories yet</div>
        :cats.map(cat=><div key={cat} style={{display:"flex",alignItems:"center",gap:10,padding:"13px 16px",borderBottom:`1px solid ${C.tealBorder}22`}}>
          {renaming===cat
            ?<><input value={renameVal} onChange={e=>setRenameVal(e.target.value)} autoFocus onKeyDown={e=>{if(e.key==="Enter"){onRenameCat(cat,renameVal);setRenaming(null);}if(e.key==="Escape")setRenaming(null);}}
                style={{flex:1,padding:"7px 12px",borderRadius:7,border:`1px solid ${C.gold}`,background:C.bg,color:C.text,fontSize:14,fontFamily:"'Noto Serif Ethiopic',serif"}}/>
              <button onClick={()=>{onRenameCat(cat,renameVal);setRenaming(null);}} className="btn-gold" style={{padding:"7px 14px",fontSize:12}}>Save</button>
              <button onClick={()=>setRenaming(null)} className="btn-outline" style={{padding:"7px 12px",fontSize:12}}>✕</button></>
            :<><span style={{flex:1,fontSize:15,fontFamily:"'Noto Serif Ethiopic',serif",color:C.text}}>{cat}</span>
              <button onClick={()=>{setRenaming(cat);setRenameVal(cat);}} className="btn-outline" style={{padding:"6px 12px",fontSize:12}}>✏️ Rename</button>
              <button onClick={()=>onRemoveCat(cat)} className="btn-danger" style={{padding:"6px 12px",fontSize:12}}>🗑</button></>}
        </div>)}
    </Card>
  </div>;
}

// ══════════════════════════════════════════════════════════════════════════
// DASHBOARD TAB
// ══════════════════════════════════════════════════════════════════════════
function DashboardTab({members,attendance,exportExcel}){
  const [personFilter,setPersonFilter]=useState("all");
  const weekly=[];
  for(let i=7;i>=0;i--){
    const d=new Date();d.setDate(d.getDate()-i*7);
    const wk=d.toISOString().split("T")[0].slice(0,7);
    const recs=personFilter==="all"?attendance.filter(r=>r.date.startsWith(wk)):attendance.filter(r=>r.date.startsWith(wk)&&r.member_id===personFilter);
    weekly.push({label:d.toLocaleDateString("en-US",{month:"short",day:"numeric"}),count:recs.length});
  }
  const maxW=Math.max(...weekly.map(w=>w.count),1);
  const sel=personFilter!=="all"?members.find(m=>m.id===personFilter):null;
  const memberHistory=sel?attendance.filter(r=>r.member_id===sel.id).sort((a,b)=>b.date.localeCompare(a.date)):[];
  const grpData=GROUPS.map(g=>({group:g,count:members.filter(m=>m.group===g).length})).filter(g=>g.count>0).sort((a,b)=>b.count-a.count);
  const alerts=members.filter(m=>missedWeeks(m.id,attendance)>=3);
  const thisMonth=attendance.filter(r=>r.date.startsWith(new Date().toISOString().slice(0,7))).length;
  const avg=Math.round(weekly.reduce((s,w)=>s+w.count,0)/8);

  return <div>
    <div style={{display:"flex",gap:10,marginBottom:20,flexWrap:"wrap",alignItems:"center"}}>
      <div style={{flex:1,minWidth:180}}>
        <FieldSelect value={personFilter} onChange={e=>setPersonFilter(e.target.value)} full>
          <option value="all">👥 All Members</option>
          {members.map(m=><option key={m.id} value={m.id}>{m.name}</option>)}
        </FieldSelect>
      </div>
      <button onClick={()=>exportExcel(members,attendance,personFilter==="all"?null:personFilter)} className="btn-outline" style={{padding:"10px 16px",fontSize:13,whiteSpace:"nowrap"}}>
        📥 {personFilter==="all"?"Export All":"Export Person"}
      </button>
    </div>

    {sel?<>
      <div style={{background:`linear-gradient(135deg,${C.teal},${C.tealMid})`,border:`1px solid ${C.gold}44`,borderRadius:16,padding:20,marginBottom:20,display:"flex",gap:16,alignItems:"center"}}>
        <Avatar name={sel.name} size={56}/>
        <div style={{flex:1}}>
          <div style={{fontSize:18,fontFamily:"'Noto Serif Ethiopic',serif",color:C.text,fontWeight:600}}>{sel.name}</div>
          <div style={{fontSize:13,color:C.gold,marginTop:2}}>{sel.role} · {sel.group}</div>
          <div style={{fontSize:12,color:C.textDim,marginTop:3}}>📞 {sel.phone}</div>
        </div>
        {missedWeeks(sel.id,attendance)>=3&&<span style={{fontSize:12,color:C.danger,background:C.dangerBg,padding:"6px 12px",borderRadius:20,border:`1px solid ${C.danger}33`}}>⚠ {missedWeeks(sel.id,attendance)}w absent</span>}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:20}}>
        <StatCard label="Total Sessions" val={memberHistory.length} color={C.gold} icon="📅"/>
        <StatCard label="Weeks Missed" val={missedWeeks(sel.id,attendance)} color={missedWeeks(sel.id,attendance)>=3?C.danger:C.success} icon="⭕"/>
        <StatCard label="Attendance Rate" val={Math.round((memberHistory.length/Math.max(new Set(attendance.map(r=>r.date.slice(0,7))).size,1))*100)+"%"} color="#7ab3e0" icon="📈"/>
      </div>
      <Section label="Weekly activity (8 weeks)">
        <Card style={{padding:20}}>
          <div style={{display:"flex",alignItems:"flex-end",gap:6,height:90}}>
            {weekly.map((w,i)=><div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
              <div style={{width:"100%",borderRadius:"4px 4px 0 0",height:w.count>0?"80px":"6px",background:w.count>0?(i===weekly.length-1?C.gold:`${C.gold}77`):`${C.danger}44`,transition:"height .4s ease"}}/>
              <div style={{fontSize:8,color:C.textDim,textAlign:"center",lineHeight:1.3}}>{w.label}</div>
            </div>)}
          </div>
          <div style={{display:"flex",gap:12,marginTop:12,fontSize:11,color:C.textDim}}>
            <span><span style={{color:C.gold}}>■</span> Present</span>
            <span><span style={{color:`${C.danger}88`}}>■</span> Absent</span>
          </div>
        </Card>
      </Section>
      <Section label={`Session history (${memberHistory.length})`} right={<span style={{fontSize:11,color:C.textDim}}>most recent first</span>}>
        <Card>
          {memberHistory.length===0?<div style={{padding:24,textAlign:"center",color:C.textDim,fontSize:14}}>No sessions recorded</div>
            :memberHistory.slice(0,20).map((r,i)=><div key={i} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 16px",borderBottom:`1px solid ${C.tealBorder}22`}}>
              <div style={{fontSize:13}}>{fmtDate(r.date)}</div>
              <span style={{fontSize:12,color:C.success,background:C.successBg,padding:"3px 10px",borderRadius:20}}>✓ {r.time}</span>
            </div>)}
          {memberHistory.length>20&&<div style={{padding:"10px 16px",fontSize:12,color:C.textDim,textAlign:"center"}}>+{memberHistory.length-20} more — export Excel for full history</div>}
        </Card>
      </Section>
    </>:<>
      <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:12,marginBottom:22}}>
        <StatCard label="Total members"  val={members.length} icon="👥" color={C.gold}/>
        <StatCard label="This month"     val={thisMonth}      icon="📅" color={C.success}/>
        <StatCard label="Weekly average" val={avg}            icon="📈" color="#7ab3e0"/>
        <StatCard label="Need follow-up" val={alerts.length}  icon="⚠️" color={C.danger}/>
      </div>
      <Section label="Weekly attendance trend (8 weeks)">
        <Card style={{padding:20}}>
          <div style={{display:"flex",alignItems:"flex-end",gap:6,height:110}}>
            {weekly.map((w,i)=><div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
              <div style={{fontSize:10,color:C.gold,fontWeight:500}}>{w.count||""}</div>
              <div style={{width:"100%",borderRadius:"4px 4px 0 0",height:maxW>0?`${Math.max(4,(w.count/maxW)*80)}px`:"4px",background:i===weekly.length-1?C.gold:`${C.gold}55`,transition:"height .4s ease"}}/>
              <div style={{fontSize:9,color:C.textDim,textAlign:"center",lineHeight:1.3}}>{w.label}</div>
            </div>)}
          </div>
        </Card>
      </Section>
      <Section label="Groups">
        <Card style={{padding:20}}>
          {grpData.map(g=><div key={g.group} style={{marginBottom:13}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
              <span style={{fontSize:13}}>{g.group}</span><span style={{fontSize:12,color:C.textDim}}>{g.count}</span>
            </div>
            <div style={{height:6,background:C.bg,borderRadius:4,overflow:"hidden"}}>
              <div style={{height:"100%",borderRadius:4,background:`linear-gradient(90deg,${C.gold}88,${C.gold})`,width:`${members.length?(g.count/members.length)*100:0}%`,transition:"width .4s ease"}}/>
            </div>
          </div>)}
        </Card>
      </Section>
      {alerts.length>0&&<Section label={`Follow-up needed (${alerts.length})`}>
        <Card>
          {alerts.map(m=>{const missed=missedWeeks(m.id,attendance);return <div key={m.id} style={{display:"flex",alignItems:"center",gap:12,padding:"12px 16px",borderBottom:`1px solid ${C.tealBorder}22`}}>
            <Avatar name={m.name} size={38} color={C.danger}/>
            <div style={{flex:1}}><div style={{fontSize:14}}>{m.name}</div><div style={{fontSize:12,color:C.textDim}}>{m.group} · 📞 {m.phone}</div></div>
            <span style={{fontSize:12,color:C.danger,background:C.dangerBg,padding:"4px 10px",borderRadius:20,border:`1px solid ${C.danger}33`}}>{missed}w absent</span>
          </div>;})}
        </Card>
      </Section>}
    </>}
  </div>;
}
