// @ts-nocheck
import { useState, useEffect, useCallback, Fragment } from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────────────────────────────────────
//  CONFIG
// ─────────────────────────────────────────────────────────────────────────────

const APP_PASSWORD = "USLsoccer1";
const SCRIPT_URL   = "https://script.google.com/macros/s/AKfycbxL5UHUZTpqiMoDRe05rf32KeLN-PvuRwfoBFFJWCSa3tsXYVsLJAKLd7wC8V5pwbcBOg/exec";

async function dbFetch() {
  const res = await fetch(`${SCRIPT_URL}?action=fetch`, { method:"GET", redirect:"follow" });
  return JSON.parse(await res.text());
}
async function dbFetchDeliverables() {
  const res = await fetch(`${SCRIPT_URL}?action=deliverables`, { method:"GET", redirect:"follow" });
  return JSON.parse(await res.text());
}
async function dbFetchClubs() {
  const res = await fetch(`${SCRIPT_URL}?action=clubs`, { method:"GET", redirect:"follow" });
  return JSON.parse(await res.text());
}

// Single entry insert
async function dbInsert(entry) {
  await fetch(SCRIPT_URL, { method:"POST", mode:"no-cors", body:JSON.stringify({ action:"insert", row:entry }) });
}

// Batch insert — sends all rows in one request
async function dbInsertBatch(entries) {
  await fetch(SCRIPT_URL, { method:"POST", mode:"no-cors", body:JSON.stringify({ action:"insertBatch", rows:entries }) });
}

async function dbDelete(ids) {
  await fetch(SCRIPT_URL, { method:"POST", mode:"no-cors", body:JSON.stringify({ action:"delete", ids }) });
}

// ─────────────────────────────────────────────────────────────────────────────
//  STAFF
// ─────────────────────────────────────────────────────────────────────────────

const STAFF = [
  "Amita Singh","Frances Baldridge","Garrett Mitchell","Julian Crockett",
  "Kendra Hodgdon","Kevin Couture","Ryan Halter","Steven Bell","Tim McCarthy",
];

// ─────────────────────────────────────────────────────────────────────────────
//  FALLBACK CLUBS (used until sheet loads)
// ─────────────────────────────────────────────────────────────────────────────

const FALLBACK_CLUBS = {
  "Championship": ["Birmingham Legion","Charleston Battery","Colorado Springs Switchbacks","Detroit City FC","El Paso Locomotive","FC Tulsa","Hartford Athletic","Indy Eleven","Las Vegas Lights","Lexington SC","Loudoun United","Louisville City FC","Miami FC","Monterey Bay FC","New Mexico United","North Carolina FC","Oakland Roots","Orange County SC","Phoenix Rising FC","Pittsburgh Riverhounds","Rhode Island FC","Sacramento Republic FC","San Antonio FC","Tampa Bay Rowdies"],
  "League One":   ["AV Alta","Chattanooga Red Wolves","Charlotte Independence","Corpus Christi","FC Naples","Forward Madison","Ft. Wayne Football Club","Greenville Triumph","One Knox SC","Portland Hearts of Pine","Richmond Kickers","South Georgia Tormenta","Spokane Velocity","Texoma FC","Union Omaha","Westchester SC"],
  "Super League": ["Brooklyn FC","Carolina Ascent","Dallas Trinity","DC Power","Ft. Lauderdale United","Lexington SC (Super League)","Sporting Jax","Tampa Bay Sun FC"],
  "Expansion":    ["Atletico Dallas","Ft. Lauderdale United (League One)","Oklahoma City Energy","Reno Pro Soccer","Santa Barbara Sky","Sarasota Paradise","Sporting Cascades","Sporting Jax (Championship)","USL Buffalo","USL Pro Iowa"],
};

// ─────────────────────────────────────────────────────────────────────────────
//  INTERNAL RECIPIENTS
// ─────────────────────────────────────────────────────────────────────────────

const INTERNAL_RECIPIENTS = {
  "Corp Partnerships":   ["HQ Corp Partnerships","League Operations"],
  "Marketing":           ["HQ Marketing / Comms","Expansion","Onboarding"],
  "Consumer Products":   ["Miscellaneous","Onboarding"],
  "Ticketing":           ["Miscellaneous","League Operations","Onboarding"],
  "Youth / Facilities":  ["League Operations","League Initiatives"],
  "League Initiatives":  ["League Operations"],
};

// ─────────────────────────────────────────────────────────────────────────────
//  DEPT CONFIG
// ─────────────────────────────────────────────────────────────────────────────

const DEPT_CONFIG = {
  "Corp Partnerships":  { color:"#7C3AED", light:"#F5F3FF", border:"#DDD6FE", emoji:"🤝", hasToggle:true  },
  "Marketing":          { color:"#0369A1", light:"#F0F9FF", border:"#BAE6FD", emoji:"📣", hasToggle:true  },
  "Consumer Products":  { color:"#B45309", light:"#FFFBEB", border:"#FDE68A", emoji:"🛍️", hasToggle:true  },
  "Ticketing":          { color:"#047857", light:"#F0FDF4", border:"#A7F3D0", emoji:"🎟️", hasToggle:true  },
  "Youth / Facilities": { color:"#0891B2", light:"#ECFEFF", border:"#A5F3FC", emoji:"⚽", hasToggle:true  },
  "League Initiatives": { color:"#4338CA", light:"#EEF2FF", border:"#C7D2FE", emoji:"🏛️", hasToggle:false },
};

const ALL_DEPT_NAMES = Object.keys(DEPT_CONFIG);
const TABS = ["Dashboard", ...ALL_DEPT_NAMES, "Activity Explorer"];

// ─────────────────────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────────────────────

const fmt$ = n => "$" + Number(n).toLocaleString();

function normalizeDept(dept) {
  const MAP = {
    "Corporate Partnerships":"Corp Partnerships","Corp. Partnerships":"Corp Partnerships",
    "Youth_Facilities":"Youth / Facilities","Youth/Facilities":"Youth / Facilities",
    "Institutional Strategy":"League Initiatives",
  };
  return MAP[String(dept).trim()] || String(dept).trim();
}

function buildClubsByLeague(clubsData) {
  const result = {};
  clubsData.filter(c => String(c.status||"").toLowerCase() !== "inactive").forEach(c => {
    if (!result[c.league]) result[c.league] = [];
    result[c.league].push(c.club);
  });
  Object.keys(result).forEach(l => result[l].sort());
  return result;
}


// ─────────────────────────────────────────────────────────────────────────────
//  FALLBACK DELIVERABLES (used when sheet fetch fails or returns empty)
// ─────────────────────────────────────────────────────────────────────────────

const FALLBACK_DEPT_ITEMS = {
  "Corp Partnerships": {
    external:[
      {name:"Stadium Naming Rights Valuation",rate:20000,cat:"Valuation",index_score:5,recurring:false,examples:["Naming rights valuation for new stadium","Renewal pricing analysis"]},
      {name:"Rate Card and Inventory Analysis",rate:4250,cat:"Valuation",index_score:4,recurring:false,examples:["Full sponsorship inventory audit","Rate card benchmarking"]},
      {name:"Kit Asset Valuation",rate:5000,cat:"Valuation",index_score:5,recurring:false,examples:["Jersey sleeve valuation","Kit partner ROI analysis"]},
      {name:"LED & Field Board Pricing",rate:3000,cat:"Valuation",index_score:3,recurring:false,examples:["LED board inventory pricing","Digital signage valuation"]},
      {name:"Standard Intelligence Product",rate:250,cat:"Analysis",index_score:2,recurring:true,examples:["Cluster report","Category analysis","Benchmark report"]},
      {name:"Partnership Performance Product",rate:250,cat:"Analysis",index_score:2,recurring:true,examples:["KORE recap","Annual partnership recap"]},
      {name:"Category Analysis",rate:75,cat:"Analysis",index_score:1,recurring:false,examples:["Insurance category","Automotive category analysis"]},
      {name:"KORE Recap Report",rate:5000,cat:"Analysis",index_score:4,recurring:false,examples:["Full KORE platform recap","Sponsorship ROI summary"]},
    ],
    internal:[
      {name:"Corporate Partnerships Advisory",rate:75,cat:"Advisory",index_score:1,recurring:true,examples:["Partner check-in","Strategy advisory call"]},
      {name:"Partnership Strategy Session",rate:150,cat:"Strategy",index_score:2,recurring:false,examples:["Category exclusivity strategy","Inventory bundling session"]},
    ],
  },
  "Marketing": {
    external:[
      {name:"Market Analysis",rate:2500,cat:"Research",index_score:4,recurring:false,examples:["DMA market sizing","Fan base demographic study"]},
      {name:"Fanbase Trending Report",rate:1000,cat:"Research",index_score:3,recurring:false,examples:["Social sentiment analysis","Fan engagement trends"]},
      {name:"Broadcast Report",rate:250,cat:"Analysis",index_score:2,recurring:true,examples:["Weekly broadcast metrics","TV viewership recap"]},
      {name:"Social Media Report",rate:250,cat:"Analysis",index_score:2,recurring:true,examples:["Monthly social report","Platform performance recap"]},
      {name:"Annual Club Recap Report",rate:250,cat:"Analysis",index_score:2,recurring:true,examples:["Season recap report","Year in review"]},
      {name:"Marketing Audit",rate:75,cat:"Analysis",index_score:3,recurring:false,examples:["Digital presence audit","Campaign effectiveness review"]},
      {name:"Go to Market Strategy",rate:500,cat:"Strategy",index_score:4,recurring:false,examples:["Season launch strategy","New market entry plan"]},
      {name:"Web / CX Audit",rate:750,cat:"Analysis",index_score:3,recurring:false,examples:["Website UX audit","Digital experience review"]},
      {name:"Marketing Playbook",rate:1000,cat:"Strategy",index_score:3,recurring:false,examples:["Campaign playbook","Content strategy guide"]},
    ],
    internal:[
      {name:"Social Media Report",rate:250,cat:"Analysis",index_score:2,recurring:true,leagueSelect:true,examples:["Championship social recap","League One social metrics"]},
      {name:"Marketing Advisory",rate:75,cat:"Advisory",index_score:1,recurring:true,examples:["Marketing check-in call","Campaign advisory"]},
      {name:"Monthly Marketing Newsletter",rate:100,cat:"Creative",index_score:1,recurring:true,examples:["Monthly league newsletter","Club marketing digest"]},
    ],
  },
  "Consumer Products": {
    external:[
      {name:"CPG Newsletter",rate:100,cat:"Creative",index_score:1,recurring:true,examples:["Monthly CPG newsletter","League-wide digest"]},
      {name:"Food and Beverage Analysis",rate:1500,cat:"Analysis",index_score:3,recurring:false,examples:["Concessions revenue analysis","F&B benchmarking study"]},
      {name:"Consumer Segmentation Report",rate:1000,cat:"Research",index_score:4,recurring:false,examples:["Fan segmentation study","Customer persona analysis"]},
    ],
    internal:[
      {name:"Consumer Products Advisory",rate:75,cat:"Advisory",index_score:1,recurring:true,examples:["Product check-in","Advisory call"]},
    ],
  },
  "Ticketing": {
    external:[
      {name:"Ticketing Analysis",rate:500,cat:"Analysis",index_score:3,recurring:false,examples:["Season ticket audit","Pricing analysis"]},
      {name:"Club Revenue Pathway Model",rate:2500,cat:"Valuation",index_score:5,recurring:false,examples:["Multi-year revenue model","Ticket revenue pathway"]},
      {name:"Dynamic Pricing Model",rate:3500,cat:"Valuation",index_score:4,recurring:false,examples:["Demand-based pricing build","Variable pricing structure"]},
      {name:"Season Ticket Revenue Model",rate:3000,cat:"Valuation",index_score:4,recurring:false,examples:["ST revenue projection","Package revenue model"]},
      {name:"Stadium Pricing & Yield Management Analysis",rate:3500,cat:"Analysis",index_score:4,recurring:false,examples:["Section-by-section pricing","Yield optimization model"]},
      {name:"Attendance Recovery Plan",rate:2500,cat:"Strategy",index_score:4,recurring:false,examples:["Win-back strategy","Attendance turnaround plan"]},
      {name:"Ticketing Playbook / Framework",rate:1500,cat:"Strategy",index_score:3,recurring:false,examples:["Best-practice playbook","ST sales framework"]},
      {name:"Season Ticket Retention Report",rate:250,cat:"Analysis",index_score:2,recurring:true,examples:["Renewal rate tracking","Churn analysis"]},
      {name:"Group Sales Performance Report",rate:250,cat:"Analysis",index_score:2,recurring:true,examples:["Group sales recap","Segment performance"]},
      {name:"Ticketing Executive Summary",rate:250,cat:"Analysis",index_score:2,recurring:true,examples:["Leadership ticketing brief","Board-ready summary"]},
    ],
    internal:[
      {name:"Ticketing Advisory",rate:75,cat:"Advisory",index_score:1,recurring:true,examples:["Ticketing check-in","Advisory call"]},
      {name:"Ticketing Check-In Call",rate:75,cat:"Advisory",index_score:1,recurring:true,examples:["Weekly check-in","Status update call"]},
    ],
  },
  "Youth / Facilities": {
    external:[
      {name:"Youth Program Analysis",rate:1500,cat:"Analysis",index_score:3,recurring:false,examples:["Academy program review","Youth league audit"]},
      {name:"Facilities Assessment",rate:2000,cat:"Analysis",index_score:3,recurring:false,examples:["Facility condition report","Standards compliance review"]},
      {name:"Youth Strategy Deck",rate:2500,cat:"Strategy",index_score:4,recurring:false,examples:["Academy strategy plan","Youth development roadmap"]},
      {name:"Facilities Feasibility Study",rate:3000,cat:"Analysis",index_score:4,recurring:false,examples:["New facility feasibility","Site evaluation report"]},
      {name:"Facilities Asset Valuation",rate:4000,cat:"Valuation",index_score:4,recurring:false,examples:["Facility valuation","Asset financial model"]},
      {name:"Facilities Revenue Strategy",rate:3500,cat:"Strategy",index_score:4,recurring:false,examples:["Revenue generation plan","Sponsorship integration"]},
    ],
    internal:[
      {name:"Youth / Facilities Advisory",rate:75,cat:"Advisory",index_score:1,recurring:true,examples:["Advisory call","Check-in session"]},
    ],
  },
  "League Initiatives": {
    internal:[
      {name:"Leadership Briefing",rate:500,cat:"Strategy",index_score:4,recurring:false,examples:["Executive briefing","Board presentation"]},
      {name:"Strategic Planning Session",rate:750,cat:"Strategy",index_score:4,recurring:false,examples:["League strategy session","Annual planning meeting"]},
      {name:"Institutional Advisory",rate:75,cat:"Advisory",index_score:1,recurring:true,examples:["Leadership advisory","Governance guidance"]},
      {name:"Cross-Department Initiative",rate:500,cat:"Strategy",index_score:4,recurring:false,examples:["Cross-functional project","League-wide initiative"]},
    ],
  },
};

function buildDeptItems(deliverables) {
  const map = {};
  deliverables.forEach(d => {
    const dept = normalizeDept(d.dept);
    if (!dept || !DEPT_CONFIG[dept]) return;
    if (!map[dept]) map[dept] = { external:[], internal:[] };
    const isRecurring = String(d.recurring||"").toLowerCase() === "true";
    const item = {
      name:        String(d.name||"").trim(),
      rate:        Number(d.rate)||0,
      cat:         String(d.cat||d.category||"Other").trim(),
      index_score: Number(d.index_score)||1,
      recurring:   isRecurring,
      examples:    d.examples ? String(d.examples).split(";").map(e=>e.trim()).filter(Boolean) : [],
      leagueSelect: String(d.name||"").trim()==="Social Media Report" && dept==="Marketing",
    };
    if (!item.name) return;
    const t = String(d.type||d["type (internal or external)"]||"").toLowerCase().trim();
    const isBoth = t.includes("external") && t.includes("internal");
    if (isBoth) { map[dept].external.push(item); map[dept].internal.push({...item}); }
    else if (t==="internal") { map[dept].internal.push(item); }
    else { map[dept].external.push(item); }
  });
  return map;
}

function leagueForClub(club, cbl) {
  const src = cbl || FALLBACK_CLUBS;
  for (const [league, clubs] of Object.entries(src))
    if (clubs.includes(club)) return league;
  return "League-wide";
}

const LEAGUE_STYLES = {
  "Championship":  { color:"#1D4ED8", bg:"#EFF6FF", border:"#BFDBFE" },
  "League One":    { color:"#047857", bg:"#F0FDF4", border:"#A7F3D0" },
  "Super League":  { color:"#7C3AED", bg:"#F5F3FF", border:"#DDD6FE" },
  "Expansion":     { color:"#B45309", bg:"#FFFBEB", border:"#FDE68A" },
  "League-wide":   { color:"#6B7280", bg:"#F3F4F6", border:"#D1D5DB" },
  "Internal":      { color:"#047857", bg:"#F0FDF4", border:"#A7F3D0" },
};

const CAT_COLORS = {
  "Valuation":"#7C3AED","Analysis":"#0369A1","Research":"#047857","Strategy":"#B45309",
  "Creative":"#DB2777","Advisory":"#475569","Commercial Intelligence":"#0891B2",
  "Valuation & Pricing":"#7C3AED","Revenue Strategy":"#B45309","Sales Enablement":"#047857",
  "Performance & Advisory":"#475569","Benchmarking & Pricing":"#0369A1","Revenue Modeling":"#4338CA",
  "Executive Summaries/POV":"#6B7280","Sales Training":"#DB2777",
  "Stadium Pricing/Yield Management Analysis":"#0891B2","Ticketing Analysis/Consultancy":"#047857",
  "Strategy/Recovery Plans":"#B45309","Staffing Analysis":"#475569","Ticketing Model":"#4338CA",
  "Calls & Comms":"#6B7280","Frameworks & Playbooks":"#0369A1","Strategy & Advisory":"#4338CA",
  "Analysis & Research":"#0369A1","Business & Valuation":"#7C3AED","Programming & Operations":"#047857",
  "Community Development":"#0891B2","Other":"#6B7280",
};

// ─────────────────────────────────────────────────────────────────────────────
//  PASSWORD SCREEN
// ─────────────────────────────────────────────────────────────────────────────

function PasswordScreen({ onUnlock }) {
  const [input, setInput] = useState("");
  const [error, setError] = useState(false);
  const attempt = () => {
    if (input===APP_PASSWORD) { sessionStorage.setItem("dept_unlocked","1"); onUnlock(); }
    else { setError(true); setInput(""); }
  };
  return (
    <div style={{minHeight:"100vh",background:"#011e5c",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif"}}>
      <div style={{background:"#fff",borderRadius:20,padding:"48px 40px",width:"100%",maxWidth:400,textAlign:"center",boxShadow:"0 24px 64px rgba(0,0,0,.4)"}}>
        <div style={{fontSize:36,marginBottom:12}}>🏟️</div>
        <div style={{fontFamily:"'DM Serif Display',serif",fontSize:26,color:"#011e5c",marginBottom:6}}>CPG Deliverable Tracker</div>
        <div style={{fontSize:14,color:"#6B7280",marginBottom:32}}>Enter your password to continue</div>
        <input type="password" value={input} autoFocus
          onChange={e=>{setInput(e.target.value);setError(false);}}
          onKeyDown={e=>e.key==="Enter"&&attempt()} placeholder="Password"
          style={{width:"100%",fontFamily:"'DM Sans',sans-serif",fontSize:15,border:`2px solid ${error?"#EF4444":"#E5E7EB"}`,borderRadius:10,padding:"12px 16px",outline:"none",marginBottom:8,background:error?"#FEF2F2":"#fff",color:"#111",textAlign:"center",letterSpacing:2}}
        />
        {error&&<div style={{color:"#EF4444",fontSize:13,fontWeight:600,marginBottom:8}}>Incorrect password. Try again.</div>}
        <button onClick={attempt} style={{width:"100%",background:"#011e5c",color:"#fff",border:"none",borderRadius:10,padding:13,fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:15,cursor:"pointer",marginTop:8}}>Enter →</button>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  SHARED UI
// ─────────────────────────────────────────────────────────────────────────────

function Badge({ type }) {
  const ext = type==="External";
  return <span style={{fontSize:11,fontWeight:700,letterSpacing:.5,padding:"2px 8px",borderRadius:20,background:ext?"#EFF6FF":"#F0FDF4",color:ext?"#1D4ED8":"#047857",border:`1px solid ${ext?"#BFDBFE":"#A7F3D0"}`}}>{type}</span>;
}

function RecurringBadge() {
  return <span style={{fontSize:10,fontWeight:700,padding:"2px 7px",borderRadius:6,background:"#FEF9C3",color:"#854D0E",border:"1px solid #FDE68A"}}>Recurring</span>;
}

function CatBadge({ cat }) {
  const color = CAT_COLORS[cat]||"#6B7280";
  return <span style={{fontSize:10,fontWeight:700,letterSpacing:.4,padding:"2px 7px",borderRadius:6,background:color+"18",color,border:`1px solid ${color}33`}}>{cat}</span>;
}

function LeagueBadge({ league }) {
  const s = LEAGUE_STYLES[league]||LEAGUE_STYLES["League-wide"];
  return <span style={{fontSize:11,fontWeight:700,padding:"2px 8px",borderRadius:6,background:s.bg,color:s.color,border:`1px solid ${s.border}`}}>{league}</span>;
}

function DeptChip({ dept }) {
  const cfg = DEPT_CONFIG[dept];
  if (!cfg) return <span>{dept}</span>;
  return <span style={{background:cfg.light,color:cfg.color,border:`1px solid ${cfg.border}`,borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>{dept}</span>;
}

function KpiCard({ label, value, color, sub, tag }) {
  return (
    <div style={{background:"#fff",borderRadius:14,padding:"20px 24px",border:"1px solid #E5E7EB",flex:1,minWidth:150,boxShadow:"0 1px 6px rgba(0,0,0,.06)"}}>
      <div style={{fontSize:11,fontWeight:700,letterSpacing:1,color:"#9CA3AF",marginBottom:6,display:"flex",alignItems:"center",gap:6}}>
        {label}{tag&&<span style={{fontSize:10,background:"#FEF9C3",color:"#854D0E",border:"1px solid #FDE68A",borderRadius:4,padding:"1px 5px",fontWeight:700}}>{tag}</span>}
      </div>
      <div style={{fontSize:28,fontWeight:800,color,fontFamily:"'DM Serif Display',serif",lineHeight:1}}>{value}</div>
      {sub&&<div style={{fontSize:12,color:"#6B7280",marginTop:4}}>{sub}</div>}
    </div>
  );
}

function TblWrap({ children }) {
  return (
    <div style={{background:"#fff",borderRadius:14,border:"1px solid #E5E7EB",overflow:"hidden",boxShadow:"0 1px 6px rgba(0,0,0,.05)"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,fontFamily:"'DM Sans',sans-serif"}}>{children}</table>
    </div>
  );
}

const TH = ({ children, right }) => <th style={{padding:"11px 14px",fontWeight:700,fontSize:11,letterSpacing:.5,color:"#6B7280",borderBottom:"1px solid #E5E7EB",background:"#F9FAFB",textAlign:right?"right":"left"}}>{children}</th>;
const TD = ({ children, right, bold, color }) => <td style={{padding:"11px 14px",textAlign:right?"right":"left",fontWeight:bold?700:400,color:color||"#374151"}}>{children}</td>;

function Toast({ msg, color, onDone }) {
  useEffect(()=>{ const t=setTimeout(onDone,2800); return()=>clearTimeout(t); },[onDone]);
  return <div style={{position:"fixed",bottom:28,right:28,zIndex:9999,background:color,color:"#fff",padding:"14px 22px",borderRadius:12,fontWeight:600,fontSize:15,boxShadow:"0 8px 32px rgba(0,0,0,.22)",fontFamily:"'DM Sans',sans-serif"}}>{msg}</div>;
}

// ─────────────────────────────────────────────────────────────────────────────
//  EXTERNAL MODAL
// ─────────────────────────────────────────────────────────────────────────────

function ExternalModal({ item, deptCfg, clubsByLeague, onConfirm, onCancel }) {
  const [selection,    setSelection]   = useState("single");
  const [chosenLeague, setChosenLeague]= useState("");
  const [chosenClubs,  setChosenClubs] = useState([]);
  const [chosenClub,   setChosenClub]  = useState("");
  const [notes,        setNotes]       = useState("");

  const allFlat = Object.entries(clubsByLeague).flatMap(([league,clubs])=>clubs.map(club=>({club,league})));
  const getEntries = () => {
    if (selection==="all")    return allFlat;
    if (selection==="league") return chosenLeague?(clubsByLeague[chosenLeague]||[]).map(club=>({club,league:chosenLeague})):[];
    if (selection==="multi")  return chosenClubs.map(club=>({club,league:leagueForClub(club,clubsByLeague)}));
    return chosenClub?[{club:chosenClub,league:leagueForClub(chosenClub,clubsByLeague)}]:[];
  };
  const entries=getEntries(), count=entries.length, totalVal=count*item.rate;
  const toggleClub=club=>setChosenClubs(prev=>prev.includes(club)?prev.filter(c=>c!==club):[...prev,club]);
  const ss={width:"100%",fontFamily:"'DM Sans',sans-serif",fontSize:14,border:`2px solid ${deptCfg.color}`,borderRadius:8,padding:"9px 12px",background:"#fff",color:"#111",cursor:"pointer",outline:"none",marginTop:8,appearance:"none"};

  const OBox=({id,title,sub,children})=>(
    <div onClick={()=>setSelection(id)} style={{border:`2px solid ${selection===id?deptCfg.color:"#E5E7EB"}`,borderRadius:10,padding:"12px 14px",marginBottom:8,cursor:"pointer",background:selection===id?deptCfg.light:"#fff"}}>
      <div style={{display:"flex",alignItems:"center",gap:10}}>
        <div style={{width:16,height:16,borderRadius:"50%",border:`2px solid ${selection===id?deptCfg.color:"#D1D5DB"}`,background:selection===id?deptCfg.color:"#fff",flexShrink:0}}/>
        <div><div style={{fontWeight:700,fontSize:14,color:"#111827"}}>{title}</div><div style={{fontSize:12,color:"#6B7280"}}>{sub}</div></div>
      </div>
      {selection===id&&children&&<div style={{marginTop:10}}>{children}</div>}
    </div>
  );

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:16,padding:"28px 28px 24px",width:"100%",maxWidth:500,boxShadow:"0 24px 64px rgba(0,0,0,.22)",fontFamily:"'DM Sans',sans-serif",maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{fontFamily:"'DM Serif Display',serif",fontSize:19,color:"#111827",marginBottom:4}}>Log External Deliverable</div>
        <div style={{fontSize:13,color:"#6B7280",marginBottom:20}}>{item.name} · <strong>{fmt$(item.rate)} each</strong>{item.recurring&&<span style={{marginLeft:8}}><RecurringBadge/></span>}</div>
        <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:10}}>ATTRIBUTE TO</label>

        <OBox id="single" title="🏟️ Specific Club" sub="Log for one individual club">
          <select value={chosenClub} onChange={e=>setChosenClub(e.target.value)} style={ss}>
            <option value="">Select a club…</option>
            {Object.entries(clubsByLeague).map(([league,clubs])=>(
              <optgroup key={league} label={`── ${league} ──`}>
                {clubs.map(c=><option key={c} value={c}>{c}</option>)}
              </optgroup>
            ))}
          </select>
        </OBox>

        <OBox id="multi" title="✅ Multiple Clubs" sub="Pick two or more specific clubs">
          <div style={{marginTop:8,maxHeight:220,overflowY:"auto",border:`1px solid ${deptCfg.border}`,borderRadius:8,padding:"4px 0"}}>
            {Object.entries(clubsByLeague).map(([league,clubs])=>(
              <div key={league}>
                <div style={{fontSize:10,fontWeight:700,color:"#6B7280",letterSpacing:.5,padding:"5px 10px",background:"#F8FAFC"}}>{league.toUpperCase()}</div>
                {clubs.map(c=>{
                  const checked=chosenClubs.includes(c);
                  return(<div key={c} onClick={e=>{e.stopPropagation();toggleClub(c);}} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 10px",cursor:"pointer",background:checked?deptCfg.light:"transparent"}}>
                    <div style={{width:15,height:15,borderRadius:4,border:`2px solid ${checked?deptCfg.color:"#D1D5DB"}`,background:checked?deptCfg.color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                      {checked&&<span style={{color:"#fff",fontSize:10,fontWeight:900}}>✓</span>}
                    </div>
                    <span style={{fontSize:13,color:"#374151"}}>{c}</span>
                  </div>);
                })}
              </div>
            ))}
          </div>
          {chosenClubs.length>0&&<div style={{fontSize:12,color:deptCfg.color,fontWeight:600,marginTop:6}}>{chosenClubs.length} club{chosenClubs.length>1?"s":""} selected</div>}
        </OBox>

        <OBox id="league" title="📋 Entire League" sub="Logs one line item per club in that league">
          <select value={chosenLeague} onChange={e=>setChosenLeague(e.target.value)} style={ss}>
            <option value="">Select a league…</option>
            {Object.keys(clubsByLeague).map(l=><option key={l} value={l}>{l} — {clubsByLeague[l].length} clubs</option>)}
          </select>
          {chosenLeague&&(
            <div style={{marginTop:8,background:"#F8FAFC",borderRadius:8,padding:"10px 12px"}}>
              <div style={{fontSize:12,fontWeight:700,color:"#374151",marginBottom:6}}>Will log {(clubsByLeague[chosenLeague]||[]).length} entries:</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:4}}>{(clubsByLeague[chosenLeague]||[]).map(c=><span key={c} style={{fontSize:11,background:"#fff",border:"1px solid #E5E7EB",borderRadius:6,padding:"2px 7px"}}>{c}</span>)}</div>
            </div>
          )}
        </OBox>

        <OBox id="all" title="🌐 All Clubs" sub={`Logs one line item for every club (${allFlat.length} total)`}>
          <div style={{background:"#F8FAFC",borderRadius:8,padding:"10px 12px"}}>
            {Object.entries(clubsByLeague).map(([league,clubs])=>(
              <div key={league} style={{marginBottom:6}}>
                <div style={{fontSize:11,fontWeight:700,color:"#6B7280",marginBottom:3}}>{league} ({clubs.length})</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:3}}>{clubs.map(c=><span key={c} style={{fontSize:10,background:"#fff",border:"1px solid #E5E7EB",borderRadius:5,padding:"1px 6px"}}>{c}</span>)}</div>
              </div>
            ))}
          </div>
        </OBox>

        {count>0&&(
          <div style={{background:deptCfg.light,border:`1px solid ${deptCfg.border}`,borderRadius:10,padding:"12px 14px",marginBottom:4}}>
            <div style={{fontSize:13,fontWeight:700,color:deptCfg.color}}>
              {count===1?`1 entry · ${fmt$(item.rate)}`:`${count} entries × ${fmt$(item.rate)} = `}
              {count>1&&<span style={{fontSize:16}}>{fmt$(totalVal)} total</span>}
            </div>
          </div>
        )}
        {/* Notes field */}
        <div style={{marginTop:8}}>
          <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:6}}>NOTES <span style={{fontWeight:400,color:"#9CA3AF"}}>(optional)</span></label>
          <textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Any additional context for this deliverable..." rows={2}
            style={{width:"100%",fontFamily:"'DM Sans',sans-serif",fontSize:13,border:"1.5px solid #E5E7EB",borderRadius:8,padding:"9px 12px",outline:"none",resize:"vertical",color:"#374151",background:"#FAFAFA"}}/>
        </div>

        <div style={{display:"flex",gap:10,marginTop:12}}>
          <button onClick={()=>count&&onConfirm(entries,notes)} disabled={!count} style={{flex:1,padding:11,border:"none",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:14,color:"#fff",background:count?deptCfg.color:"#D1D5DB",cursor:count?"pointer":"not-allowed"}}>
            {count>1?`Log ${count} Entries ▶`:"Confirm & Log ▶"}
          </button>
          <button onClick={onCancel} style={{padding:"11px 18px",border:"1.5px solid #E5E7EB",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:14,color:"#6B7280",background:"#fff",cursor:"pointer"}}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  INTERNAL MODAL
// ─────────────────────────────────────────────────────────────────────────────

function InternalModal({ item, deptCfg, dept, clubsByLeague, onConfirm, onCancel }) {
  const [recipient,    setRecipient]   = useState("");
  const [chosenLeague, setChosenLeague]= useState("");
  const [notes,        setNotes]       = useState("");
  const recipients = INTERNAL_RECIPIENTS[dept]||[];
  const isLeagueSelect = item.leagueSelect;

  const getEntries = () => {
    if (isLeagueSelect&&chosenLeague) return (clubsByLeague[chosenLeague]||[]).map(club=>({club,league:chosenLeague}));
    return recipient?[{club:recipient,league:"Internal"}]:[];
  };
  const entries=getEntries(), count=entries.length;

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:16,padding:"28px 28px 24px",width:"100%",maxWidth:460,boxShadow:"0 24px 64px rgba(0,0,0,.22)",fontFamily:"'DM Sans',sans-serif",maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{fontFamily:"'DM Serif Display',serif",fontSize:19,color:"#111827",marginBottom:4}}>Log Internal Deliverable</div>
        <div style={{fontSize:13,color:"#6B7280",marginBottom:20}}>{item.name} · <strong>{fmt$(item.rate)}</strong>{item.recurring&&<span style={{marginLeft:8}}><RecurringBadge/></span>}</div>

        {isLeagueSelect?(
          <>
            <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:10}}>SELECT LEAGUE</label>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {Object.keys(clubsByLeague).map(l=>(
                <div key={l} onClick={()=>setChosenLeague(l)} style={{border:`2px solid ${chosenLeague===l?deptCfg.color:"#E5E7EB"}`,borderRadius:10,padding:"12px 14px",cursor:"pointer",background:chosenLeague===l?deptCfg.light:"#fff",display:"flex",alignItems:"center",gap:10}}>
                  <div style={{width:16,height:16,borderRadius:"50%",border:`2px solid ${chosenLeague===l?deptCfg.color:"#D1D5DB"}`,background:chosenLeague===l?deptCfg.color:"#fff",flexShrink:0}}/>
                  <div>
                    <div style={{fontWeight:700,fontSize:14,color:"#111827"}}>{l} Social</div>
                    <div style={{fontSize:12,color:"#6B7280"}}>{(clubsByLeague[l]||[]).length} clubs</div>
                  </div>
                </div>
              ))}
            </div>
            {chosenLeague&&<div style={{marginTop:12,background:deptCfg.light,border:`1px solid ${deptCfg.border}`,borderRadius:10,padding:"10px 14px"}}><div style={{fontSize:13,fontWeight:700,color:deptCfg.color}}>{count} entries × {fmt$(item.rate)} = {fmt$(count*item.rate)} total</div></div>}
          </>
        ):(
          <>
            <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:10}}>INTERNAL RECIPIENT</label>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {recipients.map(r=>(
                <div key={r} onClick={()=>setRecipient(r)} style={{border:`2px solid ${recipient===r?deptCfg.color:"#E5E7EB"}`,borderRadius:10,padding:"12px 14px",cursor:"pointer",background:recipient===r?deptCfg.light:"#fff",display:"flex",alignItems:"center",gap:10}}>
                  <div style={{width:16,height:16,borderRadius:"50%",border:`2px solid ${recipient===r?deptCfg.color:"#D1D5DB"}`,background:recipient===r?deptCfg.color:"#fff",flexShrink:0}}/>
                  <span style={{fontWeight:600,fontSize:14,color:"#111827"}}>{r}</span>
                </div>
              ))}
            </div>
          </>
        )}

        {/* Notes field */}
        <div style={{marginTop:12}}>
          <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:6}}>NOTES <span style={{fontWeight:400,color:"#9CA3AF"}}>(optional)</span></label>
          <textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Any additional context..." rows={2}
            style={{width:"100%",fontFamily:"'DM Sans',sans-serif",fontSize:13,border:"1.5px solid #E5E7EB",borderRadius:8,padding:"9px 12px",outline:"none",resize:"vertical",color:"#374151",background:"#FAFAFA"}}/>
        </div>

        <div style={{display:"flex",gap:10,marginTop:12}}>
          <button onClick={()=>count&&onConfirm(entries,notes)} disabled={!count} style={{flex:1,padding:11,border:"none",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:14,color:"#fff",background:count?deptCfg.color:"#D1D5DB",cursor:count?"pointer":"not-allowed"}}>
            {count>1?`Log ${count} Entries ▶`:"Confirm & Log ▶"}
          </button>
          <button onClick={onCancel} style={{padding:"11px 18px",border:"1.5px solid #E5E7EB",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:14,color:"#6B7280",background:"#fff",cursor:"pointer"}}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  CATEGORY GROUPED CARDS
// ─────────────────────────────────────────────────────────────────────────────

function CategoryGroupedCards({ items, log, dept, cfg, pulsingIdx, onLogClick }) {
  const [collapsed, setCollapsed] = useState({});
  const groups=[], seen={};
  items.forEach((item,idx)=>{
    if(!seen[item.cat]){seen[item.cat]=true;groups.push({cat:item.cat,entries:[]});}
    groups.find(g=>g.cat===item.cat).entries.push({item,idx});
  });
  const toggleCat=cat=>setCollapsed(prev=>({...prev,[cat]:!prev[cat]}));

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      {groups.map(({cat,entries})=>{
        const isCollapsed=collapsed[cat];
        const color=CAT_COLORS[cat]||"#6B7280";
        const groupCount=entries.reduce((s,{item})=>s+log.filter(e=>e.dept===dept&&e.name===item.name).length,0);
        return(
          <div key={cat}>
            <div onClick={()=>toggleCat(cat)} style={{display:"flex",alignItems:"center",gap:12,marginBottom:isCollapsed?0:12,cursor:"pointer",userSelect:"none"}}>
              <div style={{width:22,height:22,borderRadius:6,background:color+"18",border:`1px solid ${color}33`,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,transition:"transform .2s",transform:isCollapsed?"rotate(-90deg)":"rotate(0deg)"}}>
                <span style={{color,fontSize:10,fontWeight:900,lineHeight:1}}>▾</span>
              </div>
              <div style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:"#111827"}}>{cat}</div>
              <div style={{flex:1,height:1,background:`${color}30`}}/>
              <div style={{display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
                {groupCount>0&&<span style={{fontSize:12,fontWeight:700,color,background:color+"18",border:`1px solid ${color}33`,borderRadius:20,padding:"2px 10px"}}>{groupCount} logged</span>}
                <span style={{fontSize:12,color:"#9CA3AF",fontFamily:"'DM Sans',sans-serif"}}>{entries.length} {entries.length===1?"deliverable":"deliverables"}</span>
              </div>
            </div>
            {!isCollapsed&&(
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(310px,1fr))",gap:12}}>
                {entries.map(({item,idx})=>{
                  const count=log.filter(e=>e.dept===dept&&e.name===item.name).length;
                  const isPulsing=pulsingIdx===idx;
                  return(
                    <div key={item.name} style={{background:"#fff",border:`1.5px solid ${isPulsing?cfg.color:"#E5E7EB"}`,borderRadius:12,padding:16,display:"flex",flexDirection:"column",gap:8,boxShadow:isPulsing?`0 0 0 4px ${cfg.color}33`:"0 1px 4px rgba(0,0,0,.06)",transition:"border-color .2s,box-shadow .2s"}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
                        <span style={{fontFamily:"'DM Serif Display',serif",fontSize:15,color:"#111827",lineHeight:1.3,flex:1}}>{item.name}</span>
                        {item.recurring&&<RecurringBadge/>}
                      </div>
                      {item.examples&&item.examples.length>0&&(
                        <div style={{fontSize:11,color:"#6B7280",fontStyle:"italic",lineHeight:1.5}}>
                          {item.examples.slice(0,4).join(" · ")}
                          {item.examples.length>4&&<span style={{color:cfg.color,fontStyle:"normal",fontWeight:600,marginLeft:4}}>+{item.examples.length-4} more</span>}
                        </div>
                      )}
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:2}}>
                        <span style={{fontSize:13,color:"#6B7280"}}>
                          {item.rate>0&&<strong style={{color:"#374151"}}>{fmt$(item.rate)}</strong>}
                          <span style={{marginLeft:8,fontSize:11,background:"#F3F4F6",borderRadius:6,padding:"1px 7px",color:"#475569"}}>Index {item.index_score}</span>
                        </span>
                        <div style={{display:"flex",alignItems:"center",gap:8}}>
                          {count>0&&<span style={{background:cfg.light,color:cfg.color,border:`1px solid ${cfg.border}`,borderRadius:20,padding:"2px 10px",fontWeight:700,fontSize:13}}>×{count}</span>}
                          <button onClick={()=>onLogClick(idx)} style={{background:cfg.color,color:"#fff",border:"none",borderRadius:8,padding:"9px 18px",fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>LOG IT ▶</button>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  DEPT TAB
// ─────────────────────────────────────────────────────────────────────────────

function DeptTab({ dept, log, onLog, onBulkLog, onBulkComplete, deptItems, clubsByLeague }) {
  const cfg=DEPT_CONFIG[dept];
  const isToggle=cfg.hasToggle, isIntOnly=!cfg.hasToggle;
  const [mode,setMode]=useState("external");
  const [staffName,setStaffName]=useState("");
  const [nameErr,setNameErr]=useState(false);
  const [modal,setModal]=useState(null);
  const [pulsingIdx,setPulsingIdx]=useState(null);

  const deptData=deptItems[dept]||{external:[],internal:[]};
  const activeMode=isIntOnly?"internal":mode;
  const items=activeMode==="internal"?deptData.internal:deptData.external;
  const deptLog=log.filter(e=>e.dept===dept);
  const deptTotal=deptLog.reduce((s,e)=>s+e.rate,0);

  const handleLogClick=idx=>{if(!staffName){setNameErr(true);return;}setNameErr(false);setModal(idx);};

  const handleConfirm=async (entries, notes="")=>{
    const item=items[modal];
    setModal(null);setPulsingIdx(modal);setTimeout(()=>setPulsingIdx(null),400);
    const total=entries.length;
    if(total===1){
      const{club,league}=entries[0];
      await onLog({dept,name:item.name,rate:item.rate,type:activeMode==="internal"?"Internal":"External",cat:item.cat,index_score:item.index_score||1,recurring:item.recurring||false,staff:staffName,club,league,notes,bulkSilent:false});
    } else {
      // Build all entries first, add to UI optimistically, then send as one batch request
      const newEntries = entries.map(({club,league})=>({
        dept,name:item.name,rate:item.rate,
        type:activeMode==="internal"?"Internal":"External",
        cat:item.cat,index_score:item.index_score||1,
        recurring:item.recurring||false,staff:staffName,club,league,notes,
        id:`${Date.now()}-${Math.random().toString(36).slice(2)}`,
        ts:Date.now(),
      }));
      // Add all to UI at once
      onBulkLog(newEntries);
      // Send as single batch request to Apps Script
      try {
        await dbInsertBatch(newEntries.map(({bulkSilent,...e})=>e));
      } catch(err) {
        console.error("Batch insert failed:", err);
      }
      onBulkComplete(total, item.name, item.rate*total);
    }
  };

  return(
    <div>
      <div style={{background:cfg.light,border:`1px solid ${cfg.border}`,borderRadius:12,padding:"14px 18px",marginBottom:20,display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
        <span style={{fontFamily:"'DM Serif Display',serif",fontSize:14,color:"#374151",whiteSpace:"nowrap"}}>Your name:</span>
        <div style={{position:"relative"}}>
          <select value={staffName} onChange={e=>{setStaffName(e.target.value);setNameErr(false);}}
            style={{fontFamily:"'DM Sans',sans-serif",fontSize:14,fontWeight:600,border:`2px solid ${nameErr?"#EF4444":cfg.color}`,borderRadius:8,padding:"8px 36px 8px 12px",outline:"none",background:"#fff",color:staffName?"#111":"#9CA3AF",cursor:"pointer",minWidth:200,appearance:"none",boxShadow:nameErr?"0 0 0 3px #FEE2E2":`0 0 0 3px ${cfg.color}22`}}>
            <option value="" disabled>Select your name…</option>
            {STAFF.map(n=><option key={n} value={n}>{n}</option>)}
          </select>
          <span style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",pointerEvents:"none",color:cfg.color,fontSize:12}}>▼</span>
        </div>
        {nameErr&&<span style={{color:"#EF4444",fontSize:13,fontWeight:600}}>⚠ Select your name first</span>}
        {isToggle&&(
          <div style={{display:"flex",background:"#F1F5F9",borderRadius:10,padding:3,gap:2}}>
            {["external","internal"].map(m=>(
              <button key={m} onClick={()=>setMode(m)} style={{padding:"7px 16px",borderRadius:8,border:"none",fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:12,cursor:"pointer",background:activeMode===m?cfg.color:"transparent",color:activeMode===m?"#fff":"#64748B",transition:"all .15s"}}>
                {m==="external"?"External":"Internal"}
              </button>
            ))}
          </div>
        )}
        {isIntOnly&&<span style={{fontSize:12,fontWeight:700,color:cfg.color,background:cfg.light,border:`1px solid ${cfg.border}`,borderRadius:8,padding:"5px 12px"}}>Internal Only</span>}
        <div style={{marginLeft:"auto",display:"flex",gap:24}}>
          <div style={{textAlign:"center"}}>
            <div style={{fontSize:22,fontWeight:800,color:cfg.color,fontFamily:"'DM Serif Display',serif"}}>{deptLog.length}</div>
            <div style={{fontSize:11,color:"#6B7280",letterSpacing:.4}}>LOGGED</div>
          </div>
          <div style={{textAlign:"center"}}>
            <div style={{fontSize:22,fontWeight:800,color:cfg.color,fontFamily:"'DM Serif Display',serif"}}>{fmt$(deptTotal)}</div>
            <div style={{fontSize:11,color:"#6B7280",letterSpacing:.4}}>TOTAL VALUE</div>
          </div>
        </div>
      </div>

      {isToggle&&(
        <div style={{marginBottom:14,display:"flex",alignItems:"center",gap:8}}>
          <Badge type={activeMode==="external"?"External":"Internal"}/>
          <span style={{fontSize:13,color:"#6B7280",fontFamily:"'DM Sans',sans-serif"}}>
            {activeMode==="external"?"Showing external deliverables — attributed to clubs":"Showing internal deliverables — attributed to USL teams"}
          </span>
        </div>
      )}

      {items.length===0
        ?<div style={{textAlign:"center",padding:"40px 0",color:"#9CA3AF",fontFamily:"'DM Sans',sans-serif"}}>No {activeMode} deliverables found — check your Google Sheet.</div>
        :<CategoryGroupedCards items={items} log={log} dept={dept} cfg={cfg} pulsingIdx={pulsingIdx} onLogClick={handleLogClick}/>
      }

      {modal!==null&&(
        activeMode==="internal"
          ?<InternalModal item={items[modal]} deptCfg={cfg} dept={dept} clubsByLeague={clubsByLeague} onConfirm={handleConfirm} onCancel={()=>setModal(null)}/>
          :<ExternalModal item={items[modal]} deptCfg={cfg} clubsByLeague={clubsByLeague} onConfirm={handleConfirm} onCancel={()=>setModal(null)}/>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  PIE CHART
// ─────────────────────────────────────────────────────────────────────────────

function PieChart({ slices, title, centerLabel, centerSub, size=200 }) {
  const [hovered,setHovered]=useState(null);
  const r=size/2-10,cx=size/2,cy=size/2;
  const total=slices.reduce((s,sl)=>s+sl.value,0);
  if(!total) return null;
  let cumAngle=-Math.PI/2;
  const arcs=slices.map((sl,i)=>{
    const angle=(sl.value/total)*2*Math.PI;
    const x1=cx+r*Math.cos(cumAngle),y1=cy+r*Math.sin(cumAngle);
    cumAngle+=angle;
    const x2=cx+r*Math.cos(cumAngle),y2=cy+r*Math.sin(cumAngle);
    return{...sl,x1,y1,x2,y2,large:angle>Math.PI?1:0,midAngle:cumAngle-angle/2,pct:sl.value/total,index:i};
  });
  const innerR=r*0.54;
  return(
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:14}}>
      {title&&<div style={{fontFamily:"'DM Serif Display',serif",fontSize:16,color:"#111827",textAlign:"center"}}>{title}</div>}
      <svg width={size} height={size} style={{overflow:"visible"}}>
        {arcs.map((arc,i)=>{
          const isH=hovered===i,off=isH?6:0,ox=off*Math.cos(arc.midAngle),oy=off*Math.sin(arc.midAngle);
          return<path key={i} d={`M ${cx+ox} ${cy+oy} L ${arc.x1+ox} ${arc.y1+oy} A ${r} ${r} 0 ${arc.large} 1 ${arc.x2+ox} ${arc.y2+oy} Z`} fill={arc.color} stroke="#fff" strokeWidth={2} style={{cursor:"pointer",filter:isH?`drop-shadow(0 4px 8px ${arc.color}66)`:"none"}} onMouseEnter={()=>setHovered(i)} onMouseLeave={()=>setHovered(null)}/>;
        })}
        <circle cx={cx} cy={cy} r={innerR} fill="#fff"/>
        {centerLabel&&<>
          <text x={cx} y={cy-6} textAnchor="middle" dominantBaseline="middle" style={{fontFamily:"'DM Serif Display',serif",fontSize:18,fontWeight:800,fill:"#111827"}}>{hovered!==null?`${(arcs[hovered].pct*100).toFixed(1)}%`:centerLabel}</text>
          <text x={cx} y={cy+14} textAnchor="middle" dominantBaseline="middle" style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,fill:"#6B7280"}}>{hovered!==null?arcs[hovered].label:centerSub}</text>
        </>}
      </svg>
      <div style={{display:"flex",flexDirection:"column",gap:6,width:"100%"}}>
        {arcs.map((arc,i)=>(
          <div key={i} onMouseEnter={()=>setHovered(i)} onMouseLeave={()=>setHovered(null)} style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",padding:"3px 6px",borderRadius:6,background:hovered===i?"#F8FAFC":"transparent"}}>
            <div style={{width:10,height:10,borderRadius:"50%",background:arc.color,flexShrink:0}}/>
            <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:12,color:"#374151",flex:1}}>{arc.label}</span>
            <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:12,fontWeight:700,color:arc.color}}>{(arc.pct*100).toFixed(1)}%</span>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  DASHBOARD
// ─────────────────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────────────────
//  RACK RATE PLACEHOLDERS (used when sheet rate is 0)
// ─────────────────────────────────────────────────────────────────────────────

const RACK_RATE_MAP = {
  // Corp Partnerships
  "Stadium Naming Rights Valuation": 20000,
  "Rate Card and Inventory Analysis": 4250,
  "Kit Asset Valuation": 5000,
  "LED & Field Board Pricing": 3000,
  "Standard Intelligence Product": 250,
  "Partnership Performance Product": 250,
  "Category Analysis": 75,
  "KORE Recap Report": 5000,
  "Strategic Advisory": 75,
  "Partnership Strategy Session": 150,
  // Marketing
  "Market Analysis": 2500,
  "Fanbase Trending Report": 1000,
  "Broadcast Report": 250,
  "Social Media Report": 250,
  "Annual Club Recap Report": 250,
  "Marketing Audit": 75,
  "Go to Market Strategy": 500,
  "Monthly Marketing Newsletter": 100,
  "1-on-1 Consultant Call": 75,
  "Marketing All-Call": 75,
  "Web / CX Audit": 750,
  "Marketing Budget Audit": 750,
  "Marketing Playbook": 1000,
  "Marketing Case Study": 500,
  "Quarterly Marketing Training": 300,
  "Revenue Leaders Call": 150,
  "Social Media Strategy Call": 150,
  "Expansion Marketing Consulting": 1500,
  "In-Person Marketing Visit": 300,
  "CPG Newsletter": 100,
  // Ticketing
  "Ticketing Analysis": 500,
  "Club Revenue Pathway Model": 2500,
  "Season Ticket Retention Report": 250,
  "Group Sales Performance Report": 250,
  "Comp Ticket Benchmarking Study": 750,
  "Dynamic Pricing Model": 3500,
  "Season Ticket Revenue Model": 3000,
  "Group Sales Revenue Model": 2000,
  "Ticketing Executive Summary": 250,
  "Ticket Sales Training Session": 500,
  "Stadium Pricing & Yield Management Analysis": 3500,
  "Ticketing Analysis & Consultancy": 500,
  "Attendance Recovery Plan": 2500,
  "Ticketing Department Staffing Analysis": 1000,
  "Custom Ticketing Model Build": 3000,
  "Ticketing Check-In Call": 75,
  "Ticketing Playbook / Framework": 1500,
  // Youth / Facilities
  "Youth Program Analysis": 1500,
  "Facilities Assessment": 2000,
  "Youth Strategy Deck": 2500,
  "Strategic Planning Session": 750,
  "Club Assessment Report": 1500,
  "Market Research Study": 2000,
  "Facilities Benchmarking Report": 1500,
  "Facilities Feasibility Study": 3000,
  "Facilities Asset Valuation": 4000,
  "Facilities Revenue Strategy": 3500,
  "Youth Programming Design": 2000,
  "Academy Staffing & Operations Plan": 1500,
  "Community Engagement Strategy": 1000,
  "Fan & Volunteer Development Program": 750,
  // Default
  "default": 250,
};

function getRackRate(entry) {
  if (entry.rate && entry.rate > 0) return entry.rate;
  return RACK_RATE_MAP[entry.name] || RACK_RATE_MAP["default"];
}

// ─────────────────────────────────────────────────────────────────────────────
//  SVG BAR CHART
// ─────────────────────────────────────────────────────────────────────────────

function HBarChart({ data, color, valueLabel, height=56 }) {
  const [hovered, setHovered] = useState(null);
  if (!data||!data.length) return <div style={{textAlign:"center",padding:"32px 0",color:"#9CA3AF",fontSize:13}}>No data</div>;
  const max = Math.max(...data.map(d=>d.value||0), 1);
  return (
    <div style={{display:"flex",flexDirection:"column",gap:12}}>
      {data.map((d,i) => {
        const pct = d.value/max*100;
        const isH = hovered===i;
        const barColor = d.color || color;
        return (
          <div key={d.label} onMouseEnter={()=>setHovered(i)} onMouseLeave={()=>setHovered(null)}
            style={{display:"flex",alignItems:"center",gap:12,cursor:"default"}}>
            <div style={{width:140,flexShrink:0,textAlign:"right",fontFamily:"'DM Sans',sans-serif",fontSize:13,fontWeight:600,color:isH?"#111827":"#374151",transition:"color .15s"}}>{d.label}</div>
            <div style={{flex:1,height:height,background:"#F3F4F6",borderRadius:8,overflow:"hidden",position:"relative"}}>
              <div style={{width:`${pct}%`,height:"100%",background:barColor,borderRadius:8,opacity:isH?1:.82,transition:"all .3s",boxShadow:isH?`0 2px 12px ${barColor}66`:"none"}}/>
            </div>
            <div style={{width:80,flexShrink:0,fontFamily:"'DM Serif Display',serif",fontSize:20,fontWeight:800,color:barColor}}>{d.value.toLocaleString()}</div>
            {d.sub && <div style={{width:60,flexShrink:0,fontSize:11,color:"#9CA3AF",fontFamily:"'DM Sans',sans-serif"}}>{d.sub}</div>}
          </div>
        );
      })}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  SVG DONUT CHART
// ─────────────────────────────────────────────────────────────────────────────

function DonutChart({ slices, size=220, title, subtitle }) {
  const [hovered, setHovered] = useState(null);
  const safeSlices = (slices||[]).filter(s=>s&&Number(s.value)>0);
  const total = safeSlices.reduce((s,sl)=>s+sl.value,0);
  if (!total||!safeSlices.length) return <div style={{textAlign:"center",padding:"40px 0",color:"#9CA3AF",fontSize:13}}>No data</div>;
  const r = size/2-16, cx=size/2, cy=size/2;
  let cumAngle = -Math.PI/2;
  const arcs = safeSlices.map((sl,i) => {
    const angle = (sl.value/total)*2*Math.PI;
    const x1=cx+r*Math.cos(cumAngle), y1=cy+r*Math.sin(cumAngle);
    cumAngle+=angle;
    const x2=cx+r*Math.cos(cumAngle), y2=cy+r*Math.sin(cumAngle);
    const midAngle=cumAngle-angle/2;
    return{...sl,x1,y1,x2,y2,large:angle>Math.PI?1:0,midAngle,pct:sl.value/total,index:i};
  });
  const innerR=r*0.56;
  return (
    <div style={{display:"flex",alignItems:"center",gap:24,flexWrap:"wrap"}}>
      <div style={{position:"relative",flexShrink:0}}>
        <svg width={size} height={size} style={{overflow:"visible"}}>
          {arcs.map((arc,i)=>{
            const isH=hovered===i;
            const off=isH?7:0, ox=off*Math.cos(arc.midAngle), oy=off*Math.sin(arc.midAngle);
            return(
              <path key={i}
                d={`M ${cx+ox} ${cy+oy} L ${arc.x1+ox} ${arc.y1+oy} A ${r} ${r} 0 ${arc.large} 1 ${arc.x2+ox} ${arc.y2+oy} Z`}
                fill={arc.color} stroke="#fff" strokeWidth={2.5}
                style={{cursor:"pointer",filter:isH?`drop-shadow(0 4px 12px ${arc.color}88)`:"none",transition:"filter .15s"}}
                onMouseEnter={()=>setHovered(i)} onMouseLeave={()=>setHovered(null)}
              />
            );
          })}
          <circle cx={cx} cy={cy} r={innerR} fill="#fff"/>
          <text x={cx} y={cy-10} textAnchor="middle" dominantBaseline="middle"
            style={{fontFamily:"'DM Serif Display',serif",fontSize:22,fontWeight:800,fill:"#111827"}}>
            {hovered!==null?`${(arcs[hovered].pct*100).toFixed(0)}%`:title}
          </text>
          <text x={cx} y={cy+14} textAnchor="middle" dominantBaseline="middle"
            style={{fontFamily:"'DM Sans',sans-serif",fontSize:12,fill:"#6B7280"}}>
            {hovered!==null?arcs[hovered].label:subtitle}
          </text>
        </svg>
      </div>
      {/* Legend */}
      <div style={{display:"flex",flexDirection:"column",gap:8,flex:1,minWidth:150}}>
        {arcs.map((arc,i)=>(
          <div key={i} onMouseEnter={()=>setHovered(i)} onMouseLeave={()=>setHovered(null)}
            style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",padding:"4px 8px",borderRadius:8,background:hovered===i?arc.color+"12":"transparent",transition:"background .15s"}}>
            <div style={{width:12,height:12,borderRadius:"50%",background:arc.color,flexShrink:0}}/>
            <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:13,color:"#374151",flex:1}}>{arc.label}</span>
            <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:13,fontWeight:700,color:arc.color}}>{(arc.pct*100).toFixed(0)}%</span>
            <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:"#9CA3AF",width:40,textAlign:"right"}}>{arc.value.toLocaleString()}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  DASHBOARD
// ─────────────────────────────────────────────────────────────────────────────

function Dashboard({ log, onExport, clubsByLeague }) {
  const [league,       setLeague]      = useState("all");
  const [filterYear,   setFilterYear]  = useState("all");
  const [filterMonth,  setFilterMonth] = useState("all");
  const [entryType,    setEntryType]   = useState("all");
  const [metric,       setMetric]      = useState("index");
  const [showAllClubs, setShowAllClubs]= useState(false);

  const LEAGUES = ["all","Championship","League One","Super League","Expansion"];
  const LEAGUE_COLORS = {"Championship":"#1D4ED8","League One":"#047857","Super League":"#7C3AED","Expansion":"#B45309","all":"#011e5c"};
  const DEPT_COLORS = {"Corp Partnerships":"#7C3AED","Marketing":"#0369A1","Consumer Products":"#B45309","Ticketing":"#047857","Youth / Facilities":"#0891B2","League Initiatives":"#4338CA","General":"#475569"};
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const availableYears = [...new Set((log||[]).map(e=>new Date(e.ts).getFullYear()))].sort((a,b)=>b-a);

  // Filters
  const timeFiltered = (log||[]).filter(e=>{
    const d=new Date(e.ts);
    if(filterYear!=="all"&&d.getFullYear()!==parseInt(filterYear)) return false;
    if(filterMonth!=="all"&&d.getMonth()!==parseInt(filterMonth)) return false;
    return true;
  });
  const typeFiltered = timeFiltered.filter(e=>{
    if(entryType==="strategic") return !e.recurring;
    if(entryType==="recurring") return e.recurring;
    return true;
  });
  const strategicLog = typeFiltered.filter(e=>!e.recurring);
  const recurringLog  = typeFiltered.filter(e=>e.recurring);
  const leagueFiltered = league==="all" ? typeFiltered : typeFiltered.filter(e=>e.league===league);
  const externalLog = leagueFiltered.filter(e=>e.type==="External");

  // Stats
  const total     = typeFiltered.length;
  const strategic = strategicLog.length;
  const recurring = recurringLog.length;
  const extCnt    = typeFiltered.filter(e=>e.type==="External").length;
  const intCnt    = typeFiltered.filter(e=>e.type==="Internal").length;
  const stratIdx  = typeFiltered.filter(e=>!e.recurring&&e.type==="External"&&e.index_score>0);
  const avgIndex  = stratIdx.length>0?(stratIdx.reduce((s,e)=>s+Number(e.index_score||1),0)/stratIdx.length).toFixed(1):"—";

  // Charts
  const byLeague = ["Championship","League One","Super League","Expansion"].map(l=>({
    label:l, value:typeFiltered.filter(e=>e.league===l).length, color:LEAGUE_COLORS[l]
  })).filter(d=>d.value>0);

  const deptCounts = {};
  leagueFiltered.forEach(e=>{ const k=e.dept||"General"; deptCounts[k]=(deptCounts[k]||0)+1; });
  const byDept = Object.entries(deptCounts).sort((a,b)=>b[1]-a[1]).map(([dept,count])=>({label:dept,value:count,color:DEPT_COLORS[dept]||"#475569"}));

  const usageByLeague = ["Championship","League One","Super League","Expansion"].map(l=>({
    label:l, value:typeFiltered.filter(e=>e.league===l).length, color:LEAGUE_COLORS[l]
  })).filter(d=>d.value>0);

  // Club map
  const clubMap={};
  externalLog.forEach(e=>{
    const c=e.club||"Unknown";
    if(!clubMap[c])clubMap[c]={club:c,league:e.league||"",count:0,value:0,indexSum:0,indexCount:0};
    clubMap[c].count++;
    clubMap[c].value+=getRackRate(e);
    if(!e.recurring){clubMap[c].indexSum+=Number(e.index_score||1);clubMap[c].indexCount++;}
  });
  Object.values(clubMap).forEach(c=>{c.indexAvg=c.indexCount>0?c.indexSum/c.indexCount:0;});
  const clubRows=Object.values(clubMap).sort((a,b)=>metric==="value"?b.value-a.value:b.indexAvg-a.indexAvg);
  const displayClubs=showAllClubs?clubRows:clubRows.slice(0,10);
  const maxClub=clubRows.length?Math.max(...clubRows.map(c=>metric==="value"?(c.value||0):(c.indexAvg||0)),1):1;

  if(!log||!log.length) return(
    <div style={{textAlign:"center",padding:"80px 0",fontFamily:"'DM Sans',sans-serif",color:"#9CA3AF"}}>
      No entries yet — go to a department tab and click LOG IT to get started.
    </div>
  );

  const Card=({children,title,action})=>(
    <div style={{background:"#fff",border:"1px solid #E5E7EB",borderRadius:16,padding:"24px 28px",boxShadow:"0 1px 6px rgba(0,0,0,.05)"}}>
      {(title||action)&&(
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20,flexWrap:"wrap",gap:12}}>
          {title&&<div style={{fontFamily:"'DM Serif Display',serif",fontSize:18,color:"#111827"}}>{title}</div>}
          {action}
        </div>
      )}
      {children}
    </div>
  );

  return(
    <div style={{display:"flex",flexDirection:"column",gap:20}}>

      {/* Filter Bar */}
      <div style={{background:"#011e5c",borderRadius:16,padding:"16px 24px",boxShadow:"0 4px 24px rgba(0,0,0,.15)",display:"flex",flexDirection:"column",gap:12}}>

        {/* Row 1: League */}
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          <span style={{fontSize:9,fontWeight:700,letterSpacing:1.2,color:"#4a6fa8",width:52,flexShrink:0}}>LEAGUE</span>
          {LEAGUES.map(l=>(
            <button key={l} onClick={()=>setLeague(l)} style={{padding:"6px 14px",border:`1.5px solid ${league===l?(LEAGUE_COLORS[l]||"#fff"):"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:league===l?700:400,fontSize:12,cursor:"pointer",background:league===l?"#fff":"transparent",color:league===l?"#011e5c":"#64748B",transition:"all .15s",whiteSpace:"nowrap"}}>
              {l==="all"?"All Leagues":l}
            </button>
          ))}
        </div>

        <div style={{height:1,background:"#0a2d6e"}}/>

        {/* Row 2: Period */}
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          <span style={{fontSize:9,fontWeight:700,letterSpacing:1.2,color:"#4a6fa8",width:52,flexShrink:0}}>PERIOD</span>
          <button onClick={()=>{setFilterYear("all");setFilterMonth("all");}} style={{padding:"6px 12px",border:`1.5px solid ${filterYear==="all"?"#f51200":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:filterYear==="all"?700:400,fontSize:12,cursor:"pointer",background:filterYear==="all"?"#f51200":"transparent",color:filterYear==="all"?"#fff":"#64748B",transition:"all .15s"}}>All Time</button>
          {availableYears.map(y=>(
            <button key={y} onClick={()=>{setFilterYear(String(y));setFilterMonth("all");}} style={{padding:"6px 12px",border:`1.5px solid ${filterYear===String(y)?"#f51200":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:filterYear===String(y)?700:400,fontSize:12,cursor:"pointer",background:filterYear===String(y)?"#f51200":"transparent",color:filterYear===String(y)?"#fff":"#64748B",transition:"all .15s"}}>{y}</button>
          ))}
          {filterYear!=="all"&&(
            <>
              <div style={{width:1,height:16,background:"#0a2d6e",margin:"0 4px"}}/>
              <button onClick={()=>setFilterMonth("all")} style={{padding:"5px 10px",border:`1.5px solid ${filterMonth==="all"?"#fff":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:filterMonth==="all"?700:400,fontSize:11,cursor:"pointer",background:filterMonth==="all"?"#fff":"transparent",color:filterMonth==="all"?"#011e5c":"#64748B",transition:"all .15s"}}>All</button>
              {MONTHS.map((m,i)=>(
                <button key={m} onClick={()=>setFilterMonth(String(i))} style={{padding:"5px 10px",border:`1.5px solid ${filterMonth===String(i)?"#fff":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:filterMonth===String(i)?700:400,fontSize:11,cursor:"pointer",background:filterMonth===String(i)?"#fff":"transparent",color:filterMonth===String(i)?"#011e5c":"#64748B",transition:"all .15s"}}>{m}</button>
              ))}
            </>
          )}
        </div>

        <div style={{height:1,background:"#0a2d6e"}}/>

        {/* Row 3: Type + Metric + count + export */}
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          <span style={{fontSize:9,fontWeight:700,letterSpacing:1.2,color:"#4a6fa8",width:52,flexShrink:0}}>TYPE</span>
          {[["all","All"],["strategic","Strategic"],["recurring","Recurring"]].map(([v,l])=>(
            <button key={v} onClick={()=>setEntryType(v)} style={{padding:"6px 14px",border:`1.5px solid ${entryType===v?"#f51200":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:entryType===v?700:400,fontSize:12,cursor:"pointer",background:entryType===v?"#f51200":"transparent",color:entryType===v?"#fff":"#64748B",transition:"all .15s"}}>
              {l}
            </button>
          ))}
          <div style={{width:1,height:16,background:"#0a2d6e",margin:"0 6px"}}/>
          <span style={{fontSize:9,fontWeight:700,letterSpacing:1.2,color:"#4a6fa8",flexShrink:0}}>METRIC</span>
          {[["index","CPG Index"],["value","Value Delivered"]].map(([v,l])=>(
            <button key={v} onClick={()=>setMetric(v)} style={{padding:"6px 14px",border:`1.5px solid ${metric===v?"#f51200":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:metric===v?700:400,fontSize:12,cursor:"pointer",background:metric===v?"#f51200":"transparent",color:metric===v?"#fff":"#64748B",transition:"all .15s"}}>
              {l}
            </button>
          ))}
          <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:10}}>
            <span style={{fontSize:12,color:"#475569",whiteSpace:"nowrap"}}><span style={{color:"#fff",fontWeight:700}}>{total.toLocaleString()}</span> entries</span>
            <button onClick={onExport} style={{background:"#0a2d6e",color:"#64748B",border:"1px solid #1e3a7a",borderRadius:8,padding:"6px 14px",fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:12,cursor:"pointer",whiteSpace:"nowrap"}}>⬇ Export</button>
          </div>
        </div>
      </div>

      {/* KPI Cards */}
      {(()=>{
        const totalValue=typeFiltered.reduce((s,e)=>s+getRackRate(e),0);
        const stratValue=strategicLog.reduce((s,e)=>s+getRackRate(e),0);
        const recurValue=recurringLog.reduce((s,e)=>s+getRackRate(e),0);
        const extValue=typeFiltered.filter(e=>e.type==="External").reduce((s,e)=>s+getRackRate(e),0);
        const isValue=metric==="value";
        const kpis=isValue?[
          {label:"TOTAL VALUE DELIVERED",value:"$"+Math.round(totalValue).toLocaleString(),color:"#111827",sub:`${total} deliverables`},
          {label:"STRATEGIC VALUE",value:"$"+Math.round(stratValue).toLocaleString(),color:"#4338CA",sub:`${strategic} deliverables`},
          {label:"RECURRING VALUE",value:"$"+Math.round(recurValue).toLocaleString(),color:"#854D0E",sub:`${recurring} deliverables`},
          {label:"EXT VALUE DELIVERED",value:"$"+Math.round(extValue).toLocaleString(),color:"#0369A1",sub:`${extCnt} external entries`},
          {label:"CLUBS ENGAGED",value:Object.keys(clubMap).length.toLocaleString(),color:"#047857",sub:league==="all"?"all leagues":league},
        ]:[
          {label:"TOTAL DELIVERABLES",value:total.toLocaleString(),color:"#111827",sub:`${extCnt} external · ${intCnt} internal`},
          {label:"STRATEGIC",value:strategic.toLocaleString(),color:"#4338CA",sub:`${total?Math.round(strategic/total*100):0}% of total`},
          {label:"RECURRING",value:recurring.toLocaleString(),color:"#854D0E",sub:`${total?Math.round(recurring/total*100):0}% of total`},
          {label:"AVG CPG INDEX",value:avgIndex,color:"#047857",sub:"strategic external"},
          {label:"CLUBS ENGAGED",value:Object.keys(clubMap).length.toLocaleString(),color:"#0369A1",sub:league==="all"?"all leagues":league},
        ];
        return(
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12}}>
            {kpis.map(({label,value,color,sub})=>(
              <div key={label} style={{background:"#fff",borderRadius:12,padding:"16px 20px",border:"1px solid #E5E7EB",boxShadow:"0 1px 4px rgba(0,0,0,.04)"}}>
                <div style={{fontSize:10,fontWeight:700,letterSpacing:1,color:"#9CA3AF",marginBottom:6}}>{label}</div>
                <div style={{fontSize:26,fontWeight:800,color,fontFamily:"'DM Serif Display',serif",lineHeight:1}}>{value}</div>
                <div style={{fontSize:11,color:"#6B7280",marginTop:4}}>{sub}</div>
              </div>
            ))}
          </div>
        );
      })()}

      {/* Row 1: Bar + Pie */}
      <div style={{display:"grid",gridTemplateColumns:"1.4fr 1fr",gap:16}}>
        <Card title="Total Deliverables by League">
          {byLeague.length>0?<HBarChart data={byLeague} color="#1D4ED8" height={52}/>:
            <div style={{textAlign:"center",padding:"32px 0",color:"#9CA3AF",fontSize:13}}>No data for this filter.</div>}
        </Card>
        <Card title="CPG Usage by League">
          {usageByLeague.length>0?<DonutChart slices={usageByLeague} size={200} title={`${total}`} subtitle="total"/>:
            <div style={{textAlign:"center",padding:"32px 0",color:"#9CA3AF",fontSize:13}}>No data.</div>}
        </Card>
      </div>

      {/* Row 2: Vertical pie */}
      <Card title={`CPG Engagement by Vertical${league!=="all"?` — ${league}`:""}`}>
        {byDept.length>0?<DonutChart slices={byDept} size={220} title={`${leagueFiltered.length}`} subtitle="deliverables"/>:
          <div style={{textAlign:"center",padding:"32px 0",color:"#9CA3AF",fontSize:13}}>No data for this filter.</div>}
      </Card>

      {/* Row 3: Top Clubs */}
      <Card
        title={`Top Clubs — ${metric==="value"?"Value Delivered":"CPG Index Score"}`}
        action={metric==="value"&&<span style={{fontSize:11,color:"#F59E0B",background:"#FEF9C3",border:"1px solid #FDE68A",borderRadius:6,padding:"2px 8px",fontWeight:600}}>Conservative est. rates</span>}
      >
        {clubRows.length===0?(
          <div style={{textAlign:"center",padding:"32px 0",color:"#9CA3AF",fontSize:13}}>No external entries for this filter.</div>
        ):(
          <>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {displayClubs.map((c,i)=>{
                const ls=LEAGUE_STYLES[c.league]||LEAGUE_STYLES["League-wide"];
                const metricVal=metric==="value"?c.value:c.indexAvg;
                const metricLabel=metric==="value"?"$"+Math.round(c.value).toLocaleString():c.indexAvg.toFixed(2);
                const pct=maxClub>0?metricVal/maxClub*100:0;
                return(
                  <div key={c.club} style={{display:"flex",alignItems:"center",gap:10}}>
                    <span style={{width:24,textAlign:"right",fontSize:12,fontWeight:700,color:"#9CA3AF",flexShrink:0}}>#{i+1}</span>
                    <div style={{width:190,flexShrink:0}}>
                      <div style={{fontSize:13,fontWeight:700,color:"#111827",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{c.club}</div>
                      <LeagueBadge league={c.league||"Unknown"}/>
                    </div>
                    <div style={{flex:1,height:28,background:"#F3F4F6",borderRadius:6,overflow:"hidden"}}>
                      <div style={{width:`${pct}%`,height:"100%",background:ls.color,borderRadius:6,opacity:.85,transition:"width .4s"}}/>
                    </div>
                    <div style={{width:110,textAlign:"right",flexShrink:0}}>
                      <div style={{fontSize:15,fontWeight:800,color:ls.color,fontFamily:"'DM Serif Display',serif"}}>{metricLabel}</div>
                      <div style={{fontSize:10,color:"#9CA3AF"}}>{c.count} logged{metric==="value"?` · idx ${c.indexAvg.toFixed(1)}`:""}</div>
                    </div>
                  </div>
                );
              })}
            </div>
            {clubRows.length>10&&(
              <button onClick={()=>setShowAllClubs(!showAllClubs)} style={{marginTop:16,width:"100%",background:"#F8FAFC",border:"1px solid #E5E7EB",borderRadius:8,padding:"9px",fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:13,color:"#374151",cursor:"pointer"}}>
                {showAllClubs?`Show Top 10 ▲`:`Show All ${clubRows.length} Clubs ▼`}
              </button>
            )}
          </>
        )}
      </Card>

      {/* Row 4: Workload */}
      <Card title="Internal vs. External Workload by Department">
        {(()=>{
          const deptIntExt=ALL_DEPT_NAMES.map(dept=>{
            const en=typeFiltered.filter(e=>e.dept===dept);
            return{dept,ext:en.filter(e=>e.type==="External").length,int:en.filter(e=>e.type==="Internal").length,total:en.length};
          }).filter(d=>d.total>0&&DEPT_CONFIG[d.dept]);
          if(!deptIntExt.length) return <div style={{textAlign:"center",padding:"24px 0",color:"#9CA3AF",fontSize:13}}>No data.</div>;
          return(
            <div>
              <div style={{marginBottom:20,padding:"12px 16px",background:"#F8FAFC",borderRadius:10}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
                  <span style={{fontSize:12,color:"#6B7280",fontFamily:"'DM Sans',sans-serif"}}>Overall</span>
                  <span style={{fontSize:12,fontFamily:"'DM Sans',sans-serif"}}>
                    <span style={{color:"#0369A1",fontWeight:700}}>{extCnt} ext</span>
                    <span style={{color:"#9CA3AF",margin:"0 6px"}}>·</span>
                    <span style={{color:"#047857",fontWeight:700}}>{intCnt} int</span>
                  </span>
                </div>
                <div style={{display:"flex",height:10,borderRadius:99,overflow:"hidden",background:"#E5E7EB"}}>
                  {extCnt>0&&<div style={{width:`${total?(extCnt/total)*100:0}%`,background:"#0369A1",transition:"width .3s"}}/>}
                  {intCnt>0&&<div style={{width:`${total?(intCnt/total)*100:0}%`,background:"#047857",transition:"width .3s"}}/>}
                </div>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:10}}>
                {deptIntExt.map(({dept,ext,int:intN,total:tot})=>{
                  const cfg=DEPT_CONFIG[dept];
                  if(!cfg) return null;
                  return(
                    <div key={dept} style={{display:"flex",alignItems:"center",gap:12}}>
                      <span style={{width:160,flexShrink:0,fontSize:12,fontWeight:600,color:"#374151",fontFamily:"'DM Sans',sans-serif"}}>{cfg.emoji} {dept}</span>
                      <div style={{flex:1,height:8,borderRadius:99,overflow:"hidden",background:"#F3F4F6",display:"flex"}}>
                        {ext>0&&<div style={{width:`${(ext/tot)*100}%`,background:"#0369A1"}}/>}
                        {intN>0&&<div style={{width:`${(intN/tot)*100}%`,background:"#047857"}}/>}
                      </div>
                      <span style={{width:80,flexShrink:0,fontSize:11,color:"#9CA3AF",fontFamily:"'DM Sans',sans-serif",textAlign:"right"}}>{ext} ext · {intN} int</span>
                    </div>
                  );
                })}
              </div>
              <div style={{display:"flex",gap:16,marginTop:14,paddingTop:12,borderTop:"1px solid #F3F4F6"}}>
                <div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:10,height:10,borderRadius:"50%",background:"#0369A1"}}/><span style={{fontSize:11,color:"#6B7280"}}>External — club-facing</span></div>
                <div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:10,height:10,borderRadius:"50%",background:"#047857"}}/><span style={{fontSize:11,color:"#6B7280"}}>Internal — USL-facing</span></div>
              </div>
            </div>
          );
        })()}
      </Card>

    </div>
  );
}


function ActivityExplorer({ log, onRemove, onExportView }) {
  const [view,setView]=useState("club");
  const [expandedClub, setExpandedClub]=useState(null);
  const total=log.reduce((s,e)=>s+e.rate,0);

  // By Club
  const clubMap={};
  log.filter(e=>e.type==="External").forEach(e=>{
    const c=e.club||"Unknown";
    if(!clubMap[c])clubMap[c]={club:c,league:e.league||"",count:0,value:0,indexSum:0,indexCount:0,entries:[]};
    clubMap[c].count++;clubMap[c].value+=e.rate;
    clubMap[c].entries.push(e);
    if(!e.recurring&&e.index_score){clubMap[c].indexSum+=Number(e.index_score);clubMap[c].indexCount++;}
  });
  Object.values(clubMap).forEach(c=>{c.indexAvg=c.indexCount>0?(c.indexSum/c.indexCount).toFixed(1):"—";});
  const clubRows=Object.values(clubMap).sort((a,b)=>b.value-a.value);

  // By League
  const leagueMap={};
  log.filter(e=>e.type==="External").forEach(e=>{
    const l=e.league||"Unknown";
    if(!leagueMap[l])leagueMap[l]={league:l,count:0,value:0,indexSum:0,indexCount:0,clubs:new Set()};
    leagueMap[l].count++;leagueMap[l].value+=e.rate;leagueMap[l].clubs.add(e.club);
    if(!e.recurring&&e.index_score){leagueMap[l].indexSum+=Number(e.index_score);leagueMap[l].indexCount++;}
  });
  Object.values(leagueMap).forEach(l=>{l.indexAvg=l.indexCount>0?(l.indexSum/l.indexCount).toFixed(1):"—";l.clubCount=l.clubs.size;});
  const leagueRows=Object.values(leagueMap).sort((a,b)=>b.count-a.count);

  // By Dept
  const deptMap={};
  log.forEach(e=>{
    if(!deptMap[e.dept])deptMap[e.dept]={dept:e.dept,count:0,value:0,ext:0,int:0};
    deptMap[e.dept].count++;deptMap[e.dept].value+=e.rate;
    if(e.type==="External")deptMap[e.dept].ext++;else deptMap[e.dept].int++;
  });
  const deptRows=Object.values(deptMap).sort((a,b)=>b.value-a.value);

  // By Staff
  const staffMap={};
  log.forEach(e=>{
    if(!staffMap[e.staff])staffMap[e.staff]={staff:e.staff,count:0,value:0,ext:0,int:0,indexSum:0,indexCount:0};
    staffMap[e.staff].count++;staffMap[e.staff].value+=e.rate;
    if(e.type==="External")staffMap[e.staff].ext++;else staffMap[e.staff].int++;
    if(!e.recurring&&e.index_score&&e.type==="External"){staffMap[e.staff].indexSum+=Number(e.index_score);staffMap[e.staff].indexCount++;}
  });
  Object.values(staffMap).forEach(s=>{s.indexAvg=s.indexCount>0?(s.indexSum/s.indexCount).toFixed(1):"—";});
  const staffRows=Object.values(staffMap).sort((a,b)=>b.value-a.value);

  const tBtn=active=>({padding:"8px 18px",border:"none",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer",transition:"all .15s",background:active?"#011e5c":"transparent",color:active?"#fff":"#6B7280"});

  const [selected,setSelected]=useState(new Set());
  const [confirm,setConfirm]=useState(false);
  const reversed=[...log].reverse();
  const allSelected=selected.size===log.length&&log.length>0;
  const toggle=id=>{setSelected(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n;});setConfirm(false);};
  const toggleAll=()=>{setSelected(allSelected?new Set():new Set(log.map(e=>e.id)));setConfirm(false);};
  const handleDelete=()=>{onRemove([...selected]);setSelected(new Set());setConfirm(false);};
  const hasSelection=selected.size>0;

  const Checkbox=({checked,indeterminate,onClick})=>(
    <div onClick={onClick} style={{width:18,height:18,borderRadius:5,border:`2px solid ${checked||indeterminate?"#EF4444":"#D1D5DB"}`,background:checked?"#EF4444":indeterminate?"#FEE2E2":"#fff",display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",flexShrink:0,transition:"all .15s"}}>
      {checked&&<span style={{color:"#fff",fontSize:11,fontWeight:900,lineHeight:1}}>✓</span>}
      {indeterminate&&!checked&&<span style={{color:"#EF4444",fontSize:11,fontWeight:900,lineHeight:1}}>—</span>}
    </div>
  );

  if(!log.length) return <div style={{textAlign:"center",padding:"60px 0",color:"#9CA3AF",fontFamily:"'DM Sans',sans-serif"}}>No entries logged yet.</div>;

  return(
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:24,flexWrap:"wrap",gap:12}}>
        <div style={{display:"flex",background:"#F1F5F9",borderRadius:12,padding:4,gap:2}}>
          {[["club","By Club"],["league","By League"],["dept","By Department"],["staff","By Staff"],["log","Full Log"]].map(([v,label])=>(
            <button key={v} onClick={()=>{setView(v);setExpandedClub(null);}} style={tBtn(view===v)}>{label}</button>
          ))}
        </div>
        <button onClick={()=>onExportView(view,{clubRows,deptRows,staffRows,log})} style={{background:"#0369A1",color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>⬇ Export This View</button>
      </div>

      {view==="club"&&(
        <TblWrap>
          <thead><tr><TH>Club</TH><TH>League</TH><TH right># Logged</TH><TH right>Avg Index</TH><TH right>Total Value</TH><TH right>% of Total</TH></tr></thead>
          <tbody>
            {clubRows.map(c=>{
              const pct=total?c.value/total:0;
              const isExpanded=expandedClub===c.club;
              const delivMap={};
              c.entries.forEach(e=>{if(!delivMap[e.name])delivMap[e.name]={name:e.name,cat:e.cat,count:0,rate:e.rate};delivMap[e.name].count++;});
              const delivList=Object.values(delivMap).sort((a,b)=>b.count-a.count);
              return(
                <Fragment key={c.club}>
                  <tr onClick={()=>setExpandedClub(isExpanded?null:c.club)} style={{cursor:"pointer",background:isExpanded?"#F5F3FF":undefined}}>
                    <TD bold color="#111827">
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontSize:10,color:"#9CA3AF",transition:"transform .15s",display:"inline-block",transform:isExpanded?"rotate(90deg)":"rotate(0deg)"}}>▶</span>
                        {c.club}
                      </div>
                    </TD>
                    <TD><LeagueBadge league={c.league||"League-wide"}/></TD>
                    <TD right>{c.count}</TD>
                    <TD right><span style={{background:"#EEF2FF",color:"#4338CA",borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700}}>{c.indexAvg}</span></TD>
                    <TD right bold>{fmt$(c.value)}</TD>
                    <TD right><div style={{display:"flex",alignItems:"center",justifyContent:"flex-end",gap:8}}><div style={{width:60,height:6,background:"#F3F4F6",borderRadius:99,overflow:"hidden"}}><div style={{width:`${pct*100}%`,height:"100%",background:"#6366F1",borderRadius:99}}/></div>{(pct*100).toFixed(1)}%</div></TD>
                  </tr>
                  {isExpanded&&(
                    <tr key={`${c.club}-expand`}>
                      <td colSpan={6} style={{padding:"0 0 4px 0",background:"#F8F5FF"}}>
                        <div style={{padding:"12px 20px 12px 36px"}}>
                          <div style={{fontSize:11,fontWeight:700,color:"#7C3AED",letterSpacing:.5,marginBottom:8}}>DELIVERABLES LOGGED FOR {c.club.toUpperCase()}</div>
                          <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                            {delivList.map(d=>(
                              <div key={d.name} style={{background:"#fff",border:"1px solid #DDD6FE",borderRadius:8,padding:"6px 12px",fontSize:12,color:"#374151"}}>
                                <span style={{fontWeight:600}}>{d.name}</span>
                                <span style={{color:"#9CA3AF",marginLeft:6}}>×{d.count}</span>
                                {d.rate>0&&<span style={{color:"#7C3AED",marginLeft:6,fontSize:11}}>{fmt$(d.rate)}</span>}
                              </div>
                            ))}
                          </div>
                        </div>
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
          </tbody>
        </TblWrap>
      )}

      {view==="league"&&(
        <TblWrap>
          <thead><tr><TH>League</TH><TH right>Clubs Engaged</TH><TH right># Logged</TH><TH right>Avg Index</TH><TH right>Total Value</TH><TH right>% of Total</TH></tr></thead>
          <tbody>
            {leagueRows.map(l=>{
              const pct=total?l.value/total:0;
              const ls=LEAGUE_STYLES[l.league]||LEAGUE_STYLES["League-wide"];
              return(
                <tr key={l.league}>
                  <TD bold><LeagueBadge league={l.league}/></TD>
                  <TD right color="#6B7280">{l.clubCount} clubs</TD>
                  <TD right>{l.count}</TD>
                  <TD right><span style={{background:"#EEF2FF",color:"#4338CA",borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700}}>{l.indexAvg}</span></TD>
                  <TD right bold color={ls.color}>{fmt$(l.value)}</TD>
                  <TD right><div style={{display:"flex",alignItems:"center",justifyContent:"flex-end",gap:8}}><div style={{width:60,height:6,background:"#F3F4F6",borderRadius:99,overflow:"hidden"}}><div style={{width:`${pct*100}%`,height:"100%",background:ls.color,borderRadius:99}}/></div>{(pct*100).toFixed(1)}%</div></TD>
                </tr>
              );
            })}
          </tbody>
        </TblWrap>
      )}

      {view==="dept"&&(
        <TblWrap>
          <thead><tr><TH>Department</TH><TH right># Logged</TH><TH right>Total Value</TH><TH right>External</TH><TH right>Internal</TH><TH right>% of Total</TH></tr></thead>
          <tbody>
            {deptRows.map(d=>{
              const cfg=DEPT_CONFIG[d.dept],pct=total?d.value/total:0;
              return(<tr key={d.dept}><TD bold><span style={{marginRight:6}}>{cfg?.emoji}</span>{d.dept}</TD><TD right>{d.count}</TD><TD right bold color={cfg?.color||"#374151"}>{fmt$(d.value)}</TD><TD right color="#0369A1">{d.ext}</TD><TD right color="#047857">{d.int}</TD><TD right><div style={{display:"flex",alignItems:"center",justifyContent:"flex-end",gap:8}}><div style={{width:60,height:6,background:"#F3F4F6",borderRadius:99,overflow:"hidden"}}><div style={{width:`${pct*100}%`,height:"100%",background:cfg?.color||"#6366F1",borderRadius:99}}/></div>{(pct*100).toFixed(1)}%</div></TD></tr>);
            })}
          </tbody>
        </TblWrap>
      )}

      {view==="staff"&&(
        <TblWrap>
          <thead><tr><TH>Staff Member</TH><TH right># Logged</TH><TH right>Avg Index</TH><TH right>Total Value</TH><TH right>External</TH><TH right>Internal</TH><TH right>% of Total</TH></tr></thead>
          <tbody>
            {staffRows.map(s=>{
              const pct=total?s.value/total:0;
              const initials=s.staff.split(" ").map(n=>n[0]).join("").slice(0,2);
              return(<tr key={s.staff}>
                <TD bold><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{width:28,height:28,borderRadius:"50%",background:"#EEF2FF",color:"#4338CA",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,flexShrink:0}}>{initials}</div>{s.staff}</div></TD>
                <TD right>{s.count}</TD>
                <TD right><span style={{background:"#EEF2FF",color:"#4338CA",borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700}}>{s.indexAvg}</span></TD>
                <TD right bold>{fmt$(s.value)}</TD>
                <TD right color="#0369A1">{s.ext}</TD>
                <TD right color="#047857">{s.int}</TD>
                <TD right><div style={{display:"flex",alignItems:"center",justifyContent:"flex-end",gap:8}}><div style={{width:60,height:6,background:"#F3F4F6",borderRadius:99,overflow:"hidden"}}><div style={{width:`${pct*100}%`,height:"100%",background:"#4338CA",borderRadius:99}}/></div>{(pct*100).toFixed(1)}%</div></TD>
              </tr>);
            })}
          </tbody>
        </TblWrap>
      )}

      {view==="log"&&(
        <div>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:10,fontFamily:"'DM Sans',sans-serif"}}>
            <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:13,color:"#374151",fontWeight:500,userSelect:"none"}}>
              <Checkbox checked={allSelected} indeterminate={selected.size>0&&!allSelected} onClick={toggleAll}/>
              {selected.size===0?"Select entries to delete":`${selected.size} of ${log.length} selected`}
            </label>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              {hasSelection&&!confirm&&<span style={{fontSize:12,color:"#6B7280"}}>{selected.size} {selected.size===1?"entry":"entries"} · {fmt$(log.filter(e=>selected.has(e.id)).reduce((s,e)=>s+e.rate,0))} value</span>}
              {!confirm?(
                <button onClick={()=>hasSelection&&setConfirm(true)} style={{background:hasSelection?"#FEE2E2":"#F3F4F6",color:hasSelection?"#991B1B":"#9CA3AF",border:`1px solid ${hasSelection?"#FECACA":"#E5E7EB"}`,borderRadius:8,padding:"8px 16px",fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:13,cursor:hasSelection?"pointer":"not-allowed"}}>🗑 Delete Selected</button>
              ):(
                <div style={{display:"flex",gap:8,alignItems:"center"}}>
                  <span style={{fontSize:13,color:"#991B1B",fontWeight:600}}>Delete {selected.size} {selected.size===1?"entry":"entries"}?</span>
                  <button onClick={handleDelete} style={{background:"#EF4444",color:"#fff",border:"none",borderRadius:8,padding:"8px 14px",fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>Yes, delete</button>
                  <button onClick={()=>setConfirm(false)} style={{background:"#F3F4F6",color:"#374151",border:"none",borderRadius:8,padding:"8px 14px",fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:13,cursor:"pointer"}}>Cancel</button>
                </div>
              )}
            </div>
          </div>
          <TblWrap>
            <thead><tr>
              <th style={{padding:"11px 14px",width:40,background:"#F9FAFB",borderBottom:"1px solid #E5E7EB"}}/>
              <TH>#</TH><TH>Time</TH><TH>Staff</TH><TH>Dept</TH><TH>Deliverable</TH><TH>Type</TH><TH>Club / Recipient</TH><TH>League</TH><TH right>Rack Rate</TH><TH>Notes</TH>
            </tr></thead>
            <tbody>
              {reversed.map((e,i)=>{
                const checked=selected.has(e.id);
                return(
                  <tr key={e.id} style={{background:checked?"#FEF2F2":undefined,transition:"background .15s"}}>
                    <td style={{padding:"10px 14px",textAlign:"center"}}><Checkbox checked={checked} onClick={()=>toggle(e.id)}/></td>
                    <td style={{padding:"10px 14px",color:"#9CA3AF",fontSize:11}}>{log.length-i}</td>
                    <td style={{padding:"10px 14px",color:"#6B7280",fontSize:12}}>{new Date(e.ts).toLocaleString()}</td>
                    <TD bold color="#111827">{e.staff}</TD>
                    <TD><DeptChip dept={e.dept}/></TD>
                    <TD>
                      <div style={{display:"flex",alignItems:"center",gap:6}}>
                        {e.name}
                        {e.recurring&&<RecurringBadge/>}
                      </div>
                    </TD>
                    <TD><Badge type={e.type}/></TD>
                    <TD color="#374151">{e.club}</TD>
                    <TD><LeagueBadge league={e.league||"League-wide"}/></TD>
                    <TD right bold color={DEPT_CONFIG[e.dept]?.color||"#374151"}>{e.rate>0?fmt$(e.rate):"—"}</TD>
                    <TD color="#6B7280">{e.notes||"—"}</TD>
                  </tr>
                );
              })}
            </tbody>
          </TblWrap>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  EXCEL EXPORT
// ─────────────────────────────────────────────────────────────────────────────

function exportExcel(log) {
  const wb=XLSX.utils.book_new(),now=new Date();
  const ds=now.toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"});
  const total=log.reduce((s,e)=>s+e.rate,0);
  const strategic=log.filter(e=>!e.recurring),recurring=log.filter(e=>e.recurring);
  const extVal=log.filter(e=>e.type==="External").reduce((s,e)=>s+e.rate,0);
  const intVal=log.filter(e=>e.type==="Internal").reduce((s,e)=>s+e.rate,0);

  const dash=XLSX.utils.aoa_to_sheet([[`CPG DEPARTMENT PERFORMANCE SUMMARY — ${ds}`],[],
    ["Total Rack Rate Value","$"+total.toLocaleString()],
    ["Strategic Value","$"+strategic.reduce((s,e)=>s+e.rate,0).toLocaleString()],
    ["Recurring Value","$"+recurring.reduce((s,e)=>s+e.rate,0).toLocaleString()],
    ["External Value","$"+extVal.toLocaleString()],["Internal Value","$"+intVal.toLocaleString()],
    ["Total Deliverables",log.length],["Strategic Deliverables",strategic.length],["Recurring Deliverables",recurring.length],[],
    ["BREAKDOWN BY DEPARTMENT"],["Department","# Logged","Strategic","Recurring","Total Value","Ext","Int","% of Total"],
    ...ALL_DEPT_NAMES.map(dept=>{
      const en=log.filter(e=>e.dept===dept),v=en.reduce((s,e)=>s+e.rate,0);
      return[dept,en.length,en.filter(e=>!e.recurring).length,en.filter(e=>e.recurring).length,v,
        en.filter(e=>e.type==="External").length,en.filter(e=>e.type==="Internal").length,
        total?(v/total*100).toFixed(1)+"%":"0%"];
    })]);
  dash["!cols"]=[{wch:26},{wch:10},{wch:12},{wch:12},{wch:16},{wch:8},{wch:8},{wch:12}];
  XLSX.utils.book_append_sheet(wb,dash,"Dashboard Summary");

  const clubMap={};
  log.filter(e=>e.type==="External").forEach(e=>{
    const c=e.club||"Unknown";
    if(!clubMap[c])clubMap[c]={club:c,league:e.league||"",count:0,value:0,indexSum:0,indexCount:0};
    clubMap[c].count++;clubMap[c].value+=e.rate;
    if(!e.recurring){clubMap[c].indexSum+=Number(e.index_score||1);clubMap[c].indexCount++;}
  });
  const clubWs=XLSX.utils.aoa_to_sheet([[`BY CLUB — ${ds}`],
    ["Club","League","# Logged","Total Value","CPG Index","% of Total"],
    ...Object.values(clubMap).sort((a,b)=>b.value-a.value).map(c=>[
      c.club,c.league,c.count,c.value,
      c.indexCount>0?(c.indexSum/c.indexCount).toFixed(2):"—",
      total?(c.value/total*100).toFixed(1)+"%":"0%"])]);
  clubWs["!cols"]=[{wch:32},{wch:16},{wch:10},{wch:14},{wch:10},{wch:12}];
  XLSX.utils.book_append_sheet(wb,clubWs,"By Club");

  const dc={};
  log.forEach(e=>{const k=`${e.dept}||${e.name}`;if(!dc[k])dc[k]={dept:e.dept,name:e.name,rate:e.rate,type:e.type,cat:e.cat,recurring:e.recurring,count:0};dc[k].count++;});
  const ctr=XLSX.utils.aoa_to_sheet([["DELIVERABLE COUNTER"],
    ["Department","Deliverable","Category","Type","Recurring","# Logged","Rate","Total Value"],
    ...Object.values(dc).sort((a,b)=>b.count-a.count).map(r=>[r.dept,r.name,r.cat,r.type,r.recurring?"Yes":"No",r.count,r.rate>0?r.rate:"—",r.rate>0?r.count*r.rate:"—"])]);
  ctr["!cols"]=[{wch:22},{wch:40},{wch:22},{wch:12},{wch:10},{wch:10},{wch:12},{wch:14}];
  XLSX.utils.book_append_sheet(wb,ctr,"Deliverable Counter");

  const logWs=XLSX.utils.aoa_to_sheet([["FULL ACTIVITY LOG"],
    ["#","Timestamp","Staff","Department","Deliverable","Category","Club / Recipient","League","Type","Recurring","Rack Rate","CPG Index"],
    ...log.map((e,i)=>[i+1,new Date(e.ts).toLocaleString(),e.staff,e.dept,e.name,e.cat||"",e.club,e.league||"",e.type,e.recurring?"Yes":"No",e.rate>0?e.rate:"—",e.index_score||""])]);
  logWs["!cols"]=[{wch:5},{wch:20},{wch:18},{wch:22},{wch:40},{wch:22},{wch:32},{wch:16},{wch:12},{wch:10},{wch:12},{wch:10}];
  XLSX.utils.book_append_sheet(wb,logWs,"Full Log");

  const staffMap={};
  log.forEach(e=>{if(!staffMap[e.staff])staffMap[e.staff]={staff:e.staff,count:0,value:0,strategic:0,recurring:0};staffMap[e.staff].count++;staffMap[e.staff].value+=e.rate;if(e.recurring)staffMap[e.staff].recurring++;else staffMap[e.staff].strategic++;});
  const staffWs=XLSX.utils.aoa_to_sheet([[`BY STAFF — ${ds}`],["Staff Member","# Logged","Strategic","Recurring","Total Value"],...Object.values(staffMap).sort((a,b)=>b.value-a.value).map(s=>[s.staff,s.count,s.strategic,s.recurring,s.value])]);
  staffWs["!cols"]=[{wch:22},{wch:10},{wch:12},{wch:12},{wch:16}];
  XLSX.utils.book_append_sheet(wb,staffWs,"By Staff");

  const wbout=XLSX.write(wb,{bookType:"xlsx",type:"array"});
  const blob=new Blob([wbout],{type:"application/octet-stream"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");
  const pad=n=>String(n).padStart(2,"0");
  a.download=`cpg_deliverable_tracker_${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())}.xlsx`;
  a.href=url;document.body.appendChild(a);a.click();
  setTimeout(()=>{URL.revokeObjectURL(url);a.remove();},1000);
}

function exportView(view,data) {
  const wb=XLSX.utils.book_new(),now=new Date();
  const ds=now.toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"});
  let ws,name;
  if(view==="club"){ws=XLSX.utils.aoa_to_sheet([[`BY CLUB — ${ds}`],["Club","League","# Logged","Total Value"],...data.clubRows.map(c=>[c.club,c.league,c.count,c.value])]);name="By Club";}
  else if(view==="dept"){ws=XLSX.utils.aoa_to_sheet([[`BY DEPARTMENT — ${ds}`],["Department","# Logged","Total Value","External","Internal"],...data.deptRows.map(d=>[d.dept,d.count,d.value,d.ext,d.int])]);name="By Department";}
  else if(view==="staff"){ws=XLSX.utils.aoa_to_sheet([[`BY STAFF — ${ds}`],["Staff Member","# Logged","Total Value","External","Internal"],...data.staffRows.map(s=>[s.staff,s.count,s.value,s.ext,s.int])]);name="By Staff";}
  else{ws=XLSX.utils.aoa_to_sheet([["FULL LOG"],["#","Timestamp","Staff","Department","Deliverable","Club","Type","Recurring","Rate"],...data.log.map((e,i)=>[i+1,new Date(e.ts).toLocaleString(),e.staff,e.dept,e.name,e.club,e.type,e.recurring?"Yes":"No",e.rate>0?e.rate:"—"])]);name="Full Log";}
  ws["!cols"]=[{wch:32},{wch:16},{wch:12},{wch:16},{wch:12}];
  XLSX.utils.book_append_sheet(wb,ws,name);
  const wbout=XLSX.write(wb,{bookType:"xlsx",type:"array"});
  const blob=new Blob([wbout],{type:"application/octet-stream"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");
  const pad=n=>String(n).padStart(2,"0");
  a.download=`cpg_${view}_${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())}.xlsx`;
  a.href=url;document.body.appendChild(a);a.click();
  setTimeout(()=>{URL.revokeObjectURL(url);a.remove();},1000);
}

// ─────────────────────────────────────────────────────────────────────────────
//  ROOT APP
// ─────────────────────────────────────────────────────────────────────────────


// ─────────────────────────────────────────────────────────────────────────────
//  APP NAV — LEFT SIDEBAR
// ─────────────────────────────────────────────────────────────────────────────

const NAV_ITEMS = [
  { id:"Dashboard",          label:"Dashboard",          emoji:"",    group:"main" },
  { id:"Activity Explorer",  label:"Activity Explorer",  emoji:"",    group:"main" },
  { id:"Corp Partnerships",  label:"Corp Partnerships",  emoji:"🤝",  group:"dept" },
  { id:"Marketing",          label:"Marketing",          emoji:"📣",  group:"dept" },
  { id:"Consumer Products",  label:"Consumer Products",  emoji:"🛍️",  group:"dept" },
  { id:"Ticketing",          label:"Ticketing",          emoji:"🎟️",  group:"dept" },
  { id:"Youth / Facilities", label:"Youth / Facilities", emoji:"⚽",  group:"dept" },
  { id:"League Initiatives", label:"League Initiatives", emoji:"🏛️",  group:"dept" },
];

function AppNav({ activeTab, setActiveTab, loading }) {
  return (
    <div style={{width:220,flexShrink:0,background:"#011e5c",minHeight:"100vh",display:"flex",flexDirection:"column",position:"sticky",top:0,zIndex:100,borderRight:"1px solid #1E293B"}}>

      {/* Brand */}
      <div style={{padding:"20px 16px 16px",borderBottom:"1px solid #1E293B"}}>
        <div style={{fontFamily:"'DM Serif Display',serif",fontSize:18,color:"#fff",letterSpacing:"-.3px",marginBottom:6}}>CPG Deliverable Tracker</div>
        <div style={{display:"flex",alignItems:"center",gap:6}}>
          <div style={{width:6,height:6,borderRadius:"50%",background:loading?"#F59E0B":"#22C55E",boxShadow:loading?"none":"0 0 8px #22C55E66"}}/>
          <span style={{fontSize:11,color:"#475569",fontFamily:"'DM Sans',sans-serif"}}>{loading?"syncing…":"live"}</span>
        </div>
        <div style={{fontSize:10,color:"#1e3a7a",fontFamily:"'DM Sans',sans-serif",marginTop:4,letterSpacing:.3}}>USL Club Performance Group</div>
      </div>

      {/* Nav items */}
      <div style={{padding:"10px 8px",flex:1,display:"flex",flexDirection:"column",gap:2}}>

        {/* Main views */}
        {NAV_ITEMS.filter(i=>i.group==="main").map(item=>{
          const isActive=activeTab===item.id;
          return(
            <button key={item.id} onClick={()=>setActiveTab(item.id)}
              style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"9px 12px",border:"none",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontSize:13,fontWeight:isActive?700:400,cursor:"pointer",textAlign:"left",background:isActive?"#fff":"transparent",color:isActive?"#011e5c":"#64748B",transition:"all .15s"}}>
              <span style={{fontSize:14,opacity:isActive?1:.7}}>{item.emoji||"◈"}</span>
              {item.label}
              {isActive&&<div style={{marginLeft:"auto",width:5,height:5,borderRadius:"50%",background:"#011e5c",flexShrink:0}}/>}
            </button>
          );
        })}

        {/* Dept section */}
        <div style={{margin:"10px 0 4px",padding:"0 12px"}}>
          <span style={{fontSize:9,fontWeight:700,letterSpacing:1.2,color:"#1e3a7a",textTransform:"uppercase"}}>Departments</span>
        </div>

        {NAV_ITEMS.filter(i=>i.group==="dept").map(item=>{
          const cfg=DEPT_CONFIG[item.id];
          const isActive=activeTab===item.id;
          return(
            <button key={item.id} onClick={()=>setActiveTab(item.id)}
              style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"9px 12px",border:"none",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontSize:13,fontWeight:isActive?700:400,cursor:"pointer",textAlign:"left",background:isActive?(cfg?.color+"22"||"#ffffff22"):"transparent",color:isActive?(cfg?.color||"#A5B4FC"):"#64748B",transition:"all .15s"}}>
              <span style={{fontSize:14}}>{item.emoji}</span>
              <span style={{flex:1,lineHeight:1.2}}>{item.label}</span>
              {isActive&&<div style={{width:5,height:5,borderRadius:"50%",background:cfg?.color||"#6366F1",flexShrink:0}}/>}
            </button>
          );
        })}
      </div>

      {/* Footer */}
      <div style={{padding:"12px 16px",borderTop:"1px solid #1E293B"}}>
        <div style={{fontSize:10,color:"#0a2d6e",fontFamily:"'DM Sans',sans-serif"}}>v2.0 · {new Date().getFullYear()}</div>
      </div>
    </div>
  );
}


export default function App() {
  const [unlocked,      setUnlocked]      = useState(()=>sessionStorage.getItem("dept_unlocked")==="1");
  const [activeTab,     setActiveTab]     = useState("Dashboard");
  const [log,           setLog]           = useState([]);
  const [deptItems,     setDeptItems]     = useState(FALLBACK_DEPT_ITEMS);
  const [clubsByLeague, setClubsByLeague] = useState(FALLBACK_CLUBS);
  const [loading,       setLoading]       = useState(true);
  const [toast,         setToast]         = useState(null);

  const loadAll=useCallback(async()=>{
    try{
      const [logData,delivData,clubData]=await Promise.all([
        dbFetch(),
        dbFetchDeliverables().catch(()=>[]),
        dbFetchClubs().catch(()=>[]),
      ]);
      setLog(logData);
      if(delivData.length) {
        const built = buildDeptItems(delivData);
        // Only use sheet data if it actually produced items for known depts
        const hasItems = ALL_DEPT_NAMES.some(d => built[d]?.external?.length || built[d]?.internal?.length);
        setDeptItems(hasItems ? built : FALLBACK_DEPT_ITEMS);
      } else {
        setDeptItems(FALLBACK_DEPT_ITEMS);
      }
      if(clubData.length)  setClubsByLeague(buildClubsByLeague(clubData));
    }catch(e){console.error("Load error:",e);}
    finally{setLoading(false);}
  },[]);

  useEffect(()=>{
    if(!unlocked) return;
    loadAll();
    const interval=setInterval(()=>{
      dbFetch().then(setLog).catch(()=>{});
    },15000);
    return()=>clearInterval(interval);
  },[unlocked,loadAll]);

  const handleLog=useCallback(async entry=>{
    const newEntry={...entry,id:`${Date.now()}-${Math.random().toString(36).slice(2)}`,ts:Date.now()};
    // Strip internal flags before storing
    const {bulkSilent,...logEntry} = newEntry;
    setLog(prev=>[...prev,logEntry]);
    if(!bulkSilent) {
      setToast({msg:`✅ Logged: ${entry.name} → ${entry.club}${entry.rate>0?" · "+fmt$(entry.rate):""}`,color:DEPT_CONFIG[entry.dept]?.color||"#374151"});
    }
    try{ await dbInsert(logEntry); }
    catch(e){ setToast({msg:"⚠️ Entry may not have saved — check your Google Sheet.",color:"#F59E0B"}); }
  },[]);

  const handleBulkLog=useCallback((entries)=>{
    // Strip any internal flags, add all entries to state at once
    const clean = entries.map(({bulkSilent,...e})=>e);
    setLog(prev=>[...prev,...clean]);
  },[]);

  const handleBulkComplete=useCallback((count, name, totalVal)=>{
    setToast({msg:`✅ Logged ${count} entries — ${name}${totalVal>0?" · "+fmt$(totalVal):""}`,color:"#047857"});
  },[]);

  const handleRemove=useCallback(async ids=>{
    const idSet=new Set(ids);
    setLog(prev=>prev.filter(e=>!idSet.has(e.id)));
    dbDelete(ids).catch(()=>{loadAll();setToast({msg:"⚠️ Delete may not have synced.",color:"#F59E0B"});});
  },[loadAll]);

  if(!unlocked) return <PasswordScreen onUnlock={()=>setUnlocked(true)}/>;

  return(
    <>
      <style>{`
        
        *{box-sizing:border-box;margin:0;padding:0;}
        body{background:#F8F9FB;font-family:'DM Sans',sans-serif;}
        ::-webkit-scrollbar{width:6px;}::-webkit-scrollbar-thumb{background:#D1D5DB;border-radius:99px;}
        tr:nth-child(even) td{background:#FAFAFA;}tr:last-child td{border-bottom:none!important;}
        td{border-bottom:1px solid #F3F4F6;}select{appearance:none;}
        @keyframes spin{to{transform:rotate(360deg)}}
      `}</style>

      <div style={{display:"flex",minHeight:"100vh"}}>
        <AppNav activeTab={activeTab} setActiveTab={setActiveTab} loading={loading}/>
        <div style={{flex:1,background:"#F8F9FB",minWidth:0}}>
          {loading?(
            <div style={{textAlign:"center",padding:"80px 0",fontFamily:"'DM Sans',sans-serif",color:"#6B7280"}}>
              <div style={{width:32,height:32,border:"3px solid #E5E7EB",borderTopColor:"#6366F1",borderRadius:"50%",animation:"spin 0.8s linear infinite",margin:"0 auto 16px"}}/>
              Loading…
            </div>
          ):(
            <div style={{maxWidth:1200,padding:"28px 28px"}}>
              {activeTab==="Dashboard"         &&<Dashboard log={log} onExport={()=>exportExcel(log)} clubsByLeague={clubsByLeague}/>}
              {activeTab==="Activity Explorer" &&<ActivityExplorer log={log} onRemove={handleRemove} onExportView={exportView}/>}
              {ALL_DEPT_NAMES.includes(activeTab)&&<DeptTab dept={activeTab} log={log} onLog={handleLog} onBulkLog={handleBulkLog} onBulkComplete={handleBulkComplete} deptItems={deptItems} clubsByLeague={clubsByLeague}/>}
            </div>
          )}
        </div>
      </div>

      {toast&&<Toast msg={toast.msg} color={toast.color} onDone={()=>setToast(null)}/>}
    </>
  );
}
