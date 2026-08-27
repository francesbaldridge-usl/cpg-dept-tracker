// @ts-nocheck
import { useState, useEffect, useCallback, useRef, Fragment } from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────────────────────────────────────
//  CONFIG
// ─────────────────────────────────────────────────────────────────────────────

const APP_PASSWORD = "USLsoccer1";
const SCRIPT_URL   = "https://script.google.com/macros/s/AKfycbzOLmVlucUFx23aC-11rEioSC23FFVDnJwAc8mHColhOdHn1k68tvWJ7BRtIupX0Qi2DQ/exec";

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
  "Kendra Hodgdon","Kevin Couture","Ryan Halter","Steven Bell",
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

const LEAGUE_TIER_OPTIONS = ["Championship","League One","Premier","Super League","Expansion","USL HQ"];

const INTERNAL_RECIPIENTS = {
  "Corp Partnerships":   ["HQ Corp Partnerships","League Operations","Expansion"],
  "Marketing":           ["HQ Marketing / Comms","Corp Partnerships","Expansion","Onboarding"],
  "Consumer Products":   ["Miscellaneous","Onboarding"],
  "Ticketing":           ["Miscellaneous","League Operations","Expansion","Onboarding"],
  "League Initiatives":  ["League Operations","Expansion"],
};

// ─────────────────────────────────────────────────────────────────────────────
//  DEPT CONFIG
// ─────────────────────────────────────────────────────────────────────────────

const DEPT_CONFIG = {
  "Corp Partnerships":  { color:"#7C3AED", light:"#F5F3FF", border:"#DDD6FE", emoji:"🤝", hasToggle:true  },
  "Marketing":          { color:"#0369A1", light:"#F0F9FF", border:"#BAE6FD", emoji:"📣", hasToggle:true  },
  "Consumer Products":  { color:"#B45309", light:"#FFFBEB", border:"#FDE68A", emoji:"🛍️", hasToggle:true  },
  "Ticketing":          { color:"#047857", light:"#F0FDF4", border:"#A7F3D0", emoji:"🎟️", hasToggle:true  },
  "League Initiatives": { color:"#4338CA", light:"#EEF2FF", border:"#C7D2FE", emoji:"🏛️", hasToggle:true },
};

const ALL_DEPT_NAMES = Object.keys(DEPT_CONFIG);
const TABS = ["Dashboard", ...ALL_DEPT_NAMES, "Activity Explorer"];

// Shared cluster color palette — used everywhere a cluster gets a color (the Dashboard's
// Total Deliverables by Cluster chart, the YoY trend card's Cluster compare mode, etc.)
// so a given cluster is always the same color no matter where you're looking at it.
const CLUSTER_PALETTE = ["#1D4ED8","#047857","#7C3AED","#B45309","#DB2777","#0891B2","#4338CA","#B91C1C"];
const clusterColorFrom = (cl, sortedClusterList) => cl==="No Cluster" ? "#9CA3AF" : CLUSTER_PALETTE[sortedClusterList.indexOf(cl) % CLUSTER_PALETTE.length];
const LEAGUE_ORDER = ["Championship","League One","Premier","Super League","Expansion"];
const sortByLeagueOrder = (a,b) => { const ia=LEAGUE_ORDER.indexOf(a), ib=LEAGUE_ORDER.indexOf(b); return (ia<0?99:ia)-(ib<0?99:ib); };

// ─────────────────────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────────────────────

const fmt$ = n => "$" + Number(n).toLocaleString();

// Retroactive logging: convert a plain YYYY-MM-DD picker value into a timestamp set
// to noon local time (avoids a midnight value silently rolling to the wrong day
// depending on the browser's timezone/DST), and vice versa for the date input's default.
const todayDateString = () => {
  const d = new Date(), pad = n => String(n).padStart(2,"0");
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
};
const dateStringToTs = dateStr => {
  if (!dateStr) return Date.now();
  const [y,m,d] = dateStr.split("-").map(Number);
  return new Date(y, m-1, d, 12, 0, 0).getTime();
};

function normalizeDept(dept) {
  const MAP = {
    "Corporate Partnerships":"Corp Partnerships","Corp. Partnerships":"Corp Partnerships",
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

function buildClubClusters(clubsData) {
  const result = {};
  clubsData.forEach(c => {
    const cluster = String(c.cluster||"").trim();
    if (c.club && cluster) result[c.club] = cluster;
  });
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
  "League Initiatives": {
    external:[
      {name:"Cluster Calls",rate:150,cat:"Calls & Consultancy",index_score:2,recurring:true,examples:["Cluster check-in call","Cluster-wide update call"]},
      {name:"Cluster Reports",rate:500,cat:"Partnerships Analysis",index_score:3,recurring:true,examples:["Quarterly cluster report","Cluster performance summary"]},
      {name:"Year End Recap Report",rate:1500,cat:"Partnerships Analysis",index_score:4,recurring:false,examples:["Annual league recap","Season-end summary report"]},
      {name:"Broadcast Reports",rate:500,cat:"Partnerships Analysis",index_score:3,recurring:true,examples:["Monthly broadcast metrics","Broadcast performance recap"]},
    ],
    internal:[
      {name:"Leadership Briefing",rate:500,cat:"Strategy",index_score:4,recurring:false,examples:["Executive briefing","Board presentation"]},
      {name:"Strategic Planning Session",rate:750,cat:"Strategy",index_score:4,recurring:false,examples:["League strategy session","Annual planning meeting"]},
      {name:"Institutional Advisory",rate:75,cat:"Advisory",index_score:1,recurring:true,examples:["Leadership advisory","Governance guidance"]},
      {name:"Cross-Department Initiative",rate:500,cat:"Strategy",index_score:4,recurring:false,examples:["Cross-functional project","League-wide initiative"]},
    ],
  },
};

// Normalize the flat fallback items into the same {name,cat,leagueSelect,subcategories:[...]}
// shape buildDeptItems() produces, so DeptTab/CategoryGroupedCards never have to branch on source.
// League Initiatives external deliverables attributed to whole leagues (or USL HQ),
// not individual clubs — same name-match approach as the Social Media Report leagueSelect flag.
const MULTI_LEAGUE_ITEMS = ["Cluster Calls","Cluster Reports","Year End Recap Report","Broadcast Reports"];
const isMultiLeagueItem = (dept,name) => dept==="League Initiatives" && MULTI_LEAGUE_ITEMS.includes(name);

(function normalizeFallback(){
  Object.entries(FALLBACK_DEPT_ITEMS).forEach(([dept,modes])=>{
    ["external","internal"].forEach(mode=>{
      if (!modes[mode]) return;
      modes[mode] = modes[mode].map(item=>({
        name: item.name,
        cat: item.cat,
        leagueSelect: !!item.leagueSelect,
        multiLeagueSelect: isMultiLeagueItem(dept,item.name),
        subcategories: [{
          subcat: item.name,
          rate: item.rate,
          index_score: item.index_score,
          recurring: item.recurring,
          examples: item.examples||[],
        }],
      }));
    });
  });
})();

function buildDeptItems(deliverables) {
  const map = {};
  deliverables.forEach(d => {
    const dept = normalizeDept(d.dept);
    if (!dept || !DEPT_CONFIG[dept]) return;
    const name = String(d.name||"").trim();
    if (!name) return;
    if (!map[dept]) map[dept] = { external:{}, internal:{} };
    const isRecurring = String(d.recurring||"").toLowerCase() === "true";
    const cat = String(d.cat||d.category||"Other").trim();
    const leagueSelect = name==="Social Media Report" && dept==="Marketing";
    const multiLeagueSelect = isMultiLeagueItem(dept,name);
    const subEntry = {
      subcat:      String(d.subcat||"").trim() || name,
      rate:        Number(d.rate)||0,
      index_score: Number(d.index_score)||1,
      recurring:   isRecurring,
      examples:    d.examples ? String(d.examples).split(";").map(e=>e.trim()).filter(Boolean) : [],
    };
    const t = String(d.type||d["type (internal or external)"]||"").toLowerCase().trim();
    const isBoth = t.includes("external") && t.includes("internal");
    const modes = isBoth ? ["external","internal"] : (t==="internal" ? ["internal"] : ["external"]);
    modes.forEach(mode=>{
      if (!map[dept][mode][name]) map[dept][mode][name] = { name, cat, leagueSelect, multiLeagueSelect, subcategories:[] };
      map[dept][mode][name].subcategories.push({...subEntry});
    });
  });
  const result = {};
  Object.entries(map).forEach(([dept,modes])=>{
    result[dept] = { external: Object.values(modes.external), internal: Object.values(modes.internal) };
  });
  return result;
}

function leagueForClub(club, cbl) {
  const src = cbl || FALLBACK_CLUBS;
  for (const [league, clubs] of Object.entries(src))
    if (clubs.includes(club)) return league;
  return "League-wide";
}

const LEAGUE_STYLES = {
  "Championship":  { color:"#b28350", bg:"#FAF3EC", border:"#E3CBAE" },
  "League One":    { color:"#00becc", bg:"#E6FBFC", border:"#99EEF2" },
  "Super League":  { color:"#ff8533", bg:"#FFF1E6", border:"#FFCDA3" },
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

// ─────────────────────────────────────────────────────────────────────────────
//  LINE ITEM TABLE — shared drill-down used by Club / League / Dept / Staff views
// ─────────────────────────────────────────────────────────────────────────────

function LineItemTable({ entries, hide=[], clubClusters={} }) {
  const ALL_COLS = [
    { key:"date",  label:"Date",              render:e=>new Date(e.ts).toLocaleDateString() },
    { key:"staff", label:"Staff",             render:e=>e.staff },
    { key:"dept",  label:"Dept",              render:e=><DeptChip dept={e.dept}/> },
    { key:"name",  label:"Deliverable",       render:e=><div><div style={{display:"flex",alignItems:"center",gap:6}}>{e.name}{e.recurring&&<RecurringBadge/>}</div>{e.subcat&&e.subcat!==e.name&&<div style={{fontSize:11,color:"#9CA3AF",marginTop:2}}>{e.subcat}</div>}</div> },
    { key:"type",  label:"Type",              render:e=><Badge type={e.type}/> },
    { key:"club",  label:"Club / Recipient",  render:e=>e.club },
    { key:"cluster",label:"Cluster",          render:e=>clubClusters[e.club]||"—" },
    { key:"league",label:"League",            render:e=><LeagueBadge league={e.league||"League-wide"}/> },
    { key:"rate",  label:"Rate",   right:true,render:e=>e.rate>0?fmt$(e.rate):"—" },
    { key:"index", label:"Index",  right:true,render:e=>e.index_score>0?<span style={{background:"#EEF2FF",color:"#4338CA",borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>{e.index_score}</span>:"—" },
    { key:"notes", label:"Notes",             render:e=>e.notes||"—" },
  ];
  const cols = ALL_COLS.filter(c=>!hide.includes(c.key));
  const sorted = [...entries].sort((a,b)=>b.ts-a.ts);
  const thStyle = right => ({padding:"8px 12px",textAlign:right?"right":"left",fontWeight:700,fontSize:10,letterSpacing:.5,color:"#7C3AED",background:"#F5F3FF",borderBottom:"1px solid #DDD6FE"});
  const tdStyle = (key,right) => ({padding:"8px 12px",textAlign:right?"right":"left",whiteSpace:(key==="date"||key==="staff"||key==="rate")?"nowrap":"normal",color:key==="rate"?"#7C3AED":key==="staff"?"#374151":key==="name"?"#111827":"#6B7280",fontWeight:(key==="rate"||key==="staff")?700:400,borderBottom:"1px solid #F3F4F6"});

  return (
    <div style={{background:"#fff",border:"1px solid #DDD6FE",borderRadius:10,overflow:"hidden"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'DM Sans',sans-serif"}}>
        <thead>
          <tr>{cols.map(c=><th key={c.key} style={thStyle(c.right)}>{c.label}</th>)}</tr>
        </thead>
        <tbody>
          {sorted.map((e,i)=>(
            <tr key={e.id} style={{background:i%2?"#FAFAFA":"#fff"}}>
              {cols.map(c=><td key={c.key} style={tdStyle(c.key,c.right)}>{c.render(e)}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function ExpandPanel({ heading, children }) {
  return (
    <div style={{padding:"12px 20px 16px 36px"}}>
      <div style={{fontSize:11,fontWeight:700,color:"#7C3AED",letterSpacing:.5,marginBottom:10}}>{heading}</div>
      {children}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  MULTI-SELECT FILTER — searchable checkbox dropdown used across Activity Explorer
// ─────────────────────────────────────────────────────────────────────────────

function MultiSelectFilter({ label, options, selected, onChange, width=170 }) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const ref = useRef(null);

  useEffect(()=>{
    const onDocClick = e => { if(ref.current && !ref.current.contains(e.target)) { setOpen(false); setSearch(""); } };
    document.addEventListener("mousedown", onDocClick);
    return ()=>document.removeEventListener("mousedown", onDocClick);
  },[]);

  const filtered = search ? options.filter(o=>o.toLowerCase().includes(search.toLowerCase())) : options;
  const toggle = v => onChange(selected.includes(v) ? selected.filter(x=>x!==v) : [...selected, v]);
  const active = selected.length>0;
  const buttonLabel = !active ? `All ${label}` : selected.length===1 ? selected[0] : `${selected.length} ${label} selected`;

  return (
    <div ref={ref} style={{position:"relative",width}}>
      <button onClick={()=>setOpen(o=>!o)} style={{width:"100%",display:"flex",alignItems:"center",justifyContent:"space-between",gap:6,fontFamily:"'DM Sans',sans-serif",fontSize:13,border:`2px solid ${active?"#011e5c":"#E5E7EB"}`,borderRadius:8,padding:"7px 10px",background:active?"#EEF2FF":"#fff",color:active?"#011e5c":"#374151",cursor:"pointer",fontWeight:active?700:400,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
        <span style={{overflow:"hidden",textOverflow:"ellipsis"}}>{buttonLabel}</span>
        <span style={{fontSize:10,flexShrink:0,color:active?"#011e5c":"#9CA3AF"}}>▾</span>
      </button>
      {open&&(
        <div style={{position:"absolute",top:"calc(100% + 4px)",left:0,zIndex:200,width:Math.max(width,220),background:"#fff",border:"1px solid #E5E7EB",borderRadius:10,boxShadow:"0 8px 24px rgba(0,0,0,.15)",overflow:"hidden"}}>
          {options.length>6&&(
            <div style={{padding:8,borderBottom:"1px solid #F3F4F6"}}>
              <input autoFocus value={search} onChange={e=>setSearch(e.target.value)} placeholder={`Search ${label.toLowerCase()}…`}
                style={{width:"100%",fontFamily:"'DM Sans',sans-serif",fontSize:13,border:"1.5px solid #E5E7EB",borderRadius:6,padding:"6px 8px",outline:"none"}}/>
            </div>
          )}
          <div style={{maxHeight:240,overflowY:"auto",padding:"4px 0"}}>
            {filtered.length===0
              ?<div style={{padding:"10px 12px",fontSize:13,color:"#9CA3AF"}}>No matches</div>
              :filtered.map(o=>{
                const checked=selected.includes(o);
                return(
                  <div key={o} onClick={()=>toggle(o)} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 12px",cursor:"pointer",background:checked?"#F5F3FF":"transparent"}}>
                    <div style={{width:15,height:15,borderRadius:4,border:`2px solid ${checked?"#011e5c":"#D1D5DB"}`,background:checked?"#011e5c":"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                      {checked&&<span style={{color:"#fff",fontSize:10,fontWeight:900}}>✓</span>}
                    </div>
                    <span style={{fontSize:13,color:"#374151",overflow:"hidden",textOverflow:"ellipsis"}}>{o}</span>
                  </div>
                );
              })}
          </div>
          {active&&(
            <div style={{borderTop:"1px solid #F3F4F6",padding:8}}>
              <button onClick={()=>onChange([])} style={{width:"100%",fontSize:12,fontWeight:600,color:"#EF4444",background:"transparent",border:"none",cursor:"pointer",padding:"4px 0"}}>Clear ({selected.length})</button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function Toast({ msg, color, onDone }) {
  useEffect(()=>{ const t=setTimeout(onDone,2800); return()=>clearTimeout(t); },[onDone]);
  return <div style={{position:"fixed",bottom:28,right:28,zIndex:9999,background:color,color:"#fff",padding:"14px 22px",borderRadius:12,fontWeight:600,fontSize:15,boxShadow:"0 8px 32px rgba(0,0,0,.22)",fontFamily:"'DM Sans',sans-serif"}}>{msg}</div>;
}

// ─────────────────────────────────────────────────────────────────────────────
//  EXTERNAL MODAL
// ─────────────────────────────────────────────────────────────────────────────

// A stable, top-level component (not redefined inside ExternalModal's render body).
// It used to be defined inline as "OBox" inside ExternalModal — meaning React saw a
// brand-new component type on every keystroke in the club search box, remounting the
// whole subtree (including the input itself) and dropping focus after every letter.
function SelectionBox({ id, title, sub, selection, onSelect, deptCfg, children }) {
  return (
    <div onClick={()=>onSelect(id)} style={{border:`2px solid ${selection===id?deptCfg.color:"#E5E7EB"}`,borderRadius:10,padding:"12px 14px",marginBottom:8,cursor:"pointer",background:selection===id?deptCfg.light:"#fff"}}>
      <div style={{display:"flex",alignItems:"center",gap:10}}>
        <div style={{width:16,height:16,borderRadius:"50%",border:`2px solid ${selection===id?deptCfg.color:"#D1D5DB"}`,background:selection===id?deptCfg.color:"#fff",flexShrink:0}}/>
        <div><div style={{fontWeight:700,fontSize:14,color:"#111827"}}>{title}</div><div style={{fontSize:12,color:"#6B7280"}}>{sub}</div></div>
      </div>
      {selection===id&&children&&<div style={{marginTop:10}}>{children}</div>}
    </div>
  );
}

function ExternalModal({ item, deptCfg, clubsByLeague, onConfirm, onCancel }) {
  const [selection,    setSelection]   = useState("single");
  const [chosenLeague, setChosenLeague]= useState("");
  const [chosenClubs,  setChosenClubs] = useState([]);
  const [chosenClub,   setChosenClub]  = useState("");
  const [chosenTiers,  setChosenTiers] = useState([]);
  const [singleSearch, setSingleSearch]= useState("");
  const [multiSearch,  setMultiSearch] = useState("");
  const [notes,        setNotes]       = useState("");
  const [chosenDate,   setChosenDate]  = useState(todayDateString());
  const allFlat = Object.entries(clubsByLeague).flatMap(([league,clubs])=>clubs.map(club=>({club,league})));
  const toggleTier = l => setChosenTiers(prev=>prev.includes(l)?prev.filter(x=>x!==l):[...prev,l]);
  const getEntries = () => {
    if (item.multiLeagueSelect) return chosenTiers.map(l=>({club:l,league:l}));
    if (selection==="all")    return allFlat;
    if (selection==="league") return chosenLeague?(clubsByLeague[chosenLeague]||[]).map(club=>({club,league:chosenLeague})):[];
    if (selection==="multi")  return chosenClubs.map(club=>({club,league:leagueForClub(club,clubsByLeague)}));
    return chosenClub?[{club:chosenClub,league:leagueForClub(chosenClub,clubsByLeague)}]:[];
  };
  const entries=getEntries(), count=entries.length, totalVal=count*item.rate;
  const toggleClub=club=>setChosenClubs(prev=>prev.includes(club)?prev.filter(c=>c!==club):[...prev,club]);
  const ss={width:"100%",fontFamily:"'DM Sans',sans-serif",fontSize:14,border:`2px solid ${deptCfg.color}`,borderRadius:8,padding:"9px 12px",background:"#fff",color:"#111",cursor:"pointer",outline:"none",marginTop:8,appearance:"none"};
  const searchStyle={...ss,cursor:"text",appearance:"auto"};

  const singleMatches = singleSearch ? allFlat.filter(({club})=>club.toLowerCase().includes(singleSearch.toLowerCase())) : [];
  const multiFiltered = Object.entries(clubsByLeague).map(([league,clubs])=>[
    league, multiSearch ? clubs.filter(c=>c.toLowerCase().includes(multiSearch.toLowerCase())) : clubs
  ]).filter(([,clubs])=>clubs.length>0);

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:16,padding:"28px 28px 24px",width:"100%",maxWidth:500,boxShadow:"0 24px 64px rgba(0,0,0,.22)",fontFamily:"'DM Sans',sans-serif",maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{fontFamily:"'DM Serif Display',serif",fontSize:19,color:"#111827",marginBottom:4}}>Log External Deliverable</div>
        <div style={{fontSize:13,color:"#6B7280",marginBottom:20}}>{item.name} · <strong>{fmt$(item.rate)} each</strong>{item.recurring&&<span style={{marginLeft:8}}><RecurringBadge/></span>}</div>
        <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:10}}>ATTRIBUTE TO</label>

        {item.multiLeagueSelect?(
          <div style={{marginBottom:8}}>
            <div style={{fontSize:12,color:"#6B7280",marginBottom:10}}>This is league-wide work, not club-specific — select one or more leagues (or USL HQ). One entry is logged per selection.</div>
            <div style={{display:"flex",flexDirection:"column",gap:6}}>
              {LEAGUE_TIER_OPTIONS.map(l=>{
                const checked=chosenTiers.includes(l);
                return(
                  <div key={l} onClick={()=>toggleTier(l)} style={{display:"flex",alignItems:"center",gap:8,padding:"10px 12px",cursor:"pointer",border:`2px solid ${checked?deptCfg.color:"#E5E7EB"}`,borderRadius:8,background:checked?deptCfg.light:"#fff"}}>
                    <div style={{width:16,height:16,borderRadius:4,border:`2px solid ${checked?deptCfg.color:"#D1D5DB"}`,background:checked?deptCfg.color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                      {checked&&<span style={{color:"#fff",fontSize:11,fontWeight:900}}>✓</span>}
                    </div>
                    <span style={{fontSize:14,fontWeight:600,color:"#111827"}}>{l}</span>
                  </div>
                );
              })}
            </div>
          </div>
        ):(
        <>
        <SelectionBox id="single" title="🏟️ Specific Club" sub="Log for one individual club" selection={selection} onSelect={setSelection} deptCfg={deptCfg}>
          {chosenClub?(
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:deptCfg.light,border:`1.5px solid ${deptCfg.color}`,borderRadius:8,padding:"9px 12px"}}>
              <span style={{fontSize:14,fontWeight:700,color:"#111827"}}>{chosenClub}</span>
              <span onClick={e=>{e.stopPropagation();setChosenClub("");setSingleSearch("");}} style={{cursor:"pointer",color:"#6B7280",fontSize:13,fontWeight:700}}>✕</span>
            </div>
          ):(
            <>
              <input type="text" value={singleSearch} onChange={e=>setSingleSearch(e.target.value)} onClick={e=>e.stopPropagation()}
                placeholder="Type to search clubs…" style={searchStyle}/>
              {singleSearch&&(
                <div style={{marginTop:6,maxHeight:200,overflowY:"auto",border:`1px solid ${deptCfg.border}`,borderRadius:8}}>
                  {singleMatches.length===0
                    ?<div style={{padding:"10px 12px",fontSize:13,color:"#9CA3AF"}}>No clubs match "{singleSearch}"</div>
                    :singleMatches.map(({club,league})=>(
                      <div key={club} onClick={e=>{e.stopPropagation();setChosenClub(club);}} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 12px",cursor:"pointer",borderBottom:"1px solid #F3F4F6"}}>
                        <span style={{fontSize:13,color:"#374151"}}>{club}</span>
                        <LeagueBadge league={league}/>
                      </div>
                    ))}
                </div>
              )}
            </>
          )}
        </SelectionBox>

        <SelectionBox id="multi" title="✅ Multiple Clubs" sub="Pick two or more specific clubs" selection={selection} onSelect={setSelection} deptCfg={deptCfg}>
          <input type="text" value={multiSearch} onChange={e=>setMultiSearch(e.target.value)} onClick={e=>e.stopPropagation()}
            placeholder="Search clubs…" style={searchStyle}/>
          <div style={{marginTop:8,maxHeight:220,overflowY:"auto",border:`1px solid ${deptCfg.border}`,borderRadius:8,padding:"4px 0"}}>
            {multiFiltered.length===0&&<div style={{padding:"10px 12px",fontSize:13,color:"#9CA3AF"}}>No clubs match "{multiSearch}"</div>}
            {multiFiltered.map(([league,clubs])=>(
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
        </SelectionBox>

        <SelectionBox id="league" title="📋 Entire League" sub="Logs one line item per club in that league" selection={selection} onSelect={setSelection} deptCfg={deptCfg}>
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
        </SelectionBox>

        <SelectionBox id="all" title="🌐 All Clubs" sub={`Logs one line item for every club (${allFlat.length} total)`} selection={selection} onSelect={setSelection} deptCfg={deptCfg}>
          <div style={{background:"#F8FAFC",borderRadius:8,padding:"10px 12px"}}>
            {Object.entries(clubsByLeague).map(([league,clubs])=>(
              <div key={league} style={{marginBottom:6}}>
                <div style={{fontSize:11,fontWeight:700,color:"#6B7280",marginBottom:3}}>{league} ({clubs.length})</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:3}}>{clubs.map(c=><span key={c} style={{fontSize:10,background:"#fff",border:"1px solid #E5E7EB",borderRadius:5,padding:"1px 6px"}}>{c}</span>)}</div>
              </div>
            ))}
          </div>
        </SelectionBox>
        </>
        )}

        {count>0&&(
          <div style={{background:deptCfg.light,border:`1px solid ${deptCfg.border}`,borderRadius:10,padding:"12px 14px",marginBottom:4}}>
            <div style={{fontSize:13,fontWeight:700,color:deptCfg.color}}>
              {count===1?`1 entry · ${fmt$(item.rate)}`:`${count} entries × ${fmt$(item.rate)} = `}
              {count>1&&<span style={{fontSize:16}}>{fmt$(totalVal)} total</span>}
            </div>
          </div>
        )}
        {/* Date field — defaults to today, overridable for retroactive logging */}
        <div style={{marginTop:8}}>
          <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:6}}>DATE OF WORK</label>
          <input type="date" value={chosenDate} onChange={e=>setChosenDate(e.target.value)} max={todayDateString()}
            style={{fontFamily:"'DM Sans',sans-serif",fontSize:13,border:"1.5px solid #E5E7EB",borderRadius:8,padding:"9px 12px",outline:"none",color:"#374151",background:"#FAFAFA"}}/>
        </div>
        {/* Notes field */}
        <div style={{marginTop:8}}>
          <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:6}}>NOTES <span style={{fontWeight:400,color:"#9CA3AF"}}>(optional)</span></label>
          <textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Any additional context for this deliverable..." rows={2}
            style={{width:"100%",fontFamily:"'DM Sans',sans-serif",fontSize:13,border:"1.5px solid #E5E7EB",borderRadius:8,padding:"9px 12px",outline:"none",resize:"vertical",color:"#374151",background:"#FAFAFA"}}/>
        </div>

        <div style={{display:"flex",gap:10,marginTop:12}}>
          <button onClick={()=>count&&onConfirm(entries,notes,dateStringToTs(chosenDate))} disabled={!count} style={{flex:1,padding:11,border:"none",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:14,color:"#fff",background:count?deptCfg.color:"#D1D5DB",cursor:count?"pointer":"not-allowed"}}>
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

// ─────────────────────────────────────────────────────────────────────────────
//  BATCH LOG MODAL — multiple deliverables, one shared club/recipient + notes
// ─────────────────────────────────────────────────────────────────────────────

function BatchLogModal({ selection, activeMode, dept, deptCfg, clubsByLeague, onConfirm, onCancel }) {
  const [subcatChoices, setSubcatChoices] = useState(() => {
    const initial = {};
    selection.forEach(({idx,item}) => { if(item.subcategories.length===1) initial[idx]=item.subcategories[0]; });
    return initial;
  });
  const [chosenClub, setChosenClub] = useState("");
  const [clubSearch, setClubSearch] = useState("");
  const [recipient, setRecipient] = useState("");
  const [notes, setNotes] = useState("");
  const [chosenDate, setChosenDate] = useState(todayDateString());

  const allFlat = Object.entries(clubsByLeague).flatMap(([league,clubs])=>clubs.map(club=>({club,league})));
  const clubMatches = clubSearch ? allFlat.filter(({club})=>club.toLowerCase().includes(clubSearch.toLowerCase())) : [];
  const recipients = INTERNAL_RECIPIENTS[dept]||[];

  const allSubcatsChosen = selection.every(({idx})=>subcatChoices[idx]);
  const attributionChosen = activeMode==="internal" ? !!recipient : !!chosenClub;
  const canConfirm = allSubcatsChosen && attributionChosen;

  const totalValue = selection.reduce((s,{idx})=>{
    const sub = subcatChoices[idx];
    return s + (sub && activeMode==="external" ? sub.rate : 0);
  },0);

  const handleConfirmClick = () => {
    if (!canConfirm) return;
    const club = activeMode==="internal" ? recipient : chosenClub;
    const league = activeMode==="internal" ? "Internal" : leagueForClub(chosenClub,clubsByLeague);
    const ts = dateStringToTs(chosenDate);
    const finalEntries = selection.map(({idx,item})=>{
      const sub = subcatChoices[idx];
      return {
        dept, name:item.name, subcat:sub.subcat,
        rate: activeMode==="internal" ? 0 : sub.rate,
        type: activeMode==="internal"?"Internal":"External",
        cat:item.cat, index_score:sub.index_score||1, recurring:sub.recurring||false,
        club, league, notes, ts,
      };
    });
    onConfirm(finalEntries);
  };

  const ss={width:"100%",fontFamily:"'DM Sans',sans-serif",fontSize:14,border:`2px solid ${deptCfg.color}`,borderRadius:8,padding:"9px 12px",background:"#fff",color:"#111",cursor:"text",outline:"none"};

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:16,padding:"28px 28px 24px",width:"100%",maxWidth:520,boxShadow:"0 24px 64px rgba(0,0,0,.22)",fontFamily:"'DM Sans',sans-serif",maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{fontFamily:"'DM Serif Display',serif",fontSize:19,color:"#111827",marginBottom:4}}>Log {selection.length} Deliverables</div>
        <div style={{fontSize:13,color:"#6B7280",marginBottom:20}}>All logged to the same {activeMode==="internal"?"recipient":"club"}, with one shared note.</div>

        <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:10}}>SELECTED DELIVERABLES</label>
        <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:20}}>
          {selection.map(({idx,item})=>(
            <div key={idx} style={{border:"1.5px solid #E5E7EB",borderRadius:10,padding:"10px 12px"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:10}}>
                <span style={{fontWeight:700,fontSize:13,color:"#111827"}}>{item.name}</span>
                {activeMode==="external"&&subcatChoices[idx]&&<span style={{fontSize:12,fontWeight:700,color:deptCfg.color}}>{fmt$(subcatChoices[idx].rate)}</span>}
              </div>
              {item.subcategories.length>1&&(
                <select value={subcatChoices[idx]?.subcat||""} onChange={e=>{
                  const sub=item.subcategories.find(s=>s.subcat===e.target.value);
                  setSubcatChoices(prev=>({...prev,[idx]:sub}));
                }} style={{...ss,marginTop:8,fontSize:13,padding:"7px 10px"}}>
                  <option value="" disabled>Select type…</option>
                  {item.subcategories.map(s=><option key={s.subcat} value={s.subcat}>{s.subcat} {activeMode==="external"?`— ${fmt$(s.rate)}`:""}</option>)}
                </select>
              )}
            </div>
          ))}
        </div>

        <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:10}}>{activeMode==="internal"?"INTERNAL RECIPIENT":"CLUB"}</label>
        {activeMode==="internal"?(
          <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:16}}>
            {recipients.map(r=>(
              <div key={r} onClick={()=>setRecipient(r)} style={{border:`2px solid ${recipient===r?deptCfg.color:"#E5E7EB"}`,borderRadius:10,padding:"12px 14px",cursor:"pointer",background:recipient===r?deptCfg.light:"#fff",display:"flex",alignItems:"center",gap:10}}>
                <div style={{width:16,height:16,borderRadius:"50%",border:`2px solid ${recipient===r?deptCfg.color:"#D1D5DB"}`,background:recipient===r?deptCfg.color:"#fff",flexShrink:0}}/>
                <span style={{fontWeight:600,fontSize:14,color:"#111827"}}>{r}</span>
              </div>
            ))}
          </div>
        ):(
          <div style={{marginBottom:16}}>
            {chosenClub?(
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:deptCfg.light,border:`1.5px solid ${deptCfg.color}`,borderRadius:8,padding:"9px 12px"}}>
                <span style={{fontSize:14,fontWeight:700,color:"#111827"}}>{chosenClub}</span>
                <span onClick={()=>{setChosenClub("");setClubSearch("");}} style={{cursor:"pointer",color:"#6B7280",fontSize:13,fontWeight:700}}>✕</span>
              </div>
            ):(
              <>
                <input type="text" value={clubSearch} onChange={e=>setClubSearch(e.target.value)} placeholder="Type to search clubs…" style={ss}/>
                {clubSearch&&(
                  <div style={{marginTop:6,maxHeight:180,overflowY:"auto",border:`1px solid ${deptCfg.border}`,borderRadius:8}}>
                    {clubMatches.length===0
                      ?<div style={{padding:"10px 12px",fontSize:13,color:"#9CA3AF"}}>No clubs match "{clubSearch}"</div>
                      :clubMatches.map(({club,league})=>(
                        <div key={club} onClick={()=>setChosenClub(club)} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 12px",cursor:"pointer",borderBottom:"1px solid #F3F4F6"}}>
                          <span style={{fontSize:13,color:"#374151"}}>{club}</span>
                          <LeagueBadge league={league}/>
                        </div>
                      ))}
                  </div>
                )}
              </>
            )}
          </div>
        )}

        {activeMode==="external"&&totalValue>0&&(
          <div style={{background:deptCfg.light,border:`1px solid ${deptCfg.border}`,borderRadius:10,padding:"10px 14px",marginBottom:16}}>
            <div style={{fontSize:13,fontWeight:700,color:deptCfg.color}}>{selection.length} deliverables · {fmt$(totalValue)} total</div>
          </div>
        )}

        <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:6}}>DATE OF WORK <span style={{fontWeight:400,color:"#9CA3AF"}}>(applies to all {selection.length})</span></label>
        <input type="date" value={chosenDate} onChange={e=>setChosenDate(e.target.value)} max={todayDateString()}
          style={{fontFamily:"'DM Sans',sans-serif",fontSize:13,border:"1.5px solid #E5E7EB",borderRadius:8,padding:"9px 12px",outline:"none",color:"#374151",background:"#FAFAFA",marginBottom:16}}/>

        <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:6}}>NOTES <span style={{fontWeight:400,color:"#9CA3AF"}}>(applies to all {selection.length})</span></label>
        <textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Any additional context for this batch…" rows={2}
          style={{width:"100%",fontFamily:"'DM Sans',sans-serif",fontSize:13,border:"1.5px solid #E5E7EB",borderRadius:8,padding:"9px 12px",outline:"none",resize:"vertical",color:"#374151",background:"#FAFAFA"}}/>

        <div style={{display:"flex",gap:10,marginTop:16}}>
          <button onClick={handleConfirmClick} disabled={!canConfirm} style={{flex:1,padding:11,border:"none",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:14,color:"#fff",background:canConfirm?deptCfg.color:"#D1D5DB",cursor:canConfirm?"pointer":"not-allowed"}}>
            Log {selection.length} Entries ▶
          </button>
          <button onClick={onCancel} style={{padding:"11px 18px",border:"1.5px solid #E5E7EB",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:14,color:"#6B7280",background:"#fff",cursor:"pointer"}}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

function InternalModal({ item, deptCfg, dept, clubsByLeague, onConfirm, onCancel }) {
  const [recipient,      setRecipient]     = useState("");
  const [chosenLeague,   setChosenLeague]  = useState("");
  const [chosenTiers,    setChosenTiers]   = useState([]);
  const [notes,          setNotes]         = useState("");
  const [chosenDate,     setChosenDate]    = useState(todayDateString());
  const recipients = INTERNAL_RECIPIENTS[dept]||[];
  const isLeagueSelect = item.leagueSelect;
  const isLeagueOpsTiers = recipient==="League Operations";
  const toggleTier = l => setChosenTiers(prev=>prev.includes(l)?prev.filter(x=>x!==l):[...prev,l]);

  const getEntries = () => {
    if (isLeagueSelect&&chosenLeague) return (clubsByLeague[chosenLeague]||[]).map(club=>({club,league:chosenLeague}));
    if (isLeagueOpsTiers) return chosenTiers.map(l=>({club:recipient,league:l}));
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
                <div key={r} onClick={()=>{setRecipient(r);setChosenTiers([]);}} style={{border:`2px solid ${recipient===r?deptCfg.color:"#E5E7EB"}`,borderRadius:10,padding:"12px 14px",cursor:"pointer",background:recipient===r?deptCfg.light:"#fff",display:"flex",alignItems:"center",gap:10}}>
                  <div style={{width:16,height:16,borderRadius:"50%",border:`2px solid ${recipient===r?deptCfg.color:"#D1D5DB"}`,background:recipient===r?deptCfg.color:"#fff",flexShrink:0}}/>
                  <span style={{fontWeight:600,fontSize:14,color:"#111827"}}>{r}</span>
                </div>
              ))}
            </div>

            {isLeagueOpsTiers&&(
              <div style={{marginTop:16}}>
                <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:10}}>APPLIES TO LEAGUE(S)</label>
                <div style={{display:"flex",flexDirection:"column",gap:6,maxHeight:220,overflowY:"auto",border:`1px solid ${deptCfg.border}`,borderRadius:8,padding:"4px 0"}}>
                  {LEAGUE_TIER_OPTIONS.map(l=>{
                    const checked=chosenTiers.includes(l);
                    return(
                      <div key={l} onClick={()=>toggleTier(l)} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",cursor:"pointer",background:checked?deptCfg.light:"transparent"}}>
                        <div style={{width:15,height:15,borderRadius:4,border:`2px solid ${checked?deptCfg.color:"#D1D5DB"}`,background:checked?deptCfg.color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                          {checked&&<span style={{color:"#fff",fontSize:10,fontWeight:900}}>✓</span>}
                        </div>
                        <span style={{fontSize:13,color:"#374151"}}>{l}</span>
                      </div>
                    );
                  })}
                </div>
                {chosenTiers.length>0&&<div style={{fontSize:12,color:deptCfg.color,fontWeight:600,marginTop:6}}>{chosenTiers.length} league{chosenTiers.length>1?"s":""} selected · {count} entries × {fmt$(item.rate)} = {fmt$(count*item.rate)} total</div>}
              </div>
            )}
          </>
        )}

        {/* Date field — defaults to today, overridable for retroactive logging */}
        <div style={{marginTop:12}}>
          <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:6}}>DATE OF WORK</label>
          <input type="date" value={chosenDate} onChange={e=>setChosenDate(e.target.value)} max={todayDateString()}
            style={{fontFamily:"'DM Sans',sans-serif",fontSize:13,border:"1.5px solid #E5E7EB",borderRadius:8,padding:"9px 12px",outline:"none",color:"#374151",background:"#FAFAFA"}}/>
        </div>

        {/* Notes field */}
        <div style={{marginTop:12}}>
          <label style={{display:"block",fontSize:12,fontWeight:700,color:"#6B7280",letterSpacing:.5,marginBottom:6}}>NOTES <span style={{fontWeight:400,color:"#9CA3AF"}}>(optional)</span></label>
          <textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Any additional context..." rows={2}
            style={{width:"100%",fontFamily:"'DM Sans',sans-serif",fontSize:13,border:"1.5px solid #E5E7EB",borderRadius:8,padding:"9px 12px",outline:"none",resize:"vertical",color:"#374151",background:"#FAFAFA"}}/>
        </div>

        <div style={{display:"flex",gap:10,marginTop:12}}>
          <button onClick={()=>count&&onConfirm(entries,notes,dateStringToTs(chosenDate))} disabled={!count} style={{flex:1,padding:11,border:"none",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:14,color:"#fff",background:count?deptCfg.color:"#D1D5DB",cursor:count?"pointer":"not-allowed"}}>
            {count>1?`Log ${count} Entries ▶`:"Confirm & Log ▶"}
          </button>
          <button onClick={onCancel} style={{padding:"11px 18px",border:"1.5px solid #E5E7EB",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:14,color:"#6B7280",background:"#fff",cursor:"pointer"}}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  SUBCATEGORY PICKER — shown when an item has more than one priced subcategory
// ─────────────────────────────────────────────────────────────────────────────

function SubcategoryPickerModal({ item, deptCfg, onSelect, onCancel }) {
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:16,padding:"28px 28px 24px",width:"100%",maxWidth:440,boxShadow:"0 24px 64px rgba(0,0,0,.22)",fontFamily:"'DM Sans',sans-serif",maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{fontFamily:"'DM Serif Display',serif",fontSize:19,color:"#111827",marginBottom:4}}>{item.name}</div>
        <div style={{fontSize:13,color:"#6B7280",marginBottom:20}}>Select the specific type to log</div>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {item.subcategories.map((s,i)=>(
            <div key={i} onClick={()=>onSelect(s)} style={{border:"2px solid #E5E7EB",borderRadius:10,padding:"12px 14px",cursor:"pointer",background:"#fff",transition:"border-color .15s"}}
              onMouseEnter={e=>e.currentTarget.style.borderColor=deptCfg.color} onMouseLeave={e=>e.currentTarget.style.borderColor="#E5E7EB"}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8}}>
                <span style={{fontWeight:700,fontSize:14,color:"#111827"}}>{s.subcat}</span>
                {s.recurring&&<RecurringBadge/>}
              </div>
              <div style={{fontSize:12,color:"#6B7280",marginTop:4}}>
                {s.rate>0&&<strong style={{color:"#374151"}}>{fmt$(s.rate)}</strong>}
                <span style={{marginLeft:8,fontSize:11,background:"#F3F4F6",borderRadius:6,padding:"1px 7px",color:"#475569"}}>Index {s.index_score}</span>
              </div>
            </div>
          ))}
        </div>
        <button onClick={onCancel} style={{marginTop:16,width:"100%",padding:"11px 18px",border:"1.5px solid #E5E7EB",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:14,color:"#6B7280",background:"#fff",cursor:"pointer"}}>Cancel</button>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  CATEGORY GROUPED CARDS
// ─────────────────────────────────────────────────────────────────────────────

function CategoryGroupedCards({ items, log, dept, cfg, pulsingIdx, onLogClick, batchSelected, onToggleBatch }) {
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
                  const subs=item.subcategories||[];
                  const multi=subs.length>1;
                  const rates=subs.map(s=>s.rate).filter(r=>r>0);
                  const rateDisplay = rates.length===0 ? null
                    : !multi ? fmt$(subs[0].rate)
                    : (Math.min(...rates)===Math.max(...rates) ? fmt$(rates[0]) : `From ${fmt$(Math.min(...rates))}`);
                  const idxVals=subs.map(s=>s.index_score);
                  const indexDisplay = !multi ? subs[0]?.index_score
                    : (Math.min(...idxVals)===Math.max(...idxVals) ? idxVals[0] : `${Math.min(...idxVals)}–${Math.max(...idxVals)}`);
                  const allRecurring = subs.length>0 && subs.every(s=>s.recurring);
                  const teaser = multi
                    ? subs.map(s=>s.subcat).join(" · ")
                    : (subs[0]?.examples||[]).slice(0,4).join(" · ");
                  const teaserExtra = multi ? 0 : Math.max(0,(subs[0]?.examples?.length||0)-4);
                  // Items that need special multi-target attribution (league-wide social,
                  // League Ops tiers) don't fit the "one shared attribution" batch flow.
                  const batchEligible = !item.leagueSelect && !item.multiLeagueSelect;
                  const isChecked = batchSelected && batchSelected.includes(idx);
                  return(
                    <div key={item.name} style={{background:"#fff",border:`1.5px solid ${isPulsing?cfg.color:isChecked?cfg.color:"#E5E7EB"}`,borderRadius:12,padding:16,display:"flex",flexDirection:"column",gap:8,boxShadow:isPulsing?`0 0 0 4px ${cfg.color}33`:isChecked?`0 0 0 3px ${cfg.color}33`:"0 1px 4px rgba(0,0,0,.06)",transition:"border-color .2s,box-shadow .2s"}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
                        <div style={{display:"flex",alignItems:"flex-start",gap:8,flex:1}}>
                          {batchEligible&&onToggleBatch&&(
                            <div onClick={()=>onToggleBatch(idx)} style={{width:16,height:16,borderRadius:4,border:`2px solid ${isChecked?cfg.color:"#D1D5DB"}`,background:isChecked?cfg.color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,cursor:"pointer",marginTop:2}}>
                              {isChecked&&<span style={{color:"#fff",fontSize:10,fontWeight:900}}>✓</span>}
                            </div>
                          )}
                          <span style={{fontFamily:"'DM Serif Display',serif",fontSize:15,color:"#111827",lineHeight:1.3,flex:1}}>{item.name}</span>
                        </div>
                        {allRecurring&&<RecurringBadge/>}
                      </div>
                      {teaser&&(
                        <div style={{fontSize:11,color:"#6B7280",fontStyle:multi?"normal":"italic",lineHeight:1.5}}>
                          {teaser}
                          {teaserExtra>0&&<span style={{color:cfg.color,fontStyle:"normal",fontWeight:600,marginLeft:4}}>+{teaserExtra} more</span>}
                        </div>
                      )}
                      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:2}}>
                        <span style={{fontSize:13,color:"#6B7280"}}>
                          {rateDisplay&&<strong style={{color:"#374151"}}>{rateDisplay}</strong>}
                          {indexDisplay!=null&&<span style={{marginLeft:8,fontSize:11,background:"#F3F4F6",borderRadius:6,padding:"1px 7px",color:"#475569"}}>Index {indexDisplay}</span>}
                          {multi&&<span style={{marginLeft:8,fontSize:11,background:cfg.light,color:cfg.color,borderRadius:6,padding:"1px 7px",fontWeight:600}}>{subs.length} types</span>}
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
  const [subcatPick,setSubcatPick]=useState(null);
  const [selectedSubcat,setSelectedSubcat]=useState(null);
  const [modal,setModal]=useState(null);
  const [pulsingIdx,setPulsingIdx]=useState(null);
  const [batchSelected,setBatchSelected]=useState({external:[],internal:[]});
  const [showBatchModal,setShowBatchModal]=useState(false);

  const deptData=deptItems[dept]||{external:[],internal:[]};
  const activeMode=isIntOnly?"internal":mode;
  const currentBatchSelected = batchSelected[activeMode]||[];
  const otherMode = activeMode==="internal" ? "external" : "internal";
  const otherModeCount = (batchSelected[otherMode]||[]).length;
  const items=activeMode==="internal"?deptData.internal:deptData.external;
  // Internal work never shows a price on the browsing cards either — the write path
  // already forces $0 regardless, this just keeps what's displayed consistent with that.
  const displayItems = activeMode==="internal"
    ? items.map(it=>({...it, subcategories: it.subcategories.map(s=>({...s,rate:0}))}))
    : items;
  const deptLog=log.filter(e=>e.dept===dept);
  const deptTotal=deptLog.reduce((s,e)=>s+e.rate,0);

  const handleLogClick=idx=>{
    if(!staffName){setNameErr(true);return;}
    setNameErr(false);
    const item=items[idx];
    if(item.subcategories.length>1){ setSubcatPick(idx); }
    else { setSelectedSubcat(item.subcategories[0]); setModal(idx); }
  };

  const handleSubcatSelect=sub=>{
    setSelectedSubcat(sub);
    setModal(subcatPick);
    setSubcatPick(null);
  };

  const handleConfirm=async (entries, notes="", ts=Date.now())=>{
    const item=items[modal];
    const sub=selectedSubcat||item.subcategories[0];
    // Internal work never carries a dollar figure — enforced here at write time so the
    // sheet itself stays clean, not just masked by the app's display layer.
    const rate = activeMode==="internal" ? 0 : sub.rate;
    setModal(null);setSelectedSubcat(null);setPulsingIdx(modal);setTimeout(()=>setPulsingIdx(null),400);
    const total=entries.length;
    if(total===1){
      const{club,league}=entries[0];
      await onLog({dept,name:item.name,subcat:sub.subcat,rate,type:activeMode==="internal"?"Internal":"External",cat:item.cat,index_score:sub.index_score||1,recurring:sub.recurring||false,staff:staffName,club,league,notes,ts,bulkSilent:false});
    } else {
      // Build all entries first, add to UI optimistically, then send as one batch request
      const newEntries = entries.map(({club,league})=>({
        dept,name:item.name,subcat:sub.subcat,rate,
        type:activeMode==="internal"?"Internal":"External",
        cat:item.cat,index_score:sub.index_score||1,
        recurring:sub.recurring||false,staff:staffName,club,league,notes,
        id:`${Date.now()}-${Math.random().toString(36).slice(2)}`,
        ts,
      }));
      // Add all to UI at once
      onBulkLog(newEntries);
      // Send as single batch request to Apps Script
      try {
        await dbInsertBatch(newEntries.map(({bulkSilent,...e})=>e));
      } catch(err) {
        console.error("Batch insert failed:", err);
      }
      onBulkComplete(total, item.name, rate*total);
    }
  };

  const toggleBatchSelect = idx => {
    if(!staffName){setNameErr(true);return;}
    setNameErr(false);
    setBatchSelected(prev => {
      const cur = prev[activeMode]||[];
      return { ...prev, [activeMode]: cur.includes(idx) ? cur.filter(i=>i!==idx) : [...cur,idx] };
    });
  };

  const handleBatchConfirm = async finalEntries => {
    setShowBatchModal(false);
    setBatchSelected(prev => ({...prev, [activeMode]: []}));
    const newEntries = finalEntries.map(e=>({
      ...e, staff:staffName,
      id:`${Date.now()}-${Math.random().toString(36).slice(2)}`,
      ts: e.ts || Date.now(),
    }));
    onBulkLog(newEntries);
    try {
      await dbInsertBatch(newEntries);
    } catch(err) {
      console.error("Batch insert failed:", err);
    }
    const total = newEntries.reduce((s,e)=>s+e.rate,0);
    const uniqueNames = [...new Set(newEntries.map(e=>e.name))];
    const label = uniqueNames.length<=2 ? uniqueNames.join(", ") : `${uniqueNames.length} deliverable types`;
    onBulkComplete(newEntries.length, label, total);
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
            {["external","internal"].map(m=>{
              const pending=(batchSelected[m]||[]).length;
              return(
                <button key={m} onClick={()=>setMode(m)} style={{position:"relative",padding:"7px 16px",borderRadius:8,border:"none",fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:12,cursor:"pointer",background:activeMode===m?cfg.color:"transparent",color:activeMode===m?"#fff":"#64748B",transition:"all .15s"}}>
                  {m==="external"?"External":"Internal"}
                  {pending>0&&<span style={{position:"absolute",top:-6,right:-6,background:"#f51200",color:"#fff",borderRadius:99,fontSize:10,fontWeight:700,minWidth:16,height:16,display:"flex",alignItems:"center",justifyContent:"center",padding:"0 4px"}}>{pending}</span>}
                </button>
              );
            })}
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
        :<CategoryGroupedCards items={displayItems} log={log} dept={dept} cfg={cfg} pulsingIdx={pulsingIdx} onLogClick={handleLogClick} batchSelected={currentBatchSelected} onToggleBatch={toggleBatchSelect}/>
      }

      {currentBatchSelected.length>0&&(
        <div style={{position:"sticky",bottom:20,marginTop:20,display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,background:"#011e5c",borderRadius:12,padding:"12px 20px",boxShadow:"0 8px 24px rgba(0,0,0,.25)",flexWrap:"wrap"}}>
          <span style={{color:"#fff",fontSize:14,fontWeight:600}}>
            {currentBatchSelected.length} deliverable{currentBatchSelected.length>1?"s":""} selected
            {otherModeCount>0&&<span style={{color:"#9DB4E0",fontWeight:400,marginLeft:8}}>· {otherModeCount} {otherMode} still waiting</span>}
          </span>
          <div style={{display:"flex",gap:10}}>
            <button onClick={()=>setBatchSelected(prev=>({...prev,[activeMode]:[]}))} style={{background:"transparent",border:"1.5px solid #3B5A94",color:"#C7D2FE",borderRadius:8,padding:"8px 16px",fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:13,cursor:"pointer"}}>Clear</button>
            <button onClick={()=>setShowBatchModal(true)} style={{background:"#f51200",border:"none",color:"#fff",borderRadius:8,padding:"8px 18px",fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>Log Selected ▶</button>
          </div>
        </div>
      )}

      {showBatchModal&&(
        <BatchLogModal
          selection={currentBatchSelected.map(idx=>({idx,item:items[idx]}))}
          activeMode={activeMode} dept={dept} deptCfg={cfg} clubsByLeague={clubsByLeague}
          onConfirm={handleBatchConfirm} onCancel={()=>setShowBatchModal(false)}
        />
      )}

      {subcatPick!==null&&(
        <SubcategoryPickerModal
          item={activeMode==="internal"
            ? {...items[subcatPick], subcategories: items[subcatPick].subcategories.map(s=>({...s,rate:0}))}
            : items[subcatPick]}
          deptCfg={cfg} onSelect={handleSubcatSelect} onCancel={()=>setSubcatPick(null)}/>
      )}

      {modal!==null&&(()=>{
        const sub=selectedSubcat||items[modal].subcategories[0];
        const effectiveItem={
          name: items[modal].subcategories.length>1 ? `${items[modal].name} — ${sub.subcat}` : items[modal].name,
          cat: items[modal].cat,
          leagueSelect: items[modal].leagueSelect,
          multiLeagueSelect: items[modal].multiLeagueSelect,
          rate: activeMode==="internal" ? 0 : sub.rate,
          index_score: sub.index_score,
          recurring: sub.recurring,
          examples: sub.examples,
        };
        return activeMode==="internal"
          ? <InternalModal item={effectiveItem} deptCfg={cfg} dept={dept} clubsByLeague={clubsByLeague} onConfirm={handleConfirm} onCancel={()=>{setModal(null);setSelectedSubcat(null);}}/>
          : <ExternalModal item={effectiveItem} deptCfg={cfg} clubsByLeague={clubsByLeague} onConfirm={handleConfirm} onCancel={()=>{setModal(null);setSelectedSubcat(null);}}/>;
      })()}
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
  "Strategic Planning Session": 750,
  // Default
  "default": 250,
};

function getRackRate(entry) {
  if (entry.type==="Internal") return 0;
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
            {d.sub && <div style={{width:90,flexShrink:0,fontSize:11,color:"#9CA3AF",fontFamily:"'DM Sans',sans-serif",whiteSpace:"nowrap"}}>{d.sub}</div>}
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
    const startAngle = cumAngle;
    const angle = (sl.value/total)*2*Math.PI;
    const x1=cx+r*Math.cos(cumAngle), y1=cy+r*Math.sin(cumAngle);
    cumAngle+=angle;
    const x2=cx+r*Math.cos(cumAngle), y2=cy+r*Math.sin(cumAngle);
    const midAngle=cumAngle-angle/2;
    // A single 100% slice has identical start/end points, which collapses a normal SVG arc
    // to nothing visible — draw it as two semicircles instead when that's the case.
    const isFull = angle >= 2*Math.PI - 0.001;
    const halfAngle = startAngle + angle/2;
    const xh=cx+r*Math.cos(halfAngle), yh=cy+r*Math.sin(halfAngle);
    return{...sl,x1,y1,x2,y2,xh,yh,isFull,large:angle>Math.PI?1:0,midAngle,pct:sl.value/total,index:i};
  });
  const innerR=r*0.56;
  return (
    <div style={{display:"flex",alignItems:"center",gap:24,flexWrap:"wrap"}}>
      <div style={{position:"relative",flexShrink:0}}>
        <svg width={size} height={size} style={{overflow:"visible"}}>
          {arcs.map((arc,i)=>{
            const isH=hovered===i;
            const off=isH?7:0, ox=off*Math.cos(arc.midAngle), oy=off*Math.sin(arc.midAngle);
            const d = arc.isFull
              ? `M ${arc.x1+ox} ${arc.y1+oy} A ${r} ${r} 0 1 1 ${arc.xh+ox} ${arc.yh+oy} A ${r} ${r} 0 1 1 ${arc.x2+ox} ${arc.y2+oy} Z`
              : `M ${cx+ox} ${cy+oy} L ${arc.x1+ox} ${arc.y1+oy} A ${r} ${r} 0 ${arc.large} 1 ${arc.x2+ox} ${arc.y2+oy} Z`;
            return(
              <path key={i}
                d={d}
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
      <div style={{display:"flex",flexDirection:"column",gap:8,width:"fit-content",minWidth:150,maxWidth:230}}>
        {arcs.map((arc,i)=>(
          <div key={i} onMouseEnter={()=>setHovered(i)} onMouseLeave={()=>setHovered(null)}
            style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer",padding:"4px 8px",borderRadius:8,background:hovered===i?arc.color+"12":"transparent",transition:"background .15s"}}>
            <div style={{width:12,height:12,borderRadius:"50%",background:arc.color,flexShrink:0}}/>
            <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:13,color:"#374151",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{arc.label}</span>
            <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:13,fontWeight:700,color:arc.color,marginLeft:"auto",flexShrink:0}}>{(arc.pct*100).toFixed(0)}%</span>
            <span style={{fontFamily:"'DM Sans',sans-serif",fontSize:11,color:"#9CA3AF",width:34,textAlign:"right",flexShrink:0}}>{arc.value.toLocaleString()}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  YEAR-OVER-YEAR TREND — self-contained filters, independent of the main
//  dashboard's Period/League/Cluster state so it can always compare across years.
// ─────────────────────────────────────────────────────────────────────────────

function YoYTrendCard({ log, clubClusters }) {
  const [metric, setMetric] = useState("count");
  const [leagueScope, setLeagueScope] = useState("all");
  const [compareMode, setCompareMode] = useState("years"); // 'years' | 'clusters' | 'clubs'
  const [compareYear, setCompareYear] = useState(new Date().getFullYear());
  const [selectedClusters, setSelectedClusters] = useState([]);
  const [selectedClubs, setSelectedClubs] = useState([]);
  const [hiddenKeys, setHiddenKeys] = useState(new Set());
  const [showTable, setShowTable] = useState(false);

  const MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const currentYear = new Date().getFullYear();
  const currentMonth = new Date().getMonth();
  const COMPARE_PALETTE = ["#94A3B8","#0369A1","#7C3AED","#DB2777","#0891B2","#B45309","#047857","#4338CA","#B91C1C","#0F766E"];

  const scoped = log.filter(e => leagueScope==="all" || e.league===leagueScope);
  const leagueOptions = [...new Set(log.map(e=>e.league))].filter(Boolean).sort();
  const yearOptions = [...new Set(log.map(e=>new Date(e.ts).getFullYear()))].sort((a,b)=>b-a);
  const clusterOptions = [...new Set(Object.values(clubClusters))].filter(Boolean).sort();
  const clubOptions = [...new Set(scoped.map(e=>e.club))].filter(Boolean).sort();

  // Same correct methodology as the rest of the dashboard: average each club's own
  // average first, then average those — never a raw sum divided by entity count.
  const avgIndexForEntries = entries => {
    const perClub={};
    entries.filter(e=>!e.recurring&&e.type==="External").forEach(e=>{
      const c=e.club||"Unknown";
      if(!perClub[c]) perClub[c]={sum:0,count:0};
      perClub[c].sum+=Number(e.index_score||1); perClub[c].count++;
    });
    const avgs=Object.values(perClub).map(c=>c.sum/c.count);
    return avgs.length ? avgs.reduce((s,v)=>s+v,0)/avgs.length : 0;
  };
  const metricValue = entries => {
    if(metric==="count") return entries.length;
    if(metric==="value") return entries.reduce((s,e)=>s+getRackRate(e),0);
    return avgIndexForEntries(entries);
  };
  const fmtMetric = v => metric==="value" ? fmt$(Math.round(v)) : metric==="index" ? (v>0?v.toFixed(1):"—") : Math.round(v).toLocaleString();

  // Bucket scoped entries into month-arrays keyed by whichever dimension is being compared.
  const monthBuckets = {};
  const pushInto = (key,e) => { if(!monthBuckets[key]) monthBuckets[key]=Array.from({length:12},()=>[]); monthBuckets[key][new Date(e.ts).getMonth()].push(e); };
  if (compareMode==="years") {
    scoped.forEach(e=>pushInto(new Date(e.ts).getFullYear(),e));
  } else if (compareMode==="leagues") {
    log.filter(e=>new Date(e.ts).getFullYear()===compareYear && e.league).forEach(e=>pushInto(e.league,e));
  } else if (compareMode==="clusters") {
    scoped.filter(e=>new Date(e.ts).getFullYear()===compareYear && clubClusters[e.club]).forEach(e=>pushInto(clubClusters[e.club],e));
  } else {
    scoped.filter(e=>new Date(e.ts).getFullYear()===compareYear && selectedClubs.includes(e.club)).forEach(e=>pushInto(e.club,e));
  }
  // null = no entries that month at all (a real gap), as opposed to a computed value
  // of 0 for a month that has entries but sums to zero — those stay on the line.
  const seriesForKey = key => (monthBuckets[key]||Array.from({length:12},()=>[])).map(entries => entries.length ? metricValue(entries) : null);
  const buildSegments = series => {
    const segments = []; let current = [];
    series.forEach((v,m) => {
      if (v==null) { if (current.length) segments.push(current); current = []; }
      else current.push({ m, v });
    });
    if (current.length) segments.push(current);
    return segments;
  };

  let keys = [];
  if (compareMode==="years") keys = Object.keys(monthBuckets).map(Number).sort((a,b)=>a-b);
  else if (compareMode==="leagues") keys = Object.keys(monthBuckets).filter(l=>l!=="USL HQ").sort(sortByLeagueOrder);
  else if (compareMode==="clusters") {
    const present = Object.keys(monthBuckets).sort();
    keys = selectedClusters.length ? selectedClusters.filter(c=>clusterOptions.includes(c)) : present;
  } else {
    keys = selectedClubs;
  }
  const colorForKey = (k,i) => {
    if (compareMode==="years" && k===currentYear) return "#f51200";
    if (compareMode==="leagues") return LEAGUE_STYLES[k]?.color || COMPARE_PALETTE[i % COMPARE_PALETTE.length];
    if (compareMode==="clusters") return clusterColorFrom(k, clusterOptions);
    return COMPARE_PALETTE[i % COMPARE_PALETTE.length];
  };
  const toggleKey = k => setHiddenKeys(prev=>{ const n=new Set(prev); n.has(k)?n.delete(k):n.add(k); return n; });
  const visibleKeys = keys.filter(k=>!hiddenKeys.has(k));
  const maxVal = Math.max(1, ...visibleKeys.flatMap(k=>seriesForKey(k)).filter(v=>v!=null));

  // YTD comparison only applies to years mode — there's no "prior period" to compare
  // against when comparing clusters/teams within a single chosen year.
  const ytdFor = y => {
    const yearEntries = scoped.filter(e=>new Date(e.ts).getFullYear()===y && new Date(e.ts).getMonth()<=currentMonth);
    if (!yearEntries.length && !scoped.some(e=>new Date(e.ts).getFullYear()===y)) return null;
    return metric==="index" ? avgIndexForEntries(yearEntries) : (metric==="value" ? yearEntries.reduce((s,e)=>s+getRackRate(e),0) : yearEntries.length);
  };
  const thisYTD = compareMode==="years" ? ytdFor(currentYear) : null;
  const lastYTD = compareMode==="years" ? ytdFor(currentYear-1) : null;
  const pctChange = (thisYTD!=null && lastYTD) ? ((thisYTD-lastYTD)/lastYTD*100) : null;

  const W=560, H=190, padL=6, padB=22, padT=8, padR=6;
  const plotW=W-padL-padR, plotH=H-padT-padB;
  const xForMonth = m => padL + (plotW/11)*m;
  const yForVal = v => padT + plotH - (v/maxVal)*plotH;
  const selectStyle={fontFamily:"'DM Sans',sans-serif",fontSize:13,border:"2px solid #E5E7EB",borderRadius:8,padding:"7px 10px",background:"#fff",color:"#111",cursor:"pointer",outline:"none"};
  const keyLabel = compareMode==="years" ? "Year" : compareMode==="leagues" ? "League" : compareMode==="clusters" ? "Cluster" : "Club";

  return (
    <div style={{background:"#fff",border:"1px solid #E5E7EB",borderRadius:16,padding:"24px 28px",boxShadow:"0 1px 6px rgba(0,0,0,.05)"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:12}}>
        <div style={{fontFamily:"'DM Serif Display',serif",fontSize:18,color:"#111827"}}>
          {compareMode==="years"?"Year-over-Year Trend":compareMode==="leagues"?`League Comparison — ${compareYear}`:compareMode==="clusters"?`Cluster Comparison — ${compareYear}`:`Team Comparison — ${compareYear}`}
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <select value={metric} onChange={e=>setMetric(e.target.value)} style={selectStyle}>
            <option value="count">Deliverables Logged</option>
            <option value="value">Total Value Delivered</option>
            <option value="index">Avg CPG Index</option>
          </select>
          {compareMode!=="leagues"&&(
            <select value={leagueScope} onChange={e=>setLeagueScope(e.target.value)} style={selectStyle}>
              <option value="all">All Leagues</option>
              {leagueOptions.map(l=><option key={l} value={l}>{l}</option>)}
            </select>
          )}
        </div>
      </div>

      <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:16,paddingBottom:16,borderBottom:"1px solid #F3F4F6"}}>
        <span style={{fontSize:11,fontWeight:700,color:"#9CA3AF",letterSpacing:.5}}>COMPARE BY</span>
        {[["years","Years"],["leagues","Leagues"],["clusters","Clusters"],["clubs","Teams"]].map(([v,l])=>(
          <button key={v} onClick={()=>{setCompareMode(v);setHiddenKeys(v==="leagues"?new Set(["Expansion"]):new Set());}} style={{padding:"6px 14px",borderRadius:8,border:`1.5px solid ${compareMode===v?"#011e5c":"#E5E7EB"}`,fontFamily:"'DM Sans',sans-serif",fontWeight:compareMode===v?700:400,fontSize:12,cursor:"pointer",background:compareMode===v?"#011e5c":"#fff",color:compareMode===v?"#fff":"#64748B"}}>{l}</button>
        ))}
        {compareMode!=="years"&&(
          <>
            <div style={{width:1,height:20,background:"#E5E7EB",margin:"0 4px"}}/>
            <span style={{fontSize:11,fontWeight:700,color:"#9CA3AF",letterSpacing:.5}}>YEAR</span>
            <select value={compareYear} onChange={e=>setCompareYear(Number(e.target.value))} style={selectStyle}>
              {yearOptions.map(y=><option key={y} value={y}>{y}</option>)}
            </select>
          </>
        )}
        {compareMode==="clusters"&&(
          <MultiSelectFilter label="Clusters" options={clusterOptions} selected={selectedClusters} onChange={setSelectedClusters} width={160}/>
        )}
        {compareMode==="clubs"&&(
          <MultiSelectFilter label="Teams" options={clubOptions} selected={selectedClubs} onChange={setSelectedClubs} width={180}/>
        )}
      </div>

      {compareMode==="clubs"&&selectedClubs.length===0 ? (
        <div style={{textAlign:"center",padding:"32px 0",color:"#9CA3AF",fontSize:13}}>Pick one or more teams above to compare.</div>
      ) : keys.length===0 ? (
        <div style={{textAlign:"center",padding:"32px 0",color:"#9CA3AF",fontSize:13}}>No data for this filter.</div>
      ) : (
        <>
          {compareMode==="years"&&(
            <div style={{display:"flex",alignItems:"center",gap:16,background:"#F8FAFC",borderRadius:10,padding:"12px 16px",marginBottom:16,flexWrap:"wrap"}}>
              <div>
                <div style={{fontSize:10,fontWeight:700,color:"#9CA3AF",letterSpacing:.5}}>{currentYear} YTD (THROUGH {MONTHS_SHORT[currentMonth].toUpperCase()})</div>
                <div style={{fontSize:22,fontWeight:800,color:"#111827",fontFamily:"'DM Serif Display',serif"}}>{thisYTD!=null?fmtMetric(thisYTD):"—"}</div>
              </div>
              {keys.includes(currentYear-1)&&(
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{fontSize:12,color:"#6B7280"}}>vs {currentYear-1} YTD: {fmtMetric(lastYTD)}</span>
                  {pctChange!==null&&(
                    <span style={{fontSize:13,fontWeight:700,color:pctChange>=0?"#047857":"#EF4444",background:pctChange>=0?"#F0FDF4":"#FEF2F2",borderRadius:6,padding:"2px 8px"}}>
                      {pctChange>=0?"▲":"▼"} {Math.abs(pctChange).toFixed(0)}%
                    </span>
                  )}
                </div>
              )}
            </div>
          )}

          <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:10}}>
            {keys.map((k,i)=>{
              const hidden=hiddenKeys.has(k);
              const color=colorForKey(k,i);
              return(
                <button key={k} onClick={()=>toggleKey(k)} style={{display:"flex",alignItems:"center",gap:6,padding:"4px 10px",borderRadius:20,border:`1.5px solid ${hidden?"#E5E7EB":color}`,background:hidden?"#fff":color+"18",cursor:"pointer",opacity:hidden?.6:1}}>
                  <div style={{width:8,height:8,borderRadius:"50%",background:color}}/>
                  <span style={{fontSize:12,fontWeight:600,color:hidden?"#9CA3AF":"#111827"}}>{k}</span>
                </button>
              );
            })}
          </div>

          <svg width="100%" viewBox={`0 0 ${W} ${H}`} style={{overflow:"visible"}}>
            {[0,.25,.5,.75,1].map(f=>(
              <line key={f} x1={padL} x2={W-padR} y1={padT+plotH*(1-f)} y2={padT+plotH*(1-f)} stroke="#F3F4F6" strokeWidth={1}/>
            ))}
            {MONTHS_SHORT.map((m,i)=>(
              <text key={m} x={xForMonth(i)} y={H-4} textAnchor="middle" style={{fontSize:9,fill:"#9CA3AF",fontFamily:"'DM Sans',sans-serif"}}>{m}</text>
            ))}
            {visibleKeys.map(k=>{
              const idx=keys.indexOf(k);
              const color=colorForKey(k,idx);
              const isCurrent=compareMode==="years"&&k===currentYear;
              const segments=buildSegments(seriesForKey(k));
              return (
                <g key={k}>
                  {segments.map((seg,si)=>(
                    <polyline key={si} points={seg.map(p=>`${xForMonth(p.m)},${yForVal(p.v)}`).join(" ")} fill="none" stroke={color} strokeWidth={isCurrent?3:1.5} strokeLinejoin="round" strokeLinecap="round" opacity={isCurrent?1:.8}/>
                  ))}
                  {segments.flat().map(p=>(
                    <circle key={p.m} cx={xForMonth(p.m)} cy={yForVal(p.v)} r={isCurrent?3:2} fill={color} opacity={isCurrent?1:.8}/>
                  ))}
                </g>
              );
            })}
          </svg>

          <button onClick={()=>setShowTable(s=>!s)} style={{marginTop:12,fontSize:12,fontWeight:600,color:"#0369A1",background:"transparent",border:"none",cursor:"pointer",padding:0}}>
            {showTable?"Hide":"Show"} monthly data table {showTable?"▲":"▼"}
          </button>

          {showTable&&(
            <div style={{marginTop:12,overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'DM Sans',sans-serif"}}>
                <thead>
                  <tr>
                    <th style={{padding:"6px 10px",textAlign:"left",fontWeight:700,color:"#6B7280",borderBottom:"1px solid #E5E7EB"}}>{keyLabel}</th>
                    {MONTHS_SHORT.map(m=><th key={m} style={{padding:"6px 10px",textAlign:"right",fontWeight:700,color:"#6B7280",borderBottom:"1px solid #E5E7EB"}}>{m}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {keys.map((k,i)=>(
                    <tr key={k} style={{opacity:hiddenKeys.has(k)?.4:1}}>
                      <td style={{padding:"6px 10px",fontWeight:700,color:colorForKey(k,i)}}>{k}</td>
                      {seriesForKey(k).map((v,m)=><td key={m} style={{padding:"6px 10px",textAlign:"right",color:"#374151"}}>{v>0?fmtMetric(v):"—"}</td>)}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  VALUE BY CLUSTER — nested League → Cluster → Year stacked bars, segments by
//  department. Self-contained filters, same pattern as the YoY trend card.
// ─────────────────────────────────────────────────────────────────────────────

function ValueByClusterCard({ log, clubClusters }) {
  const [metricMode, setMetricMode] = useState("total"); // 'total' | 'avg'
  const [leagueFilter, setLeagueFilter] = useState("all");
  const [hiddenDepts, setHiddenDepts] = useState(new Set());
  const [hover, setHover] = useState(null); // {x,y,dept,value,league,cluster,year}

  const toggleDept = d => setHiddenDepts(prev=>{ const n=new Set(prev); n.has(d)?n.delete(d):n.add(d); return n; });

  // Only external, club-attributed, cluster-assigned entries make sense on this chart.
  const scoped = log.filter(e=>e.type==="External" && e.club && clubClusters[e.club] && (leagueFilter==="all"||e.league===leagueFilter));

  const tree = {}; // league -> cluster -> year -> { byDept:{}, clubs:Set }
  scoped.forEach(e=>{
    const league = e.league||"Unknown", cluster = clubClusters[e.club], year = new Date(e.ts).getFullYear();
    if(!tree[league]) tree[league]={};
    if(!tree[league][cluster]) tree[league][cluster]={};
    if(!tree[league][cluster][year]) tree[league][cluster][year]={byDept:{},clubs:new Set()};
    const bucket = tree[league][cluster][year];
    bucket.byDept[e.dept] = (bucket.byDept[e.dept]||0) + getRackRate(e);
    bucket.clubs.add(e.club);
  });

  const leaguesPresent = Object.keys(tree).sort(sortByLeagueOrder);

  // Flatten into left-to-right bars, tracking group spans for the two-level axis labels.
  const bars = [];
  leaguesPresent.forEach(league=>{
    const clusters = Object.keys(tree[league]).sort();
    clusters.forEach(cluster=>{
      const years = Object.keys(tree[league][cluster]).map(Number).sort((a,b)=>a-b);
      years.forEach(year=>{
        const bucket = tree[league][cluster][year];
        const clubCount = bucket.clubs.size||1;
        const byDept = {};
        Object.entries(bucket.byDept).forEach(([d,v])=>{ byDept[d] = metricMode==="avg" ? v/clubCount : v; });
        bars.push({ league, cluster, year, byDept });
      });
    });
  });

  const deptsPresent = [...new Set(bars.flatMap(b=>Object.keys(b.byDept)))];
  const visibleDepts = deptsPresent.filter(d=>!hiddenDepts.has(d));
  const barTotal = b => visibleDepts.reduce((s,d)=>s+(b.byDept[d]||0),0);
  const maxTotal = Math.max(1, ...bars.map(barTotal));

  // Geometry
  const barW=34, barGap=6, clusterGap=18, leagueGap=32, chartH=260, padT=16, padL=54;
  let x = padL;
  const positioned = bars.map((b,i)=>{
    const prev = bars[i-1];
    if (prev) {
      if (prev.cluster!==b.cluster || prev.league!==b.league) x += (prev.league!==b.league ? leagueGap : clusterGap);
      else x += barGap;
    }
    const pos = { ...b, x };
    x += barW;
    return pos;
  });
  const chartW = x + 20;

  // Group spans for cluster/league labels
  const clusterSpans=[], leagueSpans=[];
  positioned.forEach(b=>{
    const cs = clusterSpans[clusterSpans.length-1];
    if (cs && cs.league===b.league && cs.cluster===b.cluster) { cs.x2=b.x+barW; }
    else clusterSpans.push({ league:b.league, cluster:b.cluster, x1:b.x, x2:b.x+barW });
    const ls = leagueSpans[leagueSpans.length-1];
    if (ls && ls.league===b.league) { ls.x2=b.x+barW; }
    else leagueSpans.push({ league:b.league, x1:b.x, x2:b.x+barW });
  });

  const yForVal = v => padT + chartH - (v/maxTotal)*chartH;
  const fmtAxis = v => v>=1000 ? `$${Math.round(v/1000)}K` : `$${Math.round(v)}`;
  const selectStyle={fontFamily:"'DM Sans',sans-serif",fontSize:13,border:"2px solid #E5E7EB",borderRadius:8,padding:"7px 10px",background:"#fff",color:"#111",cursor:"pointer",outline:"none"};
  const leagueOptions = [...new Set(log.map(e=>e.league))].filter(Boolean).sort();

  return (
    <div style={{background:"#fff",border:"1px solid #E5E7EB",borderRadius:16,padding:"24px 28px",boxShadow:"0 1px 6px rgba(0,0,0,.05)"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6,flexWrap:"wrap",gap:12}}>
        <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
          <div style={{fontFamily:"'DM Serif Display',serif",fontSize:18,color:"#111827"}}>Value Delivered by Cluster</div>
          <span style={{fontSize:11,fontWeight:700,color:"#4338CA",background:"#EEF2FF",border:"1px solid #C7D2FE",borderRadius:20,padding:"2px 10px"}}>{metricMode==="avg"?"Average per club":"Total"}</span>
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <select value={metricMode} onChange={e=>setMetricMode(e.target.value)} style={selectStyle}>
            <option value="total">Total Value</option>
            <option value="avg">Avg Value per Club</option>
          </select>
          <select value={leagueFilter} onChange={e=>setLeagueFilter(e.target.value)} style={selectStyle}>
            <option value="all">All Leagues</option>
            {leagueOptions.map(l=><option key={l} value={l}>{l}</option>)}
          </select>
        </div>
      </div>

      {bars.length===0 ? (
        <div style={{textAlign:"center",padding:"32px 0",color:"#9CA3AF",fontSize:13}}>No cluster-assigned club data for this filter.</div>
      ) : (
        <>
          <div style={{display:"flex",flexWrap:"wrap",gap:8,margin:"14px 0"}}>
            {deptsPresent.map(d=>{
              const hidden=hiddenDepts.has(d);
              const color=DEPT_CONFIG[d]?.color||"#6B7280";
              return(
                <button key={d} onClick={()=>toggleDept(d)} style={{display:"flex",alignItems:"center",gap:6,padding:"4px 10px",borderRadius:20,border:`1.5px solid ${hidden?"#E5E7EB":color}`,background:hidden?"#fff":color+"18",cursor:"pointer",opacity:hidden?.6:1}}>
                  <div style={{width:8,height:8,borderRadius:2,background:color}}/>
                  <span style={{fontSize:12,fontWeight:600,color:hidden?"#9CA3AF":"#111827"}}>{d}</span>
                </button>
              );
            })}
          </div>

          <div style={{position:"relative",overflowX:"auto"}}>
            <svg width={chartW} height={chartH+70} style={{display:"block"}}>
              {[0,.25,.5,.75,1].map(f=>(
                <g key={f}>
                  <line x1={padL-6} x2={chartW-16} y1={padT+chartH*(1-f)} y2={padT+chartH*(1-f)} stroke="#F3F4F6" strokeWidth={1}/>
                  <text x={padL-12} y={padT+chartH*(1-f)+3} textAnchor="end" style={{fontSize:9,fill:"#9CA3AF",fontFamily:"'DM Sans',sans-serif"}}>{fmtAxis(maxTotal*f)}</text>
                </g>
              ))}
              {positioned.map((b,i)=>{
                let yCursor=padT+chartH;
                return (
                  <g key={i}>
                    {visibleDepts.map(d=>{
                      const v=b.byDept[d]||0;
                      if(!v) return null;
                      const h=(v/maxTotal)*chartH;
                      const y=yCursor-h;
                      yCursor=y;
                      const isHovered = hover && hover.x===b.x && hover.dept===d;
                      return (
                        <rect key={d} x={b.x} y={y} width={barW} height={h}
                          fill={DEPT_CONFIG[d]?.color||"#6B7280"}
                          stroke={isHovered?"#111827":"none"} strokeWidth={isHovered?1.5:0}
                          opacity={hover&&!isHovered?.6:1}
                          style={{cursor:"pointer",transition:"opacity .1s"}}
                          onMouseEnter={()=>setHover({x:b.x,y,dept:d,value:v,league:b.league,cluster:b.cluster,year:b.year})}
                          onMouseLeave={()=>setHover(null)}
                        />
                      );
                    })}
                    <text x={b.x+barW/2} y={padT+chartH+13} textAnchor="middle" style={{fontSize:9,fill:"#6B7280",fontFamily:"'DM Sans',sans-serif"}}>{b.year}</text>
                  </g>
                );
              })}
              {clusterSpans.map((cs,i)=>(
                <text key={i} x={(cs.x1+cs.x2)/2} y={padT+chartH+28} textAnchor="middle" style={{fontSize:10,fontWeight:700,fill:"#374151",fontFamily:"'DM Sans',sans-serif"}}>{cs.cluster}</text>
              ))}
              {leagueSpans.map((ls,i)=>(
                <text key={i} x={(ls.x1+ls.x2)/2} y={padT+chartH+44} textAnchor="middle" style={{fontSize:10,fontWeight:700,fill:"#011e5c",fontFamily:"'DM Sans',sans-serif"}}>{ls.league}</text>
              ))}
            </svg>
            {hover&&(
              <div style={{position:"absolute",left:hover.x+barW/2,top:hover.y,transform:"translate(-50%,-100%)",marginTop:-8,background:"#111827",color:"#fff",borderRadius:8,padding:"8px 12px",fontSize:12,fontFamily:"'DM Sans',sans-serif",whiteSpace:"nowrap",pointerEvents:"none",boxShadow:"0 4px 16px rgba(0,0,0,.25)",zIndex:10}}>
                <div style={{display:"flex",alignItems:"center",gap:6,fontWeight:700,marginBottom:2}}>
                  <div style={{width:8,height:8,borderRadius:2,background:DEPT_CONFIG[hover.dept]?.color||"#6B7280"}}/>
                  {hover.dept}
                </div>
                <div style={{fontSize:14,fontWeight:800}}>{fmt$(Math.round(hover.value))}</div>
                <div style={{fontSize:11,color:"#9CA3AF",marginTop:2}}>{hover.cluster} · {hover.league} · {hover.year}</div>
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  DASHBOARD
// ─────────────────────────────────────────────────────────────────────────────

function Dashboard({ log, onExport, clubsByLeague, clubClusters }) {
  const [league,       setLeague]      = useState("all");
  const [cluster,      setCluster]     = useState("all");
  const [filterYear,   setFilterYear]  = useState("all");
  const [filterMonth,  setFilterMonth] = useState("all");
  const [customStart,  setCustomStart] = useState("");
  const [customEnd,    setCustomEnd]   = useState("");
  const [entryType,    setEntryType]   = useState("all");
  const [scope,        setScope]       = useState("all");
  const [metric,       setMetric]      = useState("index");
  const [showAllClubs, setShowAllClubs]= useState(false);

  const LEAGUES = ["all","Championship","League One","Super League","Expansion"];
  const LEAGUE_COLORS = {"Championship":"#b28350","League One":"#00becc","Super League":"#ff8533","Expansion":"#B45309","all":"#011e5c"};
  const DEPT_COLORS = {"Corp Partnerships":"#7C3AED","Marketing":"#0369A1","Consumer Products":"#B45309","Ticketing":"#047857","League Initiatives":"#4338CA","General":"#475569"};
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const availableYears = [...new Set((log||[]).map(e=>new Date(e.ts).getFullYear()))].sort((a,b)=>b-a);

  // Filters
  const hasCustomRange = !!(customStart || customEnd);
  const timeFiltered = (log||[]).filter(e=>{
    if (hasCustomRange) {
      if (customStart && e.ts < new Date(customStart+"T00:00:00").getTime()) return false;
      if (customEnd && e.ts > new Date(customEnd+"T23:59:59.999").getTime()) return false;
      return true;
    }
    const d=new Date(e.ts);
    if(filterYear!=="all"&&d.getFullYear()!==parseInt(filterYear)) return false;
    if(filterMonth!=="all"&&d.getMonth()!==parseInt(filterMonth)) return false;
    return true;
  });
  const typeFiltered = timeFiltered.filter(e=>{
    if(entryType==="strategic") return !e.recurring;
    if(entryType==="recurring") return e.recurring;
    return true;
  }).filter(e=>{
    if(scope==="external") return e.type==="External";
    if(scope==="internal") return e.type==="Internal";
    return true;
  });
  // Cluster narrows by club, independent of league — lets "all Cluster 1 clubs across leagues"
  // and "Championship Cluster 1 only" both work depending on whether League is also set.
  const clusterFiltered = cluster==="all" ? typeFiltered : typeFiltered.filter(e=>clubClusters[e.club]===cluster);
  // leagueFiltered is the fully-combined set (time + type + scope + cluster + league) and is
  // the base for every KPI/table below — everything filters together, as it should.
  const leagueFiltered = league==="all" ? clusterFiltered : clusterFiltered.filter(e=>e.league===league);
  const strategicLog = leagueFiltered.filter(e=>!e.recurring);
  const recurringLog  = leagueFiltered.filter(e=>e.recurring);
  const externalLog = leagueFiltered.filter(e=>e.type==="External");
  // Attribution log: clubs for external work, recipients for internal work — pivots the
  // "engaged" KPI and Top Clubs/Recipients table based on which scope is selected.
  const attributionLog = scope==="internal" ? leagueFiltered.filter(e=>e.type==="Internal") : externalLog;
  const attributionLabel = scope==="internal" ? "Recipients" : "Clubs";
  const clusterList = [...new Set(Object.values(clubClusters))].filter(Boolean).sort();
  // leagueOnlyFiltered mirrors clusterFiltered but the other way around: respects League,
  // ignores Cluster — the base for the "by Cluster" view of the league-shaped charts below.
  const leagueOnlyFiltered = league==="all" ? typeFiltered : typeFiltered.filter(e=>e.league===league);
  // Shared correct math for "average CPG index": average each club's own average first,
  // then average those together. Summing every entry's raw score and dividing by club
  // count is NOT bounded to the 1-5 scale once any club has more than one entry.
  const avgIndexPerClub = entries => {
    const perClub = {};
    entries.forEach(e=>{
      const c=e.club||"Unknown";
      if(!perClub[c]) perClub[c]={sum:0,count:0};
      perClub[c].sum+=Number(e.index_score||1); perClub[c].count++;
    });
    const clubAvgs = Object.values(perClub).map(c=>c.sum/c.count);
    return clubAvgs.length>0 ? clubAvgs.reduce((s,v)=>s+v,0)/clubAvgs.length : null;
  };
  const [chartAxis, setChartAxis] = useState("league");
  useEffect(()=>{ setChartAxis(cluster==="all" ? "league" : "cluster"); }, [cluster]);
  const resetFilters = () => { setLeague("all"); setCluster("all"); setFilterYear("all"); setFilterMonth("all"); setCustomStart(""); setCustomEnd(""); setEntryType("all"); setScope("all"); setMetric("index"); setChartAxis("league"); };

  // Stats
  const total     = leagueFiltered.length;
  const strategic = strategicLog.length;
  const recurring = recurringLog.length;
  const extCnt    = leagueFiltered.filter(e=>e.type==="External").length;
  const intCnt    = leagueFiltered.filter(e=>e.type==="Internal").length;
  const stratIdx  = leagueFiltered.filter(e=>!e.recurring&&e.type==="External");
  const avgIndexVal = avgIndexPerClub(stratIdx);
  const avgIndex  = avgIndexVal!==null ? avgIndexVal.toFixed(1) : "—";

  // Charts — by-league breakdowns intentionally use clusterFiltered (not leagueFiltered) so
  // selecting a single league doesn't collapse these into a trivial one-bar chart; they still
  // respect Cluster, Type, and Scope, just not the League toggle itself.
  const byLeague = ["Championship","League One","Super League","Expansion"].map(l=>{
    const leagueEntries = clusterFiltered.filter(e=>e.league===l);
    const externalLeagueEntries = leagueEntries.filter(e=>e.type==="External");
    const leagueClubs = new Set(externalLeagueEntries.map(e=>e.club));
    const clubDivisor = leagueClubs.size || 1;
    let sub;
    if (metric==="value") {
      const totalVal = leagueEntries.reduce((s,e)=>s+getRackRate(e),0);
      sub = leagueClubs.size>0 ? `(avg ${fmt$(Math.round(totalVal/clubDivisor))})` : "(avg —)";
    } else {
      const stratExt = externalLeagueEntries.filter(e=>!e.recurring);
      const avgVal = avgIndexPerClub(stratExt);
      sub = (leagueClubs.size>0 && avgVal!==null)
        ? `(avg ${avgVal.toFixed(1)})`
        : "(avg —)";
    }
    return { label:l, value:leagueEntries.length, color:LEAGUE_COLORS[l], sub };
  }).filter(d=>d.value>0);

  // Rule: once a specific real cluster is selected (not "All Clusters"), never show
  // Unassigned/No Cluster as a bucket alongside it — unless that's the exact thing picked.
  const isUnassignedLike = cl => cl==="No Cluster" || cl==="Unassigned";
  const CLUSTER_BUCKETS = [...clusterList, "No Cluster"].filter(cl =>
    !isUnassignedLike(cl) || cluster==="all" || cluster===cl
  );
  const clusterColor = cl => clusterColorFrom(cl, clusterList);
  const matchesCluster = (e,cl) => cl==="No Cluster" ? !clubClusters[e.club] : clubClusters[e.club]===cl;

  const byCluster = CLUSTER_BUCKETS.map(cl=>{
    const clEntries = leagueOnlyFiltered.filter(e=>matchesCluster(e,cl));
    const clExternal = clEntries.filter(e=>e.type==="External");
    const clClubs = new Set(clExternal.map(e=>e.club));
    const clubDivisor = clClubs.size || 1;
    let sub;
    if (metric==="value") {
      const totalVal = clEntries.reduce((s,e)=>s+getRackRate(e),0);
      sub = clClubs.size>0 ? `(avg ${fmt$(Math.round(totalVal/clubDivisor))})` : "(avg —)";
    } else {
      const stratExt = clExternal.filter(e=>!e.recurring);
      const avgVal = avgIndexPerClub(stratExt);
      sub = (clClubs.size>0 && avgVal!==null)
        ? `(avg ${avgVal.toFixed(1)})`
        : "(avg —)";
    }
    return { label:cl, value:clEntries.length, color:clusterColor(cl), sub };
  }).filter(d=>d.value>0);

  const usageByCluster = CLUSTER_BUCKETS.map(cl=>({
    label:cl, value:leagueOnlyFiltered.filter(e=>matchesCluster(e,cl)).length, color:clusterColor(cl)
  })).filter(d=>d.value>0);

  const deptCounts = {};
  leagueFiltered.forEach(e=>{ const k=e.dept||"General"; deptCounts[k]=(deptCounts[k]||0)+1; });
  const byDept = Object.entries(deptCounts).sort((a,b)=>b[1]-a[1]).map(([dept,count])=>({label:dept,value:count,color:DEPT_COLORS[dept]||"#475569"}));

  const usageByLeague = ["Championship","League One","Super League","Expansion"].map(l=>({
    label:l, value:clusterFiltered.filter(e=>e.league===l).length, color:LEAGUE_COLORS[l]
  })).filter(d=>d.value>0);

  // Club/recipient map — pivots to internal recipients when scope==="internal"
  const clubMap={};
  attributionLog.forEach(e=>{
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
          <button onClick={resetFilters} style={{marginLeft:"auto",padding:"6px 14px",border:"1.5px solid #0a2d6e",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:12,cursor:"pointer",background:"transparent",color:"#F87171",transition:"all .15s",whiteSpace:"nowrap"}}>↺ Reset Filters</button>
        </div>

        {clusterList.length>0&&(<>
        <div style={{height:1,background:"#0a2d6e"}}/>
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          <span style={{fontSize:9,fontWeight:700,letterSpacing:1.2,color:"#4a6fa8",width:52,flexShrink:0}}>CLUSTER</span>
          <button onClick={()=>setCluster("all")} style={{padding:"6px 14px",border:`1.5px solid ${cluster==="all"?"#fff":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:cluster==="all"?700:400,fontSize:12,cursor:"pointer",background:cluster==="all"?"#fff":"transparent",color:cluster==="all"?"#011e5c":"#64748B",transition:"all .15s",whiteSpace:"nowrap"}}>All Clusters</button>
          {clusterList.map(c=>(
            <button key={c} onClick={()=>setCluster(c)} style={{padding:"6px 14px",border:`1.5px solid ${cluster===c?"#fff":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:cluster===c?700:400,fontSize:12,cursor:"pointer",background:cluster===c?"#fff":"transparent",color:cluster===c?"#011e5c":"#64748B",transition:"all .15s",whiteSpace:"nowrap"}}>
              {c}
            </button>
          ))}
        </div>
        </>)}

        <div style={{height:1,background:"#0a2d6e"}}/>

        {/* Row 2: Period */}
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          <span style={{fontSize:9,fontWeight:700,letterSpacing:1.2,color:"#4a6fa8",width:52,flexShrink:0}}>PERIOD</span>
          <button onClick={()=>{setFilterYear("all");setFilterMonth("all");setCustomStart("");setCustomEnd("");}} style={{padding:"6px 12px",border:`1.5px solid ${(!hasCustomRange&&filterYear==="all")?"#f51200":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:(!hasCustomRange&&filterYear==="all")?700:400,fontSize:12,cursor:"pointer",background:(!hasCustomRange&&filterYear==="all")?"#f51200":"transparent",color:(!hasCustomRange&&filterYear==="all")?"#fff":"#64748B",transition:"all .15s"}}>All Time</button>
          {availableYears.map(y=>(
            <button key={y} onClick={()=>{setFilterYear(String(y));setFilterMonth("all");setCustomStart("");setCustomEnd("");}} style={{padding:"6px 12px",border:`1.5px solid ${(!hasCustomRange&&filterYear===String(y))?"#f51200":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:(!hasCustomRange&&filterYear===String(y))?700:400,fontSize:12,cursor:"pointer",background:(!hasCustomRange&&filterYear===String(y))?"#f51200":"transparent",color:(!hasCustomRange&&filterYear===String(y))?"#fff":"#64748B",transition:"all .15s"}}>{y}</button>
          ))}
          {!hasCustomRange&&filterYear!=="all"&&(
            <>
              <div style={{width:1,height:16,background:"#0a2d6e",margin:"0 4px"}}/>
              <button onClick={()=>setFilterMonth("all")} style={{padding:"5px 10px",border:`1.5px solid ${filterMonth==="all"?"#fff":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:filterMonth==="all"?700:400,fontSize:11,cursor:"pointer",background:filterMonth==="all"?"#fff":"transparent",color:filterMonth==="all"?"#011e5c":"#64748B",transition:"all .15s"}}>All</button>
              {MONTHS.map((m,i)=>(
                <button key={m} onClick={()=>setFilterMonth(String(i))} style={{padding:"5px 10px",border:`1.5px solid ${filterMonth===String(i)?"#fff":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:filterMonth===String(i)?700:400,fontSize:11,cursor:"pointer",background:filterMonth===String(i)?"#fff":"transparent",color:filterMonth===String(i)?"#011e5c":"#64748B",transition:"all .15s"}}>{m}</button>
              ))}
            </>
          )}
          <div style={{width:1,height:16,background:"#0a2d6e",margin:"0 4px"}}/>
          <span style={{fontSize:11,fontWeight:hasCustomRange?700:400,color:hasCustomRange?"#fff":"#64748B"}}>Custom range:</span>
          <input type="date" value={customStart} onChange={e=>{setCustomStart(e.target.value);setFilterYear("all");setFilterMonth("all");}}
            style={{padding:"5px 8px",border:`1.5px solid ${hasCustomRange?"#fff":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontSize:12,background:"#fff",color:"#111",outline:"none"}}/>
          <span style={{fontSize:12,color:"#64748B"}}>to</span>
          <input type="date" value={customEnd} onChange={e=>{setCustomEnd(e.target.value);setFilterYear("all");setFilterMonth("all");}}
            style={{padding:"5px 8px",border:`1.5px solid ${hasCustomRange?"#fff":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontSize:12,background:"#fff",color:"#111",outline:"none"}}/>
          {hasCustomRange&&<button onClick={()=>{setCustomStart("");setCustomEnd("");}} style={{padding:"5px 10px",border:"1.5px solid #0a2d6e",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:600,fontSize:11,cursor:"pointer",background:"transparent",color:"#F87171"}}>Clear</button>}
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
          <span style={{fontSize:9,fontWeight:700,letterSpacing:1.2,color:"#4a6fa8",flexShrink:0}}>SCOPE</span>
          {[["all","All"],["external","External"],["internal","Internal"]].map(([v,l])=>(
            <button key={v} onClick={()=>setScope(v)} style={{padding:"6px 14px",border:`1.5px solid ${scope===v?"#f51200":"#0a2d6e"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:scope===v?700:400,fontSize:12,cursor:"pointer",background:scope===v?"#f51200":"transparent",color:scope===v?"#fff":"#64748B",transition:"all .15s"}}>
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
        const totalValue=leagueFiltered.reduce((s,e)=>s+getRackRate(e),0);
        const stratValue=strategicLog.reduce((s,e)=>s+getRackRate(e),0);
        const recurValue=recurringLog.reduce((s,e)=>s+getRackRate(e),0);
        const isValue=metric==="value";
        const engagedCount=Object.keys(clubMap).length;
        const avgValuePerEntity=engagedCount>0?Math.round(totalValue/engagedCount):0;
        const scopeNoun = scope==="internal" ? "recipient" : "club";
        const kpis=isValue?[
          {label:"TOTAL VALUE DELIVERED",value:"$"+Math.round(totalValue).toLocaleString(),color:"#111827",sub:`${total} deliverables`},
          {label:"STRATEGIC VALUE",value:"$"+Math.round(stratValue).toLocaleString(),color:"#4338CA",sub:`${strategic} deliverables`},
          {label:"RECURRING VALUE",value:"$"+Math.round(recurValue).toLocaleString(),color:"#854D0E",sub:`${recurring} deliverables`},
          {label:`AVG VALUE PER ${scopeNoun.toUpperCase()}`,value:engagedCount>0?"$"+avgValuePerEntity.toLocaleString():"—",color:"#0369A1",sub:`${engagedCount} ${scopeNoun}${engagedCount===1?"":"s"} engaged`},
          {label:`${attributionLabel.toUpperCase()} ENGAGED`,value:engagedCount.toLocaleString(),color:"#047857",sub:league==="all"?"all leagues":league},
        ]:[
          {label:"TOTAL DELIVERABLES",value:total.toLocaleString(),color:"#111827",sub:`${extCnt} external · ${intCnt} internal`},
          {label:"STRATEGIC",value:strategic.toLocaleString(),color:"#4338CA",sub:`${total?Math.round(strategic/total*100):0}% of total`},
          {label:"RECURRING",value:recurring.toLocaleString(),color:"#854D0E",sub:`${total?Math.round(recurring/total*100):0}% of total`},
          {label:"AVG CPG INDEX",value:avgIndex,color:"#047857",sub:"strategic external"},
          {label:`${attributionLabel.toUpperCase()} ENGAGED`,value:engagedCount.toLocaleString(),color:"#0369A1",sub:league==="all"?"all leagues":league},
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
      {clusterList.length>0&&(
        <div style={{display:"flex",justifyContent:"flex-end",gap:6}}>
          <span style={{fontSize:11,fontWeight:700,color:"#9CA3AF",letterSpacing:.5,alignSelf:"center",marginRight:4}}>VIEW BY</span>
          {[["league","League"],["cluster","Cluster"]].map(([v,l])=>(
            <button key={v} onClick={()=>setChartAxis(v)} style={{padding:"5px 12px",border:`1.5px solid ${chartAxis===v?"#011e5c":"#E5E7EB"}`,borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:chartAxis===v?700:400,fontSize:12,cursor:"pointer",background:chartAxis===v?"#011e5c":"#fff",color:chartAxis===v?"#fff":"#64748B"}}>{l}</button>
          ))}
        </div>
      )}
      <div style={{display:"grid",gridTemplateColumns:"1.4fr 1fr",gap:16}}>
        <Card title={(()=>{
          const isCluster = chartAxis==="cluster";
          const base = isCluster ? leagueOnlyFiltered : clusterFiltered;
          const allExternal = base.filter(e=>e.type==="External");
          const allClubs = new Set(allExternal.map(e=>e.club));
          const divisor = allClubs.size || 1;
          let macroSub = "avg —";
          if (allClubs.size>0) {
            if (metric==="value") {
              const totalVal = base.reduce((s,e)=>s+getRackRate(e),0);
              macroSub = `avg ${fmt$(Math.round(totalVal/divisor))}`;
            } else {
              const stratExt = allExternal.filter(e=>!e.recurring);
              const avgVal = avgIndexPerClub(stratExt);
              if (avgVal!==null) macroSub = `avg ${avgVal.toFixed(1)}`;
            }
          }
          return <span>Total Deliverables by {isCluster?"Cluster":"League"} <span style={{fontSize:13,fontWeight:400,color:"#9CA3AF"}}>({macroSub} per club)</span></span>;
        })()}>
          {(chartAxis==="cluster"?byCluster:byLeague).length>0?<HBarChart data={chartAxis==="cluster"?byCluster:byLeague} color="#1D4ED8" height={52}/>:
            <div style={{textAlign:"center",padding:"32px 0",color:"#9CA3AF",fontSize:13}}>No data for this filter.</div>}
        </Card>
        <Card title={`CPG Usage by ${chartAxis==="cluster"?"Cluster":"League"}`}>
          {(chartAxis==="cluster"?usageByCluster:usageByLeague).length>0?<DonutChart slices={chartAxis==="cluster"?usageByCluster:usageByLeague} size={200} title={`${(chartAxis==="cluster"?usageByCluster:usageByLeague).reduce((s,d)=>s+d.value,0)}`} subtitle="total"/>:
            <div style={{textAlign:"center",padding:"32px 0",color:"#9CA3AF",fontSize:13}}>No data.</div>}
        </Card>
      </div>

      {/* Row 2: Vertical pie + Year-over-Year Trend */}
      <div style={{display:"grid",gridTemplateColumns:"minmax(280px,340px) 1fr",gap:16}}>
        <Card title={`CPG Engagement by Vertical${league!=="all"?` — ${league}`:""}`}>
          {byDept.length>0?<DonutChart slices={byDept} size={170} title={`${leagueFiltered.length}`} subtitle="deliverables"/>:
            <div style={{textAlign:"center",padding:"32px 0",color:"#9CA3AF",fontSize:13}}>No data for this filter.</div>}
        </Card>
        <YoYTrendCard log={log} clubClusters={clubClusters}/>
      </div>

      {/* Row 2b: Value by Cluster */}
      <ValueByClusterCard log={log} clubClusters={clubClusters}/>

      {/* Row 3: Top Clubs / Recipients */}
      <Card
        title={`Top ${attributionLabel} — ${metric==="value"?"Value Delivered":"CPG Index Score"}`}
        action={metric==="value"&&<span style={{fontSize:11,color:"#F59E0B",background:"#FEF9C3",border:"1px solid #FDE68A",borderRadius:6,padding:"2px 8px",fontWeight:600}}>Conservative est. rates</span>}
      >
        {clubRows.length===0?(
          <div style={{textAlign:"center",padding:"32px 0",color:"#9CA3AF",fontSize:13}}>No {scope==="internal"?"internal":"external"} entries for this filter.</div>
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
            const en=leagueFiltered.filter(e=>e.dept===dept);
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


function ActivityExplorer({ log, onRemove, onExportView, clubClusters={} }) {
  const [view,setView]=useState("club");
  const [expandedClub, setExpandedClub]=useState(null);
  const [expandedLeague, setExpandedLeague]=useState(null);
  const [expandedDept, setExpandedDept]=useState(null);
  const [expandedStaff, setExpandedStaff]=useState(null);
  const [staffFilter, setStaffFilter]=useState([]);
  const [deliverableFilter, setDeliverableFilter]=useState([]);
  const [clubFilter, setClubFilter]=useState([]);
  const [leagueFilter, setLeagueFilter]=useState([]);
  const [typeFilter, setTypeFilter]=useState([]);
  const [deptFilter, setDeptFilter]=useState([]);
  const [sortKey, setSortKey]=useState("value");
  const [sortDir, setSortDir]=useState("desc");

  const staffOptions = [...new Set(log.map(e=>e.staff))].filter(Boolean).sort();
  const deliverableOptions = [...new Set(log.map(e=>e.name))].filter(Boolean).sort();
  const clubOptions = [...new Set(log.map(e=>e.club))].filter(Boolean).sort();
  const leagueOptions = [...new Set(log.map(e=>e.league))].filter(Boolean).sort();
  const typeOptions = [...new Set(log.map(e=>e.type))].filter(Boolean).sort();
  const deptOptions = [...new Set(log.map(e=>e.dept))].filter(Boolean).sort();

  const hasActiveFilters = staffFilter.length||deliverableFilter.length||clubFilter.length||leagueFilter.length||typeFilter.length||deptFilter.length;
  const clearAllFilters = () => { setStaffFilter([]);setDeliverableFilter([]);setClubFilter([]);setLeagueFilter([]);setTypeFilter([]);setDeptFilter([]); };

  const filteredLog = log
    .filter(e=>!staffFilter.length||staffFilter.includes(e.staff))
    .filter(e=>!deliverableFilter.length||deliverableFilter.includes(e.name))
    .filter(e=>!clubFilter.length||clubFilter.includes(e.club))
    .filter(e=>!leagueFilter.length||leagueFilter.includes(e.league))
    .filter(e=>!typeFilter.length||typeFilter.includes(e.type))
    .filter(e=>!deptFilter.length||deptFilter.includes(e.dept));

  const total=filteredLog.reduce((s,e)=>s+e.rate,0);

  // Column definitions per view — drives both the clickable headers and the sort logic.
  const VIEW_COLS = {
    club:   [["club","Club","text"],["league","League","text"],["cluster","Cluster","text"],["count","# Logged","num"],["indexAvg","Avg Index","num"],["value","Total Value","num"]],
    league: [["league","League","text"],["clubCount","Clubs Engaged","num"],["count","# Logged","num"],["indexAvg","Avg Index","num"],["value","Total Value","num"]],
    dept:   [["dept","Department","text"],["count","# Logged","num"],["value","Total Value","num"],["ext","External","num"],["int","Internal","num"]],
    staff:  [["staff","Staff Member","text"],["count","# Logged","num"],["indexAvg","Avg Index","num"],["value","Total Value","num"],["ext","External","num"],["int","Internal","num"]],
  };
  const handleSort = (key, type) => {
    if (key===sortKey) { setSortDir(d=>d==="asc"?"desc":"asc"); }
    else { setSortKey(key); setSortDir(type==="text"?"asc":"desc"); }
  };
  const sortRows = rows => [...rows].sort((a,b)=>{
    const col = (VIEW_COLS[view]||[]).find(c=>c[0]===sortKey);
    if (col && col[2]==="text") {
      const av=String(a[sortKey]||"").toLowerCase(), bv=String(b[sortKey]||"").toLowerCase();
      return sortDir==="asc" ? av.localeCompare(bv) : bv.localeCompare(av);
    }
    const av=parseFloat(a[sortKey])||0, bv=parseFloat(b[sortKey])||0;
    return sortDir==="asc" ? av-bv : bv-av;
  });
  const SortTH = ({colKey,label,type,right}) => (
    <th onClick={()=>handleSort(colKey,type)} style={{padding:"11px 14px",fontWeight:700,fontSize:11,letterSpacing:.5,color:sortKey===colKey?"#011e5c":"#6B7280",borderBottom:"1px solid #E5E7EB",background:sortKey===colKey?"#F5F3FF":"#F9FAFB",textAlign:right?"right":"left",cursor:"pointer",userSelect:"none",whiteSpace:"nowrap"}}>
      {label}{sortKey===colKey&&<span style={{marginLeft:4,fontSize:10}}>{sortDir==="asc"?"▲":"▼"}</span>}
    </th>
  );

  // By Club
  const clubMap={};
  filteredLog.filter(e=>e.type==="External").forEach(e=>{
    const c=e.club||"Unknown";
    if(!clubMap[c])clubMap[c]={club:c,league:e.league||"",cluster:clubClusters[c]||"—",count:0,value:0,indexSum:0,indexCount:0,entries:[]};
    clubMap[c].count++;clubMap[c].value+=e.rate;
    clubMap[c].entries.push(e);
    if(!e.recurring){clubMap[c].indexSum+=Number(e.index_score||1);clubMap[c].indexCount++;}
  });
  Object.values(clubMap).forEach(c=>{c.indexAvg=c.indexCount>0?(c.indexSum/c.indexCount).toFixed(1):"—";});
  const clubRows=sortRows(Object.values(clubMap));

  // By League
  const leagueMap={};
  filteredLog.filter(e=>e.type==="External").forEach(e=>{
    const l=e.league||"Unknown";
    if(!leagueMap[l])leagueMap[l]={league:l,count:0,value:0,indexSum:0,indexCount:0,clubs:new Set(),entries:[]};
    leagueMap[l].count++;leagueMap[l].value+=e.rate;leagueMap[l].clubs.add(e.club);
    leagueMap[l].entries.push(e);
    if(!e.recurring){leagueMap[l].indexSum+=Number(e.index_score||1);leagueMap[l].indexCount++;}
  });
  Object.values(leagueMap).forEach(l=>{l.indexAvg=l.indexCount>0?(l.indexSum/l.indexCount).toFixed(1):"—";l.clubCount=l.clubs.size;});
  const leagueRows=sortRows(Object.values(leagueMap));

  // By Dept
  const deptMap={};
  filteredLog.forEach(e=>{
    if(!deptMap[e.dept])deptMap[e.dept]={dept:e.dept,count:0,value:0,ext:0,int:0,entries:[]};
    deptMap[e.dept].count++;deptMap[e.dept].value+=e.rate;
    deptMap[e.dept].entries.push(e);
    if(e.type==="External")deptMap[e.dept].ext++;else deptMap[e.dept].int++;
  });
  const deptRows=sortRows(Object.values(deptMap));

  // By Staff
  const staffMap={};
  filteredLog.forEach(e=>{
    if(!staffMap[e.staff])staffMap[e.staff]={staff:e.staff,count:0,value:0,ext:0,int:0,indexSum:0,indexCount:0,entries:[]};
    staffMap[e.staff].count++;staffMap[e.staff].value+=e.rate;
    staffMap[e.staff].entries.push(e);
    if(e.type==="External")staffMap[e.staff].ext++;else staffMap[e.staff].int++;
    if(!e.recurring&&e.type==="External"){staffMap[e.staff].indexSum+=Number(e.index_score||1);staffMap[e.staff].indexCount++;}
  });
  Object.values(staffMap).forEach(s=>{s.indexAvg=s.indexCount>0?(s.indexSum/s.indexCount).toFixed(1):"—";});
  const staffRows=sortRows(Object.values(staffMap));

  const tBtn=active=>({padding:"8px 18px",border:"none",borderRadius:8,fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer",transition:"all .15s",background:active?"#011e5c":"transparent",color:active?"#fff":"#6B7280"});

  const [selected,setSelected]=useState(new Set());
  const [confirm,setConfirm]=useState(false);
  const reversed=[...filteredLog].reverse();
  const allSelected=selected.size===filteredLog.length&&filteredLog.length>0;
  const toggle=id=>{setSelected(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n;});setConfirm(false);};
  const toggleAll=()=>{setSelected(allSelected?new Set():new Set(filteredLog.map(e=>e.id)));setConfirm(false);};
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
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:12}}>
        <div style={{display:"flex",background:"#F1F5F9",borderRadius:12,padding:4,gap:2}}>
          {[["club","By Club"],["league","By League"],["dept","By Department"],["staff","By Staff"],["log","Full Log"]].map(([v,label])=>(
            <button key={v} onClick={()=>{setView(v);setExpandedClub(null);setExpandedLeague(null);setExpandedDept(null);setExpandedStaff(null);setSortKey("value");setSortDir("desc");}} style={tBtn(view===v)}>{label}</button>
          ))}
        </div>
        <button onClick={()=>onExportView(view,{clubRows,deptRows,staffRows,log:filteredLog})} style={{background:"#0369A1",color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",fontFamily:"'DM Sans',sans-serif",fontWeight:700,fontSize:13,cursor:"pointer"}}>⬇ Export This View</button>
      </div>

      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:20,flexWrap:"wrap"}}>
        <span style={{fontSize:11,fontWeight:700,color:"#9CA3AF",letterSpacing:.5}}>FILTER</span>
        <MultiSelectFilter label="Staff" options={staffOptions} selected={staffFilter} onChange={setStaffFilter} width={150}/>
        <MultiSelectFilter label="Deliverables" options={deliverableOptions} selected={deliverableFilter} onChange={setDeliverableFilter} width={170}/>
        <MultiSelectFilter label="Clubs" options={clubOptions} selected={clubFilter} onChange={setClubFilter} width={150}/>
        <MultiSelectFilter label="Leagues" options={leagueOptions} selected={leagueFilter} onChange={setLeagueFilter} width={140}/>
        <MultiSelectFilter label="Types" options={typeOptions} selected={typeFilter} onChange={setTypeFilter} width={120}/>
        <MultiSelectFilter label="Depts" options={deptOptions} selected={deptFilter} onChange={setDeptFilter} width={150}/>
        {hasActiveFilters?(
          <button onClick={clearAllFilters} style={{fontSize:12,fontWeight:600,color:"#EF4444",background:"transparent",border:"1.5px solid #FECACA",borderRadius:8,padding:"6px 12px",cursor:"pointer"}}>Clear all filters</button>
        ):null}
      </div>

      {view==="club"&&(
        <TblWrap>
          <thead><tr>
            <SortTH colKey="club" label="Club" type="text"/>
            <SortTH colKey="league" label="League" type="text"/>
            <SortTH colKey="cluster" label="Cluster" type="text"/>
            <SortTH colKey="count" label="# Logged" type="num" right/>
            <SortTH colKey="indexAvg" label="Avg Index" type="num" right/>
            <SortTH colKey="value" label="Total Value" type="num" right/>
          </tr></thead>
          <tbody>
            {clubRows.map(c=>{
              const isExpanded=expandedClub===c.club;
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
                    <TD color="#6B7280">{c.cluster}</TD>
                    <TD right>{c.count}</TD>
                    <TD right><span style={{background:"#EEF2FF",color:"#4338CA",borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700}}>{c.indexAvg}</span></TD>
                    <TD right bold>{fmt$(c.value)}</TD>
                  </tr>
                  {isExpanded&&(
                    <tr key={`${c.club}-expand`}>
                      <td colSpan={6} style={{padding:"0 0 4px 0",background:"#F8F5FF"}}>
                        <ExpandPanel heading={`DELIVERABLES LOGGED FOR ${c.club.toUpperCase()}`}>
                          <LineItemTable entries={c.entries} hide={["club"]} clubClusters={clubClusters}/>
                        </ExpandPanel>
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
          <thead><tr>
            <SortTH colKey="league" label="League" type="text"/>
            <SortTH colKey="clubCount" label="Clubs Engaged" type="num" right/>
            <SortTH colKey="count" label="# Logged" type="num" right/>
            <SortTH colKey="indexAvg" label="Avg Index" type="num" right/>
            <SortTH colKey="value" label="Total Value" type="num" right/>
          </tr></thead>
          <tbody>
            {leagueRows.map(l=>{
              const ls=LEAGUE_STYLES[l.league]||LEAGUE_STYLES["League-wide"];
              const isExpanded=expandedLeague===l.league;
              return(
                <Fragment key={l.league}>
                  <tr onClick={()=>setExpandedLeague(isExpanded?null:l.league)} style={{cursor:"pointer",background:isExpanded?"#F5F3FF":undefined}}>
                    <TD bold>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontSize:10,color:"#9CA3AF",transition:"transform .15s",display:"inline-block",transform:isExpanded?"rotate(90deg)":"rotate(0deg)"}}>▶</span>
                        <LeagueBadge league={l.league}/>
                      </div>
                    </TD>
                    <TD right color="#6B7280">{l.clubCount} clubs</TD>
                    <TD right>{l.count}</TD>
                    <TD right><span style={{background:"#EEF2FF",color:"#4338CA",borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700}}>{l.indexAvg}</span></TD>
                    <TD right bold color={ls.color}>{fmt$(l.value)}</TD>
                  </tr>
                  {isExpanded&&(
                    <tr key={`${l.league}-expand`}>
                      <td colSpan={5} style={{padding:"0 0 4px 0",background:"#F8F5FF"}}>
                        <ExpandPanel heading={`DELIVERABLES LOGGED FOR ${l.league.toUpperCase()}`}>
                          <LineItemTable entries={l.entries} hide={["league"]} clubClusters={clubClusters}/>
                        </ExpandPanel>
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
          </tbody>
        </TblWrap>
      )}

      {view==="dept"&&(
        <TblWrap>
          <thead><tr>
            <SortTH colKey="dept" label="Department" type="text"/>
            <SortTH colKey="count" label="# Logged" type="num" right/>
            <SortTH colKey="value" label="Total Value" type="num" right/>
            <SortTH colKey="ext" label="External" type="num" right/>
            <SortTH colKey="int" label="Internal" type="num" right/>
          </tr></thead>
          <tbody>
            {deptRows.map(d=>{
              const cfg=DEPT_CONFIG[d.dept];
              const isExpanded=expandedDept===d.dept;
              return(
                <Fragment key={d.dept}>
                  <tr onClick={()=>setExpandedDept(isExpanded?null:d.dept)} style={{cursor:"pointer",background:isExpanded?"#F5F3FF":undefined}}>
                    <TD bold>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontSize:10,color:"#9CA3AF",transition:"transform .15s",display:"inline-block",transform:isExpanded?"rotate(90deg)":"rotate(0deg)"}}>▶</span>
                        <span>{cfg?.emoji}</span>{d.dept}
                      </div>
                    </TD>
                    <TD right>{d.count}</TD>
                    <TD right bold color={cfg?.color||"#374151"}>{fmt$(d.value)}</TD>
                    <TD right color="#0369A1">{d.ext}</TD>
                    <TD right color="#047857">{d.int}</TD>
                  </tr>
                  {isExpanded&&(
                    <tr key={`${d.dept}-expand`}>
                      <td colSpan={5} style={{padding:"0 0 4px 0",background:"#F8F5FF"}}>
                        <ExpandPanel heading={`DELIVERABLES LOGGED FOR ${d.dept.toUpperCase()}`}>
                          <LineItemTable entries={d.entries} hide={["dept"]} clubClusters={clubClusters}/>
                        </ExpandPanel>
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
          </tbody>
        </TblWrap>
      )}

      {view==="staff"&&(
        <TblWrap>
          <thead><tr>
            <SortTH colKey="staff" label="Staff Member" type="text"/>
            <SortTH colKey="count" label="# Logged" type="num" right/>
            <SortTH colKey="indexAvg" label="Avg Index" type="num" right/>
            <SortTH colKey="value" label="Total Value" type="num" right/>
            <SortTH colKey="ext" label="External" type="num" right/>
            <SortTH colKey="int" label="Internal" type="num" right/>
          </tr></thead>
          <tbody>
            {staffRows.map(s=>{
              const initials=s.staff.split(" ").map(n=>n[0]).join("").slice(0,2);
              const isExpanded=expandedStaff===s.staff;
              return(
                <Fragment key={s.staff}>
                  <tr onClick={()=>setExpandedStaff(isExpanded?null:s.staff)} style={{cursor:"pointer",background:isExpanded?"#F5F3FF":undefined}}>
                    <TD bold>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontSize:10,color:"#9CA3AF",transition:"transform .15s",display:"inline-block",transform:isExpanded?"rotate(90deg)":"rotate(0deg)"}}>▶</span>
                        <div style={{width:28,height:28,borderRadius:"50%",background:"#EEF2FF",color:"#4338CA",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,flexShrink:0}}>{initials}</div>
                        {s.staff}
                      </div>
                    </TD>
                    <TD right>{s.count}</TD>
                    <TD right><span style={{background:"#EEF2FF",color:"#4338CA",borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700}}>{s.indexAvg}</span></TD>
                    <TD right bold>{fmt$(s.value)}</TD>
                    <TD right color="#0369A1">{s.ext}</TD>
                    <TD right color="#047857">{s.int}</TD>
                  </tr>
                  {isExpanded&&(
                    <tr key={`${s.staff}-expand`}>
                      <td colSpan={6} style={{padding:"0 0 4px 0",background:"#F8F5FF"}}>
                        <ExpandPanel heading={`DELIVERABLES LOGGED BY ${s.staff.toUpperCase()}`}>
                          <LineItemTable entries={s.entries} hide={["staff"]} clubClusters={clubClusters}/>
                        </ExpandPanel>
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
          </tbody>
        </TblWrap>
      )}

      {view==="log"&&(
        <div>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:10,fontFamily:"'DM Sans',sans-serif"}}>
            <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:13,color:"#374151",fontWeight:500,userSelect:"none"}}>
              <Checkbox checked={allSelected} indeterminate={selected.size>0&&!allSelected} onClick={toggleAll}/>
              {selected.size===0?"Select entries to delete":`${selected.size} of ${filteredLog.length} selected`}
            </label>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              {hasSelection&&!confirm&&<span style={{fontSize:12,color:"#6B7280"}}>{selected.size} {selected.size===1?"entry":"entries"} · {fmt$(filteredLog.filter(e=>selected.has(e.id)).reduce((s,e)=>s+e.rate,0))} value</span>}
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
              <TH>#</TH><TH>Time</TH><TH>Staff</TH><TH>Dept</TH><TH>Deliverable</TH><TH>Type</TH><TH>Club / Recipient</TH><TH>Cluster</TH><TH>League</TH><TH right>Rack Rate</TH><TH right>Index</TH><TH>Notes</TH>
            </tr></thead>
            <tbody>
              {reversed.map((e,i)=>{
                const checked=selected.has(e.id);
                return(
                  <tr key={e.id} style={{background:checked?"#FEF2F2":undefined,transition:"background .15s"}}>
                    <td style={{padding:"10px 14px",textAlign:"center"}}><Checkbox checked={checked} onClick={()=>toggle(e.id)}/></td>
                    <td style={{padding:"10px 14px",color:"#9CA3AF",fontSize:11}}>{filteredLog.length-i}</td>
                    <td style={{padding:"10px 14px",color:"#6B7280",fontSize:12}}>{new Date(e.ts).toLocaleString()}</td>
                    <TD bold color="#111827">{e.staff}</TD>
                    <TD><DeptChip dept={e.dept}/></TD>
                    <TD>
                      <div>
                        <div style={{display:"flex",alignItems:"center",gap:6}}>
                          {e.name}
                          {e.recurring&&<RecurringBadge/>}
                        </div>
                        {e.subcat&&e.subcat!==e.name&&<div style={{fontSize:11,color:"#9CA3AF",marginTop:2}}>{e.subcat}</div>}
                      </div>
                    </TD>
                    <TD><Badge type={e.type}/></TD>
                    <TD color="#374151">{e.club}</TD>
                    <TD color="#6B7280">{clubClusters[e.club]||"—"}</TD>
                    <TD><LeagueBadge league={e.league||"League-wide"}/></TD>
                    <TD right bold color={DEPT_CONFIG[e.dept]?.color||"#374151"}>{e.rate>0?fmt$(e.rate):"—"}</TD>
                    <TD right>{e.index_score>0?<span style={{background:"#EEF2FF",color:"#4338CA",borderRadius:6,padding:"2px 8px",fontSize:12,fontWeight:700}}>{e.index_score}</span>:"—"}</TD>
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


// Staff no longer tracked in dashboards/reporting — their historical rows stay in the
// Google Sheet untouched (pull them directly there if ever needed), just excluded here.
const RETIRED_STAFF = ["Tim McCarthy"];
const excludeRetiredStaff = data => data.filter(e => !RETIRED_STAFF.includes(e.staff));

// Internal work never carries a dollar figure, no matter what's sitting in the sheet
// (covers historical rows logged before this rule existed). New entries are also
// written as $0 at logging time — see handleConfirm in DeptTab — so this is a
// display-layer safety net, not the only place the rule is enforced.
const zeroOutInternalRate = data => data.map(e => e.type==="Internal" ? { ...e, rate:0 } : e);

export default function App() {
  const [unlocked,      setUnlocked]      = useState(()=>sessionStorage.getItem("dept_unlocked")==="1");
  const [activeTab,     setActiveTab]     = useState("Dashboard");
  const [log,           setLog]           = useState([]);
  const [deptItems,     setDeptItems]     = useState(FALLBACK_DEPT_ITEMS);
  const [clubsByLeague, setClubsByLeague] = useState(FALLBACK_CLUBS);
  const [clubClusters,  setClubClusters]  = useState({});
  const [loading,       setLoading]       = useState(true);
  const [toast,         setToast]         = useState(null);

  const loadAll=useCallback(async()=>{
    try{
      const [logData,delivData,clubData]=await Promise.all([
        dbFetch(),
        dbFetchDeliverables().catch(()=>[]),
        dbFetchClubs().catch(()=>[]),
      ]);
      setLog(zeroOutInternalRate(excludeRetiredStaff(logData)));
      if(delivData.length) {
        const built = buildDeptItems(delivData);
        // Only use sheet data if it actually produced items for known depts
        const hasItems = ALL_DEPT_NAMES.some(d => built[d]?.external?.length || built[d]?.internal?.length);
        setDeptItems(hasItems ? built : FALLBACK_DEPT_ITEMS);
      } else {
        setDeptItems(FALLBACK_DEPT_ITEMS);
      }
      if(clubData.length) { setClubsByLeague(buildClubsByLeague(clubData)); setClubClusters(buildClubClusters(clubData)); }
    }catch(e){console.error("Load error:",e);}
    finally{setLoading(false);}
  },[]);

  useEffect(()=>{
    if(!unlocked) return;
    loadAll();
    const interval=setInterval(()=>{
      dbFetch().then(d=>setLog(zeroOutInternalRate(excludeRetiredStaff(d)))).catch(()=>{});
    },15000);
    return()=>clearInterval(interval);
  },[unlocked,loadAll]);

  const handleLog=useCallback(async entry=>{
    const newEntry={...entry,id:`${Date.now()}-${Math.random().toString(36).slice(2)}`,ts:entry.ts||Date.now()};
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
            <div style={{maxWidth:"100%",padding:"28px 40px"}}>
              {activeTab==="Dashboard"         &&<Dashboard log={log} onExport={()=>exportExcel(log)} clubsByLeague={clubsByLeague} clubClusters={clubClusters}/>}
              {activeTab==="Activity Explorer" &&<ActivityExplorer log={log} onRemove={handleRemove} onExportView={exportView} clubClusters={clubClusters}/>}
              {ALL_DEPT_NAMES.includes(activeTab)&&<DeptTab dept={activeTab} log={log} onLog={handleLog} onBulkLog={handleBulkLog} onBulkComplete={handleBulkComplete} deptItems={deptItems} clubsByLeague={clubsByLeague}/>}
            </div>
          )}
        </div>
      </div>

      {toast&&<Toast msg={toast.msg} color={toast.color} onDone={()=>setToast(null)}/>}
    </>
  );
}
