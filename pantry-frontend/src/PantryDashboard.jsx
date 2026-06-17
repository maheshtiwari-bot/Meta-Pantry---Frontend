import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import "./PantryDashboard.css";

const API_BASE = "https://meta-pantry-backend.onrender.com/api";

// ─── Palette ───────────────────────────────────────────────────────────────
const C = {
  bg: "#eef1f8", bar: "#ffffff", card: "#ffffff", panel: "#f4f6fc",
  border: "#e2e6f3", ink: "#1e2333", sub: "#6b7280",
  green: "#15803d", greenBg: "#dcfce7",
  amber: "#b45309", amberBg: "#fef3c7",
  red: "#b91c1c", redBg: "#fee2e2",
  blue: "#4338ca", blueBg: "#e0e7ff",
  indigo: "#4f46e5", indigoBg: "#eef2ff",
  accent: "#0ea5e9",
  gradient: "linear-gradient(135deg, #4338ca 0%, #6366f1 55%, #0ea5e9 100%)",
};

// ─── Role → Tabs ───────────────────────────────────────────────────────────
const ALL_TABS = ["Dashboard","Stock Inward","Stock Update","Product Performance","Forecast","Vendiman Dashboard","Client Approvals","Orders","Reports","Admin"];
const ROLE_TABS = {
  admin:           ALL_TABS,
  client:          ["Dashboard","Stock Inward","Stock Update","Product Performance","Forecast","Client Approvals","Reports"],
  clientApprover:  ["Dashboard","Stock Inward","Stock Update","Product Performance","Forecast","Client Approvals","Reports"],
  vendiman:        ["Dashboard","Stock Inward","Stock Update","Product Performance","Forecast","Vendiman Dashboard","Orders","Reports"],
};

// ─── Shared styles ─────────────────────────────────────────────────────────
const inputStyle = { border:`1px solid ${C.border}`, borderRadius:10, padding:"10px 14px", fontSize:15, width:"100%", background:"#fff", color:C.ink, boxSizing:"border-box" };
const smInput = { ...inputStyle, padding:"8px 10px", fontSize:14 };
const btn = (primary) => ({ border:primary?"none":`1px solid ${C.border}`, background:primary?C.gradient:"#fff", color:primary?"#fff":C.ink, borderRadius:12, padding:"10px 18px", fontSize:14, fontWeight:600, cursor:"pointer", boxShadow:primary?"0 4px 12px rgba(67,56,202,0.25)":"none", whiteSpace:"nowrap" });
const th = { textAlign:"left", fontSize:11, letterSpacing:"0.06em", color:C.sub, fontWeight:700, padding:"10px 10px", textTransform:"uppercase", borderBottom:`2px solid ${C.border}`, background:C.panel, whiteSpace:"nowrap" };
const td = { padding:"11px 10px", fontSize:13, color:C.ink, borderBottom:`1px solid ${C.border}` };
const trAlt = (i) => ({ background: i%2===1 ? C.panel : "transparent" });
const cellInput = { border:`1px solid ${C.border}`, borderRadius:8, padding:"6px 8px", fontSize:13, width:"100%", background:"#fff", color:C.ink, boxSizing:"border-box" };
const card = { background:C.card, borderRadius:16, padding:"18px 16px", border:`1px solid ${C.border}`, marginBottom:16 };

const STATUS_COLORS = {
  Critical: { color:C.red, background:C.redBg },
  Optimal:  { color:C.green, background:C.greenBg },
  Surplus:  { color:C.blue, background:C.blueBg },
  "Pending Approval": { color:C.amber, background:C.amberBg },
  Approved: { color:C.green, background:C.greenBg },
  Rejected: { color:C.red, background:C.redBg },
  Feedback: { color:C.indigo, background:C.indigoBg },
};

function Pill({ label }) {
  const s = STATUS_COLORS[label] || { color:C.sub, background:C.panel };
  return <span className="pd-pill" style={{ ...s, fontSize:11, fontWeight:700, padding:"3px 9px", borderRadius:999, display:"inline-block" }}>{label}</span>;
}
function Tag({ label }) {
  return <span className="pd-tag" style={{ background:C.panel, border:`1px solid ${C.border}`, color:C.sub, fontSize:11, padding:"3px 9px", borderRadius:999, whiteSpace:"nowrap", display:"inline-block" }}>{label}</span>;
}

function SectionHeader({ title, sub, action }) {
  return (
    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8, marginBottom:14 }}>
      <div>
        <div style={{ fontSize:12, letterSpacing:"0.08em", fontWeight:700, color:C.ink }}>{title}</div>
        {sub && <div style={{ fontSize:12, color:C.sub, marginTop:2 }}>{sub}</div>}
      </div>
      {action}
    </div>
  );
}

// Excel download helper
function dlExcel(rows, filename) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, filename);
}

// ─── Login Screen ──────────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [loginId, setLoginId] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = (e) => {
    e.preventDefault();
    setLoading(true); setError("");
    fetch(`${API_BASE}/auth/login`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ loginId, password }) })
      .then(r => r.json())
      .then(data => { if (data.error) { setError(data.error); setLoading(false); } else onLogin(data); })
      .catch(() => { setError("Cannot connect to server"); setLoading(false); });
  };

  return (
    <div style={{ minHeight:"100vh", background:C.gradient, display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
      <div style={{ background:"#fff", borderRadius:20, padding:"36px 28px", width:"100%", maxWidth:380, boxShadow:"0 20px 60px rgba(67,56,202,0.25)" }}>
        <div style={{ textAlign:"center", marginBottom:28 }}>
          <div style={{ fontSize:28, marginBottom:6 }}>🏢</div>
          <div style={{ fontSize:20, fontWeight:700, color:C.ink }}>Meta India Pantry</div>
          <div style={{ fontSize:13, color:C.sub }}>Sign in to continue</div>
          <div style={{ marginTop:10, padding:"10px 14px", background:"#f0f4ff", borderRadius:10, border:"1px solid #c7d2fe", fontSize:12, color:"#4338ca", textAlign:"left" }}>
            <b>First time?</b> Use the default admin account below to sign in, then go to <b>Admin → User Access</b> to create new users.<br/>
            <span style={{ display:"inline-block", marginTop:6, fontFamily:"monospace", background:"#e0e7ff", padding:"3px 8px", borderRadius:6 }}>Login ID: admin &nbsp;|&nbsp; Password: admin123</span>
          </div>
        </div>
        <form onSubmit={handleSubmit}>
          <div style={{ marginBottom:14 }}>
            <div style={{ fontSize:12, color:C.sub, marginBottom:6 }}>Login ID</div>
            <input style={inputStyle} value={loginId} onChange={e=>setLoginId(e.target.value)} placeholder="Enter login ID" autoFocus />
          </div>
          <div style={{ marginBottom:20 }}>
            <div style={{ fontSize:12, color:C.sub, marginBottom:6 }}>Password</div>
            <input type="password" style={inputStyle} value={password} onChange={e=>setPassword(e.target.value)} placeholder="Enter password" />
          </div>
          {error && <div style={{ color:C.red, fontSize:13, marginBottom:12, padding:"8px 12px", background:C.redBg, borderRadius:8 }}>{error}</div>}
          <button type="submit" disabled={loading} style={{ ...btn(true), width:"100%", padding:"13px" }}>{loading?"Signing in…":"Sign In"}</button>
        </form>
      </div>
    </div>
  );
}

// ─── Dashboard ─────────────────────────────────────────────────────────────
function Dashboard({ sites, statusDefs }) {
  const [selSites, setSelSites] = useState([]);
  const [data, setData] = useState(null);

  const load = (ids) => {
    const q = ids.length ? ids.join(",") : (sites.map(s=>s.id).join(",") || "");
    if (!q) return;
    fetch(`${API_BASE}/dashboard?siteIds=${q}`).then(r=>r.json()).then(setData).catch(()=>{});
  };

  useEffect(() => { if (sites.length) load(selSites); }, [sites, selSites]);

  const toggleSite = (id) => {
    setSelSites(prev => prev.includes(id) ? prev.filter(x=>x!==id) : [...prev, id]);
  };

  const stats = data ? [
    { label:"Total Stock Value", value:`₹${Number(data.totalStockValue).toLocaleString("en-IN")}`, color:C.blue },
    { label:"Days of Stock", value:data.avgDaysOfStock, color:C.blue },
    { label:"Critical Stocks", value:data.criticalCount, color:C.red },
    { label:"Optimal Stocks", value:data.optimalCount, color:C.green },
    { label:"Surplus Stocks", value:data.surplusCount, color:C.indigo },
  ] : [];

  return (
    <div>
      {/* Site multi-select */}
      <div style={{ ...card, padding:"14px 16px" }}>
        <div style={{ fontSize:12, color:C.sub, marginBottom:8 }}>FILTER BY SITE (select one or more)</div>
        <div style={{ display:"flex", flexWrap:"wrap", gap:8 }}>
          {sites.map(s=>(
            <button key={s.id} onClick={()=>toggleSite(s.id)} style={{ ...btn(selSites.includes(s.id)), padding:"7px 14px", fontSize:13 }}>
              {selSites.includes(s.id) ? "✓ " : ""}{s.name}
            </button>
          ))}
          {selSites.length>0 && <button onClick={()=>setSelSites([])} style={{ ...btn(false), padding:"7px 14px", fontSize:13, color:C.sub }}>Clear</button>}
        </div>
      </div>

      {/* Stat cards */}
      {data && (
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:10, marginBottom:16 }}>
          {stats.map(s=>(
            <div key={s.label} className="pd-tile" style={{ background:C.panel, borderRadius:14, padding:"16px 14px", border:`1px solid ${C.border}` }}>
              <div style={{ fontSize:24, fontWeight:700, color:s.color }}>{s.value}</div>
              <div style={{ fontSize:11, color:C.sub, marginTop:4 }}>{s.label}</div>
            </div>
          ))}
        </div>
      )}

      {/* Stock table */}
      {data && (
        <div className="pd-card" style={card}>
          <SectionHeader title="STOCK OVERVIEW" action={
            <button style={{ ...btn(false), fontSize:12, padding:"7px 12px" }} onClick={()=>dlExcel(data.rows,"dashboard.xlsx")}>⬇ Excel</button>
          }/>
          <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", minWidth:700 }}>
              <thead><tr>
                {["Product Code","Product Name","Brand","Category","Sub Category","MRP","Available Qty","Pref. Days","Avail. Days","Status"].map(h=><th key={h} style={th}>{h}</th>)}
              </tr></thead>
              <tbody>
                {data.rows.map((r,i)=>(
                  <tr key={`${r.siteId}-${r.productCode}`} style={trAlt(i)}>
                    <td style={td}>{r.productCode}</td>
                    <td style={{ ...td, fontWeight:600 }}>{r.productName}</td>
                    <td style={td}>{r.brand}</td>
                    <td style={td}><Tag label={r.category}/></td>
                    <td style={td}>{r.subCategory}</td>
                    <td style={td}>₹{r.mrp}</td>
                    <td style={{ ...td, fontWeight:600 }}>{r.availableQty}</td>
                    <td style={td}>{r.preferredDays}</td>
                    <td style={{ ...td, fontWeight:600, color:r.status==="Critical"?C.red:r.status==="Surplus"?C.blue:C.green }}>
                      {r.availableDays??"-"}
                    </td>
                    <td style={td}><Pill label={r.status}/></td>
                  </tr>
                ))}
                {data.rows.length===0 && <tr><td colSpan={10} style={{ ...td, color:C.sub }}>No stock data. Add products and perform stock update first.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Stock Inward ─────────────────────────────────────────────────────────
function StockInward({ sites, vendors, products, showToast, currentUser }) {
  const [mode, setMode] = useState(null); // null | "auto" | "manual"
  const [dcCode, setDcCode] = useState("");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [vendor, setVendor] = useState("");
  const [items, setItems] = useState([]);
  const [nameSearch, setNameSearch] = useState("");
  const [suggestions, setSuggestions] = useState([]);
  const [history, setHistory] = useState([]);
  const [submitted, setSubmitted] = useState(false);

  const loadHistory = () => fetch(`${API_BASE}/stock-inward/history`).then(r=>r.json()).then(setHistory).catch(()=>{});
  useEffect(()=>{ loadHistory(); }, []);

  const vendorProducts = () => {
    if (!vendor) return products;
    const v = vendors.find(v=>v.id===Number(vendor));
    if (!v || v.products==="all") return products;
    return products.filter(p=>v.products.includes(p.code));
  };

  const handleNameSearch = (val) => {
    setNameSearch(val);
    const vp = vendorProducts();
    setSuggestions(val.trim().length>0 ? vp.filter(p=>p.name.toLowerCase().includes(val.toLowerCase())).slice(0,6) : []);
  };

  const addItem = (prod) => {
    if (items.find(it=>it.productCode===prod.code)) { showToast("Product already added"); setSuggestions([]); setNameSearch(""); return; }
    setItems(prev=>[...prev, { productCode:prod.code, productName:prod.name, hsnCode:prod.hsnCode||"", mrp:prod.mrp||0, gst:prod.gst||0, qty:0 }]);
    setSuggestions([]); setNameSearch("");
  };

  const updateItem = (code, field, val) => setItems(prev=>prev.map(it=>it.productCode===code?{...it,[field]:val}:it));
  const removeItem = (code) => setItems(prev=>prev.filter(it=>it.productCode!==code));

  const handleConfirm = () => {
    if (!dcCode.trim()) { showToast("DC Code is required"); return; }
    if (!from) { showToast("Select From (Vendor)"); return; }
    if (!to) { showToast("Select To (Site)"); return; }
    if (items.length===0) { showToast("Add at least one product"); return; }
    const validItems = items.filter(it=>Number(it.qty)>0);
    if (validItems.length===0) { showToast("Enter quantity for at least one product"); return; }

    fetch(`${API_BASE}/stock-inward`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ dcCode, from:Number(from), to:Number(to), vendor:Number(vendor||from), items:validItems.map(it=>({...it,qty:Number(it.qty)})) }) })
      .then(r=>r.json())
      .then(()=>{ showToast("Stock inwarded successfully"); setSubmitted(true); loadHistory(); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const handleAddMore = () => { setItems([]); setNameSearch(""); setSubmitted(false); };
  const handleReset = () => { setMode(null); setDcCode(""); setFrom(""); setTo(""); setVendor(""); setItems([]); setSubmitted(false); };

  const fromVendorName = (id) => vendors.find(v=>v.id===Number(id))?.name || id;
  const toSiteName = (id) => sites.find(s=>s.id===Number(id))?.name || id;

  return (
    <div>
      {/* Mode selector */}
      {!mode && (
        <div className="pd-card" style={card}>
          <SectionHeader title="STOCK INWARD" sub="Choose how to inward stock" />
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))", gap:12 }}>
            <button onClick={()=>setMode("auto")} style={{ ...btn(true), padding:"20px 16px", borderRadius:14, display:"flex", flexDirection:"column", alignItems:"flex-start", gap:6 }}>
              <span style={{ fontSize:22 }}>📡</span>
              <span style={{ fontSize:15 }}>Auto Fetch</span>
              <span style={{ fontSize:11, opacity:.8, fontWeight:400 }}>Enter DC Code to fetch inward details</span>
            </button>
            <button onClick={()=>setMode("manual")} style={{ ...btn(false), padding:"20px 16px", borderRadius:14, display:"flex", flexDirection:"column", alignItems:"flex-start", gap:6, border:`2px solid ${C.border}` }}>
              <span style={{ fontSize:22 }}>✏️</span>
              <span style={{ fontSize:15 }}>Manual Add</span>
              <span style={{ fontSize:11, color:C.sub, fontWeight:400 }}>Manually fill DC and product details</span>
            </button>
          </div>
        </div>
      )}

      {/* Auto Fetch */}
      {mode==="auto" && !submitted && (
        <div className="pd-card" style={card}>
          <SectionHeader title="AUTO FETCH STOCK INWARD" action={<button style={{ ...btn(false), fontSize:12, padding:"6px 12px" }} onClick={handleReset}>← Back</button>} />
          <div style={{ marginBottom:14 }}>
            <div style={{ fontSize:12, color:C.sub, marginBottom:6 }}>DC Code</div>
            <div style={{ display:"flex", gap:10 }}>
              <input style={{ ...inputStyle, flex:1 }} value={dcCode} onChange={e=>setDcCode(e.target.value)} placeholder="Enter DC Code e.g. DC-2024-001" />
              <button style={btn(true)} onClick={()=>showToast("Auto-fetch integration pending — please use Manual Add for now")}>Fetch</button>
            </div>
          </div>
          <div style={{ padding:"16px", background:C.panel, borderRadius:10, color:C.sub, fontSize:13 }}>Auto-fetch requires integration with your DC system. Use Manual Add in the meantime.</div>
        </div>
      )}

      {/* Manual Add */}
      {mode==="manual" && !submitted && (
        <div className="pd-card" style={card}>
          <SectionHeader title="MANUAL STOCK INWARD" action={<button style={{ ...btn(false), fontSize:12, padding:"6px 12px" }} onClick={handleReset}>← Back</button>} />
          {/* Common fields */}
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))", gap:12, marginBottom:16 }}>
            {[
              ["DC Code *", <input style={inputStyle} value={dcCode} onChange={e=>setDcCode(e.target.value)} placeholder="DC-2024-001"/>],
              ["From (Vendor) *", <select style={inputStyle} value={from} onChange={e=>{ setFrom(e.target.value); setVendor(e.target.value); }}><option value="">Select vendor</option>{vendors.map(v=><option key={v.id} value={v.id}>{v.name}</option>)}</select>],
              ["To (Site) *",    <select style={inputStyle} value={to} onChange={e=>setTo(e.target.value)}><option value="">Select site</option>{sites.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}</select>],
            ].map(([label,ctrl])=>(
              <div key={label}><div style={{ fontSize:12, color:C.sub, marginBottom:6 }}>{label}</div>{ctrl}</div>
            ))}
          </div>

          {/* Product search */}
          <div style={{ marginBottom:14, position:"relative" }}>
            <div style={{ fontSize:12, color:C.sub, marginBottom:6 }}>Search & Add Product</div>
            <input style={inputStyle} value={nameSearch} onChange={e=>handleNameSearch(e.target.value)} placeholder="Type product name…" />
            {suggestions.length>0 && (
              <div style={{ position:"absolute", top:"100%", left:0, right:0, background:"#fff", border:`1px solid ${C.border}`, borderRadius:10, marginTop:4, boxShadow:"0 8px 24px rgba(40,50,110,.12)", zIndex:20, overflow:"hidden" }}>
                {suggestions.map(p=>(
                  <div key={p.code} onClick={()=>addItem(p)} style={{ padding:"10px 14px", cursor:"pointer", fontSize:13, borderBottom:`1px solid ${C.border}` }}
                    onMouseDown={e=>e.preventDefault()}>
                    <div style={{ fontWeight:600 }}>{p.name}</div>
                    <div style={{ fontSize:11, color:C.sub }}>{p.code} · {p.category}</div>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Items table */}
          {items.length>0 && (
            <div style={{ overflowX:"auto", marginBottom:14 }}>
              <table style={{ width:"100%", borderCollapse:"collapse", minWidth:500 }}>
                <thead><tr>{["HSN Code","Product Code","Product Name","MRP","GST%","Qty",""].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                <tbody>
                  {items.map((it,i)=>(
                    <tr key={it.productCode} style={trAlt(i)}>
                      <td style={td}><input style={{ ...cellInput, width:80 }} value={it.hsnCode} onChange={e=>updateItem(it.productCode,"hsnCode",e.target.value)}/></td>
                      <td style={td}>{it.productCode}</td>
                      <td style={{ ...td, fontWeight:600 }}>{it.productName}</td>
                      <td style={td}><input type="number" style={{ ...cellInput, width:70 }} value={it.mrp} onChange={e=>updateItem(it.productCode,"mrp",e.target.value)}/></td>
                      <td style={td}><input type="number" style={{ ...cellInput, width:60 }} value={it.gst} onChange={e=>updateItem(it.productCode,"gst",e.target.value)}/></td>
                      <td style={td}><input type="number" min={0} style={{ ...cellInput, width:70 }} value={it.qty} onChange={e=>updateItem(it.productCode,"qty",e.target.value)}/></td>
                      <td style={td}><button style={{ border:`1px solid ${C.redBg}`, background:"#fff", color:C.red, borderRadius:8, padding:"5px 8px", cursor:"pointer" }} onClick={()=>removeItem(it.productCode)}>✕</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          {items.length===0 && <div style={{ padding:"14px", background:C.panel, borderRadius:10, color:C.sub, fontSize:13, marginBottom:14 }}>No products added yet. Search above to add products.</div>}
          <div style={{ display:"flex", justifyContent:"flex-end" }}>
            <button style={btn(true)} onClick={handleConfirm}>Confirm Inward</button>
          </div>
        </div>
      )}

      {/* Success + Add More */}
      {submitted && (
        <div className="pd-card" style={{ ...card, background:C.greenBg, border:`1px solid ${C.green}` }}>
          <div style={{ fontSize:16, fontWeight:700, color:C.green, marginBottom:8 }}>✓ Stock Inwarded Successfully</div>
          <div style={{ fontSize:13, color:C.green, marginBottom:16 }}>Products have been added to {toSiteName(to)}'s stock.</div>
          <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
            <button style={btn(true)} onClick={handleAddMore}>+ Add More Products</button>
            <button style={btn(false)} onClick={handleReset}>Start New Inward</button>
          </div>
        </div>
      )}

      {/* History */}
      <div className="pd-card" style={card}>
        <SectionHeader title="INWARD HISTORY" />
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse" }}>
            <thead><tr>{["Date","DC Code","From","To","Items","Status"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
            <tbody>
              {history.slice(0,20).map((h,i)=>(
                <tr key={h.id} style={trAlt(i)}>
                  <td style={td}>{new Date(h.date).toLocaleDateString("en-IN")}</td>
                  <td style={{ ...td, fontWeight:600 }}>{h.dcCode}</td>
                  <td style={td}>{fromVendorName(h.from)}</td>
                  <td style={td}>{toSiteName(h.to)}</td>
                  <td style={td}>{h.items?.length||0}</td>
                  <td style={td}><Pill label={h.status||"Confirmed"}/></td>
                </tr>
              ))}
              {history.length===0 && <tr><td colSpan={6} style={{ ...td, color:C.sub }}>No inward history yet.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── Stock Update ─────────────────────────────────────────────────────────
function StockUpdate({ sites, showToast, currentUser }) {
  const [siteId, setSiteId] = useState("");
  const [rows, setRows] = useState([]);
  const [history, setHistory] = useState([]);
  const [confirm, setConfirm] = useState(null);
  const [editDays, setEditDays] = useState({});

  useEffect(()=>{ if(sites.length && !siteId) setSiteId(String(sites[0].id)); }, [sites]);

  const loadRows = (sid) => {
    fetch(`${API_BASE}/stock-update?siteId=${sid}`).then(r=>r.json()).then(d=>{ setRows(d); const ed={}; d.forEach(r=>{ if(r.preferredDaysOverride!=null) ed[r.productCode]=r.preferredDaysOverride; }); setEditDays(ed); }).catch(()=>{});
    fetch(`${API_BASE}/stock-update/history?siteId=${sid}`).then(r=>r.json()).then(setHistory).catch(()=>{});
  };

  useEffect(()=>{ if(siteId) loadRows(siteId); }, [siteId]);

  const handleQtyChange = (row, val) => {
    setConfirm({ productCode:row.productCode, productName:row.productName, prevQty:row.qty, newQty:Number(val), newDays:editDays[row.productCode]??row.preferredDays });
  };

  const confirmUpdate = () => {
    if (!confirm) return;
    fetch(`${API_BASE}/stock-update`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ siteId:Number(siteId), productCode:confirm.productCode, qty:confirm.newQty, preferredDaysOverride:editDays[confirm.productCode]??null, updatedBy:currentUser.loginId }) })
      .then(()=>{ showToast("Stock updated"); setConfirm(null); loadRows(siteId); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const handleDaysChange = (code, val) => setEditDays(prev=>({...prev,[code]:Number(val)}));

  const exportExcel = () => {
    if (!rows.length) return;
    dlExcel(rows.map(r=>({ "Product Code":r.productCode, "Product Name":r.productName, Brand:r.brand, Category:r.category, "Sub Category":r.subCategory, MRP:r.mrp, Qty:r.qty, "Preferred Days":editDays[r.productCode]??r.preferredDays, "Updated At":r.updatedAt, "Updated By":r.updatedBy })), "stock-update.xlsx");
  };

  return (
    <div>
      {/* Confirmation modal */}
      {confirm && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.4)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:100, padding:16 }}>
          <div style={{ background:"#fff", borderRadius:16, padding:"24px 20px", maxWidth:360, width:"100%", boxShadow:"0 20px 60px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize:15, fontWeight:700, marginBottom:8 }}>Confirm Stock Update</div>
            <div style={{ fontSize:13, color:C.sub, marginBottom:16 }}>
              <b>{confirm.productName}</b><br/>
              Previous qty: <b>{confirm.prevQty}</b> → New qty: <b>{confirm.newQty}</b>
            </div>
            <div style={{ display:"flex", gap:10 }}>
              <button style={{ ...btn(true), flex:1 }} onClick={confirmUpdate}>Confirm</button>
              <button style={{ ...btn(false), flex:1 }} onClick={()=>setConfirm(null)}>Cancel</button>
            </div>
          </div>
        </div>
      )}

      {/* Stock Update section */}
      <div className="pd-card" style={card}>
        <SectionHeader title="STOCK UPDATE" action={
          <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
            <select style={{ ...smInput, width:180 }} value={siteId} onChange={e=>setSiteId(e.target.value)}>
              <option value="">Select site</option>
              {sites.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            <button style={{ ...btn(false), fontSize:12, padding:"7px 12px" }} onClick={exportExcel}>⬇ Excel</button>
          </div>
        }/>
        {siteId ? (
          <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", minWidth:680 }}>
              <thead><tr>{["Product Code","Product Name","Brand","Category","Sub Category","MRP","Current Qty","Updated Qty","Pref. Days"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
              <tbody>
                {rows.map((r,i)=>(
                  <tr key={r.productCode} style={trAlt(i)}>
                    <td style={td}>{r.productCode}</td>
                    <td style={{ ...td, fontWeight:600 }}>{r.productName}</td>
                    <td style={td}>{r.brand}</td>
                    <td style={td}><Tag label={r.category}/></td>
                    <td style={td}>{r.subCategory}</td>
                    <td style={td}>₹{r.mrp}</td>
                    <td style={{ ...td, fontWeight:600 }}>{r.qty}</td>
                    <td style={td}>
                      <input type="number" min={0} style={{ ...cellInput, width:80 }} defaultValue={r.qty}
                        onBlur={e=>{ if(Number(e.target.value)!==r.qty) handleQtyChange(r, e.target.value); }} />
                    </td>
                    <td style={td}>
                      <input type="number" min={0} style={{ ...cellInput, width:60 }} value={editDays[r.productCode]??r.preferredDays}
                        onChange={e=>handleDaysChange(r.productCode, e.target.value)} />
                    </td>
                  </tr>
                ))}
                {rows.length===0 && <tr><td colSpan={9} style={{ ...td, color:C.sub }}>No stock data. Inward some products first.</td></tr>}
              </tbody>
            </table>
          </div>
        ) : <div style={{ padding:"16px", color:C.sub, fontSize:13 }}>Select a site to view stock.</div>}
      </div>

      {/* History */}
      <div className="pd-card" style={card}>
        <SectionHeader title="STOCK UPDATE HISTORY" />
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse" }}>
            <thead><tr>{["Date","Product","Prev Qty","New Qty","Updated By"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
            <tbody>
              {history.slice(0,30).map((h,i)=>(
                <tr key={h.id} style={trAlt(i)}>
                  <td style={td}>{new Date(h.updatedAt).toLocaleString("en-IN")}</td>
                  <td style={{ ...td, fontWeight:600 }}>{h.productName}</td>
                  <td style={td}>{h.prevQty}</td>
                  <td style={{ ...td, fontWeight:600 }}>{h.qty}</td>
                  <td style={td}>{h.updatedBy||"—"}</td>
                </tr>
              ))}
              {history.length===0 && <tr><td colSpan={5} style={{ ...td, color:C.sub }}>No history yet.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── Product Performance ──────────────────────────────────────────────────
function ProductPerformance({ sites, showToast }) {
  const [siteId, setSiteId] = useState("");
  const [rows, setRows] = useState([]);
  const [editQpd, setEditQpd] = useState({});
  const [saving, setSaving] = useState({});

  useEffect(()=>{ if(sites.length && !siteId) setSiteId(String(sites[0].id)); }, [sites]);

  const load = (sid) => {
    fetch(`${API_BASE}/product-performance?siteId=${sid}`).then(r=>r.json()).then(d=>{ setRows(d); const e={}; d.forEach(r=>{ if(r.qtyPerDay!=null) e[r.productCode]=r.qtyPerDay; }); setEditQpd(e); }).catch(()=>{});
  };
  useEffect(()=>{ if(siteId) load(siteId); }, [siteId]);

  const saveQpd = (productCode, method) => {
    const val = editQpd[productCode];
    if (val==null) return;
    setSaving(p=>({...p,[productCode]:true}));
    fetch(`${API_BASE}/product-performance`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ productCode, siteId:Number(siteId), qtyPerDay:Number(val), method }) })
      .then(()=>{ setSaving(p=>({...p,[productCode]:false})); load(siteId); showToast("Saved"); })
      .catch(()=>{ setSaving(p=>({...p,[productCode]:false})); showToast("Could not reach backend"); });
  };

  const autoCalc = () => {
    fetch(`${API_BASE}/product-performance/auto-calculate`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ siteId:Number(siteId) }) })
      .then(r=>r.json()).then(d=>{ showToast(`Auto-calculated ${d.updated?.length||0} products`); load(siteId); }).catch(()=>showToast("Could not reach backend"));
  };

  const exportExcel = () => {
    dlExcel(rows.map(r=>({ "Product Code":r.productCode, "Product Name":r.productName, Brand:r.brand, Category:r.category, "Sub Category":r.subCategory, MRP:r.mrp, "Qty Per Day":r.qtyPerDay??"-", "Method":r.method||"-", "Last Updated":r.updatedAt||"-" })), "product-performance.xlsx");
  };

  return (
    <div>
      <div className="pd-card" style={card}>
        <SectionHeader title="PRODUCT PERFORMANCE" sub="Qty per day consumed at selected site" action={
          <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
            <select style={{ ...smInput, width:180 }} value={siteId} onChange={e=>setSiteId(e.target.value)}>
              <option value="">Select site</option>
              {sites.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            {siteId && <><button style={{ ...btn(false), fontSize:12, padding:"7px 12px" }} onClick={autoCalc}>⚡ Auto Generate</button>
            <button style={{ ...btn(false), fontSize:12, padding:"7px 12px" }} onClick={exportExcel}>⬇ Excel</button></>}
          </div>
        }/>
        {siteId ? (
          <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", minWidth:640 }}>
              <thead><tr>{["Product Code","Product Name","Brand","Category","Sub Category","MRP","Qty / Day","Method",""].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
              <tbody>
                {rows.map((r,i)=>(
                  <tr key={r.productCode} style={trAlt(i)}>
                    <td style={td}>{r.productCode}</td>
                    <td style={{ ...td, fontWeight:600 }}>{r.productName}</td>
                    <td style={td}>{r.brand}</td>
                    <td style={td}><Tag label={r.category}/></td>
                    <td style={td}>{r.subCategory}</td>
                    <td style={td}>₹{r.mrp}</td>
                    <td style={td}>
                      <input type="number" min={0} step={0.01} style={{ ...cellInput, width:80 }} value={editQpd[r.productCode]??""} onChange={e=>setEditQpd(p=>({...p,[r.productCode]:e.target.value}))}/>
                    </td>
                    <td style={td}>{r.method ? <Tag label={r.method==="auto"?"Auto-calc":"Manual"}/> : "—"}</td>
                    <td style={td}>
                      <button style={{ ...btn(true), padding:"5px 10px", fontSize:12 }} disabled={saving[r.productCode]} onClick={()=>saveQpd(r.productCode,"manual")}>Save</button>
                    </td>
                  </tr>
                ))}
                {rows.length===0 && <tr><td colSpan={9} style={{ ...td, color:C.sub }}>No products at this site yet.</td></tr>}
              </tbody>
            </table>
          </div>
        ) : <div style={{ padding:"16px", color:C.sub, fontSize:13 }}>Select a site to manage performance data.</div>}
      </div>
    </div>
  );
}

// ─── Forecast ─────────────────────────────────────────────────────────────
function Forecast({ sites, showToast, currentUser }) {
  const [siteId, setSiteId] = useState("");
  const [rows, setRows] = useState([]);
  const [orderItems, setOrderItems] = useState(null);
  const [pendingOrders, setPendingOrders] = useState([]);
  const [showManual, setShowManual] = useState(false);
  const [manualItems, setManualItems] = useState([{ productCode:"", productName:"", qty:0, mrp:0 }]);
  const newOrderId = `ORD-${Date.now().toString(36).toUpperCase()}`;

  useEffect(()=>{ if(sites.length && !siteId) setSiteId(String(sites[0].id)); }, [sites]);

  const load = (sid) => {
    fetch(`${API_BASE}/forecast?siteId=${sid}`).then(r=>r.json()).then(setRows).catch(()=>{});
    fetch(`${API_BASE}/orders?siteId=${sid}`).then(r=>r.json()).then(d=>setPendingOrders(d.filter(o=>o.status==="Pending Approval"))).catch(()=>{});
  };
  useEffect(()=>{ if(siteId) load(siteId); }, [siteId]);

  const pendingProductCodes = new Set(pendingOrders.flatMap(o=>(o.items||[]).map(it=>it.productCode)));

  const autoGenerate = () => {
    const critical = rows.filter(r=>r.status==="Critical"&&r.orderQty>0).map(r=>({...r,editQty:r.orderQty}));
    if (!critical.length) { showToast("No Critical products to order"); return; }
    const dups = critical.filter(r=>pendingProductCodes.has(r.productCode));
    if (dups.length) {
      showToast(`⚠ ${dups.map(d=>d.productCode).join(", ")} already have a pending order`);
      const safe = critical.filter(r=>!pendingProductCodes.has(r.productCode));
      if (!safe.length) return;
      setOrderItems(safe);
    } else {
      setOrderItems(critical);
    }
  };

  const updateOrderQty = (code, val) => setOrderItems(prev=>prev.map(it=>it.productCode===code?{...it,editQty:Number(val)}:it));
  const removeOrderItem = (code) => setOrderItems(prev=>prev.filter(it=>it.productCode!==code));

  const sendForApproval = (items, oid) => {
    const filtered = items.filter(it=>it.editQty>0||it.qty>0).map(it=>({ productCode:it.productCode, productName:it.productName||it.productCode, brand:it.brand||"", qty:it.editQty||it.qty||0, mrp:it.mrp||0 }));
    if (!filtered.length) { showToast("No items to send"); return; }
    fetch(`${API_BASE}/orders`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ id:oid||newOrderId, siteId:Number(siteId), placedBy:currentUser.loginId, items:filtered, status:"Pending Approval" }) })
      .then(()=>{ showToast(`Order sent for approval`); setOrderItems(null); setShowManual(false); setManualItems([{ productCode:"",productName:"",qty:0,mrp:0 }]); load(siteId); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const exportExcel = () => dlExcel(rows.map(r=>({ "Product Code":r.productCode,"Product Name":r.productName,Brand:r.brand,Category:r.category,"Sub Category":r.subCategory,MRP:r.mrp,"Available Qty":r.availableQty,"Preferred Days":r.preferredDays,"Available Days":r.availableDays??"-",Status:r.status,"Order Qty":r.orderQty||0 })),"forecast.xlsx");

  const updateManual = (i,f,v) => setManualItems(prev=>prev.map((it,idx)=>idx===i?{...it,[f]:v}:it));
  const addManualRow = () => setManualItems(prev=>[...prev,{ productCode:"",productName:"",qty:0,mrp:0 }]);
  const removeManualRow = (i) => setManualItems(prev=>prev.filter((_,idx)=>idx!==i));

  return (
    <div>
      <div className="pd-card" style={card}>
        <SectionHeader title="FORECAST" action={
          <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
            <select style={{ ...smInput, width:180 }} value={siteId} onChange={e=>setSiteId(e.target.value)}>
              {sites.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            {siteId && <>
              <button style={{ ...btn(true), fontSize:12, padding:"7px 12px" }} onClick={autoGenerate}>⚡ Auto Generate Order</button>
              <button style={{ ...btn(false), fontSize:12, padding:"7px 12px" }} onClick={()=>setShowManual(m=>!m)}>✏ Manual Order</button>
              <button style={{ ...btn(false), fontSize:12, padding:"7px 12px" }} onClick={exportExcel}>⬇ Excel</button>
            </>}
          </div>
        }/>
        {pendingOrders.length>0 && (
          <div style={{ padding:"8px 12px", background:C.amberBg, borderRadius:8, marginBottom:12, fontSize:12, color:C.amber }}>
            ⚠ {pendingOrders.length} pending order(s) for this site. Products already on order are highlighted.
          </div>
        )}
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", minWidth:700 }}>
            <thead><tr>{["Product Code","Product Name","Brand","Category","Sub Category","MRP","Avail. Qty","Pref. Days","Avail. Days","Status","Order Qty"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
            <tbody>
              {rows.map((r,i)=>(
                <tr key={r.productCode} style={{ ...trAlt(i), opacity:pendingProductCodes.has(r.productCode)?0.55:1 }}>
                  <td style={td}>{r.productCode}{pendingProductCodes.has(r.productCode)&&<span style={{ fontSize:9, color:C.amber, marginLeft:4 }}>ON ORDER</span>}</td>
                  <td style={{ ...td, fontWeight:600 }}>{r.productName}</td>
                  <td style={td}>{r.brand}</td>
                  <td style={td}><Tag label={r.category}/></td>
                  <td style={td}>{r.subCategory}</td>
                  <td style={td}>₹{r.mrp}</td>
                  <td style={{ ...td, fontWeight:600 }}>{r.availableQty}</td>
                  <td style={td}>{r.preferredDays}</td>
                  <td style={{ ...td, color:r.status==="Critical"?C.red:r.status==="Surplus"?C.blue:C.green, fontWeight:600 }}>{r.availableDays??"-"}</td>
                  <td style={td}><Pill label={r.status}/></td>
                  <td style={{ ...td, fontWeight:700, color:r.orderQty>0?C.amber:C.sub }}>{r.orderQty||"-"}</td>
                </tr>
              ))}
              {rows.length===0 && <tr><td colSpan={11} style={{ ...td, color:C.sub }}>No data. Add stock and performance data first.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      {/* Manual Order */}
      {showManual && !orderItems && (
        <div className="pd-card" style={{ ...card, border:`1px solid ${C.green}` }}>
          <SectionHeader title="CREATE MANUAL ORDER" action={<button style={{ ...btn(false), fontSize:12, padding:"6px 12px" }} onClick={()=>setShowManual(false)}>Cancel</button>}/>
          <div style={{ overflowX:"auto", marginBottom:12 }}>
            <table style={{ width:"100%", borderCollapse:"collapse" }}>
              <thead><tr>{["Product Code","Product Name","Qty","MRP",""].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
              <tbody>
                {manualItems.map((it,i)=>(
                  <tr key={i}>
                    <td style={td}><input style={{ ...cellInput, width:120 }} value={it.productCode} onChange={e=>updateManual(i,"productCode",e.target.value)} placeholder="Code"/></td>
                    <td style={td}><input style={{ ...cellInput, width:160 }} value={it.productName} onChange={e=>updateManual(i,"productName",e.target.value)} placeholder="Name"/></td>
                    <td style={td}><input type="number" min={0} style={{ ...cellInput, width:80 }} value={it.qty} onChange={e=>updateManual(i,"qty",Number(e.target.value))}/></td>
                    <td style={td}><input type="number" min={0} style={{ ...cellInput, width:80 }} value={it.mrp} onChange={e=>updateManual(i,"mrp",Number(e.target.value))}/></td>
                    <td style={td}><button style={{ border:`1px solid ${C.redBg}`, background:"#fff", color:C.red, borderRadius:8, padding:"5px 8px", cursor:"pointer" }} onClick={()=>removeManualRow(i)}>✕</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ display:"flex", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
            <button style={{ ...btn(false), fontSize:12 }} onClick={addManualRow}>+ Add Row</button>
            <button style={btn(true)} onClick={()=>sendForApproval(manualItems.map(it=>({...it,editQty:it.qty})),newOrderId)}>Send for Approval</button>
          </div>
        </div>
      )}

      {/* Auto-generated order panel */}
      {orderItems && (
        <div className="pd-card" style={{ ...card, border:`1px solid ${C.indigo}` }}>
          <SectionHeader title={`ORDER — ${newOrderId}`} sub="Review quantities before sending for approval"
            action={<button style={{ ...btn(false), fontSize:12, padding:"6px 12px" }} onClick={()=>setOrderItems(null)}>Discard</button>}/>
          <div style={{ overflowX:"auto", marginBottom:14 }}>
            <table style={{ width:"100%", borderCollapse:"collapse" }}>
              <thead><tr>{["Product Code","Product Name","MRP","Order Qty",""].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
              <tbody>
                {orderItems.map((it,i)=>(
                  <tr key={it.productCode} style={trAlt(i)}>
                    <td style={td}>{it.productCode}</td>
                    <td style={{ ...td, fontWeight:600 }}>{it.productName}</td>
                    <td style={td}>₹{it.mrp}</td>
                    <td style={td}><input type="number" min={0} style={{ ...cellInput, width:80 }} value={it.editQty} onChange={e=>updateOrderQty(it.productCode,e.target.value)}/></td>
                    <td style={td}><button style={{ border:`1px solid ${C.redBg}`, background:"#fff", color:C.red, borderRadius:8, padding:"5px 8px", cursor:"pointer" }} onClick={()=>removeOrderItem(it.productCode)}>✕</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ display:"flex", justifyContent:"flex-end", gap:10 }}>
            <button style={{ ...btn(false), fontSize:12, padding:"7px 12px" }} onClick={()=>dlExcel(orderItems.map(it=>({ "Order ID":newOrderId,"Product Code":it.productCode,"Product Name":it.productName,MRP:it.mrp,"Order Qty":it.editQty })),"order.xlsx")}>⬇ Excel</button>
            <button style={btn(true)} onClick={()=>sendForApproval(orderItems,newOrderId)}>Send for Approval</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Vendiman Dashboard ───────────────────────────────────────────────────
function VendimanDashboard({ sites, showToast }) {
  const [siteId, setSiteId] = useState("");
  const [rows, setRows] = useState([]);
  const [whFile, setWhFile] = useState(null);
  const [whData, setWhData] = useState({});

  const load = (sid) => fetch(`${API_BASE}/dashboard?siteIds=${sid}`).then(r=>r.json()).then(d=>setRows(d.rows||[])).catch(()=>{});
  useEffect(()=>{ if(siteId) load(siteId); }, [siteId]);

  const handleWhUpload = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type:"array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet);
        const map = {};
        data.forEach(r=>{ const code=String(r["Product Code"]||"").trim(); if(code) map[code]=Number(r["WH Qty"])||0; });
        setWhData(map); setWhFile(file.name); showToast("WH stock uploaded");
      } catch { showToast("Could not read file"); }
    };
    reader.readAsArrayBuffer(file); e.target.value="";
  };

  return (
    <div>
      <div className="pd-card" style={{ ...card, background:C.indigoBg, border:`1px solid ${C.indigo}` }}>
        <div style={{ fontSize:13, color:C.indigo, fontWeight:600 }}>🔧 Vendiman Dashboard — Under Development</div>
        <div style={{ fontSize:12, color:C.sub, marginTop:4 }}>This section maps Vendiman WH availability against site stock levels.</div>
      </div>
      <div className="pd-card" style={card}>
        <SectionHeader title="VENDIMAN WH vs SITE STOCK" action={
          <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
            <select style={{ ...smInput, width:180 }} value={siteId} onChange={e=>setSiteId(e.target.value)}>
              <option value="">Select site</option>
              {sites.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            <label style={{ ...btn(false), display:"inline-flex", alignItems:"center", cursor:"pointer", fontSize:12, padding:"7px 12px" }}>
              ⬆ Upload WH Stocks
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleWhUpload} style={{ display:"none" }}/>
            </label>
          </div>
        }/>
        {whFile && <div style={{ fontSize:12, color:C.green, marginBottom:10 }}>✓ WH file loaded: {whFile}</div>}
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", minWidth:600 }}>
            <thead><tr>{["Product Code","Product Name","Brand","Category","Avail. Qty","Avail. Days","Status","WH Qty"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
            <tbody>
              {rows.map((r,i)=>(
                <tr key={`${r.siteId}-${r.productCode}`} style={trAlt(i)}>
                  <td style={td}>{r.productCode}</td>
                  <td style={{ ...td, fontWeight:600 }}>{r.productName}</td>
                  <td style={td}>{r.brand}</td>
                  <td style={td}><Tag label={r.category}/></td>
                  <td style={td}>{r.availableQty}</td>
                  <td style={{ ...td, color:r.status==="Critical"?C.red:r.status==="Surplus"?C.blue:C.green }}>{r.availableDays??"-"}</td>
                  <td style={td}><Pill label={r.status}/></td>
                  <td style={{ ...td, fontWeight:600 }}>{whData[r.productCode]??"-"}</td>
                </tr>
              ))}
              {rows.length===0 && <tr><td colSpan={8} style={{ ...td, color:C.sub }}>Select a site to view data.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── Client Approvals ─────────────────────────────────────────────────────
function ClientApprovals({ showToast, currentUser, sites }) {
  const [pending, setPending] = useState([]);
  const [approved, setApproved] = useState([]);
  const [expanded, setExpanded] = useState(null);
  const [feedback, setFeedback] = useState({});
  const [editQtys, setEditQtys] = useState({}); // orderId → {productCode → qty}

  const load = () => {
    fetch(`${API_BASE}/orders`).then(r=>r.json()).then(all=>{
      setPending(all.filter(o=>o.status==="Pending Approval"||o.status==="Feedback"));
      setApproved(all.filter(o=>o.status==="Approved"||o.status==="Rejected"));
    }).catch(()=>{});
  };
  useEffect(()=>{ load(); }, []);

  const siteName = (id) => sites.find(s=>s.id===id)?.name||`Site ${id}`;

  const update = (id, status, fb, items) => {
    const payload = { status, feedback:fb||"", approvedBy:status==="Approved"?currentUser.loginId:undefined };
    if (items) payload.items = items;
    fetch(`${API_BASE}/orders/${id}`, { method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload) })
      .then(()=>{ showToast(`Order ${status}`); load(); setExpanded(null); }).catch(()=>showToast("Could not reach backend"));
  };

  const approveWithEdits = (o) => {
    const qmap = editQtys[o.id]||{};
    const updatedItems = o.items.map(it=>({ ...it, qty:Number(qmap[it.productCode]??it.qty) }));
    update(o.id,"Approved","",updatedItems);
  };

  const canApprove = currentUser.role==="admin"||currentUser.role==="clientApprover";

  const OrderCard = ({ o, i, showActions }) => (
    <div key={o.id} style={{ border:`1px solid ${C.border}`, borderRadius:12, marginBottom:12, overflow:"hidden" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8, padding:"12px 14px", background:i%2===0?C.panel:"#fff" }}>
        <div>
          <div style={{ fontWeight:700, fontSize:14 }}>{o.id}</div>
          <div style={{ fontSize:12, color:C.sub }}>
            Requestor: {o.placedBy} · Site: {siteName(o.siteId)} · {new Date(o.createdAt).toLocaleDateString("en-IN")}
            {o.approvedBy && ` · Approved by: ${o.approvedBy}`}
          </div>
        </div>
        <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
          <Pill label={o.status}/>
          <button style={{ ...btn(false), fontSize:12, padding:"6px 10px" }} onClick={()=>setExpanded(expanded===o.id?null:o.id)}>{expanded===o.id?"Hide":"View"}</button>
          {showActions && canApprove && <>
            <button style={{ ...btn(true), fontSize:12, padding:"6px 10px" }} onClick={()=>approveWithEdits(o)}>Approve</button>
            <button style={{ ...btn(false), fontSize:12, padding:"6px 10px", color:C.red }} onClick={()=>update(o.id,"Rejected","")}>Reject</button>
          </>}
        </div>
      </div>
      {expanded===o.id && (
        <div style={{ padding:"14px" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", marginBottom:10 }}>
            <thead><tr>{["Product Code","Product Name",showActions&&canApprove?"Edit Qty":"Qty","MRP"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
            <tbody>{o.items?.map((it,j)=>(
              <tr key={it.productCode} style={trAlt(j)}>
                <td style={td}>{it.productCode}</td>
                <td style={{ ...td,fontWeight:600 }}>{it.productName}</td>
                <td style={td}>
                  {showActions && canApprove
                    ? <input type="number" min={0} style={{ ...cellInput, width:80 }}
                        value={editQtys[o.id]?.[it.productCode]??it.qty}
                        onChange={e=>setEditQtys(p=>({ ...p,[o.id]:{ ...(p[o.id]||{}), [it.productCode]:e.target.value } }))}/>
                    : it.qty}
                </td>
                <td style={td}>₹{it.mrp}</td>
              </tr>
            ))}</tbody>
          </table>
          {showActions && canApprove && (
            <div style={{ marginTop:8 }}>
              <div style={{ fontSize:12, color:C.sub, marginBottom:6 }}>Give Feedback</div>
              <div style={{ display:"flex", gap:8 }}>
                <input style={{ ...inputStyle, flex:1 }} value={feedback[o.id]||""} onChange={e=>setFeedback(p=>({...p,[o.id]:e.target.value}))} placeholder="Add comment…"/>
                <button style={btn(false)} onClick={()=>update(o.id,"Feedback",feedback[o.id])}>Send Feedback</button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );

  return (
    <div>
      <div className="pd-card" style={card}>
        <SectionHeader title="PENDING APPROVALS" sub={`${pending.length} order(s) awaiting review`}/>
        {pending.length===0
          ? <div style={{ padding:"16px", color:C.sub, fontSize:13 }}>No orders pending approval.</div>
          : pending.map((o,i)=><OrderCard key={o.id} o={o} i={i} showActions={true}/>)}
      </div>
      <div className="pd-card" style={card}>
        <SectionHeader title="APPROVED / REJECTED HISTORY" sub="All completed approvals"/>
        {approved.length===0
          ? <div style={{ padding:"16px", color:C.sub, fontSize:13 }}>No approved/rejected orders yet.</div>
          : approved.map((o,i)=><OrderCard key={o.id} o={o} i={i} showActions={false}/>)}
      </div>
    </div>
  );
}

// ─── Orders ───────────────────────────────────────────────────────────────
function Orders({ sites, showToast }) {
  const [siteId, setSiteId] = useState("");
  const [orders, setOrders] = useState([]);
  const [expanded, setExpanded] = useState(null);
  const [partialData, setPartialData] = useState({}); // orderId → {fulfilledQty, pendingQty, remarks}
  const FULFILL_OPTIONS = ["Fulfilled 100%","Fulfilled partially","Dropped","Others"];

  const load = () => {
    const url = siteId ? `${API_BASE}/orders?siteId=${siteId}` : `${API_BASE}/orders`;
    fetch(url).then(r=>r.json()).then(setOrders).catch(()=>{});
  };
  useEffect(()=>{ load(); }, [siteId]);

  const siteName = (id) => sites.find(s=>s.id===id)?.name||`Site ${id}`;

  const updateFulfillment = (id, status) => {
    const pd = partialData[id]||{};
    fetch(`${API_BASE}/orders/${id}`, { method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ fulfillmentStatus:status, fulfilledQty:pd.fulfilledQty!=null?Number(pd.fulfilledQty):null, pendingQty:pd.pendingQty!=null?Number(pd.pendingQty):null, fulfillmentRemarks:pd.remarks||"" }) })
      .then(()=>{ showToast("Status updated"); load(); }).catch(()=>showToast("Could not reach backend"));
  };

  const printOrder = (o) => {
    const w = window.open("","_blank","width=700,height=600");
    w.document.write(`<html><body style="font-family:sans-serif;padding:24px">
      <h2>Order: ${o.id}</h2>
      <p><b>Requestor:</b> ${o.placedBy||"—"} &nbsp;|&nbsp; <b>Approver:</b> ${o.approvedBy||"—"} &nbsp;|&nbsp; <b>Site:</b> ${siteName(o.siteId)} &nbsp;|&nbsp; <b>Date:</b> ${new Date(o.createdAt).toLocaleDateString()}</p>
      <table border="1" cellpadding="8" style="border-collapse:collapse;width:100%">
        <tr><th>Product Code</th><th>Product Name</th><th>Qty</th><th>MRP</th></tr>
        ${o.items?.map(it=>`<tr><td>${it.productCode}</td><td>${it.productName}</td><td>${it.qty}</td><td>₹${it.mrp}</td></tr>`).join("")}
      </table>
      ${o.fulfillmentStatus?`<p><b>Fulfillment:</b> ${o.fulfillmentStatus} | Fulfilled Qty: ${o.fulfilledQty??"-"} | Pending Qty: ${o.pendingQty??"-"} | Remarks: ${o.fulfillmentRemarks||"-"}</p>`:""}
    </body></html>`);
    w.document.close(); w.print();
  };

  const exportOrder = (o) => dlExcel([
    ...o.items.map(it=>({ "Order ID":o.id,"Requestor":o.placedBy,"Approver":o.approvedBy||"","Site":siteName(o.siteId),"Product Code":it.productCode,"Product Name":it.productName,"Qty":it.qty,"MRP":it.mrp })),
  ], `order-${o.id}.xlsx`);

  const activeOrders = orders.filter(o=>o.status==="Approved"&&!o.fulfillmentStatus);
  const historyOrders = orders.filter(o=>o.fulfillmentStatus||o.status==="Rejected");

  return (
    <div>
      <div className="pd-card" style={card}>
        <SectionHeader title="ORDERS" action={
          <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
            <select style={{ ...smInput, width:180 }} value={siteId} onChange={e=>setSiteId(e.target.value)}>
              <option value="">All sites</option>
              {sites.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            <button style={{ ...btn(false), fontSize:12, padding:"7px 12px" }} onClick={()=>dlExcel(orders.map(o=>({ "Order ID":o.id,"Requestor":o.placedBy,"Approver":o.approvedBy||"","Site":siteName(o.siteId),Status:o.status,"Fulfillment":o.fulfillmentStatus||"","Date":new Date(o.createdAt).toLocaleDateString() })),"order-history.xlsx")}>⬇ History Excel</button>
          </div>
        }/>
        {orders.filter(o=>!o.fulfillmentStatus&&o.status!=="Rejected").length===0
          ? <div style={{ padding:"16px", color:C.sub, fontSize:13 }}>No active orders.</div>
          : orders.filter(o=>!o.fulfillmentStatus&&o.status!=="Rejected").map((o,i)=>(
            <div key={o.id} style={{ border:`1px solid ${C.border}`, borderRadius:12, marginBottom:12, overflow:"hidden" }}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:8, padding:"12px 14px", background:i%2===0?C.panel:"#fff" }}>
                <div>
                  <div style={{ fontWeight:700, fontSize:14 }}>{o.id}</div>
                  <div style={{ fontSize:12, color:C.sub }}>
                    Requestor: {o.placedBy||"—"} · Approver: {o.approvedBy||"Pending"} · Site: {siteName(o.siteId)} · {new Date(o.createdAt).toLocaleDateString("en-IN")}
                  </div>
                </div>
                <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"center" }}>
                  <Pill label={o.status}/>
                  <button style={{ ...btn(false), fontSize:12, padding:"6px 10px" }} onClick={()=>setExpanded(expanded===o.id?null:o.id)}>View</button>
                  <button style={{ ...btn(false), fontSize:12, padding:"6px 10px" }} onClick={()=>printOrder(o)}>PDF</button>
                  <button style={{ ...btn(false), fontSize:12, padding:"6px 10px" }} onClick={()=>exportOrder(o)}>Excel</button>
                </div>
              </div>
              {expanded===o.id && (
                <div style={{ padding:"14px" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", marginBottom:12 }}>
                    <thead><tr>{["Product Code","Product Name","Qty","MRP"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                    <tbody>{o.items?.map((it,j)=><tr key={it.productCode} style={trAlt(j)}><td style={td}>{it.productCode}</td><td style={{ ...td,fontWeight:600 }}>{it.productName}</td><td style={td}>{it.qty}</td><td style={td}>₹{it.mrp}</td></tr>)}</tbody>
                  </table>
                  {o.status==="Approved" && (
                    <div>
                      <div style={{ fontSize:12, color:C.sub, marginBottom:8, fontWeight:600 }}>Update Fulfillment Status</div>
                      <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:10 }}>
                        {FULFILL_OPTIONS.map(s=>(
                          <button key={s} style={{ ...btn(false), padding:"7px 12px", fontSize:12 }} onClick={()=>{ if(s!=="Fulfilled partially"){ updateFulfillment(o.id,s); } else { setPartialData(p=>({...p,[o.id]:{...p[o.id],show:true}})); } }}>{s}</button>
                        ))}
                      </div>
                      {partialData[o.id]?.show && (
                        <div style={{ padding:"12px", background:C.amberBg, borderRadius:10, border:`1px solid ${C.amber}` }}>
                          <div style={{ fontSize:12, fontWeight:700, color:C.amber, marginBottom:10 }}>Partial Fulfillment Details</div>
                          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:10, marginBottom:10 }}>
                            {[["fulfilledQty","Fulfilled Qty"],["pendingQty","Pending/Dropped Qty"]].map(([f,l])=>(
                              <div key={f}><div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>{l}</div>
                                <input type="number" min={0} style={smInput} value={partialData[o.id]?.[f]||""} onChange={e=>setPartialData(p=>({...p,[o.id]:{...p[o.id],[f]:e.target.value}}))} placeholder="0"/>
                              </div>
                            ))}
                            <div style={{ gridColumn:"1/-1" }}><div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>Remarks / Reason</div>
                              <input style={smInput} value={partialData[o.id]?.remarks||""} onChange={e=>setPartialData(p=>({...p,[o.id]:{...p[o.id],remarks:e.target.value}}))} placeholder="Reason for partial fulfillment…"/>
                            </div>
                          </div>
                          <div style={{ display:"flex", justifyContent:"flex-end", gap:8 }}>
                            <button style={{ ...btn(false), fontSize:12 }} onClick={()=>setPartialData(p=>({...p,[o.id]:{...p[o.id],show:false}}))}>Cancel</button>
                            <button style={btn(true)} onClick={()=>updateFulfillment(o.id,"Fulfilled partially")}>Confirm Partial Fulfillment</button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}
            </div>
          ))
        }
      </div>

      {/* Order History */}
      <div className="pd-card" style={card}>
        <SectionHeader title="ORDER HISTORY" sub="Fulfilled, dropped and rejected orders"/>
        {historyOrders.length===0
          ? <div style={{ padding:"16px", color:C.sub, fontSize:13 }}>No historical orders yet.</div>
          : <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse" }}>
                <thead><tr>{["Order ID","Site","Requestor","Approver","Status","Fulfillment","Fulfilled Qty","Pending Qty","Remarks","Date",""].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                <tbody>
                  {historyOrders.map((o,i)=>(
                    <tr key={o.id} style={trAlt(i)}>
                      <td style={{ ...td, fontWeight:700 }}>{o.id}</td>
                      <td style={td}>{siteName(o.siteId)}</td>
                      <td style={td}>{o.placedBy}</td>
                      <td style={td}>{o.approvedBy||"—"}</td>
                      <td style={td}><Pill label={o.status}/></td>
                      <td style={td}>{o.fulfillmentStatus||"—"}</td>
                      <td style={td}>{o.fulfilledQty??"-"}</td>
                      <td style={td}>{o.pendingQty??"-"}</td>
                      <td style={td}>{o.fulfillmentRemarks||"—"}</td>
                      <td style={td}>{new Date(o.createdAt).toLocaleDateString("en-IN")}</td>
                      <td style={td}><button style={{ ...btn(false), fontSize:12, padding:"5px 8px" }} onClick={()=>exportOrder(o)}>Excel</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
        }
      </div>
    </div>
  );
}

// ─── Reports ──────────────────────────────────────────────────────────────
function Reports({ sites, showToast }) {
  const today = new Date().toISOString().slice(0,10);
  const monthStart = new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString().slice(0,10);
  const [activeReport, setActiveReport] = useState("inward");
  const [startDate, setStartDate] = useState(monthStart);
  const [endDate, setEndDate] = useState(today);
  const [selSites, setSelSites] = useState([]);
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);

  const toggleSite = (id) => setSelSites(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);

  const load = () => {
    const siteQ = selSites.length ? `&siteIds=${selSites.join(",")}` : "";
    const url = `${API_BASE}/reports/${activeReport}?startDate=${startDate}&endDate=${endDate}${siteQ}`;
    setLoading(true);
    fetch(url).then(r=>r.json()).then(d=>{ setData(Array.isArray(d)?d:[]); setLoading(false); }).catch(()=>{ showToast("Could not reach backend"); setLoading(false); });
  };
  useEffect(()=>{ setData([]); }, [activeReport]);

  const REPORTS = [
    { key:"inward",  label:"Stock Inward Report",  icon:"📦" },
    { key:"stock-offtake", label:"Stock Offtake Report", icon:"📉" },
    { key:"orders",  label:"Orders Report",         icon:"🛒" },
  ];

  const summaryStats = () => {
    if (!data.length) return [];
    if (activeReport==="inward") return [
      { label:"Total Inward Records", value:data.length },
      { label:"Total Qty Inwarded", value:data.reduce((s,r)=>s+(r.qty||0),0).toLocaleString("en-IN") },
      { label:"Unique Products", value:new Set(data.map(r=>r.productCode)).size },
      { label:"Unique Vendors", value:new Set(data.map(r=>r.vendorName).filter(Boolean)).size },
    ];
    if (activeReport==="stock-offtake") return [
      { label:"Total Products", value:data.length },
      { label:"Total Offtake", value:data.reduce((s,r)=>s+(r.offtake||0),0).toLocaleString("en-IN") },
      { label:"Avg Qty/Day", value:(data.reduce((s,r)=>s+(r.qtyPerDay||0),0)/Math.max(1,data.length)).toFixed(2) },
      { label:"Days in Period", value:data[0]?.days||"-" },
    ];
    if (activeReport==="orders") return [
      { label:"Total Orders", value:data.length },
      { label:"Approved", value:data.filter(o=>o.status==="Approved").length },
      { label:"Pending", value:data.filter(o=>o.status==="Pending Approval").length },
      { label:"Fulfilled", value:data.filter(o=>o.fulfillmentStatus).length },
    ];
    return [];
  };

  const exportReport = () => {
    if (!data.length) { showToast("No data to export"); return; }
    let rows;
    if (activeReport==="inward") rows = data.map(r=>({ "DC Code":r.dcCode,"Date":new Date(r.date).toLocaleDateString("en-IN"),"Vendor":r.vendorName,"Site":r.siteName,"Product Code":r.productCode,"Product Name":r.productName,"Brand":r.brand,"Category":r.category,"Sub Category":r.subCategory,"HSN Code":r.hsnCode,"MRP":r.mrp,"GST%":r.gst,"UOM":r.uom,"Weight":r.weight,"Qty Inwarded":r.qty }));
    else if (activeReport==="stock-offtake") rows = data.map(r=>({ "Product Code":r.productCode,"Product Name":r.productName,"Brand":r.brand,"Category":r.category,"Sub Category":r.subCategory,"MRP":r.mrp,"UOM":r.uom,"Opening Stock":r.openStock,"Inwarded":r.inwarded,"Closing Stock":r.closeStock,"Offtake":r.offtake,"Days in Period":r.days,"Qty Per Day":r.qtyPerDay }));
    else rows = data.map(r=>({ "Order ID":r.id,"Site":r.siteName,"Requestor":r.placedBy,"Approver":r.approvedBy||"","Status":r.status,"Fulfillment":r.fulfillmentStatus||"","Fulfilled Qty":r.fulfilledQty||"","Pending Qty":r.pendingQty||"","Remarks":r.fulfillmentRemarks||"","Date":new Date(r.createdAt).toLocaleDateString("en-IN") }));
    dlExcel(rows, `${activeReport}-report-${startDate}-${endDate}.xlsx`);
  };

  const INWARD_COLS  = ["Date","DC Code","Vendor","Site","Product Code","Product Name","Brand","Category","Sub Category","Qty"];
  const OFFTAKE_COLS = ["Product Code","Product Name","Brand","Category","MRP","Opening Stock","Inwarded","Closing Stock","Offtake","Qty/Day"];
  const ORDERS_COLS  = ["Order ID","Site","Requestor","Approver","Status","Fulfillment","Fulfilled Qty","Date"];

  return (
    <div>
      {/* Report selector */}
      <div className="pd-card" style={card}>
        <div style={{ display:"flex", gap:10, flexWrap:"wrap", marginBottom:16 }}>
          {REPORTS.map(r=>(
            <button key={r.key} onClick={()=>setActiveReport(r.key)} style={{ ...btn(activeReport===r.key), padding:"10px 16px", fontSize:13 }}>
              {r.icon} {r.label}
            </button>
          ))}
        </div>
        {/* Filters */}
        <div style={{ display:"flex", gap:10, flexWrap:"wrap", alignItems:"flex-end", marginBottom:12 }}>
          <div><div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>From Date</div><input type="date" style={{ ...smInput, width:160 }} value={startDate} onChange={e=>setStartDate(e.target.value)}/></div>
          <div><div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>To Date</div><input type="date" style={{ ...smInput, width:160 }} value={endDate} onChange={e=>setEndDate(e.target.value)}/></div>
          <button style={btn(true)} onClick={load} disabled={loading}>{loading?"Loading…":"Generate Report"}</button>
          {data.length>0 && <button style={{ ...btn(false), fontSize:12, padding:"9px 14px" }} onClick={exportReport}>⬇ Download Excel</button>}
        </div>
        <div>
          <div style={{ fontSize:11, color:C.sub, marginBottom:6 }}>Filter by Site (optional)</div>
          <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
            {sites.map(s=>(
              <button key={s.id} onClick={()=>toggleSite(s.id)} style={{ ...btn(selSites.includes(s.id)), padding:"5px 12px", fontSize:12 }}>{s.name}</button>
            ))}
            {selSites.length>0 && <button onClick={()=>setSelSites([])} style={{ ...btn(false), padding:"5px 12px", fontSize:12, color:C.sub }}>Clear</button>}
          </div>
        </div>
      </div>

      {/* Summary stats */}
      {data.length>0 && (
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:10, marginBottom:14 }}>
          {summaryStats().map(s=>(
            <div key={s.label} className="pd-tile" style={{ background:C.panel, borderRadius:14, padding:"16px 14px", border:`1px solid ${C.border}` }}>
              <div style={{ fontSize:22, fontWeight:700, color:C.blue }}>{s.value}</div>
              <div style={{ fontSize:11, color:C.sub, marginTop:4 }}>{s.label}</div>
            </div>
          ))}
        </div>
      )}

      {/* Data table */}
      {data.length>0 && (
        <div className="pd-card" style={card}>
          <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", minWidth:600 }}>
              <thead><tr>
                {(activeReport==="inward"?INWARD_COLS:activeReport==="stock-offtake"?OFFTAKE_COLS:ORDERS_COLS).map(h=><th key={h} style={th}>{h}</th>)}
              </tr></thead>
              <tbody>
                {data.map((r,i)=>(
                  <tr key={i} style={trAlt(i)}>
                    {activeReport==="inward" && [new Date(r.date).toLocaleDateString("en-IN"),r.dcCode,r.vendorName,r.siteName,r.productCode,r.productName,r.brand,r.category,r.subCategory,r.qty].map((v,j)=><td key={j} style={td}>{v||"—"}</td>)}
                    {activeReport==="stock-offtake" && [r.productCode,r.productName,r.brand,r.category,`₹${r.mrp}`,r.openStock,r.inwarded,r.closeStock,r.offtake,r.qtyPerDay].map((v,j)=><td key={j} style={{ ...td, fontWeight:j===9?700:400 }}>{v??"-"}</td>)}
                    {activeReport==="orders" && [r.id,r.siteName,r.placedBy,r.approvedBy||"—",<Pill key="s" label={r.status}/>,r.fulfillmentStatus||"—",r.fulfilledQty??"-",new Date(r.createdAt).toLocaleDateString("en-IN")].map((v,j)=><td key={j} style={td}>{v}</td>)}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
      {!loading && data.length===0 && <div style={{ padding:"20px 16px", color:C.sub, fontSize:13, textAlign:"center" }}>Select filters and click "Generate Report" to view data.</div>}
    </div>
  );
}

// ─── Admin ─────────────────────────────────────────────────────────────────
function Admin({ sites, setSites, vendors, setVendors, products, setProducts, statusDefs, setStatusDefs, users, setUsers, showToast }) {
  const [subTab, setSubTab] = useState("Products");
  const subTabs = ["Products","Sites","Vendors","Status Definition","User Access"];

  return (
    <div>
      {/* Sub-tab nav */}
      <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginBottom:16, padding:"12px 0 0" }}>
        {subTabs.map(t=>(
          <button key={t} onClick={()=>setSubTab(t)} style={{ ...btn(subTab===t), padding:"8px 14px", fontSize:13 }}>{t}</button>
        ))}
      </div>
      {subTab==="Products"       && <AdminProducts products={products} setProducts={setProducts} showToast={showToast}/>}
      {subTab==="Sites"          && <AdminSites sites={sites} setSites={setSites} showToast={showToast}/>}
      {subTab==="Vendors"        && <AdminVendors vendors={vendors} setVendors={setVendors} products={products} sites={sites} showToast={showToast}/>}
      {subTab==="Status Definition" && <AdminStatusDef sites={sites} showToast={showToast}/>}
      {subTab==="User Access"    && <AdminUserAccess users={users} setUsers={setUsers} sites={sites} showToast={showToast}/>}
    </div>
  );
}

// ── Admin > Products ───────────────────────────────────────────────────────
const PRODUCT_FIELDS = ["code","name","brand","category","subCategory","mrp","caseSize","uom","hsnCode","gst","weight","status"];
const PRODUCT_LABELS = { code:"Product Code", name:"Product Name", brand:"Brand", category:"Category", subCategory:"Sub Category", mrp:"MRP", caseSize:"Case Size", uom:"UOM", hsnCode:"HSN Code", gst:"GST%", weight:"Weight", status:"Status" };
const EMPTY_PROD = { code:"", name:"", brand:"", category:"", subCategory:"", mrp:"", caseSize:"", uom:"", hsnCode:"", gst:"", weight:"", status:"Active" };

function AdminProducts({ products, setProducts, showToast }) {
  const [sortCol, setSortCol] = useState("name");
  const [sortDir, setSortDir] = useState(1);
  const [newProd, setNewProd] = useState(EMPTY_PROD);
  const [nameSugs, setNameSugs] = useState([]);
  const [errors, setErrors] = useState({});
  const [editCode, setEditCode] = useState(null); // code of row being edited
  const [editData, setEditData] = useState({});

  const load = () => fetch(`${API_BASE}/products`).then(r=>r.json()).then(setProducts).catch(()=>{});
  useEffect(()=>{ load(); }, []);

  const sorted = [...products].sort((a,b)=>{
    const av=a[sortCol]||"", bv=b[sortCol]||"";
    return (av<bv?-1:av>bv?1:0)*sortDir;
  });
  const toggleSort = (col) => { if(sortCol===col) setSortDir(d=>-d); else { setSortCol(col); setSortDir(1); }};

  const handleNameSearch = (val) => {
    setNewProd(p=>({...p,name:val}));
    setNameSugs(val.trim().length>1 ? products.filter(p=>p.name.toLowerCase().includes(val.toLowerCase()) && p.code!==editCode).slice(0,5) : []);
  };

  const validate = (data, isNew) => {
    const e = {};
    if (!data.code?.trim()) e.code="Required";
    if (!data.name?.trim()) e.name="Required";
    if (isNew && products.find(p=>p.code===data.code?.trim())) e.code=`Code "${data.code}" already exists`;
    if (products.find(p=>p.name?.toLowerCase()===data.name?.trim().toLowerCase() && (isNew || p.code!==editCode))) e.name=`Name "${data.name}" already exists`;
    setErrors(e);
    return Object.keys(e).length===0;
  };

  const addProduct = () => {
    if (!validate(newProd, true)) return;
    fetch(`${API_BASE}/products`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ ...newProd, mrp:Number(newProd.mrp)||0, caseSize:Number(newProd.caseSize)||0, gst:Number(newProd.gst)||0 }) })
      .then(r=>r.json())
      .then(d=>{ if(d.error){ showToast(d.error); return; } setProducts(p=>[...p,d]); setNewProd(EMPTY_PROD); setNameSugs([]); setErrors({}); showToast("Product added"); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const startEdit = (p) => { setEditCode(p.code); setEditData({...p}); setErrors({}); };
  const cancelEdit = () => { setEditCode(null); setEditData({}); setErrors({}); };

  const saveEdit = () => {
    if (!validate(editData, false)) return;
    fetch(`${API_BASE}/products/${editCode}`, { method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ ...editData, mrp:Number(editData.mrp)||0, caseSize:Number(editData.caseSize)||0, gst:Number(editData.gst)||0 }) })
      .then(r=>r.json())
      .then(d=>{ if(d.error){ showToast(d.error); return; } setProducts(p=>p.map(x=>x.code===editCode?d:x)); cancelEdit(); showToast("Product updated"); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const deleteProduct = (code) => {
    if (!window.confirm(`Delete product "${code}"?`)) return;
    fetch(`${API_BASE}/products/${code}`, { method:"DELETE" }).then(()=>{ setProducts(p=>p.filter(x=>x.code!==code)); showToast("Deleted"); }).catch(()=>showToast("Could not reach backend"));
  };

  const exportExcel = () => dlExcel(products.map(p=>({ "Product Code":p.code,"Product Name":p.name,Brand:p.brand,Category:p.category,"Sub Category":p.subCategory,MRP:p.mrp,"Case Size":p.caseSize,UOM:p.uom,"HSN Code":p.hsnCode,"GST%":p.gst,"Weight":p.weight,Status:p.status })),"products.xlsx");
  const SortTh = ({col,label}) => (<th style={{ ...th, cursor:"pointer", userSelect:"none" }} onClick={()=>toggleSort(col)}>{label}{sortCol===col?(sortDir===1?" ↑":" ↓"):""}</th>);

  const ProductForm = ({ data, setData, isEdit }) => (
    <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:10, marginBottom:14 }}>
      {PRODUCT_FIELDS.map(f=>(
        <div key={f}>
          <div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>{PRODUCT_LABELS[f]}</div>
          {f==="status" ? (
            <select style={smInput} value={data[f]||"Active"} onChange={e=>setData(p=>({...p,[f]:e.target.value}))}>
              <option>Active</option><option>Inactive</option>
            </select>
          ) : f==="code" && isEdit ? (
            <input style={{ ...smInput, background:C.panel, color:C.sub }} value={data[f]||""} disabled />
          ) : f==="name" ? (
            <div style={{ position:"relative" }}>
              <input style={{ ...smInput, borderColor:errors[f]?C.red:C.border }} value={data[f]||""} onChange={e=>{ setData(p=>({...p,[f]:e.target.value})); if(!isEdit) handleNameSearch(e.target.value); }} placeholder={PRODUCT_LABELS[f]}/>
              {errors[f] && <div style={{ fontSize:10, color:C.red, marginTop:2 }}>{errors[f]}</div>}
              {!isEdit && nameSugs.length>0 && (
                <div style={{ position:"absolute", top:"100%", left:0, right:0, background:"#fff", border:`1px solid ${C.border}`, borderRadius:8, zIndex:10, boxShadow:"0 4px 12px rgba(0,0,0,.1)" }}>
                  {nameSugs.map(s=><div key={s.code} onMouseDown={e=>e.preventDefault()} style={{ padding:"8px 10px", fontSize:12, color:C.amber, borderBottom:`1px solid ${C.border}`, cursor:"pointer" }} onClick={()=>{ showToast(`"${s.name}" already exists (${s.code})`); setNameSugs([]); }}>⚠ {s.name} ({s.code})</div>)}
                </div>
              )}
            </div>
          ) : (
            <input style={{ ...smInput, borderColor:errors[f]?C.red:C.border }} value={data[f]||""} onChange={e=>setData(p=>({...p,[f]:e.target.value}))} placeholder={PRODUCT_LABELS[f]}/>
          )}
          {f!=="name" && errors[f] && <div style={{ fontSize:10, color:C.red, marginTop:2 }}>{errors[f]}</div>}
        </div>
      ))}
    </div>
  );

  return (
    <div>
      {/* Edit panel */}
      {editCode && (
        <div className="pd-card" style={{ ...card, border:`2px solid ${C.blue}`, background:C.blueBg }}>
          <SectionHeader title={`EDIT PRODUCT — ${editCode}`} action={<button style={{ ...btn(false), fontSize:12, padding:"6px 12px" }} onClick={cancelEdit}>Cancel</button>}/>
          <ProductForm data={editData} setData={setEditData} isEdit={true}/>
          <div style={{ display:"flex", justifyContent:"flex-end", gap:10 }}>
            <button style={btn(false)} onClick={cancelEdit}>Cancel</button>
            <button style={btn(true)} onClick={saveEdit}>Save Changes</button>
          </div>
        </div>
      )}

      {/* Add form */}
      {!editCode && (
        <div className="pd-card" style={card}>
          <SectionHeader title="ADD PRODUCT" />
          <ProductForm data={newProd} setData={setNewProd} isEdit={false}/>
          <div style={{ display:"flex", justifyContent:"flex-end" }}>
            <button style={btn(true)} onClick={addProduct}>+ Add Product</button>
          </div>
        </div>
      )}

      {/* Products table */}
      <div className="pd-card" style={card}>
        <SectionHeader title={`PRODUCTS (${products.length})`} action={
          <button style={{ ...btn(false), fontSize:12, padding:"7px 12px" }} onClick={exportExcel}>⬇ Excel</button>
        }/>
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", minWidth:900 }}>
            <thead><tr>
              {PRODUCT_FIELDS.map(f=><SortTh key={f} col={f} label={PRODUCT_LABELS[f]}/>)}
              <th style={th}>Actions</th>
            </tr></thead>
            <tbody>
              {sorted.map((p,i)=>(
                <tr key={p.code} style={{ ...trAlt(i), outline: editCode===p.code?`2px solid ${C.blue}`:"none" }}>
                  {PRODUCT_FIELDS.map(f=><td key={f} style={td}>{f==="status"?<Pill label={p[f]}/>:p[f]||"—"}</td>)}
                  <td style={td}>
                    <div style={{ display:"flex", gap:6 }}>
                      <button style={{ ...btn(false), padding:"5px 10px", fontSize:12, color:C.blue, borderColor:C.blue+"44" }} onClick={()=>startEdit(p)}>Edit</button>
                      <button style={{ border:`1px solid ${C.redBg}`, background:"#fff", color:C.red, borderRadius:8, padding:"5px 8px", cursor:"pointer", fontSize:12 }} onClick={()=>deleteProduct(p.code)}>Delete</button>
                    </div>
                  </td>
                </tr>
              ))}
              {products.length===0 && <tr><td colSpan={PRODUCT_FIELDS.length+1} style={{ ...td, color:C.sub }}>No products yet.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ── Admin > Sites ──────────────────────────────────────────────────────────
function AdminSites({ sites, setSites, showToast }) {
  const [name, setName] = useState("");
  const [days, setDays] = useState(5);
  const [editId, setEditId] = useState(null);
  const [editData, setEditData] = useState({});

  const load = () => fetch(`${API_BASE}/sites`).then(r=>r.json()).then(setSites).catch(()=>{});
  useEffect(()=>{ load(); }, []);

  const add = () => {
    if (!name.trim()) { showToast("Site name required"); return; }
    fetch(`${API_BASE}/sites`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ name:name.trim(), preferredDays:Number(days)||5 }) })
      .then(r=>r.json()).then(d=>{ if(d.error){ showToast(d.error); return; } setSites(s=>[...s,d]); setName(""); setDays(5); showToast("Site added"); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const startEdit = (s) => { setEditId(s.id); setEditData({ name:s.name, preferredDays:s.preferredDays||5 }); };
  const cancelEdit = () => { setEditId(null); setEditData({}); };

  const saveEdit = () => {
    if (!editData.name?.trim()) { showToast("Site name required"); return; }
    fetch(`${API_BASE}/sites/${editId}`, { method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ name:editData.name.trim(), preferredDays:Number(editData.preferredDays)||5 }) })
      .then(r=>r.json()).then(d=>{ if(d.error){ showToast(d.error); return; } setSites(s=>s.map(x=>x.id===editId?d:x)); cancelEdit(); showToast("Site updated"); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const del = (id) => {
    if (!window.confirm("Delete this site?")) return;
    fetch(`${API_BASE}/sites/${id}`, { method:"DELETE" }).then(()=>setSites(s=>s.filter(x=>x.id!==id))).catch(()=>{});
  };

  return (
    <div>
      <div className="pd-card" style={card}>
        <SectionHeader title="ADD SITE" />
        <div style={{ display:"flex", gap:10, flexWrap:"wrap", marginBottom:10 }}>
          <div style={{ flex:"1 1 180px" }}><div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>Site Name</div><input style={inputStyle} value={name} onChange={e=>setName(e.target.value)} placeholder="e.g. Bengaluru" /></div>
          <div style={{ flex:"0 0 160px" }}><div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>Preferred Days of Stock</div><input type="number" min={1} style={inputStyle} value={days} onChange={e=>setDays(e.target.value)} /></div>
        </div>
        <div style={{ display:"flex", justifyContent:"flex-end" }}><button style={btn(true)} onClick={add}>+ Add Site</button></div>
      </div>
      <div className="pd-card" style={card}>
        <SectionHeader title="SITES" />
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse" }}>
            <thead><tr>{["Site Name","Preferred Days of Stock","Actions"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
            <tbody>
              {sites.map((s,i)=>(
                <tr key={s.id} style={{ ...trAlt(i), outline:editId===s.id?`2px solid ${C.blue}`:"none" }}>
                  <td style={{ ...td, fontWeight:600 }}>
                    {editId===s.id
                      ? <input style={cellInput} value={editData.name} onChange={e=>setEditData(p=>({...p,name:e.target.value}))}/>
                      : s.name}
                  </td>
                  <td style={td}>
                    {editId===s.id
                      ? <input type="number" min={1} style={{ ...cellInput, width:80 }} value={editData.preferredDays} onChange={e=>setEditData(p=>({...p,preferredDays:e.target.value}))}/>
                      : s.preferredDays||5}
                  </td>
                  <td style={td}>
                    {editId===s.id ? (
                      <div style={{ display:"flex", gap:6 }}>
                        <button style={{ ...btn(true), padding:"5px 10px", fontSize:12 }} onClick={saveEdit}>Save</button>
                        <button style={{ ...btn(false), padding:"5px 10px", fontSize:12 }} onClick={cancelEdit}>Cancel</button>
                      </div>
                    ) : (
                      <div style={{ display:"flex", gap:6 }}>
                        <button style={{ ...btn(false), padding:"5px 10px", fontSize:12, color:C.blue, borderColor:C.blue+"44" }} onClick={()=>startEdit(s)}>Edit</button>
                        <button style={{ border:`1px solid ${C.redBg}`, background:"#fff", color:C.red, borderRadius:8, padding:"5px 8px", cursor:"pointer", fontSize:12 }} onClick={()=>del(s.id)}>Delete</button>
                      </div>
                    )}
                  </td>
                </tr>
              ))}
              {sites.length===0 && <tr><td colSpan={3} style={{ ...td, color:C.sub }}>No sites yet.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ── Admin > Vendors ────────────────────────────────────────────────────────
function AdminVendors({ vendors, setVendors, products, sites, showToast }) {
  const [name, setName] = useState("");
  const [alloc, setAlloc] = useState("all");
  const [selProds, setSelProds] = useState([]);
  const [siteAlloc, setSiteAlloc] = useState("all");
  const [selSites, setSelSites] = useState([]);
  const [editId, setEditId] = useState(null);
  const [editData, setEditData] = useState({});

  const load = () => fetch(`${API_BASE}/vendors`).then(r=>r.json()).then(setVendors).catch(()=>{});
  useEffect(()=>{ load(); }, []);

  const buildPayload = (n, allc, sp, sallc, ss) => ({
    name: n.trim(),
    products: allc==="all" ? "all" : sp,
    sites: sallc==="all" ? "all" : ss,
  });

  const add = () => {
    if (!name.trim()) { showToast("Vendor name required"); return; }
    fetch(`${API_BASE}/vendors`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(buildPayload(name,alloc,selProds,siteAlloc,selSites)) })
      .then(r=>r.json()).then(d=>{ setVendors(v=>[...v,d]); setName(""); setAlloc("all"); setSelProds([]); setSiteAlloc("all"); setSelSites([]); showToast("Vendor added"); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const startEdit = (v) => {
    setEditId(v.id);
    const isAllProds = v.products==="all"||!Array.isArray(v.products);
    const isAllSites = v.sites==="all"||!Array.isArray(v.sites);
    setEditData({
      name: v.name,
      alloc: isAllProds?"all":"specific",
      selProds: isAllProds?[]:[...v.products],
      siteAlloc: isAllSites?"all":"specific",
      selSites: isAllSites?[]:[...v.sites],
    });
  };
  const cancelEdit = () => { setEditId(null); setEditData({}); };

  const saveEdit = () => {
    if (!editData.name?.trim()) { showToast("Vendor name required"); return; }
    fetch(`${API_BASE}/vendors/${editId}`, { method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(buildPayload(editData.name,editData.alloc,editData.selProds,editData.siteAlloc,editData.selSites)) })
      .then(r=>r.json()).then(d=>{ setVendors(v=>v.map(x=>x.id===editId?d:x)); cancelEdit(); showToast("Vendor updated"); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const del = (id) => {
    if (!window.confirm("Delete this vendor?")) return;
    fetch(`${API_BASE}/vendors/${id}`, { method:"DELETE" }).then(()=>setVendors(v=>v.filter(x=>x.id!==id))).catch(()=>{});
  };

  const VendorForm = ({ data, setData, isEdit }) => (
    <div>
      <div style={{ marginBottom:10 }}>
        <div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>Vendor Name</div>
        <input style={inputStyle} value={data.name||""} onChange={e=>setData(p=>({...p,name:e.target.value}))} placeholder="Vendor name"/>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:10 }}>
        <div>
          <div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>Product Allocation</div>
          <select style={{ ...inputStyle, marginBottom:8 }} value={data.alloc||"all"} onChange={e=>setData(p=>({...p,alloc:e.target.value}))}>
            <option value="all">All Products (default)</option>
            <option value="specific">Specific Products</option>
          </select>
          {(data.alloc||"all")==="specific" && (
            <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
              {products.map(p=>(
                <button key={p.code} onClick={()=>setData(prev=>({ ...prev, selProds: prev.selProds?.includes(p.code)?prev.selProds.filter(x=>x!==p.code):[...(prev.selProds||[]),p.code] }))}
                  style={{ ...btn((data.selProds||[]).includes(p.code)), padding:"5px 10px", fontSize:12 }}>{p.name}</button>
              ))}
            </div>
          )}
        </div>
        <div>
          <div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>Site Allocation</div>
          <select style={{ ...inputStyle, marginBottom:8 }} value={data.siteAlloc||"all"} onChange={e=>setData(p=>({...p,siteAlloc:e.target.value}))}>
            <option value="all">All Sites (default)</option>
            <option value="specific">Specific Sites</option>
          </select>
          {(data.siteAlloc||"all")==="specific" && (
            <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
              {sites.map(s=>(
                <button key={s.id} onClick={()=>setData(prev=>({ ...prev, selSites: prev.selSites?.includes(s.id)?prev.selSites.filter(x=>x!==s.id):[...(prev.selSites||[]),s.id] }))}
                  style={{ ...btn((data.selSites||[]).includes(s.id)), padding:"5px 10px", fontSize:12 }}>{s.name}</button>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );

  return (
    <div>
      {editId && (
        <div className="pd-card" style={{ ...card, border:`2px solid ${C.blue}`, background:C.blueBg }}>
          <SectionHeader title="EDIT VENDOR" action={<button style={{ ...btn(false), fontSize:12, padding:"6px 12px" }} onClick={cancelEdit}>Cancel</button>}/>
          <VendorForm data={editData} setData={setEditData} isEdit={true}/>
          <div style={{ display:"flex", justifyContent:"flex-end", gap:10 }}>
            <button style={btn(false)} onClick={cancelEdit}>Cancel</button>
            <button style={btn(true)} onClick={saveEdit}>Save Changes</button>
          </div>
        </div>
      )}
      {!editId && (
        <div className="pd-card" style={card}>
          <SectionHeader title="ADD VENDOR" />
          <VendorForm data={{ name, alloc, selProds, siteAlloc, selSites }} setData={d=>{ if(d.name!==undefined)setName(d.name||""); if(d.alloc!==undefined)setAlloc(d.alloc); if(d.selProds!==undefined)setSelProds(d.selProds); if(d.siteAlloc!==undefined)setSiteAlloc(d.siteAlloc); if(d.selSites!==undefined)setSelSites(d.selSites); }} isEdit={false}/>
          <div style={{ display:"flex", justifyContent:"flex-end" }}><button style={btn(true)} onClick={add}>+ Add Vendor</button></div>
        </div>
      )}
      <div className="pd-card" style={card}>
        <SectionHeader title="VENDORS" />
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse" }}>
            <thead><tr>{["Vendor Name","Products","Sites","Actions"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
            <tbody>
              {vendors.map((v,i)=>(
                <tr key={v.id} style={trAlt(i)}>
                  <td style={{ ...td, fontWeight:600 }}>{v.name}</td>
                  <td style={td}>{v.products==="all"?<Tag label="All Products"/>:(Array.isArray(v.products)?`${v.products.length} products`:"All Products")}</td>
                  <td style={td}>{v.sites==="all"?<Tag label="All Sites"/>:(Array.isArray(v.sites)?v.sites.map(id=>sites.find(s=>s.id===id)?.name||id).join(", "):"All Sites")}</td>
                  <td style={td}>
                    <div style={{ display:"flex", gap:6 }}>
                      <button style={{ ...btn(false), padding:"5px 10px", fontSize:12, color:C.blue, borderColor:C.blue+"44" }} onClick={()=>startEdit(v)}>Edit</button>
                      <button style={{ border:`1px solid ${C.redBg}`, background:"#fff", color:C.red, borderRadius:8, padding:"5px 8px", cursor:"pointer", fontSize:12 }} onClick={()=>del(v.id)}>Delete</button>
                    </div>
                  </td>
                </tr>
              ))}
              {vendors.length===0 && <tr><td colSpan={4} style={{ ...td, color:C.sub }}>No vendors yet.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ── Admin > Status Definition (site-specific) ─────────────────────────────
function AdminStatusDef({ sites, showToast }) {
  const EMPTY_DEFS = { critical:{from:0,to:2}, optimal:{from:2,to:7}, surplus:{from:7,to:9999} };
  const [selSiteId, setSelSiteId] = useState("global");
  const [local, setLocal] = useState(EMPTY_DEFS);
  const [isCustom, setIsCustom] = useState(false);

  const loadDefs = (sid) => {
    const url = sid==="global" ? `${API_BASE}/status-defs` : `${API_BASE}/status-defs?siteId=${sid}`;
    fetch(url).then(r=>r.json()).then(d=>{
      setIsCustom(!d.isGlobal && sid!=="global");
      setLocal({ critical:d.critical, optimal:d.optimal||d.healthy||{from:2,to:7}, surplus:d.surplus||d.excessive||{from:7,to:9999} });
    }).catch(()=>{});
  };

  useEffect(()=>{ loadDefs(selSiteId); }, [selSiteId]);

  const save = () => {
    const payload = { ...local, ...(selSiteId!=="global" ? { siteId:Number(selSiteId) } : {}) };
    fetch(`${API_BASE}/status-defs`, { method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload) })
      .then(()=>{ loadDefs(selSiteId); showToast(`Status definitions saved for ${selSiteId==="global"?"Global Default":sites.find(s=>s.id===Number(selSiteId))?.name||""}`); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const resetToGlobal = () => {
    fetch(`${API_BASE}/status-defs/${selSiteId}`, { method:"DELETE" })
      .then(()=>{ loadDefs(selSiteId); showToast("Reset to global defaults"); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const STATUS_DEF_ROWS = [
    { key:"critical", label:"Critical", color:C.red,   bg:C.redBg,   icon:"🔴" },
    { key:"optimal",  label:"Optimal",  color:C.green, bg:C.greenBg, icon:"🟢" },
    { key:"surplus",  label:"Surplus",  color:C.blue,  bg:C.blueBg,  icon:"🔵" },
  ];

  return (
    <div>
      <div className="pd-card" style={card}>
        <SectionHeader title="STATUS DEFINITION" sub="Define day ranges per site. If a site has no custom definition, the Global Default applies." />
        <div style={{ display:"flex", gap:10, alignItems:"center", marginBottom:16, flexWrap:"wrap" }}>
          <div style={{ fontSize:12, color:C.sub }}>Viewing for:</div>
          <select style={{ ...smInput, width:220 }} value={selSiteId} onChange={e=>setSelSiteId(e.target.value)}>
            <option value="global">🌐 Global Default</option>
            {sites.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
          {selSiteId!=="global" && isCustom && (
            <span style={{ fontSize:12, color:C.blue, fontWeight:600 }}>✓ Has custom definition</span>
          )}
          {selSiteId!=="global" && !isCustom && (
            <span style={{ fontSize:12, color:C.sub }}>Using global default</span>
          )}
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:12, marginBottom:18 }}>
          {STATUS_DEF_ROWS.map(row=>(
            <div key={row.key} style={{ display:"flex", alignItems:"center", gap:12, flexWrap:"wrap", padding:"14px 16px", background:row.bg, borderRadius:12, border:`1px solid ${row.color}30` }}>
              <div style={{ fontSize:18 }}>{row.icon}</div>
              <div style={{ fontWeight:700, color:row.color, minWidth:80, fontSize:14 }}>{row.label}</div>
              <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
                <div style={{ fontSize:12, color:row.color }}>From (days)</div>
                <input type="number" min={0} style={{ ...smInput, width:80, borderColor:row.color }} value={local[row.key]?.from??0}
                  onChange={e=>setLocal(p=>({...p,[row.key]:{...p[row.key],from:Number(e.target.value)}}))}/>
                <div style={{ fontSize:12, color:row.color }}>To (days)</div>
                <input type="number" min={0} style={{ ...smInput, width:80, borderColor:row.color }}
                  value={local[row.key]?.to===9999?"∞":local[row.key]?.to??9999}
                  onChange={e=>{ const v=e.target.value; setLocal(p=>({...p,[row.key]:{...p[row.key],to:v===""||v==="∞"?9999:Number(v)}})); }}
                  placeholder="∞"/>
              </div>
              <div style={{ fontSize:12, color:row.color, marginLeft:"auto" }}>
                {local[row.key]?.from}–{local[row.key]?.to===9999?"∞":local[row.key]?.to} days = <b>{row.label}</b>
              </div>
            </div>
          ))}
        </div>
        <div style={{ display:"flex", justifyContent:"flex-end", gap:10 }}>
          {selSiteId!=="global" && isCustom && (
            <button style={{ ...btn(false), fontSize:12, color:C.red, borderColor:C.red+"44" }} onClick={resetToGlobal}>Reset to Global Default</button>
          )}
          <button style={btn(true)} onClick={save}>
            Save for {selSiteId==="global"?"Global Default":sites.find(s=>s.id===Number(selSiteId))?.name||"Site"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Admin > User Access ────────────────────────────────────────────────────
const ROLES = ["admin","client","clientApprover","vendiman"];
const ROLE_LABELS = { admin:"Admin", client:"Client", clientApprover:"Client Approver", vendiman:"Vendiman" };
const ROLE_DESC = {
  admin: "Full access to everything",
  client: "Dashboard, Stock Inward, Stock Update, Product Performance, Forecast, Client Approvals",
  clientApprover: "Same as Client + can approve orders",
  vendiman: "Dashboard, Stock Inward, Stock Update, Product Performance, Forecast, Vendiman Dashboard, Orders",
};

function AdminUserAccess({ users, setUsers, sites, showToast }) {
  const EMPTY_FORM = { loginId:"", name:"", password:"", role:"client", siteAlloc:"all", selSites:[] };
  const [form, setForm] = useState(EMPTY_FORM);
  const [errors, setErrors] = useState({});
  const [editId, setEditId] = useState(null);
  const [editData, setEditData] = useState({});

  const load = () => fetch(`${API_BASE}/users`).then(r=>r.json()).then(setUsers).catch(()=>{});
  useEffect(()=>{ load(); }, []);

  const validate = (data) => {
    const e = {};
    if (!data.loginId?.trim()) e.loginId="Required";
    if (!data.name?.trim()) e.name="Required";
    if (!editId && !data.password?.trim()) e.password="Required";
    setErrors(e);
    return Object.keys(e).length===0;
  };

  const buildSites = (alloc, sel) => alloc==="all" ? "all" : sel;

  const add = () => {
    if (!validate(form)) return;
    fetch(`${API_BASE}/users`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ ...form, sites:buildSites(form.siteAlloc, form.selSites) }) })
      .then(r=>r.json()).then(d=>{ if(d.error){ showToast(d.error); return; } setUsers(u=>[...u,d]); setForm(EMPTY_FORM); setErrors({}); showToast("User added"); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const startEdit = (u) => {
    const isAllSites = u.sites==="all"||!Array.isArray(u.sites);
    setEditId(u.id);
    setEditData({ loginId:u.loginId, name:u.name, password:"", role:u.role, siteAlloc:isAllSites?"all":"specific", selSites:isAllSites?[]:[...u.sites] });
    setErrors({});
  };
  const cancelEdit = () => { setEditId(null); setEditData({}); setErrors({}); };

  const saveEdit = () => {
    if (!validate(editData)) return;
    const payload = { loginId:editData.loginId, name:editData.name, role:editData.role, sites:buildSites(editData.siteAlloc,editData.selSites) };
    if (editData.password?.trim()) payload.password = editData.password;
    fetch(`${API_BASE}/users/${editId}`, { method:"PUT", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload) })
      .then(r=>r.json()).then(d=>{ if(d.error){ showToast(d.error); return; } setUsers(u=>u.map(x=>x.id===editId?d:x)); cancelEdit(); showToast("User updated"); })
      .catch(()=>showToast("Could not reach backend"));
  };

  const del = (u) => {
    if (u.loginId==="admin") { showToast("Cannot delete the Admin account"); return; }
    if (!window.confirm("Delete this user?")) return;
    fetch(`${API_BASE}/users/${u.id}`, { method:"DELETE" }).then(()=>setUsers(p=>p.filter(x=>x.id!==u.id))).catch(()=>{});
  };

  const UserForm = ({ data, setData, isEdit }) => (
    <div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))", gap:10, marginBottom:12 }}>
        {[["loginId","Login ID"],["name","Name"],["password",isEdit?"New Password (leave blank to keep)":"Password"]].map(([f,l])=>(
          <div key={f}><div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>{l}</div>
            <input style={{ ...smInput, borderColor:errors[f]?C.red:C.border }} type={f==="password"?"password":"text"} value={data[f]||""} onChange={e=>setData(p=>({...p,[f]:e.target.value}))} placeholder={l}/>
            {errors[f]&&<div style={{ fontSize:10, color:C.red, marginTop:2 }}>{errors[f]}</div>}
          </div>
        ))}
        <div><div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>Role</div>
          <select style={smInput} value={data.role||"client"} onChange={e=>setData(p=>({...p,role:e.target.value}))}>
            {ROLES.map(r=><option key={r} value={r}>{ROLE_LABELS[r]}</option>)}
          </select>
        </div>
      </div>
      <div style={{ marginBottom:10 }}>
        <div style={{ fontSize:11, color:C.sub, marginBottom:4 }}>Site Access</div>
        <select style={{ ...smInput, width:"auto", marginBottom:8 }} value={data.siteAlloc||"all"} onChange={e=>setData(p=>({...p,siteAlloc:e.target.value}))}>
          <option value="all">All Sites</option>
          <option value="specific">Specific Sites</option>
        </select>
        {(data.siteAlloc||"all")==="specific" && (
          <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
            {sites.map(s=>(
              <button key={s.id} onClick={()=>setData(prev=>({ ...prev, selSites:prev.selSites?.includes(s.id)?prev.selSites.filter(x=>x!==s.id):[...(prev.selSites||[]),s.id] }))}
                style={{ ...btn((data.selSites||[]).includes(s.id)), padding:"5px 12px", fontSize:12 }}>{s.name}</button>
            ))}
          </div>
        )}
      </div>
      <div style={{ fontSize:12, color:C.sub, padding:"10px 12px", background:C.panel, borderRadius:8 }}>
        <b>{ROLE_LABELS[data.role||"client"]}:</b> {ROLE_DESC[data.role||"client"]}
      </div>
    </div>
  );

  return (
    <div>
      {editId && (
        <div className="pd-card" style={{ ...card, border:`2px solid ${C.blue}`, background:C.blueBg }}>
          <SectionHeader title="EDIT USER" action={<button style={{ ...btn(false), fontSize:12, padding:"6px 12px" }} onClick={cancelEdit}>Cancel</button>}/>
          <UserForm data={editData} setData={setEditData} isEdit={true}/>
          <div style={{ display:"flex", justifyContent:"flex-end", gap:10, marginTop:14 }}>
            <button style={btn(false)} onClick={cancelEdit}>Cancel</button>
            <button style={btn(true)} onClick={saveEdit}>Save Changes</button>
          </div>
        </div>
      )}
      {!editId && (
        <div className="pd-card" style={card}>
          <SectionHeader title="ADD USER" />
          <UserForm data={form} setData={setForm} isEdit={false}/>
          <div style={{ display:"flex", justifyContent:"flex-end", marginTop:14 }}><button style={btn(true)} onClick={add}>+ Add User</button></div>
        </div>
      )}
      <div className="pd-card" style={card}>
        <SectionHeader title="USERS" />
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse" }}>
            <thead><tr>{["Login ID","Name","Role","Site Access","Actions"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
            <tbody>
              {users.map((u,i)=>(
                <tr key={u.id} style={trAlt(i)}>
                  <td style={{ ...td, fontWeight:600 }}>{u.loginId}{u.loginId==="admin"&&<span style={{ fontSize:10, marginLeft:6, color:C.green, fontWeight:700 }}>ADMIN</span>}</td>
                  <td style={td}>{u.name||"—"}</td>
                  <td style={td}><Tag label={ROLE_LABELS[u.role]||u.role}/></td>
                  <td style={td}>{u.sites==="all"?<Tag label="All Sites"/>:(Array.isArray(u.sites)?u.sites.map(id=>sites.find(s=>s.id===id)?.name||id).join(", "):"All Sites")}</td>
                  <td style={td}>
                    <div style={{ display:"flex", gap:6 }}>
                      <button style={{ ...btn(false), padding:"5px 10px", fontSize:12, color:C.blue, borderColor:C.blue+"44" }} onClick={()=>startEdit(u)}>Edit</button>
                      {u.loginId!=="admin" && <button style={{ border:`1px solid ${C.redBg}`, background:"#fff", color:C.red, borderRadius:8, padding:"5px 8px", cursor:"pointer", fontSize:12 }} onClick={()=>del(u)}>Delete</button>}
                    </div>
                  </td>
                </tr>
              ))}
              {users.length===0 && <tr><td colSpan={5} style={{ ...td, color:C.sub }}>No users yet.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ─── Main App ──────────────────────────────────────────────────────────────
export default function PantryDashboard() {
  const [currentUser, setCurrentUser] = useState(null);
  const [tab, setTab] = useState("Dashboard");
  const [toast, setToast] = useState("");
  const [menuOpen, setMenuOpen] = useState(false);

  // Global data
  const [sites, setSites] = useState([]);
  const [products, setProducts] = useState([]);
  const [vendors, setVendors] = useState([]);
  const [statusDefs, setStatusDefs] = useState({ critical:{from:0,to:2}, healthy:{from:2,to:7}, excessive:{from:7,to:9999} });
  const [users, setUsers] = useState([]);

  const showToast = (msg) => { setToast(msg); setTimeout(()=>setToast(""),3000); };

  useEffect(()=>{
    if (!currentUser) return;
    Promise.all([
      fetch(`${API_BASE}/sites`).then(r=>r.json()),
      fetch(`${API_BASE}/products`).then(r=>r.json()),
      fetch(`${API_BASE}/vendors`).then(r=>r.json()),
      fetch(`${API_BASE}/status-defs`).then(r=>r.json()),
      fetch(`${API_BASE}/users`).then(r=>r.json()),
    ]).then(([s,p,v,sd,u])=>{
      setSites(s); setProducts(p); setVendors(v); setStatusDefs(sd); setUsers(u);
    }).catch(()=>showToast("Could not load data from backend"));
  }, [currentUser]);

  if (!currentUser) return <LoginScreen onLogin={(user)=>{ setCurrentUser(user); setTab("Dashboard"); }}/>;

  const tabs = ROLE_TABS[currentUser.role] || ROLE_TABS.client;
  if (!tabs.includes(tab)) setTab(tabs[0]);

  const tabIcon = { Dashboard:"📊", "Stock Inward":"📦", "Stock Update":"🔄", "Product Performance":"📈", Forecast:"🔮", "Vendiman Dashboard":"🏭", "Client Approvals":"✅", Orders:"🛒", Reports:"📋", Admin:"⚙️" };

  return (
    <div style={{ minHeight:"100vh", background:C.bg, fontFamily:"'Segoe UI', system-ui, sans-serif", color:C.ink }}>
      {/* Header */}
      <div style={{ background:C.gradient, padding:"14px 16px", position:"sticky", top:0, zIndex:50, boxShadow:"0 4px 18px rgba(67,56,202,0.25)" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", maxWidth:1200, margin:"0 auto" }}>
          <div>
            <div style={{ fontSize:16, fontWeight:700, color:"#fff" }}>🏢 Meta India Pantry</div>
            <div style={{ fontSize:11, color:"rgba(255,255,255,0.75)" }}>Hi {currentUser.name} · {ROLE_LABELS[currentUser.role]}</div>
          </div>
          <div style={{ display:"flex", gap:10, alignItems:"center" }}>
            <span className="pd-live" style={{ background:"rgba(255,255,255,0.18)", color:"#fff", fontSize:11, fontWeight:600, padding:"4px 10px", borderRadius:999, border:"1px solid rgba(255,255,255,0.35)", display:"none" }}>📶 Live</span>
            {/* Hamburger for mobile */}
            <button onClick={()=>setMenuOpen(m=>!m)} style={{ background:"rgba(255,255,255,0.2)", border:"none", color:"#fff", borderRadius:10, padding:"8px 10px", cursor:"pointer", fontSize:16 }}>☰</button>
            <button onClick={()=>setCurrentUser(null)} style={{ background:"rgba(255,255,255,0.15)", border:"1px solid rgba(255,255,255,0.3)", color:"#fff", borderRadius:10, padding:"7px 12px", cursor:"pointer", fontSize:12, fontWeight:600 }}>Sign Out</button>
          </div>
        </div>
      </div>

      {/* Mobile menu overlay */}
      {menuOpen && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex:49 }} onClick={()=>setMenuOpen(false)}>
          <div style={{ position:"fixed", top:0, left:0, bottom:0, width:260, background:"#fff", padding:"70px 16px 24px", overflowY:"auto" }} onClick={e=>e.stopPropagation()}>
            {tabs.map(t=>(
              <button key={t} onClick={()=>{ setTab(t); setMenuOpen(false); }} style={{ display:"flex", alignItems:"center", gap:10, width:"100%", padding:"12px 14px", marginBottom:4, background:tab===t?C.blueBg:"transparent", color:tab===t?C.blue:C.ink, border:tab===t?`1px solid ${C.blue}30`:"none", borderRadius:10, fontSize:14, fontWeight:tab===t?700:400, cursor:"pointer", textAlign:"left" }}>
                <span>{tabIcon[t]}</span>{t}
              </button>
            ))}
          </div>
        </div>
      )}

      {/* Desktop tab nav */}
      <div style={{ background:C.bar, borderBottom:`1px solid ${C.border}`, padding:"10px 16px", display:"none" }} className="pd-desktopnav">
        <div style={{ display:"flex", gap:6, flexWrap:"wrap", maxWidth:1200, margin:"0 auto" }}>
          {tabs.map(t=>(
            <button key={t} className={`pd-tab ${tab===t?"pd-tab-active":""}`} onClick={()=>setTab(t)} style={{ ...btn(tab===t), padding:"8px 14px", fontSize:13 }}>
              {tabIcon[t]} {t}
            </button>
          ))}
        </div>
      </div>

      {/* Bottom tab bar for mobile */}
      <div style={{ position:"fixed", bottom:0, left:0, right:0, background:"#fff", borderTop:`1px solid ${C.border}`, zIndex:40, display:"flex", overflowX:"auto", padding:"4px 0" }}>
        {tabs.map(t=>(
          <button key={t} onClick={()=>setTab(t)} style={{ flex:"0 0 auto", display:"flex", flexDirection:"column", alignItems:"center", gap:2, padding:"6px 10px", background:"transparent", border:"none", cursor:"pointer", color:tab===t?C.blue:C.sub, fontSize:18, minWidth:60 }}>
            <span>{tabIcon[t]}</span>
            <span style={{ fontSize:9, fontWeight:tab===t?700:400, whiteSpace:"nowrap" }}>{t.length>10?t.slice(0,9)+"…":t}</span>
          </button>
        ))}
      </div>

      {/* Page content */}
      <div key={tab} className="pd-content" style={{ maxWidth:1200, margin:"0 auto", padding:"16px 12px 90px" }}>
        {tab==="Dashboard"          && <Dashboard sites={sites} statusDefs={statusDefs}/>}
        {tab==="Stock Inward"       && <StockInward sites={sites} vendors={vendors} products={products} showToast={showToast} currentUser={currentUser}/>}
        {tab==="Stock Update"       && <StockUpdate sites={sites} showToast={showToast} currentUser={currentUser}/>}
        {tab==="Product Performance"&& <ProductPerformance sites={sites} showToast={showToast}/>}
        {tab==="Forecast"           && <Forecast sites={sites} showToast={showToast} currentUser={currentUser}/>}
        {tab==="Vendiman Dashboard" && <VendimanDashboard sites={sites} showToast={showToast}/>}
        {tab==="Client Approvals"   && <ClientApprovals showToast={showToast} currentUser={currentUser} sites={sites}/>}
        {tab==="Reports"            && <Reports sites={sites} showToast={showToast}/>}
        {tab==="Orders"             && <Orders sites={sites} showToast={showToast}/>}
        {tab==="Admin"              && <Admin sites={sites} setSites={setSites} vendors={vendors} setVendors={setVendors} products={products} setProducts={setProducts} statusDefs={statusDefs} setStatusDefs={setStatusDefs} users={users} setUsers={setUsers} showToast={showToast}/>}
      </div>

      {/* Toast */}
      {toast && (
        <div className="pd-toast" style={{ position:"fixed", bottom:72, left:"50%", transform:"translateX(-50%)", background:C.gradient, color:"#fff", padding:"10px 20px", borderRadius:999, fontSize:13, fontWeight:600, boxShadow:"0 8px 24px rgba(67,56,202,0.35)", zIndex:60, maxWidth:"90vw", textAlign:"center" }}>
          {toast}
        </div>
      )}
    </div>
  );
}
