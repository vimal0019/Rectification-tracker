import { useState, useRef, useCallback } from "react";

// ── helpers ──────────────────────────────────────────────────────────────────
const uid = () => `TKT-${Date.now().toString(36).toUpperCase()}`;
const now = () => new Date().toLocaleString("en-IN", { hour12: true });
const STATUS = ["Open", "In Progress", "Resolved", "Closed"];
const STATUS_COLOR = {
  Open: { bg: "#FEF3C7", text: "#92400E", dot: "#F59E0B" },
  "In Progress": { bg: "#DBEAFE", text: "#1E40AF", dot: "#3B82F6" },
  Resolved: { bg: "#D1FAE5", text: "#065F46", dot: "#10B981" },
  Closed: { bg: "#F3F4F6", text: "#374151", dot: "#6B7280" },
};
const PRIORITY = ["Low", "Medium", "High", "Critical"];
const PRIORITY_COLOR = { Low: "#6B7280", Medium: "#F59E0B", High: "#EF4444", Critical: "#7C3AED" };
const SHOP_CLASSES = ["A+", "A", "B", "C", "D", "E"];
const SHOP_CLASS_COLOR = {
  "A+": { bg: "#EDE9FE", text: "#5B21B6", border: "#7C3AED" },
  "A":  { bg: "#DBEAFE", text: "#1E40AF", border: "#3B82F6" },
  "B":  { bg: "#D1FAE5", text: "#065F46", border: "#10B981" },
  "C":  { bg: "#FEF3C7", text: "#92400E", border: "#F59E0B" },
  "D":  { bg: "#FEE2E2", text: "#991B1B", border: "#EF4444" },
  "E":  { bg: "#F3F4F6", text: "#374151", border: "#6B7280" },
};

const EMPTY_FORM = {
  title: "", category: "Mechanical", priority: "Medium",
  shopClass: "A", shopName: "",
  tlName: "", tlContact: "",
  vbaName: "", vbaContact: "",
  asmName: "",
  description: "", rootCause: "", resolution: "", notes: "",
};

// ── PPT Generator ─────────────────────────────────────────────────────────────
async function generatePPTScript(tickets) {
  const summary = tickets.map((t, i) => ({
    num: i + 1, id: t.id, title: t.title, category: t.category,
    priority: t.priority, status: t.status,
    shopClass: t.shopClass, shopName: t.shopName,
    tlName: t.tlName, tlContact: t.tlContact,
    vbaName: t.vbaName, vbaContact: t.vbaContact,
    asmName: t.asmName,
    description: t.description,
    rootCause: t.rootCause || "Under investigation",
    resolution: t.resolution || "Pending",
    createdAt: t.createdAt, resolvedAt: t.resolvedAt || "—",
    hasImages: !!(t.beforeImage || t.afterImage),
  }));

  const prompt = `You are a PowerPoint slide writer. Generate a JSON array of slide objects for a Problem Rectification Report presentation.

Tickets data:
${JSON.stringify(summary, null, 2)}

Return ONLY a valid JSON array (no markdown, no extra text) with this structure:
[
  { "slideType": "title", "title": "...", "subtitle": "..." },
  { "slideType": "summary", "title": "...", "stats": [{"label":"...","value":"..."}] },
  {
    "slideType": "ticket",
    "ticketId": "...", "title": "...", "priority": "...", "status": "...",
    "shopClass": "...", "shopName": "...",
    "tlName": "...", "tlContact": "...",
    "vbaName": "...", "vbaContact": "...", "asmName": "...",
    "description": "...", "rootCause": "...", "resolution": "...",
    "dates": "Created: ... | Resolved: ..."
  }
]

Create: 1 title slide, 1 summary/stats slide (include class breakdown), then one slide per ticket. Use professional language.`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514", max_tokens: 1500,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  const data = await response.json();
  const raw = data.content.map((b) => b.text || "").join("");
  const clean = raw.replace(/```json|```/g, "").trim();
  return JSON.parse(clean);
}

// ── PPT Builder ───────────────────────────────────────────────────────────────
async function buildPPTX(slides, tickets) {
  if (!window.PptxGenJS) {
    await new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundled.js";
      s.onload = res; s.onerror = rej;
      document.head.appendChild(s);
    });
  }
  const pptx = new window.PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  pptx.title = "Problem Rectification Report";

  const NAVY="1E3A5F", TEAL="0D9488", LIGHT="F0F9FF", WHITE="FFFFFF";
  const GRAY="64748B", RED="EF4444", GREEN="10B981", AMBER="F59E0B";

  for (const slide of slides) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };

    if (slide.slideType === "title") {
      s.addShape(pptx.shapes.RECTANGLE, { x:0,y:0,w:10,h:3.2, fill:{color:NAVY} });
      s.addShape(pptx.shapes.RECTANGLE, { x:0,y:3.2,w:10,h:0.15, fill:{color:TEAL} });
      s.addText("⚙", { x:0.4,y:0.35,w:1,h:1, fontSize:48, color:TEAL });
      s.addText(slide.title||"Problem Rectification Report", { x:1.5,y:0.5,w:8,h:1.3, fontSize:34,bold:true,color:WHITE,fontFace:"Cambria" });
      s.addText(slide.subtitle||"Ticket-Based Tracking & Resolution Report", { x:1.5,y:1.9,w:8,h:0.7, fontSize:16,color:"A5C8E1",fontFace:"Calibri" });
      s.addText(`Generated: ${now()}`, { x:1.5,y:2.65,w:8,h:0.4, fontSize:11,color:"7FB3D3",fontFace:"Calibri" });
      s.addText("CONFIDENTIAL — INTERNAL USE ONLY", { x:0,y:5.3,w:10,h:0.3, fontSize:9,color:GRAY,align:"center" });

    } else if (slide.slideType === "summary") {
      s.addShape(pptx.shapes.RECTANGLE, { x:0,y:0,w:10,h:1.1, fill:{color:NAVY} });
      s.addText(slide.title||"Executive Summary", { x:0.5,y:0.2,w:9,h:0.7, fontSize:28,bold:true,color:WHITE,fontFace:"Cambria" });
      const stats = (slide.stats||[]).slice(0,4);
      const cw = 8.5/Math.max(stats.length,1);
      stats.forEach((st,i) => {
        const cx = 0.75 + i*(cw+0.1);
        s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x:cx,y:1.4,w:cw-0.05,h:2, fill:{color:LIGHT}, rectRadius:0.08,
          shadow:{type:"outer",color:"000000",blur:6,offset:2,angle:45,opacity:0.1} });
        s.addText(st.value, { x:cx,y:1.6,w:cw-0.05,h:0.9, fontSize:36,bold:true,color:TEAL,align:"center",fontFace:"Cambria" });
        s.addText(st.label, { x:cx,y:2.5,w:cw-0.05,h:0.5, fontSize:12,color:GRAY,align:"center",fontFace:"Calibri" });
      });

    } else if (slide.slideType === "ticket") {
      const prColor = slide.priority==="Critical"?"7C3AED":slide.priority==="High"?RED:slide.priority==="Medium"?AMBER:GRAY;
      const stColor = (slide.status==="Resolved"||slide.status==="Closed")?GREEN:slide.status==="In Progress"?"3B82F6":AMBER;

      s.addShape(pptx.shapes.RECTANGLE, { x:0,y:0,w:10,h:1.05, fill:{color:NAVY} });
      s.addText(slide.ticketId, { x:0.3,y:0.04,w:2.5,h:0.38, fontSize:10,color:TEAL,bold:true,fontFace:"Calibri" });
      s.addText(slide.title, { x:0.3,y:0.42,w:6.8,h:0.55, fontSize:17,bold:true,color:WHITE,fontFace:"Cambria" });

      // Priority + Status badges
      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x:7.3,y:0.1,w:1.1,h:0.33, fill:{color:prColor}, rectRadius:0.05 });
      s.addText(slide.priority, { x:7.3,y:0.1,w:1.1,h:0.33, fontSize:9,color:WHITE,align:"center",bold:true });
      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x:8.55,y:0.1,w:1.2,h:0.33, fill:{color:stColor}, rectRadius:0.05 });
      s.addText(slide.status, { x:8.55,y:0.1,w:1.2,h:0.33, fontSize:9,color:WHITE,align:"center",bold:true });

      // Shop Class badge
      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x:7.3,y:0.55,w:0.5,h:0.38, fill:{color:"F59E0B"}, rectRadius:0.05 });
      s.addText(`Class\n${slide.shopClass||"—"}`, { x:7.3,y:0.55,w:0.5,h:0.38, fontSize:7,color:WHITE,align:"center",bold:true });
      s.addText(`Shop: ${slide.shopName||"—"}`, { x:7.85,y:0.6,w:1.9,h:0.28, fontSize:9,color:"A5C8E1",fontFace:"Calibri" });

      // Contact info row
      s.addShape(pptx.shapes.RECTANGLE, { x:0,y:1.08,w:10,h:0.55, fill:{color:"EFF6FF"} });
      const contacts = [
        `TL: ${slide.tlName||"—"}  📞 ${slide.tlContact||"—"}`,
        `VBA: ${slide.vbaName||"—"}  📞 ${slide.vbaContact||"—"}`,
        `ASM: ${slide.asmName||"—"}`,
      ];
      contacts.forEach((c,i)=>{
        s.addText(c, { x:0.3+i*3.25,y:1.1,w:3.1,h:0.38, fontSize:9,color:"1E3A5F",fontFace:"Calibri",bold:i===2 });
      });

      // Content boxes 2x2
      const boxes = [
        { label:"Problem Description", value:slide.description, x:0.3, y:1.75 },
        { label:"Root Cause Analysis",  value:slide.rootCause,   x:5.2, y:1.75 },
        { label:"Resolution / Corrective Action", value:slide.resolution, x:0.3, y:3.5 },
        { label:"Timeline", value:slide.dates, x:5.2, y:3.5 },
      ];
      boxes.forEach((box)=>{
        s.addText(box.label, { x:box.x,y:box.y,w:4.5,h:0.28, fontSize:9,bold:true,color:TEAL,fontFace:"Calibri" });
        s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x:box.x,y:box.y+0.28,w:4.5,h:1.55, fill:{color:LIGHT}, rectRadius:0.05 });
        s.addText(box.value||"—", { x:box.x+0.1,y:box.y+0.35,w:4.3,h:1.38, fontSize:10,color:"1E293B",fontFace:"Calibri",valign:"top" });
      });
    }
  }

  // Before/After slides
  for (const ticket of tickets) {
    if (!ticket.beforeImage && !ticket.afterImage) continue;
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    s.addShape(pptx.shapes.RECTANGLE, { x:0,y:0,w:10,h:0.9, fill:{color:NAVY} });
    s.addText(`${ticket.id} — Before & After`, { x:0.3,y:0.15,w:7,h:0.6, fontSize:20,bold:true,color:WHITE,fontFace:"Cambria" });
    // Shop class tag on this slide too
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x:7.6,y:0.15,w:0.6,h:0.5, fill:{color:"F59E0B"}, rectRadius:0.06 });
    s.addText(`Class\n${ticket.shopClass||"?"}`, { x:7.6,y:0.15,w:0.6,h:0.5, fontSize:8,color:WHITE,align:"center",bold:true });
    s.addText(ticket.shopName||"", { x:8.3,y:0.25,w:1.5,h:0.35, fontSize:9,color:"A5C8E1" });

    if (ticket.beforeImage) {
      s.addText("BEFORE", { x:0.5,y:1.0,w:4,h:0.35, fontSize:12,bold:true,color:RED,align:"center" });
      s.addShape(pptx.shapes.RECTANGLE, { x:0.5,y:1.4,w:4,h:3.5, fill:{color:"F9FAFB"}, line:{color:"E5E7EB",width:1} });
      s.addImage({ data:ticket.beforeImage, x:0.55,y:1.45,w:3.9,h:3.4, sizing:{type:"contain",w:3.9,h:3.4} });
    }
    if (ticket.afterImage) {
      s.addText("AFTER", { x:5.5,y:1.0,w:4,h:0.35, fontSize:12,bold:true,color:GREEN,align:"center" });
      s.addShape(pptx.shapes.RECTANGLE, { x:5.5,y:1.4,w:4,h:3.5, fill:{color:"F9FAFB"}, line:{color:"E5E7EB",width:1} });
      s.addImage({ data:ticket.afterImage, x:5.55,y:1.45,w:3.9,h:3.4, sizing:{type:"contain",w:3.9,h:3.4} });
    }
    if (ticket.beforeImage && ticket.afterImage) {
      s.addText("→", { x:4.5,y:2.7,w:1,h:0.8, fontSize:36,color:TEAL,align:"center" });
    }
  }

  return pptx;
}

// ── Reusable field components ─────────────────────────────────────────────────
const Label = ({ children }) => (
  <label style={{ fontSize:11, fontWeight:700, color:"#0D9488", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:0.5 }}>
    {children}
  </label>
);
const Field = ({ label, children }) => (
  <div><Label>{label}</Label>{children}</div>
);
const inputStyle = { width:"100%", padding:"8px 10px", borderRadius:7, border:"1.5px solid #E2E8F0", fontSize:13, boxSizing:"border-box", background:"#F8FAFC" };
const selectStyle = { ...inputStyle, cursor:"pointer" };

// ── ImageUpload ───────────────────────────────────────────────────────────────
function ImageUpload({ label, value, onChange, color }) {
  const ref = useRef();
  const handleDrop = useCallback((e) => { e.preventDefault(); const f=e.dataTransfer.files[0]; if(f) read(f); }, []);
  const read = (file) => { const r=new FileReader(); r.onload=(e)=>onChange(e.target.result); r.readAsDataURL(file); };
  return (
    <div>
      <div style={{ fontWeight:600, fontSize:12, color, marginBottom:6 }}>{label}</div>
      <div onDragOver={(e)=>e.preventDefault()} onDrop={handleDrop} onClick={()=>ref.current.click()}
        style={{ border:`2px dashed ${color}55`, borderRadius:8, padding:10, cursor:"pointer", textAlign:"center",
          background:color+"08", minHeight:130, display:"flex", alignItems:"center", justifyContent:"center" }}>
        {value ? (
          <div style={{ position:"relative" }}>
            <img src={value} alt={label} style={{ maxWidth:"100%", maxHeight:160, borderRadius:6 }} />
            <button onClick={(e)=>{e.stopPropagation();onChange(null);}}
              style={{ position:"absolute",top:4,right:4,background:"#EF4444",color:"#fff",border:"none",borderRadius:4,padding:"2px 6px",cursor:"pointer",fontSize:11 }}>✕</button>
          </div>
        ) : (
          <div style={{ color:color+"99" }}>
            <div style={{ fontSize:28, marginBottom:4 }}>📷</div>
            <div style={{ fontSize:11 }}>Click or drag & drop</div>
          </div>
        )}
      </div>
      <input ref={ref} type="file" accept="image/*" style={{ display:"none" }} onChange={(e)=>e.target.files[0]&&read(e.target.files[0])} />
    </div>
  );
}

// ── ShopClassBadge ────────────────────────────────────────────────────────────
function ShopClassBadge({ cls }) {
  const c = SHOP_CLASS_COLOR[cls] || SHOP_CLASS_COLOR["E"];
  return (
    <span style={{ background:c.bg, color:c.text, border:`1.5px solid ${c.border}`, borderRadius:6,
      padding:"2px 10px", fontSize:12, fontWeight:800, letterSpacing:1 }}>
      Class {cls}
    </span>
  );
}

// ── TicketCard ────────────────────────────────────────────────────────────────
function TicketCard({ ticket, onSelect, selected }) {
  const sc = STATUS_COLOR[ticket.status] || STATUS_COLOR.Open;
  return (
    <div onClick={()=>onSelect(ticket)}
      style={{ border:selected?"2px solid #0D9488":"1.5px solid #E2E8F0", borderRadius:10, padding:"12px 14px",
        cursor:"pointer", background:selected?"#F0FDFA":"#FFF", marginBottom:10, transition:"all .15s",
        boxShadow:selected?"0 0 0 3px #CCFBF1":"0 1px 4px rgba(0,0,0,.06)" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:4 }}>
        <span style={{ fontWeight:700, fontSize:11, color:"#0D9488", letterSpacing:1 }}>{ticket.id}</span>
        <div style={{ display:"flex", gap:5 }}>
          <ShopClassBadge cls={ticket.shopClass||"E"} />
          <span style={{ background:sc.bg,color:sc.text,borderRadius:20,padding:"2px 9px",fontSize:10,fontWeight:700,display:"flex",alignItems:"center",gap:3 }}>
            <span style={{ width:5,height:5,borderRadius:"50%",background:sc.dot,display:"inline-block" }} />{ticket.status}
          </span>
        </div>
      </div>
      <div style={{ fontWeight:600, fontSize:13, color:"#1E293B", marginBottom:4 }}>{ticket.title}</div>
      <div style={{ fontSize:11, color:"#64748B", display:"flex", flexWrap:"wrap", gap:"6px 14px" }}>
        {ticket.shopName && <span>🏪 {ticket.shopName}</span>}
        {ticket.tlName && <span>👤 TL: {ticket.tlName}</span>}
        <span style={{ background:PRIORITY_COLOR[ticket.priority]+"22",color:PRIORITY_COLOR[ticket.priority],borderRadius:20,padding:"1px 8px",fontSize:10,fontWeight:700 }}>{ticket.priority}</span>
        {(ticket.beforeImage||ticket.afterImage) && <span style={{ color:"#0D9488" }}>📷</span>}
      </div>
    </div>
  );
}

// ── ShopInfoPanel ─────────────────────────────────────────────────────────────
function ShopInfoPanel({ ticket, updateTicket, readOnly }) {
  const edit = (f,v) => updateTicket && updateTicket(f,v);
  return (
    <div style={{ background:"#FFF",borderRadius:10,border:"1.5px solid #E2E8F0",padding:18,marginBottom:18 }}>
      <h3 style={{ fontSize:14,fontWeight:700,color:"#1E3A5F",marginTop:0,marginBottom:14,display:"flex",alignItems:"center",gap:8 }}>
        🏪 Shop & Contact Information
        <ShopClassBadge cls={ticket.shopClass||"E"} />
      </h3>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12, marginBottom:12 }}>
        <Field label="Shop Class">
          {readOnly ? <ShopClassBadge cls={ticket.shopClass||"E"} /> :
            <select value={ticket.shopClass||"A"} onChange={(e)=>edit("shopClass",e.target.value)} style={selectStyle}>
              {SHOP_CLASSES.map(c=><option key={c}>{c}</option>)}
            </select>}
        </Field>
        <Field label="Shop Name">
          {readOnly ? <div style={{ fontSize:13,color:"#1E293B",padding:"6px 0" }}>{ticket.shopName||"—"}</div> :
            <input value={ticket.shopName||""} onChange={(e)=>edit("shopName",e.target.value)} style={inputStyle} placeholder="Enter shop name" />}
        </Field>
        <Field label="ASM Name">
          {readOnly ? <div style={{ fontSize:13,color:"#1E293B",padding:"6px 0" }}>{ticket.asmName||"—"}</div> :
            <input value={ticket.asmName||""} onChange={(e)=>edit("asmName",e.target.value)} style={inputStyle} placeholder="ASM name" />}
        </Field>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr 1fr", gap:12 }}>
        <Field label="TL Name">
          {readOnly ? <div style={{ fontSize:13,color:"#1E293B",padding:"6px 0" }}>{ticket.tlName||"—"}</div> :
            <input value={ticket.tlName||""} onChange={(e)=>edit("tlName",e.target.value)} style={inputStyle} placeholder="Team Lead name" />}
        </Field>
        <Field label="TL Contact">
          {readOnly ? <div style={{ fontSize:13,color:"#1E293B",padding:"6px 0" }}>{ticket.tlContact||"—"}</div> :
            <input value={ticket.tlContact||""} onChange={(e)=>edit("tlContact",e.target.value)} style={inputStyle} placeholder="Mobile / ext" type="tel" />}
        </Field>
        <Field label="VBA Name">
          {readOnly ? <div style={{ fontSize:13,color:"#1E293B",padding:"6px 0" }}>{ticket.vbaName||"—"}</div> :
            <input value={ticket.vbaName||""} onChange={(e)=>edit("vbaName",e.target.value)} style={inputStyle} placeholder="VBA name" />}
        </Field>
        <Field label="VBA Contact">
          {readOnly ? <div style={{ fontSize:13,color:"#1E293B",padding:"6px 0" }}>{ticket.vbaContact||"—"}</div> :
            <input value={ticket.vbaContact||""} onChange={(e)=>edit("vbaContact",e.target.value)} style={inputStyle} placeholder="Mobile / ext" type="tel" />}
        </Field>
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [tickets, setTickets] = useState([
    {
      id:"TKT-DEMO1", title:"Corroded pipe flange at Unit-3", category:"Mechanical",
      priority:"High", status:"Resolved",
      shopClass:"A", shopName:"Main Workshop – Unit 3",
      tlName:"Ramesh Kumar", tlContact:"98400 11122",
      vbaName:"Priya Devi", vbaContact:"98400 33344",
      asmName:"Suresh Anand",
      description:"Severe corrosion on 4-inch pipe flange causing minor seepage at Unit-3 cooling line.",
      rootCause:"Moisture ingress due to failed insulation jacket.",
      resolution:"Replaced flange and re-applied insulation. Torqued to spec.",
      notes:"",
      createdAt:"09/06/2026, 09:00 AM", resolvedAt:"09/06/2026, 02:30 PM",
      beforeImage:null, afterImage:null,
    },
  ]);
  const [selected, setSelected] = useState(tickets[0]);
  const [view, setView] = useState("detail");
  const [form, setForm] = useState(EMPTY_FORM);
  const [filter, setFilter] = useState("All");
  const [classFilter, setClassFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [pptState, setPptState] = useState("idle");
  const [pptMsg, setPptMsg] = useState("");

  const updateTicket = (field, value) => {
    setTickets(prev => prev.map(t => t.id===selected.id ? {...t,[field]:value} : t));
    setSelected(prev => ({...prev,[field]:value}));
  };

  const createTicket = () => {
    const t = { ...form, id:uid(), status:"Open", createdAt:now(), resolvedAt:null, beforeImage:null, afterImage:null };
    setTickets(prev=>[t,...prev]);
    setSelected(t);
    setForm(EMPTY_FORM);
    setView("detail");
  };

  const handleGeneratePPT = async () => {
    if (!tickets.length) return;
    setPptState("generating"); setPptMsg("Claude is writing slide content…");
    try {
      const slides = await generatePPTScript(tickets);
      setPptMsg("Building PPTX file…");
      const pptx = await buildPPTX(slides, tickets);
      await pptx.writeFile({ fileName:"Rectification_Report.pptx" });
      setPptState("done"); setPptMsg("✅ Rectification_Report.pptx downloaded!");
    } catch(err) {
      setPptState("error"); setPptMsg("❌ Error: "+err.message);
    }
    setTimeout(()=>{ setPptState("idle"); setPptMsg(""); }, 7000);
  };

  const filtered = tickets.filter(t => {
    if (filter!=="All" && t.status!==filter) return false;
    if (classFilter!=="All" && t.shopClass!==classFilter) return false;
    if (search && !t.title.toLowerCase().includes(search.toLowerCase()) &&
        !t.id.includes(search.toUpperCase()) && !(t.shopName||"").toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  });

  const stats = {
    total: tickets.length,
    open: tickets.filter(t=>t.status==="Open").length,
    inProgress: tickets.filter(t=>t.status==="In Progress").length,
    resolved: tickets.filter(t=>t.status==="Resolved"||t.status==="Closed").length,
  };
  const classCounts = SHOP_CLASSES.reduce((acc,c)=>({ ...acc,[c]:tickets.filter(t=>t.shopClass===c).length }),{});

  return (
    <div style={{ fontFamily:"'Segoe UI',system-ui,sans-serif", background:"#F8FAFC", minHeight:"100vh", display:"flex", flexDirection:"column" }}>
      {/* Header */}
      <div style={{ background:"#1E3A5F", padding:"12px 20px", display:"flex", alignItems:"center", justifyContent:"space-between", boxShadow:"0 2px 8px rgba(0,0,0,.2)" }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <span style={{ fontSize:24 }}>⚙️</span>
          <div>
            <div style={{ color:"#FFF", fontWeight:700, fontSize:16 }}>Problem Rectification Tracker</div>
            <div style={{ color:"#7FB3D3", fontSize:10 }}>Shop-Class Ticket Management · TL · VBA · ASM · Auto PPT</div>
          </div>
        </div>
        <div style={{ display:"flex", gap:8 }}>
          <button onClick={()=>setView("new")}
            style={{ background:"#0D9488",color:"#fff",border:"none",borderRadius:7,padding:"7px 16px",fontWeight:700,fontSize:12,cursor:"pointer" }}>
            + New Ticket
          </button>
          <button onClick={handleGeneratePPT} disabled={pptState==="generating"||!tickets.length}
            style={{ background:pptState==="generating"?"#9CA3AF":"#F59E0B",color:"#fff",border:"none",borderRadius:7,padding:"7px 16px",fontWeight:700,fontSize:12,cursor:pptState==="generating"?"default":"pointer" }}>
            {pptState==="generating"?"⏳ Generating…":"📊 Export PPT"}
          </button>
        </div>
      </div>

      {pptMsg && (
        <div style={{ background:pptState==="done"?"#D1FAE5":pptState==="error"?"#FEE2E2":"#DBEAFE",
          color:pptState==="done"?"#065F46":pptState==="error"?"#991B1B":"#1E40AF",
          padding:"8px 20px", fontSize:12, fontWeight:600, borderBottom:"1px solid #E2E8F0" }}>
          {pptMsg}
        </div>
      )}

      {/* Stats Bar */}
      <div style={{ background:"#FFF", borderBottom:"1px solid #E2E8F0", padding:"10px 20px", display:"flex", gap:0, flexWrap:"wrap", alignItems:"center" }}>
        {[
          { label:"Total", value:stats.total, color:"#1E3A5F" },
          { label:"Open", value:stats.open, color:"#F59E0B" },
          { label:"In Progress", value:stats.inProgress, color:"#3B82F6" },
          { label:"Resolved", value:stats.resolved, color:"#10B981" },
        ].map(s=>(
          <div key={s.label} style={{ textAlign:"center", minWidth:72, paddingRight:16, borderRight:"1px solid #F1F5F9", marginRight:16 }}>
            <div style={{ fontSize:22, fontWeight:800, color:s.color }}>{s.value}</div>
            <div style={{ fontSize:10, color:"#64748B", fontWeight:600 }}>{s.label}</div>
          </div>
        ))}
        {/* Class mini-badges */}
        <div style={{ display:"flex", gap:6, flexWrap:"wrap", flex:1 }}>
          {SHOP_CLASSES.map(c=>{
            const cc = SHOP_CLASS_COLOR[c];
            return (
              <div key={c} style={{ background:cc.bg, border:`1px solid ${cc.border}`, borderRadius:6, padding:"2px 10px", textAlign:"center", cursor:"pointer" }}
                onClick={()=>setClassFilter(classFilter===c?"All":c)}>
                <div style={{ fontSize:13, fontWeight:800, color:cc.text }}>Class {c}</div>
                <div style={{ fontSize:10, color:cc.text, opacity:0.8 }}>{classCounts[c]||0} tickets</div>
                {classFilter===c && <div style={{ fontSize:8, color:cc.border, fontWeight:700 }}>● FILTERED</div>}
              </div>
            );
          })}
        </div>
        <div style={{ display:"flex", gap:5, flexWrap:"wrap" }}>
          {["All","Open","In Progress","Resolved","Closed"].map(f=>(
            <button key={f} onClick={()=>setFilter(f)}
              style={{ background:filter===f?"#1E3A5F":"#F1F5F9", color:filter===f?"#FFF":"#475569",
                border:"none", borderRadius:20, padding:"4px 12px", fontSize:11, fontWeight:600, cursor:"pointer" }}>{f}</button>
          ))}
        </div>
      </div>

      {/* Body */}
      <div style={{ display:"flex", flex:1, overflow:"hidden" }}>
        {/* List */}
        <div style={{ width:320, background:"#F8FAFC", borderRight:"1px solid #E2E8F0", padding:12, overflowY:"auto", flexShrink:0 }}>
          <input placeholder="🔍 Search by title, ID, shop…" value={search} onChange={(e)=>setSearch(e.target.value)}
            style={{ width:"100%", padding:"7px 11px", borderRadius:8, border:"1.5px solid #E2E8F0", fontSize:12, marginBottom:10, boxSizing:"border-box" }} />
          {filtered.length===0 && <div style={{ textAlign:"center", color:"#94A3B8", padding:30, fontSize:13 }}>No tickets found</div>}
          {filtered.map(t=>(
            <TicketCard key={t.id} ticket={t} selected={selected?.id===t.id}
              onSelect={(tk)=>{ setSelected(tk); setView("detail"); }} />
          ))}
        </div>

        {/* Main Panel */}
        <div style={{ flex:1, padding:20, overflowY:"auto" }}>

          {/* ── New Ticket Form ── */}
          {view==="new" && (
            <div style={{ maxWidth:780 }}>
              <h2 style={{ fontSize:19, fontWeight:700, color:"#1E3A5F", marginTop:0, marginBottom:16 }}>🎫 Create New Ticket</h2>

              {/* Shop Info */}
              <ShopInfoPanel ticket={form} updateTicket={(f,v)=>setForm(p=>({...p,[f]:v}))} />

              {/* Basic fields */}
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:14, marginBottom:14 }}>
                <div style={{ gridColumn:"1/-1" }}>
                  <Label>Problem Title *</Label>
                  <input value={form.title} onChange={(e)=>setForm(p=>({...p,title:e.target.value}))} style={inputStyle} placeholder="Brief description of the problem" />
                </div>
                <div>
                  <Label>Category</Label>
                  <select value={form.category} onChange={(e)=>setForm(p=>({...p,category:e.target.value}))} style={selectStyle}>
                    {["Mechanical","Electrical","Civil","Software","Safety","Quality","Other"].map(o=><option key={o}>{o}</option>)}
                  </select>
                </div>
                <div>
                  <Label>Priority</Label>
                  <select value={form.priority} onChange={(e)=>setForm(p=>({...p,priority:e.target.value}))} style={selectStyle}>
                    {PRIORITY.map(o=><option key={o}>{o}</option>)}
                  </select>
                </div>
              </div>

              {[{ label:"Problem Description", field:"description" },{ label:"Root Cause", field:"rootCause" },{ label:"Resolution / Action Taken", field:"resolution" }].map(f=>(
                <div key={f.field} style={{ marginBottom:12 }}>
                  <Label>{f.label}</Label>
                  <textarea value={form[f.field]} onChange={(e)=>setForm(p=>({...p,[f.field]:e.target.value}))} rows={3}
                    style={{ ...inputStyle, resize:"vertical" }} />
                </div>
              ))}

              <div style={{ display:"flex", gap:10 }}>
                <button onClick={createTicket} disabled={!form.title}
                  style={{ background:form.title?"#0D9488":"#9CA3AF",color:"#fff",border:"none",borderRadius:7,padding:"9px 22px",fontWeight:700,fontSize:13,cursor:form.title?"pointer":"default" }}>
                  Create Ticket
                </button>
                <button onClick={()=>setView(selected?"detail":"list")}
                  style={{ background:"#F1F5F9",color:"#374151",border:"none",borderRadius:7,padding:"9px 22px",fontWeight:600,fontSize:13,cursor:"pointer" }}>
                  Cancel
                </button>
              </div>
            </div>
          )}

          {/* ── Ticket Detail ── */}
          {view==="detail" && selected && (
            <div style={{ maxWidth:860 }}>
              {/* Title row */}
              <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", marginBottom:16 }}>
                <div>
                  <div style={{ fontSize:11, fontWeight:700, color:"#0D9488", letterSpacing:1, marginBottom:2 }}>{selected.id}</div>
                  <h2 style={{ fontSize:20, fontWeight:700, color:"#1E3A5F", margin:"0 0 8px" }}>{selected.title}</h2>
                  <div style={{ display:"flex", gap:7, flexWrap:"wrap" }}>
                    <ShopClassBadge cls={selected.shopClass||"E"} />
                    <span style={{ background:STATUS_COLOR[selected.status]?.bg,color:STATUS_COLOR[selected.status]?.text,borderRadius:20,padding:"3px 11px",fontSize:11,fontWeight:700 }}>{selected.status}</span>
                    <span style={{ background:PRIORITY_COLOR[selected.priority]+"20",color:PRIORITY_COLOR[selected.priority],borderRadius:20,padding:"3px 11px",fontSize:11,fontWeight:700 }}>{selected.priority}</span>
                    <span style={{ background:"#F1F5F9",color:"#475569",borderRadius:20,padding:"3px 11px",fontSize:11,fontWeight:600 }}>{selected.category}</span>
                  </div>
                </div>
                <div style={{ display:"flex", gap:8 }}>
                  <select value={selected.status} onChange={(e)=>{
                    const v=e.target.value; updateTicket("status",v);
                    if(v==="Resolved"||v==="Closed") updateTicket("resolvedAt",now());
                  }} style={{ padding:"7px 10px",borderRadius:7,border:"1.5px solid #E2E8F0",fontSize:12,fontWeight:600,cursor:"pointer" }}>
                    {STATUS.map(s=><option key={s}>{s}</option>)}
                  </select>
                  <button onClick={()=>{ setTickets(p=>p.filter(t=>t.id!==selected.id)); setSelected(null); setView("list"); }}
                    style={{ background:"#FEE2E2",color:"#991B1B",border:"none",borderRadius:7,padding:"7px 13px",fontWeight:600,fontSize:12,cursor:"pointer" }}>🗑</button>
                </div>
              </div>

              {/* Shop Info (editable) */}
              <ShopInfoPanel ticket={selected} updateTicket={updateTicket} />

              {/* Problem fields */}
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:16 }}>
                {[
                  { label:"Problem Description", field:"description" },
                  { label:"Root Cause Analysis", field:"rootCause" },
                  { label:"Resolution / Corrective Action", field:"resolution" },
                  { label:"Notes", field:"notes" },
                ].map(f=>(
                  <div key={f.field}>
                    <Label>{f.label}</Label>
                    <textarea value={selected[f.field]||""} onChange={(e)=>updateTicket(f.field,e.target.value)} rows={3}
                      style={{ ...inputStyle, resize:"vertical" }} />
                  </div>
                ))}
              </div>

              {/* Before / After */}
              <div style={{ background:"#FFF", borderRadius:10, border:"1.5px solid #E2E8F0", padding:18, marginBottom:16 }}>
                <h3 style={{ fontSize:14, fontWeight:700, color:"#1E3A5F", marginTop:0, marginBottom:14 }}>📷 Before & After Rectification Photos</h3>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20 }}>
                  <ImageUpload label="📸 BEFORE — Problem State" value={selected.beforeImage} onChange={(v)=>updateTicket("beforeImage",v)} color="#EF4444" />
                  <ImageUpload label="✅ AFTER — Rectified State" value={selected.afterImage} onChange={(v)=>updateTicket("afterImage",v)} color="#10B981" />
                </div>
                {(selected.beforeImage||selected.afterImage) && (
                  <div style={{ marginTop:10, padding:"7px 12px", background:"#F0FDFA", borderRadius:7, fontSize:12, color:"#065F46", fontWeight:600 }}>
                    ✅ These images will appear as a dedicated Before/After slide in the PPT export.
                  </div>
                )}
              </div>

              {/* Timeline */}
              <div style={{ background:"#F8FAFC", borderRadius:8, padding:"10px 14px", fontSize:12, color:"#64748B", display:"flex", gap:20 }}>
                <span>📅 Created: <strong>{selected.createdAt}</strong></span>
                {selected.resolvedAt && <span>✅ Resolved: <strong>{selected.resolvedAt}</strong></span>}
              </div>
            </div>
          )}

          {/* Empty state */}
          {view==="list" && (
            <div style={{ textAlign:"center", paddingTop:80, color:"#94A3B8" }}>
              <div style={{ fontSize:60, marginBottom:14 }}>🎫</div>
              <div style={{ fontSize:17, fontWeight:700, color:"#475569" }}>Select a ticket to view details</div>
              <div style={{ fontSize:12, marginTop:6 }}>or create a new one using the button above</div>
            </div>
          )}
        </div>
      </div>

      {/* Footer */}
      <div style={{ background:"#1E3A5F", color:"#7FB3D3", fontSize:10, padding:"7px 20px", display:"flex", justifyContent:"space-between" }}>
        <span>Problem Rectification Tracker · Shop Class A+/A/B/C/D/E · TL · VBA · ASM</span>
        <span>📊 Export PPT → Claude AI auto-generates slides with class info + Before/After images</span>
      </div>
    </div>
  );
}
