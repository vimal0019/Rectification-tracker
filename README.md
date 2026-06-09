<!DOCTYPE html>
<html>
<head>
  <title>Rectification Tracker</title>
</head>
<body>
  <h1>Rectification Tracker</h1>
  <p>Website is working.</p>
</body>
</html>
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
const PRIORITY_COLOR = {
  Low: "#6B7280",
  Medium: "#F59E0B",
  High: "#EF4444",
  Critical: "#7C3AED",
};

// ── PPT Generator (server-side via Anthropic API) ─────────────────────────────
async function generatePPTScript(tickets) {
  const summary = tickets.map((t, i) => ({
    num: i + 1,
    id: t.id,
    title: t.title,
    category: t.category,
    priority: t.priority,
    status: t.status,
    description: t.description,
    rootCause: t.rootCause || "Under investigation",
    resolution: t.resolution || "Pending",
    createdAt: t.createdAt,
    resolvedAt: t.resolvedAt || "—",
    hasImages: !!(t.beforeImage || t.afterImage),
  }));

  const prompt = `You are a PowerPoint slide writer. Generate a JSON array of slide objects for a Problem Rectification Report presentation.

Tickets data:
${JSON.stringify(summary, null, 2)}

Return ONLY a valid JSON array (no markdown, no extra text) with this structure:
[
  {
    "slideType": "title",
    "title": "...",
    "subtitle": "..."
  },
  {
    "slideType": "summary",
    "title": "...",
    "stats": [{"label":"...","value":"..."}]
  },
  {
    "slideType": "ticket",
    "ticketId": "...",
    "title": "...",
    "priority": "...",
    "status": "...",
    "description": "...",
    "rootCause": "...",
    "resolution": "...",
    "dates": "Created: ... | Resolved: ..."
  }
]

Create: 1 title slide, 1 summary/stats slide, then one slide per ticket. Use professional language.`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  const data = await response.json();
  const raw = data.content.map((b) => b.text || "").join("");
  const clean = raw.replace(/```json|```/g, "").trim();
  return JSON.parse(clean);
}

// ── PPT Builder (client-side canvas → PPTX via pptxgenjs CDN) ────────────────
async function buildPPTX(slides, tickets) {
  // Dynamically load pptxgenjs
  if (!window.PptxGenJS) {
    await new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundled.js";
      s.onload = res;
      s.onerror = rej;
      document.head.appendChild(s);
    });
  }

  const pptx = new window.PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  pptx.title = "Problem Rectification Report";

  const NAVY = "1E3A5F";
  const TEAL = "0D9488";
  const LIGHT = "F0F9FF";
  const WHITE = "FFFFFF";
  const GRAY = "64748B";
  const RED = "EF4444";
  const GREEN = "10B981";
  const AMBER = "F59E0B";

  for (const slide of slides) {
    const s = pptx.addSlide();
    s.background = { color: WHITE };

    if (slide.slideType === "title") {
      // Full-bleed header band
      s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 3.2, fill: { color: NAVY } });
      s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 3.2, w: 10, h: 0.15, fill: { color: TEAL } });
      s.addText("⚙", { x: 0.4, y: 0.35, w: 1, h: 1, fontSize: 48, color: TEAL });
      s.addText(slide.title || "Problem Rectification Report", {
        x: 1.5, y: 0.5, w: 8, h: 1.3, fontSize: 34, bold: true, color: WHITE, fontFace: "Cambria",
      });
      s.addText(slide.subtitle || "Ticket-Based Tracking & Resolution Report", {
        x: 1.5, y: 1.9, w: 8, h: 0.7, fontSize: 16, color: "A5C8E1", fontFace: "Calibri",
      });
      s.addText(`Generated: ${now()}`, {
        x: 1.5, y: 2.65, w: 8, h: 0.4, fontSize: 11, color: "7FB3D3", fontFace: "Calibri",
      });
      // bottom info
      s.addText("CONFIDENTIAL — INTERNAL USE ONLY", {
        x: 0, y: 5.3, w: 10, h: 0.3, fontSize: 9, color: GRAY, align: "center",
      });
    } else if (slide.slideType === "summary") {
      s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.1, fill: { color: NAVY } });
      s.addText(slide.title || "Executive Summary", {
        x: 0.5, y: 0.2, w: 9, h: 0.7, fontSize: 28, bold: true, color: WHITE, fontFace: "Cambria",
      });
      const stats = slide.stats || [];
      const cols = Math.min(stats.length, 4);
      const cw = 8.5 / cols;
      stats.slice(0, 4).forEach((st, i) => {
        const cx = 0.75 + i * (cw + 0.1);
        s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: cx, y: 1.4, w: cw - 0.05, h: 2,
          fill: { color: LIGHT },
          rectRadius: 0.08,
          shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 45, opacity: 0.1 },
        });
        s.addText(st.value, {
          x: cx, y: 1.6, w: cw - 0.05, h: 0.9,
          fontSize: 36, bold: true, color: TEAL, align: "center", fontFace: "Cambria",
        });
        s.addText(st.label, {
          x: cx, y: 2.5, w: cw - 0.05, h: 0.5,
          fontSize: 12, color: GRAY, align: "center", fontFace: "Calibri",
        });
      });
    } else if (slide.slideType === "ticket") {
      const prColor = slide.priority === "Critical" ? "7C3AED"
        : slide.priority === "High" ? RED
        : slide.priority === "Medium" ? AMBER : GRAY;
      const stColor = slide.status === "Resolved" || slide.status === "Closed" ? GREEN
        : slide.status === "In Progress" ? "3B82F6" : AMBER;

      s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: NAVY } });
      s.addText(slide.ticketId, {
        x: 0.3, y: 0.05, w: 2, h: 0.45, fontSize: 11, color: TEAL, bold: true, fontFace: "Calibri",
      });
      s.addText(slide.title, {
        x: 0.3, y: 0.45, w: 7, h: 0.5, fontSize: 18, bold: true, color: WHITE, fontFace: "Cambria",
      });
      // badges
      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 7.5, y: 0.1, w: 1.1, h: 0.35, fill: { color: prColor }, rectRadius: 0.05 });
      s.addText(slide.priority, { x: 7.5, y: 0.1, w: 1.1, h: 0.35, fontSize: 9, color: WHITE, align: "center", bold: true });
      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 8.7, y: 0.1, w: 1.1, h: 0.35, fill: { color: stColor }, rectRadius: 0.05 });
      s.addText(slide.status, { x: 8.7, y: 0.1, w: 1.1, h: 0.35, fontSize: 9, color: WHITE, align: "center", bold: true });

      // Content boxes
      const boxes = [
        { label: "Problem Description", value: slide.description, x: 0.3, y: 1.15 },
        { label: "Root Cause Analysis", value: slide.rootCause, x: 5.2, y: 1.15 },
        { label: "Resolution / Corrective Action", value: slide.resolution, x: 0.3, y: 3.1 },
        { label: "Timeline", value: slide.dates, x: 5.2, y: 3.1 },
      ];
      boxes.forEach((box) => {
        s.addText(box.label, {
          x: box.x, y: box.y, w: 4.5, h: 0.3,
          fontSize: 9, bold: true, color: TEAL, fontFace: "Calibri",
        });
        s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: box.x, y: box.y + 0.3, w: 4.5, h: 1.6,
          fill: { color: LIGHT }, rectRadius: 0.05,
        });
        s.addText(box.value || "—", {
          x: box.x + 0.1, y: box.y + 0.4, w: 4.3, h: 1.4,
          fontSize: 11, color: "1E293B", fontFace: "Calibri", valign: "top",
        });
      });
    }
  }

  // Before/After slides for tickets that have images
  for (const ticket of tickets) {
    if (!ticket.beforeImage && !ticket.afterImage) continue;
    const s = pptx.addSlide();
    s.background = { color: WHITE };
    s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: NAVY } });
    s.addText(`${ticket.id} — Before & After Rectification`, {
      x: 0.3, y: 0.15, w: 9, h: 0.6, fontSize: 20, bold: true, color: WHITE, fontFace: "Cambria",
    });
    if (ticket.beforeImage) {
      s.addText("BEFORE", { x: 0.5, y: 1.0, w: 4, h: 0.35, fontSize: 12, bold: true, color: RED, align: "center" });
      s.addShape(pptx.shapes.RECTANGLE, { x: 0.5, y: 1.4, w: 4, h: 3.5, fill: { color: "F9FAFB" }, line: { color: "E5E7EB", width: 1 } });
      s.addImage({ data: ticket.beforeImage, x: 0.55, y: 1.45, w: 3.9, h: 3.4, sizing: { type: "contain", w: 3.9, h: 3.4 } });
    }
    if (ticket.afterImage) {
      s.addText("AFTER", { x: 5.5, y: 1.0, w: 4, h: 0.35, fontSize: 12, bold: true, color: GREEN, align: "center" });
      s.addShape(pptx.shapes.RECTANGLE, { x: 5.5, y: 1.4, w: 4, h: 3.5, fill: { color: "F9FAFB" }, line: { color: "E5E7EB", width: 1 } });
      s.addImage({ data: ticket.afterImage, x: 5.55, y: 1.45, w: 3.9, h: 3.4, sizing: { type: "contain", w: 3.9, h: 3.4 } });
    }
    // Arrow between
    if (ticket.beforeImage && ticket.afterImage) {
      s.addText("→", { x: 4.5, y: 2.7, w: 1, h: 0.8, fontSize: 36, color: TEAL, align: "center" });
    }
  }

  return pptx;
}

// ── Ticket Card ───────────────────────────────────────────────────────────────
function TicketCard({ ticket, onSelect, selected }) {
  const sc = STATUS_COLOR[ticket.status] || STATUS_COLOR.Open;
  return (
    <div
      onClick={() => onSelect(ticket)}
      style={{
        border: selected ? "2px solid #0D9488" : "1.5px solid #E2E8F0",
        borderRadius: 10, padding: "14px 16px", cursor: "pointer",
        background: selected ? "#F0FDFA" : "#FFFFFF",
        marginBottom: 10, transition: "all .15s",
        boxShadow: selected ? "0 0 0 3px #CCFBF1" : "0 1px 4px rgba(0,0,0,.06)",
      }}
    >
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <span style={{ fontWeight: 700, fontSize: 11, color: "#0D9488", letterSpacing: 1 }}>{ticket.id}</span>
        <div style={{ display: "flex", gap: 6 }}>
          <span style={{
            background: sc.bg, color: sc.text, borderRadius: 20, padding: "2px 10px", fontSize: 10, fontWeight: 700,
            display: "flex", alignItems: "center", gap: 4,
          }}>
            <span style={{ width: 6, height: 6, borderRadius: "50%", background: sc.dot, display: "inline-block" }} />
            {ticket.status}
          </span>
          <span style={{ background: PRIORITY_COLOR[ticket.priority] + "22", color: PRIORITY_COLOR[ticket.priority], borderRadius: 20, padding: "2px 10px", fontSize: 10, fontWeight: 700 }}>
            {ticket.priority}
          </span>
        </div>
      </div>
      <div style={{ fontWeight: 600, fontSize: 14, color: "#1E293B", marginTop: 6 }}>{ticket.title}</div>
      <div style={{ fontSize: 12, color: "#64748B", marginTop: 4, display: "flex", gap: 10 }}>
        <span>📁 {ticket.category}</span>
        <span>🕐 {ticket.createdAt}</span>
        {(ticket.beforeImage || ticket.afterImage) && <span style={{ color: "#0D9488" }}>📷 Images</span>}
      </div>
    </div>
  );
}

// ── Image Upload Box ──────────────────────────────────────────────────────────
function ImageUpload({ label, value, onChange, color }) {
  const ref = useRef();
  const handleDrop = useCallback((e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) readFile(file);
  }, []);
  const readFile = (file) => {
    const r = new FileReader();
    r.onload = (e) => onChange(e.target.result);
    r.readAsDataURL(file);
  };
  return (
    <div>
      <div style={{ fontWeight: 600, fontSize: 12, color, marginBottom: 6 }}>{label}</div>
      <div
        onDragOver={(e) => e.preventDefault()} onDrop={handleDrop}
        onClick={() => ref.current.click()}
        style={{
          border: `2px dashed ${color}55`, borderRadius: 8, padding: 10,
          cursor: "pointer", textAlign: "center", background: color + "08",
          minHeight: 120, display: "flex", alignItems: "center", justifyContent: "center",
        }}
      >
        {value ? (
          <div style={{ position: "relative" }}>
            <img src={value} alt={label} style={{ maxWidth: "100%", maxHeight: 160, borderRadius: 6 }} />
            <button
              onClick={(e) => { e.stopPropagation(); onChange(null); }}
              style={{ position: "absolute", top: 4, right: 4, background: "#EF4444", color: "#fff", border: "none", borderRadius: 4, padding: "2px 6px", cursor: "pointer", fontSize: 11 }}
            >✕</button>
          </div>
        ) : (
          <div style={{ color: color + "99" }}>
            <div style={{ fontSize: 28, marginBottom: 4 }}>📷</div>
            <div style={{ fontSize: 11 }}>Click or drag & drop</div>
          </div>
        )}
      </div>
      <input ref={ref} type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => e.target.files[0] && readFile(e.target.files[0])} />
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [tickets, setTickets] = useState([
    {
      id: "TKT-DEMO1", title: "Corroded pipe flange at Unit-3", category: "Mechanical",
      priority: "High", status: "Resolved", description: "Severe corrosion on 4-inch pipe flange causing minor seepage at Unit-3 cooling line.",
      rootCause: "Moisture ingress due to failed insulation jacket.", resolution: "Replaced flange and re-applied insulation. Torqued to spec.",
      createdAt: "09/06/2026, 09:00 AM", resolvedAt: "09/06/2026, 02:30 PM",
      beforeImage: null, afterImage: null,
    },
  ]);
  const [selected, setSelected] = useState(tickets[0]);
  const [view, setView] = useState("list"); // list | detail | new
  const [form, setForm] = useState({ title: "", category: "Mechanical", priority: "Medium", description: "", rootCause: "", resolution: "" });
  const [filter, setFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [pptState, setPptState] = useState("idle"); // idle | generating | done | error
  const [pptMsg, setPptMsg] = useState("");

  const updateTicket = (field, value) => {
    setTickets((prev) => prev.map((t) => t.id === selected.id ? { ...t, [field]: value } : t));
    setSelected((prev) => ({ ...prev, [field]: value }));
  };

  const createTicket = () => {
    const t = {
      ...form,
      id: uid(),
      status: "Open",
      createdAt: now(),
      resolvedAt: null,
      beforeImage: null,
      afterImage: null,
    };
    setTickets((prev) => [t, ...prev]);
    setSelected(t);
    setForm({ title: "", category: "Mechanical", priority: "Medium", description: "", rootCause: "", resolution: "" });
    setView("detail");
  };

  const handleGeneratePPT = async () => {
    if (!tickets.length) return;
    setPptState("generating");
    setPptMsg("Asking Claude to write slide content…");
    try {
      const slides = await generatePPTScript(tickets);
      setPptMsg("Building PPTX file…");
      const pptx = await buildPPTX(slides, tickets);
      await pptx.writeFile({ fileName: "Rectification_Report.pptx" });
      setPptState("done");
      setPptMsg("✅ Rectification_Report.pptx downloaded!");
    } catch (err) {
      setPptState("error");
      setPptMsg("❌ Error: " + err.message);
    }
    setTimeout(() => setPptState("idle"), 6000);
  };

  const filtered = tickets.filter((t) => {
    if (filter !== "All" && t.status !== filter) return false;
    if (search && !t.title.toLowerCase().includes(search.toLowerCase()) && !t.id.includes(search.toUpperCase())) return false;
    return true;
  });

  const stats = {
    total: tickets.length,
    open: tickets.filter((t) => t.status === "Open").length,
    inProgress: tickets.filter((t) => t.status === "In Progress").length,
    resolved: tickets.filter((t) => t.status === "Resolved" || t.status === "Closed").length,
  };

  return (
    <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", background: "#F8FAFC", minHeight: "100vh", display: "flex", flexDirection: "column" }}>
      {/* Header */}
      <div style={{ background: "#1E3A5F", padding: "14px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", boxShadow: "0 2px 8px rgba(0,0,0,.2)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <span style={{ fontSize: 26 }}>⚙️</span>
          <div>
            <div style={{ color: "#FFF", fontWeight: 700, fontSize: 17, letterSpacing: 0.5 }}>Problem Rectification Tracker</div>
            <div style={{ color: "#7FB3D3", fontSize: 11 }}>Ticket-based issue management & PPT reporting</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button
            onClick={() => setView("new")}
            style={{ background: "#0D9488", color: "#fff", border: "none", borderRadius: 7, padding: "8px 18px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}
          >+ New Ticket</button>
          <button
            onClick={handleGeneratePPT}
            disabled={pptState === "generating" || !tickets.length}
            style={{
              background: pptState === "generating" ? "#9CA3AF" : "#F59E0B",
              color: "#fff", border: "none", borderRadius: 7, padding: "8px 18px",
              fontWeight: 700, fontSize: 13, cursor: pptState === "generating" ? "default" : "pointer",
            }}
          >
            {pptState === "generating" ? "⏳ Generating…" : "📊 Export PPT"}
          </button>
        </div>
      </div>

      {pptMsg && (
        <div style={{
          background: pptState === "done" ? "#D1FAE5" : pptState === "error" ? "#FEE2E2" : "#DBEAFE",
          color: pptState === "done" ? "#065F46" : pptState === "error" ? "#991B1B" : "#1E40AF",
          padding: "10px 24px", fontSize: 13, fontWeight: 600, borderBottom: "1px solid #E2E8F0",
        }}>
          {pptMsg}
        </div>
      )}

      {/* Stats Bar */}
      <div style={{ background: "#FFF", borderBottom: "1px solid #E2E8F0", padding: "12px 24px", display: "flex", gap: 24 }}>
        {[
          { label: "Total Tickets", value: stats.total, color: "#1E3A5F" },
          { label: "Open", value: stats.open, color: "#F59E0B" },
          { label: "In Progress", value: stats.inProgress, color: "#3B82F6" },
          { label: "Resolved / Closed", value: stats.resolved, color: "#10B981" },
        ].map((s) => (
          <div key={s.label} style={{ textAlign: "center", minWidth: 80 }}>
            <div style={{ fontSize: 24, fontWeight: 800, color: s.color }}>{s.value}</div>
            <div style={{ fontSize: 10, color: "#64748B", fontWeight: 600 }}>{s.label}</div>
          </div>
        ))}
        <div style={{ flex: 1 }} />
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          {["All", "Open", "In Progress", "Resolved", "Closed"].map((f) => (
            <button key={f} onClick={() => setFilter(f)}
              style={{
                background: filter === f ? "#1E3A5F" : "#F1F5F9",
                color: filter === f ? "#FFF" : "#475569",
                border: "none", borderRadius: 20, padding: "5px 13px",
                fontSize: 11, fontWeight: 600, cursor: "pointer",
              }}>{f}</button>
          ))}
        </div>
      </div>

      {/* Body */}
      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>
        {/* Ticket List */}
        <div style={{ width: 330, background: "#F8FAFC", borderRight: "1px solid #E2E8F0", padding: 14, overflowY: "auto" }}>
          <input
            placeholder="🔍 Search tickets…"
            value={search} onChange={(e) => setSearch(e.target.value)}
            style={{ width: "100%", padding: "8px 12px", borderRadius: 8, b
