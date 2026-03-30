import { useState, useRef } from "react";
import * as XLSX from "xlsx";

const C = {
  bg: "#0d1117",
  surface: "#161b24",
  surfaceHigh: "#1e2535",
  border: "#252d3d",
  accent: "#4f8ef7",
  missing: "#e05252",
  missingDim: "#2a1010",
  extra: "#4ecb8b",
  extraDim: "#0e2a1a",
  match: "#7aa2f7",
  matchDim: "#131d3a",
  warn: "#f0a050",
  ai: "#c084fc",
  aiDim: "#1a0f2e",
  text: "#cdd6f4",
  textSub: "#a0aacc",
  textMuted: "#606880",
  mono: "'JetBrains Mono', monospace",
  sans: "'Inter', system-ui, sans-serif",
};

// ── SheetJS ──────────────────────────────────────────────────────────────────
const readXlsx = (file) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        res(wb.SheetNames.map((name) => ({
          name,
          rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "" }),
        })));
      } catch (err) { rej(err); }
    };
    r.onerror = () => rej(new Error("Leesschout"));
    r.readAsArrayBuffer(file);
  });

const guessCol = (rows) => {
  if (!rows?.length) return 0;
  const headers = rows[0] || [];
  const colCount = Math.max(...rows.map((r) => r.length), 0);
  let best = 0, bestScore = -1;
  for (let c = 0; c < colCount; c++) {
    let score = 0;
    const h = String(headers[c] || "").toLowerCase();
    if (/document|naam|titel|omschrijving|product|deliverable|name|title/.test(h)) score += 10;
    for (let r = 1; r < Math.min(rows.length, 20); r++) {
      const v = String(rows[r][c] || "").trim();
      if (v.length > 3) score += 1;
      if (v.length > 10) score += 1;
    }
    if (score > bestScore) { bestScore = score; best = c; }
  }
  return best;
};

const extractCol = (rows, col, startRow = 1) =>
  rows.slice(startRow).map((r) => String(r[col] || "").trim()).filter((v) => v.length > 1);

// ── Anthropic API ─────────────────────────────────────────────────────────────
const callClaude = async (messages, system) => {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      system,
      messages,
    }),
  });
  const data = await res.json();
  return data.content?.map((b) => b.text || "").join("") || "";
};

const SYSTEM = `Je bent een expert in documentbeheer en informatiebeheer voor bouwprojecten en engineering.
Je taak is om documentnamen semantisch te vergelijken. 

Gegeven een document uit de ILP (Informatie Leveringsplanning) en een eis uit de productenlijst van de opdrachtgever, beoordeel je of ze hetzelfde document beschrijven, ook als de namen anders zijn geformuleerd.

Antwoord ALLEEN met geldig JSON in dit formaat:
{
  "match": true/false,
  "confidence": 0.0-1.0,
  "reason": "korte uitleg in het Nederlands (max 1 zin)"
}`;

// Batch AI matching: voor elke eis kijk je of er een ILP item bij past
const aiMatchAll = async (ilpList, plList, onProgress) => {
  const results = []; // { plItem, matchedIlp, confidence, reason, type }

  for (let pi = 0; pi < plList.length; pi++) {
    const plItem = plList[pi];
    onProgress(`Analyseer eis ${pi + 1}/${plList.length}: "${plItem.substring(0, 40)}…"`);

    let bestMatch = null;
    let bestConf = 0;
    let bestReason = "";

    // First check exact/fuzzy match to save API calls
    const normPl = plItem.toLowerCase().replace(/[\s\-_\/\.]+/g, " ").trim();
    const exactIdx = ilpList.findIndex(
      (i) => i.toLowerCase().replace(/[\s\-_\/\.]+/g, " ").trim() === normPl
    );
    if (exactIdx >= 0) {
      results.push({ plItem, matchedIlp: ilpList[exactIdx], confidence: 1.0, reason: "Exacte overeenkomst", type: "match" });
      continue;
    }

    // AI check against all ILP items
    const prompt = `Productenlijst eis: "${plItem}"

ILP documenten:
${ilpList.map((item, i) => `${i + 1}. ${item}`).join("\n")}

Welk ILP document (geef het nummer) beschrijft hetzelfde als de eis, ook als de naam anders is? 
Als er geen match is, geef dan match: false.
Als er een match is, geef dan het nummer van het beste overeenkomende ILP document.

Antwoord in JSON:
{
  "match": true/false,
  "ilp_index": <nummer 1-based, of null>,
  "confidence": 0.0-1.0,
  "reason": "korte uitleg in het Nederlands"
}`;

    try {
      const raw = await callClaude([{ role: "user", content: prompt }],
        `Je bent een expert documentbeheerder. Vergelijk semantisch of een eis en een ILP document hetzelfde beschrijven. Antwoord altijd in geldig JSON.`
      );
      const clean = raw.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);
      if (parsed.match && parsed.ilp_index) {
        const idx = parsed.ilp_index - 1;
        if (idx >= 0 && idx < ilpList.length) {
          bestMatch = ilpList[idx];
          bestConf = parsed.confidence;
          bestReason = parsed.reason;
        }
      } else {
        bestReason = parsed.reason || "Geen overeenkomend document gevonden";
      }
    } catch (_) {
      bestReason = "Analyse mislukt";
    }

    if (bestMatch) {
      results.push({ plItem, matchedIlp: bestMatch, confidence: bestConf, reason: bestReason, type: bestConf > 0.6 ? "match_ai" : "uncertain" });
    } else {
      results.push({ plItem, matchedIlp: null, confidence: 0, reason: bestReason, type: "missing" });
    }
  }

  // Find extra ILP items (not matched to any PL item)
  const matchedIlpItems = new Set(results.filter((r) => r.matchedIlp).map((r) => r.matchedIlp));
  const extraItems = ilpList.filter((item) => !matchedIlpItems.has(item));

  return { results, extraItems };
};

// ── UI Components ─────────────────────────────────────────────────────────────
const Stat = ({ label, value, color }) => (
  <div style={{ flex: 1, minWidth: 90, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, padding: "12px 16px", textAlign: "center" }}>
    <div style={{ fontFamily: C.mono, fontSize: 24, fontWeight: 700, color }}>{value}</div>
    <div style={{ fontSize: 11, color: C.textMuted, marginTop: 2 }}>{label}</div>
  </div>
);

const MatchRow = ({ item }) => {
  const [open, setOpen] = useState(false);
  const { plItem, matchedIlp, confidence, reason, type } = item;
  const cfg = {
    match:     { color: C.match,   bg: C.matchDim,   icon: "✓", label: "Overeenkomst" },
    match_ai:  { color: C.ai,      bg: C.aiDim,      icon: "≈", label: `AI match (${Math.round(confidence * 100)}%)` },
    uncertain: { color: C.warn,    bg: "#1a1400",     icon: "?", label: `Onzeker (${Math.round(confidence * 100)}%)` },
    missing:   { color: C.missing, bg: C.missingDim,  icon: "✗", label: "Ontbreekt" },
  }[type] || { color: C.textMuted, bg: C.surface, icon: "·", label: "" };

  return (
    <div style={{ marginBottom: 4 }}>
      <div
        onClick={() => type !== "match" && setOpen(!open)}
        style={{
          display: "flex", alignItems: "center", gap: 10,
          padding: "8px 12px", background: cfg.bg,
          borderLeft: `3px solid ${cfg.color}`, borderRadius: 4,
          cursor: type !== "match" ? "pointer" : "default",
        }}
      >
        <span style={{ color: cfg.color, fontFamily: C.mono, fontWeight: 700, width: 14, fontSize: 13 }}>{cfg.icon}</span>
        <span style={{ color: C.text, flex: 1, fontFamily: C.mono, fontSize: 12 }}>{plItem}</span>
        <span style={{ color: cfg.color, fontSize: 10, fontFamily: C.mono, whiteSpace: "nowrap" }}>{cfg.label}</span>
        {type !== "match" && <span style={{ color: C.textMuted, fontSize: 10 }}>{open ? "▲" : "▼"}</span>}
      </div>
      {open && (
        <div style={{
          background: `${cfg.color}0a`, border: `1px solid ${cfg.color}22`,
          borderRadius: "0 0 4px 4px", padding: "8px 14px 10px 38px",
          fontFamily: C.mono, fontSize: 11,
        }}>
          {matchedIlp && (
            <div style={{ marginBottom: 4 }}>
              <span style={{ color: C.textMuted }}>ILP document: </span>
              <span style={{ color: C.text }}>{matchedIlp}</span>
            </div>
          )}
          <div>
            <span style={{ color: C.textMuted }}>Analyse: </span>
            <span style={{ color: C.textSub }}>{reason}</span>
          </div>
        </div>
      )}
    </div>
  );
};

const ExtraRow = ({ name }) => (
  <div style={{
    display: "flex", alignItems: "center", gap: 10,
    padding: "8px 12px", marginBottom: 4,
    background: C.extraDim, borderLeft: `3px solid ${C.extra}`,
    borderRadius: 4, fontFamily: C.mono, fontSize: 12,
  }}>
    <span style={{ color: C.extra, fontWeight: 700, width: 14 }}>+</span>
    <span style={{ color: C.text, flex: 1 }}>{name}</span>
    <span style={{ color: C.extra, fontSize: 10 }}>Aanvullend</span>
  </div>
);

const Section = ({ title, children, count, color, defaultOpen = true }) => {
  const [open, setOpen] = useState(defaultOpen);
  if (!count) return null;
  return (
    <div style={{ marginBottom: 18 }}>
      <div onClick={() => setOpen(!open)} style={{
        display: "flex", alignItems: "center", gap: 10, cursor: "pointer",
        padding: "8px 0", borderBottom: `1px solid ${C.border}`, userSelect: "none",
      }}>
        <span style={{ flex: 1, fontWeight: 600, fontSize: 13, color: C.text }}>{title}</span>
        <span style={{ background: `${color}18`, color, border: `1px solid ${color}33`, borderRadius: 3, padding: "1px 8px", fontSize: 11, fontFamily: C.mono }}>{count}</span>
        <span style={{ color: C.textMuted, fontSize: 11 }}>{open ? "▲" : "▼"}</span>
      </div>
      {open && <div style={{ marginTop: 8 }}>{children}</div>}
    </div>
  );
};

// ── File Upload Panel ─────────────────────────────────────────────────────────
const FilePanel = ({ label, color, data, onLoad, onClear }) => {
  const inputRef = useRef();
  const [dragging, setDragging] = useState(false);

  const handleFile = async (file) => {
    try {
      const sheets = await readXlsx(file);
      onLoad({ file, sheets, activeSheet: 0, activeCol: guessCol(sheets[0].rows), startRow: 1 });
    } catch (e) { alert("Kon bestand niet lezen: " + e.message); }
  };

  if (data) {
    const { sheets, activeSheet, activeCol, startRow } = data;
    const sheet = sheets[activeSheet];
    const headers = sheet.rows[0] || [];
    const preview = extractCol(sheet.rows, activeCol, startRow).slice(0, 4);
    const total = extractCol(sheet.rows, activeCol, startRow).length;

    return (
      <div style={{ flex: 1, background: C.surface, border: `1px solid ${color}55`, borderRadius: 8, padding: 14, minWidth: 0 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <div>
            <div style={{ fontWeight: 700, color, fontSize: 13 }}>{label}</div>
            <div style={{ fontFamily: C.mono, color: C.textMuted, fontSize: 10, marginTop: 1 }}>{data.file.name} · {total} items</div>
          </div>
          <button onClick={onClear} style={{ background: "transparent", border: `1px solid ${C.border}`, color: C.textMuted, borderRadius: 4, padding: "2px 8px", fontSize: 10, fontFamily: C.mono, cursor: "pointer" }}>✕</button>
        </div>

        {sheets.length > 1 && (
          <div style={{ marginBottom: 8 }}>
            <div style={{ fontFamily: C.mono, fontSize: 9, color: C.textMuted, marginBottom: 3 }}>TABBLAD</div>
            <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
              {sheets.map((s, i) => (
                <button key={i} onClick={() => onLoad({ ...data, activeSheet: i, activeCol: guessCol(s.rows) })}
                  style={{ background: i === activeSheet ? color + "22" : "transparent", border: `1px solid ${i === activeSheet ? color : C.border}`, color: i === activeSheet ? color : C.textMuted, borderRadius: 4, padding: "2px 8px", fontSize: 10, fontFamily: C.mono, cursor: "pointer" }}>
                  {s.name}
                </button>
              ))}
            </div>
          </div>
        )}

        <div style={{ marginBottom: 8 }}>
          <div style={{ fontFamily: C.mono, fontSize: 9, color: C.textMuted, marginBottom: 3 }}>KOLOM</div>
          <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
            {headers.map((h, i) => (
              <button key={i} onClick={() => onLoad({ ...data, activeCol: i })}
                style={{ background: i === activeCol ? color + "22" : "transparent", border: `1px solid ${i === activeCol ? color : C.border}`, color: i === activeCol ? color : C.textMuted, borderRadius: 4, padding: "2px 8px", fontSize: 10, fontFamily: C.mono, cursor: "pointer" }}>
                {String(h || `Kol ${i + 1}`).substring(0, 20)}
              </button>
            ))}
          </div>
        </div>

        <div style={{ marginBottom: 8, display: "flex", alignItems: "center", gap: 8 }}>
          <div style={{ fontFamily: C.mono, fontSize: 9, color: C.textMuted }}>VANAF RIJ</div>
          <input type="number" min={1} value={startRow + 1}
            onChange={(e) => onLoad({ ...data, startRow: Math.max(0, parseInt(e.target.value, 10) - 1 || 0) })}
            style={{ width: 48, background: C.bg, border: `1px solid ${C.border}`, color: C.text, borderRadius: 4, padding: "2px 6px", fontFamily: C.mono, fontSize: 11, textAlign: "center" }} />
        </div>

        <div style={{ background: C.bg, borderRadius: 4, padding: "7px 10px", fontFamily: C.mono, fontSize: 10 }}>
          {preview.map((v, i) => <div key={i} style={{ color: C.textSub, marginBottom: 1 }}>· {v}</div>)}
          {total > 4 && <div style={{ color: C.textMuted, marginTop: 2 }}>… +{total - 4} meer</div>}
        </div>
      </div>
    );
  }

  return (
    <div
      onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => { e.preventDefault(); setDragging(false); handleFile(e.dataTransfer.files[0]); }}
      onClick={() => inputRef.current?.click()}
      style={{
        flex: 1, minWidth: 0, border: `2px dashed ${dragging ? color : C.border}`,
        borderRadius: 8, padding: "28px 20px", textAlign: "center", cursor: "pointer",
        background: dragging ? `${color}08` : C.surface, transition: "all 0.15s",
        display: "flex", flexDirection: "column", alignItems: "center", gap: 8,
      }}
    >
      <input ref={inputRef} type="file" accept=".xlsx,.xls,.ods" style={{ display: "none" }} onChange={(e) => handleFile(e.target.files[0])} />
      <div style={{ fontSize: 26 }}>📂</div>
      <div style={{ fontWeight: 700, color, fontSize: 13 }}>{label}</div>
      <div style={{ fontFamily: C.mono, color: C.textMuted, fontSize: 10 }}>.xlsx · .xls · .ods — sleep of klik</div>
    </div>
  );
};

// ── Main ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [ilpData, setIlpData] = useState(null);
  const [plData, setPlData] = useState(null);
  const [result, setResult] = useState(null);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState("");
  const [filter, setFilter] = useState("all");
  const [useAI, setUseAI] = useState(true);

  const run = async () => {
    setLoading(true);
    setResult(null);
    setProgress("Bestanden inlezen…");

    const ilpSheet = ilpData.sheets[ilpData.activeSheet];
    const plSheet = plData.sheets[plData.activeSheet];
    const ilpList = extractCol(ilpSheet.rows, ilpData.activeCol, ilpData.startRow);
    const plList = extractCol(plSheet.rows, plData.activeCol, plData.startRow);

    try {
      if (useAI) {
        const { results, extraItems } = await aiMatchAll(ilpList, plList, setProgress);
        const missing = results.filter((r) => r.type === "missing");
        const matched = results.filter((r) => r.type === "match");
        const matchedAi = results.filter((r) => r.type === "match_ai");
        const uncertain = results.filter((r) => r.type === "uncertain");
        setResult({ results, extraItems, missing, matched, matchedAi, uncertain, ilpCount: ilpList.length, plCount: plList.length, mode: "ai" });
      } else {
        // Simple exact/fuzzy compare
        const norm = (s) => s.toLowerCase().replace(/[\s\-_\/\.]+/g, " ").trim();
        const ilpNorm = ilpList.map(norm);
        const plNorm = plList.map(norm);
        const results = plList.map((plItem, i) => {
          const idx = ilpNorm.indexOf(plNorm[i]);
          return idx >= 0
            ? { plItem, matchedIlp: ilpList[idx], confidence: 1, reason: "Exacte overeenkomst", type: "match" }
            : { plItem, matchedIlp: null, confidence: 0, reason: "Niet gevonden", type: "missing" };
        });
        const matchedIlpSet = new Set(results.filter((r) => r.matchedIlp).map((r) => r.matchedIlp));
        const extraItems = ilpList.filter((i) => !matchedIlpSet.has(i));
        setResult({
          results, extraItems,
          missing: results.filter((r) => r.type === "missing"),
          matched: results.filter((r) => r.type === "match"),
          matchedAi: [], uncertain: [],
          ilpCount: ilpList.length, plCount: plList.length, mode: "exact",
        });
      }
    } catch (e) {
      alert("Fout bij analyse: " + e.message);
    }

    setLoading(false);
    setProgress("");
    setFilter("all");
  };

  const filterItems = (res) => {
    if (!res) return [];
    if (filter === "missing") return res.results.filter((r) => r.type === "missing");
    if (filter === "extra") return null; // handled separately
    if (filter === "match") return res.results.filter((r) => r.type === "match" || r.type === "match_ai");
    if (filter === "uncertain") return res.results.filter((r) => r.type === "uncertain");
    return res.results;
  };

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.text, fontFamily: C.sans, padding: "28px 20px" }}>
      <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&family=JetBrains+Mono:wght@400;600;700&display=swap" rel="stylesheet" />

      <div style={{ maxWidth: 920, margin: "0 auto" }}>
        {/* Header */}
        <div style={{ marginBottom: 26 }}>
          <div style={{ fontFamily: C.mono, fontSize: 10, letterSpacing: 3, color: C.accent, marginBottom: 5 }}>ILP VERGELIJKER</div>
          <h1 style={{ margin: 0, fontSize: 22, fontWeight: 700 }}>Informatie Leveringsplanning</h1>
          <div style={{ color: C.textMuted, fontSize: 13, marginTop: 4 }}>
            Upload twee xlsx-bestanden — AI interpreteert eisen en producten semantisch
          </div>
        </div>

        {/* Upload */}
        <div style={{ display: "flex", gap: 14, marginBottom: 14, flexWrap: "wrap" }}>
          <FilePanel label="Informatie Leveringsplanning (ILP)" color={C.accent} data={ilpData}
            onLoad={setIlpData} onClear={() => { setIlpData(null); setResult(null); }} />
          <FilePanel label="Productenlijst Opdrachtgever" color={C.warn} data={plData}
            onLoad={setPlData} onClear={() => { setPlData(null); setResult(null); }} />
        </div>

        {/* Mode + button */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 24, flexWrap: "wrap", gap: 10 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ fontFamily: C.mono, fontSize: 11, color: C.textMuted }}>ANALYSEMODUS</div>
            <button onClick={() => setUseAI(true)} style={{
              background: useAI ? C.ai + "22" : "transparent", border: `1px solid ${useAI ? C.ai : C.border}`,
              color: useAI ? C.ai : C.textMuted, borderRadius: 4, padding: "4px 14px",
              fontFamily: C.mono, fontSize: 11, cursor: "pointer",
            }}>✦ AI semantisch</button>
            <button onClick={() => setUseAI(false)} style={{
              background: !useAI ? C.accent + "22" : "transparent", border: `1px solid ${!useAI ? C.accent : C.border}`,
              color: !useAI ? C.accent : C.textMuted, borderRadius: 4, padding: "4px 14px",
              fontFamily: C.mono, fontSize: 11, cursor: "pointer",
            }}>Exacte tekst</button>
          </div>
          <button onClick={run} disabled={!ilpData || !plData || loading} style={{
            background: ilpData && plData && !loading ? C.accent : C.border,
            border: "none", color: ilpData && plData && !loading ? "#fff" : C.textMuted,
            borderRadius: 6, padding: "10px 32px", fontFamily: C.mono,
            fontSize: 12, fontWeight: 700, cursor: ilpData && plData && !loading ? "pointer" : "not-allowed",
            letterSpacing: 1, transition: "background 0.2s", minWidth: 180,
          }}>
            {loading ? "⏳ Analyseren…" : "VERGELIJKEN →"}
          </button>
        </div>

        {/* Progress */}
        {loading && (
          <div style={{
            background: C.aiDim, border: `1px solid ${C.ai}33`, borderRadius: 8,
            padding: "14px 18px", marginBottom: 20, fontFamily: C.mono, fontSize: 12, color: C.ai,
            display: "flex", alignItems: "center", gap: 10,
          }}>
            <span style={{ animation: "spin 1s linear infinite", display: "inline-block" }}>◌</span>
            <span>{progress}</span>
          </div>
        )}

        {/* Results */}
        {result && !loading && (
          <>
            {/* Stats */}
            <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
              <Stat label="ILP documenten" value={result.ilpCount} color={C.accent} />
              <Stat label="Opdrachtgever eisen" value={result.plCount} color={C.warn} />
              <Stat label="Ontbreekt in ILP" value={result.missing.length} color={C.missing} />
              <Stat label="Aanvullend in ILP" value={result.extraItems.length} color={C.extra} />
              <Stat label="Exacte match" value={result.matched.length} color={C.match} />
              {result.matchedAi.length > 0 && <Stat label="AI match" value={result.matchedAi.length} color={C.ai} />}
              {result.uncertain.length > 0 && <Stat label="Onzeker" value={result.uncertain.length} color={C.warn} />}
            </div>

            {result.mode === "ai" && (
              <div style={{
                background: C.aiDim, border: `1px solid ${C.ai}33`, borderRadius: 6,
                padding: "8px 14px", marginBottom: 16, fontFamily: C.mono, fontSize: 11, color: C.ai,
              }}>
                ✦ AI semantische analyse — documenten worden vergeleken op betekenis, niet alleen op tekst
              </div>
            )}

            {/* Filter tabs */}
            <div style={{ display: "flex", gap: 6, marginBottom: 14, flexWrap: "wrap" }}>
              {[
                { key: "all", label: "Alles" },
                { key: "missing", label: `Ontbreekt (${result.missing.length})` },
                { key: "extra", label: `Aanvullend (${result.extraItems.length})` },
                { key: "match", label: `Match (${result.matched.length + result.matchedAi.length})` },
                ...(result.uncertain.length ? [{ key: "uncertain", label: `Onzeker (${result.uncertain.length})` }] : []),
              ].map(({ key, label }) => (
                <button key={key} onClick={() => setFilter(key)} style={{
                  background: filter === key ? C.accent + "22" : "transparent",
                  border: `1px solid ${filter === key ? C.accent : C.border}`,
                  color: filter === key ? C.accent : C.textMuted,
                  borderRadius: 4, padding: "4px 14px", fontFamily: C.mono, fontSize: 11, cursor: "pointer",
                }}>{label}</button>
              ))}
            </div>

            {/* Result panels */}
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10, padding: "18px 20px" }}>
              {filter !== "extra" && (
                <>
                  {(filter === "all" || filter === "missing") && (
                    <Section title="Ontbrekende documenten in ILP" count={result.missing.length} color={C.missing}>
                      {result.missing.map((item, i) => <MatchRow key={i} item={item} />)}
                    </Section>
                  )}
                  {(filter === "all" || filter === "uncertain") && result.uncertain.length > 0 && (
                    <Section title="Onzekere overeenkomsten (controleer handmatig)" count={result.uncertain.length} color={C.warn}>
                      {result.uncertain.map((item, i) => <MatchRow key={i} item={item} />)}
                    </Section>
                  )}
                  {(filter === "all" || filter === "match") && (
                    <>
                      {result.matchedAi.length > 0 && (
                        <Section title="AI-herkende overeenkomsten" count={result.matchedAi.length} color={C.ai} defaultOpen={false}>
                          {result.matchedAi.map((item, i) => <MatchRow key={i} item={item} />)}
                        </Section>
                      )}
                      <Section title="Exacte overeenkomsten" count={result.matched.length} color={C.match} defaultOpen={false}>
                        {result.matched.map((item, i) => <MatchRow key={i} item={item} />)}
                      </Section>
                    </>
                  )}
                </>
              )}
              {(filter === "all" || filter === "extra") && (
                <Section title="Aanvullende documenten in ILP (niet gevraagd door opdrachtgever)" count={result.extraItems.length} color={C.extra}>
                  {result.extraItems.map((name, i) => <ExtraRow key={i} name={name} />)}
                </Section>
              )}
              {result.missing.length === 0 && result.extraItems.length === 0 && result.uncertain.length === 0 && (
                <div style={{ textAlign: "center", padding: "24px 0", color: C.extra, fontFamily: C.mono, fontSize: 13 }}>
                  ✓ Alle eisen zijn gedekt — geen ontbrekende documenten
                </div>
              )}
            </div>

            {/* Legend */}
            <div style={{ display: "flex", gap: 16, marginTop: 12, flexWrap: "wrap", fontFamily: C.mono, fontSize: 10, color: C.textMuted }}>
              <span><span style={{ color: C.missing }}>✗</span> Ontbreekt in ILP</span>
              <span><span style={{ color: C.extra }}>+</span> Aanvullend in ILP</span>
              <span><span style={{ color: C.match }}>✓</span> Exacte match</span>
              <span><span style={{ color: C.ai }}>≈</span> AI semantische match</span>
              <span><span style={{ color: C.warn }}>?</span> Onzeker — handmatig controleren</span>
            </div>
          </>
        )}
      </div>
      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}
