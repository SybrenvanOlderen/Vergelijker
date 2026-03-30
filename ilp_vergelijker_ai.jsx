const { useState, useRef } = React;
const XLSX = window.XLSX;

// ── Kleuren ─────────────────────────────────────────────────
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

// ── XLSX lezen ───────────────────────────────────────────────
const readXlsx = (file) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        res(
          wb.SheetNames.map((name) => ({
            name,
            rows: XLSX.utils.sheet_to_json(wb.Sheets[name], {
              header: 1,
              defval: "",
            }),
          }))
        );
      } catch (err) {
        rej(err);
      }
    };
    r.onerror = () => rej(new Error("Leesfout"));
    r.readAsArrayBuffer(file);
  });

const guessCol = (rows) => {
  if (!rows?.length) return 0;
  const headers = rows[0] || [];
  const colCount = Math.max(...rows.map((r) => r.length), 0);
  let best = 0,
    bestScore = -1;
  for (let c = 0; c < colCount; c++) {
    let score = 0;
    const h = String(headers[c] || "").toLowerCase();
    if (
      /document|naam|titel|omschrijving|product|deliverable|name|title/.test(h)
    )
      score += 10;
    for (let r = 1; r < Math.min(rows.length, 20); r++) {
      const v = String(rows[r][c] || "").trim();
      if (v.length > 3) score += 1;
      if (v.length > 10) score += 1;
    }
    if (score > bestScore) {
      bestScore = score;
      best = c;
    }
  }
  return best;
};

const extractCol = (rows, col, startRow = 1) =>
  rows
    .slice(startRow)
    .map((r) => String(r[col] || "").trim())
    .filter((v) => v.length > 1);

// ── Simpele vergelijking (zonder AI backend) ─────────────────
const compareLists = (ilpList, plList) => {
  const norm = (s) =>
    s.toLowerCase().replace(/[\s\-_\/\.]+/g, " ").trim();

  const ilpNorm = ilpList.map(norm);

  const results = plList.map((plItem) => {
    const idx = ilpNorm.indexOf(norm(plItem));
    return idx >= 0
      ? {
          plItem,
          matchedIlp: ilpList[idx],
          confidence: 1,
          reason: "Exacte overeenkomst",
          type: "match",
        }
      : {
          plItem,
          matchedIlp: null,
          confidence: 0,
          reason: "Niet gevonden",
          type: "missing",
        };
  });

  const matchedIlpSet = new Set(
    results.filter((r) => r.matchedIlp).map((r) => r.matchedIlp)
  );
  const extraItems = ilpList.filter((i) => !matchedIlpSet.has(i));

  return {
    results,
    extraItems,
    missing: results.filter((r) => r.type === "missing"),
    matched: results.filter((r) => r.type === "match"),
    ilpCount: ilpList.length,
    plCount: plList.length,
  };
};

// ── UI Components ─────────────────────────────────────────────
const Stat = ({ label, value, color }) => (
  <div
    style={{
      flex: 1,
      minWidth: 90,
      background: C.surface,
      border: `1px solid ${C.border}`,
      borderRadius: 8,
      padding: "12px 16px",
      textAlign: "center",
    }}
  >
    <div
      style={{
        fontFamily: C.mono,
        fontSize: 24,
        fontWeight: 700,
        color,
      }}
    >
      {value}
    </div>
    <div style={{ fontSize: 11, color: C.textMuted }}>{label}</div>
  </div>
);

// ── File Upload Panel ─────────────────────────────────────────
const FilePanel = ({ label, color, data, onLoad, onClear }) => {
  const inputRef = useRef();

  const handleFile = async (file) => {
    try {
      const sheets = await readXlsx(file);
      onLoad({
        file,
        sheets,
        activeSheet: 0,
        activeCol: guessCol(sheets[0].rows),
        startRow: 1,
      });
    } catch (e) {
      alert("Kon bestand niet lezen");
    }
  };

  if (data) {
    return (
      <div
        style={{
          flex: 1,
          background: C.surface,
          border: `1px solid ${color}55`,
          borderRadius: 8,
          padding: 14,
        }}
      >
        <div style={{ marginBottom: 10 }}>
          <div style={{ fontWeight: 700, color }}>{label}</div>
          <div style={{ fontSize: 10, color: C.textMuted }}>
            {data.file.name}
          </div>
        </div>

        <button onClick={onClear}>Verwijder</button>
      </div>
    );
  }

  return (
    <div
      onClick={() => inputRef.current?.click()}
      style={{
        flex: 1,
        border: `2px dashed ${color}`,
        borderRadius: 8,
        padding: "30px",
        textAlign: "center",
        cursor: "pointer",
        background: C.surface,
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls"
        style={{ display: "none" }}
        onChange={(e) => handleFile(e.target.files[0])}
      />
      <div>{label}</div>
    </div>
  );
};

// ── Main App ──────────────────────────────────────────────────
function App() {
  const [ilpData, setIlpData] = useState(null);
  const [plData, setPlData] = useState(null);
  const [result, setResult] = useState(null);

  const run = async () => {
    const ilpSheet = ilpData.sheets[ilpData.activeSheet];
    const plSheet = plData.sheets[plData.activeSheet];

    const ilpList = extractCol(
      ilpSheet.rows,
      ilpData.activeCol,
      ilpData.startRow
    );
    const plList = extractCol(
      plSheet.rows,
      plData.activeCol,
      plData.startRow
    );

    setResult(compareLists(ilpList, plList));
  };

  return (
    <div
      style={{
        minHeight: "100vh",
        background: C.bg,
        color: C.text,
        fontFamily: C.sans,
        padding: 20,
      }}
    >
      <h1>ILP Vergelijker</h1>

      <div style={{ display: "flex", gap: 10 }}>
        <FilePanel
          label="ILP bestand"
          color={C.accent}
          data={ilpData}
          onLoad={setIlpData}
          onClear={() => setIlpData(null)}
        />
        <FilePanel
          label="Productenlijst"
          color={C.warn}
          data={plData}
          onLoad={setPlData}
          onClear={() => setPlData(null)}
        />
      </div>

      <button
        onClick={run}
        disabled={!ilpData || !plData}
        style={{ marginTop: 20, padding: "10px 20px" }}
      >
        Vergelijken
      </button>

      {result && (
        <div style={{ marginTop: 20 }}>
          <div style={{ display: "flex", gap: 10 }}>
            <Stat
              label="ILP documenten"
              value={result.ilpCount}
              color={C.accent}
            />
            <Stat
              label="Eisen"
              value={result.plCount}
              color={C.warn}
            />
            <Stat
              label="Ontbreekt"
              value={result.missing.length}
              color={C.missing}
            />
            <Stat
              label="Aanvullend"
              value={result.extraItems.length}
              color={C.extra}
            />
          </div>
        </div>
      )}
    </div>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(<App />);
