import React, { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  Upload, Check, X, Download, LogIn, UserX, Search, Plus,
  ArrowRight, ArrowLeft, Save, ClipboardList, Trash2, FileSpreadsheet,
  ChevronDown, RotateCcw
} from "lucide-react";

// ---------- theme tokens ----------
const INK = "#232920";
const PAPER = "#F4EFE1";
const PAPER_DARK = "#EAE2CC";
const LEDGER = "#26463A";
const LEDGER_DARK = "#173026";
const STAMP = "#A63D31";
const PRESENT = "#3E6B4F";
const LINE = "#CBBF9E";

const serif = { fontFamily: "Georgia, 'Times New Roman', serif" };
const mono = { fontFamily: "ui-monospace, SFMono-Regular, Menlo, Consolas, monospace" };

const KEYWORDS = ["name","roll","no","number","id","student","class","section","regd","register","admission","enroll"];

function detectHeaderRow(aoa) {
  let best = 0, bestScore = -1;
  const limit = Math.min(10, aoa.length);
  for (let i = 0; i < limit; i++) {
    const row = aoa[i] || [];
    let score = 0;
    row.forEach(cell => {
      const c = String(cell || "").trim().toLowerCase();
      if (c && KEYWORDS.some(k => c.includes(k))) score++;
    });
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return bestScore > 0 ? best : 0;
}

function detectColumn(headers, matchers, fallback) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").trim().toLowerCase();
    if (matchers.some(m => h.includes(m))) return i;
  }
  return fallback;
}

function todayLabel() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}-${mm}-${d.getFullYear()}`;
}

function downloadAOA(aoa, filename) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Attendance");
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([out], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---------- responsive helper ----------
function useIsMobile(breakpoint = 640) {
  const [isMobile, setIsMobile] = useState(
    typeof window !== "undefined" ? window.innerWidth <= breakpoint : false
  );
  useEffect(() => {
    const mq = window.matchMedia(`(max-width: ${breakpoint}px)`);
    const onChange = () => setIsMobile(mq.matches);
    onChange();
    mq.addEventListener ? mq.addEventListener("change", onChange) : mq.addListener(onChange);
    return () => {
      mq.removeEventListener ? mq.removeEventListener("change", onChange) : mq.removeListener(onChange);
    };
  }, [breakpoint]);
  return isMobile;
}

// step tab
function Tab({ n, label, active, done, isMobile }) {
  return (
    <div style={{
      padding: isMobile ? "7px 10px" : "8px 16px", borderRadius: "6px 6px 0 0",
      background: active ? PAPER : PAPER_DARK,
      border: `1px solid ${LINE}`, borderBottom: active ? `1px solid ${PAPER}` : `1px solid ${LINE}`,
      color: active ? LEDGER_DARK : "#8a806a", fontSize: isMobile ? 10.5 : 12, letterSpacing: "0.06em",
      textTransform: "uppercase", ...serif, fontWeight: active ? 700 : 400,
      marginBottom: -1, whiteSpace: "nowrap", flexShrink: 0
    }}>
      {done ? "✓ " : `${n}. `}{label}
    </div>
  );
}

export default function AttendanceLedger() {
  const isMobile = useIsMobile();
  const [step, setStep] = useState("home");
  const [email, setEmail] = useState("");
  const [session, setSession] = useState(null); // {email} or {guest:true}
  const [emailInput, setEmailInput] = useState("");
  const [savedSessions, setSavedSessions] = useState([]);
  const [loadingSessions, setLoadingSessions] = useState(false);

  const [fileName, setFileName] = useState("");
  const [rawAOA, setRawAOA] = useState(null);
  const [headerRowIndex, setHeaderRowIndex] = useState(0);
  const [nameColIdx, setNameColIdx] = useState(null);
  const [rollColIdx, setRollColIdx] = useState(null);

  const [attendance, setAttendance] = useState({});
  const [search, setSearch] = useState("");
  const [pollEnabled, setPollEnabled] = useState(false);
  const [pollTitle, setPollTitle] = useState("");
  const [pollOptionsText, setPollOptionsText] = useState("");
  const [pollAnswers, setPollAnswers] = useState({});

  const [error, setError] = useState("");
  const fileInputRef = useRef(null);

  const headers = rawAOA ? (rawAOA[headerRowIndex] || []) : [];
  const dataRows = rawAOA ? rawAOA.slice(headerRowIndex + 1) : [];
  const pollOptions = useMemo(
    () => pollOptionsText.split(",").map(s => s.trim()).filter(Boolean),
    [pollOptionsText]
  );

  const students = useMemo(() => {
    return dataRows.map((row, i) => ({
      idx: i,
      row,
      name: nameColIdx != null ? String(row[nameColIdx] ?? "").trim() : "",
      roll: rollColIdx != null ? String(row[rollColIdx] ?? "").trim() : "",
    })).filter(s => s.name || s.roll || s.row.some(c => String(c ?? "").trim()));
  }, [dataRows, nameColIdx, rollColIdx]);

  const filteredStudents = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return students;
    return students.filter(s =>
      s.name.toLowerCase().includes(q) || s.roll.toLowerCase().includes(q)
    );
  }, [students, search]);

  const presentCount = Object.values(attendance).filter(v => v === "Present").length;
  // Anything not explicitly marked "Present" is treated as absent by default —
  // both here and at export time — so a teacher only needs to tap the students
  // who showed up (or tap "Absent" explicitly if they prefer to record it that way).
  const absentCount = students.length - presentCount;

  // ---------- storage (simulated per-email drive) ----------
  async function fetchSessions(em) {
    if (!em) return;
    setLoadingSessions(true);
    try {
      const res = await window.storage.get(`sessions:${em}`, true);
      const list = res ? JSON.parse(res.value) : [];
      setSavedSessions(list);
    } catch {
      setSavedSessions([]);
    }
    setLoadingSessions(false);
  }

  async function saveSessionToAccount(finalHeaders, finalAOA, dateLabel) {
    if (!session || !session.email) return;
    try {
      const entry = {
        id: `${Date.now()}`,
        fileName: fileName || "attendance",
        date: dateLabel,
        savedAt: new Date().toISOString(),
        headers: finalHeaders,
        aoa: finalAOA,
      };
      const existingRes = await window.storage.get(`sessions:${session.email}`, true).catch(() => null);
      const list = existingRes ? JSON.parse(existingRes.value) : [];
      const updated = [entry, ...list].slice(0, 20);
      await window.storage.set(`sessions:${session.email}`, JSON.stringify(updated), true);
      setSavedSessions(updated);
    } catch (e) {
      setError("Couldn't save to your account storage, but your file downloaded fine.");
    }
  }

  function reDownloadSession(entry) {
    downloadAOA(entry.aoa, `${entry.fileName.replace(/\.xlsx?$/i, "")}_${entry.date}.xlsx`);
  }

  async function deleteSession(id) {
    const updated = savedSessions.filter(s => s.id !== id);
    setSavedSessions(updated);
    if (session?.email) {
      await window.storage.set(`sessions:${session.email}`, JSON.stringify(updated), true).catch(() => {});
    }
  }

  // ---------- file handling ----------
  function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = new Uint8Array(ev.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (!aoa.length) { setError("That sheet looks empty."); return; }
        const hIdx = detectHeaderRow(aoa);
        const hdrs = aoa[hIdx] || [];
        const nIdx = detectColumn(hdrs, ["name"], 0);
        const rIdx = detectColumn(hdrs, ["roll", "regd", "admission", "enroll", "number", "id"], hdrs.length > 1 ? 1 : null);
        setFileName(file.name);
        setRawAOA(aoa);
        setHeaderRowIndex(hIdx);
        setNameColIdx(nIdx);
        setRollColIdx(rIdx === nIdx ? null : rIdx);
        setAttendance({});
        setPollAnswers({});
        setStep("mapping");
      } catch (err) {
        setError("Couldn't read that file. Please upload a .xlsx or .xls file.");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function setMark(idx, value) {
    setAttendance(prev => ({ ...prev, [idx]: prev[idx] === value ? "" : value }));
  }
  function markAll(value) {
    const next = {};
    filteredStudents.forEach(s => { next[s.idx] = value; });
    setAttendance(prev => ({ ...prev, ...next }));
  }
  function resetMarks() {
    setAttendance({});
    setPollAnswers({});
  }

  function buildExport() {
    const dateLabel = todayLabel();
    const finalHeaders = [...headers, dateLabel, ...(pollEnabled && pollTitle.trim() ? [pollTitle.trim()] : [])];
    const finalAOA = rawAOA.map((row, i) => {
      if (i === headerRowIndex) return finalHeaders;
      if (i > headerRowIndex) {
        const dIdx = i - headerRowIndex - 1;
        const extra = [attendance[dIdx] || "Absent"];
        if (pollEnabled && pollTitle.trim()) extra.push(pollAnswers[dIdx] || "");
        return [...row, ...extra];
      }
      // rows above header: pad
      const pad = finalHeaders.length - row.length;
      return pad > 0 ? [...row, ...Array(pad).fill("")] : row;
    });
    return { finalHeaders, finalAOA, dateLabel };
  }

  async function handleExport(alsoSave) {
    const { finalHeaders, finalAOA, dateLabel } = buildExport();
    const base = fileName ? fileName.replace(/\.xlsx?$/i, "") : "attendance";
    downloadAOA(finalAOA, `${base}_${dateLabel}.xlsx`);
    if (alsoSave && session?.email) {
      await saveSessionToAccount(finalHeaders, finalAOA, dateLabel);
    }
    setStep("done");
  }

  function startOver() {
    setStep("upload");
    setFileName(""); setRawAOA(null); setHeaderRowIndex(0);
    setNameColIdx(null); setRollColIdx(null); setAttendance({});
    setPollAnswers({}); setPollEnabled(false); setPollTitle(""); setPollOptionsText("");
    setSearch(""); setError("");
  }

  const steps = ["upload", "mapping", "configure", "take", "export", "done"];
  const stepIndex = steps.indexOf(step);

  const pageBg = {
    minHeight: "100%",
    background: PAPER_DARK,
    backgroundImage: `repeating-linear-gradient(${PAPER_DARK}, ${PAPER_DARK} 27px, ${LINE}55 28px)`,
    padding: isMobile ? "12px 8px" : "28px 16px",
  };

  const card = {
    maxWidth: 880, margin: "0 auto", background: PAPER,
    border: `1px solid ${LINE}`, borderRadius: isMobile ? 8 : 10,
    boxShadow: "0 6px 24px rgba(23,48,38,0.15)",
    overflow: "hidden",
    width: "100%",
    boxSizing: "border-box",
  };

  // ---------- HOME ----------
  if (step === "home") {
    return (
      <div style={pageBg}>
        <div style={{ ...card, maxWidth: 620 }}>
          <div style={{ background: LEDGER, color: PAPER, padding: "28px 32px", position: "relative" }}>
            <div style={{ ...serif, fontSize: 28, fontWeight: 700, letterSpacing: "0.02em" }}>The Roll Ledger</div>
            <div style={{ ...serif, fontSize: 13, opacity: 0.8, marginTop: 4 }}>
              Upload a roster, take attendance, export a dated column — every time.
            </div>
          </div>
          <div style={{ padding: 32 }}>
            <div style={{ ...serif, fontSize: 14, color: INK, marginBottom: 20, lineHeight: 1.5 }}>
              Work without an account, or use an email as a name-tag so today's
              register gets saved here for next time.
            </div>

            <button
              onClick={() => { setSession({ guest: true }); setStep("upload"); }}
              style={{
                width: "100%", padding: "12px 16px", marginBottom: 14, borderRadius: 8,
                border: `1px solid ${LEDGER}`, background: PAPER, color: LEDGER_DARK,
                display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
                cursor: "pointer", ...serif, fontSize: 15, fontWeight: 700,
              }}
            >
              <UserX size={17} /> Continue without an account
            </button>

            <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "18px 0", color: "#8a806a" }}>
              <div style={{ flex: 1, height: 1, background: LINE }} />
              <span style={{ fontSize: 12, ...serif }}>or</span>
              <div style={{ flex: 1, height: 1, background: LINE }} />
            </div>

            <div style={{ display: "flex", gap: 8 }}>
              <input
                type="email"
                value={emailInput}
                onChange={e => setEmailInput(e.target.value)}
                placeholder="you@school.edu"
                style={{
                  flex: 1, padding: "10px 12px", borderRadius: 8, border: `1px solid ${LINE}`,
                  background: "#fff", ...mono, fontSize: 14, color: INK
                }}
              />
              <button
                disabled={!emailInput.trim()}
                onClick={() => {
                  const em = emailInput.trim().toLowerCase();
                  setSession({ email: em });
                  fetchSessions(em);
                  setStep("upload");
                }}
                style={{
                  padding: "10px 16px", borderRadius: 8, border: "none",
                  background: emailInput.trim() ? STAMP : "#c9bf9e", color: PAPER,
                  ...serif, fontWeight: 700, cursor: emailInput.trim() ? "pointer" : "not-allowed",
                  display: "flex", alignItems: "center", gap: 6,
                }}
              >
                <LogIn size={16} /> Continue
              </button>
            </div>
            <div style={{ fontSize: 11.5, color: "#8a806a", marginTop: 10, lineHeight: 1.5, ...serif }}>
              This is a simple name-tag, not a secured account — anyone who
              types the same email sees the same saved registers. Don't use it
              for anything sensitive.
            </div>
          </div>
        </div>
      </div>
    );
  }

  const emailBadge = session?.email ? (
    <div style={{
      position: "absolute", top: 16, right: 20, fontSize: 11.5, color: PAPER,
      opacity: 0.85, ...mono
    }}>{session.email}</div>
  ) : (
    <div style={{
      position: "absolute", top: 16, right: 20, fontSize: 11.5, color: PAPER,
      opacity: 0.85, ...serif
    }}>guest</div>
  );

  // ---------- shared shell ----------
  return (
    <div style={pageBg}>
      <div style={{
        maxWidth: 880, margin: "0 auto 0", display: "flex", gap: 2, paddingLeft: 4,
        overflowX: "auto", WebkitOverflowScrolling: "touch"
      }}>
        {["Upload", "Map columns", "Set up day", "Take roll", "Export"].map((label, i) => (
          <Tab key={label} n={i + 1} label={label} active={i === Math.min(stepIndex, 4)} done={i < stepIndex} isMobile={isMobile} />
        ))}
      </div>
      <div style={card}>
        <div style={{ background: LEDGER, color: PAPER, padding: isMobile ? "14px 16px" : "18px 28px", position: "relative" }}>
          <div style={{ ...serif, fontSize: isMobile ? 17 : 20, fontWeight: 700 }}>The Roll Ledger</div>
          {emailBadge}
        </div>

        <div style={{ padding: isMobile ? 16 : 28 }}>
          {error && (
            <div style={{
              background: "#f6e2df", border: `1px solid ${STAMP}`, color: STAMP,
              borderRadius: 8, padding: "10px 14px", marginBottom: 18, fontSize: 13.5, ...serif
            }}>{error}</div>
          )}

          {/* ---------- UPLOAD ---------- */}
          {step === "upload" && (
            <div>
              <h2 style={{ ...serif, fontSize: 18, color: LEDGER_DARK, marginBottom: 6 }}>Upload today's roster</h2>
              <p style={{ ...serif, fontSize: 13.5, color: "#5b5343", marginBottom: 20 }}>
                Bring in an .xlsx or .xls file. If you're re-using yesterday's
                export, that works too — the ledger just adds another column.
              </p>

              {session?.email && (
                <div style={{ marginBottom: 24 }}>
                  <div style={{ ...serif, fontSize: 13, fontWeight: 700, color: LEDGER_DARK, marginBottom: 8 }}>
                    Your saved registers
                  </div>
                  {loadingSessions && <div style={{ fontSize: 13, color: "#8a806a" }}>Loading…</div>}
                  {!loadingSessions && savedSessions.length === 0 && (
                    <div style={{ fontSize: 13, color: "#8a806a", ...serif }}>Nothing saved yet under this email.</div>
                  )}
                  {savedSessions.map(s => (
                    <div key={s.id} style={{
                      display: "flex", alignItems: "center", justifyContent: "space-between",
                      padding: "8px 12px", border: `1px solid ${LINE}`, borderRadius: 8, marginBottom: 6,
                      background: "#fff"
                    }}>
                      <div style={{ fontSize: 13, ...mono, color: INK }}>{s.fileName} — {s.date}</div>
                      <div style={{ display: "flex", gap: 6 }}>
                        <button onClick={() => reDownloadSession(s)} style={{ border: "none", background: "none", cursor: "pointer", color: LEDGER }}>
                          <Download size={16} />
                        </button>
                        <button onClick={() => deleteSession(s.id)} style={{ border: "none", background: "none", cursor: "pointer", color: STAMP }}>
                          <Trash2 size={16} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}

              <div
                onClick={() => fileInputRef.current?.click()}
                style={{
                  border: `2px dashed ${LEDGER}`, borderRadius: 10, padding: "40px 20px",
                  textAlign: "center", cursor: "pointer", background: "#fdfbf4"
                }}
              >
                <Upload size={28} style={{ color: LEDGER, marginBottom: 10 }} />
                <div style={{ ...serif, color: LEDGER_DARK, fontWeight: 700 }}>Click to choose a spreadsheet</div>
                <div style={{ fontSize: 12, color: "#8a806a", marginTop: 4 }}>.xlsx or .xls</div>
                <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
              </div>
            </div>
          )}

          {/* ---------- MAPPING ---------- */}
          {step === "mapping" && rawAOA && (
            <div>
              <h2 style={{ ...serif, fontSize: 18, color: LEDGER_DARK, marginBottom: 6 }}>Confirm the header row</h2>
              <p style={{ ...serif, fontSize: 13.5, color: "#5b5343", marginBottom: 16 }}>
                We guessed which row holds the column titles, and which columns
                are name and roll number. Adjust anything that's off.
              </p>

              <div style={{ marginBottom: 16 }}>
                <label style={{ fontSize: 12.5, fontWeight: 700, color: LEDGER_DARK, ...serif }}>Header row</label>
                <select
                  value={headerRowIndex}
                  onChange={e => setHeaderRowIndex(Number(e.target.value))}
                  style={selectStyle}
                >
                  {rawAOA.slice(0, 10).map((row, i) => (
                    <option key={i} value={i}>Row {i + 1}: {row.slice(0, 5).join(" | ") || "(blank)"}</option>
                  ))}
                </select>
              </div>

              <div style={{ display: "flex", flexDirection: isMobile ? "column" : "row", gap: 16, marginBottom: 20 }}>
                <div style={{ flex: 1 }}>
                  <label style={{ fontSize: 12.5, fontWeight: 700, color: LEDGER_DARK, ...serif }}>Name column</label>
                  <select value={nameColIdx ?? ""} onChange={e => setNameColIdx(e.target.value === "" ? null : Number(e.target.value))} style={selectStyle}>
                    <option value="">— none —</option>
                    {headers.map((h, i) => <option key={i} value={i}>{h || `Column ${i + 1}`}</option>)}
                  </select>
                </div>
                <div style={{ flex: 1 }}>
                  <label style={{ fontSize: 12.5, fontWeight: 700, color: LEDGER_DARK, ...serif }}>Roll number column</label>
                  <select value={rollColIdx ?? ""} onChange={e => setRollColIdx(e.target.value === "" ? null : Number(e.target.value))} style={selectStyle}>
                    <option value="">— none —</option>
                    {headers.map((h, i) => <option key={i} value={i}>{h || `Column ${i + 1}`}</option>)}
                  </select>
                </div>
              </div>

              <div style={{ border: `1px solid ${LINE}`, borderRadius: 8, overflow: "hidden", marginBottom: 20 }}>
                <div style={{ maxHeight: 220, overflow: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, ...mono }}>
                    <thead>
                      <tr style={{ background: PAPER_DARK }}>
                        {headers.map((h, i) => (
                          <th key={i} style={{ padding: "6px 10px", textAlign: "left", borderBottom: `1px solid ${LINE}` }}>
                            {h || `Col ${i + 1}`}{i === nameColIdx ? " •name" : ""}{i === rollColIdx ? " •roll" : ""}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {dataRows.slice(0, 5).map((row, ri) => (
                        <tr key={ri}>
                          {headers.map((_, ci) => (
                            <td key={ci} style={{ padding: "6px 10px", borderBottom: `1px solid ${LINE}55` }}>{String(row[ci] ?? "")}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <NavButtons back={() => setStep("upload")} next={() => setStep("configure")} nextLabel="Set up today" />
            </div>
          )}

          {/* ---------- CONFIGURE ---------- */}
          {step === "configure" && (
            <div>
              <h2 style={{ ...serif, fontSize: 18, color: LEDGER_DARK, marginBottom: 6 }}>Set up today's columns</h2>
              <p style={{ ...serif, fontSize: 13.5, color: "#5b5343", marginBottom: 16 }}>
                The export always gets a column named with today's date: <b>{todayLabel()}</b>,
                marked Present or Absent per student. You can also add one more
                column that only accepts a fixed set of answers — like a quick poll.
              </p>

              <label style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, cursor: "pointer" }}>
                <input type="checkbox" checked={pollEnabled} onChange={e => setPollEnabled(e.target.checked)} />
                <span style={{ ...serif, fontSize: 14, color: INK, fontWeight: 700 }}>Add an extra poll-style column</span>
              </label>

              {pollEnabled && (
                <div style={{ paddingLeft: 4, marginBottom: 10 }}>
                  <label style={{ fontSize: 12.5, fontWeight: 700, color: LEDGER_DARK, ...serif }}>Column title</label>
                  <input
                    value={pollTitle} onChange={e => setPollTitle(e.target.value)}
                    placeholder="e.g. Lunch choice"
                    style={{ ...selectStyle, marginBottom: 12 }}
                  />
                  <label style={{ fontSize: 12.5, fontWeight: 700, color: LEDGER_DARK, ...serif }}>
                    Allowed answers (comma-separated)
                  </label>
                  <input
                    value={pollOptionsText} onChange={e => setPollOptionsText(e.target.value)}
                    placeholder="e.g. Veg, Non-veg, Skipping"
                    style={selectStyle}
                  />
                  {pollOptions.length > 0 && (
                    <div style={{ marginTop: 8, fontSize: 12.5, color: "#5b5343", ...serif }}>
                      Options: {pollOptions.map(o => <span key={o} style={{ background: PAPER_DARK, padding: "2px 8px", borderRadius: 12, marginRight: 6 }}>{o}</span>)}
                    </div>
                  )}
                </div>
              )}

              <NavButtons back={() => setStep("mapping")} next={() => setStep("take")} nextLabel="Start taking roll" />
            </div>
          )}

          {/* ---------- TAKE ---------- */}
          {step === "take" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
                <h2 style={{ ...serif, fontSize: 18, color: LEDGER_DARK, margin: 0 }}>Take roll — {todayLabel()}</h2>
                <div style={{ fontSize: 12.5, ...serif, color: "#5b5343" }}>
                  <span style={{ color: PRESENT, fontWeight: 700 }}>{presentCount} present</span>
                  {"  ·  "}
                  <span style={{ color: STAMP, fontWeight: 700 }}>{absentCount} absent</span>
                </div>
              </div>
              <div style={{ fontSize: 11.5, color: "#8a806a", marginTop: -8, marginBottom: 14, ...serif }}>
                Anyone you haven't marked present will be recorded as absent automatically.
              </div>

              <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
                <div style={{ position: "relative", flex: "1 1 220px" }}>
                  <Search size={15} style={{ position: "absolute", left: 10, top: 10, color: "#8a806a" }} />
                  <input
                    value={search} onChange={e => setSearch(e.target.value)}
                    placeholder="Search by name or roll number"
                    style={{ ...selectStyle, paddingLeft: 32, marginBottom: 0 }}
                  />
                </div>
                <button onClick={() => markAll("Present")} style={pillBtn(PRESENT)}>Mark visible present</button>
                <button onClick={() => markAll("Absent")} style={pillBtn(STAMP)}>Mark visible absent</button>
                <button onClick={resetMarks} style={pillBtn("#8a806a")}><RotateCcw size={13} style={{ marginRight: 4 }} />Reset</button>
              </div>

              <div style={{ border: `1px solid ${LINE}`, borderRadius: 8, maxHeight: 360, overflow: "auto" }}>
                <table style={{ width: "100%", minWidth: isMobile ? 460 : "auto", borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ background: PAPER_DARK, position: "sticky", top: 0 }}>
                      <th style={thStyle}>Name</th>
                      <th style={thStyle}>Roll</th>
                      <th style={{ ...thStyle, textAlign: "center" }}>Attendance</th>
                      {pollEnabled && pollTitle.trim() && <th style={thStyle}>{pollTitle}</th>}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredStudents.map(s => (
                      <tr key={s.idx} style={{ borderBottom: `1px solid ${LINE}55` }}>
                        <td style={{ ...tdStyle, ...serif }}>{s.name || "—"}</td>
                        <td style={{ ...tdStyle, ...mono }}>{s.roll || "—"}</td>
                        <td style={{ ...tdStyle, textAlign: "center" }}>
                          <button onClick={() => setMark(s.idx, "Present")} style={markBtn(attendance[s.idx] === "Present", PRESENT, isMobile)}>
                            <Check size={14} />
                          </button>
                          <button onClick={() => setMark(s.idx, "Absent")} style={markBtn(attendance[s.idx] === "Absent", STAMP, isMobile)}>
                            <X size={14} />
                          </button>
                        </td>
                        {pollEnabled && pollTitle.trim() && (
                          <td style={tdStyle}>
                            <select
                              value={pollAnswers[s.idx] || ""}
                              onChange={e => setPollAnswers(prev => ({ ...prev, [s.idx]: e.target.value }))}
                              style={{ ...selectStyle, marginBottom: 0, padding: "4px 8px", fontSize: 12.5 }}
                            >
                              <option value="">—</option>
                              {pollOptions.map(o => <option key={o} value={o}>{o}</option>)}
                            </select>
                          </td>
                        )}
                      </tr>
                    ))}
                    {filteredStudents.length === 0 && (
                      <tr><td colSpan={4} style={{ padding: 20, textAlign: "center", color: "#8a806a", ...serif }}>No matches.</td></tr>
                    )}
                  </tbody>
                </table>
              </div>

              <div style={{ marginTop: 18 }}>
                <NavButtons back={() => setStep("configure")} next={() => setStep("export")} nextLabel="Review & export" />
              </div>
            </div>
          )}

          {/* ---------- EXPORT ---------- */}
          {step === "export" && (
            <div>
              <h2 style={{ ...serif, fontSize: 18, color: LEDGER_DARK, marginBottom: 6 }}>Export today's register</h2>
              <p style={{ ...serif, fontSize: 13.5, color: "#5b5343", marginBottom: 16 }}>
                Your file will keep every original column and add <b>{todayLabel()}</b>
                {pollEnabled && pollTitle.trim() ? <> and <b>{pollTitle}</b></> : null} at the end.
              </p>

              <div style={{ display: "flex", gap: 20, marginBottom: 10, flexWrap: "wrap" }}>
                <StatBlock label="Students" value={students.length} />
                <StatBlock label="Present" value={presentCount} color={PRESENT} />
                <StatBlock label="Absent" value={absentCount} color={STAMP} />
              </div>
              <div style={{ fontSize: 11.5, color: "#8a806a", marginBottom: 20, ...serif }}>
                Students not explicitly marked present are counted and exported as absent.
              </div>

              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button onClick={() => handleExport(false)} style={primaryBtn(LEDGER)}>
                  <Download size={16} style={{ marginRight: 6 }} /> Download Excel file
                </button>
                {session?.email && (
                  <button onClick={() => handleExport(true)} style={primaryBtn(STAMP)}>
                    <Save size={16} style={{ marginRight: 6 }} /> Download &amp; save to my email
                  </button>
                )}
              </div>

              <div style={{ marginTop: 18 }}>
                <button onClick={() => setStep("take")} style={{ ...linkBtn }}><ArrowLeft size={14} style={{ marginRight: 4 }} />Back to roll</button>
              </div>
            </div>
          )}

          {/* ---------- DONE ---------- */}
          {step === "done" && (
            <div style={{ textAlign: "center", padding: "20px 0" }}>
              <div style={{
                width: 64, height: 64, borderRadius: "50%", border: `3px solid ${STAMP}`,
                color: STAMP, display: "flex", alignItems: "center", justifyContent: "center",
                margin: "0 auto 16px", transform: "rotate(-8deg)", ...serif, fontWeight: 700, fontSize: 12
              }}>
                {todayLabel()}
              </div>
              <h2 style={{ ...serif, fontSize: 19, color: LEDGER_DARK, marginBottom: 6 }}>Register exported</h2>
              <p style={{ ...serif, fontSize: 13.5, color: "#5b5343", marginBottom: 22 }}>
                The file downloaded to your device{session?.email ? " and a copy was saved to your email tag" : ""}.
              </p>
              <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
                <button onClick={startOver} style={primaryBtn(LEDGER)}>
                  <FileSpreadsheet size={16} style={{ marginRight: 6 }} /> Take another register
                </button>
                {session?.email && (
                  <button onClick={() => { setStep("upload"); fetchSessions(session.email); }} style={primaryBtn(STAMP)}>
                    <ClipboardList size={16} style={{ marginRight: 6 }} /> View saved registers
                  </button>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ---------- shared bits ----------
const selectStyle = {
  width: "100%", padding: "8px 10px", borderRadius: 7, border: `1px solid ${LINE}`,
  background: "#fff", fontSize: 13.5, color: INK, marginTop: 4, marginBottom: 4,
};
const thStyle = { padding: "8px 12px", textAlign: "left", fontSize: 12, letterSpacing: "0.04em", textTransform: "uppercase", color: "#5b5343", ...serif, borderBottom: `1px solid ${LINE}` };
const tdStyle = { padding: "8px 12px", fontSize: 13.5, color: INK };

function pillBtn(color) {
  return {
    padding: "8px 12px", borderRadius: 20, border: `1px solid ${color}`, background: "#fff",
    color, fontSize: 12.5, cursor: "pointer", ...serif, fontWeight: 700, display: "flex", alignItems: "center",
  };
}
function markBtn(active, color, mobile) {
  return {
    width: mobile ? 38 : 30, height: mobile ? 38 : 30, borderRadius: 6, border: `1.5px solid ${color}`,
    background: active ? color : "#fff", color: active ? "#fff" : color,
    marginRight: 6, cursor: "pointer", display: "inline-flex", alignItems: "center", justifyContent: "center",
    flexShrink: 0,
  };
}
function primaryBtn(bg) {
  return {
    padding: "11px 18px", borderRadius: 8, border: "none", background: bg, color: PAPER,
    fontSize: 14, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", ...serif,
  };
}
const linkBtn = { border: "none", background: "none", color: "#5b5343", cursor: "pointer", fontSize: 13, display: "flex", alignItems: "center", ...serif, padding: 0 };

function StatBlock({ label, value, color }) {
  return (
    <div>
      <div style={{ fontSize: 22, fontWeight: 700, color: color || INK, ...serif }}>{value}</div>
      <div style={{ fontSize: 11.5, color: "#8a806a", textTransform: "uppercase", letterSpacing: "0.05em" }}>{label}</div>
    </div>
  );
}

function NavButtons({ back, next, nextLabel }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", marginTop: 8 }}>
      <button onClick={back} style={{ ...linkBtn }}><ArrowLeft size={14} style={{ marginRight: 4 }} />Back</button>
      <button onClick={next} style={primaryBtn(LEDGER)}>{nextLabel} <ArrowRight size={16} style={{ marginLeft: 6 }} /></button>
    </div>
  );
}
