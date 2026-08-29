import React, { useState, useEffect, useRef, useCallback, useMemo } from "react";
import * as mammoth from "mammoth";
import {
  BookOpen, BookMarked, Library, Upload, FileText, Download, Trash2, Plus,
  ArrowLeft, X, Check, AlertCircle, Loader2, Pencil, Save, RotateCcw,
  ChevronLeft, ChevronRight, LayoutDashboard, BookCopy, Clock, CheckCircle2
} from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, ResponsiveContainer, Tooltip, Cell } from "recharts";

/* ============================== PALETTE ============================== */
const INK = "#241A15";        // near-black leather ink
const LEATHER = "#5B2A2E";    // deep burgundy primary
const LEATHER_DARK = "#3E1D20";
const GOLD = "#AD8A56";       // aged gold foil accent
const PAPER = "#F6F1E6";      // warm parchment background
const CARD = "#FFFDF8";
const MUTED = "#8A7B68";
const GREEN = "#3F7A57";      // reading progress / edited-live green
const RED = "#A23B3B";
const BORDER = "#E4DAC7";

const SPINES = ["#5B2A2E", "#2F4858", "#4C6444", "#7A5230", "#5C4A72"];
function spineColor(seed) {
  let h = 0;
  for (let i = 0; i < seed.length; i++) h = (h * 31 + seed.charCodeAt(i)) >>> 0;
  return SPINES[h % SPINES.length];
}

/* ============================== HELPERS ============================== */
const MAX_FILE_BYTES = 3 * 1024 * 1024;

function fmtBytes(b) {
  if (b < 1024) return b + " B";
  if (b < 1024 * 1024) return (b / 1024).toFixed(0) + " KB";
  return (b / (1024 * 1024)).toFixed(1) + " MB";
}
function fmtStamp(iso) {
  if (!iso) return { date: "—", time: "" };
  const d = new Date(iso);
  return {
    date: d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" }),
    time: d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" }),
  };
}
function extOf(name) {
  const m = /\.([a-zA-Z0-9]+)$/.exec(name || "");
  return m ? m[1].toLowerCase() : "";
}
function cleanBaseName(filename) {
  return filename.replace(/\.[^/.]+$/, "").replace(/[_-]+/g, " ").trim();
}
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result.split(",")[1]);
    r.onerror = () => reject(new Error("Could not read file"));
    r.readAsDataURL(file);
  });
}
function base64ToArrayBuffer(base64) {
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}
function base64ToBlobUrl(base64, mime) {
  const bin = atob(base64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const blob = new Blob([bytes], { type: mime });
  return URL.createObjectURL(blob);
}
function mimeFor(ext) {
  if (ext === "pdf") return "application/pdf";
  if (ext === "doc") return "application/msword";
  return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
}
function progressOf(book) {
  const total = book.totalPages || 1;
  return Math.min(100, Math.round(((book.lastReadPage || 0) / total) * 100));
}
function statusOf(book) {
  const p = progressOf(book);
  if (p <= 0) return { label: "Not started", color: MUTED };
  if (p >= 100) return { label: "Finished", color: GREEN };
  return { label: "Reading", color: GOLD };
}

const DEFAULT_FOLDERS = [
  { id: "lean-six-sigma", name: "Lean Six Sigma", createdAt: new Date().toISOString(), isDefault: true },
  { id: "pmp", name: "PMP", createdAt: new Date().toISOString(), isDefault: true },
];

/* ============================== ROOT APP ============================== */
export default function FolioApp() {
  const [ready, setReady] = useState(false);
  const [folders, setFolders] = useState([]);
  const [booksByFolder, setBooksByFolder] = useState({});
  const [view, setView] = useState("dashboard"); // dashboard | shelf
  const [activeFolderId, setActiveFolderId] = useState(null);

  const [showNewFolder, setShowNewFolder] = useState(false);
  const [newFolderName, setNewFolderName] = useState("");

  const [pendingFile, setPendingFile] = useState(null);
  const [pendingName, setPendingName] = useState("");
  const [pendingPages, setPendingPages] = useState("");
  const [uploadError, setUploadError] = useState("");
  const [saving, setSaving] = useState(false);
  const fileInputRef = useRef(null);

  const [confirmDelete, setConfirmDelete] = useState(null);
  const [toast, setToast] = useState("");
  const [readingTarget, setReadingTarget] = useState(null); // {folderId, book}

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("folders-index", true);
        let list = DEFAULT_FOLDERS;
        if (res && res.value) list = JSON.parse(res.value);
        else await window.storage.set("folders-index", JSON.stringify(DEFAULT_FOLDERS), true);
        setFolders(list);
        const entries = await Promise.all(list.map(async (f) => {
          try {
            const r = await window.storage.get(`books:${f.id}`, true);
            return [f.id, r && r.value ? JSON.parse(r.value) : []];
          } catch (e) { return [f.id, []]; }
        }));
        setBooksByFolder(Object.fromEntries(entries));
      } catch (e) {
        setFolders(DEFAULT_FOLDERS);
      }
      setReady(true);
    })();
  }, []);

  const showToast = (msg) => { setToast(msg); setTimeout(() => setToast(""), 2800); };

  const saveFolders = async (next) => {
    setFolders(next);
    try { await window.storage.set("folders-index", JSON.stringify(next), true); }
    catch (e) { showToast("Could not save shelf list."); }
  };

  const saveBooks = async (folderId, list) => {
    setBooksByFolder((prev) => ({ ...prev, [folderId]: list }));
    try {
      const result = await window.storage.set(`books:${folderId}`, JSON.stringify(list), true);
      if (!result) throw new Error("no result");
    } catch (e) { showToast("Storage error — this may not have saved."); }
  };

  const createFolder = async () => {
    const name = newFolderName.trim();
    if (!name) return;
    const id = name.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") + "-" + Date.now().toString(36);
    const next = [...folders, { id, name, createdAt: new Date().toISOString(), isDefault: false }];
    await saveFolders(next);
    setBooksByFolder((prev) => ({ ...prev, [id]: [] }));
    setShowNewFolder(false); setNewFolderName("");
    showToast(`Shelf "${name}" created`);
  };

  const deleteFolder = async (id) => {
    await saveFolders(folders.filter((f) => f.id !== id));
    try { await window.storage.delete(`books:${id}`, true); } catch (e) {}
    setBooksByFolder((prev) => { const c = { ...prev }; delete c[id]; return c; });
    if (activeFolderId === id) { setView("dashboard"); setActiveFolderId(null); }
    setConfirmDelete(null);
    showToast("Shelf deleted");
  };

  const openFolder = (id) => { setActiveFolderId(id); setView("shelf"); };

  const onPickFile = () => fileInputRef.current && fileInputRef.current.click();
  const onFileSelected = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setUploadError("");
    const ext = extOf(file.name);
    if (!["pdf", "doc", "docx"].includes(ext)) { setUploadError("Only PDF and Word (.doc, .docx) files are supported."); return; }
    if (file.size > MAX_FILE_BYTES) { setUploadError(`File is too large (${fmtBytes(file.size)}). Keep files under ${fmtBytes(MAX_FILE_BYTES)}.`); return; }
    try {
      const base64 = await fileToBase64(file);
      setPendingFile({ file, base64, ext, size: file.size });
      setPendingName(cleanBaseName(file.name));
      setPendingPages("");
    } catch (err) { setUploadError("Could not read this file. Please try again."); }
  };
  const cancelUpload = () => { setPendingFile(null); setPendingName(""); setPendingPages(""); setUploadError(""); };

  const confirmUploadSave = async () => {
    const pages = parseInt(pendingPages, 10);
    if (!pendingFile || !pendingName.trim()) return;
    if (!pages || pages < 1) { setUploadError("Please enter the total number of pages (a whole number, 1 or more)."); return; }
    setSaving(true);
    const existing = booksByFolder[activeFolderId] || [];
    const combinedSize = existing.reduce((s, c) => s + (c.size || 0), 0) + pendingFile.size;
    if (combinedSize > 4.5 * 1024 * 1024) {
      setUploadError("This shelf is close to its storage limit. Remove an older file first, or use a smaller file.");
      setSaving(false); return;
    }
    const entry = {
      id: "b_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      name: pendingName.trim(),
      fileName: pendingFile.file.name,
      fileType: pendingFile.ext,
      size: pendingFile.size,
      uploadedAt: new Date().toISOString(),
      dataUrl: pendingFile.base64,
      totalPages: pages,
      lastReadPage: 0,
      lastReadAt: null,
      editedHtml: null,
      editedAt: null,
    };
    await saveBooks(activeFolderId, [entry, ...existing]);
    setSaving(false); cancelUpload();
    showToast(`"${entry.name}" added to shelf`);
  };

  const deleteBook = async (folderId, bookId) => {
    const next = (booksByFolder[folderId] || []).filter((c) => c.id !== bookId);
    await saveBooks(folderId, next);
    setConfirmDelete(null);
    showToast("Book removed");
  };

  const updateBook = async (folderId, bookId, patch) => {
    const list = booksByFolder[folderId] || [];
    const next = list.map((b) => (b.id === bookId ? { ...b, ...patch } : b));
    await saveBooks(folderId, next);
  };

  const activeFolder = folders.find((f) => f.id === activeFolderId);
  const activeBooks = booksByFolder[activeFolderId] || [];

  const allBooksFlat = useMemo(() => {
    const out = [];
    folders.forEach((f) => (booksByFolder[f.id] || []).forEach((b) => out.push({ ...b, folderId: f.id, folderName: f.name })));
    return out;
  }, [folders, booksByFolder]);

  return (
    <div className="min-h-screen w-full" style={{ background: PAPER }}>
      <style>{`
        @keyframes riseIn { from { opacity:0; transform: translateY(14px); } to { opacity:1; transform: translateY(0); } }
        @keyframes fadeIn { from { opacity:0; } to { opacity:1; } }
        @keyframes popIn { from { opacity:0; transform: scale(.94); } to { opacity:1; transform: scale(1); } }
        .rise { animation: riseIn .5s ease both; }
        .fade { animation: fadeIn .25s ease both; }
        .pop { animation: popIn .22s cubic-bezier(.2,.9,.3,1.2) both; }
        .book-card { transition: transform .28s ease, box-shadow .28s ease; }
        .book-card:hover { transform: translateY(-6px); box-shadow: 0 18px 30px -14px rgba(36,26,21,0.35); }
        .shelf-card { transition: transform .25s ease, box-shadow .25s ease; }
        .shelf-card:hover { transform: translateY(-4px); box-shadow: 0 14px 26px -12px rgba(36,26,21,0.3); }
        .btn-lift { transition: transform .15s ease, filter .15s ease; }
        .btn-lift:hover { transform: translateY(-1px); filter: brightness(1.06); }
        .progress-fill { transition: width .8s cubic-bezier(.2,.8,.3,1); }
        .editing-live[data-editable="true"] { color: ${GREEN} !important; }
        ::-webkit-scrollbar { width: 10px; height:10px; }
        ::-webkit-scrollbar-thumb { background: ${BORDER}; border-radius: 8px; }
      `}</style>

      <input ref={fileInputRef} type="file" accept=".pdf,.doc,.docx" className="hidden" onChange={onFileSelected} />

      <TopBar view={view} onHome={() => { setView("dashboard"); setActiveFolderId(null); }} />

      <main className="max-w-6xl mx-auto px-6 py-8">
        {!ready ? (
          <div className="flex items-center gap-2 py-24 justify-center" style={{ color: MUTED }}>
            <Loader2 className="animate-spin" size={20} /> Opening your library…
          </div>
        ) : view === "dashboard" ? (
          <Dashboard
            folders={folders}
            booksByFolder={booksByFolder}
            allBooksFlat={allBooksFlat}
            onOpenFolder={openFolder}
            onNewFolder={() => setShowNewFolder(true)}
            onDeleteFolder={(f) => setConfirmDelete({ type: "folder", id: f.id, name: f.name })}
            onContinueReading={(b) => setReadingTarget({ folderId: b.folderId, book: b })}
          />
        ) : (
          <ShelfView
            folder={activeFolder}
            books={activeBooks}
            onBack={() => { setView("dashboard"); setActiveFolderId(null); }}
            onUploadClick={onPickFile}
            onRead={(book) => setReadingTarget({ folderId: activeFolderId, book })}
            onDeleteBook={(b) => setConfirmDelete({ type: "book", id: b.id, name: b.name, folderId: activeFolderId })}
          />
        )}
      </main>

      {showNewFolder && (
        <Modal onClose={() => { setShowNewFolder(false); setNewFolderName(""); }}>
          <h3 className="font-serif text-xl font-bold mb-4" style={{ color: LEATHER_DARK }}>New shelf</h3>
          <label className="text-sm block mb-1" style={{ color: MUTED }}>Shelf name</label>
          <input autoFocus value={newFolderName} onChange={(e) => setNewFolderName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && createFolder()} placeholder="e.g. Business Analysis"
            className="w-full border rounded-md px-3 py-2 mb-5 focus:outline-none focus:ring-2"
            style={{ borderColor: BORDER }} />
          <div className="flex justify-end gap-3">
            <button onClick={() => { setShowNewFolder(false); setNewFolderName(""); }} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
            <button onClick={createFolder} disabled={!newFolderName.trim()} className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white disabled:opacity-40" style={{ background: LEATHER }}>Create shelf</button>
          </div>
        </Modal>
      )}

      {pendingFile && (
        <Modal onClose={cancelUpload}>
          <h3 className="font-serif text-xl font-bold mb-1" style={{ color: LEATHER_DARK }}>Add this book</h3>
          <p className="text-sm mb-4" style={{ color: MUTED }}>File: <span className="font-mono">{pendingFile.file.name}</span> ({fmtBytes(pendingFile.size)})</p>
          <label className="text-sm block mb-1" style={{ color: MUTED }}>Book / document name</label>
          <input autoFocus value={pendingName} onChange={(e) => setPendingName(e.target.value)}
            className="w-full border rounded-md px-3 py-2 mb-4 focus:outline-none focus:ring-2" style={{ borderColor: BORDER }} />
          <label className="text-sm block mb-1" style={{ color: MUTED }}>Total pages</label>
          <input type="number" min="1" value={pendingPages} onChange={(e) => setPendingPages(e.target.value)}
            placeholder="e.g. 120" onKeyDown={(e) => e.key === "Enter" && confirmUploadSave()}
            className="w-full border rounded-md px-3 py-2 mb-2 focus:outline-none focus:ring-2" style={{ borderColor: BORDER }} />
          <p className="text-xs mb-3" style={{ color: MUTED }}>Used to calculate the reading-progress bar shown on the shelf and dashboard.</p>
          {uploadError && <p className="text-sm flex items-center gap-1.5 mb-2" style={{ color: RED }}><AlertCircle size={14} /> {uploadError}</p>}
          <div className="flex justify-end gap-3 mt-3">
            <button onClick={cancelUpload} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
            <button onClick={confirmUploadSave} disabled={!pendingName.trim() || saving}
              className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white disabled:opacity-40 flex items-center gap-2" style={{ background: LEATHER }}>
              {saving ? <Loader2 className="animate-spin" size={14} /> : <Check size={14} />} Save to shelf
            </button>
          </div>
        </Modal>
      )}

      {confirmDelete && (
        <Modal onClose={() => setConfirmDelete(null)}>
          <h3 className="font-serif text-xl font-bold mb-2" style={{ color: LEATHER_DARK }}>Delete {confirmDelete.type === "folder" ? "shelf" : "book"}?</h3>
          <p className="text-sm mb-6" style={{ color: MUTED }}>
            "{confirmDelete.name}" will be permanently removed{confirmDelete.type === "folder" ? ", along with every book on it" : ""}. This cannot be undone.
          </p>
          <div className="flex justify-end gap-3">
            <button onClick={() => setConfirmDelete(null)} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
            <button onClick={() => confirmDelete.type === "folder" ? deleteFolder(confirmDelete.id) : deleteBook(confirmDelete.folderId, confirmDelete.id)}
              className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white flex items-center gap-2" style={{ background: RED }}>
              <Trash2 size={14} /> Delete
            </button>
          </div>
        </Modal>
      )}

      {readingTarget && (
        <BookReader
          folderId={readingTarget.folderId}
          book={readingTarget.book}
          onClose={() => setReadingTarget(null)}
          onCommitProgress={(page) => updateBook(readingTarget.folderId, readingTarget.book.id, { lastReadPage: page, lastReadAt: new Date().toISOString() })}
          onSaveEdit={(html) => updateBook(readingTarget.folderId, readingTarget.book.id, { editedHtml: html, editedAt: new Date().toISOString() })}
        />
      )}

      {toast && (
        <div className="fixed bottom-6 left-1/2 -translate-x-1/2 px-4 py-2.5 rounded-lg text-sm text-white shadow-lg flex items-center gap-2 z-50 pop" style={{ background: LEATHER_DARK }}>
          <Check size={14} style={{ color: GOLD }} /> {toast}
        </div>
      )}
    </div>
  );
}

/* ============================== TOP BAR ============================== */
function TopBar({ view, onHome }) {
  return (
    <header style={{ background: LEATHER_DARK, borderBottom: `3px solid ${GOLD}` }}>
      <div className="max-w-6xl mx-auto px-6 py-5 flex items-center justify-between">
        <button onClick={onHome} className="flex items-center gap-3">
          <div className="h-9 w-9 rounded-md flex items-center justify-center" style={{ background: GOLD }}>
            <BookOpen size={19} style={{ color: LEATHER_DARK }} />
          </div>
          <div className="text-left">
            <h1 className="font-serif text-2xl font-bold text-white leading-none">Folio</h1>
            <p className="text-[11px] tracking-widest uppercase font-mono" style={{ color: "#D8C7A8" }}>Read · Track · Remember</p>
          </div>
        </button>
        <button onClick={onHome} className="flex items-center gap-1.5 text-sm px-3 py-1.5 rounded-md"
          style={{ color: "#F1E6D2", background: view === "dashboard" ? "rgba(255,255,255,0.08)" : "transparent", fontWeight: view === "dashboard" ? 600 : 400 }}>
          <LayoutDashboard size={15} /> Dashboard
        </button>
      </div>
    </header>
  );
}

/* ============================== MODAL ============================== */
function Modal({ children, onClose }) {
  return (
    <div className="fixed inset-0 bg-black/45 flex items-center justify-center z-50 px-4 fade" onClick={onClose}>
      <div className="rounded-xl shadow-2xl w-full max-w-md p-6 pop" style={{ background: CARD }} onClick={(e) => e.stopPropagation()}>
        {children}
      </div>
    </div>
  );
}

/* ============================== CIRCLE PROGRESS ============================== */
function CircleProgress({ percent, size = 54, stroke = 6, color = GREEN }) {
  const r = (size - stroke) / 2;
  const c = 2 * Math.PI * r;
  const off = c - (percent / 100) * c;
  return (
    <svg width={size} height={size} className="shrink-0">
      <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke={BORDER} strokeWidth={stroke} />
      <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke={color} strokeWidth={stroke}
        strokeDasharray={c} strokeDashoffset={off} strokeLinecap="round"
        transform={`rotate(-90 ${size / 2} ${size / 2})`} style={{ transition: "stroke-dashoffset .8s ease" }} />
      <text x="50%" y="50%" textAnchor="middle" dy="0.35em" fontSize={size * 0.26} fontWeight="700" fill={INK}>{percent}%</text>
    </svg>
  );
}

/* ============================== DASHBOARD ============================== */
function Dashboard({ folders, booksByFolder, allBooksFlat, onOpenFolder, onNewFolder, onDeleteFolder, onContinueReading }) {
  const totalBooks = allBooksFlat.length;
  const totalPagesAll = allBooksFlat.reduce((s, b) => s + (b.totalPages || 0), 0);
  const readPagesAll = allBooksFlat.reduce((s, b) => s + Math.min(b.lastReadPage || 0, b.totalPages || 0), 0);
  const overallPct = totalPagesAll > 0 ? Math.round((readPagesAll / totalPagesAll) * 100) : 0;
  const reading = allBooksFlat.filter((b) => { const p = progressOf(b); return p > 0 && p < 100; }).length;
  const finished = allBooksFlat.filter((b) => progressOf(b) >= 100).length;

  const chartData = folders.map((f) => {
    const list = booksByFolder[f.id] || [];
    const tp = list.reduce((s, b) => s + (b.totalPages || 0), 0);
    const rp = list.reduce((s, b) => s + Math.min(b.lastReadPage || 0, b.totalPages || 0), 0);
    return { name: f.name, percent: tp > 0 ? Math.round((rp / tp) * 100) : 0, count: list.length };
  });

  const continueReading = allBooksFlat
    .filter((b) => { const p = progressOf(b); return p > 0 && p < 100; })
    .sort((a, b) => new Date(b.lastReadAt || 0) - new Date(a.lastReadAt || 0))
    .slice(0, 4);

  return (
    <div>
      <div className="mb-8 rise">
        <p className="text-xs uppercase tracking-widest font-mono" style={{ color: GOLD }}>Your library at a glance</p>
        <h2 className="font-serif text-3xl font-bold" style={{ color: LEATHER_DARK }}>Dashboard</h2>
      </div>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-10">
        <StatCard icon={<BookCopy size={18} />} label="Total books" value={totalBooks} delay={0} />
        <StatCard icon={<Library size={18} />} label="Shelves" value={folders.length} delay={60} />
        <StatCard icon={<Clock size={18} />} label="Currently reading" value={reading} delay={120} />
        <StatCard icon={<CheckCircle2 size={18} />} label="Finished" value={finished} delay={180} />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-5 gap-6 mb-10">
        <div className="lg:col-span-2 rounded-xl p-6 rise" style={{ background: CARD, border: `1px solid ${BORDER}`, animationDelay: "80ms" }}>
          <p className="text-sm mb-1" style={{ color: MUTED }}>Overall completion</p>
          <div className="flex items-center gap-5 mt-3">
            <CircleProgress percent={overallPct} size={92} stroke={9} color={GREEN} />
            <div>
              <p className="font-serif text-2xl font-bold" style={{ color: LEATHER_DARK }}>{readPagesAll.toLocaleString()} <span className="text-base font-sans font-normal" style={{ color: MUTED }}>/ {totalPagesAll.toLocaleString()} pages</span></p>
              <p className="text-sm mt-1" style={{ color: MUTED }}>across every shelf</p>
            </div>
          </div>
        </div>

        <div className="lg:col-span-3 rounded-xl p-6 rise" style={{ background: CARD, border: `1px solid ${BORDER}`, animationDelay: "140ms" }}>
          <p className="text-sm mb-3" style={{ color: MUTED }}>Completion by shelf</p>
          {chartData.length === 0 ? (
            <p className="text-sm py-8 text-center" style={{ color: MUTED }}>No shelves yet.</p>
          ) : (
            <ResponsiveContainer width="100%" height={Math.max(120, chartData.length * 44)}>
              <BarChart data={chartData} layout="vertical" margin={{ left: 4, right: 24 }}>
                <XAxis type="number" domain={[0, 100]} hide />
                <YAxis dataKey="name" type="category" width={130} tick={{ fontSize: 12, fill: INK }} axisLine={false} tickLine={false} />
                <Tooltip formatter={(v) => [`${v}%`, "Complete"]} contentStyle={{ borderRadius: 8, borderColor: BORDER }} />
                <Bar dataKey="percent" radius={[0, 6, 6, 0]} barSize={16}>
                  {chartData.map((d, i) => <Cell key={i} fill={spineColor(folders[i]?.id || String(i))} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>
      </div>

      {continueReading.length > 0 && (
        <div className="mb-10">
          <h3 className="font-serif text-lg font-bold mb-4" style={{ color: LEATHER_DARK }}>Continue reading</h3>
          <div className="flex gap-4 overflow-x-auto pb-2">
            {continueReading.map((b, i) => {
              const pct = progressOf(b);
              return (
                <button key={b.id} onClick={() => onContinueReading(b)}
                  className="book-card text-left shrink-0 w-52 rounded-lg overflow-hidden rise" style={{ background: CARD, border: `1px solid ${BORDER}`, animationDelay: `${i * 70}ms` }}>
                  <div className="h-3" style={{ background: spineColor(b.folderId) }} />
                  <div className="p-4">
                    <p className="text-xs font-mono mb-1" style={{ color: MUTED }}>{b.folderName}</p>
                    <p className="font-serif font-bold text-sm mb-3" style={{ color: LEATHER_DARK }}>{b.name}</p>
                    <div className="h-1.5 rounded-full overflow-hidden mb-1" style={{ background: BORDER }}>
                      <div className="h-full progress-fill rounded-full" style={{ width: `${pct}%`, background: GOLD }} />
                    </div>
                    <p className="text-xs" style={{ color: MUTED }}>{pct}% · page {b.lastReadPage} of {b.totalPages}</p>
                  </div>
                </button>
              );
            })}
          </div>
        </div>
      )}

      <div className="flex items-end justify-between mb-5">
        <h3 className="font-serif text-lg font-bold" style={{ color: LEATHER_DARK }}>Your shelves</h3>
        <button onClick={onNewFolder} className="btn-lift flex items-center gap-1.5 px-4 py-2 rounded-md text-sm font-medium text-white" style={{ background: LEATHER }}>
          <Plus size={16} /> New shelf
        </button>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-5">
        {folders.map((f, i) => {
          const list = booksByFolder[f.id] || [];
          const tp = list.reduce((s, b) => s + (b.totalPages || 0), 0);
          const rp = list.reduce((s, b) => s + Math.min(b.lastReadPage || 0, b.totalPages || 0), 0);
          const pct = tp > 0 ? Math.round((rp / tp) * 100) : 0;
          return (
            <ShelfCard key={f.id} folder={f} count={list.length} percent={pct} delay={i * 70}
              onOpen={() => onOpenFolder(f.id)} onDelete={() => onDeleteFolder(f)} />
          );
        })}
      </div>
    </div>
  );
}

function StatCard({ icon, label, value, delay }) {
  return (
    <div className="rounded-xl p-5 rise" style={{ background: CARD, border: `1px solid ${BORDER}`, animationDelay: `${delay}ms` }}>
      <div className="flex items-center gap-2 mb-2" style={{ color: GOLD }}>{icon}<span className="text-xs uppercase tracking-wide font-mono" style={{ color: MUTED }}>{label}</span></div>
      <p className="font-serif text-3xl font-bold" style={{ color: LEATHER_DARK }}>{value}</p>
    </div>
  );
}

function ShelfCard({ folder, count, percent, delay, onOpen, onDelete }) {
  const color = spineColor(folder.id);
  const [hover, setHover] = useState(false);
  return (
    <div className="shelf-card rise cursor-pointer" style={{ animationDelay: `${delay}ms` }}
      onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)} onClick={onOpen}>
      <div className="rounded-xl overflow-hidden" style={{ background: CARD, border: `1px solid ${BORDER}` }}>
        <div className="h-2.5" style={{ background: color }} />
        <div className="p-5 flex items-start gap-4">
          <CircleProgress percent={percent} size={54} stroke={6} color={color} />
          <div className="flex-1 min-w-0">
            <div className="flex items-start justify-between gap-2">
              <h4 className="font-serif font-bold text-base" style={{ color: LEATHER_DARK }}>{folder.name}</h4>
              {!folder.isDefault && hover && (
                <button onClick={(e) => { e.stopPropagation(); onDelete(); }} className="p-1 -m-1 shrink-0" style={{ color: MUTED }} title="Delete shelf">
                  <Trash2 size={14} />
                </button>
              )}
            </div>
            <p className="text-xs font-mono mt-1" style={{ color: MUTED }}>{count} book{count !== 1 ? "s" : ""}</p>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================== SHELF VIEW ============================== */
function ShelfView({ folder, books, onBack, onUploadClick, onRead, onDeleteBook }) {
  if (!folder) return null;
  const color = spineColor(folder.id);
  return (
    <div className="fade">
      <button onClick={onBack} className="flex items-center gap-1.5 text-sm mb-5" style={{ color: MUTED }}><ArrowLeft size={15} /> Dashboard</button>
      <div className="flex items-end justify-between mb-6">
        <div className="flex items-center gap-3">
          <div className="h-11 w-11 rounded-lg flex items-center justify-center" style={{ background: color }}>
            <Library size={20} color="white" />
          </div>
          <div>
            <h2 className="font-serif text-2xl font-bold" style={{ color: LEATHER_DARK }}>{folder.name}</h2>
            <p className="text-sm" style={{ color: MUTED }}>{books.length} book{books.length !== 1 ? "s" : ""} on this shelf</p>
          </div>
        </div>
        <button onClick={onUploadClick} className="btn-lift flex items-center gap-1.5 px-4 py-2 rounded-md text-sm font-medium text-white shrink-0" style={{ background: LEATHER }}>
          <Upload size={16} /> Upload Book / Document
        </button>
      </div>

      {books.length === 0 ? (
        <div className="text-center py-16 border-2 border-dashed rounded-lg" style={{ borderColor: BORDER }}>
          <BookMarked size={30} className="mx-auto mb-3" style={{ color: BORDER }} />
          <p className="mb-4" style={{ color: MUTED }}>This shelf is empty.</p>
          <button onClick={onUploadClick} className="text-sm font-medium underline" style={{ color: LEATHER }}>Upload the first PDF or Word document</button>
        </div>
      ) : (
        <div className="grid grid-cols-1 gap-4">
          {books.map((b, i) => <BookRow key={b.id} book={b} color={color} delay={i * 50} onRead={() => onRead(b)} onDelete={() => onDeleteBook(b)} />)}
        </div>
      )}
    </div>
  );
}

function BookRow({ book, color, delay, onRead, onDelete }) {
  const pct = progressOf(book);
  const status = statusOf(book);
  const stamp = fmtStamp(book.uploadedAt);
  return (
    <div className="rise rounded-xl p-5 flex flex-col sm:flex-row sm:items-center gap-4" style={{ background: CARD, border: `1px solid ${BORDER}`, animationDelay: `${delay}ms` }}>
      <div className="h-14 w-11 rounded-sm shrink-0 flex items-center justify-center" style={{ background: color }}>
        <FileText size={18} color="white" />
      </div>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2 flex-wrap">
          <p className="font-serif font-bold" style={{ color: LEATHER_DARK }}>{book.name}</p>
          <span className="text-xs px-2 py-0.5 rounded-full font-medium" style={{ background: status.color + "22", color: status.color }}>{status.label}</span>
        </div>
        <p className="text-xs font-mono mt-0.5" style={{ color: MUTED }}>{book.fileName} · {fmtBytes(book.size)} · added {stamp.date}</p>
        <div className="flex items-center gap-3 mt-2.5">
          <div className="h-2 rounded-full overflow-hidden flex-1 max-w-xs" style={{ background: BORDER }}>
            <div className="h-full progress-fill rounded-full" style={{ width: `${pct}%`, background: pct >= 100 ? GREEN : GOLD }} />
          </div>
          <span className="text-xs font-mono w-28 shrink-0" style={{ color: MUTED }}>{pct}% · pg {book.lastReadPage}/{book.totalPages}</span>
        </div>
      </div>
      <div className="flex items-center gap-1.5 shrink-0 self-end sm:self-center">
        <button onClick={onRead} className="btn-lift flex items-center gap-1.5 px-3 py-2 rounded-md text-sm font-medium text-white" style={{ background: color }}>
          <BookOpen size={15} /> Read
        </button>
        <button onClick={() => downloadOriginal(book)} title="Download" className="p-2 rounded-md hover:bg-black/5" style={{ color: MUTED }}><Download size={17} /></button>
        <button onClick={onDelete} title="Delete" className="p-2 rounded-md hover:bg-black/5" style={{ color: MUTED }}><Trash2 size={17} /></button>
      </div>
    </div>
  );
}

function downloadOriginal(book) {
  try {
    const url = base64ToBlobUrl(book.dataUrl, mimeFor(book.fileType));
    const a = document.createElement("a");
    a.href = url; a.download = `${book.name}.${book.fileType}`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  } catch (e) { console.error(e); }
}

/* ============================== BOOK READER ============================== */
function BookReader({ folderId, book, onClose, onCommitProgress, onSaveEdit }) {
  const isPdf = book.fileType === "pdf";
  const isDocx = book.fileType === "docx";
  const isLegacyDoc = book.fileType === "doc";

  const [pdfUrl, setPdfUrl] = useState(null);
  const [pdfPage, setPdfPage] = useState(book.lastReadPage || 1);

  const [html, setHtml] = useState(null);
  const [htmlError, setHtmlError] = useState("");
  const [editMode, setEditMode] = useState(false);
  const [isDirty, setIsDirty] = useState(false);
  const [scrollPage, setScrollPage] = useState(book.lastReadPage || 1);
  const contentRef = useRef(null);
  const scrollRef = useRef(null);

  const [stage, setStage] = useState("reading"); // reading | discardCheck | confirmPage | manualPage
  const [manualPage, setManualPage] = useState("");

  useEffect(() => {
    if (isPdf) {
      const url = base64ToBlobUrl(book.dataUrl, "application/pdf");
      setPdfUrl(url);
      return () => URL.revokeObjectURL(url);
    }
  }, [isPdf, book.dataUrl]);

  useEffect(() => {
    if (isDocx) {
      (async () => {
        try {
          const initialHtml = book.editedHtml || null;
          if (initialHtml) { setHtml(initialHtml); return; }
          const buf = base64ToArrayBuffer(book.dataUrl);
          const result = await mammoth.convertToHtml({ arrayBuffer: buf });
          setHtml(result.value || "<p><em>(Empty document)</em></p>");
        } catch (e) {
          setHtmlError("This document could not be previewed. You can still download it.");
        }
      })();
    }
  }, [isDocx, book.dataUrl, book.editedHtml]);

  useEffect(() => {
    if (isDocx && html !== null && contentRef.current) {
      contentRef.current.innerHTML = html;
    }
  }, [isDocx, html]);

  const handleScroll = useCallback(() => {
    const el = scrollRef.current;
    if (!el) return;
    const max = el.scrollHeight - el.clientHeight;
    const pct = max > 0 ? el.scrollTop / max : 1;
    const page = Math.min(book.totalPages, Math.max(1, Math.round(pct * book.totalPages) || 1));
    setScrollPage(page);
  }, [book.totalPages]);

  const handleInput = () => {
    if (!isDirty) setIsDirty(true);
  };

  const enterEdit = () => setEditMode(true);
  const saveEdit = () => {
    const newHtml = contentRef.current ? contentRef.current.innerHTML : html;
    setHtml(newHtml);
    onSaveEdit(newHtml);
    setIsDirty(false);
    setEditMode(false);
  };
  const discardEdit = () => {
    if (contentRef.current) contentRef.current.innerHTML = html;
    setIsDirty(false);
    setEditMode(false);
  };

  const currentPage = isPdf ? pdfPage : scrollPage;

  const requestClose = () => {
    if (isDocx && editMode && isDirty) { setStage("discardCheck"); return; }
    setStage("confirmPage");
  };

  const finalizeClose = (page) => { onCommitProgress(page); onClose(); };

  return (
    <div className="fixed inset-0 z-50 flex flex-col fade" style={{ background: "rgba(36,26,21,0.94)" }}>
      <div className="flex items-center justify-between px-5 py-3 shrink-0" style={{ background: LEATHER_DARK, borderBottom: `2px solid ${GOLD}` }}>
        <div className="min-w-0">
          <p className="font-serif font-bold text-white truncate">{book.name}</p>
          <p className="text-xs font-mono" style={{ color: "#D8C7A8" }}>{book.fileName}</p>
        </div>
        <div className="flex items-center gap-2 shrink-0">
          {isDocx && !isLegacyDoc && !editMode && (
            <button onClick={enterEdit} className="btn-lift flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ background: GOLD, color: LEATHER_DARK }}>
              <Pencil size={14} /> Edit
            </button>
          )}
          {isDocx && editMode && (
            <>
              <button onClick={discardEdit} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm" style={{ color: "#F1E6D2" }}><RotateCcw size={14} /> Discard</button>
              <button onClick={saveEdit} disabled={!isDirty} className="btn-lift flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium disabled:opacity-40" style={{ background: GREEN, color: "white" }}>
                <Save size={14} /> Save changes
              </button>
            </>
          )}
          <button onClick={requestClose} className="p-2 rounded-md hover:bg-white/10 text-white"><X size={20} /></button>
        </div>
      </div>

      {editMode && (
        <div className="text-center text-xs py-1.5" style={{ background: "#2F241D", color: GOLD }}>
          Editing enabled — edited text appears in <span style={{ color: GREEN, fontWeight: 700 }}>green</span> until you save.
        </div>
      )}

      <div className="flex-1 overflow-hidden flex flex-col items-center px-4 py-4">
        {isPdf && pdfUrl && (
          <div className="w-full h-full max-w-4xl flex flex-col">
            <div className="flex-1 rounded-lg overflow-hidden shadow-2xl" style={{ background: "white" }}>
              <embed src={pdfUrl} type="application/pdf" className="w-full h-full" />
            </div>
            <div className="flex items-center justify-center gap-4 py-3">
              <button onClick={() => setPdfPage((p) => Math.max(1, p - 1))} className="p-2 rounded-md text-white hover:bg-white/10"><ChevronLeft size={18} /></button>
              <div className="flex items-center gap-2 text-sm" style={{ color: "#F1E6D2" }}>
                Page
                <input type="number" value={pdfPage} min={1} max={book.totalPages}
                  onChange={(e) => setPdfPage(Math.min(book.totalPages, Math.max(1, parseInt(e.target.value, 10) || 1)))}
                  className="w-16 text-center rounded-md px-2 py-1 text-black" />
                of {book.totalPages}
              </div>
              <button onClick={() => setPdfPage((p) => Math.min(book.totalPages, p + 1))} className="p-2 rounded-md text-white hover:bg-white/10"><ChevronRight size={18} /></button>
            </div>
            <p className="text-center text-xs pb-1" style={{ color: "#B7A88C" }}>Use the page controls above to bookmark where you are as you read.</p>
          </div>
        )}

        {isDocx && (
          <div className="w-full h-full max-w-3xl flex flex-col">
            {htmlError ? (
              <div className="flex-1 flex items-center justify-center text-center px-6" style={{ color: "#F1E6D2" }}>
                <div><AlertCircle className="mx-auto mb-3" /><p>{htmlError}</p></div>
              </div>
            ) : html === null ? (
              <div className="flex-1 flex items-center justify-center gap-2" style={{ color: "white" }}><Loader2 className="animate-spin" /> Preparing document…</div>
            ) : (
              <div ref={scrollRef} onScroll={handleScroll} className="flex-1 overflow-y-auto rounded-lg shadow-2xl px-10 py-8" style={{ background: "white" }}>
                <div
                  ref={contentRef}
                  data-editable="true"
                  contentEditable={editMode}
                  suppressContentEditableWarning
                  onInput={handleInput}
                  className={isDirty ? "editing-live" : ""}
                  style={{ outline: "none", lineHeight: 1.7, color: INK, fontFamily: "Georgia, serif" }}
                />
              </div>
            )}
            <p className="text-center text-xs py-2" style={{ color: "#B7A88C" }}>Scroll position estimates your page — page {scrollPage} of {book.totalPages}.</p>
          </div>
        )}

        {isLegacyDoc && (
          <div className="flex-1 flex items-center justify-center text-center px-6" style={{ color: "#F1E6D2" }}>
            <div>
              <FileText className="mx-auto mb-3" size={28} />
              <p className="mb-3">Legacy .doc files can't be previewed in-browser.</p>
              <button onClick={() => downloadOriginal(book)} className="underline text-sm">Download to view</button>
            </div>
          </div>
        )}
      </div>

      {stage === "discardCheck" && (
        <Modal onClose={() => setStage("reading")}>
          <h3 className="font-serif text-xl font-bold mb-2" style={{ color: LEATHER_DARK }}>Unsaved edits</h3>
          <p className="text-sm mb-6" style={{ color: MUTED }}>You have edits that haven't been saved yet. What would you like to do?</p>
          <div className="flex justify-end gap-3 flex-wrap">
            <button onClick={() => setStage("reading")} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Keep editing</button>
            <button onClick={() => { discardEdit(); setStage("confirmPage"); }} className="px-4 py-2 rounded-md text-sm font-medium" style={{ color: RED }}>Discard edits</button>
            <button onClick={() => { saveEdit(); setStage("confirmPage"); }} className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white" style={{ background: GREEN }}>Save & continue</button>
          </div>
        </Modal>
      )}

      {stage === "confirmPage" && (
        <Modal onClose={() => finalizeClose(currentPage)}>
          <h3 className="font-serif text-xl font-bold mb-2" style={{ color: LEATHER_DARK }}>Save your reading progress?</h3>
          <p className="text-sm mb-6" style={{ color: MUTED }}>
            You've read up to page <span className="font-semibold" style={{ color: LEATHER_DARK }}>{currentPage}</span> of <span className="font-semibold" style={{ color: LEATHER_DARK }}>{book.totalPages}</span>. Confirm?
          </p>
          <div className="flex justify-end gap-3">
            <button onClick={() => { setManualPage(String(currentPage)); setStage("manualPage"); }} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>No, let me correct it</button>
            <button onClick={() => finalizeClose(currentPage)} className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white flex items-center gap-2" style={{ background: GREEN }}><Check size={14} /> Yes, save</button>
          </div>
        </Modal>
      )}

      {stage === "manualPage" && (
        <Modal onClose={() => finalizeClose(currentPage)}>
          <h3 className="font-serif text-xl font-bold mb-3" style={{ color: LEATHER_DARK }}>Enter the page you reached</h3>
          <input type="number" min={0} max={book.totalPages} autoFocus value={manualPage} onChange={(e) => setManualPage(e.target.value)}
            className="w-full border rounded-md px-3 py-2 mb-4 focus:outline-none focus:ring-2" style={{ borderColor: BORDER }} />
          <div className="flex justify-end gap-3">
            <button onClick={() => setStage("confirmPage")} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Back</button>
            <button onClick={() => finalizeClose(Math.min(book.totalPages, Math.max(0, parseInt(manualPage, 10) || 0)))}
              className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white flex items-center gap-2" style={{ background: GREEN }}><Check size={14} /> Save progress</button>
          </div>
        </Modal>
      )}
    </div>
  );
}
