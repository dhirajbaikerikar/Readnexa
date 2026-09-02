import React, { useState, useEffect, useRef, useCallback, useMemo } from "react";
import * as docxPreview from "docx-preview";
import { GlobalWorkerOptions, getDocument } from "pdfjs-dist";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { supabase, BOOKS_BUCKET, ASSETS_BUCKET } from "./supabaseClient";
import {
  BookOpen, BookMarked, Library, Upload, FileText, Download, Trash2, Plus,
  ArrowLeft, X, Check, AlertCircle, Loader2, Pencil, Save, RotateCcw,
  ChevronLeft, ChevronRight, LayoutDashboard, BookCopy, Clock, CheckCircle2,
  Settings as SettingsIcon, Sparkles, Sun, Moon, ZoomIn, ZoomOut,
  Play, Pause, SkipBack, SkipForward, Volume2, VolumeX,
  Star, StickyNote, FolderKanban, Share2, Search
} from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, ResponsiveContainer, Tooltip, Cell } from "recharts";

GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

/* ============================== PALETTE ============================== */
// Primary/accent are themeable at runtime via CSS variables (set from Settings).
// The hex values here are just the DEFAULT values written into :root below.
const DEFAULT_PRIMARY = "#5B2A2E";
const DEFAULT_PRIMARY_DARK = "#3E1D20";
const DEFAULT_ACCENT = "#AD8A56";

const INK = "var(--ink)";
const LEATHER = "var(--leather)";
const LEATHER_DARK = "var(--leather-dark)";
const GOLD = "var(--gold)";
const PAPER = "var(--bg)";
const CARD = "var(--card)";
const MUTED = "var(--muted)";
const GREEN = "#3F7A57";
const RED = "#C24444";
const BORDER = "var(--border)";

function darkenHex(hex, factor = 0.7) {
  try {
    const h = hex.replace("#", "");
    const r = Math.round(parseInt(h.substring(0, 2), 16) * factor);
    const g = Math.round(parseInt(h.substring(2, 4), 16) * factor);
    const b = Math.round(parseInt(h.substring(4, 6), 16) * factor);
    return `#${[r, g, b].map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, "0")).join("")}`;
  } catch (e) { return DEFAULT_PRIMARY_DARK; }
}

const SPINES = ["#5B2A2E", "#2F4858", "#4C6444", "#7A5230", "#5C4A72"];
function spineColor(seed) {
  let h = 0;
  for (let i = 0; i < seed.length; i++) h = (h * 31 + seed.charCodeAt(i)) >>> 0;
  return SPINES[h % SPINES.length];
}

/* ============================== HELPERS ============================== */
const MAX_FILE_BYTES = 15 * 1024 * 1024; // Supabase free tier allows much bigger files than artifact storage

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
function progressOf(book) {
  const total = book.total_pages || 1;
  return Math.min(100, Math.round(((book.last_read_page || 0) / total) * 100));
}
function statusOf(book) {
  const p = progressOf(book);
  if (p <= 0) return { label: "Not started", color: MUTED };
  if (p >= 100) return { label: "Finished", color: GREEN };
  return { label: "Reading", color: GOLD };
}

/* ---- Real page rendering & counting ----
   PDF: reads the real page count via PDF.js — the file is shown natively,
        never converted, so this was always accurate.
   DOCX: rendered with docx-preview, which lays the document out into actual
         page elements matching Word's own pagination (not a stripped-down
         HTML reflow like a basic converter would produce, and not a guess).
         The page count comes directly from counting those real rendered
         pages, so it can never drift from what the reader actually shows.
         Legacy .doc files aren't supported for in-browser rendering. */
async function detectPdfPageCount(file) {
  const buf = await file.arrayBuffer();
  const doc = await getDocument({ data: buf }).promise;
  return doc.numPages || null;
}

// Renders a .docx into a hidden, laid-out (but off-screen) container so the
// browser computes real page breaks, then returns each page's HTML plus the
// true page count. Used both for upload-time detection and for the reader
// itself, so the two can never disagree.
async function renderDocxToPages(arrayBuffer) {
  const holder = document.createElement("div");
  holder.style.position = "fixed";
  holder.style.left = "-99999px";
  holder.style.top = "0";
  holder.style.visibility = "hidden";
  holder.style.pointerEvents = "none";
  document.body.appendChild(holder);
  try {
    await docxPreview.renderAsync(arrayBuffer, holder, null, {
      inWrapper: true,
      breakPages: true,
      ignoreLastRenderedPageBreak: true,
      experimental: true,
      renderHeaders: true,
      renderFooters: true,
    });
    const wrapper = holder.querySelector(".docx-wrapper") || holder;
    const pageEls = Array.from(wrapper.children).filter((el) => el.nodeType === 1);
    const pagesHtml = pageEls.length ? pageEls.map((el) => el.outerHTML) : [holder.innerHTML || ""];
    return pagesHtml;
  } finally {
    document.body.removeChild(holder);
  }
}

async function detectDocxPageCount(file) {
  try {
    const buf = await file.arrayBuffer();
    const pages = await renderDocxToPages(buf);
    return pages.length || null;
  } catch (e) {
    return null;
  }
}

async function detectPageCount(file, ext) {
  try {
    if (ext === "pdf") return await detectPdfPageCount(file);
    if (ext === "docx") return await detectDocxPageCount(file);
  } catch (e) { return null; }
  return null; // legacy .doc — not supported in-browser
}

/* ============================== ROOT APP ============================== */
export default function App() {
  const [ready, setReady] = useState(false);
  const [folders, setFolders] = useState([]);
  const [booksByFolder, setBooksByFolder] = useState({});
  const [view, setView] = useState("dashboard");
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
  const [readingTarget, setReadingTarget] = useState(null);
  const [loadError, setLoadError] = useState("");

  const [theme, setTheme] = useState({
    site_name: "Readnexa",
    tagline: "Read · Track · Remember",
    logo_url: null,
    primary_color: DEFAULT_PRIMARY,
    accent_color: DEFAULT_ACCENT,
  });
  const [showSettings, setShowSettings] = useState(false);

  const [showEditFolder, setShowEditFolder] = useState(null); // {id, name}
  const [editFolderName, setEditFolderName] = useState("");
  const [editFolderError, setEditFolderError] = useState("");

  const [editingBook, setEditingBook] = useState(null); // {folderId, id, name, total_pages}
  const [confirmReset, setConfirmReset] = useState(null); // {folderId, id, name}

  const resetProgress = async (folderId, id) => {
    await updateBook(folderId, id, { last_read_page: 0, last_read_at: null });
    setConfirmReset(null);
    showToast("Progress reset to 0%");
  };

  const [mode, setMode] = useState(() => {
    try { return localStorage.getItem("readnexa-mode") || "light"; } catch (e) { return "light"; }
  });
  useEffect(() => {
    document.documentElement.classList.toggle("dark", mode === "dark");
    try { localStorage.setItem("readnexa-mode", mode); } catch (e) {}
  }, [mode]);
  const toggleMode = () => setMode((m) => (m === "light" ? "dark" : "light"));

  // Page-count detection state for the upload flow
  const [pageStage, setPageStage] = useState("entry"); // detecting | confirm | entry
  const [detectedPages, setDetectedPages] = useState(null);
  const [detectFailed, setDetectFailed] = useState(false);

  const [notes, setNotes] = useState([]);
  const [collections, setCollections] = useState([]);
  const [showNoteEditor, setShowNoteEditor] = useState(null); // {id?, title, content} | null
  const [activeCollectionId, setActiveCollectionId] = useState(null);
  const [showManageCollectionBooks, setShowManageCollectionBooks] = useState(false);
  const [newCollectionName, setNewCollectionName] = useState("");
  const [showNewCollection, setShowNewCollection] = useState(false);
  const [confirmPermanentDelete, setConfirmPermanentDelete] = useState(null); // {folderId, id, name}

  const showToast = (msg) => { setToast(msg); setTimeout(() => setToast(""), 2800); };

  const applyTheme = (t) => {
    document.documentElement.style.setProperty("--leather", t.primary_color || DEFAULT_PRIMARY);
    document.documentElement.style.setProperty("--leather-dark", darkenHex(t.primary_color || DEFAULT_PRIMARY, 0.68));
    document.documentElement.style.setProperty("--gold", t.accent_color || DEFAULT_ACCENT);
  };

  const loadAll = useCallback(async () => {
    setLoadError("");
    const { data: folderRows, error: fErr } = await supabase.from("folders").select("*").order("created_at");
    if (fErr) { setLoadError(fErr.message); setReady(true); return; }
    setFolders(folderRows || []);

    const { data: bookRows, error: bErr } = await supabase.from("books").select("*").order("uploaded_at", { ascending: false });
    if (bErr) { setLoadError(bErr.message); setReady(true); return; }
    const grouped = {};
    (folderRows || []).forEach((f) => (grouped[f.id] = []));
    (bookRows || []).forEach((b) => {
      if (!grouped[b.folder_id]) grouped[b.folder_id] = [];
      grouped[b.folder_id].push(b);
    });
    setBooksByFolder(grouped);

    const { data: settingsRow } = await supabase.from("site_settings").select("*").eq("id", 1).maybeSingle();
    if (settingsRow) {
      setTheme(settingsRow);
      applyTheme(settingsRow);
    } else {
      applyTheme(theme);
    }

    const { data: noteRows } = await supabase.from("notes").select("*").order("updated_at", { ascending: false });
    setNotes(noteRows || []);
    const { data: collectionRows } = await supabase.from("collections").select("*").order("created_at");
    setCollections(collectionRows || []);

    setReady(true);
  }, []);

  useEffect(() => { loadAll(); }, [loadAll]);

  const isDuplicateFolderName = (name, excludeId = null) =>
    folders.some((f) => f.id !== excludeId && f.name.trim().toLowerCase() === name.trim().toLowerCase());

  const saveSettings = async (patch) => {
    const next = { ...theme, ...patch };
    const { error } = await supabase.from("site_settings").upsert({ id: 1, ...next });
    if (error) { showToast("Could not save settings: " + error.message); return; }
    setTheme(next);
    applyTheme(next);
    showToast("Settings saved");
  };

  const renameFolder = async () => {
    const name = editFolderName.trim();
    if (!name) return;
    if (isDuplicateFolderName(name, showEditFolder.id)) {
      setEditFolderError("A shelf with this name already exists.");
      return;
    }
    const { error } = await supabase.from("folders").update({ name }).eq("id", showEditFolder.id);
    if (error) { setEditFolderError(error.message); return; }
    setFolders((prev) => prev.map((f) => (f.id === showEditFolder.id ? { ...f, name } : f)));
    setShowEditFolder(null); setEditFolderName(""); setEditFolderError("");
    showToast("Shelf renamed");
  };

  const saveBookEdit = async () => {
    if (!editingBook) return;
    const name = editingBook.name.trim();
    const pages = parseInt(editingBook.total_pages, 10);
    if (!name) return;
    if (!pages || pages < 1) { showToast("Total pages must be 1 or more."); return; }
    const list = booksByFolder[editingBook.folderId] || [];
    const current = list.find((b) => b.id === editingBook.id);
    const clampedLastRead = Math.min(current ? current.last_read_page || 0 : 0, pages);
    const patch = { name, total_pages: pages, last_read_page: clampedLastRead };
    const { error } = await supabase.from("books").update(patch).eq("id", editingBook.id);
    if (error) { showToast("Could not save: " + error.message); return; }
    setBooksByFolder((prev) => ({
      ...prev,
      [editingBook.folderId]: (prev[editingBook.folderId] || []).map((b) => (b.id === editingBook.id ? { ...b, ...patch } : b)),
    }));
    setEditingBook(null);
    showToast("Book updated");
  };

  const [newFolderError, setNewFolderError] = useState("");

  const createFolder = async () => {
    const name = newFolderName.trim();
    if (!name) return;
    if (isDuplicateFolderName(name)) {
      setNewFolderError("A shelf with this name already exists. Choose a different name.");
      return;
    }
    const id = name.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") + "-" + Date.now().toString(36);
    const { error } = await supabase.from("folders").insert({ id, name });
    if (error) { setNewFolderError("Could not create shelf: " + error.message); return; }
    setFolders((prev) => [...prev, { id, name, created_at: new Date().toISOString() }]);
    setBooksByFolder((prev) => ({ ...prev, [id]: [] }));
    setShowNewFolder(false); setNewFolderName(""); setNewFolderError("");
    showToast(`Shelf "${name}" created`);
  };

  const deleteFolder = async (id) => {
    const books = booksByFolder[id] || [];
    for (const b of books) {
      if (b.storage_path) await supabase.storage.from(BOOKS_BUCKET).remove([b.storage_path]);
    }
    await supabase.from("books").delete().eq("folder_id", id);
    await supabase.from("folders").delete().eq("id", id);
    setFolders((prev) => prev.filter((f) => f.id !== id));
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
    setPendingFile({ file, ext, size: file.size });
    setPendingName(cleanBaseName(file.name));
    setPendingPages("");
    setDetectedPages(null);
    setDetectFailed(false);
    setPageStage("detecting");

    const detected = await detectPageCount(file, ext);
    if (detected && detected > 0) {
      setDetectedPages(detected);
      setPendingPages(String(detected));
      setPageStage("confirm");
    } else {
      setDetectFailed(true);
      setPageStage("entry");
    }
  };
  const cancelUpload = () => {
    setPendingFile(null); setPendingName(""); setPendingPages(""); setUploadError("");
    setPageStage("entry"); setDetectedPages(null); setDetectFailed(false);
  };
  const confirmDetectedPages = (accepted) => {
    if (!accepted) setPendingPages("");
    setPageStage("entry");
  };

  const confirmUploadSave = async () => {
    const pages = parseInt(pendingPages, 10);
    if (!pendingFile || !pendingName.trim()) return;
    if (!pages || pages < 1) { setUploadError("Please enter the total number of pages (a whole number, 1 or more)."); return; }
    setSaving(true);

    const bookId = "b_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    const storagePath = `${activeFolderId}/${bookId}-${pendingFile.file.name}`;

    const { error: upErr } = await supabase.storage.from(BOOKS_BUCKET).upload(storagePath, pendingFile.file, {
      cacheControl: "3600",
      upsert: false,
    });
    if (upErr) { setUploadError("Upload failed: " + upErr.message); setSaving(false); return; }

    const { data: pub } = supabase.storage.from(BOOKS_BUCKET).getPublicUrl(storagePath);
    const fileUrl = pub.publicUrl;

    const row = {
      id: bookId,
      folder_id: activeFolderId,
      name: pendingName.trim(),
      file_name: pendingFile.file.name,
      file_type: pendingFile.ext,
      size: pendingFile.size,
      total_pages: pages,
      last_read_page: 0,
      uploaded_at: new Date().toISOString(),
      last_read_at: null,
      file_url: fileUrl,
      storage_path: storagePath,
      edited_html: null,
      edited_at: null,
    };

    const { error: insErr } = await supabase.from("books").insert(row);
    if (insErr) { setUploadError("Could not save book record: " + insErr.message); setSaving(false); return; }

    setBooksByFolder((prev) => ({ ...prev, [activeFolderId]: [row, ...(prev[activeFolderId] || [])] }));
    setSaving(false); cancelUpload();
    showToast(`"${row.name}" added to shelf`);
  };

  const deleteBook = async (folderId, bookId) => {
    // Soft delete — moves the book to Trash rather than removing it immediately.
    await updateBook(folderId, bookId, { deleted_at: new Date().toISOString() });
    setConfirmDelete(null);
    showToast("Moved to Trash");
  };

  const restoreBook = async (folderId, bookId) => {
    await updateBook(folderId, bookId, { deleted_at: null });
    showToast("Book restored");
  };

  const permanentlyDeleteBook = async (folderId, bookId) => {
    const book = (booksByFolder[folderId] || []).find((b) => b.id === bookId);
    if (book && book.storage_path) await supabase.storage.from(BOOKS_BUCKET).remove([book.storage_path]);
    await supabase.from("books").delete().eq("id", bookId);
    setBooksByFolder((prev) => ({ ...prev, [folderId]: (prev[folderId] || []).filter((b) => b.id !== bookId) }));
    // Also remove it from any collections it was part of.
    setCollections((prev) => prev.map((c) => ({ ...c, book_ids: (c.book_ids || []).filter((id) => id !== bookId) })));
    collections.forEach((c) => {
      if ((c.book_ids || []).includes(bookId)) {
        supabase.from("collections").update({ book_ids: (c.book_ids || []).filter((id) => id !== bookId) }).eq("id", c.id);
      }
    });
    setConfirmPermanentDelete(null);
    showToast("Deleted permanently");
  };

  const toggleFavorite = async (folderId, bookId, current) => {
    await updateBook(folderId, bookId, { is_favorite: !current });
  };

  // ---- Notes ----
  const saveNote = async (note) => {
    const now = new Date().toISOString();
    if (note.id) {
      const patch = { title: note.title.trim() || "Untitled note", content: note.content, updated_at: now };
      const { error } = await supabase.from("notes").update(patch).eq("id", note.id);
      if (error) { showToast("Could not save note: " + error.message); return; }
      setNotes((prev) => prev.map((n) => (n.id === note.id ? { ...n, ...patch } : n)));
    } else {
      const row = { id: "n_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), title: note.title.trim() || "Untitled note", content: note.content, created_at: now, updated_at: now };
      const { error } = await supabase.from("notes").insert(row);
      if (error) { showToast("Could not create note: " + error.message); return; }
      setNotes((prev) => [row, ...prev]);
    }
    setShowNoteEditor(null);
    showToast("Note saved");
  };
  const deleteNote = async (id) => {
    await supabase.from("notes").delete().eq("id", id);
    setNotes((prev) => prev.filter((n) => n.id !== id));
    showToast("Note deleted");
  };

  // ---- Collections ----
  const createCollection = async () => {
    const name = newCollectionName.trim();
    if (!name) return;
    const row = { id: "col_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), name, book_ids: [], created_at: new Date().toISOString() };
    const { error } = await supabase.from("collections").insert(row);
    if (error) { showToast("Could not create collection: " + error.message); return; }
    setCollections((prev) => [...prev, row]);
    setShowNewCollection(false);
    setNewCollectionName("");
    showToast(`Collection "${name}" created`);
  };
  const deleteCollection = async (id) => {
    await supabase.from("collections").delete().eq("id", id);
    setCollections((prev) => prev.filter((c) => c.id !== id));
    if (activeCollectionId === id) setActiveCollectionId(null);
    showToast("Collection deleted");
  };
  const toggleBookInCollection = async (collectionId, bookId) => {
    const col = collections.find((c) => c.id === collectionId);
    if (!col) return;
    const has = (col.book_ids || []).includes(bookId);
    const nextIds = has ? col.book_ids.filter((id) => id !== bookId) : [...(col.book_ids || []), bookId];
    const { error } = await supabase.from("collections").update({ book_ids: nextIds }).eq("id", collectionId);
    if (error) { showToast("Could not update collection: " + error.message); return; }
    setCollections((prev) => prev.map((c) => (c.id === collectionId ? { ...c, book_ids: nextIds } : c)));
  };

  const updateBook = async (folderId, bookId, patch) => {
    const { error } = await supabase.from("books").update(patch).eq("id", bookId);
    if (error) { showToast("Could not save: " + error.message); return; }
    setBooksByFolder((prev) => ({
      ...prev,
      [folderId]: (prev[folderId] || []).map((b) => (b.id === bookId ? { ...b, ...patch } : b)),
    }));
  };

  const activeFolder = folders.find((f) => f.id === activeFolderId);
  const activeBooks = (booksByFolder[activeFolderId] || []).filter((b) => !b.deleted_at);

  const allBooksFlat = useMemo(() => {
    const out = [];
    folders.forEach((f) => (booksByFolder[f.id] || []).forEach((b) => { if (!b.deleted_at) out.push({ ...b, folderName: f.name }); }));
    return out;
  }, [folders, booksByFolder]);

  const trashedBooksFlat = useMemo(() => {
    const out = [];
    folders.forEach((f) => (booksByFolder[f.id] || []).forEach((b) => { if (b.deleted_at) out.push({ ...b, folderName: f.name }); }));
    return out.sort((a, b) => new Date(b.deleted_at) - new Date(a.deleted_at));
  }, [folders, booksByFolder]);

  return (
    <div className="min-h-screen w-full" style={{ background: PAPER }}>
      <style>{`
        :root {
          --leather: ${DEFAULT_PRIMARY};
          --leather-dark: ${DEFAULT_PRIMARY_DARK};
          --gold: ${DEFAULT_ACCENT};
          --bg: #F6F1E6;
          --card: #FFFDF8;
          --ink: #241A15;
          --muted: #8A7B68;
          --border: #E4DAC7;
          --soft: #EFE7D5;
        }
        html.dark {
          --bg: #171310;
          --card: #221C18;
          --ink: #F1EAE0;
          --muted: #A79A89;
          --border: #3A322B;
          --soft: #2A241E;
        }
        html.dark body { background: var(--bg); }
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

      <TopBar view={view} theme={theme} mode={mode} onToggleMode={toggleMode}
        onHome={() => { setView("dashboard"); setActiveFolderId(null); }}
        onOpenSettings={() => setShowSettings(true)}
        allBooksFlat={allBooksFlat}
        onSearchSelect={(b) => setReadingTarget({ folderId: b.folder_id, book: b })} />

      <div className="flex items-start">
        <Sidebar view={view} onNavigate={(v) => { setView(v); setActiveFolderId(null); setActiveCollectionId(null); }} />

        <main className="flex-1 min-w-0 px-6 py-8 max-w-6xl mx-auto">
        {loadError && (
          <div className="mb-6 p-4 rounded-lg text-sm flex items-start gap-2" style={{ background: "#FBEAEA", color: RED }}>
            <AlertCircle size={16} className="mt-0.5 shrink-0" />
            <div>
              <p className="font-medium">Could not load your library from Supabase.</p>
              <p className="mt-1 font-mono text-xs">{loadError}</p>
              <p className="mt-1">Check that your Supabase URL/key are set correctly and that the `folders` and `books` tables exist.</p>
            </div>
          </div>
        )}
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
            onEditFolder={(f) => { setShowEditFolder(f); setEditFolderName(f.name); setEditFolderError(""); }}
            onDeleteFolder={(f) => setConfirmDelete({ type: "folder", id: f.id, name: f.name })}
            onContinueReading={(b) => setReadingTarget({ folderId: b.folder_id, book: b })}
          />
        ) : view === "library" ? (
          <LibraryView folders={folders} booksByFolder={booksByFolder} onOpenFolder={openFolder}
            onNewFolder={() => setShowNewFolder(true)}
            onEditFolder={(f) => { setShowEditFolder(f); setEditFolderName(f.name); setEditFolderError(""); }}
            onDeleteFolder={(f) => setConfirmDelete({ type: "folder", id: f.id, name: f.name })} />
        ) : view === "recent" ? (
          <FlatBookListView
            title="Recent"
            subtitle="Your most recently added books and documents, across every shelf."
            books={[...allBooksFlat].sort((a, b) => new Date(b.uploaded_at) - new Date(a.uploaded_at)).slice(0, 30)}
            onRead={(b) => setReadingTarget({ folderId: b.folder_id, book: b })}
            onEditBook={(b) => setEditingBook({ folderId: b.folder_id, id: b.id, name: b.name, total_pages: b.total_pages })}
            onResetProgress={(b) => setConfirmReset({ folderId: b.folder_id, id: b.id, name: b.name })}
            onDeleteBook={(b) => setConfirmDelete({ type: "book", id: b.id, name: b.name, folderId: b.folder_id })}
            onToggleFavorite={(b) => toggleFavorite(b.folder_id, b.id, b.is_favorite)}
          />
        ) : view === "shelf" ? (
          <ShelfView
            folder={activeFolder}
            books={activeBooks}
            onBack={() => { setView("dashboard"); setActiveFolderId(null); }}
            onUploadClick={onPickFile}
            onRead={(book) => setReadingTarget({ folderId: activeFolderId, book })}
            onEditBook={(b) => setEditingBook({ folderId: activeFolderId, id: b.id, name: b.name, total_pages: b.total_pages })}
            onResetProgress={(b) => setConfirmReset({ folderId: activeFolderId, id: b.id, name: b.name })}
            onDeleteBook={(b) => setConfirmDelete({ type: "book", id: b.id, name: b.name, folderId: activeFolderId })}
            onToggleFavorite={(b) => toggleFavorite(activeFolderId, b.id, b.is_favorite)}
          />
        ) : view === "favorites" ? (
          <FlatBookListView
            title="Favorites"
            subtitle="Books and documents you've marked as favorites."
            books={allBooksFlat.filter((b) => b.is_favorite)}
            onRead={(b) => setReadingTarget({ folderId: b.folder_id, book: b })}
            onEditBook={(b) => setEditingBook({ folderId: b.folder_id, id: b.id, name: b.name, total_pages: b.total_pages })}
            onResetProgress={(b) => setConfirmReset({ folderId: b.folder_id, id: b.id, name: b.name })}
            onDeleteBook={(b) => setConfirmDelete({ type: "book", id: b.id, name: b.name, folderId: b.folder_id })}
            onToggleFavorite={(b) => toggleFavorite(b.folder_id, b.id, b.is_favorite)}
            emptyText="No favorites yet — tap the star on any book to add it here."
          />
        ) : view === "notes" ? (
          <NotesView notes={notes} onNew={() => setShowNoteEditor({ title: "", content: "" })}
            onEdit={(n) => setShowNoteEditor(n)} onDelete={deleteNote} />
        ) : view === "collections" ? (
          activeCollectionId ? (
            <CollectionDetailView
              collection={collections.find((c) => c.id === activeCollectionId)}
              books={allBooksFlat.filter((b) => (collections.find((c) => c.id === activeCollectionId)?.book_ids || []).includes(b.id))}
              onBack={() => setActiveCollectionId(null)}
              onManageBooks={() => setShowManageCollectionBooks(true)}
              onDelete={() => deleteCollection(activeCollectionId)}
              onRead={(b) => setReadingTarget({ folderId: b.folder_id, book: b })}
              onEditBook={(b) => setEditingBook({ folderId: b.folder_id, id: b.id, name: b.name, total_pages: b.total_pages })}
              onResetProgress={(b) => setConfirmReset({ folderId: b.folder_id, id: b.id, name: b.name })}
              onDeleteBook={(b) => setConfirmDelete({ type: "book", id: b.id, name: b.name, folderId: b.folder_id })}
              onToggleFavorite={(b) => toggleFavorite(b.folder_id, b.id, b.is_favorite)}
            />
          ) : (
            <CollectionsView collections={collections} onOpen={(id) => setActiveCollectionId(id)} onNew={() => setShowNewCollection(true)} />
          )
        ) : view === "trash" ? (
          <TrashView books={trashedBooksFlat}
            onRestore={(b) => restoreBook(b.folder_id, b.id)}
            onPermanentDelete={(b) => setConfirmPermanentDelete({ folderId: b.folder_id, id: b.id, name: b.name })} />
        ) : (
          <ComingSoonView view={view} />
        )}
        </main>
      </div>

      {showNewFolder && (
        <Modal onClose={() => { setShowNewFolder(false); setNewFolderName(""); setNewFolderError(""); }}>
          <h3 className="font-serif text-xl font-bold mb-4" style={{ color: LEATHER_DARK }}>New shelf</h3>
          <label className="text-sm block mb-1" style={{ color: MUTED }}>Shelf name</label>
          <input autoFocus value={newFolderName} onChange={(e) => { setNewFolderName(e.target.value); setNewFolderError(""); }}
            onKeyDown={(e) => e.key === "Enter" && createFolder()} placeholder="e.g. Business Analysis"
            className="w-full border rounded-md px-3 py-2 mb-2 focus:outline-none focus:ring-2"
            style={{ borderColor: newFolderError ? RED : BORDER }} />
          {newFolderError && <p className="text-sm flex items-center gap-1.5 mb-3" style={{ color: RED }}><AlertCircle size={14} /> {newFolderError}</p>}
          <div className="flex justify-end gap-3 mt-3">
            <button onClick={() => { setShowNewFolder(false); setNewFolderName(""); setNewFolderError(""); }} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
            <button onClick={createFolder} disabled={!newFolderName.trim()} className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white disabled:opacity-40" style={{ background: LEATHER }}>Create shelf</button>
          </div>
        </Modal>
      )}

      {showEditFolder && (
        <Modal onClose={() => { setShowEditFolder(null); setEditFolderError(""); }}>
          <h3 className="font-serif text-xl font-bold mb-4" style={{ color: LEATHER_DARK }}>Rename shelf</h3>
          <label className="text-sm block mb-1" style={{ color: MUTED }}>Shelf name</label>
          <input autoFocus value={editFolderName} onChange={(e) => { setEditFolderName(e.target.value); setEditFolderError(""); }}
            onKeyDown={(e) => e.key === "Enter" && renameFolder()}
            className="w-full border rounded-md px-3 py-2 mb-2 focus:outline-none focus:ring-2"
            style={{ borderColor: editFolderError ? RED : BORDER }} />
          {editFolderError && <p className="text-sm flex items-center gap-1.5 mb-3" style={{ color: RED }}><AlertCircle size={14} /> {editFolderError}</p>}
          <div className="flex justify-end gap-3 mt-3">
            <button onClick={() => { setShowEditFolder(null); setEditFolderError(""); }} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
            <button onClick={renameFolder} disabled={!editFolderName.trim()} className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white disabled:opacity-40" style={{ background: LEATHER }}>Save name</button>
          </div>
        </Modal>
      )}

      {editingBook && (
        <Modal onClose={() => setEditingBook(null)}>
          <h3 className="font-serif text-xl font-bold mb-4" style={{ color: LEATHER_DARK }}>Edit book details</h3>
          <label className="text-sm block mb-1" style={{ color: MUTED }}>Book / document name</label>
          <input autoFocus value={editingBook.name} onChange={(e) => setEditingBook({ ...editingBook, name: e.target.value })}
            className="w-full border rounded-md px-3 py-2 mb-4 focus:outline-none focus:ring-2" style={{ borderColor: BORDER }} />
          <label className="text-sm block mb-1" style={{ color: MUTED }}>Total pages</label>
          <input type="number" min="1" value={editingBook.total_pages} onChange={(e) => setEditingBook({ ...editingBook, total_pages: e.target.value })}
            onKeyDown={(e) => e.key === "Enter" && saveBookEdit()}
            className="w-full border rounded-md px-3 py-2 mb-2 focus:outline-none focus:ring-2" style={{ borderColor: BORDER }} />
          <p className="text-xs mb-3" style={{ color: MUTED }}>If you lower the page count below your current reading progress, progress will be adjusted automatically.</p>
          <div className="flex justify-end gap-3 mt-3">
            <button onClick={() => setEditingBook(null)} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
            <button onClick={saveBookEdit} disabled={!editingBook.name.trim()} className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white disabled:opacity-40" style={{ background: LEATHER }}>Save changes</button>
          </div>
        </Modal>
      )}

      {confirmReset && (
        <Modal onClose={() => setConfirmReset(null)}>
          <h3 className="font-serif text-xl font-bold mb-2" style={{ color: LEATHER_DARK }}>Reset reading progress?</h3>
          <p className="text-sm mb-6" style={{ color: MUTED }}>
            "{confirmReset.name}" will go back to <span className="font-semibold">0%</span> — its current progress will be lost. This cannot be undone.
          </p>
          <div className="flex justify-end gap-3">
            <button onClick={() => setConfirmReset(null)} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
            <button onClick={() => resetProgress(confirmReset.folderId, confirmReset.id)}
              className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white flex items-center gap-2" style={{ background: LEATHER }}>
              <RotateCcw size={14} /> Reset to 0%
            </button>
          </div>
        </Modal>
      )}

      {showSettings && (
        <SettingsModal theme={theme} onClose={() => setShowSettings(false)} onSave={saveSettings} />
      )}

      {showNoteEditor && (
        <Modal onClose={() => setShowNoteEditor(null)}>
          <h3 className="font-serif text-xl font-bold mb-4" style={{ color: LEATHER_DARK }}>{showNoteEditor.id ? "Edit note" : "New note"}</h3>
          <input autoFocus value={showNoteEditor.title} onChange={(e) => setShowNoteEditor({ ...showNoteEditor, title: e.target.value })}
            placeholder="Title" className="w-full border rounded-md px-3 py-2 mb-3 font-medium focus:outline-none focus:ring-2" style={{ borderColor: BORDER }} />
          <textarea value={showNoteEditor.content} onChange={(e) => setShowNoteEditor({ ...showNoteEditor, content: e.target.value })}
            placeholder="Write your note…" rows={7}
            className="w-full border rounded-md px-3 py-2 mb-4 focus:outline-none focus:ring-2 resize-none" style={{ borderColor: BORDER }} />
          <div className="flex justify-end gap-3">
            <button onClick={() => setShowNoteEditor(null)} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
            <button onClick={() => saveNote(showNoteEditor)} className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white flex items-center gap-2" style={{ background: LEATHER }}>
              <Save size={14} /> Save note
            </button>
          </div>
        </Modal>
      )}

      {showNewCollection && (
        <Modal onClose={() => { setShowNewCollection(false); setNewCollectionName(""); }}>
          <h3 className="font-serif text-xl font-bold mb-4" style={{ color: LEATHER_DARK }}>New collection</h3>
          <label className="text-sm block mb-1" style={{ color: MUTED }}>Collection name</label>
          <input autoFocus value={newCollectionName} onChange={(e) => setNewCollectionName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && createCollection()} placeholder="e.g. Interview Prep"
            className="w-full border rounded-md px-3 py-2 mb-5 focus:outline-none focus:ring-2" style={{ borderColor: BORDER }} />
          <div className="flex justify-end gap-3">
            <button onClick={() => { setShowNewCollection(false); setNewCollectionName(""); }} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
            <button onClick={createCollection} disabled={!newCollectionName.trim()} className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white disabled:opacity-40" style={{ background: LEATHER }}>Create</button>
          </div>
        </Modal>
      )}

      {showManageCollectionBooks && activeCollectionId && (
        <Modal onClose={() => setShowManageCollectionBooks(false)}>
          <h3 className="font-serif text-xl font-bold mb-4" style={{ color: LEATHER_DARK }}>Manage books in this collection</h3>
          <div className="max-h-80 overflow-y-auto -mx-1 px-1">
            {allBooksFlat.length === 0 ? (
              <p className="text-sm" style={{ color: MUTED }}>No books uploaded yet.</p>
            ) : allBooksFlat.map((b) => {
              const col = collections.find((c) => c.id === activeCollectionId);
              const checked = (col?.book_ids || []).includes(b.id);
              return (
                <label key={b.id} className="flex items-center gap-3 px-2 py-2 rounded-md hover:bg-black/5 cursor-pointer">
                  <input type="checkbox" checked={checked} onChange={() => toggleBookInCollection(activeCollectionId, b.id)} />
                  <span className="text-sm truncate flex-1" style={{ color: INK }}>{b.name}</span>
                  <span className="text-xs shrink-0" style={{ color: MUTED }}>{b.folderName}</span>
                </label>
              );
            })}
          </div>
          <div className="flex justify-end mt-5">
            <button onClick={() => setShowManageCollectionBooks(false)} className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white" style={{ background: LEATHER }}>Done</button>
          </div>
        </Modal>
      )}

      {confirmPermanentDelete && (
        <Modal onClose={() => setConfirmPermanentDelete(null)}>
          <h3 className="font-serif text-xl font-bold mb-2" style={{ color: LEATHER_DARK }}>Delete permanently?</h3>
          <p className="text-sm mb-6" style={{ color: MUTED }}>
            "{confirmPermanentDelete.name}" and its file will be permanently deleted. This cannot be undone.
          </p>
          <div className="flex justify-end gap-3">
            <button onClick={() => setConfirmPermanentDelete(null)} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
            <button onClick={() => permanentlyDeleteBook(confirmPermanentDelete.folderId, confirmPermanentDelete.id)}
              className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white flex items-center gap-2" style={{ background: RED }}>
              <Trash2 size={14} /> Delete permanently
            </button>
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

          {pageStage === "detecting" && (
            <div className="flex items-center gap-2 mb-4 p-3 rounded-md" style={{ background: "var(--soft)" }}>
              <Loader2 className="animate-spin shrink-0" size={16} style={{ color: LEATHER }} />
              <p className="text-sm" style={{ color: MUTED }}>Analyzing document to detect page count…</p>
            </div>
          )}

          {pageStage === "confirm" && (
            <div className="mb-4 p-3 rounded-md" style={{ background: "var(--soft)" }}>
              <p className="text-sm flex items-center gap-1.5 mb-3" style={{ color: LEATHER_DARK }}>
                <Sparkles size={15} style={{ color: GOLD }} /> We detected <span className="font-bold">{detectedPages}</span> pages in this document. Is that correct?
              </p>
              {pendingFile && pendingFile.ext === "docx" && (
                <p className="text-xs mb-3" style={{ color: MUTED }}>Based on how it renders in the browser — for an exact match to Word's own count, upload as PDF instead.</p>
              )}
              <div className="flex gap-2">
                <button onClick={() => confirmDetectedPages(false)} className="px-3 py-1.5 rounded-md text-sm" style={{ border: `1px solid ${BORDER}`, color: LEATHER_DARK }}>No, let me correct it</button>
                <button onClick={() => confirmDetectedPages(true)} className="btn-lift px-3 py-1.5 rounded-md text-sm font-medium text-white flex items-center gap-1.5" style={{ background: GREEN }}>
                  <Check size={14} /> Yes, that's correct
                </button>
              </div>
            </div>
          )}

          {pageStage === "entry" && (
            <>
              <label className="text-sm block mb-1" style={{ color: MUTED }}>Total pages</label>
              <input type="number" min="1" value={pendingPages} onChange={(e) => setPendingPages(e.target.value)}
                placeholder="e.g. 120" autoFocus={detectFailed || detectedPages === null}
                onKeyDown={(e) => e.key === "Enter" && confirmUploadSave()}
                className="w-full border rounded-md px-3 py-2 mb-2 focus:outline-none focus:ring-2" style={{ borderColor: BORDER }} />
              {detectedPages !== null ? (
                <p className="text-xs mb-3 flex items-center gap-1" style={{ color: MUTED }}><Pencil size={11} /> Enter the correct page count below.</p>
              ) : detectFailed ? (
                <p className="text-xs mb-3" style={{ color: MUTED }}>We couldn't automatically detect the page count for this file — please enter it manually.</p>
              ) : (
                <p className="text-xs mb-3" style={{ color: MUTED }}>Used to calculate the reading-progress bar shown on the shelf and dashboard.</p>
              )}
            </>
          )}

          {uploadError && <p className="text-sm flex items-center gap-1.5 mb-2" style={{ color: RED }}><AlertCircle size={14} /> {uploadError}</p>}
          <div className="flex justify-end gap-3 mt-3">
            <button onClick={cancelUpload} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
            <button onClick={confirmUploadSave} disabled={!pendingName.trim() || saving || pageStage !== "entry" || !pendingPages}
              className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white disabled:opacity-40 flex items-center gap-2" style={{ background: LEATHER }}>
              {saving ? <Loader2 className="animate-spin" size={14} /> : <Check size={14} />} Save to shelf
            </button>
          </div>
        </Modal>
      )}

      {confirmDelete && (
        <Modal onClose={() => setConfirmDelete(null)}>
          <h3 className="font-serif text-xl font-bold mb-2" style={{ color: LEATHER_DARK }}>
            {confirmDelete.type === "folder" ? "Delete shelf?" : "Move to Trash?"}
          </h3>
          <p className="text-sm mb-6" style={{ color: MUTED }}>
            {confirmDelete.type === "folder"
              ? `"${confirmDelete.name}" will be permanently removed, along with every book on it. This cannot be undone.`
              : `"${confirmDelete.name}" will be moved to Trash, where you can restore it or delete it permanently later.`}
          </p>
          <div className="flex justify-end gap-3">
            <button onClick={() => setConfirmDelete(null)} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
            <button onClick={() => confirmDelete.type === "folder" ? deleteFolder(confirmDelete.id) : deleteBook(confirmDelete.folderId, confirmDelete.id)}
              className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white flex items-center gap-2" style={{ background: RED }}>
              <Trash2 size={14} /> {confirmDelete.type === "folder" ? "Delete" : "Move to Trash"}
            </button>
          </div>
        </Modal>
      )}

      {readingTarget && (
        <BookReader
          book={readingTarget.book}
          onClose={() => setReadingTarget(null)}
          onCommitProgress={(page) => updateBook(readingTarget.folderId, readingTarget.book.id, { last_read_page: page, last_read_at: new Date().toISOString() })}
          onSaveEdit={(html) => updateBook(readingTarget.folderId, readingTarget.book.id, { edited_html: html, edited_at: new Date().toISOString() })}
          onFixTotalPages={(realTotal) => updateBook(readingTarget.folderId, readingTarget.book.id, { total_pages: realTotal, last_read_page: Math.min(readingTarget.book.last_read_page || 0, realTotal) })}
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
function TopBar({ view, theme, mode, onToggleMode, onHome, onOpenSettings, allBooksFlat, onSearchSelect }) {
  const [query, setQuery] = useState("");
  const results = query.trim().length > 0
    ? allBooksFlat.filter((b) => b.name.toLowerCase().includes(query.trim().toLowerCase())).slice(0, 8)
    : [];
  return (
    <header style={{ background: LEATHER_DARK, borderBottom: `3px solid ${GOLD}` }}>
      <div className="max-w-6xl mx-auto px-6 py-5 flex items-center justify-between gap-4">
        <button onClick={onHome} className="flex items-center gap-3 min-w-0 shrink-0">
          {theme.logo_url ? (
            <img src={theme.logo_url} alt="" className="h-9 w-9 rounded-md object-cover shrink-0" />
          ) : (
            <div className="h-9 w-9 rounded-md flex items-center justify-center shrink-0" style={{ background: GOLD }}>
              <BookOpen size={19} style={{ color: LEATHER_DARK }} />
            </div>
          )}
          <div className="text-left min-w-0 hidden md:block">
            <h1 className="font-serif text-2xl font-bold text-white leading-none truncate">{theme.site_name || "Readnexa"}</h1>
            <p className="text-[11px] tracking-widest uppercase font-mono truncate" style={{ color: "#D8C7A8" }}>{theme.tagline || "Read · Track · Remember"}</p>
          </div>
        </button>

        <div className="relative flex-1 max-w-md">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: "#B7A88C" }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)}
            placeholder="Search your library…"
            className="w-full rounded-md pl-9 pr-3 py-2 text-sm focus:outline-none"
            style={{ background: "rgba(255,255,255,0.08)", color: "white" }} />
          {results.length > 0 && (
            <div className="absolute top-full left-0 right-0 mt-1 rounded-md shadow-2xl overflow-hidden z-30" style={{ background: CARD, border: `1px solid ${BORDER}` }}>
              {results.map((b) => (
                <button key={b.id} onClick={() => { onSearchSelect(b); setQuery(""); }}
                  className="w-full text-left px-3 py-2 text-sm hover:bg-black/5 flex items-center gap-2" style={{ color: INK }}>
                  <FileText size={14} style={{ color: MUTED }} />
                  <span className="truncate">{b.name}</span>
                  <span className="text-xs ml-auto shrink-0" style={{ color: MUTED }}>{b.folderName}</span>
                </button>
              ))}
            </div>
          )}
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <button onClick={onToggleMode} title={mode === "dark" ? "Switch to light mode" : "Switch to dark mode"} className="p-2 rounded-md hover:bg-white/10" style={{ color: "#F1E6D2" }}>
            {mode === "dark" ? <Sun size={17} /> : <Moon size={17} />}
          </button>
          <button onClick={onOpenSettings} title="Site settings" className="p-2 rounded-md hover:bg-white/10" style={{ color: "#F1E6D2" }}>
            <SettingsIcon size={17} />
          </button>
        </div>
      </div>
    </header>
  );
}

/* ============================== SETTINGS MODAL ============================== */
function SettingsModal({ theme, onClose, onSave }) {
  const [siteName, setSiteName] = useState(theme.site_name || "Readnexa");
  const [tagline, setTagline] = useState(theme.tagline || "");
  const [primary, setPrimary] = useState(theme.primary_color || DEFAULT_PRIMARY);
  const [accent, setAccent] = useState(theme.accent_color || DEFAULT_ACCENT);
  const [logoFile, setLogoFile] = useState(null);
  const [logoPreview, setLogoPreview] = useState(theme.logo_url || null);
  const [removeLogo, setRemoveLogo] = useState(false);
  const [saving, setSaving] = useState(false);
  const [err, setErr] = useState("");
  const logoInputRef = useRef(null);

  const onPickLogo = (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    if (!file.type.startsWith("image/")) { setErr("Please choose an image file for the logo."); return; }
    if (file.size > 1.5 * 1024 * 1024) { setErr("Logo image should be under 1.5MB."); return; }
    setErr("");
    setLogoFile(file);
    setRemoveLogo(false);
    setLogoPreview(URL.createObjectURL(file));
  };

  const handleSave = async () => {
    if (!siteName.trim()) { setErr("Site name can't be empty."); return; }
    setSaving(true);
    let logo_url = theme.logo_url || null;
    try {
      if (removeLogo) {
        logo_url = null;
      } else if (logoFile) {
        const path = `logo-${Date.now()}-${logoFile.name}`;
        const { error: upErr } = await supabase.storage.from(ASSETS_BUCKET).upload(path, logoFile, { upsert: false });
        if (upErr) throw upErr;
        const { data: pub } = supabase.storage.from(ASSETS_BUCKET).getPublicUrl(path);
        logo_url = pub.publicUrl;
      }
      await onSave({ site_name: siteName.trim(), tagline: tagline.trim(), primary_color: primary, accent_color: accent, logo_url });
      onClose();
    } catch (e) {
      setErr("Could not save logo: " + (e.message || "unknown error") + ". Make sure the 'assets' storage bucket exists and is public.");
    }
    setSaving(false);
  };

  return (
    <Modal onClose={onClose}>
      <h3 className="font-serif text-xl font-bold mb-4" style={{ color: LEATHER_DARK }}>Site settings</h3>

      <label className="text-sm block mb-1" style={{ color: MUTED }}>Site name</label>
      <input value={siteName} onChange={(e) => setSiteName(e.target.value)}
        className="w-full border rounded-md px-3 py-2 mb-4 focus:outline-none focus:ring-2" style={{ borderColor: BORDER }} />

      <label className="text-sm block mb-1" style={{ color: MUTED }}>Tagline</label>
      <input value={tagline} onChange={(e) => setTagline(e.target.value)} placeholder="Read · Track · Remember"
        className="w-full border rounded-md px-3 py-2 mb-4 focus:outline-none focus:ring-2" style={{ borderColor: BORDER }} />

      <label className="text-sm block mb-2" style={{ color: MUTED }}>Logo</label>
      <div className="flex items-center gap-3 mb-4">
        <div className="h-12 w-12 rounded-md flex items-center justify-center shrink-0 overflow-hidden" style={{ background: primary }}>
          {logoPreview && !removeLogo ? <img src={logoPreview} className="h-full w-full object-cover" alt="" /> : <BookOpen size={20} color="white" />}
        </div>
        <button onClick={() => logoInputRef.current.click()} className="px-3 py-1.5 rounded-md text-sm" style={{ border: `1px solid ${BORDER}`, color: LEATHER_DARK }}>Upload image</button>
        {logoPreview && !removeLogo && (
          <button onClick={() => { setRemoveLogo(true); setLogoPreview(null); setLogoFile(null); }} className="text-sm" style={{ color: RED }}>Remove</button>
        )}
        <input ref={logoInputRef} type="file" accept="image/*" className="hidden" onChange={onPickLogo} />
      </div>

      <div className="grid grid-cols-2 gap-4 mb-4">
        <div>
          <label className="text-sm block mb-1" style={{ color: MUTED }}>Primary color</label>
          <div className="flex items-center gap-2">
            <input type="color" value={primary} onChange={(e) => setPrimary(e.target.value)} className="h-9 w-10 rounded border p-0.5" style={{ borderColor: BORDER }} />
            <input value={primary} onChange={(e) => setPrimary(e.target.value)} className="w-full border rounded-md px-2 py-1.5 text-sm font-mono" style={{ borderColor: BORDER }} />
          </div>
        </div>
        <div>
          <label className="text-sm block mb-1" style={{ color: MUTED }}>Accent color</label>
          <div className="flex items-center gap-2">
            <input type="color" value={accent} onChange={(e) => setAccent(e.target.value)} className="h-9 w-10 rounded border p-0.5" style={{ borderColor: BORDER }} />
            <input value={accent} onChange={(e) => setAccent(e.target.value)} className="w-full border rounded-md px-2 py-1.5 text-sm font-mono" style={{ borderColor: BORDER }} />
          </div>
        </div>
      </div>

      {err && <p className="text-sm flex items-center gap-1.5 mb-3" style={{ color: RED }}><AlertCircle size={14} /> {err}</p>}

      <div className="flex justify-end gap-3 mt-2">
        <button onClick={onClose} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Cancel</button>
        <button onClick={handleSave} disabled={saving} className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white flex items-center gap-2 disabled:opacity-50" style={{ background: primary }}>
          {saving ? <Loader2 className="animate-spin" size={14} /> : <Check size={14} />} Save settings
        </button>
      </div>
    </Modal>
  );
}

/* ============================== SIDEBAR ============================== */
const NAV_ITEMS = [
  { id: "dashboard", label: "Home", icon: LayoutDashboard, enabled: true },
  { id: "library", label: "My Library", icon: Library, enabled: true },
  { id: "recent", label: "Recent", icon: Clock, enabled: true },
  { id: "favorites", label: "Favorites", icon: Star, enabled: true },
  { id: "notes", label: "Notes", icon: StickyNote, enabled: true },
  { id: "collections", label: "Collections", icon: FolderKanban, enabled: true },
  { id: "shared", label: "Shared", icon: Share2, enabled: false },
  { id: "trash", label: "Trash", icon: Trash2, enabled: true },
];

function Sidebar({ view, onNavigate }) {
  return (
    <aside className="w-56 shrink-0 hidden sm:flex flex-col py-6 px-3 gap-0.5" style={{ borderRight: `1px solid ${BORDER}`, background: CARD, minHeight: "calc(100vh - 89px)" }}>
      {NAV_ITEMS.map((item) => {
        const Icon = item.icon;
        const active = view === item.id || (item.id === "library" && view === "shelf");
        return (
          <button key={item.id} onClick={() => onNavigate(item.id)}
            className="flex items-center gap-2.5 px-3 py-2 rounded-md text-sm text-left"
            style={{
              background: active ? "var(--soft)" : "transparent",
              color: active ? LEATHER_DARK : MUTED,
              fontWeight: active ? 600 : 400,
            }}>
            <Icon size={16} />
            <span className="flex-1">{item.label}</span>
            {!item.enabled && (
              <span className="text-[10px] px-1.5 py-0.5 rounded-full font-mono" style={{ background: BORDER, color: MUTED }}>Soon</span>
            )}
          </button>
        );
      })}
    </aside>
  );
}

function ComingSoonView({ view }) {
  const item = NAV_ITEMS.find((n) => n.id === view);
  const Icon = item ? item.icon : Sparkles;
  return (
    <div className="flex flex-col items-center justify-center text-center py-24">
      <div className="h-14 w-14 rounded-full flex items-center justify-center mb-4" style={{ background: "var(--soft)" }}>
        <Icon size={24} style={{ color: LEATHER }} />
      </div>
      <h2 className="font-serif text-xl font-bold mb-1" style={{ color: LEATHER_DARK }}>{item ? item.label : "Coming soon"}</h2>
      <p className="text-sm max-w-sm" style={{ color: MUTED }}>This feature is on the roadmap and isn't built yet — say the word and it's next.</p>
    </div>
  );
}

function LibraryView({ folders, booksByFolder, onOpenFolder, onNewFolder, onEditFolder, onDeleteFolder }) {
  return (
    <div>
      <div className="flex items-end justify-between mb-6 rise">
        <div>
          <p className="text-xs uppercase tracking-widest font-mono" style={{ color: GOLD }}>Every shelf, in one place</p>
          <h2 className="font-serif text-3xl font-bold" style={{ color: LEATHER_DARK }}>My Library</h2>
        </div>
        <button onClick={onNewFolder} className="btn-lift flex items-center gap-1.5 px-4 py-2 rounded-md text-sm font-medium text-white" style={{ background: LEATHER }}>
          <Plus size={16} /> New shelf
        </button>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-5">
        {folders.map((f, i) => {
          const list = (booksByFolder[f.id] || []).filter((b) => !b.deleted_at);
          const tp = list.reduce((s, b) => s + (b.total_pages || 0), 0);
          const rp = list.reduce((s, b) => s + Math.min(b.last_read_page || 0, b.total_pages || 0), 0);
          const pct = tp > 0 ? Math.round((rp / tp) * 100) : 0;
          return (
            <ShelfCard key={f.id} folder={f} count={list.length} percent={pct} delay={i * 60}
              onOpen={() => onOpenFolder(f.id)} onEdit={() => onEditFolder(f)} onDelete={() => onDeleteFolder(f)} />
          );
        })}
      </div>
    </div>
  );
}

function FlatBookListView({ title, subtitle, books, onRead, onEditBook, onResetProgress, onDeleteBook, onToggleFavorite, emptyText }) {
  return (
    <div>
      <div className="mb-6 rise">
        <h2 className="font-serif text-3xl font-bold" style={{ color: LEATHER_DARK }}>{title}</h2>
        <p className="text-sm mt-1" style={{ color: MUTED }}>{subtitle}</p>
      </div>
      {books.length === 0 ? (
        <p className="text-center py-16" style={{ color: MUTED }}>{emptyText || "Nothing here yet."}</p>
      ) : (
        <div className="grid grid-cols-1 gap-4">
          {books.map((b, i) => (
            <BookRow key={b.id} book={b} color={spineColor(b.folder_id)} delay={i * 40}
              onRead={() => onRead(b)} onEdit={() => onEditBook(b)} onReset={() => onResetProgress(b)} onDelete={() => onDeleteBook(b)}
              onToggleFavorite={onToggleFavorite ? () => onToggleFavorite(b) : undefined} />
          ))}
        </div>
      )}
    </div>
  );
}

function NotesView({ notes, onNew, onEdit, onDelete }) {
  return (
    <div>
      <div className="flex items-end justify-between mb-6 rise">
        <div>
          <p className="text-xs uppercase tracking-widest font-mono" style={{ color: GOLD }}>Your personal notepad</p>
          <h2 className="font-serif text-3xl font-bold" style={{ color: LEATHER_DARK }}>Notes</h2>
        </div>
        <button onClick={onNew} className="btn-lift flex items-center gap-1.5 px-4 py-2 rounded-md text-sm font-medium text-white" style={{ background: LEATHER }}>
          <Plus size={16} /> New note
        </button>
      </div>
      {notes.length === 0 ? (
        <p className="text-center py-16" style={{ color: MUTED }}>No notes yet — jot down a thought.</p>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
          {notes.map((n, i) => (
            <div key={n.id} onClick={() => onEdit(n)} className="rise cursor-pointer rounded-xl p-4" style={{ background: CARD, border: `1px solid ${BORDER}`, animationDelay: `${i * 40}ms` }}>
              <div className="flex items-start justify-between gap-2 mb-2">
                <h4 className="font-serif font-bold truncate" style={{ color: LEATHER_DARK }}>{n.title || "Untitled note"}</h4>
                <button onClick={(e) => { e.stopPropagation(); onDelete(n.id); }} className="p-1 -m-1 shrink-0" style={{ color: MUTED }}><Trash2 size={14} /></button>
              </div>
              <p className="text-sm line-clamp-4 whitespace-pre-wrap" style={{ color: MUTED }}>{n.content || "No content."}</p>
              <p className="text-xs font-mono mt-3" style={{ color: MUTED }}>{fmtStamp(n.updated_at).date}</p>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function CollectionsView({ collections, onOpen, onNew }) {
  return (
    <div>
      <div className="flex items-end justify-between mb-6 rise">
        <div>
          <p className="text-xs uppercase tracking-widest font-mono" style={{ color: GOLD }}>Group books across shelves</p>
          <h2 className="font-serif text-3xl font-bold" style={{ color: LEATHER_DARK }}>Collections</h2>
        </div>
        <button onClick={onNew} className="btn-lift flex items-center gap-1.5 px-4 py-2 rounded-md text-sm font-medium text-white" style={{ background: LEATHER }}>
          <Plus size={16} /> New collection
        </button>
      </div>
      {collections.length === 0 ? (
        <p className="text-center py-16" style={{ color: MUTED }}>No collections yet — group related books together, regardless of which shelf they're on.</p>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-5">
          {collections.map((c, i) => (
            <button key={c.id} onClick={() => onOpen(c.id)} className="shelf-card rise text-left rounded-xl overflow-hidden" style={{ background: CARD, border: `1px solid ${BORDER}`, animationDelay: `${i * 60}ms` }}>
              <div className="h-2.5" style={{ background: spineColor(c.id) }} />
              <div className="p-5">
                <FolderKanban size={22} style={{ color: LEATHER }} />
                <h4 className="font-serif font-bold text-base mt-3" style={{ color: LEATHER_DARK }}>{c.name}</h4>
                <p className="text-xs font-mono mt-1" style={{ color: MUTED }}>{(c.book_ids || []).length} book{(c.book_ids || []).length !== 1 ? "s" : ""}</p>
              </div>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function CollectionDetailView({ collection, books, onBack, onManageBooks, onDelete, onRead, onEditBook, onResetProgress, onDeleteBook, onToggleFavorite }) {
  if (!collection) return null;
  return (
    <div className="fade">
      <button onClick={onBack} className="flex items-center gap-1.5 text-sm mb-5" style={{ color: MUTED }}><ArrowLeft size={15} /> Collections</button>
      <div className="flex items-end justify-between mb-6">
        <div>
          <h2 className="font-serif text-2xl font-bold" style={{ color: LEATHER_DARK }}>{collection.name}</h2>
          <p className="text-sm" style={{ color: MUTED }}>{books.length} book{books.length !== 1 ? "s" : ""} in this collection</p>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={onManageBooks} className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white" style={{ background: LEATHER }}>Manage books</button>
          <button onClick={onDelete} className="p-2 rounded-md hover:bg-black/5" style={{ color: MUTED }}><Trash2 size={17} /></button>
        </div>
      </div>
      {books.length === 0 ? (
        <p className="text-center py-16" style={{ color: MUTED }}>No books in this collection yet.</p>
      ) : (
        <div className="grid grid-cols-1 gap-4">
          {books.map((b, i) => (
            <BookRow key={b.id} book={b} color={spineColor(b.folder_id)} delay={i * 40}
              onRead={() => onRead(b)} onEdit={() => onEditBook(b)} onReset={() => onResetProgress(b)} onDelete={() => onDeleteBook(b)}
              onToggleFavorite={() => onToggleFavorite(b)} />
          ))}
        </div>
      )}
    </div>
  );
}

function TrashView({ books, onRestore, onPermanentDelete }) {
  return (
    <div>
      <div className="mb-6 rise">
        <h2 className="font-serif text-3xl font-bold" style={{ color: LEATHER_DARK }}>Trash</h2>
        <p className="text-sm mt-1" style={{ color: MUTED }}>Deleted books stay here until you permanently remove them.</p>
      </div>
      {books.length === 0 ? (
        <p className="text-center py-16" style={{ color: MUTED }}>Trash is empty.</p>
      ) : (
        <div className="grid grid-cols-1 gap-4">
          {books.map((b, i) => (
            <div key={b.id} className="rise rounded-xl p-5 flex items-center gap-4" style={{ background: CARD, border: `1px solid ${BORDER}`, animationDelay: `${i * 40}ms` }}>
              <div className="h-14 w-11 rounded-sm shrink-0 flex items-center justify-center" style={{ background: spineColor(b.folder_id) }}>
                <FileText size={18} color="white" />
              </div>
              <div className="flex-1 min-w-0">
                <p className="font-serif font-bold truncate" style={{ color: LEATHER_DARK }}>{b.name}</p>
                <p className="text-xs font-mono" style={{ color: MUTED }}>{b.folderName} · deleted {fmtStamp(b.deleted_at).date}</p>
              </div>
              <button onClick={() => onRestore(b)} className="btn-lift flex items-center gap-1.5 px-3 py-2 rounded-md text-sm font-medium text-white shrink-0" style={{ background: GREEN }}>
                <RotateCcw size={14} /> Restore
              </button>
              <button onClick={() => onPermanentDelete(b)} className="p-2 rounded-md hover:bg-black/5 shrink-0" style={{ color: RED }} title="Delete permanently"><Trash2 size={17} /></button>
            </div>
          ))}
        </div>
      )}
    </div>
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
function Dashboard({ folders, booksByFolder, allBooksFlat, onOpenFolder, onNewFolder, onEditFolder, onDeleteFolder, onContinueReading }) {
  const totalBooks = allBooksFlat.length;
  const totalPagesAll = allBooksFlat.reduce((s, b) => s + (b.total_pages || 0), 0);
  const readPagesAll = allBooksFlat.reduce((s, b) => s + Math.min(b.last_read_page || 0, b.total_pages || 0), 0);
  const overallPct = totalPagesAll > 0 ? Math.round((readPagesAll / totalPagesAll) * 100) : 0;
  const reading = allBooksFlat.filter((b) => { const p = progressOf(b); return p > 0 && p < 100; }).length;
  const finished = allBooksFlat.filter((b) => progressOf(b) >= 100).length;

  const chartData = folders.map((f) => {
    const list = (booksByFolder[f.id] || []).filter((b) => !b.deleted_at);
    const tp = list.reduce((s, b) => s + (b.total_pages || 0), 0);
    const rp = list.reduce((s, b) => s + Math.min(b.last_read_page || 0, b.total_pages || 0), 0);
    return { name: f.name, percent: tp > 0 ? Math.round((rp / tp) * 100) : 0, count: list.length };
  });

  const continueReading = allBooksFlat
    .filter((b) => { const p = progressOf(b); return p > 0 && p < 100; })
    .sort((a, b) => new Date(b.last_read_at || 0) - new Date(a.last_read_at || 0))
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
                  <div className="h-3" style={{ background: spineColor(b.folder_id) }} />
                  <div className="p-4">
                    <p className="text-xs font-mono mb-1" style={{ color: MUTED }}>{b.folderName}</p>
                    <p className="font-serif font-bold text-sm mb-3" style={{ color: LEATHER_DARK }}>{b.name}</p>
                    <div className="h-1.5 rounded-full overflow-hidden mb-1" style={{ background: BORDER }}>
                      <div className="h-full progress-fill rounded-full" style={{ width: `${pct}%`, background: GOLD }} />
                    </div>
                    <p className="text-xs" style={{ color: MUTED }}>{pct}% · page {b.last_read_page} of {b.total_pages}</p>
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
          const list = (booksByFolder[f.id] || []).filter((b) => !b.deleted_at);
          const tp = list.reduce((s, b) => s + (b.total_pages || 0), 0);
          const rp = list.reduce((s, b) => s + Math.min(b.last_read_page || 0, b.total_pages || 0), 0);
          const pct = tp > 0 ? Math.round((rp / tp) * 100) : 0;
          return (
            <ShelfCard key={f.id} folder={f} count={list.length} percent={pct} delay={i * 70}
              onOpen={() => onOpenFolder(f.id)} onEdit={() => onEditFolder(f)} onDelete={() => onDeleteFolder(f)} />
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

function ShelfCard({ folder, count, percent, delay, onOpen, onEdit, onDelete }) {
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
              <h4 className="font-serif font-bold text-base truncate" style={{ color: LEATHER_DARK }}>{folder.name}</h4>
              {hover && (
                <div className="flex items-center gap-1 shrink-0">
                  <button onClick={(e) => { e.stopPropagation(); onEdit(); }} className="p-1 -m-1" style={{ color: MUTED }} title="Rename shelf">
                    <Pencil size={13} />
                  </button>
                  <button onClick={(e) => { e.stopPropagation(); onDelete(); }} className="p-1 -m-1" style={{ color: MUTED }} title="Delete shelf">
                    <Trash2 size={14} />
                  </button>
                </div>
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
function ShelfView({ folder, books, onBack, onUploadClick, onRead, onEditBook, onResetProgress, onDeleteBook, onToggleFavorite }) {
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
          {books.map((b, i) => <BookRow key={b.id} book={b} color={color} delay={i * 50} onRead={() => onRead(b)} onEdit={() => onEditBook(b)} onReset={() => onResetProgress(b)} onDelete={() => onDeleteBook(b)} onToggleFavorite={() => onToggleFavorite(b)} />)}
        </div>
      )}
    </div>
  );
}

function BookRow({ book, color, delay, onRead, onEdit, onReset, onDelete, onToggleFavorite }) {
  const pct = progressOf(book);
  const status = statusOf(book);
  const stamp = fmtStamp(book.uploaded_at);
  return (
    <div className="rise rounded-xl p-5 flex flex-col sm:flex-row sm:items-center gap-4" style={{ background: CARD, border: `1px solid ${BORDER}`, animationDelay: `${delay}ms` }}>
      <div className="h-14 w-11 rounded-sm shrink-0 flex items-center justify-center" style={{ background: color }}>
        <FileText size={18} color="white" />
      </div>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2 flex-wrap">
          <p className="font-serif font-bold" style={{ color: LEATHER_DARK }}>{book.name}</p>
          <span className="text-xs px-2 py-0.5 rounded-full font-medium" style={{ background: `color-mix(in srgb, ${status.color} 22%, transparent)`, color: status.color }}>{status.label}</span>
        </div>
        <p className="text-xs font-mono mt-0.5" style={{ color: MUTED }}>{book.file_name} · {fmtBytes(book.size)} · added {stamp.date}</p>
        <div className="flex items-center gap-3 mt-2.5">
          <div className="h-2 rounded-full overflow-hidden flex-1 max-w-xs" style={{ background: BORDER }}>
            <div className="h-full progress-fill rounded-full" style={{ width: `${pct}%`, background: pct >= 100 ? GREEN : GOLD }} />
          </div>
          <span className="text-xs font-mono w-28 shrink-0" style={{ color: MUTED }}>{pct}% · pg {book.last_read_page}/{book.total_pages}</span>
        </div>
      </div>
      <div className="flex items-center gap-1.5 shrink-0 self-end sm:self-center">
        <button onClick={onRead} className="btn-lift flex items-center gap-1.5 px-3 py-2 rounded-md text-sm font-medium text-white" style={{ background: color }}>
          <BookOpen size={15} /> Read
        </button>
        {onToggleFavorite && (
          <button onClick={onToggleFavorite} title={book.is_favorite ? "Remove from favorites" : "Add to favorites"} className="p-2 rounded-md hover:bg-black/5">
            <Star size={16} fill={book.is_favorite ? GOLD : "none"} style={{ color: book.is_favorite ? GOLD : MUTED }} />
          </button>
        )}
        <button onClick={onEdit} title="Edit details" className="p-2 rounded-md hover:bg-black/5" style={{ color: MUTED }}><Pencil size={16} /></button>
        {pct > 0 && (
          <button onClick={onReset} title="Reset reading progress to 0%" className="p-2 rounded-md hover:bg-black/5" style={{ color: MUTED }}><RotateCcw size={16} /></button>
        )}
        <a href={book.file_url} download={book.file_name} title="Download" className="p-2 rounded-md hover:bg-black/5" style={{ color: MUTED }}><Download size={17} /></a>
        <button onClick={onDelete} title="Move to Trash" className="p-2 rounded-md hover:bg-black/5" style={{ color: MUTED }}><Trash2 size={17} /></button>
      </div>
    </div>
  );
}

/* ============================== BOOK READER ============================== */
function NarrationBar({ state, rate, error, supported, onPlayPause, onStop, onSkip, onRateChange }) {
  if (!supported) {
    return <p className="text-center text-xs py-2" style={{ color: "#B7A88C" }}>Voice narration isn't supported in this browser.</p>;
  }
  return (
    <div className="flex flex-col items-center gap-1.5 py-2">
      <div className="flex items-center gap-3">
        <button onClick={() => onSkip(-1)} title="Previous page" className="p-2 rounded-md hover:bg-white/10 text-white"><SkipBack size={17} /></button>
        <button onClick={onPlayPause} title={state === "playing" ? "Pause narration" : "Play narration"}
          className="btn-lift h-10 w-10 rounded-full flex items-center justify-center" style={{ background: GOLD, color: LEATHER_DARK }}>
          {state === "playing" ? <Pause size={18} /> : <Play size={18} className="ml-0.5" />}
        </button>
        <button onClick={() => onSkip(1)} title="Next page" className="p-2 rounded-md hover:bg-white/10 text-white"><SkipForward size={17} /></button>
        {state !== "stopped" && (
          <button onClick={onStop} title="Stop narration" className="p-2 rounded-md hover:bg-white/10 text-white"><VolumeX size={16} /></button>
        )}
        <div className="flex items-center gap-1 ml-2">
          {[0.75, 1, 1.25, 1.5].map((r) => (
            <button key={r} onClick={() => onRateChange(r)}
              className="px-2 py-1 rounded-md text-xs font-mono"
              style={{ background: rate === r ? GOLD : "rgba(255,255,255,0.1)", color: rate === r ? LEATHER_DARK : "#F1E6D2" }}>
              {r}×
            </button>
          ))}
        </div>
      </div>
      <p className="text-xs flex items-center gap-1.5" style={{ color: state === "stopped" ? "#B7A88C" : GOLD }}>
        <Volume2 size={12} />
        {state === "playing" ? "Narrating this page…" : state === "paused" ? "Narration paused" : "Voice narration — play to start"}
      </p>
      {error && <p className="text-xs" style={{ color: "#E08A8A" }}>{error}</p>}
    </div>
  );
}

function BookReader({ book, onClose, onCommitProgress, onSaveEdit, onFixTotalPages }) {
  const isPdf = book.file_type === "pdf";
  const isDocx = book.file_type === "docx";
  const isLegacyDoc = book.file_type === "doc";

  const [pdfPage, setPdfPage] = useState(book.last_read_page || 1);

  const [pages, setPages] = useState(null); // array of REAL rendered page HTML strings
  const [pageIndex, setPageIndex] = useState(0);
  const [htmlError, setHtmlError] = useState("");
  const [editMode, setEditMode] = useState(false);
  const [isDirty, setIsDirty] = useState(false);
  const [fontScale, setFontScale] = useState(1);
  const contentRef = useRef(null);
  const fixedRef = useRef(false);

  const [stage, setStage] = useState("reading");
  const [manualPage, setManualPage] = useState("");

  // ---- Voice narration (Web Speech API — built into the browser, no API key) ----
  const speechSupported = typeof window !== "undefined" && "speechSynthesis" in window;
  const [narrationState, setNarrationState] = useState("stopped"); // stopped | playing | paused
  const [narrationRate, setNarrationRate] = useState(1);
  const [narrationError, setNarrationError] = useState("");
  const pdfDocRef = useRef(null);
  const narrationActiveRef = useRef(false);

  useEffect(() => {
    if (isPdf) {
      (async () => {
        try { pdfDocRef.current = await getDocument(book.file_url).promise; }
        catch (e) { pdfDocRef.current = null; }
      })();
    }
  }, [isPdf, book.file_url]);

  useEffect(() => () => { if (speechSupported) window.speechSynthesis.cancel(); }, []);

  const getPageText = async (pageArg) => {
    if (isPdf) {
      if (!pdfDocRef.current) return "";
      try {
        const page = await pdfDocRef.current.getPage(pageArg); // 1-indexed
        const content = await page.getTextContent();
        return content.items.map((it) => it.str).join(" ");
      } catch (e) { return ""; }
    }
    if (isDocx && pages) {
      const html = pages[pageArg] || ""; // 0-indexed
      try { return new DOMParser().parseFromString(html, "text/html").body.textContent || ""; }
      catch (e) { return ""; }
    }
    return "";
  };

  const speakPage = async (pageArg, rateOverride) => {
    setNarrationError("");
    const text = await getPageText(pageArg);
    if (!text || !text.trim()) { goToNextForNarration(pageArg); return; }
    window.speechSynthesis.cancel();
    const utter = new SpeechSynthesisUtterance(text);
    utter.rate = rateOverride != null ? rateOverride : narrationRate;
    utter.onend = () => { if (narrationActiveRef.current) goToNextForNarration(pageArg); };
    utter.onerror = () => { setNarrationError("Narration was interrupted."); };
    setNarrationState("playing");
    window.speechSynthesis.speak(utter);
  };

  const goToNextForNarration = (fromPage) => {
    const atEnd = isPdf ? fromPage >= book.total_pages : fromPage >= ((pages ? pages.length : 1) - 1);
    if (atEnd) { narrationActiveRef.current = false; setNarrationState("stopped"); return; }
    const next = fromPage + 1;
    if (isPdf) setPdfPage(next); else setPageIndex(next);
    speakPage(next);
  };

  const handleNarrationPlayPause = () => {
    if (!speechSupported) { setNarrationError("Voice narration isn't supported in this browser."); return; }
    if (narrationState === "stopped") {
      narrationActiveRef.current = true;
      speakPage(isPdf ? pdfPage : pageIndex);
    } else if (narrationState === "playing") {
      window.speechSynthesis.pause();
      setNarrationState("paused");
    } else if (narrationState === "paused") {
      window.speechSynthesis.resume();
      setNarrationState("playing");
    }
  };
  const handleNarrationStop = () => {
    narrationActiveRef.current = false;
    if (speechSupported) window.speechSynthesis.cancel();
    setNarrationState("stopped");
  };
  const handleNarrationSkip = (dir) => {
    if (speechSupported) window.speechSynthesis.cancel();
    const current = isPdf ? pdfPage : pageIndex;
    const min = isPdf ? 1 : 0;
    const max = isPdf ? book.total_pages : ((pages ? pages.length : 1) - 1);
    const target = Math.max(min, Math.min(max, current + dir));
    if (isPdf) setPdfPage(target); else setPageIndex(target);
    if (narrationState !== "stopped") { narrationActiveRef.current = true; speakPage(target); }
  };
  const changeNarrationRate = (r) => {
    setNarrationRate(r);
    if (narrationState === "playing") { narrationActiveRef.current = true; speakPage(isPdf ? pdfPage : pageIndex, r); }
  };

  useEffect(() => {
    if (isDocx) {
      (async () => {
        try {
          let realPages = null;
          if (book.edited_html) {
            try {
              const maybe = JSON.parse(book.edited_html);
              if (Array.isArray(maybe) && maybe.length) realPages = maybe;
            } catch (e) { /* fall through to a fresh render */ }
          }
          if (!realPages) {
            const res = await fetch(book.file_url);
            const buf = await res.arrayBuffer();
            realPages = await renderDocxToPages(buf);
          }
          setPages(realPages);
          const startIdx = Math.max(0, Math.min(realPages.length - 1, (book.last_read_page || 1) - 1));
          setPageIndex(startIdx);
          // Self-heal: this is now the real, authoritative page count. If the
          // stored total doesn't match it (e.g. from an older/incorrect entry
          // or a manual override that didn't match reality), correct it once.
          if (!fixedRef.current && realPages.length && realPages.length !== book.total_pages) {
            fixedRef.current = true;
            onFixTotalPages(realPages.length);
          }
        } catch (e) {
          setHtmlError("This document could not be previewed. You can still download it.");
        }
      })();
    }
  }, [isDocx, book.file_url, book.edited_html]);

  useEffect(() => {
    if (isDocx && pages && contentRef.current) {
      contentRef.current.innerHTML = pages[pageIndex] || "";
    }
  }, [isDocx, pages, pageIndex]);

  const totalPages = isDocx && pages ? pages.length : book.total_pages;

  const handleInput = () => { if (!isDirty) setIsDirty(true); };

  const enterEdit = () => { handleNarrationStop(); setEditMode(true); };
  const saveEdit = () => {
    const newHtml = contentRef.current ? contentRef.current.innerHTML : pages[pageIndex];
    const nextPages = pages.map((p, i) => (i === pageIndex ? newHtml : p));
    setPages(nextPages);
    onSaveEdit(JSON.stringify(nextPages));
    setIsDirty(false);
    setEditMode(false);
  };
  const discardEdit = () => {
    if (contentRef.current) contentRef.current.innerHTML = pages[pageIndex];
    setIsDirty(false);
    setEditMode(false);
  };

  const goPrev = () => { if (!editMode) setPageIndex((p) => Math.max(0, p - 1)); };
  const goNext = () => { if (!editMode && pages) setPageIndex((p) => Math.min(pages.length - 1, p + 1)); };

  const currentPage = isPdf ? pdfPage : pageIndex + 1;
  const requestClose = () => {
    handleNarrationStop();
    if (isDocx && editMode && isDirty) { setStage("discardCheck"); return; }
    setStage("confirmPage");
  };
  const finalizeClose = (page) => { onCommitProgress(page); onClose(); };

  return (
    <div className="fixed inset-0 z-50 flex flex-col fade" style={{ background: "rgba(20,18,16,0.96)" }}>
      <div className="flex items-center justify-between px-5 py-3 shrink-0" style={{ background: LEATHER_DARK, borderBottom: `2px solid ${GOLD}` }}>
        <div className="min-w-0">
          <p className="font-serif font-bold text-white truncate">{book.name}</p>
          <p className="text-xs font-mono" style={{ color: "#D8C7A8" }}>{book.file_name}</p>
        </div>
        <div className="flex items-center gap-2 shrink-0">
          {isDocx && !isLegacyDoc && !editMode && (
            <>
              <button onClick={() => setFontScale((s) => Math.max(0.7, +(s - 0.1).toFixed(2)))} title="Zoom out" className="p-2 rounded-md hover:bg-white/10 text-white"><ZoomOut size={16} /></button>
              <button onClick={() => setFontScale((s) => Math.min(1.6, +(s + 0.1).toFixed(2)))} title="Zoom in" className="p-2 rounded-md hover:bg-white/10 text-white"><ZoomIn size={16} /></button>
              <button onClick={enterEdit} className="btn-lift flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium" style={{ background: GOLD, color: LEATHER_DARK }}>
                <Pencil size={14} /> Edit
              </button>
            </>
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
          Editing enabled — edited text appears in <span style={{ color: GREEN, fontWeight: 700 }}>green</span> until you save. Page navigation is paused while editing.
        </div>
      )}

      <div className="flex-1 overflow-hidden flex flex-col items-center px-4 py-4">
        {isPdf && (
          <div className="w-full h-full max-w-4xl flex flex-col">
            <div className="flex-1 rounded-lg overflow-hidden shadow-2xl" style={{ background: "white" }}>
              <embed src={book.file_url} type="application/pdf" className="w-full h-full" />
            </div>
            <div className="flex items-center justify-center gap-4 py-3">
              <button onClick={() => setPdfPage((p) => Math.max(1, p - 1))} className="p-2 rounded-md text-white hover:bg-white/10"><ChevronLeft size={18} /></button>
              <div className="flex items-center gap-2 text-sm" style={{ color: "#F1E6D2" }}>
                Page
                <input type="number" value={pdfPage} min={1} max={book.total_pages}
                  onChange={(e) => setPdfPage(Math.min(book.total_pages, Math.max(1, parseInt(e.target.value, 10) || 1)))}
                  className="w-16 text-center rounded-md px-2 py-1 text-black" />
                of {book.total_pages}
              </div>
              <button onClick={() => setPdfPage((p) => Math.min(book.total_pages, p + 1))} className="p-2 rounded-md text-white hover:bg-white/10"><ChevronRight size={18} /></button>
            </div>
            <NarrationBar state={narrationState} rate={narrationRate} error={narrationError} supported={speechSupported}
              onPlayPause={handleNarrationPlayPause} onStop={handleNarrationStop}
              onSkip={handleNarrationSkip} onRateChange={changeNarrationRate} />
            <p className="text-center text-xs pb-1" style={{ color: "#B7A88C" }}>The PDF above is your real, unmodified file. Use the controls to bookmark where you are.</p>
          </div>
        )}

        {isDocx && (
          <div className="w-full h-full flex flex-col">
            {htmlError ? (
              <div className="flex-1 flex items-center justify-center text-center px-6" style={{ color: "#F1E6D2" }}>
                <div><AlertCircle className="mx-auto mb-3" /><p>{htmlError}</p></div>
              </div>
            ) : pages === null ? (
              <div className="flex-1 flex items-center justify-center gap-2" style={{ color: "white" }}><Loader2 className="animate-spin" /> Rendering your document's real pages…</div>
            ) : (
              <div className="flex-1 overflow-auto flex items-start justify-center py-6" style={{ background: "#3A3630", borderRadius: 10 }}>
                <div key={pageIndex} className="fade" style={{
                  transform: `scale(${fontScale})`, transformOrigin: "top center",
                  boxShadow: "0 12px 34px rgba(0,0,0,0.45)",
                }}>
                  <div
                    ref={contentRef}
                    data-editable="true"
                    contentEditable={editMode}
                    suppressContentEditableWarning
                    onInput={handleInput}
                    className={isDirty ? "editing-live" : ""}
                    style={{ outline: "none", background: "white" }}
                  />
                </div>
              </div>
            )}
            <div className="flex items-center justify-center gap-4 py-3">
              <button onClick={goPrev} disabled={editMode || pageIndex === 0} className="p-2 rounded-md text-white hover:bg-white/10 disabled:opacity-30"><ChevronLeft size={18} /></button>
              <div className="flex items-center gap-2 text-sm" style={{ color: "#F1E6D2" }}>
                Page
                <input type="number" value={pageIndex + 1} min={1} max={totalPages} disabled={editMode}
                  onChange={(e) => setPageIndex(Math.min(totalPages - 1, Math.max(0, (parseInt(e.target.value, 10) || 1) - 1)))}
                  className="w-16 text-center rounded-md px-2 py-1 text-black disabled:opacity-60" />
                of {totalPages}
              </div>
              <button onClick={goNext} disabled={editMode || !pages || pageIndex >= pages.length - 1} className="p-2 rounded-md text-white hover:bg-white/10 disabled:opacity-30"><ChevronRight size={18} /></button>
            </div>
            {!editMode && (
              <NarrationBar state={narrationState} rate={narrationRate} error={narrationError} supported={speechSupported}
                onPlayPause={handleNarrationPlayPause} onStop={handleNarrationStop}
                onSkip={handleNarrationSkip} onRateChange={changeNarrationRate} />
            )}
            <p className="text-center text-xs pb-1" style={{ color: "#B7A88C" }}>Page count is measured from how this document renders in your browser — it may differ slightly from Word's own count. For exact page numbers, upload as PDF instead.</p>
          </div>
        )}

        {isLegacyDoc && (
          <div className="flex-1 flex items-center justify-center text-center px-6" style={{ color: "#F1E6D2" }}>
            <div>
              <FileText className="mx-auto mb-3" size={28} />
              <p className="mb-3">Legacy .doc files can't be previewed in-browser.</p>
              <a href={book.file_url} download={book.file_name} className="underline text-sm">Download to view</a>
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
            You've read up to page <span className="font-semibold" style={{ color: LEATHER_DARK }}>{currentPage}</span> of <span className="font-semibold" style={{ color: LEATHER_DARK }}>{totalPages}</span>. Confirm?
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
          <input type="number" min={0} max={totalPages} autoFocus value={manualPage} onChange={(e) => setManualPage(e.target.value)}
            className="w-full border rounded-md px-3 py-2 mb-4 focus:outline-none focus:ring-2" style={{ borderColor: BORDER }} />
          <div className="flex justify-end gap-3">
            <button onClick={() => setStage("confirmPage")} className="px-4 py-2 rounded-md text-sm" style={{ color: MUTED }}>Back</button>
            <button onClick={() => finalizeClose(Math.min(totalPages, Math.max(0, parseInt(manualPage, 10) || 0)))}
              className="btn-lift px-4 py-2 rounded-md text-sm font-medium text-white flex items-center gap-2" style={{ background: GREEN }}><Check size={14} /> Save progress</button>
          </div>
        </Modal>
      )}
    </div>
  );
}
