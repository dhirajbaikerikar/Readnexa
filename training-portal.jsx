import React, { useState, useEffect, useRef, useCallback } from "react";
import { FolderOpen, Folder, Upload, FileText, File, Download, Trash2, Plus, ArrowLeft, X, Check, AlertCircle, Loader2 } from "lucide-react";

const FOLDER_TABS = ["#C98A3E", "#3C6E52", "#7A4A9C", "#2E75B6", "#B5473E"];
const MAX_FILE_BYTES = 3 * 1024 * 1024; // 3MB raw per file (storage-safe after base64 inflation)

function tabColor(seed) {
  let h = 0;
  for (let i = 0; i < seed.length; i++) h = (h * 31 + seed.charCodeAt(i)) >>> 0;
  return FOLDER_TABS[h % FOLDER_TABS.length];
}

function fmtBytes(b) {
  if (b < 1024) return b + " B";
  if (b < 1024 * 1024) return (b / 1024).toFixed(0) + " KB";
  return (b / (1024 * 1024)).toFixed(1) + " MB";
}

function fmtStamp(iso) {
  const d = new Date(iso);
  const date = d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
  const time = d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
  return { date, time };
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

function downloadCourse(course) {
  try {
    const mime =
      course.fileType === "pdf" ? "application/pdf" :
      course.fileType === "doc" ? "application/msword" :
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    const bin = atob(course.dataUrl);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const blob = new Blob([bytes], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${course.name}.${course.fileType}`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  } catch (e) {
    console.error("Download failed", e);
  }
}

const DEFAULT_FOLDERS = [
  { id: "lean-six-sigma", name: "Lean Six Sigma", createdAt: new Date().toISOString(), isDefault: true },
  { id: "pmp", name: "PMP", createdAt: new Date().toISOString(), isDefault: true },
];

export default function TrainingPortal() {
  const [ready, setReady] = useState(false);
  const [folders, setFolders] = useState([]);
  const [coursesByFolder, setCoursesByFolder] = useState({});
  const [loadingFolder, setLoadingFolder] = useState(false);
  const [view, setView] = useState("home");
  const [activeFolderId, setActiveFolderId] = useState(null);

  const [showNewFolder, setShowNewFolder] = useState(false);
  const [newFolderName, setNewFolderName] = useState("");

  const [pendingFile, setPendingFile] = useState(null); // {file, base64, ext}
  const [pendingName, setPendingName] = useState("");
  const [uploadError, setUploadError] = useState("");
  const [saving, setSaving] = useState(false);
  const fileInputRef = useRef(null);

  const [confirmDelete, setConfirmDelete] = useState(null); // {type:'folder'|'course', id, folderId}
  const [toast, setToast] = useState("");

  // ---------- Load folder index on mount ----------
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("folders-index", true);
        if (res && res.value) {
          setFolders(JSON.parse(res.value));
        } else {
          await window.storage.set("folders-index", JSON.stringify(DEFAULT_FOLDERS), true);
          setFolders(DEFAULT_FOLDERS);
        }
      } catch (e) {
        setFolders(DEFAULT_FOLDERS);
      }
      setReady(true);
    })();
  }, []);

  const showToast = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(""), 2600);
  };

  const saveFolders = async (next) => {
    setFolders(next);
    try {
      await window.storage.set("folders-index", JSON.stringify(next), true);
    } catch (e) {
      showToast("Could not save folder list. Please try again.");
    }
  };

  const loadFolderCourses = useCallback(async (folderId) => {
    if (coursesByFolder[folderId]) return;
    setLoadingFolder(true);
    try {
      const res = await window.storage.get(`courses:${folderId}`, true);
      setCoursesByFolder((prev) => ({ ...prev, [folderId]: res && res.value ? JSON.parse(res.value) : [] }));
    } catch (e) {
      setCoursesByFolder((prev) => ({ ...prev, [folderId]: [] }));
    }
    setLoadingFolder(false);
  }, [coursesByFolder]);

  const openFolder = (id) => {
    setActiveFolderId(id);
    setView("folder");
    loadFolderCourses(id);
  };

  const saveCourses = async (folderId, list) => {
    setCoursesByFolder((prev) => ({ ...prev, [folderId]: list }));
    try {
      const result = await window.storage.set(`courses:${folderId}`, JSON.stringify(list), true);
      if (!result) throw new Error("no result");
    } catch (e) {
      showToast("Storage error — the file may not have saved. Try a smaller file.");
    }
  };

  // ---------- New folder ----------
  const createFolder = async () => {
    const name = newFolderName.trim();
    if (!name) return;
    const id = name.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") + "-" + Date.now().toString(36);
    const next = [...folders, { id, name, createdAt: new Date().toISOString(), isDefault: false }];
    await saveFolders(next);
    setShowNewFolder(false);
    setNewFolderName("");
    showToast(`Folder "${name}" created`);
  };

  const deleteFolder = async (id) => {
    const next = folders.filter((f) => f.id !== id);
    await saveFolders(next);
    try { await window.storage.delete(`courses:${id}`, true); } catch (e) {}
    setCoursesByFolder((prev) => {
      const copy = { ...prev };
      delete copy[id];
      return copy;
    });
    if (activeFolderId === id) { setView("home"); setActiveFolderId(null); }
    setConfirmDelete(null);
    showToast("Folder deleted");
  };

  // ---------- Upload flow ----------
  const onPickFile = () => fileInputRef.current && fileInputRef.current.click();

  const onFileSelected = async (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    setUploadError("");
    const ext = extOf(file.name);
    if (!["pdf", "doc", "docx"].includes(ext)) {
      setUploadError("Only PDF and Word (.doc, .docx) files are supported.");
      return;
    }
    if (file.size > MAX_FILE_BYTES) {
      setUploadError(`File is too large (${fmtBytes(file.size)}). Please keep files under ${fmtBytes(MAX_FILE_BYTES)}.`);
      return;
    }
    try {
      const base64 = await fileToBase64(file);
      setPendingFile({ file, base64, ext, size: file.size });
      setPendingName(cleanBaseName(file.name));
    } catch (err) {
      setUploadError("Could not read this file. Please try again.");
    }
  };

  const cancelUpload = () => {
    setPendingFile(null);
    setPendingName("");
    setUploadError("");
  };

  const confirmUploadSave = async () => {
    if (!pendingFile || !pendingName.trim()) return;
    setSaving(true);
    const existing = coursesByFolder[activeFolderId] || [];
    const combinedSize = existing.reduce((s, c) => s + (c.size || 0), 0) + pendingFile.size;
    if (combinedSize > 4.5 * 1024 * 1024) {
      setUploadError("This folder is close to its storage limit. Try removing an older file first, or use a smaller file.");
      setSaving(false);
      return;
    }
    const entry = {
      id: "c_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      name: pendingName.trim(),
      fileName: pendingFile.file.name,
      fileType: pendingFile.ext,
      size: pendingFile.size,
      uploadedAt: new Date().toISOString(),
      dataUrl: pendingFile.base64,
    };
    const next = [entry, ...existing];
    await saveCourses(activeFolderId, next);
    setSaving(false);
    cancelUpload();
    showToast(`"${entry.name}" uploaded`);
  };

  const deleteCourse = async (courseId) => {
    const next = (coursesByFolder[activeFolderId] || []).filter((c) => c.id !== courseId);
    await saveCourses(activeFolderId, next);
    setConfirmDelete(null);
    showToast("Course file removed");
  };

  const activeFolder = folders.find((f) => f.id === activeFolderId);
  const activeCourses = coursesByFolder[activeFolderId] || [];

  // ---------- Render ----------
  return (
    <div className="min-h-screen w-full" style={{ background: "#EFEDE6" }}>
      <input ref={fileInputRef} type="file" accept=".pdf,.doc,.docx" className="hidden" onChange={onFileSelected} />

      {/* Header */}
      <header className="border-b" style={{ borderColor: "#D9D4C7", background: "#16323A" }}>
        <div className="max-w-5xl mx-auto px-6 py-6 flex items-center justify-between">
          <div>
            <p className="uppercase tracking-widest text-xs font-mono" style={{ color: "#C98A3E" }}>Learning &amp; Development</p>
            <h1 className="font-serif text-3xl font-bold text-white mt-1">Training Library</h1>
          </div>
          <div className="hidden sm:flex flex-col items-end text-right">
            <span className="text-white/70 text-xs font-mono">{folders.length} course folder{folders.length !== 1 ? "s" : ""}</span>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-10">
        {!ready ? (
          <div className="flex items-center gap-2 text-[#5B6360] py-20 justify-center">
            <Loader2 className="animate-spin" size={20} /> Loading training library…
          </div>
        ) : view === "home" ? (
          <HomeView
            folders={folders}
            onOpen={openFolder}
            onNewFolder={() => setShowNewFolder(true)}
            onDeleteFolder={(f) => setConfirmDelete({ type: "folder", id: f.id, name: f.name })}
          />
        ) : (
          <FolderView
            folder={activeFolder}
            courses={activeCourses}
            loading={loadingFolder}
            onBack={() => { setView("home"); setActiveFolderId(null); }}
            onUploadClick={onPickFile}
            onDownload={downloadCourse}
            onDeleteCourse={(c) => setConfirmDelete({ type: "course", id: c.id, name: c.name })}
          />
        )}
      </main>

      {/* New folder modal */}
      {showNewFolder && (
        <Modal onClose={() => { setShowNewFolder(false); setNewFolderName(""); }}>
          <h3 className="font-serif text-xl font-bold mb-4" style={{ color: "#16323A" }}>New course folder</h3>
          <label className="text-sm text-[#5B6360] block mb-1">Folder name</label>
          <input
            autoFocus
            value={newFolderName}
            onChange={(e) => setNewFolderName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && createFolder()}
            placeholder="e.g. Business Analysis"
            className="w-full border rounded-md px-3 py-2 mb-5 focus:outline-none focus:ring-2"
            style={{ borderColor: "#D9D4C7", "--tw-ring-color": "#C98A3E" }}
          />
          <div className="flex justify-end gap-3">
            <button onClick={() => { setShowNewFolder(false); setNewFolderName(""); }} className="px-4 py-2 rounded-md text-sm text-[#5B6360] hover:bg-black/5">Cancel</button>
            <button onClick={createFolder} disabled={!newFolderName.trim()}
              className="px-4 py-2 rounded-md text-sm font-medium text-white disabled:opacity-40"
              style={{ background: "#16323A" }}>Create folder</button>
          </div>
        </Modal>
      )}

      {/* Upload naming modal */}
      {pendingFile && (
        <Modal onClose={cancelUpload}>
          <h3 className="font-serif text-xl font-bold mb-1" style={{ color: "#16323A" }}>Name this course</h3>
          <p className="text-sm text-[#5B6360] mb-4">
            File: <span className="font-mono">{pendingFile.file.name}</span> ({fmtBytes(pendingFile.size)})
          </p>
          <label className="text-sm text-[#5B6360] block mb-1">Course name</label>
          <input
            autoFocus
            value={pendingName}
            onChange={(e) => setPendingName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && confirmUploadSave()}
            className="w-full border rounded-md px-3 py-2 mb-2 focus:outline-none focus:ring-2"
            style={{ borderColor: "#D9D4C7", "--tw-ring-color": "#C98A3E" }}
          />
          {uploadError && (
            <p className="text-sm flex items-center gap-1.5 mb-2" style={{ color: "#B5473E" }}>
              <AlertCircle size={14} /> {uploadError}
            </p>
          )}
          <p className="text-xs text-[#8A8F8C] mb-5">This name is what appears in the training library. The original file name is kept for reference.</p>
          <div className="flex justify-end gap-3">
            <button onClick={cancelUpload} className="px-4 py-2 rounded-md text-sm text-[#5B6360] hover:bg-black/5">Cancel</button>
            <button onClick={confirmUploadSave} disabled={!pendingName.trim() || saving}
              className="px-4 py-2 rounded-md text-sm font-medium text-white disabled:opacity-40 flex items-center gap-2"
              style={{ background: "#16323A" }}>
              {saving ? <Loader2 className="animate-spin" size={14} /> : <Check size={14} />}
              Save to folder
            </button>
          </div>
        </Modal>
      )}

      {/* Delete confirmation */}
      {confirmDelete && (
        <Modal onClose={() => setConfirmDelete(null)}>
          <h3 className="font-serif text-xl font-bold mb-2" style={{ color: "#16323A" }}>
            Delete {confirmDelete.type === "folder" ? "folder" : "file"}?
          </h3>
          <p className="text-sm text-[#5B6360] mb-6">
            "{confirmDelete.name}" will be permanently removed{confirmDelete.type === "folder" ? ", along with every course file inside it" : ""}. This cannot be undone.
          </p>
          <div className="flex justify-end gap-3">
            <button onClick={() => setConfirmDelete(null)} className="px-4 py-2 rounded-md text-sm text-[#5B6360] hover:bg-black/5">Cancel</button>
            <button
              onClick={() => confirmDelete.type === "folder" ? deleteFolder(confirmDelete.id) : deleteCourse(confirmDelete.id)}
              className="px-4 py-2 rounded-md text-sm font-medium text-white flex items-center gap-2"
              style={{ background: "#B5473E" }}>
              <Trash2 size={14} /> Delete
            </button>
          </div>
        </Modal>
      )}

      {/* Toast */}
      {toast && (
        <div className="fixed bottom-6 left-1/2 -translate-x-1/2 px-4 py-2.5 rounded-lg text-sm text-white shadow-lg flex items-center gap-2 z-50" style={{ background: "#16323A" }}>
          <Check size={14} style={{ color: "#C98A3E" }} /> {toast}
        </div>
      )}

      <p className="text-center text-xs text-[#8A8F8C] font-mono pb-8">
        Shared library — folders and course files here are visible to everyone using this portal.
      </p>
    </div>
  );
}

function Modal({ children, onClose }) {
  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 px-4" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-md p-6" onClick={(e) => e.stopPropagation()}>
        {children}
      </div>
    </div>
  );
}

function HomeView({ folders, onOpen, onNewFolder, onDeleteFolder }) {
  return (
    <div>
      <div className="flex items-end justify-between mb-6">
        <div>
          <h2 className="font-serif text-xl font-bold" style={{ color: "#16323A" }}>Course folders</h2>
          <p className="text-sm text-[#5B6360] mt-0.5">Open a folder to view or add training materials.</p>
        </div>
        <button onClick={onNewFolder}
          className="flex items-center gap-1.5 px-4 py-2 rounded-md text-sm font-medium text-white shrink-0"
          style={{ background: "#16323A" }}>
          <Plus size={16} /> New folder
        </button>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-5">
        {folders.map((f) => (
          <FolderCard key={f.id} folder={f} onOpen={() => onOpen(f.id)} onDelete={() => onDeleteFolder(f)} />
        ))}
      </div>

      {folders.length === 0 && (
        <p className="text-center text-[#8A8F8C] py-16">No folders yet. Create one to get started.</p>
      )}
    </div>
  );
}

function FolderCard({ folder, onOpen, onDelete }) {
  const color = tabColor(folder.id);
  const [hover, setHover] = useState(false);
  return (
    <div
      className="relative cursor-pointer group"
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      onClick={onOpen}
    >
      <div className="absolute -top-2.5 left-4 h-3 w-16 rounded-t-md" style={{ background: color }} />
      <div className="relative bg-white rounded-md rounded-tl-none border shadow-sm p-5 pt-6 transition-transform group-hover:-translate-y-0.5 group-hover:shadow-md"
        style={{ borderColor: "#D9D4C7" }}>
        <div className="flex items-start justify-between">
          <FolderOpen size={26} style={{ color }} />
          {!folder.isDefault && hover && (
            <button
              onClick={(e) => { e.stopPropagation(); onDelete(); }}
              className="text-[#8A8F8C] hover:text-[#B5473E] p-1 -m-1"
              title="Delete folder"
            >
              <Trash2 size={15} />
            </button>
          )}
        </div>
        <h3 className="font-serif font-bold text-base mt-3" style={{ color: "#16323A" }}>{folder.name}</h3>
        <p className="text-xs text-[#8A8F8C] font-mono mt-1">Added {fmtStamp(folder.createdAt).date}</p>
      </div>
    </div>
  );
}

function FolderView({ folder, courses, loading, onBack, onUploadClick, onDownload, onDeleteCourse }) {
  if (!folder) return null;
  const color = tabColor(folder.id);
  return (
    <div>
      <button onClick={onBack} className="flex items-center gap-1.5 text-sm text-[#5B6360] hover:text-[#16323A] mb-5">
        <ArrowLeft size={15} /> All folders
      </button>

      <div className="flex items-end justify-between mb-6">
        <div className="flex items-center gap-3">
          <FolderOpen size={30} style={{ color }} />
          <div>
            <h2 className="font-serif text-2xl font-bold" style={{ color: "#16323A" }}>{folder.name}</h2>
            <p className="text-sm text-[#5B6360]">{courses.length} course file{courses.length !== 1 ? "s" : ""}</p>
          </div>
        </div>
        <button onClick={onUploadClick}
          className="flex items-center gap-1.5 px-4 py-2 rounded-md text-sm font-medium text-white shrink-0"
          style={{ background: "#16323A" }}>
          <Upload size={16} /> Upload course
        </button>
      </div>

      {loading ? (
        <div className="flex items-center gap-2 text-[#5B6360] py-16 justify-center">
          <Loader2 className="animate-spin" size={18} /> Loading files…
        </div>
      ) : courses.length === 0 ? (
        <div className="text-center py-16 border-2 border-dashed rounded-lg" style={{ borderColor: "#D9D4C7" }}>
          <File size={30} className="mx-auto mb-3 text-[#C9C4B7]" />
          <p className="text-[#5B6360] mb-4">No course files here yet.</p>
          <button onClick={onUploadClick} className="text-sm font-medium underline" style={{ color: "#16323A" }}>
            Upload the first PDF or Word file
          </button>
        </div>
      ) : (
        <ul className="divide-y rounded-lg border bg-white" style={{ borderColor: "#D9D4C7" }}>
          {courses.map((c) => {
            const stamp = fmtStamp(c.uploadedAt);
            return (
              <li key={c.id} className="flex items-center gap-4 px-5 py-4">
                <FileText size={22} style={{ color }} className="shrink-0" />
                <div className="min-w-0 flex-1">
                  <p className="font-medium truncate" style={{ color: "#16323A" }}>{c.name}</p>
                  <p className="text-xs text-[#8A8F8C] font-mono truncate">
                    {c.fileName} · {fmtBytes(c.size)} · uploaded {stamp.date} at {stamp.time}
                  </p>
                </div>
                <button onClick={() => onDownload(c)} title="Download" className="p-2 rounded-md hover:bg-black/5 text-[#5B6360] shrink-0">
                  <Download size={17} />
                </button>
                <button onClick={() => onDeleteCourse(c)} title="Delete" className="p-2 rounded-md hover:bg-black/5 text-[#8A8F8C] hover:text-[#B5473E] shrink-0">
                  <Trash2 size={17} />
                </button>
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
}
