import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { Truck, Plus, Download, Upload, X, Trash2, Link2, Github, RefreshCw, Settings } from "lucide-react";

// ---------- Design tokens ----------
const C = {
  bg: "#FFFFFF",
  surface: "#F5F6F8",
  surfaceAlt: "#ECEEF1",
  surfaceHover: "#E2E5EA",
  border: "#DADDE2",
  borderLight: "#C4C9D0",
  text: "#14181C",
  textDim: "#5C6675",
  textFaint: "#8A94A3",
  amber: "#8F5808",
  amberDim: "#FBE8C7",
  green: "#1E8A4C",
  greenDim: "#DDF2E3",
  red: "#C0392B",
  redDim: "#FADBD8",
  blue: "#2B6CB0",
};

const FONT_DISPLAY = "'Space Grotesk', sans-serif";
const FONT_BODY = "'Inter', sans-serif";
const FONT_MONO = "'JetBrains Mono', monospace";

const STORAGE_KEY = "fleet-ledger-data";
const GITHUB_CONFIG_KEY = "fleet-ledger-github-config";

// GitHub is used as the shared "source of truth" data file so the same
// data shows up on any device. The config (owner/repo/branch/path/token)
// is kept only in this browser's localStorage - it is a secret and is
// never written into data.json or exported anywhere.
function loadGithubConfig() {
  try {
    const raw = localStorage.getItem(GITHUB_CONFIG_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function saveGithubConfigLocal(cfg) {
  localStorage.setItem(GITHUB_CONFIG_KEY, JSON.stringify(cfg));
}

function clearGithubConfigLocal() {
  localStorage.removeItem(GITHUB_CONFIG_KEY);
}

async function githubGetFile(cfg) {
  const url = `https://api.github.com/repos/${cfg.owner}/${cfg.repo}/contents/${cfg.path}?ref=${encodeURIComponent(cfg.branch)}`;
  const res = await fetch(url, {
    headers: {
      Authorization: `token ${cfg.token}`,
      Accept: "application/vnd.github+json",
    },
  });
  if (res.status === 404) return { exists: false, sha: null, data: null };
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(`GitHub read failed (${res.status}): ${body.message || res.statusText}`);
  }
  const json = await res.json();
  const decoded = decodeURIComponent(escape(atob(json.content.replace(/\n/g, ""))));
  return { exists: true, sha: json.sha, data: JSON.parse(decoded) };
}

async function githubPutFile(cfg, dataObj, sha) {
  const content = btoa(unescape(encodeURIComponent(JSON.stringify(dataObj, null, 2))));
  const res = await fetch(`https://api.github.com/repos/${cfg.owner}/${cfg.repo}/contents/${cfg.path}`, {
    method: "PUT",
    headers: {
      Authorization: `token ${cfg.token}`,
      Accept: "application/vnd.github+json",
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message: "Update fleet data",
      content,
      branch: cfg.branch,
      ...(sha ? { sha } : {}),
    }),
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(`GitHub write failed (${res.status}): ${body.message || res.statusText}`);
  }
  return res.json();
}


// Standalone (non-Claude) storage shim backed by the browser's localStorage,
// matching the get/set(key, value) shape the app uses. Data lives on this
// device/browser only - use Export/Import .xlsx to move it between devices.
const storage = {
  async get(key) {
    try {
      const v = localStorage.getItem(key);
      return v !== null ? { value: v } : null;
    } catch (e) {
      return null;
    }
  },
  async set(key, value) {
    localStorage.setItem(key, value);
    return { value };
  },
};

const uid = () => (crypto?.randomUUID ? crypto.randomUUID() : "id-" + Math.random().toString(36).slice(2));

// Calendar week number using the standard US convention: weeks run
// Sunday through Saturday, and week 1 is the week containing Jan 1.
// This matches Sun-Sat dispatch/billing weeks, so a period like
// 8/23/2026-8/29/2026 (a full Sun-Sat week) lands on a single week number.
function calendarWeek(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr + "T00:00:00");
  if (isNaN(d.getTime())) return "";
  const jan1 = new Date(d.getFullYear(), 0, 1);
  const jan1WeekdaySun0 = jan1.getDay(); // 0=Sun..6=Sat already
  const daysSinceJan1 = Math.round((d - jan1) / (24 * 3600 * 1000));
  return Math.floor((daysSinceJan1 + jan1WeekdaySun0) / 7) + 1;
}

function num(v) {
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function money(n) {
  return n.toLocaleString(undefined, { style: "currency", currency: "USD" });
}

const emptyBlock = () => ({ id: uid(), qty: "1", price: "", day: "", time: "" });

const emptyContract = (truckId) => ({
  id: uid(),
  truckId,
  contactId: "",
  periodStart: "",
  periodEnd: "",
  blocks: [emptyBlock(), emptyBlock(), emptyBlock()],
  basePricePerWeek: "",
  otherPrice: "",
  surcharge: "",
  paymentDate: "",
  paymentStatus: "Pending",
});

function payoutPerWeek(c) {
  return (c.blocks || []).reduce((sum, b) => sum + num(b.price), 0);
}

// Upgrades contracts saved by earlier versions of this app (which used
// block1/block2/block3 flat fields and a single "period" date) into the
// current shape (a blocks[] array of {qty, price}, periodStart/periodEnd).
function migrateContract(c) {
  if (Array.isArray(c.blocks) && c.blocks.length > 0) {
    // Already current shape - just make sure every block has an id/qty.
    return {
      ...c,
      blocks: c.blocks.map((b) => ({
        id: b.id || uid(),
        qty: b.qty ?? "1",
        price: b.price ?? "",
        note: b.note ?? "",
        day: b.day ?? "",
        time: b.time ?? "",
      })),
      periodStart: c.periodStart ?? c.period ?? "",
      periodEnd: c.periodEnd ?? "",
      basePricePerWeek: c.basePricePerWeek ?? "",
    };
  }

  // Legacy shape: block1 / block2 / block3 as flat price fields.
  const legacyBlocks = [c.block1, c.block2, c.block3]
    .filter((v) => v !== undefined)
    .map((price) => ({ id: uid(), qty: "1", price: price ?? "", note: "", day: "", time: "" }));

  return {
    id: c.id || uid(),
    truckId: c.truckId,
    contactId: c.contactId || "",
    periodStart: c.periodStart ?? c.period ?? "",
    periodEnd: c.periodEnd ?? "",
    blocks: legacyBlocks.length > 0 ? legacyBlocks : [emptyBlock()],
    basePricePerWeek: c.basePricePerWeek ?? "",
    otherPrice: c.otherPrice ?? "",
    surcharge: c.surcharge ?? "",
    paymentDate: c.paymentDate ?? "",
    paymentStatus: c.paymentStatus ?? "Pending",
  };
}

// ---------- Small UI atoms ----------
function Cell({ children, width, align = "left" }) {
  return (
    <div
      style={{
        width,
        flexShrink: 0,
        padding: "0 6px",
        display: "flex",
        alignItems: "center",
        justifyContent: align === "right" ? "flex-end" : "flex-start",
      }}
    >
      {children}
    </div>
  );
}

function EditableInput({ value, onChange, placeholder, type = "text", mono = true, align = "left" }) {
  return (
    <input
      value={value}
      onChange={(e) => onChange(e.target.value)}
      placeholder={placeholder}
      type={type}
      style={{
        width: "100%",
        background: "transparent",
        border: "none",
        outline: "none",
        color: C.text,
        fontFamily: mono ? FONT_MONO : FONT_BODY,
        fontSize: 12.5,
        padding: "8px 4px",
        textAlign: align,
      }}
      onFocus={(e) => (e.target.style.background = C.surfaceHover)}
      onBlur={(e) => (e.target.style.background = "transparent")}
    />
  );
}

// Vertically stacked block lines: each has a Qty cell and a Price cell,
// plus a remove button, with an "add block" row at the bottom.
function BlockStack({ blocks, onChangeBlock, onAddBlock, onRemoveBlock, width }) {
  const safeBlocks = Array.isArray(blocks) ? blocks : [];
  return (
    <div
      style={{
        width,
        flexShrink: 0,
        display: "flex",
        flexDirection: "column",
        borderLeft: `1px solid ${C.border}`,
        borderRight: `1px solid ${C.border}`,
      }}
    >
      {safeBlocks.map((b, i) => (
        <div
          key={b.id}
          style={{
            display: "flex",
            flexDirection: "column",
            padding: "4px 4px",
            borderBottom: `1px solid ${C.border}`,
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
            <span
              style={{
                fontFamily: FONT_MONO,
                fontSize: 10,
                color: C.textFaint,
                width: 10,
                flexShrink: 0,
              }}
            >
              {i + 1}
            </span>
            <input
              value={b.qty}
              onChange={(e) => onChangeBlock(i, "qty", e.target.value)}
              placeholder="qty"
              type="number"
              title="Number of blocks"
              style={{
                width: 34,
                flexShrink: 0,
                background: C.surfaceAlt,
                border: `1px solid ${C.border}`,
                borderRadius: 3,
                outline: "none",
                color: C.textDim,
                fontFamily: FONT_MONO,
                fontSize: 11,
                textAlign: "center",
                padding: "3px 2px",
              }}
            />
            <span style={{ color: C.textFaint, fontSize: 10 }}>×</span>
            <input
              value={b.price}
              onChange={(e) => onChangeBlock(i, "price", e.target.value)}
              placeholder="price"
              type="number"
              title="Price of block"
              style={{
                width: "100%",
                background: "transparent",
                border: "none",
                outline: "none",
                color: C.text,
                fontFamily: FONT_MONO,
                fontSize: 12,
                textAlign: "right",
                padding: "4px 2px",
              }}
              onFocus={(e) => (e.target.style.background = C.surfaceHover)}
              onBlur={(e) => (e.target.style.background = "transparent")}
            />
            <button
              onClick={() => onRemoveBlock(i)}
              title="Remove block"
              style={{ background: "transparent", border: "none", color: C.textFaint, padding: 1, flexShrink: 0 }}
            >
              <X size={10} />
            </button>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 3, paddingLeft: 14 }}>
            <select
              value={b.day || ""}
              onChange={(e) => onChangeBlock(i, "day", e.target.value)}
              title="Block start day"
              style={{
                width: "56%",
                background: "transparent",
                border: "none",
                outline: "none",
                color: b.day ? C.textFaint : C.textFaint,
                fontFamily: FONT_MONO,
                fontSize: 10.5,
                padding: "1px 2px",
              }}
            >
              <option style={{ background: C.surface, color: C.text }} value="">day</option>
              {["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"].map((d) => (
                <option key={d} style={{ background: C.surface, color: C.text }} value={d}>
                  {d}
                </option>
              ))}
            </select>
            <input
              value={b.time || ""}
              onChange={(e) => onChangeBlock(i, "time", e.target.value)}
              type="time"
              title="Block start time"
              style={{
                width: "44%",
                background: "transparent",
                border: "none",
                outline: "none",
                color: C.textFaint,
                fontFamily: FONT_MONO,
                fontSize: 10.5,
                padding: "1px 2px",
              }}
              onFocus={(e) => (e.target.style.background = C.surfaceHover)}
              onBlur={(e) => (e.target.style.background = "transparent")}
            />
          </div>
        </div>
      ))}
      <button
        onClick={onAddBlock}
        style={{
          background: "transparent",
          border: "none",
          color: C.textFaint,
          fontFamily: FONT_MONO,
          fontSize: 10.5,
          padding: "5px 4px",
          textAlign: "left",
        }}
      >
        + block
      </button>
    </div>
  );
}

function StatusPill({ value, onChange }) {
  const colors = {
    Paid: { fg: C.green, bg: C.greenDim },
    Pending: { fg: C.amber, bg: C.amberDim },
    Overdue: { fg: C.red, bg: C.redDim },
  };
  const c = colors[value] || colors.Pending;
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      style={{
        background: c.bg,
        color: c.fg,
        border: `1px solid ${c.fg}55`,
        borderRadius: 4,
        fontSize: 11,
        fontFamily: FONT_MONO,
        fontWeight: 600,
        padding: "4px 6px",
        outline: "none",
        letterSpacing: 0.3,
      }}
    >
      <option style={{ background: C.surface, color: C.text }} value="Paid">PAID</option>
      <option style={{ background: C.surface, color: C.text }} value="Pending">PENDING</option>
      <option style={{ background: C.surface, color: C.text }} value="Overdue">OVERDUE</option>
    </select>
  );
}

// ---------- Main app ----------
export default function FleetLedger() {
  const [trucks, setTrucks] = useState([]);
  const [contracts, setContracts] = useState([]);
  const [selectedTruckId, setSelectedTruckId] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [status, setStatus] = useState("");
  const [newTruckOpen, setNewTruckOpen] = useState(false);
  const [newTruckName, setNewTruckName] = useState("");
  const fileInputRef = useRef(null);
  const saveTimer = useRef(null);
  const backupTimer = useRef(null);
  const [backupHandle, setBackupHandle] = useState(null);
  const [backupName, setBackupName] = useState("");
  const [backupStatus, setBackupStatus] = useState("");
  const fileSystemAccessSupported = typeof window !== "undefined" && "showSaveFilePicker" in window;

  const [githubConfig, setGithubConfig] = useState(null);
  const [githubSha, setGithubSha] = useState(null);
  const [githubStatus, setGithubStatus] = useState("");
  const [showGithubPanel, setShowGithubPanel] = useState(false);
  const [ghOwner, setGhOwner] = useState("");
  const [ghRepo, setGhRepo] = useState("");
  const [ghBranch, setGhBranch] = useState("main");
  const [ghPath, setGhPath] = useState("data.json");
  const [ghToken, setGhToken] = useState("");
  const githubSyncTimer = useRef(null);
  const suppressNextGithubPush = useRef(false);

  const inputStyle = {
    background: C.surfaceAlt,
    border: `1px solid ${C.borderLight}`,
    borderRadius: 5,
    color: C.text,
    fontSize: 12.5,
    padding: "7px 8px",
    outline: "none",
    fontFamily: FONT_MONO,
  };

  const [isMobile, setIsMobile] = useState(
    typeof window !== "undefined" ? window.matchMedia("(max-width: 760px)").matches : false
  );
  useEffect(() => {
    const mq = window.matchMedia("(max-width: 760px)");
    const handler = (e) => setIsMobile(e.matches);
    mq.addEventListener ? mq.addEventListener("change", handler) : mq.addListener(handler);
    return () => {
      mq.removeEventListener ? mq.removeEventListener("change", handler) : mq.removeListener(handler);
    };
  }, []);

  // Load on mount - pulls from the linked GitHub data.json when a sync
  // config is saved on this device, otherwise falls back to local storage.
  useEffect(() => {
    (async () => {
      const applyData = (data) => {
        const t = data.trucks || [];
        const c = (data.contracts || []).map(migrateContract);
        setTrucks(t);
        setContracts(c);
        if (t.length > 0) setSelectedTruckId(t[0].id);
      };

      const cfg = loadGithubConfig();
      if (cfg) {
        setGithubConfig(cfg);
        setGhOwner(cfg.owner);
        setGhRepo(cfg.repo);
        setGhBranch(cfg.branch);
        setGhPath(cfg.path);
        setGhToken(cfg.token);
      }

      try {
        if (cfg) {
          setGithubStatus("Syncing...");
          const result = await githubGetFile(cfg);
          if (result.exists) {
            applyData(result.data);
            setGithubSha(result.sha);
            setGithubStatus("Synced " + new Date().toLocaleTimeString());
          } else {
            // data.json doesn't exist in the repo yet - fall back to
            // whatever's cached locally so it can be pushed up on first save.
            setGithubStatus("No data.json in repo yet");
            const localResult = await storage.get(STORAGE_KEY);
            if (localResult?.value) applyData(JSON.parse(localResult.value));
          }
        } else {
          const localResult = await storage.get(STORAGE_KEY);
          if (localResult?.value) applyData(JSON.parse(localResult.value));
        }
      } catch (e) {
        setGithubStatus("Sync error - " + e.message);
        try {
          const localResult = await storage.get(STORAGE_KEY);
          if (localResult?.value) applyData(JSON.parse(localResult.value));
        } catch (e2) {
          // no local fallback available either
        }
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  // Debounced save whenever data changes
  useEffect(() => {
    if (!loaded) return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try {
        await storage.set(STORAGE_KEY, JSON.stringify({ trucks, contracts }));
        setStatus("Saved");
        setTimeout(() => setStatus(""), 1200);
      } catch (e) {
        setStatus("Save failed");
      }
    }, 500);
    return () => clearTimeout(saveTimer.current);
  }, [trucks, contracts, loaded]);

  // Debounced push to the linked GitHub data.json whenever data changes.
  useEffect(() => {
    if (!loaded || !githubConfig) return;
    if (suppressNextGithubPush.current) {
      // This change came from a pull (GitHub -> app), not an edit -
      // don't immediately push it right back.
      suppressNextGithubPush.current = false;
      return;
    }
    if (githubSyncTimer.current) clearTimeout(githubSyncTimer.current);
    githubSyncTimer.current = setTimeout(async () => {
      setGithubStatus("Syncing...");
      try {
        let sha = githubSha;
        try {
          const fresh = await githubGetFile(githubConfig);
          sha = fresh.exists ? fresh.sha : null;
        } catch (e) {
          // fall back to the last-known sha if the pre-check fails
        }
        const result = await githubPutFile(githubConfig, { trucks, contracts }, sha);
        setGithubSha(result.content.sha);
        setGithubStatus("Synced " + new Date().toLocaleTimeString());
      } catch (e) {
        setGithubStatus("Sync error - " + e.message);
      }
    }, 1200);
    return () => clearTimeout(githubSyncTimer.current);
  }, [trucks, contracts, loaded, githubConfig]);

  // Re-pull from GitHub whenever the tab/window regains focus, so opening
  // this page on another device picks up changes made elsewhere.
  useEffect(() => {
    if (!githubConfig) return;
    const pullLatest = async () => {
      if (document.visibilityState && document.visibilityState !== "visible") return;
      setGithubStatus("Syncing...");
      try {
        const result = await githubGetFile(githubConfig);
        if (result.exists) {
          suppressNextGithubPush.current = true;
          setTrucks(result.data.trucks || []);
          setContracts((result.data.contracts || []).map(migrateContract));
          setGithubSha(result.sha);
          setGithubStatus("Synced " + new Date().toLocaleTimeString());
        }
      } catch (e) {
        setGithubStatus("Sync error - " + e.message);
      }
    };
    document.addEventListener("visibilitychange", pullLatest);
    window.addEventListener("focus", pullLatest);
    return () => {
      document.removeEventListener("visibilitychange", pullLatest);
      window.removeEventListener("focus", pullLatest);
    };
  }, [githubConfig]);

  const connectGithub = async () => {
    if (!ghOwner.trim() || !ghRepo.trim() || !ghToken.trim()) {
      setGithubStatus("Fill in owner, repo, and token");
      return;
    }
    const cfg = {
      owner: ghOwner.trim(),
      repo: ghRepo.trim(),
      branch: ghBranch.trim() || "main",
      path: ghPath.trim() || "data.json",
      token: ghToken.trim(),
    };
    setGithubStatus("Connecting...");
    try {
      const result = await githubGetFile(cfg);
      if (result.exists) {
        suppressNextGithubPush.current = true;
        setTrucks(result.data.trucks || []);
        setContracts((result.data.contracts || []).map(migrateContract));
        setGithubSha(result.sha);
      } else {
        const put = await githubPutFile(cfg, { trucks, contracts }, null);
        setGithubSha(put.content.sha);
      }
      saveGithubConfigLocal(cfg);
      setGithubConfig(cfg);
      setGithubStatus("Connected");
      setShowGithubPanel(false);
    } catch (e) {
      setGithubStatus("Connection failed - " + e.message);
    }
  };

  const disconnectGithub = () => {
    clearGithubConfigLocal();
    setGithubConfig(null);
    setGithubSha(null);
    setGithubStatus("");
  };

  const manualSyncGithub = async () => {
    if (!githubConfig) return;
    setGithubStatus("Syncing...");
    try {
      const result = await githubGetFile(githubConfig);
      if (result.exists) {
        suppressNextGithubPush.current = true;
        setTrucks(result.data.trucks || []);
        setContracts((result.data.contracts || []).map(migrateContract));
        setGithubSha(result.sha);
      }
      setGithubStatus("Synced " + new Date().toLocaleTimeString());
    } catch (e) {
      setGithubStatus("Sync error - " + e.message);
    }
  };

  // Debounced write to the linked live-backup .xlsx file, when one is set.
  useEffect(() => {
    if (!loaded || !backupHandle) return;
    if (backupTimer.current) clearTimeout(backupTimer.current);
    backupTimer.current = setTimeout(async () => {
      try {
        const wb = buildWorkbook();
        const arrayBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
        const writable = await backupHandle.createWritable();
        await writable.write(arrayBuffer);
        await writable.close();
        setBackupStatus("Backed up " + new Date().toLocaleTimeString());
      } catch (e) {
        setBackupStatus("Backup failed - relink the file");
        setBackupHandle(null);
      }
    }, 700);
    return () => clearTimeout(backupTimer.current);
  }, [trucks, contracts, loaded, backupHandle]);

  const linkBackupFile = async () => {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: "fleet-ledger-backup.xlsx",
        types: [
          {
            description: "Excel Workbook",
            accept: {
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
            },
          },
        ],
      });
      setBackupHandle(handle);
      setBackupName(handle.name);
      setBackupStatus("Linked - writing on every change");
    } catch (e) {
      // user cancelled the picker - do nothing
    }
  };

  const unlinkBackupFile = () => {
    setBackupHandle(null);
    setBackupName("");
    setBackupStatus("");
  };

  const addTruck = () => {
    if (!newTruckName.trim()) return;
    const t = { id: uid(), name: newTruckName.trim() };
    setTrucks((prev) => [...prev, t]);
    setSelectedTruckId(t.id);
    setNewTruckName("");
    setNewTruckOpen(false);
  };

  const deleteTruck = (id) => {
    setTrucks((prev) => prev.filter((t) => t.id !== id));
    setContracts((prev) => prev.filter((c) => c.truckId !== id));
    setSelectedTruckId((prev) => (prev === id ? null : prev));
  };

  const addContract = (truckId) => {
    setContracts((prev) => [...prev, emptyContract(truckId)]);
  };

  const updateContract = (id, field, value) => {
    setContracts((prev) => prev.map((c) => (c.id === id ? { ...c, [field]: value } : c)));
  };

  const updateBlock = (contractId, blockIndex, field, value) => {
    setContracts((prev) =>
      prev.map((c) => {
        if (c.id !== contractId) return c;
        const blocks = c.blocks.map((b, i) => (i === blockIndex ? { ...b, [field]: value } : b));
        return { ...c, blocks };
      })
    );
  };

  const addBlock = (contractId) => {
    setContracts((prev) =>
      prev.map((c) => (c.id === contractId ? { ...c, blocks: [...c.blocks, emptyBlock()] } : c))
    );
  };

  const removeBlock = (contractId, blockIndex) => {
    setContracts((prev) =>
      prev.map((c) => {
        if (c.id !== contractId) return c;
        const blocks = c.blocks.filter((_, i) => i !== blockIndex);
        return { ...c, blocks: blocks.length > 0 ? blocks : [emptyBlock()] };
      })
    );
  };

  const deleteContract = (id) => {
    setContracts((prev) => prev.filter((c) => c.id !== id));
  };

  const truckContracts = contracts.filter((c) => c.truckId === selectedTruckId);
  const selectedTruck = trucks.find((t) => t.id === selectedTruckId);

  const truckTotals = (truckId) => {
    const list = contracts.filter((c) => c.truckId === truckId);
    return list.reduce((sum, c) => sum + payoutPerWeek(c), 0);
  };

  const buildWorkbook = () => {
    const contractRows = [];
    contracts.forEach((c) => {
      const truck = trucks.find((t) => t.id === c.truckId);
      const payout = payoutPerWeek(c);
      (c.blocks || []).forEach((b, i) => {
        contractRows.push({
          "Truck": truck ? truck.name : "(unknown truck)",
          "Contact ID": c.contactId,
          "Period Start": c.periodStart,
          "Period End": c.periodEnd,
          "Week": calendarWeek(c.periodStart),
          "Block #": i + 1,
          "Number of Blocks": num(b.qty),
          "Price of Block": num(b.price),
          "Block Day": b.day || "",
          "Block Time": b.time || "",
          "Payout per Week": payout,
          "Base Price per Week": num(c.basePricePerWeek),
          "Payment Date": c.paymentDate,
          "Payment Status": c.paymentStatus,
        });
      });
    });

    const truckSheet = trucks.map((t) => ({
      "Truck": t.name,
      "Contracts": contracts.filter((c) => c.truckId === t.id).length,
      "Total Payout per Week": truckTotals(t.id),
    }));

    const wb = XLSX.utils.book_new();
    // "Contracts" is added first so it's the sheet that opens by default -
    // it mirrors the on-screen table column-for-column.
    const wsContracts = XLSX.utils.json_to_sheet(contractRows);
    const wsTrucks = XLSX.utils.json_to_sheet(truckSheet);
    XLSX.utils.book_append_sheet(wb, wsContracts, "Contracts");
    XLSX.utils.book_append_sheet(wb, wsTrucks, "Truck Summary");
    return wb;
  };

  const exportToExcel = () => {
    const wb = buildWorkbook();
    XLSX.writeFile(wb, "fleet-ledger.xlsx");
  };

  const importFromExcel = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });

        const contractsSheetName =
          wb.SheetNames.find((n) => n.toLowerCase() === "contracts") || wb.SheetNames[0];
        const rawContracts = contractsSheetName ? XLSX.utils.sheet_to_json(wb.Sheets[contractsSheetName]) : [];

        // Trucks are derived from the truck names found in the Contracts
        // sheet itself, in the order they first appear.
        const truckByName = new Map();
        const newTrucks = [];
        const truckIdFor = (name) => {
          const key = name || "(unknown truck)";
          if (!truckByName.has(key)) {
            const t = { id: uid(), name: key };
            truckByName.set(key, t.id);
            newTrucks.push(t);
          }
          return truckByName.get(key);
        };

        // Group imported block rows back into contracts, keyed by
        // truck + contact + period start/end.
        const grouped = new Map();
        rawContracts.forEach((r) => {
          const truckName = r["Truck"] || r["Truck Name"] || "(unknown truck)";
          const truckId = truckIdFor(truckName);
          const key = [truckName, r["Contact ID"], r["Period Start"], r["Period End"]].join("|");
          if (!grouped.has(key)) {
            grouped.set(key, {
              id: uid(),
              truckId,
              contactId: String(r["Contact ID"] ?? ""),
              periodStart: r["Period Start"] ? String(r["Period Start"]) : "",
              periodEnd: r["Period End"] ? String(r["Period End"]) : "",
              blocks: [],
              basePricePerWeek: r["Base Price per Week"] ?? "",
              otherPrice: "",
              surcharge: "",
              paymentDate: r["Payment Date"] ? String(r["Payment Date"]) : "",
              paymentStatus: r["Payment Status"] || "Pending",
            });
          }
          const entry = grouped.get(key);
          entry.blocks.push({
            id: uid(),
            qty: r["Number of Blocks"] ?? "1",
            price: r["Price of Block"] ?? "",
            day: r["Block Day"] ?? "",
            time: r["Block Time"] ?? "",
          });
        });

        const newContracts = Array.from(grouped.values()).map((c) => ({
          ...c,
          blocks: c.blocks.length > 0 ? c.blocks : [emptyBlock()],
        }));

        setTrucks(newTrucks);
        setContracts(newContracts);
        if (newTrucks.length > 0) setSelectedTruckId(newTrucks[0].id);
        setStatus("Imported");
        setTimeout(() => setStatus(""), 1500);
      } catch (err) {
        setStatus("Import failed");
        setTimeout(() => setStatus(""), 2000);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const colWidths = {
    contact: 140,
    periodStart: 138,
    periodEnd: 138,
    week: 42,
    block: 210,
    payout: 92,
    base: 92,
    payDate: 100,
    status: 96,
    del: 28,
  };

  return (
    <div
      style={{
        fontFamily: FONT_BODY,
        background: C.bg,
        color: C.text,
        minHeight: "100vh",
        display: "flex",
        flexDirection: "column",
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;600&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { height: 8px; width: 8px; }
        ::-webkit-scrollbar-thumb { background: ${C.borderLight}; border-radius: 4px; }
        ::-webkit-scrollbar-track { background: transparent; }
        input::placeholder { color: ${C.textFaint}; }
        button { cursor: pointer; font-family: ${FONT_BODY}; }
        input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button { opacity: 0.3; }
      `}</style>

      {/* Header */}
      <div
        style={{
          borderBottom: `1px solid ${C.border}`,
          background: C.surface,
          padding: isMobile ? "12px 14px" : "16px 24px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          flexWrap: "wrap",
          gap: 10,
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div
            style={{
              width: 34,
              height: 34,
              borderRadius: 6,
              background: C.amber,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              flexShrink: 0,
            }}
          >
            <Truck size={19} color={C.bg} strokeWidth={2.4} />
          </div>
          <div>
            <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 700, fontSize: 18, letterSpacing: 0.2 }}>
              Fleet Ledger
            </div>
            {!isMobile && (
              <div style={{ fontSize: 11.5, color: C.textDim, fontFamily: FONT_MONO }}>
                trucks &amp; contracts · optional live .xlsx backup
              </div>
            )}
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: isMobile ? 6 : 10, flexWrap: "wrap" }}>
          {status && !isMobile && (
            <span style={{ fontSize: 11.5, color: C.textDim, fontFamily: FONT_MONO }}>{status}</span>
          )}

          {githubConfig ? (
            <div
              title={githubStatus}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 7,
                background: C.surfaceAlt,
                border: `1px solid ${C.borderLight}`,
                borderRadius: 6,
                padding: "7px 10px",
              }}
            >
              <span
                style={{
                  width: 6,
                  height: 6,
                  borderRadius: 99,
                  background: githubStatus.startsWith("Sync error") ? C.red : C.green,
                  flexShrink: 0,
                }}
              />
              {!isMobile && (
                <span style={{ fontFamily: FONT_MONO, fontSize: 11, color: C.textDim, maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {githubConfig.owner}/{githubConfig.repo}
                </span>
              )}
              <button
                onClick={manualSyncGithub}
                title="Sync now"
                style={{ background: "transparent", border: "none", color: C.textFaint, padding: 0, display: "flex" }}
              >
                <RefreshCw size={12} />
              </button>
              <button
                onClick={() => setShowGithubPanel(true)}
                title="GitHub sync settings"
                style={{ background: "transparent", border: "none", color: C.textFaint, padding: 0, display: "flex" }}
              >
                <Settings size={12} />
              </button>
            </div>
          ) : (
            <button
              onClick={() => setShowGithubPanel(true)}
              title="Sync this app's data through a file in a GitHub repo"
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                background: "transparent",
                border: `1px dashed ${C.borderLight}`,
                color: C.textDim,
                borderRadius: 6,
                padding: "8px 12px",
                fontSize: 12.5,
              }}
            >
              <Github size={14} /> {!isMobile && "Sync via GitHub"}
            </button>
          )}

          {!isMobile && (fileSystemAccessSupported ? (
            backupHandle ? (
              <div
                title={backupStatus}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 7,
                  background: C.surfaceAlt,
                  border: `1px solid ${C.borderLight}`,
                  borderRadius: 6,
                  padding: "7px 10px",
                }}
              >
                <span style={{ width: 6, height: 6, borderRadius: 99, background: C.green, flexShrink: 0 }} />
                <span style={{ fontFamily: FONT_MONO, fontSize: 11, color: C.textDim, maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {backupName}
                </span>
                <button
                  onClick={unlinkBackupFile}
                  title="Unlink backup file"
                  style={{ background: "transparent", border: "none", color: C.textFaint, padding: 0 }}
                >
                  <X size={12} />
                </button>
              </div>
            ) : (
              <button
                onClick={linkBackupFile}
                title="Pick or create an .xlsx file - every change will be written to it automatically"
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  background: "transparent",
                  border: `1px dashed ${C.borderLight}`,
                  color: C.textDim,
                  borderRadius: 6,
                  padding: "8px 12px",
                  fontSize: 12.5,
                }}
              >
                <Link2 size={14} /> Link backup .xlsx
              </button>
            )
          ) : (
            <span
              title="Live auto-backup to a file needs Chrome or Edge on a computer. On phones and Safari, use Export .xlsx below."
              style={{ fontSize: 11, color: C.textFaint, fontFamily: FONT_MONO, maxWidth: 150, lineHeight: 1.3 }}
            >
              Live backup: Chrome/Edge only
            </span>
          ))}

          <input
            type="file"
            accept=".xlsx,.xls"
            ref={fileInputRef}
            onChange={importFromExcel}
            style={{ display: "none" }}
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            title="Import .xlsx"
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              background: "transparent",
              border: `1px solid ${C.borderLight}`,
              color: C.text,
              borderRadius: 6,
              padding: isMobile ? "8px 10px" : "8px 12px",
              fontSize: 12.5,
              fontWeight: 500,
            }}
          >
            <Upload size={14} /> {!isMobile && "Import .xlsx"}
          </button>
          <button
            onClick={exportToExcel}
            title="Export .xlsx"
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              background: C.amber,
              border: "none",
              color: C.bg,
              borderRadius: 6,
              padding: isMobile ? "8px 10px" : "8px 12px",
              fontSize: 12.5,
              fontWeight: 600,
            }}
          >
            <Download size={14} /> {!isMobile && "Export .xlsx"}
          </button>
        </div>
      </div>

      {showGithubPanel && (
        <div
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(0,0,0,0.55)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 50,
          }}
          onClick={() => setShowGithubPanel(false)}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              background: C.surface,
              border: `1px solid ${C.borderLight}`,
              borderRadius: 10,
              padding: 22,
              width: 380,
              maxWidth: "90vw",
            }}
          >
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 4 }}>
              <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 700, fontSize: 16, display: "flex", alignItems: "center", gap: 8 }}>
                <Github size={17} /> Sync via GitHub
              </div>
              <button
                onClick={() => setShowGithubPanel(false)}
                style={{ background: "transparent", border: "none", color: C.textFaint }}
              >
                <X size={16} />
              </button>
            </div>
            <div style={{ fontSize: 11.5, color: C.textDim, marginBottom: 16, lineHeight: 1.4 }}>
              Stores your data in a <code>data.json</code> file in your repo, so any
              device connected here shows the same data.
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              <label style={{ fontSize: 10.5, color: C.textFaint, fontFamily: FONT_MONO }}>GitHub username</label>
              <input
                value={ghOwner}
                onChange={(e) => setGhOwner(e.target.value)}
                placeholder="jaouadbico"
                style={inputStyle}
              />
              <label style={{ fontSize: 10.5, color: C.textFaint, fontFamily: FONT_MONO }}>Repository name</label>
              <input
                value={ghRepo}
                onChange={(e) => setGhRepo(e.target.value)}
                placeholder="fleet-ledger"
                style={inputStyle}
              />
              <div style={{ display: "flex", gap: 8 }}>
                <div style={{ flex: 1 }}>
                  <label style={{ fontSize: 10.5, color: C.textFaint, fontFamily: FONT_MONO }}>Branch</label>
                  <input value={ghBranch} onChange={(e) => setGhBranch(e.target.value)} style={{ ...inputStyle, width: "100%" }} />
                </div>
                <div style={{ flex: 1 }}>
                  <label style={{ fontSize: 10.5, color: C.textFaint, fontFamily: FONT_MONO }}>File path</label>
                  <input value={ghPath} onChange={(e) => setGhPath(e.target.value)} style={{ ...inputStyle, width: "100%" }} />
                </div>
              </div>
              <label style={{ fontSize: 10.5, color: C.textFaint, fontFamily: FONT_MONO }}>Personal access token</label>
              <input
                value={ghToken}
                onChange={(e) => setGhToken(e.target.value)}
                placeholder="github_pat_..."
                type="password"
                style={inputStyle}
              />
            </div>

            {githubStatus && (
              <div style={{ fontSize: 11, color: C.textDim, marginTop: 10, fontFamily: FONT_MONO }}>{githubStatus}</div>
            )}

            <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
              <button
                onClick={connectGithub}
                style={{
                  flex: 1,
                  background: C.amber,
                  color: C.bg,
                  border: "none",
                  borderRadius: 6,
                  padding: "9px 0",
                  fontSize: 13,
                  fontWeight: 600,
                }}
              >
                {githubConfig ? "Reconnect" : "Connect"}
              </button>
              {githubConfig && (
                <button
                  onClick={() => {
                    disconnectGithub();
                    setShowGithubPanel(false);
                  }}
                  style={{
                    background: "transparent",
                    color: C.red,
                    border: `1px solid ${C.redDim}`,
                    borderRadius: 6,
                    padding: "9px 14px",
                    fontSize: 13,
                  }}
                >
                  Disconnect
                </button>
              )}
            </div>
          </div>
        </div>
      )}

      <div style={{ display: "flex", flex: 1, minHeight: 0, flexDirection: isMobile ? "column" : "row" }}>
        {/* Sidebar / truck roster */}
        <div
          style={{
            width: isMobile ? "100%" : 260,
            flexShrink: 0,
            borderRight: isMobile ? "none" : `1px solid ${C.border}`,
            borderBottom: isMobile ? `1px solid ${C.border}` : "none",
            background: C.surface,
            display: "flex",
            flexDirection: "column",
          }}
        >
          {!isMobile && (
            <div
              style={{
                padding: "14px 16px 8px",
                fontSize: 11,
                fontFamily: FONT_MONO,
                color: C.textFaint,
                letterSpacing: 0.6,
                textTransform: "uppercase",
              }}
            >
              Fleet Roster · {trucks.length}
            </div>
          )}

          <div
            style={
              isMobile
                ? { display: "flex", flexDirection: "row", overflowX: "auto", gap: 8, padding: "10px 12px" }
                : { flex: 1, overflowY: "auto", padding: "0 10px" }
            }
          >
            {trucks.map((t) => {
              const active = t.id === selectedTruckId;
              const tContracts = contracts.filter((c) => c.truckId === t.id);
              return (
                <div
                  key={t.id}
                  onClick={() => setSelectedTruckId(t.id)}
                  style={
                    isMobile
                      ? {
                          background: active ? C.surfaceAlt : "transparent",
                          border: `1px solid ${active ? C.amberDim : C.border}`,
                          borderRadius: 8,
                          padding: "8px 12px",
                          flexShrink: 0,
                          minWidth: 120,
                          cursor: "pointer",
                        }
                      : {
                          background: active ? C.surfaceAlt : "transparent",
                          border: `1px solid ${active ? C.amberDim : "transparent"}`,
                          borderRadius: 8,
                          padding: "10px 10px",
                          marginBottom: 6,
                          cursor: "pointer",
                        }
                  }
                >
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 6 }}>
                    <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 600, fontSize: 13.5, color: active ? C.amber : C.text, whiteSpace: isMobile ? "nowrap" : "normal" }}>
                      {t.name}
                    </div>
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        deleteTruck(t.id);
                      }}
                      style={{
                        background: "transparent",
                        border: "none",
                        color: C.textFaint,
                        padding: 2,
                        flexShrink: 0,
                      }}
                      title="Remove truck"
                    >
                      <X size={13} />
                    </button>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 10, marginTop: isMobile ? 4 : 8, fontSize: 11, fontFamily: FONT_MONO, color: C.textDim, whiteSpace: "nowrap" }}>
                    <span>{tContracts.length} contract{tContracts.length === 1 ? "" : "s"}</span>
                    <span style={{ color: C.green }}>{money(truckTotals(t.id))}</span>
                  </div>
                </div>
              );
            })}

            {trucks.length === 0 && !isMobile && (
              <div style={{ fontSize: 12, color: C.textFaint, padding: "12px 6px" }}>
                No trucks yet. Add your first truck below.
              </div>
            )}
          </div>

          <div style={{ padding: 12, borderTop: isMobile ? "none" : `1px solid ${C.border}` }}>
            {!newTruckOpen ? (
              <button
                onClick={() => setNewTruckOpen(true)}
                style={{
                  width: "100%",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  gap: 6,
                  background: "transparent",
                  border: `1px dashed ${C.borderLight}`,
                  color: C.textDim,
                  borderRadius: 6,
                  padding: "9px 10px",
                  fontSize: 12.5,
                }}
              >
                <Plus size={14} /> Add truck
              </button>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                <input
                  autoFocus
                  value={newTruckName}
                  onChange={(e) => setNewTruckName(e.target.value)}
                  placeholder="Truck name (e.g. Unit 12 — Cascadia)"
                  onKeyDown={(e) => e.key === "Enter" && addTruck()}
                  style={{
                    background: C.surfaceAlt,
                    border: `1px solid ${C.borderLight}`,
                    borderRadius: 5,
                    color: C.text,
                    fontSize: 12.5,
                    padding: "7px 8px",
                    outline: "none",
                    fontFamily: FONT_BODY,
                  }}
                />
                <div style={{ display: "flex", gap: 6 }}>
                  <button
                    onClick={addTruck}
                    style={{
                      flex: 1,
                      background: C.amber,
                      color: C.bg,
                      border: "none",
                      borderRadius: 5,
                      padding: "7px 0",
                      fontSize: 12,
                      fontWeight: 600,
                    }}
                  >
                    Add
                  </button>
                  <button
                    onClick={() => {
                      setNewTruckOpen(false);
                      setNewTruckName("");
                    }}
                    style={{
                      background: "transparent",
                      color: C.textDim,
                      border: `1px solid ${C.borderLight}`,
                      borderRadius: 5,
                      padding: "7px 10px",
                      fontSize: 12,
                    }}
                  >
                    Cancel
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Main content */}
        <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column" }}>
          {!selectedTruck ? (
            <div
              style={{
                flex: 1,
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
                color: C.textFaint,
                gap: 8,
              }}
            >
              <Truck size={32} color={C.textFaint} />
              <div style={{ fontSize: 13.5 }}>Select or add a truck to track its contracts.</div>
            </div>
          ) : (
            <>
              {/* Truck header */}
              <div
                style={{
                  padding: isMobile ? "14px 14px 12px" : "18px 24px 14px",
                  borderBottom: `1px solid ${C.border}`,
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: isMobile ? "flex-start" : "flex-end",
                  flexWrap: "wrap",
                  gap: 10,
                }}
              >
                <div>
                  <div style={{ fontSize: 11, fontFamily: FONT_MONO, color: C.textFaint, letterSpacing: 0.6, textTransform: "uppercase" }}>
                    Truck
                  </div>
                  <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 700, fontSize: isMobile ? 19 : 22, display: "flex", alignItems: "center", gap: 10 }}>
                    {selectedTruck.name}
                  </div>
                </div>
                <div style={{ display: "flex", gap: isMobile ? 14 : 20 }}>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: 10.5, color: C.textFaint, fontFamily: FONT_MONO, textTransform: "uppercase" }}>Contracts</div>
                    <div style={{ fontFamily: FONT_DISPLAY, fontSize: 18, fontWeight: 600 }}>{truckContracts.length}</div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: 10.5, color: C.textFaint, fontFamily: FONT_MONO, textTransform: "uppercase" }}>Total revenue</div>
                    <div style={{ fontFamily: FONT_DISPLAY, fontSize: 18, fontWeight: 600, color: C.green }}>
                      {money(truckTotals(selectedTruck.id))}
                    </div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: 10.5, color: C.textFaint, fontFamily: FONT_MONO, textTransform: "uppercase" }}>Pending</div>
                    <div style={{ fontFamily: FONT_DISPLAY, fontSize: 18, fontWeight: 600, color: C.amber }}>
                      {truckContracts.filter((c) => c.paymentStatus !== "Paid").length}
                    </div>
                  </div>
                </div>
              </div>

              {/* Contracts */}
              {isMobile ? (
                <div style={{ flex: 1, overflow: "auto", padding: "14px 14px 24px" }}>
                  {truckContracts.map((c) => {
                    const payout = payoutPerWeek(c);
                    return (
                      <div
                        key={c.id}
                        style={{
                          background: C.surfaceAlt,
                          border: `1px solid ${C.border}`,
                          borderRadius: 10,
                          padding: 12,
                          marginBottom: 12,
                        }}
                      >
                        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                          <div style={{ flex: 1 }}>
                            <div style={{ fontSize: 9.5, color: C.textFaint, fontFamily: FONT_MONO, textTransform: "uppercase", marginBottom: 2 }}>
                              Contact ID
                            </div>
                            <input
                              value={c.contactId}
                              onChange={(e) => updateContract(c.id, "contactId", e.target.value)}
                              placeholder="C-1001"
                              style={{ ...inputStyle, width: "100%", fontSize: 14, fontWeight: 600 }}
                            />
                          </div>
                          <button
                            onClick={() => deleteContract(c.id)}
                            style={{ background: "transparent", border: "none", color: C.textFaint, padding: 6, flexShrink: 0, marginTop: 14 }}
                            title="Delete contract"
                          >
                            <Trash2 size={16} />
                          </button>
                        </div>

                        <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                          <div style={{ flex: 1 }}>
                            <div style={{ fontSize: 9.5, color: C.textFaint, fontFamily: FONT_MONO, textTransform: "uppercase", marginBottom: 2 }}>
                              Period Start
                            </div>
                            <input
                              type="date"
                              value={c.periodStart}
                              onChange={(e) => updateContract(c.id, "periodStart", e.target.value)}
                              style={{ ...inputStyle, width: "100%" }}
                            />
                          </div>
                          <div style={{ flex: 1 }}>
                            <div style={{ fontSize: 9.5, color: C.textFaint, fontFamily: FONT_MONO, textTransform: "uppercase", marginBottom: 2 }}>
                              Period End
                            </div>
                            <input
                              type="date"
                              value={c.periodEnd}
                              onChange={(e) => updateContract(c.id, "periodEnd", e.target.value)}
                              style={{ ...inputStyle, width: "100%" }}
                            />
                          </div>
                          <div style={{ width: 52, flexShrink: 0 }}>
                            <div style={{ fontSize: 9.5, color: C.textFaint, fontFamily: FONT_MONO, textTransform: "uppercase", marginBottom: 2 }}>
                              Wk
                            </div>
                            <div style={{ fontFamily: FONT_MONO, fontSize: 13, color: C.textDim, padding: "7px 0" }}>
                              {calendarWeek(c.periodStart) || "—"}
                            </div>
                          </div>
                        </div>

                        <div style={{ fontSize: 9.5, color: C.textFaint, fontFamily: FONT_MONO, textTransform: "uppercase", marginBottom: 4 }}>
                          Blocks (qty × price / day · time)
                        </div>
                        <div style={{ border: `1px solid ${C.border}`, borderRadius: 6, marginBottom: 10, overflow: "hidden" }}>
                          <BlockStack
                            width="100%"
                            blocks={c.blocks}
                            onChangeBlock={(i, field, v) => updateBlock(c.id, i, field, v)}
                            onAddBlock={() => addBlock(c.id)}
                            onRemoveBlock={(i) => removeBlock(c.id, i)}
                          />
                        </div>

                        <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                          <div style={{ flex: 1 }}>
                            <div style={{ fontSize: 9.5, color: C.textFaint, fontFamily: FONT_MONO, textTransform: "uppercase", marginBottom: 2 }}>
                              Payout/wk
                            </div>
                            <div style={{ fontFamily: FONT_MONO, fontSize: 14, fontWeight: 600, color: C.green, padding: "7px 0" }}>
                              {money(payout)}
                            </div>
                          </div>
                          <div style={{ flex: 1 }}>
                            <div style={{ fontSize: 9.5, color: C.textFaint, fontFamily: FONT_MONO, textTransform: "uppercase", marginBottom: 2 }}>
                              Base price/wk
                            </div>
                            <input
                              type="number"
                              value={c.basePricePerWeek}
                              onChange={(e) => updateContract(c.id, "basePricePerWeek", e.target.value)}
                              placeholder="0"
                              style={{ ...inputStyle, width: "100%" }}
                            />
                          </div>
                        </div>

                        <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
                          <div style={{ flex: 1 }}>
                            <div style={{ fontSize: 9.5, color: C.textFaint, fontFamily: FONT_MONO, textTransform: "uppercase", marginBottom: 2 }}>
                              Payment Date
                            </div>
                            <input
                              type="date"
                              value={c.paymentDate}
                              onChange={(e) => updateContract(c.id, "paymentDate", e.target.value)}
                              style={{ ...inputStyle, width: "100%" }}
                            />
                          </div>
                          <StatusPill value={c.paymentStatus} onChange={(v) => updateContract(c.id, "paymentStatus", v)} />
                        </div>
                      </div>
                    );
                  })}

                  <button
                    onClick={() => addContract(selectedTruck.id)}
                    style={{
                      width: "100%",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      gap: 6,
                      background: "transparent",
                      border: `1px dashed ${C.borderLight}`,
                      color: C.textDim,
                      borderRadius: 6,
                      padding: "12px 14px",
                      fontSize: 13,
                    }}
                  >
                    <Plus size={14} /> Add contract
                  </button>
                </div>
              ) : (
              <div style={{ flex: 1, overflow: "auto", padding: "0 24px 24px" }}>
                <div style={{ minWidth: 1250, marginTop: 14 }}>
                  {/* header row */}
                  <div
                    style={{
                      display: "flex",
                      fontSize: 10.5,
                      fontFamily: FONT_MONO,
                      color: C.textFaint,
                      textTransform: "uppercase",
                      letterSpacing: 0.4,
                      padding: "0 0 8px",
                      borderBottom: `1px solid ${C.border}`,
                    }}
                  >
                    <Cell width={colWidths.contact}>Contact ID</Cell>
                    <Cell width={colWidths.periodStart}>Period Start</Cell>
                    <Cell width={colWidths.periodEnd}>Period End</Cell>
                    <Cell width={colWidths.week}>Wk</Cell>
                    <Cell width={colWidths.block}>Qty × Price / Start Day · Time</Cell>
                    <Cell width={colWidths.payout} align="right">Payout/wk</Cell>
                    <Cell width={colWidths.base} align="right">Base price/wk</Cell>
                    <Cell width={colWidths.payDate}>Payment Date</Cell>
                    <Cell width={colWidths.status}>Status</Cell>
                    <Cell width={colWidths.del}></Cell>
                  </div>

                  {truckContracts.map((c) => {
                    const payout = payoutPerWeek(c);
                    return (
                      <div
                        key={c.id}
                        style={{
                          display: "flex",
                          alignItems: "stretch",
                          borderBottom: `1px solid ${C.border}`,
                        }}
                      >
                        <Cell width={colWidths.contact}>
                          <EditableInput value={c.contactId} onChange={(v) => updateContract(c.id, "contactId", v)} placeholder="C-1001" />
                        </Cell>
                        <Cell width={colWidths.periodStart}>
                          <EditableInput type="date" value={c.periodStart} onChange={(v) => updateContract(c.id, "periodStart", v)} />
                        </Cell>
                        <Cell width={colWidths.periodEnd}>
                          <EditableInput type="date" value={c.periodEnd} onChange={(v) => updateContract(c.id, "periodEnd", v)} />
                        </Cell>
                        <Cell width={colWidths.week}>
                          <span style={{ fontFamily: FONT_MONO, fontSize: 12, color: C.textDim }}>{calendarWeek(c.periodStart) || "—"}</span>
                        </Cell>
                        <BlockStack
                          width={colWidths.block}
                          blocks={c.blocks}
                          onChangeBlock={(i, field, v) => updateBlock(c.id, i, field, v)}
                          onAddBlock={() => addBlock(c.id)}
                          onRemoveBlock={(i) => removeBlock(c.id, i)}
                        />
                        <Cell width={colWidths.payout} align="right">
                          <span style={{ fontFamily: FONT_MONO, fontSize: 12.5, color: C.text }}>{money(payout)}</span>
                        </Cell>
                        <Cell width={colWidths.base} align="right">
                          <EditableInput align="right" type="number" value={c.basePricePerWeek} onChange={(v) => updateContract(c.id, "basePricePerWeek", v)} placeholder="0" />
                        </Cell>
                        <Cell width={colWidths.payDate}>
                          <EditableInput type="date" value={c.paymentDate} onChange={(v) => updateContract(c.id, "paymentDate", v)} />
                        </Cell>
                        <Cell width={colWidths.status}>
                          <StatusPill value={c.paymentStatus} onChange={(v) => updateContract(c.id, "paymentStatus", v)} />
                        </Cell>
                        <Cell width={colWidths.del}>
                          <button
                            onClick={() => deleteContract(c.id)}
                            style={{ background: "transparent", border: "none", color: C.textFaint, padding: 4 }}
                            title="Delete contract"
                          >
                            <Trash2 size={14} />
                          </button>
                        </Cell>
                      </div>
                    );
                  })}

                  <button
                    onClick={() => addContract(selectedTruck.id)}
                    style={{
                      marginTop: 12,
                      display: "flex",
                      alignItems: "center",
                      gap: 6,
                      background: "transparent",
                      border: `1px dashed ${C.borderLight}`,
                      color: C.textDim,
                      borderRadius: 6,
                      padding: "9px 14px",
                      fontSize: 12.5,
                    }}
                  >
                    <Plus size={14} /> Add contract
                  </button>
                </div>
              </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
