import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { Truck, Plus, Download, Upload, X, Trash2, Link2 } from "lucide-react";

// ---------- Design tokens ----------
const C = {
  bg: "#12161B",
  surface: "#1A2028",
  surfaceAlt: "#212934",
  surfaceHover: "#262F3B",
  border: "#2B333E",
  borderLight: "#374150",
  text: "#E7EAEE",
  textDim: "#8A94A3",
  textFaint: "#5C6675",
  amber: "#F2A93B",
  amberDim: "#8A6420",
  green: "#5FBE83",
  greenDim: "#2C4E3A",
  red: "#E2685C",
  redDim: "#4E2C2C",
  blue: "#5B93E0",
};

const FONT_DISPLAY = "'Space Grotesk', sans-serif";
const FONT_BODY = "'Inter', sans-serif";
const FONT_MONO = "'JetBrains Mono', monospace";

const STORAGE_KEY = "fleet-ledger-data";

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

function isoWeek(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr + "T00:00:00");
  if (isNaN(d.getTime())) return "";
  const target = new Date(d.valueOf());
  const dayNr = (d.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  const firstThursday = new Date(target.getFullYear(), 0, 4);
  const diff = target - firstThursday;
  return 1 + Math.round(diff / (7 * 24 * 3600 * 1000));
}

function num(v) {
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function money(n) {
  return n.toLocaleString(undefined, { style: "currency", currency: "USD" });
}

const emptyBlock = () => ({ id: uid(), qty: "1", price: "" });

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
      })),
      periodStart: c.periodStart ?? c.period ?? "",
      periodEnd: c.periodEnd ?? "",
      basePricePerWeek: c.basePricePerWeek ?? "",
    };
  }

  // Legacy shape: block1 / block2 / block3 as flat price fields.
  const legacyBlocks = [c.block1, c.block2, c.block3]
    .filter((v) => v !== undefined)
    .map((price) => ({ id: uid(), qty: "1", price: price ?? "" }));

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
            alignItems: "center",
            gap: 4,
            height: 34,
            padding: "0 4px",
            borderBottom: `1px solid ${C.border}`,
          }}
        >
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
  const [newTruckPlate, setNewTruckPlate] = useState("");
  const fileInputRef = useRef(null);
  const saveTimer = useRef(null);
  const backupTimer = useRef(null);
  const [backupHandle, setBackupHandle] = useState(null);
  const [backupName, setBackupName] = useState("");
  const [backupStatus, setBackupStatus] = useState("");
  const fileSystemAccessSupported = typeof window !== "undefined" && "showSaveFilePicker" in window;

  // Load from storage on mount
  useEffect(() => {
    (async () => {
      try {
        const result = await storage.get(STORAGE_KEY);
        if (result?.value) {
          const parsed = JSON.parse(result.value);
          setTrucks(parsed.trucks || []);
          setContracts((parsed.contracts || []).map(migrateContract));
          if ((parsed.trucks || []).length > 0) {
            setSelectedTruckId(parsed.trucks[0].id);
          }
        }
      } catch (e) {
        // no existing data yet
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
    const t = { id: uid(), name: newTruckName.trim(), plate: newTruckPlate.trim() };
    setTrucks((prev) => [...prev, t]);
    setSelectedTruckId(t.id);
    setNewTruckName("");
    setNewTruckPlate("");
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
    const truckSheet = trucks.map((t) => ({
      "Truck ID": t.id,
      "Truck Name": t.name,
      "Plate": t.plate,
      "Contracts": contracts.filter((c) => c.truckId === t.id).length,
      "Total Revenue": truckTotals(t.id),
    }));

    const contractRows = [];
    contracts.forEach((c) => {
      const truck = trucks.find((t) => t.id === c.truckId);
      const payout = payoutPerWeek(c);
      (c.blocks || []).forEach((b, i) => {
        contractRows.push({
          "Truck Name": truck ? truck.name : "",
          "Truck ID": c.truckId,
          "Contact ID": c.contactId,
          "Period Start": c.periodStart,
          "Period End": c.periodEnd,
          "Calendar Week": isoWeek(c.periodStart),
          "Block #": i + 1,
          "Number of Blocks": num(b.qty),
          "Price of Block": num(b.price),
          "Payout per Week": i === 0 ? payout : "",
          "Base Price per Week": i === 0 ? num(c.basePricePerWeek) : "",
          "Payment Date": i === 0 ? c.paymentDate : "",
          "Payment Status": i === 0 ? c.paymentStatus : "",
        });
      });
    });

    const wb = XLSX.utils.book_new();
    const wsTrucks = XLSX.utils.json_to_sheet(truckSheet);
    const wsContracts = XLSX.utils.json_to_sheet(contractRows);
    XLSX.utils.book_append_sheet(wb, wsTrucks, "Trucks");
    XLSX.utils.book_append_sheet(wb, wsContracts, "Contracts");
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

        const trucksSheetName = wb.SheetNames.find((n) => n.toLowerCase() === "trucks") || wb.SheetNames[0];
        const contractsSheetName = wb.SheetNames.find((n) => n.toLowerCase() === "contracts") || wb.SheetNames[1];

        const rawTrucks = trucksSheetName ? XLSX.utils.sheet_to_json(wb.Sheets[trucksSheetName]) : [];
        const rawContracts = contractsSheetName ? XLSX.utils.sheet_to_json(wb.Sheets[contractsSheetName]) : [];

        const idMap = {};
        const newTrucks = rawTrucks.map((r) => {
          const id = uid();
          idMap[r["Truck ID"]] = id;
          return {
            id,
            name: r["Truck Name"] || "Unnamed Truck",
            plate: r["Plate"] || "",
          };
        });

        // Group imported block rows back into contracts, keyed by
        // truck + contact + period start/end.
        const grouped = new Map();
        rawContracts.forEach((r) => {
          const key = [r["Truck ID"], r["Contact ID"], r["Period Start"], r["Period End"]].join("|");
          if (!grouped.has(key)) {
            grouped.set(key, {
              id: uid(),
              truckId: idMap[r["Truck ID"]] || newTrucks[0]?.id || null,
              contactId: String(r["Contact ID"] ?? ""),
              periodStart: r["Period Start"] ? String(r["Period Start"]) : "",
              periodEnd: r["Period End"] ? String(r["Period End"]) : "",
              blocks: [],
              basePricePerWeek: r["Base Price per Week"] || "",
              otherPrice: r["Other Contract Price"] || "",
              surcharge: r["Surcharges"] || "",
              paymentDate: r["Payment Date"] ? String(r["Payment Date"]) : "",
              paymentStatus: r["Payment Status"] || "Pending",
            });
          }
          const entry = grouped.get(key);
          entry.blocks.push({
            id: uid(),
            qty: r["Number of Blocks"] ?? "1",
            price: r["Price of Block"] ?? "",
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
    contact: 88,
    periodStart: 92,
    periodEnd: 92,
    week: 42,
    block: 150,
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
          padding: "16px 24px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
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
            }}
          >
            <Truck size={19} color={C.bg} strokeWidth={2.4} />
          </div>
          <div>
            <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 700, fontSize: 18, letterSpacing: 0.2 }}>
              Fleet Ledger
            </div>
            <div style={{ fontSize: 11.5, color: C.textDim, fontFamily: FONT_MONO }}>
              trucks &amp; contracts · optional live .xlsx backup
            </div>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {status && (
            <span style={{ fontSize: 11.5, color: C.textDim, fontFamily: FONT_MONO }}>{status}</span>
          )}

          {fileSystemAccessSupported ? (
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
          )}

          <input
            type="file"
            accept=".xlsx,.xls"
            ref={fileInputRef}
            onChange={importFromExcel}
            style={{ display: "none" }}
          />
          <button
            onClick={() => fileInputRef.current?.click()}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              background: "transparent",
              border: `1px solid ${C.borderLight}`,
              color: C.text,
              borderRadius: 6,
              padding: "8px 12px",
              fontSize: 12.5,
              fontWeight: 500,
            }}
          >
            <Upload size={14} /> Import .xlsx
          </button>
          <button
            onClick={exportToExcel}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              background: C.amber,
              border: "none",
              color: C.bg,
              borderRadius: 6,
              padding: "8px 12px",
              fontSize: 12.5,
              fontWeight: 600,
            }}
          >
            <Download size={14} /> Export .xlsx
          </button>
        </div>
      </div>

      <div style={{ display: "flex", flex: 1, minHeight: 0 }}>
        {/* Sidebar */}
        <div
          style={{
            width: 260,
            flexShrink: 0,
            borderRight: `1px solid ${C.border}`,
            background: C.surface,
            display: "flex",
            flexDirection: "column",
          }}
        >
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

          <div style={{ flex: 1, overflowY: "auto", padding: "0 10px" }}>
            {trucks.map((t) => {
              const active = t.id === selectedTruckId;
              const tContracts = contracts.filter((c) => c.truckId === t.id);
              return (
                <div
                  key={t.id}
                  onClick={() => setSelectedTruckId(t.id)}
                  style={{
                    background: active ? C.surfaceAlt : "transparent",
                    border: `1px solid ${active ? C.amberDim : "transparent"}`,
                    borderRadius: 8,
                    padding: "10px 10px",
                    marginBottom: 6,
                    cursor: "pointer",
                  }}
                >
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                    <div>
                      <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 600, fontSize: 13.5, color: active ? C.amber : C.text }}>
                        {t.name}
                      </div>
                      <div style={{ fontFamily: FONT_MONO, fontSize: 11, color: C.textDim, marginTop: 2 }}>
                        {t.plate || "no plate"}
                      </div>
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
                      }}
                      title="Remove truck"
                    >
                      <X size={13} />
                    </button>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", marginTop: 8, fontSize: 11, fontFamily: FONT_MONO, color: C.textDim }}>
                    <span>{tContracts.length} contract{tContracts.length === 1 ? "" : "s"}</span>
                    <span style={{ color: C.green }}>{money(truckTotals(t.id))}</span>
                  </div>
                </div>
              );
            })}

            {trucks.length === 0 && (
              <div style={{ fontSize: 12, color: C.textFaint, padding: "12px 6px" }}>
                No trucks yet. Add your first truck below.
              </div>
            )}
          </div>

          <div style={{ padding: 12, borderTop: `1px solid ${C.border}` }}>
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
                <input
                  value={newTruckPlate}
                  onChange={(e) => setNewTruckPlate(e.target.value)}
                  placeholder="Plate / VIN (optional)"
                  onKeyDown={(e) => e.key === "Enter" && addTruck()}
                  style={{
                    background: C.surfaceAlt,
                    border: `1px solid ${C.borderLight}`,
                    borderRadius: 5,
                    color: C.text,
                    fontSize: 12.5,
                    padding: "7px 8px",
                    outline: "none",
                    fontFamily: FONT_MONO,
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
                      setNewTruckPlate("");
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
                  padding: "18px 24px 14px",
                  borderBottom: `1px solid ${C.border}`,
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "flex-end",
                }}
              >
                <div>
                  <div style={{ fontSize: 11, fontFamily: FONT_MONO, color: C.textFaint, letterSpacing: 0.6, textTransform: "uppercase" }}>
                    Truck
                  </div>
                  <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 700, fontSize: 22, display: "flex", alignItems: "center", gap: 10 }}>
                    {selectedTruck.name}
                    {selectedTruck.plate && (
                      <span
                        style={{
                          fontFamily: FONT_MONO,
                          fontSize: 12,
                          background: C.surfaceAlt,
                          border: `1px solid ${C.borderLight}`,
                          borderRadius: 4,
                          padding: "3px 7px",
                          color: C.textDim,
                        }}
                      >
                        {selectedTruck.plate}
                      </span>
                    )}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 20 }}>
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

              {/* Table */}
              <div style={{ flex: 1, overflow: "auto", padding: "0 24px 24px" }}>
                <div style={{ minWidth: 1050, marginTop: 14 }}>
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
                    <Cell width={colWidths.block}># of blocks × Price</Cell>
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
                          <span style={{ fontFamily: FONT_MONO, fontSize: 12, color: C.textDim }}>{isoWeek(c.periodStart) || "—"}</span>
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
            </>
          )}
        </div>
      </div>
    </div>
  );
}
