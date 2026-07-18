import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  Package, Plus, Search, ArrowDownCircle, ArrowUpCircle, Pencil, Trash2,
  AlertTriangle, X, History, Boxes, Wallet, MapPin, ChevronDown, ChevronRight, Layers,
  Factory, Undo2, Calendar, FileSpreadsheet, ArrowRight, Scissors, ClipboardList, Upload,
  SlidersHorizontal, Ban, Printer, Tag, Users, LogOut, ShoppingCart, BarChart3, Shield, Eye, UserPlus,
  Settings, Globe, ZoomIn, Trash, Clock, TrendingDown, Download
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer
} from "recharts";

const UNIT_OPTIONS = ["ชิ้น", "กล่อง", "แพ็ค", "พาเลท", "ลัง", "กิโลกรัม", "เมตร"];
const CATEGORIES = ["วัตถุดิบ", "บรรจุภัณฑ์", "สินค้าสำเร็จรูป", "อื่นๆ"];
const ROLES = [
  { id: "admin", label: "ผู้ดูแลระบบ" },
  { id: "accounting", label: "แผนกบัญชี" },
  { id: "staff", label: "พนักงาน" },
  { id: "viewer", label: "ดูอย่างเดียว" },
];
function roleLabel(role) {
  return ROLES.find((r) => r.id === role)?.label || role;
}
const WASTE_REASONS = ["สินค้าเสียหาย", "หมดอายุ", "สูญหาย", "อื่นๆ"];
const ADJUST_REASONS = ["นับสต็อกจริงไม่ตรง", "แก้ไขข้อมูลผิดพลาด", "อื่นๆ"];

// ---- accounting module ----
const ACCOUNT_TYPES = [
  { id: "asset", label: "สินทรัพย์" },
  { id: "liability", label: "หนี้สิน" },
  { id: "equity", label: "ทุน" },
  { id: "revenue", label: "รายได้" },
  { id: "expense", label: "ค่าใช้จ่าย" },
];
const DEFAULT_ACCOUNTS = [
  { id: "acc-cash", code: "1000", name: "เงินสด/เงินฝากธนาคาร", type: "asset" },
  { id: "acc-ar", code: "1100", name: "ลูกหนี้การค้า", type: "asset" },
  { id: "acc-inventory", code: "1200", name: "สินค้าคงเหลือ", type: "asset" },
  { id: "acc-ap", code: "2000", name: "เจ้าหนี้การค้า", type: "liability" },
  { id: "acc-equity", code: "3000", name: "ทุน", type: "equity" },
  { id: "acc-sales", code: "4000", name: "รายได้จากการขาย", type: "revenue" },
  { id: "acc-cogs", code: "5000", name: "ต้นทุนขาย", type: "expense" },
  { id: "acc-opex", code: "5100", name: "ค่าใช้จ่ายในการดำเนินงาน", type: "expense" },
  { id: "acc-writeoff", code: "5200", name: "ของเสีย/สูญหาย/ปรับปรุงสต็อก", type: "expense" },
];
function accountTypeLabel(t) {
  return ACCOUNT_TYPES.find((a) => a.id === t)?.label || t;
}
// Asset/Expense accounts increase with a debit; Liability/Equity/Revenue accounts increase with a credit.
function isDebitNormal(type) {
  return type === "asset" || type === "expense";
}

// ---- system settings ----
const DISPLAY_SIZES = [
  { id: "small", label: "เล็ก", scale: 0.9 },
  { id: "medium", label: "กลาง", scale: 1 },
  { id: "large", label: "ใหญ่", scale: 1.15 },
];
const PAPER_SIZES = ["A4", "A5"];
const DEFAULT_SETTINGS = {
  displaySize: "medium",
  language: "th",
  printHeaderText: "",
  printPaperSize: "A4",
  approvalThreshold: 0,
  lastBackupTs: 0,
};

// Small translation table covering the main dashboard chrome. Deep forms/modals stay in
// Thai only for now — translating every label across the whole app is a much larger job.
const TRANSLATIONS = {
  appTitle: { th: "ระบบผู้ช่วยคลังสินค้า", en: "Warehouse Assistant System" },
  appSubtitle: {
    th: "จัดการสต็อกแบบแยกล็อต รับเข้า-เบิกออก แจ้งเตือนของใกล้หมด/ใกล้หมดอายุ",
    en: "Lot-tracked stock management, receiving/issuing, low-stock & expiry alerts",
  },
  dashboard: { th: "แดชบอร์ดวิเคราะห์", en: "Analytics dashboard" },
  prodPanel: { th: "ฝ่ายผลิต SSN/SD", en: "Production (SSN/SD)" },
  history: { th: "ประวัติการเคลื่อนไหว", en: "Movement history" },
  addItem: { th: "เพิ่มสินค้า", en: "Add item" },
  cutMaterial: { th: "ตัดวัตถุดิบ", en: "Issue material" },
  fgCheck: { th: "ตรวจรับสินค้าสำเร็จรูป", en: "Finished goods QC" },
  exportExcel: { th: "ส่งออก Excel", en: "Export Excel" },
  importExcel: { th: "นำเข้า Excel", en: "Import Excel" },
  countSheet: { th: "ใบนับสต็อก", en: "Count sheet" },
  po: { th: "ใบสั่งซื้อ (PO)", en: "Purchase orders (PO)" },
  statItems: { th: "รายการสินค้า", en: "Items" },
  statUnits: { th: "จำนวนรวม", en: "Total units" },
  statLow: { th: "ใกล้หมดสต็อก", en: "Low stock" },
  statExpiring: { th: "ใกล้หมดอายุ", en: "Expiring soon" },
  statValue: { th: "มูลค่าสต็อกรวม", en: "Total stock value" },
  settings: { th: "ตั้งค่าระบบ", en: "System settings" },
  colProduct: { th: "สินค้า", en: "Product" },
  colLocation: { th: "ตำแหน่ง", en: "Location" },
  colQty: { th: "คงเหลือ", en: "On hand" },
  colValue: { th: "มูลค่า/หน่วย", en: "Unit value" },
  colActions: { th: "การจัดการ", en: "Actions" },
};

// Fixed set of physical storage locations/warehouses.
const LOCATIONS = ["RM01", "RM02", "RM03", "RM11", "RM13", "RM12", "FG", "IR01", "IR02", "SSN", "SD"];
// Which of the locations above are production departments (where raw materials get
// transferred to be used in a formula/recipe, then cut/consumed once production is done).
const PROD_LOCATIONS = ["SSN", "SD"];
// Older versions of this app used one generic "PROD" bucket instead of named departments.
// Kept here so previously-saved data still shows up and can still be managed.
const LEGACY_PROD = "PROD";

function isProdLocation(loc) {
  return loc === LEGACY_PROD || PROD_LOCATIONS.includes(loc);
}

function locationLabel(loc) {
  if (loc === LEGACY_PROD) return "ฝ่ายผลิต (เดิม)";
  return loc || "ไม่ระบุ";
}

// Reverse of locationLabel: turns an imported label/code back into the internal location value.
function parseLocationLabel(value) {
  const s = String(value || "").trim();
  if (!s) return "";
  if (s === "ฝ่ายผลิต (เดิม)" || s === "ฝ่ายผลิต") return LEGACY_PROD;
  const match = LOCATIONS.find((l) => l.toLowerCase() === s.toLowerCase());
  return match || s;
}

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

// Client-side password hashing. This is a convenience deterrent (stops someone from just
// clicking a name to get in), NOT enterprise-grade security — there's no server here to
// keep a real secret from a determined user of the same browser.
async function hashPassword(pw) {
  try {
    const enc = new TextEncoder().encode(pw);
    const buf = await crypto.subtle.digest("SHA-256", enc);
    return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
  } catch {
    let h = 0;
    for (let i = 0; i < pw.length; i++) h = (h * 31 + pw.charCodeAt(i)) >>> 0;
    return "fallback-" + h.toString(16);
  }
}

function formatBaht(n) {
  return new Intl.NumberFormat("th-TH", { style: "currency", currency: "THB", maximumFractionDigits: 0 }).format(n || 0);
}

function formatNum(n) {
  return new Intl.NumberFormat("th-TH").format(n || 0);
}

function formatDateTime(ts) {
  const d = new Date(ts);
  return d.toLocaleString("th-TH", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" });
}

function todayStr() {
  const d = new Date();
  return d.toISOString().slice(0, 10);
}

function formatDateOnly(dateStr) {
  if (!dateStr) return "";
  const [y, m, d] = dateStr.split("-");
  if (!y || !m || !d) return dateStr;
  return `${d}/${m}/${y}`;
}

// Parses dd/mm/yyyy (our display/export format), a JS Date (from Excel date cells),
// or an already-ISO yyyy-mm-dd string back into the internal yyyy-mm-dd format.
function parseDateOnly(value) {
  if (!value && value !== 0) return "";
  if (value instanceof Date && !isNaN(value)) {
    return value.toISOString().slice(0, 10);
  }
  const s = String(value).trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const [, d, mo, y] = m;
    return `${y}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  return "";
}

function movementTypeLabel(type) {
  if (type === "in") return "รับเข้า";
  if (type === "out") return "เบิกออก";
  if (type === "to_prod") return "โอนไปผลิต";
  if (type === "from_prod") return "คืนจากผลิต";
  if (type === "consume_prod") return "ตัดวัตถุดิบ (ผลิต)";
  if (type === "waste") return "ตัดของเสีย/สูญหาย";
  if (type === "adjust") return "ปรับปรุงสต็อก";
  return type;
}

function daysUntil(dateStr) {
  if (!dateStr) return null;
  const target = new Date(dateStr + "T00:00:00");
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  return Math.round((target - now) / 86400000);
}

// ---- lot helpers ----
// Each item stores its stock as an array of lots: { id, location, productionDate, expiryDate, qty, receivedDate }
// so stock received on different production dates or held at different locations never gets merged together.
function totalQty(item) {
  return (item.lots || []).reduce((s, l) => s + Number(l.qty || 0), 0);
}

// Quantity available in physical warehouse locations (excludes stock currently at any production department).
function warehouseQty(item) {
  return (item.lots || []).filter((l) => !isProdLocation(l.location)).reduce((s, l) => s + Number(l.qty || 0), 0);
}

// Total quantity currently held across all production departments (SSN + SD + legacy).
function productionQty(item) {
  return (item.lots || []).filter((l) => isProdLocation(l.location)).reduce((s, l) => s + Number(l.qty || 0), 0);
}

function qtyAtLocation(item, loc) {
  return (item.lots || []).filter((l) => l.location === loc).reduce((s, l) => s + Number(l.qty || 0), 0);
}

// Distinct locations currently holding stock for an item, e.g. [["RM01", 80], ["SSN", 10]]
function locationBreakdown(item) {
  const map = new Map();
  (item.lots || []).forEach((l) => {
    if (l.qty > 0) map.set(l.location || "", (map.get(l.location || "") || 0) + Number(l.qty || 0));
  });
  return [...map.entries()];
}

function nearestExpiry(item) {
  const withExpiry = (item.lots || []).filter((l) => l.qty > 0 && l.expiryDate);
  if (!withExpiry.length) return null;
  return withExpiry.reduce((min, l) => (min === null || l.expiryDate < min ? l.expiryDate : min), null);
}

// FEFO sort: earliest expiry first, lots without expiry go last.
function sortFEFO(lots) {
  return [...lots].sort((a, b) => {
    if (!a.expiryDate && !b.expiryDate) return 0;
    if (!a.expiryDate) return 1;
    if (!b.expiryDate) return -1;
    return a.expiryDate.localeCompare(b.expiryDate);
  });
}

// Merge a moved quantity into a destination location, reusing a lot with the same
// production date if one already exists there, otherwise opening a new lot.
function mergeIntoLocation(lots, destLocation, sourceLot, qty, receivedDate) {
  const matchIdx = lots.findIndex((l) => l.location === destLocation && (l.productionDate || "") === (sourceLot.productionDate || ""));
  if (matchIdx >= 0) {
    return lots.map((l, idx) => (idx === matchIdx ? { ...l, qty: Number(l.qty || 0) + qty } : l));
  }
  return [
    ...lots,
    {
      id: uid(),
      location: destLocation,
      productionDate: sourceLot.productionDate,
      expiryDate: sourceLot.expiryDate,
      qty,
      receivedDate: receivedDate || sourceLot.receivedDate || todayStr(),
    },
  ];
}

// Migrate legacy items (saved before lot tracking / locations existed) into the current model.
function migrateItem(raw) {
  const withCategory = raw.category ? raw : { ...raw, category: CATEGORIES[0] };
  if (Array.isArray(withCategory.lots)) {
    // Ensure every lot has a location (older lot-tracking version didn't have one).
    const needsLocation = withCategory.lots.some((l) => l.location === undefined);
    if (!needsLocation) return withCategory;
    return { ...withCategory, lots: withCategory.lots.map((l) => ({ location: withCategory.location || "", ...l })) };
  }
  const legacyQty = Number(withCategory.qty || 0);
  const lots =
    legacyQty > 0
      ? [
          {
            id: uid(),
            location: withCategory.location || "",
            productionDate: withCategory.productionDate || "",
            expiryDate: withCategory.expiryDate || "",
            qty: legacyQty,
            receivedDate: "",
          },
        ]
      : [];
  const { qty, productionDate, expiryDate, location, ...rest } = withCategory;
  return { ...rest, lots };
}

const BarcodeDivider = () => (
  <div
    className="h-2 w-full rounded-sm opacity-70"
    style={{
      backgroundImage:
        "repeating-linear-gradient(90deg, #FBBF24 0px, #FBBF24 2px, transparent 2px, transparent 5px, #78716C 5px, #78716C 6px, transparent 6px, transparent 10px)",
    }}
  />
);

const LocationTag = ({ location, qty, unit }) => {
  const isProd = isProdLocation(location);
  return (
    <span
      className={`inline-flex items-center gap-1 rounded border px-2 py-0.5 font-mono text-xs tracking-wide ${
        isProd
          ? "border-indigo-500/40 bg-indigo-500/10 text-indigo-300"
          : "border-amber-500/40 bg-amber-500/10 text-amber-300"
      }`}
    >
      {isProd ? <Factory size={11} className="shrink-0" /> : <MapPin size={11} className="shrink-0" />}
      {locationLabel(location)}
      {qty !== undefined && <span className="text-neutral-400">· {formatNum(qty)}{unit ? ` ${unit}` : ""}</span>}
    </span>
  );
};

function StatCard({ icon: Icon, label, value, accent }) {
  return (
    <div className="flex items-center gap-3 rounded-lg border border-neutral-800 bg-neutral-900 px-4 py-3">
      <div className={`flex h-10 w-10 shrink-0 items-center justify-center rounded-md ${accent}`}>
        <Icon size={20} />
      </div>
      <div className="min-w-0">
        <p className="truncate text-xs uppercase tracking-wider text-neutral-500">{label}</p>
        <p className="truncate font-mono text-lg font-bold text-neutral-100">{value}</p>
      </div>
    </div>
  );
}

function Modal({ title, onClose, children, footer }) {
  return (
    <div className="fixed inset-0 z-50 overflow-y-auto bg-black/70 p-4">
      <div className="flex min-h-full items-center justify-center">
        <div className="flex max-h-[85vh] w-full max-w-md flex-col rounded-lg border border-neutral-700 bg-neutral-900 shadow-2xl">
          <div className="flex shrink-0 items-center justify-between border-b border-neutral-800 px-5 py-3">
            <h3 className="font-bold uppercase tracking-wide text-neutral-100">{title}</h3>
            <button onClick={onClose} className="rounded p-1 text-neutral-400 hover:bg-neutral-800 hover:text-neutral-100">
              <X size={18} />
            </button>
          </div>
          <div className="min-h-0 flex-1 overflow-y-auto p-5">{children}</div>
          {footer && <div className="shrink-0 border-t border-neutral-800 px-5 py-3">{footer}</div>}
        </div>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="mb-3 block">
      <span className="mb-1 block text-xs font-semibold uppercase tracking-wide text-neutral-400">{label}</span>
      {children}
    </label>
  );
}

const inputCls =
  "w-full rounded-md border border-neutral-700 bg-neutral-800 px-3 py-2 text-sm text-neutral-100 outline-none focus:border-amber-500 focus:ring-1 focus:ring-amber-500";

// Shows dates as วว/ดด/ปปปป regardless of browser/OS locale, while still using the
// native date picker underneath (kept invisible on top) for familiar pick-a-date UX.
function DateField({ value, onChange }) {
  return (
    <div className="relative">
      <div className={`${inputCls} flex items-center justify-between gap-2`}>
        <span className={value ? "text-neutral-100" : "text-neutral-500"}>
          {value ? formatDateOnly(value) : "วว/ดด/ปปปป"}
        </span>
        <Calendar size={14} className="shrink-0 text-neutral-500" />
      </div>
      <input
        type="date"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="absolute inset-0 h-full w-full cursor-pointer opacity-0"
      />
    </div>
  );
}

export default function WarehouseApp() {
  const [items, setItems] = useState([]);
  const [transactions, setTransactions] = useState([]);
  const [jobs, setJobs] = useState([]);
  const [users, setUsers] = useState([]);
  const [currentUser, setCurrentUser] = useState(null);
  const [purchaseOrders, setPurchaseOrders] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("all");
  const [categoryFilter, setCategoryFilter] = useState("all");
  const [showForm, setShowForm] = useState(false);
  const [editingItem, setEditingItem] = useState(null);
  const [movementItem, setMovementItem] = useState(null);
  const [movementType, setMovementType] = useState("in");
  const [movementAmount, setMovementAmount] = useState("");
  const [movementNote, setMovementNote] = useState("");
  const [movementReceivedDate, setMovementReceivedDate] = useState("");
  const [movementProductionDate, setMovementProductionDate] = useState("");
  const [movementExpiryDate, setMovementExpiryDate] = useState("");
  const [movementRequestedBy, setMovementRequestedBy] = useState("");
  const [movementReason, setMovementReason] = useState(WASTE_REASONS[0]);
  const [movementLocation, setMovementLocation] = useState(LOCATIONS[0]);
  const [movementDept, setMovementDept] = useState(PROD_LOCATIONS[0]);
  const [movementEntryUnit, setMovementEntryUnit] = useState("base");
  const [showHistory, setShowHistory] = useState(false);
  const [showProdPanel, setShowProdPanel] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [saveError, setSaveError] = useState("");
  const [bannerDismissed, setBannerDismissed] = useState(false);
  const [backupBannerDismissed, setBackupBannerDismissed] = useState(false);
  const [expandedItemId, setExpandedItemId] = useState(null);
  const [alertMessage, setAlertMessage] = useState("");
  const [historySearch, setHistorySearch] = useState("");
  const [historyTypeFilter, setHistoryTypeFilter] = useState("all");
  const [showJobForm, setShowJobForm] = useState(false);
  const [jobForm, setJobForm] = useState(null);
  const [jobReceipt, setJobReceipt] = useState(null);
  const [fgChecks, setFgChecks] = useState([]);
  const [showFgForm, setShowFgForm] = useState(false);
  const [fgForm, setFgForm] = useState(null);
  const [fgReceipt, setFgReceipt] = useState(null);
  const [showAdjustForm, setShowAdjustForm] = useState(false);
  const [adjustForm, setAdjustForm] = useState(null);
  const [showPoPanel, setShowPoPanel] = useState(false);
  const [showPoForm, setShowPoForm] = useState(false);
  const [poForm, setPoForm] = useState(null);
  const [receivingPo, setReceivingPo] = useState(null);
  const [receiveLocation, setReceiveLocation] = useState(LOCATIONS[0]);
  const [receiveDateInput, setReceiveDateInput] = useState(todayStr());
  const [confirmReversePo, setConfirmReversePo] = useState(null);
  const [accounts, setAccounts] = useState(DEFAULT_ACCOUNTS);
  const [journalEntries, setJournalEntries] = useState([]);
  const [showAccounting, setShowAccounting] = useState(false);
  const [accountingTab, setAccountingTab] = useState("journal");
  const [showJournalForm, setShowJournalForm] = useState(false);
  const [journalForm, setJournalForm] = useState(null);
  const [showAccountForm, setShowAccountForm] = useState(false);
  const [accountForm, setAccountForm] = useState(null);
  const [payingPo, setPayingPo] = useState(null);
  const [poPaymentForm, setPoPaymentForm] = useState(null);
  const [ledgerAccountId, setLedgerAccountId] = useState("");
  const [pnlFrom, setPnlFrom] = useState("");
  const [pnlTo, setPnlTo] = useState("");
  const [saleForItem, setSaleForItem] = useState(null);
  const [itemEdits, setItemEdits] = useState([]);
  const [showEditLog, setShowEditLog] = useState(false);
  const [editLogSearch, setEditLogSearch] = useState("");
  const [approvals, setApprovals] = useState([]);
  const [showApprovals, setShowApprovals] = useState(false);
  const [pendingMovement, setPendingMovement] = useState(null);
  const [deadStockDays, setDeadStockDays] = useState(30);
  const [showDashboard, setShowDashboard] = useState(false);
  const [loginName, setLoginName] = useState("");
  const [loginRole, setLoginRole] = useState("staff");
  const [loginPassword, setLoginPassword] = useState("");
  const [attemptingUser, setAttemptingUser] = useState(null);
  const [attemptPassword, setAttemptPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [resettingPwUser, setResettingPwUser] = useState(null);
  const [resetPwValue, setResetPwValue] = useState("");
  const [showUserManager, setShowUserManager] = useState(false);
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [showSettings, setShowSettings] = useState(false);
  const [confirmClearHistory, setConfirmClearHistory] = useState(false);
  const fileInputRef = useRef(null);
  const restoreInputRef = useRef(null);

  // ---- load persisted data ----
  useEffect(() => {
    (async () => {
      try {
        const itemsRes = await window.storage.get("wh:items", false).catch(() => null);
        const txRes = await window.storage.get("wh:transactions", false).catch(() => null);
        const jobsRes = await window.storage.get("wh:jobs", false).catch(() => null);
        const fgRes = await window.storage.get("wh:fgchecks", false).catch(() => null);
        const usersRes = await window.storage.get("wh:users", false).catch(() => null);
        const curUserRes = await window.storage.get("wh:currentuser", false).catch(() => null);
        const posRes = await window.storage.get("wh:pos", false).catch(() => null);
        const settingsRes = await window.storage.get("wh:settings", false).catch(() => null);
        const accountsRes = await window.storage.get("wh:accounts", false).catch(() => null);
        const journalRes = await window.storage.get("wh:journal", false).catch(() => null);
        const editsRes = await window.storage.get("wh:itemedits", false).catch(() => null);
        const approvalsRes = await window.storage.get("wh:approvals", false).catch(() => null);
        const rawItems = itemsRes ? JSON.parse(itemsRes.value) : [];
        setItems(rawItems.map(migrateItem));
        setTransactions(txRes ? JSON.parse(txRes.value) : []);
        setJobs(jobsRes ? JSON.parse(jobsRes.value) : []);
        setFgChecks(fgRes ? JSON.parse(fgRes.value) : []);
        setUsers(usersRes ? JSON.parse(usersRes.value) : []);
        setCurrentUser(curUserRes ? JSON.parse(curUserRes.value) : null);
        setPurchaseOrders(posRes ? JSON.parse(posRes.value) : []);
        setSettings(settingsRes ? { ...DEFAULT_SETTINGS, ...JSON.parse(settingsRes.value) } : DEFAULT_SETTINGS);
        setAccounts(accountsRes ? JSON.parse(accountsRes.value) : DEFAULT_ACCOUNTS);
        setJournalEntries(journalRes ? JSON.parse(journalRes.value) : []);
        setItemEdits(editsRes ? JSON.parse(editsRes.value) : []);
        setApprovals(approvalsRes ? JSON.parse(approvalsRes.value) : []);
      } catch (e) {
        setSaveError("โหลดข้อมูลไม่สำเร็จ เริ่มต้นด้วยคลังว่างเปล่า");
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  const persistItems = useCallback(async (next) => {
    setItems(next);
    try {
      const res = await window.storage.set("wh:items", JSON.stringify(next), false);
      if (!res) setSaveError("บันทึกข้อมูลสินค้าไม่สำเร็จ");
      else setSaveError("");
    } catch {
      setSaveError("บันทึกข้อมูลสินค้าไม่สำเร็จ");
    }
  }, []);

  const persistTransactions = useCallback(async (next) => {
    const trimmed = next.slice(0, 300);
    setTransactions(trimmed);
    try {
      const res = await window.storage.set("wh:transactions", JSON.stringify(trimmed), false);
      if (!res) setSaveError("บันทึกประวัติไม่สำเร็จ");
    } catch {
      setSaveError("บันทึกประวัติไม่สำเร็จ");
    }
  }, []);

  const persistJobs = useCallback(async (next) => {
    const trimmed = next.slice(0, 200);
    setJobs(trimmed);
    try {
      const res = await window.storage.set("wh:jobs", JSON.stringify(trimmed), false);
      if (!res) setSaveError("บันทึกใบตัดวัตถุดิบไม่สำเร็จ");
    } catch {
      setSaveError("บันทึกใบตัดวัตถุดิบไม่สำเร็จ");
    }
  }, []);

  const persistFgChecks = useCallback(async (next) => {
    const trimmed = next.slice(0, 200);
    setFgChecks(trimmed);
    try {
      const res = await window.storage.set("wh:fgchecks", JSON.stringify(trimmed), false);
      if (!res) setSaveError("บันทึกใบตรวจรับสินค้าสำเร็จรูปไม่สำเร็จ");
    } catch {
      setSaveError("บันทึกใบตรวจรับสินค้าสำเร็จรูปไม่สำเร็จ");
    }
  }, []);

  const persistUsers = useCallback(async (next) => {
    setUsers(next);
    try {
      await window.storage.set("wh:users", JSON.stringify(next), false);
    } catch {
      setSaveError("บันทึกรายชื่อผู้ใช้งานไม่สำเร็จ");
    }
  }, []);

  const loginAs = useCallback(async (user) => {
    setCurrentUser(user);
    try {
      await window.storage.set("wh:currentuser", JSON.stringify(user), false);
    } catch {
      /* non-fatal: just won't be remembered on next load */
    }
  }, []);

  const logout = useCallback(async () => {
    setCurrentUser(null);
    try {
      await window.storage.delete("wh:currentuser", false);
    } catch {
      /* non-fatal */
    }
  }, []);

  function removeUser(userId) {
    persistUsers(users.filter((u) => u.id !== userId));
  }

  function changeUserRole(userId, role) {
    const next = users.map((u) => (u.id === userId ? { ...u, role } : u));
    persistUsers(next);
    if (currentUser?.id === userId) {
      loginAs({ ...currentUser, role });
    }
  }

  const persistPurchaseOrders = useCallback(async (next) => {
    const trimmed = next.slice(0, 200);
    setPurchaseOrders(trimmed);
    try {
      const res = await window.storage.set("wh:pos", JSON.stringify(trimmed), false);
      if (!res) setSaveError("บันทึกใบสั่งซื้อไม่สำเร็จ");
    } catch {
      setSaveError("บันทึกใบสั่งซื้อไม่สำเร็จ");
    }
  }, []);

  const persistAccounts = useCallback(async (next) => {
    setAccounts(next);
    try {
      await window.storage.set("wh:accounts", JSON.stringify(next), false);
    } catch {
      setSaveError("บันทึกผังบัญชีไม่สำเร็จ");
    }
  }, []);

  const persistJournal = useCallback(async (next) => {
    const trimmed = next.slice(0, 500);
    setJournalEntries(trimmed);
    try {
      const res = await window.storage.set("wh:journal", JSON.stringify(trimmed), false);
      if (!res) setSaveError("บันทึกรายการบัญชีไม่สำเร็จ");
    } catch {
      setSaveError("บันทึกรายการบัญชีไม่สำเร็จ");
    }
  }, []);

  const persistItemEdits = useCallback(async (next) => {
    const trimmed = next.slice(0, 500);
    setItemEdits(trimmed);
    try {
      await window.storage.set("wh:itemedits", JSON.stringify(trimmed), false);
    } catch {
      setSaveError("บันทึกประวัติการแก้ไขสินค้าไม่สำเร็จ");
    }
  }, []);

  const persistApprovals = useCallback(async (next) => {
    const trimmed = next.slice(0, 300);
    setApprovals(trimmed);
    try {
      const res = await window.storage.set("wh:approvals", JSON.stringify(trimmed), false);
      if (!res) setSaveError("บันทึกคำขออนุมัติไม่สำเร็จ");
    } catch {
      setSaveError("บันทึกคำขออนุมัติไม่สำเร็จ");
    }
  }, []);

  const persistSettings = useCallback(async (next) => {
    setSettings(next);
    try {
      await window.storage.set("wh:settings", JSON.stringify(next), false);
    } catch {
      setSaveError("บันทึกการตั้งค่าไม่สำเร็จ");
    }
  }, []);

  function updateSetting(key, value) {
    persistSettings({ ...settings, [key]: value });
  }

  function clearHistory() {
    persistTransactions([]);
    setConfirmClearHistory(false);
  }

  // Translates a small set of top-level chrome strings; everything else stays in Thai.
  function t(key) {
    return TRANSLATIONS[key]?.[settings.language] || TRANSLATIONS[key]?.th || key;
  }

  const displayScale = DISPLAY_SIZES.find((d) => d.id === settings.displaySize)?.scale || 1;

  const isViewer = currentUser?.role === "viewer";
  const isAdmin = currentUser?.role === "admin";
  const isAccounting = currentUser?.role === "accounting" || isAdmin;
  const canEdit = !!currentUser && !isViewer;

  // ---- derived ----
  const filteredItems = useMemo(() => {
    const q = search.trim().toLowerCase();
    return items
      .filter((it) => {
        if (filter === "low" && totalQty(it) > it.lowStock) return false;
        if (categoryFilter !== "all" && (it.category || CATEGORIES[0]) !== categoryFilter) return false;
        if (!q) return true;
        return (
          it.name.toLowerCase().includes(q) ||
          it.sku.toLowerCase().includes(q) ||
          locationBreakdown(it).some(([loc]) => locationLabel(loc).toLowerCase().includes(q))
        );
      })
      .sort((a, b) => a.name.localeCompare(b.name, "th"));
  }, [items, search, filter, categoryFilter]);

  const filteredTransactions = useMemo(() => {
    const q = historySearch.trim().toLowerCase();
    return transactions.filter((tx) => {
      if (historyTypeFilter !== "all" && tx.type !== historyTypeFilter) return false;
      if (!q) return true;
      return (
        tx.itemName.toLowerCase().includes(q) ||
        tx.sku.toLowerCase().includes(q) ||
        locationLabel(tx.fromLocation).toLowerCase().includes(q) ||
        locationLabel(tx.toLocation).toLowerCase().includes(q)
      );
    });
  }, [transactions, historySearch, historyTypeFilter]);

  const stats = useMemo(() => {
    const totalItems = items.length;
    const totalUnits = items.reduce((s, i) => s + totalQty(i), 0);
    const lowStockCount = items.filter((i) => totalQty(i) <= i.lowStock).length;
    const totalValue = items.reduce((s, i) => s + totalQty(i) * Number(i.price || 0), 0);
    const expiringCount = items.filter((i) => {
      const d = daysUntil(nearestExpiry(i));
      return d !== null && d <= 30;
    }).length;
    return { totalItems, totalUnits, lowStockCount, totalValue, expiringCount };
  }, [items]);

  const categoryValueData = useMemo(() => {
    return CATEGORIES.map((cat) => ({
      name: cat,
      value: items
        .filter((i) => (i.category || CATEGORIES[0]) === cat)
        .reduce((s, i) => s + totalQty(i) * Number(i.price || 0), 0),
    })).filter((d) => d.value > 0);
  }, [items]);

  const dailyMovementData = useMemo(() => {
    const days = [];
    for (let i = 6; i >= 0; i--) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      days.push(d.toISOString().slice(0, 10));
    }
    return days.map((dateStr) => {
      const dayTxs = transactions.filter((tx) => new Date(tx.ts).toISOString().slice(0, 10) === dateStr);
      return {
        date: formatDateOnly(dateStr),
        รับเข้า: dayTxs.filter((t) => t.type === "in").length,
        เบิกออก: dayTxs.filter((t) => t.type === "out").length,
        ของเสีย: dayTxs.filter((t) => t.type === "waste").length,
      };
    });
  }, [transactions]);

  const deadStockItems = useMemo(() => {
    const cutoff = Date.now() - deadStockDays * 86400000;
    return items
      .map((it) => {
        const lastTx = transactions.find((tx) => tx.itemId === it.id);
        const qty = totalQty(it);
        return { item: it, qty, lastTs: lastTx ? lastTx.ts : null };
      })
      .filter((x) => x.qty > 0 && (x.lastTs === null || x.lastTs < cutoff))
      .sort((a, b) => (a.lastTs || 0) - (b.lastTs || 0));
  }, [items, transactions, deadStockDays]);

  const CHART_COLORS = ["#f59e0b", "#38bdf8", "#34d399", "#a78bfa", "#fb7185"];

  const trialBalance = useMemo(() => {
    return accounts.map((acc) => {
      let debit = 0;
      let credit = 0;
      journalEntries.forEach((j) => {
        j.lines.forEach((l) => {
          if (l.accountId === acc.id) {
            debit += Number(l.debit || 0);
            credit += Number(l.credit || 0);
          }
        });
      });
      const balance = isDebitNormal(acc.type) ? debit - credit : credit - debit;
      return { ...acc, debit, credit, balance };
    });
  }, [accounts, journalEntries]);

  const trialTotals = useMemo(
    () => ({
      debit: trialBalance.reduce((s, a) => s + a.debit, 0),
      credit: trialBalance.reduce((s, a) => s + a.credit, 0),
    }),
    [trialBalance]
  );

  const ledgerEntries = useMemo(() => {
    if (!ledgerAccountId) return [];
    const acc = accounts.find((a) => a.id === ledgerAccountId);
    const rows = [];
    let running = 0;
    [...journalEntries]
      .sort((a, b) => a.date.localeCompare(b.date) || a.ts - b.ts)
      .forEach((j) => {
        j.lines.forEach((l) => {
          if (l.accountId !== ledgerAccountId) return;
          const delta = isDebitNormal(acc?.type) ? l.debit - l.credit : l.credit - l.debit;
          running += delta;
          rows.push({ ...j, debit: l.debit, credit: l.credit, running });
        });
      });
    return rows;
  }, [journalEntries, ledgerAccountId, accounts]);

  const pnlData = useMemo(() => {
    const from = pnlFrom || "0000-00-00";
    const to = pnlTo || "9999-99-99";
    let revenue = 0;
    let cogs = 0;
    let opex = 0;
    let writeoff = 0;
    journalEntries.forEach((j) => {
      if (j.date < from || j.date > to) return;
      j.lines.forEach((l) => {
        const acc = accounts.find((a) => a.id === l.accountId);
        if (!acc) return;
        if (acc.type === "revenue") revenue += l.credit - l.debit;
        if (acc.id === "acc-cogs") cogs += l.debit - l.credit;
        if (acc.id === "acc-opex") opex += l.debit - l.credit;
        if (acc.id === "acc-writeoff") writeoff += l.debit - l.credit;
      });
    });
    const grossProfit = revenue - cogs;
    const netIncome = grossProfit - opex - writeoff;
    return { revenue, cogs, grossProfit, opex, writeoff, netIncome };
  }, [journalEntries, accounts, pnlFrom, pnlTo]);

  // ---- item actions ----
  function openAddForm() {
    setEditingItem({
      id: null,
      name: "",
      sku: "",
      unit: UNIT_OPTIONS[0],
      category: CATEGORIES[0],
      lowStock: 5,
      price: 0,
      supplier: "",
      packUnit: "",
      packFactor: "",
      initialQty: 0,
      initialLocation: LOCATIONS[0],
      initialProductionDate: "",
      initialExpiryDate: "",
    });
    setShowForm(true);
  }

  function openEditForm(item) {
    setEditingItem({ supplier: "", category: CATEGORIES[0], packUnit: "", packFactor: "", ...item });
    setShowForm(true);
  }

  function saveItemForm() {
    const it = editingItem;
    if (!it.name.trim() || !it.sku.trim()) return;

    const dupe = items.find((x) => x.sku.toLowerCase() === it.sku.trim().toLowerCase() && x.id !== it.id);
    if (dupe) {
      setAlertMessage(`SKU "${it.sku}" ถูกใช้กับสินค้า "${dupe.name}" อยู่แล้ว กรุณาใช้ SKU อื่น`);
      return;
    }

    if (it.id) {
      const before = items.find((x) => x.id === it.id);
      const after = {
        ...before,
        name: it.name,
        sku: it.sku,
        unit: it.unit,
        category: it.category,
        lowStock: Number(it.lowStock) || 0,
        price: Number(it.price) || 0,
        supplier: it.supplier,
        packUnit: it.packUnit,
        packFactor: Number(it.packFactor) || 0,
      };
      if (before) {
        const fieldLabels = {
          name: "ชื่อสินค้า",
          sku: "SKU",
          unit: "หน่วยนับ",
          category: "หมวดหมู่",
          lowStock: "จุดแจ้งเตือนสต็อกต่ำ",
          price: "ราคาต่อหน่วย",
          supplier: "ผู้ขาย",
          packUnit: "หน่วยบรรจุ",
          packFactor: "อัตราแปลงหน่วยบรรจุ",
        };
        const changes = Object.keys(fieldLabels)
          .filter((f) => String(before[f] ?? "") !== String(after[f] ?? ""))
          .map((f) => ({ field: fieldLabels[f], from: before[f] ?? "", to: after[f] ?? "" }));
        if (changes.length) {
          persistItemEdits([
            {
              id: uid(),
              itemId: it.id,
              itemName: after.name,
              sku: after.sku,
              changes,
              changedBy: currentUser?.name || "",
              ts: Date.now(),
            },
            ...itemEdits,
          ]);
        }
      }
      persistItems(items.map((x) => (x.id === it.id ? after : x)));
    } else {
      const qty = Number(it.initialQty) || 0;
      const newItem = {
        id: uid(),
        name: it.name,
        sku: it.sku,
        unit: it.unit,
        category: it.category,
        lowStock: Number(it.lowStock) || 0,
        price: Number(it.price) || 0,
        supplier: it.supplier,
        packUnit: it.packUnit,
        packFactor: Number(it.packFactor) || 0,
        lots:
          qty > 0
            ? [
                {
                  id: uid(),
                  location: it.initialLocation || LOCATIONS[0],
                  productionDate: it.initialProductionDate,
                  expiryDate: it.initialExpiryDate,
                  qty,
                  receivedDate: todayStr(),
                },
              ]
            : [],
      };
      persistItems([newItem, ...items]);
    }
    setShowForm(false);
    setEditingItem(null);
  }

  function deleteItem(id) {
    persistItems(items.filter((x) => x.id !== id));
    setConfirmDelete(null);
  }

  // ---- stock adjustment (ปรับปรุงสต็อก): correct a lot's quantity after a physical count ----
  function openAdjustForm(item) {
    const firstLot = (item.lots || [])[0];
    setAdjustForm({
      itemId: item.id,
      lotId: firstLot ? firstLot.id : "new",
      location: firstLot ? firstLot.location : LOCATIONS[0],
      productionDate: firstLot ? firstLot.productionDate : "",
      expiryDate: firstLot ? firstLot.expiryDate : "",
      actualQty: firstLot ? String(firstLot.qty) : "",
      reason: ADJUST_REASONS[0],
      note: "",
    });
    setShowAdjustForm(true);
  }

  function selectAdjustLot(lotId, item) {
    if (lotId === "new") {
      setAdjustForm((f) => ({ ...f, lotId: "new", location: LOCATIONS[0], productionDate: "", expiryDate: "", actualQty: "" }));
      return;
    }
    const lot = (item.lots || []).find((l) => l.id === lotId);
    if (!lot) return;
    setAdjustForm((f) => ({
      ...f,
      lotId,
      location: lot.location,
      productionDate: lot.productionDate,
      expiryDate: lot.expiryDate,
      actualQty: String(lot.qty),
    }));
  }

  function submitAdjustForm() {
    const f = adjustForm;
    if (f.actualQty === "" || Number(f.actualQty) < 0) return;
    const item = items.find((i) => i.id === f.itemId);
    if (!item) return;
    const newQty = Number(f.actualQty);

    let oldQty = 0;
    let location = f.location;
    let newLots;
    if (f.lotId === "new") {
      newLots = [
        ...(item.lots || []),
        { id: uid(), location: f.location, productionDate: f.productionDate, expiryDate: f.expiryDate, qty: newQty, receivedDate: todayStr() },
      ];
    } else {
      const lot = (item.lots || []).find((l) => l.id === f.lotId);
      if (!lot) return;
      oldQty = Number(lot.qty || 0);
      location = lot.location;
      newLots = (item.lots || []).map((l) => (l.id === lot.id ? { ...l, qty: newQty } : l)).filter((l) => l.qty > 0);
    }

    const diff = newQty - oldQty;
    if (diff === 0) {
      setShowAdjustForm(false);
      setAdjustForm(null);
      return;
    }

    const depleted = newLots.reduce((s, l) => s + l.qty, 0) <= 0;
    if (depleted) {
      persistItems(items.filter((i) => i.id !== item.id));
    } else {
      persistItems(items.map((i) => (i.id === item.id ? { ...i, lots: newLots } : i)));
    }

    persistTransactions([
      {
        id: uid(),
        itemId: item.id,
        itemName: item.name,
        sku: item.sku,
        type: "adjust",
        amount: diff,
        note: f.note,
        receivedDate: "",
        productionDate: "",
        requestedBy: "",
        reason: f.reason,
        location,
        fromLocation: "",
        toLocation: "",
        lotSummary: `${diff > 0 ? "+" : ""}${formatNum(diff)} ที่ ${locationLabel(location)}`,
        ts: Date.now(),
      },
      ...transactions,
    ]);

    setShowAdjustForm(false);
    setAdjustForm(null);
    if (depleted) {
      setAlertMessage(`${item.name} (${item.sku}) ปรับสต็อกเหลือ 0 และถูกลบออกจากรายการสินค้าแล้ว`);
    }
  }

  // ---- purchase orders (ใบสั่งซื้อ) ----
  function openPoForm() {
    setPoForm({
      supplier: "",
      supplierAddress: "",
      taxInvoiceNo: "",
      orderDate: todayStr(),
      expectedDate: "",
      note: "",
      lines: [{ id: uid(), itemId: "", newName: "", newSku: "", qty: "", unitPrice: "", productionDate: "", expiryDate: "" }],
    });
    setShowPoForm(true);
  }

  function addPoLine() {
    setPoForm((f) => ({
      ...f,
      lines: [...f.lines, { id: uid(), itemId: "", newName: "", newSku: "", qty: "", unitPrice: "", productionDate: "", expiryDate: "" }],
    }));
  }

  function removePoLine(lineId) {
    setPoForm((f) => ({ ...f, lines: f.lines.filter((l) => l.id !== lineId) }));
  }

  function updatePoLine(lineId, patch) {
    setPoForm((f) => ({ ...f, lines: f.lines.map((l) => (l.id === lineId ? { ...l, ...patch } : l)) }));
  }

  function submitPoForm() {
    const supplier = poForm.supplier.trim();
    const validLines = poForm.lines.filter(
      (l) => (l.itemId || (l.newName.trim() && l.newSku.trim())) && Number(l.qty) > 0
    );
    if (!supplier || validLines.length === 0) return;

    const poNumber = `PO-${String(purchaseOrders.length + 1).padStart(4, "0")}`;
    const lines = validLines.map((l) => {
      const existing = l.itemId ? items.find((i) => i.id === l.itemId) : null;
      return {
        id: uid(),
        itemId: l.itemId || "",
        itemName: existing ? existing.name : l.newName.trim(),
        sku: existing ? existing.sku : l.newSku.trim(),
        unit: existing ? existing.unit : UNIT_OPTIONS[0],
        qty: Number(l.qty),
        unitPrice: Number(l.unitPrice) || 0,
        productionDate: l.productionDate || "",
        expiryDate: l.expiryDate || "",
      };
    });

    const po = {
      id: uid(),
      poNumber,
      supplier,
      supplierAddress: poForm.supplierAddress.trim(),
      taxInvoiceNo: poForm.taxInvoiceNo.trim(),
      orderDate: poForm.orderDate || todayStr(),
      expectedDate: poForm.expectedDate,
      status: "pending",
      lines,
      note: poForm.note,
      ts: Date.now(),
    };
    persistPurchaseOrders([po, ...purchaseOrders]);
    setShowPoForm(false);
    setPoForm(null);
  }

  function openReceivePo(po) {
    setReceivingPo(po);
    setReceiveLocation(LOCATIONS[0]);
    setReceiveDateInput(todayStr());
  }

  function confirmReceivePo() {
    const po = receivingPo;
    if (!po) return;
    let workingItems = items;
    const newTxs = [];
    const receivedDate = receiveDateInput || todayStr();

    po.lines.forEach((line) => {
      let item = line.itemId ? workingItems.find((i) => i.id === line.itemId) : workingItems.find((i) => i.sku.toLowerCase() === line.sku.toLowerCase());
      if (!item) {
        item = {
          id: uid(),
          name: line.itemName,
          sku: line.sku,
          unit: line.unit,
          category: CATEGORIES[0],
          lowStock: 5,
          price: line.unitPrice,
          supplier: po.supplier,
          lots: [],
        };
        workingItems = [item, ...workingItems];
      }
      const matchIdx = (item.lots || []).findIndex(
        (l) => l.location === receiveLocation && (l.productionDate || "") === (line.productionDate || "")
      );
      const newLots =
        matchIdx >= 0
          ? item.lots.map((l, idx) =>
              idx === matchIdx ? { ...l, qty: Number(l.qty || 0) + line.qty, expiryDate: line.expiryDate || l.expiryDate } : l
            )
          : [
              ...(item.lots || []),
              {
                id: uid(),
                location: receiveLocation,
                productionDate: line.productionDate || "",
                expiryDate: line.expiryDate || "",
                qty: line.qty,
                receivedDate,
              },
            ];
      workingItems = workingItems.map((i) => (i.id === item.id ? { ...i, lots: newLots, price: line.unitPrice || i.price } : i));

      newTxs.push({
        id: uid(),
        itemId: item.id,
        itemName: line.itemName,
        sku: line.sku,
        type: "in",
        amount: line.qty,
        note: `รับเข้าจากใบสั่งซื้อ ${po.poNumber} (${po.supplier})${po.taxInvoiceNo ? ` · ใบกำกับภาษี ${po.taxInvoiceNo}` : ""}`,
        receivedDate,
        productionDate: line.productionDate || "",
        requestedBy: "",
        location: receiveLocation,
        fromLocation: "",
        toLocation: receiveLocation,
        lotSummary: "",
        ts: Date.now(),
      });
    });

    persistItems(workingItems);
    persistTransactions([...newTxs, ...transactions]);
    persistPurchaseOrders(
      purchaseOrders.map((p) => (p.id === po.id ? { ...p, status: "received", receivedDate, receivedLocation: receiveLocation } : p))
    );
    setReceivingPo(null);
  }

  // ---- reverse a PO that was mistakenly marked received: deduct the stock it added back out ----
  function reversePoReceive(po) {
    let workingItems = items;
    const newTxs = [];
    const reverseDate = todayStr();
    let blocked = null;

    po.lines.forEach((line) => {
      if (blocked) return;
      const item = workingItems.find((i) => i.id === line.itemId || i.sku.toLowerCase() === line.sku.toLowerCase());
      if (!item) return;
      const available = qtyAtLocation(item, po.receivedLocation || LOCATIONS[0]);
      const take = Math.min(line.qty, available);
      if (take <= 0) return;

      const sourceLots = sortFEFO((item.lots || []).filter((l) => l.location === (po.receivedLocation || LOCATIONS[0])));
      let remaining = take;
      let newLots = item.lots || [];
      sourceLots.forEach((target) => {
        if (remaining <= 0) return;
        const t = Math.min(target.qty, remaining);
        remaining -= t;
        if (t > 0) newLots = newLots.map((l) => (l.id === target.id ? { ...l, qty: l.qty - t } : l));
      });
      newLots = newLots.filter((l) => l.qty > 0);
      const depleted = newLots.reduce((s, l) => s + l.qty, 0) <= 0;
      workingItems = depleted
        ? workingItems.filter((i) => i.id !== item.id)
        : workingItems.map((i) => (i.id === item.id ? { ...i, lots: newLots } : i));

      newTxs.push({
        id: uid(),
        itemId: item.id,
        itemName: line.itemName,
        sku: line.sku,
        type: "adjust",
        amount: -take,
        note: `ยกเลิกใบสั่งซื้อ ${po.poNumber} ที่รับผิด ตัดสต็อกที่รับเข้าไปออก`,
        receivedDate: "",
        productionDate: "",
        requestedBy: "",
        reason: "ยกเลิกใบสั่งซื้อที่รับผิด",
        location: po.receivedLocation || "",
        fromLocation: "",
        toLocation: "",
        lotSummary: `-${formatNum(take)} ที่ ${locationLabel(po.receivedLocation || "")}`,
        ts: Date.now(),
      });
    });

    persistItems(workingItems);
    if (newTxs.length) persistTransactions([...newTxs, ...transactions]);
    persistPurchaseOrders(purchaseOrders.map((p) => (p.id === po.id ? { ...p, status: "cancelled" } : p)));
    setAlertMessage(`ยกเลิกใบสั่งซื้อ ${po.poNumber} และตัดสต็อกที่รับผิดออกจากระบบแล้ว`);
  }

  function cancelPo(poId) {
    persistPurchaseOrders(purchaseOrders.map((p) => (p.id === poId ? { ...p, status: "cancelled" } : p)));
  }

  // ---- accounting module ----
  function poTotal(po) {
    return po.lines.reduce((s, l) => s + l.qty * l.unitPrice, 0);
  }
  function poPaidTotal(po) {
    return (po.payments || []).reduce((s, p) => s + Number(p.amount || 0), 0);
  }
  function poBalanceDue(po) {
    return poTotal(po) - poPaidTotal(po);
  }
  function poPaymentStatus(po) {
    const paid = poPaidTotal(po);
    const total = poTotal(po);
    if (total <= 0) return "paid";
    if (paid <= 0) return "unpaid";
    if (paid >= total) return "paid";
    return "partial";
  }

  // Auto-derive a balanced journal entry for a warehouse transaction (perpetual-inventory style,
  // using the item's current unit price as cost). Internal transfers (to/from production) don't
  // touch the inventory asset total, so they intentionally produce no entry.
  function autoJournalForTransaction(tx) {
    const item = items.find((i) => i.id === tx.itemId);
    const unitCost = item ? Number(item.price || 0) : 0;
    const value = Math.round(Math.abs(tx.amount) * unitCost * 100) / 100;
    if (!value) return null;

    if (tx.type === "waste") {
      return [
        { accountId: "acc-writeoff", debit: value, credit: 0 },
        { accountId: "acc-inventory", debit: 0, credit: value },
      ];
    }
    if (tx.type === "out" || tx.type === "consume_prod") {
      return [
        { accountId: "acc-cogs", debit: value, credit: 0 },
        { accountId: "acc-inventory", debit: 0, credit: value },
      ];
    }
    if (tx.type === "adjust") {
      return tx.amount < 0
        ? [
            { accountId: "acc-writeoff", debit: value, credit: 0 },
            { accountId: "acc-inventory", debit: 0, credit: value },
          ]
        : [
            { accountId: "acc-inventory", debit: value, credit: 0 },
            { accountId: "acc-writeoff", debit: 0, credit: value },
          ];
    }
    if (tx.type === "in") {
      const note = tx.note || "";
      if (note.includes("รับเข้าจากใบสั่งซื้อ")) {
        return [
          { accountId: "acc-inventory", debit: value, credit: 0 },
          { accountId: "acc-ap", debit: 0, credit: value },
        ];
      }
      if (note.includes("รับเข้าจากใบตรวจรับสินค้าสำเร็จรูป") || note.includes("นำเข้าจากไฟล์ Excel")) {
        return null; // cost already recognized via material consumption, or this is a data migration
      }
      return [
        { accountId: "acc-inventory", debit: value, credit: 0 },
        { accountId: "acc-cash", debit: 0, credit: value },
      ];
    }
    return null; // to_prod / from_prod: internal transfer only
  }

  // Watches for new warehouse transactions and posts the matching journal entry automatically,
  // skipping any transaction that's already been journaled (tracked via sourceRef = tx.id).
  useEffect(() => {
    if (!loaded) return;
    const alreadyJournaled = new Set(journalEntries.map((j) => j.sourceRef).filter(Boolean));
    const newEntries = [];
    transactions.forEach((tx) => {
      if (alreadyJournaled.has(tx.id)) return;
      const lines = autoJournalForTransaction(tx);
      if (!lines) return;
      newEntries.push({
        id: uid(),
        date: new Date(tx.ts).toISOString().slice(0, 10),
        description: `${movementTypeLabel(tx.type)}: ${tx.itemName} (${tx.sku})`,
        lines,
        auto: true,
        sourceRef: tx.id,
        ts: tx.ts,
      });
    });
    if (newEntries.length) persistJournal([...newEntries, ...journalEntries]);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [transactions, items, loaded]);

  function recordPoPayment(po, form) {
    const amount = Number(form.amount);
    if (!amount || amount <= 0) {
      setAlertMessage("กรุณากรอกจำนวนเงินที่จ่ายให้มากกว่า 0");
      return;
    }
    const balance = poBalanceDue(po);
    if (amount > balance + 0.01) {
      setAlertMessage(`ยอดค้างชำระของ PO นี้เหลือเพียง ${formatBaht(balance)} ไม่สามารถบันทึกจ่ายเกินยอดนี้ได้`);
      return;
    }
    const payment = {
      id: uid(),
      date: form.date || todayStr(),
      amount,
      method: form.method,
      reference: form.reference.trim(),
      note: form.note.trim(),
      ts: Date.now(),
    };
    persistPurchaseOrders(purchaseOrders.map((p) => (p.id === po.id ? { ...p, payments: [...(p.payments || []), payment] } : p)));
    persistJournal([
      {
        id: uid(),
        date: payment.date,
        description: `จ่ายชำระ PO ${po.poNumber} (${po.supplier})${payment.reference ? " · อ้างอิง " + payment.reference : ""}`,
        lines: [
          { accountId: "acc-ap", debit: amount, credit: 0 },
          { accountId: "acc-cash", debit: 0, credit: amount },
        ],
        auto: false,
        sourceRef: `po-payment-${payment.id}`,
        ts: Date.now(),
      },
      ...journalEntries,
    ]);
    setPayingPo(null);
    setPoPaymentForm(null);
  }

  function openJournalForm() {
    setJournalForm({
      date: todayStr(),
      description: "",
      lines: [
        { id: uid(), accountId: accounts[0]?.id || "", debit: "", credit: "" },
        { id: uid(), accountId: accounts[1]?.id || "", debit: "", credit: "" },
      ],
    });
    setShowJournalForm(true);
  }
  function addJournalLine() {
    setJournalForm((f) => ({ ...f, lines: [...f.lines, { id: uid(), accountId: "", debit: "", credit: "" }] }));
  }
  function removeJournalLine(lineId) {
    setJournalForm((f) => ({ ...f, lines: f.lines.filter((l) => l.id !== lineId) }));
  }
  function updateJournalLine(lineId, patch) {
    setJournalForm((f) => ({ ...f, lines: f.lines.map((l) => (l.id === lineId ? { ...l, ...patch } : l)) }));
  }
  function submitJournalForm() {
    if (!journalForm.description.trim()) {
      setAlertMessage("กรุณากรอกคำอธิบายรายการ");
      return;
    }
    const validLines = journalForm.lines.filter((l) => l.accountId && (Number(l.debit) > 0 || Number(l.credit) > 0));
    const totalDebit = validLines.reduce((s, l) => s + (Number(l.debit) || 0), 0);
    const totalCredit = validLines.reduce((s, l) => s + (Number(l.credit) || 0), 0);
    if (validLines.length < 2) {
      setAlertMessage("ต้องมีอย่างน้อย 2 บัญชีในรายการนี้");
      return;
    }
    if (Math.round(totalDebit * 100) !== Math.round(totalCredit * 100) || totalDebit === 0) {
      setAlertMessage(`ยอดเดบิต (${formatBaht(totalDebit)}) ต้องเท่ากับยอดเครดิต (${formatBaht(totalCredit)}) และต้องมากกว่า 0`);
      return;
    }
    persistJournal([
      {
        id: uid(),
        date: journalForm.date || todayStr(),
        description: journalForm.description.trim(),
        lines: validLines.map((l) => ({ accountId: l.accountId, debit: Number(l.debit) || 0, credit: Number(l.credit) || 0 })),
        auto: false,
        sourceRef: "",
        ts: Date.now(),
      },
      ...journalEntries,
    ]);
    setShowJournalForm(false);
    setJournalForm(null);
  }
  function deleteJournalEntry(entryId) {
    persistJournal(journalEntries.filter((j) => j.id !== entryId));
  }

  function openAccountForm(account) {
    setAccountForm(account ? { ...account } : { id: "", code: "", name: "", type: "asset" });
    setShowAccountForm(true);
  }
  function submitAccountForm() {
    const f = accountForm;
    if (!f.code.trim() || !f.name.trim()) {
      setAlertMessage("กรุณากรอกรหัสบัญชีและชื่อบัญชี");
      return;
    }
    if (f.id) {
      persistAccounts(accounts.map((a) => (a.id === f.id ? { ...a, code: f.code.trim(), name: f.name.trim(), type: f.type } : a)));
    } else {
      if (accounts.some((a) => a.code === f.code.trim())) {
        setAlertMessage(`รหัสบัญชี ${f.code} มีอยู่แล้ว`);
        return;
      }
      persistAccounts([...accounts, { id: uid(), code: f.code.trim(), name: f.name.trim(), type: f.type }]);
    }
    setShowAccountForm(false);
    setAccountForm(null);
  }
  function deleteAccount(accountId) {
    const inUse = journalEntries.some((j) => j.lines.some((l) => l.accountId === accountId));
    if (inUse) {
      setAlertMessage("ไม่สามารถลบบัญชีนี้ได้ เนื่องจากมีรายการบัญชีอ้างอิงอยู่แล้ว");
      return;
    }
    persistAccounts(accounts.filter((a) => a.id !== accountId));
  }

  // ---- full system backup / restore (JSON) ----
  function backupData() {
    const payload = {
      _type: "warehouse-app-backup",
      _version: 2,
      exportedAt: new Date().toISOString(),
      items,
      transactions,
      jobs,
      fgChecks,
      users,
      purchaseOrders,
      settings,
      accounts,
      journalEntries,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `warehouse-backup-${todayStr()}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    persistSettings({ ...settings, lastBackupTs: Date.now() });
  }

  function restoreData(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        if (data._type !== "warehouse-app-backup") {
          setAlertMessage("ไฟล์นี้ไม่ใช่ไฟล์สำรองข้อมูลของระบบนี้");
          return;
        }
        persistItems((data.items || []).map(migrateItem));
        persistTransactions(data.transactions || []);
        persistJobs(data.jobs || []);
        persistFgChecks(data.fgChecks || []);
        persistUsers(data.users || []);
        persistPurchaseOrders(data.purchaseOrders || []);
        persistSettings({ ...DEFAULT_SETTINGS, ...(data.settings || {}) });
        persistAccounts(data.accounts && data.accounts.length ? data.accounts : DEFAULT_ACCOUNTS);
        persistJournal(data.journalEntries || []);
        setAlertMessage("กู้คืนข้อมูลสำเร็จ ข้อมูลเดิมทั้งหมดถูกแทนที่ด้วยไฟล์สำรองนี้แล้ว");
      } catch (err) {
        setAlertMessage("ไม่สามารถอ่านไฟล์สำรองนี้ได้ กรุณาตรวจสอบว่าไฟล์ไม่เสียหาย");
      }
    };
    reader.readAsText(file);
  }

  // ---- export to Excel ----
  function exportExcel() {
    const invRows = [];
    items.forEach((it) => {
      const lots = (it.lots || []).filter((l) => l.qty > 0);
      if (lots.length === 0) {
        invRows.push({
          "ชื่อสินค้า": it.name,
          SKU: it.sku,
          หมวดหมู่: it.category || CATEGORIES[0],
          ผู้ขาย: it.supplier || "",
          สถานที่: "",
          "วันที่ผลิต": "",
          "วันหมดอายุ": "",
          "วันที่รับเข้า": "",
          จำนวนคงเหลือ: 0,
          หน่วย: it.unit,
          "ราคาต่อหน่วย": Number(it.price || 0),
          มูลค่า: 0,
        });
      } else {
        lots.forEach((l) => {
          invRows.push({
            "ชื่อสินค้า": it.name,
            SKU: it.sku,
            หมวดหมู่: it.category || CATEGORIES[0],
            ผู้ขาย: it.supplier || "",
            สถานที่: locationLabel(l.location),
            "วันที่ผลิต": formatDateOnly(l.productionDate),
            "วันหมดอายุ": formatDateOnly(l.expiryDate),
            "วันที่รับเข้า": formatDateOnly(l.receivedDate),
            จำนวนคงเหลือ: Number(l.qty || 0),
            หน่วย: it.unit,
            "ราคาต่อหน่วย": Number(it.price || 0),
            มูลค่า: Number(l.qty || 0) * Number(it.price || 0),
          });
        });
      }
    });

    const historyRows = transactions.map((tx) => ({
      "วันที่เวลา": formatDateTime(tx.ts),
      ประเภทรายการ: movementTypeLabel(tx.type),
      "ชื่อสินค้า": tx.itemName,
      SKU: tx.sku,
      จำนวน: tx.amount,
      จาก: tx.fromLocation ? locationLabel(tx.fromLocation) : "",
      ไปยัง: tx.toLocation ? locationLabel(tx.toLocation) : "",
      ผู้ทำรายการ: tx.requestedBy || "",
      เหตุผล: tx.reason || "",
      รายละเอียดล็อต: tx.lotSummary || "",
      หมายเหตุ: tx.note || "",
    }));

    const jobRows = [];
    jobs.forEach((job) => {
      job.lines.forEach((line) => {
        jobRows.push({
          "เลขที่ใบตัด": job.id.slice(-6).toUpperCase(),
          "วันที่ผลิต": formatDateOnly(job.jobDate),
          "ผลิตสินค้า": job.jobName,
          "วัตถุดิบที่ตัด": line.itemName,
          SKU: line.sku,
          จำนวนที่ตัด: line.qty,
          หน่วย: line.unit,
          คงเหลือที่ฝ่ายผลิตหลังตัด: line.remaining,
        });
      });
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(invRows), "สต็อกคงเหลือ");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(historyRows), "ประวัติการเคลื่อนไหว");
    if (jobRows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(jobRows), "ใบตัดวัตถุดิบ");

    if (fgChecks.length) {
      const fgRows = fgChecks.map((f) => ({
        "วันที่ตรวจ": formatDateOnly(f.checkDate),
        "สินค้า": f.itemName,
        SKU: f.sku,
        จำนวน: f.qty,
        หน่วย: f.unit,
        แผนกผลิต: f.dept ? locationLabel(f.dept) : "",
        ผู้ตรวจสอบ: f.inspector,
        "ลักษณะ/สี/กลิ่น": f.checkAppearance ? "ผ่าน" : "ไม่ผ่าน",
        "น้ำหนัก/ปริมาณ": f.checkWeight ? "ผ่าน" : "ไม่ผ่าน",
        บรรจุภัณฑ์: f.checkPackaging ? "ผ่าน" : "ไม่ผ่าน",
        ผลรวม: f.result === "pass" ? "ผ่าน" : "ไม่ผ่าน",
        รับเข้าสต็อก: f.stockedIn ? "ใช่" : "ไม่",
        หมายเหตุ: f.note || "",
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(fgRows), "ใบตรวจรับสินค้าสำเร็จรูป");
    }

    if (purchaseOrders.length) {
      const poRows = [];
      purchaseOrders.forEach((po) => {
        po.lines.forEach((l) => {
          poRows.push({
            "เลขที่ PO": po.poNumber,
            ซัพพลายเออร์: po.supplier,
            "ที่อยู่ผู้ขาย": po.supplierAddress || "",
            "เลขที่ใบกำกับภาษี": po.taxInvoiceNo || "",
            "วันที่สั่งซื้อ": formatDateOnly(po.orderDate),
            "วันที่คาดว่าจะได้รับ": formatDateOnly(po.expectedDate),
            สถานะ: po.status === "pending" ? "รอรับของ" : po.status === "received" ? "รับของแล้ว" : "ยกเลิก",
            สินค้า: l.itemName,
            SKU: l.sku,
            "วันที่ผลิต": formatDateOnly(l.productionDate),
            "วันหมดอายุ": formatDateOnly(l.expiryDate),
            จำนวนสั่ง: l.qty,
            หน่วย: l.unit,
            ราคาต่อหน่วย: l.unitPrice,
            มูลค่ารวม: l.qty * l.unitPrice,
          });
        });
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(poRows), "ใบสั่งซื้อ");
    }

    if (journalEntries.length) {
      const journalRows = [];
      journalEntries
        .slice()
        .sort((a, b) => a.date.localeCompare(b.date) || a.ts - b.ts)
        .forEach((j) => {
          j.lines.forEach((l) => {
            journalRows.push({
              "วันที่": formatDateOnly(j.date),
              คำอธิบาย: j.description,
              บัญชี: accounts.find((a) => a.id === l.accountId)?.name || "",
              รหัสบัญชี: accounts.find((a) => a.id === l.accountId)?.code || "",
              เดบิต: l.debit || 0,
              เครดิต: l.credit || 0,
              ที่มา: j.auto ? "อัตโนมัติ" : "บันทึกเอง",
            });
          });
        });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(journalRows), "บันทึกรายวัน");

      const trialRows = trialBalance
        .filter((a) => a.debit || a.credit)
        .map((a) => ({
          รหัส: a.code,
          บัญชี: a.name,
          ประเภท: accountTypeLabel(a.type),
          เดบิตรวม: a.debit,
          เครดิตรวม: a.credit,
          ยอดคงเหลือ: a.balance,
        }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(trialRows), "งบทดลอง");
    }

    XLSX.writeFile(wb, `คลังสินค้า-${formatDateOnly(todayStr()).replace(/\//g, "-")}.xlsx`);
  }

  // ---- import stock from Excel (matches the "สต็อกคงเหลือ" export format by SKU) ----
  function importExcel(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheetName = wb.SheetNames.find((n) => n.includes("สต็อก") || n.includes("stock")) || wb.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: "" });

        let workingItems = items;
        const newTxs = [];
        let createdCount = 0;
        let updatedCount = 0;
        let skipped = 0;

        rows.forEach((row) => {
          const name = String(row["ชื่อสินค้า"] || "").trim();
          const sku = String(row["SKU"] || "").trim();
          const qty = Number(row["จำนวนคงเหลือ"] || 0);
          if (!name || !sku || !qty || qty <= 0) {
            skipped++;
            return;
          }

          const location = parseLocationLabel(row["สถานที่"]) || LOCATIONS[0];
          const productionDate = parseDateOnly(row["วันที่ผลิต"]);
          const expiryDate = parseDateOnly(row["วันหมดอายุ"]);
          const receivedDate = parseDateOnly(row["วันที่รับเข้า"]) || todayStr();
          const supplier = String(row["ผู้ขาย"] || "").trim();
          const unit = String(row["หน่วย"] || "").trim() || UNIT_OPTIONS[0];
          const price = Number(row["ราคาต่อหน่วย"] || 0);

          let item = workingItems.find((i) => i.sku.toLowerCase() === sku.toLowerCase());
          if (!item) {
            item = { id: uid(), name, sku, unit, lowStock: 5, price, supplier, lots: [] };
            workingItems = [item, ...workingItems];
            createdCount++;
          } else {
            updatedCount++;
          }

          const matchIdx = (item.lots || []).findIndex(
            (l) => l.location === location && (l.productionDate || "") === (productionDate || "")
          );
          const newLots =
            matchIdx >= 0
              ? item.lots.map((l, idx) =>
                  idx === matchIdx ? { ...l, qty: Number(l.qty || 0) + qty, expiryDate: expiryDate || l.expiryDate } : l
                )
              : [...(item.lots || []), { id: uid(), location, productionDate, expiryDate, qty, receivedDate }];

          workingItems = workingItems.map((i) => (i.id === item.id ? { ...i, lots: newLots } : i));

          newTxs.push({
            id: uid(),
            itemId: item.id,
            itemName: name,
            sku,
            type: "in",
            amount: qty,
            note: "นำเข้าจากไฟล์ Excel",
            receivedDate,
            productionDate,
            requestedBy: "",
            location,
            fromLocation: "",
            toLocation: location,
            lotSummary: "",
            ts: Date.now(),
          });
        });

        persistItems(workingItems);
        if (newTxs.length) persistTransactions([...newTxs, ...transactions]);

        setAlertMessage(
          `นำเข้าสำเร็จ: สร้างสินค้าใหม่ ${createdCount} รายการ, เพิ่มสต็อกสินค้าเดิม ${updatedCount} รายการ` +
            (skipped ? `, ข้าม ${skipped} แถวที่ข้อมูลไม่ครบ` : "")
        );
      } catch (err) {
        setAlertMessage("ไม่สามารถอ่านไฟล์นี้ได้ กรุณาตรวจสอบว่าเป็นไฟล์ Excel ที่ถูกต้อง");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ---- generate a printable stock-count sheet (ใบนับสต็อก) for physical stocktakes ----
  function generateCountSheet() {
    const rows = [];
    let idx = 1;
    items
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name, "th"))
      .forEach((it) => {
        const locs = locationBreakdown(it);
        const pushRow = (loc, qty) =>
          rows.push({
            ลำดับ: idx++,
            "ชื่อสินค้า": it.name,
            SKU: it.sku,
            สถานที่: loc ? locationLabel(loc) : "",
            หน่วย: it.unit,
            ยอดตามระบบ: qty,
            ยอดนับจริง: "",
            ผลต่าง: "",
            ผู้นับ: "",
            หมายเหตุ: "",
          });
        if (locs.length === 0) pushRow("", 0);
        else locs.forEach(([loc, qty]) => pushRow(loc, qty));
      });

    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [
      { wch: 6 }, { wch: 26 }, { wch: 14 }, { wch: 10 }, { wch: 8 },
      { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 18 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ใบนับสต็อก");
    XLSX.writeFile(wb, `ใบนับสต็อก-${formatDateOnly(todayStr()).replace(/\//g, "-")}.xlsx`);
  }

  // ---- material-issue slip (ใบตัดวัตถุดิบ): permanently consumes raw material held at a production department ----
  function openJobForm(presetDept) {
    setJobForm({
      jobName: "",
      recipe: "",
      dept: presetDept || PROD_LOCATIONS[0],
      jobDate: todayStr(),
      lines: [{ id: uid(), itemId: "", qty: "" }],
    });
    setShowJobForm(true);
  }

  function addJobLine() {
    setJobForm((f) => ({ ...f, lines: [...f.lines, { id: uid(), itemId: "", qty: "" }] }));
  }

  function removeJobLine(lineId) {
    setJobForm((f) => ({ ...f, lines: f.lines.filter((l) => l.id !== lineId) }));
  }

  function updateJobLine(lineId, patch) {
    setJobForm((f) => ({ ...f, lines: f.lines.map((l) => (l.id === lineId ? { ...l, ...patch } : l)) }));
  }

  function submitJobForm() {
    const jobName = jobForm.jobName.trim();
    const recipe = jobForm.recipe.trim();
    const dept = jobForm.dept;
    const validLines = jobForm.lines.filter((l) => l.itemId && Number(l.qty) > 0);
    if (!jobName) {
      setAlertMessage("กรุณากรอกชื่อสินค้า/งานผลิตก่อนบันทึก");
      return;
    }
    if (validLines.length === 0) {
      setAlertMessage("กรุณาเลือกวัตถุดิบและกรอกจำนวนอย่างน้อย 1 รายการ");
      return;
    }

    // Validate every line against current stock held at the chosen department before committing anything.
    for (const line of validLines) {
      const item = items.find((i) => i.id === line.itemId);
      const available = item ? qtyAtLocation(item, dept) : 0;
      if (!item || available <= 0) {
        setAlertMessage(`${item ? item.name : "สินค้า"} ไม่มีสต็อกค้างอยู่ที่แผนกผลิต ${locationLabel(dept)} ให้ตัด`);
        return;
      }
      if (Number(line.qty) > available) {
        setAlertMessage(
          `${item.name} (${item.sku}) ที่แผนกผลิต ${locationLabel(dept)} มีคงเหลือเพียง ${formatNum(available)} ${item.unit} ไม่สามารถตัด ${formatNum(line.qty)} ${item.unit} ได้`
        );
        return;
      }
    }

    let workingItems = items;
    const receiptLines = [];
    const newTxs = [];
    const jobId = uid();
    const jobDate = jobForm.jobDate || todayStr();

    validLines.forEach((line) => {
      const qty = Number(line.qty);
      const item = workingItems.find((i) => i.id === line.itemId);
      let remaining = qty;
      const sourceLots = sortFEFO((item.lots || []).filter((l) => l.location === dept));
      let newLots = item.lots || [];
      sourceLots.forEach((target) => {
        if (remaining <= 0) return;
        const take = Math.min(target.qty, remaining);
        remaining -= take;
        if (take > 0) {
          newLots = newLots.map((l) => (l.id === target.id ? { ...l, qty: l.qty - take } : l));
        }
      });
      newLots = newLots.filter((l) => l.qty > 0);
      const remainingAtDept = newLots.filter((l) => l.location === dept).reduce((s, l) => s + l.qty, 0);
      const depleted = newLots.reduce((s, l) => s + l.qty, 0) <= 0;

      workingItems = depleted
        ? workingItems.filter((i) => i.id !== item.id)
        : workingItems.map((i) => (i.id === item.id ? { ...i, lots: newLots } : i));

      receiptLines.push({
        itemId: item.id,
        itemName: item.name,
        sku: item.sku,
        unit: item.unit,
        qty,
        remaining: remainingAtDept,
        deleted: depleted,
      });

      newTxs.push({
        id: uid(),
        itemId: item.id,
        itemName: item.name,
        sku: item.sku,
        type: "consume_prod",
        amount: qty,
        note: jobForm.jobNote || "",
        receivedDate: "",
        productionDate: "",
        requestedBy: "",
        location: "",
        fromLocation: dept,
        toLocation: "",
        lotSummary: `ตัดที่แผนก ${locationLabel(dept)} สำหรับผลิต ${jobName}${recipe ? ` (สูตร ${recipe})` : ""}`,
        jobId,
        jobName,
        recipe,
        dept,
        jobDate,
        ts: Date.now(),
      });
    });

    persistItems(workingItems);
    persistTransactions([...newTxs, ...transactions]);
    persistJobs([{ id: jobId, jobName, recipe, dept, jobDate, lines: receiptLines, ts: Date.now() }, ...jobs]);

    setShowJobForm(false);
    setJobForm(null);
    setJobReceipt({ jobName, recipe, dept, jobDate, lines: receiptLines });
  }

  // ---- finished goods QC check / receiving document (เอกสารเช็คสินค้าสำเร็จรูป) ----
  function openFgForm(presetJobId) {
    const presetJob = presetJobId ? jobs.find((j) => j.id === presetJobId) : null;
    setFgForm({
      itemId: "",
      newName: "",
      newSku: "",
      newUnit: UNIT_OPTIONS[0],
      qty: "",
      checkDate: todayStr(),
      expiryDate: "",
      inspector: "",
      dept: presetJob ? presetJob.dept : PROD_LOCATIONS[0],
      jobId: presetJobId || "",
      checkAppearance: true,
      checkWeight: true,
      checkPackaging: true,
      result: "pass",
      note: "",
    });
    setShowFgForm(true);
  }

  function submitFgForm() {
    const f = fgForm;
    const qty = Number(f.qty);
    const inspector = f.inspector.trim();
    const usingExisting = !!f.itemId;
    const name = usingExisting ? items.find((i) => i.id === f.itemId)?.name || "" : f.newName.trim();
    const sku = usingExisting ? items.find((i) => i.id === f.itemId)?.sku || "" : f.newSku.trim();

    if (!usingExisting && (!name || !sku)) {
      setAlertMessage("กรุณาเลือกสินค้าที่มีอยู่ หรือกรอกชื่อและ SKU สำหรับสินค้าใหม่ให้ครบ");
      return;
    }
    if (!qty || qty <= 0) {
      setAlertMessage("กรุณากรอกจำนวนที่ผลิตได้ให้มากกว่า 0");
      return;
    }
    if (!inspector) {
      setAlertMessage("กรุณากรอกชื่อผู้ตรวจสอบก่อนบันทึก");
      return;
    }

    let workingItems = items;
    let stockedIn = false;

    if (f.result === "pass") {
      let item = workingItems.find((i) => i.sku.toLowerCase() === sku.toLowerCase());
      if (!item) {
        item = { id: uid(), name, sku, unit: f.newUnit, category: CATEGORIES[0], lowStock: 5, price: 0, supplier: "", packUnit: "", packFactor: 0, lots: [] };
        workingItems = [item, ...workingItems];
      }
      const matchIdx = (item.lots || []).findIndex(
        (l) => l.location === "FG" && (l.productionDate || "") === (f.checkDate || "")
      );
      const newLots =
        matchIdx >= 0
          ? item.lots.map((l, idx) => (idx === matchIdx ? { ...l, qty: Number(l.qty || 0) + qty, expiryDate: f.expiryDate || l.expiryDate } : l))
          : [...(item.lots || []), { id: uid(), location: "FG", productionDate: f.checkDate, expiryDate: f.expiryDate, qty, receivedDate: f.checkDate }];
      workingItems = workingItems.map((i) => (i.id === item.id ? { ...i, lots: newLots } : i));
      stockedIn = true;

      persistItems(workingItems);
      persistTransactions([
        {
          id: uid(),
          itemId: item.id,
          itemName: name,
          sku,
          type: "in",
          amount: qty,
          note: `รับเข้าจากใบตรวจรับสินค้าสำเร็จรูป${f.note ? " · " + f.note : ""}`,
          receivedDate: f.checkDate,
          productionDate: f.checkDate,
          requestedBy: inspector,
          location: "FG",
          fromLocation: "",
          toLocation: "FG",
          lotSummary: "",
          ts: Date.now(),
        },
        ...transactions,
      ]);
    }

    const record = {
      id: uid(),
      itemName: name,
      sku,
      qty,
      unit: usingExisting ? items.find((i) => i.id === f.itemId)?.unit || f.newUnit : f.newUnit,
      checkDate: f.checkDate,
      expiryDate: f.expiryDate,
      inspector,
      dept: f.dept,
      jobId: f.jobId,
      checkAppearance: f.checkAppearance,
      checkWeight: f.checkWeight,
      checkPackaging: f.checkPackaging,
      result: f.result,
      note: f.note,
      stockedIn,
      ts: Date.now(),
    };
    persistFgChecks([record, ...fgChecks]);

    setShowFgForm(false);
    setFgForm(null);
    setFgReceipt(record);
  }

  // ---- stock movement ----
  function openMovement(item, type, presetDept) {
    setMovementItem(item);
    setMovementType(type);
    setMovementAmount("");
    setMovementNote("");
    setMovementReceivedDate(todayStr());
    setMovementProductionDate("");
    setMovementExpiryDate("");
    setMovementRequestedBy("");
    setMovementReason(WASTE_REASONS[0]);
    setMovementEntryUnit("base");
    if (type === "to_prod") {
      const sourceLocs = locationBreakdown(item).filter(([loc]) => loc && !isProdLocation(loc));
      setMovementLocation(sourceLocs.length ? sourceLocs[0][0] : LOCATIONS[0]);
      setMovementDept(presetDept || PROD_LOCATIONS[0]);
    } else if (type === "from_prod") {
      setMovementLocation(LOCATIONS[0]);
      const deptLocs = PROD_LOCATIONS.filter((d) => qtyAtLocation(item, d) > 0);
      setMovementDept(presetDept || (deptLocs.length ? deptLocs[0] : PROD_LOCATIONS[0]));
    } else {
      setMovementLocation(LOCATIONS[0]);
    }
  }

  function submitMovement() {
    const item = items.find((i) => i.id === movementItem.id);
    const usingPackUnit = movementType === "in" && movementEntryUnit === "pack" && item?.packUnit && Number(item.packFactor) > 0;
    const amount = usingPackUnit ? Number(movementAmount) * Number(item.packFactor) : Number(movementAmount);
    const note = movementNote.trim();
    const requestedBy = movementRequestedBy.trim();
    if (!amount || amount <= 0) return;
    if (movementType === "out" && !requestedBy) return;

    if (movementType === "out" && settings.approvalThreshold > 0 && amount >= settings.approvalThreshold && !isAdmin) {
      const available = warehouseQty(item);
      if (amount > available) {
        setAlertMessage(
          `${item.name} (${item.sku}) มีคงเหลือในคลังเพียง ${formatNum(available)} ${item.unit} ไม่สามารถเบิกออก ${formatNum(amount)} ${item.unit} ได้`
        );
        return;
      }
      persistApprovals([
        {
          id: uid(),
          itemId: item.id,
          itemName: item.name,
          sku: item.sku,
          unit: item.unit,
          amount,
          requestedBy,
          note,
          status: "pending",
          ts: Date.now(),
        },
        ...approvals,
      ]);
      setAlertMessage(
        `จำนวนที่เบิก (${formatNum(amount)} ${item.unit}) เกินเกณฑ์ที่ต้องขออนุมัติ (${formatNum(settings.approvalThreshold)} ${item.unit}) ส่งคำขอไปรออนุมัติจากผู้ดูแลระบบแล้ว`
      );
      setMovementItem(null);
      return;
    }

    let lotSummary = "";
    let itemDeleted = false;

    if (movementType === "out" || movementType === "waste") {
      const available = warehouseQty(item);
      if (available <= 0) {
        setAlertMessage(`${item.name} (${item.sku}) หมดสต็อกแล้ว ไม่สามารถ${movementType === "waste" ? "ตัดของเสีย/สูญหาย" : "เบิกออก"}ได้`);
        return;
      }
      if (amount > available) {
        setAlertMessage(
          `${item.name} (${item.sku}) มีคงเหลือในคลังเพียง ${formatNum(available)} ${item.unit} ไม่สามารถ${movementType === "waste" ? "ตัด" : "เบิกออก"} ${formatNum(amount)} ${item.unit} ได้`
        );
        return;
      }
    }

    if (movementType === "to_prod") {
      const available = qtyAtLocation(item, movementLocation);
      if (available <= 0) {
        setAlertMessage(`${item.name} (${item.sku}) ไม่มีสต็อกที่ ${locationLabel(movementLocation)} ให้โอนไปฝ่ายผลิต`);
        return;
      }
      if (amount > available) {
        setAlertMessage(
          `${item.name} (${item.sku}) ที่ ${locationLabel(movementLocation)} มีคงเหลือเพียง ${formatNum(available)} ${item.unit} ไม่สามารถโอน ${formatNum(amount)} ${item.unit} ได้`
        );
        return;
      }
    }

    if (movementType === "from_prod") {
      const available = qtyAtLocation(item, movementDept);
      if (available <= 0) {
        setAlertMessage(`${item.name} (${item.sku}) ไม่มีสต็อกค้างอยู่ที่แผนกผลิต ${locationLabel(movementDept)} ให้คืน`);
        return;
      }
      if (amount > available) {
        setAlertMessage(
          `${item.name} (${item.sku}) ที่แผนกผลิต ${locationLabel(movementDept)} มีคงเหลือเพียง ${formatNum(available)} ${item.unit} ไม่สามารถคืน ${formatNum(amount)} ${item.unit} ได้`
        );
        return;
      }
    }

    if (movementType === "in") {
      const prodDate = movementProductionDate;
      const existingLots = item.lots || [];
      // Same location + same production date -> add to that lot. Otherwise open a new, separate lot.
      const matchIdx = existingLots.findIndex(
        (l) => l.location === movementLocation && (l.productionDate || "") === (prodDate || "")
      );
      let newLots;
      if (matchIdx >= 0) {
        newLots = existingLots.map((l, idx) =>
          idx === matchIdx ? { ...l, qty: Number(l.qty || 0) + amount, expiryDate: movementExpiryDate || l.expiryDate } : l
        );
        lotSummary = `รวมเข้าล็อตเดิม (${locationLabel(movementLocation)})`;
      } else {
        newLots = [
          ...existingLots,
          {
            id: uid(),
            location: movementLocation,
            productionDate: prodDate,
            expiryDate: movementExpiryDate,
            qty: amount,
            receivedDate: movementReceivedDate || todayStr(),
          },
        ];
        lotSummary = existingLots.length ? `เปิดล็อตใหม่ (${locationLabel(movementLocation)})` : "";
      }
      persistItems(items.map((i) => (i.id === item.id ? { ...i, lots: newLots } : i)));
    } else if (movementType === "out" || movementType === "waste") {
      // Withdraw for use outside the warehouse (sale/consumption) or write off as waste/loss: FEFO across physical locations only.
      let remaining = amount;
      const sorted = sortFEFO((item.lots || []).filter((l) => !isProdLocation(l.location)));
      const consumed = [];
      let newLots = item.lots || [];
      sorted.forEach((target) => {
        if (remaining <= 0) return;
        const take = Math.min(target.qty, remaining);
        remaining -= take;
        if (take > 0) {
          consumed.push({ productionDate: target.productionDate, location: target.location, amount: take });
          newLots = newLots.map((l) => (l.id === target.id ? { ...l, qty: l.qty - take } : l));
        }
      });
      lotSummary = consumed
        .map((c) => `${locationLabel(c.location)}${c.productionDate ? " " + formatDateOnly(c.productionDate) : ""} -${formatNum(c.amount)}`)
        .join(", ");
      const remainingLots = newLots.filter((l) => l.qty > 0);
      const depleted = remainingLots.reduce((s, l) => s + l.qty, 0) <= 0;
      if (depleted) {
        itemDeleted = true;
        persistItems(items.filter((i) => i.id !== item.id));
      } else {
        persistItems(items.map((i) => (i.id === item.id ? { ...i, lots: remainingLots } : i)));
      }
    } else if (movementType === "to_prod") {
      // Transfer to production: FEFO within the chosen source location, moved into the chosen production department.
      let remaining = amount;
      const sourceLots = sortFEFO((item.lots || []).filter((l) => l.location === movementLocation));
      let newLots = item.lots || [];
      sourceLots.forEach((target) => {
        if (remaining <= 0) return;
        const take = Math.min(target.qty, remaining);
        remaining -= take;
        if (take > 0) {
          newLots = newLots.map((l) => (l.id === target.id ? { ...l, qty: l.qty - take } : l));
          newLots = mergeIntoLocation(newLots, movementDept, target, take, todayStr());
        }
      });
      newLots = newLots.filter((l) => l.qty > 0);
      lotSummary = `${locationLabel(movementLocation)} → แผนกผลิต ${locationLabel(movementDept)}`;
      persistItems(items.map((i) => (i.id === item.id ? { ...i, lots: newLots } : i)));
    } else if (movementType === "from_prod") {
      // Return from production: FEFO among lots currently held at the chosen department, moved back into the chosen location.
      let remaining = amount;
      const sourceLots = sortFEFO((item.lots || []).filter((l) => l.location === movementDept));
      let newLots = item.lots || [];
      sourceLots.forEach((target) => {
        if (remaining <= 0) return;
        const take = Math.min(target.qty, remaining);
        remaining -= take;
        if (take > 0) {
          newLots = newLots.map((l) => (l.id === target.id ? { ...l, qty: l.qty - take } : l));
          newLots = mergeIntoLocation(newLots, movementLocation, target, take, todayStr());
        }
      });
      newLots = newLots.filter((l) => l.qty > 0);
      lotSummary = `แผนกผลิต ${locationLabel(movementDept)} → ${locationLabel(movementLocation)}`;
      persistItems(items.map((i) => (i.id === item.id ? { ...i, lots: newLots } : i)));
    }

    persistTransactions([
      {
        id: uid(),
        itemId: item.id,
        itemName: item.name,
        sku: item.sku,
        type: movementType,
        amount,
        note,
        receivedDate: movementType === "in" ? movementReceivedDate : "",
        productionDate: movementType === "in" ? movementProductionDate : "",
        requestedBy: movementType === "out" || movementType === "to_prod" || movementType === "from_prod" ? requestedBy : "",
        reason: movementType === "waste" ? movementReason : "",
        location: movementType !== "out" && movementType !== "waste" ? movementLocation : "",
        fromLocation: movementType === "to_prod" ? movementLocation : movementType === "from_prod" ? movementDept : "",
        toLocation: movementType === "in" ? movementLocation : movementType === "to_prod" ? movementDept : movementType === "from_prod" ? movementLocation : "",
        lotSummary,
        ts: Date.now(),
      },
      ...transactions,
    ]);
    setMovementItem(null);
    if (itemDeleted) {
      setAlertMessage(`${item.name} (${item.sku}) หมดสต็อกและถูกลบออกจากรายการสินค้าแล้ว`);
    }
  }

  // ---- approval workflow: admin approves/rejects large withdrawal requests ----
  function approveWithdrawalRequest(approval) {
    const item = items.find((i) => i.id === approval.itemId);
    if (!item) {
      setAlertMessage("ไม่พบสินค้านี้ในระบบแล้ว (อาจถูกลบไปก่อนหน้านี้)");
      persistApprovals(approvals.map((a) => (a.id === approval.id ? { ...a, status: "rejected" } : a)));
      return;
    }
    const available = warehouseQty(item);
    if (approval.amount > available) {
      setAlertMessage(`ไม่สามารถอนุมัติได้ เนื่องจากคงเหลือในคลังตอนนี้มีเพียง ${formatNum(available)} ${item.unit}`);
      return;
    }
    let remaining = approval.amount;
    const sorted = sortFEFO((item.lots || []).filter((l) => !isProdLocation(l.location)));
    const consumed = [];
    let newLots = item.lots || [];
    sorted.forEach((target) => {
      if (remaining <= 0) return;
      const take = Math.min(target.qty, remaining);
      remaining -= take;
      if (take > 0) {
        consumed.push({ productionDate: target.productionDate, location: target.location, amount: take });
        newLots = newLots.map((l) => (l.id === target.id ? { ...l, qty: l.qty - take } : l));
      }
    });
    const lotSummary = consumed
      .map((c) => `${locationLabel(c.location)}${c.productionDate ? " " + formatDateOnly(c.productionDate) : ""} -${formatNum(c.amount)}`)
      .join(", ");
    const remainingLots = newLots.filter((l) => l.qty > 0);
    const depleted = remainingLots.reduce((s, l) => s + l.qty, 0) <= 0;
    if (depleted) {
      persistItems(items.filter((i) => i.id !== item.id));
    } else {
      persistItems(items.map((i) => (i.id === item.id ? { ...i, lots: remainingLots } : i)));
    }
    persistTransactions([
      {
        id: uid(),
        itemId: item.id,
        itemName: item.name,
        sku: item.sku,
        type: "out",
        amount: approval.amount,
        note: approval.note,
        receivedDate: "",
        productionDate: "",
        requestedBy: approval.requestedBy,
        reason: "",
        location: "",
        fromLocation: "",
        toLocation: "",
        lotSummary,
        ts: Date.now(),
      },
      ...transactions,
    ]);
    persistApprovals(
      approvals.map((a) =>
        a.id === approval.id ? { ...a, status: "approved", approvedBy: currentUser?.name || "", approvedTs: Date.now() } : a
      )
    );
    if (depleted) {
      setAlertMessage(`อนุมัติสำเร็จ — ${item.name} (${item.sku}) หมดสต็อกและถูกลบออกจากรายการสินค้าแล้ว`);
    }
  }

  function rejectWithdrawalRequest(approvalId) {
    persistApprovals(approvals.map((a) => (a.id === approvalId ? { ...a, status: "rejected" } : a)));
  }

  if (!loaded) {
    return (
      <div className="flex h-64 w-full items-center justify-center text-neutral-500">
        <div className="flex items-center gap-2">
          <Boxes className="animate-pulse" size={20} />
          <span>กำลังโหลดข้อมูลคลังสินค้า...</span>
        </div>
      </div>
    );
  }

  async function createUser(name, role, password) {
    const trimmed = name.trim();
    if (!trimmed) return;
    if (!password || password.length < 4) {
      setLoginError("กรุณาตั้งรหัสผ่านอย่างน้อย 4 ตัวอักษร");
      return;
    }
    const passwordHash = await hashPassword(password);
    const user = { id: uid(), name: trimmed, role, passwordHash };
    persistUsers([...users, user]);
    loginAs(user);
    setLoginName("");
    setLoginPassword("");
    setLoginError("");
  }

  async function attemptLogin(user, password) {
    if (!user.passwordHash) {
      // Grandfathered account created before password support existed.
      loginAs(user);
      return;
    }
    const hash = await hashPassword(password);
    if (hash === user.passwordHash) {
      loginAs(user);
      setAttemptingUser(null);
      setAttemptPassword("");
      setLoginError("");
    } else {
      setLoginError("รหัสผ่านไม่ถูกต้อง");
    }
  }

  async function resetUserPassword(userId, newPassword) {
    if (!newPassword || newPassword.length < 4) {
      setAlertMessage("กรุณาตั้งรหัสผ่านอย่างน้อย 4 ตัวอักษร");
      return;
    }
    const passwordHash = await hashPassword(newPassword);
    persistUsers(users.map((u) => (u.id === userId ? { ...u, passwordHash } : u)));
    setResettingPwUser(null);
    setResetPwValue("");
    setAlertMessage("ตั้งรหัสผ่านใหม่สำเร็จ");
  }

  if (!currentUser) {
    return (
      <div className="flex w-full items-center justify-center rounded-xl border border-neutral-800 bg-neutral-950 p-8 text-neutral-100">
        <div className="w-full max-w-sm">
          <div className="mb-5 flex items-center gap-2">
            <div className="flex h-9 w-9 items-center justify-center rounded-md bg-amber-500 text-neutral-950">
              <Package size={20} />
            </div>
            <h1 className="text-lg font-black uppercase tracking-wide text-neutral-50">ระบบผู้ช่วยคลังสินค้า</h1>
          </div>

          {attemptingUser ? (
            <div>
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-neutral-400">
                รหัสผ่านของ {attemptingUser.name}
              </p>
              <input
                type="password"
                autoFocus
                className={`${inputCls} mb-2`}
                placeholder="รหัสผ่าน"
                value={attemptPassword}
                onChange={(e) => setAttemptPassword(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") attemptLogin(attemptingUser, attemptPassword);
                }}
              />
              {loginError && <p className="mb-2 text-xs text-red-400">{loginError}</p>}
              <div className="flex gap-2">
                <button
                  onClick={() => {
                    setAttemptingUser(null);
                    setAttemptPassword("");
                    setLoginError("");
                  }}
                  className="flex-1 rounded-md border border-neutral-700 px-4 py-2 text-sm text-neutral-300 hover:bg-neutral-800"
                >
                  ยกเลิก
                </button>
                <button
                  onClick={() => attemptLogin(attemptingUser, attemptPassword)}
                  className="flex-1 rounded-md bg-amber-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-amber-400"
                >
                  เข้าสู่ระบบ
                </button>
              </div>
            </div>
          ) : (
            <>
              {users.length > 0 && (
                <>
                  <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-neutral-400">เลือกผู้ใช้งาน</p>
                  <div className="mb-4 space-y-1.5">
                    {users.map((u) => (
                      <button
                        key={u.id}
                        onClick={() => {
                          setLoginError("");
                          if (u.passwordHash) {
                            setAttemptingUser(u);
                          } else {
                            loginAs(u);
                          }
                        }}
                        className="flex w-full items-center justify-between rounded-md border border-neutral-800 bg-neutral-900 px-3 py-2 text-left hover:border-amber-500/50 hover:bg-neutral-800"
                      >
                        <span className="font-semibold text-neutral-100">{u.name}</span>
                        <span className="flex items-center gap-1 rounded bg-neutral-800 px-2 py-0.5 text-[10px] uppercase text-neutral-400">
                          {u.role === "admin" ? <Shield size={10} /> : u.role === "viewer" ? <Eye size={10} /> : <Users size={10} />}
                          {roleLabel(u.role)}
                        </span>
                      </button>
                    ))}
                  </div>
                </>
              )}

              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-neutral-400">
                {users.length === 0 ? "ตั้งค่าผู้ดูแลระบบคนแรก" : "เพิ่มผู้ใช้งานใหม่"}
              </p>
              <input
                className={`${inputCls} mb-2`}
                placeholder="ชื่อผู้ใช้งาน"
                value={loginName}
                onChange={(e) => setLoginName(e.target.value)}
              />
              <input
                type="password"
                className={`${inputCls} mb-2`}
                placeholder="ตั้งรหัสผ่าน (อย่างน้อย 4 ตัวอักษร)"
                value={loginPassword}
                onChange={(e) => setLoginPassword(e.target.value)}
              />
              {users.length > 0 && (
                <select className={`${inputCls} mb-2`} value={loginRole} onChange={(e) => setLoginRole(e.target.value)}>
                  {ROLES.map((r) => (
                    <option key={r.id} value={r.id}>{r.label}</option>
                  ))}
                </select>
              )}
              {loginError && <p className="mb-2 text-xs text-red-400">{loginError}</p>}
              <button
                onClick={() => createUser(loginName, users.length === 0 ? "admin" : loginRole, loginPassword)}
                disabled={!loginName.trim()}
                className="flex w-full items-center justify-center gap-1.5 rounded-md bg-amber-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-amber-400 disabled:cursor-not-allowed disabled:opacity-40"
              >
                <UserPlus size={15} /> {users.length === 0 ? "เริ่มใช้งาน" : "เพิ่มและเข้าสู่ระบบ"}
              </button>
              <p className="mt-3 text-[11px] text-neutral-600">
                รหัสผ่านนี้เป็นการกันคนคลิกชื่อคนอื่นเข้าใช้งานเฉยๆ ไม่ใช่ระบบยืนยันตัวตนระดับองค์กร และข้อมูลทั้งหมดเก็บอยู่ในเบราว์เซอร์เครื่องนี้เท่านั้น
              </p>
            </>
          )}
        </div>
      </div>
    );
  }

  return (
    <div
      style={{ transform: `scale(${displayScale})`, transformOrigin: "top left", width: `${100 / displayScale}%` }}
    >
    <div className="w-full rounded-xl border border-neutral-800 bg-neutral-950 p-5 text-neutral-100">
      <style>{`
        @media print {
          body * { visibility: hidden; }
          .printable-content, .printable-content * { visibility: visible; }
          .printable-content {
            position: fixed; inset: 0; padding: 24px;
            background: white; color: black;
          }
          .printable-content .no-print { display: none !important; }
          @page { size: ${settings.printPaperSize}; }
        }
      `}</style>
      {/* Header */}
      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
        <div>
          <div className="flex items-center gap-2">
            <div className="flex h-9 w-9 items-center justify-center rounded-md bg-amber-500 text-neutral-950">
              <Package size={20} />
            </div>
            <h1 className="text-xl font-black uppercase tracking-wide text-neutral-50">{t("appTitle")}</h1>
          </div>
          <p className="ml-11 text-xs text-neutral-500">{t("appSubtitle")}</p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <span className="flex items-center gap-1.5 rounded-md border border-neutral-700 bg-neutral-900 px-3 py-2 text-xs font-semibold text-neutral-300">
            {currentUser.role === "admin" ? <Shield size={13} className="text-amber-400" /> : currentUser.role === "viewer" ? <Eye size={13} className="text-neutral-400" /> : <Users size={13} className="text-sky-400" />}
            {currentUser.name} · {roleLabel(currentUser.role)}
          </span>
          {isAdmin && (
            <button
              onClick={() => setShowUserManager(true)}
              className="flex items-center gap-1.5 rounded-md border border-neutral-700 bg-neutral-900 px-2 py-2 text-xs font-semibold text-neutral-300 hover:bg-neutral-800"
              title="จัดการผู้ใช้งาน"
            >
              <Users size={14} />
            </button>
          )}
          <button
            onClick={() => setShowSettings(true)}
            className="flex items-center gap-1.5 rounded-md border border-neutral-700 bg-neutral-900 px-2 py-2 text-xs font-semibold text-neutral-300 hover:bg-neutral-800"
            title={t("settings")}
          >
            <Settings size={14} />
          </button>
          <button
            onClick={logout}
            className="flex items-center gap-1.5 rounded-md border border-neutral-700 bg-neutral-900 px-2 py-2 text-xs font-semibold text-neutral-300 hover:bg-neutral-800"
            title="สลับผู้ใช้งาน"
          >
            <LogOut size={14} />
          </button>
          <button
            onClick={() => setShowDashboard((s) => !s)}
            className="flex items-center gap-1.5 rounded-md border border-emerald-600/50 bg-emerald-500/10 px-3 py-2 text-xs font-semibold uppercase tracking-wide text-emerald-300 hover:bg-emerald-500/20"
          >
            <BarChart3 size={15} /> {t("dashboard")}
          </button>
          {isAccounting && (
            <button
              onClick={() => setShowAccounting((s) => !s)}
              className="flex items-center gap-1.5 rounded-md border border-yellow-600/50 bg-yellow-500/10 px-3 py-2 text-xs font-semibold uppercase tracking-wide text-yellow-300 hover:bg-yellow-500/20"
            >
              <Wallet size={15} /> ระบบบัญชี
            </button>
          )}
          <button
            onClick={() => setShowProdPanel((s) => !s)}
            className="flex items-center gap-1.5 rounded-md border border-indigo-600/50 bg-indigo-500/10 px-3 py-2 text-xs font-semibold uppercase tracking-wide text-indigo-300 hover:bg-indigo-500/20"
          >
            <Factory size={15} /> {t("prodPanel")}
          </button>
          <button
            onClick={() => setShowHistory((s) => !s)}
            className="flex items-center gap-1.5 rounded-md border border-neutral-700 bg-neutral-900 px-3 py-2 text-xs font-semibold uppercase tracking-wide text-neutral-300 hover:bg-neutral-800"
          >
            <History size={15} /> {t("history")}
          </button>
          <button
            onClick={() => setShowEditLog((s) => !s)}
            className="flex items-center gap-1.5 rounded-md border border-neutral-700 bg-neutral-900 px-3 py-2 text-xs font-semibold uppercase tracking-wide text-neutral-300 hover:bg-neutral-800"
          >
            <Pencil size={15} /> ประวัติแก้ไขสินค้า
          </button>
          {canEdit && (
            <button
              onClick={() => setShowApprovals((s) => !s)}
              className="relative flex items-center gap-1.5 rounded-md border border-rose-600/50 bg-rose-500/10 px-3 py-2 text-xs font-semibold uppercase tracking-wide text-rose-300 hover:bg-rose-500/20"
            >
              <Clock size={15} /> รออนุมัติ
              {approvals.filter((a) => a.status === "pending").length > 0 && (
                <span className="absolute -right-1.5 -top-1.5 flex h-4 min-w-4 items-center justify-center rounded-full bg-rose-500 px-1 text-[10px] font-bold text-neutral-950">
                  {approvals.filter((a) => a.status === "pending").length}
                </span>
              )}
            </button>
          )}
        </div>
      </div>

      <BarcodeDivider />

      {saveError && (
        <div className="mt-3 flex items-center gap-2 rounded-md border border-red-500/40 bg-red-500/10 px-3 py-2 text-xs text-red-300">
          <AlertTriangle size={14} /> {saveError}
        </div>
      )}

      {!bannerDismissed && (stats.lowStockCount > 0 || stats.expiringCount > 0) && (
        <div className="mt-3 flex items-center justify-between gap-2 rounded-md border border-amber-500/40 bg-amber-500/10 px-3 py-2 text-xs text-amber-200">
          <div className="flex items-center gap-2">
            <AlertTriangle size={14} className="shrink-0" />
            <span>
              {stats.lowStockCount > 0 && <>มี {stats.lowStockCount} รายการใกล้หมดสต็อก</>}
              {stats.lowStockCount > 0 && stats.expiringCount > 0 && " และ "}
              {stats.expiringCount > 0 && <>{stats.expiringCount} รายการใกล้/หมดอายุ</>}
              {" — ตรวจสอบก่อนของขาดหรือหมดอายุ"}
            </span>
          </div>
          <button onClick={() => setBannerDismissed(true)} className="shrink-0 rounded p-1 hover:bg-amber-500/20">
            <X size={14} />
          </button>
        </div>
      )}

      {isAdmin && !backupBannerDismissed && (Date.now() - (settings.lastBackupTs || 0) > 7 * 86400000) && (
        <div className="mt-3 flex items-center justify-between gap-2 rounded-md border border-sky-500/40 bg-sky-500/10 px-3 py-2 text-xs text-sky-200">
          <div className="flex items-center gap-2">
            <Download size={14} className="shrink-0" />
            <span>
              ข้อมูลทั้งหมดเก็บอยู่ในเบราว์เซอร์ของเครื่องนี้เท่านั้น (ไม่ sync ข้ามเครื่อง/เบราว์เซอร์)
              {settings.lastBackupTs
                ? ` — สำรองข้อมูลล่าสุดเมื่อ ${formatDateOnly(new Date(settings.lastBackupTs).toISOString().slice(0, 10))}`
                : " — ยังไม่เคยสำรองข้อมูลเลย"}{" "}
              แนะนำให้กด "สำรองข้อมูล" ในหน้าตั้งค่าเป็นประจำ
            </span>
          </div>
          <button onClick={() => setBackupBannerDismissed(true)} className="shrink-0 rounded p-1 hover:bg-sky-500/20">
            <X size={14} />
          </button>
        </div>
      )}

      {/* Stats */}
      <div className="mt-4 grid grid-cols-2 gap-3 sm:grid-cols-5">
        <StatCard icon={Boxes} label={t("statItems")} value={formatNum(stats.totalItems)} accent="bg-neutral-800 text-amber-400" />
        <StatCard icon={Package} label={t("statUnits")} value={formatNum(stats.totalUnits)} accent="bg-neutral-800 text-sky-400" />
        <StatCard icon={AlertTriangle} label={t("statLow")} value={formatNum(stats.lowStockCount)} accent="bg-neutral-800 text-red-400" />
        <StatCard icon={AlertTriangle} label={t("statExpiring")} value={formatNum(stats.expiringCount)} accent="bg-neutral-800 text-amber-400" />
        <StatCard icon={Wallet} label={t("statValue")} value={formatBaht(stats.totalValue)} accent="bg-neutral-800 text-emerald-400" />
      </div>

      {/* Analytics dashboard */}
      {showDashboard && (
        <>
        <div className="mt-4 grid gap-3 rounded-lg border border-emerald-800/40 bg-emerald-950/10 p-3 sm:grid-cols-2">
          <div className="rounded-md border border-neutral-800 bg-neutral-950 p-3">
            <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-neutral-400">มูลค่าสต็อกตามหมวดหมู่</p>
            {categoryValueData.length === 0 ? (
              <p className="py-8 text-center text-xs text-neutral-500">ยังไม่มีข้อมูลเพียงพอ</p>
            ) : (
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={categoryValueData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#27272a" />
                  <XAxis dataKey="name" tick={{ fill: "#a1a1aa", fontSize: 11 }} />
                  <YAxis tick={{ fill: "#a1a1aa", fontSize: 11 }} />
                  <Tooltip
                    contentStyle={{ background: "#171717", border: "1px solid #404040", fontSize: 12 }}
                    formatter={(v) => formatBaht(v)}
                  />
                  <Bar dataKey="value" radius={[4, 4, 0, 0]}>
                    {categoryValueData.map((_, idx) => (
                      <Cell key={idx} fill={CHART_COLORS[idx % CHART_COLORS.length]} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            )}
          </div>

          <div className="rounded-md border border-neutral-800 bg-neutral-950 p-3">
            <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-neutral-400">
              จำนวนรายการเคลื่อนไหว 7 วันล่าสุด
            </p>
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={dailyMovementData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#27272a" />
                <XAxis dataKey="date" tick={{ fill: "#a1a1aa", fontSize: 11 }} />
                <YAxis tick={{ fill: "#a1a1aa", fontSize: 11 }} allowDecimals={false} />
                <Tooltip contentStyle={{ background: "#171717", border: "1px solid #404040", fontSize: 12 }} />
                <Legend wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="รับเข้า" fill="#34d399" radius={[3, 3, 0, 0]} />
                <Bar dataKey="เบิกออก" fill="#f87171" radius={[3, 3, 0, 0]} />
                <Bar dataKey="ของเสีย" fill="#fb7185" radius={[3, 3, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="mt-3 rounded-md border border-neutral-800 bg-neutral-950 p-3">
          <div className="mb-2 flex items-center justify-between gap-2">
            <p className="flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-neutral-400">
              <TrendingDown size={13} /> สินค้าไม่เคลื่อนไหว (Dead Stock)
            </p>
            <div className="flex items-center gap-1.5 text-xs text-neutral-500">
              ไม่เคลื่อนไหวเกิน
              <input
                type="number"
                min="1"
                value={deadStockDays}
                onChange={(e) => setDeadStockDays(Number(e.target.value) || 30)}
                className="w-16 rounded border border-neutral-700 bg-neutral-800 px-2 py-1 text-center text-xs text-neutral-100"
              />
              วัน
            </div>
          </div>
          {deadStockItems.length === 0 ? (
            <p className="py-4 text-center text-xs text-neutral-500">ไม่มีสินค้าที่เข้าเงื่อนไขนี้ — สต็อกหมุนเวียนดี</p>
          ) : (
            <div className="max-h-56 overflow-y-auto">
              <table className="w-full text-xs">
                <thead className="text-neutral-500">
                  <tr>
                    <th className="px-2 py-1 text-left">สินค้า</th>
                    <th className="px-2 py-1 text-right">คงเหลือ</th>
                    <th className="px-2 py-1 text-right">มูลค่า</th>
                    <th className="px-2 py-1 text-right">เคลื่อนไหวล่าสุด</th>
                  </tr>
                </thead>
                <tbody>
                  {deadStockItems.map(({ item, qty, lastTs }) => (
                    <tr key={item.id} className="border-t border-neutral-800">
                      <td className="px-2 py-1.5">
                        <p className="font-semibold text-neutral-200">{item.name}</p>
                        <p className="font-mono text-neutral-500">{item.sku}</p>
                      </td>
                      <td className="px-2 py-1.5 text-right font-mono text-neutral-100">{formatNum(qty)} {item.unit}</td>
                      <td className="px-2 py-1.5 text-right font-mono text-neutral-300">{formatBaht(qty * Number(item.price || 0))}</td>
                      <td className="px-2 py-1.5 text-right text-neutral-500">
                        {lastTs ? formatDateTime(lastTs) : "ไม่เคยมีการเคลื่อนไหว"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </>
      )}

      {/* Production department (SSN/SD) panel */}
      {showProdPanel && (
        <div className="mt-4 grid gap-3 rounded-lg border border-indigo-800/40 bg-indigo-950/10 p-3 sm:grid-cols-2">
          {PROD_LOCATIONS.map((dept) => {
            const stockHere = items
              .map((it) => ({ it, qty: qtyAtLocation(it, dept) }))
              .filter((x) => x.qty > 0)
              .sort((a, b) => a.it.name.localeCompare(b.it.name, "th"));
            const recentJobs = jobs.filter((j) => j.dept === dept).slice(0, 5);
            return (
              <div key={dept} className="rounded-md border border-indigo-700/40 bg-neutral-950 p-3">
                <div className="mb-2 flex items-center justify-between">
                  <p className="flex items-center gap-1.5 font-bold uppercase tracking-wide text-indigo-300">
                    <Factory size={16} /> แผนกผลิต {dept}
                  </p>
                  {canEdit && (
                    <button
                      onClick={() => openJobForm(dept)}
                      className="flex items-center gap-1 rounded-md border border-indigo-600/50 bg-indigo-500/10 px-2 py-1 text-xs font-semibold text-indigo-300 hover:bg-indigo-500/20"
                    >
                      <Scissors size={12} /> ตัดวัตถุดิบที่นี่
                    </button>
                  )}
                </div>

                <p className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-neutral-500">
                  วัตถุดิบที่ค้างอยู่ที่แผนกนี้
                </p>
                {stockHere.length === 0 ? (
                  <p className="mb-3 text-xs text-neutral-500">ไม่มีวัตถุดิบค้างอยู่ที่แผนกนี้</p>
                ) : (
                  <ul className="mb-3 max-h-40 space-y-1 overflow-y-auto">
                    {stockHere.map(({ it, qty }) => (
                      <li key={it.id} className="flex items-center justify-between rounded border border-neutral-800 bg-neutral-900 px-2 py-1 text-xs">
                        <span className="truncate text-neutral-200">{it.name} <span className="font-mono text-neutral-500">({it.sku})</span></span>
                        <span className="shrink-0 font-mono font-semibold text-neutral-100">{formatNum(qty)} {it.unit}</span>
                      </li>
                    ))}
                  </ul>
                )}

                <p className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-neutral-500">
                  ใบตัดวัตถุดิบล่าสุด (สูตรที่เคยใช้)
                </p>
                {recentJobs.length === 0 ? (
                  <p className="text-xs text-neutral-500">ยังไม่มีการตัดวัตถุดิบที่แผนกนี้</p>
                ) : (
                  <ul className="space-y-1">
                    {recentJobs.map((j) => (
                      <li key={j.id} className="rounded border border-neutral-800 bg-neutral-900 px-2 py-1 text-xs">
                        <p className="truncate font-semibold text-neutral-200">
                          {j.jobName} {j.recipe && <span className="text-indigo-300">· สูตร {j.recipe}</span>}
                        </p>
                        <p className="truncate text-neutral-500">
                          {formatDateOnly(j.jobDate)} · {j.lines.length} วัตถุดิบ
                        </p>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            );
          })}
        </div>
      )}

      {/* Purchase order (PO) panel */}
      {showPoPanel && (
        <div className="mt-4 rounded-lg border border-orange-800/40 bg-orange-950/10 p-3">
          <div className="mb-2 flex items-center justify-between">
            <p className="flex items-center gap-1.5 font-bold uppercase tracking-wide text-orange-300">
              <ShoppingCart size={16} /> ใบสั่งซื้อ
            </p>
            {canEdit && (
              <button
                onClick={openPoForm}
                className="flex items-center gap-1 rounded-md border border-orange-600/50 bg-orange-500/10 px-2 py-1 text-xs font-semibold text-orange-300 hover:bg-orange-500/20"
              >
                <Plus size={12} /> สร้าง PO ใหม่
              </button>
            )}
          </div>
          {purchaseOrders.length === 0 ? (
            <p className="py-3 text-center text-sm text-neutral-500">ยังไม่มีใบสั่งซื้อ</p>
          ) : (
            <ul className="max-h-80 space-y-1.5 overflow-y-auto">
              {purchaseOrders.map((po) => (
                <li key={po.id} className="rounded-md border border-neutral-800 bg-neutral-950 p-2 text-xs">
                  <div className="flex items-center justify-between gap-2">
                    <div className="min-w-0">
                      <p className="font-mono font-semibold text-neutral-200">
                        {po.poNumber} <span className="text-neutral-500">· {po.supplier}</span>
                      </p>
                      <p className="text-neutral-500">
                        สั่ง {formatDateOnly(po.orderDate)}
                        {po.expectedDate ? ` · คาดว่าจะได้รับ ${formatDateOnly(po.expectedDate)}` : ""} · {po.lines.length} รายการ
                      </p>
                      {(po.taxInvoiceNo || po.supplierAddress) && (
                        <p className="text-neutral-500">
                          {po.taxInvoiceNo ? `เลขที่ใบกำกับภาษี: ${po.taxInvoiceNo}` : ""}
                          {po.taxInvoiceNo && po.supplierAddress ? " · " : ""}
                          {po.supplierAddress ? `ที่อยู่: ${po.supplierAddress}` : ""}
                        </p>
                      )}
                    </div>
                    <div className="flex shrink-0 items-center gap-1.5">
                      <span
                        className={`rounded px-2 py-0.5 text-[10px] font-bold uppercase ${
                          po.status === "pending"
                            ? "bg-amber-500/20 text-amber-300"
                            : po.status === "received"
                            ? "bg-emerald-500/20 text-emerald-300"
                            : "bg-neutral-700 text-neutral-400"
                        }`}
                      >
                        {po.status === "pending" ? "รอรับของ" : po.status === "received" ? "รับของแล้ว" : "ยกเลิก"}
                      </span>
                      {po.status === "pending" && canEdit && (
                        <>
                          <button
                            onClick={() => openReceivePo(po)}
                            className="rounded border border-emerald-600/50 bg-emerald-500/10 px-2 py-1 font-semibold text-emerald-300 hover:bg-emerald-500/20"
                          >
                            รับของ
                          </button>
                          <button
                            onClick={() => cancelPo(po.id)}
                            className="rounded border border-neutral-700 px-2 py-1 text-neutral-400 hover:bg-neutral-800"
                          >
                            ยกเลิก
                          </button>
                        </>
                      )}
                      {po.status === "received" && canEdit && (
                        <button
                          onClick={() => setConfirmReversePo(po)}
                          title="ยกเลิกใบสั่งซื้อนี้เนื่องจากรับของผิด (จะตัดสต็อกที่รับเข้าไปออก)"
                          className="rounded border border-red-600/50 bg-red-500/10 px-2 py-1 font-semibold text-red-300 hover:bg-red-500/20"
                        >
                          ยกเลิก (รับผิด)
                        </button>
                      )}
                    </div>
                  </div>
                  <ul className="mt-1.5 space-y-0.5 border-t border-neutral-800 pt-1.5">
                    {po.lines.map((l) => (
                      <li key={l.id} className="flex items-center justify-between text-neutral-400">
                        <span className="truncate">
                          {l.itemName} <span className="font-mono text-neutral-600">({l.sku})</span>
                          {(l.productionDate || l.expiryDate) && (
                            <span className="ml-1 text-neutral-600">
                              {l.productionDate ? ` ผลิต ${formatDateOnly(l.productionDate)}` : ""}
                              {l.expiryDate ? ` หมดอายุ ${formatDateOnly(l.expiryDate)}` : ""}
                            </span>
                          )}
                        </span>
                        <span className="shrink-0 font-mono">{formatNum(l.qty)} {l.unit} × {formatBaht(l.unitPrice)}</span>
                      </li>
                    ))}
                  </ul>
                </li>
              ))}
            </ul>
          )}
        </div>
      )}

      {/* Accounting module panel */}
      {showAccounting && isAccounting && (
        <div className="mt-4 rounded-lg border border-yellow-800/40 bg-yellow-950/10 p-3">
          <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
            <p className="flex items-center gap-1.5 font-bold uppercase tracking-wide text-yellow-300">
              <Wallet size={16} /> ระบบบัญชี (แผนกบัญชี)
            </p>
            <div className="flex flex-wrap gap-1">
              {[
                { id: "journal", label: "บันทึกรายวัน" },
                { id: "accounts", label: "ผังบัญชี" },
                { id: "ledger", label: "บัญชีแยกประเภท" },
                { id: "trial", label: "งบทดลอง" },
                { id: "pnl", label: "งบกำไรขาดทุน" },
                { id: "popay", label: "การจ่ายเงิน PO" },
              ].map((tab) => (
                <button
                  key={tab.id}
                  onClick={() => setAccountingTab(tab.id)}
                  className={`rounded-md px-3 py-1.5 text-xs font-semibold ${
                    accountingTab === tab.id
                      ? "bg-yellow-500 text-neutral-950"
                      : "border border-neutral-700 bg-neutral-900 text-neutral-300 hover:bg-neutral-800"
                  }`}
                >
                  {tab.label}
                </button>
              ))}
            </div>
          </div>

          {accountingTab === "journal" && (
            <div>
              <div className="mb-2 flex justify-end">
                <button
                  onClick={openJournalForm}
                  className="flex items-center gap-1 rounded-md bg-yellow-500 px-3 py-1.5 text-xs font-bold text-neutral-950 hover:bg-yellow-400"
                >
                  <Plus size={13} /> บันทึกรายการใหม่
                </button>
              </div>
              {journalEntries.length === 0 ? (
                <p className="py-6 text-center text-xs text-neutral-500">ยังไม่มีรายการบัญชี</p>
              ) : (
                <div className="max-h-96 overflow-y-auto rounded-md border border-neutral-800">
                  <table className="w-full text-xs">
                    <thead className="bg-neutral-900 text-neutral-500">
                      <tr>
                        <th className="px-2 py-1.5 text-left">วันที่</th>
                        <th className="px-2 py-1.5 text-left">คำอธิบาย</th>
                        <th className="px-2 py-1.5 text-left">บัญชี</th>
                        <th className="px-2 py-1.5 text-right">เดบิต</th>
                        <th className="px-2 py-1.5 text-right">เครดิต</th>
                        <th className="px-2 py-1.5"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {journalEntries
                        .slice()
                        .sort((a, b) => b.ts - a.ts)
                        .map((j) => (
                          <React.Fragment key={j.id}>
                            {j.lines.map((l, idx) => (
                              <tr key={idx} className="border-t border-neutral-800">
                                {idx === 0 && (
                                  <>
                                    <td className="px-2 py-1.5 font-mono text-neutral-400" rowSpan={j.lines.length}>
                                      {formatDateOnly(j.date)}
                                    </td>
                                    <td className="px-2 py-1.5 text-neutral-300" rowSpan={j.lines.length}>
                                      {j.description}
                                      {j.auto && (
                                        <span className="ml-1 rounded bg-neutral-800 px-1 text-[10px] text-neutral-500">อัตโนมัติ</span>
                                      )}
                                    </td>
                                  </>
                                )}
                                <td className="px-2 py-1.5 text-neutral-300">
                                  {accounts.find((a) => a.id === l.accountId)?.name || "?"}
                                </td>
                                <td className="px-2 py-1.5 text-right font-mono text-emerald-400">
                                  {l.debit ? formatBaht(l.debit) : ""}
                                </td>
                                <td className="px-2 py-1.5 text-right font-mono text-red-400">
                                  {l.credit ? formatBaht(l.credit) : ""}
                                </td>
                                {idx === 0 && (
                                  <td className="px-2 py-1.5 text-right" rowSpan={j.lines.length}>
                                    {!j.auto && (
                                      <button onClick={() => deleteJournalEntry(j.id)} className="text-neutral-500 hover:text-red-400">
                                        <Trash2 size={13} />
                                      </button>
                                    )}
                                  </td>
                                )}
                              </tr>
                            ))}
                          </React.Fragment>
                        ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          {accountingTab === "accounts" && (
            <div>
              <div className="mb-2 flex justify-end">
                <button
                  onClick={() => openAccountForm(null)}
                  className="flex items-center gap-1 rounded-md bg-yellow-500 px-3 py-1.5 text-xs font-bold text-neutral-950 hover:bg-yellow-400"
                >
                  <Plus size={13} /> เพิ่มบัญชี
                </button>
              </div>
              <table className="w-full text-xs">
                <thead className="text-neutral-500">
                  <tr>
                    <th className="px-2 py-1 text-left">รหัส</th>
                    <th className="px-2 py-1 text-left">ชื่อบัญชี</th>
                    <th className="px-2 py-1 text-left">ประเภท</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {accounts
                    .slice()
                    .sort((a, b) => a.code.localeCompare(b.code))
                    .map((a) => (
                      <tr key={a.id} className="border-t border-neutral-800">
                        <td className="px-2 py-1.5 font-mono text-neutral-300">{a.code}</td>
                        <td className="px-2 py-1.5 text-neutral-200">{a.name}</td>
                        <td className="px-2 py-1.5 text-neutral-400">{accountTypeLabel(a.type)}</td>
                        <td className="px-2 py-1.5 text-right">
                          <button onClick={() => openAccountForm(a)} className="mr-1 text-neutral-500 hover:text-amber-400">
                            <Pencil size={13} />
                          </button>
                          <button onClick={() => deleteAccount(a.id)} className="text-neutral-500 hover:text-red-400">
                            <Trash2 size={13} />
                          </button>
                        </td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>
          )}

          {accountingTab === "ledger" && (
            <div>
              <Field label="เลือกบัญชี">
                <select className={inputCls} value={ledgerAccountId} onChange={(e) => setLedgerAccountId(e.target.value)}>
                  <option value="">-- เลือกบัญชี --</option>
                  {accounts.map((a) => (
                    <option key={a.id} value={a.id}>
                      {a.code} - {a.name}
                    </option>
                  ))}
                </select>
              </Field>
              {ledgerAccountId &&
                (ledgerEntries.length === 0 ? (
                  <p className="py-6 text-center text-xs text-neutral-500">ยังไม่มีรายการ</p>
                ) : (
                  <div className="max-h-96 overflow-y-auto rounded-md border border-neutral-800">
                    <table className="w-full text-xs">
                      <thead className="bg-neutral-900 text-neutral-500">
                        <tr>
                          <th className="px-2 py-1.5 text-left">วันที่</th>
                          <th className="px-2 py-1.5 text-left">คำอธิบาย</th>
                          <th className="px-2 py-1.5 text-right">เดบิต</th>
                          <th className="px-2 py-1.5 text-right">เครดิต</th>
                          <th className="px-2 py-1.5 text-right">ยอดคงเหลือ</th>
                        </tr>
                      </thead>
                      <tbody>
                        {ledgerEntries.map((r, idx) => (
                          <tr key={idx} className="border-t border-neutral-800">
                            <td className="px-2 py-1.5 font-mono text-neutral-400">{formatDateOnly(r.date)}</td>
                            <td className="px-2 py-1.5 text-neutral-300">{r.description}</td>
                            <td className="px-2 py-1.5 text-right font-mono text-emerald-400">{r.debit ? formatBaht(r.debit) : ""}</td>
                            <td className="px-2 py-1.5 text-right font-mono text-red-400">{r.credit ? formatBaht(r.credit) : ""}</td>
                            <td className="px-2 py-1.5 text-right font-mono text-neutral-100">{formatBaht(r.running)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ))}
            </div>
          )}

          {accountingTab === "trial" && (
            <div>
              <table className="w-full text-xs">
                <thead className="text-neutral-500">
                  <tr>
                    <th className="px-2 py-1 text-left">รหัส</th>
                    <th className="px-2 py-1 text-left">บัญชี</th>
                    <th className="px-2 py-1 text-right">เดบิต</th>
                    <th className="px-2 py-1 text-right">เครดิต</th>
                  </tr>
                </thead>
                <tbody>
                  {trialBalance
                    .filter((a) => a.debit || a.credit)
                    .map((a) => (
                      <tr key={a.id} className="border-t border-neutral-800">
                        <td className="px-2 py-1.5 font-mono text-neutral-400">{a.code}</td>
                        <td className="px-2 py-1.5 text-neutral-200">{a.name}</td>
                        <td className="px-2 py-1.5 text-right font-mono text-neutral-100">
                          {isDebitNormal(a.type) && a.balance > 0 ? formatBaht(a.balance) : ""}
                        </td>
                        <td className="px-2 py-1.5 text-right font-mono text-neutral-100">
                          {!isDebitNormal(a.type) && a.balance > 0 ? formatBaht(a.balance) : ""}
                        </td>
                      </tr>
                    ))}
                </tbody>
                <tfoot>
                  <tr className="border-t-2 border-neutral-700 font-bold">
                    <td className="px-2 py-1.5" colSpan={2}>รวม</td>
                    <td className="px-2 py-1.5 text-right font-mono text-neutral-100">{formatBaht(trialTotals.debit)}</td>
                    <td className="px-2 py-1.5 text-right font-mono text-neutral-100">{formatBaht(trialTotals.credit)}</td>
                  </tr>
                </tfoot>
              </table>
              <p className={`mt-2 text-xs ${Math.abs(trialTotals.debit - trialTotals.credit) < 0.01 ? "text-emerald-400" : "text-red-400"}`}>
                {Math.abs(trialTotals.debit - trialTotals.credit) < 0.01
                  ? "✓ ยอดเดบิตเท่ากับเครดิต ถูกต้อง"
                  : "⚠ ยอดเดบิตไม่เท่ากับเครดิต ตรวจสอบรายการ"}
              </p>
            </div>
          )}

          {accountingTab === "pnl" && (
            <div>
              <div className="mb-3 grid grid-cols-2 gap-3">
                <Field label="จากวันที่">
                  <DateField value={pnlFrom} onChange={setPnlFrom} />
                </Field>
                <Field label="ถึงวันที่">
                  <DateField value={pnlTo} onChange={setPnlTo} />
                </Field>
              </div>
              <div className="space-y-1 rounded-md border border-neutral-800 bg-neutral-950 p-3 text-sm">
                <div className="flex justify-between">
                  <span className="text-neutral-400">รายได้จากการขาย</span>
                  <span className="font-mono text-neutral-100">{formatBaht(pnlData.revenue)}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-neutral-400">หัก ต้นทุนขาย (COGS)</span>
                  <span className="font-mono text-red-400">-{formatBaht(pnlData.cogs)}</span>
                </div>
                <div className="flex justify-between border-t border-neutral-800 pt-1 font-semibold">
                  <span className="text-neutral-200">กำไรขั้นต้น</span>
                  <span className="font-mono text-neutral-100">{formatBaht(pnlData.grossProfit)}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-neutral-400">หัก ค่าใช้จ่ายดำเนินงาน</span>
                  <span className="font-mono text-red-400">-{formatBaht(pnlData.opex)}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-neutral-400">หัก ของเสีย/สูญหาย/ปรับปรุงสต็อก</span>
                  <span className="font-mono text-red-400">-{formatBaht(pnlData.writeoff)}</span>
                </div>
                <div className="flex justify-between border-t border-neutral-700 pt-1 text-base font-bold">
                  <span className="text-neutral-100">กำไร(ขาดทุน)สุทธิ</span>
                  <span className={`font-mono ${pnlData.netIncome >= 0 ? "text-emerald-400" : "text-red-400"}`}>
                    {formatBaht(pnlData.netIncome)}
                  </span>
                </div>
              </div>
              <p className="mt-2 text-[11px] text-neutral-500">
                หมายเหตุ: ต้นทุนขายคำนวณจากราคาต่อหน่วยปัจจุบันของสินค้า ณ เวลาที่ทำรายการ (ไม่ได้แยกต้นทุนตามล็อตแบบ FEFO อย่างละเอียด)
              </p>
            </div>
          )}

          {accountingTab === "popay" && (
            <div>
              {purchaseOrders.filter((po) => po.status !== "cancelled").length === 0 ? (
                <p className="py-6 text-center text-xs text-neutral-500">ยังไม่มีใบสั่งซื้อ</p>
              ) : (
                <ul className="space-y-1.5">
                  {purchaseOrders
                    .filter((po) => po.status !== "cancelled")
                    .map((po) => {
                      const status = poPaymentStatus(po);
                      const balance = poBalanceDue(po);
                      return (
                        <li key={po.id} className="rounded-md border border-neutral-800 bg-neutral-950 p-2 text-xs">
                          <div className="flex items-center justify-between gap-2">
                            <div>
                              <p className="font-mono font-semibold text-neutral-200">
                                {po.poNumber} <span className="text-neutral-500">· {po.supplier}</span>
                              </p>
                              <p className="text-neutral-500">
                                ยอดรวม {formatBaht(poTotal(po))} · จ่ายแล้ว {formatBaht(poPaidTotal(po))} · ค้างชำระ {formatBaht(balance)}
                              </p>
                            </div>
                            <div className="flex items-center gap-1.5">
                              <span
                                className={`rounded px-2 py-0.5 text-[10px] font-bold uppercase ${
                                  status === "paid"
                                    ? "bg-emerald-500/20 text-emerald-300"
                                    : status === "partial"
                                    ? "bg-amber-500/20 text-amber-300"
                                    : "bg-red-500/20 text-red-300"
                                }`}
                              >
                                {status === "paid" ? "จ่ายครบ" : status === "partial" ? "จ่ายบางส่วน" : "ยังไม่จ่าย"}
                              </span>
                              {balance > 0.01 && canEdit && (
                                <button
                                  onClick={() => {
                                    setPayingPo(po);
                                    setPoPaymentForm({ amount: String(balance), date: todayStr(), method: "โอนเงิน", reference: "", note: "" });
                                  }}
                                  className="rounded border border-yellow-600/50 bg-yellow-500/10 px-2 py-1 font-semibold text-yellow-300 hover:bg-yellow-500/20"
                                >
                                  บันทึกจ่ายเงิน
                                </button>
                              )}
                            </div>
                          </div>
                          {(po.payments || []).length > 0 && (
                            <ul className="mt-1.5 space-y-0.5 border-t border-neutral-800 pt-1.5 text-neutral-500">
                              {po.payments.map((p) => (
                                <li key={p.id} className="flex justify-between">
                                  <span>
                                    {formatDateOnly(p.date)} · {p.method}
                                    {p.reference ? ` · อ้างอิง ${p.reference}` : ""}
                                  </span>
                                  <span className="font-mono">{formatBaht(p.amount)}</span>
                                </li>
                              ))}
                            </ul>
                          )}
                        </li>
                      );
                    })}
                </ul>
              )}
            </div>
          )}
        </div>
      )}

      {/* History panel */}
      {showHistory && (
        <div className="mt-4 rounded-lg border border-neutral-800 bg-neutral-900 p-3">
          <div className="mb-2 flex flex-wrap gap-2">
            <div className="relative flex-1 min-w-[160px]">
              <Search size={13} className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-neutral-500" />
              <input
                value={historySearch}
                onChange={(e) => setHistorySearch(e.target.value)}
                placeholder="ค้นหาสินค้า, SKU หรือสถานที่ เพื่อดูว่าโอนไปไหน"
                className={`${inputCls} py-1.5 pl-8 text-xs`}
              />
            </div>
            <select
              value={historyTypeFilter}
              onChange={(e) => setHistoryTypeFilter(e.target.value)}
              className={`${inputCls} w-auto py-1.5 text-xs`}
            >
              <option value="all">ทุกประเภทรายการ</option>
              <option value="in">รับเข้า</option>
              <option value="out">เบิกออก</option>
              <option value="to_prod">โอนไปผลิต</option>
              <option value="from_prod">คืนจากผลิต</option>
              <option value="consume_prod">ตัดวัตถุดิบ (ผลิต)</option>
              <option value="waste">ตัดของเสีย/สูญหาย</option>
              <option value="adjust">ปรับปรุงสต็อก</option>
            </select>
          </div>
          <div className="max-h-64 overflow-y-auto">
            {filteredTransactions.length === 0 ? (
              <p className="py-4 text-center text-sm text-neutral-500">
                {transactions.length === 0 ? "ยังไม่มีประวัติการเคลื่อนไหว" : "ไม่พบรายการที่ค้นหา"}
              </p>
            ) : (
              <ul className="space-y-1.5">
                {filteredTransactions.slice(0, 100).map((tx) => (
                  <li key={tx.id} className="flex items-center justify-between gap-2 rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2 text-xs">
                    <div className="flex items-center gap-2 min-w-0">
                      {tx.type === "in" && <ArrowDownCircle size={15} className="shrink-0 text-emerald-400" />}
                      {tx.type === "out" && <ArrowUpCircle size={15} className="shrink-0 text-red-400" />}
                      {tx.type === "to_prod" && <Factory size={15} className="shrink-0 text-indigo-400" />}
                      {tx.type === "from_prod" && <Undo2 size={15} className="shrink-0 text-teal-400" />}
                      {tx.type === "consume_prod" && <Scissors size={15} className="shrink-0 text-rose-400" />}
                      {tx.type === "waste" && <Ban size={15} className="shrink-0 text-rose-500" />}
                      {tx.type === "adjust" && <SlidersHorizontal size={15} className="shrink-0 text-fuchsia-400" />}
                      <div className="min-w-0">
                        <p className="truncate font-semibold text-neutral-200">{tx.itemName}</p>
                        <p className="truncate font-mono text-neutral-500">{tx.sku}</p>
                        {tx.type === "consume_prod" && tx.jobName && (
                          <p className="truncate text-[11px] text-rose-300">ตัดสำหรับผลิต: {tx.jobName}</p>
                        )}
                        {(tx.type === "waste" || tx.type === "adjust") && tx.reason && (
                          <p className="truncate text-[11px] text-neutral-400">เหตุผล: {tx.reason}</p>
                        )}
                        {tx.type === "adjust" && tx.lotSummary && (
                          <p className="truncate text-[11px] text-fuchsia-300">{tx.lotSummary}</p>
                        )}
                        {(tx.fromLocation || tx.toLocation) && (
                          <div className="mt-0.5 flex items-center gap-1 text-[11px] text-neutral-400">
                            {tx.fromLocation && <span>{locationLabel(tx.fromLocation)}</span>}
                            {tx.fromLocation && tx.toLocation && <ArrowRight size={10} />}
                            {tx.toLocation && <span>{locationLabel(tx.toLocation)}</span>}
                          </div>
                        )}
                        <p className="truncate text-neutral-500">
                          {tx.requestedBy ? `ผู้ทำรายการ: ${tx.requestedBy}` : ""}
                          {tx.note ? ` · ${tx.note}` : ""}
                        </p>
                      </div>
                    </div>
                    <div className="shrink-0 text-right">
                      <p
                        className={`font-mono font-bold ${
                          tx.type === "in" || (tx.type === "adjust" && tx.amount > 0)
                            ? "text-emerald-400"
                            : tx.type === "out" || tx.type === "consume_prod" || tx.type === "waste" || (tx.type === "adjust" && tx.amount < 0)
                            ? "text-red-400"
                            : "text-neutral-300"
                        }`}
                      >
                        {tx.type === "adjust"
                          ? `${tx.amount > 0 ? "+" : ""}${formatNum(tx.amount)}`
                          : tx.type === "in"
                          ? `+${formatNum(tx.amount)}`
                          : tx.type === "out" || tx.type === "consume_prod" || tx.type === "waste"
                          ? `-${formatNum(tx.amount)}`
                          : formatNum(tx.amount)}
                      </p>
                      <p className="text-neutral-600">{formatDateTime(tx.ts)}</p>
                    </div>
                  </li>
                ))}
              </ul>
            )}
          </div>
        </div>
      )}

      {/* Item edit-history panel */}
      {showEditLog && (
        <div className="mt-4 rounded-lg border border-neutral-800 bg-neutral-900 p-3">
          <div className="relative mb-2">
            <Search size={13} className="pointer-events-none absolute left-2.5 top-1/2 -translate-y-1/2 text-neutral-500" />
            <input
              value={editLogSearch}
              onChange={(e) => setEditLogSearch(e.target.value)}
              placeholder="ค้นหาชื่อสินค้าหรือ SKU"
              className={`${inputCls} py-1.5 pl-8 text-xs`}
            />
          </div>
          {(() => {
            const q = editLogSearch.trim().toLowerCase();
            const filtered = itemEdits.filter(
              (e) => !q || e.itemName.toLowerCase().includes(q) || e.sku.toLowerCase().includes(q)
            );
            return filtered.length === 0 ? (
              <p className="py-4 text-center text-sm text-neutral-500">
                {itemEdits.length === 0 ? "ยังไม่มีการแก้ไขข้อมูลสินค้า" : "ไม่พบรายการที่ค้นหา"}
              </p>
            ) : (
              <div className="max-h-64 overflow-y-auto">
                <ul className="space-y-1.5">
                  {filtered.slice(0, 100).map((e) => (
                    <li key={e.id} className="rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2 text-xs">
                      <div className="mb-1 flex items-center justify-between">
                        <p className="font-semibold text-neutral-200">
                          {e.itemName} <span className="font-mono text-neutral-500">({e.sku})</span>
                        </p>
                        <span className="text-neutral-600">{formatDateTime(e.ts)}</span>
                      </div>
                      <ul className="space-y-0.5 text-neutral-400">
                        {e.changes.map((c, idx) => (
                          <li key={idx}>
                            {c.field}: <span className="text-red-400">{String(c.from) || "-"}</span> →{" "}
                            <span className="text-emerald-400">{String(c.to) || "-"}</span>
                          </li>
                        ))}
                      </ul>
                      {e.changedBy && <p className="mt-1 text-neutral-600">แก้ไขโดย: {e.changedBy}</p>}
                    </li>
                  ))}
                </ul>
              </div>
            );
          })()}
        </div>
      )}

      {/* Pending-approvals panel */}
      {showApprovals && canEdit && (
        <div className="mt-4 rounded-lg border border-rose-800/40 bg-rose-950/10 p-3">
          <p className="mb-2 flex items-center gap-1.5 font-bold uppercase tracking-wide text-rose-300">
            <Clock size={16} /> คำขอเบิกที่รออนุมัติ
          </p>
          {approvals.length === 0 ? (
            <p className="py-4 text-center text-sm text-neutral-500">ยังไม่มีคำขออนุมัติ</p>
          ) : (
            <ul className="max-h-80 space-y-1.5 overflow-y-auto">
              {approvals
                .slice()
                .sort((a, b) => b.ts - a.ts)
                .map((a) => (
                  <li key={a.id} className="rounded-md border border-neutral-800 bg-neutral-950 p-2 text-xs">
                    <div className="flex items-center justify-between gap-2">
                      <div>
                        <p className="font-semibold text-neutral-200">
                          {a.itemName} <span className="font-mono text-neutral-500">({a.sku})</span>
                        </p>
                        <p className="text-neutral-500">
                          ขอเบิก {formatNum(a.amount)} {a.unit} · ผู้ขอ: {a.requestedBy || "-"} · {formatDateTime(a.ts)}
                        </p>
                        {a.note && <p className="text-neutral-500">หมายเหตุ: {a.note}</p>}
                      </div>
                      <div className="flex shrink-0 items-center gap-1.5">
                        <span
                          className={`rounded px-2 py-0.5 text-[10px] font-bold uppercase ${
                            a.status === "pending"
                              ? "bg-amber-500/20 text-amber-300"
                              : a.status === "approved"
                              ? "bg-emerald-500/20 text-emerald-300"
                              : "bg-neutral-700 text-neutral-400"
                          }`}
                        >
                          {a.status === "pending" ? "รออนุมัติ" : a.status === "approved" ? "อนุมัติแล้ว" : "ปฏิเสธ"}
                        </span>
                        {a.status === "pending" && isAdmin && (
                          <>
                            <button
                              onClick={() => approveWithdrawalRequest(a)}
                              className="rounded border border-emerald-600/50 bg-emerald-500/10 px-2 py-1 font-semibold text-emerald-300 hover:bg-emerald-500/20"
                            >
                              อนุมัติ
                            </button>
                            <button
                              onClick={() => rejectWithdrawalRequest(a.id)}
                              className="rounded border border-neutral-700 px-2 py-1 text-neutral-400 hover:bg-neutral-800"
                            >
                              ปฏิเสธ
                            </button>
                          </>
                        )}
                      </div>
                    </div>
                  </li>
                ))}
            </ul>
          )}
        </div>
      )}

      {/* Toolbar */}
      <div className="mt-5 flex flex-wrap items-center gap-2">
        <div className="relative flex-1 min-w-[180px]">
          <Search size={15} className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-neutral-500" />
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="ค้นหาชื่อสินค้า, SKU, หรือตำแหน่งจัดเก็บ"
            className={`${inputCls} pl-9`}
          />
        </div>
        <div className="relative">
          <select
            value={filter}
            onChange={(e) => setFilter(e.target.value)}
            className={`${inputCls} appearance-none pr-8`}
          >
            <option value="all">แสดงทั้งหมด</option>
            <option value="low">เฉพาะใกล้หมด</option>
          </select>
          <ChevronDown size={14} className="pointer-events-none absolute right-2.5 top-1/2 -translate-y-1/2 text-neutral-500" />
        </div>
        <div className="relative">
          <select
            value={categoryFilter}
            onChange={(e) => setCategoryFilter(e.target.value)}
            className={`${inputCls} appearance-none pr-8`}
          >
            <option value="all">ทุกหมวดหมู่</option>
            {CATEGORIES.map((c) => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
          <ChevronDown size={14} className="pointer-events-none absolute right-2.5 top-1/2 -translate-y-1/2 text-neutral-500" />
        </div>
        {canEdit && (
          <>
            <button
              onClick={openAddForm}
              className="flex items-center gap-1.5 rounded-md bg-amber-500 px-3 py-2 text-sm font-bold text-neutral-950 hover:bg-amber-400"
            >
              <Plus size={16} /> {t("addItem")}
            </button>
            <button
              onClick={() => openJobForm()}
              className="flex items-center gap-1.5 rounded-md border border-indigo-600/50 bg-indigo-500/10 px-3 py-2 text-sm font-bold text-indigo-400 hover:bg-indigo-500/20"
            >
              <Scissors size={16} /> {t("cutMaterial")}
            </button>
            <button
              onClick={() => openFgForm()}
              className="flex items-center gap-1.5 rounded-md border border-teal-600/50 bg-teal-500/10 px-3 py-2 text-sm font-bold text-teal-400 hover:bg-teal-500/20"
            >
              <ClipboardList size={16} /> {t("fgCheck")}
            </button>
          </>
        )}
        <button
          onClick={exportExcel}
          className="flex items-center gap-1.5 rounded-md border border-emerald-600/50 bg-emerald-500/10 px-3 py-2 text-sm font-bold text-emerald-400 hover:bg-emerald-500/20"
        >
          <FileSpreadsheet size={16} /> {t("exportExcel")}
        </button>
        {canEdit && (
          <>
            <button
              onClick={() => fileInputRef.current?.click()}
              className="flex items-center gap-1.5 rounded-md border border-sky-600/50 bg-sky-500/10 px-3 py-2 text-sm font-bold text-sky-400 hover:bg-sky-500/20"
            >
              <Upload size={16} /> {t("importExcel")}
            </button>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={(e) => {
                const file = e.target.files && e.target.files[0];
                if (file) importExcel(file);
                e.target.value = "";
              }}
            />
          </>
        )}
        <button
          onClick={generateCountSheet}
          className="flex items-center gap-1.5 rounded-md border border-neutral-700 bg-neutral-900 px-3 py-2 text-sm font-bold text-neutral-300 hover:bg-neutral-800"
        >
          <ClipboardList size={16} /> {t("countSheet")}
        </button>
        <button
          onClick={() => setShowPoPanel((s) => !s)}
          className="flex items-center gap-1.5 rounded-md border border-orange-600/50 bg-orange-500/10 px-3 py-2 text-sm font-bold text-orange-400 hover:bg-orange-500/20"
        >
          <ShoppingCart size={16} /> {t("po")}
        </button>
      </div>

      {/* Table */}
      <div className="mt-4 overflow-x-auto rounded-lg border border-neutral-800">
        <table className="w-full min-w-[760px] border-collapse text-sm">
          <thead>
            <tr className="border-b border-neutral-800 bg-neutral-900 text-left text-xs uppercase tracking-wide text-neutral-500">
              <th className="px-3 py-2.5"></th>
              <th className="px-3 py-2.5">{t("colProduct")}</th>
              <th className="px-3 py-2.5">{t("colLocation")}</th>
              <th className="px-3 py-2.5 text-right">{t("colQty")}</th>
              <th className="px-3 py-2.5 text-right">{t("colValue")}</th>
              <th className="px-3 py-2.5 text-center">{t("colActions")}</th>
            </tr>
          </thead>
          <tbody>
            {filteredItems.length === 0 ? (
              <tr>
                <td colSpan={6} className="px-3 py-10 text-center text-neutral-500">
                  {items.length === 0 ? "ยังไม่มีสินค้าในคลัง เริ่มเพิ่มสินค้าชิ้นแรก" : "ไม่พบสินค้าที่ค้นหา"}
                </td>
              </tr>
            ) : (
              filteredItems.map((it) => {
                const qty = totalQty(it);
                const wQty = warehouseQty(it);
                const pQty = productionQty(it);
                const low = qty <= it.lowStock;
                const nearExp = nearestExpiry(it);
                const expDays = daysUntil(nearExp);
                const lots = (it.lots || []).filter((l) => l.qty > 0);
                const locs = locationBreakdown(it);
                const expanded = expandedItemId === it.id;
                return (
                  <React.Fragment key={it.id}>
                    <tr className="border-b border-neutral-900 hover:bg-neutral-900/60">
                      <td className="px-2 py-2.5 text-center align-top">
                        {lots.length > 0 && (
                          <button
                            onClick={() => setExpandedItemId(expanded ? null : it.id)}
                            className="rounded p-1 text-neutral-500 hover:bg-neutral-800 hover:text-neutral-200"
                            title="ดูล็อตสินค้า"
                          >
                            {expanded ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
                          </button>
                        )}
                      </td>
                      <td className="px-3 py-2.5">
                        <p className="font-semibold text-neutral-100">{it.name}</p>
                        <p className="font-mono text-xs text-neutral-500">{it.sku}</p>
                        <span className="mt-0.5 inline-flex items-center gap-1 rounded bg-neutral-800 px-1.5 py-0.5 text-[10px] text-neutral-400">
                          <Tag size={9} /> {it.category || CATEGORIES[0]}
                        </span>
                        {it.supplier && <p className="text-xs text-neutral-500">ผู้ขาย: {it.supplier}</p>}
                        {lots.length > 1 && (
                          <p className="mt-0.5 flex items-center gap-1 text-[11px] text-sky-400">
                            <Layers size={11} /> {lots.length} ล็อต
                          </p>
                        )}
                      </td>
                      <td className="px-3 py-2.5">
                        <div className="flex flex-wrap gap-1">
                          {locs.length === 0 ? (
                            <LocationTag location="" />
                          ) : (
                            locs.map(([loc, locQty]) => <LocationTag key={loc || "none"} location={loc} qty={locQty} unit={it.unit} />)
                          )}
                        </div>
                        {nearExp && (
                          <div
                            className={`mt-1 flex items-center gap-1 text-[11px] ${
                              expDays < 0 ? "text-red-400" : expDays <= 30 ? "text-amber-400" : "text-neutral-500"
                            }`}
                          >
                            <AlertTriangle size={10} className={expDays <= 30 ? "opacity-100" : "opacity-0"} />
                            {expDays < 0 ? "หมดอายุแล้ว " : "หมดอายุใกล้สุด "}
                            {formatDateOnly(nearExp)}
                          </div>
                        )}
                      </td>
                      <td className="px-3 py-2.5 text-right">
                        <span className={`font-mono font-bold ${low ? "text-red-400" : "text-neutral-100"}`}>
                          {formatNum(qty)}
                        </span>
                        <span className="ml-1 text-xs text-neutral-500">{it.unit}</span>
                        {low && (
                          <div className="mt-0.5 flex items-center justify-end gap-1 text-[10px] text-red-400">
                            <AlertTriangle size={10} /> ต่ำกว่า {formatNum(it.lowStock)}
                          </div>
                        )}
                      </td>
                      <td className="px-3 py-2.5 text-right font-mono text-neutral-300">{formatBaht(it.price)}</td>
                      <td className="px-3 py-2.5">
                        <div className="flex flex-wrap items-center justify-center gap-1">
                          {canEdit && (
                            <>
                              <button
                                title="รับเข้า"
                                onClick={() => openMovement(it, "in")}
                                className="rounded p-1.5 text-emerald-400 hover:bg-emerald-500/10"
                              >
                                <ArrowDownCircle size={17} />
                              </button>
                              <button
                                title={wQty <= 0 ? "สินค้าหมดสต็อก เบิกออกไม่ได้" : "เบิกออก"}
                                onClick={() =>
                                  wQty <= 0
                                    ? setAlertMessage(`${it.name} (${it.sku}) หมดสต็อกแล้ว ไม่สามารถเบิกออกได้`)
                                    : openMovement(it, "out")
                                }
                                disabled={wQty <= 0}
                                className={`rounded p-1.5 ${
                                  wQty <= 0 ? "cursor-not-allowed text-neutral-700" : "text-red-400 hover:bg-red-500/10"
                                }`}
                              >
                                <ArrowUpCircle size={17} />
                              </button>
                              <button
                                title={wQty <= 0 ? "ไม่มีสต็อกในคลังให้โอนไปผลิต" : "โอนไปฝ่ายผลิต"}
                                onClick={() =>
                                  wQty <= 0
                                    ? setAlertMessage(`${it.name} (${it.sku}) ไม่มีสต็อกในคลังให้โอนไปฝ่ายผลิต`)
                                    : openMovement(it, "to_prod")
                                }
                                disabled={wQty <= 0}
                                className={`rounded p-1.5 ${
                                  wQty <= 0 ? "cursor-not-allowed text-neutral-700" : "text-indigo-400 hover:bg-indigo-500/10"
                                }`}
                              >
                                <Factory size={17} />
                              </button>
                              <button
                                title={pQty <= 0 ? "ไม่มีสต็อกค้างที่ฝ่ายผลิต" : "คืนจากฝ่ายผลิต"}
                                onClick={() =>
                                  pQty <= 0
                                    ? setAlertMessage(`${it.name} (${it.sku}) ไม่มีสต็อกค้างอยู่ที่ฝ่ายผลิต`)
                                    : openMovement(it, "from_prod")
                                }
                                disabled={pQty <= 0}
                                className={`rounded p-1.5 ${
                                  pQty <= 0 ? "cursor-not-allowed text-neutral-700" : "text-teal-400 hover:bg-teal-500/10"
                                }`}
                              >
                                <Undo2 size={17} />
                              </button>
                              <button
                                title={wQty <= 0 ? "ไม่มีสต็อกให้ตัดของเสีย/สูญหาย" : "ตัดของเสีย/สูญหาย"}
                                onClick={() =>
                                  wQty <= 0
                                    ? setAlertMessage(`${it.name} (${it.sku}) หมดสต็อกแล้ว ไม่สามารถตัดของเสีย/สูญหายได้`)
                                    : openMovement(it, "waste")
                                }
                                disabled={wQty <= 0}
                                className={`rounded p-1.5 ${
                                  wQty <= 0 ? "cursor-not-allowed text-neutral-700" : "text-rose-400 hover:bg-rose-500/10"
                                }`}
                              >
                                <Ban size={16} />
                              </button>
                              <button
                                title="ปรับปรุงสต็อก (นับสต็อกจริง)"
                                onClick={() => openAdjustForm(it)}
                                className="rounded p-1.5 text-fuchsia-400 hover:bg-fuchsia-500/10"
                              >
                                <SlidersHorizontal size={16} />
                              </button>
                            </>
                          )}
                          <button
                            title="ดูประวัติการเคลื่อนไหวของสินค้านี้"
                            onClick={() => {
                              setHistorySearch(it.sku);
                              setHistoryTypeFilter("all");
                              setShowHistory(true);
                            }}
                            className="rounded p-1.5 text-sky-400 hover:bg-sky-500/10"
                          >
                            <History size={16} />
                          </button>
                          {canEdit && (
                            <button
                              title="แก้ไข"
                              onClick={() => openEditForm(it)}
                              className="rounded p-1.5 text-neutral-400 hover:bg-neutral-800 hover:text-neutral-100"
                            >
                              <Pencil size={16} />
                            </button>
                          )}
                          {isAdmin && (
                            <button
                              title="ลบ"
                              onClick={() => setConfirmDelete(it)}
                              className="rounded p-1.5 text-neutral-400 hover:bg-neutral-800 hover:text-red-400"
                            >
                              <Trash2 size={16} />
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                    {expanded && lots.length > 0 && (
                      <tr className="border-b border-neutral-900 bg-neutral-900/40">
                        <td></td>
                        <td colSpan={5} className="px-3 py-2.5">
                          <table className="w-full text-xs">
                            <thead>
                              <tr className="text-neutral-500">
                                <th className="px-2 py-1 text-left">สถานที่</th>
                                <th className="px-2 py-1 text-left">วันที่ผลิต</th>
                                <th className="px-2 py-1 text-left">วันหมดอายุ</th>
                                <th className="px-2 py-1 text-left">วันที่รับเข้า</th>
                                <th className="px-2 py-1 text-right">คงเหลือในล็อต</th>
                              </tr>
                            </thead>
                            <tbody>
                              {sortFEFO(lots).map((l) => {
                                  const d = daysUntil(l.expiryDate);
                                  return (
                                    <tr key={l.id} className="border-t border-neutral-800">
                                      <td className="px-2 py-1.5">
                                        <LocationTag location={l.location} />
                                      </td>
                                      <td className="px-2 py-1.5 font-mono text-neutral-300">{formatDateOnly(l.productionDate) || "-"}</td>
                                      <td className={`px-2 py-1.5 font-mono ${d !== null && d < 0 ? "text-red-400" : d !== null && d <= 30 ? "text-amber-400" : "text-neutral-300"}`}>
                                        {formatDateOnly(l.expiryDate) || "-"}
                                      </td>
                                      <td className="px-2 py-1.5 font-mono text-neutral-400">{formatDateOnly(l.receivedDate) || "-"}</td>
                                      <td className="px-2 py-1.5 text-right font-mono text-neutral-100">
                                        {formatNum(l.qty)} {it.unit}
                                      </td>
                                    </tr>
                                  );
                                })}
                            </tbody>
                          </table>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* Add/Edit item modal */}
      {showForm && editingItem && (
        <Modal
          title={editingItem.id ? "แก้ไขสินค้า" : "เพิ่มสินค้าใหม่"}
          onClose={() => setShowForm(false)}
          footer={
            <div className="flex justify-end gap-2">
              <button type="button" onClick={() => setShowForm(false)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={saveItemForm}
                disabled={!editingItem.name.trim() || !editingItem.sku.trim()}
                className="rounded-md bg-amber-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-amber-400 disabled:cursor-not-allowed disabled:opacity-40"
              >
                บันทึก
              </button>
            </div>
          }
        >
          <div>
            <Field label="ชื่อสินค้า">
              <input
                className={inputCls}
                value={editingItem.name}
                onChange={(e) => setEditingItem({ ...editingItem, name: e.target.value })}
                placeholder="เช่น น็อตหกเหลี่ยม M8"
              />
            </Field>
            <div className="grid grid-cols-2 gap-3">
              <Field label="รหัส SKU">
                <input
                  className={inputCls}
                  value={editingItem.sku}
                  onChange={(e) => setEditingItem({ ...editingItem, sku: e.target.value })}
                  placeholder="เช่น BOLT-M8-001"
                />
              </Field>
              <Field label="ชื่อผู้ขาย / ซัพพลายเออร์">
                <input
                  className={inputCls}
                  value={editingItem.supplier}
                  onChange={(e) => setEditingItem({ ...editingItem, supplier: e.target.value })}
                  placeholder="เช่น บริษัท เอบีซี จำกัด"
                />
              </Field>
            </div>
            <Field label="หมวดหมู่">
              <select
                className={inputCls}
                value={editingItem.category}
                onChange={(e) => setEditingItem({ ...editingItem, category: e.target.value })}
              >
                {CATEGORIES.map((c) => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </Field>

            <div className="mb-1 rounded-md border border-neutral-800 bg-neutral-950 p-3">
              <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-neutral-400">
                หน่วยบรรจุ (ไม่บังคับ) — ใช้ตอนรับเข้าเป็นลัง/กล่อง แล้วแปลงเป็นหน่วยหลักอัตโนมัติ
              </p>
              <div className="grid grid-cols-2 gap-3">
                <Field label="ชื่อหน่วยบรรจุ">
                  <input
                    className={inputCls}
                    placeholder="เช่น กล่อง, ลัง"
                    value={editingItem.packUnit}
                    onChange={(e) => setEditingItem({ ...editingItem, packUnit: e.target.value })}
                  />
                </Field>
                <Field label={`1 หน่วยบรรจุ = กี่ ${editingItem.unit}`}>
                  <input
                    type="number"
                    min="0"
                    className={inputCls}
                    placeholder="เช่น 12"
                    value={editingItem.packFactor}
                    onChange={(e) => setEditingItem({ ...editingItem, packFactor: e.target.value })}
                  />
                </Field>
              </div>
            </div>

            {!editingItem.id && (
              <div className="mb-1 rounded-md border border-neutral-800 bg-neutral-950 p-3">
                <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-neutral-400">
                  ล็อตแรกเริ่ม (ไม่บังคับ ใส่ 0 ได้ถ้ายังไม่มีสต็อก)
                </p>
                <div className="grid grid-cols-2 gap-3">
                  <Field label="จำนวน">
                    <input
                      type="number"
                      min="0"
                      className={inputCls}
                      value={editingItem.initialQty}
                      onChange={(e) => setEditingItem({ ...editingItem, initialQty: e.target.value })}
                    />
                  </Field>
                  <Field label="สถานที่จัดเก็บ">
                    <select
                      className={inputCls}
                      value={editingItem.initialLocation}
                      onChange={(e) => setEditingItem({ ...editingItem, initialLocation: e.target.value })}
                    >
                      {LOCATIONS.map((loc) => (
                        <option key={loc} value={loc}>{loc}</option>
                      ))}
                    </select>
                  </Field>
                </div>
                <div className="mt-3 grid grid-cols-2 gap-3">
                  <Field label="วันที่ผลิต">
                    <DateField
                      value={editingItem.initialProductionDate}
                      onChange={(v) => setEditingItem({ ...editingItem, initialProductionDate: v })}
                    />
                  </Field>
                  <Field label="วันหมดอายุ">
                    <DateField
                      value={editingItem.initialExpiryDate}
                      onChange={(v) => setEditingItem({ ...editingItem, initialExpiryDate: v })}
                    />
                  </Field>
                </div>
              </div>
            )}

            {editingItem.id && (
              <p className="mb-3 rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2 text-xs text-neutral-500">
                จำนวนสต็อก ล็อต และสถานที่จัดเก็บ จัดการผ่านปุ่ม รับเข้า/เบิกออก/โอนไปผลิต/คืนจากผลิต ในตารางสินค้า ไม่ได้แก้ตรงนี้
              </p>
            )}

            <div className="grid grid-cols-2 gap-3">
              <Field label="หน่วยนับ">
                <select
                  className={inputCls}
                  value={editingItem.unit}
                  onChange={(e) => setEditingItem({ ...editingItem, unit: e.target.value })}
                >
                  {UNIT_OPTIONS.map((u) => (
                    <option key={u} value={u}>{u}</option>
                  ))}
                </select>
              </Field>
              <Field label="แจ้งเตือนเมื่อต่ำกว่า">
                <input
                  type="number"
                  min="0"
                  className={inputCls}
                  value={editingItem.lowStock}
                  onChange={(e) => setEditingItem({ ...editingItem, lowStock: e.target.value })}
                />
              </Field>
            </div>
            <Field label="ราคาต่อหน่วย (บาท)">
              <input
                type="number"
                min="0"
                step="0.01"
                className={inputCls}
                value={editingItem.price}
                onChange={(e) => setEditingItem({ ...editingItem, price: e.target.value })}
              />
            </Field>
          </div>
        </Modal>
      )}

      {/* Stock movement modal */}
      {movementItem && (
        <Modal
          title={
            movementType === "in"
              ? "รับสินค้าเข้าคลัง"
              : movementType === "out"
              ? "เบิกสินค้าออกจากคลัง"
              : movementType === "to_prod"
              ? "โอนสินค้าไปฝ่ายผลิต"
              : movementType === "from_prod"
              ? "รับคืนสินค้าจากฝ่ายผลิต"
              : "ตัดของเสีย/สูญหาย"
          }
          onClose={() => setMovementItem(null)}
          footer={
            <div className="flex justify-end gap-2">
              <button type="button" onClick={() => setMovementItem(null)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={submitMovement}
                disabled={
                  !Number(movementAmount) ||
                  Number(movementAmount) <= 0 ||
                  (movementType === "out" && !movementRequestedBy.trim()) ||
                  ((movementType === "out" || movementType === "waste") && Number(movementAmount) > warehouseQty(movementItem)) ||
                  (movementType === "to_prod" && Number(movementAmount) > qtyAtLocation(movementItem, movementLocation)) ||
                  (movementType === "from_prod" && Number(movementAmount) > qtyAtLocation(movementItem, movementDept))
                }
                className={`rounded-md px-4 py-2 text-sm font-bold disabled:cursor-not-allowed disabled:opacity-40 ${
                  movementType === "in"
                    ? "bg-emerald-500 text-neutral-950 hover:bg-emerald-400"
                    : movementType === "out"
                    ? "bg-red-500 text-neutral-950 hover:bg-red-400"
                    : movementType === "to_prod"
                    ? "bg-indigo-500 text-neutral-950 hover:bg-indigo-400"
                    : movementType === "from_prod"
                    ? "bg-teal-500 text-neutral-950 hover:bg-teal-400"
                    : "bg-rose-600 text-neutral-950 hover:bg-rose-500"
                }`}
              >
                ยืนยัน
                {movementType === "in"
                  ? "รับเข้า"
                  : movementType === "out"
                  ? "เบิกออก"
                  : movementType === "to_prod"
                  ? "โอนไปผลิต"
                  : movementType === "from_prod"
                  ? "รับคืน"
                  : "ตัดของเสีย"}
              </button>
            </div>
          }
        >
          <div>
            <div className="mb-3 rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2">
              <p className="font-semibold text-neutral-100">{movementItem.name}</p>
              <p className="font-mono text-xs text-neutral-500">
                {movementItem.sku} · คงเหลือรวม {formatNum(totalQty(movementItem))} {movementItem.unit}
                {" "}(คลัง {formatNum(warehouseQty(movementItem))} / ฝ่ายผลิต {formatNum(productionQty(movementItem))})
              </p>
              {locationBreakdown(movementItem).length > 0 && (
                <div className="mt-1.5 flex flex-wrap gap-1">
                  {locationBreakdown(movementItem).map(([loc, q]) => (
                    <LocationTag key={loc || "none"} location={loc} qty={q} unit={movementItem.unit} />
                  ))}
                </div>
              )}
            </div>

            <Field
              label={`จำนวนที่ ${
                movementType === "in"
                  ? "รับเข้า"
                  : movementType === "out"
                  ? "เบิกออก"
                  : movementType === "to_prod"
                  ? "โอนไปผลิต"
                  : movementType === "from_prod"
                  ? "รับคืน"
                  : "ตัดออก"
              } (${movementType === "in" && movementEntryUnit === "pack" && movementItem.packUnit ? movementItem.packUnit : movementItem.unit})`}
            >
              <input
                type="number"
                min="1"
                autoFocus
                className={inputCls}
                placeholder="0"
                value={movementAmount}
                onChange={(e) => setMovementAmount(e.target.value)}
              />
              {movementType === "in" && movementEntryUnit === "pack" && movementItem.packUnit && Number(movementAmount) > 0 && (
                <span className="mt-1 block text-xs text-sky-400">
                  = {formatNum(Number(movementAmount) * Number(movementItem.packFactor || 0))} {movementItem.unit}
                </span>
              )}
              {(movementType === "out" || movementType === "waste") && Number(movementAmount) > warehouseQty(movementItem) && (
                <span className="mt-1 block text-xs text-red-400">
                  มีในคลังเพียง {formatNum(warehouseQty(movementItem))} {movementItem.unit} {movementType === "waste" ? "ตัด" : "เบิก"}เกินจำนวนนี้ไม่ได้
                </span>
              )}
              {movementType === "to_prod" && Number(movementAmount) > qtyAtLocation(movementItem, movementLocation) && (
                <span className="mt-1 block text-xs text-red-400">
                  ที่ {locationLabel(movementLocation)} มีเพียง {formatNum(qtyAtLocation(movementItem, movementLocation))} {movementItem.unit}
                </span>
              )}
              {movementType === "from_prod" && Number(movementAmount) > qtyAtLocation(movementItem, movementDept) && (
                <span className="mt-1 block text-xs text-red-400">
                  ที่แผนกผลิต {locationLabel(movementDept)} มีเพียง {formatNum(qtyAtLocation(movementItem, movementDept))} {movementItem.unit}
                </span>
              )}
            </Field>

            {movementType === "in" && movementItem.packUnit && Number(movementItem.packFactor) > 0 && (
              <Field label="กรอกจำนวนเป็นหน่วย">
                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={() => setMovementEntryUnit("base")}
                    className={`flex-1 rounded-md border px-3 py-2 text-sm font-semibold ${
                      movementEntryUnit === "base" ? "border-amber-500 bg-amber-500/20 text-amber-300" : "border-neutral-700 bg-neutral-800 text-neutral-400"
                    }`}
                  >
                    {movementItem.unit}
                  </button>
                  <button
                    type="button"
                    onClick={() => setMovementEntryUnit("pack")}
                    className={`flex-1 rounded-md border px-3 py-2 text-sm font-semibold ${
                      movementEntryUnit === "pack" ? "border-amber-500 bg-amber-500/20 text-amber-300" : "border-neutral-700 bg-neutral-800 text-neutral-400"
                    }`}
                  >
                    {movementItem.packUnit} (1 = {formatNum(movementItem.packFactor)} {movementItem.unit})
                  </button>
                </div>
              </Field>
            )}

            {movementType === "in" && (
              <>
                <p className="mb-2 -mt-1 text-xs text-neutral-500">
                  ถ้าวันที่ผลิตหรือสถานที่ต่างจากล็อตเดิม ระบบจะเปิดล็อตใหม่แยกให้อัตโนมัติ
                </p>
                <Field label="สถานที่จัดเก็บ">
                  <select className={inputCls} value={movementLocation} onChange={(e) => setMovementLocation(e.target.value)}>
                    {LOCATIONS.filter((loc) => !isProdLocation(loc)).map((loc) => (
                      <option key={loc} value={loc}>{loc}</option>
                    ))}
                  </select>
                </Field>
                <div className="grid grid-cols-2 gap-3">
                  <Field label="วันที่ผลิต (ของล็อตนี้)">
                    <DateField value={movementProductionDate} onChange={setMovementProductionDate} />
                  </Field>
                  <Field label="วันหมดอายุ (ของล็อตนี้)">
                    <DateField value={movementExpiryDate} onChange={setMovementExpiryDate} />
                  </Field>
                </div>
                <Field label="วันที่รับเข้า">
                  <DateField value={movementReceivedDate} onChange={setMovementReceivedDate} />
                </Field>
              </>
            )}

            {movementType === "out" && (
              <>
                <p className="mb-2 -mt-1 text-xs text-neutral-500">
                  ระบบจะตัดสต็อกจากล็อตที่ใกล้หมดอายุที่สุดก่อนโดยอัตโนมัติ (FEFO) เฉพาะสต็อกในคลัง
                </p>
                <Field label="ชื่อผู้เบิก">
                  <input
                    className={inputCls}
                    placeholder="ชื่อพนักงานที่เบิกสินค้า"
                    value={movementRequestedBy}
                    onChange={(e) => setMovementRequestedBy(e.target.value)}
                  />
                </Field>
              </>
            )}

            {movementType === "to_prod" && (
              <>
                <p className="mb-2 -mt-1 text-xs text-neutral-500">
                  ตัดสต็อกแบบ FEFO จากสถานที่ที่เลือก แล้วย้ายไปกองไว้ที่แผนกผลิตที่เลือก
                </p>
                <div className="grid grid-cols-2 gap-3">
                  <Field label="จากสถานที่">
                    <select className={inputCls} value={movementLocation} onChange={(e) => setMovementLocation(e.target.value)}>
                      {LOCATIONS.filter((loc) => !isProdLocation(loc)).map((loc) => (
                        <option key={loc} value={loc}>
                          {loc} (มี {formatNum(qtyAtLocation(movementItem, loc))} {movementItem.unit})
                        </option>
                      ))}
                    </select>
                  </Field>
                  <Field label="ไปยังแผนกผลิต">
                    <select className={inputCls} value={movementDept} onChange={(e) => setMovementDept(e.target.value)}>
                      {PROD_LOCATIONS.map((d) => (
                        <option key={d} value={d}>{d}</option>
                      ))}
                    </select>
                  </Field>
                </div>
                <Field label="ผู้รับไปผลิต (ไม่บังคับ)">
                  <input
                    className={inputCls}
                    placeholder="ชื่อผู้รับสินค้าไปผลิต"
                    value={movementRequestedBy}
                    onChange={(e) => setMovementRequestedBy(e.target.value)}
                  />
                </Field>
              </>
            )}

            {movementType === "from_prod" && (
              <>
                <p className="mb-2 -mt-1 text-xs text-neutral-500">
                  ตัดสต็อกแบบ FEFO จากของที่ค้างอยู่ที่แผนกผลิตที่เลือก แล้วนำกลับเข้าสถานที่ที่เลือก
                </p>
                <div className="grid grid-cols-2 gap-3">
                  <Field label="จากแผนกผลิต">
                    <select className={inputCls} value={movementDept} onChange={(e) => setMovementDept(e.target.value)}>
                      {PROD_LOCATIONS.map((d) => (
                        <option key={d} value={d}>
                          {d} (มี {formatNum(qtyAtLocation(movementItem, d))} {movementItem.unit})
                        </option>
                      ))}
                    </select>
                  </Field>
                  <Field label="ไปยังสถานที่">
                    <select className={inputCls} value={movementLocation} onChange={(e) => setMovementLocation(e.target.value)}>
                      {LOCATIONS.filter((loc) => !isProdLocation(loc)).map((loc) => (
                        <option key={loc} value={loc}>{loc}</option>
                      ))}
                    </select>
                  </Field>
                </div>
                <Field label="ผู้คืนสินค้า (ไม่บังคับ)">
                  <input
                    className={inputCls}
                    placeholder="ชื่อพนักงานฝ่ายผลิตที่คืนสินค้า"
                    value={movementRequestedBy}
                    onChange={(e) => setMovementRequestedBy(e.target.value)}
                  />
                </Field>
              </>
            )}

            {movementType === "waste" && (
              <>
                <p className="mb-2 -mt-1 text-xs text-neutral-500">
                  ใช้เมื่อสินค้าเสียหาย หมดอายุ หรือสูญหาย ตัดออกจากสต็อกถาวรและแยกจากยอดเบิกใช้งานปกติ เพื่อให้รายงานความเสียหายแม่นยำ
                </p>
                <Field label="เหตุผล">
                  <select className={inputCls} value={movementReason} onChange={(e) => setMovementReason(e.target.value)}>
                    {WASTE_REASONS.map((r) => (
                      <option key={r} value={r}>{r}</option>
                    ))}
                  </select>
                </Field>
              </>
            )}

            <Field label="หมายเหตุ (ไม่บังคับ)">
              <input
                className={inputCls}
                placeholder="เช่น รับจากซัพพลายเออร์ / เบิกใช้งานไซต์ A"
                value={movementNote}
                onChange={(e) => setMovementNote(e.target.value)}
              />
            </Field>
          </div>
        </Modal>
      )}

      {/* Material-issue slip (ใบตัดวัตถุดิบ) form */}
      {showJobForm && jobForm && (
        <Modal
          title={`ตัดวัตถุดิบออกจากแผนกผลิต ${locationLabel(jobForm.dept)}`}
          onClose={() => setShowJobForm(false)}
          footer={
            <div className="flex justify-end gap-2">
              <button type="button" onClick={() => setShowJobForm(false)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={submitJobForm}
                className="rounded-md bg-indigo-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-indigo-400"
              >
                ตัดวัตถุดิบและออกใบตัด
              </button>
            </div>
          }
        >
          <p className="mb-3 -mt-1 text-xs text-neutral-500">
            ใช้เมื่อผลิตสินค้าเสร็จแล้ว เพื่อตัดวัตถุดิบที่เคยโอนไปแผนกผลิตออกจากสต็อกอย่างถาวร (ไม่กลับเข้าคลัง) พร้อมออกใบตัดวัตถุดิบให้
          </p>
          <div className="grid grid-cols-2 gap-3">
            <Field label="แผนกผลิต">
              <select
                className={inputCls}
                value={jobForm.dept}
                onChange={(e) => setJobForm({ ...jobForm, dept: e.target.value, lines: [{ id: uid(), itemId: "", qty: "" }] })}
              >
                {PROD_LOCATIONS.map((d) => (
                  <option key={d} value={d}>{d}</option>
                ))}
              </select>
            </Field>
            <Field label="วันที่ผลิต">
              <DateField value={jobForm.jobDate} onChange={(v) => setJobForm({ ...jobForm, jobDate: v })} />
            </Field>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <Field label="ผลิตสินค้า / งานผลิต">
              <input
                className={inputCls}
                placeholder="เช่น ผงชีส"
                value={jobForm.jobName}
                onChange={(e) => setJobForm({ ...jobForm, jobName: e.target.value })}
              />
            </Field>
            <Field label="สูตรที่ใช้ (ไม่บังคับ)">
              <input
                className={inputCls}
                placeholder="เช่น สูตร A / SSN-001"
                value={jobForm.recipe}
                onChange={(e) => setJobForm({ ...jobForm, recipe: e.target.value })}
              />
            </Field>
          </div>

          <p className="mb-1 text-xs font-semibold uppercase tracking-wide text-neutral-400">
            วัตถุดิบที่ตัด (เฉพาะที่มีสต็อกอยู่ที่แผนก {locationLabel(jobForm.dept)})
          </p>
          <div className="space-y-2">
            {jobForm.lines.map((line) => {
              const item = items.find((i) => i.id === line.itemId);
              const available = item ? qtyAtLocation(item, jobForm.dept) : 0;
              return (
                <div key={line.id} className="rounded-md border border-neutral-800 bg-neutral-950 p-2">
                  <div className="flex items-start gap-2">
                    <select
                      className={`${inputCls} flex-1`}
                      value={line.itemId}
                      onChange={(e) => updateJobLine(line.id, { itemId: e.target.value, qty: "" })}
                    >
                      <option value="">เลือกวัตถุดิบ...</option>
                      {items
                        .filter((i) => qtyAtLocation(i, jobForm.dept) > 0)
                        .map((i) => (
                          <option key={i.id} value={i.id}>
                            {i.name} ({i.sku}) — มี {formatNum(qtyAtLocation(i, jobForm.dept))} {i.unit} ที่แผนกนี้
                          </option>
                        ))}
                    </select>
                    <button
                      onClick={() => removeJobLine(line.id)}
                      className="rounded p-2 text-neutral-500 hover:bg-neutral-800 hover:text-red-400"
                      title="ลบรายการนี้"
                    >
                      <X size={15} />
                    </button>
                  </div>
                  {line.itemId && (
                    <div className="mt-2 flex items-center gap-2">
                      <input
                        type="number"
                        min="1"
                        className={`${inputCls} flex-1`}
                        placeholder={`จำนวน (${item?.unit || ""})`}
                        value={line.qty}
                        onChange={(e) => updateJobLine(line.id, { qty: e.target.value })}
                      />
                      <button
                        type="button"
                        onClick={() => updateJobLine(line.id, { qty: String(available) })}
                        className="shrink-0 rounded-md border border-indigo-600/50 px-2 py-2 text-xs font-semibold text-indigo-400 hover:bg-indigo-500/10"
                      >
                        ตัดทั้งหมด ({formatNum(available)})
                      </button>
                    </div>
                  )}
                  {line.itemId && Number(line.qty) > available && (
                    <p className="mt-1 text-xs text-red-400">มีที่แผนก {locationLabel(jobForm.dept)} เพียง {formatNum(available)} {item?.unit}</p>
                  )}
                </div>
              );
            })}
          </div>
          <button
            type="button"
            onClick={addJobLine}
            className="mt-2 flex items-center gap-1 rounded-md border border-neutral-700 px-3 py-1.5 text-xs font-semibold text-neutral-300 hover:bg-neutral-800"
          >
            <Plus size={14} /> เพิ่มวัตถุดิบ
          </button>
        </Modal>
      )}

      {/* Material-issue slip (ใบตัดวัตถุดิบ) receipt */}
      {jobReceipt && (
        <Modal title="ใบตัดวัตถุดิบ" onClose={() => setJobReceipt(null)}>
          <div className="printable-content">
            {settings.printHeaderText && (
              <p className="mb-2 border-b border-neutral-800 pb-2 text-sm font-bold text-neutral-100">{settings.printHeaderText}</p>
            )}
            <div className="mb-3 rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2">
              <p className="font-semibold text-neutral-100">ผลิต: {jobReceipt.jobName}</p>
              {jobReceipt.recipe && <p className="text-xs text-indigo-300">สูตร: {jobReceipt.recipe}</p>}
              <p className="font-mono text-xs text-neutral-500">
                แผนกผลิต {locationLabel(jobReceipt.dept)} · วันที่ผลิต {formatDateOnly(jobReceipt.jobDate)}
              </p>
            </div>
            <table className="w-full text-xs">
              <thead>
                <tr className="text-left text-neutral-500">
                  <th className="px-2 py-1">วัตถุดิบ</th>
                  <th className="px-2 py-1 text-right">ตัดออก</th>
                  <th className="px-2 py-1 text-right">คงเหลือที่แผนกนี้</th>
                </tr>
              </thead>
              <tbody>
                {jobReceipt.lines.map((l) => (
                  <tr key={l.itemId} className="border-t border-neutral-800">
                    <td className="px-2 py-1.5">
                      <p className="font-semibold text-neutral-200">{l.itemName}</p>
                      <p className="font-mono text-neutral-500">{l.sku}</p>
                    </td>
                    <td className="px-2 py-1.5 text-right font-mono text-red-400">
                      -{formatNum(l.qty)} {l.unit}
                    </td>
                    <td className="px-2 py-1.5 text-right font-mono text-neutral-300">
                      {l.deleted ? "0 (ลบสินค้าแล้ว)" : `${formatNum(l.remaining)} ${l.unit}`}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="no-print mt-4 flex justify-end gap-2">
            <button
              onClick={() => window.print()}
              className="flex items-center gap-1.5 rounded-md border border-neutral-700 px-4 py-2 text-sm font-bold text-neutral-300 hover:bg-neutral-800"
            >
              <Printer size={15} /> พิมพ์
            </button>
            <button
              onClick={() => setJobReceipt(null)}
              className="rounded-md bg-amber-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-amber-400"
            >
              ปิด
            </button>
          </div>
        </Modal>
      )}

      {/* Finished goods QC check / receiving form (เอกสารเช็คสินค้าสำเร็จรูป) */}
      {showFgForm && fgForm && (
        <Modal
          title="ตรวจรับสินค้าสำเร็จรูป"
          onClose={() => setShowFgForm(false)}
          footer={
            <div className="flex justify-end gap-2">
              <button type="button" onClick={() => setShowFgForm(false)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={submitFgForm}
                className="rounded-md bg-teal-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-teal-400"
              >
                บันทึกผลตรวจและออกเอกสาร
              </button>
            </div>
          }
        >
          <p className="mb-3 -mt-1 text-xs text-neutral-500">
            ใช้บันทึกผลตรวจสอบสินค้าสำเร็จรูปหลังผลิตเสร็จ ถ้าผลตรวจ "ผ่าน" ระบบจะรับสินค้าเข้าสต็อกที่ FG ให้อัตโนมัติ
          </p>

          <Field label="สินค้าสำเร็จรูป">
            <select
              className={inputCls}
              value={fgForm.itemId}
              onChange={(e) => setFgForm({ ...fgForm, itemId: e.target.value })}
            >
              <option value="">-- สินค้าใหม่ (พิมพ์ชื่อ/SKU ด้านล่าง) --</option>
              {items
                .slice()
                .sort((a, b) => a.name.localeCompare(b.name, "th"))
                .map((i) => (
                  <option key={i.id} value={i.id}>{i.name} ({i.sku})</option>
                ))}
            </select>
          </Field>

          {!fgForm.itemId && (
            <div className="grid grid-cols-3 gap-3">
              <div className="col-span-2">
                <Field label="ชื่อสินค้าใหม่">
                  <input
                    className={inputCls}
                    placeholder="เช่น ผงชีส"
                    value={fgForm.newName}
                    onChange={(e) => setFgForm({ ...fgForm, newName: e.target.value })}
                  />
                </Field>
              </div>
              <Field label="SKU">
                <input
                  className={inputCls}
                  value={fgForm.newSku}
                  onChange={(e) => setFgForm({ ...fgForm, newSku: e.target.value })}
                />
              </Field>
            </div>
          )}
          {!fgForm.itemId && (
            <Field label="หน่วยนับ">
              <select className={inputCls} value={fgForm.newUnit} onChange={(e) => setFgForm({ ...fgForm, newUnit: e.target.value })}>
                {UNIT_OPTIONS.map((u) => (
                  <option key={u} value={u}>{u}</option>
                ))}
              </select>
            </Field>
          )}

          <div className="grid grid-cols-2 gap-3">
            <Field label={`จำนวนที่ผลิตได้${fgForm.itemId ? ` (${items.find((i) => i.id === fgForm.itemId)?.unit || ""})` : ""}`}>
              <input
                type="number"
                min="1"
                className={inputCls}
                value={fgForm.qty}
                onChange={(e) => setFgForm({ ...fgForm, qty: e.target.value })}
              />
            </Field>
            <Field label="วันที่ตรวจสอบ/ผลิต">
              <DateField value={fgForm.checkDate} onChange={(v) => setFgForm({ ...fgForm, checkDate: v })} />
            </Field>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <Field label="วันหมดอายุ (ถ้ามี)">
              <DateField value={fgForm.expiryDate} onChange={(v) => setFgForm({ ...fgForm, expiryDate: v })} />
            </Field>
            <Field label="แผนกผลิตที่มา (ไม่บังคับ)">
              <select className={inputCls} value={fgForm.dept} onChange={(e) => setFgForm({ ...fgForm, dept: e.target.value })}>
                {PROD_LOCATIONS.map((d) => (
                  <option key={d} value={d}>{d}</option>
                ))}
              </select>
            </Field>
          </div>

          {jobs.length > 0 && (
            <Field label="อ้างอิงใบตัดวัตถุดิบ (ไม่บังคับ)">
              <select className={inputCls} value={fgForm.jobId} onChange={(e) => setFgForm({ ...fgForm, jobId: e.target.value })}>
                <option value="">-- ไม่ระบุ --</option>
                {jobs.slice(0, 30).map((j) => (
                  <option key={j.id} value={j.id}>
                    {formatDateOnly(j.jobDate)} · {j.jobName}{j.recipe ? ` (สูตร ${j.recipe})` : ""}
                  </option>
                ))}
              </select>
            </Field>
          )}

          <Field label="ผู้ตรวจสอบ">
            <input
              className={inputCls}
              placeholder="ชื่อผู้ตรวจสอบคุณภาพ"
              value={fgForm.inspector}
              onChange={(e) => setFgForm({ ...fgForm, inspector: e.target.value })}
            />
          </Field>

          <p className="mb-1 text-xs font-semibold uppercase tracking-wide text-neutral-400">รายการตรวจสอบคุณภาพ (QC)</p>
          <div className="mb-3 space-y-1.5 rounded-md border border-neutral-800 bg-neutral-950 p-2">
            <label className="flex items-center gap-2 text-sm text-neutral-200">
              <input
                type="checkbox"
                checked={fgForm.checkAppearance}
                onChange={(e) => setFgForm({ ...fgForm, checkAppearance: e.target.checked })}
              />
              ลักษณะ สี กลิ่น ปกติ
            </label>
            <label className="flex items-center gap-2 text-sm text-neutral-200">
              <input
                type="checkbox"
                checked={fgForm.checkWeight}
                onChange={(e) => setFgForm({ ...fgForm, checkWeight: e.target.checked })}
              />
              น้ำหนัก/ปริมาณตรงตามสเปค
            </label>
            <label className="flex items-center gap-2 text-sm text-neutral-200">
              <input
                type="checkbox"
                checked={fgForm.checkPackaging}
                onChange={(e) => setFgForm({ ...fgForm, checkPackaging: e.target.checked })}
              />
              บรรจุภัณฑ์เรียบร้อย
            </label>
          </div>

          <Field label="ผลการตรวจสอบ">
            <div className="flex gap-2">
              <button
                type="button"
                onClick={() => setFgForm({ ...fgForm, result: "pass" })}
                className={`flex-1 rounded-md border px-3 py-2 text-sm font-bold ${
                  fgForm.result === "pass"
                    ? "border-emerald-500 bg-emerald-500/20 text-emerald-300"
                    : "border-neutral-700 bg-neutral-800 text-neutral-400"
                }`}
              >
                ผ่าน (รับเข้า FG)
              </button>
              <button
                type="button"
                onClick={() => setFgForm({ ...fgForm, result: "fail" })}
                className={`flex-1 rounded-md border px-3 py-2 text-sm font-bold ${
                  fgForm.result === "fail"
                    ? "border-red-500 bg-red-500/20 text-red-300"
                    : "border-neutral-700 bg-neutral-800 text-neutral-400"
                }`}
              >
                ไม่ผ่าน (ไม่รับเข้า)
              </button>
            </div>
          </Field>

          <Field label="หมายเหตุ (ไม่บังคับ)">
            <input
              className={inputCls}
              placeholder="เช่น เหตุผลที่ไม่ผ่าน หรือข้อสังเกตอื่นๆ"
              value={fgForm.note}
              onChange={(e) => setFgForm({ ...fgForm, note: e.target.value })}
            />
          </Field>
        </Modal>
      )}

      {/* Finished goods QC check receipt */}
      {fgReceipt && (
        <Modal title="ใบตรวจรับสินค้าสำเร็จรูป" onClose={() => setFgReceipt(null)}>
          <div className="printable-content">
            {settings.printHeaderText && (
              <p className="mb-2 border-b border-neutral-800 pb-2 text-sm font-bold text-neutral-100">{settings.printHeaderText}</p>
            )}
            <div className="mb-3 rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2">
              <p className="font-semibold text-neutral-100">{fgReceipt.itemName} <span className="font-mono text-xs text-neutral-500">({fgReceipt.sku})</span></p>
              <p className="font-mono text-xs text-neutral-500">
                จำนวน {formatNum(fgReceipt.qty)} {fgReceipt.unit} · วันที่ {formatDateOnly(fgReceipt.checkDate)}
                {fgReceipt.dept ? ` · จากแผนก ${locationLabel(fgReceipt.dept)}` : ""}
              </p>
              <p className="text-xs text-neutral-500">ผู้ตรวจสอบ: {fgReceipt.inspector}</p>
            </div>
            <ul className="mb-3 space-y-1 text-xs text-neutral-300">
              <li>{fgReceipt.checkAppearance ? "✓" : "✗"} ลักษณะ สี กลิ่น ปกติ</li>
              <li>{fgReceipt.checkWeight ? "✓" : "✗"} น้ำหนัก/ปริมาณตรงตามสเปค</li>
              <li>{fgReceipt.checkPackaging ? "✓" : "✗"} บรรจุภัณฑ์เรียบร้อย</li>
            </ul>
            <p
              className={`mb-2 rounded-md px-3 py-2 text-center text-sm font-bold ${
                fgReceipt.result === "pass" ? "bg-emerald-500/20 text-emerald-300" : "bg-red-500/20 text-red-300"
              }`}
            >
              ผลการตรวจสอบ: {fgReceipt.result === "pass" ? "ผ่าน — รับเข้าสต็อก FG แล้ว" : "ไม่ผ่าน — ไม่ได้รับเข้าสต็อก"}
            </p>
            {fgReceipt.note && <p className="mb-2 text-xs text-neutral-400">หมายเหตุ: {fgReceipt.note}</p>}
          </div>
          <div className="no-print mt-2 flex justify-end gap-2">
            <button
              onClick={() => window.print()}
              className="flex items-center gap-1.5 rounded-md border border-neutral-700 px-4 py-2 text-sm font-bold text-neutral-300 hover:bg-neutral-800"
            >
              <Printer size={15} /> พิมพ์
            </button>
            <button
              onClick={() => setFgReceipt(null)}
              className="rounded-md bg-amber-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-amber-400"
            >
              ปิด
            </button>
          </div>
        </Modal>
      )}

      {/* Purchase order (PO) create form */}
      {showPoForm && poForm && (
        <Modal
          title="สร้างใบสั่งซื้อใหม่"
          onClose={() => setShowPoForm(false)}
          footer={
            <div className="flex justify-end gap-2">
              <button type="button" onClick={() => setShowPoForm(false)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={submitPoForm}
                disabled={!poForm.supplier.trim() || !poForm.lines.some((l) => (l.itemId || (l.newName.trim() && l.newSku.trim())) && Number(l.qty) > 0)}
                className="rounded-md bg-orange-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-orange-400 disabled:cursor-not-allowed disabled:opacity-40"
              >
                สร้างใบสั่งซื้อ
              </button>
            </div>
          }
        >
          <div className="grid grid-cols-2 gap-3">
            <Field label="ซัพพลายเออร์">
              <input
                className={inputCls}
                placeholder="ชื่อผู้ขาย"
                value={poForm.supplier}
                onChange={(e) => setPoForm({ ...poForm, supplier: e.target.value })}
              />
            </Field>
            <Field label="วันที่สั่งซื้อ">
              <DateField value={poForm.orderDate} onChange={(v) => setPoForm({ ...poForm, orderDate: v })} />
            </Field>
          </div>
          <Field label="วันที่คาดว่าจะได้รับ (ไม่บังคับ)">
            <DateField value={poForm.expectedDate} onChange={(v) => setPoForm({ ...poForm, expectedDate: v })} />
          </Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="เลขที่ใบกำกับภาษี (ไม่บังคับ)">
              <input
                className={inputCls}
                placeholder="เช่น INV-00123"
                value={poForm.taxInvoiceNo}
                onChange={(e) => setPoForm({ ...poForm, taxInvoiceNo: e.target.value })}
              />
            </Field>
            <Field label="ที่อยู่บริษัทผู้ขาย (ไม่บังคับ)">
              <input
                className={inputCls}
                placeholder="ที่อยู่สำหรับออกเอกสาร"
                value={poForm.supplierAddress}
                onChange={(e) => setPoForm({ ...poForm, supplierAddress: e.target.value })}
              />
            </Field>
          </div>

          <p className="mb-1 text-xs font-semibold uppercase tracking-wide text-neutral-400">รายการสินค้าที่สั่ง</p>
          <div className="space-y-2">
            {poForm.lines.map((line) => (
              <div key={line.id} className="rounded-md border border-neutral-800 bg-neutral-950 p-2">
                <div className="flex items-start gap-2">
                  <select
                    className={`${inputCls} flex-1`}
                    value={line.itemId}
                    onChange={(e) => updatePoLine(line.id, { itemId: e.target.value })}
                  >
                    <option value="">-- สินค้าใหม่ (พิมพ์ชื่อ/SKU ด้านล่าง) --</option>
                    {items
                      .slice()
                      .sort((a, b) => a.name.localeCompare(b.name, "th"))
                      .map((i) => (
                        <option key={i.id} value={i.id}>{i.name} ({i.sku})</option>
                      ))}
                  </select>
                  <button onClick={() => removePoLine(line.id)} className="rounded p-2 text-neutral-500 hover:bg-neutral-800 hover:text-red-400">
                    <X size={15} />
                  </button>
                </div>
                {!line.itemId && (
                  <div className="mt-2 grid grid-cols-2 gap-2">
                    <input
                      className={inputCls}
                      placeholder="ชื่อสินค้าใหม่"
                      value={line.newName}
                      onChange={(e) => updatePoLine(line.id, { newName: e.target.value })}
                    />
                    <input
                      className={inputCls}
                      placeholder="SKU"
                      value={line.newSku}
                      onChange={(e) => updatePoLine(line.id, { newSku: e.target.value })}
                    />
                  </div>
                )}
                <div className="mt-2 grid grid-cols-2 gap-2">
                  <input
                    type="number"
                    min="1"
                    className={inputCls}
                    placeholder="จำนวนที่สั่ง"
                    value={line.qty}
                    onChange={(e) => updatePoLine(line.id, { qty: e.target.value })}
                  />
                  <input
                    type="number"
                    min="0"
                    className={inputCls}
                    placeholder="ราคาต่อหน่วย (บาท)"
                    value={line.unitPrice}
                    onChange={(e) => updatePoLine(line.id, { unitPrice: e.target.value })}
                  />
                </div>
                <div className="mt-2 grid grid-cols-2 gap-2">
                  <div>
                    <p className="mb-1 text-[10px] font-semibold uppercase tracking-wide text-neutral-500">วันที่ผลิต (ไม่บังคับ)</p>
                    <DateField value={line.productionDate} onChange={(v) => updatePoLine(line.id, { productionDate: v })} />
                  </div>
                  <div>
                    <p className="mb-1 text-[10px] font-semibold uppercase tracking-wide text-neutral-500">วันหมดอายุ (ไม่บังคับ)</p>
                    <DateField value={line.expiryDate} onChange={(v) => updatePoLine(line.id, { expiryDate: v })} />
                  </div>
                </div>
              </div>
            ))}
          </div>
          <button
            type="button"
            onClick={addPoLine}
            className="mt-2 flex items-center gap-1 rounded-md border border-neutral-700 px-3 py-1.5 text-xs font-semibold text-neutral-300 hover:bg-neutral-800"
          >
            <Plus size={14} /> เพิ่มรายการสินค้า
          </button>

          <Field label="หมายเหตุ (ไม่บังคับ)">
            <input className={inputCls} value={poForm.note} onChange={(e) => setPoForm({ ...poForm, note: e.target.value })} />
          </Field>
        </Modal>
      )}

      {/* Receive PO modal */}
      {receivingPo && (
        <Modal
          title={`รับของตามใบสั่งซื้อ ${receivingPo.poNumber}`}
          onClose={() => setReceivingPo(null)}
          footer={
            <div className="flex justify-end gap-2">
              <button type="button" onClick={() => setReceivingPo(null)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={confirmReceivePo}
                className="rounded-md bg-emerald-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-emerald-400"
              >
                ยืนยันรับของเข้าคลัง
              </button>
            </div>
          }
        >
          <p className="mb-3 -mt-1 text-xs text-neutral-500">
            สินค้าทุกรายการในใบสั่งซื้อนี้จะถูกรับเข้าคลังที่สถานที่เดียวกัน
          </p>
          <ul className="mb-3 space-y-1 text-xs text-neutral-300">
            {receivingPo.lines.map((l) => (
              <li key={l.id} className="flex items-center justify-between rounded border border-neutral-800 bg-neutral-950 px-2 py-1">
                <span>{l.itemName} ({l.sku})</span>
                <span className="font-mono">{formatNum(l.qty)} {l.unit}</span>
              </li>
            ))}
          </ul>
          <Field label="รับเข้าที่สถานที่">
            <select className={inputCls} value={receiveLocation} onChange={(e) => setReceiveLocation(e.target.value)}>
              {LOCATIONS.filter((loc) => !isProdLocation(loc)).map((loc) => (
                <option key={loc} value={loc}>{loc}</option>
              ))}
            </select>
          </Field>
          <Field label="วันที่รับเข้า">
            <DateField value={receiveDateInput} onChange={setReceiveDateInput} />
          </Field>
        </Modal>
      )}

      {/* Journal entry form (บันทึกรายวัน) */}
      {showJournalForm && journalForm && (
        <Modal
          title="บันทึกรายการบัญชีใหม่"
          onClose={() => setShowJournalForm(false)}
          footer={
            <div className="flex justify-end gap-2">
              <button type="button" onClick={() => setShowJournalForm(false)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={submitJournalForm}
                className="rounded-md bg-yellow-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-yellow-400"
              >
                บันทึกรายการ
              </button>
            </div>
          }
        >
          <div className="grid grid-cols-2 gap-3">
            <Field label="วันที่">
              <DateField value={journalForm.date} onChange={(v) => setJournalForm({ ...journalForm, date: v })} />
            </Field>
            <Field label="คำอธิบายรายการ">
              <input
                className={inputCls}
                placeholder="เช่น จ่ายค่าน้ำค่าไฟ"
                value={journalForm.description}
                onChange={(e) => setJournalForm({ ...journalForm, description: e.target.value })}
              />
            </Field>
          </div>

          <p className="mb-1 text-xs font-semibold uppercase tracking-wide text-neutral-400">รายการเดบิต/เครดิต (ยอดรวมต้องเท่ากัน)</p>
          <div className="space-y-2">
            {journalForm.lines.map((line) => (
              <div key={line.id} className="flex items-start gap-2 rounded-md border border-neutral-800 bg-neutral-950 p-2">
                <select
                  className={`${inputCls} flex-1`}
                  value={line.accountId}
                  onChange={(e) => updateJournalLine(line.id, { accountId: e.target.value })}
                >
                  <option value="">เลือกบัญชี...</option>
                  {accounts.map((a) => (
                    <option key={a.id} value={a.id}>{a.code} - {a.name}</option>
                  ))}
                </select>
                <input
                  type="number"
                  min="0"
                  className={`${inputCls} w-24`}
                  placeholder="เดบิต"
                  value={line.debit}
                  onChange={(e) => updateJournalLine(line.id, { debit: e.target.value, credit: "" })}
                />
                <input
                  type="number"
                  min="0"
                  className={`${inputCls} w-24`}
                  placeholder="เครดิต"
                  value={line.credit}
                  onChange={(e) => updateJournalLine(line.id, { credit: e.target.value, debit: "" })}
                />
                <button onClick={() => removeJournalLine(line.id)} className="rounded p-2 text-neutral-500 hover:bg-neutral-800 hover:text-red-400">
                  <X size={15} />
                </button>
              </div>
            ))}
          </div>
          <button
            type="button"
            onClick={addJournalLine}
            className="mt-2 flex items-center gap-1 rounded-md border border-neutral-700 px-3 py-1.5 text-xs font-semibold text-neutral-300 hover:bg-neutral-800"
          >
            <Plus size={14} /> เพิ่มบัญชี
          </button>
          <p className="mt-2 text-xs text-neutral-500">
            รวมเดบิต: {formatBaht(journalForm.lines.reduce((s, l) => s + (Number(l.debit) || 0), 0))} · รวมเครดิต:{" "}
            {formatBaht(journalForm.lines.reduce((s, l) => s + (Number(l.credit) || 0), 0))}
          </p>
        </Modal>
      )}

      {/* Chart-of-accounts add/edit form */}
      {showAccountForm && accountForm && (
        <Modal
          title={accountForm.id ? "แก้ไขบัญชี" : "เพิ่มบัญชีใหม่"}
          onClose={() => setShowAccountForm(false)}
          footer={
            <div className="flex justify-end gap-2">
              <button type="button" onClick={() => setShowAccountForm(false)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={submitAccountForm}
                className="rounded-md bg-yellow-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-yellow-400"
              >
                บันทึก
              </button>
            </div>
          }
        >
          <Field label="รหัสบัญชี">
            <input className={inputCls} value={accountForm.code} onChange={(e) => setAccountForm({ ...accountForm, code: e.target.value })} />
          </Field>
          <Field label="ชื่อบัญชี">
            <input className={inputCls} value={accountForm.name} onChange={(e) => setAccountForm({ ...accountForm, name: e.target.value })} />
          </Field>
          <Field label="ประเภทบัญชี">
            <select className={inputCls} value={accountForm.type} onChange={(e) => setAccountForm({ ...accountForm, type: e.target.value })}>
              {ACCOUNT_TYPES.map((t) => (
                <option key={t.id} value={t.id}>{t.label}</option>
              ))}
            </select>
          </Field>
        </Modal>
      )}

      {/* PO payment form */}
      {payingPo && poPaymentForm && (
        <Modal
          title={`บันทึกจ่ายเงิน PO ${payingPo.poNumber}`}
          onClose={() => setPayingPo(null)}
          footer={
            <div className="flex justify-end gap-2">
              <button type="button" onClick={() => setPayingPo(null)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={() => recordPoPayment(payingPo, poPaymentForm)}
                className="rounded-md bg-yellow-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-yellow-400"
              >
                บันทึกการจ่ายเงิน
              </button>
            </div>
          }
        >
          <p className="mb-3 -mt-1 text-xs text-neutral-500">
            ยอดค้างชำระปัจจุบัน: {formatBaht(poBalanceDue(payingPo))}
          </p>
          <Field label="จำนวนเงินที่จ่าย (บาท)">
            <input
              type="number"
              min="0"
              autoFocus
              className={inputCls}
              value={poPaymentForm.amount}
              onChange={(e) => setPoPaymentForm({ ...poPaymentForm, amount: e.target.value })}
            />
          </Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="วันที่จ่าย">
              <DateField value={poPaymentForm.date} onChange={(v) => setPoPaymentForm({ ...poPaymentForm, date: v })} />
            </Field>
            <Field label="วิธีการจ่าย">
              <select className={inputCls} value={poPaymentForm.method} onChange={(e) => setPoPaymentForm({ ...poPaymentForm, method: e.target.value })}>
                <option>โอนเงิน</option>
                <option>เงินสด</option>
                <option>เช็ค</option>
                <option>บัตรเครดิต</option>
              </select>
            </Field>
          </div>
          <Field label="เลขที่อ้างอิง (ไม่บังคับ)">
            <input
              className={inputCls}
              placeholder="เลขที่เช็ค/เลขอ้างอิงการโอน"
              value={poPaymentForm.reference}
              onChange={(e) => setPoPaymentForm({ ...poPaymentForm, reference: e.target.value })}
            />
          </Field>
          <Field label="หมายเหตุ (ไม่บังคับ)">
            <input
              className={inputCls}
              value={poPaymentForm.note}
              onChange={(e) => setPoPaymentForm({ ...poPaymentForm, note: e.target.value })}
            />
          </Field>
        </Modal>
      )}

      {/* Insufficient-stock alert modal */}
      {alertMessage && (
        <Modal title="แจ้งเตือนสต็อกสินค้า" onClose={() => setAlertMessage("")}>
          <div className="flex items-start gap-3">
            <AlertTriangle className="mt-0.5 shrink-0 text-red-400" size={20} />
            <p className="text-sm text-neutral-200">{alertMessage}</p>
          </div>
          <div className="mt-4 flex justify-end">
            <button
              onClick={() => setAlertMessage("")}
              className="rounded-md bg-amber-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-amber-400"
            >
              รับทราบ
            </button>
          </div>
        </Modal>
      )}

      {/* Stock adjustment modal (ปรับปรุงสต็อก) */}
      {showAdjustForm && adjustForm && (
        <Modal
          title="ปรับปรุงสต็อก (นับสต็อกจริง)"
          onClose={() => setShowAdjustForm(false)}
          footer={
            <div className="flex justify-end gap-2">
              <button type="button" onClick={() => setShowAdjustForm(false)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={submitAdjustForm}
                disabled={adjustForm.actualQty === "" || Number(adjustForm.actualQty) < 0}
                className="rounded-md bg-fuchsia-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-fuchsia-400 disabled:cursor-not-allowed disabled:opacity-40"
              >
                บันทึกการปรับปรุง
              </button>
            </div>
          }
        >
          {(() => {
            const item = items.find((i) => i.id === adjustForm.itemId);
            if (!item) return null;
            const lots = item.lots || [];
            return (
              <>
                <div className="mb-3 rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2">
                  <p className="font-semibold text-neutral-100">{item.name}</p>
                  <p className="font-mono text-xs text-neutral-500">{item.sku} · คงเหลือรวมตามระบบ {formatNum(totalQty(item))} {item.unit}</p>
                </div>

                <Field label="เลือกล็อตที่จะปรับ">
                  <select
                    className={inputCls}
                    value={adjustForm.lotId}
                    onChange={(e) => selectAdjustLot(e.target.value, item)}
                  >
                    {lots.map((l) => (
                      <option key={l.id} value={l.id}>
                        {locationLabel(l.location)}{l.productionDate ? ` · ผลิต ${formatDateOnly(l.productionDate)}` : ""} (ระบบมี {formatNum(l.qty)} {item.unit})
                      </option>
                    ))}
                    <option value="new">-- พบสต็อกใหม่ที่ไม่เคยบันทึกไว้ --</option>
                  </select>
                </Field>

                {adjustForm.lotId === "new" && (
                  <div className="grid grid-cols-2 gap-3">
                    <Field label="สถานที่">
                      <select className={inputCls} value={adjustForm.location} onChange={(e) => setAdjustForm({ ...adjustForm, location: e.target.value })}>
                        {LOCATIONS.map((loc) => (
                          <option key={loc} value={loc}>{loc}</option>
                        ))}
                      </select>
                    </Field>
                    <Field label="วันที่ผลิต (ถ้ามี)">
                      <DateField value={adjustForm.productionDate} onChange={(v) => setAdjustForm({ ...adjustForm, productionDate: v })} />
                    </Field>
                  </div>
                )}

                <Field label={`ยอดนับได้จริง (${item.unit})`}>
                  <input
                    type="number"
                    min="0"
                    autoFocus
                    className={inputCls}
                    value={adjustForm.actualQty}
                    onChange={(e) => setAdjustForm({ ...adjustForm, actualQty: e.target.value })}
                  />
                  {adjustForm.actualQty !== "" && adjustForm.lotId !== "new" && (
                    <span className={`mt-1 block text-xs ${Number(adjustForm.actualQty) >= (lots.find((l) => l.id === adjustForm.lotId)?.qty || 0) ? "text-emerald-400" : "text-red-400"}`}>
                      ผลต่างจากระบบ: {Number(adjustForm.actualQty) - (lots.find((l) => l.id === adjustForm.lotId)?.qty || 0) > 0 ? "+" : ""}
                      {formatNum(Number(adjustForm.actualQty) - (lots.find((l) => l.id === adjustForm.lotId)?.qty || 0))} {item.unit}
                    </span>
                  )}
                </Field>

                <Field label="เหตุผล">
                  <select className={inputCls} value={adjustForm.reason} onChange={(e) => setAdjustForm({ ...adjustForm, reason: e.target.value })}>
                    {ADJUST_REASONS.map((r) => (
                      <option key={r} value={r}>{r}</option>
                    ))}
                  </select>
                </Field>

                <Field label="หมายเหตุ (ไม่บังคับ)">
                  <input
                    className={inputCls}
                    placeholder="รายละเอียดเพิ่มเติม"
                    value={adjustForm.note}
                    onChange={(e) => setAdjustForm({ ...adjustForm, note: e.target.value })}
                  />
                </Field>
              </>
            );
          })()}
        </Modal>
      )}

      {/* System settings modal */}
      {showSettings && (
        <Modal
          title={t("settings")}
          onClose={() => setShowSettings(false)}
          footer={
            <div className="flex justify-end">
              <button onClick={() => setShowSettings(false)} className="rounded-md bg-amber-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-amber-400">
                ปิด
              </button>
            </div>
          }
        >
          <div className="mb-4">
            <p className="mb-2 flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-neutral-400">
              <ZoomIn size={13} /> ขนาดหน้าจอ
            </p>
            <div className="flex gap-2">
              {DISPLAY_SIZES.map((d) => (
                <button
                  key={d.id}
                  onClick={() => updateSetting("displaySize", d.id)}
                  className={`flex-1 rounded-md border px-3 py-2 text-sm font-semibold ${
                    settings.displaySize === d.id ? "border-amber-500 bg-amber-500/20 text-amber-300" : "border-neutral-700 bg-neutral-800 text-neutral-400"
                  }`}
                >
                  {d.label}
                </button>
              ))}
            </div>
          </div>

          <div className="mb-4">
            <p className="mb-2 flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-neutral-400">
              <Globe size={13} /> ภาษา (Language)
            </p>
            <div className="flex gap-2">
              <button
                onClick={() => updateSetting("language", "th")}
                className={`flex-1 rounded-md border px-3 py-2 text-sm font-semibold ${
                  settings.language === "th" ? "border-amber-500 bg-amber-500/20 text-amber-300" : "border-neutral-700 bg-neutral-800 text-neutral-400"
                }`}
              >
                ไทย
              </button>
              <button
                onClick={() => updateSetting("language", "en")}
                className={`flex-1 rounded-md border px-3 py-2 text-sm font-semibold ${
                  settings.language === "en" ? "border-amber-500 bg-amber-500/20 text-amber-300" : "border-neutral-700 bg-neutral-800 text-neutral-400"
                }`}
              >
                English
              </button>
            </div>
            <p className="mt-1.5 text-[11px] text-neutral-600">
              ตอนนี้แปลเฉพาะเมนูหลัก/แดชบอร์ด ส่วนแบบฟอร์มย่อยยังเป็นภาษาไทยเท่านั้น
            </p>
          </div>

          <div className="mb-4 rounded-md border border-neutral-800 bg-neutral-950 p-3">
            <p className="mb-2 flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-neutral-400">
              <Printer size={13} /> การพิมพ์
            </p>
            <Field label="ขนาดกระดาษ">
              <select
                className={inputCls}
                value={settings.printPaperSize}
                onChange={(e) => updateSetting("printPaperSize", e.target.value)}
              >
                {PAPER_SIZES.map((p) => (
                  <option key={p} value={p}>{p}</option>
                ))}
              </select>
            </Field>
            <Field label="ข้อความหัวกระดาษ (ไม่บังคับ)">
              <input
                className={inputCls}
                placeholder="เช่น ชื่อบริษัท / ที่อยู่โรงงาน"
                value={settings.printHeaderText}
                onChange={(e) => updateSetting("printHeaderText", e.target.value)}
              />
            </Field>
          </div>

          {isAdmin && (
            <div className="mb-4 rounded-md border border-rose-900/50 bg-rose-950/20 p-3">
              <p className="mb-2 flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-rose-300">
                <Clock size={13} /> เกณฑ์การขออนุมัติเบิกสินค้า
              </p>
              <Field label="ต้องขออนุมัติเมื่อเบิกตั้งแต่ (จำนวนหน่วย, 0 = ปิดใช้งาน)">
                <input
                  type="number"
                  min="0"
                  className={inputCls}
                  value={settings.approvalThreshold}
                  onChange={(e) => updateSetting("approvalThreshold", Number(e.target.value) || 0)}
                />
              </Field>
              <p className="text-[11px] text-neutral-500">
                ถ้าตั้งไว้มากกว่า 0 พนักงาน (ที่ไม่ใช่แอดมิน) ที่เบิกสินค้าตั้งแต่จำนวนนี้ขึ้นไปจะต้องรอแอดมินอนุมัติก่อน จึงจะตัดสต็อกจริง
              </p>
            </div>
          )}

          {isAdmin && (
            <div className="mb-4 rounded-md border border-sky-900/50 bg-sky-950/20 p-3">
              <p className="mb-2 flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-sky-300">
                <FileSpreadsheet size={13} /> สำรอง / กู้คืนข้อมูลทั้งระบบ
              </p>
              <p className="mb-2 text-xs text-neutral-400">
                สำรองข้อมูลทั้งหมด (สินค้า ล็อต ประวัติ ใบตัดวัตถุดิบ ใบตรวจ FG ผู้ใช้งาน PO และการตั้งค่า) เป็นไฟล์เดียว เพื่อย้ายเครื่องหรือกู้คืนภายหลังได้
              </p>
              <p className="mb-2 text-[11px] text-sky-300">
                ⚠ ข้อมูลทั้งหมดเก็บอยู่ในเบราว์เซอร์ของเครื่องนี้เท่านั้น ไม่ sync ข้ามเครื่อง/เบราว์เซอร์อัตโนมัติ —{" "}
                {settings.lastBackupTs
                  ? `สำรองล่าสุดเมื่อ ${formatDateOnly(new Date(settings.lastBackupTs).toISOString().slice(0, 10))}`
                  : "ยังไม่เคยสำรองข้อมูลเลย"}
              </p>
              <div className="flex flex-wrap gap-2">
                <button
                  onClick={backupData}
                  className="flex items-center gap-1.5 rounded-md border border-sky-600/50 bg-sky-500/10 px-3 py-2 text-xs font-semibold text-sky-300 hover:bg-sky-500/20"
                >
                  <FileSpreadsheet size={13} /> สำรองข้อมูล (.json)
                </button>
                <button
                  onClick={() => restoreInputRef.current?.click()}
                  className="flex items-center gap-1.5 rounded-md border border-neutral-700 bg-neutral-900 px-3 py-2 text-xs font-semibold text-neutral-300 hover:bg-neutral-800"
                >
                  <Upload size={13} /> กู้คืนจากไฟล์สำรอง
                </button>
                <input
                  ref={restoreInputRef}
                  type="file"
                  accept="application/json"
                  className="hidden"
                  onChange={(e) => {
                    const file = e.target.files && e.target.files[0];
                    if (file) restoreData(file);
                    e.target.value = "";
                  }}
                />
              </div>
              <p className="mt-1.5 text-[11px] text-neutral-600">
                การกู้คืนจะแทนที่ข้อมูลปัจจุบันทั้งหมดด้วยข้อมูลในไฟล์สำรอง ควรสำรองข้อมูลปัจจุบันไว้ก่อนถ้าไม่แน่ใจ
              </p>
            </div>
          )}

          {isAdmin && (
            <div className="rounded-md border border-red-900/50 bg-red-950/20 p-3">
              <p className="mb-2 flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-red-400">
                <Trash size={13} /> ล้างประวัติการโอนเข้าออก
              </p>
              <p className="mb-2 text-xs text-neutral-400">
                ลบประวัติการเคลื่อนไหวทั้งหมด (รับเข้า/เบิกออก/โอนย้าย/ตัดของเสีย ฯลฯ) ถาวร ไม่กระทบยอดสต็อกปัจจุบันของสินค้า
              </p>
              {!confirmClearHistory ? (
                <button
                  onClick={() => setConfirmClearHistory(true)}
                  className="rounded-md border border-red-600/50 bg-red-500/10 px-3 py-2 text-sm font-semibold text-red-300 hover:bg-red-500/20"
                >
                  ล้างประวัติการเคลื่อนไหวทั้งหมด
                </button>
              ) : (
                <div className="flex items-center gap-2">
                  <p className="text-xs text-red-300">แน่ใจนะ? ย้อนกลับไม่ได้</p>
                  <button onClick={clearHistory} className="rounded-md bg-red-500 px-3 py-1.5 text-xs font-bold text-neutral-950 hover:bg-red-400">
                    ยืนยันล้าง
                  </button>
                  <button
                    onClick={() => setConfirmClearHistory(false)}
                    className="rounded-md border border-neutral-700 px-3 py-1.5 text-xs text-neutral-300 hover:bg-neutral-800"
                  >
                    ยกเลิก
                  </button>
                </div>
              )}
            </div>
          )}
        </Modal>
      )}

      {/* User management modal (admin only) */}
      {showUserManager && (
        <Modal title="จัดการผู้ใช้งาน" onClose={() => setShowUserManager(false)}>
          {users.length === 0 ? (
            <p className="text-sm text-neutral-500">ยังไม่มีผู้ใช้งานอื่น</p>
          ) : (
            <ul className="space-y-1.5">
              {users.map((u) => (
                <li key={u.id} className="flex items-center justify-between gap-2 rounded-md border border-neutral-800 bg-neutral-950 px-3 py-2">
                  <span className="text-sm font-semibold text-neutral-200">
                    {u.name} {u.id === currentUser.id && <span className="text-xs text-amber-400">(คุณ)</span>}
                  </span>
                  <div className="flex items-center gap-2">
                    <select
                      className={`${inputCls} w-auto py-1 text-xs`}
                      value={u.role}
                      onChange={(e) => changeUserRole(u.id, e.target.value)}
                    >
                      {ROLES.map((r) => (
                        <option key={r.id} value={r.id}>{r.label}</option>
                      ))}
                    </select>
                    <button
                      onClick={() => {
                        setResettingPwUser(u);
                        setResetPwValue("");
                      }}
                      title="ตั้งรหัสผ่านใหม่"
                      className="rounded p-1.5 text-neutral-500 hover:bg-neutral-800 hover:text-amber-400"
                    >
                      <Shield size={15} />
                    </button>
                    {u.id !== currentUser.id && (
                      <button onClick={() => removeUser(u.id)} className="rounded p-1.5 text-neutral-500 hover:bg-neutral-800 hover:text-red-400">
                        <Trash2 size={15} />
                      </button>
                    )}
                  </div>
                </li>
              ))}
            </ul>
          )}
          <div className="mt-4 flex justify-end">
            <button onClick={() => setShowUserManager(false)} className="rounded-md bg-amber-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-amber-400">
              ปิด
            </button>
          </div>
        </Modal>
      )}

      {/* Reset user password modal */}
      {resettingPwUser && (
        <Modal
          title={`ตั้งรหัสผ่านใหม่ให้ ${resettingPwUser.name}`}
          onClose={() => setResettingPwUser(null)}
          footer={
            <div className="flex justify-end gap-2">
              <button type="button" onClick={() => setResettingPwUser(null)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={() => resetUserPassword(resettingPwUser.id, resetPwValue)}
                className="rounded-md bg-amber-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-amber-400"
              >
                บันทึกรหัสผ่านใหม่
              </button>
            </div>
          }
        >
          <Field label="รหัสผ่านใหม่ (อย่างน้อย 4 ตัวอักษร)">
            <input
              type="password"
              autoFocus
              className={inputCls}
              value={resetPwValue}
              onChange={(e) => setResetPwValue(e.target.value)}
            />
          </Field>
        </Modal>
      )}

      {/* Confirm reverse-received-PO modal */}
      {confirmReversePo && (
        <Modal title="ยืนยันยกเลิกใบสั่งซื้อที่รับผิด" onClose={() => setConfirmReversePo(null)}>
          <p className="text-sm text-neutral-300">
            ใบสั่งซื้อ <span className="font-mono font-semibold text-neutral-100">{confirmReversePo.poNumber}</span> ถูกทำเครื่องหมายว่า "รับของแล้ว"
            การยกเลิกนี้จะ <span className="font-semibold text-red-400">ตัดสต็อกที่เคยรับเข้าไปจากรายการนี้ออกทั้งหมด</span> (เท่าที่ยังมีเหลืออยู่ในคลัง)
            และเปลี่ยนสถานะใบสั่งซื้อเป็น "ยกเลิก" การกระทำนี้ย้อนกลับไม่ได้
          </p>
          <div className="mt-4 flex justify-end gap-2">
            <button onClick={() => setConfirmReversePo(null)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
              ยกเลิก
            </button>
            <button
              onClick={() => {
                reversePoReceive(confirmReversePo);
                setConfirmReversePo(null);
              }}
              className="rounded-md bg-red-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-red-400"
            >
              ยืนยันยกเลิกและตัดสต็อกออก
            </button>
          </div>
        </Modal>
      )}

      {/* Delete confirm modal */}
      {confirmDelete && (
        <Modal title="ยืนยันการลบสินค้า" onClose={() => setConfirmDelete(null)}>
          <p className="text-sm text-neutral-300">
            ต้องการลบ <span className="font-semibold text-neutral-100">{confirmDelete.name}</span> ออกจากคลังใช่หรือไม่?
            การลบนี้ไม่สามารถย้อนกลับได้
          </p>
          <div className="mt-4 flex justify-end gap-2">
            <button onClick={() => setConfirmDelete(null)} className="rounded-md border border-neutral-700 px-3 py-2 text-sm text-neutral-300 hover:bg-neutral-800">
              ยกเลิก
            </button>
            <button onClick={() => deleteItem(confirmDelete.id)} className="rounded-md bg-red-500 px-4 py-2 text-sm font-bold text-neutral-950 hover:bg-red-400">
              ลบสินค้า
            </button>
          </div>
        </Modal>
      )}
    </div>
    </div>
  );
}
