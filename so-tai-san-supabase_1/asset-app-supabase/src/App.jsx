import { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";
import {
  LayoutDashboard, Boxes, Users2, Wrench, History, FileText, Archive, Warehouse, Fuel, Cable, CircleDollarSign, UploadCloud, DownloadCloud,
  ClipboardList, UserCog, ScrollText, Search, RefreshCw, X, Plus,
  ChevronRight, Building2, ArrowLeftRight, ShieldCheck, Trash2, Pencil,
  ChevronDown, Package, Stamp, AlertCircle, Loader2, ClipboardCheck,
  Settings as SettingsIcon, FileDown, Printer, Check, PencilLine,
  Lock, LogOut, User, ShieldAlert, Eye, EyeOff, Download, Upload, RotateCcw, HardDrive, Save,
} from "lucide-react";

/* ============================== DESIGN TOKENS ==============================
Color — theo bộ nhận diện TMC (đỏ chủ đạo, nền sáng):
  paper       #F7F4F3  nền trang
  surface     #FFFFFF  thẻ / bảng
  ink         #241A19  chữ chính
  muted       #7C6E6D  chữ phụ
  border      #EEE1DF  viền mảnh
  brand       #C1272D  đỏ TMC — hành động chính, điều hướng đang chọn
  brandSoft   #FBE8E8  sắc đỏ nhạt
  gold        #C08A1E  vàng đồng — tiền/khấu hao/giá trị
  goldSoft    #FAF0DC
  danger      #C2410C  cam cháy — hỏng hóc/thanh lý (tách biệt với đỏ thương hiệu)
  dangerSoft  #FBEAE0
  info        #2C5F7C  xanh xám — chuyển đổi/thông tin
  infoSoft    #E6EEF2
Type:
  Display: "Space Grotesk"  (tiêu đề, số liệu lớn)
  Body:    "IBM Plex Sans"  (nội dung chung)
  Mono:    "IBM Plex Mono"  (mã tài sản, số tiền, ngày tháng)
============================================================================ */

const TOKENS = {
  paper: "#F7F4F3", surface: "#FFFFFF", ink: "#241A19", muted: "#7C6E6D",
  border: "#EEE1DF", brand: "#C1272D", brandSoft: "#FBE8E8",
  gold: "#C08A1E", goldSoft: "#FAF0DC", danger: "#C2410C", dangerSoft: "#FBEAE0",
  info: "#2C5F7C", infoSoft: "#E6EEF2",
};

const FONT_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
.aa-root{font-family:'IBM Plex Sans',sans-serif;color:${TOKENS.ink};background:${TOKENS.paper};}
.aa-display{font-family:'Space Grotesk',sans-serif;}
.aa-mono{font-family:'IBM Plex Mono',monospace;}
.aa-scroll::-webkit-scrollbar{width:8px;height:8px;}
.aa-scroll::-webkit-scrollbar-thumb{background:#D6DAD1;border-radius:4px;}
.aa-row:hover{background:${TOKENS.brandSoft}55;}
.aa-fade{animation:aaFade .15s ease-out;}
@keyframes aaFade{from{opacity:0;transform:translateY(2px)}to{opacity:1;transform:translateY(0)}}
.aa-slide{animation:aaSlide .18s ease-out;}
@keyframes aaSlide{from{transform:translateX(16px);opacity:0}to{transform:translateX(0);opacity:1}}
input:focus,select:focus,textarea:focus{outline:2px solid ${TOKENS.brand}; outline-offset:1px;}
button:focus-visible{outline:2px solid ${TOKENS.brand}; outline-offset:2px;}
`;

/* ============================== CONSTANTS ============================== */

const STATUS = {
  ASSIGNED: "Đã cấp phát",
  SHARED: "Dùng chung",
  UNUSED: "Chưa sử dụng",
  REPAIR: "Đang sửa chữa/bảo dưỡng",
  BROKEN: "Hỏng/mất",
  LIQUIDATED: "Đã thanh lý",
  TRANSFERRED_OUT: "Đã chuyển team khác",
};

const STATUS_COLOR = {
  [STATUS.ASSIGNED]: TOKENS.brand,
  [STATUS.SHARED]: TOKENS.info,
  [STATUS.UNUSED]: TOKENS.muted,
  [STATUS.REPAIR]: TOKENS.gold,
  [STATUS.BROKEN]: TOKENS.danger,
  [STATUS.LIQUIDATED]: "#8A8FA3",
  [STATUS.TRANSFERRED_OUT]: "#6E4E9E",
};

const DEPARTMENTS = ["IT", "Kế toán", "Kinh doanh", "Nhân sự", "Vận hành", "Ban giám đốc"];
const CATEGORIES = ["Laptop", "Màn hình", "Điện thoại", "Máy in", "Nội thất văn phòng", "Thiết bị mạng", "Khác"];
const FIELD_TYPES = [
  { id: "text", label: "Chữ" },
  { id: "number", label: "Số" },
  { id: "date", label: "Ngày" },
];

const ROLES = {
  admin: { label: "Toàn quyền", color: "#22594C" },
  user: { label: "Nhập liệu & xuất báo cáo", color: "#34607F" },
};

/* ---------- export helpers (Excel + in PDF) ---------- */

function exportExcel(filename, headers, rows) {
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws["!cols"] = headers.map((h, i) => ({ wch: Math.max(12, h.length + 2, ...rows.map((r) => String(r[i] ?? "").length + 1)) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Dữ liệu");
  XLSX.writeFile(wb, `${filename}.xlsx`);
}

function downloadExcelTemplate(filename, headers, sampleRows = []) {
  const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
  ws["!cols"] = headers.map((h) => ({ wch: Math.max(14, String(h).length + 3) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Mẫu nhập liệu");
  XLSX.writeFile(wb, `${filename}.xlsx`);
}

const NAV = [
  { section: "TÀI SẢN", items: [
    { id: "overview", label: "Tổng quan", icon: LayoutDashboard },
    { id: "catalog", label: "Danh mục tài sản", icon: Boxes, badge: "assets" },
    { id: "byProject", label: "Tài sản theo công trình", icon: Building2 },
    { id: "depreciation", label: "Khấu hao tài sản", icon: History },
    { id: "repair", label: "Sửa chữa", icon: Wrench },
    { id: "repairHistory", label: "Lịch sử sửa chữa", icon: ScrollText },
    { id: "liquidation", label: "Thanh lý", icon: Trash2 },
    { id: "warehouse", label: "Kho — Nhập / Xuất / Tồn", icon: Warehouse },
    { id: "costHistory", label: "Chi phí thiết bị", icon: CircleDollarSign },
  ]},
  { section: "CHỨNG TỪ", items: [
    { id: "minutes", label: "Biên bản", icon: FileText },
    { id: "transactions", label: "Lịch sử giao dịch", icon: ArrowLeftRight },
  ]},
  { section: "TỔ CHỨC", items: [
    { id: "projects", label: "Công trình", icon: Building2, badge: "projects" },
    { id: "activityLog", label: "Nhật ký thao tác", icon: ClipboardList },
    { id: "settings", label: "Cài đặt", icon: SettingsIcon, adminOnly: true },
  ]},
];

/* ============================== UTILITIES ============================== */

let __seq = 1000;
const uid = (p) => `${p}-${(__seq++).toString(36)}`;

const fmtVND = (n) => (Number(n) || 0).toLocaleString("vi-VN") + " đ";
const fmtDate = (iso) => {
  if (!iso) return "—";
  const d = new Date(iso);
  if (isNaN(d)) return "—";
  return d.toLocaleDateString("vi-VN");
};
const nowIso = () => new Date().toISOString();
const monthsBetween = (a, b) => {
  const d1 = new Date(a), d2 = new Date(b);
  return Math.max(0, (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth()));
};
const depreciationOf = (asset) => {
  const elapsed = Math.min(monthsBetween(asset.purchaseDate, nowIso()), asset.usefulLifeMonths);
  const monthly = asset.cost / asset.usefulLifeMonths;
  const depreciated = Math.round(monthly * elapsed);
  const remaining = Math.max(0, asset.cost - depreciated);
  const pct = asset.cost ? Math.round((depreciated / asset.cost) * 100) : 0;
  return { elapsed, monthly: Math.round(monthly), depreciated, remaining, pct };
};

/* ============================== SEED DATA ============================== */

function seed() {
  const projects = [
    { id: "p1", commander: "Nguyễn Văn An", name: "Công trình mẫu A", address: "Hà Nội", workItem: "Hạng mục nền móng", startDate: "2026-01-01", endDate: "2026-12-31" },
    { id: "p2", commander: "Trần Văn Bình", name: "Công trình mẫu B", address: "Bắc Giang", workItem: "Hạng mục kết cấu", startDate: "2026-03-01", endDate: "2026-10-31" },
  ];

  const mk = (code, name, category, cost, purchaseDate, life, status, empId, dept, serial) => ({
    id: uid("as"), code, name, category, cost, purchaseDate, usefulLifeMonths: life,
    status, assignedTo: null, projectId: empId || null, department: dept || "Vận hành", serial,
    assetGroup: "Thiết bị chính", ownership: "Công ty", quantity: 1,
    supplier: "Nhà cung cấp chưa ghi", warranty: true, warrantyEnd: "2027-03-01", note: "",
  });

  const assets = [
    mk("TS-LT-001", "Laptop Dell", "Laptop", 10000000, "2020-03-01", 36, STATUS.UNUSED, null, "IT", "SN-LT001"),
    mk("TS-LT-002", "Laptop Dell", "Laptop", 10000000, "2020-03-01", 36, STATUS.ASSIGNED, "e1", "IT", "SN-LT002"),
    mk("TS-LT-003", "Macbook Pro 14-inch 2021", "Laptop", 20000000, "2023-02-01", 36, STATUS.ASSIGNED, "e3", "Kinh doanh", "SN-LT003"),
    mk("TS-LT-004", "Laptop Dell", "Laptop", 10000000, "2020-03-01", 36, STATUS.UNUSED, null, "IT", "SN-LT004"),
    mk("TS-LT-005", "ThinkPad X1", "Laptop", 15000000, "2024-01-15", 36, STATUS.ASSIGNED, "e2", "Vận hành", "SN-LT005"),
    mk("TS-LT-006", "Macbook Pro 14-inch 2021", "Laptop", 20000000, "2023-02-01", 36, STATUS.ASSIGNED, "e2", "Vận hành", "SN-LT006"),
    mk("TS-LT-007", "Laptop Dell", "Laptop", 10000000, "2020-03-01", 36, STATUS.REPAIR, null, "IT", "SN-LT007"),
    mk("TS-LT-008", "Macbook Pro", "Laptop", 20000000, "2020-03-01", 36, STATUS.SHARED, null, "Kinh doanh", "SN-LT008"),
    mk("MH-001", "Màn hình ViewSonic 24 inch", "Màn hình", 10000000, "2020-03-01", 36, STATUS.UNUSED, null, "IT", "SN-MH001"),
    mk("MH-002", "Màn hình Dell", "Màn hình", 10000000, "2020-03-01", 36, STATUS.SHARED, null, "IT", "SN-MH002"),
    mk("DT-001", "iPhone 13", "Điện thoại", 18000000, "2022-06-01", 24, STATUS.LIQUIDATED, null, "Kinh doanh", "SN-DT001"),
    mk("MI-001", "Máy in Canon LBP", "Máy in", 6000000, "2021-09-01", 48, STATUS.TRANSFERRED_OUT, null, "Vận hành", "SN-MI001"),
  ];

  const transactions = [
    { id: uid("tx"), assetId: assets[1].id, type: "mua_sam", date: "2020-03-01", title: "Mua sắm", detail: "Mua từ nhà cung cấp chưa ghi", amount: 10000000 },
    { id: uid("tx"), assetId: assets[1].id, type: "cap_phat", date: "2020-03-02", title: "Cấp phát", detail: `Giao cho công trình ${projects[0].name}`, amount: 0 },
    { id: uid("tx"), assetId: assets[6].id, type: "sua_chua", date: "2026-06-10", title: "Gửi sửa chữa", detail: "Lỗi bàn phím, gửi bảo hành", amount: 800000 },
    { id: uid("tx"), assetId: assets[10].id, type: "thanh_ly", date: "2026-05-20", title: "Thanh lý", detail: "Thiết bị cũ, thanh lý thu hồi", amount: 3000000 },
  ];

  const repairs = [
    { id: uid("rp"), assetId: assets[6].id, date: "2026-06-10", description: "Lỗi bàn phím, gửi bảo hành", cost: 800000, status: "Đang sửa" },
  ];

  const liquidations = [
    { id: uid("lq"), assetId: assets[10].id, date: "2026-05-20", value: 3000000, reason: "Thiết bị cũ, hiệu năng thấp" },
  ];

  const minutes = [
    { id: uid("bb"), assetId: assets[1].id, type: "Biên bản bàn giao", date: "2020-03-02", content: `Bàn giao ${assets[1].name} cho ${projects[0].name}`, status: "Đã ký" },
  ];

  const activityLog = [
    { id: uid("lg"), date: nowIso(), user: "Hệ thống", action: "Khởi tạo dữ liệu mẫu cho Sổ tài sản" },
  ];

  const settings = {
    companyName: "TÊN CÔNG TY",
    departments: [...DEPARTMENTS],
    categories: [...CATEGORIES],
    customColumns: [], // { key, label, type: 'text'|'number'|'date' }
  };

  // Tài khoản đăng nhập giờ do Supabase Auth quản lý — không còn lưu trong app_data.
  const warehouse = assets.map((a) => ({ id: uid("wh"), voucherNo: `PN-${String(a.purchaseDate || nowIso().slice(0,10)).replaceAll("-", "")}-OPEN`, assetId: a.id, type: "nhap", quantity: Number(a.quantity || 1), date: a.purchaseDate, unitCost: a.cost, note: "Tồn đầu kỳ", projectId: a.projectId || null, locationType: a.projectId ? "project" : "warehouse", locationName: a.projectId ? projects.find(p => p.id === a.projectId)?.name || "Kho trung tâm" : "Kho trung tâm", warehouseName: a.projectId ? "" : "Kho trung tâm", itemName: a.name, itemCode: a.code, category: a.category, assetGroup: a.assetGroup, ownership: a.ownership }));
  return { assets, projects, transactions, repairs, liquidations, minutes, warehouse, costHistory: [], activityLog, settings };
}

// Fills in any fields missing from data saved by an older version of the app.
function withDefaults(d) {
  const legacyProjects = Array.isArray(d.projects) ? d.projects : (d.employees || []).map((e, i) => ({
    id: e.id || `p_legacy_${i}`, commander: e.name || "", name: e.name || "Công trình cũ", address: "", workItem: "", startDate: "", endDate: ""
  }));
  const migratedAssets = (d.assets || []).map(a => ({
    ...a,
    projectId: a.projectId || a.assignedTo || null,
    assetGroup: a.assetGroup || a.category || "Thiết bị chính",
    ownership: a.ownership || "Công ty",
    quantity: Number(a.quantity || 1)
  }));
  return {
    ...d,
    assets: migratedAssets,
    projects: legacyProjects,
    warehouse: (Array.isArray(d.warehouse) ? d.warehouse : migratedAssets.map(a => ({ id: uid("wh"), assetId: a.id, type: "nhap", quantity: Number(a.quantity || 1), date: a.purchaseDate || nowIso().slice(0,10), unitCost: Number(a.cost || 0), note: "Tồn đầu kỳ" }))).map((w, i) => { const a = migratedAssets.find(x => x.id === w.assetId); const pid = w.projectId || a?.projectId || null; const loc = w.locationName || (pid ? legacyProjects.find(p => p.id === pid)?.name : "Kho trung tâm") || "Kho trung tâm"; return { ...w, voucherNo: w.voucherNo || `PN-${String(w.date || nowIso().slice(0,10)).replaceAll("-", "")}-OPEN${String(i+1).padStart(3,"0")}`, projectId: pid, locationType: w.locationType || (pid ? "project" : "warehouse"), locationName: loc, warehouseName: w.warehouseName || (pid ? "" : loc), itemName: w.itemName || a?.name || "", itemCode: w.itemCode || a?.code || "", category: w.category || a?.category || "Khác", assetGroup: w.assetGroup || a?.assetGroup || "Thiết bị chính", ownership: w.ownership || a?.ownership || "Công ty" }; }),
    costHistory: Array.isArray(d.costHistory) ? d.costHistory : [],
    employees: undefined,
    settings: {
      companyName: "TÊN CÔNG TY",
      departments: [...DEPARTMENTS],
      categories: [...CATEGORIES],
      customColumns: [],
      ...(d.settings || {}),
    },
  };
}

// Snapshot used for backups — the whole business-data object.
function snapshotOf(d) {
  return d;
}

/* ============================== STORAGE HOOK ============================== */

function useAppData(enabled) {
  const [data, setDataState] = useState(null);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [err, setErr] = useState(null);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const { data: row, error } = await supabase.from("app_data").select("data").eq("id", 1).single();
      if (error) throw error;
      if (row && row.data && Object.keys(row.data).length > 0) {
        setDataState(withDefaults(row.data));
      } else {
        const s = seed();
        setDataState(s);
        await supabase.from("app_data").update({ data: s, updated_at: new Date().toISOString() }).eq("id", 1);
      }
      setErr(null);
    } catch (e) {
      setErr(e?.message ? `Không tải được dữ liệu: ${e.message}` : "Không thể kết nối Supabase. Kiểm tra VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY, đã chạy đúng file supabase-schema.sql, và kết nối mạng.");
    } finally {
      setLoading(false);
    }
  }, []);

  // Chỉ tải dữ liệu nghiệp vụ SAU KHI đã đăng nhập — chính sách bảo mật (RLS) trên
  // Supabase yêu cầu tài khoản phải authenticated mới được đọc bảng app_data.
  useEffect(() => {
    if (!enabled) return;
    load();
  }, [enabled, load]);

  // Đồng bộ thời gian thực — khi người khác lưu thay đổi, mọi người mở app cùng lúc sẽ tự cập nhật.
  useEffect(() => {
    if (!enabled) return;
    const channel = supabase
      .channel("app_data_changes")
      .on(
        "postgres_changes",
        { event: "UPDATE", schema: "public", table: "app_data", filter: "id=eq.1" },
        (payload) => { if (payload.new?.data) setDataState(withDefaults(payload.new.data)); }
      )
      .subscribe();
    return () => { supabase.removeChannel(channel); };
  }, [enabled]);

  const persist = useCallback(async (next) => {
    setDataState(next);
    setSaving(true);
    try {
      const { error } = await supabase.from("app_data").update({ data: next, updated_at: new Date().toISOString() }).eq("id", 1);
      if (error) throw error;
      setErr(null);
    } catch (e) {
      setErr("Lưu thất bại — kiểm tra kết nối mạng, thay đổi có thể chưa lưu lên máy chủ.");
    } finally {
      setSaving(false);
    }
  }, []);

  const reload = useCallback(async () => { await load(); }, [load]);

  return { data, setData: persist, loading, saving, err, reload };
}

/* ============================== SMALL UI PARTS ============================== */

function StatusDot({ status }) {
  const c = STATUS_COLOR[status] || TOKENS.muted;
  return (
    <span className="inline-flex items-center gap-1.5">
      <span style={{ width: 8, height: 8, borderRadius: 999, background: c, display: "inline-block" }} />
      <span style={{ color: TOKENS.ink }} className="text-[13px]">{status}</span>
    </span>
  );
}

function Tag({ children }) {
  return (
    <span
      className="aa-mono text-[12px] px-2 py-0.5 rounded"
      style={{ background: TOKENS.brandSoft, color: TOKENS.brand, border: `1px solid ${TOKENS.brand}22` }}
    >
      {children}
    </span>
  );
}

function StatCard({ label, value, sub, accent }) {
  return (
    <div className="rounded-lg p-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
      <div className="text-[12px] mb-1" style={{ color: TOKENS.muted }}>{label}</div>
      <div className="aa-display text-2xl font-semibold" style={{ color: accent || TOKENS.ink }}>{value}</div>
      {sub ? <div className="text-[11px] mt-1" style={{ color: TOKENS.muted }}>{sub}</div> : null}
    </div>
  );
}

function Btn({ children, onClick, kind = "default", icon: Icon, disabled, type = "button", small }) {
  const styles = {
    primary: { background: TOKENS.brand, color: "#fff", border: `1px solid ${TOKENS.brand}` },
    danger: { background: TOKENS.danger, color: "#fff", border: `1px solid ${TOKENS.danger}` },
    gold: { background: TOKENS.gold, color: "#fff", border: `1px solid ${TOKENS.gold}` },
    info: { background: TOKENS.info, color: "#fff", border: `1px solid ${TOKENS.info}` },
    default: { background: TOKENS.surface, color: TOKENS.ink, border: `1px solid ${TOKENS.border}` },
    ghost: { background: "transparent", color: TOKENS.ink, border: `1px solid transparent` },
  };
  return (
    <button
      type={type}
      disabled={disabled}
      onClick={onClick}
      className={`inline-flex items-center gap-1.5 rounded-md font-medium transition-opacity ${small ? "px-2 py-1 text-[12px]" : "px-3 py-1.5 text-[13px]"} ${disabled ? "opacity-50 cursor-not-allowed" : "hover:opacity-85"}`}
      style={styles[kind]}
    >
      {Icon ? <Icon size={small ? 13 : 14} /> : null}
      {children}
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label className="block mb-3">
      <div className="text-[12px] mb-1" style={{ color: TOKENS.muted }}>{label}</div>
      {children}
    </label>
  );
}

const inputCls = "w-full rounded-md px-2.5 py-1.5 text-[13px] bg-white";
const inputStyle = { border: `1px solid ${TOKENS.border}` };

function Modal({ title, onClose, children, wide }) {
  return (
    <div className="fixed inset-0 z-40 flex items-center justify-center p-4" style={{ background: "#1B232066" }} onClick={onClose}>
      <div
        className={`aa-fade rounded-lg w-full ${wide ? "max-w-2xl" : "max-w-md"} max-h-[85vh] overflow-y-auto aa-scroll`}
        style={{ background: TOKENS.surface }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-3.5 sticky top-0" style={{ background: TOKENS.surface, borderBottom: `1px solid ${TOKENS.border}` }}>
          <div className="aa-display font-semibold text-[15px]">{title}</div>
          <button onClick={onClose} className="p-1 rounded hover:bg-black/5"><X size={16} /></button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function EmptyState({ text, sub }) {
  return (
    <div className="flex flex-col items-center justify-center py-16 text-center">
      <Package size={30} style={{ color: TOKENS.border }} />
      <div className="mt-3 text-[13px] font-medium" style={{ color: TOKENS.ink }}>{text}</div>
      {sub ? <div className="text-[12px] mt-1" style={{ color: TOKENS.muted }}>{sub}</div> : null}
    </div>
  );
}

/* ============================== LOGIN ============================== */

function LoginScreen({ onLogin }) {
  const [mode, setMode] = useState("login"); // 'login' | 'signup'
  const [email, setEmail] = useState("");
  const [name, setName] = useState("");
  const [password, setPassword] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError(""); setInfo(""); setBusy(true);
    try {
      if (mode === "login") {
        const { data: sess, error: err } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
        if (err) throw err;
        await loadProfileAndLogin(sess.user, onLogin, setError);
      } else {
        const { data: sess, error: err } = await supabase.auth.signUp({
          email: email.trim(), password, options: { data: { name: name.trim() || email.split("@")[0] } },
        });
        if (err) throw err;
        if (sess.user && !sess.session) {
          setInfo("Đã tạo tài khoản — kiểm tra email để xác nhận trước khi đăng nhập.");
          setMode("login");
        } else if (sess.user) {
          await loadProfileAndLogin(sess.user, onLogin, setError);
        }
      }
    } catch (err) {
      setError(err.message === "Invalid login credentials" ? "Sai email hoặc mật khẩu" : err.message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="aa-root min-h-screen flex items-center justify-center p-4" style={{ background: TOKENS.paper }}>
      <style>{FONT_CSS}</style>
      <div className="w-full max-w-sm">
        <div className="flex flex-col items-center mb-6">
          <div className="w-11 h-11 rounded-lg flex items-center justify-center mb-3" style={{ background: TOKENS.brand }}>
            <Stamp size={20} color="#fff" />
          </div>
          <div className="aa-display font-bold text-[18px]" style={{ color: TOKENS.brand }}>SỔ TÀI SẢN</div>
          <div className="text-[12px]" style={{ color: TOKENS.muted }}>Quản lý tài sản nội bộ</div>
        </div>

        <form onSubmit={submit} className="rounded-lg p-6" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
          <div className="flex gap-1.5 mb-4">
            {[["login", "Đăng nhập"], ["signup", "Tạo tài khoản"]].map(([id, label]) => (
              <button key={id} type="button" onClick={() => { setMode(id); setError(""); setInfo(""); }}
                className="px-2.5 py-1 rounded-md text-[12.5px] font-medium"
                style={{ background: mode === id ? TOKENS.brand : TOKENS.paper, color: mode === id ? "#fff" : TOKENS.muted }}>
                {label}
              </button>
            ))}
          </div>

          {mode === "signup" && (
            <Field label="Họ tên hiển thị"><input className={inputCls} style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="Nguyễn Văn A" /></Field>
          )}
          <Field label="Email">
            <input autoFocus type="email" className={inputCls} style={inputStyle} value={email} onChange={(e) => setEmail(e.target.value)} placeholder="ban@congty.vn" />
          </Field>
          <Field label="Mật khẩu">
            <div className="relative">
              <input type={showPw ? "text" : "password"} className={inputCls} style={{ ...inputStyle, paddingRight: 34 }} value={password}
                onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" />
              <button type="button" onClick={() => setShowPw((s) => !s)} className="absolute right-2 top-1/2 -translate-y-1/2 p-0.5" style={{ color: TOKENS.muted }}>
                {showPw ? <EyeOff size={14} /> : <Eye size={14} />}
              </button>
            </div>
          </Field>

          {error && (
            <div className="flex items-center gap-1.5 text-[12px] rounded-md px-2.5 py-1.5 mb-3" style={{ background: TOKENS.dangerSoft, color: TOKENS.danger }}>
              <AlertCircle size={13} /> {error}
            </div>
          )}
          {info && (
            <div className="flex items-center gap-1.5 text-[12px] rounded-md px-2.5 py-1.5 mb-3" style={{ background: TOKENS.brandSoft, color: TOKENS.brand }}>
              <ClipboardCheck size={13} /> {info}
            </div>
          )}

          <button type="submit" disabled={busy} className="w-full rounded-md py-2 text-[13px] font-medium hover:opacity-90 disabled:opacity-60"
            style={{ background: TOKENS.brand, color: "#fff" }}>
            {busy ? "Đang xử lý…" : mode === "login" ? "Đăng nhập" : "Tạo tài khoản"}
          </button>

          {mode === "signup" && (
            <div className="mt-3 text-[11px] leading-relaxed" style={{ color: TOKENS.muted }}>
              Tài khoản mới mặc định ở quyền "Nhập liệu &amp; xuất báo cáo". Nhờ quản trị viên nâng quyền trong Cài đặt → Tài khoản nếu cần Toàn quyền.
            </div>
          )}
        </form>
      </div>
    </div>
  );
}

async function loadProfileAndLogin(user, onLogin, setError) {
  const { data: profile, error } = await supabase.from("profiles").select("*").eq("id", user.id).single();
  if (error || !profile) { setError("Không đọc được hồ sơ người dùng — thử tải lại trang."); return; }
  onLogin({ id: user.id, email: user.email, name: profile.name, role: profile.role });
}

/* ============================== APP ============================== */

export default function AssetManagementApp() {
  const [active, setActive] = useState("overview");
  const [query, setQuery] = useState("");
  const [selectedAssetId, setSelectedAssetId] = useState(null);
  const [modal, setModal] = useState(null); // {type, assetId?}
  const [toast, setToast] = useState(null);
  const [printJob, setPrintJob] = useState(null); // {title, headers, rows}
  const [currentUser, setCurrentUser] = useState(null); // {id, email, name, role} — restored from Supabase session
  const [authChecked, setAuthChecked] = useState(false);
  const [backups, setBackups] = useState([]); // [{id, label, created_by, created_at}] — from the `backups` table
  const [profiles, setProfiles] = useState([]); // [{id, email, name, role}] — from the `profiles` table (admin view)

  // Dữ liệu nghiệp vụ chỉ tải SAU KHI có currentUser (đăng nhập xong) — xem lý do trong useAppData.
  const { data, setData, loading, saving, err, reload } = useAppData(!!currentUser);

  // Khôi phục phiên đăng nhập khi mở lại trang, và lắng nghe đăng xuất/hết hạn phiên.
  useEffect(() => {
    let active = true;
    supabase.auth.getSession().then(async ({ data: { session } }) => {
      if (session?.user) await loadProfileAndLogin(session.user, (u) => { if (active) setCurrentUser(u); }, () => {});
      if (active) setAuthChecked(true);
    });
    const { data: sub } = supabase.auth.onAuthStateChange((event) => {
      if (event === "SIGNED_OUT") setCurrentUser(null);
    });
    return () => { active = false; sub.subscription.unsubscribe(); };
  }, []);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToast(null), 2400);
    return () => clearTimeout(t);
  }, [toast]);

  useEffect(() => {
    if (!printJob) return;
    const t = setTimeout(() => window.print(), 80);
    const onAfter = () => setPrintJob(null);
    window.addEventListener("afterprint", onAfter);
    return () => { clearTimeout(t); window.removeEventListener("afterprint", onAfter); };
  }, [printJob]);

  // NOTE: these must run on every render (React hook rules), so they sit above
  // the loading/login early-returns below and guard internally for missing data.
  const counts = useMemo(() => {
    if (!data) return { total: 0, warranty: 0, repairCostThisYear: 0, openRepairs: 0 };
    const c = { total: data.assets.length };
    Object.values(STATUS).forEach((s) => (c[s] = data.assets.filter((a) => a.status === s).length));
    c.warranty = data.assets.filter((a) => a.warranty).length;
    c.repairCostThisYear = data.repairs.filter((r) => r.date.startsWith("2026")).reduce((s, r) => s + r.cost, 0);
    c.openRepairs = data.repairs.filter((r) => r.status === "Đang sửa").length;
    return c;
  }, [data]);

  const filteredAssets = useMemo(() => {
    if (!data) return [];
    if (!query.trim()) return data.assets;
    const q = query.toLowerCase();
    return data.assets.filter((a) =>
      a.code.toLowerCase().includes(q) || a.name.toLowerCase().includes(q) ||
      (a.serial || "").toLowerCase().includes(q) ||
      (data.projects.find((e) => e.id === a.projectId)?.name || "").toLowerCase().includes(q)
    );
  }, [data, query]);

  const MAX_BACKUPS = 15;

  const loadBackups = useCallback(async () => {
    const { data: rows, error } = await supabase.from("backups").select("id, label, created_by, created_at").order("created_at", { ascending: false }).limit(MAX_BACKUPS);
    if (!error) setBackups(rows || []);
    return rows || [];
  }, []);

  // Tải danh sách bản sao lưu + tự động sao lưu 1 lần/ngày khi admin mở app.
  useEffect(() => {
    if (!currentUser) return;
    (async () => {
      const rows = await loadBackups();
      if (currentUser.role === "admin") {
        const today = nowIso().slice(0, 10);
        const last = rows[0];
        if (!last || last.created_at.slice(0, 10) !== today) {
          createBackup("Tự động hàng ngày", { silent: true });
        }
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentUser]);

  // Danh sách tài khoản — chỉ admin cần, dùng ở Cài đặt → Tài khoản đăng nhập.
  useEffect(() => {
    if (!currentUser || currentUser.role !== "admin") return;
    supabase.from("profiles").select("id, email, name, role").order("created_at", { ascending: true })
      .then(({ data: rows, error }) => { if (!error) setProfiles(rows || []); });
  }, [currentUser]);

  // 1) Chưa xác định xong trạng thái đăng nhập — chờ.
  if (!authChecked) {
    return (
      <div className="aa-root min-h-screen flex items-center justify-center" style={{ background: TOKENS.paper }}>
        <style>{FONT_CSS}</style>
        <div className="flex flex-col items-center gap-2">
          <Loader2 className="animate-spin" size={22} style={{ color: TOKENS.brand }} />
          <div className="text-[13px]" style={{ color: TOKENS.muted }}>Đang kiểm tra đăng nhập…</div>
        </div>
      </div>
    );
  }

  // 2) Chưa đăng nhập — hiện màn hình đăng nhập (không cần dữ liệu tài sản trước đó).
  if (!currentUser) {
    return <LoginScreen onLogin={setCurrentUser} />;
  }

  // 3) Đã đăng nhập nhưng đang tải / lỗi khi tải dữ liệu tài sản từ Supabase.
  if (loading || !data) {
    return (
      <div className="aa-root min-h-screen flex items-center justify-center p-6" style={{ background: TOKENS.paper }}>
        <style>{FONT_CSS}</style>
        {err ? (
          <div className="w-full max-w-sm rounded-lg p-6 text-center" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
            <AlertCircle size={22} style={{ color: TOKENS.danger, margin: "0 auto" }} />
            <div className="text-[13px] font-medium mt-2" style={{ color: TOKENS.ink }}>Không tải được dữ liệu</div>
            <div className="text-[12px] mt-1.5 leading-relaxed" style={{ color: TOKENS.muted }}>{err}</div>
            <div className="flex justify-center gap-2 mt-4">
              <Btn onClick={reload}>Thử lại</Btn>
              <Btn kind="ghost" onClick={async () => { await supabase.auth.signOut(); setCurrentUser(null); }}>Đăng xuất</Btn>
            </div>
          </div>
        ) : (
          <div className="flex flex-col items-center gap-2">
            <Loader2 className="animate-spin" size={22} style={{ color: TOKENS.brand }} />
            <div className="text-[13px]" style={{ color: TOKENS.muted }}>Đang tải sổ tài sản…</div>
          </div>
        )}
      </div>
    );
  }

  const isAdmin = currentUser.role === "admin";
  const requireAdmin = () => {
    if (!isAdmin) { setToast("Bạn không có quyền thực hiện thao tác này"); return false; }
    return true;
  };

  const projectName = (id) => data.projects.find((e) => e.id === id)?.name || "—";
  const assetsById = Object.fromEntries(data.assets.map((a) => [a.id, a]));

  const logAction = (list, action) => [...list, { id: uid("lg"), date: nowIso(), user: currentUser.name, action }];

  const notify = (text) => setToast(text);

  /* ---------- mutations ---------- */

  const addAsset = (form) => {
    const asset = {
      id: uid("as"), code: form.code, name: form.name, category: form.category,
      cost: Number(form.cost) || 0, purchaseDate: form.purchaseDate, usefulLifeMonths: Number(form.usefulLifeMonths) || 36,
      status: STATUS.UNUSED, assignedTo: null, projectId: form.projectId || null, department: form.department, serial: form.serial,
      assetGroup: form.assetGroup || form.category, ownership: form.ownership || "Công ty", quantity: Number(form.quantity) || 1,
      supplier: form.supplier, warranty: !!form.warranty, warrantyEnd: form.warrantyEnd, note: form.note || "",
      customFields: { ...(form.customFields || {}) },
    };
    const tx = { id: uid("tx"), assetId: asset.id, type: "mua_sam", date: form.purchaseDate, title: "Mua sắm", detail: `Mua từ ${form.supplier || "nhà cung cấp chưa ghi"}`, amount: asset.cost };
    setData({
      ...data,
      assets: [asset, ...data.assets],
      transactions: [tx, ...data.transactions],
      activityLog: logAction(data.activityLog, `Thêm tài sản mới ${asset.code} — ${asset.name}`),
    });
    notify("Đã thêm tài sản");
  };

  const editAsset = (id, form) => {
    if (!requireAdmin()) return;
    setData({
      ...data,
      assets: data.assets.map((a) => (a.id === id ? { ...a, ...form, cost: Number(form.cost) || 0, usefulLifeMonths: Number(form.usefulLifeMonths) || 36 } : a)),
      activityLog: logAction(data.activityLog, `Cập nhật thông tin tài sản ${assetsById[id]?.code}`),
    });
    notify("Đã lưu thay đổi");
  };

  const deleteAsset = (id) => {
    if (!requireAdmin()) return;
    const asset = assetsById[id];
    if (!asset) return;
    if (!window.confirm(`Xoá tài sản ${asset.code} — ${asset.name}?`)) return;
    setData({ ...data, assets: data.assets.filter(a => a.id !== id), transactions: data.transactions.filter(t => t.assetId !== id), repairs: data.repairs.filter(r => r.assetId !== id), liquidations: data.liquidations.filter(r => r.assetId !== id), minutes: data.minutes.filter(m => m.assetId !== id), warehouse: data.warehouse.filter(w => w.assetId !== id), costHistory: data.costHistory.filter(c => c.assetId !== id), activityLog: logAction(data.activityLog, `Xoá tài sản ${asset.code}`) });
    setSelectedAssetId(null); notify("Đã xoá tài sản");
  };

  const addWarehouseTx = (form) => {
    const asset = assetsById[form.assetId];
    if (!asset) { notify("Vui lòng chọn tài sản"); return; }
    const qty = Number(form.quantity) || 0;
    if (qty <= 0) { notify("Số lượng phải lớn hơn 0"); return; }
    const unitCost = Number(form.unitCost) || 0;
    const currentRows = data.warehouse || [];
    const locationName = form.locationType === "project"
      ? (data.projects.find(p => p.id === form.projectId)?.name || "")
      : (form.warehouseName || "Kho trung tâm");
    if (!locationName) { notify("Vui lòng chọn công trình hoặc nhập tên kho"); return; }
    const prefix = form.type === "nhap" ? "PN" : "PX";
    const dateKey = String(form.date || nowIso().slice(0,10)).slice(0,10).replaceAll("-", "");
    const seq = currentRows.filter(w => String(w.voucherNo || "").startsWith(`${prefix}-${dateKey}-`)).length + 1;
    const voucherNo = form.voucherNo?.trim() || `${prefix}-${dateKey}-${String(seq).padStart(3,"0")}`;
    const tx = {
      id: uid("wh"), voucherNo, assetId: asset.id, type: form.type, quantity: qty, date: form.date,
      unitCost, total: qty * unitCost, unit: form.unit || "Cái", receiver: form.receiver || "", note: form.note || "",
      category: asset.category || "Khác", assetGroup: asset.assetGroup || "Thiết bị chính", ownership: asset.ownership || "Công ty",
      locationType: form.locationType || "project", locationName, warehouseName: form.locationType === "warehouse" ? locationName : "",
      projectId: form.projectId || null, itemName: asset.name, itemCode: asset.code,
    };
    if (form.type === "xuat") {
      const currentQty = currentRows.reduce((sum, w) => {
        if (w.assetId !== asset.id) return sum;
        const loc = w.locationName || (w.projectId ? projectName(w.projectId) : "Kho trung tâm");
        if (loc !== locationName || String(w.date || "") > String(form.date || "")) return sum;
        return sum + (w.type === "nhap" ? Number(w.quantity || 0) : -Number(w.quantity || 0));
      }, 0);
      if (qty > currentQty) { notify(`Không đủ tồn tại ${locationName}. ${asset.code} hiện còn ${currentQty} ${tx.unit}`); return; }
    }
    setData({ ...data, warehouse: [tx, ...currentRows], transactions: [{ id: uid("tx"), assetId: asset.id, type: form.type === "nhap" ? "nhap_kho" : "xuat_kho", date: form.date, title: `${form.type === "nhap" ? "Nhập kho" : "Xuất kho"} ${voucherNo}`, detail: `${asset.name} · ${locationName} · ${form.receiver || ""}`, amount: tx.total }, ...data.transactions], activityLog: logAction(data.activityLog, `${form.type === "nhap" ? "Nhập" : "Xuất"} ${voucherNo} — ${asset.code} — ${locationName}`) });
    notify(`${form.type === "nhap" ? "Đã lập phiếu nhập" : "Đã lập phiếu xuất"} ${voucherNo}`);
  };

  const importWarehouseExcel = (file) => {
    if (!requireAdmin()) return;
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
        if (!rows.length) return notify("File Excel không có dữ liệu");
        const byCode = Object.fromEntries(data.assets.map(a => [String(a.code).trim().toLowerCase(), a]));
        const byName = Object.fromEntries(data.assets.map(a => [String(a.name).trim().toLowerCase(), a]));
        const imported = [];
        const errors = [];
        const existing = data.warehouse || [];
        rows.forEach((r, i) => {
          const rawType = String(r["Loại phiếu"] || "Nhập kho").trim().toLowerCase();
          const type = rawType.includes("xuất") || rawType === "xuat" ? "xuat" : "nhap";
          const code = String(r["Mã hàng"] || "").trim().toLowerCase();
          const name = String(r["Tên tài sản"] || "").trim().toLowerCase();
          const asset = byCode[code] || byName[name];
          if (!asset) { errors.push(`Dòng ${i + 2}: không tìm thấy tài sản theo Mã hàng/Tên tài sản`); return; }
          const qty = Number(r["Số lượng"] || 0);
          if (qty <= 0) { errors.push(`Dòng ${i + 2}: số lượng không hợp lệ`); return; }
          const unitCost = Number(r["Đơn giá"] || 0);
          const date = String(r["Ngày tháng"] || nowIso().slice(0, 10)).slice(0, 10);
          const locationText = String(r["Kho/Công trình"] || "").trim();
          const project = data.projects.find(p => p.name.trim().toLowerCase() === locationText.toLowerCase());
          const locationType = String(r["Loại địa điểm"] || "").toLowerCase().includes("kho") ? "warehouse" : (project ? "project" : "warehouse");
          const locationName = project ? project.name : (locationText || "Kho trung tâm");
          const prefix = type === "nhap" ? "PN" : "PX";
          const dateKey = date.replaceAll("-", "");
          const seq = [...existing, ...imported].filter(w => String(w.voucherNo || "").startsWith(`${prefix}-${dateKey}-`)).length + 1;
          const voucherNo = String(r["Số phiếu"] || "").trim() || `${prefix}-${dateKey}-${String(seq).padStart(3,"0")}`;
          imported.push({ id: uid("wh"), voucherNo, assetId: asset.id, type, quantity: qty, date, unitCost, total: qty * unitCost,
            unit: String(r["Đơn vị tính"] || "Cái"), receiver: String(r["Người giao/nhận"] || ""), note: String(r["Ghi chú"] || ""),
            category: asset.category || "Khác", assetGroup: asset.assetGroup || "Thiết bị chính", ownership: asset.ownership || "Công ty",
            locationType, locationName, warehouseName: locationType === "warehouse" ? locationName : "", projectId: project?.id || null,
            itemName: asset.name, itemCode: asset.code });
        });
        if (errors.length) notify(`Import ${imported.length} dòng; ${errors.length} dòng lỗi. ${errors[0]}`);
        if (!imported.length) return;
        setData({ ...data, warehouse: [...imported, ...existing], transactions: [...imported.map(tx => ({ id: uid("tx"), assetId: tx.assetId, type: tx.type === "nhap" ? "nhap_kho" : "xuat_kho", date: tx.date, title: `${tx.type === "nhap" ? "Nhập kho" : "Xuất kho"} ${tx.voucherNo}`, detail: `${tx.itemName} · ${tx.locationName} · ${tx.receiver}`, amount: tx.total })), ...data.transactions], activityLog: logAction(data.activityLog, `Import ${imported.length} phiếu kho từ Excel`) });
        if (!errors.length) notify(`Đã import ${imported.length} phiếu kho`);
      } catch (err) { notify("Không đọc được file Excel phiếu kho"); }
    };
    reader.readAsArrayBuffer(file);
  };

  const addCostHistory = (form) => {
    const item = { id: uid("cp"), assetId: form.assetId, type: form.type, date: form.date, amount: Number(form.amount)||0, description: form.description || "", vendor: form.vendor || "" };
    setData({ ...data, costHistory: [item, ...data.costHistory], transactions: [{ id: uid("tx"), assetId: form.assetId, type: form.type, date: form.date, title: form.type, detail: form.description || "", amount: item.amount }, ...data.transactions], activityLog: logAction(data.activityLog, `Thêm chi phí ${form.type} cho ${assetsById[form.assetId]?.code || ""}`) });
    notify("Đã lưu chi phí thiết bị");
  };

  const assignAsset = (id, projectId, department) => {
    const asset = assetsById[id];
    const tx = { id: uid("tx"), assetId: id, type: "cap_phat", date: nowIso().slice(0, 10), title: "Cấp phát", detail: `Giao cho công trình ${projectName(projectId)}`, amount: 0 };
    setData({
      ...data,
      assets: data.assets.map((a) => (a.id === id ? { ...a, status: STATUS.ASSIGNED, assignedTo: null, projectId, department } : a)),
      transactions: [tx, ...data.transactions],
      activityLog: logAction(data.activityLog, `Giao ${asset.code} cho công trình ${projectName(projectId)}`),
    });
    notify("Đã cấp phát tài sản");
  };

  const transferAsset = (id, department) => {
    const asset = assetsById[id];
    const tx = { id: uid("tx"), assetId: id, type: "chuyen_bo_phan", date: nowIso().slice(0, 10), title: "Chuyển bộ phận", detail: `Chuyển sang ${department}`, amount: 0 };
    setData({
      ...data,
      assets: data.assets.map((a) => (a.id === id ? { ...a, department } : a)),
      transactions: [tx, ...data.transactions],
      activityLog: logAction(data.activityLog, `Chuyển bộ phận ${asset.code} sang ${department}`),
    });
    notify("Đã chuyển bộ phận");
  };

  const createRepair = (id, description, cost) => {
    const asset = assetsById[id];
    const rp = { id: uid("rp"), assetId: id, date: nowIso().slice(0, 10), description, cost: Number(cost) || 0, status: "Đang sửa" };
    const tx = { id: uid("tx"), assetId: id, type: "sua_chua", date: rp.date, title: "Gửi sửa chữa", detail: description, amount: rp.cost };
    setData({
      ...data,
      assets: data.assets.map((a) => (a.id === id ? { ...a, status: STATUS.REPAIR } : a)),
      repairs: [rp, ...data.repairs],
      transactions: [tx, ...data.transactions],
      activityLog: logAction(data.activityLog, `Tạo phiếu sửa chữa cho ${asset.code}`),
    });
    notify("Đã tạo phiếu sửa chữa");
  };

  const completeRepair = (repairId) => {
    const rp = data.repairs.find((r) => r.id === repairId);
    setData({
      ...data,
      repairs: data.repairs.map((r) => (r.id === repairId ? { ...r, status: "Hoàn thành", completeDate: nowIso().slice(0, 10) } : r)),
      assets: data.assets.map((a) => (a.id === rp.assetId ? { ...a, status: STATUS.UNUSED } : a)),
      activityLog: logAction(data.activityLog, `Hoàn thành sửa chữa ${assetsById[rp.assetId]?.code}`),
    });
    notify("Đã hoàn thành sửa chữa");
  };

  const liquidateAsset = (id, value, reason) => {
    if (!requireAdmin()) return;
    const asset = assetsById[id];
    const lq = { id: uid("lq"), assetId: id, date: nowIso().slice(0, 10), value: Number(value) || 0, reason };
    const tx = { id: uid("tx"), assetId: id, type: "thanh_ly", date: lq.date, title: "Thanh lý", detail: reason, amount: lq.value };
    setData({
      ...data,
      assets: data.assets.map((a) => (a.id === id ? { ...a, status: STATUS.LIQUIDATED } : a)),
      liquidations: [lq, ...data.liquidations],
      transactions: [tx, ...data.transactions],
      activityLog: logAction(data.activityLog, `Thanh lý tài sản ${asset.code}`),
    });
    notify("Đã thanh lý tài sản");
  };

  const createMinutes = (id, type, content) => {
    const asset = assetsById[id];
    const bb = { id: uid("bb"), assetId: id, type, date: nowIso().slice(0, 10), content, status: "Nháp" };
    setData({
      ...data,
      minutes: [bb, ...data.minutes],
      activityLog: logAction(data.activityLog, `Lập ${type.toLowerCase()} cho ${asset.code}`),
    });
    notify("Đã lập biên bản");
  };

  const addProject = (form) => {
    const project = { id: uid("p"), commander: form.commander, name: form.name, address: form.address, workItem: form.workItem, startDate: form.startDate, endDate: form.endDate };
    setData({ ...data, projects: [project, ...data.projects], activityLog: logAction(data.activityLog, `Thêm công trình ${project.name}`) });
    notify("Đã thêm công trình");
  };
  const deleteProject = (id) => { if (!requireAdmin()) return; const p = data.projects.find(x => x.id === id); if (data.assets.some(a => a.projectId === id)) { notify("Không thể xoá công trình đang có tài sản"); return; } if (!window.confirm(`Xoá công trình ${p?.name || id}?`)) return; setData({ ...data, projects: data.projects.filter(x => x.id !== id), activityLog: logAction(data.activityLog, `Xoá công trình ${p?.name || id}`) }); notify("Đã xoá công trình"); };

  /* ---------- settings ---------- */

  const setCompanyName = (name) => {
    if (!requireAdmin()) return;
    setData({ ...data, settings: { ...data.settings, companyName: name }, activityLog: logAction(data.activityLog, `Đổi tên công ty thành "${name}"`) });
    notify("Đã cập nhật tên công ty");
  };

  const addCategory = (name) => {
    if (!requireAdmin()) return;
    if (!name.trim() || data.settings.categories.includes(name.trim())) return;
    setData({ ...data, settings: { ...data.settings, categories: [...data.settings.categories, name.trim()] }, activityLog: logAction(data.activityLog, `Thêm danh mục "${name}"`) });
    notify("Đã thêm danh mục");
  };
  const removeCategory = (name) => {
    if (!requireAdmin()) return;
    setData({ ...data, settings: { ...data.settings, categories: data.settings.categories.filter((c) => c !== name) }, activityLog: logAction(data.activityLog, `Xoá danh mục "${name}"`) });
  };

  const addDepartment = (name) => {
    if (!requireAdmin()) return;
    if (!name.trim() || data.settings.departments.includes(name.trim())) return;
    setData({ ...data, settings: { ...data.settings, departments: [...data.settings.departments, name.trim()] }, activityLog: logAction(data.activityLog, `Thêm bộ phận "${name}"`) });
    notify("Đã thêm bộ phận");
  };
  const removeDepartment = (name) => {
    if (!requireAdmin()) return;
    setData({ ...data, settings: { ...data.settings, departments: data.settings.departments.filter((d) => d !== name) }, activityLog: logAction(data.activityLog, `Xoá bộ phận "${name}"`) });
  };

  const addCustomColumn = (label, type) => {
    if (!requireAdmin()) return;
    if (!label.trim()) return;
    const key = "cf_" + label.trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "") + "_" + uid("").slice(-4);
    setData({
      ...data,
      settings: { ...data.settings, customColumns: [...data.settings.customColumns, { key, label: label.trim(), type }] },
      activityLog: logAction(data.activityLog, `Thêm cột tuỳ chỉnh "${label}"`),
    });
    notify("Đã thêm cột mới");
  };
  const removeCustomColumn = (key) => {
    if (!requireAdmin()) return;
    setData({
      ...data,
      settings: { ...data.settings, customColumns: data.settings.customColumns.filter((c) => c.key !== key) },
      activityLog: logAction(data.activityLog, `Xoá cột tuỳ chỉnh`),
    });
  };

  /* ---------- users (Supabase Auth + profiles table) ---------- */

  const setUserRole = async (profileId, role) => {
    if (!requireAdmin()) return;
    const { error } = await supabase.from("profiles").update({ role }).eq("id", profileId);
    if (error) { notify("Đổi quyền thất bại"); return; }
    setProfiles((ps) => ps.map((p) => (p.id === profileId ? { ...p, role } : p)));
    setData({ ...data, activityLog: logAction(data.activityLog, `Đổi quyền một tài khoản thành "${ROLES[role].label}"`) });
    notify("Đã cập nhật quyền");
  };

  const changeOwnPassword = async (oldPw, newPw) => {
    const { error: reauthErr } = await supabase.auth.signInWithPassword({ email: currentUser.email, password: oldPw });
    if (reauthErr) { notify("Mật khẩu hiện tại không đúng"); return false; }
    const { error } = await supabase.auth.updateUser({ password: newPw });
    if (error) { notify("Đổi mật khẩu thất bại: " + error.message); return false; }
    setData({ ...data, activityLog: logAction(data.activityLog, `Đổi mật khẩu của "${currentUser.name}"`) });
    notify("Đã đổi mật khẩu");
    return true;
  };

  const sendPasswordReset = async (email) => {
    if (!requireAdmin()) return;
    const { error } = await supabase.auth.resetPasswordForEmail(email, { redirectTo: window.location.origin });
    if (error) { notify("Gửi email thất bại"); return; }
    notify(`Đã gửi email đặt lại mật khẩu tới ${email}`);
  };

  /* ---------- backup & restore (server-side, Supabase table) ---------- */

  const createBackup = async (label, { silent } = {}) => {
    if (currentUser.role !== "admin") { if (!silent) setToast("Bạn không có quyền thực hiện thao tác này"); return; }
    try {
      const { error } = await supabase.from("backups").insert({
        label: label || "Thủ công", created_by: currentUser.name, data: snapshotOf(data),
      });
      if (error) throw error;
      await loadBackups();
      // dọn bớt bản cũ nếu vượt quá giới hạn
      const { data: all } = await supabase.from("backups").select("id, created_at").order("created_at", { ascending: false });
      const toDelete = (all || []).slice(MAX_BACKUPS);
      if (toDelete.length) await supabase.from("backups").delete().in("id", toDelete.map((b) => b.id));
      await loadBackups();
      if (!silent) {
        setData({ ...data, activityLog: logAction(data.activityLog, `Tạo bản sao lưu máy chủ${label ? ` (${label})` : ""}`) });
        notify("Đã sao lưu vào máy chủ");
      }
    } catch (e) {
      if (!silent) notify("Sao lưu thất bại — thử lại sau");
    }
  };

  const restoreBackup = async (id) => {
    if (!requireAdmin()) return;
    if (!window.confirm("Khôi phục sẽ GHI ĐÈ toàn bộ dữ liệu hiện tại bằng bản sao lưu này. Bạn chắc chắn chứ?")) return;
    try {
      const { data: row, error } = await supabase.from("backups").select("data").eq("id", id).single();
      if (error || !row) { notify("Không tìm thấy bản sao lưu này trên máy chủ"); return; }
      const restored = withDefaults(row.data);
      restored.activityLog = logAction(restored.activityLog, "Khôi phục dữ liệu từ bản sao lưu máy chủ");
      await setData(restored);
      notify("Đã khôi phục dữ liệu từ máy chủ");
    } catch (e) {
      notify("Khôi phục thất bại — dữ liệu sao lưu có thể bị lỗi");
    }
  };

  const deleteBackup = async (id) => {
    if (!requireAdmin()) return;
    try { await supabase.from("backups").delete().eq("id", id); } catch (e) { /* ignore */ }
    await loadBackups();
    setData({ ...data, activityLog: logAction(data.activityLog, "Xoá một bản sao lưu máy chủ") });
    notify("Đã xoá bản sao lưu");
  };

  const downloadBackupFile = () => {
    try {
      const snapshot = JSON.stringify(snapshotOf(data), null, 2);
      const blob = new Blob([snapshot], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `so-tai-san-backup-${nowIso().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      notify("Đã tải file sao lưu về máy");
    } catch (e) {
      notify("Tải file thất bại");
    }
  };

  const restoreFromFile = (file) => {
    if (!requireAdmin()) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const snapshot = JSON.parse(e.target.result);
        if (!snapshot.assets || !Array.isArray(snapshot.assets)) throw new Error("invalid shape");
        if (!window.confirm("Khôi phục từ file sẽ GHI ĐÈ toàn bộ dữ liệu hiện tại. Bạn chắc chắn chứ?")) return;
        const restored = withDefaults(snapshot);
        restored.activityLog = logAction(restored.activityLog, `Khôi phục dữ liệu từ file tải lên (${file.name})`);
        setData(restored);
        notify("Đã khôi phục dữ liệu từ file");
      } catch (err) {
        notify("File không hợp lệ hoặc bị hỏng");
      }
    };
    reader.readAsText(file);
  };

  const importExcel = (file, kind) => {
    if (!requireAdmin()) return;
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
        if (!rows.length) return notify("File Excel không có dữ liệu");
        if (kind === "assets") {
          const imported = rows.map((r,i)=>({ id: uid("as"), code: String(r["Mã quản lý"]||r.code||`IMP-${Date.now()}-${i}`), name: String(r["Tên tài sản"]||r.name||""), category: String(r["Loại"]||r.category||"Khác"), assetGroup: String(r["Nhóm tài sản"]||r.assetGroup||r["Loại"]||"Thiết bị chính"), ownership: String(r["Nguồn"]||r.ownership||"Công ty"), cost:Number(r["Nguyên giá"]||r.cost||0), purchaseDate:String(r["Ngày mua"]||r.purchaseDate||nowIso().slice(0,10)).slice(0,10), usefulLifeMonths:Number(r["Thời gian SD"]||r.usefulLifeMonths||36), status:STATUS.UNUSED, assignedTo:null, projectId:null, department:String(r["Bộ phận"]||"Vận hành"), serial:String(r["Serial"]||""), supplier:String(r["Nhà cung cấp"]||""), warranty:false, warrantyEnd:"", note:String(r["Ghi chú"]||""), quantity:Number(r["Số lượng"]||r.quantity||1), customFields:{} }));
          const wh = imported.map(a=>({id:uid("wh"),assetId:a.id,type:"nhap",quantity:a.quantity||1,date:a.purchaseDate,unitCost:a.cost,note:"Nhập từ Excel"}));
          setData({...data, assets:[...imported,...data.assets], warehouse:[...wh,...data.warehouse], activityLog:logAction(data.activityLog,`Import ${imported.length} tài sản từ Excel`)}); notify(`Đã import ${imported.length} tài sản`);
        } else if (kind === "projects") {
          const imported=rows.map(r=>({id:uid("p"),commander:String(r["Chỉ huy trưởng"]||r.commander||""),name:String(r["Tên công trình"]||r.name||""),address:String(r["Địa chỉ"]||r.address||""),workItem:String(r["Hạng mục thi công"]||r.workItem||""),startDate:String(r["Ngày bắt đầu"]||r.startDate||"").slice(0,10),endDate:String(r["Ngày kết thúc"]||r.endDate||"").slice(0,10)}));
          setData({...data,projects:[...imported,...data.projects],activityLog:logAction(data.activityLog,`Import ${imported.length} công trình từ Excel`)}); notify(`Đã import ${imported.length} công trình`);
        }
      } catch(err){ notify("Không đọc được file Excel"); }
    }; reader.readAsArrayBuffer(file);
  };

  /* ---------- export ---------- */

  const doExportExcel = (filename, headers, rows) => {
    try { exportExcel(filename, headers, rows); notify("Đã xuất file Excel"); }
    catch (e) { notify("Xuất Excel thất bại"); }
  };
  const doExportPdf = (title, headers, rows) => {
    setPrintJob({ title, headers, rows: rows.map((r) => r.map((c) => (c == null ? "" : String(c)))) });
  };

  /* ---------- derived ---------- */

  const selectedAsset = selectedAssetId ? assetsById[selectedAssetId] : null;

  /* ---------- render ---------- */

  const settings = data.settings;

  return (
    <>
      <div className="aa-root aa-noprint min-h-screen flex" style={{ background: TOKENS.paper }}>
        <style>{FONT_CSS}</style>
        <style>{PRINT_CSS}</style>

        <Sidebar active={active} setActive={(id) => { setActive(id); setSelectedAssetId(null); }} data={data} counts={counts}
          currentUser={currentUser} isAdmin={isAdmin} onLogout={async () => { await supabase.auth.signOut(); setCurrentUser(null); }} onChangePassword={() => setModal({ type: "changePassword" })} />

        <div className="flex-1 min-w-0 flex flex-col">
          <TopBar query={query} setQuery={setQuery} data={data} saving={saving} onRefresh={reload} />

          {err && (
            <div className="px-6 pt-3">
              <div className="flex items-center gap-2 text-[12px] rounded-md px-3 py-2" style={{ background: TOKENS.dangerSoft, color: TOKENS.danger }}>
                <AlertCircle size={14} /> {err}
              </div>
            </div>
          )}

          <div className="flex-1 min-w-0 flex">
            <div className="flex-1 min-w-0 p-6 overflow-y-auto aa-scroll">
              {active === "overview" && <Overview data={data} counts={counts} projectName={projectName} onNavigate={setActive} />}
              {active === "catalog" && (
                <AssetCatalog
                  assets={filteredAssets} projectName={projectName} onSelect={setSelectedAssetId} selectedAssetId={selectedAssetId}
                  onAdd={() => setModal({ type: "addAsset" })} customColumns={settings.customColumns} isAdmin={isAdmin} onDelete={deleteAsset}
                  onExportExcel={doExportExcel} onExportPdf={doExportPdf}
                />
              )}
              {active === "byProject" && <ByProject data={data} projectName={projectName} onSelect={(id) => { setActive("catalog"); setSelectedAssetId(id); }} />}
              {active === "depreciation" && <Depreciation assets={data.assets} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "repair" && <RepairView repairs={data.repairs.filter((r) => r.status === "Đang sửa")} assetsById={assetsById} onComplete={completeRepair} />}
              {active === "repairHistory" && <RepairHistory repairs={data.repairs} assetsById={assetsById} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "liquidation" && <LiquidationView liquidations={data.liquidations} assetsById={assetsById} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "minutes" && <MinutesView minutes={data.minutes} assetsById={assetsById} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "transactions" && <TransactionsView transactions={data.transactions} assetsById={assetsById} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
{active === "warehouse" && <WarehouseHub
                warehouse={data.warehouse || []} assets={data.assets} projects={data.projects} settings={settings}
                onAdd={(type) => setModal({ type: type === "nhap" ? "warehouseIn" : "warehouseOut" })}
                onExportExcel={doExportExcel} onExportPdf={doExportPdf}
                onImport={importWarehouseExcel} isAdmin={isAdmin}
              />}
              {active === "costHistory" && <CostHistoryView costHistory={data.costHistory} assetsById={assetsById} onAdd={() => setModal({ type: "costHistory" })} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "projects" && <ProjectsView projects={data.projects} assets={data.assets} isAdmin={isAdmin} onDelete={deleteProject} onAdd={() => setModal({ type: "addProject" })} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "activityLog" && <ActivityLogView log={data.activityLog} />}
              {active === "settings" && isAdmin && (
                <SettingsView
                  settings={settings} onSetCompanyName={setCompanyName}
                  onAddCategory={addCategory} onRemoveCategory={removeCategory}
                  onAddDepartment={addDepartment} onRemoveDepartment={removeDepartment}
                  onAddColumn={addCustomColumn} onRemoveColumn={removeCustomColumn}
                  categoryUsage={(name) => data.assets.filter((a) => a.category === name).length}
                  deptUsage={(name) => data.assets.filter((a) => a.department === name).length}
                  users={profiles} currentUser={currentUser} onSetUserRole={setUserRole} onSendPasswordReset={sendPasswordReset}
                  backups={backups}
                  onCreateBackup={createBackup} onRestoreBackup={restoreBackup} onDeleteBackup={deleteBackup}
                  onDownloadBackupFile={downloadBackupFile} onRestoreFromFile={restoreFromFile} onImportExcel={(file, kind) => importExcel(file, kind)}
                />
              )}
              {active === "settings" && !isAdmin && (
                <div className="aa-fade max-w-md">
                  <div className="rounded-lg p-6 text-center" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
                    <ShieldAlert size={22} style={{ color: TOKENS.danger, margin: "0 auto" }} />
                    <div className="text-[13px] font-medium mt-2">Bạn không có quyền truy cập Cài đặt</div>
                    <div className="text-[12px] mt-1" style={{ color: TOKENS.muted }}>Liên hệ tài khoản Toàn quyền để được hỗ trợ.</div>
                  </div>
                </div>
              )}
            </div>

            {selectedAsset && active === "catalog" && (
              <AssetDetail
                asset={selectedAsset} data={data} projectName={projectName} isAdmin={isAdmin} onDelete={deleteAsset} onClose={() => setSelectedAssetId(null)}
                onAction={(type) => setModal({ type, assetId: selectedAsset.id })}
              />
            )}
          </div>
        </div>

        {modal?.type === "addAsset" && (
          <AssetFormModal title="Thêm tài sản mới" categories={settings.categories} departments={settings.departments} projects={data.projects} customColumns={settings.customColumns}
            onClose={() => setModal(null)} onSubmit={(f) => { addAsset(f); setModal(null); }} />
        )}
        {modal?.type === "editAsset" && (
          <AssetFormModal title={`Sửa tài sản — ${assetsById[modal.assetId]?.code}`} initial={assetsById[modal.assetId]}
            categories={settings.categories} departments={settings.departments} projects={data.projects} customColumns={settings.customColumns}
            onClose={() => setModal(null)} onSubmit={(f) => { editAsset(modal.assetId, f); setModal(null); }} />
        )}
        {modal?.type === "assign" && (
          <AssignModal asset={assetsById[modal.assetId]} projects={data.projects} departments={settings.departments} onClose={() => setModal(null)}
            onSubmit={(projectId, dept) => { assignAsset(modal.assetId, projectId, dept); setModal(null); }} />
        )}
        {modal?.type === "transfer" && (
          <TransferModal asset={assetsById[modal.assetId]} departments={settings.departments} onClose={() => setModal(null)}
            onSubmit={(dept) => { transferAsset(modal.assetId, dept); setModal(null); }} />
        )}
        {modal?.type === "repair" && (
          <RepairModal asset={assetsById[modal.assetId]} onClose={() => setModal(null)}
            onSubmit={(desc, cost) => { createRepair(modal.assetId, desc, cost); setModal(null); }} />
        )}
        {modal?.type === "liquidate" && (
          <LiquidateModal asset={assetsById[modal.assetId]} onClose={() => setModal(null)}
            onSubmit={(value, reason) => { liquidateAsset(modal.assetId, value, reason); setModal(null); }} />
        )}
        {modal?.type === "minutes" && (
          <MinutesModal asset={assetsById[modal.assetId]} onClose={() => setModal(null)}
            onSubmit={(type, content) => { createMinutes(modal.assetId, type, content); setModal(null); }} />
        )}
        {modal?.type === "addProject" && (
          <ProjectFormModal onClose={() => setModal(null)} onSubmit={(f) => { addProject(f); setModal(null); }} />
        )}
        {(modal?.type === "warehouseIn" || modal?.type === "warehouseOut") && <WarehouseTxModal
          title={modal.type === "warehouseIn" ? "Lập phiếu nhập kho" : "Lập phiếu xuất kho"}
          fixedType={modal.type === "warehouseIn" ? "nhap" : "xuat"}
          assets={data.assets} projects={data.projects} onClose={()=>setModal(null)} onSubmit={f=>{addWarehouseTx(f);setModal(null)}}
        />}
        {modal?.type === "costHistory" && <CostHistoryModal assets={data.assets} onClose={()=>setModal(null)} onSubmit={f=>{addCostHistory(f);setModal(null)}} />}
        {modal?.type === "changePassword" && (
          <ChangePasswordModal onClose={() => setModal(null)} onSubmit={async (oldPw, newPw) => { const ok = await changeOwnPassword(oldPw, newPw); if (ok) setModal(null); }} />
        )}

        {toast && (
          <div className="fixed bottom-5 left-1/2 -translate-x-1/2 z-50 aa-fade px-4 py-2 rounded-md text-[13px] flex items-center gap-2"
            style={{ background: TOKENS.ink, color: "#fff" }}>
            <ClipboardCheck size={14} /> {toast}
          </div>
        )}
      </div>

      {printJob && <PrintArea job={printJob} companyName={settings.companyName} />}
    </>
  );
}

const PRINT_CSS = `
#aa-print-area{ position:fixed; left:-99999px; top:0; }
@media print {
  .aa-noprint{ display:none !important; }
  #aa-print-area{ position:static !important; left:auto !important; }
}
`;

function PrintArea({ job, companyName }) {
  return (
    <div id="aa-print-area" style={{ background: "#fff", color: "#111", padding: 28, width: 1000, fontFamily: "IBM Plex Sans, sans-serif" }}>
      <div style={{ fontSize: 11, color: "#666", marginBottom: 2 }}>{companyName}</div>
      <div style={{ fontFamily: "Space Grotesk, sans-serif", fontWeight: 700, fontSize: 19, marginBottom: 2 }}>{job.title}</div>
      <div style={{ fontSize: 11, color: "#666", marginBottom: 14 }}>Xuất ngày {new Date().toLocaleString("vi-VN")}</div>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11.5 }}>
        <thead>
          <tr>{job.headers.map((h) => <th key={h} style={{ border: "1px solid #999", padding: "5px 7px", textAlign: "left", background: "#eee" }}>{h}</th>)}</tr>
        </thead>
        <tbody>
          {job.rows.map((r, i) => (
            <tr key={i}>{r.map((c, j) => <td key={j} style={{ border: "1px solid #ccc", padding: "5px 7px" }}>{c}</td>)}</tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* ============================== SIDEBAR / TOPBAR ============================== */

function Sidebar({ active, setActive, data, counts, currentUser, isAdmin, onLogout, onChangePassword }) {
  return (
    <div className="w-[240px] shrink-0 flex flex-col" style={{ background: TOKENS.surface, color: TOKENS.ink, borderRight: `1px solid ${TOKENS.border}` }}>
      <div className="px-5 pt-5 pb-4 flex items-center gap-2.5" style={{ borderBottom: `1px solid ${TOKENS.border}` }}>
        <div className="w-9 h-9 rounded-lg flex items-center justify-center shrink-0" style={{ background: TOKENS.brand }}>
          <Stamp size={17} color="#fff" />
        </div>
        <div className="min-w-0">
          <div className="aa-display font-bold text-[15px] leading-tight" style={{ color: TOKENS.brand, letterSpacing: "0.02em" }}>SỔ TÀI SẢN</div>
          <div className="text-[11px] truncate" style={{ color: TOKENS.muted }} title={data.settings.companyName}>{data.settings.companyName}</div>
        </div>
      </div>

      <div className="flex-1 overflow-y-auto aa-scroll py-3">
        {NAV.map((sec) => (
          <div key={sec.section} className="mb-4">
            <div className="px-5 mb-1.5 text-[10.5px] tracking-wider font-semibold" style={{ color: "#C9B4B3" }}>{sec.section}</div>
            {sec.items.filter((it) => !it.adminOnly || isAdmin).map((it) => {
              const isActive = active === it.id;
              const count = it.badge === "assets" ? counts.total : it.badge === "projects" ? data.projects.length : null;
              return (
                <button
                  key={it.id}
                  onClick={() => setActive(it.id)}
                  className="w-full flex items-center gap-2.5 px-5 py-2 text-[13px] relative transition-colors"
                  style={{
                    background: isActive ? TOKENS.brandSoft : "transparent",
                    color: isActive ? TOKENS.brand : "#5C5150",
                    fontWeight: isActive ? 600 : 400,
                  }}
                >
                  {isActive && <span style={{ position: "absolute", left: 0, top: 6, bottom: 6, width: 3, background: TOKENS.brand, borderRadius: 2 }} />}
                  <it.icon size={15} />
                  <span className="flex-1 text-left">{it.label}</span>
                  {count != null && <span className="aa-mono text-[11px]" style={{ color: isActive ? TOKENS.brand : TOKENS.muted }}>{count}</span>}
                </button>
              );
            })}
          </div>
        ))}
      </div>

      <div className="px-5 py-3" style={{ borderTop: `1px solid ${TOKENS.border}` }}>
        <div className="flex items-center gap-2.5 mb-2.5">
          <div className="w-7 h-7 rounded-full flex items-center justify-center aa-display font-semibold text-[11px] shrink-0"
            style={{ background: currentUser.role === "admin" ? TOKENS.brand : TOKENS.info, color: "#fff" }}>
            {currentUser.name[0]?.toUpperCase()}
          </div>
          <div className="min-w-0">
            <div className="text-[12.5px] font-medium truncate" style={{ color: TOKENS.ink }}>{currentUser.name}</div>
            <div className="text-[10.5px] truncate" style={{ color: currentUser.role === "admin" ? TOKENS.brand : TOKENS.info }}>{ROLES[currentUser.role].label}</div>
          </div>
          <button onClick={onChangePassword} title="Đổi mật khẩu" className="ml-auto p-1.5 rounded hover:bg-black/5 shrink-0"><Lock size={13} /></button>
          <button onClick={onLogout} title="Đăng xuất" className="p-1.5 rounded hover:bg-black/5 shrink-0"><LogOut size={14} /></button>
        </div>
        <div className="text-[10px]" style={{ color: TOKENS.muted }}>
          Dữ liệu dùng chung — mọi thành viên đều thấy cùng một sổ tài sản.
        </div>
      </div>
    </div>
  );
}

function TopBar({ query, setQuery, data, saving, onRefresh }) {
  return (
    <div className="flex items-center justify-between px-6 py-3.5 shrink-0" style={{ background: TOKENS.surface, borderBottom: `1px solid ${TOKENS.border}` }}>
      <div className="relative w-[340px] max-w-full">
        <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2" style={{ color: TOKENS.muted }} />
        <input
          value={query} onChange={(e) => setQuery(e.target.value)}
          placeholder="Tìm mã, tên, serial, người dùng…"
          className="w-full rounded-md pl-8 pr-2.5 py-1.5 text-[13px]"
          style={{ border: `1px solid ${TOKENS.border}`, background: TOKENS.paper }}
        />
      </div>
      <div className="flex items-center gap-3">
        <span className="text-[11px]" style={{ color: TOKENS.muted }}>
          {saving ? "Đang lưu…" : "Đã lưu"}
        </span>
        <Btn icon={RefreshCw} onClick={onRefresh}>Làm mới</Btn>
      </div>
    </div>
  );
}

/* ============================== OVERVIEW ============================== */

/* ---------- lightweight SVG charts (no extra dependency) ---------- */

const CHART_PALETTE = [TOKENS.brand, "#E07A5F", TOKENS.gold, TOKENS.info, "#8B5CF6", "#4C956C", "#7C6E6D"];

function DonutChart({ segments, size = 168, thickness = 22 }) {
  const total = segments.reduce((s, x) => s + x.value, 0) || 1;
  const r = (size - thickness) / 2;
  const cx = size / 2, cy = size / 2;
  const circumference = 2 * Math.PI * r;
  let offset = 0;
  return (
    <div className="flex items-center gap-5">
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
        <circle cx={cx} cy={cy} r={r} fill="none" stroke={TOKENS.paper} strokeWidth={thickness} />
        {segments.filter((s) => s.value > 0).map((s, i) => {
          const frac = s.value / total;
          const dash = frac * circumference;
          const el = (
            <circle
              key={s.label}
              cx={cx} cy={cy} r={r} fill="none"
              stroke={s.color}
              strokeWidth={thickness}
              strokeDasharray={`${dash} ${circumference - dash}`}
              strokeDashoffset={-offset}
              transform={`rotate(-90 ${cx} ${cy})`}
              strokeLinecap={segments.filter((x) => x.value > 0).length === 1 ? "butt" : "butt"}
            />
          );
          offset += dash;
          return el;
        })}
        <text x={cx} y={cy - 3} textAnchor="middle" className="aa-display" style={{ fontSize: 20, fontWeight: 700, fill: TOKENS.ink }}>{total}</text>
        <text x={cx} y={cy + 14} textAnchor="middle" style={{ fontSize: 10, fill: TOKENS.muted }}>thiết bị</text>
      </svg>
      <div className="space-y-1.5 min-w-0">
        {segments.filter((s) => s.value > 0).map((s) => (
          <div key={s.label} className="flex items-center gap-1.5 text-[12px]">
            <span style={{ width: 8, height: 8, borderRadius: 2, background: s.color, display: "inline-block", flexShrink: 0 }} />
            <span className="truncate" style={{ color: TOKENS.ink }}>{s.label}</span>
            <span className="aa-mono shrink-0" style={{ color: TOKENS.muted }}>{Math.round((s.value / total) * 100)}%</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function BarChart({ bars, height = 168 }) {
  const max = Math.max(1, ...bars.map((b) => b.value));
  return (
    <div className="flex items-end gap-3" style={{ height }}>
      {bars.map((b) => (
        <div key={b.label} className="flex-1 flex flex-col items-center gap-1.5 min-w-0">
          <div className="text-[11px] aa-mono" style={{ color: TOKENS.muted }}>{b.value}</div>
          <div className="w-full rounded-t-md" style={{ height: Math.max(4, (b.value / max) * (height - 40)), background: b.color || TOKENS.brand }} />
          <div className="text-[10.5px] text-center truncate w-full" style={{ color: TOKENS.muted }} title={b.label}>{b.label}</div>
        </div>
      ))}
    </div>
  );
}

function daysUntil(dateStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr);
  if (isNaN(d)) return null;
  return Math.ceil((d - new Date()) / (1000 * 60 * 60 * 24));
}

function MaintenanceBadge({ days }) {
  if (days == null) return <span style={{ color: TOKENS.muted }}>—</span>;
  if (days < 0) return <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: TOKENS.dangerSoft, color: TOKENS.danger }}>Quá hạn</span>;
  if (days <= 30) return <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: TOKENS.goldSoft, color: TOKENS.gold }}>Sắp đến hạn</span>;
  return <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: TOKENS.brandSoft, color: TOKENS.brand }}>Còn hạn</span>;
}

function QuickActionCard({ icon: Icon, label, onClick }) {
  return (
    <button onClick={onClick} className="flex flex-col gap-2.5 rounded-lg p-4 text-left hover:shadow-sm transition-shadow"
      style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
      <div className="w-9 h-9 rounded-lg flex items-center justify-center" style={{ background: TOKENS.brandSoft }}>
        <Icon size={16} style={{ color: TOKENS.brand }} />
      </div>
      <div className="text-[12.5px] font-medium" style={{ color: TOKENS.ink }}>{label}</div>
      <div className="flex items-center gap-1 text-[11px]" style={{ color: TOKENS.brand }}>
        Đi tới <ChevronRight size={12} />
      </div>
    </button>
  );
}

function Overview({ data, counts, projectName, onNavigate }) {
  const breakdown = [
    { label: STATUS.ASSIGNED, value: counts[STATUS.ASSIGNED] },
    { label: STATUS.SHARED, value: counts[STATUS.SHARED] },
    { label: STATUS.UNUSED, value: counts[STATUS.UNUSED] },
    { label: STATUS.REPAIR, value: counts[STATUS.REPAIR] },
    { label: STATUS.BROKEN, value: counts[STATUS.BROKEN] },
    { label: STATUS.LIQUIDATED, value: counts[STATUS.LIQUIDATED] },
    { label: STATUS.TRANSFERRED_OUT, value: counts[STATUS.TRANSFERRED_OUT] },
  ];
  const max = Math.max(1, ...breakdown.map((b) => b.value));

  const maintenanceDue = data.assets
    .filter((a) => a.warranty && a.warrantyEnd)
    .map((a) => ({ ...a, daysLeft: daysUntil(a.warrantyEnd) }))
    .filter((a) => a.daysLeft != null && a.daysLeft <= 60)
    .sort((a, b) => a.daysLeft - b.daysLeft)
    .slice(0, 6);

  const recentTx = data.transactions.filter((t) => t.type === "xuat_kho" || t.type === "nhap_kho").slice(0, 5);

  const quickActions = [
    { icon: PackageMinus, label: "Phiếu xuất kho", target: "warehouse" },
    { icon: PackagePlus, label: "Phiếu nhập kho", target: "warehouse" },
    { icon: FileSpreadsheet, label: "Báo cáo nhập xuất tồn", target: "warehouse" },
    { icon: HardDrive, label: "Chi phí thiết bị", target: "depreciation" },
    { icon: Boxes, label: "Danh mục tài sản", target: "catalog" },
    { icon: History, label: "Khấu hao tài sản", target: "depreciation" },
  ];

  return (
    <div className="aa-fade max-w-6xl">
      <div className="flex items-center justify-between mb-4">
        <h1 className="aa-display text-xl font-semibold">Tổng quan</h1>
        <div className="flex items-center gap-2 rounded-full px-3 py-1" style={{ background: TOKENS.brandSoft, color: TOKENS.brand }}>
          <Stamp size={14} />
          <span className="aa-display text-[13px] font-semibold">{counts.total} tài sản</span>
        </div>
      </div>

      {/* Hàng thẻ số liệu chính */}
      <div className="grid grid-cols-3 gap-4 mb-4">
        <div className="rounded-lg p-5 flex items-center gap-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
          <div className="w-11 h-11 rounded-lg flex items-center justify-center shrink-0" style={{ background: TOKENS.brandSoft }}>
            <Boxes size={19} style={{ color: TOKENS.brand }} />
          </div>
          <div>
            <div className="text-[12px]" style={{ color: TOKENS.muted }}>Tổng tài sản</div>
            <div className="aa-display text-xl font-bold">{counts.total}</div>
          </div>
        </div>
        <div className="rounded-lg p-5 flex items-center gap-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
          <div className="w-11 h-11 rounded-lg flex items-center justify-center shrink-0" style={{ background: TOKENS.goldSoft }}>
            <HardDrive size={19} style={{ color: TOKENS.gold }} />
          </div>
          <div>
            <div className="text-[12px]" style={{ color: TOKENS.muted }}>Tổng nguyên giá</div>
            <div className="aa-display text-xl font-bold">{fmtVND(data.assets.reduce((s, a) => s + (a.cost || 0), 0))}</div>
          </div>
        </div>
        <div className="rounded-lg p-5 flex items-center gap-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
          <div className="w-11 h-11 rounded-lg flex items-center justify-center shrink-0" style={{ background: TOKENS.infoSoft }}>
            <HardHat size={19} style={{ color: TOKENS.info }} />
          </div>
          <div>
            <div className="text-[12px]" style={{ color: TOKENS.muted }}>Đang tại công trình</div>
            <div className="aa-display text-xl font-bold">{counts[STATUS.ASSIGNED]}</div>
          </div>
        </div>
      </div>

      {/* Bố cục chính: cột trái (biểu đồ + bảo trì), cột phải (kho) */}
      <div className="grid gap-4 mb-4" style={{ gridTemplateColumns: "1.6fr 1fr" }}>
        <div className="space-y-4">
          <div className="grid grid-cols-2 gap-4">
            <div className="rounded-lg p-5" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
              <div className="text-[13px] font-medium mb-4">Tài sản theo loại</div>
              <DonutChart
                segments={Object.entries(
                  data.assets.reduce((acc, a) => { acc[a.category] = (acc[a.category] || 0) + 1; return acc; }, {})
                ).map(([label, value], i) => ({ label, value, color: CHART_PALETTE[i % CHART_PALETTE.length] }))}
              />
            </div>
            <div className="rounded-lg p-5" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
              <div className="text-[13px] font-medium mb-4">Tài sản theo công trình</div>
              <BarChart
                bars={data.projects.map((p, i) => ({
                  label: p.name, value: data.assets.filter((a) => a.assignedTo === p.id).length, color: CHART_PALETTE[i % CHART_PALETTE.length],
                }))}
              />
              {data.projects.length === 0 && <EmptyState text="Chưa có công trình nào" />}
            </div>
          </div>

          <div className="rounded-lg overflow-hidden" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
            <div className="flex items-center justify-between px-5 pt-4 pb-1">
              <div className="text-[13px] font-medium">Tài sản sắp hết hạn bảo trì</div>
              <button onClick={() => onNavigate("catalog")} className="text-[11.5px]" style={{ color: TOKENS.brand }}>Xem tất cả</button>
            </div>
            <table className="w-full mt-2">
              <thead><tr><Th>Mã tài sản</Th><Th>Tên tài sản</Th><Th>Loại</Th><Th>Hạn bảo trì</Th><Th right>Trạng thái</Th></tr></thead>
              <tbody>
                {maintenanceDue.map((a) => (
                  <tr key={a.id} className="aa-row">
                    <Td mono><Tag>{a.code}</Tag></Td><Td>{a.name}</Td><Td>{a.category}</Td>
                    <Td mono>{fmtDate(a.warrantyEnd)}</Td>
                    <Td right><MaintenanceBadge days={a.daysLeft} /></Td>
                  </tr>
                ))}
              </tbody>
            </table>
            {maintenanceDue.length === 0 && <EmptyState text="Không có tài sản nào sắp hết hạn bảo trì" sub="Trong vòng 60 ngày tới." />}
          </div>
        </div>

        <div className="space-y-4">
          <div className="rounded-lg p-5" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
            <div className="text-[13px] font-medium mb-1">Kho — thao tác nhanh</div>
            <div className="text-[11.5px] mb-3" style={{ color: TOKENS.muted }}>Lập phiếu nhập/xuất kho hoặc xem báo cáo tồn.</div>
            <div className="grid grid-cols-2 gap-2 mb-4">
              <Btn kind="primary" icon={PackageMinus} onClick={() => onNavigate("warehouse")}>Xuất kho</Btn>
              <Btn kind="info" icon={PackagePlus} onClick={() => onNavigate("warehouse")}>Nhập kho</Btn>
            </div>
            <div className="text-[12px] font-medium mb-2" style={{ color: TOKENS.ink }}>Giao dịch kho gần đây</div>
            <div className="space-y-2">
              {recentTx.map((t) => (
                <div key={t.id} className="flex items-start gap-2 text-[12px]">
                  <span style={{ width: 7, height: 7, borderRadius: 999, marginTop: 5, background: t.type === "nhap_kho" ? TOKENS.info : TOKENS.brand, flexShrink: 0 }} />
                  <div className="min-w-0">
                    <div className="truncate" style={{ color: TOKENS.ink }}>{t.title} · {t.detail}</div>
                    <div className="aa-mono text-[11px]" style={{ color: TOKENS.muted }}>{fmtDate(t.date)}</div>
                  </div>
                </div>
              ))}
              {recentTx.length === 0 && <div className="text-[12px]" style={{ color: TOKENS.muted }}>Chưa có giao dịch kho nào.</div>}
            </div>
            <button onClick={() => onNavigate("warehouse")} className="w-full text-center mt-3 text-[11.5px] py-1.5 rounded-md" style={{ color: TOKENS.brand, background: TOKENS.brandSoft }}>
              Xem báo cáo nhập xuất tồn
            </button>
          </div>

          <div className="rounded-lg p-5" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
            <div className="text-[13px] font-medium mb-4">Phân bổ tài sản</div>
            <div className="space-y-2.5">
              {breakdown.map((b) => (
                <div key={b.label} className="flex items-center gap-3">
                  <div className="w-[128px] text-[11.5px] shrink-0 truncate" style={{ color: TOKENS.muted }}>{b.label}</div>
                  <div className="flex-1 h-2 rounded-full" style={{ background: TOKENS.paper }}>
                    <div className="h-2 rounded-full" style={{ width: `${(b.value / max) * 100}%`, background: STATUS_COLOR[b.label] }} />
                  </div>
                  <div className="w-6 text-right aa-mono text-[12px]">{b.value}</div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>

      {/* Dải thẻ truy cập nhanh */}
      <div className="grid grid-cols-6 gap-3">
        {quickActions.map((qa) => (
          <QuickActionCard key={qa.label} icon={qa.icon} label={qa.label} onClick={() => onNavigate(qa.target)} />
        ))}
      </div>
    </div>
  );
}

/* ============================== ASSET CATALOG ============================== */

function Th({ children, right }) {
  return <th className={`text-[11px] font-medium px-3 py-2 ${right ? "text-right" : "text-left"}`} style={{ color: TOKENS.muted, borderBottom: `1px solid ${TOKENS.border}` }}>{children}</th>;
}
function Td({ children, right, mono, style }) {
  return <td className={`px-3 py-2.5 text-[13px] ${right ? "text-right" : "text-left"} ${mono ? "aa-mono" : ""}`} style={{ borderBottom: `1px solid ${TOKENS.border}`, ...style }}>{children}</td>;
}

function ExportBar({ onExcel, onPdf }) {
  return (
    <div className="flex items-center gap-2">
      <Btn icon={FileDown} onClick={onExcel}>Xuất Excel</Btn>
      <Btn icon={Printer} onClick={onPdf}>Xuất PDF</Btn>
    </div>
  );
}

function AssetCatalog({ assets, projectName, onSelect, selectedAssetId, onAdd, onDelete, isAdmin, customColumns = [], onExportExcel, onExportPdf }) {
  const headers = ["Mã quản lý", "Tên tài sản", "Loại", "Nhóm tài sản", "Công trình", "Nguồn", "Trạng thái", "Nguyên giá", ...customColumns.map((c) => c.label)];
  const buildRows = () => assets.map((a) => [
    a.code, a.name, a.category, a.assetGroup || "", projectName(a.projectId), a.ownership || "Công ty", a.status, a.cost,
    ...customColumns.map((c) => a.customFields?.[c.key] ?? ""),
  ]);

  return (
    <div className="aa-fade">
      <div className="flex items-center justify-between mb-4">
        <h1 className="aa-display text-xl font-semibold">Danh mục tài sản</h1>
        <div className="flex items-center gap-2">
          <ExportBar onExcel={() => onExportExcel("danh-muc-tai-san", headers, buildRows())} onPdf={() => onExportPdf("Danh mục tài sản", headers, buildRows())} />
          <Btn kind="primary" icon={Plus} onClick={onAdd}>Thêm mới</Btn>
        </div>
      </div>
      <div className="rounded-lg overflow-x-auto aa-scroll" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <table className="w-full">
          <thead><tr>
            <Th>Mã quản lý</Th><Th>Tên tài sản</Th><Th>Loại</Th><Th>Nhóm tài sản</Th><Th>Công trình</Th><Th>Nguồn</Th><Th>Trạng thái</Th><Th right>Nguyên giá</Th>
            {customColumns.map((c) => <Th key={c.key}>{c.label}</Th>)}
          </tr></thead>
          <tbody>
            {assets.map((a) => (
              <tr key={a.id} className="aa-row cursor-pointer" onClick={() => onSelect(a.id)}
                style={{ background: selectedAssetId === a.id ? TOKENS.brandSoft : "transparent" }}>
                <Td mono><Tag>{a.code}</Tag></Td>
                <Td>{a.name}</Td>
                <Td>{a.category}</Td>
                <Td>{a.assetGroup || "—"}</Td>
                <Td>{projectName(a.projectId)}</Td>
                <Td>{a.ownership || "Công ty"}</Td>
                <Td><StatusDot status={a.status} /></Td>
                <Td right mono>{fmtVND(a.cost)}</Td>
                {customColumns.map((c) => (
                  <Td key={c.key} mono={c.type !== "text"}>
                    {c.type === "date" ? fmtDate(a.customFields?.[c.key]) : (a.customFields?.[c.key] || "—")}
                  </Td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
        {assets.length === 0 && <EmptyState text="Không tìm thấy tài sản phù hợp" sub="Thử từ khoá khác hoặc bỏ bộ lọc." />}
      </div>
    </div>
  );
}

function AssetDetail({ asset, data, projectName, isAdmin, onDelete, onClose, onAction }) {
  const [tab, setTab] = useState("vongdoi");
  const history = data.transactions.filter((t) => t.assetId === asset.id);
  const repairs = data.repairs.filter((r) => r.assetId === asset.id);
  const mins = data.minutes.filter((m) => m.assetId === asset.id);
  const dep = depreciationOf(asset);

  const tabs = [
    { id: "vongdoi", label: "Vòng đời" }, { id: "thongtin", label: "Thông tin" },
    { id: "phieu", label: "Phiếu" }, { id: "bienban", label: "Biên bản" }, { id: "tep", label: "Tệp" },
  ];

  return (
    <div className="w-[380px] shrink-0 aa-slide overflow-y-auto aa-scroll" style={{ background: TOKENS.surface, borderLeft: `1px solid ${TOKENS.border}` }}>
      <div className="px-5 pt-4 pb-3" style={{ borderBottom: `1px solid ${TOKENS.border}` }}>
        <div className="flex items-start justify-between">
          <div>
            <Tag>{asset.code}</Tag>
            <div className="aa-display font-semibold text-[16px] mt-1.5">{asset.name}</div>
            <div className="mt-1"><StatusDot status={asset.status} /></div>
          </div>
          <button onClick={onClose} className="p-1 rounded hover:bg-black/5"><X size={16} /></button>
        </div>

        <div className="flex gap-1.5 mt-3">
          {tabs.map((t) => (
            <button key={t.id} onClick={() => setTab(t.id)}
              className="px-2.5 py-1 rounded-md text-[12px] font-medium"
              style={{ background: tab === t.id ? TOKENS.brand : TOKENS.paper, color: tab === t.id ? "#fff" : TOKENS.muted }}>
              {t.label}
            </button>
          ))}
        </div>
      </div>

      <div className="px-5 py-4">
        {tab === "vongdoi" && (
          <div>
            <div className="grid grid-cols-3 gap-2 mb-4">
              <div className="text-center rounded-md py-2" style={{ background: TOKENS.paper }}>
                <div className="aa-display font-semibold text-[15px]">{history.filter(h=>h.type==="cap_phat").length}</div>
                <div className="text-[10.5px]" style={{ color: TOKENS.muted }}>lần bàn giao</div>
              </div>
              <div className="text-center rounded-md py-2" style={{ background: TOKENS.paper }}>
                <div className="aa-display font-semibold text-[15px]">{repairs.length}</div>
                <div className="text-[10.5px]" style={{ color: TOKENS.muted }}>lần sửa chữa</div>
              </div>
              <div className="text-center rounded-md py-2" style={{ background: TOKENS.paper }}>
                <div className="aa-display font-semibold text-[15px]">{mins.length}</div>
                <div className="text-[10.5px]" style={{ color: TOKENS.muted }}>biên bản</div>
              </div>
            </div>
            {history.length === 0 ? <EmptyState text="Chưa có giao dịch" /> : (
              <div className="space-y-4">
                {history.map((h) => (
                  <div key={h.id} className="flex gap-3">
                    <div className="flex flex-col items-center pt-0.5">
                      <span style={{ width: 7, height: 7, borderRadius: 999, background: TOKENS.brand }} />
                      <span className="flex-1 w-px my-0.5" style={{ background: TOKENS.border }} />
                    </div>
                    <div className="pb-1">
                      <div className="text-[11px] aa-mono" style={{ color: TOKENS.muted }}>{fmtDate(h.date)}</div>
                      <div className="text-[13px] font-medium">{h.title}</div>
                      <div className="text-[12px]" style={{ color: TOKENS.muted }}>{h.detail}</div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {tab === "thongtin" && (
          <div className="space-y-2.5 text-[13px]">
            {[
              ["Serial", asset.serial], ["Danh mục", asset.category], ["Bộ phận", asset.department],
              ["Người dùng", projectName(asset.projectId)], ["Ngày mua", fmtDate(asset.purchaseDate)],
              ["Nguyên giá", fmtVND(asset.cost)], ["Thời gian SD", `${asset.usefulLifeMonths} tháng`],
              ["Bảo hành đến", asset.warranty ? fmtDate(asset.warrantyEnd) : "Không bảo hành"],
              ["Nhà cung cấp", asset.supplier || "—"],
            ].map(([k, v]) => (
              <div key={k} className="flex justify-between gap-3 py-1" style={{ borderBottom: `1px dashed ${TOKENS.border}` }}>
                <span style={{ color: TOKENS.muted }}>{k}</span><span className="font-medium text-right aa-mono">{v}</span>
              </div>
            ))}
            <div className="mt-3 rounded-md p-3" style={{ background: TOKENS.goldSoft }}>
              <div className="text-[11px] font-medium mb-1" style={{ color: TOKENS.gold }}>Khấu hao hiện tại</div>
              <div className="flex justify-between text-[12px]"><span>Đã khấu hao</span><span className="aa-mono">{fmtVND(dep.depreciated)} ({dep.pct}%)</span></div>
              <div className="flex justify-between text-[12px]"><span>Giá trị còn lại</span><span className="aa-mono">{fmtVND(dep.remaining)}</span></div>
            </div>
          </div>
        )}

        {tab === "phieu" && (
          repairs.length === 0 ? <EmptyState text="Chưa có phiếu sửa chữa" /> : (
            <div className="space-y-2">
              {repairs.map((r) => (
                <div key={r.id} className="rounded-md p-3 text-[12.5px]" style={{ background: TOKENS.paper }}>
                  <div className="flex justify-between"><span className="font-medium">{r.description}</span><span className="aa-mono">{fmtVND(r.cost)}</span></div>
                  <div className="flex justify-between mt-1" style={{ color: TOKENS.muted }}><span>{fmtDate(r.date)}</span><span>{r.status}</span></div>
                </div>
              ))}
            </div>
          )
        )}

        {tab === "bienban" && (
          mins.length === 0 ? <EmptyState text="Chưa có biên bản" /> : (
            <div className="space-y-2">
              {mins.map((m) => (
                <div key={m.id} className="rounded-md p-3 text-[12.5px]" style={{ background: TOKENS.paper }}>
                  <div className="font-medium">{m.type}</div>
                  <div style={{ color: TOKENS.muted }} className="mt-0.5">{m.content}</div>
                  <div className="flex justify-between mt-1" style={{ color: TOKENS.muted }}><span>{fmtDate(m.date)}</span><span>{m.status}</span></div>
                </div>
              ))}
            </div>
          )
        )}

        {tab === "tep" && <EmptyState text="Chưa có tệp đính kèm" sub="Ảnh, hoá đơn hoặc chứng từ liên quan sẽ hiện ở đây." />}
      </div>

      <div className="px-5 py-3.5 grid grid-cols-2 gap-2 sticky bottom-0" style={{ background: TOKENS.surface, borderTop: `1px solid ${TOKENS.border}` }}>
        <Btn kind="primary" onClick={() => onAction("assign")}>Cấp phát</Btn>
        <Btn kind="gold" onClick={() => onAction("minutes")}>Biên bản</Btn>
        <Btn kind="info" icon={ArrowLeftRight} onClick={() => onAction("transfer")}>Chuyển bộ phận</Btn>
        <Btn icon={Pencil} onClick={() => onAction("editAsset")}>Sửa</Btn>
        {isAdmin && <Btn kind="danger" icon={Trash2} onClick={() => onDelete(asset.id)}>Xoá</Btn>}
        <Btn kind="gold" icon={Wrench} onClick={() => onAction("repair")}>Phiếu sửa</Btn>
        {isAdmin ? (
          <Btn kind="danger" icon={Trash2} onClick={() => onAction("liquidate")}>Thanh lý</Btn>
        ) : (
          <div className="flex items-center justify-center gap-1 text-[11.5px] rounded-md" style={{ color: TOKENS.muted }}>
            <ShieldAlert size={12} /> Cần Toàn quyền
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================== BY EMPLOYEE ============================== */

function ByProject({ data, projectName, onSelect }) {
  const grouped = data.projects.map((project) => ({ project, assets: data.assets.filter((a) => a.projectId === project.id) }));
  return (
    <div className="aa-fade">
      <h1 className="aa-display text-xl font-semibold mb-4">Tài sản theo công trình</h1>
      <div className="space-y-4">
        {grouped.map(({ project, assets }) => (
          <div key={project.id} className="rounded-lg overflow-hidden" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
            <div className="flex items-center justify-between px-4 py-3" style={{ borderBottom: `1px solid ${TOKENS.border}` }}>
              <div className="flex items-center gap-2.5">
                <div className="w-8 h-8 rounded-full flex items-center justify-center aa-display font-semibold text-[12px]" style={{ background: TOKENS.brandSoft, color: TOKENS.brand }}>
                  {project.name?.split(" ").slice(-1)[0]?.[0] || "C"}
                </div>
                <div>
                  <div className="text-[13px] font-medium">{project.name}</div>
                  <div className="text-[11.5px]" style={{ color: TOKENS.muted }}>{project.commander} · {project.address}</div>
                </div>
              </div>
              <div className="aa-mono text-[12px]" style={{ color: TOKENS.muted }}>{assets.length} tài sản</div>
            </div>
            {assets.length > 0 && (
              <div className="px-4 py-2 flex flex-wrap gap-2">
                {assets.map((a) => (
                  <button key={a.id} onClick={() => onSelect(a.id)} className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-md text-[12px]" style={{ background: TOKENS.paper }}>
                    <Tag>{a.code}</Tag><span>{a.name}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

/* ============================== DEPRECIATION ============================== */

function Depreciation({ assets, onExportExcel, onExportPdf }) {
  const headers = ["Mã quản lý", "Tên tài sản", "Nguyên giá", "Thời gian SD (tháng)", "Bắt đầu tính", "Đã KH (tháng)", "Giá trị còn lại", "% còn lại"];
  const buildRows = () => assets.map((a) => {
    const d = depreciationOf(a);
    return [a.code, a.name, a.cost, a.usefulLifeMonths, fmtDate(a.purchaseDate), d.elapsed, d.remaining, `${100 - d.pct}%`];
  });
  return (
    <div className="aa-fade">
      <div className="flex items-center justify-between mb-4">
        <h1 className="aa-display text-xl font-semibold">Khấu hao tài sản</h1>
        <ExportBar onExcel={() => onExportExcel("khau-hao-tai-san", headers, buildRows())} onPdf={() => onExportPdf("Khấu hao tài sản", headers, buildRows())} />
      </div>
      <div className="rounded-lg overflow-hidden" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <table className="w-full">
          <thead><tr>
            <Th>Mã quản lý</Th><Th>Tên tài sản</Th><Th right>Nguyên giá</Th><Th>Thời gian SD</Th><Th>Bắt đầu tính</Th>
            <Th right>Đã KH</Th><Th right>Giá trị còn lại</Th><Th right>% còn lại</Th>
          </tr></thead>
          <tbody>
            {assets.map((a) => {
              const d = depreciationOf(a);
              return (
                <tr key={a.id} className="aa-row">
                  <Td mono><Tag>{a.code}</Tag></Td>
                  <Td>{a.name}</Td>
                  <Td right mono>{fmtVND(a.cost)}</Td>
                  <Td>{a.usefulLifeMonths} tháng</Td>
                  <Td>{fmtDate(a.purchaseDate)}</Td>
                  <Td right mono>{d.elapsed} th</Td>
                  <Td right mono style={{ color: TOKENS.gold }}>{fmtVND(d.remaining)}</Td>
                  <Td right mono>{100 - d.pct}%</Td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================== REPAIR ============================== */

function RepairView({ repairs, assetsById, onComplete }) {
  return (
    <div className="aa-fade">
      <h1 className="aa-display text-xl font-semibold mb-4">Sửa chữa — đang xử lý</h1>
      {repairs.length === 0 ? (
        <div className="rounded-lg" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
          <EmptyState text="Không có phiếu sửa chữa nào đang mở" sub="Tạo phiếu sửa chữa từ trang chi tiết tài sản." />
        </div>
      ) : (
        <div className="grid grid-cols-2 gap-3">
          {repairs.map((r) => {
            const a = assetsById[r.assetId];
            return (
              <div key={r.id} className="rounded-lg p-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
                <div className="flex justify-between items-start">
                  <div><Tag>{a?.code}</Tag><div className="text-[13px] font-medium mt-1">{a?.name}</div></div>
                  <span className="text-[11px] px-2 py-0.5 rounded-full" style={{ background: TOKENS.goldSoft, color: TOKENS.gold }}>{r.status}</span>
                </div>
                <div className="text-[12.5px] mt-2" style={{ color: TOKENS.muted }}>{r.description}</div>
                <div className="flex justify-between items-center mt-3">
                  <span className="aa-mono text-[12px]">{fmtDate(r.date)} · {fmtVND(r.cost)}</span>
                  <Btn small kind="primary" onClick={() => onComplete(r.id)}>Hoàn thành</Btn>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function RepairHistory({ repairs, assetsById, onExportExcel, onExportPdf }) {
  const headers = ["Mã quản lý", "Tài sản", "Mô tả", "Ngày gửi", "Chi phí", "Trạng thái"];
  const buildRows = () => repairs.map((r) => {
    const a = assetsById[r.assetId];
    return [a?.code, a?.name, r.description, fmtDate(r.date), r.cost, r.status];
  });
  return (
    <div className="aa-fade">
      <div className="flex items-center justify-between mb-4">
        <h1 className="aa-display text-xl font-semibold">Lịch sử sửa chữa</h1>
        <ExportBar onExcel={() => onExportExcel("lich-su-sua-chua", headers, buildRows())} onPdf={() => onExportPdf("Lịch sử sửa chữa", headers, buildRows())} />
      </div>
      <div className="rounded-lg overflow-hidden" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <table className="w-full">
          <thead><tr><Th>Mã quản lý</Th><Th>Tài sản</Th><Th>Mô tả</Th><Th>Ngày gửi</Th><Th right>Chi phí</Th><Th>Trạng thái</Th></tr></thead>
          <tbody>
            {repairs.map((r) => {
              const a = assetsById[r.assetId];
              return (
                <tr key={r.id} className="aa-row">
                  <Td mono><Tag>{a?.code}</Tag></Td><Td>{a?.name}</Td><Td>{r.description}</Td>
                  <Td>{fmtDate(r.date)}</Td><Td right mono>{fmtVND(r.cost)}</Td>
                  <Td><span style={{ color: r.status === "Hoàn thành" ? TOKENS.brand : TOKENS.gold }}>{r.status}</span></Td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {repairs.length === 0 && <EmptyState text="Chưa có lịch sử sửa chữa" />}
      </div>
    </div>
  );
}

/* ============================== LIQUIDATION ============================== */

function LiquidationView({ liquidations, assetsById, onExportExcel, onExportPdf }) {
  const headers = ["Mã quản lý", "Tài sản", "Ngày thanh lý", "Giá trị thu hồi", "Lý do"];
  const buildRows = () => liquidations.map((l) => {
    const a = assetsById[l.assetId];
    return [a?.code, a?.name, fmtDate(l.date), l.value, l.reason];
  });
  return (
    <div className="aa-fade">
      <div className="flex items-center justify-between mb-4">
        <h1 className="aa-display text-xl font-semibold">Thanh lý tài sản</h1>
        <ExportBar onExcel={() => onExportExcel("thanh-ly-tai-san", headers, buildRows())} onPdf={() => onExportPdf("Thanh lý tài sản", headers, buildRows())} />
      </div>
      <div className="rounded-lg overflow-hidden" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <table className="w-full">
          <thead><tr><Th>Mã quản lý</Th><Th>Tài sản</Th><Th>Ngày thanh lý</Th><Th right>Giá trị thu hồi</Th><Th>Lý do</Th></tr></thead>
          <tbody>
            {liquidations.map((l) => {
              const a = assetsById[l.assetId];
              return (
                <tr key={l.id} className="aa-row">
                  <Td mono><Tag>{a?.code}</Tag></Td><Td>{a?.name}</Td><Td>{fmtDate(l.date)}</Td>
                  <Td right mono style={{ color: TOKENS.gold }}>{fmtVND(l.value)}</Td><Td>{l.reason}</Td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {liquidations.length === 0 && <EmptyState text="Chưa có tài sản thanh lý" />}
      </div>
    </div>
  );
}

/* ============================== MINUTES / TRANSACTIONS / EMPLOYEES / LOG ============================== */

function MinutesView({ minutes, assetsById, onExportExcel, onExportPdf }) {
  const headers = ["Loại biên bản", "Mã quản lý", "Nội dung", "Ngày lập", "Trạng thái"];
  const buildRows = () => minutes.map((m) => {
    const a = assetsById[m.assetId];
    return [m.type, a?.code, m.content, fmtDate(m.date), m.status];
  });
  return (
    <div className="aa-fade">
      <div className="flex items-center justify-between mb-4">
        <h1 className="aa-display text-xl font-semibold">Biên bản</h1>
        <ExportBar onExcel={() => onExportExcel("bien-ban", headers, buildRows())} onPdf={() => onExportPdf("Biên bản", headers, buildRows())} />
      </div>
      <div className="rounded-lg overflow-hidden" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <table className="w-full">
          <thead><tr><Th>Loại biên bản</Th><Th>Tài sản</Th><Th>Nội dung</Th><Th>Ngày lập</Th><Th>Trạng thái</Th></tr></thead>
          <tbody>
            {minutes.map((m) => {
              const a = assetsById[m.assetId];
              return (
                <tr key={m.id} className="aa-row">
                  <Td>{m.type}</Td><Td mono><Tag>{a?.code}</Tag></Td><Td>{m.content}</Td>
                  <Td>{fmtDate(m.date)}</Td><Td>{m.status}</Td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {minutes.length === 0 && <EmptyState text="Chưa có biên bản nào" />}
      </div>
    </div>
  );
}

const TX_LABEL = { mua_sam: "Mua sắm", cap_phat: "Cấp phát", chuyen_bo_phan: "Chuyển bộ phận", sua_chua: "Sửa chữa", thanh_ly: "Thanh lý" };
const TX_COLOR = { mua_sam: TOKENS.info, cap_phat: TOKENS.brand, chuyen_bo_phan: "#6E4E9E", sua_chua: TOKENS.gold, thanh_ly: TOKENS.danger };

function TransactionsView({ transactions, assetsById, onExportExcel, onExportPdf }) {
  const headers = ["Loại", "Mã quản lý", "Chi tiết", "Ngày", "Giá trị"];
  const buildRows = () => transactions.map((t) => {
    const a = assetsById[t.assetId];
    return [TX_LABEL[t.type] || t.title, a?.code, t.detail, fmtDate(t.date), t.amount || 0];
  });
  return (
    <div className="aa-fade">
      <div className="flex items-center justify-between mb-4">
        <h1 className="aa-display text-xl font-semibold">Lịch sử giao dịch</h1>
        <ExportBar onExcel={() => onExportExcel("lich-su-giao-dich", headers, buildRows())} onPdf={() => onExportPdf("Lịch sử giao dịch", headers, buildRows())} />
      </div>
      <div className="rounded-lg overflow-hidden" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <table className="w-full">
          <thead><tr><Th>Loại</Th><Th>Tài sản</Th><Th>Chi tiết</Th><Th>Ngày</Th><Th right>Giá trị</Th></tr></thead>
          <tbody>
            {transactions.map((t) => {
              const a = assetsById[t.assetId];
              return (
                <tr key={t.id} className="aa-row">
                  <Td><span className="text-[12px] px-2 py-0.5 rounded-full" style={{ background: `${TX_COLOR[t.type]}18`, color: TX_COLOR[t.type] }}>{TX_LABEL[t.type] || t.title}</span></Td>
                  <Td mono><Tag>{a?.code}</Tag></Td><Td>{t.detail}</Td><Td>{fmtDate(t.date)}</Td>
                  <Td right mono>{t.amount ? fmtVND(t.amount) : "—"}</Td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {transactions.length === 0 && <EmptyState text="Chưa có giao dịch nào" />}
      </div>
    </div>
  );
}

function ProjectsView({ projects, assets, onAdd, onExportExcel, onExportPdf, isAdmin, onDelete }) {
  const headers = ["Chỉ huy trưởng","Tên công trình","Địa chỉ","Hạng mục thi công","Ngày bắt đầu","Ngày kết thúc","Tài sản"];
  const rows = projects.map(p => [p.commander,p.name,p.address,p.workItem,fmtDate(p.startDate),fmtDate(p.endDate),assets.filter(a=>a.projectId===p.id).length]);
  return <div className="aa-fade"><div className="flex items-center justify-between mb-4"><h1 className="aa-display text-xl font-semibold">Công trình</h1><div className="flex gap-2"><ExportBar onExcel={()=>onExportExcel("cong-trinh",headers,rows)} onPdf={()=>onExportPdf("Công trình",headers,rows)}/><Btn kind="primary" icon={Plus} onClick={onAdd}>Thêm công trình</Btn></div></div><div className="rounded-lg overflow-hidden" style={{background:TOKENS.surface,border:`1px solid ${TOKENS.border}`}}><table className="w-full"><thead><tr>{headers.map(h=><Th key={h}>{h}</Th>)}{isAdmin&&<Th>Thao tác</Th>}</tr></thead><tbody>{projects.map(p=><tr key={p.id} className="aa-row"><Td>{p.commander}</Td><Td>{p.name}</Td><Td>{p.address}</Td><Td>{p.workItem}</Td><Td>{fmtDate(p.startDate)}</Td><Td>{fmtDate(p.endDate)}</Td><Td right mono>{assets.filter(a=>a.projectId===p.id).length}</Td>{isAdmin&&<Td><button className="p-1 rounded hover:bg-black/10" title="Xoá công trình" onClick={()=>onDelete(p.id)}><Trash2 size={14}/></button></Td>}</tr>)}</tbody></table>{projects.length===0&&<EmptyState text="Chưa có công trình"/>}</div></div>;
}

function ActivityLogView({ log }) {
  return (
    <div className="aa-fade">
      <h1 className="aa-display text-xl font-semibold mb-4">Nhật ký thao tác</h1>
      <div className="rounded-lg p-5" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <div className="space-y-3">
          {[...log].reverse().map((l) => (
            <div key={l.id} className="flex gap-3 text-[13px]">
              <span className="aa-mono shrink-0" style={{ color: TOKENS.muted, width: 130 }}>{fmtDate(l.date)} {new Date(l.date).toLocaleTimeString("vi-VN", { hour: "2-digit", minute: "2-digit" })}</span>
              <span className="font-medium shrink-0">{l.user}</span>
              <span style={{ color: TOKENS.muted }}>{l.action}</span>
            </div>
          ))}
        </div>
        {log.length === 0 && <EmptyState text="Chưa có hoạt động nào" />}
      </div>
    </div>
  );
}

function WarehouseHub({ warehouse, assets, projects, settings, onAdd, onExportExcel, onExportPdf, onImport, isAdmin }) {
  const [tab, setTab] = useState("in");
  const [filter, setFilter] = useState({ category:"", group:"", ownership:"", projectId:"", locationName:"", asOfDate:nowIso().slice(0,10) });
  const [query, setQuery] = useState("");
  const projectName = id => projects.find(p => p.id === id)?.name || "";
  const getLocation = w => w.locationName || (w.projectId ? projectName(w.projectId) : "Kho trung tâm");
  const categories=[...new Set(assets.map(a=>a.category).filter(Boolean))];
  const groupsList=[...new Set(assets.map(a=>a.assetGroup).filter(Boolean))];
  const ownerships=[...new Set(assets.map(a=>a.ownership).filter(Boolean))];
  const locations=[...new Set(warehouse.map(getLocation).filter(Boolean))];
  const filteredTx=warehouse.filter(w=>{const a=assets.find(x=>x.id===w.assetId);if(!a)return false;if(filter.category&&(w.category||a.category)!==filter.category)return false;if(filter.group&&(w.assetGroup||a.assetGroup)!==filter.group)return false;if(filter.ownership&&(w.ownership||a.ownership)!==filter.ownership)return false;if(filter.projectId&&(w.projectId||a.projectId)!==filter.projectId)return false;if(filter.locationName&&getLocation(w)!==filter.locationName)return false;if(query&&!`${w.voucherNo||""} ${a.name} ${a.code}`.toLowerCase().includes(query.toLowerCase()))return false;return true;});
  const headers=["Số phiếu","Ngày tháng","Tên tài sản","Mã hàng","Kho/Công trình","Loại tài sản","Nhóm tài sản","Nguồn gốc","Người giao/nhận","Số lượng","ĐVT","Đơn giá","Thành tiền","Ghi chú"];
  const reportTx=warehouse.filter(w=>String(w.date||"")<=String(filter.asOfDate));
  const balances={};
  reportTx.forEach(w=>{const a=assets.find(x=>x.id===w.assetId);if(!a)return;const loc=getLocation(w), key=`${a.id}¦${loc}`;if(!balances[key])balances[key]={assetId:a.id,code:a.code,name:a.name,location:loc,category:w.category||a.category||"Khác",group:w.assetGroup||a.assetGroup||"Thiết bị chính",ownership:w.ownership||a.ownership||"Công ty",unit:w.unit||"Cái",inQty:0,outQty:0,inValue:0,outValue:0};const q=Number(w.quantity||0),v=Number(w.total??q*Number(w.unitCost||0));if(w.type==="nhap"){balances[key].inQty+=q;balances[key].inValue+=v}else{balances[key].outQty+=q;balances[key].outValue+=v;}});
  const detailRows=Object.values(balances).filter(r=>r.inQty-r.outQty>0).filter(r=>!filter.category||r.category===filter.category).filter(r=>!filter.group||r.group===filter.group).filter(r=>!filter.ownership||r.ownership===filter.ownership).filter(r=>!filter.locationName||r.location===filter.locationName).filter(r=>!filter.projectId||(projects.find(p=>p.name===r.location)?.id===filter.projectId)).sort((a,b)=>`${a.location}${a.group}${a.category}${a.code}`.localeCompare(`${b.location}${b.group}${b.category}${b.code}`));
  const reportHeaders=["Kho/Công trình","Loại tài sản","Nhóm tài sản","Nguồn gốc","Mã hàng","Tên tài sản","ĐVT","Nhập lũy kế","Xuất lũy kế","Tồn đến ngày","Giá trị tồn"];
  const reportRows=detailRows.map(r=>[r.location,r.category,r.group,r.ownership,r.code,r.name,r.unit,r.inQty,r.outQty,r.inQty-r.outQty,r.inValue-r.outValue]);
  const summary={qty:detailRows.reduce((s,r)=>s+r.inQty-r.outQty,0),value:detailRows.reduce((s,r)=>s+r.inValue-r.outValue,0)};
  const tabBtn=(id,label,icon)=><button onClick={()=>setTab(id)} className="flex items-center gap-2 px-4 py-2.5 text-[13px] font-medium" style={{color:tab===id?TOKENS.brand:TOKENS.muted,borderBottom:`2px solid ${tab===id?TOKENS.brand:"transparent"}`}}>{icon}{label}</button>;
  const downloadTemplate=()=>downloadExcelTemplate("Mau_Import_Phieu_Nhap_Xuat_Kho",["Loại phiếu","Số phiếu","Ngày tháng","Tên tài sản","Mã hàng","Kho/Công trình","Loại địa điểm","Loại tài sản","Nhóm tài sản","Nguồn gốc","Người giao/nhận","Số lượng","Đơn vị tính","Đơn giá","Thành tiền","Ghi chú"],[["Nhập kho","PN-20260822-001","2026-08-22","Máy xúc 01","MX-01",projects[0]?.name||"Kho trung tâm",projects[0]?"Công trình":"Kho","Máy xúc","Thiết bị chính","Thuê","Nguyễn Văn A",1,"Cái",500000000,500000000,"Nhập đầu kỳ"]]);
  return <div className="aa-fade"><div className="flex items-start justify-between mb-4 gap-4"><div><h1 className="aa-display text-xl font-semibold">Kho</h1><div className="text-[12px] mt-1" style={{color:TOKENS.muted}}>Quản lý phiếu nhập, phiếu xuất và tồn thiết bị theo từng kho/công trình</div></div><div className="flex gap-2 flex-wrap justify-end"><Btn icon={Download} onClick={downloadTemplate}>Tải mẫu Excel</Btn>{isAdmin&&<label className="inline-flex items-center gap-1.5 rounded-md font-medium px-3 py-1.5 text-[13px] cursor-pointer" style={{background:TOKENS.info,color:"white"}}><UploadCloud size={14}/>Import Excel<input type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)onImport(f);e.target.value=""}}/></label>}</div></div><div className="rounded-lg overflow-hidden" style={{background:TOKENS.surface,border:`1px solid ${TOKENS.border}`}}><div className="flex border-b" style={{borderColor:TOKENS.border}}>{tabBtn("in","Phiếu nhập kho",<DownloadCloud size={15}/>)}{tabBtn("out","Phiếu xuất kho",<UploadCloud size={15}/>)}{tabBtn("report","Báo cáo nhập xuất tồn",<ClipboardList size={15}/> )}</div><div className="p-4">{(tab==="in"||tab==="out")&&<><div className="flex items-center justify-between gap-3 mb-3"><Btn kind="primary" icon={Plus} onClick={()=>onAdd(tab==="in"?"nhap":"xuat")}>{tab==="in"?"Lập phiếu nhập":"Lập phiếu xuất"}</Btn><input className={inputCls} style={{...inputStyle,maxWidth:290}} placeholder="Tìm số phiếu / tên tài sản / mã" value={query} onChange={e=>setQuery(e.target.value)}/></div><WarehouseFilter filter={filter} setFilter={setFilter} categories={categories} groups={groupsList} ownerships={ownerships} projects={projects} locations={locations}/><div className="rounded-lg overflow-auto" style={{border:`1px solid ${TOKENS.border}`}}><table className="w-full min-w-[1500px]"><thead><tr>{headers.map(h=><Th key={h}>{h}</Th>)}</tr></thead><tbody>{filteredTx.filter(w=>w.type===(tab==="in"?"nhap":"xuat")).map(w=>{const a=assets.find(x=>x.id===w.assetId)||{};return <tr key={w.id} className="aa-row"><Td><Tag>{w.voucherNo||"—"}</Tag></Td><Td mono>{fmtDate(w.date)}</Td><Td>{w.itemName||a.name}</Td><Td><Tag>{w.itemCode||a.code}</Tag></Td><Td>{getLocation(w)}</Td><Td>{w.category||a.category}</Td><Td>{w.assetGroup||a.assetGroup}</Td><Td>{w.ownership||a.ownership}</Td><Td>{w.receiver||"—"}</Td><Td right mono>{w.quantity}</Td><Td>{w.unit||"Cái"}</Td><Td right mono>{fmtVND(w.unitCost)}</Td><Td right mono>{fmtVND(w.total??Number(w.quantity||0)*Number(w.unitCost||0))}</Td><Td>{w.note||"—"}</Td></tr>})}</tbody></table>{!filteredTx.filter(w=>w.type===(tab==="in"?"nhap":"xuat")).length&&<EmptyState text={tab==="in"?"Chưa có phiếu nhập kho":"Chưa có phiếu xuất kho"}/>}</div></>}{tab==="report"&&<><div className="rounded-lg p-4 mb-4" style={{background:TOKENS.paper,border:`1px solid ${TOKENS.border}`}}><div className="flex flex-wrap gap-2 items-end"><WarehouseFilter filter={filter} setFilter={setFilter} categories={categories} groups={groupsList} ownerships={ownerships} projects={projects} locations={locations}/><label className="text-[11px]" style={{color:TOKENS.muted}}>Tồn đến ngày<input type="date" className={inputCls} style={{...inputStyle,width:155}} value={filter.asOfDate} onChange={e=>setFilter({...filter,asOfDate:e.target.value})}/></label></div><div className="flex flex-wrap gap-2 mt-3 items-center"><span className="text-[12px] px-2 py-1 rounded" style={{background:TOKENS.brandSoft,color:TOKENS.brand}}>Tổng thiết bị tồn: <b>{summary.qty}</b></span><span className="text-[12px] px-2 py-1 rounded" style={{background:TOKENS.goldSoft,color:TOKENS.gold}}>Giá trị tồn: <b>{fmtVND(summary.value)}</b></span><div className="ml-auto"><ExportBar onExcel={()=>onExportExcel("bao-cao-ton-thiet-bi-theo-kho-cong-trinh",reportHeaders,reportRows)} onPdf={()=>onExportPdf(`Báo cáo tồn thiết bị đến ${fmtDate(filter.asOfDate)}`,reportHeaders,reportRows)}/></div></div></div><div className="mb-3 text-[13px] font-medium">Chi tiết thiết bị đang tồn tại thời điểm {fmtDate(filter.asOfDate)}</div><div className="rounded-lg overflow-auto" style={{border:`1px solid ${TOKENS.border}`}}><table className="w-full min-w-[1450px]"><thead><tr>{reportHeaders.map(h=><Th key={h} right={h.includes("lũy kế")||h.includes("Tồn")||h.includes("Giá trị")}>{h}</Th>)}</tr></thead><tbody>{reportRows.map((r,i)=><tr key={i} className="aa-row">{r.map((c,j)=><Td key={j} right={j>=7} mono={j>=7}>{j===10?fmtVND(c):c}</Td>)}</tr>)}</tbody></table>{!reportRows.length&&<EmptyState text="Không có thiết bị đang tồn theo điều kiện lọc"/>}</div></>}</div></div></div>;
}
function WarehouseFilter({filter,setFilter,categories,groups,ownerships,projects,locations}){const set=k=>e=>setFilter({...filter,[k]:e.target.value});return <div className="flex flex-wrap gap-2 mb-3"><select className={inputCls} style={{...inputStyle,width:175}} value={filter.category} onChange={set("category")}><option value="">Tất cả loại tài sản</option>{categories.map(x=><option key={x}>{x}</option>)}</select><select className={inputCls} style={{...inputStyle,width:175}} value={filter.group} onChange={set("group")}><option value="">Tất cả nhóm</option>{groups.map(x=><option key={x}>{x}</option>)}</select><select className={inputCls} style={{...inputStyle,width:155}} value={filter.ownership} onChange={set("ownership")}><option value="">Tất cả nguồn gốc</option>{ownerships.map(x=><option key={x}>{x}</option>)}</select><select className={inputCls} style={{...inputStyle,width:210}} value={filter.projectId} onChange={set("projectId")}><option value="">Tất cả công trình</option>{projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select><select className={inputCls} style={{...inputStyle,width:220}} value={filter.locationName} onChange={set("locationName")}><option value="">Tất cả kho / công trình</option>{locations.map(x=><option key={x}>{x}</option>)}</select></div>;
}
function CostHistoryView({ costHistory, assetsById, onAdd, onExportExcel, onExportPdf }) { const headers=["Tài sản","Loại chi phí","Ngày","Số tiền","Nội dung","Nhà cung cấp"]; const rows=costHistory.map(c=>[assetsById[c.assetId]?.code,c.type,fmtDate(c.date),c.amount,c.description,c.vendor]); return <div className="aa-fade"><div className="flex items-center justify-between mb-4"><h1 className="aa-display text-xl font-semibold">Chi phí thiết bị</h1><div className="flex gap-2"><ExportBar onExcel={()=>onExportExcel("lich-su-chi-phi-thiet-bi",headers,rows)} onPdf={()=>onExportPdf("Lịch sử chi phí thiết bị",headers,rows)}/><Btn kind="primary" icon={Plus} onClick={onAdd}>Thêm chi phí</Btn></div></div><div className="rounded-lg overflow-hidden" style={{background:TOKENS.surface,border:`1px solid ${TOKENS.border}`}}><table className="w-full"><thead><tr>{headers.map(h=><Th key={h}>{h}</Th>)}</tr></thead><tbody>{costHistory.map(c=><tr key={c.id} className="aa-row"><Td mono><Tag>{assetsById[c.assetId]?.code}</Tag></Td><Td>{c.type}</Td><Td>{fmtDate(c.date)}</Td><Td right mono>{fmtVND(c.amount)}</Td><Td>{c.description}</Td><Td>{c.vendor}</Td></tr>)}</tbody></table>{!costHistory.length&&<EmptyState text="Chưa có chi phí"/>}</div></div>; }
function WarehouseTxModal({ assets, projects, onClose, onSubmit, title, fixedType }) {
  const first=assets[0]; const [f,setF]=useState({assetId:first?.id||"",type:fixedType||"nhap",voucherNo:"",quantity:1,date:nowIso().slice(0,10),unit:"Cái",unitCost:first?.cost||0,receiver:"",locationType:first?.projectId?"project":"warehouse",warehouseName:"Kho trung tâm",category:first?.category||"",assetGroup:first?.assetGroup||"",ownership:first?.ownership||"Công ty",projectId:first?.projectId||"",note:""});
  const asset=assets.find(a=>a.id===f.assetId)||{}; const set=k=>e=>setF({...f,[k]:e.target.value}); const pick=e=>{const a=assets.find(x=>x.id===e.target.value)||{};setF({...f,assetId:e.target.value,category:a.category||"",assetGroup:a.assetGroup||"",ownership:a.ownership||"Công ty",projectId:a.projectId||"",unitCost:a.cost||0,locationType:a.projectId?"project":"warehouse"})}; const total=(Number(f.quantity)||0)*(Number(f.unitCost)||0);
  return <Modal title={title||"Lập phiếu kho"} onClose={onClose} wide><div className="grid grid-cols-2 gap-x-4"><Field label="Số phiếu"><input className={inputCls} style={inputStyle} value={f.voucherNo} onChange={set("voucherNo")} placeholder={fixedType==="xuat"?"PX-20260822-001":"PN-20260822-001"}/></Field><Field label="Ngày tháng"><input type="date" className={inputCls} style={inputStyle} value={f.date} onChange={set("date")}/></Field><Field label="Tên tài sản"><select className={inputCls} style={inputStyle} value={f.assetId} onChange={pick}>{assets.map(a=><option key={a.id} value={a.id}>{a.name}</option>)}</select></Field><Field label="Mã hàng tự động"><input readOnly className={inputCls} style={{...inputStyle,background:TOKENS.paper}} value={asset.code||""}/></Field><Field label="Kho / Công trình"><select className={inputCls} style={inputStyle} value={f.locationType} onChange={e=>setF({...f,locationType:e.target.value,projectId:e.target.value==="project"?f.projectId:""})}><option value="project">Công trình</option><option value="warehouse">Kho</option></select></Field><Field label={f.locationType==="project"?"Tên công trình":"Tên kho"}>{f.locationType==="project"?<select className={inputCls} style={inputStyle} value={f.projectId} onChange={set("projectId")}><option value="">Chọn công trình</option>{projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select>:<input className={inputCls} style={inputStyle} value={f.warehouseName} onChange={set("warehouseName")} placeholder="VD: Kho vật tư Xà Lá"/>}</Field><Field label="Người giao/nhận"><input className={inputCls} style={inputStyle} value={f.receiver} onChange={set("receiver")} placeholder="Người giao hoặc người nhận"/></Field><Field label="Loại tài sản"><input readOnly className={inputCls} style={{...inputStyle,background:TOKENS.paper}} value={asset.category||""}/></Field><Field label="Nhóm tài sản"><input readOnly className={inputCls} style={{...inputStyle,background:TOKENS.paper}} value={asset.assetGroup||""}/></Field><Field label="Nguồn gốc"><input readOnly className={inputCls} style={{...inputStyle,background:TOKENS.paper}} value={asset.ownership||""}/></Field><Field label="Số lượng"><input type="number" min="0.01" step="0.01" className={inputCls} style={inputStyle} value={f.quantity} onChange={set("quantity")}/></Field><Field label="Đơn vị tính"><input className={inputCls} style={inputStyle} value={f.unit} onChange={set("unit")}/></Field><Field label="Đơn giá"><input type="number" min="0" className={inputCls} style={inputStyle} value={f.unitCost} onChange={set("unitCost")}/></Field><Field label="Thành tiền"><input readOnly className={inputCls} style={{...inputStyle,background:TOKENS.paper,fontWeight:600}} value={fmtVND(total)}/></Field></div><Field label="Ghi chú"><textarea className={inputCls} style={inputStyle} value={f.note} onChange={set("note")} placeholder="Số chứng từ, diễn giải..."/></Field><div className="flex justify-end gap-2 mt-3"><Btn onClick={onClose}>Huỷ</Btn><Btn kind="primary" onClick={()=>onSubmit(f)}>Lưu phiếu</Btn></div></Modal>;
}
function CostHistoryModal({ assets, onClose, onSubmit }) { const [f,setF]=useState({assetId:assets[0]?.id||"",type:"Cấp dầu",date:nowIso().slice(0,10),amount:0,description:"",vendor:""}); const set=k=>e=>setF({...f,[k]:e.target.value}); return <Modal title="Thêm chi phí thiết bị" onClose={onClose}><Field label="Tài sản"><select className={inputCls} style={inputStyle} value={f.assetId} onChange={set("assetId")}>{assets.map(a=><option key={a.id} value={a.id}>{a.code} — {a.name}</option>)}</select></Field><Field label="Loại chi phí"><select className={inputCls} style={inputStyle} value={f.type} onChange={set("type")}><option>Cấp dầu</option><option>Thay cáp</option><option>Sửa chữa</option><option>Chi phí khác</option></select></Field><Field label="Ngày"><input type="date" className={inputCls} style={inputStyle} value={f.date} onChange={set("date")}/></Field><Field label="Số tiền"><input type="number" className={inputCls} style={inputStyle} value={f.amount} onChange={set("amount")}/></Field><Field label="Nội dung"><textarea className={inputCls} style={inputStyle} value={f.description} onChange={set("description")}/></Field><Field label="Nhà cung cấp"><input className={inputCls} style={inputStyle} value={f.vendor} onChange={set("vendor")}/></Field><div className="flex justify-end gap-2"><Btn onClick={onClose}>Huỷ</Btn><Btn kind="primary" onClick={()=>onSubmit(f)}>Lưu chi phí</Btn></div></Modal>; }

/* ============================== SETTINGS ============================== */

function EditableTagList({ items, onAdd, onRemove, usage, placeholder }) {
  const [val, setVal] = useState("");
  const submit = () => { if (val.trim()) { onAdd(val.trim()); setVal(""); } };
  return (
    <div>
      <div className="flex flex-wrap gap-2 mb-3">
        {items.map((it) => {
          const n = usage ? usage(it) : 0;
          return (
            <span key={it} className="inline-flex items-center gap-1.5 pl-2.5 pr-1.5 py-1 rounded-md text-[12.5px]" style={{ background: TOKENS.paper, border: `1px solid ${TOKENS.border}` }}>
              {it}
              {usage && <span className="aa-mono text-[10.5px]" style={{ color: TOKENS.muted }}>({n})</span>}
              <button
                title={n > 0 ? `Đang dùng bởi ${n} tài sản` : "Xoá"}
                onClick={() => { if (n > 0 && !window.confirm(`"${it}" đang được ${n} tài sản sử dụng. Vẫn xoá khỏi danh sách lựa chọn?`)) return; onRemove(it); }}
                className="p-0.5 rounded hover:bg-black/10"
              ><X size={12} /></button>
            </span>
          );
        })}
        {items.length === 0 && <span className="text-[12px]" style={{ color: TOKENS.muted }}>Chưa có mục nào</span>}
      </div>
      <div className="flex gap-2">
        <input className={inputCls} style={inputStyle} value={val} onChange={(e) => setVal(e.target.value)}
          placeholder={placeholder} onKeyDown={(e) => e.key === "Enter" && submit()} />
        <Btn kind="primary" icon={Plus} onClick={submit}>Thêm</Btn>
      </div>
    </div>
  );
}

function SettingsView({
  settings, onSetCompanyName, onAddCategory, onRemoveCategory, onAddDepartment, onRemoveDepartment,
  onAddColumn, onRemoveColumn, categoryUsage, deptUsage, users, currentUser, onSetUserRole, onSendPasswordReset,
  backups, onCreateBackup, onRestoreBackup, onDeleteBackup, onDownloadBackupFile, onRestoreFromFile, onImportExcel,
}) {
  const [name, setName] = useState(settings.companyName);
  const [colLabel, setColLabel] = useState("");
  const [colType, setColType] = useState("text");

  return (
    <div className="aa-fade max-w-3xl">
      <h1 className="aa-display text-xl font-semibold mb-4">Cài đặt</h1>

      <div className="rounded-lg p-5 mb-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <div className="text-[13px] font-medium mb-3">Tên công ty</div>
        <div className="flex gap-2 max-w-md">
          <input className={inputCls} style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} />
          <Btn kind="primary" icon={Check} disabled={!name.trim() || name === settings.companyName} onClick={() => onSetCompanyName(name.trim())}>Lưu</Btn>
        </div>
        <div className="text-[11.5px] mt-2" style={{ color: TOKENS.muted }}>Tên này hiển thị ở thanh điều hướng bên trái và trên các bản xuất PDF.</div>
      </div>

      <div className="rounded-lg p-5 mb-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <div className="text-[13px] font-medium mb-1">Nhập liệu hàng loạt từ Excel</div>
        <div className="text-[11.5px] mb-3" style={{ color: TOKENS.muted }}>Dùng để đổ lô dữ liệu nhanh vào danh mục tài sản hoặc công trình. Dòng đầu tiên phải là tiêu đề cột.</div>
        <div className="flex gap-2">
          <label className="inline-flex"><input type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0]; if(f) onImportExcel(f,"assets"); e.target.value="";}}/><span className="inline-flex items-center gap-1.5 rounded-md font-medium px-3 py-1.5 text-[13px] cursor-pointer" style={{background:TOKENS.brand,color:"#fff"}}><UploadCloud size={14}/>Import tài sản</span></label>
          <label className="inline-flex"><input type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0]; if(f) onImportExcel(f,"projects"); e.target.value="";}}/><span className="inline-flex items-center gap-1.5 rounded-md font-medium px-3 py-1.5 text-[13px] cursor-pointer" style={{background:TOKENS.info,color:"#fff"}}><UploadCloud size={14}/>Import công trình</span></label>
        </div>
      </div>

      <div className="rounded-lg p-5 mb-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <div className="text-[13px] font-medium mb-1">Danh mục tài sản</div>
        <div className="text-[11.5px] mb-3" style={{ color: TOKENS.muted }}>Các loại tài sản có thể chọn khi thêm mới (Laptop, Màn hình…).</div>
        <EditableTagList items={settings.categories} onAdd={onAddCategory} onRemove={onRemoveCategory} usage={categoryUsage} placeholder="VD: Máy chiếu" />
      </div>

      <div className="rounded-lg p-5 mb-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <div className="text-[13px] font-medium mb-1">Bộ phận / đơn vị sử dụng</div>
        <div className="text-[11.5px] mb-3" style={{ color: TOKENS.muted }}>Danh sách bộ phận dùng khi cấp phát hoặc chuyển tài sản.</div>
        <EditableTagList items={settings.departments} onAdd={onAddDepartment} onRemove={onRemoveDepartment} usage={deptUsage} placeholder="VD: Marketing" />
      </div>

      <div className="rounded-lg p-5" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <div className="text-[13px] font-medium mb-1">Cột tuỳ chỉnh cho tài sản</div>
        <div className="text-[11.5px] mb-3" style={{ color: TOKENS.muted }}>Thêm cột riêng của công ty bạn — sẽ hiện trong Danh mục tài sản và khi thêm/sửa tài sản.</div>

        <div className="space-y-2 mb-3">
          {settings.customColumns.map((c) => (
            <div key={c.key} className="flex items-center justify-between px-3 py-2 rounded-md" style={{ background: TOKENS.paper }}>
              <div className="flex items-center gap-2 text-[13px]">
                <span className="font-medium">{c.label}</span>
                <span className="text-[11px] px-1.5 py-0.5 rounded" style={{ background: TOKENS.brandSoft, color: TOKENS.brand }}>{FIELD_TYPES.find((t) => t.id === c.type)?.label}</span>
              </div>
              <button onClick={() => onRemoveColumn(c.key)} className="p-1 rounded hover:bg-black/10"><Trash2 size={13} /></button>
            </div>
          ))}
          {settings.customColumns.length === 0 && <div className="text-[12px]" style={{ color: TOKENS.muted }}>Chưa có cột tuỳ chỉnh nào</div>}
        </div>

        <div className="flex gap-2">
          <input className={inputCls} style={inputStyle} value={colLabel} onChange={(e) => setColLabel(e.target.value)} placeholder="Tên cột, VD: Vị trí lắp đặt" />
          <select className={inputCls} style={{ ...inputStyle, width: 120 }} value={colType} onChange={(e) => setColType(e.target.value)}>
            {FIELD_TYPES.map((t) => <option key={t.id} value={t.id}>{t.label}</option>)}
          </select>
          <Btn kind="primary" icon={Plus} onClick={() => { if (colLabel.trim()) { onAddColumn(colLabel.trim(), colType); setColLabel(""); } }}>Thêm cột</Btn>
        </div>
      </div>

      <UserManager users={users} currentUser={currentUser} onSetUserRole={onSetUserRole} onSendPasswordReset={onSendPasswordReset} />

      <BackupSettings
        backups={backups} onCreateBackup={onCreateBackup} onRestoreBackup={onRestoreBackup}
        onDeleteBackup={onDeleteBackup} onDownloadBackupFile={onDownloadBackupFile} onRestoreFromFile={onRestoreFromFile}
      />
    </div>
  );
}

function BackupSettings({ backups, onCreateBackup, onRestoreBackup, onDeleteBackup, onDownloadBackupFile, onRestoreFromFile }) {
  const [label, setLabel] = useState("");
  const [busy, setBusy] = useState(false);
  const fileInputRef = useRef(null);

  const handleCreate = async () => {
    setBusy(true);
    await onCreateBackup(label.trim() || undefined);
    setLabel("");
    setBusy(false);
  };

  const handleFilePick = (e) => {
    const file = e.target.files?.[0];
    if (file) onRestoreFromFile(file);
    e.target.value = "";
  };

  return (
    <div className="rounded-lg p-5 mt-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
      <div className="flex items-center gap-2 mb-1">
        <HardDrive size={15} style={{ color: TOKENS.brand }} />
        <div className="text-[13px] font-medium">Sao lưu &amp; khôi phục máy chủ</div>
      </div>
      <div className="text-[11.5px] mb-3" style={{ color: TOKENS.muted }}>
        Toàn bộ sổ tài sản (tài sản, nhân sự, giao dịch, cài đặt, tài khoản) được lưu chung trên máy chủ ngay khi bạn thao tác.
        Mục này tạo thêm các <b>bản chụp theo mốc thời gian</b> để có thể khôi phục nếu lỡ xoá/sửa nhầm dữ liệu.
        Hệ thống tự động sao lưu 1 lần/ngày khi có quản trị viên mở ứng dụng, giữ tối đa 15 bản gần nhất.
      </div>

      <div className="flex flex-wrap items-center gap-2 mb-4 pb-4" style={{ borderBottom: `1px dashed ${TOKENS.border}` }}>
        <input className={inputCls} style={{ ...inputStyle, width: 220 }} placeholder="Nhãn bản sao lưu (tuỳ chọn)" value={label} onChange={(e) => setLabel(e.target.value)} />
        <Btn kind="primary" icon={Save} disabled={busy} onClick={handleCreate}>{busy ? "Đang sao lưu…" : "Sao lưu ngay"}</Btn>
        <span className="w-px h-5" style={{ background: TOKENS.border }} />
        <Btn icon={Download} onClick={onDownloadBackupFile}>Tải file .json về máy</Btn>
        <Btn icon={Upload} onClick={() => fileInputRef.current?.click()}>Khôi phục từ file</Btn>
        <input ref={fileInputRef} type="file" accept="application/json" className="hidden" onChange={handleFilePick} />
      </div>

      <div className="space-y-2">
        {backups.length === 0 && <div className="text-[12px]" style={{ color: TOKENS.muted }}>Chưa có bản sao lưu nào trên máy chủ.</div>}
        {backups.map((b) => (
          <div key={b.key} className="flex items-center justify-between px-3 py-2 rounded-md" style={{ background: TOKENS.paper }}>
            <div className="flex items-center gap-2.5">
              <div className="w-7 h-7 rounded-full flex items-center justify-center" style={{ background: TOKENS.brandSoft }}>
                <HardDrive size={13} style={{ color: TOKENS.brand }} />
              </div>
              <div>
                <div className="text-[13px] font-medium">{b.label}</div>
                <div className="text-[11px] aa-mono" style={{ color: TOKENS.muted }}>{fmtDate(b.date)} {new Date(b.date).toLocaleTimeString("vi-VN", { hour: "2-digit", minute: "2-digit" })} · {b.by}</div>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <Btn small icon={RotateCcw} onClick={() => onRestoreBackup(b.key)}>Khôi phục</Btn>
              <button onClick={() => onDeleteBackup(b.key)} className="p-1 rounded hover:bg-black/10" title="Xoá bản sao lưu"><Trash2 size={13} /></button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function UserManager({ users, currentUser, onSetUserRole, onSendPasswordReset }) {
  const adminCount = users.filter((u) => u.role === "admin").length;

  return (
    <div className="rounded-lg p-5 mt-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
      <div className="text-[13px] font-medium mb-1">Tài khoản đăng nhập</div>
      <div className="text-[11.5px] mb-3" style={{ color: TOKENS.muted }}>
        <b>Toàn quyền</b> dùng được mọi chức năng, kể cả Cài đặt. <b>Nhập liệu &amp; xuất báo cáo</b> chỉ nhập/cập nhật dữ liệu nghiệp vụ và xuất Excel/PDF; không được xoá tài sản, công trình, danh mục hay thanh lý.
        Người mới tự tạo tài khoản ở màn hình đăng nhập (tab "Tạo tài khoản") — bạn nâng quyền cho họ ở đây. Muốn khoá hẳn một tài khoản, xoá ở Supabase Dashboard → Authentication → Users.
      </div>

      <div className="space-y-2">
        {users.map((u) => {
          const isLastAdmin = u.role === "admin" && adminCount === 1;
          return (
            <div key={u.id} className="flex items-center justify-between px-3 py-2 rounded-md" style={{ background: TOKENS.paper }}>
              <div className="flex items-center gap-2.5 min-w-0">
                <div className="w-7 h-7 rounded-full flex items-center justify-center aa-display font-semibold text-[11px] shrink-0"
                  style={{ background: u.role === "admin" ? TOKENS.goldSoft : TOKENS.infoSoft, color: u.role === "admin" ? TOKENS.gold : TOKENS.info }}>
                  {u.name?.[0]?.toUpperCase() || "?"}
                </div>
                <div className="min-w-0">
                  <div className="text-[13px] font-medium truncate">{u.name} {u.id === currentUser.id && <span className="text-[11px]" style={{ color: TOKENS.muted }}>(bạn)</span>}</div>
                  <div className="text-[11.5px] aa-mono truncate" style={{ color: TOKENS.muted }}>{u.email}</div>
                </div>
              </div>
              <div className="flex items-center gap-2 shrink-0">
                <select
                  value={u.role}
                  disabled={isLastAdmin}
                  title={isLastAdmin ? "Cần giữ lại ít nhất 1 tài khoản Toàn quyền" : "Đổi quyền"}
                  onChange={(e) => onSetUserRole(u.id, e.target.value)}
                  className="text-[11.5px] rounded-md px-2 py-1 disabled:opacity-50"
                  style={{ border: `1px solid ${TOKENS.border}`, background: u.role === "admin" ? TOKENS.goldSoft : TOKENS.infoSoft, color: u.role === "admin" ? TOKENS.gold : TOKENS.info }}
                >
                  <option value="user">Nhập liệu &amp; xuất báo cáo</option>
                  <option value="admin">Toàn quyền</option>
                </select>
                <button onClick={() => onSendPasswordReset(u.email)} className="p-1 rounded hover:bg-black/10" title="Gửi email đặt lại mật khẩu"><Lock size={13} /></button>
              </div>
            </div>
          );
        })}
        {users.length === 0 && <div className="text-[12px]" style={{ color: TOKENS.muted }}>Chưa có tài khoản nào ngoài bạn.</div>}
      </div>
    </div>
  );
}

/* ============================== FORMS / MODALS ============================== */

function AssetFormModal({ title, initial, categories, departments, projects = [], customColumns = [], onClose, onSubmit }) {
  const [f, setF] = useState(() => initial ? {
    code: initial.code, name: initial.name, category: initial.category, cost: initial.cost, assetGroup: initial.assetGroup || initial.category, ownership: initial.ownership || "Công ty", quantity: initial.quantity || 1, projectId: initial.projectId || "",
    purchaseDate: initial.purchaseDate, usefulLifeMonths: initial.usefulLifeMonths, department: initial.department,
    serial: initial.serial, supplier: initial.supplier, warranty: initial.warranty, warrantyEnd: initial.warrantyEnd, note: initial.note,
    customFields: { ...(initial.customFields || {}) },
  } : {
    code: "", name: "", category: categories[0] || "", cost: "", purchaseDate: nowIso().slice(0, 10),
    usefulLifeMonths: 36, department: departments[0] || "", projectId: "", assetGroup: categories[0] || "Thiết bị chính", ownership: "Công ty", quantity: 1, serial: "", supplier: "", warranty: true, warrantyEnd: "", note: "",
    customFields: {},
  });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const setCustom = (key) => (e) => setF({ ...f, customFields: { ...f.customFields, [key]: e.target.value } });
  const valid = f.code.trim() && f.name.trim() && Number(f.cost) > 0;

  return (
    <Modal title={title} onClose={onClose} wide>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Mã quản lý"><input className={inputCls} style={inputStyle} value={f.code} onChange={set("code")} placeholder="VD: TS-LT-009" /></Field>
        <Field label="Tên tài sản"><input className={inputCls} style={inputStyle} value={f.name} onChange={set("name")} placeholder="VD: Laptop Dell" /></Field>
        <Field label="Danh mục">
          <select className={inputCls} style={inputStyle} value={f.category} onChange={set("category")}>
            {categories.map((c) => <option key={c}>{c}</option>)}
          </select>
        </Field>
        <Field label="Nhóm tài sản"><input className={inputCls} style={inputStyle} value={f.assetGroup} onChange={set("assetGroup")} placeholder="VD: Thiết bị chính" /></Field>
        <Field label="Nguồn gốc"><select className={inputCls} style={inputStyle} value={f.ownership} onChange={set("ownership")}><option>Công ty</option><option>Thuê</option></select></Field>
        <Field label="Công trình"><select className={inputCls} style={inputStyle} value={f.projectId} onChange={set("projectId")}><option value="">Kho / chưa giao</option>{projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></Field>
        <Field label="Số lượng"><input type="number" min="1" className={inputCls} style={inputStyle} value={f.quantity} onChange={set("quantity")} /></Field>
        <Field label="Bộ phận">
          <select className={inputCls} style={inputStyle} value={f.department} onChange={set("department")}>
            {departments.map((d) => <option key={d}>{d}</option>)}
          </select>
        </Field>
        <Field label="Nguyên giá (đ)"><input type="number" className={inputCls} style={inputStyle} value={f.cost} onChange={set("cost")} placeholder="10000000" /></Field>
        <Field label="Thời gian sử dụng (tháng)"><input type="number" className={inputCls} style={inputStyle} value={f.usefulLifeMonths} onChange={set("usefulLifeMonths")} /></Field>
        <Field label="Ngày mua"><input type="date" className={inputCls} style={inputStyle} value={f.purchaseDate} onChange={set("purchaseDate")} /></Field>
        <Field label="Serial"><input className={inputCls} style={inputStyle} value={f.serial} onChange={set("serial")} /></Field>
        <Field label="Nhà cung cấp"><input className={inputCls} style={inputStyle} value={f.supplier} onChange={set("supplier")} /></Field>
        <Field label="Bảo hành đến"><input type="date" className={inputCls} style={inputStyle} value={f.warrantyEnd} onChange={set("warrantyEnd")} /></Field>
        {customColumns.map((c) => (
          <Field key={c.key} label={c.label}>
            <input type={c.type === "number" ? "number" : c.type === "date" ? "date" : "text"} className={inputCls} style={inputStyle}
              value={f.customFields[c.key] || ""} onChange={setCustom(c.key)} />
          </Field>
        ))}
      </div>
      <Field label="Ghi chú"><textarea rows={2} className={inputCls} style={inputStyle} value={f.note} onChange={set("note")} /></Field>
      <div className="flex justify-end gap-2 mt-4">
        <Btn onClick={onClose}>Huỷ</Btn>
        <Btn kind="primary" disabled={!valid} onClick={() => onSubmit(f)}>Lưu tài sản</Btn>
      </div>
    </Modal>
  );
}

function AssignModal({ asset, projects, departments, onClose, onSubmit }) { const [projectId,setProjectId]=useState(asset.projectId||projects[0]?.id||""); const [dept,setDept]=useState(asset.department||""); return <Modal title={`Giao tài sản — ${asset.code}`} onClose={onClose}><Field label="Công trình"><select className={inputCls} style={inputStyle} value={projectId} onChange={e=>setProjectId(e.target.value)}>{projects.map(p=><option key={p.id} value={p.id}>{p.name} — {p.commander}</option>)}</select></Field><Field label="Bộ phận sử dụng"><select className={inputCls} style={inputStyle} value={dept} onChange={e=>setDept(e.target.value)}>{departments.map(d=><option key={d}>{d}</option>)}</select></Field><div className="flex justify-end gap-2 mt-4"><Btn onClick={onClose}>Huỷ</Btn><Btn kind="primary" disabled={!projectId} onClick={()=>onSubmit(projectId,dept)}>Xác nhận giao</Btn></div></Modal>; }

function TransferModal({ asset, departments, onClose, onSubmit }) {
  const [dept, setDept] = useState(asset.department);
  return (
    <Modal title={`Chuyển bộ phận — ${asset.code}`} onClose={onClose}>
      <Field label="Bộ phận mới">
        <select className={inputCls} style={inputStyle} value={dept} onChange={(e) => setDept(e.target.value)}>
          {departments.map((d) => <option key={d}>{d}</option>)}
        </select>
      </Field>
      <div className="flex justify-end gap-2 mt-4">
        <Btn onClick={onClose}>Huỷ</Btn>
        <Btn kind="info" onClick={() => onSubmit(dept)}>Xác nhận chuyển</Btn>
      </div>
    </Modal>
  );
}

function RepairModal({ asset, onClose, onSubmit }) {
  const [desc, setDesc] = useState("");
  const [cost, setCost] = useState("");
  return (
    <Modal title={`Phiếu sửa chữa — ${asset.code}`} onClose={onClose}>
      <Field label="Mô tả sự cố"><textarea rows={3} className={inputCls} style={inputStyle} value={desc} onChange={(e) => setDesc(e.target.value)} placeholder="VD: Lỗi bàn phím, không lên nguồn…" /></Field>
      <Field label="Chi phí dự kiến (đ)"><input type="number" className={inputCls} style={inputStyle} value={cost} onChange={(e) => setCost(e.target.value)} /></Field>
      <div className="flex justify-end gap-2 mt-4">
        <Btn onClick={onClose}>Huỷ</Btn>
        <Btn kind="gold" disabled={!desc.trim()} onClick={() => onSubmit(desc, cost)}>Tạo phiếu sửa</Btn>
      </div>
    </Modal>
  );
}

function LiquidateModal({ asset, onClose, onSubmit }) {
  const [value, setValue] = useState("");
  const [reason, setReason] = useState("");
  return (
    <Modal title={`Thanh lý — ${asset.code}`} onClose={onClose}>
      <Field label="Giá trị thu hồi (đ)"><input type="number" className={inputCls} style={inputStyle} value={value} onChange={(e) => setValue(e.target.value)} /></Field>
      <Field label="Lý do thanh lý"><textarea rows={2} className={inputCls} style={inputStyle} value={reason} onChange={(e) => setReason(e.target.value)} /></Field>
      <div className="flex justify-end gap-2 mt-4">
        <Btn onClick={onClose}>Huỷ</Btn>
        <Btn kind="danger" disabled={!reason.trim()} onClick={() => onSubmit(value, reason)}>Xác nhận thanh lý</Btn>
      </div>
    </Modal>
  );
}

const MINUTES_TYPES = ["Biên bản bàn giao", "Biên bản thu hồi", "Biên bản kiểm kê", "Biên bản thanh lý"];

function MinutesModal({ asset, onClose, onSubmit }) {
  const [type, setType] = useState(MINUTES_TYPES[0]);
  const [content, setContent] = useState("");
  return (
    <Modal title={`Lập biên bản — ${asset.code}`} onClose={onClose}>
      <Field label="Loại biên bản">
        <select className={inputCls} style={inputStyle} value={type} onChange={(e) => setType(e.target.value)}>
          {MINUTES_TYPES.map((t) => <option key={t}>{t}</option>)}
        </select>
      </Field>
      <Field label="Nội dung"><textarea rows={3} className={inputCls} style={inputStyle} value={content} onChange={(e) => setContent(e.target.value)} placeholder="Nội dung biên bản…" /></Field>
      <div className="flex justify-end gap-2 mt-4">
        <Btn onClick={onClose}>Huỷ</Btn>
        <Btn kind="gold" disabled={!content.trim()} onClick={() => onSubmit(type, content)}>Lập biên bản</Btn>
      </div>
    </Modal>
  );
}

function ProjectFormModal({ onClose, onSubmit }) { const [f,setF]=useState({commander:"",name:"",address:"",workItem:"",startDate:"",endDate:""}); const set=k=>e=>setF({...f,[k]:e.target.value}); return <Modal title="Thêm công trình" onClose={onClose} wide><div className="grid grid-cols-2 gap-x-4"><Field label="Chỉ huy trưởng"><input className={inputCls} style={inputStyle} value={f.commander} onChange={set("commander")}/></Field><Field label="Tên công trình"><input className={inputCls} style={inputStyle} value={f.name} onChange={set("name")}/></Field><Field label="Địa chỉ"><input className={inputCls} style={inputStyle} value={f.address} onChange={set("address")}/></Field><Field label="Hạng mục thi công"><input className={inputCls} style={inputStyle} value={f.workItem} onChange={set("workItem")}/></Field><Field label="Ngày bắt đầu"><input type="date" className={inputCls} style={inputStyle} value={f.startDate} onChange={set("startDate")}/></Field><Field label="Ngày kết thúc"><input type="date" className={inputCls} style={inputStyle} value={f.endDate} onChange={set("endDate")}/></Field></div><div className="flex justify-end gap-2 mt-4"><Btn onClick={onClose}>Huỷ</Btn><Btn kind="primary" disabled={!f.name.trim()} onClick={()=>onSubmit(f)}>Thêm công trình</Btn></div></Modal>; }

function ChangePasswordModal({ onClose, onSubmit }) {
  const [oldPw, setOldPw] = useState("");
  const [newPw, setNewPw] = useState("");
  const [confirmPw, setConfirmPw] = useState("");
  const mismatch = newPw && confirmPw && newPw !== confirmPw;
  const valid = oldPw && newPw.length >= 4 && newPw === confirmPw;
  return (
    <Modal title="Đổi mật khẩu" onClose={onClose}>
      <Field label="Mật khẩu hiện tại"><input type="password" className={inputCls} style={inputStyle} value={oldPw} onChange={(e) => setOldPw(e.target.value)} /></Field>
      <Field label="Mật khẩu mới (tối thiểu 4 ký tự)"><input type="password" className={inputCls} style={inputStyle} value={newPw} onChange={(e) => setNewPw(e.target.value)} /></Field>
      <Field label="Xác nhận mật khẩu mới"><input type="password" className={inputCls} style={inputStyle} value={confirmPw} onChange={(e) => setConfirmPw(e.target.value)} /></Field>
      {mismatch && <div className="text-[12px] mb-2" style={{ color: TOKENS.danger }}>Mật khẩu xác nhận không khớp</div>}
      <div className="flex justify-end gap-2 mt-4">
        <Btn onClick={onClose}>Huỷ</Btn>
        <Btn kind="primary" disabled={!valid} onClick={() => onSubmit(oldPw, newPw)}>Đổi mật khẩu</Btn>
      </div>
    </Modal>
  );
}
