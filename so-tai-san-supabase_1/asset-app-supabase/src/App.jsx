import { Component, useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";
import {
  LayoutDashboard, Boxes, Users2, Wrench, History, FileText, Archive, Warehouse, Fuel, Cable, CircleDollarSign, UploadCloud, DownloadCloud,
  ClipboardList, UserCog, ScrollText, Search, RefreshCw, X, Plus,
  ChevronRight, Building2, ArrowLeftRight, ShieldCheck, Trash2, Pencil,
  ChevronDown, Package, Stamp, AlertCircle, Loader2, ClipboardCheck,
  Settings as SettingsIcon, FileDown, Printer, Check, PencilLine,
  Lock, LogOut, User, ShieldAlert, Eye, EyeOff, Download, Upload, RotateCcw, HardDrive, Save,
  PackagePlus, PackageMinus, FileSpreadsheet, ArrowUp, ArrowDown, SlidersHorizontal,
} from "lucide-react";

const CORE_VERSION = "v17.0.0-verified";

/* ============================== DESIGN TOKENS ==============================
Color:
  paper       #F5F6F2  page background, unbleached ledger paper
  surface     #FFFFFF  cards / table
  ink         #1B2320  primary text, near-black warm green-charcoal
  muted       #707A73  secondary text
  border      #E2E5DE  hairlines
  brand       #22594C  deep ledger green (primary actions, active nav)
  brandSoft   #E6EEEA  brand tint
  gold        #A9832E  brass accent — money / depreciation / value
  goldSoft    #F5EEDC  gold tint
  danger      #A23F3F  faults / liquidation
  dangerSoft  #F5E6E5
  info        #34607F  transfers / info
  infoSoft    #E7EEF3
Type:
  Display: "Space Grotesk"  (nav headers, page titles, stamp numerals)
  Body:    "IBM Plex Sans"  (everything else)
  Mono:    "IBM Plex Mono"  (asset codes, money, dates — ledger data)
Signature: sidebar rendered as a bound ledger spine (notch ticks + stamped
  total), asset codes shown as mono "inventory tag" chips, status shown as
  a wax-seal-style dot rather than a pill.
============================================================================ */

const TOKENS = {
  paper: "#F7F8FA", surface: "#FFFFFF", ink: "#17181A", muted: "#6B7280",
  border: "#E5E7EB", brand: "#D71920", brandSoft: "#FFF0F1",
  gold: "#D71920", goldSoft: "#FFF0F1", danger: "#B42318", dangerSoft: "#FEF3F2",
  info: "#475467", infoSoft: "#F2F4F7",
};

const FONT_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700&family=Inter:wght@400;500;600;700;800&display=swap');
.aa-root{font-family:'Be Vietnam Pro','Inter',sans-serif;color:${TOKENS.ink};background:${TOKENS.paper};letter-spacing:-.01em;}
.aa-display{font-family:'Inter','Be Vietnam Pro',sans-serif;letter-spacing:-.025em;}
.aa-mono{font-family:'Inter',monospace;}
.aa-scroll::-webkit-scrollbar{width:8px;height:8px;}
.aa-scroll::-webkit-scrollbar-track{background:transparent;}
.aa-scroll::-webkit-scrollbar-thumb{background:#D1D5DB;border-radius:999px;}
.aa-scroll::-webkit-scrollbar-thumb:hover{background:#9CA3AF;}
.aa-row{transition:background .16s ease, box-shadow .16s ease;}
.aa-row:hover{background:${TOKENS.brandSoft}66;}
.aa-fade{animation:aaFade .18s ease-out;}
@keyframes aaFade{from{opacity:0;transform:translateY(3px)}to{opacity:1;transform:translateY(0)}}
.aa-slide{animation:aaSlide .2s ease-out;}
@keyframes aaSlide{from{transform:translateX(16px);opacity:0}to{transform:translateX(0);opacity:1}}
input,select,textarea{font-family:inherit;transition:border-color .15s ease,box-shadow .15s ease,background .15s ease;}
input:focus,select:focus,textarea:focus{outline:none!important;border-color:${TOKENS.brand}!important;box-shadow:0 0 0 3px ${TOKENS.brandSoft};}
button{font-family:inherit;}
button:focus-visible{outline:2px solid ${TOKENS.brand};outline-offset:2px;}
.aa-card{box-shadow:0 1px 2px rgba(16,24,40,.04),0 8px 24px rgba(16,24,40,.035);}
.aa-sidebar-item{transition:background .15s ease,color .15s ease,transform .15s ease;}
.aa-sidebar-item:hover{background:#F8FAFC!important;color:${TOKENS.ink}!important;}
.aa-sidebar-item.aa-active:hover{background:${TOKENS.brand}!important;color:#fff!important;}
`;


/* ============================== CONSTANTS ============================== */

const STATUS = {
  ASSIGNED: "Đã cấp phát",
  SHARED: "Dùng chung",
  UNUSED: "Chưa sử dụng",
  REPAIR: "Đang sửa chữa/bảo dưỡng",
  BROKEN: "Hỏng/mất",
  LIQUIDATED: "Đã thanh lý",
  TRANSFERRED_OUT: "Đã điều chuyển",
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
  admin: { label: "Toàn quyền", color: "#D71920" },
  user: { label: "Nhập liệu & xuất báo cáo", color: "#34607F" },
};

const WAREHOUSE_OPERATIONS = {
  nhap: [
    { id: "mua_moi", label: "Mua mới bên ngoài" },
    { id: "luan_chuyen_den", label: "Luân chuyển từ kho/công trình khác" },
    { id: "thu_hoi_sua_chua", label: "Thu hồi sau sửa chữa" },
    { id: "thu_hoi_cong_trinh", label: "Thu hồi từ công trình" },
    { id: "nhap_khac", label: "Nhập khác" },
  ],
  xuat: [
    { id: "su_dung_cong_trinh", label: "Xuất dùng tại công trình" },
    { id: "luan_chuyen_di", label: "Luân chuyển sang kho/công trình khác" },
    { id: "sua_chua", label: "Xuất đi sửa chữa" },
    { id: "xuat_khac", label: "Xuất khác" },
  ],
};
const OPERATION_LABELS = Object.fromEntries(Object.values(WAREHOUSE_OPERATIONS).flat().map(x => [x.id, x.label]));

/* ---------- export helpers (Excel + in PDF) ---------- */

function exportExcel(filename, headers, rows, title = "BÁO CÁO") {
  const generatedAt = `Ngày xuất: ${new Date().toLocaleString("vi-VN")}`;
  const aoa = [[title], [generatedAt], [], headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  if (headers.length > 1) ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: headers.length - 1 } }];
  ws["!cols"] = headers.map((h, i) => ({ wch: Math.min(40, Math.max(12, String(h).length + 3, ...rows.map((r) => String(r[i] ?? "").length + 2))) }));
  ws["!rows"] = [{ hpt: 28 }, { hpt: 19 }, { hpt: 8 }, { hpt: 26 }];
  ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 3, c: 0 }, e: { r: Math.max(3, rows.length + 3), c: Math.max(0, headers.length - 1) } }) };
  ws["!freeze"] = { xSplit: 0, ySplit: 4, topLeftCell: "A5", activePane: "bottomLeft", state: "frozen" };
  const border = { top: { style: "thin", color: { rgb: "B42318" } }, bottom: { style: "thin", color: { rgb: "D0D5DD" } }, left: { style: "thin", color: { rgb: "D0D5DD" } }, right: { style: "thin", color: { rgb: "D0D5DD" } } };
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
  for (let R = range.s.r; R <= range.e.r; ++R) for (let C = range.s.c; C <= range.e.c; ++C) {
    const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })]; if (!cell) continue;
    if (R === 0) cell.s = { font: { bold: true, sz: 16, color: { rgb: "B42318" } }, alignment: { horizontal: "center", vertical: "center" } };
    else if (R === 1) cell.s = { font: { italic: true, sz: 10, color: { rgb: "667085" } }, alignment: { horizontal: "center" } };
    else if (R === 3) cell.s = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "D71920" } }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border };
    else if (R >= 4) cell.s = { border, alignment: { vertical: "top", wrapText: true }, fill: R % 2 ? { fgColor: { rgb: "FFF7F7" } } : undefined };
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Báo cáo");
  XLSX.writeFile(wb, `${filename}.xlsx`);
}


async function exportStyledExcel(filename, title, headers, rows, companyName = "MYHL - Quản lý tài sản") {
  // Xuất Excel dùng thư viện XLSX đã có sẵn trong project.
  // Trình thiết kế báo cáo vẫn quyết định chính xác cột và thứ tự trước khi gọi hàm này.
  exportExcel(filename, headers, rows, title || companyName || "BÁO CÁO");
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
    { id: "assetCategories", label: "Danh mục / loại tài sản", icon: Boxes },
    { id: "departments", label: "Bộ phận / đơn vị", icon: Users2 },
    { id: "suppliers", label: "Danh mục nhà cung cấp", icon: Users2 },
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
const normalizeText = (v) => String(v ?? "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d");

// Chuẩn hóa dữ liệu cũ trước khi render. Một số bản dữ liệu cũ có thể lưu
// supplier/category/location dưới dạng object thay vì chuỗi, làm React lỗi
// "Objects are not valid as a React child" khi mở module Kho.
const safeText = (v, fallback = "") => {
  if (v == null) return fallback;
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return String(v);
  if (v instanceof Date) return isNaN(v) ? fallback : v.toISOString().slice(0, 10);
  if (typeof v === "object") {
    const preferred = v.name ?? v.label ?? v.code ?? v.title ?? v.value;
    if (preferred != null && preferred !== v) return safeText(preferred, fallback);
    try { return JSON.stringify(v); } catch { return fallback; }
  }
  try { return String(v); } catch { return fallback; }
};

const safeNumber = (v, fallback = 0) => {
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
};
const parseOperationType = (raw, type) => {
  const q = normalizeText(raw);
  const ops = WAREHOUSE_OPERATIONS[type] || [];
  const exact = ops.find(o => normalizeText(o.id) === q || normalizeText(o.label) === q);
  if (exact) return exact.id;
  const found = ops.find(o => normalizeText(o.label).includes(q) || q.includes(normalizeText(o.label)));
  return found?.id || (type === "nhap" ? "mua_moi" : "su_dung_cong_trinh");
};
const parseDateValue = (v) => {
  if (v == null || v === "") return nowIso().slice(0,10);
  if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0,10);
  if (typeof v === "number" || /^\d+(\.\d+)?$/.test(String(v).trim())) {
    const n = Number(v);
    if (n > 20000 && n < 80000) { const d = XLSX.SSF.parse_date_code(n); if (d) return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`; }
  }
  const x=String(v).trim();
  let m=x.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if(m) return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
  m=x.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if(m) return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  const d=new Date(x); return isNaN(d)?nowIso().slice(0,10):d.toISOString().slice(0,10);
};
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
    companyName: "MYHL - Quản lý tài sản",
    departments: [...DEPARTMENTS],
    categories: [...CATEGORIES],
    suppliers: [
      { id: "sup-001", code: "NCC-001", name: "Nhà cung cấp chưa ghi", taxCode: "", phone: "", address: "", contact: "", note: "" },
    ],
    customColumns: [], // { key, label, type: 'text'|'number'|'date' }
  };

  // Tài khoản đăng nhập giờ do Supabase Auth quản lý — không còn lưu trong app_data.
  const warehouse = assets.map((a) => ({ id: uid("wh"), voucherNo: `PN-${String(a.purchaseDate || nowIso().slice(0,10)).replaceAll("-", "")}-OPEN`, assetId: a.id, type: "nhap", quantity: Number(a.quantity || 1), date: a.purchaseDate, unitCost: a.cost, note: "Tồn đầu kỳ", projectId: a.projectId || null, locationType: a.projectId ? "project" : "warehouse", locationName: a.projectId ? projects.find(p => p.id === a.projectId)?.name || "Kho trung tâm" : "Kho trung tâm", warehouseName: a.projectId ? "" : "Kho trung tâm", itemName: a.name, itemCode: a.code, category: a.category, assetGroup: a.assetGroup, ownership: a.ownership, supplier: a.supplier || "" }));
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
    quantity: Number(a.quantity || 1), unit: a.unit || "Cái"
  }));
  return {
    ...d,
    assets: migratedAssets,
    projects: legacyProjects,
    warehouse: (Array.isArray(d.warehouse) ? d.warehouse : migratedAssets.map(a => ({ id: uid("wh"), assetId: a.id, type: "nhap", quantity: Number(a.quantity || 1), date: a.purchaseDate || nowIso().slice(0,10), unitCost: Number(a.cost || 0), note: "Tồn đầu kỳ" }))).map((w, i) => { const a = migratedAssets.find(x => x.id === w.assetId); const pid = w.projectId || a?.projectId || null; const loc = w.locationName || (pid ? legacyProjects.find(p => p.id === pid)?.name : "Kho trung tâm") || "Kho trung tâm"; return { ...w, voucherNo: w.voucherNo || `PN-${String(w.date || nowIso().slice(0,10)).replaceAll("-", "")}-OPEN${String(i+1).padStart(3,"0")}`, projectId: pid, locationType: w.locationType || (pid ? "project" : "warehouse"), locationName: loc, warehouseName: w.warehouseName || (pid ? "" : loc), itemName: w.itemName || a?.name || "", itemCode: w.itemCode || a?.code || "", category: w.category || a?.category || "Khác", assetGroup: w.assetGroup || a?.assetGroup || "Thiết bị chính", ownership: w.ownership || a?.ownership || "Công ty", operationType: w.operationType || (w.type === "nhap" ? "mua_moi" : "su_dung_cong_trinh"), operationLabel: w.operationLabel || OPERATION_LABELS[w.operationType] || (w.type === "nhap" ? "Mua mới bên ngoài" : "Xuất dùng tại công trình"), counterpartyLocation: w.counterpartyLocation || "", repairVendor: w.repairVendor || "", transferId: w.transferId || "" }; }),
    costHistory: Array.isArray(d.costHistory) ? d.costHistory : [],
    employees: undefined,
    settings: {
      ...(d.settings || {}),
      companyName: d.settings?.companyName || "MYHL - Quản lý tài sản",
      departments: [...new Set([...(Array.isArray(d.settings?.departments) && d.settings.departments.length ? d.settings.departments : DEPARTMENTS), ...migratedAssets.map(a => a.department).filter(Boolean)])],
      categories: [...new Set([...(Array.isArray(d.settings?.categories) && d.settings.categories.filter(Boolean).length ? d.settings.categories.filter(Boolean) : CATEGORIES), ...migratedAssets.map(a => a.category).filter(Boolean)])],
      suppliers: Array.isArray(d.settings?.suppliers) && d.settings.suppliers.length ? d.settings.suppliers : [...new Set([...(d.assets || []).map(a => a.supplier), ...(d.warehouse || []).map(w => w.supplier)].filter(Boolean))].map((name, i) => ({ id: `sup_legacy_${i}`, code: `NCC-${String(i+1).padStart(3,"0")}`, name, taxCode: "", phone: "", address: "", contact: "", note: "Tự tạo từ dữ liệu cũ" })),
      customColumns: Array.isArray(d.settings?.customColumns) ? d.settings.customColumns : [],
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
    <div className="aa-card rounded-2xl p-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
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
        className={`aa-fade rounded-lg w-full ${wide ? "max-w-6xl" : "max-w-md"} max-h-[85vh] overflow-y-auto aa-scroll`}
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
          <div className="w-12 h-12 rounded-2xl flex items-center justify-center mb-3" style={{ background: TOKENS.brand, boxShadow: `0 8px 18px ${TOKENS.brand}30` }}>
            <ShieldCheck size={22} color="#fff" />
          </div>
          <div className="aa-display font-semibold text-[18px]" style={{ color: TOKENS.ink }}>MYHL – QUẢN LÝ TÀI SẢN</div>
          <div className="text-[12px]" style={{ color: TOKENS.muted }}>Quản lý tài sản nội bộ</div>
        </div>

        <form onSubmit={submit} className="aa-card rounded-2xl p-6" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
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
  const [reportDesigner, setReportDesigner] = useState(null);
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

  useEffect(() => { if(data?.settings?.companyName) document.title=data.settings.companyName; }, [data?.settings?.companyName]);

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
      assetGroup: form.assetGroup || form.category, ownership: form.ownership || "Công ty", quantity: Number(form.quantity) || 1, unit: form.unit || "Cái",
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

  const deleteAssets = (ids) => {
    if (!requireAdmin()) return;
    const unique = [...new Set(ids || [])].filter(Boolean);
    if (!unique.length) return;
    if (!window.confirm(`Xoá hàng loạt ${unique.length} tài sản và toàn bộ dữ liệu liên quan?`)) return;
    const idSet = new Set(unique);
    setData({
      ...data,
      assets: data.assets.filter(a => !idSet.has(a.id)),
      transactions: data.transactions.filter(t => !idSet.has(t.assetId)),
      repairs: data.repairs.filter(r => !idSet.has(r.assetId)),
      liquidations: data.liquidations.filter(r => !idSet.has(r.assetId)),
      minutes: data.minutes.filter(m => !idSet.has(m.assetId)),
      warehouse: (data.warehouse || []).filter(w => !idSet.has(w.assetId)),
      costHistory: data.costHistory.filter(c => !idSet.has(c.assetId)),
      activityLog: logAction(data.activityLog, `Xoá hàng loạt ${unique.length} tài sản`),
    });
    setSelectedAssetId(null); notify(`Đã xoá ${unique.length} tài sản`);
  };

  const addSupplier = (form) => {
    if (!requireAdmin()) return;
    const name = String(form.name || "").trim(); if (!name) return notify("Vui lòng nhập tên nhà cung cấp");
    const supplier = { id: uid("sup"), code: String(form.code || `NCC-${Date.now().toString().slice(-6)}`).trim(), name, taxCode: String(form.taxCode||"").trim(), phone: String(form.phone||"").trim(), address: String(form.address||"").trim(), contact: String(form.contact||"").trim(), note: String(form.note||"").trim() };
    setData({ ...data, settings: { ...data.settings, suppliers: [supplier, ...(data.settings.suppliers || [])] }, activityLog: logAction(data.activityLog, `Thêm nhà cung cấp ${supplier.code} — ${supplier.name}`) }); notify("Đã thêm nhà cung cấp");
  };
  const deleteSuppliers = (ids) => {
    if (!requireAdmin()) return; const unique=[...new Set(ids||[])]; if(!unique.length)return;
    if (!window.confirm(`Xoá ${unique.length} nhà cung cấp đã chọn? Dữ liệu phiếu cũ vẫn giữ tên nhà cung cấp.`)) return;
    const setIds=new Set(unique); setData({ ...data, settings:{...data.settings, suppliers:(data.settings.suppliers||[]).filter(x=>!setIds.has(x.id))}, activityLog:logAction(data.activityLog,`Xoá hàng loạt ${unique.length} nhà cung cấp`) }); notify(`Đã xoá ${unique.length} nhà cung cấp`);
  };

  const editSupplier = (id, form) => {
    if (!requireAdmin()) return;
    const old = (data.settings.suppliers || []).find(x => x.id === id);
    if (!old) return;
    const name = String(form.name || "").trim();
    if (!name) return notify("Tên nhà cung cấp không được để trống");
    const next = { ...old, ...form, name, code: String(form.code || old.code || "").trim(), taxCode: String(form.taxCode || "").trim(), phone: String(form.phone || "").trim(), address: String(form.address || "").trim(), contact: String(form.contact || "").trim(), note: String(form.note || "").trim() };
    setData({ ...data, settings: { ...data.settings, suppliers: (data.settings.suppliers || []).map(x => x.id === id ? next : x) }, activityLog: logAction(data.activityLog, `Sửa nhà cung cấp ${next.code} — ${next.name}`) });
    notify("Đã cập nhật nhà cung cấp");
  };

  const renameCategory = (oldName, newName) => {
    if (!requireAdmin()) return;
    const next = String(newName || "").trim();
    if (!next || next === oldName) return;
    if (data.settings.categories.includes(next) && next !== oldName) return notify("Danh mục mới đã tồn tại");
    setData({
      ...data,
      assets: data.assets.map(a => a.category === oldName ? { ...a, category: next } : a),
      warehouse: (data.warehouse || []).map(w => w.category === oldName ? { ...w, category: next } : w),
      settings: { ...data.settings, categories: data.settings.categories.map(x => x === oldName ? next : x) },
      activityLog: logAction(data.activityLog, `Đổi danh mục tài sản "${oldName}" thành "${next}"`)
    });
    notify("Đã sửa danh mục tài sản");
  };

  const renameDepartment = (oldName, newName) => {
    if (!requireAdmin()) return;
    const next = String(newName || "").trim();
    if (!next || next === oldName) return;
    if (data.settings.departments.includes(next) && next !== oldName) return notify("Bộ phận mới đã tồn tại");
    setData({
      ...data,
      assets: data.assets.map(a => a.department === oldName ? { ...a, department: next } : a),
      settings: { ...data.settings, departments: data.settings.departments.map(x => x === oldName ? next : x) },
      activityLog: logAction(data.activityLog, `Đổi bộ phận "${oldName}" thành "${next}"`)
    });
    notify("Đã sửa bộ phận");
  };

  const addWarehouseTx = (form) => {
    const items = Array.isArray(form.items) ? form.items.filter(x=>x.assetId && Number(x.quantity)>0) : [form];
    if (!items.length) { notify("Phiếu chưa có tài sản"); return false; }
    const currentRows = data.warehouse || [];
    const locationName = form.locationType === "project" ? (data.projects.find(p=>p.id===form.projectId)?.name || "") : (form.warehouseName || "Kho trung tâm");
    if (!locationName) { notify("Vui lòng chọn công trình hoặc kho"); return false; }
    const operationType = form.operationType || (form.type === "nhap" ? "mua_moi" : "su_dung_cong_trinh");
    const operationLabel = OPERATION_LABELS[operationType] || operationType;
    if ((operationType === "nhap_khac" || operationType === "xuat_khac") && !String(form.description||"").trim()) { notify("Nhập/Xuất khác bắt buộc nhập diễn giải lý do"); return false; }
    const counterpartyLocation = form.counterpartyLocationType === "project"
      ? (data.projects.find(p=>p.id===form.counterpartyProjectId)?.name || "")
      : String(form.counterpartyWarehouseName || "").trim();
    if ((operationType === "luan_chuyen_di" || operationType === "luan_chuyen_den" || operationType === "thu_hoi_cong_trinh") && !counterpartyLocation) { notify("Vui lòng chọn kho/công trình nguồn hoặc đích"); return false; }
    if (operationType === "luan_chuyen_di" && counterpartyLocation === locationName) { notify("Kho/công trình đi và đến không được trùng nhau"); return false; }
    if (operationType === "sua_chua" && !String(form.repairVendor||"").trim()) { notify("Vui lòng nhập đơn vị sửa chữa"); return false; }
    if (operationType === "mua_moi" && !String(form.supplier||"").trim()) { notify("Phiếu mua mới cần chọn/nhập nhà cung cấp"); return false; }

    const date=parseDateValue(form.date);
    const prefixMap={mua_moi:"PN",luan_chuyen_den:"LC",thu_hoi_sua_chua:"PN-SC",thu_hoi_cong_trinh:"PN-TH",nhap_khac:"PN-K",su_dung_cong_trinh:"PX",luan_chuyen_di:"LC",sua_chua:"PX-SC",xuat_khac:"PX-K"};
    const prefix=prefixMap[operationType] || (form.type === "nhap" ? "PN" : "PX"), dateKey=date.replaceAll("-","");
    const seq=currentRows.filter(w=>String(w.voucherNo||"").startsWith(`${prefix}-${dateKey}-`)).length+1;
    const voucherNo=form.voucherNo?.trim() || `${prefix}-${dateKey}-${String(seq).padStart(3,"0")}`;
    const rows=[], pairedRows=[], transferId=operationType === "luan_chuyen_di" ? uid("tr") : "";

    for (const item of items) {
      const asset=assetsById[item.assetId]; if(!asset) continue;
      const qty=Number(item.quantity)||0, enteredUnitCost=Number(item.unitCost)||0;
      const fifoLotsFor=(assetId,loc,untilDate)=>{
        const source=[...currentRows,...rows].filter(w=>w&&w.assetId===assetId).filter(w=>{const wl=w.locationName||(w.projectId?projectName(w.projectId):w.warehouseName||"Kho trung tâm");return wl===loc&&String(w.date||"").slice(0,10)<=untilDate;}).slice().sort((a,b)=>{const da=String(a.date||""),db=String(b.date||"");if(da!==db)return da.localeCompare(db);const ta=a.type==="nhap"?0:1,tb=b.type==="nhap"?0:1;if(ta!==tb)return ta-tb;return String(a.id||"").localeCompare(String(b.id||""));});
        const lots=[];
        source.forEach(w=>{const q=Math.max(0,Number(w.quantity)||0);if(w.type==="nhap")lots.push({rowId:w.id,voucherNo:w.voucherNo,date:w.date,qty:q,remain:q,unitCost:Number(w.unitCost)||0,sourceLocation:loc});else{let left=q;for(const lot of lots){if(left<=0)break;const take=Math.min(left,Math.max(0,lot.remain));lot.remain-=take;left-=take;}}});
        return lots.filter(l=>l.remain>0.0000001);
      };
      let fifoAllocations=[],effectiveUnitCost=enteredUnitCost;
      if(form.type === "xuat") {
        const lots=fifoLotsFor(asset.id,locationName,date);
        const available=lots.reduce((n,l)=>n+l.remain,0);
        if(qty>available+0.0000001){ notify(`Không đủ tồn ${asset.code} tại ${locationName}. Hiện còn ${available}`); return false; }
        let left=qty,totalCost=0;
        for(const lot of lots){if(left<=0)break;const take=Math.min(left,lot.remain);if(take>0){fifoAllocations.push({inRowId:lot.rowId,inVoucherNo:lot.voucherNo,inDate:lot.date,quantity:take,unitCost:lot.unitCost,sourceLocation:locationName});totalCost+=take*lot.unitCost;left-=take;}}
        effectiveUnitCost=qty>0?totalCost/qty:0;
      }
      const base={id:uid("wh"),voucherNo,assetId:asset.id,type:form.type,quantity:qty,date,unitCost:effectiveUnitCost,total:qty*effectiveUnitCost,unit:asset.unit||item.unit||"Cái",receiver:form.receiver||"",supplier:form.supplier||"",description:form.description||form.note||"",note:form.note||"",category:asset.category||"Khác",assetGroup:asset.assetGroup||"Thiết bị chính",ownership:asset.ownership||"Công ty",locationType:form.locationType||"project",locationName,warehouseName:form.locationType==="warehouse"?locationName:"",projectId:form.projectId||null,itemName:asset.name,itemCode:asset.code,operationType,operationLabel,counterpartyLocation,repairVendor:form.repairVendor||"",transferId,address:form.address||"",referenceNo:form.referenceNo||"",attachedDoc:form.attachedDoc||"",transportPerson:form.transportPerson||"",vehicle:form.vehicle||"",orderNo:form.orderNo||"",fifoAllocations};
      rows.push(base);
      if(operationType === "luan_chuyen_di") {
        fifoAllocations.forEach((alloc,ai)=>pairedRows.push({...base,id:uid("wh"),type:"nhap",quantity:alloc.quantity,unitCost:alloc.unitCost,total:alloc.quantity*alloc.unitCost,operationType:"luan_chuyen_den",operationLabel:OPERATION_LABELS.luan_chuyen_den,locationType:form.counterpartyLocationType||"project",locationName:counterpartyLocation,warehouseName:form.counterpartyLocationType==="warehouse"?counterpartyLocation:"",projectId:form.counterpartyProjectId||null,counterpartyLocation:locationName,supplier:"",receiver:form.receiver||"",sourceInboundVoucher:alloc.inVoucherNo||"",sourceInboundDate:alloc.inDate||"",fifoAllocations:[alloc],transferLayerIndex:ai+1}));
      }
    }
    if(!rows.length) { notify("Không có dòng tài sản hợp lệ"); return false; }

    const allRows=[...rows,...pairedRows];
    const txs=allRows.map(r=>({id:uid("tx"),assetId:r.assetId,type:r.operationType?.startsWith("luan_chuyen")?"luan_chuyen":(r.type==="nhap"?"nhap_kho":"xuat_kho"),date:r.date,title:`${r.operationLabel} ${voucherNo}`,detail:`${r.itemName} · ${r.locationName}${r.counterpartyLocation?` ↔ ${r.counterpartyLocation}`:""} · ${r.description||""}`,amount:r.total}));
    let nextAssets=data.assets, nextRepairs=data.repairs;
    if(operationType === "sua_chua") {
      const affected=new Set(rows.map(r=>r.assetId));
      nextAssets=data.assets.map(a=>affected.has(a.id)?{...a,status:STATUS.REPAIR}:a);
      nextRepairs=[...rows.map(r=>({id:uid("rp"),assetId:r.assetId,date:r.date,description:r.description||`Xuất sửa chữa theo ${voucherNo}`,cost:0,status:"Đang sửa",vendor:form.repairVendor||"",warehouseVoucherNo:voucherNo})),...data.repairs];
    } else if(operationType === "thu_hoi_sua_chua") {
      const affected=new Set(rows.map(r=>r.assetId));
      nextAssets=data.assets.map(a=>affected.has(a.id)?{...a,status:STATUS.UNUSED}:a);
      nextRepairs=data.repairs.map(r=>affected.has(r.assetId)&&r.status==="Đang sửa"?{...r,status:"Hoàn thành",completeDate:date,returnVoucherNo:voucherNo}:r);
    }
    setData({...data,assets:nextAssets,repairs:nextRepairs,warehouse:[...allRows,...currentRows],transactions:[...txs,...data.transactions],activityLog:logAction(data.activityLog,`${operationLabel} ${voucherNo} — ${rows.length} mã tài sản — ${locationName}${counterpartyLocation?` ↔ ${counterpartyLocation}`:""}`)});
    notify(`Đã lưu ${operationLabel.toLowerCase()} ${voucherNo} (${rows.length} mã)`); return true;
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
          const operationType = parseOperationType(r["Loại nghiệp vụ"] || "", type);
          const operationLabel = OPERATION_LABELS[operationType] || (type === "nhap" ? "Mua mới bên ngoài" : "Xuất dùng tại công trình");
          const code = String(r["Mã hàng"] || "").trim().toLowerCase();
          const name = String(r["Tên tài sản"] || "").trim().toLowerCase();
          const asset = byCode[code] || byName[name];
          if (!asset) { errors.push(`Dòng ${i + 2}: không tìm thấy tài sản theo Mã hàng/Tên tài sản`); return; }
          const qty = Number(r["Số lượng"] || 0);
          if (qty <= 0) { errors.push(`Dòng ${i + 2}: số lượng không hợp lệ`); return; }
          const unitCost = Number(r["Đơn giá"] || 0);
          const date = parseDateValue(r["Ngày tháng"]);
          const locationText = String(r["Kho/Công trình"] || "").trim();
          const project = data.projects.find(p => p.name.trim().toLowerCase() === locationText.toLowerCase());
          const locationType = String(r["Loại địa điểm"] || "").toLowerCase().includes("kho") ? "warehouse" : (project ? "project" : "warehouse");
          const locationName = project ? project.name : (locationText || "Kho trung tâm");
          const prefix = type === "nhap" ? "PN" : "PX";
          const dateKey = date.replaceAll("-", "");
          const seq = [...existing, ...imported].filter(w => String(w.voucherNo || "").startsWith(`${prefix}-${dateKey}-`)).length + 1;
          const voucherNo = String(r["Số phiếu"] || "").trim() || `${prefix}-${dateKey}-${String(seq).padStart(3,"0")}`;
          imported.push({ id: uid("wh"), voucherNo, assetId: asset.id, type, quantity: qty, date, unitCost, total: qty * unitCost,
            unit: String(r["Đơn vị tính"] || asset.unit || "Cái"), receiver: String(r["Người nhận"] || r["Người giao/nhận"] || ""), supplier: String(r["Nhà cung cấp"] || ""), repairVendor: String(r["Đơn vị sửa chữa"] || ""), description: String(r["Diễn giải"] || r["Ghi chú"] || ""), note: String(r["Ghi chú"] || ""),
            category: asset.category || "Khác", assetGroup: asset.assetGroup || "Thiết bị chính", ownership: asset.ownership || "Công ty", operationType, operationLabel, counterpartyLocation: String(r["Kho/Công trình đối ứng"] || ""), transferId: "",
            locationType, locationName, warehouseName: locationType === "warehouse" ? locationName : "", projectId: project?.id || null,
            itemName: asset.name, itemCode: asset.code });
        });
        if (errors.length) notify(`Import ${imported.length} dòng; ${errors.length} dòng lỗi. ${errors[0]}`);
        if (!imported.length) return;
        setData({ ...data, warehouse: [...imported, ...existing], transactions: [...imported.map(tx => ({ id: uid("tx"), assetId: tx.assetId, type: tx.operationType?.startsWith("luan_chuyen") ? "luan_chuyen" : (tx.type === "nhap" ? "nhap_kho" : "xuat_kho"), date: tx.date, title: `${tx.operationLabel || (tx.type === "nhap" ? "Nhập kho" : "Xuất kho")} ${tx.voucherNo}`, detail: `${tx.itemName} · ${tx.locationName}${tx.counterpartyLocation ? ` ↔ ${tx.counterpartyLocation}` : ""} · ${tx.description || tx.receiver}`, amount: tx.total })), ...data.transactions], activityLog: logAction(data.activityLog, `Import ${imported.length} phiếu kho từ Excel`) });
        if (!errors.length) notify(`Đã import ${imported.length} phiếu kho`);
      } catch (err) { notify("Không đọc được file Excel phiếu kho"); }
    };
    reader.readAsArrayBuffer(file);
  };

  const deleteWarehouseRows = (ids) => {
    if (!requireAdmin()) return; if(!ids.length) return;
    if(!window.confirm(`Xóa ${ids.length} dòng phiếu đã chọn? Thao tác này sẽ thay đổi số tồn kho.`)) return;
    const removed=(data.warehouse||[]).filter(w=>ids.includes(w.id));
    const keys=new Set(removed.map(w=>`${w.assetId}|${w.voucherNo}|${w.type}|${w.date}`));
    setData({...data,warehouse:(data.warehouse||[]).filter(w=>!ids.includes(w.id)),transactions:data.transactions.filter(t=>!removed.some(w=>t.assetId===w.assetId&&t.date===w.date&&String(t.title||"").includes(w.voucherNo))),activityLog:logAction(data.activityLog,`Xóa hàng loạt ${ids.length} dòng phiếu kho`)}); notify(`Đã xóa ${ids.length} dòng`);
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
          const imported = rows.map((r,i)=>({ id: uid("as"), code: String(r["Mã quản lý"]||r.code||`IMP-${Date.now()}-${i}`), name: String(r["Tên tài sản"]||r.name||""), category: String(r["Loại"]||r.category||"Khác"), assetGroup: String(r["Nhóm tài sản"]||r.assetGroup||r["Loại"]||"Thiết bị chính"), ownership: String(r["Nguồn"]||r.ownership||"Công ty"), cost:Number(r["Nguyên giá"]||r.cost||0), purchaseDate:parseDateValue(r["Ngày mua"]||r.purchaseDate), usefulLifeMonths:Number(r["Thời gian SD"]||r.usefulLifeMonths||36), status:STATUS.UNUSED, assignedTo:null, projectId:null, department:String(r["Bộ phận"]||"Vận hành"), serial:String(r["Serial"]||""), supplier:String(r["Nhà cung cấp"]||""), warranty:false, warrantyEnd:"", note:String(r["Ghi chú"]||""), quantity:Number(r["Số lượng"]||r.quantity||1), unit:String(r["Đơn vị tính"]||r.unit||"Cái"), customFields:{} }));
          const wh = imported.map(a=>({id:uid("wh"),assetId:a.id,type:"nhap",quantity:a.quantity||1,date:a.purchaseDate,unitCost:a.cost,note:"Nhập từ Excel"}));
          setData({...data, assets:[...imported,...data.assets], warehouse:[...wh,...data.warehouse], activityLog:logAction(data.activityLog,`Import ${imported.length} tài sản từ Excel`)}); notify(`Đã import ${imported.length} tài sản`);
        } else if (kind === "projects") {
          const imported=rows.map(r=>({id:uid("p"),commander:String(r["Chỉ huy trưởng"]||r.commander||""),name:String(r["Tên công trình"]||r.name||""),address:String(r["Địa chỉ"]||r.address||""),workItem:String(r["Hạng mục thi công"]||r.workItem||""),startDate:parseDateValue(r["Ngày bắt đầu"]||r.startDate),endDate:parseDateValue(r["Ngày kết thúc"]||r.endDate)})).filter(x=>x.name.trim());
          setData({...data,projects:[...imported,...data.projects],activityLog:logAction(data.activityLog,`Import ${imported.length} công trình từ Excel`)}); notify(`Đã import ${imported.length} công trình`);
        } else if (kind === "suppliers") {
          const existing = new Set((data.settings.suppliers || []).map(x => normalizeText(x.code || x.name)));
          const imported = rows.map((r,i)=>({
            id:uid("sup"), code:String(r["Mã NCC"]||r["Mã nhà cung cấp"]||r.code||`NCC-${String(Date.now()).slice(-5)}-${i+1}`).trim(),
            name:String(r["Tên nhà cung cấp"]||r["Nhà cung cấp"]||r.name||"").trim(), taxCode:String(r["Mã số thuế"]||r.taxCode||"").trim(),
            phone:String(r["Điện thoại"]||r.phone||"").trim(), address:String(r["Địa chỉ"]||r.address||"").trim(),
            contact:String(r["Người liên hệ"]||r.contact||"").trim(), note:String(r["Ghi chú"]||r.note||"").trim()
          })).filter(x=>x.name && !existing.has(normalizeText(x.code || x.name)));
          setData({...data,settings:{...data.settings,suppliers:[...imported,...(data.settings.suppliers||[])]},activityLog:logAction(data.activityLog,`Import ${imported.length} nhà cung cấp từ Excel`)}); notify(`Đã import ${imported.length} nhà cung cấp`);
        } else if (kind === "categories") {
          const vals=[...new Set(rows.map(r=>String(r["Danh mục"]||r["Loại tài sản"]||r["Tên danh mục"]||r.name||"").trim()).filter(Boolean))];
          const added=vals.filter(x=>!data.settings.categories.includes(x));
          setData({...data,settings:{...data.settings,categories:[...data.settings.categories,...added]},activityLog:logAction(data.activityLog,`Import ${added.length} danh mục tài sản từ Excel`)}); notify(`Đã import ${added.length} danh mục`);
        } else if (kind === "departments") {
          const vals=[...new Set(rows.map(r=>String(r["Bộ phận"]||r["Tên bộ phận"]||r["Đơn vị sử dụng"]||r.name||"").trim()).filter(Boolean))];
          const added=vals.filter(x=>!data.settings.departments.includes(x));
          setData({...data,settings:{...data.settings,departments:[...data.settings.departments,...added]},activityLog:logAction(data.activityLog,`Import ${added.length} bộ phận từ Excel`)}); notify(`Đã import ${added.length} bộ phận`);
        }
      } catch(err){ notify("Không đọc được file Excel"); }
    }; reader.readAsArrayBuffer(file);
  };

  /* ---------- export ---------- */

  const openReportDesigner = (mode, filename, title, headers, rows) => {
    setReportDesigner({ mode, filename, title, headers, rows, selected: headers.map((_,i)=>i) });
  };
  const doExportExcel = (filename, headers, rows) => openReportDesigner("excel", filename, filename.replaceAll("-", " ").toUpperCase(), headers, rows);
  const doExportPdf = (title, headers, rows) => openReportDesigner("pdf", title.toLowerCase().replace(/[^a-z0-9]+/gi,"-"), title, headers, rows);

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
            <div className="flex-1 min-w-0 p-7 overflow-y-auto aa-scroll">
              {active === "overview" && <Overview data={data} counts={counts} projectName={projectName} />}
              {active === "catalog" && (
                <AssetCatalog
                  assets={filteredAssets} projectName={projectName} onSelect={setSelectedAssetId} selectedAssetId={selectedAssetId}
                  onAdd={() => setModal({ type: "addAsset" })} customColumns={settings.customColumns} isAdmin={isAdmin} onDelete={deleteAsset}
                  onExportExcel={doExportExcel} onExportPdf={doExportPdf} onDeleteMany={deleteAssets} onImportExcel={(file)=>importExcel(file,"assets")}
                />
              )}
              {active === "byProject" && <ByProject data={data} projectName={projectName} onSelect={(id) => { setActive("catalog"); setSelectedAssetId(id); }} />}
              {active === "depreciation" && <Depreciation assets={data.assets} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "repair" && <RepairView repairs={data.repairs.filter((r) => r.status === "Đang sửa")} assetsById={assetsById} onComplete={completeRepair} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "repairHistory" && <RepairHistory repairs={data.repairs} assetsById={assetsById} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "liquidation" && <LiquidationView liquidations={data.liquidations} assetsById={assetsById} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "minutes" && <MinutesView minutes={data.minutes} assetsById={assetsById} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "transactions" && <TransactionsView transactions={data.transactions} assetsById={assetsById} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
{active === "warehouse" && <WarehouseBoundary><WarehouseHub
                warehouse={data.warehouse || []} assets={data.assets} projects={data.projects} settings={settings} suppliers={settings.suppliers || []}
                onAdd={(type) => setModal({ type: type === "nhap" ? "warehouseIn" : type === "transfer" ? "warehouseTransfer" : "warehouseOut" })}
                onExportExcel={doExportExcel} onExportPdf={doExportPdf}
                onImport={importWarehouseExcel} onDeleteRows={deleteWarehouseRows} isAdmin={isAdmin}
              /></WarehouseBoundary>}
              {active === "assetCategories" && <MasterDataPage
                title="Danh mục / loại tài sản"
                description="Quản lý danh sách loại tài sản dùng trực tiếp tại trường Danh mục khi thêm/sửa tài sản và trong bộ lọc báo cáo."
                items={settings.categories || []}
                usage={(name) => data.assets.filter((a) => a.category === name).length}
                onAdd={addCategory} onRename={renameCategory} onRemove={removeCategory}
                importKind="categories" onImportExcel={(file, kind)=>importExcel(file, kind)}
                templateHeader="Danh mục" placeholder="VD: TBT, VTT, Máy móc thiết bị"
                isAdmin={isAdmin}
              />}
              {active === "departments" && <MasterDataPage
                title="Bộ phận / đơn vị sử dụng"
                description="Quản lý danh sách bộ phận dùng trực tiếp khi thêm/sửa, cấp phát và điều chuyển tài sản."
                items={settings.departments || []}
                usage={(name) => data.assets.filter((a) => a.department === name).length}
                onAdd={addDepartment} onRename={renameDepartment} onRemove={removeDepartment}
                importKind="departments" onImportExcel={(file, kind)=>importExcel(file, kind)}
                templateHeader="Bộ phận" placeholder="VD: Vận hành, Kế toán, Cơ điện"
                isAdmin={isAdmin}
              />}
              {active === "suppliers" && <SupplierCatalog suppliers={settings.suppliers || []} isAdmin={isAdmin} onAdd={addSupplier} onEdit={editSupplier} onDeleteMany={deleteSuppliers} onImportExcel={(file)=>importExcel(file,"suppliers")} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "costHistory" && <CostHistoryView costHistory={data.costHistory} assetsById={assetsById} onAdd={() => setModal({ type: "costHistory" })} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "projects" && <ProjectsView projects={data.projects} assets={data.assets} isAdmin={isAdmin} onDelete={deleteProject} onAdd={() => setModal({ type: "addProject" })} onExportExcel={doExportExcel} onExportPdf={doExportPdf} />}
              {active === "activityLog" && <ActivityLogView log={data.activityLog} />}
              {active === "settings" && isAdmin && (
                <SettingsView
                  settings={settings} onSetCompanyName={setCompanyName}
                  onAddCategory={addCategory} onRemoveCategory={removeCategory} onRenameCategory={renameCategory}
                  onAddDepartment={addDepartment} onRemoveDepartment={removeDepartment} onRenameDepartment={renameDepartment}
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
          <AssetFormModal title="Thêm tài sản mới" categories={settings.categories} departments={settings.departments} projects={data.projects} suppliers={settings.suppliers || []} customColumns={settings.customColumns}
            onClose={() => setModal(null)} onSubmit={(f) => { addAsset(f); setModal(null); }} />
        )}
        {modal?.type === "editAsset" && (
          <AssetFormModal title={`Sửa tài sản — ${assetsById[modal.assetId]?.code}`} initial={assetsById[modal.assetId]}
            categories={settings.categories} departments={settings.departments} projects={data.projects} suppliers={settings.suppliers || []} customColumns={settings.customColumns}
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
        {(modal?.type === "warehouseIn" || modal?.type === "warehouseOut" || modal?.type === "warehouseTransfer") && <WarehouseTxModal
          title={modal.type === "warehouseIn" ? "Phiếu nhập kho" : modal.type === "warehouseTransfer" ? "Phiếu chuyển kho" : "Phiếu xuất kho"}
          fixedType={modal.type === "warehouseIn" ? "nhap" : "xuat"}
          fixedOperation={modal.type === "warehouseTransfer" ? "luan_chuyen_di" : ""}
          assets={data.assets} projects={data.projects} suppliers={settings.suppliers || []} onClose={()=>setModal(null)} onSubmit={f=>{if(addWarehouseTx(f))setModal(null)}}
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

      {reportDesigner && <ReportDesigner job={reportDesigner} companyName={settings.companyName} onClose={()=>setReportDesigner(null)} onPrint={(job)=>{setReportDesigner(null);setPrintJob(job)}} onExcel={async(job)=>{try{await exportStyledExcel(job.filename,job.title,job.headers,job.rows,settings.companyName);notify("Đã xuất Excel theo mẫu tùy chỉnh");setReportDesigner(null)}catch(e){console.error(e);notify("Xuất Excel thất bại")}}} />}
      {printJob && <PrintArea job={printJob} companyName={settings.companyName} />}
    </>
  );
}


function ReportDesigner({ job, companyName, onClose, onExcel, onPrint }) {
  const [selected,setSelected]=useState(job.selected || job.headers.map((_,i)=>i));
  const toggle=(i)=>setSelected(v=>v.includes(i)?v.filter(x=>x!==i):[...v,i]);
  const move=(pos,dir)=>setSelected(v=>{const n=[...v],to=pos+dir;if(to<0||to>=n.length)return n;[n[pos],n[to]]=[n[to],n[pos]];return n});
  const chosen=selected.map(i=>job.headers[i]);
  const rows=job.rows.map(r=>selected.map(i=>r[i]));
  const run=()=>{if(!selected.length)return alert("Hãy chọn ít nhất 1 cột báo cáo."); const out={filename:job.filename,title:job.title,headers:chosen,rows}; job.mode==="excel"?onExcel(out):onPrint({...out,rows:rows.map(r=>r.map(c=>c==null?"":String(c)))})};
  return <Modal title="Thiết kế báo cáo trước khi xuất" onClose={onClose} wide>
    <div className="grid grid-cols-[1fr_1.15fr] gap-5">
      <div><div className="flex items-center justify-between mb-2"><div><div className="text-[13px] font-semibold">Cột thông tin</div><div className="text-[11px]" style={{color:TOKENS.muted}}>Chọn cột và sắp xếp thứ tự xuất.</div></div><button className="text-[11px] font-semibold" style={{color:TOKENS.brand}} onClick={()=>setSelected(job.headers.map((_,i)=>i))}>Mẫu mặc định</button></div>
        <div className="rounded-lg overflow-hidden" style={{border:`1px solid ${TOKENS.border}`}}>{job.headers.map((h,i)=>{const pos=selected.indexOf(i),on=pos>=0;return <div key={`${h}-${i}`} className="flex items-center gap-2 px-3 py-2" style={{borderBottom:`1px solid ${TOKENS.border}`,background:on?TOKENS.brandSoft:"#fff"}}><input type="checkbox" checked={on} onChange={()=>toggle(i)}/><span className="flex-1 text-[12px]">{h}</span>{on&&<><span className="text-[10px] aa-mono" style={{color:TOKENS.muted}}>#{pos+1}</span><button disabled={pos===0} onClick={()=>move(pos,-1)} className="p-1 disabled:opacity-25"><ArrowUp size={14}/></button><button disabled={pos===selected.length-1} onClick={()=>move(pos,1)} className="p-1 disabled:opacity-25"><ArrowDown size={14}/></button></>}</div>})}</div>
      </div>
      <div><div className="flex items-center gap-2 mb-2"><SlidersHorizontal size={15}/><div className="text-[13px] font-semibold">Xem trước mẫu báo cáo</div></div><div className="rounded-lg bg-white p-4 overflow-auto max-h-[520px]" style={{border:`1px solid ${TOKENS.border}`}}><div className="text-[10px] font-bold" style={{color:TOKENS.brand}}>{String(companyName||"MYHL").toUpperCase()}</div><div className="text-center font-extrabold text-[15px] uppercase my-2">{job.title}</div><div className="text-center text-[9px] mb-3" style={{color:TOKENS.muted}}>Ngày xuất: {new Date().toLocaleString("vi-VN")}</div><table className="w-full border-collapse text-[8px]"><thead><tr>{chosen.map((h,i)=><th key={i} className="p-1 text-white text-center" style={{background:TOKENS.brand,border:"1px solid #D0D5DD"}}>{h}</th>)}</tr></thead><tbody>{rows.slice(0,6).map((r,ri)=><tr key={ri}>{r.map((c,ci)=><td key={ci} className="p-1" style={{border:"1px solid #D0D5DD",background:ri%2?"#FFF7F7":"#fff"}}>{String(c??"")}</td>)}</tr>)}</tbody></table><div className="text-[8px] mt-2" style={{color:TOKENS.muted}}>Xem trước 6/{rows.length} dòng • File chính thức có kẻ bảng, tiêu đề, bộ lọc và căn cột.</div></div></div>
    </div>
    <div className="flex justify-between items-center mt-5 pt-4" style={{borderTop:`1px solid ${TOKENS.border}`}}><div className="text-[11px]" style={{color:TOKENS.muted}}>Đã chọn <b>{selected.length}/{job.headers.length}</b> cột.</div><div className="flex gap-2"><Btn onClick={onClose}>Hủy</Btn><Btn kind="primary" icon={job.mode==="excel"?FileDown:Printer} onClick={run}>{job.mode==="excel"?"Xuất Excel theo mẫu":"Xuất PDF theo mẫu"}</Btn></div></div>
  </Modal>
}

const PRINT_CSS = `
#aa-print-area{ position:fixed; left:-99999px; top:0; }
@media print {
  @page { size: A4 landscape; margin: 10mm; }
  .aa-noprint{ display:none !important; }
  #aa-print-area{ position:static !important; left:auto !important; width:auto !important; }
  #aa-print-area thead{display:table-header-group;}
  #aa-print-area tr{page-break-inside:avoid;}
}
`;

function PrintArea({ job, companyName }) {
  return (
    <div id="aa-print-area" style={{ background: "#fff", color: "#111827", padding: 24, width: 1120, fontFamily: "Arial, sans-serif" }}>
      <div style={{ borderBottom: `3px solid ${TOKENS.brand}`, paddingBottom: 10, marginBottom: 12, display:"flex", justifyContent:"space-between", alignItems:"flex-end" }}>
        <div><div style={{ fontWeight: 800, color: TOKENS.brand, fontSize: 13, textTransform:"uppercase" }}>{companyName}</div><div style={{ fontWeight: 800, fontSize: 21, marginTop: 3, textTransform:"uppercase" }}>{job.title}</div></div>
        <div style={{ fontSize: 10.5, color: "#667085", textAlign:"right" }}>Ngày xuất báo cáo<br/><b>{new Date().toLocaleString("vi-VN")}</b></div>
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10.5 }}>
        <thead><tr>{job.headers.map((h) => <th key={h} style={{ border: `1px solid ${TOKENS.brand}`, padding: "7px 6px", textAlign: "center", background: TOKENS.brand, color:"#fff", fontWeight:700 }}>{h}</th>)}</tr></thead>
        <tbody>{job.rows.map((r, i) => <tr key={i} style={{background:i%2?"#FFF7F7":"#fff"}}>{r.map((c, j) => <td key={j} style={{ border: "1px solid #D0D5DD", padding: "6px 6px", verticalAlign:"top" }}>{c}</td>)}</tr>)}</tbody>
      </table>
      <div style={{marginTop:12,fontSize:9.5,color:"#667085",display:"flex",justifyContent:"space-between"}}><span>MYHL - Hệ thống quản lý tài sản</span><span>Tổng số dòng: {job.rows.length}</span></div>
    </div>
  );
}

/* ============================== SIDEBAR / TOPBAR ============================== */

function Sidebar({ active, setActive, data, counts, currentUser, isAdmin, onLogout, onChangePassword }) {
  return (
    <aside className="w-[252px] shrink-0 flex flex-col bg-white" style={{ borderRight: `1px solid ${TOKENS.border}`, boxShadow: "4px 0 18px rgba(16,24,40,.025)" }}>
      <div className="px-5 pt-5 pb-4" style={{ borderBottom: `1px solid ${TOKENS.border}` }}>
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl flex items-center justify-center shrink-0" style={{ background: TOKENS.brand, boxShadow: `0 6px 14px ${TOKENS.brand}35` }}>
            <ShieldCheck size={20} color="#fff" strokeWidth={2.3} />
          </div>
          <div className="min-w-0">
            <div className="aa-display font-extrabold text-[15px] leading-tight tracking-[-.03em]">MYHL</div>
            <div className="text-[10.5px] font-semibold tracking-wide truncate" style={{ color: TOKENS.muted }}>QUẢN LÝ TÀI SẢN</div>
          </div>
        </div>
        <div className="mt-4 px-3 py-2.5 rounded-xl" style={{ background: TOKENS.brandSoft }}>
          <div className="text-[10px] font-semibold uppercase tracking-wider" style={{ color: TOKENS.brand }}>Đơn vị quản lý</div>
          <div className="text-[12px] font-medium truncate mt-0.5" style={{ color: TOKENS.ink }} title={data.settings.companyName}>{data.settings.companyName}</div>
        </div>
      </div>

      <div className="flex-1 overflow-y-auto aa-scroll py-4 px-3">
        {NAV.map((sec) => (
          <div key={sec.section} className="mb-5">
            <div className="px-3 mb-2 text-[10px] tracking-[.14em] font-bold" style={{ color: "#98A2B3" }}>{sec.section}</div>
            <div className="space-y-1">
              {sec.items.filter((it) => !it.adminOnly || isAdmin).map((it) => {
                const isActive = active === it.id;
                const count = it.badge === "assets" ? counts.total : it.badge === "projects" ? data.projects.length : null;
                return (
                  <button key={it.id} onClick={() => setActive(it.id)}
                    className={`aa-sidebar-item w-full flex items-center gap-3 px-3 py-2.5 rounded-xl text-[12.5px] relative ${isActive ? "aa-active" : ""}`}
                    style={{ background: isActive ? TOKENS.brand : "transparent", color: isActive ? "#fff" : TOKENS.ink, fontWeight: isActive ? 700 : 500 }}>
                    <it.icon size={16} strokeWidth={isActive ? 2.2 : 1.9} />
                    <span className="flex-1 text-left">{it.label}</span>
                    {count != null && <span className="aa-mono text-[10px] rounded-full px-1.5 py-0.5" style={{ background: isActive ? "#ffffff26" : TOKENS.infoSoft, color: isActive ? "#fff" : TOKENS.muted }}>{count}</span>}
                    {isActive && <ChevronRight size={14} className="opacity-80" />}
                  </button>
                );
              })}
            </div>
          </div>
        ))}
      </div>

      <div className="px-4 py-4" style={{ borderTop: `1px solid ${TOKENS.border}` }}>
        <div className="flex items-center gap-2.5 p-2.5 rounded-xl" style={{ background: TOKENS.paper }}>
          <div className="w-9 h-9 rounded-full flex items-center justify-center aa-display font-bold text-[12px] shrink-0" style={{ background: TOKENS.brandSoft, color: TOKENS.brand }}>
            {currentUser.name[0]?.toUpperCase()}
          </div>
          <div className="min-w-0 flex-1">
            <div className="text-[12px] font-semibold truncate">{currentUser.name}</div>
            <div className="text-[10px] truncate" style={{ color: TOKENS.muted }}>{ROLES[currentUser.role].label}</div>
          </div>
          <button onClick={onChangePassword} title="Đổi mật khẩu" className="p-1.5 rounded-lg hover:bg-white" style={{ color: TOKENS.muted }}><Lock size={14} /></button>
          <button onClick={onLogout} title="Đăng xuất" className="p-1.5 rounded-lg hover:bg-white" style={{ color: TOKENS.muted }}><LogOut size={14} /></button>
        </div>
        <div className="mt-1 text-center aa-mono text-[8px]" style={{ color: "#98A2B3" }}>Core {CORE_VERSION}</div>
      </div>
    </aside>
  );
}

function TopBar({ query, setQuery, data, saving, onRefresh }) {
  return (
    <header className="flex items-center justify-between gap-4 px-7 py-4 shrink-0 bg-white" style={{ borderBottom: `1px solid ${TOKENS.border}` }}>
      <div className="min-w-0">
        <div className="aa-display text-[17px] font-bold truncate">Xin chào, Admin! <span aria-hidden="true">👋</span></div>
        <div className="text-[11px] mt-0.5" style={{ color: TOKENS.muted }}>Quản lý tài sản, kho và công trình trên một hệ thống.</div>
      </div>
      <div className="flex items-center gap-3 shrink-0">
        <div className="relative w-[330px] max-w-[36vw]">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: TOKENS.muted }} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Tìm tài sản, mã, công trình…"
            className="w-full rounded-xl pl-9 pr-3 py-2 text-[12.5px] bg-white" style={{ border: `1px solid ${TOKENS.border}` }} />
        </div>
        <div className="hidden md:flex items-center gap-2 px-3 py-2 rounded-xl" style={{ background: TOKENS.paper, color: TOKENS.muted }}>
          <span className="w-2 h-2 rounded-full" style={{ background: saving ? "#F79009" : "#12B76A" }} />
          <span className="text-[11px] font-medium">{saving ? "Đang lưu…" : "Đã đồng bộ"}</span>
        </div>
        <Btn icon={RefreshCw} onClick={onRefresh} small>Làm mới</Btn>
      </div>
    </header>
  );
}

/* ============================== OVERVIEW ============================== */

function Overview({ data, counts, projectName }) {
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

  return (
    <div className="aa-fade max-w-5xl">
      <div className="flex items-center justify-between mb-4">
        <h1 className="aa-display text-xl font-semibold">Tổng quan</h1>
        <div className="flex items-center gap-2 rounded-full px-3 py-1" style={{ background: TOKENS.goldSoft, color: TOKENS.gold }}>
          <Stamp size={14} />
          <span className="aa-display text-[13px] font-semibold">{counts.total} tài sản</span>
        </div>
      </div>

      <div className="grid grid-cols-4 gap-3 mb-3">
        <StatCard label="Đang rảnh" value={counts[STATUS.UNUSED]} />
        <StatCard label="Dùng chung" value={counts[STATUS.SHARED]} accent={TOKENS.info} />
        <StatCard label="Đang sửa/bảo dưỡng" value={counts[STATUS.REPAIR]} accent={TOKENS.gold} />
        <StatCard label="Hỏng/mất" value={counts[STATUS.BROKEN]} accent={TOKENS.danger} />
      </div>
      <div className="grid grid-cols-4 gap-3 mb-6">
        <StatCard label="Đã điều chuyển" value={counts[STATUS.TRANSFERRED_OUT]} />
        <StatCard label="Còn bảo hành" value={counts.warranty} accent={TOKENS.brand} />
        <StatCard label="Tổng công trình" value={data.projects.length} />
        <StatCard label="Chi phí sửa chữa 2026" value={fmtVND(counts.repairCostThisYear)} accent={TOKENS.gold} />
      </div>

      <div className="grid grid-cols-2 gap-4">
        <div className="rounded-lg p-5" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
          <div className="text-[13px] font-medium mb-4">Phân bổ tài sản</div>
          <div className="space-y-2.5">
            {breakdown.map((b) => (
              <div key={b.label} className="flex items-center gap-3">
                <div className="w-[150px] text-[12px] shrink-0" style={{ color: TOKENS.muted }}>{b.label}</div>
                <div className="flex-1 h-2 rounded-full" style={{ background: TOKENS.paper }}>
                  <div className="h-2 rounded-full" style={{ width: `${(b.value / max) * 100}%`, background: STATUS_COLOR[b.label] }} />
                </div>
                <div className="w-6 text-right aa-mono text-[12px]">{b.value}</div>
              </div>
            ))}
          </div>
        </div>

        <div className="rounded-lg p-5" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
          <div className="text-[13px] font-medium mb-4">Phiếu đang xử lý</div>
          {counts.openRepairs === 0 ? (
            <EmptyState text="Không có phiếu sửa chữa đang mở" />
          ) : (
            <div className="text-[13px]" style={{ color: TOKENS.muted }}>{counts.openRepairs} phiếu sửa chữa đang chờ xử lý.</div>
          )}
          <div className="mt-5 text-[13px] font-medium mb-2">Gần đây</div>
          <div className="space-y-2">
            {data.activityLog.slice(-4).reverse().map((l) => (
              <div key={l.id} className="text-[12px] flex items-start gap-2" style={{ color: TOKENS.muted }}>
                <span className="aa-mono shrink-0">{fmtDate(l.date)}</span>
                <span>{l.action}</span>
              </div>
            ))}
          </div>
        </div>
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

function AssetCatalog({ assets, projectName, onSelect, selectedAssetId, onAdd, onDelete, onDeleteMany, isAdmin, customColumns = [], onExportExcel, onExportPdf, onImportExcel }) {
  const [selected,setSelected]=useState([]);
  const fileRef=useRef(null);
  const headers = ["Mã quản lý", "Tên tài sản", "ĐVT", "Loại", "Nhóm tài sản", "Công trình", "Nguồn", "Trạng thái", "Nguyên giá", ...customColumns.map((c) => c.label)];
  const buildRows = () => assets.map((a) => [a.code, a.name, a.unit||"Cái", a.category, a.assetGroup || "", projectName(a.projectId), a.ownership || "Công ty", a.status, a.cost, ...customColumns.map((c) => a.customFields?.[c.key] ?? "")]);
  const visibleIds=assets.map(a=>a.id), allSelected=visibleIds.length>0&&visibleIds.every(id=>selected.includes(id));
  const toggleAll=()=>setSelected(allSelected?selected.filter(id=>!visibleIds.includes(id)):[...new Set([...selected,...visibleIds])]);
  const downloadTemplate=()=>downloadExcelTemplate("Mau_Import_Danh_Muc_Tai_San",["Mã quản lý","Tên tài sản","Đơn vị tính","Loại","Nhóm tài sản","Nguồn","Số lượng","Nguyên giá","Ngày mua","Thời gian SD","Bộ phận","Serial","Nhà cung cấp","Ghi chú"],[["TS-001","Máy khoan mẫu","Cái","TBT","Thiết bị điện","Công ty",1,3500000,"24/08/2026",36,"Vận hành","SN001","NCC A",""]]);
  return <div className="aa-fade"><div className="flex items-center justify-between mb-4 gap-3"><div><h1 className="aa-display text-xl font-semibold">Danh mục tài sản</h1><div className="text-[11px] mt-1" style={{color:TOKENS.muted}}>Có thể đổ lô Excel, chọn nhiều dòng và xóa hàng loạt.</div></div><div className="flex items-center gap-2 flex-wrap justify-end"><ExportBar onExcel={() => onExportExcel("danh-muc-tai-san", headers, buildRows())} onPdf={() => onExportPdf("Danh mục tài sản", headers, buildRows())} /><Btn icon={Download} onClick={downloadTemplate}>Tải mẫu</Btn>{isAdmin&&<><input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)onImportExcel?.(f);e.target.value=""}}/><Btn kind="info" icon={UploadCloud} onClick={()=>fileRef.current?.click()}>Đổ Excel</Btn></>}{isAdmin&&selected.length>0&&<Btn kind="danger" icon={Trash2} onClick={()=>{onDeleteMany(selected);setSelected([])}}>Xóa {selected.length} tài sản</Btn>}<Btn kind="primary" icon={Plus} onClick={onAdd}>Thêm mới</Btn></div></div><div className="rounded-lg overflow-x-auto aa-scroll" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}><table className="w-full min-w-[1200px]"><thead><tr><Th><input type="checkbox" checked={allSelected} onChange={toggleAll}/></Th><Th>Mã quản lý</Th><Th>Tên tài sản</Th><Th>ĐVT</Th><Th>Loại</Th><Th>Nhóm tài sản</Th><Th>Công trình</Th><Th>Nguồn</Th><Th>Trạng thái</Th><Th right>Nguyên giá</Th>{customColumns.map((c) => <Th key={c.key}>{c.label}</Th>)}</tr></thead><tbody>{assets.map((a) => <tr key={a.id} className="aa-row cursor-pointer" onClick={() => onSelect(a.id)} style={{ background: selectedAssetId === a.id ? TOKENS.brandSoft : "transparent" }}><Td><input type="checkbox" checked={selected.includes(a.id)} onClick={e=>e.stopPropagation()} onChange={e=>setSelected(e.target.checked?[...selected,a.id]:selected.filter(id=>id!==a.id))}/></Td><Td mono><Tag>{a.code}</Tag></Td><Td>{a.name}</Td><Td>{a.unit||"Cái"}</Td><Td>{a.category}</Td><Td>{a.assetGroup || "—"}</Td><Td>{projectName(a.projectId)}</Td><Td>{a.ownership || "Công ty"}</Td><Td><StatusDot status={a.status} /></Td><Td right mono>{fmtVND(a.cost)}</Td>{customColumns.map((c) => <Td key={c.key} mono={c.type !== "text"}>{c.type === "date" ? fmtDate(a.customFields?.[c.key]) : (a.customFields?.[c.key] || "—")}</Td>)}</tr>)}</tbody></table>{assets.length === 0 && <EmptyState text="Không tìm thấy tài sản phù hợp" sub="Thử từ khoá khác hoặc bỏ bộ lọc." />}</div></div>;
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

function RepairView({ repairs, assetsById, onComplete, onExportExcel, onExportPdf }) {
  const headers=["Mã tài sản","Tên tài sản","Ngày gửi","Mô tả / lỗi","Chi phí","Trạng thái"];
  const rows=repairs.map(r=>{const a=assetsById[r.assetId]||{};return[a.code||"",a.name||"",fmtDate(r.date),r.description||"",Number(r.cost||0),r.status||""]});
  return (
    <div className="aa-fade">
      <div className="flex items-center justify-between mb-4"><div><h1 className="aa-display text-xl font-semibold">Sửa chữa — đang xử lý</h1><div className="text-[11px] mt-1" style={{color:TOKENS.muted}}>Có thể xuất danh sách đang sửa ra Excel/PDF.</div></div><ExportBar onExcel={()=>onExportExcel("sua-chua-dang-xu-ly",headers,rows)} onPdf={()=>onExportPdf("Sửa chữa — đang xử lý",headers,rows)}/></div>
      {repairs.length === 0 ? (
        <div className="rounded-lg" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
          <EmptyState text="Không có phiếu sửa chữa nào đang mở" sub="Tạo phiếu sửa chữa từ trang chi tiết tài sản hoặc xuất đi sửa chữa từ Kho." />
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

class WarehouseBoundary extends Component {
  constructor(props){super(props);this.state={error:null};}
  static getDerivedStateFromError(error){return {error};}
  componentDidCatch(error,info){console.error("Warehouse module error",error,info);}
  render(){if(this.state.error)return <div className="aa-fade rounded-lg p-6" style={{background:TOKENS.surface,border:`1px solid ${TOKENS.border}`}}><AlertCircle size={24} style={{color:TOKENS.danger}}/><div className="font-semibold mt-2">Không thể mở module Kho</div><div className="text-[12px] mt-1" style={{color:TOKENS.muted}}>Module Kho đã có lớp tương thích dữ liệu cũ. Nếu màn hình này vẫn xuất hiện, mã lỗi kỹ thuật bên dưới sẽ cho biết chính xác nguyên nhân.</div><div className="aa-mono text-[11px] mt-3 p-3 rounded" style={{background:TOKENS.paper,color:TOKENS.danger,border:`1px solid ${TOKENS.border}`}}>{safeText(this.state.error?.message,"Lỗi không xác định")}</div><div className="mt-3"><Btn onClick={()=>this.setState({error:null})}>Thử mở lại</Btn></div></div>;return this.props.children;}
}

function WarehouseHub({ warehouse = [], assets = [], projects = [], settings = {}, suppliers = [], onAdd, onExportExcel, onExportPdf, onImport, onDeleteRows, isAdmin }) {
  // Không render trực tiếp dữ liệu thô. Dữ liệu từ các core cũ có thể có field
  // là object/null; chuẩn hóa toàn bộ ngay tại biên module Kho để module luôn mở được.
  const safeAssets = (Array.isArray(assets) ? assets : []).filter(a => a && typeof a === "object").map((a, i) => ({
    ...a,
    id: safeText(a.id, `legacy_asset_${i}`),
    code: safeText(a.code), name: safeText(a.name), serial: safeText(a.serial),
    category: safeText(a.category, "Khác"), assetGroup: safeText(a.assetGroup, "—"),
    ownership: safeText(a.ownership, "Công ty"), unit: safeText(a.unit, "Cái"),
    projectId: safeText(a.projectId),
  }));
  const safeProjects = (Array.isArray(projects) ? projects : []).filter(p => p && typeof p === "object").map((p, i) => ({
    ...p, id: safeText(p.id, `legacy_project_${i}`), name: safeText(p.name, `Công trình ${i + 1}`)
  }));
  const safeWarehouse = (Array.isArray(warehouse) ? warehouse : []).filter(w => w && typeof w === "object").map((w, i) => ({
    ...w,
    id: safeText(w.id, `legacy_wh_${i}`), assetId: safeText(w.assetId), projectId: safeText(w.projectId),
    voucherNo: safeText(w.voucherNo, "—"), date: safeText(w.date),
    type: safeText(w.type) === "xuat" ? "xuat" : "nhap",
    operationType: safeText(w.operationType), operationLabel: safeText(w.operationLabel),
    itemName: safeText(w.itemName), itemCode: safeText(w.itemCode),
    locationName: safeText(w.locationName), warehouseName: safeText(w.warehouseName),
    counterpartyLocation: safeText(w.counterpartyLocation),
    supplier: safeText(w.supplier), repairVendor: safeText(w.repairVendor), receiver: safeText(w.receiver),
    category: safeText(w.category), assetGroup: safeText(w.assetGroup), ownership: safeText(w.ownership),
    unit: safeText(w.unit, "Cái"), description: safeText(w.description), note: safeText(w.note), address:safeText(w.address),referenceNo:safeText(w.referenceNo),attachedDoc:safeText(w.attachedDoc),transportPerson:safeText(w.transportPerson),vehicle:safeText(w.vehicle),orderNo:safeText(w.orderNo),
    quantity: safeNumber(w.quantity), unitCost: safeNumber(w.unitCost),
    total: Number.isFinite(Number(w.total)) ? Number(w.total) : safeNumber(w.quantity) * safeNumber(w.unitCost),
  }));
  const safeSuppliers = (Array.isArray(suppliers) ? suppliers : []).map((s, i) => typeof s === "string" ? { id: `sup_${i}`, name: s } : (s && typeof s === "object" ? { ...s, id: safeText(s.id, `sup_${i}`), name: safeText(s.name || s.label || s.code) } : null)).filter(s => s && s.name);

  const [tab,setTab]=useState("in");
  const [query,setQuery]=useState("");
  const [selectedRows,setSelectedRows]=useState([]);
  const [filter,setFilter]=useState({category:"",group:"",ownership:"",projectId:"",locationName:"",supplier:"",operationType:"",asOfDate:nowIso().slice(0,10)});
  const assetMap=Object.fromEntries(safeAssets.map(a=>[a.id,a]));
  const projectMap=Object.fromEntries(safeProjects.map(p=>[p.id,p]));
  const getLocation=(w)=>safeText(w?.locationName || (w?.projectId ? projectMap[safeText(w.projectId)]?.name : "") || w?.warehouseName, "Kho trung tâm");
  const categories=[...new Set(safeAssets.map(a=>safeText(a.category)).filter(Boolean))];
  const groupsList=[...new Set(safeAssets.map(a=>safeText(a.assetGroup)).filter(Boolean))];
  const ownerships=[...new Set(safeAssets.map(a=>safeText(a.ownership)).filter(Boolean))];
  const locations=[...new Set(safeWarehouse.map(w=>getLocation(w)).filter(Boolean))];
  const supplierNames=[...new Set([...safeSuppliers.map(s=>safeText(s.name)),...safeWarehouse.map(w=>safeText(w.supplier)),...safeWarehouse.map(w=>safeText(w.repairVendor))].filter(Boolean))];
  const operationOptions=[...(WAREHOUSE_OPERATIONS.nhap||[]),...(WAREHOUSE_OPERATIONS.xuat||[])];
  const matchesCommon=(w)=>{
    if(!w) return false;
    const a=assetMap[safeText(w.assetId)]||{};
    if(filter.category && safeText(w.category||a.category)!==filter.category)return false;
    if(filter.group && safeText(w.assetGroup||a.assetGroup)!==filter.group)return false;
    if(filter.ownership && safeText(w.ownership||a.ownership)!==filter.ownership)return false;
    if(filter.projectId && safeText(w.projectId||a.projectId)!==filter.projectId)return false;
    if(filter.locationName && getLocation(w)!==filter.locationName)return false;
    if(filter.supplier && safeText(w.supplier)!==filter.supplier && safeText(w.repairVendor)!==filter.supplier)return false;
    if(filter.operationType && safeText(w.operationType)!==filter.operationType)return false;
    if(query && !normalizeText(`${safeText(w.voucherNo)} ${safeText(w.itemName||a.name)} ${safeText(w.itemCode||a.code)} ${safeText(w.supplier)} ${safeText(w.receiver)} ${safeText(w.operationLabel||OPERATION_LABELS[w.operationType])} ${safeText(w.description||w.note)}`).includes(normalizeText(query))) return false;
    return true;
  };
  const filteredTx=safeWarehouse.filter(matchesCommon);
  const txHeaders=["","Số phiếu","Ngày chứng từ","Loại nghiệp vụ","Mã hàng","Tên hàng","Xuất/Nhập tại Kho-Công trình","Kho-Công trình đối ứng","ĐVT","Số lượng","Đơn giá / Giá vốn","Thành tiền","Đối tượng / NCC","Người giao nhận","Tham chiếu","Diễn giải"];
  const asOf=filter.asOfDate||nowIso().slice(0,10);
  const reportTx=safeWarehouse.filter(w=>!w.date || safeText(w.date).slice(0,10)<=asOf).slice().sort((a,b)=>{const da=safeText(a.date),db=safeText(b.date);if(da!==db)return da.localeCompare(db);const ta=a.type==="nhap"?0:1,tb=b.type==="nhap"?0:1;if(ta!==tb)return ta-tb;return `${safeText(a.voucherNo)}|${safeText(a.id)}`.localeCompare(`${safeText(b.voucherNo)}|${safeText(b.id)}`);});
  const balances={};
  reportTx.forEach(w=>{
    const a=assetMap[safeText(w.assetId)]||{};
    const identity=safeText(w.assetId||w.itemCode||a.code);
    if(!identity)return;
    const loc=getLocation(w), key=`${identity}¦${loc}`;
    if(!balances[key]) balances[key]={code:safeText(w.itemCode||a.code),name:safeText(w.itemName||a.name),location:loc,category:safeText(w.category||a.category,"Khác"),group:safeText(w.assetGroup||a.assetGroup,"—"),ownership:safeText(w.ownership||a.ownership,"Công ty"),unit:safeText(w.unit||a.unit,"Cái"),inQty:0,outQty:0,inValue:0,outValue:0,lastDescription:"",operations:[]};
    const q=safeNumber(w.quantity), v=Number.isFinite(Number(w.total))?Number(w.total):q*safeNumber(w.unitCost);
    if(w.type==="nhap"){balances[key].inQty+=q;balances[key].inValue+=v}else if(w.type==="xuat"){balances[key].outQty+=q;balances[key].outValue+=v}
    const op=safeText(w.operationLabel||OPERATION_LABELS[w.operationType]||(w.type==="nhap"?"Nhập kho":"Xuất kho")); if(op&&!balances[key].operations.includes(op))balances[key].operations.push(op);
    if(w.description||w.note)balances[key].lastDescription=safeText(w.description||w.note);
  });
  const summary=Object.values(balances).filter(r=>!filter.category||r.category===filter.category).filter(r=>!filter.group||r.group===filter.group).filter(r=>!filter.ownership||r.ownership===filter.ownership).filter(r=>!filter.locationName||r.location===filter.locationName).sort((a,b)=>`${a.location}|${a.code}`.localeCompare(`${b.location}|${b.code}`));
  const reportHeaders=["Kho/Công trình","Loại tài sản","Nhóm tài sản","Nguồn gốc","Mã hàng","Tên tài sản","ĐVT","Nhập lũy kế","Xuất lũy kế","Tồn đến ngày","Giá trị tồn","Nghiệp vụ phát sinh","Diễn giải gần nhất"];
  const reportRows=summary.map(r=>[r.location,r.category,r.group,r.ownership,r.code,r.name,r.unit,r.inQty,r.outQty,r.inQty-r.outQty,r.inValue-r.outValue,r.operations.join("; "),r.lastDescription||"—"]);
  const running={};
  const txDetailRows=reportTx.filter(matchesCommon).map(w=>{const a=assetMap[safeText(w.assetId)]||{},loc=getLocation(w),key=`${safeText(w.assetId||w.itemCode||a.code)}¦${loc}`,q=safeNumber(w.quantity);running[key]=(running[key]||0)+(w.type==="nhap"?q:-q);return[fmtDate(w.date),w.type==="nhap"?"Nhập":"Xuất",safeText(w.operationLabel||OPERATION_LABELS[w.operationType],"—"),safeText(w.voucherNo,"—"),loc,safeText(w.counterpartyLocation,"—"),safeText(w.supplier||w.repairVendor,"—"),safeText(w.receiver,"—"),safeText(w.itemCode||a.code),safeText(w.itemName||a.name),safeText(w.unit||a.unit,"Cái"),safeText(w.category||a.category),safeText(w.ownership||a.ownership),w.type==="nhap"?q:0,w.type==="xuat"?q:0,running[key],safeText(w.description||w.note,"—")]});
  const txDetailHeaders=["Ngày","Loại phiếu","Loại nghiệp vụ","Số phiếu","Kho/Công trình","Nguồn/Đích đối ứng","Nhà cung cấp / ĐV sửa","Người nhận/giao","Mã tài sản","Tên tài sản","ĐVT","Loại tài sản","Nguồn gốc","Nhập","Xuất","Tồn sau giao dịch","Diễn giải"];
  const lotMap={};
  reportTx.forEach(w=>{const key=`${safeText(w.assetId||w.itemCode)}¦${getLocation(w)}`,q=Math.max(0,safeNumber(w.quantity));if(!lotMap[key])lotMap[key]=[];if(w.type==="nhap")lotMap[key].push({...w,remain:q});else if(w.type==="xuat"){let left=q;for(const lot of lotMap[key]){if(left<=0)break;const take=Math.min(left,Math.max(0,safeNumber(lot.remain)));lot.remain-=take;left-=take;}}});
  const lotDetail=Object.values(lotMap).flatMap(lots=>lots.filter(l=>safeNumber(l.remain)>0).map(l=>{const a=assetMap[safeText(l.assetId)]||{};return{date:l.date,voucherNo:safeText(l.voucherNo,"—"),operationLabel:safeText(l.operationLabel||OPERATION_LABELS[l.operationType],"Nhập kho"),location:getLocation(l),counterparty:safeText(l.counterpartyLocation),supplier:safeText(l.supplier),receiver:safeText(l.receiver),description:safeText(l.description||l.note),code:safeText(l.itemCode||a.code),name:safeText(l.itemName||a.name),unit:safeText(l.unit||a.unit,"Cái"),category:safeText(l.category||a.category),ownership:safeText(l.ownership||a.ownership),qtyIn:safeNumber(l.quantity),qtyRemain:safeNumber(l.remain),unitCost:safeNumber(l.unitCost)}})).filter(r=>!filter.category||r.category===filter.category).filter(r=>!filter.ownership||r.ownership===filter.ownership).filter(r=>!filter.locationName||r.location===filter.locationName).filter(r=>!filter.supplier||r.supplier===filter.supplier);
  const lotHeaders=["Ngày nhập","Phiếu nhập","Loại nghiệp vụ","Kho/Công trình nhận","Nguồn đối ứng","Nhà cung cấp","Mã tài sản","Tên tài sản","ĐVT","Loại tài sản","Nguồn gốc","SL nhập phiếu","SL còn tồn từ phiếu","Đơn giá","Giá trị tồn","Người nhận/giao","Diễn giải"];
  const lotRows=lotDetail.map(r=>[fmtDate(r.date),r.voucherNo,r.operationLabel,r.location,r.counterparty||"—",r.supplier||"—",r.code,r.name,r.unit,r.category,r.ownership,r.qtyIn,r.qtyRemain,r.unitCost,r.qtyRemain*r.unitCost,r.receiver||"—",r.description||"—"]);
  // Đối chiếu FIFO theo TỪNG kho/công trình: phiếu xuất chỉ được lấy từ các lô nhập của chính địa điểm đó.
  const fifoState={}; const fifoLinked=[];
  reportTx.forEach(w=>{
    if(!matchesCommon(w))return; const a=assetMap[safeText(w.assetId)]||{},loc=getLocation(w),key=`${safeText(w.assetId||w.itemCode||a.code)}¦${loc}`,q=Math.max(0,safeNumber(w.quantity));
    if(!fifoState[key])fifoState[key]=[]; const lots=fifoState[key];
    if(w.type==="nhap"){lots.push({row:w,remain:q});return;}
    let left=q,alloc=[];
    for(const lot of lots){if(left<=0)break;const take=Math.min(left,Math.max(0,lot.remain));if(take>0){lot.remain-=take;left-=take;alloc.push({lot,qty:take});}}
    if(alloc.length){const inV=[...new Set(alloc.map(x=>safeText(x.lot.row.voucherNo)).filter(Boolean))];const inD=[...new Set(alloc.map(x=>fmtDate(x.lot.row.date)).filter(Boolean))];const inCosts=alloc.map(x=>safeNumber(x.lot.row.unitCost));const cost=alloc.reduce((n,x)=>n+x.qty*safeNumber(x.lot.row.unitCost),0);fifoLinked.push([loc,safeText(w.itemCode||a.code),safeText(w.itemName||a.name),safeText(w.unit||a.unit,"Cái"),safeText(w.category||a.category),inV.join(" - "),inD.join(" - "),alloc.map(x=>x.qty).join(" - "),inCosts.map(x=>Number(x).toLocaleString("vi-VN")).join(" - "),safeText(w.voucherNo),fmtDate(w.date),safeText(w.operationLabel||OPERATION_LABELS[w.operationType]),safeText(w.receiver),q,cost,left>0?`Thiếu nguồn FIFO ${left}`:"",safeText(w.description||w.note)]);}
  });
  // Lô nhập chưa phát sinh xuất: để trống toàn bộ thông tin phiếu xuất.
  Object.entries(fifoState).forEach(([key,lots])=>lots.forEach(({row,remain})=>{if(remain<=0)return;const a=assetMap[safeText(row.assetId)]||{},loc=getLocation(row);fifoLinked.push([loc,safeText(row.itemCode||a.code),safeText(row.itemName||a.name),safeText(row.unit||a.unit,"Cái"),safeText(row.category||a.category),safeText(row.voucherNo),fmtDate(row.date),remain,safeNumber(row.unitCost),"","","","",0,remain*safeNumber(row.unitCost),"Chưa xuất",safeText(row.description||row.note)]);}));
  const fifoLinkHeaders=["Kho/Công trình","Mã tài sản","Tên tài sản","ĐVT","Loại tài sản","Phiếu nhập FIFO","Ngày nhập","SL lấy từ lô nhập","Đơn giá nhập","Phiếu xuất","Ngày xuất","Mục đích xuất","Người nhận","SL xuất","Giá vốn FIFO / Giá trị còn","Trạng thái","Diễn giải"];
  const fifoLinkRows=fifoLinked.sort((a,b)=>`${a[0]}|${a[1]}|${a[6]}|${a[9]}`.localeCompare(`${b[0]}|${b[1]}|${b[6]}|${b[9]}`));
  const tabBtn=(id,label,icon)=><button type="button" onClick={()=>setTab(id)} className="flex items-center gap-2 px-4 py-2.5 text-[13px] font-medium whitespace-nowrap" style={{color:tab===id?TOKENS.brand:TOKENS.muted,borderBottom:`2px solid ${tab===id?TOKENS.brand:"transparent"}`}}>{icon}{label}</button>;
  const downloadTemplate=()=>downloadExcelTemplate("Mau_Import_Phieu_Nhap_Xuat_Kho",["Loại phiếu","Loại nghiệp vụ","Số phiếu","Ngày tháng","Tên tài sản","Mã hàng","Kho/Công trình","Loại địa điểm","Kho/Công trình đối ứng","Nhà cung cấp","Đơn vị sửa chữa","Người giao/nhận","Số lượng","Đơn vị tính","Đơn giá","Diễn giải","Ghi chú"],[["Nhập kho","Mua mới bên ngoài","PN-20260824-001","24/08/2026","Máy khoan mẫu","TS-001",safeProjects[0]?.name||"Kho trung tâm",safeProjects[0]?"Công trình":"Kho","",supplierNames[0]||"Nhà cung cấp A","","Nguyễn Văn A",1,"Cái",3500000,"Nhập mua mới",""]]);
  const filterProps={filter,setFilter,categories,groups:groupsList,ownerships,projects:safeProjects,locations,suppliers:supplierNames,operations:operationOptions};
  const visibleType=tab==="in"?"nhap":"xuat";
  const visibleRows=tab==="transfer"?filteredTx.filter(w=>w.type==="xuat"&&w.operationType==="luan_chuyen_di"):filteredTx.filter(w=>w.type===visibleType&&w.operationType!=="luan_chuyen_di"&&w.operationType!=="luan_chuyen_den");
  const allVisible=visibleRows.length>0&&visibleRows.every(w=>selectedRows.includes(w.id));
  return <div className="aa-fade">
    <div className="flex items-start justify-between mb-4 gap-4"><div><h1 className="aa-display text-xl font-semibold">Kho — Nhập / Xuất / Tồn</h1><div className="text-[12px] mt-1" style={{color:TOKENS.muted}}>Quản lý phiếu nhiều mã tài sản, loại nghiệp vụ, luân chuyển, sửa chữa và báo cáo chi tiết.</div></div><div className="flex gap-2 flex-wrap justify-end"><Btn icon={Download} onClick={downloadTemplate}>Tải mẫu Excel</Btn>{isAdmin&&<label className="inline-flex items-center gap-1.5 rounded-md font-medium px-3 py-1.5 text-[13px] cursor-pointer" style={{background:TOKENS.info,color:"white"}}><UploadCloud size={14}/>Đổ phiếu Excel<input type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const file=e.target.files?.[0];if(file)onImport?.(file);e.target.value=""}}/></label>}</div></div>
    <div className="rounded-lg overflow-hidden" style={{background:TOKENS.surface,border:`1px solid ${TOKENS.border}`}}><div className="flex border-b overflow-x-auto" style={{borderColor:TOKENS.border}}>{tabBtn("in","Nhập kho",<PackagePlus size={15}/>)}{tabBtn("out","Xuất kho",<PackageMinus size={15}/>)}{tabBtn("transfer","Chuyển kho",<ArrowLeftRight size={15}/>)}{tabBtn("report","Tổng hợp N-X-T",<ClipboardList size={15}/>)}{tabBtn("txdetail","Chi tiết N-X-T",<FileSpreadsheet size={15}/>)}{tabBtn("fifolink","Đối chiếu FIFO nhập → xuất",<Archive size={15}/>)}{tabBtn("lotdetail","Tồn theo phiếu nhập",<Archive size={15}/>)}</div><div className="p-4">
      {(tab==="in"||tab==="out"||tab==="transfer")&&<><div className="flex items-center justify-between gap-3 mb-3"><div className="flex gap-2"><Btn kind="primary" icon={Plus} onClick={()=>onAdd?.(tab==="transfer"?"transfer":visibleType)}>{tab==="in"?"Lập phiếu nhập":tab==="out"?"Lập phiếu xuất":"Lập phiếu chuyển kho"}</Btn>{visibleRows.length>0&&<label className="inline-flex items-center gap-1.5 text-[12px]"><input type="checkbox" checked={allVisible} onChange={()=>setSelectedRows(allVisible?selectedRows.filter(id=>!visibleRows.some(w=>w.id===id)):[...new Set([...selectedRows,...visibleRows.map(w=>w.id)])])}/> Chọn tất cả trang lọc</label>}</div><div className="flex gap-2"><input className={inputCls} style={{...inputStyle,maxWidth:360}} placeholder="Tìm phiếu, tài sản, NCC, diễn giải..." value={query} onChange={e=>setQuery(e.target.value)}/>{isAdmin&&selectedRows.length>0&&<Btn kind="danger" icon={Trash2} onClick={()=>{onDeleteRows?.(selectedRows);setSelectedRows([])}}>Xóa {selectedRows.length} dòng</Btn>}</div></div><WarehouseFilter {...filterProps}/><div className="rounded-lg overflow-auto" style={{border:`1px solid ${TOKENS.border}`}}><table className="w-full min-w-[2050px]"><thead><tr>{txHeaders.map((h,i)=><Th key={`${i}-${h}`}>{h}</Th>)}</tr></thead><tbody>{visibleRows.map((w,i)=>{const a=assetMap[w.assetId]||{};return <tr key={w.id||`row-${i}`} className="aa-row"><Td><input type="checkbox" checked={selectedRows.includes(w.id)} onChange={e=>setSelectedRows(e.target.checked?[...selectedRows,w.id]:selectedRows.filter(id=>id!==w.id))}/></Td><Td><Tag>{safeText(w.voucherNo,"—")}</Tag></Td><Td mono>{fmtDate(w.date)}</Td><Td>{safeText(w.operationLabel||OPERATION_LABELS[w.operationType],"—")}</Td><Td><Tag>{safeText(w.itemCode||a.code,"—")}</Tag></Td><Td>{safeText(w.itemName||a.name,"—")}</Td><Td>{getLocation(w)}</Td><Td>{safeText(w.counterpartyLocation,"—")}</Td><Td>{safeText(w.unit||a.unit,"Cái")}</Td><Td right mono>{safeNumber(w.quantity)}</Td><Td right mono>{fmtVND(safeNumber(w.unitCost))}</Td><Td right mono>{fmtVND(Number.isFinite(Number(w.total))?Number(w.total):safeNumber(w.quantity)*safeNumber(w.unitCost))}</Td><Td>{safeText(w.supplier||w.repairVendor,"—")}</Td><Td>{safeText(w.receiver,"—")}</Td><Td>{safeText(w.referenceNo,"—")}</Td><Td>{safeText(w.description||w.note,"—")}</Td></tr>})}</tbody></table>{!visibleRows.length&&<EmptyState text={tab==="in"?"Chưa có phiếu nhập phù hợp":tab==="out"?"Chưa có phiếu xuất phù hợp":"Chưa có phiếu chuyển kho phù hợp"}/>}</div></>}
      {(tab==="report"||tab==="txdetail"||tab==="fifolink"||tab==="lotdetail")&&<><div className="rounded-lg p-4 mb-4" style={{background:TOKENS.paper,border:`1px solid ${TOKENS.border}`}}><WarehouseFilter {...filterProps}/><div className="flex gap-2 items-end justify-between"><label className="text-[11px]" style={{color:TOKENS.muted}}>Đến ngày<input type="date" className={inputCls} style={{...inputStyle,width:165}} value={asOf} onChange={e=>setFilter({...filter,asOfDate:e.target.value})}/></label><div>{tab==="report"&&<ExportBar onExcel={()=>onExportExcel("bao-cao-tong-hop-nhap-xuat-ton",reportHeaders,reportRows)} onPdf={()=>onExportPdf(`Báo cáo tổng hợp nhập xuất tồn đến ${fmtDate(asOf)}`,reportHeaders,reportRows)}/>} {tab==="txdetail"&&<ExportBar onExcel={()=>onExportExcel("bao-cao-chi-tiet-nhap-xuat-ton-theo-phieu",txDetailHeaders,txDetailRows)} onPdf={()=>onExportPdf("Báo cáo chi tiết nhập xuất tồn theo phiếu",txDetailHeaders,txDetailRows)}/>} {tab==="fifolink"&&<ExportBar onExcel={()=>onExportExcel("bao-cao-doi-chieu-fifo-nhap-xuat-theo-cong-trinh",fifoLinkHeaders,fifoLinkRows)} onPdf={()=>onExportPdf("Báo cáo đối chiếu FIFO nhập - xuất theo công trình",fifoLinkHeaders,fifoLinkRows)}/>} {tab==="lotdetail"&&<ExportBar onExcel={()=>onExportExcel("bao-cao-ton-chi-tiet-theo-phieu-nhap",lotHeaders,lotRows)} onPdf={()=>onExportPdf("Báo cáo tồn chi tiết theo phiếu nhập",lotHeaders,lotRows)}/>}</div></div></div>{tab==="report"&&<ReportTable headers={reportHeaders} rows={reportRows} moneyCols={[10]} empty="Không có dữ liệu tổng hợp"/>}{tab==="txdetail"&&<ReportTable headers={txDetailHeaders} rows={txDetailRows} empty="Không có giao dịch theo điều kiện lọc"/>}{tab==="fifolink"&&<><div className="text-[12px] mb-2 rounded-md p-3" style={{background:TOKENS.brandSoft,color:TOKENS.ink}}>FIFO được tính <b>riêng cho từng Kho/Công trình</b>. Phiếu xuất của Công trình A không được lấy lô nhập của Công trình B. Nếu một phiếu xuất lấy từ nhiều phiếu nhập, các số phiếu nhập được nối bằng dấu <b> - </b>. Lô chưa xuất để trống thông tin phiếu xuất.</div><ReportTable headers={fifoLinkHeaders} rows={fifoLinkRows} moneyCols={[14]} empty="Không có dữ liệu đối chiếu FIFO"/></>}{tab==="lotdetail"&&<><div className="text-[12px] mb-2" style={{color:TOKENS.muted}}>Theo FIFO: thể hiện tài sản còn tồn được nhập <b>ngày nào, phiếu nhập nào</b> và diễn giải nghiệp vụ.</div><ReportTable headers={lotHeaders} rows={lotRows} moneyCols={[13,14]} empty="Không có tồn theo phiếu nhập"/></>}</>}
    </div></div>
  </div>;
}

function ReportTable({headers,rows,moneyCols=[],empty}){return <div className="rounded-lg overflow-auto" style={{border:`1px solid ${TOKENS.border}`}}><table className="w-full min-w-[1450px]"><thead><tr>{headers.map(h=><Th key={h}>{h}</Th>)}</tr></thead><tbody>{rows.map((r,i)=><tr key={i} className="aa-row">{r.map((c,j)=><Td key={j} mono={typeof c==="number"} right={typeof c==="number"}>{moneyCols.includes(j)?fmtVND(c):c}</Td>)}</tr>)}</tbody></table>{!rows.length&&<EmptyState text={empty||"Không có dữ liệu"}/>}</div>}
function WarehouseFilter({filter,setFilter,categories=[],groups=[],ownerships=[],projects=[],locations=[],suppliers=[],operations=[]}){const set=k=>e=>setFilter({...filter,[k]:e.target.value});const opts=a=>(Array.isArray(a)?a:[]).map(safeText).filter(Boolean);return <div className="flex flex-wrap gap-2 mb-3"><select className={inputCls} style={{...inputStyle,width:210}} value={safeText(filter.operationType)} onChange={set("operationType")}><option value="">Tất cả loại nghiệp vụ</option>{(Array.isArray(operations)?operations:[]).filter(Boolean).map((x,i)=><option key={safeText(x.id,`op-${i}`)} value={safeText(x.id)}>{safeText(x.label,x.id)}</option>)}</select><select className={inputCls} style={{...inputStyle,width:175}} value={safeText(filter.category)} onChange={set("category")}><option value="">Tất cả loại tài sản</option>{opts(categories).map((x,i)=><option key={`${i}-${x}`} value={x}>{x}</option>)}</select><select className={inputCls} style={{...inputStyle,width:175}} value={safeText(filter.group)} onChange={set("group")}><option value="">Tất cả nhóm</option>{opts(groups).map((x,i)=><option key={`${i}-${x}`} value={x}>{x}</option>)}</select><select className={inputCls} style={{...inputStyle,width:155}} value={safeText(filter.ownership)} onChange={set("ownership")}><option value="">Tất cả nguồn gốc</option>{opts(ownerships).map((x,i)=><option key={`${i}-${x}`} value={x}>{x}</option>)}</select><select className={inputCls} style={{...inputStyle,width:210}} value={safeText(filter.projectId)} onChange={set("projectId")}><option value="">Tất cả công trình</option>{(Array.isArray(projects)?projects:[]).filter(Boolean).map((p,i)=><option key={safeText(p.id,`p-${i}`)} value={safeText(p.id)}>{safeText(p.name,`Công trình ${i+1}`)}</option>)}</select><select className={inputCls} style={{...inputStyle,width:220}} value={safeText(filter.locationName)} onChange={set("locationName")}><option value="">Tất cả kho / công trình</option>{opts(locations).map((x,i)=><option key={`${i}-${x}`} value={x}>{x}</option>)}</select><select className={inputCls} style={{...inputStyle,width:220}} value={safeText(filter.supplier)} onChange={set("supplier")}><option value="">Tất cả nhà cung cấp / đơn vị sửa</option>{opts(suppliers).map((x,i)=><option key={`${i}-${x}`} value={x}>{x}</option>)}</select></div>;}
function SupplierSearchPicker({suppliers,value,onPick}){const selected=suppliers.find(s=>s.name===value||s.id===value);const[q,setQ]=useState(selected?.name||value||"");const[open,setOpen]=useState(false);const matches=suppliers.filter(s=>normalizeText(`${s.code||""} ${s.name||""} ${s.taxCode||""} ${s.phone||""}`).includes(normalizeText(q))).slice(0,12);return <div className="relative"><input className={inputCls} style={inputStyle} value={q} onFocus={()=>setOpen(true)} onChange={e=>{setQ(e.target.value);onPick(e.target.value);setOpen(true)}} placeholder="Gõ tên, mã NCC, MST..."/>{open&&q&&<div className="absolute z-50 left-0 right-0 top-full mt-1 max-h-56 overflow-auto rounded-md bg-white shadow-lg" style={{border:`1px solid ${TOKENS.border}`}}>{matches.map(s=><button type="button" key={s.id} className="block w-full text-left px-3 py-2 text-[12px] hover:bg-red-50" onClick={()=>{onPick(s.name);setQ(s.name);setOpen(false)}}><b>{s.code||"NCC"}</b> — {s.name}<div className="text-[10px]" style={{color:TOKENS.muted}}>{s.taxCode?`MST: ${s.taxCode}`:""}{s.phone?` · ${s.phone}`:""}</div></button>)}{!matches.length&&<div className="px-3 py-2 text-[12px]" style={{color:TOKENS.muted}}>Không tìm thấy trong danh mục — vẫn có thể dùng tên vừa nhập</div>}</div>}</div>}

function SupplierCatalog({suppliers,isAdmin,onAdd,onEdit,onDeleteMany,onImportExcel,onExportExcel,onExportPdf}){
  const[q,setQ]=useState(""),[selected,setSelected]=useState([]),[editing,setEditing]=useState(null),fileRef=useRef(null);
  const empty={code:"",name:"",taxCode:"",phone:"",address:"",contact:"",note:""}; const[f,setF]=useState(empty);
  const rows=suppliers.filter(x=>normalizeText(`${x.code} ${x.name} ${x.taxCode} ${x.phone} ${x.address}`).includes(normalizeText(q)));
  const headers=["Mã NCC","Tên nhà cung cấp","Mã số thuế","Điện thoại","Địa chỉ","Người liên hệ","Ghi chú"],dataRows=rows.map(x=>[x.code,x.name,x.taxCode,x.phone,x.address,x.contact,x.note]);
  const all=rows.length>0&&rows.every(x=>selected.includes(x.id));
  const save=()=>{if(editing){onEdit(editing,f);setEditing(null)}else onAdd(f);setF(empty)};
  const downloadTemplate=()=>downloadExcelTemplate("Mau_Import_Nha_Cung_Cap",headers,[["NCC-001","Công ty ABC","0312345678","0909123456","TP.HCM","Nguyễn Văn A",""]]);
  return <div className="aa-fade"><div className="flex items-start justify-between mb-4 gap-3"><div><h1 className="aa-display text-xl font-semibold">Danh mục nhà cung cấp</h1><div className="text-[12px] mt-1" style={{color:TOKENS.muted}}>Tìm nhanh trong phiếu nhập/xuất, đổ lô Excel, sửa và xóa hàng loạt.</div></div><div className="flex gap-2 flex-wrap justify-end"><ExportBar onExcel={()=>onExportExcel("danh-muc-nha-cung-cap",headers,dataRows)} onPdf={()=>onExportPdf("Danh mục nhà cung cấp",headers,dataRows)}/><Btn icon={Download} onClick={downloadTemplate}>Tải mẫu</Btn>{isAdmin&&<><input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const file=e.target.files?.[0];if(file)onImportExcel?.(file);e.target.value=""}}/><Btn kind="info" icon={UploadCloud} onClick={()=>fileRef.current?.click()}>Đổ Excel</Btn></>}</div></div>{isAdmin&&<div className="rounded-lg p-4 mb-4 grid grid-cols-4 gap-2" style={{background:TOKENS.surface,border:`1px solid ${TOKENS.border}`}}><input className={inputCls} style={inputStyle} placeholder="Mã NCC" value={f.code} onChange={e=>setF({...f,code:e.target.value})}/><input className={inputCls} style={inputStyle} placeholder="Tên nhà cung cấp *" value={f.name} onChange={e=>setF({...f,name:e.target.value})}/><input className={inputCls} style={inputStyle} placeholder="Mã số thuế" value={f.taxCode} onChange={e=>setF({...f,taxCode:e.target.value})}/><input className={inputCls} style={inputStyle} placeholder="Điện thoại" value={f.phone} onChange={e=>setF({...f,phone:e.target.value})}/><input className={inputCls} style={inputStyle} placeholder="Địa chỉ" value={f.address} onChange={e=>setF({...f,address:e.target.value})}/><input className={inputCls} style={inputStyle} placeholder="Người liên hệ" value={f.contact} onChange={e=>setF({...f,contact:e.target.value})}/><input className={inputCls} style={inputStyle} placeholder="Ghi chú" value={f.note} onChange={e=>setF({...f,note:e.target.value})}/><div className="flex gap-2"><Btn kind="primary" icon={editing?Save:Plus} onClick={save}>{editing?"Lưu sửa":"Thêm NCC"}</Btn>{editing&&<Btn onClick={()=>{setEditing(null);setF(empty)}}>Hủy</Btn>}</div></div>}<div className="flex justify-between gap-2 mb-3"><input className={inputCls} style={{...inputStyle,maxWidth:360}} placeholder="Tìm theo tên, mã NCC, MST, điện thoại..." value={q} onChange={e=>setQ(e.target.value)}/>{isAdmin&&selected.length>0&&<Btn kind="danger" icon={Trash2} onClick={()=>{onDeleteMany(selected);setSelected([])}}>Xóa {selected.length} NCC</Btn>}</div><div className="rounded-lg overflow-auto" style={{background:TOKENS.surface,border:`1px solid ${TOKENS.border}`}}><table className="w-full min-w-[1200px]"><thead><tr><Th><input type="checkbox" checked={all} onChange={()=>setSelected(all?selected.filter(id=>!rows.some(x=>x.id===id)):[...new Set([...selected,...rows.map(x=>x.id)])])}/></Th>{headers.map(h=><Th key={h}>{h}</Th>)}{isAdmin&&<Th>Thao tác</Th>}</tr></thead><tbody>{rows.map(x=><tr key={x.id} className="aa-row"><Td><input type="checkbox" checked={selected.includes(x.id)} onChange={e=>setSelected(e.target.checked?[...selected,x.id]:selected.filter(id=>id!==x.id))}/></Td><Td mono><Tag>{x.code||"—"}</Tag></Td><Td>{x.name}</Td><Td>{x.taxCode||"—"}</Td><Td>{x.phone||"—"}</Td><Td>{x.address||"—"}</Td><Td>{x.contact||"—"}</Td><Td>{x.note||"—"}</Td>{isAdmin&&<Td><Btn small icon={Pencil} onClick={()=>{setEditing(x.id);setF({...x})}}>Sửa</Btn></Td>}</tr>)}</tbody></table>{!rows.length&&<EmptyState text="Không có nhà cung cấp phù hợp"/>}</div></div>
}

function CostHistoryView({ costHistory, assetsById, onAdd, onExportExcel, onExportPdf }) { const headers=["Tài sản","Loại chi phí","Ngày","Số tiền","Nội dung","Nhà cung cấp"]; const rows=costHistory.map(c=>[assetsById[c.assetId]?.code,c.type,fmtDate(c.date),c.amount,c.description,c.vendor]); return <div className="aa-fade"><div className="flex items-center justify-between mb-4"><h1 className="aa-display text-xl font-semibold">Chi phí thiết bị</h1><div className="flex gap-2"><ExportBar onExcel={()=>onExportExcel("lich-su-chi-phi-thiet-bi",headers,rows)} onPdf={()=>onExportPdf("Lịch sử chi phí thiết bị",headers,rows)}/><Btn kind="primary" icon={Plus} onClick={onAdd}>Thêm chi phí</Btn></div></div><div className="rounded-lg overflow-hidden" style={{background:TOKENS.surface,border:`1px solid ${TOKENS.border}`}}><table className="w-full"><thead><tr>{headers.map(h=><Th key={h}>{h}</Th>)}</tr></thead><tbody>{costHistory.map(c=><tr key={c.id} className="aa-row"><Td mono><Tag>{assetsById[c.assetId]?.code}</Tag></Td><Td>{c.type}</Td><Td>{fmtDate(c.date)}</Td><Td right mono>{fmtVND(c.amount)}</Td><Td>{c.description}</Td><Td>{c.vendor}</Td></tr>)}</tbody></table>{!costHistory.length&&<EmptyState text="Chưa có chi phí"/>}</div></div>; }
function AssetSearchPicker({assets,value,onPick}) {
  const selected=assets.find(a=>a.id===value); const [q,setQ]=useState(selected?`${selected.code} — ${selected.name}`:""); const [open,setOpen]=useState(false);
  const matches=assets.filter(a=>normalizeText(`${a.code} ${a.name} ${a.serial||""}`).includes(normalizeText(q))).slice(0,10);
  return <div className="relative"><input className={inputCls} style={inputStyle} value={q} onFocus={()=>setOpen(true)} onChange={e=>{setQ(e.target.value);setOpen(true)}} placeholder="Gõ tên, mã hoặc serial..."/>{open&&q&&<div className="absolute z-50 left-0 right-0 top-full mt-1 max-h-56 overflow-auto rounded-md bg-white shadow-lg" style={{border:`1px solid ${TOKENS.border}`}}>{matches.map(a=><button type="button" key={a.id} className="block w-full text-left px-3 py-2 text-[12px] hover:bg-red-50" onClick={()=>{onPick(a);setQ(`${a.code} — ${a.name}`);setOpen(false)}}><b>{a.code}</b> — {a.name}<div className="text-[10px]" style={{color:TOKENS.muted}}>{a.unit||"Cái"} · {a.category||"Khác"} · {a.ownership||"Công ty"}</div></button>)}{!matches.length&&<div className="px-3 py-2 text-[12px]" style={{color:TOKENS.muted}}>Không tìm thấy tài sản</div>}</div>}</div>;
}
function WarehouseTxModal({ assets, projects, suppliers = [], onClose, onSubmit, title, fixedType, fixedOperation = "" }) {
  const blank=()=>({id:uid("line"),assetId:"",quantity:1,unitCost:0});
  const initialOperation=fixedOperation || (fixedType==="xuat"?"su_dung_cong_trinh":"mua_moi");
  const [f,setF]=useState({type:fixedType||"nhap",operationType:initialOperation,voucherNo:"",date:nowIso().slice(0,10),receiver:"",supplier:"",description:"",locationType:"project",warehouseName:"Kho trung tâm",projectId:"",counterpartyLocationType:"project",counterpartyProjectId:"",counterpartyWarehouseName:"",repairVendor:"",note:"",address:"",referenceNo:"",attachedDoc:"",transportPerson:"",vehicle:"",orderNo:"",items:Array.from({length:8},blank)});
  const set=k=>e=>setF({...f,[k]:e.target.value});
  const update=(id,patch)=>setF({...f,items:f.items.map(x=>x.id===id?{...x,...patch}:x)});
  const remove=id=>setF({...f,items:f.items.filter(x=>x.id!==id)});
  const ops=WAREHOUSE_OPERATIONS[f.type]||[];
  const isTransfer=f.operationType==="luan_chuyen_di";
  const isRepairOut=f.operationType==="sua_chua";
  const isRepairReturn=f.operationType==="thu_hoi_sua_chua";
  const isOther=f.operationType==="nhap_khac"||f.operationType==="xuat_khac";
  const isPurchase=f.operationType==="mua_moi";
  const needsCounterparty=isTransfer||f.operationType==="thu_hoi_cong_trinh"||f.operationType==="luan_chuyen_den";
  const sourceName=f.locationType==="project"?(projects.find(p=>p.id===f.projectId)?.name||""):(f.warehouseName||"");
  const destName=f.counterpartyLocationType==="project"?(projects.find(p=>p.id===f.counterpartyProjectId)?.name||""):(f.counterpartyWarehouseName||"");
  const enteredTotal=f.items.reduce((n,x)=>n+(Number(x.quantity)||0)*(Number(x.unitCost)||0),0);
  const Box=({legend,children})=><fieldset className="rounded-md p-3" style={{border:`1px solid ${TOKENS.border}`,background:TOKENS.surface}}><legend className="px-2 text-[12px] font-semibold" style={{color:TOKENS.brand}}>{legend}</legend>{children}</fieldset>;
  return <Modal title={title||"Phiếu kho"} onClose={onClose} wide>
    <div className="grid grid-cols-12 gap-3 mb-3">
      <div className="col-span-9">
        <Box legend="Thông tin chung">
          {fixedOperation?<div className="flex items-center gap-2 mb-3"><span className="px-2.5 py-1 rounded text-[12px] font-semibold" style={{background:TOKENS.brandSoft,color:TOKENS.brand}}>CHUYỂN KHO NỘI BỘ</span><span className="text-[11px]" style={{color:TOKENS.muted}}>Tự động xuất nơi đi, nhập nơi đến và giữ nguyên giá vốn FIFO.</span></div>:<Field label={f.type==="nhap"?"Loại phiếu nhập":"Loại phiếu xuất"}><select className={inputCls} style={inputStyle} value={f.operationType} onChange={e=>setF({...f,operationType:e.target.value})}>{ops.map(o=><option key={o.id} value={o.id}>{o.label}</option>)}</select></Field>}
          {isTransfer?<>
            <div className="grid grid-cols-2 gap-3 mb-2">
              <div className="rounded-md p-3" style={{background:TOKENS.paper,border:`1px solid ${TOKENS.border}`}}><div className="text-[12px] font-semibold mb-2">Xuất tại kho / công trình</div><div className="grid grid-cols-2 gap-2"><select className={inputCls} style={inputStyle} value={f.locationType} onChange={set("locationType")}><option value="project">Công trình</option><option value="warehouse">Kho</option></select>{f.locationType==="project"?<select className={inputCls} style={inputStyle} value={f.projectId} onChange={set("projectId")}><option value="">Chọn công trình đi</option>{projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select>:<input className={inputCls} style={inputStyle} value={f.warehouseName} onChange={set("warehouseName")} placeholder="Tên kho đi"/>}</div></div>
              <div className="rounded-md p-3" style={{background:TOKENS.paper,border:`1px solid ${TOKENS.border}`}}><div className="text-[12px] font-semibold mb-2">Nhập tại kho / công trình</div><div className="grid grid-cols-2 gap-2"><select className={inputCls} style={inputStyle} value={f.counterpartyLocationType} onChange={set("counterpartyLocationType")}><option value="project">Công trình</option><option value="warehouse">Kho</option></select>{f.counterpartyLocationType==="project"?<select className={inputCls} style={inputStyle} value={f.counterpartyProjectId} onChange={set("counterpartyProjectId")}><option value="">Chọn công trình đến</option>{projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select>:<input className={inputCls} style={inputStyle} value={f.counterpartyWarehouseName} onChange={set("counterpartyWarehouseName")} placeholder="Tên kho đến"/>}</div></div>
            </div>
            <div className="grid grid-cols-2 gap-x-4"><Field label="Lệnh điều động số"><input className={inputCls} style={inputStyle} value={f.orderNo} onChange={set("orderNo")}/></Field><Field label="Người vận chuyển"><input className={inputCls} style={inputStyle} value={f.transportPerson} onChange={set("transportPerson")}/></Field><Field label="Về việc"><input className={inputCls} style={inputStyle} value={f.description} onChange={set("description")} placeholder="VD: Chuyển vật tư từ CT A đến CT B"/></Field><Field label="Phương tiện"><input className={inputCls} style={inputStyle} value={f.vehicle} onChange={set("vehicle")}/></Field><Field label="Tham chiếu"><input className={inputCls} style={inputStyle} value={f.referenceNo} onChange={set("referenceNo")}/></Field><Field label="Người giao / nhận"><input className={inputCls} style={inputStyle} value={f.receiver} onChange={set("receiver")}/></Field></div>
          </>:<>
            <div className="grid grid-cols-2 gap-x-4">
              <Field label={f.type==="nhap"?"Kho/Công trình nhận":"Kho/Công trình xuất"}><select className={inputCls} style={inputStyle} value={f.locationType} onChange={set("locationType")}><option value="project">Công trình</option><option value="warehouse">Kho</option></select></Field>
              <Field label={f.locationType==="project"?"Tên công trình":"Tên kho"}>{f.locationType==="project"?<select className={inputCls} style={inputStyle} value={f.projectId} onChange={set("projectId")}><option value="">Chọn công trình</option>{projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select>:<input className={inputCls} style={inputStyle} value={f.warehouseName} onChange={set("warehouseName")} />}</Field>
              {needsCounterparty&&<><Field label={f.type==="nhap"?"Nguồn đối ứng":"Đích đối ứng"}><select className={inputCls} style={inputStyle} value={f.counterpartyLocationType} onChange={set("counterpartyLocationType")}><option value="project">Công trình</option><option value="warehouse">Kho</option></select></Field><Field label="Kho/Công trình đối ứng">{f.counterpartyLocationType==="project"?<select className={inputCls} style={inputStyle} value={f.counterpartyProjectId} onChange={set("counterpartyProjectId")}><option value="">Chọn công trình</option>{projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select>:<input className={inputCls} style={inputStyle} value={f.counterpartyWarehouseName} onChange={set("counterpartyWarehouseName")}/>}</Field></>}
              <Field label={isPurchase?"Đối tượng / Nhà cung cấp *":"Đối tượng / Nhà cung cấp"}><SupplierSearchPicker suppliers={suppliers} value={f.supplier} onPick={v=>setF({...f,supplier:v})}/></Field>
              <Field label={f.type==="nhap"?"Người giao hàng":"Người nhận"}><input className={inputCls} style={inputStyle} value={f.receiver} onChange={set("receiver")}/></Field>
              {isRepairOut&&<Field label="Đơn vị sửa chữa *"><SupplierSearchPicker suppliers={suppliers} value={f.repairVendor} onPick={v=>setF({...f,repairVendor:v})}/></Field>}
              {f.type==="xuat"&&<Field label="Địa chỉ"><input className={inputCls} style={inputStyle} value={f.address} onChange={set("address")}/></Field>}
              <Field label={isOther?"Diễn giải / Lý do *":"Diễn giải / Lý do"}><input className={inputCls} style={inputStyle} value={f.description} onChange={set("description")} /></Field>
              <Field label="Kèm theo"><input className={inputCls} style={inputStyle} value={f.attachedDoc} onChange={set("attachedDoc")} placeholder="Chứng từ gốc / biên bản..."/></Field>
              <Field label="Tham chiếu"><input className={inputCls} style={inputStyle} value={f.referenceNo} onChange={set("referenceNo")}/></Field>
            </div>
          </>}
        </Box>
      </div>
      <div className="col-span-3"><Box legend="Chứng từ"><Field label="Ngày hạch toán"><input type="date" className={inputCls} style={inputStyle} value={f.date} onChange={set("date")}/></Field><Field label="Ngày chứng từ"><input type="date" className={inputCls} style={inputStyle} value={f.date} onChange={set("date")}/></Field><Field label="Số chứng từ"><input className={inputCls} style={inputStyle} value={f.voucherNo} onChange={set("voucherNo")} placeholder="Để trống để tự sinh"/></Field><div className="text-[11px] leading-relaxed" style={{color:TOKENS.muted}}>Trong cùng ngày, hệ thống mặc định <b>Nhập trước - Xuất sau</b> khi tính FIFO.</div></Box></div>
    </div>
    <div className="flex items-center justify-between px-3 py-2 rounded-t-md" style={{background:TOKENS.brandSoft,border:`1px solid ${TOKENS.border}`}}><div className="flex items-center gap-4"><b className="text-[13px]">1. Hàng tiền</b><span className="text-[12px]" style={{color:TOKENS.muted}}>2. Thống kê</span><span className="text-[12px]" style={{color:TOKENS.muted}}>3. Khác</span></div><div className="flex gap-2"><Btn small onClick={()=>setF({...f,items:[...f.items,...Array.from({length:5},blank)]})}>+5 dòng</Btn><Btn small icon={Plus} onClick={()=>setF({...f,items:[...f.items,blank()]})}>Thêm dòng</Btn></div></div>
    <div className="overflow-auto" style={{border:`1px solid ${TOKENS.border}`,borderTop:0}}><table className="w-full min-w-[1250px]"><thead><tr>{(isTransfer?["STT","Mã hàng","Tên hàng","Xuất tại kho","Nhập tại kho","ĐVT","Số lượng","Giá vốn FIFO","Thành tiền","Tìm/chọn"]:["STT","Mã hàng","Tên hàng","Kho/Công trình","ĐVT","Số lượng",f.type==="nhap"?"Đơn giá":"Giá vốn FIFO","Thành tiền","Tìm/chọn"]).map(h=><Th key={h}>{h}</Th>)}</tr></thead><tbody>{f.items.map((x,i)=>{const a=assets.find(z=>z.id===x.assetId)||{};const amount=(Number(x.quantity)||0)*(Number(x.unitCost)||0);return <tr key={x.id} className="aa-row"><Td>{i+1}</Td><Td mono>{a.code||"—"}</Td><Td>{a.name||"—"}</Td><Td>{sourceName||"—"}</Td>{isTransfer&&<Td>{destName||"—"}</Td>}<Td>{a.unit||"Cái"}</Td><Td><input type="number" min="0.01" step="0.01" className={inputCls} style={{...inputStyle,width:90}} value={x.quantity} onChange={e=>update(x.id,{quantity:e.target.value})}/></Td><Td>{f.type==="nhap"&&!isTransfer?<input type="number" min="0" className={inputCls} style={{...inputStyle,width:120}} value={x.unitCost} onChange={e=>update(x.id,{unitCost:e.target.value})}/>:<span className="text-[11px]" style={{color:TOKENS.muted}}>Tự tính FIFO khi lưu</span>}</Td><Td right mono>{f.type==="nhap"&&!isTransfer?fmtVND(amount):"—"}</Td><Td><div className="flex items-center gap-2 min-w-[290px]"><AssetSearchPicker assets={assets} value={x.assetId} onPick={picked=>update(x.id,{assetId:picked.id,unitCost:f.type==="nhap"&&!isTransfer?(picked.cost||0):0})}/><button type="button" onClick={()=>remove(x.id)} title="Xóa dòng"><X size={14}/></button></div></Td></tr>})}</tbody></table></div>
    <div className="flex justify-between items-end mt-3"><Field label="Ghi chú"><textarea className={inputCls} style={{...inputStyle,width:520}} value={f.note} onChange={set("note")}/></Field><div className="text-right"><div className="text-[11px]" style={{color:TOKENS.muted}}>{f.type==="nhap"&&!isTransfer?"Tổng giá trị phiếu":"Giá vốn sẽ được xác định theo FIFO của đúng Kho/Công trình khi lưu"}</div>{f.type==="nhap"&&!isTransfer&&<div className="aa-display text-xl font-bold" style={{color:TOKENS.brand}}>{fmtVND(enteredTotal)}</div>}<div className="flex gap-2 mt-3"><Btn onClick={onClose}>Đóng</Btn><Btn kind="primary" onClick={()=>onSubmit({...f,operationType:fixedOperation||f.operationType})}>Lưu chứng từ</Btn></div></div></div>
  </Modal>;
}

function CostHistoryModal({ assets, onClose, onSubmit }) { const [f,setF]=useState({assetId:assets[0]?.id||"",type:"Cấp dầu",date:nowIso().slice(0,10),amount:0,description:"",vendor:""}); const set=k=>e=>setF({...f,[k]:e.target.value}); return <Modal title="Thêm chi phí thiết bị" onClose={onClose}><Field label="Tài sản"><select className={inputCls} style={inputStyle} value={f.assetId} onChange={set("assetId")}>{assets.map(a=><option key={a.id} value={a.id}>{a.code} — {a.name}</option>)}</select></Field><Field label="Loại chi phí"><select className={inputCls} style={inputStyle} value={f.type} onChange={set("type")}><option>Cấp dầu</option><option>Thay cáp</option><option>Sửa chữa</option><option>Chi phí khác</option></select></Field><Field label="Ngày"><input type="date" className={inputCls} style={inputStyle} value={f.date} onChange={set("date")}/></Field><Field label="Số tiền"><input type="number" className={inputCls} style={inputStyle} value={f.amount} onChange={set("amount")}/></Field><Field label="Nội dung"><textarea className={inputCls} style={inputStyle} value={f.description} onChange={set("description")}/></Field><Field label="Nhà cung cấp"><input className={inputCls} style={inputStyle} value={f.vendor} onChange={set("vendor")}/></Field><div className="flex justify-end gap-2"><Btn onClick={onClose}>Huỷ</Btn><Btn kind="primary" onClick={()=>onSubmit(f)}>Lưu chi phí</Btn></div></Modal>; }

/* ============================== SETTINGS ============================== */


function MasterDataPage({ title, description, items = [], usage, onAdd, onRename, onRemove, importKind, onImportExcel, templateHeader, placeholder, isAdmin }) {
  const [q,setQ]=useState("");
  const [val,setVal]=useState("");
  const [editing,setEditing]=useState(null);
  const [editVal,setEditVal]=useState("");
  const [selected,setSelected]=useState([]);
  const fileRef=useRef(null);
  const safeItems=Array.isArray(items)?items.filter(Boolean).map(String):[];
  const rows=safeItems.filter(x=>normalizeText(x).includes(normalizeText(q)));
  const all=rows.length>0&&rows.every(x=>selected.includes(x));
  const add=()=>{const v=val.trim();if(!v)return;onAdd?.(v);setVal("");};
  const removeOne=(name)=>{
    const n=usage?.(name)||0;
    if(n>0&&!window.confirm(`"${name}" đang được ${n} tài sản sử dụng. Xóa khỏi danh sách lựa chọn? Dữ liệu tài sản hiện có vẫn được giữ.`))return;
    onRemove?.(name);setSelected(x=>x.filter(v=>v!==name));
  };
  const removeMany=()=>{
    if(!selected.length)return;
    if(!window.confirm(`Xóa ${selected.length} mục đã chọn khỏi danh sách lựa chọn?`))return;
    selected.forEach(name=>onRemove?.(name));
    setSelected([]);
  };
  const headers=["Tên danh mục","Số tài sản đang dùng"];
  const exportRows=rows.map(x=>[x,usage?.(x)||0]);
  const downloadTemplate=()=>downloadExcelTemplate(`Mau_${title.replaceAll(" ","_")}`,[templateHeader||"Tên"],[[placeholder?.replace(/^VD:\\s*/,"")||"Mẫu"]]);
  return <div className="aa-fade">
    <div className="flex items-start justify-between gap-3 mb-4">
      <div><h1 className="aa-display text-xl font-semibold">{title}</h1><div className="text-[12px] mt-1" style={{color:TOKENS.muted}}>{description}</div></div>
      <div className="flex gap-2 flex-wrap justify-end">
        <Btn icon={Download} onClick={downloadTemplate}>Tải mẫu Excel</Btn>
        {isAdmin&&<><input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)onImportExcel?.(f,importKind);e.target.value=""}}/><Btn kind="info" icon={UploadCloud} onClick={()=>fileRef.current?.click()}>Đổ Excel</Btn></>}
      </div>
    </div>
    {isAdmin&&<div className="rounded-lg p-4 mb-4 flex gap-2" style={{background:TOKENS.surface,border:`1px solid ${TOKENS.border}`}}>
      <input className={inputCls} style={inputStyle} value={val} onChange={e=>setVal(e.target.value)} placeholder={placeholder} onKeyDown={e=>e.key==="Enter"&&add()}/>
      <Btn kind="primary" icon={Plus} onClick={add}>Thêm mới</Btn>
    </div>}
    <div className="flex items-center justify-between gap-3 mb-3">
      <input className={inputCls} style={{...inputStyle,maxWidth:420}} value={q} onChange={e=>setQ(e.target.value)} placeholder={`Tìm ${title.toLowerCase()}...`}/>
      {isAdmin&&selected.length>0&&<Btn kind="danger" icon={Trash2} onClick={removeMany}>Xóa {selected.length} mục</Btn>}
    </div>
    <div className="rounded-lg overflow-hidden" style={{background:TOKENS.surface,border:`1px solid ${TOKENS.border}`}}>
      <table className="w-full">
        <thead><tr><Th><input type="checkbox" checked={all} onChange={()=>setSelected(all?selected.filter(x=>!rows.includes(x)):[...new Set([...selected,...rows])])}/></Th><Th>Tên</Th><Th right>Số tài sản đang dùng</Th>{isAdmin&&<Th>Thao tác</Th>}</tr></thead>
        <tbody>{rows.map(name=><tr key={name} className="aa-row">
          <Td><input type="checkbox" checked={selected.includes(name)} onChange={e=>setSelected(e.target.checked?[...selected,name]:selected.filter(x=>x!==name))}/></Td>
          <Td>{editing===name?<input autoFocus className={inputCls} style={inputStyle} value={editVal} onChange={e=>setEditVal(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&editVal.trim()){onRename?.(name,editVal.trim());setEditing(null)}}}/>:<span className="font-medium">{name}</span>}</Td>
          <Td right mono>{usage?.(name)||0}</Td>
          {isAdmin&&<Td><div className="flex gap-1">{editing===name?<><Btn small kind="primary" icon={Save} onClick={()=>{if(editVal.trim()){onRename?.(name,editVal.trim());setEditing(null)}}}>Lưu</Btn><Btn small onClick={()=>setEditing(null)}>Hủy</Btn></>:<Btn small icon={Pencil} onClick={()=>{setEditing(name);setEditVal(name)}}>Sửa</Btn>}<Btn small kind="danger" icon={Trash2} onClick={()=>removeOne(name)}>Xóa</Btn></div></Td>}
        </tr>)}</tbody>
      </table>
      {!rows.length&&<EmptyState text="Chưa có dữ liệu phù hợp"/>}
    </div>
  </div>;
}

function EditableMasterList({ title, items, onAdd, onRemove, onRename, usage, placeholder, importKind, onImportExcel }) {
  const [val,setVal]=useState(""),[editing,setEditing]=useState(null),[editVal,setEditVal]=useState(""),fileRef=useRef(null);
  const submit=()=>{if(val.trim()){onAdd(val.trim());setVal("")}};
  const downloadTemplate=()=>downloadExcelTemplate(`Mau_${title.replaceAll(" ","_")}`,[title.includes("Bộ phận")?"Bộ phận":"Danh mục"],[[title.includes("Bộ phận")?"Vận hành":"Máy móc thiết bị"]]);
  return <div><div className="flex items-center justify-between gap-2 mb-3"><div><div className="text-[13px] font-medium">{title}</div><div className="text-[11px] mt-0.5" style={{color:TOKENS.muted}}>Thêm, sửa, xóa và đổ lô Excel; khi đổi tên sẽ cập nhật dữ liệu tài sản đang sử dụng.</div></div><div className="flex gap-2"><Btn small icon={Download} onClick={downloadTemplate}>Mẫu Excel</Btn><input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0];if(f)onImportExcel?.(f,importKind);e.target.value=""}}/><Btn small kind="info" icon={UploadCloud} onClick={()=>fileRef.current?.click()}>Đổ Excel</Btn></div></div><div className="rounded-lg overflow-hidden mb-3" style={{border:`1px solid ${TOKENS.border}`}}><table className="w-full"><thead><tr><Th>Tên</Th><Th right>Số tài sản đang dùng</Th><Th>Thao tác</Th></tr></thead><tbody>{items.map(it=>{const n=usage?usage(it):0;return <tr key={it} className="aa-row"><Td>{editing===it?<input autoFocus className={inputCls} style={inputStyle} value={editVal} onChange={e=>setEditVal(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"&&editVal.trim()){onRename(it,editVal.trim());setEditing(null)}}}/>:it}</Td><Td right mono>{n}</Td><Td><div className="flex gap-1">{editing===it?<><Btn small kind="primary" icon={Save} onClick={()=>{if(editVal.trim()){onRename(it,editVal.trim());setEditing(null)}}}>Lưu</Btn><Btn small onClick={()=>setEditing(null)}>Hủy</Btn></>:<Btn small icon={Pencil} onClick={()=>{setEditing(it);setEditVal(it)}}>Sửa</Btn>}<Btn small kind="danger" icon={Trash2} onClick={()=>{if(n>0&&!window.confirm(`"${it}" đang được ${n} tài sản sử dụng. Vẫn xoá khỏi danh sách lựa chọn?`))return;onRemove(it)}}>Xóa</Btn></div></Td></tr>})}</tbody></table>{!items.length&&<EmptyState text="Chưa có dữ liệu"/>}</div><div className="flex gap-2"><input className={inputCls} style={inputStyle} value={val} onChange={e=>setVal(e.target.value)} placeholder={placeholder} onKeyDown={e=>e.key==="Enter"&&submit()}/><Btn kind="primary" icon={Plus} onClick={submit}>Thêm</Btn></div></div>
}

function SettingsView({
  settings, onSetCompanyName, onAddCategory, onRemoveCategory, onRenameCategory, onAddDepartment, onRemoveDepartment, onRenameDepartment,
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
        <div className="text-[13px] font-medium mb-3">Tên ứng dụng</div>
        <div className="flex gap-2 max-w-md">
          <input className={inputCls} style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} />
          <Btn kind="primary" icon={Check} disabled={!name.trim() || name === settings.companyName} onClick={() => onSetCompanyName(name.trim())}>Lưu</Btn>
        </div>
        <div className="text-[11.5px] mt-2" style={{ color: TOKENS.muted }}>Mặc định: MYHL - Quản lý tài sản. Chỉ quản trị viên được đổi tên; tên hiển thị trên app và báo cáo PDF.</div>
      </div>

      <div className="rounded-lg p-5 mb-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <div className="text-[13px] font-medium mb-1">Nhập liệu hàng loạt từ Excel</div>
        <div className="text-[11.5px] mb-3" style={{ color: TOKENS.muted }}>Dùng để đổ lô dữ liệu nhanh vào danh mục tài sản hoặc công trình. Dòng đầu tiên phải là tiêu đề cột.</div>
        <div className="flex gap-2">
          <label className="inline-flex"><input type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0]; if(f) onImportExcel(f,"assets"); e.target.value="";}}/><span className="inline-flex items-center gap-1.5 rounded-md font-medium px-3 py-1.5 text-[13px] cursor-pointer" style={{background:TOKENS.brand,color:"#fff"}}><UploadCloud size={14}/>Import tài sản</span></label>
          <label className="inline-flex"><input type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0]; if(f) onImportExcel(f,"projects"); e.target.value="";}}/><span className="inline-flex items-center gap-1.5 rounded-md font-medium px-3 py-1.5 text-[13px] cursor-pointer" style={{background:TOKENS.info,color:"#fff"}}><UploadCloud size={14}/>Import công trình</span></label>
          <label className="inline-flex"><input type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{const f=e.target.files?.[0]; if(f) onImportExcel(f,"suppliers"); e.target.value="";}}/><span className="inline-flex items-center gap-1.5 rounded-md font-medium px-3 py-1.5 text-[13px] cursor-pointer" style={{background:TOKENS.gold,color:"#fff"}}><UploadCloud size={14}/>Import nhà cung cấp</span></label>
        </div>
      </div>

      <div className="rounded-lg p-5 mb-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <div className="text-[13px] font-medium mb-1">Danh mục tài sản</div>
        <div className="text-[11.5px] mb-3" style={{ color: TOKENS.muted }}>Các loại tài sản có thể chọn khi thêm mới (Laptop, Màn hình…).</div>
        <EditableMasterList title="Danh mục tài sản" items={settings.categories} onAdd={onAddCategory} onRemove={onRemoveCategory} onRename={onRenameCategory} usage={categoryUsage} placeholder="VD: Máy móc thiết bị" importKind="categories" onImportExcel={onImportExcel} />
      </div>

      <div className="rounded-lg p-5 mb-4" style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}>
        <div className="text-[13px] font-medium mb-1">Bộ phận / đơn vị sử dụng</div>
        <div className="text-[11.5px] mb-3" style={{ color: TOKENS.muted }}>Danh sách bộ phận dùng khi cấp phát hoặc chuyển tài sản.</div>
        <EditableMasterList title="Bộ phận / đơn vị sử dụng" items={settings.departments} onAdd={onAddDepartment} onRemove={onRemoveDepartment} onRename={onRenameDepartment} usage={deptUsage} placeholder="VD: Vận hành" importKind="departments" onImportExcel={onImportExcel} />
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

function AssetFormModal({ title, initial, categories = [], departments = [], projects = [], suppliers = [], customColumns = [], onClose, onSubmit }) {
  const [f, setF] = useState(() => initial ? {
    code: initial.code, name: initial.name, category: initial.category, cost: initial.cost, assetGroup: initial.assetGroup || initial.category, ownership: initial.ownership || "Công ty", quantity: initial.quantity || 1, unit: initial.unit || "Cái", projectId: initial.projectId || "",
    purchaseDate: initial.purchaseDate, usefulLifeMonths: initial.usefulLifeMonths, department: initial.department,
    serial: initial.serial, supplier: initial.supplier, warranty: initial.warranty, warrantyEnd: initial.warrantyEnd, note: initial.note,
    customFields: { ...(initial.customFields || {}) },
  } : {
    code: "", name: "", category: categories[0] || "", cost: "", purchaseDate: nowIso().slice(0, 10),
    usefulLifeMonths: 36, department: departments[0] || "", projectId: "", assetGroup: categories[0] || "Thiết bị chính", ownership: "Công ty", quantity: 1, unit: "Cái", serial: "", supplier: "", warranty: true, warrantyEnd: "", note: "",
    customFields: {},
  });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const setCustom = (key) => (e) => setF({ ...f, customFields: { ...f.customFields, [key]: e.target.value } });
  const valid = f.code.trim() && f.name.trim() && f.category && Number(f.quantity) > 0;

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
        <Field label="Đơn vị tính"><input className={inputCls} style={inputStyle} value={f.unit} onChange={set("unit")} placeholder="Cái / Bộ / Mét..." /></Field>
        <Field label="Bộ phận">
          <select className={inputCls} style={inputStyle} value={f.department} onChange={set("department")}>
            {departments.map((d) => <option key={d}>{d}</option>)}
          </select>
        </Field>
        <Field label="Nguyên giá (đ)"><input type="number" className={inputCls} style={inputStyle} value={f.cost} onChange={set("cost")} placeholder="10000000" /></Field>
        <Field label="Thời gian sử dụng (tháng)"><input type="number" className={inputCls} style={inputStyle} value={f.usefulLifeMonths} onChange={set("usefulLifeMonths")} /></Field>
        <Field label="Ngày mua"><input type="date" className={inputCls} style={inputStyle} value={f.purchaseDate} onChange={set("purchaseDate")} /></Field>
        <Field label="Serial"><input className={inputCls} style={inputStyle} value={f.serial} onChange={set("serial")} /></Field>
        <Field label="Nhà cung cấp"><SupplierSearchPicker suppliers={suppliers} value={f.supplier} onPick={v=>setF({...f,supplier:v})}/></Field>
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
