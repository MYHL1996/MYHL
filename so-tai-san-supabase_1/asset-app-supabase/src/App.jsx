/*
 MYHL-QUẢN LÝ TÀI SẢN
 WAREHOUSE CORE V10 FINAL — 24/08/2026

 Mục tiêu:
 - Danh mục tài sản = MASTER DATA, tuyệt đối không tạo phiếu kho.
 - Phiếu nhập/xuất = giao dịch duy nhất làm thay đổi tồn.
 - Tìm kiếm tài sản tự do theo MÃ + TÊN, không phân biệt hoa thường/dấu.
 - Sửa phiếu, xóa phiếu, xóa hàng loạt.
 - Import Excel phiếu kho: xử lý ngày Excel an toàn, không lệch 1 ngày.
 - Báo cáo tồn theo TÀI SẢN + KHO/CÔNG TRÌNH + ngày.

 Cách dùng:
 import { WarehouseCoreV10 } from './MYHL_WAREHOUSE_CORE_V10_FINAL';
 <WarehouseCoreV10 data={data} setData={setData} requireAdmin={requireAdmin} notify={notify} />

 LƯU Ý: Đây là module kho. Nếu App.jsx vẫn render WarehouseTxModal/select cũ
 thì giao diện cũ sẽ tiếp tục xuất hiện. Phải thay phần Kho cũ bằng WarehouseCoreV10.
*/

import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

const BRAND = '#C1272D';
const BRAND_SOFT = '#FBE8E8';
const SURFACE = '#FFFFFF';
const PAPER = '#F8F5F4';
const INK = '#241A19';
const MUTED = '#756B69';
const BORDER = '#E9DEDB';
const DANGER = '#B42318';

const text = (v) => String(v ?? '').trim();

export function normalizeSearch(v) {
  return text(v)
    .toLocaleLowerCase('vi-VN')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd');
}

export function pad2(n) { return String(n).padStart(2, '0'); }

export function isoDate(y, m, d) {
  if (!y || !m || !d) return '';
  const yy = Number(y), mm = Number(m), dd = Number(d);
  if (yy < 1900 || yy > 2200 || mm < 1 || mm > 12 || dd < 1 || dd > 31) return '';
  const out = `${yy}-${pad2(mm)}-${pad2(dd)}`;
  // Reject impossible dates such as 31/02/2026.
  const dt = new Date(Date.UTC(yy, mm - 1, dd));
  if (dt.getUTCFullYear() !== yy || dt.getUTCMonth() !== mm - 1 || dt.getUTCDate() !== dd) return '';
  return out;
}

// UI/internal dates: yyyy-mm-dd, dd/mm/yyyy, yyyy/mm/dd.
export function dateKey(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number' && Number.isFinite(v)) {
    try {
      const p = XLSX.SSF.parse_date_code(v);
      if (p?.y && p?.m && p?.d) return isoDate(p.y, p.m, p.d);
    } catch (_) {}
  }
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    // Only used for UI values. Excel import is intentionally read as raw values
    // and therefore does not go through this branch.
    return isoDate(v.getFullYear(), v.getMonth() + 1, v.getDate());
  }
  const s = text(v);
  let m = s.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})$/);
  if (m) return isoDate(m[3], m[2], m[1]);
  m = s.match(/^(\d{4})[/.\-](\d{1,2})[/.\-](\d{1,2})$/);
  if (m) return isoDate(m[1], m[2], m[3]);
  m = s.match(/^(\d{4}-\d{2}-\d{2})(?:T|\s)/);
  if (m) return dateKey(m[1]);
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? dateKey(s.slice(0, 10)) : '';
}

// Excel-import-only date parser. readExcelRowsV10 uses raw:true/cellDates:false,
// so an Excel date cell remains a serial number and cannot be shifted by timezone.
export function excelDateKey(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number' && Number.isFinite(v)) {
    try {
      const p = XLSX.SSF.parse_date_code(v);
      return p?.y && p?.m && p?.d ? isoDate(p.y, p.m, p.d) : '';
    } catch (_) { return ''; }
  }
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return isoDate(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate());
  }
  const s = text(v);
  // Excel/CSV may contain ISO date-time.
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:T|\s)/);
  if (iso) return isoDate(iso[1], iso[2], iso[3]);
  return dateKey(s);
}

export function dateVN(v) {
  const d = dateKey(v);
  if (!d) return '';
  const [y, m, day] = d.split('-');
  return `${day}/${m}/${y}`;
}

const uid = (prefix = 'id') => `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;

function isNhap(tx) {
  const t = normalizeSearch(tx?.type || tx?.loaiPhieu || tx?.loaiChungTu);
  return t === 'nhap' || t.includes('nhap');
}
function isXuat(tx) {
  const t = normalizeSearch(tx?.type || tx?.loaiPhieu || tx?.loaiChungTu);
  return t === 'xuat' || t.includes('xuat');
}

function assetCode(a) { return text(a?.code || a?.itemCode || a?.maHang || a?.maQuanLy); }
function assetName(a) { return text(a?.name || a?.itemName || a?.tenTaiSan || a?.tenHang); }
function assetUnit(a) { return text(a?.unit || a?.donViTinh || a?.dvt) || 'Cái'; }
function assetCategory(a) { return text(a?.category || a?.loaiTaiSan || a?.loai) || 'Khác'; }
function assetGroup(a) { return text(a?.assetGroup || a?.group || a?.nhomTaiSan || a?.nhom) || 'Thiết bị chính'; }
function assetOwner(a) { return text(a?.ownership || a?.origin || a?.nguonGoc) || 'Công ty'; }
function assetCost(a) { return Number(a?.cost ?? a?.unitCost ?? a?.donGia ?? a?.nguyenGia ?? 0) || 0; }

function resolveAsset(tx, assets) {
  if (tx?.assetId) {
    const byId = assets.find(a => a.id === tx.assetId);
    if (byId) return byId;
  }
  const code = text(tx?.itemCode || tx?.code || tx?.maHang || tx?.maQuanLy);
  if (code) {
    const byCode = assets.find(a => normalizeSearch(assetCode(a)) === normalizeSearch(code));
    if (byCode) return byCode;
  }
  const name = text(tx?.itemName || tx?.name || tx?.tenTaiSan || tx?.tenHang);
  if (name) {
    const byName = assets.find(a => normalizeSearch(assetName(a)) === normalizeSearch(name));
    if (byName) return byName;
  }
  return null;
}

function locationOf(tx, projects = []) {
  if (text(tx?.locationName)) return text(tx.locationName);
  if (tx?.projectId) {
    const p = projects.find(x => x.id === tx.projectId);
    if (p) return text(p.name || p.tenCongTrinh);
  }
  if (text(tx?.warehouseName)) return text(tx.warehouseName);
  return '';
}

function stockAt({ warehouse, assets, projects, assetId, locationName, asOfDate, excludeId = '' }) {
  const cut = dateKey(asOfDate) || '9999-12-31';
  const loc = normalizeSearch(locationName);
  return (warehouse || []).reduce((sum, tx) => {
    if (excludeId && tx.id === excludeId) return sum;
    const d = dateKey(tx.date || tx.ngayThang || tx.ngayChungTu);
    if (!d || d > cut) return sum;
    if (normalizeSearch(locationOf(tx, projects)) !== loc) return sum;
    const a = resolveAsset(tx, assets);
    if (!a || a.id !== assetId) return sum;
    const q = Number(tx.quantity ?? tx.soLuong ?? 0) || 0;
    if (isNhap(tx)) return sum + q;
    if (isXuat(tx)) return sum - q;
    return sum;
  }, 0);
}

function nextVoucher(warehouse, type, date) {
  const prefix = type === 'xuat' ? 'PX' : 'PN';
  const dk = date.replaceAll('-', '');
  const used = new Set((warehouse || []).map(w => text(w.voucherNo).toUpperCase()));
  let n = 1;
  while (used.has(`${prefix}-${dk}-${String(n).padStart(3, '0')}`)) n += 1;
  return `${prefix}-${dk}-${String(n).padStart(3, '0')}`;
}

function selectStyle() {
  return { width: '100%', boxSizing: 'border-box', border: `1px solid ${BORDER}`, borderRadius: 8, padding: '9px 10px', background: '#fff', fontSize: 13 };
}

function Button({ children, onClick, kind = 'default', disabled = false }) {
  const css = kind === 'primary'
    ? { background: BRAND, color: '#fff', borderColor: BRAND }
    : kind === 'danger'
      ? { background: '#fff', color: DANGER, borderColor: '#E7B2AD' }
      : { background: '#fff', color: INK, borderColor: BORDER };
  return <button type="button" disabled={disabled} onClick={onClick}
    style={{ ...css, border: '1px solid', borderRadius: 8, padding: '8px 12px', fontSize: 13, fontWeight: 650, cursor: disabled ? 'not-allowed' : 'pointer', opacity: disabled ? .55 : 1 }}>
    {children}
  </button>;
}

function Input({ value, onChange, placeholder = '', type = 'text', autoFocus = false }) {
  return <input autoFocus={autoFocus} type={type} value={value ?? ''} onChange={e => onChange(e.target.value)} placeholder={placeholder}
    style={{ width: '100%', boxSizing: 'border-box', border: `1px solid ${BORDER}`, borderRadius: 8, padding: '9px 10px', background: '#fff', fontSize: 13 }} />;
}
function Field({ label, children }) {
  return <label style={{ display: 'block' }}><div style={{ fontSize: 12, color: MUTED, fontWeight: 650, marginBottom: 5 }}>{label}</div>{children}</label>;
}
function Modal({ title, onClose, children, wide = false }) {
  return <div style={{ position: 'fixed', inset: 0, zIndex: 3000, background: 'rgba(20,15,14,.38)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 18 }}>
    <div style={{ width: wide ? 'min(980px,100%)' : 'min(900px,100%)', maxHeight: '92vh', overflow: 'auto', background: '#fff', borderRadius: 13, boxShadow: '0 24px 70px rgba(0,0,0,.24)' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '15px 18px', borderBottom: `1px solid ${BORDER}` }}>
        <strong style={{ fontSize: 17 }}>{title}</strong>
        <button type="button" onClick={onClose} style={{ border: 0, background: 'transparent', fontSize: 24, cursor: 'pointer', color: MUTED }}>×</button>
      </div>
      <div style={{ padding: 18 }}>{children}</div>
    </div>
  </div>;
}

function AssetSearchBox({ assets, value, onChange }) {
  const selected = assets.find(a => a.id === value) || null;
  const [q, setQ] = useState(selected ? `${assetCode(selected)} — ${assetName(selected)}` : '');
  const [open, setOpen] = useState(false);
  const rootRef = useRef(null);

  useEffect(() => {
    if (selected) setQ(`${assetCode(selected)} — ${assetName(selected)}`);
    else if (!value) setQ('');
  }, [selected?.id, value]);

  useEffect(() => {
    const close = (e) => { if (rootRef.current && !rootRef.current.contains(e.target)) setOpen(false); };
    document.addEventListener('mousedown', close);
    return () => document.removeEventListener('mousedown', close);
  }, []);

  const results = useMemo(() => {
    const s = normalizeSearch(q.replace(/\s+—\s+.*$/, '').trim() || q);
    if (!s) return assets.slice(0, 30);
    return assets.map(a => {
      const code = normalizeSearch(assetCode(a));
      const name = normalizeSearch(assetName(a));
      const serial = normalizeSearch(a.serial || a.soSerial);
      let score = 0;
      if (code === s) score += 1000;
      if (name === s) score += 950;
      if (code.startsWith(s)) score += 700;
      if (name.startsWith(s)) score += 650;
      if (code.includes(s)) score += 350;
      if (name.includes(s)) score += 300;
      if (serial.includes(s)) score += 120;
      return { a, score };
    }).filter(x => x.score > 0).sort((x, y) => y.score - x.score).slice(0, 30).map(x => x.a);
  }, [assets, q]);

  const choose = (a) => {
    onChange(a.id);
    setQ(`${assetCode(a)} — ${assetName(a)}`);
    setOpen(false);
  };

  return <div ref={rootRef} style={{ position: 'relative' }}>
    <div style={{ position: 'relative' }}>
      <input value={q} onFocus={() => setOpen(true)} onChange={e => { setQ(e.target.value); setOpen(true); if (!e.target.value.trim()) onChange(''); }}
        placeholder="Gõ tự do: máy xúc / MX-01 / máy khoan..."
        style={{ width: '100%', boxSizing: 'border-box', border: `1px solid ${open ? BRAND : BORDER}`, borderRadius: 8, padding: '10px 34px 10px 11px', background: '#fff', fontSize: 13, outline: 'none' }} />
      {q && <button type="button" onClick={() => { setQ(''); onChange(''); setOpen(true); }} style={{ position: 'absolute', right: 7, top: 5, border: 0, background: 'transparent', fontSize: 18, color: MUTED, cursor: 'pointer' }}>×</button>}
    </div>
    {open && <div style={{ position: 'absolute', zIndex: 5000, left: 0, right: 0, top: 'calc(100% + 4px)', background: '#fff', border: `1px solid ${BORDER}`, borderRadius: 9, boxShadow: '0 15px 38px rgba(0,0,0,.15)', maxHeight: 330, overflowY: 'auto' }}>
      <div style={{ padding: '7px 10px', borderBottom: `1px solid ${BORDER}`, fontSize: 11, color: MUTED }}>Tìm theo mã hoặc tên · {results.length} gợi ý</div>
      {results.length ? results.map(a => <button key={a.id} type="button" onClick={() => choose(a)} style={{ width: '100%', display: 'block', textAlign: 'left', border: 0, borderBottom: `1px solid #F2ECEA`, background: '#fff', padding: '10px 11px', cursor: 'pointer' }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: INK }}>{assetCode(a) || '—'} <span style={{ color: MUTED }}>—</span> {assetName(a) || '—'}</div>
        <div style={{ fontSize: 11, color: MUTED, marginTop: 3 }}>{assetCategory(a)} · {assetGroup(a)} · {assetOwner(a)} · {assetUnit(a)}</div>
      </button>) : <div style={{ padding: 15, color: MUTED, fontSize: 12 }}>Không tìm thấy tài sản. Hãy kiểm tra lại mã hoặc tên trong Danh mục tài sản.</div>}
    </div>}
  </div>;
}

function emptyForm(type, assets, warehouse) {
  const today = new Date();
  const d = `${today.getFullYear()}-${pad2(today.getMonth() + 1)}-${pad2(today.getDate())}`;
  return { id: '', type, voucherNo: '', date: d, assetId: '', quantity: 1, unitCost: 0, receiver: '', locationType: 'project', projectId: '', warehouseName: '', note: '' };
}

function validateForm(form, data, editingId = '') {
  const assets = data.assets || [], projects = data.projects || [], warehouse = data.warehouse || [];
  const asset = assets.find(a => a.id === form.assetId);
  if (!asset) return { error: 'Bạn chưa chọn tài sản. Hãy gõ mã hoặc tên và chọn một dòng gợi ý.' };
  const d = dateKey(form.date);
  if (!d) return { error: 'Ngày chứng từ không hợp lệ.' };
  const qty = Number(form.quantity);
  if (!(qty > 0)) return { error: 'Số lượng phải lớn hơn 0.' };
  let loc = '';
  let projectId = null;
  if (form.locationType === 'project') {
    const p = projects.find(x => x.id === form.projectId);
    if (!p) return { error: 'Vui lòng chọn Công trình.' };
    loc = text(p.name || p.tenCongTrinh);
    projectId = p.id;
  } else {
    loc = text(form.warehouseName);
    if (!loc) return { error: 'Vui lòng nhập Tên kho.' };
  }
  const voucher = text(form.voucherNo) || nextVoucher(warehouse.filter(w => w.id !== editingId), form.type, d);
  const duplicate = warehouse.some(w => w.id !== editingId && normalizeSearch(w.voucherNo) === normalizeSearch(voucher));
  if (duplicate) return { error: `Số phiếu ${voucher} đã tồn tại.` };

  if (form.type === 'xuat') {
    const stock = stockAt({ warehouse, assets, projects, assetId: asset.id, locationName: loc, asOfDate: d, excludeId: editingId });
    if (qty > stock) return { error: `Không đủ tồn tại ${loc}. ${assetCode(asset)} — ${assetName(asset)} hiện còn ${stock} ${assetUnit(asset)}.` };
  }
  return { asset, date: d, quantity: qty, locationName: loc, projectId, voucher };
}

export function saveWarehouseTransactionV10({ form, data, setData, requireAdmin, notify }) {
  if (requireAdmin && !requireAdmin()) return false;
  const editingId = text(form.id);
  const checked = validateForm(form, data, editingId);
  if (checked.error) { notify?.(checked.error); return false; }
  const asset = checked.asset;
  const old = (data.warehouse || []).find(w => w.id === editingId);
  const tx = {
    id: editingId || uid('wh'),
    voucherNo: checked.voucher,
    assetId: asset.id,
    type: form.type,
    quantity: checked.quantity,
    date: checked.date,
    unitCost: Number(form.unitCost) || 0,
    total: checked.quantity * (Number(form.unitCost) || 0),
    unit: assetUnit(asset),
    receiver: text(form.receiver),
    note: text(form.note),
    category: assetCategory(asset),
    assetGroup: assetGroup(asset),
    ownership: assetOwner(asset),
    locationType: form.locationType,
    locationName: checked.locationName,
    warehouseName: form.locationType === 'warehouse' ? checked.locationName : '',
    projectId: checked.projectId,
    itemName: assetName(asset),
    itemCode: assetCode(asset),
    serial: text(asset.serial || asset.soSerial)
  };
  const warehouse = data.warehouse || [];
  const nextWarehouse = editingId ? warehouse.map(w => w.id === editingId ? tx : w) : [tx, ...warehouse];
  const oldHistoryIds = new Set((data.transactions || []).filter(t => t.warehouseTxId === editingId).map(t => t.id));
  const history = {
    id: oldHistoryIds.size ? [...oldHistoryIds][0] : uid('tx'),
    warehouseTxId: tx.id,
    assetId: tx.assetId,
    type: tx.type === 'xuat' ? 'xuat_kho' : 'nhap_kho',
    date: tx.date,
    title: `${tx.type === 'xuat' ? 'Xuất kho' : 'Nhập kho'} ${tx.voucherNo}`,
    detail: `${tx.itemName} · ${tx.locationName} · ${tx.receiver || ''}`,
    amount: tx.total
  };
  const oldTransactions = (data.transactions || []).filter(t => !oldHistoryIds.has(t.id));
  setData({ ...data, warehouse: nextWarehouse, transactions: [history, ...oldTransactions] });
  notify?.(editingId ? `Đã sửa phiếu ${tx.voucherNo}.` : `Đã lập phiếu ${tx.voucherNo}.`);
  return true;
}

export function deleteWarehouseTransactionsV10({ ids = [], data, setData, requireAdmin, notify, skipConfirm = false }) {
  if (requireAdmin && !requireAdmin()) return false;
  const selected = new Set(ids.filter(Boolean));
  const rows = (data.warehouse || []).filter(w => selected.has(w.id));
  if (!rows.length) { notify?.('Chưa chọn phiếu cần xóa.'); return false; }
  const names = [...new Set(rows.map(w => text(w.voucherNo)).filter(Boolean))];
  if (!skipConfirm && !window.confirm(`Bạn chắc chắn muốn xóa ${rows.length} dòng / ${names.length} số phiếu?\n\n${names.join(', ')}\n\nDữ liệu xóa sẽ làm thay đổi báo cáo nhập - xuất - tồn.`)) return false;
  const idsSet = new Set(rows.map(w => w.id));
  setData({
    ...data,
    warehouse: (data.warehouse || []).filter(w => !idsSet.has(w.id)),
    transactions: (data.transactions || []).filter(t => !idsSet.has(t.warehouseTxId))
  });
  notify?.(`Đã xóa ${names.length} phiếu kho.`);
  return true;
}

export function readExcelRowsV10(file) {
  return file.arrayBuffer().then(buffer => {
    const wb = XLSX.read(buffer, { type: 'array', cellDates: false, raw: true, dense: false });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true });
  });
}

function rowValue(row, names) {
  for (const n of names) if (row[n] !== undefined && row[n] !== '') return row[n];
  return '';
}

function findAssetFromRow(row, assets) {
  const code = text(rowValue(row, ['Mã hàng', 'Mã quản lý', 'Mã tài sản', 'Mã thiết bị', 'code']));
  const name = text(rowValue(row, ['Tên tài sản', 'Tên hàng', 'Tên vật tư', 'Tên thiết bị', 'name']));
  const byCode = code ? assets.find(a => normalizeSearch(assetCode(a)) === normalizeSearch(code)) : null;
  const byName = name ? assets.find(a => normalizeSearch(assetName(a)) === normalizeSearch(name)) : null;
  if (byCode && byName && byCode.id !== byName.id) return { error: `Mã ${code} và Tên ${name} không cùng một tài sản.` };
  if (!byCode && !byName) return { error: `Không tìm thấy tài sản theo mã ${code || '(trống)'} hoặc tên ${name || '(trống)'}.` };
  return { asset: byCode || byName };
}

export async function importAssetCatalogV10({ file, data, setData, requireAdmin, notify }) {
  if (requireAdmin && !requireAdmin()) return false;
  try {
    const rows = await readExcelRowsV10(file);
    const existing = data.assets || [];
    const byCode = new Map(existing.map(a => [normalizeSearch(assetCode(a)), a]));
    const imported = [], seen = new Set(), errors = [];
    rows.forEach((r, i) => {
      const line = i + 2;
      const code = text(rowValue(r, ['Mã quản lý', 'Mã hàng', 'Mã tài sản', 'code']));
      const name = text(rowValue(r, ['Tên tài sản', 'Tên hàng', 'Tên thiết bị', 'name']));
      if (!code || !name) { errors.push(`Dòng ${line}: thiếu mã hoặc tên.`); return; }
      const key = normalizeSearch(code);
      if (seen.has(key)) { errors.push(`Dòng ${line}: mã ${code} bị trùng trong file.`); return; }
      seen.add(key);
      const old = byCode.get(key);
      imported.push({
        ...old,
        id: old?.id || uid('as'),
        code, name,
        category: text(rowValue(r, ['Loại tài sản', 'Loại'])) || old?.category || 'Khác',
        assetGroup: text(rowValue(r, ['Nhóm tài sản', 'Nhóm'])) || old?.assetGroup || 'Thiết bị chính',
        ownership: text(rowValue(r, ['Nguồn gốc'])) || old?.ownership || 'Công ty',
        unit: text(rowValue(r, ['Đơn vị tính', 'ĐVT'])) || old?.unit || 'Cái',
        serial: text(rowValue(r, ['Serial', 'Số serial'])) || old?.serial || '',
        cost: Number(rowValue(r, ['Đơn giá', 'Nguyên giá'])) || old?.cost || 0
      });
    });
    if (!imported.length) { notify?.(errors[0] || 'Không có dòng danh mục hợp lệ.'); return false; }
    const keys = new Set(imported.map(a => normalizeSearch(assetCode(a))));
    // IMPORTANT: only assets change. warehouse and transactions are untouched.
    setData({ ...data, assets: [...imported, ...existing.filter(a => !keys.has(normalizeSearch(assetCode(a))))] });
    notify?.(`Đã cập nhật ${imported.length} danh mục tài sản. Không tạo phiếu nhập/xuất.`);
    return true;
  } catch (e) {
    console.error(e);
    notify?.(`Không đọc được file danh mục Excel: ${e?.message || e}`);
    return false;
  }
}

export async function importWarehouseExcelV10({ file, data, setData, requireAdmin, notify }) {
  if (requireAdmin && !requireAdmin()) return false;
  try {
    const rows = await readExcelRowsV10(file);
    const assets = data.assets || [], projects = data.projects || [], existing = data.warehouse || [];
    const imported = [], errors = [];
    const vouchers = new Set(existing.map(w => normalizeSearch(w.voucherNo)).filter(Boolean));

    rows.forEach((r, i) => {
      const line = i + 2;
      const rawType = normalizeSearch(rowValue(r, ['Loại phiếu', 'Loại chứng từ', 'Loại']));
      const type = rawType.includes('xuat') ? 'xuat' : 'nhap';
      const d = excelDateKey(rowValue(r, ['Ngày tháng', 'Ngày chứng từ', 'Ngày', 'date']));
      if (!d) { errors.push(`Dòng ${line}: ngày không hợp lệ. Hãy dùng dd/mm/yyyy hoặc yyyy-mm-dd.`); return; }
      const found = findAssetFromRow(r, assets);
      if (found.error) { errors.push(`Dòng ${line}: ${found.error}`); return; }
      const a = found.asset;
      const q = Number(rowValue(r, ['Số lượng', 'SL', 'quantity']));
      if (!(q > 0)) { errors.push(`Dòng ${line}: số lượng phải > 0.`); return; }
      const loc = text(rowValue(r, ['Kho/Công trình', 'Công trình', 'Tên kho', 'Kho', 'locationName']));
      if (!loc) { errors.push(`Dòng ${line}: thiếu Kho/Công trình.`); return; }
      const project = projects.find(p => normalizeSearch(p.name || p.tenCongTrinh) === normalizeSearch(loc));
      const rawLocationType = normalizeSearch(rowValue(r, ['Loại địa điểm', 'Địa điểm']));
      const locationType = rawLocationType.includes('cong trinh') ? 'project' : rawLocationType.includes('kho') ? 'warehouse' : project ? 'project' : 'warehouse';
      const voucherInput = text(rowValue(r, ['Số phiếu', 'Số chứng từ', 'Voucher']));
      const voucher = voucherInput || `${type === 'xuat' ? 'PX' : 'PN'}-${d.replaceAll('-', '')}-${String(existing.length + imported.length + 1).padStart(3, '0')}`;
      const vk = normalizeSearch(voucher);
      if (vouchers.has(vk)) { errors.push(`Dòng ${line}: số phiếu ${voucher} bị trùng.`); return; }
      vouchers.add(vk);
      const unitCost = Number(rowValue(r, ['Đơn giá', 'Đơn giá nhập', 'Đơn giá xuất'])) || 0;
      imported.push({
        id: uid('wh'), voucherNo: voucher, assetId: a.id, type, quantity: q, date: d,
        unitCost, total: q * unitCost, unit: text(rowValue(r, ['Đơn vị tính', 'ĐVT'])) || assetUnit(a),
        receiver: text(rowValue(r, ['Người giao/nhận', 'Người giao', 'Người nhận'])),
        note: text(rowValue(r, ['Ghi chú', 'Diễn giải'])), category: assetCategory(a), assetGroup: assetGroup(a), ownership: assetOwner(a),
        locationType, locationName: loc, warehouseName: locationType === 'warehouse' ? loc : '', projectId: locationType === 'project' ? project?.id || null : null,
        itemName: assetName(a), itemCode: assetCode(a), serial: text(a.serial || a.soSerial)
      });
    });

    if (errors.length) {
      notify?.(`Không import để tránh dữ liệu dở dang: ${errors.length} lỗi. ${errors[0]}`);
      return false;
    }

    // Validate all imported OUT rows against stock AFTER prior imported IN rows.
    const working = [...existing];
    for (const tx of imported) {
      if (tx.type === 'xuat') {
        const stock = stockAt({ warehouse: working, assets, projects, assetId: tx.assetId, locationName: tx.locationName, asOfDate: tx.date });
        if (tx.quantity > stock) {
          notify?.(`Import dừng: ${tx.voucherNo} xuất ${tx.quantity} nhưng ${tx.locationName} chỉ còn ${stock}.`);
          return false;
        }
      }
      working.push(tx);
    }

    const histories = imported.map(tx => ({
      id: uid('tx'), warehouseTxId: tx.id, assetId: tx.assetId,
      type: tx.type === 'xuat' ? 'xuat_kho' : 'nhap_kho', date: tx.date,
      title: `${tx.type === 'xuat' ? 'Xuất kho' : 'Nhập kho'} ${tx.voucherNo}`,
      detail: `${tx.itemName} · ${tx.locationName} · ${tx.receiver || ''}`, amount: tx.total
    }));
    setData({ ...data, warehouse: [...imported, ...existing], transactions: [...histories, ...(data.transactions || [])] });
    notify?.(`Đã import ${imported.length} phiếu kho. Ngày Excel được giữ nguyên theo ô ngày.`);
    return true;
  } catch (e) {
    console.error(e);
    notify?.(`Không đọc được file Excel phiếu kho: ${e?.message || e}`);
    return false;
  }
}

export function makeWarehouseReportV10({ warehouse = [], assets = [], projects = [], filter = {} }) {
  const cut = dateKey(filter.asOfDate) || '9999-12-31';
  const map = new Map();
  for (const tx of warehouse) {
    const d = dateKey(tx.date || tx.ngayThang || tx.ngayChungTu);
    if (!d || d > cut) continue;
    const a = resolveAsset(tx, assets);
    if (!a) continue;
    const loc = locationOf(tx, projects) || 'Chưa xác định';
    const key = `${a.id}¦${normalizeSearch(loc)}`;
    if (!map.has(key)) map.set(key, { assetId: a.id, code: assetCode(a) || tx.itemCode, name: assetName(a) || tx.itemName, location: loc, category: assetCategory(a), group: assetGroup(a), ownership: assetOwner(a), unit: assetUnit(a), inQty: 0, outQty: 0, inValue: 0, outValue: 0 });
    const r = map.get(key);
    const q = Number(tx.quantity ?? tx.soLuong ?? 0) || 0;
    const cost = Number(tx.total ?? tx.thanhTien) || q * (Number(tx.unitCost ?? tx.unitPrice ?? tx.donGia) || 0);
    if (isNhap(tx)) { r.inQty += q; r.inValue += cost; }
    else if (isXuat(tx)) { r.outQty += q; r.outValue += cost; }
  }
  let rows = [...map.values()].map(r => ({ ...r, balanceQty: r.inQty - r.outQty, balanceValue: r.inValue - r.outValue }));
  if (filter.category) rows = rows.filter(r => normalizeSearch(r.category) === normalizeSearch(filter.category));
  if (filter.group) rows = rows.filter(r => normalizeSearch(r.group) === normalizeSearch(filter.group));
  if (filter.ownership) rows = rows.filter(r => normalizeSearch(r.ownership) === normalizeSearch(filter.ownership));
  if (filter.locationName) rows = rows.filter(r => normalizeSearch(r.location) === normalizeSearch(filter.locationName));
  rows = rows.filter(r => r.balanceQty > 0).sort((a, b) => `${a.location}|${a.group}|${a.category}|${a.code}|${a.name}`.localeCompare(`${b.location}|${b.group}|${b.category}|${b.code}|${b.name}`, 'vi'));
  return { rows, summary: { qty: rows.reduce((s, r) => s + r.balanceQty, 0), value: rows.reduce((s, r) => s + r.balanceValue, 0), assets: rows.length }, asOfDate: cut === '9999-12-31' ? '' : cut, asOfDateVN: cut === '9999-12-31' ? '' : dateVN(cut) };
}

function TransactionForm({ mode, initial, data, onClose, onSaved, requireAdmin, notify, setData }) {
  const [form, setForm] = useState(initial || emptyForm(mode, data.assets || [], data.warehouse || []));
  const a = (data.assets || []).find(x => x.id === form.assetId) || null;
  const patch = (k, v) => setForm(f => ({ ...f, [k]: v }));
  useEffect(() => {
    if (a && !form.id) setForm(f => ({ ...f, unitCost: f.unitCost ? f.unitCost : assetCost(a) }));
  }, [a?.id]);
  const total = (Number(form.quantity) || 0) * (Number(form.unitCost) || 0);
  const editing = Boolean(form.id);
  return <Modal wide title={editing ? `Sửa phiếu ${form.voucherNo || ''}` : `Lập phiếu ${mode === 'nhap' ? 'nhập kho' : 'xuất kho'}`} onClose={onClose}>
    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
      <Field label="Số phiếu"><Input value={form.voucherNo} onChange={v => patch('voucherNo', v)} placeholder={mode === 'nhap' ? 'PN-20260824-001' : 'PX-20260824-001'} /></Field>
      <Field label="Ngày tháng"><Input type="date" value={form.date} onChange={v => patch('date', v)} /></Field>
      <div style={{ gridColumn: '1 / -1' }}><Field label="Tên tài sản / Mã hàng — tìm kiếm tự do"><AssetSearchBox assets={data.assets || []} value={form.assetId} onChange={id => { const x = (data.assets || []).find(z => z.id === id); setForm(f => ({ ...f, assetId: id, unitCost: f.unitCost || assetCost(x) })); }} /></Field></div>
      <Field label="Mã hàng"><Input value={assetCode(a)} onChange={() => {}} /></Field>
      <Field label="Loại tài sản"><Input value={assetCategory(a)} onChange={() => {}} /></Field>
      <Field label="Nhóm tài sản"><Input value={assetGroup(a)} onChange={() => {}} /></Field>
      <Field label="Nguồn gốc"><Input value={assetOwner(a)} onChange={() => {}} /></Field>
      <Field label="Đơn vị tính"><Input value={assetUnit(a)} onChange={() => {}} /></Field>
      <Field label="Kho / Công trình">
        <select value={form.locationType === 'project' ? form.projectId : '__warehouse__'} onChange={e => e.target.value === '__warehouse__' ? setForm(f => ({ ...f, locationType: 'warehouse', projectId: '' })) : setForm(f => ({ ...f, locationType: 'project', projectId: e.target.value, warehouseName: '' }))} style={selectStyle()}>
          <option value="">-- Chọn công trình --</option>
          {(data.projects || []).map(p => <option key={p.id} value={p.id}>{p.name || p.tenCongTrinh}</option>)}
          <option value="__warehouse__">Kho trung tâm / kho khác</option>
        </select>
      </Field>
      {form.locationType === 'warehouse' && <Field label="Tên kho"><Input value={form.warehouseName} onChange={v => patch('warehouseName', v)} placeholder="Kho Long Thành" /></Field>}
      <Field label="Người giao / nhận"><Input value={form.receiver} onChange={v => patch('receiver', v)} /></Field>
      <Field label="Số lượng"><Input type="number" value={form.quantity} onChange={v => patch('quantity', v)} /></Field>
      <Field label="Đơn giá"><Input type="number" value={form.unitCost} onChange={v => patch('unitCost', v)} /></Field>
      <Field label="Thành tiền"><Input value={total.toLocaleString('vi-VN')} onChange={() => {}} /></Field>
      <div style={{ gridColumn: '1 / -1' }}><Field label="Ghi chú"><textarea value={form.note} onChange={e => patch('note', e.target.value)} rows={3} style={{ width: '100%', boxSizing: 'border-box', border: `1px solid ${BORDER}`, borderRadius: 8, padding: 10, fontSize: 13 }} /></Field></div>
    </div>
    <div style={{ marginTop: 16, display: 'flex', justifyContent: 'flex-end', gap: 8 }}><Button onClick={onClose}>Hủy</Button><Button kind="primary" onClick={() => { const ok = saveWarehouseTransactionV10({ form: { ...form, type: mode }, data, setData, requireAdmin, notify }); if (ok) onSaved?.(); }}> {editing ? 'Lưu thay đổi' : 'Lưu phiếu'} </Button></div>
  </Modal>;
}

function TransactionTable({ type, data, setData, requireAdmin, notify }) {
  const assets = data.assets || [], projects = data.projects || [], warehouse = data.warehouse || [];
  const [q, setQ] = useState(''), [loc, setLoc] = useState(''), [selected, setSelected] = useState([]), [edit, setEdit] = useState(null);
  const rows = useMemo(() => {
    const s = normalizeSearch(q);
    return warehouse.filter(w => {
      if (type === 'nhap' && !isNhap(w)) return false;
      if (type === 'xuat' && !isXuat(w)) return false;
      if (loc && normalizeSearch(locationOf(w, projects)) !== normalizeSearch(loc)) return false;
      if (!s) return true;
      const a = resolveAsset(w, assets);
      const hay = [w.voucherNo, w.itemCode, w.itemName, w.receiver, locationOf(w, projects), assetCode(a), assetName(a)].map(normalizeSearch).join(' ');
      return hay.includes(s);
    }).sort((a, b) => (dateKey(b.date) || '').localeCompare(dateKey(a.date) || '') || String(b.id).localeCompare(String(a.id)));
  }, [warehouse, assets, projects, type, q, loc]);
  const locations = [...new Set(warehouse.map(w => locationOf(w, projects)).filter(Boolean))].sort((a, b) => normalizeSearch(a).localeCompare(normalizeSearch(b), 'vi'));
  const all = rows.length > 0 && rows.every(r => selected.includes(r.id));
  const toggleAll = () => setSelected(all ? selected.filter(id => !rows.some(r => r.id === id)) : [...new Set([...selected, ...rows.map(r => r.id)])]);
  const remove = () => { const ok = deleteWarehouseTransactionsV10({ ids: selected, data, setData, requireAdmin, notify }); if (ok) setSelected([]); };
  return <div>
    <div style={{ display: 'grid', gridTemplateColumns: 'minmax(300px,1fr) 240px auto auto', gap: 8, marginBottom: 10, alignItems: 'center' }}>
      <Input value={q} onChange={setQ} placeholder="🔎 Gõ tên tài sản hoặc mã hàng, số phiếu, công trình..." />
      <select value={loc} onChange={e => setLoc(e.target.value)} style={selectStyle()}><option value="">Tất cả Kho/Công trình</option>{locations.map(x => <option key={x}>{x}</option>)}</select>
      <Button disabled={!selected.length} onClick={remove} kind="danger">Xóa đã chọn ({selected.length})</Button>
      <span style={{ fontSize: 12, color: MUTED }}>{rows.length} phiếu</span>
    </div>
    <div style={{ overflow: 'auto', border: `1px solid ${BORDER}`, borderRadius: 10, background: '#fff' }}>
      <table style={{ width: '100%', minWidth: 1180, borderCollapse: 'collapse' }}>
        <thead><tr style={{ background: BRAND_SOFT }}>
          <th style={{ padding: 9 }}><input type="checkbox" checked={all} onChange={toggleAll} /></th>
          {['Số phiếu', 'Ngày', 'Mã hàng', 'Tên tài sản', 'Kho/Công trình', 'Nhóm', 'Nguồn gốc', 'SL', 'ĐVT', 'Người giao/nhận', 'Thao tác'].map(h => <th key={h} style={{ padding: 9, textAlign: 'left', fontSize: 12, color: MUTED }}>{h}</th>)}
        </tr></thead>
        <tbody>{rows.map(w => { const a = resolveAsset(w, assets); return <tr key={w.id} style={{ borderTop: `1px solid ${BORDER}` }}>
          <td style={{ padding: 9 }}><input type="checkbox" checked={selected.includes(w.id)} onChange={() => setSelected(s => s.includes(w.id) ? s.filter(x => x !== w.id) : [...s, w.id])} /></td>
          <td style={{ padding: 9, fontFamily: 'monospace' }}>{w.voucherNo || '—'}</td>
          <td style={{ padding: 9 }}>{dateVN(w.date)}</td>
          <td style={{ padding: 9, fontWeight: 650 }}>{assetCode(a) || w.itemCode || '—'}</td>
          <td style={{ padding: 9 }}>{assetName(a) || w.itemName || '—'}</td>
          <td style={{ padding: 9 }}>{locationOf(w, projects) || '—'}</td>
          <td style={{ padding: 9 }}>{assetGroup(a) || w.assetGroup}</td>
          <td style={{ padding: 9 }}>{assetOwner(a) || w.ownership}</td>
          <td style={{ padding: 9 }}>{Number(w.quantity || 0).toLocaleString('vi-VN')}</td>
          <td style={{ padding: 9 }}>{w.unit || assetUnit(a)}</td>
          <td style={{ padding: 9 }}>{w.receiver || '—'}</td>
          <td style={{ padding: 9, whiteSpace: 'nowrap' }}><Button onClick={() => setEdit(w)}>Sửa</Button></td>
        </tr>; })}
        {!rows.length && <tr><td colSpan={12} style={{ padding: 40, textAlign: 'center', color: MUTED }}>Không có phiếu phù hợp với điều kiện tìm kiếm.</td></tr>}</tbody>
      </table>
    </div>
    {edit && <TransactionForm mode={isXuat(edit) ? 'xuat' : 'nhap'} initial={{ ...edit, id: edit.id }} data={data} setData={setData} requireAdmin={requireAdmin} notify={notify} onClose={() => setEdit(null)} onSaved={() => setEdit(null)} />}
  </div>;
}

function Report({ data, filter, setFilter }) {
  const report = useMemo(() => makeWarehouseReportV10({ warehouse: data.warehouse || [], assets: data.assets || [], projects: data.projects || [], filter }), [data.warehouse, data.assets, data.projects, filter]);
  const categories = [...new Set((data.assets || []).map(assetCategory).filter(Boolean))];
  const groups = [...new Set((data.assets || []).map(assetGroup).filter(Boolean))];
  const owners = [...new Set((data.assets || []).map(assetOwner).filter(Boolean))];
  const locations = [...new Set((data.warehouse || []).map(w => locationOf(w, data.projects || [])).filter(Boolean))];
  return <div>
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5,minmax(150px,1fr))', gap: 10, padding: 14, border: `1px solid ${BORDER}`, borderRadius: 10, background: PAPER, marginBottom: 12 }}>
      <Field label="Loại tài sản"><select value={filter.category} onChange={e => setFilter(f => ({ ...f, category: e.target.value }))} style={selectStyle()}><option value="">Tất cả</option>{categories.map(x => <option key={x}>{x}</option>)}</select></Field>
      <Field label="Nhóm tài sản"><select value={filter.group} onChange={e => setFilter(f => ({ ...f, group: e.target.value }))} style={selectStyle()}><option value="">Tất cả</option>{groups.map(x => <option key={x}>{x}</option>)}</select></Field>
      <Field label="Nguồn gốc"><select value={filter.ownership} onChange={e => setFilter(f => ({ ...f, ownership: e.target.value }))} style={selectStyle()}><option value="">Tất cả</option>{owners.map(x => <option key={x}>{x}</option>)}</select></Field>
      <Field label="Kho/Công trình"><select value={filter.locationName} onChange={e => setFilter(f => ({ ...f, locationName: e.target.value }))} style={selectStyle()}><option value="">Tất cả</option>{locations.map(x => <option key={x}>{x}</option>)}</select></Field>
      <Field label="Tồn đến ngày"><Input type="date" value={filter.asOfDate} onChange={v => setFilter(f => ({ ...f, asOfDate: v }))} /></Field>
    </div>
    <div style={{ display: 'flex', gap: 10, marginBottom: 10, flexWrap: 'wrap' }}>
      <span style={{ background: BRAND_SOFT, color: BRAND, borderRadius: 8, padding: '7px 10px', fontSize: 12, fontWeight: 700 }}>Thiết bị tồn: {report.summary.qty.toLocaleString('vi-VN')}</span>
      <span style={{ background: '#FFF4D6', color: '#805A00', borderRadius: 8, padding: '7px 10px', fontSize: 12, fontWeight: 700 }}>Giá trị tồn: {Math.round(report.summary.value).toLocaleString('vi-VN')} đ</span>
      <span style={{ color: MUTED, fontSize: 12, paddingTop: 8 }}>Tồn đến {report.asOfDateVN || 'hiện tại'}</span>
    </div>
    <div style={{ overflow: 'auto', border: `1px solid ${BORDER}`, borderRadius: 10, background: '#fff' }}>
      <table style={{ width: '100%', minWidth: 1250, borderCollapse: 'collapse' }}><thead><tr style={{ background: BRAND_SOFT }}>
        {['Kho/Công trình', 'Loại tài sản', 'Nhóm tài sản', 'Nguồn gốc', 'Mã hàng', 'Tên tài sản', 'ĐVT', 'Nhập lũy kế', 'Xuất lũy kế', 'Tồn đến ngày', 'Giá trị tồn'].map(h => <th key={h} style={{ padding: 9, textAlign: 'left', fontSize: 12, color: MUTED }}>{h}</th>)}
      </tr></thead><tbody>
        {report.rows.map(r => <tr key={`${r.assetId}-${r.location}`} style={{ borderTop: `1px solid ${BORDER}` }}>{[r.location, r.category, r.group, r.ownership, r.code, r.name, r.unit, r.inQty, r.outQty, r.balanceQty, r.balanceValue].map((v, i) => <td key={i} style={{ padding: 9, fontSize: 12 }}>{typeof v === 'number' ? (i === 10 ? Math.round(v).toLocaleString('vi-VN') : v.toLocaleString('vi-VN')) : v}</td>)}</tr>)}
        {!report.rows.length && <tr><td colSpan={11} style={{ padding: 45, textAlign: 'center', color: MUTED }}>Không có thiết bị tồn theo điều kiện lọc.</td></tr>}
      </tbody></table>
    </div>
  </div>;
}

export function downloadWarehouseTemplateV10() {
  const headers = ['Loại phiếu', 'Số phiếu', 'Ngày tháng', 'Tên tài sản', 'Mã hàng', 'Kho/Công trình', 'Loại địa điểm', 'Loại tài sản', 'Nhóm tài sản', 'Nguồn gốc', 'Người giao/nhận', 'Số lượng', 'Đơn vị tính', 'Đơn giá', 'Thành tiền', 'Ghi chú'];
  const rows = [headers, ['Nhập kho', 'PN-20260824-001', '24/08/2026', 'Máy xúc 01', 'MX-01', 'Cao Xà Lá - Thanh Xuân', 'Công trình', 'Máy xúc', 'Thiết bị chính', 'Thuê', 'Nguyễn Văn A', 1, 'Cái', 0, 0, '']];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['C2'].z = 'dd/mm/yyyy';
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Phiếu kho');
  XLSX.writeFile(wb, 'Mau_Import_Phieu_Nhap_Xuat_Kho_V10.xlsx');
}

export function WarehouseCoreV10({ data, setData, requireAdmin, notify }) {
  const [tab, setTab] = useState('nhap');
  const [modal, setModal] = useState(null);
  const [filter, setFilter] = useState({ asOfDate: `${new Date().getFullYear()}-${pad2(new Date().getMonth() + 1)}-${pad2(new Date().getDate())}`, category: '', group: '', ownership: '', locationName: '' });
  const importInput = useRef(null);
  const doImport = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    await importWarehouseExcelV10({ file, data, setData, requireAdmin, notify });
    e.target.value = '';
  };
  return <div style={{ minHeight: '100%', background: PAPER, color: INK }}>
    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 10, flexWrap: 'wrap', marginBottom: 14 }}>
      <div><h2 style={{ margin: 0, fontSize: 22 }}>Kho — Nhập / Xuất / Tồn</h2><div style={{ color: MUTED, fontSize: 12, marginTop: 3 }}>Quản lý phiếu và tồn thiết bị theo Kho/Công trình</div></div>
      <div style={{ display: 'flex', gap: 8 }}><Button onClick={downloadWarehouseTemplateV10}>Tải mẫu Excel</Button><label style={{ display: 'inline-flex', alignItems: 'center', border: `1px solid ${BRAND}`, color: '#fff', background: BRAND, borderRadius: 8, padding: '8px 12px', fontSize: 13, fontWeight: 650, cursor: 'pointer' }}><input ref={importInput} hidden type="file" accept=".xlsx,.xls,.csv" onChange={doImport} />Import phiếu Excel</label></div>
    </div>
    <div style={{ display: 'flex', borderBottom: `1px solid ${BORDER}`, gap: 3, marginBottom: 14 }}>
      {[['nhap', 'Phiếu nhập kho'], ['xuat', 'Phiếu xuất kho'], ['report', 'Báo cáo nhập xuất tồn']].map(([id, label]) => <button key={id} type="button" onClick={() => setTab(id)} style={{ border: 0, borderBottom: tab === id ? `3px solid ${BRAND}` : '3px solid transparent', background: 'transparent', padding: '11px 15px', fontSize: 14, fontWeight: tab === id ? 750 : 500, color: tab === id ? BRAND : INK, cursor: 'pointer' }}>{label}</button>)}
    </div>
    {tab !== 'report' && <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 10 }}><Button kind="primary" onClick={() => setModal({ type: tab, initial: emptyForm(tab, data.assets || [], data.warehouse || []) })}>+ Lập phiếu {tab === 'nhap' ? 'nhập' : 'xuất'}</Button></div>}
    {tab === 'nhap' && <TransactionTable type="nhap" data={data} setData={setData} requireAdmin={requireAdmin} notify={notify} />}
    {tab === 'xuat' && <TransactionTable type="xuat" data={data} setData={setData} requireAdmin={requireAdmin} notify={notify} />}
    {tab === 'report' && <Report data={data} filter={filter} setFilter={setFilter} />}
    {modal && <TransactionForm mode={modal.type} initial={modal.initial} data={data} setData={setData} requireAdmin={requireAdmin} notify={notify} onClose={() => setModal(null)} onSaved={() => setModal(null)} />}
  </div>;
}

export default WarehouseCoreV10;
