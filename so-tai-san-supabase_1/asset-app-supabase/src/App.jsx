/*
  MYHL-QUẢN LÝ TÀI SẢN
  WAREHOUSE CORE V8 — 24/08/2026

  Mục tiêu V8
  ===========
  1. DANH MỤC TÀI SẢN = MASTER DATA.
     Import danh mục chỉ cập nhật định danh tài sản.
     TUYỆT ĐỐI KHÔNG tạo phiếu nhập/xuất và không tạo tồn.

  2. PHIẾU NHẬP / PHIẾU XUẤT = TRANSACTION.
     Chỉ giao dịch kho mới làm phát sinh tồn.

  3. Khi lập phiếu:
     - Gõ tự do tên tài sản hoặc mã hàng.
     - Có danh sách gợi ý autocomplete.
     - Tìm không phân biệt hoa/thường và dấu tiếng Việt.
       Ví dụ: "may xuc", "Máy xúc", "MX-01" đều tìm được.
     - Chọn tài sản => tự lấy mã, loại, nhóm, nguồn gốc, ĐVT.
     - Không phải nhập lại master data.

  4. Tồn được tính theo:
       Tài sản + Kho/Công trình + đến ngày.

  5. XÓA PHIẾU:
     - Admin mới được xóa.
     - Chọn nhiều phiếu rồi xóa hàng loạt.
     - Có "Chọn tất cả kết quả đang lọc".
     - Có thể xóa theo số phiếu.
     - Khi xóa warehouse transaction thì đồng thời xóa transaction lịch sử
       tương ứng, tránh còn "lịch sử giao dịch ma".

  6. Import:
     - Import DANH MỤC => chỉ assets.
     - Import PHIẾU KHO => chỉ warehouse + transactions.
     - Import phiếu kiểm tra toàn bộ trước; có lỗi => không ghi dòng nào.

  7. Ngày:
     - Không dùng new Date("YYYY-MM-DD") để tính ngày nghiệp vụ.
     - Giữ calendar date dạng YYYY-MM-DD.
     - 24/08/2026 luôn là 24/08/2026.

  8. Không tự tạo "Tồn đầu kỳ" từ assets.
     Có trong danh mục không đồng nghĩa có trong kho/công trình.

  Core này dùng các helper đã có trong App.jsx:
    Modal, Field, Btn, Th, Td, Tag, EmptyState, ExportBar,
    inputCls, inputStyle, TOKENS, uid, nowIso, fmtVND,
    XLSX, Search, X, UploadCloud, DownloadCloud.

  ============================================================
  TÍCH HỢP VỚI APP.JSX
  ============================================================
  A. Thay WarehouseTxModal cũ bằng WarehouseTxModalV8.

  B. Handler lưu phiếu:
       const ok = createWarehouseTransactionV8({
         form, data, setData, requireAdmin, notify, logAction
       });
       if (ok) setModal(null);

  C. Danh mục:
       importAssetCatalogV8({
         file, data, setData, requireAdmin, notify, logAction
       });

  D. Phiếu Excel:
       importWarehouseExcelV8({
         file, data, setData, requireAdmin, notify, logAction
       });

  E. Xóa hàng loạt:
       deleteWarehouseTransactionsV8({
         ids: selectedIds,
         data,
         setData,
         requireAdmin,
         notify,
         logAction
       });

  F. Báo cáo:
       const report = makeWarehouseReportV8({
         warehouse: data.warehouse,
         assets: data.assets,
         projects: data.projects,
         filter
       });

  G. QUAN TRỌNG:
       Không giữ handler addWarehouseTx / importWarehouseExcel / WarehouseTxModal
       cũ song song với V8 nếu chúng vẫn được render.
       Chỉ một nguồn logic kho được sử dụng.

  H. KHÔNG:
       - seed lại dữ liệu
       - DROP TABLE
       - DELETE app_data
       - đổi VITE_SUPABASE_URL
       - đổi VITE_SUPABASE_ANON_KEY
*/

/* ============================================================
   1. TEXT / DATE
============================================================ */

function wh8Text(v) {
  return String(v ?? "").trim();
}

function wh8SearchText(v) {
  return wh8Text(v)
    .toLocaleLowerCase("vi-VN")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .normalize("NFC");
}

function wh8Key(v) {
  return wh8SearchText(v);
}

function wh8Date(value) {
  if (value === null || value === undefined || value === "") return "";

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return [
      value.getFullYear(),
      String(value.getMonth() + 1).padStart(2, "0"),
      String(value.getDate()).padStart(2, "0"),
    ].join("-");
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    try {
      const d = XLSX.SSF.parse_date_code(value);
      if (d && d.y && d.m && d.d) {
        return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
      }
    } catch (_) {}
  }

  const s = wh8Text(value);
  if (!s) return "";

  let m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if (m) {
    return `${m[3]}-${String(m[2]).padStart(2, "0")}-${String(m[1]).padStart(2, "0")}`;
  }

  m = s.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})$/);
  if (m) {
    return `${m[1]}-${String(m[2]).padStart(2, "0")}-${String(m[3]).padStart(2, "0")}`;
  }

  m = s.match(/^(\d{4}-\d{2}-\d{2})(?:T|\s)/);
  if (m) return m[1];

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  return "";
}

function wh8DateVN(value) {
  const d = wh8Date(value);
  if (!d) return "";
  const [y, m, day] = d.split("-");
  return `${day}/${m}/${y}`;
}

function wh8IsNhap(tx) {
  const t = wh8Key(tx?.type || tx?.loaiPhieu);
  return t === "nhap" || t.includes("nhap");
}

function wh8IsXuat(tx) {
  const t = wh8Key(tx?.type || tx?.loaiPhieu);
  return t === "xuat" || t.includes("xuat");
}

/* ============================================================
   2. LOCATION
============================================================ */

function wh8Location(tx, projects = []) {
  if (wh8Text(tx?.locationName)) return wh8Text(tx.locationName);

  if (tx?.locationType === "project" && tx?.projectId) {
    const p = projects.find(x => x.id === tx.projectId);
    if (p?.name) return wh8Text(p.name);
  }

  if (wh8Text(tx?.warehouseName)) return wh8Text(tx.warehouseName);

  return "Kho trung tâm";
}

function wh8ProjectByName(projects, name) {
  const k = wh8Key(name);
  return (projects || []).find(p => wh8Key(p.name) === k) || null;
}

/* ============================================================
   3. RESOLVE ASSET
============================================================ */

function resolveWarehouseAssetV8({ row, assets = [] }) {
  const code = wh8Text(
    row?.["Mã hàng"] ||
    row?.["Mã quản lý"] ||
    row?.["Mã tài sản"] ||
    row?.code
  );

  const name = wh8Text(
    row?.["Tên tài sản"] ||
    row?.["Tên hàng"] ||
    row?.["Tên thiết bị"] ||
    row?.name
  );

  const byCode = code
    ? assets.find(a => wh8Key(a.code) === wh8Key(code))
    : null;

  const byName = name
    ? assets.find(a => wh8Key(a.name) === wh8Key(name))
    : null;

  if (byCode && byName && byCode.id !== byName.id) {
    return {
      asset: null,
      error: `Mã "${code}" và Tên "${name}" không cùng một tài sản.`
    };
  }

  const asset = byCode || byName;

  if (!asset) {
    return {
      asset: null,
      error: `Không tìm thấy tài sản theo mã "${code}" hoặc tên "${name}".`
    };
  }

  return { asset, error: "" };
}

function resolveExistingWarehouseAssetV8(tx, assets = []) {
  if (tx?.assetId) {
    const byId = assets.find(a => a.id === tx.assetId);
    if (byId) return byId;
  }

  const code = tx?.itemCode || tx?.code || tx?.maHang || tx?.maQuanLy;
  if (code) {
    const byCode = assets.find(a => wh8Key(a.code) === wh8Key(code));
    if (byCode) return byCode;
  }

  const name = tx?.itemName || tx?.name || tx?.tenTaiSan || tx?.tenHang;
  if (name) {
    const byName = assets.find(a => wh8Key(a.name) === wh8Key(name));
    if (byName) return byName;
  }

  return null;
}

/* ============================================================
   4. SEARCH AUTOCOMPLETE
============================================================ */

function searchWarehouseAssetsV8(assets = [], query = "", limit = 20) {
  const q = wh8Key(query);

  if (!q) return assets.slice(0, limit);

  return assets
    .map(a => {
      const fields = [
        a.code,
        a.name,
        a.serial,
        a.category,
        a.assetGroup,
        a.ownership,
      ].map(wh8SearchText);

      let score = 0;
      if (fields[0] === q) score += 1000;
      if (fields[1] === q) score += 900;
      if (fields[0].startsWith(q)) score += 500;
      if (fields[1].startsWith(q)) score += 450;
      if (fields.some(x => x.includes(q))) score += 100;
      if (fields[0].includes(q)) score += 60;
      if (fields[1].includes(q)) score += 50;

      return { asset: a, score };
    })
    .filter(x => x.score > 0)
    .sort((a, b) => b.score - a.score || wh8Key(a.asset.name).localeCompare(wh8Key(b.asset.name), "vi"))
    .slice(0, limit)
    .map(x => x.asset);
}

/* ============================================================
   5. AUTOCOMPLETE UI
============================================================ */

function WarehouseAssetAutocompleteV8({
  assets = [],
  value = "",
  onChange,
  disabled = false,
}) {
  const selected = assets.find(a => a.id === value) || null;
  const [query, setQuery] = React.useState("");
  const [open, setOpen] = React.useState(false);

  React.useEffect(() => {
    if (selected) {
      setQuery(`${selected.code || ""} — ${selected.name || ""}`);
    } else if (!value) {
      setQuery("");
    }
  }, [value, selected?.id]);

  const results = React.useMemo(
    () => searchWarehouseAssetsV8(assets, query, 25),
    [assets, query]
  );

  const choose = asset => {
    setQuery(`${asset.code || ""} — ${asset.name || ""}`);
    setOpen(false);
    onChange?.(asset.id);
  };

  return (
    <div className="relative">
      <input
        className={inputCls}
        style={inputStyle}
        disabled={disabled}
        value={query}
        autoComplete="off"
        placeholder="Gõ tên hoặc mã tài sản..."
        onFocus={() => setOpen(true)}
        onChange={e => {
          setQuery(e.target.value);
          setOpen(true);
          if (!e.target.value.trim()) onChange?.("");
        }}
        onKeyDown={e => {
          if (e.key === "Escape") setOpen(false);
        }}
      />

      {open && !disabled && (
        <div
          className="absolute z-[100] left-0 right-0 mt-1 max-h-72 overflow-auto rounded-lg shadow-xl"
          style={{
            background: TOKENS.surface,
            border: `1px solid ${TOKENS.border}`,
          }}
          onMouseDown={e => e.preventDefault()}
        >
          {results.length ? (
            results.map(a => (
              <button
                type="button"
                key={a.id}
                className="w-full text-left px-3 py-2.5 hover:bg-black/[0.03]"
                onClick={() => choose(a)}
              >
                <div className="text-[13px] font-medium">
                  {a.code || "—"} — {a.name || "—"}
                </div>
                <div
                  className="text-[11px] mt-0.5"
                  style={{ color: TOKENS.muted }}
                >
                  {a.category || "Chưa phân loại"} ·{" "}
                  {a.assetGroup || "Chưa có nhóm"} ·{" "}
                  {a.ownership || "Chưa có nguồn gốc"}
                </div>
              </button>
            ))
          ) : (
            <div
              className="px-3 py-4 text-[12px]"
              style={{ color: TOKENS.muted }}
            >
              Không tìm thấy tài sản theo mã hoặc tên.
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ============================================================
   6. STOCK ENGINE
============================================================ */

function getWarehouseStockV8({
  warehouse = [],
  assets = [],
  projects = [],
  assetId,
  locationName,
  asOfDate,
}) {
  const cutoff = wh8Date(asOfDate) || "9999-12-31";
  const locKey = wh8Key(locationName);

  return warehouse.reduce((sum, tx) => {
    const date = wh8Date(tx.date || tx.ngayThang);
    if (!date || date > cutoff) return sum;

    const asset = resolveExistingWarehouseAssetV8(tx, assets);
    if (!asset || asset.id !== assetId) return sum;

    if (wh8Key(wh8Location(tx, projects)) !== locKey) return sum;

    const qty = Number(tx.quantity ?? tx.soLuong ?? 0) || 0;

    if (wh8IsNhap(tx)) return sum + qty;
    if (wh8IsXuat(tx)) return sum - qty;
    return sum;
  }, 0);
}

function buildWarehouseBalancesV8({
  warehouse = [],
  assets = [],
  projects = [],
  asOfDate,
}) {
  const cutoff = wh8Date(asOfDate) || "9999-12-31";
  const balances = {};

  warehouse.forEach(tx => {
    const date = wh8Date(tx.date || tx.ngayThang);
    if (!date || date > cutoff) return;

    const asset = resolveExistingWarehouseAssetV8(tx, assets);
    if (!asset) return;

    const location = wh8Location(tx, projects);
    const key = `${asset.id}¦${wh8Key(location)}`;

    if (!balances[key]) {
      balances[key] = {
        assetId: asset.id,
        code: asset.code || tx.itemCode || "",
        name: asset.name || tx.itemName || "",
        location,
        category: asset.category || tx.category || "Khác",
        group: asset.assetGroup || tx.assetGroup || "Thiết bị chính",
        ownership: asset.ownership || tx.ownership || "Công ty",
        unit: asset.unit || tx.unit || "Cái",
        inQty: 0,
        outQty: 0,
        inValue: 0,
        outValue: 0,
      };
    }

    const qty = Number(tx.quantity ?? tx.soLuong ?? 0) || 0;
    const unitCost = Number(tx.unitCost ?? tx.unitPrice ?? tx.donGia ?? 0) || 0;
    const rawTotal = Number(tx.total ?? tx.thanhTien);
    const total = Number.isFinite(rawTotal) ? rawTotal : qty * unitCost;

    if (wh8IsNhap(tx)) {
      balances[key].inQty += qty;
      balances[key].inValue += total;
    } else if (wh8IsXuat(tx)) {
      balances[key].outQty += qty;
      balances[key].outValue += total;
    }
  });

  return Object.values(balances).map(r => ({
    ...r,
    balanceQty: r.inQty - r.outQty,
    balanceValue: r.inValue - r.outValue,
  }));
}

function makeWarehouseReportV8({
  warehouse = [],
  assets = [],
  projects = [],
  filter = {},
}) {
  const all = buildWarehouseBalancesV8({
    warehouse,
    assets,
    projects,
    asOfDate: filter.asOfDate,
  });

  const rows = all
    .filter(r => !filter.category || wh8Key(r.category) === wh8Key(filter.category))
    .filter(r => !filter.group || wh8Key(r.group) === wh8Key(filter.group))
    .filter(r => !filter.ownership || wh8Key(r.ownership) === wh8Key(filter.ownership))
    .filter(r => !filter.locationName || wh8Key(r.location) === wh8Key(filter.locationName))
    .filter(r => {
      if (!filter.projectId) return true;
      const p = projects.find(x => x.id === filter.projectId);
      return !!p && wh8Key(p.name) === wh8Key(r.location);
    })
    .filter(r => r.balanceQty > 0)
    .sort((a, b) =>
      `${a.location}|${a.group}|${a.category}|${a.code}|${a.name}`
        .localeCompare(
          `${b.location}|${b.group}|${b.category}|${b.code}|${b.name}`,
          "vi"
        )
    );

  return {
    rows,
    summary: {
      qty: rows.reduce((s, r) => s + r.balanceQty, 0),
      value: rows.reduce((s, r) => s + r.balanceValue, 0),
      assets: rows.length,
      locations: new Set(rows.map(r => r.location)).size,
    },
    headers: [
      "Kho/Công trình",
      "Loại tài sản",
      "Nhóm tài sản",
      "Nguồn gốc",
      "Mã hàng",
      "Tên tài sản",
      "ĐVT",
      "Nhập lũy kế",
      "Xuất lũy kế",
      "Tồn đến ngày",
      "Giá trị tồn",
    ],
    values: rows.map(r => [
      r.location,
      r.category,
      r.group,
      r.ownership,
      r.code,
      r.name,
      r.unit,
      r.inQty,
      r.outQty,
      r.balanceQty,
      r.balanceValue,
    ]),
    asOfDate: wh8Date(filter.asOfDate),
    asOfDateVN: wh8DateVN(filter.asOfDate),
  };
}

/* ============================================================
   7. CREATE TRANSACTION
============================================================ */

function createWarehouseTransactionV8({
  form,
  data,
  setData,
  requireAdmin,
  notify,
  logAction,
}) {
  if (requireAdmin && !requireAdmin()) return false;

  const assets = data.assets || [];
  const projects = data.projects || [];
  const warehouse = data.warehouse || [];

  const asset = assets.find(a => a.id === form.assetId);
  if (!asset) {
    notify("Vui lòng chọn tài sản từ danh sách gợi ý.");
    return false;
  }

  const date = wh8Date(form.date);
  if (!date) {
    notify("Ngày chứng từ không hợp lệ.");
    return false;
  }

  const qty = Number(form.quantity);
  if (!(qty > 0)) {
    notify("Số lượng phải lớn hơn 0.");
    return false;
  }

  const location =
    form.locationType === "project"
      ? projects.find(p => p.id === form.projectId)?.name || ""
      : wh8Text(form.warehouseName);

  if (!location) {
    notify("Vui lòng chọn Kho/Công trình.");
    return false;
  }

  const isNhap = form.type === "nhap";
  const prefix = isNhap ? "PN" : "PX";
  const dateKey = date.replaceAll("-", "");

  const sameDay = warehouse.filter(w =>
    String(w.voucherNo || "").startsWith(`${prefix}-${dateKey}-`)
  );

  const voucherNo =
    wh8Text(form.voucherNo) ||
    `${prefix}-${dateKey}-${String(sameDay.length + 1).padStart(3, "0")}`;

  const duplicateVoucher = warehouse.some(
    w => wh8Key(w.voucherNo) === wh8Key(voucherNo)
  );

  if (duplicateVoucher) {
    notify(`Số phiếu ${voucherNo} đã tồn tại.`);
    return false;
  }

  const unit = asset.unit || form.unit || "Cái";
  const unitCost = Number(form.unitCost || 0) || 0;

  if (!isNhap) {
    const stock = getWarehouseStockV8({
      warehouse,
      assets,
      projects,
      assetId: asset.id,
      locationName: location,
      asOfDate: date,
    });

    if (qty > stock) {
      notify(
        `Không đủ tồn tại "${location}". ${asset.code} — ${asset.name} hiện còn ${stock} ${unit}.`
      );
      return false;
    }
  }

  const tx = {
    id: uid("wh"),
    voucherNo,
    assetId: asset.id,
    type: isNhap ? "nhap" : "xuat",
    quantity: qty,
    date,
    unitCost,
    total: qty * unitCost,
    unit,
    receiver: wh8Text(form.receiver),
    note: wh8Text(form.note),
    category: asset.category || "Khác",
    assetGroup: asset.assetGroup || "Thiết bị chính",
    ownership: asset.ownership || "Công ty",
    locationType: form.locationType || "project",
    locationName: location,
    warehouseName: form.locationType === "warehouse" ? location : "",
    projectId: form.locationType === "project" ? form.projectId || null : null,
    itemName: asset.name,
    itemCode: asset.code,
    serial: asset.serial || "",
  };

  const historyTx = {
    id: uid("tx"),
    warehouseTxId: tx.id,
    assetId: asset.id,
    type: isNhap ? "nhap_kho" : "xuat_kho",
    date,
    title: `${isNhap ? "Nhập kho" : "Xuất kho"} ${voucherNo}`,
    detail: `${asset.name} · ${location} · ${tx.receiver || ""}`,
    amount: tx.total,
  };

  setData({
    ...data,
    warehouse: [tx, ...warehouse],
    transactions: [historyTx, ...(data.transactions || [])],
    activityLog:
      typeof logAction === "function"
        ? logAction(
            data.activityLog,
            `${isNhap ? "Nhập" : "Xuất"} ${voucherNo} — ${asset.code} — ${location}`
          )
        : data.activityLog,
  });

  notify(`Đã lập ${isNhap ? "phiếu nhập" : "phiếu xuất"} ${voucherNo}`);
  return true;
}

/* ============================================================
   8. MODAL PHIẾU NHẬP / XUẤT V8
============================================================ */

function WarehouseTxModalV8({
  assets = [],
  projects = [],
  onClose,
  onSubmit,
  title,
  fixedType = "nhap",
}) {
  const [f, setF] = React.useState({
    assetId: "",
    voucherNo: "",
    quantity: 1,
    date: wh8Date(nowIso()),
    unit: "Cái",
    unitCost: 0,
    receiver: "",
    locationType: "project",
    warehouseName: "",
    projectId: "",
    note: "",
    type: fixedType,
  });

  const asset = assets.find(a => a.id === f.assetId) || {};

  const set = (key, value) =>
    setF(prev => ({ ...prev, [key]: value }));

  const selectAsset = id => {
    const a = assets.find(x => x.id === id);

    setF(prev => ({
      ...prev,
      assetId: id,
      unit: a?.unit || "Cái",
      unitCost: Number(a?.cost || 0),
    }));
  };

  const locationName =
    f.locationType === "project"
      ? projects.find(p => p.id === f.projectId)?.name || ""
      : wh8Text(f.warehouseName);

  const total =
    (Number(f.quantity) || 0) *
    (Number(f.unitCost) || 0);

  const valid =
    !!f.assetId &&
    !!wh8Date(f.date) &&
    Number(f.quantity) > 0 &&
    !!locationName;

  return (
    <Modal
      title={title || (fixedType === "xuat" ? "Lập phiếu xuất kho" : "Lập phiếu nhập kho")}
      onClose={onClose}
      wide
    >
      <div
        className="rounded-lg p-3 mb-4"
        style={{
          background: TOKENS.brandSoft,
          border: `1px solid ${TOKENS.brand}22`,
        }}
      >
        <div
          className="text-[12px] font-semibold"
          style={{ color: TOKENS.brand }}
        >
          {fixedType === "xuat" ? "PHIẾU XUẤT KHO" : "PHIẾU NHẬP KHO"}
        </div>
        <div
          className="text-[11px] mt-1"
          style={{ color: TOKENS.muted }}
        >
          Gõ tự do mã hoặc tên tài sản để tìm. Thông tin loại, nhóm,
          nguồn gốc và đơn vị tính lấy từ Danh mục tài sản.
        </div>
      </div>

      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Số phiếu">
          <input
            className={inputCls}
            style={inputStyle}
            value={f.voucherNo}
            onChange={e => set("voucherNo", e.target.value)}
            placeholder={fixedType === "xuat" ? "PX-20260824-001" : "PN-20260824-001"}
          />
        </Field>

        <Field label="Ngày tháng">
          <input
            type="date"
            className={inputCls}
            style={inputStyle}
            value={wh8Date(f.date)}
            onChange={e => set("date", e.target.value)}
          />
        </Field>

        <Field label="Tên tài sản / Mã hàng">
          <WarehouseAssetAutocompleteV8
            assets={assets}
            value={f.assetId}
            onChange={selectAsset}
          />
        </Field>

        <Field label="Mã hàng">
          <input
            readOnly
            className={inputCls}
            style={{ ...inputStyle, background: TOKENS.paper }}
            value={asset.code || ""}
          />
        </Field>

        <Field label="Loại tài sản">
          <input
            readOnly
            className={inputCls}
            style={{ ...inputStyle, background: TOKENS.paper }}
            value={asset.category || ""}
          />
        </Field>

        <Field label="Nhóm tài sản">
          <input
            readOnly
            className={inputCls}
            style={{ ...inputStyle, background: TOKENS.paper }}
            value={asset.assetGroup || ""}
          />
        </Field>

        <Field label="Nguồn gốc">
          <input
            readOnly
            className={inputCls}
            style={{ ...inputStyle, background: TOKENS.paper }}
            value={asset.ownership || ""}
          />
        </Field>

        <Field label="Đơn vị tính">
          <input
            readOnly
            className={inputCls}
            style={{ ...inputStyle, background: TOKENS.paper }}
            value={asset.unit || "Cái"}
          />
        </Field>

        <Field label="Loại địa điểm">
          <select
            className={inputCls}
            style={inputStyle}
            value={f.locationType}
            onChange={e => set("locationType", e.target.value)}
          >
            <option value="project">Công trình</option>
            <option value="warehouse">Kho</option>
          </select>
        </Field>

        {f.locationType === "project" ? (
          <Field label="Kho / Công trình">
            <select
              className={inputCls}
              style={inputStyle}
              value={f.projectId}
              onChange={e => set("projectId", e.target.value)}
            >
              <option value="">-- Chọn công trình --</option>
              {projects.map(p => (
                <option key={p.id} value={p.id}>
                  {p.name}
                </option>
              ))}
            </select>
          </Field>
        ) : (
          <Field label="Tên kho">
            <input
              className={inputCls}
              style={inputStyle}
              value={f.warehouseName}
              onChange={e => set("warehouseName", e.target.value)}
              placeholder="Kho trung tâm"
            />
          </Field>
        )}

        <Field label="Người giao / nhận">
          <input
            className={inputCls}
            style={inputStyle}
            value={f.receiver}
            onChange={e => set("receiver", e.target.value)}
          />
        </Field>

        <Field label="Số lượng">
          <input
            type="number"
            min="0"
            step="any"
            className={inputCls}
            style={inputStyle}
            value={f.quantity}
            onChange={e => set("quantity", e.target.value)}
          />
        </Field>

        <Field label="Đơn giá">
          <input
            type="number"
            min="0"
            step="any"
            className={inputCls}
            style={inputStyle}
            value={f.unitCost}
            onChange={e => set("unitCost", e.target.value)}
          />
        </Field>

        <Field label="Thành tiền">
          <input
            readOnly
            className={inputCls}
            style={{ ...inputStyle, background: TOKENS.paper }}
            value={typeof fmtVND === "function" ? fmtVND(total) : total.toLocaleString("vi-VN")}
          />
        </Field>

        <div className="col-span-2">
          <Field label="Ghi chú">
            <textarea
              className={inputCls}
              style={inputStyle}
              rows={3}
              value={f.note}
              onChange={e => set("note", e.target.value)}
            />
          </Field>
        </div>
      </div>

      <div className="flex justify-end gap-2 mt-4">
        <Btn onClick={onClose}>Hủy</Btn>
        <Btn
          kind="primary"
          disabled={!valid}
          onClick={() =>
            onSubmit?.({
              ...f,
              type: fixedType,
              date: wh8Date(f.date),
              unit: asset.unit || "Cái",
              category: asset.category || "",
              assetGroup: asset.assetGroup || "",
              ownership: asset.ownership || "",
            })
          }
        >
          Lưu phiếu
        </Btn>
      </div>
    </Modal>
  );
}

/* ============================================================
   9. BULK DELETE — XÓA PHIẾU NHẬP / XUẤT
============================================================ */

function deleteWarehouseTransactionsV8({
  ids = [],
  data,
  setData,
  requireAdmin,
  notify,
  logAction,
}) {
  if (requireAdmin && !requireAdmin()) return false;

  const selected = new Set(ids.filter(Boolean));
  if (!selected.size) {
    notify("Chưa chọn phiếu cần xóa.");
    return false;
  }

  const warehouse = data.warehouse || [];
  const rows = warehouse.filter(w => selected.has(w.id));

  if (!rows.length) {
    notify("Không tìm thấy các phiếu đã chọn.");
    return false;
  }

  const voucherNos = new Set(rows.map(w => wh8Key(w.voucherNo)).filter(Boolean));
  const warehouseIds = new Set(rows.map(w => w.id));

  const confirmed = window.confirm(
    `Bạn có chắc muốn xóa ${rows.length} dòng giao dịch thuộc ${voucherNos.size} số phiếu?\n\n` +
    [...voucherNos].join(", ") +
    `\n\nThao tác này sẽ làm thay đổi báo cáo nhập - xuất - tồn.`
  );

  if (!confirmed) return false;

  const remainingWarehouse = warehouse.filter(w => !selected.has(w.id));

  const remainingTransactions = (data.transactions || []).filter(t => {
    if (t.warehouseTxId && warehouseIds.has(t.warehouseTxId)) return false;

    const title = wh8Key(t.title);
    const matchesVoucher = [...voucherNos].some(v => v && title.includes(v));
    return !matchesVoucher;
  });

  setData({
    ...data,
    warehouse: remainingWarehouse,
    transactions: remainingTransactions,
    activityLog:
      typeof logAction === "function"
        ? logAction(
            data.activityLog,
            `Xóa ${rows.length} dòng phiếu kho (${[...voucherNos].join(", ")})`
          )
        : data.activityLog,
  });

  notify(`Đã xóa ${rows.length} dòng giao dịch kho.`);
  return true;
}

function deleteWarehouseVouchersV8({
  voucherNos = [],
  data,
  setData,
  requireAdmin,
  notify,
  logAction,
}) {
  if (requireAdmin && !requireAdmin()) return false;

  const keys = new Set(voucherNos.map(wh8Key).filter(Boolean));
  if (!keys.size) {
    notify("Chưa chọn số phiếu cần xóa.");
    return false;
  }

  const warehouse = data.warehouse || [];
  const rows = warehouse.filter(w => keys.has(wh8Key(w.voucherNo)));

  if (!rows.length) {
    notify("Không tìm thấy phiếu cần xóa.");
    return false;
  }

  const ok = window.confirm(
    `Xóa ${keys.size} phiếu (${rows.length} dòng giao dịch)?\n\n` +
    `Các phiếu: ${voucherNos.join(", ")}`
  );

  if (!ok) return false;

  const rowIds = new Set(rows.map(w => w.id));

  setData({
    ...data,
    warehouse: warehouse.filter(w => !keys.has(wh8Key(w.voucherNo))),
    transactions: (data.transactions || []).filter(t => {
      if (t.warehouseTxId && rowIds.has(t.warehouseTxId)) return false;
      const title = wh8Key(t.title);
      return ![...keys].some(k => k && title.includes(k));
    }),
    activityLog:
      typeof logAction === "function"
        ? logAction(
            data.activityLog,
            `Xóa hàng loạt ${keys.size} phiếu kho`
          )
        : data.activityLog,
  });

  notify(`Đã xóa ${keys.size} phiếu kho.`);
  return true;
}

/* ============================================================
   10. TABLE PHIẾU + SEARCH + BULK SELECT
============================================================ */

function WarehouseTransactionTableV8({
  warehouse = [],
  assets = [],
  projects = [],
  type = "",
  requireAdmin,
  onDeleteSelected,
}) {
  const [query, setQuery] = React.useState("");
  const [locationFilter, setLocationFilter] = React.useState("");
  const [selectedIds, setSelectedIds] = React.useState([]);

  const filtered = React.useMemo(() => {
    const q = wh8Key(query);

    return warehouse
      .filter(w => {
        if (type && !(
          (type === "nhap" && wh8IsNhap(w)) ||
          (type === "xuat" && wh8IsXuat(w))
        )) return false;

        if (locationFilter && wh8Key(wh8Location(w, projects)) !== wh8Key(locationFilter)) {
          return false;
        }

        if (!q) return true;

        const asset = resolveExistingWarehouseAssetV8(w, assets);

        const haystack = [
          w.voucherNo,
          w.itemCode,
          w.itemName,
          asset?.code,
          asset?.name,
          wh8Location(w, projects),
          w.receiver,
        ].map(wh8SearchText).join(" ");

        return haystack.includes(q);
      })
      .sort((a, b) => {
        const da = wh8Date(a.date);
        const db = wh8Date(b.date);
        return db.localeCompare(da) || String(b.id).localeCompare(String(a.id));
      });
  }, [warehouse, assets, projects, type, query, locationFilter]);

  const filteredIds = filtered.map(w => w.id);
  const allFilteredSelected =
    filteredIds.length > 0 &&
    filteredIds.every(id => selectedIds.includes(id));

  const toggleAll = () => {
    if (allFilteredSelected) {
      setSelectedIds(prev => prev.filter(id => !filteredIds.includes(id)));
    } else {
      setSelectedIds(prev => Array.from(new Set([...prev, ...filteredIds])));
    }
  };

  const toggleOne = id => {
    setSelectedIds(prev =>
      prev.includes(id)
        ? prev.filter(x => x !== id)
        : [...prev, id]
    );
  };

  const locations = Array.from(
    new Set(warehouse.map(w => wh8Location(w, projects)))
  ).sort((a, b) => wh8Key(a).localeCompare(wh8Key(b), "vi"));

  const doDelete = () => {
    if (!selectedIds.length) return;
    onDeleteSelected?.(selectedIds);
    setSelectedIds([]);
  };

  return (
    <div>
      <div className="flex flex-wrap gap-2 items-center mb-3">
        <div className="flex-1 min-w-[260px]">
          <input
            className={inputCls}
            style={inputStyle}
            value={query}
            onChange={e => setQuery(e.target.value)}
            placeholder="Tìm số phiếu, mã hàng, tên tài sản, công trình..."
          />
        </div>

        <select
          className={inputCls}
          style={{ ...inputStyle, width: 220 }}
          value={locationFilter}
          onChange={e => setLocationFilter(e.target.value)}
        >
          <option value="">Tất cả Kho/Công trình</option>
          {locations.map(x => (
            <option key={x} value={x}>{x}</option>
          ))}
        </select>

        <Btn
          kind="danger"
          disabled={!selectedIds.length}
          onClick={doDelete}
        >
          Xóa {selectedIds.length ? `(${selectedIds.length})` : ""} phiếu
        </Btn>
      </div>

      <div
        className="overflow-auto rounded-lg"
        style={{
          background: TOKENS.surface,
          border: `1px solid ${TOKENS.border}`,
        }}
      >
        <table className="w-full min-w-[1150px]">
          <thead>
            <tr>
              <Th>
                <input
                  type="checkbox"
                  checked={allFilteredSelected}
                  onChange={toggleAll}
                  aria-label="Chọn tất cả kết quả"
                />
              </Th>
              <Th>Số phiếu</Th>
              <Th>Ngày</Th>
              <Th>Mã hàng</Th>
              <Th>Tên tài sản</Th>
              <Th>Kho/Công trình</Th>
              <Th>Loại</Th>
              <Th right>Số lượng</Th>
              <Th>ĐVT</Th>
              <Th>Người giao/nhận</Th>
            </tr>
          </thead>
          <tbody>
            {filtered.map(w => {
              const asset = resolveExistingWarehouseAssetV8(w, assets);
              const checked = selectedIds.includes(w.id);

              return (
                <tr key={w.id} className="aa-row">
                  <Td>
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={() => toggleOne(w.id)}
                    />
                  </Td>
                  <Td mono>
                    <Tag>{w.voucherNo || "—"}</Tag>
                  </Td>
                  <Td>{wh8DateVN(w.date)}</Td>
                  <Td mono>{asset?.code || w.itemCode || "—"}</Td>
                  <Td>{asset?.name || w.itemName || "—"}</Td>
                  <Td>{wh8Location(w, projects)}</Td>
                  <Td>
                    {wh8IsNhap(w) ? "Nhập kho" : "Xuất kho"}
                  </Td>
                  <Td right mono>
                    {Number(w.quantity || 0).toLocaleString("vi-VN")}
                  </Td>
                  <Td>{w.unit || asset?.unit || "Cái"}</Td>
                  <Td>{w.receiver || "—"}</Td>
                </tr>
              );
            })}

            {!filtered.length && (
              <tr>
                <td colSpan={10} className="py-12 text-center">
                  <EmptyState text="Không có phiếu phù hợp điều kiện lọc." />
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <div
        className="text-[12px] mt-2"
        style={{ color: TOKENS.muted }}
      >
        Đang hiển thị {filtered.length} dòng · Đã chọn {selectedIds.length} dòng.
        Tích ô đầu bảng để chọn toàn bộ kết quả đang lọc.
      </div>
    </div>
  );
}

/* ============================================================
   11. IMPORT DANH MỤC — KHÔNG TẠO PHIẾU
============================================================ */

function importAssetCatalogV8({
  file,
  data,
  setData,
  requireAdmin,
  notify,
  logAction,
}) {
  if (requireAdmin && !requireAdmin()) return;

  const reader = new FileReader();

  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, {
        type: "array",
        cellDates: true,
      });

      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      if (!rows.length) {
        notify("File danh mục không có dữ liệu.");
        return;
      }

      const existing = data.assets || [];
      const existingByCode = Object.fromEntries(
        existing.map(a => [wh8Key(a.code), a])
      );

      const imported = [];
      const errors = [];
      const seen = new Set();

      rows.forEach((r, i) => {
        const rowNo = i + 2;

        const code = wh8Text(
          r["Mã quản lý"] ||
          r["Mã hàng"] ||
          r["Mã tài sản"] ||
          r.code
        );

        const name = wh8Text(
          r["Tên tài sản"] ||
          r["Tên hàng"] ||
          r["Tên thiết bị"] ||
          r.name
        );

        if (!code || !name) {
          errors.push(`Dòng ${rowNo}: thiếu Mã quản lý hoặc Tên tài sản.`);
          return;
        }

        const key = wh8Key(code);

        if (seen.has(key)) {
          errors.push(`Dòng ${rowNo}: mã ${code} bị trùng trong file.`);
          return;
        }

        seen.add(key);
        const old = existingByCode[key];

        imported.push({
          ...old,
          id: old?.id || uid("as"),
          code,
          name,
          category: wh8Text(r["Loại tài sản"] || r["Loại"] || old?.category || ""),
          assetGroup: wh8Text(r["Nhóm tài sản"] || r["Nhóm"] || old?.assetGroup || ""),
          ownership: wh8Text(r["Nguồn gốc"] || old?.ownership || "Công ty"),
          unit: wh8Text(r["Đơn vị tính"] || r["ĐVT"] || old?.unit || "Cái"),
          serial: wh8Text(r["Serial"] || r["Số serial"] || old?.serial || ""),
          cost: Number(r["Đơn giá"] || r["Nguyên giá"] || old?.cost || 0) || 0,
          usefulLifeMonths: Number(
            r["Thời gian sử dụng"] || old?.usefulLifeMonths || 36
          ) || 36,
        });
      });

      if (!imported.length) {
        notify(errors[0] || "Không có dòng danh mục hợp lệ.");
        return;
      }

      const importedCodes = new Set(imported.map(a => wh8Key(a.code)));
      const remaining = existing.filter(
        a => !importedCodes.has(wh8Key(a.code))
      );

      setData({
        ...data,
        assets: [...imported, ...remaining],
        warehouse: Array.isArray(data.warehouse) ? data.warehouse : [],
        transactions: Array.isArray(data.transactions) ? data.transactions : [],
        activityLog:
          typeof logAction === "function"
            ? logAction(
                data.activityLog,
                `Import/cập nhật ${imported.length} danh mục tài sản — KHÔNG tạo phiếu kho`
              )
            : data.activityLog,
      });

      notify(
        errors.length
          ? `Đã cập nhật ${imported.length} danh mục; ${errors.length} dòng lỗi. ${errors[0]}`
          : `Đã cập nhật ${imported.length} danh mục tài sản. Không tạo phiếu kho.`
      );
    } catch (err) {
      console.error(err);
      notify("Không đọc được file Excel danh mục tài sản.");
    }
  };

  reader.readAsArrayBuffer(file);
}

/* ============================================================
   12. IMPORT PHIẾU KHO
============================================================ */

function importWarehouseExcelV8({
  file,
  data,
  setData,
  requireAdmin,
  notify,
  logAction,
}) {
  if (requireAdmin && !requireAdmin()) return;

  const reader = new FileReader();

  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, {
        type: "array",
        cellDates: true,
      });

      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      if (!rows.length) {
        notify("File Excel phiếu kho không có dữ liệu.");
        return;
      }

      const assets = data.assets || [];
      const projects = data.projects || [];
      const existing = data.warehouse || [];

      const imported = [];
      const errors = [];

      rows.forEach((r, i) => {
        const rowNo = i + 2;

        const rawType = wh8Key(
          r["Loại phiếu"] ||
          r["Loại chứng từ"] ||
          "Nhập kho"
        );

        const type = rawType.includes("xuat") ? "xuat" : "nhap";

        const date = wh8Date(
          r["Ngày tháng"] ||
          r["Ngày chứng từ"] ||
          r["Ngày"]
        );

        if (!date) {
          errors.push(`Dòng ${rowNo}: ngày không hợp lệ.`);
          return;
        }

        const resolved = resolveWarehouseAssetV8({
          row: r,
          assets,
        });

        if (resolved.error) {
          errors.push(`Dòng ${rowNo}: ${resolved.error}`);
          return;
        }

        const asset = resolved.asset;

        const qty = Number(
          r["Số lượng"] ??
          r["SL"] ??
          0
        );

        if (!(qty > 0)) {
          errors.push(`Dòng ${rowNo}: số lượng phải lớn hơn 0.`);
          return;
        }

        const locationText = wh8Text(
          r["Kho/Công trình"] ||
          r["Công trình"] ||
          r["Tên kho"] ||
          r["Kho"]
        );

        if (!locationText) {
          errors.push(`Dòng ${rowNo}: thiếu Kho/Công trình.`);
          return;
        }

        const project = wh8ProjectByName(projects, locationText);
        const rawLocationType = wh8Key(
          r["Loại địa điểm"] ||
          r["Loại nơi"] ||
          ""
        );

        const locationType =
          rawLocationType.includes("cong trinh")
            ? "project"
            : rawLocationType.includes("kho")
              ? "warehouse"
              : project
                ? "project"
                : "warehouse";

        const unitCost = Number(
          r["Đơn giá"] ??
          r["Đơn giá nhập"] ??
          r["Đơn giá xuất"] ??
          0
        ) || 0;

        const prefix = type === "nhap" ? "PN" : "PX";
        const dateKey = date.replaceAll("-", "");

        const seq =
          [...existing, ...imported].filter(w =>
            String(w.voucherNo || "").startsWith(`${prefix}-${dateKey}-`)
          ).length + 1;

        const voucherNo =
          wh8Text(r["Số phiếu"]) ||
          `${prefix}-${dateKey}-${String(seq).padStart(3, "0")}`;

        const unit =
          wh8Text(r["Đơn vị tính"] || r["ĐVT"]) ||
          asset.unit ||
          "Cái";

        imported.push({
          id: uid("wh"),
          voucherNo,
          assetId: asset.id,
          type,
          quantity: qty,
          date,
          unitCost,
          total: qty * unitCost,
          unit,
          receiver: wh8Text(r["Người giao/nhận"]),
          note: wh8Text(r["Ghi chú"]),
          category: asset.category || "Khác",
          assetGroup: asset.assetGroup || "Thiết bị chính",
          ownership: asset.ownership || "Công ty",
          locationType,
          locationName: locationText,
          warehouseName: locationType === "warehouse" ? locationText : "",
          projectId: locationType === "project" ? project?.id || null : null,
          itemName: asset.name,
          itemCode: asset.code,
          serial: asset.serial || "",
        });
      });

      const existingVoucherKeys = new Set(
        existing.map(w => wh8Key(w.voucherNo))
      );

      imported.forEach(tx => {
        if (existingVoucherKeys.has(wh8Key(tx.voucherNo))) {
          errors.push(
            `Số phiếu ${tx.voucherNo} đã tồn tại. Không import trùng phiếu.`
          );
        }
      });

      /*
        Kiểm tra xuất:
        - theo đúng tài sản
        - đúng Kho/Công trình
        - đúng ngày
        - tính cả các dòng import trước đó
      */
      for (const tx of imported) {
        if (!wh8IsXuat(tx)) continue;

        const rowsAtLocation = [
          ...existing,
          ...imported,
        ].filter(w => {
          const a = resolveExistingWarehouseAssetV8(w, assets);
          return (
            a?.id === tx.assetId &&
            wh8Key(wh8Location(w, projects)) === wh8Key(tx.locationName) &&
            wh8Date(w.date) <= tx.date
          );
        });

        const available = rowsAtLocation.reduce((sum, w) => {
          if (wh8IsNhap(w)) return sum + (Number(w.quantity) || 0);
          if (wh8IsXuat(w)) return sum - (Number(w.quantity) || 0);
          return sum;
        }, 0);

        if (tx.quantity > available) {
          errors.push(
            `${tx.voucherNo}: xuất ${tx.quantity} ${tx.unit} tại "${tx.locationName}" vượt tồn ${available}.`
          );
        }
      }

      if (errors.length) {
        notify(`Import bị dừng: ${errors.length} lỗi. ${errors[0]}`);
        return;
      }

      const history = imported.map(tx => ({
        id: uid("tx"),
        warehouseTxId: tx.id,
        assetId: tx.assetId,
        type: tx.type === "nhap" ? "nhap_kho" : "xuat_kho",
        date: tx.date,
        title: `${tx.type === "nhap" ? "Nhập kho" : "Xuất kho"} ${tx.voucherNo}`,
        detail: `${tx.itemName} · ${tx.locationName} · ${tx.receiver || ""}`,
        amount: tx.total,
      }));

      setData({
        ...data,
        warehouse: [...imported, ...existing],
        transactions: [...history, ...(data.transactions || [])],
        activityLog:
          typeof logAction === "function"
            ? logAction(
                data.activityLog,
                `Import ${imported.length} phiếu nhập/xuất kho`
              )
            : data.activityLog,
      });

      notify(`Đã import ${imported.length} phiếu kho.`);
    } catch (err) {
      console.error(err);
      notify("Không đọc được file Excel phiếu nhập/xuất kho.");
    }
  };

  reader.readAsArrayBuffer(file);
}

/* ============================================================
   13. TEMPLATE EXCEL
============================================================ */

function downloadWarehouseTemplateV8({
  projects = [],
  downloadExcelTemplate,
}) {
  const headers = [
    "Loại phiếu",
    "Số phiếu",
    "Ngày tháng",
    "Tên tài sản",
    "Mã hàng",
    "Kho/Công trình",
    "Loại địa điểm",
    "Loại tài sản",
    "Nhóm tài sản",
    "Nguồn gốc",
    "Người giao/nhận",
    "Số lượng",
    "Đơn vị tính",
    "Đơn giá",
    "Thành tiền",
    "Ghi chú",
  ];

  const p = projects[0]?.name || "Kho trung tâm";

  const rows = [[
    "Nhập kho",
    "PN-20260824-001",
    "2026-08-24",
    "Máy xúc 01",
    "MX-01",
    p,
    projects.length ? "Công trình" : "Kho",
    "Máy xúc",
    "Thiết bị chính",
    "Thuê",
    "Nguyễn Văn A",
    1,
    "Cái",
    0,
    0,
    "",
  ]];

  if (typeof downloadExcelTemplate === "function") {
    downloadExcelTemplate(
      "Mau_Import_Phieu_Nhap_Xuat_Kho_V8",
      headers,
      rows
    );
    return;
  }

  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Phiếu kho");
  XLSX.writeFile(wb, "Mau_Import_Phieu_Nhap_Xuat_Kho_V8.xlsx");
}

/* ============================================================
   14. AUDIT DATA
============================================================ */

function auditWarehouseDataV8({
  warehouse = [],
  assets = [],
  projects = [],
}) {
  const orphanTransactions = [];
  const invalidDates = [];
  const invalidLocations = [];

  warehouse.forEach(tx => {
    if (!resolveExistingWarehouseAssetV8(tx, assets)) {
      orphanTransactions.push({
        id: tx.id,
        voucherNo: tx.voucherNo,
        itemCode: tx.itemCode,
        itemName: tx.itemName,
      });
    }

    if (!wh8Date(tx.date)) {
      invalidDates.push({
        id: tx.id,
        voucherNo: tx.voucherNo,
        date: tx.date,
      });
    }

    if (!wh8Location(tx, projects)) {
      invalidLocations.push(tx.voucherNo);
    }
  });

  const balances = buildWarehouseBalancesV8({
    warehouse,
    assets,
    projects,
    asOfDate: "9999-12-31",
  });

  return {
    totalAssets: assets.length,
    totalWarehouseTransactions: warehouse.length,
    orphanTransactions,
    invalidDates,
    invalidLocations,
    negativeLocations: balances
      .filter(r => r.balanceQty < 0)
      .map(r => ({
        code: r.code,
        name: r.name,
        location: r.location,
        balanceQty: r.balanceQty,
      })),
  };
}

/* ============================================================
   15. INTEGRATION EXAMPLE
============================================================ */

/*
  PHIẾU NHẬP:

  {modal?.type === "warehouseIn" && (
    <WarehouseTxModalV8
      title="Lập phiếu nhập kho"
      fixedType="nhap"
      assets={data.assets}
      projects={data.projects}
      onClose={() => setModal(null)}
      onSubmit={form => {
        const ok = createWarehouseTransactionV8({
          form,
          data,
          setData,
          requireAdmin,
          notify,
          logAction,
        });
        if (ok) setModal(null);
      }}
    />
  )}

  PHIẾU XUẤT:

  {modal?.type === "warehouseOut" && (
    <WarehouseTxModalV8
      title="Lập phiếu xuất kho"
      fixedType="xuat"
      assets={data.assets}
      projects={data.projects}
      onClose={() => setModal(null)}
      onSubmit={form => {
        const ok = createWarehouseTransactionV8({
          form,
          data,
          setData,
          requireAdmin,
          notify,
          logAction,
        });
        if (ok) setModal(null);
      }}
    />
  )}

  TAB PHIẾU NHẬP:

  <WarehouseTransactionTableV8
    warehouse={data.warehouse}
    assets={data.assets}
    projects={data.projects}
    type="nhap"
    requireAdmin={requireAdmin}
    onDeleteSelected={ids =>
      deleteWarehouseTransactionsV8({
        ids,
        data,
        setData,
        requireAdmin,
        notify,
        logAction,
      })
    }
  />

  TAB PHIẾU XUẤT:
  tương tự nhưng type="xuat".

  IMPORT DANH MỤC:
  importAssetCatalogV8(...)

  IMPORT PHIẾU:
  importWarehouseExcelV8(...)

  BÁO CÁO:
  makeWarehouseReportV8(...)

  LƯU Ý:
  - Không render WarehouseTxModal cũ.
  - Không gọi addWarehouseTx cũ.
  - Không gọi importWarehouseExcel cũ.
  - Không có đoạn assets import => warehouse.
*/
