/*
  MYHL-QUẢN LÝ TÀI SẢN
  WAREHOUSE CORE V7 — 24/08/2026

  Mục tiêu nghiệp vụ:
  1. Danh mục tài sản = MASTER DATA, chỉ định danh tài sản.
     Import danh mục TUYỆT ĐỐI KHÔNG tạo phiếu nhập/xuất.
  2. Phiếu nhập/xuất = TRANSACTION. Chỉ giao dịch mới làm phát sinh tồn.
  3. Tồn được tính theo:
       Tài sản + Kho/Công trình + đến ngày.
  4. Báo cáo cho phép lọc:
       Loại tài sản / Nhóm tài sản / Nguồn gốc / Công trình / Kho-Công trình.
  5. Tìm tài sản bằng Mã hàng + Tên + Serial + Loại + Nhóm + Nguồn gốc.
  6. Excel date được chuẩn hóa an toàn: Date, serial, dd/mm/yyyy, yyyy-mm-dd.
  7. Không để timezone làm ngày 24/08 thành 23/08.
  8. Nếu transaction cũ mất assetId, hệ thống vẫn khôi phục bằng Mã hàng/Tên tài sản.
  9. Import phiếu kho kiểm tra toàn bộ trước khi ghi; có lỗi => không ghi dòng nào.
 10. Xuất kho phải kiểm tra tồn đúng tài sản + đúng kho/công trình + đúng ngày.

  Core này dùng các helper có sẵn trong App.jsx:
  Modal, Field, Btn, Th, Td, Tag, EmptyState,
  inputCls, inputStyle, TOKENS, uid, nowIso, fmtVND,
  XLSX, Search, X, UploadCloud, DownloadCloud.
*/

/* =========================================================
   1. CHUẨN HÓA CHUỖI / NGÀY
========================================================= */

function whText(v) {
  return String(v ?? "").trim();
}

function whKey(v) {
  return whText(v)
    .toLocaleLowerCase("vi-VN")
    .normalize("NFC");
}

/*
  Không dùng new Date("2026-08-24") để so sánh/hiển thị.
  Chuỗi YYYY-MM-DD được giữ nguyên để tránh lệch timezone.
*/
function whDate(value) {
  if (value === null || value === undefined || value === "") return "";

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    // Excel cellDates -> lấy LOCAL calendar date, không dùng toISOString().
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

  const s = whText(value);
  if (!s) return "";

  let m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if (m) {
    const d = Number(m[1]);
    const mo = Number(m[2]);
    const y = Number(m[3]);
    if (d >= 1 && d <= 31 && mo >= 1 && mo <= 12) {
      return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    }
  }

  m = s.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})$/);
  if (m) {
    const y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);
    if (d >= 1 && d <= 31 && mo >= 1 && mo <= 12) {
      return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    }
  }

  m = s.match(/^(\d{4}-\d{2}-\d{2})(?:T|\s)/);
  if (m) return m[1];

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  return "";
}

function whDateVN(value) {
  const d = whDate(value);
  if (!d) return "";
  const [y, m, day] = d.split("-");
  return `${day}/${m}/${y}`;
}

/* =========================================================
   2. XÁC ĐỊNH KHO / CÔNG TRÌNH
========================================================= */

function whLocation(tx, projects = []) {
  if (whText(tx.locationName)) return whText(tx.locationName);

  if (tx.locationType === "project" && tx.projectId) {
    const p = projects.find(x => x.id === tx.projectId);
    if (p?.name) return whText(p.name);
  }

  if (whText(tx.warehouseName)) return whText(tx.warehouseName);

  return "Kho trung tâm";
}

function whProjectByName(projects, name) {
  const k = whKey(name);
  return (projects || []).find(p => whKey(p.name) === k) || null;
}

/* =========================================================
   3. TÌM TÀI SẢN AN TOÀN
   - Ưu tiên Mã hàng.
   - Sau đó Tên.
   - Không cho phép Mã và Tên trỏ tới 2 tài sản khác nhau.
========================================================= */

function resolveWarehouseAsset({ row, assets }) {
  const code = whText(
    row["Mã hàng"] ||
    row["Mã quản lý"] ||
    row["Mã tài sản"] ||
    row.code
  );

  const name = whText(
    row["Tên tài sản"] ||
    row["Tên hàng"] ||
    row["Tên thiết bị"] ||
    row.name
  );

  const byCode = code
    ? (assets || []).find(a => whKey(a.code) === whKey(code))
    : null;

  const byName = name
    ? (assets || []).find(a => whKey(a.name) === whKey(name))
    : null;

  if (byCode && byName && byCode.id !== byName.id) {
    return {
      asset: null,
      error: `Mã "${code}" và Tên "${name}" không cùng một tài sản`,
    };
  }

  const asset = byCode || byName;

  if (!asset) {
    return {
      asset: null,
      error: `Không tìm thấy tài sản: ${code || name || "(thiếu mã/tên)"}`,
    };
  }

  return { asset, error: "" };
}

/*
  Khôi phục transaction cũ:
  - assetId
  - nếu không có thì itemCode/code
  - nếu vẫn không có thì itemName/name
*/
function resolveExistingWarehouseAsset(tx, assets) {
  const byId = tx.assetId
    ? (assets || []).find(a => a.id === tx.assetId)
    : null;

  if (byId) return byId;

  const code = tx.itemCode || tx.code || tx.maHang || tx.maQuanLy;
  const byCode = code
    ? (assets || []).find(a => whKey(a.code) === whKey(code))
    : null;

  if (byCode) return byCode;

  const name = tx.itemName || tx.name || tx.tenTaiSan || tx.tenHang;
  return name
    ? (assets || []).find(a => whKey(a.name) === whKey(name))
    : null;
}

/* =========================================================
   4. TÍNH TỒN — NGUỒN SỰ THẬT DUY NHẤT
========================================================= */

function buildWarehouseBalancesV7({
  warehouse = [],
  assets = [],
  projects = [],
  asOfDate,
}) {
  const cutoff = whDate(asOfDate) || "9999-12-31";
  const balances = {};

  warehouse.forEach(tx => {
    const date = whDate(tx.date || tx.ngayThang);
    if (!date || date > cutoff) return;

    const asset = resolveExistingWarehouseAsset(tx, assets);
    if (!asset) return;

    const location = whLocation(tx, projects);
    const locationKey = whKey(location);

    /*
      Dùng asset.id nếu có.
      Nếu dữ liệu cũ thiếu id, dùng code để vẫn gom đúng.
    */
    const assetKey = asset.id || whKey(asset.code || asset.name);
    const key = `${assetKey}¦${locationKey}`;

    if (!balances[key]) {
      balances[key] = {
        assetId: asset.id,
        code: asset.code || tx.itemCode || "",
        name: asset.name || tx.itemName || "",
        location,
        locationType: tx.locationType || "",
        projectId: tx.projectId || null,
        category: asset.category || tx.category || "Khác",
        group: asset.assetGroup || tx.assetGroup || "Thiết bị chính",
        ownership: asset.ownership || tx.ownership || "Công ty",
        unit: asset.unit || tx.unit || "Cái",
        serial: asset.serial || tx.serial || "",
        inQty: 0,
        outQty: 0,
        inValue: 0,
        outValue: 0,
      };
    }

    const qty = Number(tx.quantity ?? tx.soLuong ?? 0) || 0;
    const unitCost = Number(tx.unitCost ?? tx.unitPrice ?? tx.donGia ?? 0) || 0;
    const rawTotal = Number(tx.total ?? tx.thanhTien);
    const total = Number.isFinite(rawTotal) && rawTotal !== 0
      ? rawTotal
      : qty * unitCost;

    const type = whKey(tx.type || tx.loaiPhieu);

    if (type === "nhap" || type.includes("nhập")) {
      balances[key].inQty += qty;
      balances[key].inValue += total;
    } else if (type === "xuat" || type.includes("xuất")) {
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

/* =========================================================
   5. TỒN CỦA MỘT TÀI SẢN TẠI MỘT KHO/CÔNG TRÌNH
========================================================= */

function getWarehouseStockV7({
  warehouse = [],
  assets = [],
  projects = [],
  assetId,
  locationName,
  asOfDate,
}) {
  const rows = buildWarehouseBalancesV7({
    warehouse,
    assets,
    projects,
    asOfDate,
  });

  return rows.find(r =>
    r.assetId === assetId &&
    whKey(r.location) === whKey(locationName)
  )?.balanceQty || 0;
}

/* =========================================================
   6. AUTOCOMPLETE MÃ / TÊN
========================================================= */

function WarehouseAssetAutocompleteV7({
  assets = [],
  value,
  onChange,
}) {
  const selected = assets.find(a => a.id === value) || null;
  const [q, setQ] = React.useState("");
  const [open, setOpen] = React.useState(false);

  React.useEffect(() => {
    if (selected) {
      setQ(`${selected.code || ""} — ${selected.name || ""}`);
    }
  }, [selected?.id]);

  const options = React.useMemo(() => {
    const s = whKey(q);

    if (!s || (selected && s === whKey(`${selected.code} — ${selected.name}`))) {
      return assets.slice(0, 40);
    }

    return assets
      .filter(a => {
        const hay = [
          a.code,
          a.name,
          a.serial,
          a.category,
          a.assetGroup,
          a.ownership,
        ].map(whKey).join(" ");
        return hay.includes(s);
      })
      .slice(0, 40);
  }, [assets, q, selected?.id]);

  const choose = a => {
    onChange(a.id);
    setQ(`${a.code || ""} — ${a.name || ""}`);
    setOpen(false);
  };

  return (
    <div className="relative">
      <div className="relative">
        <Search
          size={14}
          className="absolute left-2.5 top-1/2 -translate-y-1/2"
          style={{ color: TOKENS.muted }}
        />
        <input
          className={inputCls}
          style={{ ...inputStyle, paddingLeft: 32 }}
          value={q}
          placeholder="Gõ mã hàng hoặc tên tài sản..."
          autoComplete="off"
          onFocus={() => setOpen(true)}
          onChange={e => {
            setQ(e.target.value);
            setOpen(true);

            if (!e.target.value.trim()) {
              onChange("");
            }
          }}
        />

        {selected && (
          <button
            type="button"
            className="absolute right-2 top-1/2 -translate-y-1/2"
            onClick={() => {
              onChange("");
              setQ("");
              setOpen(true);
            }}
          >
            <X size={14} style={{ color: TOKENS.muted }} />
          </button>
        )}
      </div>

      {open && (
        <div
          className="absolute z-[100] left-0 right-0 mt-1 max-h-72 overflow-auto rounded-lg shadow-xl"
          style={{
            background: TOKENS.surface,
            border: `1px solid ${TOKENS.border}`,
          }}
          onMouseDown={e => e.preventDefault()}
        >
          {options.map(a => (
            <button
              type="button"
              key={a.id}
              className="w-full text-left px-3 py-2.5 hover:bg-black/5"
              onClick={() => choose(a)}
            >
              <div className="flex items-center gap-2">
                <Tag>{a.code || "—"}</Tag>
                <span className="text-[13px] font-medium">
                  {a.name || "Không tên"}
                </span>
              </div>

              <div
                className="text-[11px] mt-1"
                style={{ color: TOKENS.muted }}
              >
                {(a.category || "Khác")} ·
                {" "}{(a.assetGroup || "Thiết bị chính")} ·
                {" "}{(a.ownership || "Công ty")} ·
                {" "}{a.unit || "Cái"}
                {a.serial ? ` · Serial: ${a.serial}` : ""}
              </div>
            </button>
          ))}

          {!options.length && (
            <div
              className="px-3 py-5 text-[12px]"
              style={{ color: TOKENS.muted }}
            >
              Không tìm thấy tài sản phù hợp.
            </div>
          )}

          <button
            type="button"
            className="w-full text-center py-2 text-[11px]"
            style={{
              color: TOKENS.muted,
              borderTop: `1px solid ${TOKENS.border}`,
            }}
            onClick={() => setOpen(false)}
          >
            Đóng
          </button>
        </div>
      )}
    </div>
  );
}

/* =========================================================
   7. MODAL PHIẾU NHẬP / XUẤT
========================================================= */

function WarehouseTxModalV7({
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
    date: whDate(nowIso()),
    unit: "Cái",
    unitCost: 0,
    receiver: "",
    locationType: "project",
    warehouseName: "",
    projectId: "",
    note: "",
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
      : whText(f.warehouseName);

  const total =
    (Number(f.quantity) || 0) *
    (Number(f.unitCost) || 0);

  const valid =
    !!f.assetId &&
    !!whDate(f.date) &&
    Number(f.quantity) > 0 &&
    !!locationName;

  return (
    <Modal
      title={title || (fixedType === "xuat"
        ? "Lập phiếu xuất kho"
        : "Lập phiếu nhập kho")}
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
          {fixedType === "xuat"
            ? "PHIẾU XUẤT KHO"
            : "PHIẾU NHẬP KHO"}
        </div>
        <div
          className="text-[11px] mt-1"
          style={{ color: TOKENS.muted }}
        >
          Danh mục tài sản chỉ dùng để định danh. Phiếu này mới làm phát sinh tồn.
        </div>
      </div>

      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Số phiếu">
          <input
            className={inputCls}
            style={inputStyle}
            value={f.voucherNo}
            onChange={e => set("voucherNo", e.target.value)}
            placeholder={
              fixedType === "xuat"
                ? "PX-20260824-001"
                : "PN-20260824-001"
            }
          />
        </Field>

        <Field label="Ngày tháng">
          <input
            type="date"
            className={inputCls}
            style={inputStyle}
            value={whDate(f.date)}
            onChange={e => set("date", e.target.value)}
          />
        </Field>

        <Field label="Tên tài sản / Mã hàng">
          <WarehouseAssetAutocompleteV7
            assets={assets}
            value={f.assetId}
            onChange={selectAsset}
          />
        </Field>

        <Field label="Mã hàng tự động">
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
            value={asset.unit || f.unit || "Cái"}
          />
        </Field>

        <Field label="Kho / Công trình">
          <select
            className={inputCls}
            style={inputStyle}
            value={f.locationType}
            onChange={e => {
              setF(prev => ({
                ...prev,
                locationType: e.target.value,
                projectId: "",
                warehouseName: "",
              }));
            }}
          >
            <option value="project">Công trình</option>
            <option value="warehouse">Kho</option>
          </select>
        </Field>

        {f.locationType === "project" ? (
          <Field label="Tên công trình">
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
            value={fmtVND(total)}
          />
        </Field>

        <div className="col-span-2">
          <Field label="Ghi chú">
            <textarea
              rows={3}
              className={inputCls}
              style={inputStyle}
              value={f.note}
              onChange={e => set("note", e.target.value)}
            />
          </Field>
        </div>
      </div>

      <div className="flex justify-end gap-2 mt-2">
        <Btn onClick={onClose}>Huỷ</Btn>

        <Btn
          kind="primary"
          disabled={!valid}
          onClick={() =>
            onSubmit({
              ...f,
              type: fixedType,
              date: whDate(f.date),
              locationName,
              unit: asset.unit || f.unit || "Cái",
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

/* =========================================================
   8. TẠO GIAO DỊCH NHẬP / XUẤT
   Dùng cho cả form tay và import Excel.
========================================================= */

function createWarehouseTransactionV7({
  form,
  data,
  requireAdmin,
  notify,
  setData,
}) {
  if (requireAdmin && !requireAdmin()) return false;

  const assets = data.assets || [];
  const projects = data.projects || [];
  const warehouse = data.warehouse || [];

  const asset = assets.find(a => a.id === form.assetId);

  if (!asset) {
    notify("Không xác định được tài sản.");
    return false;
  }

  const date = whDate(form.date);

  if (!date) {
    notify("Ngày chứng từ không hợp lệ.");
    return false;
  }

  const qty = Number(form.quantity) || 0;

  if (qty <= 0) {
    notify("Số lượng phải lớn hơn 0.");
    return false;
  }

  const location =
    whText(form.locationName) ||
    (form.locationType === "project"
      ? projects.find(p => p.id === form.projectId)?.name
      : form.warehouseName);

  if (!location) {
    notify("Vui lòng chọn Kho/Công trình.");
    return false;
  }

  const type = whKey(form.type);
  const isNhap = type === "nhap" || type.includes("nhập");

  const prefix = isNhap ? "PN" : "PX";
  const dateKey = date.replaceAll("-", "");

  const sameDay = warehouse.filter(w =>
    String(w.voucherNo || "").startsWith(`${prefix}-${dateKey}-`)
  );

  const voucherNo =
    whText(form.voucherNo) ||
    `${prefix}-${dateKey}-${String(sameDay.length + 1).padStart(3, "0")}`;

  const tx = {
    id: uid("wh"),
    voucherNo,
    assetId: asset.id,
    type: isNhap ? "nhap" : "xuat",
    quantity: qty,
    date,
    unitCost: Number(form.unitCost || 0),
    total: qty * Number(form.unitCost || 0),
    unit: asset.unit || form.unit || "Cái",
    receiver: whText(form.receiver),
    note: whText(form.note),

    // Luôn lấy master data.
    category: asset.category || "Khác",
    assetGroup: asset.assetGroup || "Thiết bị chính",
    ownership: asset.ownership || "Công ty",

    locationType: form.locationType || "project",
    locationName: location,
    warehouseName:
      form.locationType === "warehouse" ? location : "",
    projectId: form.locationType === "project"
      ? form.projectId || null
      : null,

    itemName: asset.name,
    itemCode: asset.code,
    serial: asset.serial || "",
  };

  if (!isNhap) {
    const stock = getWarehouseStockV7({
      warehouse,
      assets,
      projects,
      assetId: asset.id,
      locationName: location,
      asOfDate: date,
    });

    if (qty > stock) {
      notify(
        `Không đủ tồn tại "${location}". ${asset.code} — ${asset.name} hiện còn ${stock} ${tx.unit}.`
      );
      return false;
    }
  }

  const transaction = {
    id: uid("tx"),
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
    transactions: [
      transaction,
      ...(data.transactions || []),
    ],
    activityLog: typeof logAction === "function"
      ? logAction(
          data.activityLog,
          `${isNhap ? "Nhập" : "Xuất"} ${voucherNo} — ${asset.code} — ${location}`
        )
      : data.activityLog,
  });

  notify(
    `${isNhap ? "Đã lập phiếu nhập" : "Đã lập phiếu xuất"} ${voucherNo}`
  );

  return true;
}

/* =========================================================
   9. IMPORT DANH MỤC TÀI SẢN
   QUAN TRỌNG: KHÔNG ĐỤNG warehouse / transactions.
========================================================= */

function importAssetCatalogV7({
  file,
  data,
  setData,
  requireAdmin,
  notify,
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

      const rows = XLSX.utils.sheet_to_json(sheet, {
        defval: "",
      });

      if (!rows.length) {
        notify("File danh mục không có dữ liệu.");
        return;
      }

      const existing = data.assets || [];

      const existingByCode = Object.fromEntries(
        existing.map(a => [whKey(a.code), a])
      );

      const imported = [];
      const errors = [];
      const seen = new Set();

      rows.forEach((r, i) => {
        const rowNo = i + 2;

        const code = whText(
          r["Mã quản lý"] ||
          r["Mã hàng"] ||
          r["Mã tài sản"] ||
          r.code
        );

        const name = whText(
          r["Tên tài sản"] ||
          r["Tên hàng"] ||
          r["Tên thiết bị"] ||
          r.name
        );

        if (!code || !name) {
          errors.push(`Dòng ${rowNo}: thiếu Mã quản lý hoặc Tên tài sản.`);
          return;
        }

        if (seen.has(whKey(code))) {
          errors.push(`Dòng ${rowNo}: mã ${code} bị trùng trong file.`);
          return;
        }

        seen.add(whKey(code));

        const old = existingByCode[whKey(code)];

        imported.push({
          ...old,

          // Giữ ID cũ để KHÔNG làm mất liên kết với phiếu kho.
          id: old?.id || uid("as"),

          code,
          name,

          category: whText(
            r["Loại tài sản"] ||
            r["Loại"] ||
            r.category
          ) || old?.category || "Khác",

          assetGroup: whText(
            r["Nhóm tài sản"] ||
            r.assetGroup
          ) || old?.assetGroup || "Thiết bị chính",

          ownership: whText(
            r["Nguồn gốc"] ||
            r["Nguồn"] ||
            r.ownership
          ) || old?.ownership || "Công ty",

          unit: whText(
            r["Đơn vị tính"] ||
            r["ĐVT"] ||
            r.unit
          ) || old?.unit || "Cái",

          cost: Number(
            r["Nguyên giá"] ??
            r.cost ??
            old?.cost ??
            0
          ) || 0,

          purchaseDate: whDate(
            r["Ngày mua"] ||
            r.purchaseDate ||
            old?.purchaseDate
          ),

          serial: whText(
            r["Serial"] ||
            r["Số serial"] ||
            r.serial
          ) || old?.serial || "",

          supplier: whText(
            r["Nhà cung cấp"] ||
            r.supplier
          ) || old?.supplier || "",

          warranty:
            String(
              r["Bảo hành"] ??
              old?.warranty ??
              ""
            ).toLowerCase() === "có",

          warrantyEnd: whDate(
            r["Hạn bảo hành"] ||
            r.warrantyEnd ||
            old?.warrantyEnd
          ),

          note: whText(
            r["Ghi chú"] ||
            r.note
          ) || old?.note || "",

          // Danh mục chỉ là master data.
          // Không tự đưa tài sản xuống công trình.
          projectId: old?.projectId || null,
          assignedTo: old?.assignedTo || null,
        });
      });

      if (!imported.length) {
        notify(errors[0] || "Không có dòng danh mục hợp lệ.");
        return;
      }

      const importedCodes = new Set(
        imported.map(a => whKey(a.code))
      );

      const remaining = existing.filter(
        a => !importedCodes.has(whKey(a.code))
      );

      /*
        ĐÂY LÀ ĐIỂM SỬA LỖI QUAN TRỌNG:
        warehouse giữ nguyên 100%.
        transactions giữ nguyên 100%.
        Import danh mục không được tạo PN.
      */
      setData({
        ...data,
        assets: [...imported, ...remaining],
        warehouse: data.warehouse || [],
        transactions: data.transactions || [],
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
          : `Đã cập nhật ${imported.length} danh mục tài sản. Không tạo phiếu nhập.`
      );
    } catch (err) {
      console.error(err);
      notify("Không đọc được file Excel danh mục tài sản.");
    }
  };

  reader.readAsArrayBuffer(file);
}

/* =========================================================
   10. IMPORT PHIẾU NHẬP / XUẤT KHO
========================================================= */

function importWarehouseExcelV7({
  file,
  data,
  setData,
  requireAdmin,
  notify,
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

      const rows = XLSX.utils.sheet_to_json(sheet, {
        defval: "",
      });

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

        const rawType = whKey(
          r["Loại phiếu"] ||
          r["Loại chứng từ"] ||
          "Nhập kho"
        );

        const type =
          rawType.includes("xuất") ||
          rawType === "xuat"
            ? "xuat"
            : "nhap";

        const date = whDate(
          r["Ngày tháng"] ||
          r["Ngày chứng từ"] ||
          r["Ngày"] 
        );

        if (!date) {
          errors.push(`Dòng ${rowNo}: ngày không hợp lệ.`);
          return;
        }

        const resolved = resolveWarehouseAsset({
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
          errors.push(
            `Dòng ${rowNo}: Số lượng phải lớn hơn 0.`
          );
          return;
        }

        const locationText = whText(
          r["Kho/Công trình"] ||
          r["Công trình"] ||
          r["Tên kho"] ||
          r["Kho"]
        );

        if (!locationText) {
          errors.push(
            `Dòng ${rowNo}: thiếu Kho/Công trình.`
          );
          return;
        }

        const project = whProjectByName(
          projects,
          locationText
        );

        const rawLocationType = whKey(
          r["Loại địa điểm"] ||
          r["Loại nơi"] ||
          ""
        );

        let locationType;

        if (rawLocationType.includes("công trình")) {
          locationType = "project";
        } else if (rawLocationType.includes("kho")) {
          locationType = "warehouse";
        } else {
          locationType = project ? "project" : "warehouse";
        }

        const unitCost = Number(
          r["Đơn giá"] ??
          r["Đơn giá nhập"] ??
          r["Đơn giá xuất"] ??
          0
        ) || 0;

        const suppliedTotal = Number(
          r["Thành tiền"] ??
          r["Giá trị"] ??
          0
        );

        const total =
          Number.isFinite(suppliedTotal) &&
          suppliedTotal !== 0
            ? suppliedTotal
            : qty * unitCost;

        const prefix = type === "nhap" ? "PN" : "PX";
        const dateKey = date.replaceAll("-", "");

        const seq =
          [...existing, ...imported].filter(w =>
            String(w.voucherNo || "").startsWith(
              `${prefix}-${dateKey}-`
            )
          ).length + 1;

        const voucherNo =
          whText(r["Số phiếu"]) ||
          `${prefix}-${dateKey}-${String(seq).padStart(3, "0")}`;

        imported.push({
          id: uid("wh"),

          voucherNo,
          assetId: asset.id,
          type,
          quantity: qty,
          date,
          unitCost,
          total,

          // Luôn lấy master data từ Danh mục.
          itemName: asset.name,
          itemCode: asset.code,
          category: asset.category || "Khác",
          assetGroup: asset.assetGroup || "Thiết bị chính",
          ownership: asset.ownership || "Công ty",
          unit:
            asset.unit ||
            whText(r["Đơn vị tính"] || r["ĐVT"]) ||
            "Cái",

          receiver: whText(
            r["Người giao/nhận"] ||
            r["Người giao"] ||
            r["Người nhận"]
          ),

          note: whText(r["Ghi chú"]),

          locationType,
          locationName: locationText,
          warehouseName:
            locationType === "warehouse"
              ? locationText
              : "",

          projectId:
            locationType === "project"
              ? project?.id || null
              : null,

          serial: asset.serial || "",
        });
      });

      /*
        Validate xuất theo đúng:
        Tài sản + Kho/Công trình + ngày.
        Không được lấy tổng tồn toàn hệ thống.
      */
      for (const tx of imported) {
        if (tx.type !== "xuat") continue;

        const rowsAtLocation = [
          ...existing,
          ...imported,
        ].filter(w => {
          const a = resolveExistingWarehouseAsset(w, assets);

          return (
            a?.id === tx.assetId &&
            whKey(whLocation(w, projects)) ===
              whKey(tx.locationName) &&
            whDate(w.date) <= tx.date
          );
        });

        const available = rowsAtLocation.reduce(
          (sum, w) => {
            const type = whKey(w.type);

            if (
              type === "nhap" ||
              type.includes("nhập")
            ) {
              return sum + (Number(w.quantity) || 0);
            }

            if (
              type === "xuat" ||
              type.includes("xuất")
            ) {
              return sum - (Number(w.quantity) || 0);
            }

            return sum;
          },
          0
        );

        if (tx.quantity > available) {
          errors.push(
            `${tx.voucherNo}: xuất ${tx.quantity} ${tx.unit} tại "${tx.locationName}" vượt tồn ${available}.`
          );
        }
      }

      /*
        Transaction trùng số phiếu không được ghi lần 2.
      */
      const existingVoucherKeys = new Set(
        existing.map(w => whKey(w.voucherNo))
      );

      const duplicateVoucher = imported.find(
        tx => existingVoucherKeys.has(whKey(tx.voucherNo))
      );

      if (duplicateVoucher) {
        errors.push(
          `Số phiếu ${duplicateVoucher.voucherNo} đã tồn tại. Không import trùng phiếu.`
        );
      }

      if (errors.length) {
        notify(
          `Import bị dừng. ${errors.length} lỗi. ${errors[0]}`
        );
        return;
      }

      const transactions = imported.map(tx => ({
        id: uid("tx"),
        assetId: tx.assetId,
        type:
          tx.type === "nhap"
            ? "nhap_kho"
            : "xuat_kho",
        date: tx.date,
        title:
          `${tx.type === "nhap" ? "Nhập kho" : "Xuất kho"} ${tx.voucherNo}`,
        detail:
          `${tx.itemName} · ${tx.locationName} · ${tx.receiver || ""}`,
        amount: tx.total,
      }));

      setData({
        ...data,

        /*
          Import phiếu kho CHỈ thêm warehouse transactions.
          Không thay đổi assets.
        */
        warehouse: [
          ...imported,
          ...existing,
        ],

        transactions: [
          ...transactions,
          ...(data.transactions || []),
        ],

        activityLog:
          typeof logAction === "function"
            ? logAction(
                data.activityLog,
                `Import ${imported.length} phiếu nhập/xuất kho`
              )
            : data.activityLog,
      });

      notify(
        `Đã import ${imported.length} phiếu kho. Báo cáo tồn đã có thể tính đến ngày chứng từ.`
      );
    } catch (err) {
      console.error(err);
      notify("Không đọc được file Excel phiếu nhập/xuất kho.");
    }
  };

  reader.readAsArrayBuffer(file);
}

/* =========================================================
   11. FILTER BÁO CÁO
========================================================= */

function WarehouseFilterV7({
  filter,
  setFilter,
  categories = [],
  groups = [],
  ownerships = [],
  projects = [],
  locations = [],
}) {
  const set = key => e =>
    setFilter(prev => ({
      ...prev,
      [key]: e.target.value,
    }));

  return (
    <div className="flex flex-wrap gap-2">
      <select
        className={inputCls}
        style={{ ...inputStyle, width: 175 }}
        value={filter.category || ""}
        onChange={set("category")}
      >
        <option value="">Tất cả loại tài sản</option>
        {categories.map(x => (
          <option key={x}>{x}</option>
        ))}
      </select>

      <select
        className={inputCls}
        style={{ ...inputStyle, width: 175 }}
        value={filter.group || ""}
        onChange={set("group")}
      >
        <option value="">Tất cả nhóm</option>
        {groups.map(x => (
          <option key={x}>{x}</option>
        ))}
      </select>

      <select
        className={inputCls}
        style={{ ...inputStyle, width: 155 }}
        value={filter.ownership || ""}
        onChange={set("ownership")}
      >
        <option value="">Tất cả nguồn gốc</option>
        {ownerships.map(x => (
          <option key={x}>{x}</option>
        ))}
      </select>

      <select
        className={inputCls}
        style={{ ...inputStyle, width: 220 }}
        value={filter.projectId || ""}
        onChange={set("projectId")}
      >
        <option value="">Tất cả công trình</option>
        {projects.map(p => (
          <option key={p.id} value={p.id}>
            {p.name}
          </option>
        ))}
      </select>

      <select
        className={inputCls}
        style={{ ...inputStyle, width: 240 }}
        value={filter.locationName || ""}
        onChange={set("locationName")}
      >
        <option value="">
          Tất cả kho / công trình
        </option>
        {locations.map(x => (
          <option key={x}>{x}</option>
        ))}
      </select>

      <input
        type="date"
        className={inputCls}
        style={{ ...inputStyle, width: 170 }}
        value={whDate(filter.asOfDate)}
        onChange={e =>
          setFilter(prev => ({
            ...prev,
            asOfDate: e.target.value,
          }))
        }
      />
    </div>
  );
}

/* =========================================================
   12. HÀM LỌC BÁO CÁO
========================================================= */

function filterWarehouseReportV7({
  rows = [],
  filter = {},
  projects = [],
}) {
  return rows
    .filter(r =>
      !filter.category ||
      whKey(r.category) === whKey(filter.category)
    )
    .filter(r =>
      !filter.group ||
      whKey(r.group) === whKey(filter.group)
    )
    .filter(r =>
      !filter.ownership ||
      whKey(r.ownership) === whKey(filter.ownership)
    )
    .filter(r =>
      !filter.locationName ||
      whKey(r.location) === whKey(filter.locationName)
    )
    .filter(r => {
      if (!filter.projectId) return true;

      const project = projects.find(
        p => p.id === filter.projectId
      );

      return (
        project &&
        whKey(project.name) === whKey(r.location)
      );
    })
    .filter(r => r.balanceQty > 0)
    .sort((a, b) =>
      `${a.location}|${a.group}|${a.category}|${a.code}|${a.name}`
        .localeCompare(
          `${b.location}|${b.group}|${b.category}|${b.code}|${b.name}`,
          "vi"
        )
    );
}

/* =========================================================
   13. FORMAT BÁO CÁO
========================================================= */

function makeWarehouseReportV7({
  warehouse,
  assets,
  projects,
  filter,
}) {
  const allRows = buildWarehouseBalancesV7({
    warehouse,
    assets,
    projects,
    asOfDate: filter.asOfDate,
  });

  const rows = filterWarehouseReportV7({
    rows: allRows,
    filter,
    projects,
  });

  return {
    rows,

    summary: {
      qty: rows.reduce(
        (s, r) => s + r.balanceQty,
        0
      ),
      value: rows.reduce(
        (s, r) => s + r.balanceValue,
        0
      ),
      locations: new Set(
        rows.map(r => r.location)
      ).size,
      assets: rows.length,
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

    asOfDate: whDate(filter.asOfDate),
    asOfDateVN: whDateVN(filter.asOfDate),
  };
}

/* =========================================================
   14. KIỂM TRA NHANH DỮ LIỆU KHO
   Dùng trong console/debug khi cần.
========================================================= */

function auditWarehouseDataV7({
  warehouse = [],
  assets = [],
  projects = [],
}) {
  const result = {
    totalAssets: assets.length,
    totalWarehouseTransactions: warehouse.length,
    orphanTransactions: [],
    invalidDates: [],
    invalidLocations: [],
    negativeLocations: [],
  };

  warehouse.forEach(tx => {
    if (!resolveExistingWarehouseAsset(tx, assets)) {
      result.orphanTransactions.push({
        voucherNo: tx.voucherNo,
        itemCode: tx.itemCode,
        itemName: tx.itemName,
      });
    }

    if (!whDate(tx.date)) {
      result.invalidDates.push({
        voucherNo: tx.voucherNo,
        date: tx.date,
      });
    }

    if (!whLocation(tx, projects)) {
      result.invalidLocations.push(tx.voucherNo);
    }
  });

  const balances = buildWarehouseBalancesV7({
    warehouse,
    assets,
    projects,
    asOfDate: "9999-12-31",
  });

  result.negativeLocations = balances
    .filter(r => r.balanceQty < 0)
    .map(r => ({
      code: r.code,
      name: r.name,
      location: r.location,
      balanceQty: r.balanceQty,
    }));

  return result;
}

/* =========================================================
   15. INTEGRATION BẮT BUỘC VÀO App.jsx
========================================================= */

/*
A. IMPORT DANH MỤC
--------------------------------------------
Trong handler import danh mục tài sản, KHÔNG được có:

  const wh = imported.map(...)

hoặc:

  warehouse: [...wh, ...data.warehouse]

hoặc:

  transactions: [...]

Thay bằng:

  importAssetCatalogV7({
    file,
    data,
    setData,
    requireAdmin,
    notify,
  });


B. IMPORT PHIẾU KHO
--------------------------------------------

  importWarehouseExcelV7({
    file,
    data,
    setData,
    requireAdmin,
    notify,
  });


C. FORM NHẬP
--------------------------------------------

  <WarehouseTxModalV7
    title="Lập phiếu nhập kho"
    fixedType="nhap"
    assets={data.assets}
    projects={data.projects}
    onClose={() => setModal(null)}
    onSubmit={form => {
      createWarehouseTransactionV7({
        form,
        data,
        requireAdmin,
        notify,
        setData,
      });
      setModal(null);
    }}
  />


D. FORM XUẤT
--------------------------------------------

  <WarehouseTxModalV7
    title="Lập phiếu xuất kho"
    fixedType="xuat"
    assets={data.assets}
    projects={data.projects}
    onClose={() => setModal(null)}
    onSubmit={form => {
      createWarehouseTransactionV7({
        form,
        data,
        requireAdmin,
        notify,
        setData,
      });
      setModal(null);
    }}
  />


E. BÁO CÁO
--------------------------------------------

  const report = React.useMemo(() =>
    makeWarehouseReportV7({
      warehouse: data.warehouse,
      assets: data.assets,
      projects: data.projects,
      filter,
    }),
    [
      data.warehouse,
      data.assets,
      data.projects,
      filter,
    ]
  );

  Hiển thị:

  report.rows

  và tiêu đề:

  Chi tiết thiết bị đang tồn tại đến ngày ${report.asOfDateVN}


F. QUAN TRỌNG — KHÔNG dùng:

  new Date(filter.asOfDate).toLocaleDateString()

cho tiêu đề báo cáo.

Dùng:

  whDateVN(filter.asOfDate)

để ngày 24/08/2026 luôn hiển thị là 24/08/2026.


G. FILTER MẶC ĐỊNH
--------------------------------------------

  const [filter, setFilter] = React.useState({
    category: "",
    group: "",
    ownership: "",
    projectId: "",
    locationName: "",
    asOfDate: whDate(nowIso()),
  });


H. XÓA / VÔ HIỆU HÓA MIGRATION SAI
--------------------------------------------

Trong withDefaults() cũ, KHÔNG được tạo:

  warehouse: migratedAssets.map(...)

từ danh mục tài sản.

Danh mục tài sản không đồng nghĩa với tồn kho.

Nếu dữ liệu cũ không có warehouse thì dùng:

  warehouse: Array.isArray(d.warehouse)
    ? d.warehouse
    : []

Không tự sinh "Tồn đầu kỳ" từ toàn bộ assets.


I. KHÔNG chạy:
  seed()
  DROP TABLE
  DELETE app_data

J. KHÔNG đổi:
  VITE_SUPABASE_URL
  VITE_SUPABASE_ANON_KEY
*/

/* =========================================================
   16. NGHIỆP VỤ CHUẨN CỦA APP SAU V7
========================================================= */

/*
DANH MỤC TÀI SẢN
    MX-01 | Máy xúc 01 | Máy xúc | Thiết bị chính | Thuê | Cái
                         |
                         v
              CHỈ LÀ MASTER DATA
                         |
          +--------------+--------------+
          |                             |
       PHIẾU NHẬP                    PHIẾU XUẤT
          |                             |
          v                             v
  Kho/Công trình A               Kho/Công trình A
          |                             |
          +--------------+--------------+
                         |
                         v
                BÁO CÁO NHẬP-XUẤT-TỒN
                         |
                         v
      Tài sản + Kho/Công trình + Đến ngày

Ví dụ:
  MX-01 Máy xúc 01 — Thuê — Thiết bị chính
  MX-02 Máy xúc 02 — Thuê — Thiết bị chính
  MC-04 Máy cẩu 04 — Thuê — Thiết bị chính
  MK-05 Máy khoan 05 — Thuê — Thiết bị chính

Lọc:
  Nhóm = Thiết bị chính
  Nguồn gốc = Thuê
  Công trình = Cao Xà Lá - Thanh Xuân
  Đến ngày = 24/08/2026

=> Báo cáo trả từng tài sản:
  MX-01
  MX-02
  MC-04
  MK-05
  ...
với số lượng tồn thực tế tại đúng công trình.
*/
