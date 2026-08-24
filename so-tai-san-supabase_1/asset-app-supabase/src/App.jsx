/*
MYHL-QUẢN LÝ TÀI SẢN — WAREHOUSE CORE V6

Mục tiêu:
1) Danh mục tài sản = MASTER DATA, tuyệt đối KHÔNG tự tạo phiếu nhập.
2) Phiếu nhập/xuất = TRANSACTION, là nguồn để tính tồn.
3) Chuẩn hóa ngày Excel: Date object / serial / dd/mm/yyyy / yyyy-mm-dd.
4) Tìm tài sản theo Mã + Tên + Serial bằng autocomplete.
5) Báo cáo tồn theo Kho/Công trình + Loại + Nhóm + Nguồn gốc + đến ngày.
6) Import phiếu kho kiểm tra lỗi trước; có lỗi thì không ghi dữ liệu.

Core này dùng các helper/component đã có trong App.jsx:
Modal, Field, Btn, Th, Td, Tag, EmptyState, ExportBar,
inputCls, inputStyle, TOKENS, uid, nowIso, fmtDate, fmtVND,
XLSX, Search, X, DownloadCloud, UploadCloud, Plus, ClipboardList.
*/

function normalizeWarehouseDate(value) {
  if (value === null || value === undefined || value === "") return "";

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    try {
      const d = XLSX.SSF.parse_date_code(value);
      if (d && d.y && d.m && d.d) {
        return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
      }
    } catch (_) {}
  }

  const s = String(value).trim();
  if (!s) return "";

  let m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if (m) return `${m[3]}-${String(m[2]).padStart(2, "0")}-${String(m[1]).padStart(2, "0")}`;

  m = s.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})$/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, "0")}-${String(m[3]).padStart(2, "0")}`;

  const iso = s.match(/^(\d{4}-\d{2}-\d{2})[T\s]/);
  if (iso) return iso[1];

  return "";
}

const warehouseText = value => String(value ?? "").trim();
const warehouseKey = value => warehouseText(value).toLocaleLowerCase("vi-VN");

function warehouseLocation(tx, projects = []) {
  if (tx.locationName) return warehouseText(tx.locationName);

  if (tx.locationType === "project" && tx.projectId) {
    return warehouseText(projects.find(p => p.id === tx.projectId)?.name) || "Không xác định";
  }

  return warehouseText(tx.warehouseName) || "Kho trung tâm";
}

/* =========================
   AUTOCOMPLETE TÀI SẢN
========================= */

function WarehouseAssetAutocomplete({ assets = [], value, onChange }) {
  const selected = assets.find(a => a.id === value) || null;
  const [open, setOpen] = React.useState(false);
  const [q, setQ] = React.useState("");

  const options = React.useMemo(() => {
    const search = warehouseKey(q);
    if (!search) return assets.slice(0, 30);

    return assets.filter(a => {
      const hay = [
        a.code, a.name, a.serial, a.category,
        a.assetGroup, a.ownership
      ].map(warehouseKey).join(" ");
      return hay.includes(search);
    }).slice(0, 30);
  }, [assets, q]);

  React.useEffect(() => {
    if (selected) setQ(`${selected.code} — ${selected.name}`);
  }, [selected?.id]);

  return (
    <div className="relative">
      <div className="relative">
        <Search size={14}
          className="absolute left-2.5 top-1/2 -translate-y-1/2"
          style={{ color: TOKENS.muted }}
        />
        <input
          className={inputCls}
          style={{ ...inputStyle, paddingLeft: 32 }}
          value={q}
          placeholder="Gõ mã hoặc tên tài sản..."
          onFocus={() => setOpen(true)}
          onChange={e => {
            setQ(e.target.value);
            setOpen(true);
            if (!e.target.value.trim()) onChange("");
          }}
        />
        {selected && (
          <button type="button"
            className="absolute right-2 top-1/2 -translate-y-1/2"
            onClick={() => {
              onChange("");
              setQ("");
              setOpen(true);
            }}>
            <X size={14} style={{ color: TOKENS.muted }} />
          </button>
        )}
      </div>

      {open && (
        <div
          className="absolute z-50 left-0 right-0 mt-1 max-h-64 overflow-auto rounded-md shadow-lg"
          style={{ background: TOKENS.surface, border: `1px solid ${TOKENS.border}` }}
          onMouseDown={e => e.preventDefault()}
        >
          {options.map(a => (
            <button type="button" key={a.id}
              className="w-full text-left px-3 py-2 hover:bg-black/5"
              onClick={() => {
                onChange(a.id);
                setQ(`${a.code} — ${a.name}`);
                setOpen(false);
              }}>
              <div className="flex items-center gap-2">
                <Tag>{a.code}</Tag>
                <span className="text-[13px] font-medium">{a.name}</span>
              </div>
              <div className="text-[11px] mt-0.5" style={{ color: TOKENS.muted }}>
                {a.category || "Khác"} · {a.assetGroup || "Thiết bị chính"} · {a.ownership || "Công ty"} · {a.serial || "Không có serial"}
              </div>
            </button>
          ))}
          {!options.length && (
            <div className="px-3 py-4 text-[12px]" style={{ color: TOKENS.muted }}>
              Không tìm thấy tài sản phù hợp.
            </div>
          )}
          <button type="button"
            className="w-full text-center py-1.5 text-[11px]"
            style={{ color: TOKENS.muted, borderTop: `1px solid ${TOKENS.border}` }}
            onClick={() => setOpen(false)}>
            Đóng
          </button>
        </div>
      )}
    </div>
  );
}

/* =========================
   PHIẾU NHẬP / XUẤT
========================= */

function WarehouseTxModalV6({ assets, projects, onClose, onSubmit, title, fixedType }) {
  const first = assets[0];
  const [f, setF] = React.useState({
    assetId: first?.id || "",
    type: fixedType || "nhap",
    voucherNo: "",
    quantity: 1,
    date: nowIso().slice(0, 10),
    unit: first?.unit || "Cái",
    unitCost: first?.cost || 0,
    receiver: "",
    locationType: "project",
    warehouseName: "Kho trung tâm",
    projectId: "",
    note: "",
  });

  const asset = assets.find(a => a.id === f.assetId) || {};
  const set = (k, v) => setF(prev => ({ ...prev, [k]: v }));

  const selectAsset = id => {
    const a = assets.find(x => x.id === id);
    setF(prev => ({
      ...prev,
      assetId: id,
      unit: a?.unit || "Cái",
      unitCost: a?.cost || 0,
    }));
  };

  const total = (Number(f.quantity) || 0) * (Number(f.unitCost) || 0);

  const valid =
    f.assetId &&
    normalizeWarehouseDate(f.date) &&
    Number(f.quantity) > 0 &&
    (f.locationType === "warehouse" ? warehouseText(f.warehouseName) : f.projectId);

  return (
    <Modal title={title || "Lập phiếu kho"} onClose={onClose} wide>
      <div className="rounded-lg p-3 mb-4"
        style={{ background: TOKENS.brandSoft, border: `1px solid ${TOKENS.brand}22` }}>
        <div className="text-[12px] font-medium" style={{ color: TOKENS.brand }}>
          {fixedType === "xuat" ? "PHIẾU XUẤT KHO" : "PHIẾU NHẬP KHO"}
        </div>
        <div className="text-[11px] mt-0.5" style={{ color: TOKENS.muted }}>
          Mã, loại, nhóm, nguồn gốc và đơn vị tính được lấy từ Danh mục tài sản.
        </div>
      </div>

      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Số phiếu">
          <input className={inputCls} style={inputStyle}
            value={f.voucherNo}
            onChange={e => set("voucherNo", e.target.value)}
            placeholder={fixedType === "xuat" ? "PX-20260824-001" : "PN-20260824-001"} />
        </Field>

        <Field label="Ngày tháng">
          <input type="date" className={inputCls} style={inputStyle}
            value={f.date}
            onChange={e => set("date", e.target.value)} />
        </Field>

        <Field label="Tên tài sản / Mã hàng">
          <WarehouseAssetAutocomplete
            assets={assets}
            value={f.assetId}
            onChange={selectAsset}
          />
        </Field>

        <Field label="Mã quản lý">
          <input readOnly className={inputCls}
            style={{ ...inputStyle, background: TOKENS.paper }}
            value={asset.code || ""} />
        </Field>

        <Field label="Loại tài sản">
          <input readOnly className={inputCls}
            style={{ ...inputStyle, background: TOKENS.paper }}
            value={asset.category || ""} />
        </Field>

        <Field label="Nhóm tài sản">
          <input readOnly className={inputCls}
            style={{ ...inputStyle, background: TOKENS.paper }}
            value={asset.assetGroup || ""} />
        </Field>

        <Field label="Nguồn gốc">
          <input readOnly className={inputCls}
            style={{ ...inputStyle, background: TOKENS.paper }}
            value={asset.ownership || ""} />
        </Field>

        <Field label="Đơn vị tính">
          <input readOnly className={inputCls}
            style={{ ...inputStyle, background: TOKENS.paper }}
            value={asset.unit || f.unit || "Cái"} />
        </Field>

        <Field label="Loại địa điểm">
          <select className={inputCls} style={inputStyle}
            value={f.locationType}
            onChange={e => set("locationType", e.target.value)}>
            <option value="project">Công trình</option>
            <option value="warehouse">Kho</option>
          </select>
        </Field>

        {f.locationType === "project" ? (
          <Field label="Kho / Công trình">
            <select className={inputCls} style={inputStyle}
              value={f.projectId}
              onChange={e => set("projectId", e.target.value)}>
              <option value="">-- Chọn công trình --</option>
              {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </Field>
        ) : (
          <Field label="Tên kho">
            <input className={inputCls} style={inputStyle}
              value={f.warehouseName}
              onChange={e => set("warehouseName", e.target.value)}
              placeholder="Kho trung tâm" />
          </Field>
        )}

        <Field label="Người giao / nhận">
          <input className={inputCls} style={inputStyle}
            value={f.receiver}
            onChange={e => set("receiver", e.target.value)} />
        </Field>

        <Field label="Số lượng">
          <input type="number" min="0" step="any"
            className={inputCls} style={inputStyle}
            value={f.quantity}
            onChange={e => set("quantity", e.target.value)} />
        </Field>

        <Field label="Đơn giá">
          <input type="number" min="0" step="any"
            className={inputCls} style={inputStyle}
            value={f.unitCost}
            onChange={e => set("unitCost", e.target.value)} />
        </Field>

        <Field label="Thành tiền">
          <input readOnly className={inputCls}
            style={{ ...inputStyle, background: TOKENS.paper }}
            value={fmtVND(total)} />
        </Field>

        <div className="col-span-2">
          <Field label="Ghi chú">
            <textarea className={inputCls} style={inputStyle}
              rows={3}
              value={f.note}
              onChange={e => set("note", e.target.value)} />
          </Field>
        </div>
      </div>

      <div className="flex justify-end gap-2">
        <Btn onClick={onClose}>Huỷ</Btn>
        <Btn kind="primary" disabled={!valid}
          onClick={() => onSubmit({
            ...f,
            date: normalizeWarehouseDate(f.date),
            unit: asset.unit || f.unit || "Cái",
            category: asset.category || "",
            assetGroup: asset.assetGroup || "",
            ownership: asset.ownership || "",
          })}>
          Lưu phiếu
        </Btn>
      </div>
    </Modal>
  );
}

/* =========================
   FILTER
========================= */

function WarehouseFilterV6({ filter, setFilter, categories, groups, ownerships, projects, locations }) {
  const set = key => e => setFilter(prev => ({ ...prev, [key]: e.target.value }));

  return (
    <div className="flex flex-wrap gap-2">
      <select className={inputCls} style={{ ...inputStyle, width: 175 }}
        value={filter.category} onChange={set("category")}>
        <option value="">Tất cả loại tài sản</option>
        {categories.map(x => <option key={x}>{x}</option>)}
      </select>

      <select className={inputCls} style={{ ...inputStyle, width: 175 }}
        value={filter.group} onChange={set("group")}>
        <option value="">Tất cả nhóm</option>
        {groups.map(x => <option key={x}>{x}</option>)}
      </select>

      <select className={inputCls} style={{ ...inputStyle, width: 155 }}
        value={filter.ownership} onChange={set("ownership")}>
        <option value="">Tất cả nguồn gốc</option>
        {ownerships.map(x => <option key={x}>{x}</option>)}
      </select>

      <select className={inputCls} style={{ ...inputStyle, width: 220 }}
        value={filter.projectId} onChange={set("projectId")}>
        <option value="">Tất cả công trình</option>
        {projects.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
      </select>

      <select className={inputCls} style={{ ...inputStyle, width: 230 }}
        value={filter.locationName} onChange={set("locationName")}>
        <option value="">Tất cả kho / công trình</option>
        {locations.map(x => <option key={x}>{x}</option>)}
      </select>
    </div>
  );
}

/* =========================
   TÍNH TỒN
========================= */

function buildWarehouseBalancesV6({ warehouse = [], assets = [], projects = [], asOfDate }) {
  const cutoff = normalizeWarehouseDate(asOfDate) || "9999-12-31";
  const assetMap = Object.fromEntries(assets.map(a => [a.id, a]));
  const balances = {};

  warehouse
    .map(w => ({ ...w, normalizedDate: normalizeWarehouseDate(w.date) }))
    .filter(w => w.normalizedDate && w.normalizedDate <= cutoff)
    .forEach(w => {
      const a = assetMap[w.assetId];
      if (!a) return;

      const location = warehouseLocation(w, projects);
      const key = `${a.id}¦${warehouseKey(location)}`;

      if (!balances[key]) {
        balances[key] = {
          assetId: a.id,
          code: a.code || w.itemCode || "",
          name: a.name || w.itemName || "",
          location,
          category: w.category || a.category || "Khác",
          group: w.assetGroup || a.assetGroup || "Thiết bị chính",
          ownership: w.ownership || a.ownership || "Công ty",
          unit: w.unit || a.unit || "Cái",
          inQty: 0,
          outQty: 0,
          inValue: 0,
          outValue: 0,
        };
      }

      const qty = Number(w.quantity) || 0;
      const unitCost = Number(w.unitCost) || 0;
      const total = Number.isFinite(Number(w.total))
        ? Number(w.total)
        : qty * unitCost;

      if (w.type === "nhap") {
        balances[key].inQty += qty;
        balances[key].inValue += total;
      } else if (w.type === "xuat") {
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

/* =========================
   IMPORT DANH MỤC — KHÔNG TẠO PHIẾU
========================= */

function importAssetCatalogV6({ file, data, setData, requireAdmin, notify }) {
  if (!requireAdmin()) return;

  const reader = new FileReader();

  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, {
        type: "array",
        cellDates: true,
      });

      const rows = XLSX.utils.sheet_to_json(
        wb.Sheets[wb.SheetNames[0]],
        { defval: "" }
      );

      if (!rows.length) {
        notify("File Excel không có dữ liệu");
        return;
      }

      const existing = data.assets || [];
      const existingByCode = Object.fromEntries(
        existing.map(a => [warehouseKey(a.code), a])
      );

      const imported = [];
      const errors = [];

      rows.forEach((r, i) => {
        const rowNo = i + 2;

        const code = warehouseText(
          r["Mã quản lý"] || r["Mã hàng"] || r.code
        );
        const name = warehouseText(
          r["Tên tài sản"] || r["Tên hàng"] || r.name
        );

        if (!code || !name) {
          errors.push(`Dòng ${rowNo}: thiếu Mã quản lý hoặc Tên tài sản`);
          return;
        }

        imported.push({
          id: existingByCode[warehouseKey(code)]?.id || uid("as"),
          code,
          name,
          category: warehouseText(
            r["Loại tài sản"] || r["Loại"] || r.category
          ) || "Khác",
          assetGroup: warehouseText(
            r["Nhóm tài sản"] || r.assetGroup
          ) || "Thiết bị chính",
          ownership: warehouseText(
            r["Nguồn gốc"] || r["Nguồn"] || r.ownership
          ) || "Công ty",
          unit: warehouseText(
            r["Đơn vị tính"] || r["ĐVT"] || r.unit
          ) || "Cái",
          cost: Number(r["Nguyên giá"] || r.cost || 0),
          purchaseDate: normalizeWarehouseDate(
            r["Ngày mua"] || r.purchaseDate
          ),
          usefulLifeMonths: Number(
            r["Thời gian SD"] ||
            r["Thời gian sử dụng"] ||
            r.usefulLifeMonths ||
            36
          ),
          status: existingByCode[warehouseKey(code)]?.status || STATUS.UNUSED,
          assignedTo: null,
          projectId: null,
          department: warehouseText(r["Bộ phận"]) || "Vận hành",
          serial: warehouseText(r["Serial"]),
          supplier: warehouseText(r["Nhà cung cấp"]),
          warranty: String(r["Bảo hành"] || "").toLowerCase() === "có",
          warrantyEnd: normalizeWarehouseDate(r["Hạn bảo hành"]),
          note: warehouseText(r["Ghi chú"]),
          quantity: Number(r["Số lượng"] || 1),
          customFields: {},
        });
      });

      if (!imported.length) {
        notify(errors[0] || "Không có dòng danh mục hợp lệ");
        return;
      }

      const importedCodes = new Set(
        imported.map(a => warehouseKey(a.code))
      );

      const remaining = existing.filter(
        a => !importedCodes.has(warehouseKey(a.code))
      );

      setData({
        ...data,
        assets: [...imported, ...remaining],
        // QUAN TRỌNG: KHÔNG tạo warehouse transaction.
        warehouse: data.warehouse || [],
        activityLog: logAction(
          data.activityLog,
          `Import/cập nhật ${imported.length} danh mục tài sản`
        ),
      });

      notify(
        errors.length
          ? `Đã cập nhật ${imported.length} danh mục; ${errors.length} dòng lỗi. ${errors[0]}`
          : `Đã cập nhật ${imported.length} danh mục tài sản`
      );
    } catch (err) {
      console.error(err);
      notify("Không đọc được file Excel danh mục tài sản");
    }
  };

  reader.readAsArrayBuffer(file);
}

/* =========================
   IMPORT PHIẾU KHO
========================= */

function importWarehouseExcelV6({ file, data, setData, requireAdmin, notify }) {
  if (!requireAdmin()) return;

  const reader = new FileReader();

  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, {
        type: "array",
        cellDates: true,
      });

      const rows = XLSX.utils.sheet_to_json(
        wb.Sheets[wb.SheetNames[0]],
        { defval: "" }
      );

      if (!rows.length) {
        notify("File Excel không có dữ liệu");
        return;
      }

      const assets = data.assets || [];
      const projects = data.projects || [];
      const existing = data.warehouse || [];

      const byCode = Object.fromEntries(
        assets.map(a => [warehouseKey(a.code), a])
      );
      const byName = Object.fromEntries(
        assets.map(a => [warehouseKey(a.name), a])
      );

      const imported = [];
      const errors = [];

      rows.forEach((r, i) => {
        const rowNo = i + 2;

        const rawType = warehouseKey(r["Loại phiếu"]);
        const type =
          rawType.includes("xuất") || rawType === "xuat"
            ? "xuat"
            : "nhap";

        const code = warehouseKey(
          r["Mã hàng"] || r["Mã quản lý"]
        );
        const name = warehouseKey(
          r["Tên tài sản"] || r["Tên hàng"]
        );

        const asset = byCode[code] || byName[name];

        if (!asset) {
          errors.push(
            `Dòng ${rowNo}: không tìm thấy tài sản theo Mã hàng/Tên tài sản`
          );
          return;
        }

        const date = normalizeWarehouseDate(
          r["Ngày tháng"] || r["Ngày chứng từ"]
        );

        if (!date) {
          errors.push(`Dòng ${rowNo}: ngày không hợp lệ`);
          return;
        }

        const qty = Number(r["Số lượng"] || 0);
        if (!(qty > 0)) {
          errors.push(`Dòng ${rowNo}: số lượng phải lớn hơn 0`);
          return;
        }

        const unitCost = Number(r["Đơn giá"] || 0);
        const locationText = warehouseText(
          r["Kho/Công trình"] || r["Công trình"]
        );

        if (!locationText) {
          errors.push(`Dòng ${rowNo}: thiếu Kho/Công trình`);
          return;
        }

        const project = projects.find(
          p => warehouseKey(p.name) === warehouseKey(locationText)
        );

        const rawLocationType = warehouseKey(r["Loại địa điểm"]);
        const locationType =
          rawLocationType.includes("kho")
            ? "warehouse"
            : project
              ? "project"
              : "warehouse";

        const prefix = type === "nhap" ? "PN" : "PX";
        const dateKey = date.replaceAll("-", "");

        const seq =
          [...existing, ...imported].filter(w =>
            String(w.voucherNo || "").startsWith(`${prefix}-${dateKey}-`)
          ).length + 1;

        const voucherNo =
          warehouseText(r["Số phiếu"]) ||
          `${prefix}-${dateKey}-${String(seq).padStart(3, "0")}`;

        imported.push({
          id: uid("wh"),
          voucherNo,
          assetId: asset.id,
          type,
          quantity: qty,
          date,
          unitCost,
          total: qty * unitCost,
          unit: warehouseText(r["Đơn vị tính"]) || asset.unit || "Cái",
          receiver: warehouseText(r["Người giao/nhận"]),
          note: warehouseText(r["Ghi chú"]),

          // Luôn lấy master data từ Danh mục.
          category: asset.category || "Khác",
          assetGroup: asset.assetGroup || "Thiết bị chính",
          ownership: asset.ownership || "Công ty",

          locationType,
          locationName: locationText,
          warehouseName: locationType === "warehouse" ? locationText : "",
          projectId: project?.id || null,

          itemName: asset.name,
          itemCode: asset.code,
        });
      });

      if (errors.length) {
        notify(`Import bị dừng: ${errors.length} lỗi. ${errors[0]}`);
        return;
      }

      // Kiểm tra xuất vượt tồn tại đúng địa điểm và đúng ngày.
      for (const tx of imported) {
        if (tx.type !== "xuat") continue;

        const rowsAtLocation = [...existing, ...imported].filter(w =>
          w.assetId === tx.assetId &&
          warehouseKey(warehouseLocation(w, projects)) === warehouseKey(tx.locationName) &&
          normalizeWarehouseDate(w.date) <= tx.date
        );

        const available = rowsAtLocation.reduce(
          (sum, w) =>
            sum +
            (w.type === "nhap"
              ? Number(w.quantity) || 0
              : -(Number(w.quantity) || 0)),
          0
        );

        if (tx.quantity > available) {
          errors.push(
            `${tx.voucherNo}: xuất ${tx.quantity} ${tx.unit} vượt tồn tại ${tx.locationName}; tồn tính đến ${tx.date} chỉ còn ${available}`
          );
        }
      }

      if (errors.length) {
        notify(`Import bị dừng: ${errors.length} lỗi. ${errors[0]}`);
        return;
      }

      const txs = imported.map(tx => ({
        id: uid("tx"),
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
        transactions: [...txs, ...(data.transactions || [])],
        activityLog: logAction(
          data.activityLog,
          `Import ${imported.length} phiếu nhập/xuất kho`
        ),
      });

      notify(`Đã import ${imported.length} phiếu kho`);
    } catch (err) {
      console.error(err);
      notify("Không đọc được file Excel phiếu kho");
    }
  };

  reader.readAsArrayBuffer(file);
}

/* =========================
   PHẦN BẮT BUỘC TRONG APP
========================= */

/*
1) XÓA logic cũ trong importExcel(kind === "assets"):

   const wh = imported.map(...)

   và:

   warehouse: [...wh, ...data.warehouse]

   Thay bằng:

   importAssetCatalogV6({
     file,
     data,
     setData,
     requireAdmin,
     notify,
   });

2) Modal trong App đổi thành:

   {modal?.type === "warehouseIn" && (
     <WarehouseTxModalV6
       title="Lập phiếu nhập kho"
       fixedType="nhap"
       assets={data.assets}
       projects={data.projects}
       onClose={() => setModal(null)}
       onSubmit={f => {
         addWarehouseTx(f);
         setModal(null);
       }}
     />
   )}

   {modal?.type === "warehouseOut" && (
     <WarehouseTxModalV6
       title="Lập phiếu xuất kho"
       fixedType="xuat"
       assets={data.assets}
       projects={data.projects}
       onClose={() => setModal(null)}
       onSubmit={f => {
         addWarehouseTx(f);
         setModal(null);
       }}
     />
   )}

3) Trong phần render Warehouse dùng WarehouseViewV6.

4) Import phiếu kho:

   onImport={file =>
     importWarehouseExcelV6({
       file,
       data,
       setData,
       requireAdmin,
       notify,
     })
   }

5) Báo cáo tồn dùng:

   buildWarehouseBalancesV6({
     warehouse: data.warehouse,
     assets: data.assets,
     projects: data.projects,
     asOfDate: filter.asOfDate,
   });

6) KHÔNG chạy lại seed().
7) KHÔNG DROP TABLE.
8) KHÔNG DELETE app_data.
9) Không cần thay VITE_SUPABASE_URL / ANON KEY.
*/
