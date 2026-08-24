/*
 MYHL-QUẢN LÝ TÀI SẢN
 WAREHOUSE CORE V9 FIXED — 24/08/2026

 This file is a STANDALONE module. It does NOT depend on App.jsx helper names.
 Import it from App.jsx:
   import {
     WarehouseCoreV9,
     makeWarehouseReportV9,
     importAssetCatalogV9,
     importWarehouseExcelV9,
     downloadWarehouseTemplateV9,
   } from "./MYHL_WAREHOUSE_CORE_V9_FIXED";

 Important:
 - Do not paste this file into App.jsx.
 - Do not keep the old V8 block active.
 - The catalog is MASTER DATA only and never creates warehouse transactions.
 - Warehouse transactions alone create stock.
*/

import React, { useMemo, useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

const BRAND = "#C1272D";
const BRAND_SOFT = "#FBE8E8";
const PAPER = "#F7F4F3";
const SURFACE = "#FFFFFF";
const INK = "#241A19";
const MUTED = "#7C6E6D";
const BORDER = "#EEE1DF";
const DANGER = "#B42318";
const GOLD = "#C08A1E";

function text(v) { return String(v ?? "").trim(); }

function searchText(v) {
  return text(v)
    .toLocaleLowerCase("vi-VN")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d");
}

function dateKey(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date && !Number.isNaN(v.getTime())) {
    return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,"0")}-${String(v.getDate()).padStart(2,"0")}`;
  }
  if (typeof v === "number" && Number.isFinite(v)) {
    try {
      const d = XLSX.SSF.parse_date_code(v);
      if (d?.y && d?.m && d?.d) return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
    } catch {}
  }
  const s = text(v);
  let m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if (m) return `${m[3]}-${String(m[2]).padStart(2,"0")}-${String(m[1]).padStart(2,"0")}`;
  m = s.match(/^(\d{4})[\/.\-](\d{1,2})[\/.\-](\d{1,2})$/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2,"0")}-${String(m[3]).padStart(2,"0")}`;
  m = s.match(/^(\d{4}-\d{2}-\d{2})(?:T|\s)/);
  if (m) return m[1];
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : "";
}

function dateVN(v) {
  const d = dateKey(v);
  if (!d) return "";
  const [y,m,day] = d.split("-");
  return `${day}/${m}/${y}`;
}

function uid(prefix="id") {
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2,9)}`;
}

function locationOf(tx, projects=[]) {
  if (text(tx?.locationName)) return text(tx.locationName);
  if (tx?.projectId) {
    const p = projects.find(x => x.id === tx.projectId);
    if (p?.name) return text(p.name);
  }
  if (text(tx?.warehouseName)) return text(tx.warehouseName);
  return "Kho trung tâm";
}

function isIn(tx) {
  const t = searchText(tx?.type || tx?.loaiPhieu);
  return t === "nhap" || t.includes("nhap");
}

function isOut(tx) {
  const t = searchText(tx?.type || tx?.loaiPhieu);
  return t === "xuat" || t.includes("xuat");
}

function resolveAsset(tx, assets=[]) {
  if (tx?.assetId) {
    const a = assets.find(x => x.id === tx.assetId);
    if (a) return a;
  }
  const code = tx?.itemCode || tx?.code || tx?.maHang || tx?.maQuanLy;
  if (code) {
    const a = assets.find(x => searchText(x.code) === searchText(code));
    if (a) return a;
  }
  const name = tx?.itemName || tx?.name || tx?.tenTaiSan || tx?.tenHang;
  if (name) {
    const a = assets.find(x => searchText(x.name) === searchText(name));
    if (a) return a;
  }
  return null;
}

function resolveAssetFromRow(row, assets=[]) {
  const code = text(row?.["Mã hàng"] || row?.["Mã quản lý"] || row?.["Mã tài sản"] || row?.code);
  const name = text(row?.["Tên tài sản"] || row?.["Tên hàng"] || row?.["Tên thiết bị"] || row?.name);
  const byCode = code ? assets.find(a => searchText(a.code) === searchText(code)) : null;
  const byName = name ? assets.find(a => searchText(a.name) === searchText(name)) : null;
  if (byCode && byName && byCode.id !== byName.id) return { error: `Mã "${code}" và Tên "${name}" không cùng một tài sản.` };
  if (!byCode && !byName) return { error: `Không tìm thấy tài sản theo mã "${code}" hoặc tên "${name}".` };
  return { asset: byCode || byName };
}

export function getWarehouseStockV9({ warehouse=[], assets=[], projects=[], assetId, locationName, asOfDate }) {
  const cut = dateKey(asOfDate) || "9999-12-31";
  const loc = searchText(locationName);
  return warehouse.reduce((sum, tx) => {
    const d = dateKey(tx.date || tx.ngayThang);
    if (!d || d > cut) return sum;
    const a = resolveAsset(tx, assets);
    if (!a || a.id !== assetId) return sum;
    if (searchText(locationOf(tx, projects)) !== loc) return sum;
    const q = Number(tx.quantity ?? tx.soLuong ?? 0) || 0;
    return isIn(tx) ? sum + q : isOut(tx) ? sum - q : sum;
  }, 0);
}

export function makeWarehouseReportV9({ warehouse=[], assets=[], projects=[], filter={} }) {
  const cut = dateKey(filter.asOfDate) || "9999-12-31";
  const map = new Map();

  for (const tx of warehouse) {
    const d = dateKey(tx.date || tx.ngayThang);
    if (!d || d > cut) continue;
    const a = resolveAsset(tx, assets);
    if (!a) continue;
    const loc = locationOf(tx, projects);
    const key = `${a.id}¦${searchText(loc)}`;
    if (!map.has(key)) {
      map.set(key, {
        assetId:a.id, code:a.code || tx.itemCode || "", name:a.name || tx.itemName || "",
        location:loc, category:a.category || tx.category || "Khác",
        group:a.assetGroup || tx.assetGroup || "Thiết bị chính",
        ownership:a.ownership || tx.ownership || "Công ty",
        unit:a.unit || tx.unit || "Cái", inQty:0,outQty:0,inValue:0,outValue:0
      });
    }
    const r = map.get(key);
    const q = Number(tx.quantity ?? tx.soLuong ?? 0) || 0;
    const valueRaw = Number(tx.total ?? tx.thanhTien);
    const value = Number.isFinite(valueRaw) ? valueRaw : q * (Number(tx.unitCost ?? tx.unitPrice ?? tx.donGia ?? 0) || 0);
    if (isIn(tx)) { r.inQty += q; r.inValue += value; }
    else if (isOut(tx)) { r.outQty += q; r.outValue += value; }
  }

  let rows = [...map.values()].map(r => ({
    ...r, balanceQty:r.inQty-r.outQty, balanceValue:r.inValue-r.outValue
  }));

  if (filter.category) rows = rows.filter(r => searchText(r.category) === searchText(filter.category));
  if (filter.group) rows = rows.filter(r => searchText(r.group) === searchText(filter.group));
  if (filter.ownership) rows = rows.filter(r => searchText(r.ownership) === searchText(filter.ownership));
  if (filter.locationName) rows = rows.filter(r => searchText(r.location) === searchText(filter.locationName));
  if (filter.projectId) {
    const p = projects.find(x => x.id === filter.projectId);
    rows = rows.filter(r => p && searchText(r.location) === searchText(p.name));
  }
  rows = rows.filter(r => r.balanceQty > 0).sort((a,b) =>
    `${a.location}|${a.group}|${a.category}|${a.code}|${a.name}`.localeCompare(
      `${b.location}|${b.group}|${b.category}|${b.code}|${b.name}`, "vi"
    )
  );

  return {
    rows,
    summary: {
      qty: rows.reduce((s,r)=>s+r.balanceQty,0),
      value: rows.reduce((s,r)=>s+r.balanceValue,0),
      assets: rows.length,
      locations: new Set(rows.map(r=>r.location)).size
    },
    asOfDate: cut === "9999-12-31" ? "" : cut,
    asOfDateVN: cut === "9999-12-31" ? "" : dateVN(cut),
    headers:["Kho/Công trình","Loại tài sản","Nhóm tài sản","Nguồn gốc","Mã hàng","Tên tài sản","ĐVT","Nhập lũy kế","Xuất lũy kế","Tồn đến ngày","Giá trị tồn"],
    values: rows.map(r=>[r.location,r.category,r.group,r.ownership,r.code,r.name,r.unit,r.inQty,r.outQty,r.balanceQty,r.balanceValue])
  };
}

function Button({children,onClick,kind="default",disabled=false}) {
  const style = kind === "primary"
    ? {background:BRAND,color:"#fff",borderColor:BRAND}
    : kind === "danger"
    ? {background:"#fff",color:DANGER,borderColor:"#F1B5AE"}
    : {background:"#fff",color:INK,borderColor:BORDER};
  return <button type="button" disabled={disabled} onClick={onClick}
    style={{...style,border:"1px solid",borderRadius:8,padding:"8px 12px",fontSize:13,fontWeight:600,cursor:disabled?"not-allowed":"pointer",opacity:disabled?.55:1}}>
    {children}
  </button>;
}

function Input({value,onChange,placeholder,type="text"}) {
  return <input type={type} value={value ?? ""} onChange={e=>onChange(e.target.value)}
    placeholder={placeholder} style={{width:"100%",boxSizing:"border-box",border:`1px solid ${BORDER}`,borderRadius:8,padding:"9px 10px",fontSize:13,outline:"none",background:"#fff"}} />;
}

function Field({label,children}) {
  return <label style={{display:"block",marginBottom:12}}>
    <div style={{fontSize:12,color:MUTED,fontWeight:600,marginBottom:5}}>{label}</div>{children}
  </label>;
}

function SearchAsset({assets,value,onChange}) {
  const selected = assets.find(a=>a.id===value);
  const [q,setQ] = useState(selected ? `${selected.code||""} — ${selected.name||""}` : "");
  const [open,setOpen] = useState(false);
  useEffect(()=>{ if(selected) setQ(`${selected.code||""} — ${selected.name||""}`); else if(!value) setQ(""); },[value,selected?.id]);

  const results = useMemo(()=>{
    const s=searchText(q);
    if(!s) return assets.slice(0,25);
    return assets.map(a=>{
      const code=searchText(a.code), name=searchText(a.name), serial=searchText(a.serial);
      let score=0;
      if(code===s) score+=1000; if(name===s) score+=900;
      if(code.startsWith(s)) score+=500; if(name.startsWith(s)) score+=450;
      if(code.includes(s)) score+=80; if(name.includes(s)) score+=70; if(serial.includes(s)) score+=40;
      return {a,score};
    }).filter(x=>x.score).sort((a,b)=>b.score-a.score).slice(0,25).map(x=>x.a);
  },[assets,q]);

  return <div style={{position:"relative"}}>
    <Input value={q} onChange={v=>{setQ(v);setOpen(true);if(!v.trim())onChange("");}} placeholder="Gõ tên hoặc mã tài sản..." />
    {open && <div style={{position:"absolute",zIndex:1000,left:0,right:0,top:"100%",marginTop:3,maxHeight:260,overflow:"auto",background:"#fff",border:`1px solid ${BORDER}`,borderRadius:8,boxShadow:"0 12px 30px rgba(0,0,0,.12)"}}>
      {results.length ? results.map(a=><button key={a.id} type="button" onMouseDown={e=>e.preventDefault()} onClick={()=>{onChange(a.id);setQ(`${a.code||""} — ${a.name||""}`);setOpen(false);}}
        style={{display:"block",width:"100%",textAlign:"left",border:0,background:"#fff",padding:"9px 11px",cursor:"pointer"}}>
        <div style={{fontWeight:600,fontSize:13,color:INK}}>{a.code||"—"} — {a.name||"—"}</div>
        <div style={{fontSize:11,color:MUTED,marginTop:2}}>{a.category||"Chưa phân loại"} · {a.assetGroup||"Chưa có nhóm"} · {a.ownership||"Chưa có nguồn gốc"}</div>
      </button>) : <div style={{padding:12,fontSize:12,color:MUTED}}>Không tìm thấy theo mã hoặc tên.</div>}
    </div>}
  </div>;
}

function Modal({title,onClose,children}) {
  return <div style={{position:"fixed",inset:0,zIndex:2000,background:"rgba(0,0,0,.35)",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
    <div style={{width:"min(920px,100%)",maxHeight:"90vh",overflow:"auto",background:"#fff",borderRadius:12,boxShadow:"0 20px 60px rgba(0,0,0,.2)"}}>
      <div style={{padding:"16px 18px",borderBottom:`1px solid ${BORDER}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <strong style={{fontSize:17}}>{title}</strong><button type="button" onClick={onClose} style={{border:0,background:"transparent",fontSize:22,cursor:"pointer"}}>×</button>
      </div>
      <div style={{padding:18}}>{children}</div>
    </div>
  </div>;
}

export function createWarehouseTransactionV9({form,data,requireAdmin,notify,setData}) {
  if (requireAdmin && !requireAdmin()) return false;
  const assets=data.assets||[], projects=data.projects||[], warehouse=data.warehouse||[];
  const asset=assets.find(a=>a.id===form.assetId);
  if(!asset){notify?.("Vui lòng chọn tài sản từ danh sách.");return false;}
  const d=dateKey(form.date); if(!d){notify?.("Ngày chứng từ không hợp lệ.");return false;}
  const qty=Number(form.quantity); if(!(qty>0)){notify?.("Số lượng phải lớn hơn 0.");return false;}
  const loc=form.locationType==="project" ? projects.find(p=>p.id===form.projectId)?.name||"" : text(form.warehouseName);
  if(!loc){notify?.("Vui lòng chọn Kho/Công trình.");return false;}
  const out=form.type==="xuat";
  const prefix=out?"PX":"PN", dk=d.replaceAll("-","");
  const voucher=text(form.voucherNo)||`${prefix}-${dk}-${String(warehouse.filter(w=>String(w.voucherNo||"").startsWith(`${prefix}-${dk}-`)).length+1).padStart(3,"0")}`;
  if(warehouse.some(w=>searchText(w.voucherNo)===searchText(voucher))){notify?.(`Số phiếu ${voucher} đã tồn tại.`);return false;}
  const unit=asset.unit||"Cái", unitCost=Number(form.unitCost)||0;
  if(out){
    const stock=getWarehouseStockV9({warehouse,assets,projects,assetId:asset.id,locationName:loc,asOfDate:d});
    if(qty>stock){notify?.(`Không đủ tồn tại "${loc}". ${asset.code} — ${asset.name} hiện còn ${stock} ${unit}.`);return false;}
  }
  const tx={id:uid("wh"),voucherNo:voucher,assetId:asset.id,type:out?"xuat":"nhap",quantity:qty,date:d,unitCost,total:qty*unitCost,unit,
    receiver:text(form.receiver),note:text(form.note),category:asset.category||"Khác",assetGroup:asset.assetGroup||"Thiết bị chính",
    ownership:asset.ownership||"Công ty",locationType:form.locationType||"project",locationName:loc,
    warehouseName:form.locationType==="warehouse"?loc:"",projectId:form.locationType==="project"?form.projectId||null:null,
    itemName:asset.name,itemCode:asset.code,serial:asset.serial||""};
  const history={id:uid("tx"),warehouseTxId:tx.id,assetId:asset.id,type:out?"xuat_kho":"nhap_kho",date:d,title:`${out?"Xuất kho":"Nhập kho"} ${voucher}`,
    detail:`${asset.name} · ${loc} · ${tx.receiver||""}`,amount:tx.total};
  setData({...data,warehouse:[tx,...warehouse],transactions:[history,...(data.transactions||[])]});
  notify?.(`Đã lập ${out?"phiếu xuất":"phiếu nhập"} ${voucher}`);
  return true;
}

export function deleteWarehouseTransactionsV9({ids=[],data,requireAdmin,notify,setData}) {
  if(requireAdmin && !requireAdmin()) return false;
  const selected=new Set(ids.filter(Boolean)), warehouse=data.warehouse||[];
  const rows=warehouse.filter(w=>selected.has(w.id));
  if(!rows.length){notify?.("Chưa chọn phiếu cần xóa.");return false;}
  const vouchers=[...new Set(rows.map(w=>text(w.voucherNo)).filter(Boolean))];
  if(!window.confirm(`Xóa ${rows.length} dòng thuộc ${vouchers.length} số phiếu?\n\n${vouchers.join(", ")}\n\nThao tác này sẽ thay đổi báo cáo nhập - xuất - tồn.`)) return false;
  const idsSet=new Set(rows.map(w=>w.id)), voucherKeys=new Set(vouchers.map(searchText));
  setData({...data,
    warehouse:warehouse.filter(w=>!selected.has(w.id)),
    transactions:(data.transactions||[]).filter(t=>!(t.warehouseTxId&&idsSet.has(t.warehouseTxId)) && ![...voucherKeys].some(k=>k&&searchText(t.title).includes(k)))
  });
  notify?.(`Đã xóa ${vouchers.length} phiếu kho.`);
  return true;
}

function TxModal({assets,projects,data,onClose,onSaved,type,requireAdmin,notify,setData}) {
  const [f,setF]=useState({assetId:"",voucherNo:"",date:new Date().toLocaleDateString("en-CA"),quantity:1,unitCost:0,receiver:"",locationType:"project",projectId:"",warehouseName:"",note:"",type});
  const a=assets.find(x=>x.id===f.assetId)||{};
  const set=(k,v)=>setF(p=>({...p,[k]:v}));
  const total=(Number(f.quantity)||0)*(Number(f.unitCost)||0);
  return <Modal title={type==="nhap"?"Lập phiếu nhập kho":"Lập phiếu xuất kho"} onClose={onClose}>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      <Field label="Số phiếu"><Input value={f.voucherNo} onChange={v=>set("voucherNo",v)} placeholder={type==="nhap"?"PN-20260824-001":"PX-20260824-001"} /></Field>
      <Field label="Ngày tháng"><Input type="date" value={f.date} onChange={v=>set("date",v)} /></Field>
      <Field label="Tên tài sản / Mã hàng"><SearchAsset assets={assets} value={f.assetId} onChange={id=>{const x=assets.find(z=>z.id===id)||{};setF(p=>({...p,assetId:id,unitCost:Number(x.cost)||0}));}} /></Field>
      <Field label="Mã hàng tự động"><Input value={a.code||""} onChange={()=>{}} /></Field>
      <Field label="Loại tài sản"><Input value={a.category||""} onChange={()=>{}} /></Field>
      <Field label="Nhóm tài sản"><Input value={a.assetGroup||""} onChange={()=>{}} /></Field>
      <Field label="Nguồn gốc"><Input value={a.ownership||""} onChange={()=>{}} /></Field>
      <Field label="Đơn vị tính"><Input value={a.unit||"Cái"} onChange={()=>{}} /></Field>
      <Field label="Kho / Công trình">
        <select value={f.locationType==="project"?f.projectId:"__warehouse__"} onChange={e=>e.target.value==="__warehouse__"?setF(p=>({...p,locationType:"warehouse",projectId:""})):setF(p=>({...p,locationType:"project",projectId:e.target.value}))}
          style={{width:"100%",border:`1px solid ${BORDER}`,borderRadius:8,padding:9,fontSize:13}}>
          <option value="">-- Chọn công trình --</option>
          {projects.map(p=><option key={p.id} value={p.id}>{p.name}</option>)}
          <option value="__warehouse__">Kho trung tâm / kho khác</option>
        </select>
      </Field>
      {f.locationType==="warehouse" && <Field label="Tên kho"><Input value={f.warehouseName} onChange={v=>set("warehouseName",v)} placeholder="Kho trung tâm" /></Field>}
      <Field label="Người giao / nhận"><Input value={f.receiver} onChange={v=>set("receiver",v)} /></Field>
      <Field label="Số lượng"><Input type="number" value={f.quantity} onChange={v=>set("quantity",v)} /></Field>
      <Field label="Đơn giá"><Input type="number" value={f.unitCost} onChange={v=>set("unitCost",v)} /></Field>
      <Field label="Thành tiền"><Input value={total.toLocaleString("vi-VN")} onChange={()=>{}} /></Field>
      <div style={{gridColumn:"1 / -1"}}><Field label="Ghi chú"><textarea value={f.note} onChange={e=>set("note",e.target.value)} rows={3} style={{width:"100%",boxSizing:"border-box",border:`1px solid ${BORDER}`,borderRadius:8,padding:9}} /></Field></div>
    </div>
    <div style={{display:"flex",justifyContent:"flex-end",gap:8}}><Button onClick={onClose}>Hủy</Button>
      <Button kind="primary" disabled={!f.assetId||Number(f.quantity)<=0||(!f.projectId&&f.locationType==="project")||(!f.warehouseName&&f.locationType==="warehouse")}
        onClick={()=>{const ok=createWarehouseTransactionV9({form:f,data,setData,requireAdmin,notify});if(ok){onSaved?.();}}}>Lưu phiếu</Button>
    </div>
  </Modal>;
}

function TxTable({warehouse,assets,projects,data,type,requireAdmin,notify,setData}) {
  const [q,setQ]=useState(""),[loc,setLoc]=useState(""),[selected,setSelected]=useState([]);
  const rows=useMemo(()=>{
    const s=searchText(q);
    return warehouse.filter(w=>(!type||(type==="nhap"?isIn(w):isOut(w)))&&(!loc||searchText(locationOf(w,projects))===searchText(loc))&&(!s||[w.voucherNo,w.itemCode,w.itemName,w.receiver,locationOf(w,projects)].map(searchText).join(" ").includes(s)))
      .sort((a,b)=>(dateKey(b.date)||"").localeCompare(dateKey(a.date)||"")||String(b.id).localeCompare(String(a.id)));
  },[warehouse,assets,projects,type,q,loc]);
  const locations=[...new Set(warehouse.map(w=>locationOf(w,projects)))].sort((a,b)=>searchText(a).localeCompare(searchText(b),"vi"));
  const all=rows.length>0&&rows.every(r=>selected.includes(r.id));
  const toggleAll=()=>setSelected(all?selected.filter(id=>!rows.some(r=>r.id===id)):[...new Set([...selected,...rows.map(r=>r.id)])]);
  return <div>
    <div style={{display:"flex",gap:8,marginBottom:10,flexWrap:"wrap"}}>
      <div style={{flex:1,minWidth:260}}><Input value={q} onChange={setQ} placeholder="Tìm số phiếu, mã hàng, tên tài sản, công trình..." /></div>
      <select value={loc} onChange={e=>setLoc(e.target.value)} style={{minWidth:220,border:`1px solid ${BORDER}`,borderRadius:8,padding:9}}><option value="">Tất cả Kho/Công trình</option>{locations.map(x=><option key={x}>{x}</option>)}</select>
      <Button kind="danger" disabled={!selected.length} onClick={()=>{const ok=deleteWarehouseTransactionsV9({ids:selected,data,setData,requireAdmin,notify});if(ok)setSelected([]);}}>Xóa ({selected.length})</Button>
    </div>
    <div style={{overflow:"auto",border:`1px solid ${BORDER}`,borderRadius:10,background:"#fff"}}>
      <table style={{width:"100%",minWidth:1050,borderCollapse:"collapse"}}><thead><tr style={{background:BRAND_SOFT}}>
        {["","Số phiếu","Ngày","Mã hàng","Tên tài sản","Kho/Công trình","Loại","SL","ĐVT","Người giao/nhận"].map((h,i)=><th key={i} style={{padding:9,textAlign:"left",fontSize:12,color:MUTED}}>{i===0?<input type="checkbox" checked={all} onChange={toggleAll}/>:h}</th>)}
      </tr></thead><tbody>{rows.map(w=>{const a=resolveAsset(w,assets);return <tr key={w.id} style={{borderTop:`1px solid ${BORDER}`}}>
        <td style={{padding:9}}><input type="checkbox" checked={selected.includes(w.id)} onChange={()=>setSelected(s=>s.includes(w.id)?s.filter(x=>x!==w.id):[...s,w.id])}/></td>
        <td style={{padding:9,fontFamily:"monospace"}}>{w.voucherNo||"—"}</td><td style={{padding:9}}>{dateVN(w.date)}</td><td style={{padding:9}}>{a?.code||w.itemCode||"—"}</td><td style={{padding:9}}>{a?.name||w.itemName||"—"}</td><td style={{padding:9}}>{locationOf(w,projects)}</td><td style={{padding:9}}>{isIn(w)?"Nhập kho":"Xuất kho"}</td><td style={{padding:9}}>{Number(w.quantity||0).toLocaleString("vi-VN")}</td><td style={{padding:9}}>{w.unit||a?.unit||"Cái"}</td><td style={{padding:9}}>{w.receiver||"—"}</td>
      </tr>})}{!rows.length&&<tr><td colSpan={10} style={{padding:35,textAlign:"center",color:MUTED}}>Không có phiếu phù hợp.</td></tr>}</tbody></table>
    </div>
  </div>;
}

export async function readExcelRows(file) {
  const buffer=await file.arrayBuffer();
  const wb=XLSX.read(buffer,{type:"array",cellDates:true});
  const sheet=wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet,{defval:""});
}

export async function importAssetCatalogV9({file,data,setData,requireAdmin,notify}) {
  if(requireAdmin && !requireAdmin()) return false;
  try{
    const rows=await readExcelRows(file), existing=data.assets||[], byCode=new Map(existing.map(a=>[searchText(a.code),a]));
    const imported=[],seen=new Set(),errors=[];
    rows.forEach((r,i)=>{
      const code=text(r["Mã quản lý"]||r["Mã hàng"]||r["Mã tài sản"]||r.code), name=text(r["Tên tài sản"]||r["Tên hàng"]||r["Tên thiết bị"]||r.name);
      if(!code||!name){errors.push(`Dòng ${i+2}: thiếu mã hoặc tên.`);return;}
      const k=searchText(code); if(seen.has(k)){errors.push(`Dòng ${i+2}: mã ${code} trùng.`);return;} seen.add(k);
      const old=byCode.get(k);
      imported.push({...old,id:old?.id||uid("as"),code,name,
        category:text(r["Loại tài sản"]||r["Loại"]||old?.category||"Khác"),
        assetGroup:text(r["Nhóm tài sản"]||r["Nhóm"]||old?.assetGroup||"Thiết bị chính"),
        ownership:text(r["Nguồn gốc"]||old?.ownership||"Công ty"),
        unit:text(r["Đơn vị tính"]||r["ĐVT"]||old?.unit||"Cái"),
        serial:text(r["Serial"]||r["Số serial"]||old?.serial||""),
        cost:Number(r["Đơn giá"]||r["Nguyên giá"]||old?.cost||0)||0});
    });
    if(!imported.length){notify?.(errors[0]||"Không có dòng hợp lệ.");return false;}
    const codes=new Set(imported.map(a=>searchText(a.code)));
    setData({...data,assets:[...imported,...existing.filter(a=>!codes.has(searchText(a.code)))],
      warehouse:Array.isArray(data.warehouse)?data.warehouse:[],transactions:Array.isArray(data.transactions)?data.transactions:[]});
    notify?.(`Đã cập nhật ${imported.length} danh mục. Không tạo phiếu nhập/xuất.`);
    return true;
  }catch(e){console.error(e);notify?.("Không đọc được file Excel danh mục.");return false;}
}

export async function importWarehouseExcelV9({file,data,setData,requireAdmin,notify}) {
  if(requireAdmin && !requireAdmin()) return false;
  try{
    const rows=await readExcelRows(file), assets=data.assets||[], projects=data.projects||[], existing=data.warehouse||[], imported=[],errors=[];
    rows.forEach((r,i)=>{
      const row=i+2, raw=searchText(r["Loại phiếu"]||r["Loại chứng từ"]||"Nhập kho"), type=raw.includes("xuat")?"xuat":"nhap";
      const d=dateKey(r["Ngày tháng"]||r["Ngày chứng từ"]||r["Ngày"]); if(!d){errors.push(`Dòng ${row}: ngày không hợp lệ.`);return;}
      const resolved=resolveAssetFromRow(r,assets); if(resolved.error){errors.push(`Dòng ${row}: ${resolved.error}`);return;}
      const a=resolved.asset,q=Number(r["Số lượng"]??r["SL"]??0); if(!(q>0)){errors.push(`Dòng ${row}: số lượng phải > 0.`);return;}
      const loc=text(r["Kho/Công trình"]||r["Công trình"]||r["Tên kho"]||r["Kho"]); if(!loc){errors.push(`Dòng ${row}: thiếu Kho/Công trình.`);return;}
      const p=projects.find(x=>searchText(x.name)===searchText(loc));
      const locationType=searchText(r["Loại địa điểm"]||"").includes("cong trinh")?"project":searchText(r["Loại địa điểm"]||"").includes("kho")?"warehouse":p?"project":"warehouse";
      const voucher=text(r["Số phiếu"])||`${type==="nhap"?"PN":"PX"}-${d.replaceAll("-","")}-${String(existing.length+imported.length+1).padStart(3,"0")}`;
      if(existing.some(w=>searchText(w.voucherNo)===searchText(voucher))||imported.some(w=>searchText(w.voucherNo)===searchText(voucher))){errors.push(`Dòng ${row}: số phiếu ${voucher} bị trùng.`);return;}
      imported.push({id:uid("wh"),voucherNo:voucher,assetId:a.id,type,quantity:q,date:d,unitCost:Number(r["Đơn giá"]??r["Đơn giá nhập"]??r["Đơn giá xuất"]??0)||0,
        total:q*(Number(r["Đơn giá"]??r["Đơn giá nhập"]??r["Đơn giá xuất"]??0)||0),unit:text(r["Đơn vị tính"]||r["ĐVT"])||a.unit||"Cái",
        receiver:text(r["Người giao/nhận"]),note:text(r["Ghi chú"]),category:a.category||"Khác",assetGroup:a.assetGroup||"Thiết bị chính",ownership:a.ownership||"Công ty",
        locationType,locationName:loc,warehouseName:locationType==="warehouse"?loc:"",projectId:locationType==="project"?p?.id||null:null,itemName:a.name,itemCode:a.code,serial:a.serial||""});
    });
    if(errors.length){notify?.(`Import bị dừng: ${errors.length} lỗi. ${errors[0]}`);return false;}
    const histories=imported.map(tx=>({id:uid("tx"),warehouseTxId:tx.id,assetId:tx.assetId,type:tx.type==="nhap"?"nhap_kho":"xuat_kho",date:tx.date,title:`${tx.type==="nhap"?"Nhập kho":"Xuất kho"} ${tx.voucherNo}`,detail:`${tx.itemName} · ${tx.locationName} · ${tx.receiver||""}`,amount:tx.total}));
    setData({...data,warehouse:[...imported,...existing],transactions:[...histories,...(data.transactions||[])]});
    notify?.(`Đã import ${imported.length} phiếu kho.`);
    return true;
  }catch(e){console.error(e);notify?.("Không đọc được file Excel phiếu kho.");return false;}
}

export function downloadWarehouseTemplateV9() {
  const headers=["Loại phiếu","Số phiếu","Ngày tháng","Tên tài sản","Mã hàng","Kho/Công trình","Loại địa điểm","Loại tài sản","Nhóm tài sản","Nguồn gốc","Người giao/nhận","Số lượng","Đơn vị tính","Đơn giá","Thành tiền","Ghi chú"];
  const rows=[["Nhập kho","PN-20260824-001","2026-08-24","Máy xúc 01","MX-01","Cao Xà Lá - Thanh Xuân","Công trình","Máy xúc","Thiết bị chính","Thuê","Nguyễn Văn A",1,"Cái",0,0,""]];
  const ws=XLSX.utils.aoa_to_sheet([headers,...rows]); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Phiếu kho"); XLSX.writeFile(wb,"Mau_Import_Phieu_Nhap_Xuat_Kho_V9.xlsx");
}

export function WarehouseReportV9({warehouse,assets,projects,filter,onFilterChange}) {
  const report=useMemo(()=>makeWarehouseReportV9({warehouse,assets,projects,filter}),[warehouse,assets,projects,filter]);
  return <div style={{background:SURFACE,border:`1px solid ${BORDER}`,borderRadius:10,overflow:"hidden"}}>
    <div style={{padding:14,borderBottom:`1px solid ${BORDER}`,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
      <strong>Báo cáo nhập xuất tồn</strong>
      <span style={{fontSize:12,color:MUTED}}>Tồn đến {report.asOfDateVN||"ngày hiện tại"}</span>
      <span style={{marginLeft:"auto",fontSize:12,color:BRAND,fontWeight:700}}>Tồn: {report.summary.qty.toLocaleString("vi-VN")}</span>
    </div>
    <div style={{overflow:"auto"}}>
      <table style={{width:"100%",minWidth:1100,borderCollapse:"collapse"}}><thead><tr style={{background:BRAND_SOFT}}>{report.headers.map(h=><th key={h} style={{padding:9,textAlign:"left",fontSize:12,color:MUTED}}>{h}</th>)}</tr></thead>
      <tbody>{report.rows.map(r=><tr key={`${r.assetId}-${r.location}`} style={{borderTop:`1px solid ${BORDER}`}}>{[r.location,r.category,r.group,r.ownership,r.code,r.name,r.unit,r.inQty,r.outQty,r.balanceQty,r.balanceValue].map((v,i)=><td key={i} style={{padding:9,fontSize:12}}>{typeof v==="number"?(i===10?Math.round(v).toLocaleString("vi-VN"):v.toLocaleString("vi-VN")):v}</td>)}</tr>)}{!report.rows.length&&<tr><td colSpan={11} style={{padding:35,textAlign:"center",color:MUTED}}>Không có thiết bị tồn theo điều kiện lọc.</td></tr>}</tbody></table>
    </div>
  </div>;
}

export function WarehouseCoreV9({data,setData,requireAdmin,notify,logAction}) {
  const [tab,setTab]=useState("nhap"),[modal,setModal]=useState(null),[filter,setFilter]=useState({asOfDate:new Date().toLocaleDateString("en-CA"),category:"",group:"",ownership:"",locationName:"",projectId:""});
  const projects=data.projects||[],assets=data.assets||[],warehouse=data.warehouse||[];
  const report=useMemo(()=>makeWarehouseReportV9({warehouse,assets,projects,filter}),[warehouse,assets,projects,filter]);
  const categories=[...new Set(assets.map(a=>a.category).filter(Boolean))],groups=[...new Set(assets.map(a=>a.assetGroup).filter(Boolean))],owners=[...new Set(assets.map(a=>a.ownership).filter(Boolean))];
  const locations=[...new Set(warehouse.map(w=>locationOf(w,projects)))].sort((a,b)=>searchText(a).localeCompare(searchText(b),"vi"));
  const importFile=async(e,kind)=>{const f=e.target.files?.[0];if(!f)return; if(kind==="assets") await importAssetCatalogV9({file:f,data,setData,requireAdmin,notify}); else await importWarehouseExcelV9({file:f,data,setData,requireAdmin,notify}); e.target.value="";};
  return <div style={{background:PAPER,minHeight:"100%",color:INK}}>
    <div style={{display:"flex",justifyContent:"space-between",gap:10,alignItems:"center",marginBottom:14,flexWrap:"wrap"}}>
      <div><h2 style={{margin:"0 0 3px",fontSize:22}}>Kho — Nhập / Xuất / Tồn</h2><div style={{fontSize:12,color:MUTED}}>Theo dõi thiết bị theo Kho/Công trình</div></div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}><Button onClick={downloadWarehouseTemplateV9}>Tải mẫu Excel</Button><label style={{...btnLabel()}}><input hidden type="file" accept=".xlsx,.xls,.csv" onChange={e=>importFile(e,"warehouse")} />Import phiếu Excel</label></div>
    </div>
    <div style={{display:"flex",borderBottom:`1px solid ${BORDER}`,gap:4,marginBottom:14}}>
      {[["nhap","Phiếu nhập kho"],["xuat","Phiếu xuất kho"],["report","Báo cáo nhập xuất tồn"]].map(([id,label])=><button key={id} type="button" onClick={()=>setTab(id)} style={{border:0,borderBottom:tab===id?`3px solid ${BRAND}`:"3px solid transparent",background:"transparent",padding:"10px 14px",fontWeight:tab===id?700:500,color:tab===id?BRAND:INK}}>{label}</button>)}
    </div>
    {tab!=="report"&&<div style={{display:"flex",justifyContent:"flex-end",marginBottom:10}}><Button kind="primary" onClick={()=>setModal(tab)}>+ Lập phiếu {tab==="nhap"?"nhập":"xuất"}</Button></div>}
    {tab==="nhap"&&<TxTable warehouse={warehouse} assets={assets} projects={projects} data={data} type="nhap" requireAdmin={requireAdmin} notify={notify} setData={setData}/>}
    {tab==="xuat"&&<TxTable warehouse={warehouse} assets={assets} projects={projects} data={data} type="xuat" requireAdmin={requireAdmin} notify={notify} setData={setData}/>}
    {tab==="report"&&<div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:10,marginBottom:12}}>
        <Field label="Loại tài sản"><select value={filter.category} onChange={e=>setFilter(f=>({...f,category:e.target.value}))} style={sel()}><option value="">Tất cả loại</option>{categories.map(x=><option key={x}>{x}</option>)}</select></Field>
        <Field label="Nhóm tài sản"><select value={filter.group} onChange={e=>setFilter(f=>({...f,group:e.target.value}))} style={sel()}><option value="">Tất cả nhóm</option>{groups.map(x=><option key={x}>{x}</option>)}</select></Field>
        <Field label="Nguồn gốc"><select value={filter.ownership} onChange={e=>setFilter(f=>({...f,ownership:e.target.value}))} style={sel()}><option value="">Tất cả nguồn gốc</option>{owners.map(x=><option key={x}>{x}</option>)}</select></Field>
        <Field label="Kho/Công trình"><select value={filter.locationName} onChange={e=>setFilter(f=>({...f,locationName:e.target.value}))} style={sel()}><option value="">Tất cả</option>{locations.map(x=><option key={x}>{x}</option>)}</select></Field>
        <Field label="Tồn đến ngày"><Input type="date" value={filter.asOfDate} onChange={v=>setFilter(f=>({...f,asOfDate:v}))}/></Field>
      </div>
      <WarehouseReportV9 warehouse={warehouse} assets={assets} projects={projects} filter={filter}/>
    </div>}
    {modal&&<TxModal assets={assets} projects={projects} data={data} type={modal} requireAdmin={requireAdmin} notify={notify} setData={setData} onClose={()=>setModal(null)} onSaved={()=>setModal(null)}/>}
  </div>;
}

function btnLabel(){return {display:"inline-flex",alignItems:"center",border:`1px solid ${BORDER}`,borderRadius:8,padding:"8px 12px",fontSize:13,fontWeight:600,cursor:"pointer",background:"#fff"}}
function sel(){return {width:"100%",boxSizing:"border-box",border:`1px solid ${BORDER}`,borderRadius:8,padding:9,fontSize:13,background:"#fff"}}
