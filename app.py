# Streamlit web app cho "Cholimex Display Checker"
from __future__ import annotations
import io, os, re, json
from typing import Dict, List, Tuple, Optional
from datetime import datetime

import streamlit as st
import pandas as pd

# =========================
# DEFAULT CONFIG (fallback)
# =========================
DEFAULT_CONFIG = {
    "muc_toi_thieu": {
        "NMCD": 150000, "DHLM": 100000, "KOS_XXTG": 300000, "LTLKC": 80000,
        "GVIG": 300000, "GVIG_BMTR": 300000, "KOS_XXTG_BS": 200000,
        "CAKOS": 50000, "XBM_MN": 36000, "XBM_MB": 36000
    },
    "program_names": {
        "NMCD": "Trưng bày Nước mắm Cholimex 30, 35, 40 độ đạm 500ml + 750ml",
        "DHLM": "Trưng bày Dầu hào 820g, Nước tương Lên men 700ml",
        "LTLKC": "Trưng bày Xốt Lẩu thái 280g & Xốt Lẩu kim chi 280g",
        "KOS_XXTG": "Trưng bày cá KOS và Xúc xích - Miền Nam",
        "KOS_XXTG_BS": "Trưng bày cá KOS và Xúc xích - Miền Bắc & Bắc Miền Trung",
        "XBM_MN": "Trưng bày Xe Bánh Mì - Miền Nam",
        "XBM_MB": "Trưng bày Xe Bánh Mì - Miền Bắc",
        "CAKOS": "Trưng bày cá KOS",
        "GVIG": "Trưng bày Gia vị gói - Miền Bắc",
        "GVIG_BMTR": "Trưng bày Gia vị gói - Bắc Miền Trung"
    },
    "region_map": {
        "HCME": ["HCM", "MD"],
        "MTRUNG": ["MTR", "MB_MT3"],
        "MTAY": ["MTA"],
        "MBAC": ["MB"],
        "TOAN_QUOC": "ALL"
    },
    # map riêng cho Xe Bánh Mì
    "xbm_map": {"M70": "XBM_MN", "M110": "XBM_MN", "M80": "XBM_MB", "M120": "XBM_MB"}
}

# ============ CONFIG LOADER ============
def _load_json_text(text: str) -> Optional[dict]:
    try:
        return json.loads(text)
    except Exception:
        return None

@st.cache_data(show_spinner=False)
def load_config(overrides: dict | None = None) -> Dict:
    cfg = DEFAULT_CONFIG.copy()
    if overrides:
        for k in ["muc_toi_thieu","program_names","region_map","xbm_map"]:
            if k in overrides and isinstance(overrides[k], dict):
                cfg[k] = overrides[k]
    return cfg

# =============== UTILITIES ===============
def parse_stage_value(giai_doan: str) -> Tuple[int,int,str]:
    """
    Chuyển 'Tháng 11/2025' → (2025, 11, 'Tháng 11/2025'), dùng để sort.
    Nếu không bắt được, trả (0,0,raw).
    """
    if not isinstance(giai_doan, str): return (0,0,str(giai_doan))
    m = re.search(r"(\d{1,2}).*?(\d{4})", giai_doan)
    if not m: return (0,0,giai_doan)
    mm, yy = int(m.group(1)), int(m.group(2))
    return (yy, mm, giai_doan)

def fmt_money(x):
    try:
        return f"{int(round(float(x))):,}".replace(",", ".")
    except Exception:
        return x

# =============== CORE ===============
def xu_ly_file(file: bytes, muc_toi_thieu: Dict[str, float], xbm_map: Dict[str,str]):
    df = pd.read_excel(io.BytesIO(file), header=1, dtype={"Mã khách hàng": str, "Mã NPP": str})
    cols_in = ["Mức đăng ký","Miền","Vùng","Mã NPP","Tên NPP","Giai đoạn","Mã NVBH","Tên NVBH",
               "Mã khách hàng","Tên khách hàng","Thứ bán hàng","Tuyến","Số suất đăng kí","Doanh số tích lũy hiện tại"]
    df = df[[c for c in cols_in if c in df.columns]].copy()

    df.rename(columns={
        "Mức đăng ký":"MucDK","Miền":"Mien","Vùng":"Vung","Mã NPP":"MaNPP","Tên NPP":"TenNPP",
        "Giai đoạn":"GiaiDoan","Mã NVBH":"MaNVBH","Tên NVBH":"TenNVBH",
        "Mã khách hàng":"MaKH","Tên khách hàng":"TenKH",
        "Thứ bán hàng":"ThuBanHang","Tuyến":"Tuyen",
        "Số suất đăng kí":"SoSuat","Doanh số tích lũy hiện tại":"DoanhSo"
    }, inplace=True)

    if "Tuyen" not in df.columns: df["Tuyen"] = None
    if "ThuBanHang" not in df.columns: df["ThuBanHang"] = None

    muc_map = df["MucDK"].astype(str).str.strip().map(xbm_map).fillna(df["MucDK"].astype(str).str.strip())
    base = muc_map.map(muc_toi_thieu).fillna(0).astype(float)
    df["NguongToiThieu"] = base * pd.to_numeric(df["SoSuat"], errors="coerce").fillna(0).astype(float)

    giai_doan = str(df["GiaiDoan"].iloc[0]).strip()
    df[f"SoSuat_{giai_doan}"] = df["SoSuat"]
    df[f"DoanhSo_{giai_doan}"] = df["DoanhSo"]
    df[f"Nguong_{giai_doan}"] = df["NguongToiThieu"]
    return df, giai_doan

def xu_ly_chuong_trinh(file_t1: bytes, file_t2: bytes, muc_toi_thieu, program_names, xbm_map,
                        file_t0: bytes | None = None,
                        filter_ketqua: Optional[set] = None,
                        filter_tuyen_tokens: Optional[List[str]] = None):
    df1, g1 = xu_ly_file(file_t1, muc_toi_thieu, xbm_map)
    df2, g2 = xu_ly_file(file_t2, muc_toi_thieu, xbm_map)

    new_in_T1_keys = set()
    if file_t0:
        df0, g0 = xu_ly_file(file_t0, muc_toi_thieu, xbm_map)
        keys_t0 = set(zip(df0["MaKH"], df0["MucDK"]))
        keys_t1 = set(zip(df1["MaKH"], df1["MucDK"]))
        new_in_T1_keys = keys_t1 - keys_t0
    else:
        df0, g0 = None, None

    df = pd.merge(df1, df2, on=["MaKH","MucDK"], how="outer", suffixes=("_T1","_T2"))
    if df0 is not None:
        df = df.merge(df0[["MaKH", f"SoSuat_{g0}", f"DoanhSo_{g0}"]], on="MaKH", how="left")

    for col in [f"SoSuat_{g1}", f"SoSuat_{g2}", f"DoanhSo_{g1}", f"DoanhSo_{g2}", f"Nguong_{g1}", f"Nguong_{g2}"]:
        if col in df.columns: df[col] = df[col].fillna(0)

    def xet(row):
        ds1, ds2 = row.get(f"DoanhSo_{g1}",0) or 0, row.get(f"DoanhSo_{g2}",0) or 0
        ss1, ss2 = row.get(f"SoSuat_{g1}",0) or 0, row.get(f"SoSuat_{g2}",0) or 0
        n1, n2 = row.get(f"Nguong_{g1}",0) or 0, row.get(f"Nguong_{g2}",0) or 0
        key = (row.get("MaKH"), row.get("MucDK"))
        if ss1 > 0 and ss2 == 0: return "XOA", "Tháng trước có tham gia, tháng sau không tham gia"
        if ss1 > 0 and key in new_in_T1_keys: return "Đạt", "Khách mới tháng trước (DS xét chu kỳ 11/T0→10/T1)"
        if ss1 == 0 and ss2 > 0: return "Không xét", "Khách hàng mới tháng sau (không xét kết quả kỳ này)"
        if ss2 > ss1 > 0: return "Đạt", f"Nâng suất {int(ss1)}→{int(ss2)}"
        if ss2 < ss1:
            if (ds1 >= n1) or (ds2 >= n2): return "Đạt", f"Giảm suất {int(ss1)}→{int(ss2)} (đủ 1 trong 2)"
            else: return "Không đạt", f"Giảm suất {int(ss1)}→{int(ss2)} (thiếu)"
        if (ds1 >= n1) or (ds2 >= n2): return "Đạt",""
        return "Không đạt","Thiếu"

    df[["KetQua","GhiChu"]] = df.apply(lambda r: pd.Series(xet(r)), axis=1)

    df_removed = df[df["KetQua"]=="XOA"].copy()
    df_final  = df[df["KetQua"]!="XOA"].copy()

    # lọc theo kết quả
    if filter_ketqua is not None:
        df_final = df_final[df_final["KetQua"].isin(filter_ketqua)]

    # lọc theo 'Thứ bán hàng' (fallback 'Tuyến') — KHÔNG xuất cột 'Tuyến'
    route_col = "ThuBanHang_T2" if "ThuBanHang_T2" in df_final.columns else ("Tuyen_T2" if "Tuyen_T2" in df_final.columns else None)
    if filter_tuyen_tokens and route_col:
        toks = [t.lower() for t in filter_tuyen_tokens if t]
        df_final = df_final[df_final[route_col].astype(str).str.lower().apply(lambda s: any(t in s for t in toks))]

    # chọn cột xuất ra
    cols_out = [
        "MucDK","Mien_T2","Vung_T2","MaNPP_T2","TenNPP_T2","MaNVBH_T2","TenNVBH_T2",
        "MaKH","TenKH_T2","ThuBanHang_T2",
        f"SoSuat_{g1}", f"SoSuat_{g2}",
        f"DoanhSo_{g1}", f"DoanhSo_{g2}",
        f"Nguong_{g2}", "KetQua","GhiChu"
    ]
    if df0 is not None:
        cols_out.insert(10, f"SoSuat_{g0}")
        cols_out.insert(11, f"DoanhSo_{g0}")

    rename = {
        "MucDK":"Mức đăng ký","Mien_T2":"Miền","Vung_T2":"Vùng",
        "MaNPP_T2":"Mã NPP","TenNPP_T2":"Tên NPP","MaNVBH_T2":"Mã NVBH","TenNVBH_T2":"Tên NVBH",
        "MaKH":"Mã khách hàng","TenKH_T2":"Tên khách hàng","ThuBanHang_T2":"Thứ bán hàng",
        f"SoSuat_{g1}":f"Số suất đăng ký {g1}", f"SoSuat_{g2}":f"Số suất đăng ký {g2}",
        f"DoanhSo_{g1}":f"Doanh số tích lũy {g1}", f"DoanhSo_{g2}":f"Doanh số tích lũy {g2}",
        f"Nguong_{g2}":"Ngưỡng tối thiểu","KetQua":"Kết quả","GhiChu":"Ghi chú"
    }
    if df0 is not None:
        rename[f"SoSuat_{g0}"] = f"Số suất đăng ký {g0}"
        rename[f"DoanhSo_{g0}"] = f"Doanh số tích lũy {g0}"

    out = df_final[cols_out].copy().rename(columns=rename)
    removed_out = df_removed[cols_out].copy().rename(columns=rename)
    return out, removed_out

# ======== GROUP FILES BẰNG CỘT "GIAI ĐOẠN" + "MỨC ĐĂNG KÝ" (KHỎI ĐỔI TÊN FILE) ========
def derive_ct_key(df: pd.DataFrame, xbm_map: Dict[str,str]) -> str:
    # Tự xác định CT từ "Mức đăng ký" (đặc biệt XBM)
    mucs = df["Mức đăng ký"] if "Mức đăng ký" in df.columns else df.get("MucDK")
    if mucs is None or mucs.empty:
        return "UNKNOWN"
    first = str(mucs.iloc[0]).strip()
    mapped = xbm_map.get(first, first)
    # nếu vẫn là mã XBM_MN/MB hay tên CT khác thì dùng luôn
    return mapped

def group_files_by_content(uploaded_files, xbm_map: Dict[str,str]):
    """
    Trả về dict: { CT: {stage_key: file_bytes} }
    stage_key được sắp theo thời gian dựa vào cột 'Giai đoạn'
    """
    groups: Dict[str, Dict[str, bytes]] = {}
    for uf in uploaded_files:
        df_preview = pd.read_excel(uf, header=1, nrows=5)
        ct = derive_ct_key(df_preview, xbm_map)
        # lấy giai đoạn
        g = str(df_preview["Giai đoạn"].iloc[0]).strip() if "Giai đoạn" in df_preview.columns else "Tháng ?/?"
        yy, mm, label = parse_stage_value(g)
        key = f"{yy:04d}-{mm:02d}|{label}"  # sort được
        # cần bytes (vì streamlit trả file-like)
        uf.seek(0)
        data = uf.read()
        groups.setdefault(ct, {})[key] = data
    return groups

# =============== UI ===============
st.set_page_config(page_title="Cholimex Display Checker", page_icon="📊", layout="wide")
st.title("Cholimex Display Checker (Web)")

with st.expander("⚙️ Tuỳ chọn cấu hình (không bắt buộc)"):
    cfg_text = st.text_area("Dán JSON override cho config (muc_toi_thieu / program_names / region_map / xbm_map):", height=120, placeholder='{"xbm_map":{"M70":"XBM_MN"}}')
    overrides = _load_json_text(cfg_text) if cfg_text.strip() else None

cfg = load_config(overrides)
muc_toi_thieu = cfg["muc_toi_thieu"]
program_names = cfg["program_names"]
region_map = cfg["region_map"]
xbm_map = cfg["xbm_map"]

uploaded = st.file_uploader("Tải nhiều file Excel (.xls/.xlsx) — mỗi CT ít nhất 2 tháng", type=["xls","xlsx"], accept_multiple_files=True)

colA, colB, colC = st.columns([1.2,1.2,1.6])

with colA:
    regions = st.multiselect("② Chọn miền", list(region_map.keys()), default=[])

with colB:
    mode = st.selectbox("③ Chế độ xuất", ["MKT","GSBH"], index=0)

with colC:
    st.write("④ Bộ lọc Kết quả")
    kq_all   = st.checkbox("Tất cả", value=False)
    kq_dat   = st.checkbox("Đạt", value=False)
    kq_kdat  = st.checkbox("Không đạt", value=False)
    kq_kxet  = st.checkbox("Không xét", value=False)

do_run = st.button("▶︎ Xử lý & Xuất Excel", use_container_width=True)

if do_run:
    if not uploaded:
        st.warning("Vui lòng chọn file Excel trước.")
    elif not regions:
        st.warning("Vui lòng chọn ít nhất 1 miền.")
    else:
        with st.spinner("Đang xử lý..."):
            groups = group_files_by_content(uploaded, xbm_map)

            # xác định filter kết quả
            if kq_all or (not any([kq_dat, kq_kdat, kq_kxet])):
                selected_kq = None
            else:
                sel = set()
                if kq_dat:  sel.add("Đạt")
                if kq_kdat: sel.add("Không đạt")
                if kq_kxet: sel.add("Không xét")
                selected_kq = sel if sel else None

            # xuất 1 file/từng miền
            all_outputs = {}

            for region in regions:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer_kq:
                    writer_xoa = None
                    if mode != "GSBH":
                        xoa_buf = io.BytesIO()
                        writer_xoa = pd.ExcelWriter(xoa_buf, engine="openpyxl")

                    bao_cao_data, bao_cao_huy = [], []

                    ct_idx = 0
                    for ct, files_dict in groups.items():
                        # sắp theo thời gian
                        ordered = sorted(files_dict.items(), key=lambda x: x[0])
                        if len(ordered) < 2:
                            continue
                        # lấy T2 là cuối, T1 là kế cuối, T0 nếu có là đầu
                        f_t2 = ordered[-1][1]
                        f_t1 = ordered[-2][1]
                        f_t0 = ordered[0][1] if len(ordered) >= 3 else None

                        try:
                            df_out, df_removed_out = xu_ly_chuong_trinh(
                                file_t1=f_t1, file_t2=f_t2,
                                muc_toi_thieu=muc_toi_thieu,
                                program_names=program_names,
                                xbm_map=xbm_map,
                                file_t0=f_t0,
                                filter_ketqua=selected_kq,
                                filter_tuyen_tokens=None,
                            )
                        except Exception as e:
                            st.error(f"Lỗi xử lý CT {ct}: {e}")
                            continue

                        # lọc miền
                        if region_map.get(region) != "ALL":
                            df_out = df_out[df_out["Miền"].isin(region_map[region])]
                            df_removed_out = df_removed_out[df_removed_out["Miền"].isin(region_map[region])]

                        # GSBH: ghi chú chỉ còn "Thiếu: xxx"
                        if mode == "GSBH":
                            doanh_so_cols = sorted([c for c in df_out.columns if c.startswith("Doanh số tích lũy ")])
                            if doanh_so_cols and "Ngưỡng tối thiểu" in df_out.columns:
                                col_ds_t2 = doanh_so_cols[-1]
                                mask_nd = df_out["Kết quả"].eq("Không đạt")
                                remain = (df_out.loc[mask_nd,"Ngưỡng tối thiểu"].astype(float)
                                          - df_out.loc[mask_nd,col_ds_t2].astype(float)).clip(lower=0)
                                df_out.loc[mask_nd,"Ghi chú"] = remain.map(lambda v: f"Thiếu: {fmt_money(v)}")

                            keep = ["Mức đăng ký","Tên NPP","Mã NVBH","Tên NVBH","Mã khách hàng","Tên khách hàng","Thứ bán hàng"]
                            so_suat_cols = sorted([c for c in df_out.columns if c.startswith("Số suất đăng ký ")])
                            ds_cols = sorted([c for c in df_out.columns if c.startswith("Doanh số tích lũy ")])
                            if len(so_suat_cols)>=2: keep += [so_suat_cols[-2],so_suat_cols[-1]]
                            elif len(so_suat_cols)==1: keep += [so_suat_cols[-1]]
                            if len(ds_cols)>=2: keep += [ds_cols[-2],ds_cols[-1]]
                            elif len(ds_cols)==1: keep += [ds_cols[-1]]
                            keep += ["Ngưỡng tối thiểu","Kết quả","Ghi chú"]
                            keep = [c for c in keep if c in df_out.columns]
                            df_out = df_out[keep]

                        # ghi sheet
                        df_out.to_excel(writer_kq, sheet_name=ct, index=False)

                        if mode != "GSBH" and writer_xoa is not None:
                            df_removed_out.to_excel(writer_xoa, sheet_name=ct, index=False)

                        # tổng hợp
                        try:
                            tong = df_out.filter(like="Số suất đăng ký").iloc[:, -1].sum()
                            ko_dat = df_out.loc[df_out["Kết quả"]=="Không đạt",:].filter(like="Số suất đăng ký").iloc[:, -1].sum()
                            tile = f"{(ko_dat/tong):.1%}" if tong>0 else "0%"
                            ct_idx += 1
                            bao_cao_data.append([ct_idx, program_names.get(ct, ct), muc_toi_thieu.get(ct,0), int(tong), int(ko_dat), tile])
                            if mode != "GSBH":
                                bao_cao_huy.append([ct_idx, program_names.get(ct, ct), int(ko_dat)])
                        except Exception:
                            pass

                    # sheet tổng hợp (đơn giản)
                    if bao_cao_data:
                        df_tong = pd.DataFrame(bao_cao_data, columns=[
                            "STT","Tên chương trình","DOANH SỐ TỐI THIỂU PHÁT SINH/ SUẤT/ THÁNG (VND)",
                            "TỔNG SỐ SUẤT TRƯNG BÀY","SỐ SUẤT KHÔNG ĐẠT","TỈ LỆ"
                        ])
                        df_tong.to_excel(writer_kq, sheet_name="BaoCao_TongHop", index=False)

                    if (mode!="GSBH") and bao_cao_huy:
                        df_huy = pd.DataFrame(bao_cao_huy, columns=["STT","Tên chương trình","TỔNG SỐ SUẤT HỦY TRƯNG BÀY TRÊN HT DMS"])
                        df_huy.to_excel(writer_kq, sheet_name="BaoCao_Huy", index=False)

                # lưu file
                fname = f"TongHop_{region}{'_GSBH' if mode=='GSBH' else ''}.xlsx"
                all_outputs[fname] = output.getvalue()

                if mode != "GSBH" and writer_xoa is not None:
                    fname_x = f"TongHop_Xoa_{region}.xlsx"
                    all_outputs[fname_x] = xoa_buf.getvalue()

        # nút tải về
        for fn, data in all_outputs.items():
            st.download_button("⬇️ Tải "+fn, data=data, file_name=fn, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.success("Xong!")
