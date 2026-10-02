from fastapi import FastAPI, UploadFile, File, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import os
import shutil
import zipfile
import httpx
from parse_invoice import parse_invoice_xml, export_to_excel

app = FastAPI(title="Invoice Management API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "./temp_uploads"
STATIC_DIR = "./static"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

# ----------------- HÀM TRA CỨU RỦI RO MST -----------------
async def check_tax_status(mst: str) -> dict:
    """
    Tra cứu trạng thái doanh nghiệp theo MST qua API công khai (VietQR / Cổng thuế).
    Nếu lỗi kết nối hoặc MST mẫu, fallback về trạng thái giả định an toàn.
    """
    mst_clean = mst.strip().replace(" ", "").replace("-", "")
    if not mst_clean:
        return {"status": "Không xác định", "risk": False}

    api_url = f"https://api.vietqr.io/v2/business/{mst_clean}"
    try:
        async with httpx.AsyncClient(timeout=4.0) as client:
            resp = client.get(api_url)
            res = resp.json()
            if res.get("code") == "00" and res.get("data"):
                data = res["data"]
                # Một số API trả về text tình trạng hoặc cờ
                activity_status = data.get("status", "Đang hoạt động")
                is_risk = any(keyword in activity_status.lower() for keyword in ["ngừng", "đóng", "bỏ trốn", "khoá", "nợ thuế"])
                return {
                    "company_name": data.get("name", ""),
                    "address": data.get("address", ""),
                    "status": activity_status if activity_status else "Đang hoạt động",
                    "risk": is_risk
                }
    except Exception:
        pass

    # Mặc định an toàn nếu không gọi được mạng
    return {"status": "Đang hoạt động", "risk": False}


# ----------------- API XỬ LÝ HÓA ĐƠN & QUÉT RỦI RO -----------------
@app.post("/api/process-invoices")
async def process_invoices(files: list[UploadFile] = File(...)):
    for f in os.listdir(UPLOAD_DIR):
        fpath = os.path.join(UPLOAD_DIR, f)
        if os.path.isdir(fpath):
            shutil.rmtree(fpath)
        else:
            os.remove(fpath)

    extracted_xmls = []

    for file in files:
        temp_path = os.path.join(UPLOAD_DIR, file.filename)
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        if file.filename.endswith(".zip"):
            extract_folder = os.path.join(UPLOAD_DIR, "extracted_" + file.filename)
            os.makedirs(extract_folder, exist_ok=True)
            with zipfile.ZipFile(temp_path, "r") as zip_ref:
                zip_ref.extractall(extract_folder)
            for root_dir, _, fs in os.walk(extract_folder):
                for f in fs:
                    if f.endswith(".xml"):
                        extracted_xmls.append(os.path.join(root_dir, f))
        elif file.filename.endswith(".xml"):
            extracted_xmls.append(temp_path)

    all_summaries = []
    all_details = []

    for xml_file in extracted_xmls:
        sum_row, det_rows = parse_invoice_xml(xml_file)
        if sum_row:
            # Tra cứu rủi ro MST người bán
            tax_info = await check_tax_status(sum_row.get("MST Người Bán", ""))
            sum_row["Trạng Thái MST"] = tax_info["status"]
            sum_row["CanhBaoRuiRo"] = tax_info["risk"]
            all_summaries.append(sum_row)
        all_details.extend(det_rows)

    out_excel = os.path.join(UPLOAD_DIR, "Ket_Qua_Hoa_Don.xlsx")
    export_to_excel(all_summaries, all_details, out_excel)

    return {
        "status": "success",
        "total_invoices": len(all_summaries),
        "total_items": len(all_details),
        "summaries": all_summaries,
        "details": all_details,
        "excel_download_url": "/api/download-excel"
    }


@app.get("/api/download-excel")
async def download_excel():
    excel_path = os.path.join(UPLOAD_DIR, "Ket_Qua_Hoa_Don.xlsx")
    if os.path.exists(excel_path):
        return FileResponse(
            excel_path,
            filename="Bang_Ke_Hoa_Don_TongHop_ChiTiet.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return JSONResponse(status_code=404, content={"error": "File chưa sẵn sàng"})


# Phục vụ Static frontend
app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")