from fastapi import FastAPI, Response
from fastapi.middleware.cors import CORSMiddleware
import openpyxl
from io import BytesIO

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Database sementara di memory server (Menampung semua submit)
database_penampung = []

@app.post("/submit-data")
async def submit_data(data: dict):
    """Menampung inputan dari Grafana tanpa langsung download Excel"""
    database_penampung.append(data)
    return {
        "status": "success", 
        "message": f"Data berhasil disimpan! Total data tersimpan: {len(database_penampung)}"
    }

@app.post("/download-all")
async def download_all(payload: dict = None):
    """Mendownload SELURUH data tersimpan menjadi 1 File Excel"""
    wb = openpyxl.Workbook()
    
    # --- SHEET 1: DATA MASTER (Semua Data Submit) ---
    ws_master = wb.active
    ws_master.title = "Data Master"
    ws_master.append(["ID", "Unit Bisnis", "Kategori", "Volume Sales", "Harga/Unit", "Biaya Direct/Unit"])
    
    for idx, item in enumerate(database_penampung, start=1):
        ws_master.append([
            f"TRX-{idx:03d}",
            item.get("unit_bisnis", "Branch Utama"),
            item.get("kategori", "General"),
            float(item.get("volume", 0)),
            float(item.get("harga_unit", 0)),
            float(item.get("biaya_direct", 0))
        ])

    # --- SHEET 2: HASIL KALKULASI (Rumus Excel untuk Semua Baris) ---
    ws_calc = wb.create_sheet(title="Hasil Kalkulasi")
    ws_calc.append(["ID", "Unit Bisnis", "Gross Revenue", "Total Direct Cost", "Tax (11%)", "Net Profit", "Margin %"])
    
    tax_rate = 0.11
    total_rows = len(database_penampung)
    
    for r in range(2, total_rows + 2):
        ws_calc.append([
            f"='Data Master'!A{r}",
            f"='Data Master'!B{r}",
            f"='Data Master'!D{r}*'Data Master'!E{r}",  # Revenue
            f"='Data Master'!D{r}*'Data Master'!F{r}",  # Direct Cost
            f"=C{r}*{tax_rate}",                       # Tax
            f"=C{r}-D{r}-E{r}",                        # Net Profit
            f"=F{r}/C{r}"                               # Margin
        ])
        
    # Baris Total di Paling Bawah
    tot_row = total_rows + 2
    if total_rows > 0:
        ws_calc.cell(row=tot_row, column=1, value="TOTAL")
        ws_calc.cell(row=tot_row, column=3, value=f"=SUM(C2:C{tot_row-1})")
        ws_calc.cell(row=tot_row, column=4, value=f"=SUM(D2:D{tot_row-1})")
        ws_calc.cell(row=tot_row, column=5, value=f"=SUM(E2:E{tot_row-1})")
        ws_calc.cell(row=tot_row, column=6, value=f"=SUM(F2:F{tot_row-1})")
        ws_calc.cell(row=tot_row, column=7, value=f"=AVERAGE(G2:G{tot_row-1})")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return Response(
        content=output.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=Master_Kalkulasi_Rekap_Semua.xlsx"}
    )
