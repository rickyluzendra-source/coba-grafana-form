from fastapi import FastAPI, Response
from fastapi.middleware.cors import CORSMiddleware
import openpyxl
from io import BytesIO

app = FastAPI()

# Mengizinkan Grafana Cloud untuk mengakses API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/calculate-export")
async def calculate_export(data: dict):
    # Ambil data dari Form Grafana
    unit = data.get("unit_bisnis", "Branch Utama")
    kategori = data.get("kategori", "Hardware")
    volume = float(data.get("volume", 0))
    harga = float(data.get("harga_unit", 0))
    biaya = float(data.get("biaya_direct", 0))
    tax = float(data.get("tax_rate", 0.11))

    # Buat Workbook Excel
    wb = openpyxl.Workbook()
    
    # Sheet 1: Data Master
    ws1 = wb.active
    ws1.title = "Data Master"
    ws1.append(["Unit Bisnis", "Kategori", "Volume", "Harga/Unit", "Biaya Direct/Unit"])
    ws1.append([unit, kategori, volume, harga, biaya])
    
    # Sheet 2: Hasil Kalkulasi (dengan Rumus Excel)
    ws2 = wb.create_sheet(title="Hasil Kalkulasi")
    ws2.append(["Gross Revenue", "Total Direct Cost", "Tax Overhead", "Net Profit"])
    ws2.append([
        "='Data Master'!C2*'Data Master'!D2",
        "='Data Master'!C2*'Data Master'!E2",
        f"=A2*{tax}",
        "=A2-B2-C2"
    ])
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return Response(
        content=output.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=Grafana_Hasil_Kalkulasi.xlsx"}
    )
