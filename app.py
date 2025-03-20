from flask import Flask, render_template
import re
from openpyxl import load_workbook

app = Flask(__name__)

# Mapping sheet ke key yang diinginkan
sheet_mapping = {
    "BAY2-BS.4": "meter-4",
    "BAY3-PL.8": "meter-8",
    "BAY4-BS.7": "meter-7",
    "BAY4-BS.11": "meter-11",
    "BAY6-PL.16": "meter-16",
    "BAY7-BS.12": "meter-12",
    "BAY7-BS.15": "meter-15"
}

# Data referensi yang ingin dibandingkan
reference_data = {
    "meter-4": ["8934", "111"],
    "meter-8": ["2222", "3333"],
    "meter-7": ["4444", "5555"],
    "meter-11": ["8548", "7777"],
    "meter-12": ["8136", "8548"],
    "meter-15": ["8097", "8040"],
    "meter-16": ["8614", "9201"]
}

# Tentukan baris awal dan kolom target
start_row = 9  
truck_column = "C"

# Fungsi untuk mengekstrak angka dari teks
def extract_numbers(text):
    if text:
        numbers = re.findall(r"\d+", str(text))
        return "".join(numbers) if numbers else None
    return None

def process_excel():
    # Membuka file Excel
    wb = load_workbook("20250320.xlsx")

    truck_data = {}

    # Iterasi melalui sheet yang dipilih
    for sheet_name, meter_key in sheet_mapping.items():
        if sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]

            # Ambil angka dari kolom C, mulai dari baris 9 ke bawah
            truck_numbers = [
                extract_numbers(sheet[f"{truck_column}{row}"].value)
                for row in range(start_row, sheet.max_row + 1)
                if sheet[f"{truck_column}{row}"].value is not None
            ]

            # Filter agar tidak ada nilai None
            truck_numbers = [num for num in truck_numbers if num]

            # Simpan hasil dengan key yang sesuai
            truck_data[meter_key] = truck_numbers

    # Dictionary untuk menyimpan hasil perbandingan
    comparison_result = {}
    # Perbandingan data
    for meter, trucks in reference_data.items():
        excel_trucks = set(truck_data.get(meter, []))  # Data dari Excel (set untuk efisiensi pencarian)
        reference_trucks = set(trucks)  # Data referensi

        # Mencari data yang ada di referensi tapi tidak ada di Excel
        missing_trucks = reference_trucks - excel_trucks
        # Mencari data yang ada di Excel tapi tidak ada di referensi
        additional_trucks = excel_trucks - reference_trucks

        comparison_result[meter] = {
            "missing": list(missing_trucks),
            "additional": list(additional_trucks)
        }

    # Menemukan truk yang hilang di satu meter tetapi ada di meter lain
    missing_found_elsewhere = []

    for meter, result in comparison_result.items():
        for missing_truck in result["missing"]:
            # Cek apakah nomor truk yang hilang ditemukan di meter lain
            found_in_other_meter = [
                other_meter for other_meter, data in truck_data.items() if missing_truck in data and other_meter != meter
            ]
            
            if found_in_other_meter:
                missing_found_elsewhere.append(f"Nomor {missing_truck} mengisi di {meter} sedangkan antrian di {', '.join(found_in_other_meter)}")

    return comparison_result, missing_found_elsewhere

@app.route('/')
def index():
    comparison_result, missing_found_elsewhere = process_excel()
    return render_template('index.html', comparison_result=comparison_result, missing_found_elsewhere=missing_found_elsewhere)

if __name__ == '__main__':
    app.run(debug=True)
