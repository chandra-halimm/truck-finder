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
    "meter-4": ["8256", "8635", "8934", "9164", "9634", "8255", "8457", "8614", "8255", "8339", "9674", "9547", "9649", "9787", "8770", "9273", "8546"],
    "meter-8": ["9190", "8434", "8438", "8013", "8774", "9787", "8335", "8437", "8099", "9164", "9477", "8545", "9190", "9273", "9201", "8544", "9957", "9336", "9284", "9674", "9367", "8544", "9957", "9336", "9284", "9674", "9367", "8544", "9316", "8446", "9507", "9325", "8770", "9171", "9367", "9774", "9812", "9766", "9731", "8545", "9347", "9787", "9445", "9673", "9787", "9719", "9749"],
    "meter-7": ["9168", "9354", "9336", "8438", "8440", "8548", "9785", "9338", "8548", "9493", "9325", "9336", "8186", "8890", "9347", "9503", "8616", "9221", "8616", "9750", "9369", "8544", "9224", "9483", "9774"],
    "meter-11": ["9168", "9354", "8255", "9336", "8314", "8438", "8440", "8548", "9785", "9338", "8548", "9493", "8256", "9325", "8186", "8890", "9347", "8616", "9488", "9221", "8616", "9750", "9600", "9787", "8544", "9224", "9483", "9774"],
    "meter-12": ["9201", "9190", "9811", "8436", "9284", "8544", "9356", "9316", "8040", "9284", "9787", "9345", "9649", "9190", "9488", "8545", "9171", "8077", "9677", "9787", "9742", "8854", "8546", "9643", "9742", "9356", "9284", "9473", "9483", "9674", "9201", "9221", "9371", "9789", "9766"],
    "meter-15": ["9811", "8436", "9284", "8544", "9356", "9316", "8040", "9284", "9273", "9787", "9345", "9649", "9190", "9488", "8545", "9171", "8077", "9677", "9787", "9742", "8854", "9157", "8546", "9190", "9643", "9742", "9356", "9284", "9473", "9483", "9674", "9201", "9221", "9371", "9789", "9766"],
    "meter-16": ["9201", "9784", "8323", "8136", "8097", "8932", "8336", "9331", "8428", "9477", "8548", "9354", "9731", "9327", "9649", "9190", "9347", "9787", "9677", "9749", "9587", "9957", "8446", "9730", "9643", "8770", "9338", "8097", "8066", "9336", "8548", "9284", "9373", "9957"]
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
