from sheets.forma8_1 import run_forma8_1
from sheets.forma8_2 import run_forma8_2
from sheets.forma8_3 import run_forma8_3
from sheets.forma8_4 import run_forma8_4
from sheets.forma8_5 import run_forma8_5
from sheets.forma8_6 import run_forma8_6
from sheets.forma8_11 import run_forma8_11
from sheets.forma8_7 import run_forma8_7
from sheets.forma8_10 import run_forma8_10
from sheets.forma8_8 import run_forma8_8
import os
from openpyxl import load_workbook

# ==================== KONFİQURASİYA ====================
UCOT_FILE = r"C:\Users\samir.alizade\Desktop\Automantion\Risk\UcotA.xlsx"
TEMPLATE_FILE = r"C:\Users\samir.alizade\Desktop\Automantion\Risk\ALL.xlsx"
OUTPUT_FOLDER = r"C:\Users\samir.alizade\Desktop\Automantion\Risk\AutoTEST"
PREVIOUS_FOLDER = r"C:\Users\samir.alizade\Desktop\Automantion\Risk\Previous"
REFERENCE_DATE = "2025-10-01"

def main():
    print("=" * 60)
    print("▶ FORMA8 PROSESI BAŞLADI")
    print("=" * 60)
    
    # UCOT faylının mövcudluğunu yoxla
    if not os.path.exists(UCOT_FILE):
        print(f"❌ XƏTA: {UCOT_FILE} tapılmadı!")
        return
    
    if not os.path.exists(TEMPLATE_FILE):
        print(f"❌ XƏTA: {TEMPLATE_FILE} tapılmadı!")
        return
    
    # Forma8_1 prosesi
    print("\n▶ Forma8_1 işlənir...")
    try:
        run_forma8_1(
            ucot_file=UCOT_FILE,
            template_file=TEMPLATE_FILE,
            output_folder=OUTPUT_FOLDER
        )
        print("✅ Forma8_1 tamamlandı")
    except Exception as e:
        print(f"❌ Forma8_1 xətası: {e}")
        return
    
    # Excel fayllarını tap
    excel_files = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith(".xlsx")]
    
    if not excel_files:
        print("⚠ Heç bir Excel faylı tapılmadı!")
        return
    
    # Forma8_2 prosesi (Forma8_7-dən ƏVVƏL)
    print("\n▶ Forma8_2 işlənir...")
    success_count_8_2 = 0
    total_f_values = {}  # Hər fayl üçün total_f saxla
    
    for file in excel_files:
        excel_path = os.path.join(OUTPUT_FOLDER, file)
        try:
            # Return dəyərini götür
            total_f = run_forma8_2(
                excel_file=excel_path,
                ucot_file=UCOT_FILE,
                reference_date=REFERENCE_DATE
            )
            
            # Saxla
            if total_f is not None and total_f > 0:
                total_f_values[file] = total_f
                print(f"  ✓ {file} - Total F: {total_f}")
            
            success_count_8_2 += 1
        except Exception as e:
            print(f"  ✗ {file}: {e}")
    
    # Forma8_3 prosesi
    print("\n▶ Forma8_3 işlənir...")
    success_count_8_3 = 0
    for file in excel_files:
        excel_path = os.path.join(OUTPUT_FOLDER, file)
        try:
            run_forma8_3(
                excel_file=excel_path,
                ucot_file=UCOT_FILE,
                reference_date=REFERENCE_DATE
            )
            success_count_8_3 += 1
        except Exception as e:
            print(f"  ✗ {file}: {e}")
    
    # Forma8_4 prosesi
    print("\n▶ Forma8_4 işlənir...")
    success_count_8_4 = 0
    for file in excel_files:
        excel_path = os.path.join(OUTPUT_FOLDER, file)
        try:
            run_forma8_4(
                excel_file=excel_path,
                ucot_file=UCOT_FILE
            )
            success_count_8_4 += 1
        except Exception as e:
            print(f"  ✗ {file}: {e}")
    
    # Forma8_5 prosesi
    print("\n▶ Forma8_5 işlənir...")
    success_count_8_5 = 0
    for file in excel_files:
        excel_path = os.path.join(OUTPUT_FOLDER, file)
        try:
            run_forma8_5(
                excel_file=excel_path,
                ucot_file=UCOT_FILE,
                reference_date=REFERENCE_DATE
            )
            success_count_8_5 += 1
        except Exception as e:
            print(f"  ✗ {file}: {e}")
    
    # Forma8_6 prosesi
    print("\n▶ Forma8_6 işlənir...")
    success_count_8_6 = 0
    for file in excel_files:
        excel_path = os.path.join(OUTPUT_FOLDER, file)
        try:
            run_forma8_6(excel_file=excel_path)
            success_count_8_6 += 1
        except Exception as e:
            print(f"  ✗ {file}: {e}")

    # Forma8_11 prosesi
    print("\n▶ Forma8_11 işlənir...")
    success_count_8_11 = 0
    for file in excel_files:
        excel_path = os.path.join(OUTPUT_FOLDER, file)
        try:
            run_forma8_11(
                excel_file=excel_path,
                ucot_file=UCOT_FILE,
                reference_date=REFERENCE_DATE
            )
            success_count_8_11 += 1
        except Exception as e:
            print(f"  ✗ {file}: {e}")
    
    # Forma8_7 prosesi
    print("\n▶ Forma8_7 işlənir...")
    success_count_8_7 = 0
    for file in excel_files:
        excel_path = os.path.join(OUTPUT_FOLDER, file)
        try:
            # Total F dəyərini götür (varsa)
            total_f = total_f_values.get(file, None)
            
            run_forma8_7(
                excel_file=excel_path,
                previous_folder=PREVIOUS_FOLDER,
                reference_date=REFERENCE_DATE,
                total_f_from_forma8_2=total_f
            )
            success_count_8_7 += 1
        except Exception as e:
            print(f"  ✗ {file}: {e}")
    
    # Forma8_10 prosesi
    print("\n▶ Forma8_10 işlənir...")
    success_count_8_10 = 0
    for file in excel_files:
        excel_path = os.path.join(OUTPUT_FOLDER, file)
        try:
            run_forma8_10(
                excel_file=excel_path,
                previous_folder=PREVIOUS_FOLDER,
                reference_date=REFERENCE_DATE  # ✅ Əlavə edildi
            )
            success_count_8_10 += 1
        except Exception as e:
            print(f"  ✗ {file}: {e}")

    # Forma8_8 prosesi
    print("\n▶ Forma8_8 işlənir...")
    success_count_8_8 = 0
    for file in excel_files:
        excel_path = os.path.join(OUTPUT_FOLDER, file)
        try:
            run_forma8_8(
                excel_file=excel_path,
                reference_date=REFERENCE_DATE,
                ucot_file=UCOT_FILE  # ✅ Əlavə edildi
            )
            success_count_8_8 += 1
        except Exception as e:
            print(f"  ✗ {file}: {e}")
    
    print("\n" + "=" * 60)
    print(f"✅ PROSES TAMAMLANDI:")
    print(f"   - Forma8_2: {success_count_8_2}/{len(excel_files)} fayl")
    print(f"   - Forma8_3: {success_count_8_3}/{len(excel_files)} fayl")
    print(f"   - Forma8_4: {success_count_8_4}/{len(excel_files)} fayl")
    print(f"   - Forma8_5: {success_count_8_5}/{len(excel_files)} fayl")
    print(f"   - Forma8_6: {success_count_8_6}/{len(excel_files)} fayl")
    print(f"   - Forma8_11: {success_count_8_11}/{len(excel_files)} fayl")
    print(f"   - Forma8_7: {success_count_8_7}/{len(excel_files)} fayl")
    print(f"   - Forma8_10: {success_count_8_10}/{len(excel_files)} fayl")
    print(f"   - Forma8_8: {success_count_8_8}/{len(excel_files)} fayl")
    print("=" * 60)


if __name__ == "__main__":
    main()