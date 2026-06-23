import os
import win32com.client
import win32print
import time

original_default_printer = win32print.GetDefaultPrinter()

print_flag = "N"
if os.path.exists("piutang.conf"):
    with open("piutang.conf", "r") as f:
        lines = f.readlines()
    for i, line in enumerate(lines):
        if "[PRINT]" in line:
            for j in range(i + 1, len(lines)):
                next_line = lines[j].strip()
                if next_line:
                    print_flag = next_line.upper()
                    break
            break

if print_flag == "Y":
    printers = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    )
    
    last_printer = ""
    if os.path.exists("last_printer.txt"):
        with open("last_printer.txt", "r") as f:
            last_printer = f.read().strip()

    chosen_printer = ""
    if last_printer:
        print(f"--> Printer terakhir digunakan: {last_printer}")
        konfirmasi = input("--> Gunakan printer ini lagi? (Tekan ENTER atau Y jika ya / ketik 'n' untuk pilih ulang): ").strip().upper()
        if konfirmasi == "" or konfirmasi == "Y":
            chosen_printer = last_printer

    if not chosen_printer:
        print("--> Daftar Printer Tersedia:")
        for i, printer in enumerate(printers):
            print(f"--> {i + 1}. {printer[2]}")

        choice = int(input("--> Masukkan nomor printer yang ingin digunakan: ")) - 1
        chosen_printer = printers[choice][2]
        
        with open("last_printer.txt", "w") as f:
            f.write(chosen_printer)

    try:
        print(f"--> Mengunci driver printer: {chosen_printer}")
        win32print.SetDefaultPrinter(chosen_printer)
        
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.ScreenUpdating = True

        file_path = os.path.abspath("Print_AR.xlsx")
        workbook = excel.Workbooks.Open(file_path)
        sheet = workbook.Worksheets(1)
        workbook.Windows(1).View = 2

        max_row = sheet.UsedRange.Rows.Count
        start_rows = []
        for r in range(1, max_row + 1):
            val = sheet.Cells(r, 2).Value
            if val and "LAPORAN HASIL TAGIHAN" in str(val):
                start_rows.append(r)

        print(f"--> Ditemukan {len(start_rows)} data laporan penagih. Mulai memproses cetak...")

        for i in range(len(start_rows)):
            start_row = start_rows[i]
            if i + 1 < len(start_rows):
                end_row = start_rows[i+1] - 1
            else:
                end_row = max_row

            while end_row > start_row:
                is_empty = True
                for col in range(2, 17):
                    if sheet.Cells(end_row, col).Value is not None:
                        is_empty = False
                        break
                if is_empty:
                    end_row -= 1
                else:
                    break

            print(f"--> Mencetak data ke-{i+1} (Baris {start_row} s/d {end_row})")
            
            sheet.ResetAllPageBreaks()
            sheet.PageSetup.Zoom = True   
            sheet.PageSetup.PrintArea = f"B{start_row}:P{end_row}"
            sheet.PageSetup.Orientation = 2
            sheet.PageSetup.Zoom = False  
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = 1
            
            time.sleep(0.3)
            sheet.PrintOut()

        workbook.Close(False)
        excel.Quit()
        print("--> Proses cetak seluruh data selesai dengan sukses.")

    except Exception as e:
        print(f"--> Terjadi kesalahan saat memproses: {e}")
        
    finally:
        win32print.SetDefaultPrinter(original_default_printer)
        print("--> Sistem printer komputer Anda telah dikembalikan ke semula.")
else:
    print("--> Proses cetak dilewati berdasarkan konfigurasi piutang.conf.")