import ctypes
import os
import xlwings as xw


def cetak_laporan_ar_xlwings():
    nama_file = "Print_AR.xlsx"
    path_file = os.path.abspath(nama_file)

    if not os.path.exists(path_file):
        msg = f"File '{nama_file}' tidak ditemukan di folder!"
        print(f"--> {msg}")
        ctypes.windll.user32.MessageBoxW(0, msg, "Peringatan", 48)
        return

    app = xw.App(visible=False, add_book=False)

    try:
        wb = app.books.open(path_file)
        ws = wb.sheets.active

        app.visible = True
        printer_dipilih = app.api.Dialogs(9).Show()

        if not printer_dipilih:
            msg = "Proses pencetakan dibatalkan."
            print(f"--> {msg}")
            ctypes.windll.user32.MessageBoxW(0, msg, "Batal", 64)
            wb.close()
            return

        app.visible = False
        app.screen_updating = False

        ws.api.ResetAllPageBreaks()

        ws.api.PageSetup.Orientation = 2
        ws.api.PageSetup.PaperSize = 2
        ws.api.PageSetup.LeftMargin = app.api.InchesToPoints(0.25)
        ws.api.PageSetup.RightMargin = app.api.InchesToPoints(0.25)
        ws.api.PageSetup.TopMargin = app.api.InchesToPoints(0.25)
        ws.api.PageSetup.BottomMargin = app.api.InchesToPoints(0.25)
        ws.api.PageSetup.Zoom = False
        ws.api.PageSetup.FitToPagesWide = 1
        ws.api.PageSetup.FitToPagesTall = 1

        last_row = ws.used_range.last_cell.row
        start_row = 0
        jumlah_terhitung = 0

        for r in range(1, last_row + 1):
            ada_pembuka = False
            ada_penutup = False

            for c in range(2, 17):
                nilai_sel = str(ws.cells(r, c).value or "").upper()
                if "LAPORAN HASIL TAGIHAN" in nilai_sel:
                    ada_pembuka = True
                if "TTD SALES & COLLECTOR" in nilai_sel:
                    ada_penutup = True

            if ada_pembuka and start_row == 0:
                start_row = r

            if ada_penutup and start_row > 0:
                end_row = r
                ws.api.PageSetup.PrintArea = f"B{start_row}:P{end_row}"
                ws.api.PrintOut(From=1, To=1, Copies=1)
                jumlah_terhitung += 1
                start_row = 0

        app.screen_updating = True

        if jumlah_terhitung > 0:
            msg = f"Selesai! Total ada {jumlah_terhitung} kelompok laporan yang dicetak."
            print(f"--> {msg}")
            ctypes.windll.user32.MessageBoxW(0, msg, "Sukses", 64)
        else:
            msg = "Tidak ditemukan blok data dengan kata kunci yang sesuai."
            print(f"--> {msg}")
            ctypes.windll.user32.MessageBoxW(0, msg, "Peringatan", 48)

        wb.close()

    finally:
        app.quit()


if __name__ == "__main__":
    cetak_laporan_ar_xlwings()
