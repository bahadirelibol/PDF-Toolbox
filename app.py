import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
from docx import Document
import os


# -----------------------------------------------------------------------------#
#  Kesme Sekmesi
# -----------------------------------------------------------------------------#
class CutTab(ttk.Frame):
    """PDF sayfa kesici sekmesi"""

    def __init__(self, master: ttk.Notebook):
        super().__init__(master, padding=20)

        # --- Dosya seçiciler
        self.input_entry = self._add_file_row("Kaynak PDF", 0, self._select_input)
        self.output_entry = self._add_file_row("Çıktı PDF", 1)

        # --- Sayfa aralığı
        self.start_entry = self._add_spin_row("Başlangıç Sayfa", 2)
        self.end_entry = self._add_spin_row("Bitiş Sayfa", 3)

        # --- İşlem düğmesi
        ttk.Button(
            self, text="Kes", command=self.cut_pdf
        ).grid(row=4, column=0, columnspan=3, pady=(20, 0), ipadx=30)

    # ----------------------------- yardımcılar ---------------------------------
    def _add_file_row(self, label: str, row: int, cmd=None):
        ttk.Label(self, text=f"{label}:").grid(row=row, column=0, sticky="w", pady=8)
        entry = ttk.Entry(self, width=30)
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 8))
        if cmd:
            ttk.Button(self, text="Seç", command=cmd).grid(row=row, column=2)
        return entry

    def _add_spin_row(self, label: str, row: int):
        ttk.Label(self, text=f"{label}:").grid(row=row, column=0, sticky="w", pady=8)
        spin = ttk.Spinbox(self, from_=1, to=9999, width=5)
        spin.grid(row=row, column=1, sticky="w")
        return spin

    def _select_input(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if path:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, path)
            base = os.path.splitext(os.path.basename(path))[0]
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, f"{base}_cut.pdf")

    # ----------------------------- iş mantığı ----------------------------------
    def cut_pdf(self):
        in_path = self.input_entry.get()
        out_path = self.output_entry.get()
        try:
            start = int(self.start_entry.get())
            end = int(self.end_entry.get())
        except ValueError:
            messagebox.showerror("Hata", "Sayfa numaraları geçerli bir sayı olmalı!")
            return

        if not in_path or not out_path:
            messagebox.showerror("Hata", "Lütfen dosya yollarını girin!")
            return

        try:
            reader = PdfReader(in_path)
            total = len(reader.pages)
            if start < 1 or end < start or end > total:
                messagebox.showerror(
                    "Hata", f"Lütfen 1–{total} aralığında geçerli bir sayfa aralığı girin!"
                )
                return

            writer = PdfWriter()
            for idx in range(start - 1, end):
                writer.add_page(reader.pages[idx])

            with open(out_path, "wb") as f:
                writer.write(f)

            messagebox.showinfo("Başarılı", "Seçilen sayfalar yeni PDF'e kaydedildi.")
        except Exception as e:
            messagebox.showerror("Hata", str(e))


# -----------------------------------------------------------------------------#
#  Birleştirme Sekmesi
# -----------------------------------------------------------------------------#
class MergeTab(ttk.Frame):
    """PDF birleştirme sekmesi"""

    def __init__(self, master):
        super().__init__(master, padding=20)

        self.pdf1_entry = self._add_file_row("1. PDF", 0, self._select_pdf1)
        self.pdf2_entry = self._add_file_row("2. PDF", 1, self._select_pdf2)
        self.output_entry = self._add_file_row("Çıktı PDF", 2)

        ttk.Button(
            self, text="Birleştir", command=self.merge
        ).grid(row=3, column=0, columnspan=3, pady=(20, 0), ipadx=30)

    # ----------------------------- yardımcılar ---------------------------------
    def _add_file_row(self, label, row, cmd=None):
        ttk.Label(self, text=f"{label}:").grid(row=row, column=0, sticky="w", pady=8)
        entry = ttk.Entry(self, width=30)
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 8))
        if cmd:
            ttk.Button(self, text="Seç", command=cmd).grid(row=row, column=2)
        return entry

    def _select_pdf1(self):
        self._select_to(self.pdf1_entry, suffix="_merged")

    def _select_pdf2(self):
        self._select_to(self.pdf2_entry)

    def _select_to(self, entry: ttk.Entry, suffix=""):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)
            if suffix and not self.output_entry.get():
                base = os.path.splitext(os.path.basename(path))[0]
                self.output_entry.insert(0, f"{base}{suffix}.pdf")

    # ----------------------------- iş mantığı ----------------------------------
    def merge(self):
        p1, p2, out = self.pdf1_entry.get(), self.pdf2_entry.get(), self.output_entry.get()
        if not all([p1, p2, out]):
            messagebox.showerror("Hata", "Tüm alanları doldurun!")
            return

        try:
            reader1, reader2 = PdfReader(p1), PdfReader(p2)
            writer = PdfWriter()

            # önce 1. PDF
            for page in reader1.pages:
                writer.add_page(page)
            # sonra 2. PDF
            for page in reader2.pages:
                writer.add_page(page)

            with open(out, "wb") as f:
                writer.write(f)

            messagebox.showinfo("Başarılı", "PDF dosyaları birleştirildi.")
        except Exception as e:
            messagebox.showerror("Hata", str(e))


# -----------------------------------------------------------------------------#
#  PDF → Word Sekmesi
# -----------------------------------------------------------------------------#
class WordTab(ttk.Frame):
    """PDF-ten Word (.docx) dönüştürme sekmesi"""

    def __init__(self, master):
        super().__init__(master, padding=20)

        self.pdf_entry = self._add_file_row("PDF", 0, self._select_pdf)
        self.docx_entry = self._add_file_row("Word", 1)

        ttk.Button(
            self, text="Dönüştür", command=self.convert
        ).grid(row=2, column=0, columnspan=3, pady=(20, 0), ipadx=30)

    # ----------------------------- yardımcılar ---------------------------------
    def _add_file_row(self, label, row, cmd=None):
        ttk.Label(self, text=f"{label}:").grid(row=row, column=0, sticky="w", pady=8)
        entry = ttk.Entry(self, width=30)
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 8))
        if cmd:
            ttk.Button(self, text="Seç", command=cmd).grid(row=row, column=2)
        return entry

    def _select_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if path:
            self.pdf_entry.delete(0, tk.END)
            self.pdf_entry.insert(0, path)
            base = os.path.splitext(os.path.basename(path))[0]
            self.docx_entry.delete(0, tk.END)
            self.docx_entry.insert(0, f"{base}.docx")

    # ----------------------------- iş mantığı ----------------------------------
    def convert(self):
        pdf, docx_path = self.pdf_entry.get(), self.docx_entry.get()
        if not pdf or not docx_path:
            messagebox.showerror("Hata", "Dosya yollarını girin!")
            return
        try:
            reader = PdfReader(pdf)
            doc = Document()
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    doc.add_paragraph(text)
            doc.save(docx_path)
            messagebox.showinfo("Başarılı", "PDF, Word dosyasına dönüştürüldü.")
        except Exception as e:
            messagebox.showerror("Hata", str(e))


# -----------------------------------------------------------------------------#
#  Ana Uygulama Penceresi
# -----------------------------------------------------------------------------#
class PDFApp(tk.Tk):
    """Ana uygulama penceresi"""

    def __init__(self):
        super().__init__()
        self.title("PDF Araçları")
        self.geometry("600x420")
        self.resizable(False, False)

        # ---- Stil ayarları
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background="#f6f7fb")
        style.configure("TLabel", background="#f6f7fb", font=("Segoe UI", 11))
        style.configure("TButton", font=("Segoe UI", 11))

        # ---- Sekmeli gezinme
        notebook = ttk.Notebook(self)
        notebook.pack(expand=True, fill="both", padx=10, pady=10)

        notebook.add(CutTab(notebook), text="Kes")
        notebook.add(MergeTab(notebook), text="Birleştir")
        notebook.add(WordTab(notebook), text="Word")

        # ---- Çıkış düğmesi
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(bottom, text="Çıkış", command=self.destroy).pack(side="right")


# -----------------------------------------------------------------------------#
if __name__ == "__main__":
    PDFApp().mainloop()