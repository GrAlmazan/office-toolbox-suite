import os
import sys
import threading
import tkinter as tk
from tkinter import messagebox, filedialog
import customtkinter as ctk

# --- IMPORTS DE LÓGICA ---
from docx2pdf import convert
import comtypes.client
from pypdf import PdfReader, PdfWriter
import pythoncom
from pdf2docx import Converter 
from PIL import Image
import cv2 
from pptx import Presentation
import fitz  # PyMuPDF

# --- FUNCIÓN PARA RECURSOS (ICONO) ---
def resource_path(relative_path):
    """Obtiene la ruta absoluta al recurso, funciona para dev y para PyInstaller"""
    try:
        # PyInstaller crea una carpeta temporal en _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# --- FIX CONSOLA ---
class NullWriter:
    def write(self, text): pass
    def flush(self): pass
if sys.stdout is None: sys.stdout = NullWriter()
if sys.stderr is None: sys.stderr = NullWriter()

# --- CONFIGURACIÓN VISUAL ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class OfficeToolApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # 1. Configuración de Ventana y Nombre
        self.title("OfficeToolbox") 
        
        # --- CONFIGURACIÓN DEL ICONO ---
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except:
            pass # Si no encuentra el icono, usa el default de la ventana
        
        # Dimensiones
        w, h = 980, 700 
        
        # 2. Lógica de Centrado de Pantalla
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.resizable(True, True)

        self.grid_columnconfigure(0, weight=0) 
        self.grid_columnconfigure(1, weight=1) 
        self.grid_rowconfigure(0, weight=1)

        # --- A. SIDEBAR ---
        self.sidebar_frame = ctk.CTkFrame(self, width=250, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(20, weight=1)

        # Logo / Título
        self.logo = ctk.CTkLabel(self.sidebar_frame, text="Office\nToolbox", font=ctk.CTkFont(size=22, weight="bold")) 
        self.logo.grid(row=0, column=0, padx=20, pady=(30, 20))

        # --- NAVEGACIÓN ORGANIZADA ---
        self.curr_row = 1

        # SECCIÓN 1: DOCUMENTOS
        self.add_sidebar_label("DOCUMENTOS")
        self.add_sidebar_btn("📄 Office a PDF", self.show_convert)
        self.add_sidebar_btn("📝 PDF a Word", self.show_pdfword)
        self.add_sidebar_btn("📊 PDF a PPT", self.show_pdfppt)
        self.add_sidebar_btn("🔗 Unir PDFs", self.show_merge)

        # SECCIÓN 2: MULTIMEDIA
        self.add_sidebar_label("MULTIMEDIA")
        self.add_sidebar_btn("📷 Img a PDF", self.show_imgpdf)
        self.add_sidebar_btn("🎥 Video a Foto", self.show_video)
        self.add_sidebar_btn("📥 Extraer de PDF", self.show_extract)

        # SECCIÓN 3: OPTIMIZACIÓN
        self.add_sidebar_label("OPTIMIZACIÓN")
        self.add_sidebar_btn("🚀 Optimiz. PDF", self.show_optimize)
        self.add_sidebar_btn("🖼️ Optimiz. Img", self.show_optimg)

        # --- B. ÁREA PRINCIPAL ---
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        
        self.frames = {}
        for F in (ConvertView, PdfWordView, PdfPptView, MergeView, ImgPdfView, VideoView, OptPdfView, OptImgView, ExtractView):
            page_name = F.__name__
            frame = F(parent=self.main_frame, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.show_convert()

    def add_sidebar_label(self, text):
        lbl = ctk.CTkLabel(self.sidebar_frame, text=text, anchor="w", text_color="gray60", font=ctk.CTkFont(size=11, weight="bold"))
        lbl.grid(row=self.curr_row, column=0, sticky="ew", padx=20, pady=(15, 2))
        self.curr_row += 1

    def add_sidebar_btn(self, text, command):
        btn = ctk.CTkButton(self.sidebar_frame, text=text, command=command, 
                            fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                            anchor="w", height=32, font=ctk.CTkFont(size=13))
        btn.grid(row=self.curr_row, column=0, sticky="ew", padx=10, pady=2)
        self.curr_row += 1

    def show_view(self, name):
        self.frames[name].tkraise()

    # Puentes
    def show_convert(self): self.show_view("ConvertView")
    def show_pdfword(self): self.show_view("PdfWordView")
    def show_pdfppt(self): self.show_view("PdfPptView")
    def show_merge(self): self.show_view("MergeView")
    def show_imgpdf(self): self.show_view("ImgPdfView")
    def show_video(self): self.show_view("VideoView")
    def show_optimize(self): self.show_view("OptPdfView")
    def show_optimg(self): self.show_view("OptImgView")
    def show_extract(self): self.show_view("ExtractView")

# --- PLANTILLA BASE ---
class BaseToolView(ctk.CTkFrame):
    def __init__(self, parent, controller, title, instructions):
        super().__init__(parent, fg_color="transparent")
        self.controller = controller
        
        self.grid_rowconfigure(0, weight=0) 
        self.grid_rowconfigure(1, weight=1) 
        self.grid_columnconfigure(0, weight=1)

        # 1. PANEL DE CONTROL
        self.ctrl_frame = ctk.CTkFrame(self, corner_radius=10, fg_color=("gray85", "gray17"))
        self.ctrl_frame.grid(row=0, column=0, sticky="new", padx=5, pady=5)
        
        ctk.CTkLabel(self.ctrl_frame, text=title, font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(20, 5))
        
        if instructions:
            inst_label = ctk.CTkLabel(self.ctrl_frame, text=instructions, text_color="gray", 
                                      justify="left", font=ctk.CTkFont(size=12))
            inst_label.pack(pady=(0, 15), padx=20)

        # 2. PANEL DE LOG
        self.log_frame = ctk.CTkFrame(self, corner_radius=10)
        self.log_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        self.progress = ctk.CTkProgressBar(self.log_frame)
        self.progress.set(0)
        self.progress.pack(pady=(15, 10), padx=30, fill="x")

        self.log_box = ctk.CTkTextbox(self.log_frame, height=100)
        self.log_box.pack(pady=5, padx=15, fill="both", expand=True)

        ctk.CTkButton(self.log_frame, text="Borrar Log", command=self.clear_log, 
                      height=24, width=90, fg_color="#444", hover_color="#555").pack(pady=10, padx=15, anchor="e")

    def log(self, msg): self.log_box.after(0, lambda: self._ins(msg))
    def _ins(self, msg): self.log_box.insert("end", msg + "\n"); self.log_box.see("end")
    def clear_log(self): self.log_box.delete("0.0", "end"); self.progress.set(0)
    def run_thread(self, target, *args):
        self.controller.config(cursor="watch")
        threading.Thread(target=target, args=args, daemon=True).start()
    def finish_thread(self): self.controller.config(cursor="")

# ================= VISTAS =================

class ConvertView(BaseToolView):
    def __init__(self, parent, controller):
        inst = ("1. Selecciona archivos Word, Excel o PowerPoint.\n"
                "2. Elige si quieres carpetas individuales.\n"
                "3. El programa creará los PDFs automáticamente.")
        super().__init__(parent, controller, "Office a PDF", inst)
        self.check = ctk.CTkCheckBox(self.ctrl_frame, text="Crear carpeta individual por archivo")
        self.check.pack(pady=5)
        ctk.CTkButton(self.ctrl_frame, text="📂 Seleccionar Archivos", command=self.select, height=40, font=ctk.CTkFont(weight="bold")).pack(pady=20)

    def select(self):
        files = filedialog.askopenfilenames(filetypes=[("Office", "*.docx *.pptx *.xlsx *.xls")])
        if files: self.run_thread(self.process, files)

    def process(self, files):
        pythoncom.CoInitialize()
        total = len(files)
        indiv = self.check.get()
        self.log("--- Iniciando ---")
        for i, path in enumerate(files):
            name = os.path.basename(path); self.log(f"Procesando: {name}...")
            folder = os.path.splitext(path)[0] if indiv else os.path.join(os.path.dirname(path), "PDFs_Convertidos")
            os.makedirs(folder, exist_ok=True)
            out = os.path.join(folder, os.path.splitext(name)[0] + ".pdf")
            try:
                ext = os.path.splitext(name)[1].lower()
                if ext == ".docx": convert(path, out)
                elif ext == ".pptx": self.ppt_to_pdf(path, out)
                elif ext in [".xlsx", ".xls"]: self.excel_to_pdf(path, out)
                self.log(f"✔ Listo: {name}")
            except Exception as e: self.log(f"❌ Error: {e}")
            self.progress.set((i+1)/total)
        self.finish_thread()
    def ppt_to_pdf(self, i, o):
        ppt = comtypes.client.CreateObject("PowerPoint.Application"); ppt.Visible = 1
        deck = ppt.Presentations.Open(os.path.abspath(i)); deck.SaveAs(os.path.abspath(o), 32); deck.Close(); ppt.Quit()
    def excel_to_pdf(self, i, o):
        xl = comtypes.client.CreateObject("Excel.Application"); xl.Visible = False
        wb = xl.Workbooks.Open(os.path.abspath(i))
        for ws in wb.Worksheets: ws.PageSetup.Zoom = False; ws.PageSetup.FitToPagesWide = 1; ws.PageSetup.FitToPagesTall = False
        wb.ExportAsFixedFormat(0, os.path.abspath(o)); wb.Close(False); xl.Quit()

class PdfWordView(BaseToolView):
    def __init__(self, parent, controller):
        inst = ("1. Convierte PDFs a documentos Word editables (.docx).\n"
                "2. Ideal para documentos de texto digitales.")
        super().__init__(parent, controller, "PDF a Word", inst)
        ctk.CTkButton(self.ctrl_frame, text="📂 Seleccionar PDFs", command=self.select, height=40, fg_color="#2B6CB0").pack(pady=20)
    def select(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if files: self.run_thread(self.process, files)
    def process(self, files):
        self.log("--- Iniciando ---")
        for i, path in enumerate(files):
            self.log(f"Convirtiendo: {os.path.basename(path)}...")
            try:
                cv = Converter(path)
                cv.convert(os.path.splitext(path)[0]+".docx", start=0, end=None); cv.close(); self.log("✔ Éxito")
            except Exception as e: self.log(f"❌ Error: {e}")
            self.progress.set((i+1)/len(files))
        self.finish_thread()

class PdfPptView(BaseToolView):
    def __init__(self, parent, controller):
        inst = ("1. Crea PowerPoint idéntico visualmente al PDF.\n"
                "2. Cada página se convierte en imagen (no editable).")
        super().__init__(parent, controller, "PDF a PowerPoint", inst)
        ctk.CTkButton(self.ctrl_frame, text="📂 Seleccionar PDFs", command=self.select, height=40, fg_color="#C05621").pack(pady=20)
    def select(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if files: self.run_thread(self.process, files)
    def process(self, files):
        self.log("--- Iniciando ---")
        for i, path in enumerate(files):
            self.log(f"Procesando: {os.path.basename(path)}...")
            out = os.path.splitext(path)[0] + ".pptx"
            try:
                doc = fitz.open(path); prs = Presentation()
                if len(doc)>0:
                    prs.slide_width = int(doc[0].rect.width * 12700)
                    prs.slide_height = int(doc[0].rect.height * 12700)
                for pn in range(len(doc)):
                    pix = doc[pn].get_pixmap(matrix=fitz.Matrix(2,2))
                    tmp = f"tmp_{pn}.png"; pix.save(tmp)
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    slide.shapes.add_picture(tmp, 0, 0, width=prs.slide_width, height=prs.slide_height)
                    os.remove(tmp)
                prs.save(out); doc.close(); self.log("✔ Creado")
            except Exception as e: self.log(f"❌ Error: {e}")
            self.progress.set((i+1)/len(files))
        self.finish_thread()

class MergeView(BaseToolView):
    def __init__(self, parent, controller):
        inst = ("1. Selecciona 2 o más archivos PDF.\n"
                "2. Se unirán en un solo archivo final.")
        super().__init__(parent, controller, "Unir PDFs", inst)
        ctk.CTkButton(self.ctrl_frame, text="📂 Seleccionar PDFs", command=self.select, height=40).pack(pady=20)
    def select(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if len(files)<2: return messagebox.showwarning("Info", "Selecciona 2+ archivos")
        save = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if save: self.run_thread(self.process, files, save)
    def process(self, files, save):
        self.log("Fusionando...")
        try:
            m = PdfWriter()
            for f in files: m.append(f)
            m.write(save); m.close(); self.log(f"✔ Guardado: {save}")
        except Exception as e: self.log(f"❌ Error: {e}")
        self.progress.set(1.0); self.finish_thread()

class ImgPdfView(BaseToolView):
    def __init__(self, parent, controller):
        inst = ("1. Selecciona múltiples imágenes.\n"
                "2. Se unirán en un solo documento PDF.")
        super().__init__(parent, controller, "Imágenes a PDF", inst)
        ctk.CTkButton(self.ctrl_frame, text="📂 Seleccionar Imágenes", command=self.select, height=40, fg_color="#D97706").pack(pady=20)
    def select(self):
        files = filedialog.askopenfilenames(filetypes=[("Img", "*.jpg *.png *.bmp")])
        if files:
            save = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
            if save: self.run_thread(self.process, files, save)
    def process(self, files, save):
        self.log("Creando PDF...")
        try:
            imgs = [Image.open(f).convert('RGB') for f in files]
            imgs[0].save(save, save_all=True, append_images=imgs[1:], resolution=100.0)
            self.log(f"✔ Creado: {save}")
        except Exception as e: self.log(f"❌ Error: {e}")
        self.progress.set(1.0); self.finish_thread()

class VideoView(BaseToolView):
    def __init__(self, parent, controller):
        inst = ("1. Elige el intervalo de segundos.\n"
                "2. Selecciona un video para extraer fotos.")
        super().__init__(parent, controller, "Video a Fotos", inst)
        self.spin = ctk.CTkComboBox(self.ctrl_frame, values=["0.5", "1", "2", "5", "10", "30"])
        self.spin.set("1"); self.spin.pack(pady=5)
        ctk.CTkLabel(self.ctrl_frame, text="Segundos entre fotos", text_color="gray").pack()
        ctk.CTkButton(self.ctrl_frame, text="📂 Seleccionar Video", command=self.select, height=40, fg_color="#805AD5").pack(pady=15)
    def select(self):
        files = filedialog.askopenfilenames(filetypes=[("Video", "*.mp4 *.avi *.mov")])
        if files: self.run_thread(self.process, files)
    def process(self, files):
        sec = float(self.spin.get())
        for i, path in enumerate(files):
            self.log(f"Analizando: {os.path.basename(path)}")
            folder = os.path.splitext(path)[0] + "_Frames"; os.makedirs(folder, exist_ok=True)
            cap = cv2.VideoCapture(path); fps = cap.get(cv2.CAP_PROP_FPS) or 30; interval = int(fps * sec) or 1
            cnt = 0; saved = 0
            while cap.isOpened():
                ret, frame = cap.read()
                if not ret: break
                if cnt % interval == 0:
                    cv2.imwrite(os.path.join(folder, f"f_{saved:04d}.jpg"), frame); saved += 1
                cnt += 1
            cap.release(); self.log(f"✔ {saved} fotos extraídas.")
            self.progress.set((i+1)/len(files))
        self.finish_thread()

class OptPdfView(BaseToolView):
    def __init__(self, parent, controller):
        inst = ("1. Elige nivel de compresión (Alta reduce más peso).\n"
                "2. Alta = Menos calidad / Baja = Mejor calidad.")
        super().__init__(parent, controller, "Optimizar PDF", inst)
        
        # Opciones Claras
        self.qual = ctk.CTkComboBox(self.ctrl_frame, width=280, values=[
            "Compresión Baja (Mejor Calidad / Mayor Peso)", 
            "Compresión Media (Equilibrado)", 
            "Compresión Alta (Menor Calidad / Menor Peso)"
        ])
        self.qual.set("Compresión Media (Equilibrado)")
        self.qual.pack(pady=5)
        
        self.check = ctk.CTkCheckBox(self.ctrl_frame, text="Carpeta individual")
        self.check.pack(pady=5)
        
        ctk.CTkButton(self.ctrl_frame, text="📂 Seleccionar PDFs", command=self.select, height=40, fg_color="#E07A5F").pack(pady=15)

    def select(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if files: self.run_thread(self.process, files)

    def process(self, files):
        lvl = self.qual.get(); indiv = self.check.get()
        for i, path in enumerate(files):
            name = os.path.basename(path); self.log(f"Optimizando: {name}")
            folder = os.path.splitext(path)[0] if indiv else os.path.join(os.path.dirname(path), "PDFs_Optimizados")
            os.makedirs(folder, exist_ok=True)
            is_high = "Alta" in lvl
            out = os.path.join(folder, os.path.splitext(name)[0] + ("_min.pdf" if is_high else "_opt.pdf"))
            try:
                r = PdfReader(path); w = PdfWriter()
                for p in r.pages: w.add_page(p)
                for p in w.pages: 
                    p.compress_content_streams()
                    if is_high and "/Annots" in p: del p["/Annots"]
                if is_high: w.add_metadata({})
                else: 
                    try: w.add_metadata(r.metadata)
                    except: pass
                with open(out, "wb") as f: w.write(f)
                self.log(f"✔ Listo")
            except Exception as e: self.log(f"❌ Error: {e}")
            self.progress.set((i+1)/len(files))
        self.finish_thread()

class OptImgView(BaseToolView):
    def __init__(self, parent, controller):
        inst = ("1. Reduce el peso de imágenes JPG o PNG.\n"
                "2. Alta Compresión = Menos peso.")
        super().__init__(parent, controller, "Optimizar Imágenes", inst)
        
        # Opciones Claras
        self.qual = ctk.CTkComboBox(self.ctrl_frame, width=280, values=[
            "Compresión Baja (Mejor Calidad)", 
            "Compresión Media (Equilibrado)", 
            "Compresión Alta (Menor Peso)"
        ])
        self.qual.set("Compresión Media (Equilibrado)")
        self.qual.pack(pady=5)
        
        ctk.CTkButton(self.ctrl_frame, text="📂 Seleccionar Imágenes", command=self.select, height=40, fg_color="#D97706").pack(pady=15)

    def select(self):
        files = filedialog.askopenfilenames(filetypes=[("Img", "*.jpg *.png")])
        if files: self.run_thread(self.process, files)

    def process(self, files):
        q_sel = self.qual.get()
        # Lógica inversa: Alta Compresión = Baja Calidad (Quality number low)
        q_val = 50 if "Alta" in q_sel else (90 if "Baja" in q_sel else 75)
        
        for i, path in enumerate(files):
            name = os.path.basename(path); self.log(f"Procesando: {name}")
            folder = os.path.join(os.path.dirname(path), "_Optimizadas"); os.makedirs(folder, exist_ok=True)
            try:
                img = Image.open(path)
                img.save(os.path.join(folder, name), optimize=True, quality=q_val)
                self.log("✔ Listo")
            except Exception as e: self.log(f"❌ Error: {e}")
            self.progress.set((i+1)/len(files))
        self.finish_thread()

class ExtractView(BaseToolView):
    def __init__(self, parent, controller):
        inst = ("1. Extrae imágenes incrustadas en un PDF.\n"
                "2. Se guardan en una carpeta '_Imagenes'.")
        super().__init__(parent, controller, "Extraer Imágenes", inst)
        ctk.CTkButton(self.ctrl_frame, text="📂 Seleccionar PDFs", command=self.select, height=40, fg_color="#2CC985").pack(pady=20)
    def select(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if files: self.run_thread(self.process, files)
    def process(self, files):
        for i, path in enumerate(files):
            self.log(f"Analizando: {os.path.basename(path)}")
            folder = os.path.splitext(path)[0] + "_Imagenes"; os.makedirs(folder, exist_ok=True)
            cnt = 0
            try:
                r = PdfReader(path)
                for pn, p in enumerate(r.pages):
                    if hasattr(p, "images"):
                        for img in p.images:
                            with open(os.path.join(folder, f"P{pn}_{img.name}"), "wb") as f: f.write(img.data); cnt += 1
                self.log(f"✔ {cnt} imgs.")
                if cnt == 0: os.rmdir(folder)
            except Exception as e: self.log(f"❌ Error: {e}")
            self.progress.set((i+1)/len(files))
        self.finish_thread()

if __name__ == "__main__":
    app = OfficeToolApp()
    app.mainloop()