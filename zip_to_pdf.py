import zipfile
import os
from PIL import Image
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys

class PinkZipToPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Zip a PDF")
        self.root.geometry("550x350")
        self.root.configure(bg='#FFF0F5')  # Rosa del fondo
        
        # Elstilo para widgets ttk
        self.setup_styles()
        
        self.setup_ui()
    
    def setup_styles(self):
        style = ttk.Style()
        
        # Configurar tema
        style.theme_use('clam')
        
        # Configurar colores
        style.configure('Pink.TFrame', background='#FFF0F5')
        style.configure('Pink.TLabel', background='#FFF0F5', foreground='#8B4789', font=('Arial', 10))
        style.configure('Pink.Title.TLabel', background='#FFF0F5', foreground="#FF88D3", 
                       font=('Arial', 18, 'bold'))
        style.configure('Pink.TButton', background='#FFB6C1', foreground='#8B4789', 
                       focuscolor='none', font=('Arial', 10, 'bold'))
        style.configure('Pink.TEntry', fieldbackground='#FFF0F5', foreground='#8B4789')
        style.configure('Pink.Horizontal.TProgressbar', background='#FF69B4', troughcolor='#FFE4E1')
    
    def setup_ui(self):
        # Frame principal con color rosa
        main_frame = ttk.Frame(self.root, padding="20", style='Pink.TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título 
        title_label = ttk.Label(main_frame, text="Zip to PDF Converter", 
                               style='Pink.Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Subtítulo
        subtitle_label = ttk.Label(main_frame, text="Convierte imágenes desde archivos ZIP a PDF", 
                                  style='Pink.TLabel')
        subtitle_label.grid(row=1, column=0, columnspan=3, pady=(0, 20))
        
        # Selección de archivo ZIP
        ttk.Label(main_frame, text="Selecciona el archivo ZIP:", style='Pink.TLabel').grid(
            row=2, column=0, sticky=tk.W, pady=5)
        
        self.zip_path_var = tk.StringVar()
        zip_entry = ttk.Entry(main_frame, textvariable=self.zip_path_var, width=50, style='Pink.TEntry')
        zip_entry.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5)
        
        browse_zip_btn = ttk.Button(main_frame, text="Examinar", command=self.browse_zip_file, 
                                   style='Pink.TButton')
        browse_zip_btn.grid(row=3, column=1, padx=(10, 0), pady=5)
        
        # Selección de archivo PDF de salida
        ttk.Label(main_frame, text="Archivo PDF de salida:", style='Pink.TLabel').grid(
            row=4, column=0, sticky=tk.W, pady=5)
        
        self.pdf_path_var = tk.StringVar()
        pdf_entry = ttk.Entry(main_frame, textvariable=self.pdf_path_var, width=50, style='Pink.TEntry')
        pdf_entry.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=5)
        
        browse_pdf_btn = ttk.Button(main_frame, text="Examinar", command=self.browse_pdf_file, 
                                   style='Pink.TButton')
        browse_pdf_btn.grid(row=5, column=1, padx=(10, 0), pady=5)
        
        # Botón de conversión
        self.convert_btn = ttk.Button(main_frame, text="Convertir a PDF", 
                                     command=self.convert_zip_to_pdf, style='Pink.TButton')
        self.convert_btn.grid(row=6, column=0, columnspan=2, pady=20)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', 
                                       style='Pink.Horizontal.TProgressbar')
        self.progress.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # Etiqueta de estado
        self.status_var = tk.StringVar(value="Listo para convertir")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, style='Pink.TLabel')
        status_label.grid(row=8, column=0, columnspan=2, pady=5)
        
        # Configurar pesos de la cuadrícula
        main_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def browse_zip_file(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo ZIP",
            filetypes=[("Archivos ZIP", "*.zip"), ("Todos los archivos", "*.*")]
        )
        if file_path:
            self.zip_path_var.set(file_path)
            # Generar nombre PDF automáticamente
            base_name = os.path.splitext(file_path)[0]
            self.pdf_path_var.set(f"{base_name}.pdf")
    
    def browse_pdf_file(self):
        file_path = filedialog.asksaveasfilename(
            title="Guardar PDF como",
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
        )
        if file_path:
            self.pdf_path_var.set(file_path)
    
    def convert_zip_to_pdf(self):
        zip_path = self.zip_path_var.get()
        pdf_path = self.pdf_path_var.get()
        
        if not zip_path:
            messagebox.showerror("Error", "Por favor selecciona un archivo ZIP")
            return
        
        if not pdf_path:
            messagebox.showerror("Error", "Por favor especifica la ruta de salida del PDF")
            return
        
        # Deshabilita botón de conversión e iniciar progreso
        self.convert_btn.config(state='disabled')
        self.progress.start()
        self.status_var.set("Convirtiendo...")
        
        # Ejecuta la conversión en el hilo principal (por simplicidad)
        self.root.after(100, lambda: self.do_conversion(zip_path, pdf_path))
    
    def do_conversion(self, zip_path, pdf_path):
        try:
            success = self.zip_to_pdf(zip_path, pdf_path)
            
            if success:
                messagebox.showinfo("Éxito", f"El PDF se generó correctamente \n\nUbicación: {pdf_path}")
                self.status_var.set("Conversión completada :)")
            else:
                messagebox.showerror("Error", "Error al crear PDF. Verifica que el ZIP contenga imágenes válidas.")
                self.status_var.set("Error en la conversión")
                
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")
            self.status_var.set("Error durante la conversión")
        finally:
            # Rehabilitar botón de conversión y detener progreso
            self.convert_btn.config(state='normal')
            self.progress.stop()
    
    def zip_to_pdf(self, zip_file_path, output_pdf_path):
        """
        Convertir imágenes de un archivo zip a un solo PDF
        """
        # Crear directorio temporal para extraer imágenes
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                # Extraer archivo zip
                with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # Obtener lista de archivos de imagen
                image_files = []
                supported_formats = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.gif', '.webp')
                
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        if file.lower().endswith(supported_formats):
                            image_files.append(os.path.join(root, file))
                
                # Ordenar archivos de imagen para un orden consistente
                image_files.sort()
                
                if not image_files:
                    self.status_var.set("No se encontraron imágenes compatibles en el ZIP")
                    return False
                
                self.status_var.set(f"Procesando {len(image_files)} imágenes...")
                
                # Convertir imágenes a PDF
                images = []
                for i, image_file in enumerate(image_files):
                    try:
                        # Actualizar estado
                        self.status_var.set(f"Procesando {i+1}/{len(image_files)}: {os.path.basename(image_file)}")
                        self.root.update()
                        
                        # Abrir y convertir imagen a RGB (requerido para PDF)
                        img = Image.open(image_file)
                        if img.mode != 'RGB':
                            img = img.convert('RGB')
                        images.append(img)
                    except Exception as e:
                        print(f"Error procesando {image_file}: {e}")
                        continue
                
                if not images:
                    self.status_var.set("No se pudieron procesar las imágenes")
                    return False
                
                # Guardar como PDF
                self.status_var.set("Creando PDF...")
                first_image = images[0]
                remaining_images = images[1:]
                
                first_image.save(
                    output_pdf_path, 
                    "PDF", 
                    resolution=100.0, 
                    save_all=True, 
                    append_images=remaining_images
                )
                
                return True
                
            except zipfile.BadZipFile:
                self.status_var.set("Error: Archivo ZIP inválido")
                return False
            except Exception as e:
                self.status_var.set(f"Error: {str(e)}")
                return False

def main():
    root = tk.Tk()
    app = PinkZipToPDFConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()