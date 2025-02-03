import os
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def solicitar_version():
    def obtener_version():
        version = version_var.get()
        if version in ['3', '4']:
            procesar_xml(version)
            version_window.destroy()
        else:
            messagebox.showerror("Error", "La versión debe ser 3 o 4.")
    
    version_window = tk.Toplevel(root)
    version_window.title("Seleccionar Versión CFDI")
    version_window.geometry("300x150")
    
    version_var = tk.StringVar()

    label_version = tk.Label(version_window, text="Ingresa la versión 3 o 4 del CFDI:")
    label_version.pack(pady=10)
    
    entry_version = tk.Entry(version_window, textvariable=version_var)
    entry_version.pack(pady=10)
    
    btn_confirmar = tk.Button(version_window, text="Confirmar", command=obtener_version)
    btn_confirmar.pack(pady=10)

def procesar_xml(version):
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los archivos XML")
    if not carpeta:
        return

    datos = []
    namespace = {
        'cfdi': 'http://www.sat.gob.mx/cfd/' + version,
        'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'
    }

    for archivo in os.listdir(carpeta):
        if archivo.endswith('.xml'):
            ruta_xml = os.path.join(carpeta, archivo)
            try:
                tree = ET.parse(ruta_xml)
                root = tree.getroot()

                complemento = root.find("cfdi:Complemento", namespace)
                if complemento is not None:
                    timbre_fiscal = complemento.find("tfd:TimbreFiscalDigital", namespace)
                    uuid = timbre_fiscal.get("UUID") if timbre_fiscal is not None else "No encontrado"
                else:
                    uuid = "No encontrado"

                comprobante = root.attrib
                folio = comprobante.get("Folio", "Folio No encontrado")
                fecha = comprobante.get("Fecha", "Fecha No encontrada")
                subtotal = comprobante.get("SubTotal", "Fecha No encontrada")
                total = comprobante.get("Total", "Total No encontrado")

                # Extraer información del Emisor
                emisor = root.find("cfdi:Emisor", namespace)
                rfc_emisor = emisor.get("Rfc")
                nombre_emisor = emisor.get("Nombre")

                # Extraer información del Receptor
                receptor = root.find("cfdi:Receptor", namespace)
                rfc_receptor = receptor.get("Rfc")
                nombre_receptor = receptor.get("Nombre")

                # Agregamos los datos extraídos a la lista
                datos.append([nombre_emisor, rfc_emisor, nombre_receptor, rfc_receptor, uuid, folio, fecha, subtotal, total])
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al procesar {archivo}: {str(e)}")

    if datos:
        df = pd.DataFrame(datos, columns=["Nombre Emisor", "RFC Emisor", "Nombre Receptor", "RFC Receptor", "UUID", "Folio", "Fecha", "SubTotal", "Total"])
        ruta_excel = os.path.join(carpeta, "CFDI_Datos.xlsx")
        df.to_excel(ruta_excel, index=False, engine="openpyxl")
        messagebox.showinfo("Éxito", f"Excel guardado en:\n{ruta_excel}")
    else:
        messagebox.showwarning("Aviso", "No se encontraron archivos XML en la carpeta seleccionada.")

# Interfaz Gráfica con Tkinter
root = tk.Tk()
root.title("Extractor XML a Excel")
root.geometry("400x200")

btn_procesar = tk.Button(root, text="Seleccionar Carpeta y Procesar XML", command=solicitar_version)
btn_procesar.pack(pady=50)

root.mainloop()
