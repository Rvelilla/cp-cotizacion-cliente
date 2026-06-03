import re
import os
import logging
from datetime import datetime
from typing import Dict, Any, List
import pdfplumber
import docx 
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

class CotizacionExtractor:
    MESES_ES = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio",
        7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }

    def __init__(self, pdf_path: str, nombre_asesor: str, cargo_asesor: str):
        self.pdf_path = pdf_path
        self.datos: Dict[str, Any] = {
            "ciudad": "Medellín",
            "fecha": self._obtener_fecha_actual(),
            "nombre_cliente": "", "largo_ext": "", "ancho_ext": "", "alto_ext": "",
            "tipo_carroceria": "ESTACA", "peso_kg": "",
            "valor_c_rra": "", "iva": "", "subtotal": "",
            "nombre_asesor": nombre_asesor, "cargo_asesor": cargo_asesor,
            "especificaciones": [], "tipo_garantia": "", "tiempo_garantia": ""
        }
        self.texto_plano = ""
        self.texto_layout = ""

    def _obtener_fecha_actual(self) -> str:
        hoy = datetime.now()
        return f"{hoy.day} de {self.MESES_ES[hoy.month]} de {hoy.year}"

    def procesar(self) -> str:
        self._extraer_textos_pdf()
        self._parsear_datos_basicos()
        self._parsear_especificaciones_exactas()
        self._configurar_garantia_e_imagenes()
        return self._generar_documento()

    def _extraer_textos_pdf(self) -> None:
        with pdfplumber.open(self.pdf_path) as pdf:
            self.texto_plano = pdf.pages[0].extract_text() or ""
            self.texto_layout = pdf.pages[0].extract_text(layout=True) or ""

    def _parsear_datos_basicos(self) -> None:
        texto_unido = self.texto_plano + "\n" + self.texto_layout
        if re.search(r'\bBogot[áa]\b', texto_unido, re.IGNORECASE):
            self.datos["ciudad"] = "Bogotá"

        def extraer_valor_layout(keyword: str) -> str:
            kw_esc = re.escape(keyword)
            patron = rf'(?<!\w){kw_esc}[:\s]+([^\s]+(?:\s[^\s]+)*)'
            for linea in self.texto_layout.split('\n'):
                match = re.search(patron, linea, re.IGNORECASE)
                if match:
                    val = match.group(1).strip()
                    basuras_columna = ["CÉDULA", "CEDULA", "NIT", "TELÉFONO", "EMAIL", "DIRECCIÓN", "LIMPIAR", "VOLVER", "PEDIDO", "CAPACIDAD", "VERSION"]
                    for b in basuras_columna:
                        if b in val.upper():
                            val = val[:val.upper().find(b)].strip()
                    if val:
                        return val
            return ""

        tipo = extraer_valor_layout("Tipo:")
        if not tipo: tipo = extraer_valor_layout("Tipo")
        if tipo: self.datos["tipo_carroceria"] = tipo

        nombre = extraer_valor_layout("Nombre")
        if nombre: self.datos["nombre_cliente"] = nombre

        patrones_medidas = {
            "largo_ext": r'Largo Externo[^\d]*(\d+[,.]\d+)',
            "ancho_ext": r'Ancho Ext[^\d]*(\d+[,.]\d+)',
            "alto_ext": r'Alto Externo[^\d]*(\d+[,.]\d+)'
        }
        for clave, patron in patrones_medidas.items():
            match = re.search(patron, texto_unido, re.IGNORECASE)
            if match: self.datos[clave] = match.group(1).strip()

        match_peso = re.search(r'Peso Kg.*?(±\s*\d+\s*Kg).*?(\d+)', self.texto_plano, re.IGNORECASE | re.DOTALL)
        if match_peso:
            self.datos["peso_kg"] = f"{match_peso.group(2)} ({match_peso.group(1).replace(' ', '')})"

        def extraer_moneda(keyword: str) -> str:
            patron = rf'{keyword}[^\$]*\$?\s*(\d{{1,3}}(?:\.\d{{3}})+)'
            match = re.search(patron, texto_unido, re.IGNORECASE | re.DOTALL)
            return f"$ {match.group(1).strip()}" if match else ""

        self.datos["valor_c_rra"] = extraer_moneda("Subtotal")
        self.datos["iva"]         = extraer_moneda("Iva")
        self.datos["subtotal"]    = extraer_moneda("Total Precio Venta")

    def _parsear_especificaciones_exactas(self) -> None:
        estado = "INICIO"
        concepto_actual = ""
        detalles_actuales = []
        complementos = []

        def guardar_concepto():
            if concepto_actual and detalles_actuales:
                self.datos["especificaciones"].append({
                    "concepto": concepto_actual,
                    "detalle": " - ".join(detalles_actuales)
                })

        for linea in self.texto_layout.split('\n'):
            l = linea.strip()
            if not l: continue

            if "CONCEPTO" in l.upper() and ("CANTIDAD" in l.upper() or "MEDIDA" in l.upper() or "DETALLE" in l.upper()):
                estado = "CONCEPTOS"
                continue
            elif "OTROS" in l.upper() and len(re.split(r'\s{2,}', l)) < 4:
                guardar_concepto()
                concepto_actual = ""
                detalles_actuales = []
                estado = "OTROS"
                continue
            elif "VALOR TOTAL" in l.upper() or "SUBTOTAL" in l.upper() or "% DE DESCUENTO" in l.upper():
                break

            partes_reales = [p.strip() for p in re.split(r'\s{2,}', l) if p.strip()]
            if not partes_reales: continue

            texto_raw = partes_reales[0].strip()
            upper_partes = [p.upper() for p in partes_reales]

            if estado == "CONCEPTOS":
                if texto_raw.upper() in ["CONCEPTO", "DETALLE", "CANTIDAD"]: continue
                
                es_concepto = (len(partes_reales) == 1) and not bool(re.match(r'^[\d\s*]+', texto_raw))
                
                if es_concepto:
                    guardar_concepto()
                    concepto_actual = texto_raw
                    detalles_actuales = []
                else:
                    if len(partes_reales) == 1:
                        continue
                        
                    col_2 = upper_partes[1]
                    if col_2 in ["NO", "CANTIDAD"]: continue
                    
                    cantidad = ""
                    if col_2 == "X":
                        if len(partes_reales) > 2:
                            cantidad = partes_reales[2]
                            if not re.match(r'^\d', cantidad) and len(partes_reales) > 3:
                                if re.match(r'^\d', partes_reales[3]): cantidad = partes_reales[3]
                    else:
                        if re.match(r'^\d', col_2):
                            cantidad = col_2
                    
                    if cantidad:
                        match_num = re.search(r'^(\d+[,.]?\d*)', cantidad)
                        cantidad = match_num.group(1) if match_num else ""

                    texto_limpio = re.sub(r'^[\d\s*]+', '', texto_raw).strip()
                    texto_limpio = re.sub(r'\bX\b', '', texto_limpio, flags=re.IGNORECASE).strip()
                    texto_limpio = re.sub(r'\s{2,}', ' ', texto_limpio)
                        
                    detalle_final = f"{cantidad} {texto_limpio}".strip() if cantidad else texto_limpio
                    if detalle_final:
                        detalles_actuales.append(detalle_final)

            elif estado == "OTROS":
                if "NO" in upper_partes or "CANTIDAD" in upper_partes: continue
                
                # Unimos todo para analizarlo (Resuelve el problema de líneas largas como GP A. Espejo)
                texto_completo = " ".join(partes_reales)
                
                # Descartamos headers vacíos (Ej. "Accesorios")
                texto_check = re.sub(r'^[\d\s*]+', '', texto_completo).strip().upper()
                if not texto_check or texto_check in ["OTROS", "ACCESORIOS", "ADICIONALES LINEA", "ADICIONALES CENTRO DE SERVICIOS"]:
                    continue

                cantidad = ""
                texto_item = ""
                
                # Búsqueda Regex Inteligente: Corta la cadena si encuentra un " X [número]"
                match_x = re.search(r'(?:^|\s)X\s+(\d+[,.]?\d*)', texto_completo, re.IGNORECASE)
                if match_x:
                    cantidad = match_x.group(1)
                    texto_item = texto_completo[:match_x.start()].strip()
                else:
                    if len(partes_reales) == 1:
                        continue # Header sin cantidades a su derecha
                    
                    texto_item = partes_reales[0]
                    if not re.sub(r'^[\d\s*]+', '', texto_item).strip() and len(partes_reales) > 1:
                        texto_item = partes_reales[1]
                        
                    # Extrae la cantidad si no tiene X (Ej. Carpa)
                    for p in partes_reales[1:]:
                        if p.upper() in ["X", "NO", "UNID", "MTS", "M2"]: continue
                        if re.match(r'^\d+[,.]?\d*', p):
                            cantidad = re.search(r'^(\d+[,.]?\d*)', p).group(1)
                            break
                            
                texto_limpio = re.sub(r'^[\d\s*]+', '', texto_item).strip()
                texto_limpio = re.sub(r'\bX\b', '', texto_limpio, flags=re.IGNORECASE).strip()
                texto_limpio = re.sub(r'\s{2,}', ' ', texto_limpio)
                
                if not texto_limpio: continue

                detalle_final = f"* {cantidad} {texto_limpio}".strip() if cantidad else f"* {texto_limpio}"
                
                if detalle_final not in complementos:
                    complementos.append(detalle_final)

        guardar_concepto()
        
        if complementos:
            self.datos["especificaciones"].append({
                "concepto": "Complementos del furgón",
                "detalle": "\n".join(complementos)
            })

    def _configurar_garantia_e_imagenes(self) -> None:
        tipo = self.datos.get("tipo_carroceria", "").upper()
        if "ISO" in tipo:
            self.datos["tipo_garantia"], self.datos["tiempo_garantia"], self.prefijo_img = "furgón isotérmico", "dos años", "iso"
        elif "FURG" in tipo:
            self.datos["tipo_garantia"], self.datos["tiempo_garantia"], self.prefijo_img = "furgón", "un año", "furgon"
        else:
            self.datos["tipo_garantia"], self.datos["tiempo_garantia"], self.prefijo_img = "estaca", "un año", "estaca"

    def _generar_documento(self) -> str:
        ruta_plantilla = "plantillas/plantilla_cotizacion.docx"
        doc_tpl = DocxTemplate(ruta_plantilla)

        for i in range(1, 4):
            img_path = f"assets/imagenes_modelos/{self.prefijo_img}_{i}.jpg"
            self.datos[f"imagen_{i}"] = InlineImage(doc_tpl, img_path, width=Mm(50)) if os.path.exists(img_path) else ""

        doc_tpl.render(self.datos)
        nombre_limpio = re.sub(r'[^\w\s-]', '', self.datos["nombre_cliente"]).strip().replace(' ', '_')
        output_path = f"Cotizacion_{nombre_limpio}.docx"
        doc_tpl.save(output_path)

        doc_final = docx.Document(output_path)
        for tabla in doc_final.tables:
            if len(tabla.rows) > 0 and "Concepto" in tabla.rows[0].cells[0].text:
                while len(tabla.rows) > 1:
                    tr = tabla.rows[1]._tr
                    tr.getparent().remove(tr)
                
                for spec in self.datos["especificaciones"]:
                    nueva_fila = tabla.add_row()
                    nueva_fila.cells[0].text = spec["concepto"]
                    nueva_fila.cells[1].text = spec["detalle"]
                break
                
        doc_final.save(output_path)
        return output_path

def procesar_cotizacion(pdf_path: str, nombre_asesor: str, cargo_asesor: str) -> str:
    return CotizacionExtractor(pdf_path, nombre_asesor, cargo_asesor).procesar()