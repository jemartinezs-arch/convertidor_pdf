import fitz  # PyMuPDF
from docx import Document
import os


class OCRService:
    def __init__(self, mode="mock"):
        """
        mode:
        - mock → No usa Vision (simulado)
        - vision → Preparado para Google Vision (futuro)
        """
        self.mode = mode

    # ===============================
    # DETECTAR SI PDF TIENE TEXTO
    # ===============================
    def pdf_has_text(self, pdf_path):
        doc = fitz.open(pdf_path)
        for page in doc:
            if page.get_text().strip():
                doc.close()
                return True
        doc.close()
        return False

    # ===============================
    # PDF → TEXTO COMPLETO
    # ===============================
    def pdf_to_text(self, pdf_path):
        doc = fitz.open(pdf_path)
        full_text = ""

        for page in doc:
            text = page.get_text()

            if text.strip():
                # PDF digital
                full_text += text + "\n"
            else:
                # PDF escaneado (usar OCR)
                image_bytes = self._page_to_image_bytes(page)
                ocr_text = self.process_image(image_bytes)
                full_text += ocr_text + "\n"

        doc.close()
        return full_text.strip()

    # ===============================
    # PDF → WORD (.docx)
    # ===============================
    def pdf_to_word(self, pdf_path, output_path):
        text = self.pdf_to_text(pdf_path)

        document = Document()
        document.add_paragraph(text)
        document.save(output_path)

        return output_path

    # ===============================
    # PDF → PNG (todas las páginas)
    # ===============================
    def pdf_to_png(self, pdf_path, output_folder):
        doc = fitz.open(pdf_path)

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        image_paths = []

        for i, page in enumerate(doc):
            pix = page.get_pixmap()
            output_file = os.path.join(output_folder, f"page_{i + 1}.png")
            pix.save(output_file)
            image_paths.append(output_file)

        doc.close()
        return image_paths

    # ===============================
    # PROCESAR IMAGEN (OCR)
    # ===============================
    def process_image(self, image_bytes):
        if self.mode == "mock":
            return self._mock_ocr()
        elif self.mode == "vision":
            return self._vision_ocr(image_bytes)

    # ===============================
    # OCR SIMULADO
    # ===============================
    def _mock_ocr(self):
        return "Texto detectado por OCR (modo simulación)"

    # ===============================
    # FUTURO: GOOGLE VISION
    # ===============================
    def _vision_ocr(self, image_bytes):
        """
        Aquí irá la integración real con Google Vision API.
        No implementado aún.
        """
        raise NotImplementedError("Vision OCR aún no está activado.")

    # ===============================
    # CONVERTIR PÁGINA A IMAGEN (BYTES)
    # ===============================
    def _page_to_image_bytes(self, page):
        pix = page.get_pixmap()
        return pix.tobytes()