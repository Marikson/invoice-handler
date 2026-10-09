import os
import importlib
import re
from pathlib import Path

try:
    pdfplumber = importlib.import_module("pdfplumber")
except ImportError as exc:
    raise ImportError(
        "pdfplumber is required for PDF parsing. Install it with `pip install -r requirements.txt`."
    ) from exc


class PDFParser:
    VENDOR_KEYWORDS = ["kft", "bt", "zrt", "microsoft", "neda-bau", "fatelep", "centerstahl"]
    EXTRACT_KEYWORDS = ["nordvik", "bankszámlaszám"]
    
    INVOICE_ID_KEYWORDS = ["számlaszám", "számla sorszám", "sorszám"]
    INVOICE_ID_PATTERN = re.compile(r"\b(?=[A-Z0-9./_-]{6,}\b)(?=(?:\D*\d){5,}\D*$)[A-Z0-9./_-]+\b", re.IGNORECASE,)
    
    SUM_TOTAL_KEYWORDS = ["huf", "ft", "forint", "eur"]
    SUM_TOTAL_PATTERN = re.compile(r"\b(?:\d{1,3}(?:[ .,]\d{3})+(?:[.,]\d{2})?|\d{3,9}(?:[.,]\d{2})?)\b")

    def __init__(self, attachments_folder="attachments"):
        """Initialize the PDF parser with the path to attachments folder."""
        self.attachments_folder = attachments_folder
        self.pdf_data = []
        
    def __call__(self):
        return self.parse_all_pdfs()
    
    
    def get_pdf_files(self):
        """Get all PDF files from the attachments folder."""
        if not os.path.exists(self.attachments_folder):
            raise FileNotFoundError(f"Folder '{self.attachments_folder}' not found")
        
        pdf_files = [f for f in os.listdir(self.attachments_folder) if f.lower().endswith('.pdf')]
        return pdf_files
    
    def parse_pdf(self, filename):
        """Parse a single PDF file and extract text."""
        filepath = os.path.join(self.attachments_folder, filename)
        
        try:
            text_parts = []
            with pdfplumber.open(filepath) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text() or ""
                    if page_text:
                        text_parts.append(page_text)
            text = "\n".join(text_parts)
            return text
        except Exception as e:
            print(f"Error parsing {filename}: {str(e)}")
            return None

    def extract_header_text(self, filename, header_height=120):
        filepath = os.path.join(self.attachments_folder, filename)
        try:
            with pdfplumber.open(filepath) as pdf:
                if not pdf.pages:
                    return ""

                first_page = pdf.pages[0]
                header_region = first_page.crop((0, 0, first_page.width, header_height))
                return header_region.extract_text() or ""
        except Exception as e:
            print(f"Error extracting header from {filename}: {str(e)}")
            return ""

    def extract_lower_third_text(self, filename):
        filepath = os.path.join(self.attachments_folder, filename)
        try:
            with pdfplumber.open(filepath) as pdf:
                if not pdf.pages:
                    return ""

                first_page = pdf.pages[0]
                lower_third_top = first_page.height * (2 / 3)
                lower_third_region = first_page.crop(
                    (0, lower_third_top, first_page.width, first_page.height)
                )
                return lower_third_region.extract_text() or ""
        except Exception as e:
            print(f"Error extracting lower third from {filename}: {str(e)}")
            return ""
    
    def parse_all_pdfs(self):
        """Parse all PDF files in the attachments folder."""
        pdf_files = self.get_pdf_files()
        
        for pdf_file in pdf_files:
            full_content = self.parse_pdf(pdf_file)
            header_content = self.extract_header_text(pdf_file)
            lower_third_content = self.extract_lower_third_text(pdf_file)
            if full_content:
                vendor = self.extract_vendor(header_content, full_content)
                invoice_id = self.extract_invoice_id(header_content, full_content)
                sum_total = self.extract_sum_total(lower_third_content, full_content)
                
                
                self.pdf_data.append({
                    'filename': pdf_file,
                    'vendor': vendor,
                    'invoice_id': invoice_id,
                    'sum_total': int(float(sum_total.replace(',', '').replace(' ', ''))) if sum_total else None
                })
        
        return self.pdf_data
    
    def get_parsed_data(self):
        """Return the parsed PDF data."""
        return self.pdf_data

    def _contains_exact_keyword(self, line, keywords):
        normalized_line = line.casefold()
        for keyword in keywords:
            pattern = rf"(?<!\w){re.escape(keyword.casefold())}(?!\w)"
            if re.search(pattern, normalized_line):
                return True
        return False


    def extract_vendor(self, header_content, full_content):
        candidate_sources = [header_content, full_content]

        for content in candidate_sources:
            if not content:
                continue

            lines = [line.strip() for line in content.split("\n") if line.strip()]
            for line in lines:
                lowered_line = line.casefold()
                if any(suffix in lowered_line for suffix in self.VENDOR_KEYWORDS) and not any(keyword.casefold() in lowered_line for keyword in self.EXTRACT_KEYWORDS):
                    return line

        return None

    def extract_invoice_id(self, header_content, full_content):
        candidate_sources = [header_content, full_content]

        for content in candidate_sources:
            if not content:
                continue

            lines = [line.strip() for line in content.split("\n") if line.strip()]
            for line in lines:
                lowered_line = line.casefold()
                if any(keyword.casefold() in lowered_line for keyword in self.INVOICE_ID_KEYWORDS) and not any(keyword.casefold() in lowered_line for keyword in self.EXTRACT_KEYWORDS):
                    match = self.INVOICE_ID_PATTERN.search(line)
                    if match:
                        return match.group(0)

        return None

    def extract_sum_total(self, lower_third, full_content):
        candidate_sources = [lower_third, full_content]

        for content in candidate_sources:
            if not content:
                continue

            lines = [line.strip() for line in content.split("\n") if line.strip()]
            for line in lines:
                if self._contains_exact_keyword(line, self.SUM_TOTAL_KEYWORDS) and not self._contains_exact_keyword(line, self.EXTRACT_KEYWORDS):
                    match = self.SUM_TOTAL_PATTERN.search(line)
                    if match:
                        return match.group(0)

        return None