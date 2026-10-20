import os
import importlib
import re
from datetime import datetime
from text_normalizer import TextNormalizer as TN
from wigo_invoice_parser import WigoInvoiceParser
from currency_converter import CurrencyConverter
from xmlrpc.client import DateTime


try:
    pdfplumber = importlib.import_module("pdfplumber")
except ImportError as exc:
    raise ImportError(
        "pdfplumber is required for PDF parsing. Install it with `pip install -r requirements.txt`."
    ) from exc


class InvoiceData:
    def __init__(self, filename, vendor=None, invoice_id=None, amount=None, deadline=None):
        self.filename : str = filename
        self.vendor = vendor
        self.invoice_id : str = invoice_id
        self.amount : int = amount
        self.deadline  : DateTime = deadline


class PDFParser:
    VENDOR_KEYWORDS = ["kft", "bt", "zrt", "microsoft", "neda-bau", "fatelep", "centerstahl"]
    EXTRACT_KEYWORDS = ["nordvik", "bankszámlaszám"]
    TABLE_INVOICE_ID_KEYWORDS = ["számla sorszám", "számlaszám", "invoice number", "szamla sorszam"]
    TABLE_DEADLINE_KEYWORDS = ["fizetési", "fizetesi", "esedékesség", "esedekesseg", "due date"]
    
    INVOICE_ID_KEYWORDS = ["számlaszám", "számla sorszám", "sorszám", "tájékoztatásra"]
    INVOICE_ID_PATTERN = re.compile(r"\b(?=[A-Z0-9./_-]{6,}\b)(?=(?:\D*\d){5,}\D*$)[A-Z0-9./_-]+\b", re.IGNORECASE,)
    
    CURRENCY_KEYWORDS = ["huf", "ft", "forint", "eur", "usd", "gbp"]
    FINAL_SUM_KEYWORDS = [
        "számlaösszeg",
        "bruttó",
        "brutto",
        "fizetendő",
        "fizetendo",
        "összesen",
        "osszesen",
        "végösszeg",
        "vegosszeg",
        "payable",
        "grand total",
        "total due",
    ]
    PARTIAL_SUM_KEYWORDS = [
        "nettó",
        "netto",
        "egységár",
        "egysegar",
        "részösszeg",
        "reszosszeg",
        "line total",
        "item total",
        "subtotal",
        "sub-total",
        "áfa",
        "afa",
        "vat",
    ]
    SUM_TOTAL_PATTERN = re.compile(r"\b(?:\d{1,3}(?:[ .,]\d{3})+(?:[.,]\d{2})?|\d{3,9}(?:[.,]\d{2})?)\b")

    DEADLINE_KEYWORDS = ["fizetési határidő:", "fizetésihatáridő:", "teljesítési határidő", "lejárat", "esedékesség", "-ig"]
    DEADLINE_PATTERN = re.compile(
        r"\b(?:\d{4}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{1,2}|\d{1,2}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{4})\b[.]?"
    )
    DEADLINE_DATE_FORMATS = (
        "%Y-%m-%d",
        "%Y.%m.%d",
        "%Y/%m/%d",
        "%d-%m-%Y",
        "%d.%m.%Y",
        "%d/%m/%Y",
    )

    def __init__(self, attachments_folder="attachments"):
        """Initialize the PDF parser with the path to attachments folder."""
        self.attachments_folder = attachments_folder
        self.pdf_data : list[InvoiceData] = []
        self.normalizer = TN()
        
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
                    page_text = self.normalizer._normalize_extracted_text(page.extract_text() or "")
                    if page_text:
                        text_parts.append(page_text)
            text = "\n".join(text_parts)
            return text
        except Exception as e:
            print(f"Error parsing {filename}: {str(e)}")
            return None

    def build_pdf_context(self, filename):
        filepath = os.path.join(self.attachments_folder, filename)
        context = {
            "full_text": "",
            "header_text": "",
            "lower_third_text": "",
            "tables": [],
        }

        try:
            text_parts = []
            with pdfplumber.open(filepath) as pdf:
                for index, page in enumerate(pdf.pages):
                    page_text = self.normalizer._normalize_extracted_text(page.extract_text() or "")
                    if page_text:
                        text_parts.append(page_text)

                    page_tables = page.extract_tables() or []
                    context["tables"].extend(self.normalizer._normalize_table(table) for table in page_tables)

                    if index == 0:
                        header_region = page.crop((0, 0, page.width, 120))
                        context["header_text"] = self.normalizer._normalize_extracted_text(header_region.extract_text() or "")

                        lower_third_top = page.height * (2 / 3)
                        lower_third_region = page.crop((0, lower_third_top, page.width, page.height))
                        context["lower_third_text"] = self.normalizer._normalize_extracted_text(lower_third_region.extract_text() or "")

            context["full_text"] = "\n".join(text_parts)
            return context
        except Exception as e:
            print(f"Error building context for {filename}: {str(e)}")
            return context

    def extract_header_text(self, filename, header_height=120):
        filepath = os.path.join(self.attachments_folder, filename)
        try:
            with pdfplumber.open(filepath) as pdf:
                if not pdf.pages:
                    return ""

                first_page = pdf.pages[0]
                header_region = first_page.crop((0, 0, first_page.width, header_height))
                return self.normalizer._normalize_extracted_text(header_region.extract_text() or "")
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
                return self.normalizer._normalize_extracted_text(lower_third_region.extract_text() or "")
        except Exception as e:
            print(f"Error extracting lower third from {filename}: {str(e)}")
            return ""
    
    def parse_all_pdfs(self):
        """Parse all PDF files in the attachments folder."""
        pdf_files = self.get_pdf_files()
        
        for pdf_file in pdf_files:
            context = self.build_pdf_context(pdf_file)
            full_content = context["full_text"]
            header_content = context["header_text"]
            lower_third_content = context["lower_third_text"]
            if full_content:
                vendor = self.extract_vendor(header_content, full_content)
                vendor_specific_fields = self._extract_vendor_specific_fields(vendor, context)
                    
                invoice_id = vendor_specific_fields.get("invoice_id") or self.extract_invoice_id(header_content, full_content)
                vendor_sum_total = vendor_specific_fields.get("sum_total")
                vendor_currency = vendor_specific_fields.get("currency")

                if vendor_sum_total and vendor_currency:
                    sum_total, currency = vendor_sum_total, vendor_currency
                else:
                    sum_total, currency = self.extract_sum_total(lower_third_content, full_content)
                deadline = vendor_specific_fields.get("deadline") or self.extract_deadline(header_content, full_content)
                
                amount_in_huf = self.currency_converter(currency, sum_total) if sum_total else None
                
                self.append_to_pdf_data(
                    filename=pdf_file,
                    vendor=vendor,
                    invoice_id=invoice_id,
                    amount=amount_in_huf,
                    deadline=self._normalize_deadline_date(deadline) if deadline else None
                )
        
        return self.pdf_data
    
    
    def append_to_pdf_data(self, filename, vendor, invoice_id, amount, deadline):
        if not bool(invoice_id or amount):
            print(f"Skipping invoice with no ID and no amount in file: {filename}")
            return
        
        for pdf_data_entry in self.pdf_data:
            if pdf_data_entry.invoice_id == invoice_id:
                print(f"Duplicate invoice ID found: {invoice_id}, in file: {pdf_data_entry.filename} and {filename}")
                return

        
        
        self.pdf_data.append(InvoiceData(
            filename=filename,
            vendor=vendor,
            invoice_id=invoice_id,
            amount=amount,
            deadline=deadline
        ))
    
    def get_parsed_data(self):
        """Return the parsed PDF data."""
        return self.pdf_data

    def _extract_vendor_specific_fields(self, vendor, context):
        if not vendor or not context:
            return {}

        normalized_vendor = vendor.casefold()
        if "nelson flottalízing" in normalized_vendor:
            return WigoInvoiceParser(context)()

        return {}

    def _contains_exact_keyword(self, line, keywords):
        normalized_line = line.casefold()
        for keyword in keywords:
            pattern = rf"(?<!\w){re.escape(keyword.casefold())}(?!\w)"
            if re.search(pattern, normalized_line):
                return True, keyword
        return False, None

    def _is_final_sum_line(self, line):
        normalized_line = line.casefold()

        has_total_marker, total_keyword = self._contains_exact_keyword(normalized_line, self.FINAL_SUM_KEYWORDS)
        if not has_total_marker:
            return False, None

        has_partial_marker, _ = self._contains_exact_keyword(normalized_line, self.PARTIAL_SUM_KEYWORDS)
        if has_partial_marker:
            return False, None

        return True, total_keyword


    def extract_vendor(self, header_content, full_content):
        candidate_sources = [header_content, full_content]

        for content in candidate_sources:
            if not content:
                continue

            lines = [line.strip() for line in content.split("\n") if line.strip()]
            for line in lines:
                lowered_line = line.casefold()
                vendor_match = None
                for vendor_keyword in self.VENDOR_KEYWORDS:
                    vendor_index = lowered_line.find(vendor_keyword.casefold())
                    if vendor_index == -1:
                        continue
                    if vendor_match is None or vendor_index < vendor_match[0]:
                        vendor_match = (vendor_index, vendor_keyword)

                if not vendor_match:
                    continue

                excluded_indices = [lowered_line.find(keyword.casefold()) for keyword in self.EXTRACT_KEYWORDS]
                excluded_indices = [index for index in excluded_indices if index != -1]

                if excluded_indices:
                    first_excluded_index = min(excluded_indices)
                    if first_excluded_index <= vendor_match[0]:
                        continue
                    return line[:first_excluded_index].strip(" ,.;:-")

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
                if any(keyword.casefold() in lowered_line for keyword in self.INVOICE_ID_KEYWORDS):
                    tokens = re.findall(r"\b[A-Z0-9./_-]{6,}\b", line, flags=re.IGNORECASE)
                    for token in tokens:
                        digit_count = sum(ch.isdigit() for ch in token)
                        if digit_count >= 5:
                            return token
        return None

    def extract_sum_total(self, lower_third, full_content):
        candidate_sources = [lower_third, full_content]

        for content in candidate_sources:
            if not content:
                continue

            lines = [line.strip() for line in content.split("\n") if line.strip()]
            for index, line in enumerate(lines):
                contains_currency_keyword, currency_keyword = self._contains_exact_keyword(line, self.CURRENCY_KEYWORDS)
                if not contains_currency_keyword:
                    continue

                is_final_sum_line, _ = self._is_final_sum_line(line)
                if not is_final_sum_line and index > 0:
                    previous_line = lines[index - 1]
                    is_final_sum_line, _ = self._is_final_sum_line(previous_line)

                if not is_final_sum_line:
                    continue

                matches = self.SUM_TOTAL_PATTERN.findall(line)
                if matches:
                    return matches[-1], currency_keyword

        return None, None
    
    
    def currency_converter(self, currency, amount):
        currency_map = {
            "usd": "USD",
            "eur": "EUR",
            "gbp": "GBP",
            "huf": "HUF",
            "ft": "HUF",
            "forint": "HUF"
            
        }
        casted_amount = float(amount.replace('.', " ").replace(" ", "").replace(',', ".")) if amount else None
        if currency_map[currency] == "HUF":
            return int(casted_amount)
        else:
            if currency.lower() in currency_map:
                converter = CurrencyConverter()
                return round(converter.convert(casted_amount, currency_map[currency], "HUF"))
        return casted_amount
    
    
    def extract_deadline(self, header_content, full_content):
        candidate_sources = [header_content, full_content]
        collected_dates = []

        for content in candidate_sources:
            if not content:
                continue

            lines = [line.strip() for line in content.split("\n") if line.strip()]
            for line in lines:
                date_matches = self.DEADLINE_PATTERN.findall(line)
                if date_matches:
                    for date_match in date_matches:
                        collected_dates.append(date_match)

                contains_keyword, _ = self._contains_exact_keyword(line, self.DEADLINE_KEYWORDS)
                contains_unwanted_keyword, _ = self._contains_exact_keyword(line, self.EXTRACT_KEYWORDS)
                if contains_keyword and not contains_unwanted_keyword:
                    match = self.DEADLINE_PATTERN.search(line)
                    if match:
                        return match.group(0)
                        

        if collected_dates:
            return max(collected_dates)

        return None

    def _normalize_deadline_date(self, raw_date):
        date_candidate = raw_date.strip().rstrip(".")
        date_candidate = re.sub(r"\s+", "", date_candidate)

        for date_format in self.DEADLINE_DATE_FORMATS:
            try:
                parsed_date = datetime.strptime(date_candidate, date_format)
                return DateTime(parsed_date.strftime("%Y-%m-%d"))
            except ValueError:
                continue

        return None