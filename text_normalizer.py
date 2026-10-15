import re
import unicodedata


class TextNormalizer:
    def _normalize_table(self, table):
            normalized_rows = []
            for row in table:
                normalized_row = [self._normalize_cell(cell) for cell in row]
                if any(normalized_row):
                    normalized_rows.append(normalized_row)
            return normalized_rows
    
    def _normalize_cell(self, cell):
        if cell is None:
            return ""
        return re.sub(r"\s+", " ", self._normalize_extracted_text(cell, preserve_newlines=False)).strip()

    def _normalize_extracted_text(self, text, preserve_newlines=True):
        normalized = unicodedata.normalize("NFKC", text)
        if preserve_newlines:
            normalized = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", " ", normalized)
            normalized = re.sub(r"[^\S\n]+", " ", normalized)
            normalized = re.sub(r" *\n+ *", "\n", normalized)
            return normalized.strip()

        normalized = re.sub(r"[\x00-\x1f\x7f]", " ", normalized)
        normalized = re.sub(r"\s+", " ", normalized)
        return normalized.strip()