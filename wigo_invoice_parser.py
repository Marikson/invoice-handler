class WigoInvoiceParser:
    def __init__(self, context):
        self.context = context
    
    def __call__(self):
        return self._extract_nelson_fields(self.context)


    def _extract_nelson_fields(self, context):
            all_tables = self._get_all_tables(context)
            amount, currency = self._extract_sum_from_tables(all_tables)
            return {
                "invoice_id": self._extract_invoice_id_from_tables(all_tables),
                "deadline": self._extract_date_from_tables(all_tables),
                "sum_total": amount,
                "currency": currency,
            }

    def _extract_sum_from_tables(self, all_tables):
           if all_tables and len(all_tables) > 3 and len(all_tables[3]) > 4:
               return all_tables[3][4][0], all_tables[3][4][1].casefold()
           return None, None
       
    
    def _extract_date_from_tables(self, all_tables):
        if all_tables and len(all_tables) > 0 and len(all_tables[0]) >= 4 and len(all_tables[0][3]) >= 2:
            return all_tables[0][3][1]
        return None

    def _extract_invoice_id_from_tables(self, all_tables):
        if all_tables and len(all_tables) > 0 and len(all_tables[0]) > 0 and len(all_tables[0][1]) > 1:
            return all_tables[0][1][1]
        return None
    

    def _get_all_tables(self, context):
        if not context:
            return []
        tables = context.get("tables", [])
        return [
            [[cell for cell in row if cell != ""] for row in table]
            for table in tables
        ]