import base64
from pathlib import Path


XLS_FILE_PATH = "https://nordvikkft.sharepoint.com/:x:/s/Nordvik-Dokumentumok/IQB3bShMeKO-SZpYHsGexweMAYurSjFY_oj2cnTyMc74xFo?e=1dVd05"
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

class XlsMethods:    
    def __init__(self, session):
        self.session = session

    def _build_share_content_url(self, share_url: str) -> str:
        encoded = base64.urlsafe_b64encode(share_url.encode("utf-8")).decode("ascii")
        encoded = encoded.rstrip("=")
        return f"{GRAPH_ROOT}/shares/u!{encoded}/driveItem/content"
    
    def download_xls_file(self, save_to: str, share_url: str = XLS_FILE_PATH) -> bool:
        try:
            url = self._build_share_content_url(share_url)
            response = self.session.get(url, timeout=60)
            response.raise_for_status()
            
            destination = Path(save_to)
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(response.content)
            
            return True
        except Exception as e:
            print(f"Error downloading file: {e}")
            return False

