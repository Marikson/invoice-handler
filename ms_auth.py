from dataclasses import dataclass

import requests

DEFAULT_SCOPES = ["Mail.Read"]

@dataclass(frozen=True)
class AuthConfig:
    tenant_id: str | None
    client_id: str | None
    client_secret: str | None
    use_device_code: bool
    scopes: list[str]


class MSAuth:
    def __init__(self, tenant_id: str | None, client_id: str | None, client_secret: str | None, scopes: list[str] | None):
        self.auth_config = AuthConfig(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret,
                use_device_code=not bool(client_secret),
                scopes=scopes or DEFAULT_SCOPES,
            )
        self.token = self.acquire_token()
        self.session = self.build_session(self.token)
        
    
    def acquire_token(self) -> str:
        if self.auth_config.client_id and self.auth_config.client_secret:
            if not self.auth_config.tenant_id:
                raise SystemExit("--tenant-id is required with --client-id and --client-secret")
            token_url = f"https://login.microsoftonline.com/{self.auth_config.tenant_id}/oauth2/v2.0/token"
            data = {
                "client_id": self.auth_config.client_id,
                "client_secret": self.auth_config.client_secret,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            }
            response = requests.post(token_url, data=data, timeout=60)
            if response.status_code >= 400:
                raise SystemExit(f"Token request failed ({response.status_code}): {response.text}")
            return response.json()["access_token"]

        if not self.auth_config.client_id:
            raise SystemExit("--client-id is required for device code flow")

        if not self.auth_config.tenant_id:
            tenant = "common"
        else:
            tenant = self.auth_config.tenant_id

        device_code_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode"
        token_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
        scope = " ".join(self.auth_config.scopes)

        device = requests.post(device_code_url, data={"client_id": self.auth_config.client_id, "scope": scope}, timeout=60)
        if device.status_code >= 400:
            raise SystemExit(f"Device code request failed ({device.status_code}): {device.text}")

        payload = device.json()
        print(payload["message"])

        data = {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": self.auth_config.client_id,
            "device_code": payload["device_code"],
        }

        while True:
            token = requests.post(token_url, data=data, timeout=60)
            if token.status_code == 200:
                return token.json()["access_token"]

            error = token.json().get("error")
            if error in {"authorization_pending", "slow_down"}:
                continue
            raise SystemExit(f"Device sign-in failed: {token.text}")
        
    
    def build_session(self, access_token: str) -> requests.Session:
        session = requests.Session()
        session.headers.update({
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
        })
        return session