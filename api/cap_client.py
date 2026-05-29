# api/cap_client.py
import time
import requests


class CapConnectionError(Exception):
    pass


class CapClient:
    def __init__(self, server_url: str, timeout: int = 600, xsuaa: dict | None = None):
        self.root = server_url.rstrip("/")
        self.base = self.root + "/if-mapping"
        self.timeout = timeout
        self._xsuaa = xsuaa          # {"url": ..., "clientid": ..., "clientsecret": ...}
        self._token: str | None = None
        self._token_expiry: float = 0.0

    # ── Auth ──────────────────────────────────────────────────────────────────

    def _get_token(self) -> str | None:
        if not self._xsuaa:
            return None
        if self._token and time.time() < self._token_expiry - 30:
            return self._token
        r = requests.post(
            self._xsuaa["url"].rstrip("/") + "/oauth/token",
            data={"grant_type": "client_credentials"},
            auth=(self._xsuaa["clientid"], self._xsuaa["clientsecret"]),
            timeout=15,
        )
        r.raise_for_status()
        data = r.json()
        self._token = data["access_token"]
        self._token_expiry = time.time() + data.get("expires_in", 3600)
        return self._token

    def _headers(self, extra: dict | None = None) -> dict:
        headers: dict = {}
        token = self._get_token()
        if token:
            headers["Authorization"] = f"Bearer {token}"
        if extra:
            headers.update(extra)
        return headers

    # ── API calls ─────────────────────────────────────────────────────────────

    def ping(self) -> bool:
        try:
            r = requests.get(self.base, headers=self._headers(), timeout=5)
            return r.status_code < 500
        except requests.RequestException:
            return False

    def match(self, fields: list[dict], provider: str, language: str,
              correlation_id: str | None = None) -> list[dict]:
        extra = {"x-correlation-id": correlation_id} if correlation_id else None
        try:
            r = requests.post(
                f"{self.base}/match",
                json={"fields": fields, "provider": provider, "language": language},
                headers=self._headers(extra),
                timeout=self.timeout,
            )
            r.raise_for_status()
            return r.json().get("value", [])
        except requests.HTTPError as e:
            raise CapConnectionError(str(e)) from e

    def open_log_stream(self, correlation_id: str) -> requests.Response:
        return requests.get(
            f"{self.root}/log-stream",
            params={"correlationId": correlation_id},
            headers=self._headers(),
            stream=True,
            timeout=(5, 600),
        )

    def upload_custom_fields(self, records: list[dict], mode: str = "upsert") -> dict:
        try:
            r = requests.post(
                f"{self.base}/uploadCustomFields",
                json={"records": records, "mode": mode},
                headers=self._headers(),
                timeout=self.timeout,
            )
            r.raise_for_status()
            return r.json()
        except requests.HTTPError as e:
            raise CapConnectionError(str(e)) from e

    def get_prompts(self, language: str | None = None) -> list[dict]:
        url = f"{self.base}/PromptTemplates"
        if language:
            url += f"?$filter=language eq '{language}'"
        r = requests.get(url, headers=self._headers(), timeout=self.timeout)
        r.raise_for_status()
        return r.json().get("value", [])

    def patch_prompt(self, prompt_id: str, content: str) -> dict:
        r = requests.patch(
            f"{self.base}/PromptTemplates('{prompt_id}')",
            json={"content": content},
            headers=self._headers(),
            timeout=self.timeout,
        )
        r.raise_for_status()
        return r.json()

    def reload_prompts(self) -> None:
        r = requests.post(
            f"{self.base}/reloadPrompts",
            headers=self._headers(),
            timeout=self.timeout,
        )
        r.raise_for_status()

    def get_token_logs(self) -> list[dict]:
        r = requests.get(
            f"{self.base}/TokenLogs",
            headers=self._headers(),
            timeout=self.timeout,
        )
        r.raise_for_status()
        return r.json().get("value", [])
