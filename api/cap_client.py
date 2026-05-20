# api/cap_client.py
import requests


class CapConnectionError(Exception):
    pass


class CapClient:
    def __init__(self, server_url: str, timeout: int = 30):
        self.base = server_url.rstrip("/") + "/if-mapping"
        self.timeout = timeout

    def ping(self) -> bool:
        try:
            r = requests.get(self.base, timeout=5)
            return r.status_code < 500
        except requests.RequestException:
            return False

    def match(self, fields: list[dict], provider: str, language: str) -> list[dict]:
        try:
            r = requests.post(
                f"{self.base}/match",
                json={"fields": fields, "provider": provider, "language": language},
                timeout=self.timeout,
            )
            r.raise_for_status()
            return r.json().get("value", [])
        except requests.HTTPError as e:
            raise CapConnectionError(str(e)) from e

    def upload_custom_fields(self, records: list[dict], mode: str = "upsert") -> dict:
        try:
            r = requests.post(
                f"{self.base}/uploadCustomFields",
                json={"records": records, "mode": mode},
                timeout=self.timeout,
            )
            r.raise_for_status()
            return r.json()
        except requests.HTTPError as e:
            raise CapConnectionError(str(e)) from e

    def get_prompts(self, language: str | None = None) -> list[dict]:
        params = {}
        if language:
            params["$filter"] = f"language eq '{language}'"
        r = requests.get(f"{self.base}/PromptTemplates", params=params, timeout=self.timeout)
        r.raise_for_status()
        return r.json().get("value", [])

    def patch_prompt(self, prompt_id: str, content: str) -> dict:
        r = requests.patch(
            f"{self.base}/PromptTemplates('{prompt_id}')",
            json={"content": content},
            timeout=self.timeout,
        )
        r.raise_for_status()
        return r.json()

    def reload_prompts(self) -> None:
        r = requests.post(f"{self.base}/reloadPrompts", timeout=self.timeout)
        r.raise_for_status()

    def get_token_logs(self) -> list[dict]:
        r = requests.get(f"{self.base}/TokenLogs", timeout=self.timeout)
        r.raise_for_status()
        return r.json().get("value", [])
