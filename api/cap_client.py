# api/cap_client.py
import requests


class CapConnectionError(Exception):
    pass


class CapClient:
    def __init__(self, server_url: str, timeout: int = 600):
        self.root = server_url.rstrip("/")
        self.base = self.root + "/if-mapping"
        self.timeout = timeout

    def ping(self) -> bool:
        try:
            r = requests.get(self.base, timeout=5)
            return r.status_code < 500
        except requests.RequestException:
            return False

    def match(self, fields: list[dict], provider: str, language: str,
              correlation_id: str | None = None) -> list[dict]:
        kwargs: dict = {"json": {"fields": fields, "provider": provider, "language": language},
                        "timeout": self.timeout}
        if correlation_id:
            kwargs["headers"] = {"x-correlation-id": correlation_id}
        try:
            r = requests.post(f"{self.base}/match", **kwargs)
            r.raise_for_status()
            return r.json().get("value", [])
        except requests.HTTPError as e:
            raise CapConnectionError(str(e)) from e

    def open_log_stream(self, correlation_id: str) -> requests.Response:
        return requests.get(
            f"{self.root}/log-stream",
            params={"correlationId": correlation_id},
            stream=True,
            timeout=(5, 600),
        )

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
        url = f"{self.base}/PromptTemplates"
        if language:
            url += f"?$filter=language eq '{language}'"
        r = requests.get(url, timeout=self.timeout)
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
