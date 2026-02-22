from __future__ import annotations

import time
import requests


class WeComClient:
    def __init__(self, corp_id: str, secret: str, agent_id: str):
        self.corp_id = corp_id
        self.secret = secret
        self.agent_id = agent_id
        self._token = None
        self._token_exp = 0.0

    @property
    def token_url(self) -> str:
        return "https://qyapi.weixin.qq.com/cgi-bin/gettoken"

    def _get_token(self) -> str:
        now = time.time()
        if self._token and now < self._token_exp - 60:
            return self._token
        r = requests.get(self.token_url, params={"corpid": self.corp_id, "corpsecret": self.secret}, timeout=15)
        r.raise_for_status()
        data = r.json()
        if data.get("errcode") != 0:
            raise RuntimeError(f"gettoken failed: {data}")
        self._token = data["access_token"]
        self._token_exp = now + int(data.get("expires_in", 7200))
        return self._token

    def upload_image(self, file_path: str) -> str:
        token = self._get_token()
        url = "https://qyapi.weixin.qq.com/cgi-bin/media/upload"
        with open(file_path, "rb") as f:
            r = requests.post(url, params={"access_token": token, "type": "image"}, files={"media": f}, timeout=30)
        r.raise_for_status()
        data = r.json()
        if data.get("errcode") != 0:
            raise RuntimeError(f"upload failed: {data}")
        return data["media_id"]

    def send_image(self, user_id: str, media_id: str) -> None:
        token = self._get_token()
        url = "https://qyapi.weixin.qq.com/cgi-bin/message/send"
        payload = {
            "touser": user_id,
            "msgtype": "image",
            "agentid": self.agent_id,
            "image": {"media_id": media_id},
            "safe": 0,
        }
        r = requests.post(url, params={"access_token": token}, json=payload, timeout=30)
        r.raise_for_status()
        data = r.json()
        if data.get("errcode") != 0:
            raise RuntimeError(f"send failed: {data}")
