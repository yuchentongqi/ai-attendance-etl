import os
import requests
import time
from dotenv import load_dotenv

# 加载 .env 文件中的环境变量
load_dotenv()

APP_KEY = os.getenv("DINGTALK_APP_KEY")
APP_SECRET = os.getenv("DINGTALK_APP_SECRET")

_cached_token = None
_expires_at = 0

def get_cached_token():
    global _cached_token, _expires_at
    if _cached_token and time.time() < _expires_at:
        return _cached_token

    url = "https://oapi.dingtalk.com/gettoken"
    params = {"appkey": APP_KEY, "appsecret": APP_SECRET}
    resp = requests.get(url, params=params, timeout=10)
    data = resp.json()

    if data.get("errcode") != 0:
        raise RuntimeError(f"获取 token 失败: {data}")

    _cached_token = data["access_token"]
    _expires_at = time.time() + data.get("expires_in", 7200) - 300
    return _cached_token