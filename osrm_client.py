"""Bounded OSRM startup checks and requests; never substitutes another travel mode."""
import math
import os
import threading
import time
from urllib.parse import urlparse

import requests
import streamlit as st

_ready = {}
_lock = threading.RLock()


def setting(key, default=None):
    try:
        value = st.session_state.get(key)
    except Exception:
        value = None
    value = value or os.getenv(key)
    if value is None:
        try:
            value = st.secrets.get(key)
        except Exception:
            pass
    return default if value is None or value == '' else value


def endpoint():
    return str(setting('CAREFLOW_OSRM_MOTORCYCLE_HOST', 'http://127.0.0.1:5001')).rstrip('/')


def is_remote(host):
    return urlparse(host).hostname not in ('127.0.0.1', 'localhost', '::1')


def seconds(key, default):
    try:
        value = float(setting(key, default))
        return min(90.0, max(1.0, value)) if math.isfinite(value) else default
    except (ValueError, TypeError):
        return default


def table_limit(host=None):
    try:
        return max(2, min(90, int(setting('CAREFLOW_OSRM_TABLE_MAX_COORDS', 40 if is_remote(host or endpoint()) else 90))))
    except (ValueError, TypeError):
        return 40 if is_remote(host or endpoint()) else 90


def request_json(host, path, params=None, local_timeout=10, startup=False):
    remote = is_remote(host)
    read = seconds('CAREFLOW_OSRM_WARMUP_TIMEOUT', 65) if startup else (
        seconds('CAREFLOW_OSRM_REMOTE_READ_TIMEOUT', 30) if remote else local_timeout)
    timeout = (seconds('CAREFLOW_OSRM_CONNECT_TIMEOUT', 5), read)
    attempts = 1 if startup or not remote else 2
    for attempt in range(attempts):
        try:
            response = requests.get(host + path, params=params, timeout=timeout)
            if response.status_code in (429, 502, 503, 504) and attempt + 1 < attempts:
                continue
            response.raise_for_status()
            data = response.json()
            if not isinstance(data, dict) or data.get('code') != 'Ok':
                code = data.get('code', '缺少code') if isinstance(data, dict) else '格式不符'
                raise RuntimeError('機車路網未回傳可用結果：' + str(code))
            return data
        except (requests.Timeout, requests.ConnectionError) as exc:
            if attempt + 1 < attempts:
                continue
            _ready.pop(host, None)
            detail = '等待回應逾時' if isinstance(exc, requests.Timeout) else '連線失敗'
            raise RuntimeError(f'OSRM機車服務{detail}（連線上限{timeout[0]:g}秒、回應上限{read:g}秒）。請檢查Render服務狀態與日誌；未使用汽車或直線替代。') from exc
        except requests.HTTPError as exc:
            _ready.pop(host, None)
            raise RuntimeError(f'OSRM機車服務HTTP {response.status_code}；請檢查Render服務與路網部署。') from exc
        except RuntimeError:
            _ready.pop(host, None)
            raise
        except ValueError as exc:
            _ready.pop(host, None)
            raise RuntimeError('OSRM未回傳JSON路網資料，可能仍在啟動或部署設定錯誤；請稍後檢查連線。') from exc


def ensure_ready(host=None, force=False):
    """Probe a public landmark, once per host/5 minutes, before a large matrix."""
    host = host or endpoint()
    if not is_remote(host) and not force:
        return {'服務': urlparse(host).netloc, '狀態': '本機路網，直接查詢', '檢查耗時(秒)': 0}
    with _lock:
        if not force and time.monotonic() - _ready.get(host, -1e10) < 300:
            return {'服務': urlparse(host).netloc, '狀態': '已通過近期連線檢查', '檢查耗時(秒)': 0}
        start = time.monotonic()
        # Public streets near Taipei 101. No worker/client locations or identifiers.
        data = request_json(host, '/route/v1/motorcycle/121.5654,25.0330;121.5649,25.0340',
                            {'overview': 'false'}, startup=True)
        duration = data.get('routes', [{}])[0].get('duration') if data.get('routes') else None
        if not isinstance(duration, (int, float)) or not math.isfinite(duration) or duration < 0:
            raise RuntimeError('OSRM連線可達，但未回傳有效機車路線時間。')
        _ready[host] = time.monotonic()
        return {'服務': urlparse(host).netloc, '狀態': '機車路網可連線（尚非準確度驗證）',
                '檢查耗時(秒)': round(time.monotonic() - start, 2)}
