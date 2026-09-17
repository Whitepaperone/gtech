# -*- coding: utf-8 -*-
"""서버 API 호출과 공통 응답 검증을 담당하는 모듈."""

import json
import os
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from AppConstants import API_BASE_URL, ORDERS_API_TIMEOUT_SECONDS


API_KEY_ENV_NAME = "COA_API_KEY"


def _get_api_key() -> str:
    """Return the API key configured on the PC running this program."""
    api_key = str(os.getenv(API_KEY_ENV_NAME) or "").strip()
    if not api_key:
        raise RuntimeError(
            f"{API_KEY_ENV_NAME} 환경변수가 설정되지 않았습니다. "
            "Windows 환경변수에 API Key를 등록한 후 프로그램을 다시 실행해 주세요."
        )
    return api_key


def fetch_json(url: str, timeout: float = 30) -> dict:
    """GET 요청을 보내고 JSON 객체 응답을 반환한다."""
    url = str(url or "").strip()
    if not url:
        raise RuntimeError("API 주소가 설정되지 않았습니다.")
    if not url.lower().startswith(("http://", "https://")):
        url = f"http://{url}"

    request = Request(
        url,
        headers={
            "Accept": "application/json",
            "Authorization": f"Bearer {_get_api_key()}",
        },
    )
    try:
        with urlopen(request, timeout=timeout) as response:
            charset = response.headers.get_content_charset() or "utf-8"
            payload = json.loads(response.read().decode(charset))
    except HTTPError as exc:
        if exc.code == 401:
            raise RuntimeError(
                "API 인증에 실패했습니다. COA_API_KEY가 서버의 API_KEY와 같은지 확인해 주세요."
            ) from exc
        raise RuntimeError(f"API 호출 실패: HTTP {exc.code} ({url})") from exc
    except URLError as exc:
        reason = getattr(exc, "reason", exc)
        raise RuntimeError(f"API 서버에 연결할 수 없습니다: {reason} ({url})") from exc
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"API 응답이 올바른 JSON 형식이 아닙니다: {url}") from exc

    if not isinstance(payload, dict):
        raise RuntimeError("API 응답의 최상위 값이 JSON 객체가 아닙니다.")
    return payload


def post_json(url: str, body: dict, timeout: float = 30) -> dict:
    """Bearer 인증을 포함해 JSON POST 요청을 보내고 JSON 응답을 반환한다."""
    request = Request(
        url,
        data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
        headers={
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8",
            "Authorization": f"Bearer {_get_api_key()}",
        },
        method="POST",
    )
    try:
        with urlopen(request, timeout=timeout) as response:
            charset = response.headers.get_content_charset() or "utf-8"
            payload = json.loads(response.read().decode(charset))
    except HTTPError as exc:
        if exc.code == 401:
            raise RuntimeError(
                "API 인증에 실패했습니다. COA_API_KEY가 서버의 API_KEY와 같은지 확인해 주세요."
            ) from exc
        raise RuntimeError(f"API 호출 실패: HTTP {exc.code} ({url})") from exc
    except URLError as exc:
        reason = getattr(exc, "reason", exc)
        raise RuntimeError(f"API 서버에 연결할 수 없습니다: {reason} ({url})") from exc
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"API 응답이 올바른 JSON 형식이 아닙니다: {url}") from exc

    if not isinstance(payload, dict):
        raise RuntimeError("API 응답의 최상위 값이 JSON 객체가 아닙니다.")
    return payload


def fetch_orders(base_url: str = None, timeout: float = None) -> list[dict]:
    """설정된 /orders API를 호출하고 수주 레코드 배열을 반환한다."""
    server_address = str(base_url or API_BASE_URL).strip()
    if not server_address:
        raise RuntimeError(
            "API 서버 주소가 설정되지 않았습니다. "
            "app_config.json의 api_base_url에 아이피주소:포트를 입력해 주세요."
        )
    url = f"{server_address.rstrip('/')}/orders"

    payload = fetch_json(url, timeout or ORDERS_API_TIMEOUT_SECONDS)
    if payload.get("ok") is not True:
        raise RuntimeError("수주 API가 실패 응답을 반환했습니다.")

    records = payload.get("data")
    if not isinstance(records, list):
        raise RuntimeError("수주 API 응답의 data가 배열이 아닙니다.")
    if any(not isinstance(record, dict) for record in records):
        raise RuntimeError("수주 API의 data에 객체가 아닌 항목이 포함되어 있습니다.")
    return records


def fetch_shipment_changes(
    consumer: str, base_url: str = None, timeout: float = None
) -> dict:
    """설정된 /shipments/changes API의 검증된 전체 응답을 반환한다."""
    server_address = str(base_url or API_BASE_URL).strip()
    if not server_address:
        raise RuntimeError(
            "API 서버 주소가 설정되지 않았습니다. "
            "app_config.json의 api_base_url에 아이피주소:포트를 입력해 주세요."
        )
    query = urlencode({"consumer": consumer})
    url = f"{server_address.rstrip('/')}/shipments/changes?{query}"
    payload = fetch_json(url, timeout or ORDERS_API_TIMEOUT_SECONDS)
    if payload.get("ok") is not True:
        raise RuntimeError("출고 변경 API가 실패 응답을 반환했습니다.")

    records = payload.get("data")
    if not isinstance(records, list):
        raise RuntimeError("출고 변경 API 응답의 data가 배열이 아닙니다.")
    if any(not isinstance(record, dict) for record in records):
        raise RuntimeError("출고 변경 API의 data에 객체가 아닌 항목이 포함되어 있습니다.")
    return payload


def acknowledge_shipment_changes(
    consumer: str,
    change_ids: list[str],
    metadata: dict = None,
    base_url: str = None,
    timeout: float = None,
) -> int:
    """엑셀 저장에 성공한 출고 변경 건을 중앙 처리 이력에 기록한다."""
    if not change_ids:
        return 0
    server_address = str(base_url or API_BASE_URL).strip()
    if not server_address:
        raise RuntimeError("API 서버 주소가 설정되지 않았습니다.")
    url = f"{server_address.rstrip('/')}/shipments/acknowledgements"
    acknowledged_count = 0
    for start in range(0, len(change_ids), 100):
        payload = post_json(
            url,
            {
                "consumer": consumer,
                "changeIds": change_ids[start:start + 100],
                "metadata": metadata or {},
            },
            timeout or ORDERS_API_TIMEOUT_SECONDS,
        )
        if payload.get("ok") is not True:
            raise RuntimeError("출고 변경 처리 완료 기록에 실패했습니다.")
        acknowledged_count += int(payload.get("acknowledgedCount") or 0)
    return acknowledged_count
