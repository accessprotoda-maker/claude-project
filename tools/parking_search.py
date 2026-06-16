import requests


def geocode_address(address: str) -> tuple[float, float] | None:
    """住所・建物名称から緯度経度を取得（国土地理院API）"""
    url = "https://msearch.gsi.go.jp/address-search/AddressSearch"
    resp = requests.get(url, params={"q": address}, timeout=10)
    resp.raise_for_status()
    data = resp.json()
    if not data:
        return None
    coords = data[0]["geometry"]["coordinates"]
    return float(coords[1]), float(coords[0])  # (lat, lng)


def search_parking(lat: float, lng: float, radius_m: int = 500) -> list[dict]:
    """Overpass API で近隣駐車場を検索"""
    overpass_url = "https://overpass-api.de/api/interpreter"
    query = f"""
[out:json][timeout:25];
(
  node["amenity"="parking"](around:{radius_m},{lat},{lng});
  way["amenity"="parking"](around:{radius_m},{lat},{lng});
  relation["amenity"="parking"](around:{radius_m},{lat},{lng});
);
out center;
"""
    resp = requests.post(overpass_url, data={"data": query}, timeout=30)
    resp.raise_for_status()
    elements = resp.json().get("elements", [])

    results = []
    for el in elements:
        if el["type"] == "node":
            elat, elng = el.get("lat"), el.get("lon")
        else:
            center = el.get("center", {})
            elat, elng = center.get("lat"), center.get("lon")

        if elat is None or elng is None:
            continue

        tags = el.get("tags", {})
        fee = tags.get("fee", "")
        access = tags.get("access", "")
        parking_type = tags.get("parking", "")
        name = tags.get("name", tags.get("name:ja", ""))
        capacity = tags.get("capacity", "")
        operator = tags.get("operator", "")

        # 種別判定
        if fee == "no":
            kind = "無料駐車場"
        elif fee == "yes":
            kind = "コインパーキング"
        else:
            kind = "駐車場（詳細不明）"

        # 非公開・関係者専用は除外
        if access in ("private", "no"):
            continue

        dist = _haversine(lat, lng, elat, elng)

        results.append({
            "name": name or kind,
            "kind": kind,
            "fee": fee,
            "capacity": capacity,
            "operator": operator,
            "parking_type": parking_type,
            "lat": elat,
            "lng": elng,
            "distance_m": int(dist),
        })

    results.sort(key=lambda x: x["distance_m"])
    return results


def _haversine(lat1, lng1, lat2, lng2) -> float:
    import math
    R = 6371000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lng2 - lng1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlam / 2) ** 2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
