import folium
import streamlit as st
from streamlit_folium import st_folium

from tools.parking_search import geocode_address, search_parking

st.set_page_config(page_title="駐車場検索", layout="wide")

st.title("近隣駐車場検索")
st.write("住所または建物名称を入力すると、周辺のコインパーキング・無料駐車場を地図で表示します。")

col1, col2 = st.columns([3, 1])
with col1:
    address = st.text_input("住所・建物名称", placeholder="例: 東京都渋谷区道玄坂1-1")
with col2:
    radius = st.selectbox("検索半径", [300, 500, 800, 1000, 1500], index=1, format_func=lambda x: f"{x}m")

if st.button("検索", type="primary", disabled=not address):
    with st.spinner("検索中..."):
        coords = geocode_address(address)
        if coords is None:
            st.error("住所が見つかりませんでした。別の表記でお試しください。")
            st.stop()

        lat, lng = coords
        parkings = search_parking(lat, lng, radius_m=radius)

    # 件数サマリ
    coin = [p for p in parkings if p["fee"] == "yes"]
    free = [p for p in parkings if p["fee"] == "no"]
    other = [p for p in parkings if p["fee"] not in ("yes", "no")]

    st.success(f"**{len(parkings)}件** 見つかりました（コインパーキング: {len(coin)}件 / 無料: {len(free)}件 / 詳細不明: {len(other)}件）")

    # 地図
    m = folium.Map(location=[lat, lng], zoom_start=16)

    # 検索基点マーカー
    folium.Marker(
        [lat, lng],
        popup=folium.Popup(address, max_width=200),
        tooltip="検索地点",
        icon=folium.Icon(color="red", icon="home"),
    ).add_to(m)

    # 検索半径円
    folium.Circle(
        [lat, lng],
        radius=radius,
        color="#3186cc",
        fill=True,
        fill_opacity=0.05,
    ).add_to(m)

    color_map = {"yes": "blue", "no": "green"}
    icon_map = {"yes": "yen-sign", "no": "parking"}

    for p in parkings:
        color = color_map.get(p["fee"], "gray")
        lines = [
            f"<b>{p['name']}</b>",
            f"種別: {p['kind']}",
            f"距離: {p['distance_m']}m",
        ]
        if p["capacity"]:
            lines.append(f"収容台数: {p['capacity']}台")
        if p["operator"]:
            lines.append(f"運営: {p['operator']}")

        popup_html = "<br>".join(lines)

        folium.Marker(
            [p["lat"], p["lng"]],
            popup=folium.Popup(popup_html, max_width=250),
            tooltip=f"{p['name']} ({p['distance_m']}m)",
            icon=folium.Icon(color=color, prefix="fa", icon="p"),
        ).add_to(m)

    st_folium(m, width="100%", height=500)

    # 一覧テーブル
    if parkings:
        st.subheader("一覧")
        rows = []
        for p in parkings:
            rows.append({
                "名称": p["name"],
                "種別": p["kind"],
                "距離": f"{p['distance_m']}m",
                "収容台数": p["capacity"] or "—",
                "運営": p["operator"] or "—",
            })
        st.dataframe(rows, use_container_width=True, hide_index=True)
