import streamlit as st
import pandas as pd
import requests
import json
import difflib
import datetime
import random

# ==========================================
# 0. 初期設定 & モバイル最適化CSS
# ==========================================
st.set_page_config(page_title="献立＆買出しアプリ", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    /* タブメニューを画面上部に固定 */
    div[data-testid="stTabs"] > div:nth-child(1) {
        position: sticky;
        top: 0;
        z-index: 999;
        background-color: white;
        padding-top: 10px;
        padding-bottom: 5px;
        border-bottom: 1px solid #e6e6e6;
    }
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 5rem !important;
    }
    .day-badge {
        color: white; 
        background-color: #E03C31; 
        padding: 3px 8px; 
        border-radius: 4px; 
        font-size: 0.85em; 
        font-weight: bold;
        margin-right: 8px;
    }
    </style>
""", unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def fetch_available_models(key):
    for version in ["v1", "v1beta"]:
        url = f"https://generativelanguage.googleapis.com/{version}/models?key={key}"
        res = requests.get(url)
        if res.status_code == 200:
            models = res.json().get("models", [])
            return [m["name"].replace("models/", "") for m in models if "generateContent" in m.get("supportedGenerationMethods", []) and "2.5" not in m["name"]]
    return []

# APIキー管理
with st.sidebar:
    st.header("🔑 システム設定")
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("✅ APIキー自動読み込み完了")
    else:
        api_key = st.text_input("Gemini API Key", type="password")
        st.warning("⚠️ APIキーが未設定です")
    selected_model = "gemini-3.1-flash-lite"

# セッション状態の初期化
if "inventory_df" not in st.session_state:
    st.session_state.inventory_df = pd.DataFrame([
        {"食材名": "豚肉", "カテゴリ": "精肉", "残量": 200.0, "単位": "g", "購入日": datetime.date.today() - datetime.timedelta(days=4)},
        {"食材名": "キャベツ", "カテゴリ": "青果", "残量": 1.0, "単位": "玉", "購入日": datetime.date.today()},
        {"食材名": "玉ねぎ", "カテゴリ": "青果", "残量": 3.0, "単位": "個", "購入日": datetime.date.today() - datetime.timedelta(days=6)},
        {"食材名": "卵", "カテゴリ": "日配品", "残量": 4.0, "単位": "個", "購入日": datetime.date.today() - datetime.timedelta(days=2)},
        {"食材名": "牛乳", "カテゴリ": "日配品", "残量": 0.5, "単位": "本", "購入日": datetime.date.today() - datetime.timedelta(days=3)}
    ])

if "shopping_list_df" not in st.session_state:
    st.session_state.shopping_list_df = pd.DataFrame(columns=["買出済", "食材名", "カテゴリ", "必要量", "単位", "確定献立のDAY", "確定献立のレシピ名"])
if "draft_plan" not in st.session_state:
    st.session_state.draft_plan = []
if "final_plan" not in st.session_state:
    st.session_state.final_plan = []
if "selected_order" not in st.session_state:
    st.session_state.selected_order = []

UNIT_OPTIONS = ["g", "個", "玉", "本", "束", "パック", "枚", "ml"]
CATEGORY_OPTIONS = ["青果", "精肉", "鮮魚", "日配品", "加工食品", "その他"]
STAPLE_ITEMS = [
    {"食材名": "牛乳", "カテゴリ": "日配品", "目標量": 2.0, "単位": "本"},
    {"食材名": "卵", "カテゴリ": "日配品", "目標量": 10.0, "単位": "個"},
    {"食材名": "玉ねぎ", "カテゴリ": "青果", "目標量": 5.0, "単位": "個"}
]

# ==========================================
# 1. ユーティリティ関数
# ==========================================
def generate_via_gemini(prompt, key, model_name, sys_prompt):
    if not key or not model_name:
        st.error("APIキーが設定されていません。")
        return None
    payload = {"contents": [{"parts": [{"text": sys_prompt + "\n\n" + prompt}]}]}
    headers = {'Content-Type': 'application/json'}
    for version in ["v1", "v1beta", "v1alpha"]:
        url = f"https://generativelanguage.googleapis.com/{version}/models/{model_name}:generateContent?key={key}"
        try:
            res = requests.post(url, headers=headers, json=payload)
            if res.status_code == 200:
                text = res.json()["candidates"][0]["content"]["parts"][0]["text"]
                return json.loads(text.replace("```json", "").replace("```", "").strip())
        except Exception:
            pass
    return None

def calculate_shopping_list(plan_list):
    required = {}
    for idx, r in enumerate(plan_list):
        day_str = f"Day {idx+1}"
        recipe_name = r.get("name", "")
        for ing, amt in r.get("ingredients", {}).items():
            if ing not in required:
                required[ing] = {"amt": 0, "unit": r.get("unit_map", {}).get(ing, ""), "days": set(), "recipes": set()}
            required[ing]["amt"] += float(amt)
            required[ing]["days"].add(day_str)
            required[ing]["recipes"].add(recipe_name)
            
    inv_total = st.session_state.inventory_df.groupby("食材名")["残量"].sum().to_dict()
    shop_data = []
    for ing, data in required.items():
        inv_amt = float(inv_total.get(ing, 0))
        shortage = data["amt"] - inv_amt
        if shortage > 0:
            shop_data.append({
                "買出済": False, "食材名": ing, "カテゴリ": "その他", "必要量": shortage, "単位": data["unit"], 
                "確定献立のDAY": ", ".join(sorted(data["days"])), "確定献立のレシピ名": ", ".join(data["recipes"])
            })
    st.session_state.shopping_list_df = pd.DataFrame(shop_data) if shop_data else pd.DataFrame(columns=["買出済", "食材名", "カテゴリ", "必要量", "単位", "確定献立のDAY", "確定献立のレシピ名"])

def consume_fifo(ingredients_dict):
    """FIFO（古い購入日から順に消費）で在庫を減算"""
    df = st.session_state.inventory_df
    for ing, required_amt in ingredients_dict.items():
        remaining_to_consume = float(required_amt)
        target_indices = df[df["食材名"] == ing].sort_values("購入日").index.tolist()
        
        for idx in target_indices:
            if remaining_to_consume <= 0: break
            current_amt = float(df.at[idx, "残量"])
            if current_amt <= remaining_to_consume:
                remaining_to_consume -= current_amt
                df.at[idx, "残量"] = 0.0
            else:
                df.at[idx, "残量"] = current_amt - remaining_to_consume
                remaining_to_consume = 0.0
                
    st.session_state.inventory_df = df[df["残量"] > 0].reset_index(drop=True)

def toggle_recipe(idx):
    if idx in st.session_state.selected_order:
        st.session_state.selected_order.remove(idx)
    else:
        st.session_state.selected_order.append(idx)

# ==========================================
# UI: タブ構成 (5タブに拡張)
# ==========================================
tab_create, tab_home, tab_shop, tab_consume, tab_inv = st.tabs(["⚙️ 献立作成", "📅 献立確定", "🛒 買出し", "🍳 個別消費", "📦 在庫管理"])

# ------------------------------------------
# Tab 1: 献立作成
# ------------------------------------------
with tab_create:
    st.header("献立の作成")
    mode = st.radio("作成モード", ["在庫消費優先", "リクエスト優先"], horizontal=True)
    target_days = st.number_input("何日分作成しますか？", min_value=1, max_value=7, value=3)
    
    if mode == "在庫消費優先":
        st.caption("現在庫のうち、購入日が古い食材をAIが自動で優先的に使って献立を考えます。")
        sorted_df = st.session_state.inventory_df.sort_values(by="購入日", ascending=True)
        st.dataframe(sorted_df[["購入日", "食材名", "残量", "単位"]].head(5), use_container_width=True, hide_index=True)
        
        req_prompt = f"以下の在庫データは「購入日が古い順」に並んでいます。上位の古い食材を優先的に使い切るように、{target_days}日分のお米に合う夕食レシピを考案してください。\n"
        req_prompt += f"【在庫データ（古い順）】\n{sorted_df.to_json(orient='records', force_ascii=False)}\n"
            
    else:
        st.caption(f"※指定日数の2倍（{target_days * 2}品）のレシピを提案します。")
        example_menus = ["彩り鮮やかな野菜の黒酢あん", "ホロホロ鶏肉のトマト煮込み", "ガツンとニンニク香る豚バラ炒め", "さっぱり柚子胡椒の和風パスタ", "鮭のふっくらホイル焼き"]
        requests_list = []
        for d in range(int(target_days)):
            ex = example_menus[d % len(example_menus)]
            req = st.text_input(f"Day {d+1} のリクエスト", placeholder=f"例：{ex}", key=f"req_day_{d}")
            requests_list.append(req)
            
        req_prompt = f"以下のリクエストに基づき、合計 {target_days * 2}品 のお米に合う夕食レシピを考案してください。\n【リクエスト】\n{requests_list}\n"

    if st.button("✨ レシピ案を生成", type="primary", use_container_width=True):
        with st.spinner("AIが詳細なレシピを考案中..."):
            sys_prompt = """
            あなたはプロの料理研究家です。条件に基づき、以下のJSON配列のフォーマットを【厳密に】守って出力してください。
            ・調味料は「炒め用」「下味」「合わせ調味料」など用途別に連想配列で分類してください。
            ・手順やポイントは、必ず「見出し(title)」と「詳細説明(desc)」のセットにしてください。
            
            [
              {
                "name": "豚こまとズッキーニ、舞茸のガリバタ醤油炒め",
                "intro": "豚こま肉の旨味、ズッキーニのジューシーさ...ご飯のおかずにもぴったりの一品です。",
                "ingredients": {"豚こま切れ肉": 200, "ズッキーニ": 1, "舞茸": 1},
                "unit_map": {"豚こま切れ肉": "g", "ズッキーニ": "本", "舞茸": "パック"},
                "seasonings": {
                  "炒め用・その他": ["にんにく（みじん切り）", "バター", "サラダ油"],
                  "豚肉の下味": ["酒", "塩こしょう", "片栗粉"],
                  "合わせ調味料": ["醤油", "みりん"]
                },
                "steps": [
                  {"title": "具材の下準備", "desc": "ズッキーニは縦半分に切り、幅1cmほどの半月切りにします。舞茸は石づきを取り..."},
                  {"title": "豚肉の下処理", "desc": "豚こま切れ肉はボウルに入れ、下味の酒、塩こしょうを揉み込みます..."},
                  {"title": "香りを出して豚肉を炒める", "desc": "フライパンにサラダ油とみじん切りにしたにんにくを入れて弱火にかけます..."}
                ],
                "tips": [
                  {"title": "お肉に片栗粉をまぶす", "desc": "豚肉がパサつかず柔らかく仕上がるだけでなく、タレがしっかり絡むようになります。"},
                  {"title": "ズッキーニの焼き加減", "desc": "ズッキーニは少し焼き色がつくくらいまでしっかり炒めると、中がトロッとジューシーに仕上がります。"}
                ]
              }
            ]
            """
            prompt = f"条件: 2人前。\n" + req_prompt
            res = generate_via_gemini(prompt, api_key, selected_model, sys_prompt)
            if res:
                st.session_state.draft_plan = res
                st.session_state.selected_order = []

    if st.session_state.draft_plan:
        st.divider()
        st.markdown(f"**💡 採用するレシピを {target_days} つ選んでください**")
        st.caption("チェックした順に DAY が割り当てられます。")
        
        for i, recipe in enumerate(st.session_state.draft_plan):
            is_selected = (i in st.session_state.selected_order)
            order_badge = f"<span class='day-badge'>DAY {st.session_state.selected_order.index(i) + 1}</span>" if is_selected else ""
                
            cols = st.columns([1, 8])
            with cols[0]:
                st.checkbox(" ", value=is_selected, on_change=toggle_recipe, args=(i,), key=f"sel_{i}")
            with cols[1]:
                st.markdown(f"{order_badge} **{recipe['name']}**", unsafe_allow_html=True)
                
        if len(st.session_state.selected_order) > 0:
            st.write("")
            if st.button("✅ チェックしたレシピで確定", type="primary", use_container_width=True):
                st.session_state.final_plan = [st.session_state.draft_plan[idx] for idx in st.session_state.selected_order]
                calculate_shopping_list(st.session_state.final_plan)
                st.session_state.draft_plan = []
                st.session_state.selected_order = []
                st.success("確定しました！「献立確定」タブへ移動してください。")

# ------------------------------------------
# Tab 2: 献立確定
# ------------------------------------------
with tab_home:
    st.header("📅 確定した献立")
    if not st.session_state.final_plan:
        st.info("「献立作成」タブからメニューを確定させてください。")
    else:
        inv_total = st.session_state.inventory_df.groupby("食材名")["残量"].sum().to_dict()
        
        for i, r in enumerate(st.session_state.final_plan):
            with st.expander(f"Day {i+1}: {r['name']}", expanded=(i==0)):
                st.markdown(f"*{r.get('intro', '')}*")
                
                col_left, col_right = st.columns(2)
                with col_left:
                    st.markdown("#### 🔪 材料（2人分）")
                    for ing, amt in r.get("ingredients", {}).items():
                        unit = r.get("unit_map", {}).get(ing, "")
                        inv_amt = float(inv_total.get(ing, 0))
                        if inv_amt < float(amt):
                            st.error(f"- **{ing}**: {amt} {unit} ⚠️不足 (在庫: {inv_amt}{unit})")
                        else:
                            st.write(f"- **{ing}**: {amt} {unit}")
                
                with col_right:
                    st.markdown("#### 🧂 調味料・基本食材")
                    st.caption("※分量管理なし。有無のチェック用")
                    seasonings = r.get("seasonings", {})
                    if isinstance(seasonings, dict):
                        for cat_name, items in seasonings.items():
                            st.markdown(f"**【{cat_name}】**")
                            for s in items:
                                st.checkbox(s, key=f"seasoning_{i}_{cat_name}_{s}")
                    elif isinstance(seasonings, list):
                        for s in seasonings:
                            st.checkbox(s, key=f"seasoning_{i}_{s}")
                        
                st.markdown("#### 🍳 作り方")
                for step_idx, step in enumerate(r.get("steps", [])): 
                    if isinstance(step, dict):
                        st.markdown(f"**{step_idx+1}. {step.get('title', '')}**")
                        st.write(f"{step.get('desc', '')}")
                    else:
                        st.write(f"- {step}")
                    
                if r.get("tips"):
                    st.markdown("#### 💡 美味しく作るためのポイント")
                    for tip in r.get("tips", []): 
                        if isinstance(tip, dict):
                            st.markdown(f"- **{tip.get('title', '')}**: {tip.get('desc', '')}")
                        else:
                            st.info(tip)
                
                st.write("")
                if st.button(f"👩‍🍳 Day {i+1} 調理完了 (在庫から減算)", key=f"consume_btn_{i}", type="secondary", use_container_width=True):
                    consume_fifo(r.get("ingredients", {}))
                    st.success("古い在庫から順に材料を差し引きました！")
                    st.rerun()

# ------------------------------------------
# Tab 3: 買出しリスト
# ------------------------------------------
with tab_shop:
    st.header("🛒 買出しリスト")
    
    if st.button("🔄 定番アイテムの不足分を追加", use_container_width=True):
        inv_total = st.session_state.inventory_df.groupby("食材名")["残量"].sum().to_dict()
        added = 0
        for item in STAPLE_ITEMS:
            current = float(inv_total.get(item["食材名"], 0))
            if item["目標量"] - current > 0:
                if st.session_state.shopping_list_df[st.session_state.shopping_list_df["食材名"] == item["食材名"]].empty:
                    new_row = pd.DataFrame([{"買出済": False, "食材名": item["食材名"], "カテゴリ": item["カテゴリ"], "必要量": item["目標量"] - current, "単位": item["単位"], "確定献立のDAY": "ストック", "確定献立のレシピ名": "定番補充"}])
                    st.session_state.shopping_list_df = pd.concat([st.session_state.shopping_list_df, new_row], ignore_index=True)
                    added += 1
        st.success(f"{added}件の定番アイテムを追加しました。" if added > 0 else "ストックは十分です。")

    if len(st.session_state.shopping_list_df) == 0:
        st.info("買い出しが必要な食材はありません。")
    else:
        edited_shop = st.data_editor(
            st.session_state.shopping_list_df,
            num_rows="dynamic",
            column_config={
                "買出済": st.column_config.CheckboxColumn("買出済", default=False),
                "カテゴリ": st.column_config.SelectboxColumn("カテゴリ", options=CATEGORY_OPTIONS),
                "確定献立のDAY": st.column_config.TextColumn("使用日", disabled=True),
                "確定献立のレシピ名": st.column_config.TextColumn("対象レシピ", disabled=True)
            },
            key="editor_shop",
            use_container_width=True
        )
        st.session_state.shopping_list_df = edited_shop

        if st.button("✅ チェック済みの品を在庫へ反映", type="primary", use_container_width=True):
            purchased = edited_shop[edited_shop["買出済"] == True]
            pending = edited_shop[edited_shop["買出済"] == False]
            today = datetime.date.today()
            
            if not purchased.empty:
                inv_df = st.session_state.inventory_df
                for _, row in purchased.iterrows():
                    ing_name = row["食材名"]
                    match_idx = inv_df[(inv_df["食材名"] == ing_name) & (inv_df["購入日"] == today)].index
                    if not match_idx.empty:
                        inv_df.at[match_idx[0], "残量"] += float(row["必要量"])
                    else:
                        new_row = pd.DataFrame([{"食材名": ing_name, "カテゴリ": row.get("カテゴリ", "その他"), "残量": float(row["必要量"]), "単位": row["単位"], "購入日": today}])
                        inv_df = pd.concat([inv_df, new_row], ignore_index=True)
                st.session_state.inventory_df = inv_df
                st.session_state.shopping_list_df = pending
                st.success("在庫に追加しました！")
                st.rerun()

    st.divider()
    
    st.subheader("🛍 音声スピード登録")
    voice_input = st.text_area("購入品を入力（マイク入力推奨）", placeholder="例：特売の豚肉500g")
    
    if st.button("🪄 解析して在庫に追加", use_container_width=True):
        with st.spinner("解析中..."):
            sys_prompt = """入力から食材名、数量、単位を抽出し、以下のJSON配列で出力してください。カテゴリは "青果", "精肉", "鮮魚", "日配品", "加工食品", "その他" から推測。[{"name": "食材名", "amount": 数量(数値), "unit": "単位", "category": "カテゴリ"}]"""
            parsed_items = generate_via_gemini(voice_input, api_key, selected_model, sys_prompt)
            if parsed_items:
                inv_df = st.session_state.inventory_df
                today = datetime.date.today()
                added_str = []
                for item in parsed_items:
                    ing_name = item.get("name")
                    match_idx = inv_df[(inv_df["食材名"] == ing_name) & (inv_df["購入日"] == today)].index
                    if not match_idx.empty:
                        inv_df.at[match_idx[0], "残量"] += float(item.get("amount", 1))
                    else:
                        new_row = pd.DataFrame([{"食材名": ing_name, "カテゴリ": item.get("category", "その他"), "残量": float(item.get("amount", 1)), "単位": item.get("unit", "個"), "購入日": today}])
                        inv_df = pd.concat([inv_df, new_row], ignore_index=True)
                    added_str.append(f"{ing_name}({item.get('amount')}{item.get('unit')})")
                st.session_state.inventory_df = inv_df
                st.success(f"追加完了: {', '.join(added_str)}")


# ------------------------------------------
# Tab 4: 個別消費 (専用画面)
# ------------------------------------------
with tab_consume:
    st.header("🍳 個別消費")
    st.caption("夕食の献立以外（朝食やお弁当など）で使った食材をここで記録します。購入日が古いものから自動で減算（FIFO）されます。")
    
    with st.container(border=True):
        unique_items = sorted(st.session_state.inventory_df["食材名"].unique().tolist())
        
        cols_c = st.columns([1, 1])
        with cols_c[0]:
            consume_target = st.selectbox("🍎 どの食材を使いましたか？", unique_items if unique_items else ["(在庫なし)"])
        with cols_c[1]:
            consume_amt = st.number_input("⚖️ 使った量", min_value=0.1, value=1.0, step=0.5)
            
        st.write("") 
        if st.button("一括で消費を記録する", type="primary", use_container_width=True) and unique_items:
            consume_fifo({consume_target: consume_amt})
            st.success(f"✅ {consume_target} を {consume_amt} 消費しました。")
            st.rerun()

# ------------------------------------------
# Tab 5: 食材管理
# ------------------------------------------
with tab_inv:
    st.header("📦 在庫管理表")
    st.caption("直接編集・削除が可能です。※同じ食材でも購入日が異なれば別行として管理されます。")
    
    old_names = st.session_state.inventory_df["食材名"].tolist()
    
    edited_inv_main = st.data_editor(
        st.session_state.inventory_df,
        num_rows="dynamic",
        column_config={
            "カテゴリ": st.column_config.SelectboxColumn("カテゴリ", options=CATEGORY_OPTIONS, required=True),
            "単位": st.column_config.SelectboxColumn("単位", options=UNIT_OPTIONS, required=True),
            "購入日": st.column_config.DateColumn("購入日", format="YYYY/MM/DD")
        },
        key="editor_inv",
        use_container_width=True
    )
    
    new_names = edited_inv_main["食材名"].dropna().tolist()
    if len(new_names) > len(old_names):
        added_item = list(set(new_names) - set(old_names))
        if added_item:
            matches = difflib.get_close_matches(added_item[0], old_names, n=1, cutoff=0.6)
            if matches:
                st.warning(f"⚠️ 表記ゆれアラート: 追加された「{added_item[0]}」は、登録済みの「{matches[0]}」と統合できる可能性があります。")
                
    st.session_state.inventory_df = edited_inv_main
