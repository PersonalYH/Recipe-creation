import streamlit as st
import pandas as pd
import requests
import json
import difflib
import datetime
import random

# ==========================================
# 0. 初期設定 & 状態管理
# ==========================================
st.set_page_config(page_title="1週間献立＆買出し自動化", layout="wide", initial_sidebar_state="expanded")

# サイドバー設定（Secretsからの自動読み込み対応）
with st.sidebar:
    st.header("🔑 システム設定")
    
    # Streamlitの安全な保管庫(Secrets)にキーがあれば自動設定、なければ手入力枠を表示
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
        st.success("✅ APIキー自動読み込み完了")
    else:
        api_key = st.text_input("Gemini API Key", type="password")
        st.warning("⚠️ APIキーが未設定です")
        
    # ご指定のモデルに固定
    selected_model = "gemini-3.1-flash-lite"
    
    if api_key:
        st.info(f"🤖 使用モデル:\n{selected_model}")

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
if "must_use_temp" not in st.session_state:
    st.session_state.must_use_temp = []

UNIT_OPTIONS = ["g", "個", "玉", "本", "束", "パック", "枚", "ml", "大さじ", "小さじ", "少々"]
CATEGORY_OPTIONS = ["青果", "精肉", "鮮魚", "日配品", "加工食品", "調味料", "その他"]

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
            
    # 合計在庫を計算
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
    """FIFO（古いものから順に消費）で在庫を減算"""
    df = st.session_state.inventory_df
    for ing, required_amt in ingredients_dict.items():
        remaining_to_consume = float(required_amt)
        # 購入日が古い順にインデックスを取得
        target_indices = df[df["食材名"] == ing].sort_values("購入日").index.tolist()
        
        for idx in target_indices:
            if remaining_to_consume <= 0:
                break
            current_amt = float(df.at[idx, "残量"])
            if current_amt <= remaining_to_consume:
                remaining_to_consume -= current_amt
                df.at[idx, "残量"] = 0.0
            else:
                df.at[idx, "残量"] = current_amt - remaining_to_consume
                remaining_to_consume = 0.0
                
    # 残量が0になった行を削除
    st.session_state.inventory_df = df[df["残量"] > 0].reset_index(drop=True)

def toggle_recipe(idx):
    if idx in st.session_state.selected_order:
        st.session_state.selected_order.remove(idx)
    else:
        st.session_state.selected_order.append(idx)

# ==========================================
# UI: タブ構成
# ==========================================
tab_create, tab_home, tab_shop, tab_inv = st.tabs(["⚙️ 献立作成", "📅 献立確定", "🛒 買出しリスト", "📦 食材管理"])

# ------------------------------------------
# Tab 1: 献立作成
# ------------------------------------------
with tab_create:
    st.title("⚙️ 献立プランの作成")
    mode = st.radio("作成モード", ["在庫消費優先", "リクエスト優先（食べたいもの）"], horizontal=True)
    target_days = st.number_input("何日分の献立を作成しますか？", min_value=1, max_value=7, value=5)
    
    if mode == "在庫消費優先":
        st.subheader("📦 カテゴリ別 現在庫（FIFO表示）")
        st.caption("カテゴリ別に購入日が古い順で表示されています。今回の献立でマストで使い切る食材にチェックを入れてください。")
        
        sorted_df = st.session_state.inventory_df.sort_values(by=["カテゴリ", "購入日"], ascending=[True, True])
        grouped = sorted_df.groupby("カテゴリ")
        
        selected_must_use = []
        for cat, group in grouped:
            st.markdown(f"#### 🏷️ {cat}")
            cols = st.columns(4)
            col_idx = 0
            for idx, row in group.iterrows():
                with cols[col_idx % 4]:
                    days_old = (datetime.date.today() - row["購入日"]).days
                    alert_icon = "⚠️" if days_old >= 5 else "🥬"
                    is_checked = st.checkbox(f"{alert_icon} {row['食材名']} (残:{row['残量']}{row['単位']})", key=f"must_temp_{idx}")
                    if is_checked:
                        selected_must_use.append({"食材名": row['食材名'], "残量": row['残量'], "単位": row['単位']})
                col_idx += 1
        
        if selected_must_use:
            req_prompt = f"以下の食材は指定された残量を*必ず全て使い切る*ように、{target_days}日分のお米に合う夕食レシピを考案してください。\n"
            req_prompt += f"【マスト消費データ】\n{json.dumps(selected_must_use, ensure_ascii=False)}\n"
        else:
            req_prompt = f"一般的な家庭にある食材を想定して、{target_days}日分のお米に合う夕食レシピを考案してください。\n"
            
    else:
        st.subheader("📝 食べたいメニューのリクエスト")
        st.caption("※指定日数の2倍のレシピを提案します。")
        example_menus = ["彩り鮮やかな野菜の黒酢あん", "ホロホロ鶏肉のトマト煮込み", "ガツンとニンニク香る豚バラ炒め", "さっぱり柚子胡椒の和風パスタ", "絶品本格カオマンガイ", "鮭のふっくらホイル焼き"]
        
        requests_list = []
        for d in range(int(target_days)):
            ex = example_menus[d % len(example_menus)]
            req = st.text_input(f"Day {d+1} のリクエスト", placeholder=f"例：{ex}", key=f"req_day_{d}")
            requests_list.append(req)
            
        req_prompt = f"以下のリクエストに基づき、合計 {target_days * 2}品 のお米に合う夕食レシピを考案してください。\n【リクエスト】\n{requests_list}\n"

    if st.button("✨ レシピ案を生成", type="primary"):
        with st.spinner("AIがレシピを考案中..."):
            sys_prompt = """
            あなたはプロの料理研究家です。指定された条件に基づいてレシピを考案し、以下のJSON配列のみを出力してください。
            ・食材（肉・野菜・メインの魚など）の単位は厳密に統一してください。
            ・米、塩こしょう、醤油などの「調味料・基本食材」は、分量不要で「seasonings」の配列に入れてください。
            ・手順(steps)は、下ごしらえから火加減まで「具体的に詳細に」記述してください（最低5ステップ以上）。
            [
              {
                "name": "料理名", "intro": "紹介文", 
                "ingredients": {"豚肉": 200, "キャベツ": 1}, 
                "unit_map": {"豚肉": "g", "キャベツ": "玉"}, 
                "seasonings": ["醤油", "みりん", "酒", "塩こしょう", "サラダ油"],
                "steps": ["1. 豚肉は〇〇cm幅に切り、塩こしょうで下味をつける。", "2. フライパンにサラダ油を中火で熱し、..."], 
                "tips": ["ポイント1"]
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
        st.subheader(f"💡 提案されたレシピ ({len(st.session_state.draft_plan)}件)")
        st.write(f"採用するレシピを {target_days} つチェックしてください。（チェックした瞬間にDAYが明示されます）")
        
        for i, recipe in enumerate(st.session_state.draft_plan):
            is_selected = (i in st.session_state.selected_order)
            
            # 選択順位のバッジ表示（リアルタイム明示）
            order_badge = ""
            if is_selected:
                day_num = st.session_state.selected_order.index(i) + 1
                order_badge = f"<span style='color: white; background-color: #E03C31; padding: 2px 8px; border-radius: 4px; margin-right: 8px;'>DAY {day_num}</span>"
                
            cols = st.columns([1, 10])
            with cols[0]:
                st.checkbox("選択", value=is_selected, on_change=toggle_recipe, args=(i,), key=f"sel_{i}")
            with cols[1]:
                st.markdown(f"{order_badge} **{recipe['name']}**", unsafe_allow_html=True)
                st.caption(recipe.get("intro", ""))
                
        if len(st.session_state.selected_order) > 0:
            if st.button("✅ チェックしたレシピで献立を確定させる", use_container_width=True):
                st.session_state.final_plan = [st.session_state.draft_plan[idx] for idx in st.session_state.selected_order]
                calculate_shopping_list(st.session_state.final_plan)
                st.session_state.draft_plan = []
                st.session_state.selected_order = []
                st.success("確定しました！「献立確定」タブを確認してください。")

# ------------------------------------------
# Tab 2: 献立確定
# ------------------------------------------
with tab_home:
    st.title("📅 確定した献立")
    if not st.session_state.final_plan:
        st.info("「献立作成」タブからメニューを確定させてください。")
    else:
        inv_total = st.session_state.inventory_df.groupby("食材名")["残量"].sum().to_dict()
        
        for i, r in enumerate(st.session_state.final_plan):
            with st.expander(f"Day {i+1}: {r['name']}", expanded=True):
                st.markdown(f"*{r.get('intro', '')}*")
                
                col_left, col_right = st.columns(2)
                with col_left:
                    st.markdown("#### 🔪 必要な材料（2人分）")
                    for ing, amt in r.get("ingredients", {}).items():
                        unit = r.get("unit_map", {}).get(ing, "")
                        inv_amt = float(inv_total.get(ing, 0))
                        if inv_amt < float(amt):
                            st.error(f"- **{ing}**: {amt} {unit} ⚠️在庫不足 (現在庫: {inv_amt}{unit})")
                        else:
                            st.write(f"- **{ing}**: {amt} {unit}")
                
                with col_right:
                    st.markdown("#### 🧂 調味料・基本食材")
                    st.caption("※分量管理なし。有無のチェックのみ。")
                    for s in r.get("seasonings", []):
                        st.checkbox(s, key=f"seasoning_{i}_{s}")
                        
                st.markdown("#### 🍳 作り方 (詳細手順)")
                for step in r.get("steps", []): st.write(f"- {step}")
                    
                if r.get("tips"):
                    st.markdown("#### 💡 美味しく作るためのポイント")
                    for tip in r.get("tips", []): st.info(tip)
                
                # 調理完了（FIFO一括消費）
                if st.button(f"👩‍🍳 Day {i+1} の調理を完了する（在庫から古い順に差し引く）", key=f"consume_btn_{i}"):
                    consume_fifo(r.get("ingredients", {}))
                    st.success(f"{r['name']} の材料を古い在庫から差し引きました！")
                    st.rerun()

# ------------------------------------------
# Tab 3: 買出しリスト
# ------------------------------------------
with tab_shop:
    st.title("🛒 買出しリスト")
    
    st.subheader("🥛 定番アイテムの自動補充")
    if st.button("🔄 定番アイテムの不足分をリストに反映"):
        inv_total = st.session_state.inventory_df.groupby("食材名")["残量"].sum().to_dict()
        added_count = 0
        for item in STAPLE_ITEMS:
            current = float(inv_total.get(item["食材名"], 0))
            shortage = item["目標量"] - current
            if shortage > 0:
                existing = st.session_state.shopping_list_df[st.session_state.shopping_list_df["食材名"] == item["食材名"]]
                if existing.empty:
                    new_row = pd.DataFrame([{
                        "買出済": False, "食材名": item["食材名"], "カテゴリ": item["カテゴリ"],
                        "必要量": shortage, "単位": item["単位"], "確定献立のDAY": "ストック", "確定献立のレシピ名": "定番補充"
                    }])
                    st.session_state.shopping_list_df = pd.concat([st.session_state.shopping_list_df, new_row], ignore_index=True)
                    added_count += 1
        if added_count > 0:
            st.success(f"{added_count}件の定番アイテムをリストに追加しました。")
        else:
            st.info("すべての定番アイテムは十分なストックがあります。")

    st.divider()

    if len(st.session_state.shopping_list_df) == 0:
        st.info("現在、買い出しが必要な食材はありません。")
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

        if st.button("✅ 完了（チェック済みのリスト品を在庫に反映）", type="primary"):
            purchased = edited_shop[edited_shop["買出済"] == True]
            pending = edited_shop[edited_shop["買出済"] == False]
            today = datetime.date.today()
            
            if not purchased.empty:
                inv_df = st.session_state.inventory_df
                for _, row in purchased.iterrows():
                    ing_name = row["食材名"]
                    # 同日購入の同一食材があれば合算、なければ新規行
                    match_idx = inv_df[(inv_df["食材名"] == ing_name) & (inv_df["購入日"] == today)].index
                    if not match_idx.empty:
                        idx = match_idx[0]
                        inv_df.at[idx, "残量"] += float(row["必要量"])
                    else:
                        new_row = pd.DataFrame([{
                            "食材名": ing_name, "カテゴリ": row.get("カテゴリ", "その他"), "残量": float(row["必要量"]), 
                            "単位": row["単位"], "購入日": today
                        }])
                        inv_df = pd.concat([inv_df, new_row], ignore_index=True)
                st.session_state.inventory_df = inv_df
                st.session_state.shopping_list_df = pending
                st.success(f"{len(purchased)} 件の食材を在庫に追加しました！")
                st.rerun()

    st.divider()
    
    st.subheader("🛍 リスト外購入のスピード登録")
    voice_input = st.text_area("購入したものを入力（または音声入力）", placeholder="例：アイス3個、特売の豚肉500g")
    
    if st.button("🪄 音声/テキストを解析して在庫に直接追加"):
        with st.spinner("AIが購入品を解析中..."):
            sys_prompt = """
            ユーザーの入力から食材名、数量、単位を抽出し、以下のJSON配列で出力してください。
            カテゴリは "青果", "精肉", "鮮魚", "日配品", "加工食品", "調味料", "その他" から推測。
            [{"name": "食材名", "amount": 数量(数値), "unit": "単位", "category": "カテゴリ"}]
            """
            parsed_items = generate_via_gemini(voice_input, api_key, selected_model, sys_prompt)
            if parsed_items:
                inv_df = st.session_state.inventory_df
                today = datetime.date.today()
                added_str = []
                for item in parsed_items:
                    ing_name = item.get("name")
                    match_idx = inv_df[(inv_df["食材名"] == ing_name) & (inv_df["購入日"] == today)].index
                    if not match_idx.empty:
                        idx = match_idx[0]
                        inv_df.at[idx, "残量"] += float(item.get("amount", 1))
                    else:
                        new_row = pd.DataFrame([{
                            "食材名": ing_name, "カテゴリ": item.get("category", "その他"), "残量": float(item.get("amount", 1)), 
                            "単位": item.get("unit", "個"), "購入日": today
                        }])
                        inv_df = pd.concat([inv_df, new_row], ignore_index=True)
                    added_str.append(f"{ing_name}({item.get('amount')}{item.get('unit')})")
                
                st.session_state.inventory_df = inv_df
                st.success(f"在庫に追加完了: {', '.join(added_str)}")

# ------------------------------------------
# Tab 4: 食材管理
# ------------------------------------------
with tab_inv:
    st.title("📦 食材管理表")
    
    st.subheader("🗑 サクッと個別消費 (FIFO自動適用)")
    st.caption("調理以外の用途で使った食材を素早く減算します。古い購入日のものから自動で消化されます。")
    cols_c = st.columns([3, 2, 2])
    
    unique_items = sorted(st.session_state.inventory_df["食材名"].unique().tolist())
    with cols_c[0]:
        consume_target = st.selectbox("消費した食材", unique_items if unique_items else ["(在庫なし)"])
    with cols_c[1]:
        consume_amt = st.number_input("消費量", min_value=0.1, value=1.0, step=0.5)
    with cols_c[2]:
        st.write("") 
        st.write("")
        if st.button("減算する", type="secondary") and unique_items:
            consume_fifo({consume_target: consume_amt})
            st.success(f"{consume_target} を {consume_amt} 消費しました。")
            st.rerun()

    st.divider()

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
                st.warning(f"⚠️ 表記ゆれアラート: 追加された「{added_item[0]}」は、既に登録されている「{matches[0]}」と統合できる可能性があります。")
                
    st.session_state.inventory_df = edited_inv_main
