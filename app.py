import streamlit as st
import pandas as pd
import requests
import json
import difflib
import datetime
import io
import msal
import base64

# ==========================================
# 0. 初期設定 & モバイル最適化CSS
# ==========================================
st.set_page_config(page_title="次世代・献立アプリ", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    
""", unsafe_allow_html=True)

# 定数
UNIT_OPTIONS = ["g", "個", "玉", "本", "束", "パック", "枚", "ml", "袋", "缶", "瓶"]
CATEGORY_OPTIONS = ["青果", "精肉", "鮮魚", "日配品", "加工食品", "調味料", "その他"]
SHEETS = ["Inventory", "ShoppingList", "Staples", "Seasonings", "PremadeSauces", "TransactionLog", "Ratings", "MealPlan"]

# ==========================================
# 1. AI連携 (テキスト・JSON・画像 対応型)
# ==========================================
def generate_via_gemini(prompt, sys_prompt="", response_type="json", image_b64=None):
    key = st.secrets.get("GEMINI_API_KEY")
    if not key: return None
    
    parts = [{"text": sys_prompt + "\n\n" + prompt}]
    if image_b64:
        parts.append({"inline_data": {"mime_type": "image/jpeg", "data": image_b64}})
        
    payload = {
        "contents": [{"parts": parts}],
        "generationConfig": {"temperature": 0.7}
    }
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key={key}"
    
    try:
        res = requests.post(url, headers={'Content-Type': 'application/json'}, json=payload, timeout=60)
        if res.status_code == 200:
            text = res.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
            if response_type == "json":
                text = text.replace("```json", "").replace("```", "").strip()
                start_idx = text.find('[')
                end_idx = text.rfind(']') + 1
                if start_idx != -1 and end_idx != -1:
                    return json.loads(text[start_idx:end_idx])
                return json.loads(text)
            else:
                return text # テキストモードの場合はそのまま返す
        else:
            st.error(f"APIエラー: {res.status_code}")
    except Exception as e:
        st.error(f"AI通信エラー。詳細: {e}")
    return None

# ==========================================
# 2. 認証 & OneDrive同期
# ==========================================
def get_ms_access_token():
    client_id = st.secrets["MS_CLIENT_ID"]
    client_secret = st.secrets["MS_CLIENT_SECRET"]
    refresh_token = st.secrets["MS_REFRESH_TOKEN"]
    app = msal.ConfidentialClientApplication(client_id, authority="https://login.microsoftonline.com/consumers", client_credential=client_secret)
    return app.acquire_token_by_refresh_token(refresh_token, scopes=["Files.ReadWrite"]).get("access_token")

def save_to_excel():
    token = get_ms_access_token()
    if not token: return False
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet in SHEETS:
            df = st.session_state.get(f"df_{sheet}", pd.DataFrame())
            df.to_excel(writer, sheet_name=sheet, index=False)
    output.seek(0)
    url = "https://graph.microsoft.com/v1.0/me/drive/root:/MealAppDB.xlsx:/content"
    requests.put(url, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}, data=output.read())

def load_from_excel():
    token = get_ms_access_token()
    if not token: return False
    url = "https://graph.microsoft.com/v1.0/me/drive/root:/MealAppDB.xlsx:/content"
    res = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    if res.status_code == 200:
        excel_data = io.BytesIO(res.content)
        for sheet in SHEETS:
            try:
                df = pd.read_excel(excel_data, sheet_name=sheet)
                if '購入日' in df.columns: df['購入日'] = pd.to_datetime(df['購入日']).dt.date
                if '日時' in df.columns: df['日時'] = pd.to_datetime(df['日時'])
                st.session_state[f"df_{sheet}"] = df
            except: pass
        return True
    return False

# ==========================================
# 3. データ処理ロジック (表記ゆれ対応)
# ==========================================
# ★ 追加: 表記ゆれや「〜の素」を空気を読んで同一視する関数
def find_closest_item(target_name, item_list, cutoff=0.6):
    # 完全一致
    if target_name in item_list: return target_name
    # 類似検索
    matches = difflib.get_close_matches(target_name, item_list, n=1, cutoff=cutoff)
    if matches: return matches[0]
    # 「〇〇の素」等の部分一致チェック
    for item in item_list:
        if target_name in item or item in target_name: return item
    return None

def log_transaction(item, category, io_type, amount):
    new_log = pd.DataFrame([{"日時": datetime.datetime.now(), "食材名": item, "カテゴリ": category, "入出庫": io_type, "数量": amount}])
    st.session_state.df_TransactionLog = pd.concat([st.session_state.df_TransactionLog, new_log], ignore_index=True)

def consume_fifo(ingredients_dict):
    df = st.session_state.df_Inventory
    inv_items = df["食材名"].unique().tolist()
    
    for req_ing, req_amt in ingredients_dict.items():
        # 表記ゆれを吸収して消費対象の食材を特定
        actual_ing = find_closest_item(req_ing, inv_items)
        if not actual_ing: continue # 在庫に無い場合はスキップ
        
        remaining = float(req_amt)
        target_indices = df[df["食材名"] == actual_ing].sort_values("購入日").index.tolist()
        cat = df[df["食材名"] == actual_ing]["カテゴリ"].iloc[0]
        log_transaction(actual_ing, cat, "消費", req_amt)
        
        for idx in target_indices:
            if remaining <= 0: break
            current = float(df.at[idx, "残量"])
            if current <= remaining:
                remaining -= current
                df.at[idx, "残量"] = 0.0
            else:
                df.at[idx, "残量"] = current - remaining
                remaining = 0.0
    st.session_state.df_Inventory = df[df["残量"] > 0].reset_index(drop=True)

# Siri用バックドア
if st.query_params.get("api") == "siri":
    query = st.query_params.get("q", "現在の在庫状況を教えて")
    load_from_excel()
    db_context = {s: st.session_state.get(f"df_{s}", pd.DataFrame()).to_dict(orient="records") for s in SHEETS if s != "MealPlan"}
    answer = generate_via_gemini(f"【データ】\n{json.dumps(db_context, ensure_ascii=False)}\n\n【質問】\n{query}", "あなたは家事サポートAIです。短く音声用の日本語で回答。", "text")
    st.json({"answer": answer if answer else "エラーが発生しました"})
    st.stop()

# 初期化
with st.sidebar:
    st.header("🔑 システム設定")
    if st.button("🔄 クラウド同期・再読込", use_container_width=True):
        st.session_state.data_loaded = False
        st.rerun()

if "data_loaded" not in st.session_state or not st.session_state.data_loaded:
    with st.spinner("☁️ クラウド同期中..."):
        load_from_excel()
        defaults = {
            "df_Inventory": pd.DataFrame([{"食材名": "豚肉", "カテゴリ": "精肉", "残量": 200.0, "単位": "g", "購入日": datetime.date.today()}]),
            "df_ShoppingList": pd.DataFrame(columns=["買出済", "食材名", "カテゴリ", "必要量", "単位", "目的"]),
            "df_Staples": pd.DataFrame([{"食材名": "牛乳", "カテゴリ": "日配品", "目標量": 2.0, "単位": "本"}]),
            "df_Seasonings": pd.DataFrame([{"調味料名": "醤油", "在庫あり": True}, {"調味料名": "みりん", "在庫あり": True}]),
            "df_PremadeSauces": pd.DataFrame([{"商品名": "麻婆豆腐の素", "対応メニュー": "麻婆豆腐"}]),
            "df_TransactionLog": pd.DataFrame(columns=["日時", "食材名", "カテゴリ", "入出庫", "数量"]),
            "df_Ratings": pd.DataFrame(columns=["日時", "レシピ名", "評価"]),
            "df_MealPlan": pd.DataFrame(columns=["PlanJSON"])
        }
        needs_save = False
        for key, default_df in defaults.items():
            if key not in st.session_state:
                st.session_state[key] = default_df; needs_save = True
        if needs_save: save_to_excel()
        
        st.session_state.final_plan = []
        df_plan = st.session_state.get("df_MealPlan")
        if df_plan is not None and not df_plan.empty:
            try: st.session_state.final_plan = json.loads(df_plan.iloc[0]["PlanJSON"])
            except: pass
        st.session_state.data_loaded = True

# ==========================================
# UI: 6つのタブ
# ==========================================
if "draft_plan" not in st.session_state: st.session_state.draft_plan = []
if "selected_order" not in st.session_state: st.session_state.selected_order = []

tab_create, tab_home, tab_shop, tab_consume, tab_manage, tab_dash = st.tabs(["⚙️ 作成", "📅 確定", "🛒 買出", "🍳 消費", "📦 管理", "📊 分析"])

# ------------------------------------------
# Tab 1: 献立作成 (気分スライダー＆蛇腹搭載)
# ------------------------------------------
with tab_create:
    st.header("献立の作成")
    mode = st.radio("作成モード", ["在庫消費優先", "リクエスト優先"], horizontal=True)
    
    # ★追加: 気分スライダー（AIのプロンプトを動的に変化）
    mood = st.selectbox("今日の気分・テーマは？", ["おまかせ（バランス良く）", "⏱ 疲れた（10分以内・包丁最小限）", "🔥 大量消費（余り物一掃）", "🍷 少し冒険（洋風・エスニック）"])
    target_days = st.number_input("何日分作成しますか？", min_value=1, max_value=7, value=1)
    
    req_prompt = f"【テーマ】{mood}\n"
    if mode == "在庫消費優先":
        sorted_df = st.session_state.df_Inventory.sort_values(by="購入日", ascending=True)
        st.dataframe(sorted_df[["購入日", "食材名", "残量", "単位"]].head(5), use_container_width=True, hide_index=True)
        req_prompt += f"【在庫】{sorted_df.to_json(orient='records', force_ascii=False)}\n"
        num_proposals = 5 
    else:
        reqs = [st.text_input(f"Day {d+1} のリクエスト", key=f"r_{d}") for d in range(int(target_days))]
        req_prompt += f"【リクエスト】{reqs}\n"
        num_proposals = max(2, target_days * 2)

    if st.button("✨ レシピ案を生成", type="primary", use_container_width=True):
        with st.spinner(f"AIが {num_proposals} 品のレシピを考案中...（テーマ: {mood}）"):
            hi_rates = st.session_state.df_Ratings[st.session_state.df_Ratings["評価"] >= 4]["レシピ名"].tolist()
            premades = st.session_state.df_PremadeSauces.to_json(orient='records', force_ascii=False)
            stocked_seasonings = st.session_state.df_Seasonings[st.session_state.df_Seasonings["在庫あり"] == True]["調味料名"].tolist()
            
            sys_prompt = f"""
            あなたはプロの料理研究家です。以下の条件に従い、必ず【{num_proposals}品】のレシピを作成しJSON配列で出力してください。
            【厳守事項】サーバー負荷軽減のため「steps（手順）」は簡潔な数文字の概要のみにしてください。(詳細な手順は後で生成します)
            
            1. リクエストとテーマに基づきメニューを考案。可能な限り【家にある調味料】を活用すること。
            2. 以下の高評価データから好みを推測して反映。【高評価】: {hi_rates}
            3. 「便利調味料リスト」に合致する場合は調合せずそれを使う手順を出力。【便利調味料】: {premades}
            4. 考案の際、以下の「現在家にある調味料」を最大限考慮してください。【家にある調味料】: {stocked_seasonings}
            
            [出力フォーマット(必ず配列)]
            [{{"name": "料理名", "intro": "紹介", "ingredients": {{"豚肉": 200}}, "unit_map": {{"豚肉": "g"}}, "seasonings": {{"調味料": ["醤油"]}}, "steps": [{{"title": "下準備", "desc": "簡潔に"}}], "tips": ["ポイント"]}}]
            """
            res = generate_via_gemini(req_prompt, sys_prompt, "json")
            if res:
                st.session_state.draft_plan = res
                st.session_state.selected_order = []
            else:
                st.error("⚠️️ エラーが発生しました。時間を置いて再度お試しください。")

    if st.session_state.draft_plan:
        st.divider()
        st.markdown(f"**💡 採用するレシピを {target_days} つ選択**")
        for i, recipe in enumerate(st.session_state.draft_plan):
            is_selected = (i in st.session_state.selected_order)
            badge = f"DAY {st.session_state.selected_order.index(i) + 1}" if is_selected else ""
            cols = st.columns([1, 8])
            with cols[0]:
                if st.checkbox(" ", value=is_selected, key=f"sel_{i}"):
                    if len(st.session_state.selected_order) < target_days and i not in st.session_state.selected_order:
                        st.session_state.selected_order.append(i)
                else:
                    if i in st.session_state.selected_order: st.session_state.selected_order.remove(i)
            with cols[1]: 
                # ★追加: 蛇腹でプレビュー表示
                with st.expander(f"{recipe['name']}"):
                    st.caption(recipe.get('intro', ''))
                    ings = [f"{k} {v}{recipe.get('unit_map', {}).get(k, '')}" for k, v in recipe.get("ingredients", {}).items()]
                    st.write(f"**材料:** {', '.join(ings)}")
                
        if len(st.session_state.selected_order) > 0 and st.button("✅ 確定する", type="primary", use_container_width=True):
            st.session_state.final_plan = [st.session_state.draft_plan[idx] for idx in st.session_state.selected_order]
            st.session_state.df_MealPlan = pd.DataFrame([{"PlanJSON": json.dumps(st.session_state.final_plan, ensure_ascii=False)}])
            
            # 不足分を買い物リストに追加
            req_dict = {}
            for r in st.session_state.final_plan:
                for ing, amt in r.get("ingredients", {}).items():
                    unit = r.get("unit_map", {}).get(ing, "個")
                    if ing not in req_dict: req_dict[ing] = {"amt": 0.0, "unit": unit, "reason": "献立"}
                    req_dict[ing]["amt"] += float(amt)
            inv_total = st.session_state.df_Inventory.groupby("食材名")["残量"].sum().to_dict()
            shop_df = st.session_state.df_ShoppingList
            new_items = []
            
            # ★表記ゆれ対応の買い物リスト計算
            inv_items_list = list(inv_total.keys())
            premade_list = st.session_state.df_PremadeSauces["商品名"].tolist()
            
            for ing, data in req_dict.items():
                if find_closest_item(ing, premade_list): continue # レトルトで代用可能なら買わない
                actual_ing = find_closest_item(ing, inv_items_list)
                current_stock = float(inv_total.get(actual_ing, 0.0)) if actual_ing else 0.0
                shortage = data["amt"] - current_stock
                if shortage > 0 and shop_df[shop_df["食材名"] == ing].empty:
                    new_items.append({"買出済": False, "食材名": ing, "カテゴリ": "青果", "必要量": shortage, "単位": data["unit"], "目的": data["reason"]})
            if new_items: st.session_state.df_ShoppingList = pd.concat([shop_df, pd.DataFrame(new_items)], ignore_index=True)
            
            save_to_excel()
            st.session_state.draft_plan = []; st.session_state.selected_order = []
            st.success("確定しました！「📅 確定」タブへ移動してください。")
            st.rerun()

# ------------------------------------------
# Tab 2: 献立確定 (スマホ自由編集 & プロ料理人モード)
# ------------------------------------------
with tab_home:
    st.header("📅 確定した献立")
    if not st.session_state.final_plan:
        st.info("献立がありません。「⚙️ 作成」タブで作成してください。")
    else:
        inv_df = st.session_state.df_Inventory
        inv_items = inv_df["食材名"].tolist()
        premade_items = st.session_state.df_PremadeSauces["商品名"].tolist()
        stocked_seasonings = st.session_state.df_Seasonings[st.session_state.df_Seasonings["在庫あり"] == True]["調味料名"].tolist()
        
        for i, r in enumerate(st.session_state.final_plan):
            with st.expander(f"Day {i+1}: {r['name']}", expanded=(i==0)):
                st.markdown(f"*{r.get('intro', '')}*")
                
                # ★追加: スマホ対応の食材「自由編集」UI
                st.markdown("#### 🔪 材料（自由に量や単位を調整可能）")
                edited_ingredients = {}
                for ing, amt in r.get("ingredients", {}).items():
                    unit = r.get("unit_map", {}).get(ing, "")
                    
                    # 🍎 表記ゆれ・レトルト判定ロジック
                    status_text = ""
                    if find_closest_item(ing, premade_items):
                        status_text = "✨ レトルト代用可"
                    else:
                        actual_ing = find_closest_item(ing, inv_items)
                        current_stock = inv_df[inv_df["食材名"] == actual_ing]["残量"].sum() if actual_ing else 0.0
                        if current_stock >= float(amt): status_text = f"✔️ 在庫あり ({actual_ing})"
                        else: status_text = f"⚠️ 不足"
                    
                    # スマホで押しやすい横並びUI
                    c1, c2, c3 = st.columns([5, 3, 3])
                    with c1: 
                        st.markdown(f"**{ing}**")
                        st.caption(status_text)
                    with c2: 
                        new_amt = st.number_input("量", value=float(amt), step=0.5, key=f"amt_{i}_{ing}", label_visibility="collapsed")
                    with c3: 
                        new_unit = st.selectbox("単位", UNIT_OPTIONS, index=UNIT_OPTIONS.index(unit) if unit in UNIT_OPTIONS else 0, key=f"unt_{i}_{ing}", label_visibility="collapsed")
                    edited_ingredients[ing] = new_amt
                
                st.markdown("#### 🧂 調味料確認")
                for cat, items in r.get("seasonings", {}).items():
                    for s in items:
                        if find_closest_item(s, stocked_seasonings): st.write(f"✔️ {s}")
                        else: st.error(f"❌ {s} (在庫未登録)")
                
                # ★追加: 詳細レシピ遅延生成モード
                st.markdown("#### 🍳 作り方")
                if "detailed_steps" in r:
                    st.write(r["detailed_steps"]) # 取得済みなら詳細を表示
                else:
                    for step in r.get("steps", []): st.write(f"**{step.get('title', '')}**: {step.get('desc', '')}")
                    if st.button("👨‍🍳 プロの料理人モードで【詳細な手順】を生成", key=f"pro_{i}"):
                        with st.spinner("切り方や火加減など、完璧な手順を執筆中..."):
                            sys_prompt = "プロの料理人として、以下の料理の完璧な調理手順（切り方、火加減、加熱時間など）を、初心者でも失敗しないように詳細かつ具体的にテキストで出力してください。"
                            detail_text = generate_via_gemini(f"料理名: {r['name']}\n材料: {r['ingredients']}", sys_prompt, "text")
                            if detail_text:
                                st.session_state.final_plan[i]["detailed_steps"] = detail_text
                                st.session_state.df_MealPlan = pd.DataFrame([{"PlanJSON": json.dumps(st.session_state.final_plan, ensure_ascii=False)}])
                                save_to_excel(); st.rerun()
                
                st.divider()
                st.markdown("**🍽️ アクション**")
                cols_r = st.columns([2, 2, 2])
                with cols_r[0]: rating = st.selectbox("評価", [5,4,3,2,1], format_func=lambda x: "⭐"*x, key=f"rate_{i}", label_visibility="collapsed")
                with cols_r[1]:
                    if st.button("👩‍🍳 調理完了", key=f"btn_{i}", type="primary", use_container_width=True):
                        consume_fifo(edited_ingredients) # ★編集された最新の量で消費計算！
                        new_rating = pd.DataFrame([{"日時": datetime.datetime.now(), "レシピ名": r['name'], "評価": rating}])
                        st.session_state.df_Ratings = pd.concat([st.session_state.df_Ratings, new_rating], ignore_index=True)
                        st.session_state.final_plan.pop(i)
                        st.session_state.df_MealPlan = pd.DataFrame([{"PlanJSON": json.dumps(st.session_state.final_plan, ensure_ascii=False)}])
                        save_to_excel(); st.success("消費記録完了！"); st.rerun()
                with cols_r[2]:
                    # ★追加: 作らずにキャンセルボタン
                    if st.button("🗑️ キャンセル", key=f"del_{i}", use_container_width=True):
                        st.session_state.final_plan.pop(i)
                        st.session_state.df_MealPlan = pd.DataFrame([{"PlanJSON": json.dumps(st.session_state.final_plan, ensure_ascii=False)}])
                        save_to_excel(); st.rerun()

# ------------------------------------------
# Tab 3: 買出 & レシートAI読み取り
# ------------------------------------------
with tab_shop:
    st.header("🛒 買出しリスト")
    if st.button("🔄 定番品の不足分をリストに追加", use_container_width=True):
        inv_total = st.session_state.df_Inventory.groupby("食材名")["残量"].sum().to_dict()
        shop_df = st.session_state.df_ShoppingList
        added = 0
        for _, row in st.session_state.df_Staples.iterrows():
            ing, amt, unit = row["食材名"], float(row["目標量"]), row["単位"]
            # 🍎 定番品の在庫チェックも表記ゆれ対応
            actual_ing = find_closest_item(ing, list(inv_total.keys()))
            current = float(inv_total.get(actual_ing, 0)) if actual_ing else 0.0
            if amt - current > 0:
                shortage = amt - current
                if shop_df[shop_df["食材名"] == ing].empty:
                    new_row = pd.DataFrame([{"買出済": False, "食材名": ing, "カテゴリ": row.get("カテゴリ","日配品"), "必要量": shortage, "単位": unit, "目的": "定番補充"}])
                    shop_df = pd.concat([shop_df, new_row], ignore_index=True); added += 1
        if added > 0:
            st.session_state.df_ShoppingList = shop_df; save_to_excel(); st.rerun()

    if len(st.session_state.df_ShoppingList) > 0:
        edited_shop = st.data_editor(st.session_state.df_ShoppingList, num_rows="dynamic", use_container_width=True, key="ed_shop",
            column_config={
                "買出済": st.column_config.CheckboxColumn("買出済", default=False), 
                "目的": st.column_config.TextColumn(disabled=True),
                "カテゴリ": st.column_config.SelectboxColumn(options=CATEGORY_OPTIONS),
                "単位": st.column_config.SelectboxColumn(options=UNIT_OPTIONS)
            })
        if not edited_shop.equals(st.session_state.df_ShoppingList):
            st.session_state.df_ShoppingList = edited_shop; save_to_excel()

        if st.button("✅ チェック済みの品を在庫へ反映", type="primary", use_container_width=True):
            purchased = edited_shop[edited_shop["買出済"] == True]
            pending = edited_shop[edited_shop["買出済"] == False]
            inv_df = st.session_state.df_Inventory
            today = datetime.date.today()
            for _, row in purchased.iterrows():
                log_transaction(row["食材名"], row.get("カテゴリ", "その他"), "購入", row["必要量"])
                match_idx = inv_df[(inv_df["食材名"] == row["食材名"]) & (inv_df["購入日"] == today)].index
                if not match_idx.empty: inv_df.at[match_idx[0], "残量"] += float(row["必要量"])
                else:
                    new_row = pd.DataFrame([{"食材名": row["食材名"], "カテゴリ": row.get("カテゴリ", "その他"), "残量": float(row["必要量"]), "単位": row["単位"], "購入日": today}])
                    inv_df = pd.concat([inv_df, new_row], ignore_index=True)
            st.session_state.df_Inventory = inv_df
            st.session_state.df_ShoppingList = pending.reset_index(drop=True)
            save_to_excel(); st.rerun()
            
    st.divider()
    # ★追加: レシート写真のAI一発読み取り
    st.subheader("📷 レシートAI自動読み取り")
    uploaded_file = st.file_uploader("レシートを撮影またはアップロード", type=["jpg", "jpeg", "png"])
    
    if uploaded_file is not None:
        if st.button("🪄 レシートを解析して在庫に全自動追加", use_container_width=True):
            with st.spinner("AIがレシートを解読中..."):
                base64_img = base64.b64encode(uploaded_file.read()).decode('utf-8')
                sys_prompt = """レシートの画像から購入した【食品・食材】のみを読み取り、JSON配列で出力せよ。「袋」等の単位は常識的な数量に変換するかそのまま出力。
                [{"name": "食材名", "amount": 数量, "unit": "単位", "category": "カテゴリ"}]"""
                parsed_items = generate_via_gemini("このレシートをデータ化してください。", sys_prompt, "json", base64_img)
                
                if parsed_items:
                    inv_df = st.session_state.df_Inventory
                    today = datetime.date.today()
                    added_str = []
                    for item in parsed_items:
                        ing_name, amount = item.get("name"), float(item.get("amount", 1))
                        log_transaction(ing_name, item.get("category", "その他"), "購入", amount)
                        match_idx = inv_df[(inv_df["食材名"] == ing_name) & (inv_df["購入日"] == today)].index
                        if not match_idx.empty: inv_df.at[match_idx[0], "残量"] += amount
                        else:
                            new_row = pd.DataFrame([{"食材名": ing_name, "カテゴリ": item.get("category", "その他"), "残量": amount, "単位": item.get("unit", "個"), "購入日": today}])
                            inv_df = pd.concat([inv_df, new_row], ignore_index=True)
                        added_str.append(f"{ing_name}({amount}{item.get('unit')})")
                        
                        # 買出しリストから減らす（表記ゆれ対応）
                        shop_df = st.session_state.df_ShoppingList
                        actual_shop_ing = find_closest_item(ing_name, shop_df["食材名"].tolist())
                        if actual_shop_ing:
                            idx = shop_df[shop_df["食材名"] == actual_shop_ing].index[0]
                            new_req = float(shop_df.at[idx, "必要量"]) - amount
                            if new_req <= 0: shop_df = shop_df.drop(idx)
                            else: shop_df.at[idx, "必要量"] = new_req
                            st.session_state.df_ShoppingList = shop_df.reset_index(drop=True)
                            
                    st.session_state.df_Inventory = inv_df
                    save_to_excel(); st.success(f"追加完了: {', '.join(added_str)}")
                else:
                    st.error("解析に失敗しました。画像が不鮮明な可能性があります。")

# ------------------------------------------
# Tab 4: ざっくり個別消費 
# ------------------------------------------
with tab_consume:
    st.header("🍳 ざっくり消費")
    inv_df = st.session_state.df_Inventory
    unique_items = sorted(inv_df["食材名"].unique().tolist())
    target = st.selectbox("🍎 使った食材を選択", unique_items if unique_items else ["(なし)"])
    
    if target and target != "(なし)":
        current_stock = inv_df[inv_df["食材名"] == target]["残量"].sum()
        unit = inv_df[inv_df["食材名"] == target]["単位"].iloc[0]
        st.write(f"現在の在庫: **{current_stock} {unit}**")
        
        # ★追加: 直感的なざっくり消費ボタン（グラムを考えなくて良い）
        c_z1, c_z2, c_z3 = st.columns(3)
        amt_to_consume = 0
        if c_z1.button("🤏 少し(1/4)", use_container_width=True): amt_to_consume = current_stock * 0.25
        if c_z2.button("🌗 半分(1/2)", use_container_width=True): amt_to_consume = current_stock * 0.5
        if c_z3.button("🗑️ 全部(使い切った)", use_container_width=True): amt_to_consume = current_stock
        
        st.divider()
        st.caption("詳細に数値を指定する場合はこちら")
        cols_c = st.columns([2, 1])
        with cols_c[0]: amt_manual = st.number_input("⚖️ 使った量", min_value=0.1, max_value=float(current_stock) if current_stock>0 else 1000.0, value=1.0, step=0.5)
        with cols_c[1]: 
            st.write("") # 高さ合わせ
            if st.button("グラム消費", type="primary", use_container_width=True): amt_to_consume = amt_manual
            
        if amt_to_consume > 0:
            consume_fifo({target: amt_to_consume})
            save_to_excel(); st.success(f"✅ {target} を {amt_to_consume:.1f} 消費しました。"); st.rerun()

# ------------------------------------------
# Tab 5 & Tab 6: マスター管理 & 分析
# ------------------------------------------
with tab_manage:
    sub_inv, sub_staple, sub_seasoning, sub_premade = st.tabs(["📦 在庫", "🥛 定番", "🧂 調味料", "🍛 レトルト"])
    with sub_inv:
        edited_inv = st.data_editor(st.session_state.df_Inventory, num_rows="dynamic", use_container_width=True, key="ed_inv", column_config={"カテゴリ": st.column_config.SelectboxColumn(options=CATEGORY_OPTIONS), "単位": st.column_config.SelectboxColumn(options=UNIT_OPTIONS)})
        if not edited_inv.equals(st.session_state.df_Inventory): st.session_state.df_Inventory = edited_inv; save_to_excel()
    with sub_staple:
        edited_staple = st.data_editor(st.session_state.df_Staples, num_rows="dynamic", use_container_width=True, key="ed_sta", column_config={"カテゴリ": st.column_config.SelectboxColumn(options=CATEGORY_OPTIONS), "単位": st.column_config.SelectboxColumn(options=UNIT_OPTIONS)})
        if not edited_staple.equals(st.session_state.df_Staples): st.session_state.df_Staples = edited_staple; save_to_excel()
    with sub_seasoning:
        edited_season = st.data_editor(st.session_state.df_Seasonings, num_rows="dynamic", use_container_width=True, key="ed_sea")
        if not edited_season.equals(st.session_state.df_Seasonings): st.session_state.df_Seasonings = edited_season; save_to_excel()
    with sub_premade:
        edited_premade = st.data_editor(st.session_state.df_PremadeSauces, num_rows="dynamic", use_container_width=True, key="ed_pre")
        if not edited_premade.equals(st.session_state.df_PremadeSauces): st.session_state.df_PremadeSauces = edited_premade; save_to_excel()

with tab_dash:
    st.header("📊 入出庫トレンド")
    log_df = st.session_state.df_TransactionLog
    if log_df.empty: st.info("買出しや消費を記録するとグラフが表示されます。")
    else:
        log_df["月日"] = log_df["日時"].dt.strftime('%m/%d')
        cols_d = st.columns([1, 1])
        categories = ["すべて"] + list(log_df["カテゴリ"].unique())
        with cols_d[0]: sel_cat = st.selectbox("カテゴリ", categories)
        filtered_df = log_df if sel_cat == "すべて" else log_df[log_df["カテゴリ"] == sel_cat]
        items = ["すべて"] + list(filtered_df["食材名"].unique())
        with cols_d[1]: sel_item = st.selectbox("食材", items)
        final_df = filtered_df if sel_item == "すべて" else filtered_df[filtered_df["食材名"] == sel_item]
        chart_data = final_df.groupby(["月日", "入出庫"])["数量"].sum().unstack(fill_value=0)
        if "購入" not in chart_data: chart_data["購入"] = 0
        if "消費" not in chart_data: chart_data["消費"] = 0
        st.bar_chart(chart_data[["購入", "消費"]], color=["#2e7bcf", "#E03C31"])
        st.divider()
        st.subheader("🏆 AI 評価履歴")
        st.dataframe(st.session_state.df_Ratings.sort_values(by="日時", ascending=False).head(10), use_container_width=True, hide_index=True)
