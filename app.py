import streamlit as st
import pandas as pd
import requests
import json
import difflib
import datetime
import io
import msal
import base64
import math

# ==========================================
# 0. 初期設定 & モバイル最適化CSS
# ==========================================
st.set_page_config(page_title="真・自動化 献立アプリ", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    
""", unsafe_allow_html=True)

UNIT_OPTIONS = ["g", "個", "玉", "本", "束", "枚", "ml", "袋", "パック", "缶", "瓶", "少々", "適量"]
CATEGORY_ORDER = {"青果": 1, "精肉": 2, "鮮魚": 3, "日配品": 4, "加工食品": 5, "調味料": 6, "その他": 7}
SHEETS = ["Inventory", "ShoppingList", "Staples", "Seasonings", "PremadeSauces", "TransactionLog", "Ratings", "MealPlan", "Settings"]

# ==========================================
# 1. AI連携 & ユーティリティ
# ==========================================
def extract_json(text):
    """AIの出力から確実なJSONを抽出する絶対パースエンジン"""
    text = text.replace("```json", "").replace("```", "").strip()
    idx_list = text.find('[')
    idx_dict = text.find('{')
    
    if idx_list != -1 and (idx_dict == -1 or idx_list < idx_dict):
        start = idx_list
        end = text.rfind(']') + 1
    elif idx_dict != -1:
        start = idx_dict
        end = text.rfind('}') + 1
    else:
        return None
        
    try: return json.loads(text[start:end])
    except: return None

def generate_via_gemini(prompt, sys_prompt="", response_type="json", image_b64=None):
    key = st.secrets.get("GEMINI_API_KEY")
    if not key: return None
    parts = [{"text": sys_prompt + "\n\n" + prompt}]
    if image_b64: parts.append({"inline_data": {"mime_type": "image/jpeg", "data": image_b64}})
        
    payload = {"contents": [{"parts": parts}], "generationConfig": {"temperature": 0.5}}
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key={key}"
    
    try:
        res = requests.post(url, headers={'Content-Type': 'application/json'}, json=payload, timeout=60)
        if res.status_code == 200:
            text = res.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
            return extract_json(text) if response_type == "json" else text
    except Exception: pass
    return None

def batch_generate_recipes(req_prompt, sys_prompt_template, total_needed):
    """AIエラーを防ぐ分割バッチ生成エンジン"""
    all_recipes = []
    remaining = total_needed
    while remaining > 0:
        batch_size = min(remaining, 5)
        sys_prompt = sys_prompt_template.replace("{NUM}", str(batch_size))
        res = generate_via_gemini(req_prompt, sys_prompt, "json")
        if res and isinstance(res, list):
            all_recipes.extend(res)
            remaining -= len(res)
        else:
            break
    return all_recipes

def guess_category(item_name):
    """手動追加用・カテゴリ簡易推測エンジン"""
    if any(x in item_name for x in ["肉", "豚", "牛", "鶏"]): return "精肉"
    if any(x in item_name for x in ["魚", "鮭", "鯖", "えび", "イカ"]): return "鮮魚"
    if any(x in item_name for x in ["野菜", "玉ねぎ", "人参", "キャベツ", "トマト", "ネギ", "ピーマン", "大根"]): return "青果"
    if any(x in item_name for x in ["牛乳", "卵", "チーズ", "ヨーグルト"]): return "日配品"
    return "その他"

def round_half_step(num):
    """正確に0.5刻みへ四捨五入する補正ロジック"""
    return math.floor(num * 2 + 0.5) / 2

# ==========================================
# 2. 認証 & 堅牢同期
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
    
    # 409(競合) または 423(ロック) エラーをキャッチ
    res = requests.put(url, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}, data=output.read())
    if res.status_code in [409, 423]:
        st.error("⚠️ エクセルファイルがPC等で開かれロックされています。ファイルを閉じてから再度お試しください。")
        return False
    elif res.status_code >= 400:
        st.error(f"⚠️ 保存エラーが発生しました (コード: {res.status_code})")
        return False
    return True

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
# 3. データ処理ロジック
# ==========================================
def find_closest_item(target_name, item_list, cutoff=0.85):
    if target_name in item_list: return target_name
    matches = difflib.get_close_matches(str(target_name), [str(x) for x in item_list], n=1, cutoff=cutoff)
    return matches[0] if matches else None

def log_transaction(item, category, io_type, amount):
    new_log = pd.DataFrame([{"日時": datetime.datetime.now(), "食材名": item, "カテゴリ": category, "入出庫": io_type, "数量": amount}])
    st.session_state.df_TransactionLog = pd.concat([st.session_state.df_TransactionLog, new_log], ignore_index=True)

def get_reserved_stock():
    reserved = {}
    inv_items = st.session_state.df_Inventory["食材名"].tolist()
    for plan in st.session_state.final_plan:
        for ing, amt in plan.get("ingredients", {}).items():
            actual_ing = find_closest_item(ing, inv_items) or ing
            reserved[actual_ing] = reserved.get(actual_ing, 0.0) + float(amt)
    return reserved

def sync_shopping_list_with_plan():
    shop_df = st.session_state.df_ShoppingList
    mask_protect = (shop_df["目的"] == "手動追加") | (shop_df["買出済"] == True)
    shop_df = shop_df[mask_protect]
    
    inv_total = st.session_state.df_Inventory.groupby("食材名")["残量"].sum().to_dict()
    inv_items_list = list(inv_total.keys())
    premade_list = st.session_state.df_PremadeSauces["商品名"].tolist()
    needed_items = {}

    reserved = get_reserved_stock()
    for ing, reserved_amt in reserved.items():
        if find_closest_item(ing, premade_list): continue
        actual_ing = find_closest_item(ing, inv_items_list) or ing
        shortage = reserved_amt - float(inv_total.get(actual_ing, 0.0))
        if shortage > 0: needed_items[actual_ing] = {"amount": shortage, "unit": "個", "cat": "青果"}

    for _, row in st.session_state.df_Staples.iterrows():
        ing, target_amt, unit = row["食材名"], float(row["目標量"]), row["単位"]
        actual_ing = find_closest_item(ing, inv_items_list) or ing
        usable_stock = float(inv_total.get(actual_ing, 0.0)) - reserved.get(actual_ing, 0.0)
        shortage = target_amt - usable_stock
        if shortage > 0:
            if actual_ing in needed_items:
                needed_items[actual_ing]["amount"] += shortage
            else:
                needed_items[actual_ing] = {"amount": shortage, "unit": unit, "cat": row.get("カテゴリ", "日配品")}

    new_rows = []
    for ing, data in needed_items.items():
        cat = data["cat"]
        if "肉" in ing: cat = "精肉"
        elif "魚" in ing or "鮭" in ing: cat = "鮮魚"
        elif data["unit"] in ["g", "玉", "本", "束"]: cat = "青果"
        
        final_amt = data["amount"]
        if data["unit"] in ["本", "個", "玉", "袋", "パック"]:
            final_amt = math.ceil(final_amt)
        else:
            final_amt = round(final_amt, 1)
            
        new_rows.append({"買出済": False, "食材名": ing, "カテゴリ": cat, "必要量": final_amt, "単位": data["unit"], "目的": "自動計算"})
        
    if new_rows:
        new_df = pd.DataFrame(new_rows)
        new_df["order"] = new_df["カテゴリ"].map(CATEGORY_ORDER).fillna(99)
        new_df = new_df.sort_values("order").drop(columns=["order"])
        st.session_state.df_ShoppingList = pd.concat([shop_df, new_df], ignore_index=True)
    else:
        st.session_state.df_ShoppingList = shop_df

def consume_fifo(ingredients_dict):
    df = st.session_state.df_Inventory
    inv_items = df["食材名"].unique().tolist()
    for req_ing, req_amt in ingredients_dict.items():
        actual_ing = find_closest_item(req_ing, inv_items)
        if not actual_ing: continue 
        remaining = float(req_amt)
        target_indices = df[df["食材名"] == actual_ing].sort_values("購入日").index.tolist()
        cat = df[df["食材名"] == actual_ing]["カテゴリ"].iloc[0]
        log_transaction(actual_ing, cat, "消費", req_amt)
        
        for idx in target_indices:
            if remaining <= 0: break
            current = float(df.at[idx, "残量"])
            if current <= remaining: remaining -= current; df.at[idx, "残量"] = 0.0
            else: df.at[idx, "残量"] = current - remaining; remaining = 0.0
            
    # お掃除ロジック（計算誤差の破棄）
    df = df[df["残量"] >= 0.1].reset_index(drop=True)
    st.session_state.df_Inventory = df

def clear_editor_cache(plan_idx):
    """UIメモリ（キャッシュ）の強制リセット"""
    for k in list(st.session_state.keys()):
        if k.startswith(f"n_{plan_idx}_") or k.startswith(f"a_{plan_idx}_") or k.startswith(f"u_{plan_idx}_"):
            del st.session_state[k]

# ==========================================
# 4. 初期化 & 設定ロード
# ==========================================
with st.sidebar:
    st.header("🔑 システム設定")
    if st.button("🔄 クラウド同期", use_container_width=True):
        st.session_state.data_loaded = False; st.rerun()
    st.divider()
    
    if "df_Settings" in st.session_state:
        settings_dict = st.session_state.df_Settings.set_index("Key")["Value"].to_dict() if not st.session_state.df_Settings.empty else {}
    else: settings_dict = {}
    
    blacklist = st.text_area("🚫 絶対に避ける食材\n(アレルギー・嫌いなもの)", value=settings_dict.get("blacklist", ""))
    if st.button("設定を保存"):
        st.session_state.df_Settings = pd.DataFrame([{"Key": "blacklist", "Value": blacklist}])
        save_to_excel()
        st.success("保存しました")

if "data_loaded" not in st.session_state or not st.session_state.data_loaded:
    with st.spinner("☁️ クラウド同期中..."):
        load_from_excel()
        defaults = {
            "df_Inventory": pd.DataFrame(columns=["食材名", "カテゴリ", "残量", "単位", "購入日"]),
            "df_ShoppingList": pd.DataFrame(columns=["買出済", "食材名", "カテゴリ", "必要量", "単位", "目的"]),
            "df_Staples": pd.DataFrame(columns=["食材名", "カテゴリ", "目標量", "単位"]),
            "df_Seasonings": pd.DataFrame(columns=["調味料名", "在庫あり"]),
            "df_PremadeSauces": pd.DataFrame(columns=["商品名", "対応メニュー"]),
            "df_TransactionLog": pd.DataFrame(columns=["日時", "食材名", "カテゴリ", "入出庫", "数量"]),
            "df_Ratings": pd.DataFrame(columns=["日時", "レシピ名", "評価"]),
            "df_MealPlan": pd.DataFrame(columns=["PlanJSON"]),
            "df_Settings": pd.DataFrame([{"Key": "blacklist", "Value": ""}])
        }
        for key, default_df in defaults.items():
            if key not in st.session_state: st.session_state[key] = default_df
        
        st.session_state.final_plan = []
        df_plan = st.session_state.get("df_MealPlan")
        if df_plan is not None and not df_plan.empty:
            try:
                raw_data = df_plan.iloc[0]["PlanJSON"]
                try: st.session_state.final_plan = json.loads(base64.b64decode(raw_data).decode('utf-8'))
                except: st.session_state.final_plan = json.loads(raw_data)
            except: pass
            
        sync_shopping_list_with_plan()
        st.session_state.data_loaded = True

if "draft_plan" not in st.session_state: st.session_state.draft_plan = []
if "selected_order" not in st.session_state: st.session_state.selected_order = []
if "pending_receipt" not in st.session_state: st.session_state.pending_receipt = None

tab_create, tab_home, tab_shop, tab_consume, tab_manage, tab_dash = st.tabs(["⚙️ 作成", "📅 確定", "🛒 買出", "🍳 消費", "📦 管理", "📊 分析"])

# ------------------------------------------
# Tab 1: 献立作成
# ------------------------------------------
with tab_create:
    st.header("献立の作成")
    mode = st.radio("作成モード", ["在庫消費優先", "リクエスト優先"], horizontal=True)
    mood = st.selectbox("今日の気分", ["おまかせ", "⏱ 疲れた(10分以内)", "🔥 大量消費", "🍷 少し冒険"])
    target_days = st.number_input("何日分作成しますか？", min_value=1, max_value=7, value=1)
    
    req_prompt = f"【テーマ】{mood}\n"
    if mode == "在庫消費優先":
        inv_df = st.session_state.df_Inventory
        reserved = get_reserved_stock()
        usable_inv = []
        for _, r in inv_df.iterrows():
            amt = r['残量'] - reserved.get(r['食材名'], 0.0)
            if amt > 0: usable_inv.append({"食材名": r['食材名'], "残量": amt, "単位": r['単位']})
            
        st.caption("現在フリーで使える在庫(予約分除外)")
        st.dataframe(pd.DataFrame(usable_inv).head(5) if usable_inv else pd.DataFrame(), use_container_width=True, hide_index=True)
        req_prompt += f"【フリー在庫】{json.dumps(usable_inv, ensure_ascii=False)}\n"
        num_proposals = max(5, target_days * 2)
    else:
        reqs = [st.text_input(f"Day {d+1} リクエスト", key=f"r_{d}") for d in range(int(target_days))]
        req_prompt += f"【リクエスト】{reqs}\n"
        num_proposals = max(2, target_days * 2)

    if st.button("✨ レシピ案を生成", type="primary", use_container_width=True):
        with st.spinner("AI考案中..."):
            df_ratings = st.session_state.df_Ratings
            hi_rates = df_ratings[df_ratings["評価"] >= 4]["レシピ名"].tolist() if not df_ratings.empty else []
            low_rates = df_ratings[df_ratings["評価"] <= 2]["レシピ名"].tolist() if not df_ratings.empty else []
            
            premades = st.session_state.df_PremadeSauces.to_json(orient='records', force_ascii=False)
            sys_setting = st.session_state.df_Settings.set_index("Key")["Value"].to_dict() if not st.session_state.df_Settings.empty else {}
            blacklist_str = sys_setting.get("blacklist", "")
            
            sys_prompt_template = f"""
            プロの料理研究家として【{{NUM}}品】のレシピを作成しJSON配列で出力せよ。
            【厳格ルール】分量は必ず「半角数値のみ(小数可)」。少々や適量は不可。
            【禁止食材】以下は絶対に使用しないこと: {blacklist_str}
            【不評ブロック】以下のメニューは過去不評だったため絶対に提案しないこと: {low_rates}
            
            [フォーマット]
            [{{ "name": "料理名", "intro": "紹介", "time": "15分", "ingredients": {{"豚肉": 200}}, "unit_map": {{"豚肉": "g"}}, "seasonings": {{"調味料": ["醤油"]}}, "steps": [{{"title": "下準備", "desc": "簡潔に"}}] }}]
            """
            res = batch_generate_recipes(req_prompt, sys_prompt_template, num_proposals)
            if res:
                st.session_state.draft_plan = res; st.session_state.selected_order = []
            else: st.error("⚠ エラーが発生しました。")

    if st.session_state.draft_plan:
        st.divider()
        st.markdown(f"**💡 採用レシピ選択 ({target_days}つ)**")
        for i, recipe in enumerate(st.session_state.draft_plan):
            is_selected = (i in st.session_state.selected_order)
            cols = st.columns([1, 8])
            with cols[0]:
                if st.checkbox(" ", value=is_selected, key=f"sel_{i}"):
                    if len(st.session_state.selected_order) < target_days and i not in st.session_state.selected_order:
                        st.session_state.selected_order.append(i)
                else:
                    if i in st.session_state.selected_order: st.session_state.selected_order.remove(i)
            with cols[1]: 
                with st.expander(f"{recipe['name']}"):
                    st.caption(recipe.get('intro', ''))
                    st.write(f"**⏱ 予想調理時間:** {recipe.get('time', '不明')}")
                    ings = [f"{k} {v}{recipe.get('unit_map', {}).get(k, '')}" for k, v in recipe.get("ingredients", {}).items()]
                    st.write(f"**🛒 材料:** {', '.join(ings)}")
                    st.write("**🍳 簡単な手順:**")
                    for step in recipe.get("steps", []):
                        st.write(f"・{step.get('title', '')}: {step.get('desc', '')}")
                
        if len(st.session_state.selected_order) > 0 and st.button("✅ 確定する", type="primary", use_container_width=True):
            st.session_state.final_plan.extend([st.session_state.draft_plan[idx] for idx in st.session_state.selected_order])
            
            plan_str = json.dumps(st.session_state.final_plan, ensure_ascii=False)
            plan_b64 = base64.b64encode(plan_str.encode('utf-8')).decode('utf-8')
            st.session_state.df_MealPlan = pd.DataFrame([{"PlanJSON": plan_b64}])
            
            sync_shopping_list_with_plan()
            save_to_excel()
            st.session_state.draft_plan = []; st.session_state.selected_order = []
            st.success("確定しました！「📅 確定」タブへ。"); st.rerun()

# ------------------------------------------
# Tab 2: 献立確定
# ------------------------------------------
with tab_home:
    st.header("📅 確定した献立")
    if not st.session_state.final_plan: st.info("献立がありません。")
    else:
        for i, r in enumerate(st.session_state.final_plan):
            with st.expander(f"Day {i+1}: {r['name']}", expanded=True):
                if r.get("ai_alert"): st.warning(f"👨‍🍳 AIアドバイス:\n{r['ai_alert']}")
                
                st.markdown("#### 🔪 材料編集")
                st.caption("※削除する場合は『食材名』を空にするか、『量』を0に。一番下の空欄で追加。")
                
                current_ings = list(r.get("ingredients", {}).items())
                current_ings.append(("", 0.0)) 
                
                new_ings, new_units = {}, {}
                for ing_idx, (ing_name, ing_amt) in enumerate(current_ings):
                    ing_unit = r.get("unit_map", {}).get(ing_name, "個") if ing_name else "個"
                    c1, c2, c3 = st.columns([5, 3, 3])
                    val_name = c1.text_input("食材", value=ing_name, key=f"n_{i}_{ing_idx}", placeholder="追加する食材", label_visibility="collapsed")
                    val_amt = c2.number_input("量", value=float(ing_amt), min_value=0.0, step=0.5, key=f"a_{i}_{ing_idx}", label_visibility="collapsed")
                    val_unt = c3.selectbox("単位", UNIT_OPTIONS, index=UNIT_OPTIONS.index(ing_unit) if ing_unit in UNIT_OPTIONS else 1, key=f"u_{i}_{ing_idx}", label_visibility="collapsed")
                    
                    if val_name.strip() != "" and val_amt > 0:
                        new_ings[val_name.strip()] = val_amt
                        new_units[val_name.strip()] = val_unt
                
                if st.button("💾 材料の変更を保存 (AIアラートもリセット)", key=f"save_ing_{i}"):
                    st.session_state.final_plan[i]["ingredients"] = new_ings
                    st.session_state.final_plan[i]["unit_map"] = new_units
                    st.session_state.final_plan[i]["ai_alert"] = ""
                    st.session_state.final_plan[i]["detailed_steps"] = "" 
                    
                    plan_str = json.dumps(st.session_state.final_plan, ensure_ascii=False)
                    st.session_state.df_MealPlan = pd.DataFrame([{"PlanJSON": base64.b64encode(plan_str.encode('utf-8')).decode('utf-8')}])
                    sync_shopping_list_with_plan()
                    save_to_excel()
                    clear_editor_cache(i) # キャッシュクリア
                    st.rerun()

                st.markdown("#### 🍳 作り方")
                if r.get("detailed_steps"): st.write(r.get("detailed_steps"))
                else:
                    for step in r.get("steps", []): st.write(f"**{step.get('title', '')}**: {step.get('desc', '')}")
                    if st.button("👨‍🍳 詳細レシピを生成 (AIが材料を添削)", key=f"pro_{i}", type="primary"):
                        with st.spinner("手順を執筆中..."):
                            sys_prompt = f"""
                            以下の【ユーザーが編集した材料】に基づき詳細手順をJSONで出力せよ。
                            1. ユーザーの材料に極力従う。
                            2. 味が極端に薄い等、重大な欠陥がある場合のみ材料を補正し、その理由を `alerts` に記載。
                            {{ "alerts": "補正理由(なければ空)", "suggested_ingredients": {{"食材": 100}}, "suggested_unit_map": {{"食材": "g"}}, "detailed_steps": "手順テキスト" }}
                            """
                            req_data = f"料理名: {r['name']}\nユーザー材料: {new_ings}\n単位: {new_units}"
                            ai_res = generate_via_gemini(req_data, sys_prompt, "json")
                            if ai_res:
                                st.session_state.final_plan[i]["detailed_steps"] = ai_res.get("detailed_steps", "")
                                if ai_res.get("alerts"):
                                    st.session_state.final_plan[i]["ai_alert"] = ai_res.get("alerts", "")
                                    st.session_state.final_plan[i]["ingredients"] = ai_res.get("suggested_ingredients", new_ings)
                                    st.session_state.final_plan[i]["unit_map"] = ai_res.get("suggested_unit_map", new_units)
                                    sync_shopping_list_with_plan()
                                
                                plan_str = json.dumps(st.session_state.final_plan, ensure_ascii=False)
                                st.session_state.df_MealPlan = pd.DataFrame([{"PlanJSON": base64.b64encode(plan_str.encode('utf-8')).decode('utf-8')}])
                                save_to_excel(); st.rerun()
                
                st.divider()
                st.markdown("**評価をして調理完了**")
                st.caption("⭐5: お気に入り！ / ⭐1-2: 二度と提案しない(ブラックリスト入り)")
                cols_r = st.columns(2)
                with cols_r[0]: rating = st.selectbox("評価", [5,4,3,2,1], format_func=lambda x: "⭐"*x, key=f"rate_{i}", label_visibility="collapsed")
                with cols_r[1]:
                    if st.button("👩‍🍳 調理完了", key=f"btn_{i}", type="primary", use_container_width=True):
                        consume_fifo(r.get("ingredients", {}))
                        new_rating = pd.DataFrame([{"日時": datetime.datetime.now(), "レシピ名": r['name'], "評価": rating}])
                        st.session_state.df_Ratings = pd.concat([st.session_state.df_Ratings, new_rating], ignore_index=True)
                        st.session_state.final_plan.pop(i)
                        
                        plan_str = json.dumps(st.session_state.final_plan, ensure_ascii=False)
                        st.session_state.df_MealPlan = pd.DataFrame([{"PlanJSON": base64.b64encode(plan_str.encode('utf-8')).decode('utf-8')}])
                        sync_shopping_list_with_plan()
                        save_to_excel()
                        clear_editor_cache(i)
                        st.success("消費記録完了！"); st.rerun()
                if st.button("🗑️ この献立をキャンセル", key=f"del_{i}"):
                    st.session_state.final_plan.pop(i)
                    plan_str = json.dumps(st.session_state.final_plan, ensure_ascii=False)
                    st.session_state.df_MealPlan = pd.DataFrame([{"PlanJSON": base64.b64encode(plan_str.encode('utf-8')).decode('utf-8')}])
                    sync_shopping_list_with_plan()
                    save_to_excel()
                    clear_editor_cache(i)
                    st.rerun()

# ------------------------------------------
# Tab 3: 買出 & 手動追加 & AIプレビュー
# ------------------------------------------
with tab_shop:
    st.header("🛒 買出しリスト")
    
    c_m1, c_m2 = st.columns([4, 1])
    manual_item = c_m1.text_input("追加する品目", placeholder="パン、ラップなど", label_visibility="collapsed")
    if c_m2.button("＋追加"):
        if manual_item.strip():
            cat_guess = guess_category(manual_item.strip())
            new_row = pd.DataFrame([{"買出済": False, "食材名": manual_item.strip(), "カテゴリ": cat_guess, "必要量": 1, "単位": "個", "目的": "手動追加"}])
            st.session_state.df_ShoppingList = pd.concat([st.session_state.df_ShoppingList, new_row], ignore_index=True)
            sync_shopping_list_with_plan() # 並び替えを実行
            save_to_excel(); st.rerun()

    if len(st.session_state.df_ShoppingList) > 0:
        edited_shop = st.data_editor(st.session_state.df_ShoppingList, num_rows="dynamic", use_container_width=True, key="ed_shop",
            column_config={"カテゴリ": None, "目的": st.column_config.TextColumn(disabled=True), "買出済": st.column_config.CheckboxColumn("買出済", default=False)})
        
        if st.button("✅ チェック済の品を在庫へ (リスト状態を保存)", type="primary", use_container_width=True):
            st.session_state.df_ShoppingList = edited_shop
            purchased = edited_shop[edited_shop["買出済"] == True]
            pending = edited_shop[edited_shop["買出済"] == False]
            inv_df = st.session_state.df_Inventory
            today = datetime.date.today()
            for _, row in purchased.iterrows():
                ing = row["食材名"]
                actual_ing = find_closest_item(ing, inv_df["食材名"].tolist()) or ing
                cat = row.get("カテゴリ", "その他")
                log_transaction(actual_ing, cat, "購入", row["必要量"])
                match_idx = inv_df[(inv_df["食材名"] == actual_ing)].index
                if not match_idx.empty: inv_df.at[match_idx[0], "残量"] += float(row["必要量"])
                else:
                    new_row = pd.DataFrame([{"食材名": actual_ing, "カテゴリ": cat, "残量": float(row["必要量"]), "単位": row["単位"], "購入日": today}])
                    inv_df = pd.concat([inv_df, new_row], ignore_index=True)
            st.session_state.df_Inventory = inv_df
            st.session_state.df_ShoppingList = pending.reset_index(drop=True)
            sync_shopping_list_with_plan()
            save_to_excel(); st.rerun()
            
    st.divider()
    st.subheader("📷 レシートAI自動読み取り")
    uploaded_file = st.file_uploader("レシートを撮影", type=["jpg", "jpeg", "png"])
    if uploaded_file is not None:
        if st.button("🪄 画像を解析", use_container_width=True):
            with st.spinner("解読中..."):
                base64_img = base64.b64encode(uploaded_file.read()).decode('utf-8')
                sys_prompt = """レシート画像から食品を読み取りJSON出力せよ。
                【絶対ルール】日用品(ラップ等)、割引額、税金、袋代など、食べられないものは絶対に除外すること。
                【単位ルール】袋やパックは一般的な数値に換算せよ(ピーマン1袋→4個)。
                [{"name": "食材名", "amount": 数量, "unit": "単位", "category": "青果等"}]"""
                parsed_items = generate_via_gemini("データ化", sys_prompt, "json", base64_img)
                if parsed_items:
                    st.session_state.pending_receipt = parsed_items
                else: st.error("解読失敗")

    if st.session_state.pending_receipt:
        st.warning("以下の内容で在庫に追加しますか？（修正可能）")
        df_preview = pd.DataFrame(st.session_state.pending_receipt)
        edited_preview = st.data_editor(df_preview, num_rows="dynamic", use_container_width=True)
        c_p1, c_p2 = st.columns(2)
        if c_p1.button("✅ この内容で在庫に追加", type="primary"):
            inv_df = st.session_state.df_Inventory
            today = datetime.date.today()
            for _, row in edited_preview.iterrows():
                if pd.isna(row.get("name")): continue
                ing = row["name"]
                actual_ing = find_closest_item(ing, inv_df["食材名"].tolist()) or ing
                amount = float(row.get("amount", 1))
                log_transaction(actual_ing, row.get("category", "その他"), "購入", amount)
                match_idx = inv_df[(inv_df["食材名"] == actual_ing)].index
                if not match_idx.empty: inv_df.at[match_idx[0], "残量"] += amount
                else:
                    inv_df = pd.concat([inv_df, pd.DataFrame([{"食材名": actual_ing, "カテゴリ": row.get("category", "その他"), "残量": amount, "単位": row.get("unit", "個"), "購入日": today}])], ignore_index=True)
            st.session_state.df_Inventory = inv_df
            st.session_state.pending_receipt = None
            sync_shopping_list_with_plan()
            save_to_excel(); st.success("追加完了！"); st.rerun()
        if c_p2.button("やり直す"):
            st.session_state.pending_receipt = None; st.rerun()

# ------------------------------------------
# Tab 4: ざっくり消費
# ------------------------------------------
with tab_consume:
    st.header("🍳 ざっくり消費")
    inv_df = st.session_state.df_Inventory
    unique_items = sorted(inv_df["食材名"].unique().tolist())
    target = st.selectbox("🍎 使った食材", unique_items if unique_items else ["(なし)"])
    if target and target != "(なし)":
        current_stock = inv_df[inv_df["食材名"] == target]["残量"].sum()
        unit = inv_df[inv_df["食材名"] == target]["単位"].iloc[0]
        st.write(f"現在の在庫: **{current_stock} {unit}**")
        c_z1, c_z2, c_z3 = st.columns(3)
        amt_to_consume = 0
        if c_z1.button("🤏 少し(1/4)", use_container_width=True): amt_to_consume = current_stock * 0.25
        if c_z2.button("🌗 半分(1/2)", use_container_width=True): amt_to_consume = current_stock * 0.5
        if c_z3.button("🗑️ 全部(使い切り)", use_container_width=True): amt_to_consume = current_stock
        
        st.divider()
        cols_c = st.columns([2, 1])
        with cols_c[0]: amt_manual = st.number_input("⚖️ 数値入力", min_value=0.1, max_value=float(current_stock) if current_stock>0 else 1000.0, value=1.0, step=0.5)
        with cols_c[1]: 
            st.write("") 
            if st.button("数値消費", type="primary", use_container_width=True): amt_to_consume = amt_manual
            
        if amt_to_consume > 0:
            if unit in ["個", "本", "玉"]: amt_to_consume = round_half_step(amt_to_consume)
            consume_fifo({target: amt_to_consume})
            sync_shopping_list_with_plan()
            save_to_excel(); st.rerun()

# ------------------------------------------
# Tab 5 & 6: 管理・分析
# ------------------------------------------
with tab_manage:
    sub_inv, sub_staple, sub_seasoning, sub_premade = st.tabs(["📦 在庫", "🥛 定番", "🧂 調味料", "🍛 レトルト"])
    with sub_inv:
        edited_inv = st.data_editor(st.session_state.df_Inventory, num_rows="dynamic", use_container_width=True, key="ed_inv")
        if not edited_inv.equals(st.session_state.df_Inventory): st.session_state.df_Inventory = edited_inv; sync_shopping_list_with_plan(); save_to_excel()
    with sub_staple:
        edited_staple = st.data_editor(st.session_state.df_Staples, num_rows="dynamic", use_container_width=True, key="ed_sta")
        if not edited_staple.equals(st.session_state.df_Staples): st.session_state.df_Staples = edited_staple; sync_shopping_list_with_plan(); save_to_excel()
    with sub_seasoning:
        edited_season = st.data_editor(st.session_state.df_Seasonings, num_rows="dynamic", use_container_width=True, key="ed_sea")
        if not edited_season.equals(st.session_state.df_Seasonings): st.session_state.df_Seasonings = edited_season; save_to_excel()
    with sub_premade:
        edited_premade = st.data_editor(st.session_state.df_PremadeSauces, num_rows="dynamic", use_container_width=True, key="ed_pre")
        if not edited_premade.equals(st.session_state.df_PremadeSauces): st.session_state.df_PremadeSauces = edited_premade; save_to_excel()

with tab_dash:
    st.header("📊 トレンド")
    log_df = st.session_state.df_TransactionLog
    if not log_df.empty:
        log_df["月日"] = log_df["日時"].dt.strftime('%m/%d')
        cols_d = st.columns([1, 1])
        categories = ["すべて"] + list(log_df["カテゴリ"].unique())
        with cols_d[0]: sel_cat = st.selectbox("カテゴリ", categories)
        filtered_df = log_df if sel_cat == "すべて" else log_df[log_df["カテゴリ"] == sel_cat]
        items = ["すべて"] + list(filtered_df["食材名"].unique())
        with cols_d[1]: sel_item = st.selectbox("食材", items)
        final_df = filtered_df if sel_item == "すべて" else filtered_df[filtered_df["食材名"] == sel_item]
        chart_data = final_df.groupby(["月日", "入出庫"])["数量"].sum().unstack(fill_value=0)
        st.bar_chart(chart_data[["購入", "消費"]] if "購入" in chart_data and "消費" in chart_data else chart_data, color=["#2e7bcf", "#E03C31"] if len(chart_data.columns)==2 else None)
