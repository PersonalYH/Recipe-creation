import streamlit as st
import pandas as pd
import requests
import json
import difflib
import datetime
import io
import msal

# ==========================================
# 0. 初期設定 & モバイル最適化CSS
# ==========================================
st.set_page_config(page_title="献立＆買出しアプリ", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    div[data-testid="stTabs"] > div:nth-child(1) {
        position: sticky; top: 0; z-index: 999; background-color: white;
        padding-top: 10px; padding-bottom: 5px; border-bottom: 1px solid #e6e6e6;
    }
    .block-container { padding-top: 2rem !important; padding-bottom: 5rem !important; }
    .day-badge { color: white; background-color: #E03C31; padding: 3px 8px; border-radius: 4px; font-size: 0.85em; font-weight: bold; margin-right: 8px; }
    .st-emotion-cache-1v0mbdj { margin-top: -15px; } /* UIの隙間微調整 */
    </style>
""", unsafe_allow_html=True)

# 定数
UNIT_OPTIONS = ["g", "個", "玉", "本", "束", "パック", "枚", "ml"]
CATEGORY_OPTIONS = ["青果", "精肉", "鮮魚", "日配品", "加工食品", "調味料", "その他"]
SHEETS = ["Inventory", "ShoppingList", "Staples", "Seasonings", "PremadeSauces", "TransactionLog", "Ratings"]

# ==========================================
# 1. ユーティリティ & AI (Gemini)
# ==========================================
def generate_via_gemini(prompt, sys_prompt=""):
    key = st.secrets.get("GEMINI_API_KEY")
    if not key: return None
    payload = {"contents": [{"parts": [{"text": sys_prompt + "\n\n" + prompt}]}]}
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key={key}"
    try:
        res = requests.post(url, headers={'Content-Type': 'application/json'}, json=payload)
        if res.status_code == 200:
            text = res.json()["candidates"][0]["content"]["parts"][0]["text"]
            text = text.replace("```json", "").replace("```", "").strip()
            return json.loads(text) if text.startswith("[") or text.startswith("{") else text
    except Exception:
        pass
    return None

# ==========================================
# 2. 認証 ＆ OneDrive(Excel) 同期システム
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
# 3. Siri用 API バックドア (ハンズフリー機能)
# ==========================================
# iOSショートカットから /?api=siri&q=質問内容 でアクセスされた場合の隠し処理
if st.query_params.get("api") == "siri":
    query = st.query_params.get("q", "現在の在庫状況を教えて")
    load_from_excel()
    # 全データをJSON化してAIに渡し、Siri向けの短文回答を生成
    db_context = {s: st.session_state.get(f"df_{s}", pd.DataFrame()).to_dict(orient="records") for s in SHEETS}
    sys_prompt = "あなたはユーザーの家事サポートAIです。提供されたJSONデータベースに基づき、ユーザーの質問に音声読み上げに適した【短く簡潔な日本語】で回答してください。挨拶不要。"
    prompt = f"【データベース】\n{json.dumps(db_context, ensure_ascii=False)}\n\n【ユーザーの質問】\n{query}"
    answer = generate_via_gemini(prompt, sys_prompt)
    st.json({"answer": answer if answer else "エラーが発生しました"})
    st.stop()

# ==========================================
# 4. データ初期化 & ロジック
# ==========================================
with st.sidebar:
    st.header("🔑 システム設定")
    if "GEMINI_API_KEY" in st.secrets: st.success("✅ AI連携稼働中")
    if "MS_REFRESH_TOKEN" in st.secrets: st.success("✅ クラウド同期稼働中")
    if st.button("🔄 クラウド再読込"):
        st.session_state.data_loaded = False
        st.rerun()

if "data_loaded" not in st.session_state or not st.session_state.data_loaded:
    with st.spinner("☁️ クラウドと同期中..."):
        load_from_excel() # まず既存のExcelから読めるシートだけ読む
        
        # 旧バージョンのExcelに存在しない新シート用の初期データを定義
        defaults = {
            "df_Inventory": pd.DataFrame([{"食材名": "豚肉", "カテゴリ": "精肉", "残量": 200.0, "単位": "g", "購入日": datetime.date.today()}]),
            "df_ShoppingList": pd.DataFrame(columns=["買出済", "食材名", "カテゴリ", "必要量", "単位", "目的"]),
            "df_Staples": pd.DataFrame([{"食材名": "牛乳", "カテゴリ": "日配品", "目標量": 2.0, "単位": "本"}]),
            "df_Seasonings": pd.DataFrame([{"調味料名": "醤油", "在庫あり": True}, {"調味料名": "みりん", "在庫あり": True}]),
            "df_PremadeSauces": pd.DataFrame([{"商品名": "麻婆豆腐の素", "対応メニュー": "麻婆豆腐"}]),
            "df_TransactionLog": pd.DataFrame(columns=["日時", "食材名", "カテゴリ", "入出庫", "数量"]),
            "df_Ratings": pd.DataFrame(columns=["日時", "レシピ名", "評価"])
        }
        
        # 足りないシートのデータをセッションステートに補完
        needs_save = False
        for key, default_df in defaults.items():
            if key not in st.session_state:
                st.session_state[key] = default_df
                needs_save = True
                
        # 1つでも足りないシートを追加した場合は、クラウドのExcelを上書きして再構築
        if needs_save:
            save_to_excel()
            
        st.session_state.data_loaded = True

# 履歴記録
def log_transaction(item, category, io_type, amount):
    new_log = pd.DataFrame([{"日時": datetime.datetime.now(), "食材名": item, "カテゴリ": category, "入出庫": io_type, "数量": amount}])
    st.session_state.df_TransactionLog = pd.concat([st.session_state.df_TransactionLog, new_log], ignore_index=True)

# 厳密な再計算ロジック
def refresh_shopping_list():
    req_dict = {}
    
    # 1. 確定献立からの必要量
    for r in st.session_state.get("final_plan", []):
        for ing, amt in r.get("ingredients", {}).items():
            unit = r.get("unit_map", {}).get(ing, "個")
            if ing not in req_dict: req_dict[ing] = {"amt": 0.0, "unit": unit, "reason": "献立"}
            req_dict[ing]["amt"] += float(amt)
            
    # 2. 定番品からの必要量
    for _, row in st.session_state.df_Staples.iterrows():
        ing, amt, unit = row["食材名"], float(row["目標量"]), row["単位"]
        if ing not in req_dict: req_dict[ing] = {"amt": 0.0, "unit": unit, "reason": "定番補充"}
        else: req_dict[ing]["amt"] = max(req_dict[ing]["amt"], amt) # 定番と献立で被った場合は多い方を採用

    # 3. 在庫との引き算
    inv_total = st.session_state.df_Inventory.groupby("食材名")["残量"].sum().to_dict()
    new_shop_data = []
    for ing, data in req_dict.items():
        shortage = data["amt"] - float(inv_total.get(ing, 0.0))
        if shortage > 0:
            new_shop_data.append({"買出済": False, "食材名": ing, "カテゴリ": "青果", "必要量": shortage, "単位": data["unit"], "目的": data["reason"]})
            
    st.session_state.df_ShoppingList = pd.DataFrame(new_shop_data) if new_shop_data else pd.DataFrame(columns=["買出済", "食材名", "カテゴリ", "必要量", "単位", "目的"])
    save_to_excel()

def consume_fifo(ingredients_dict):
    df = st.session_state.df_Inventory
    for ing, req_amt in ingredients_dict.items():
        remaining = float(req_amt)
        target_indices = df[df["食材名"] == ing].sort_values("購入日").index.tolist()
        
        # カテゴリ取得用
        cat = df[df["食材名"] == ing]["カテゴリ"].iloc[0] if not df[df["食材名"] == ing].empty else "その他"
        log_transaction(ing, cat, "消費", req_amt)
        
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
    refresh_shopping_list() # 在庫が減ったので買出しリストを再計算

# ==========================================
# UI: 6つのタブ
# ==========================================
if "draft_plan" not in st.session_state: st.session_state.draft_plan = []
if "final_plan" not in st.session_state: st.session_state.final_plan = []
if "selected_order" not in st.session_state: st.session_state.selected_order = []

tab_create, tab_home, tab_shop, tab_consume, tab_manage, tab_dash = st.tabs(["⚙️ 作成", "📅 確定", "🛒 買出し", "🍳 消費", "📦 管理", "📊 分析"])

# ------------------------------------------
# Tab 1: 献立作成 (AIロジック最適化)
# ------------------------------------------
with tab_create:
    st.header("献立の作成")
    mode = st.radio("作成モード", ["在庫消費優先", "リクエスト優先"], horizontal=True)
    target_days = st.number_input("何日分作成しますか？", min_value=1, max_value=7, value=3)
    
    req_prompt = ""
    if mode == "在庫消費優先":
        sorted_df = st.session_state.df_Inventory.sort_values(by="購入日", ascending=True)
        st.dataframe(sorted_df[["購入日", "食材名", "残量", "単位"]].head(5), use_container_width=True, hide_index=True)
        req_prompt = f"在庫データ(古い順): {sorted_df.to_json(orient='records', force_ascii=False)}\n"
    else:
        reqs = [st.text_input(f"Day {d+1} のリクエスト", key=f"r_{d}") for d in range(int(target_days))]
        req_prompt = f"リクエスト: {reqs}\n"

    if st.button("✨ レシピ案を生成", type="primary", use_container_width=True):
        with st.spinner("AIが考案中..."):
            # AI成長システム用: 過去の高評価データ抽出
            hi_rates = st.session_state.df_Ratings[st.session_state.df_Ratings["評価"] >= 4]["レシピ名"].tolist()
            # レトルトデータ抽出
            premades = st.session_state.df_PremadeSauces.to_json(orient='records', force_ascii=False)
            
            sys_prompt = f"""
            プロの料理研究家として以下のJSON配列を厳密に出力してください。
            1. リクエストと消費食材に基づき、純粋に客観的で最適なメニューを考案。
            2. 以下の高評価データから味の好みを推測して反映 (ただし過去メニューそのものの再提案は厳禁)。
            【高評価推測用データ】: {hi_rates}
            3. メニュー決定後、以下の「便利調味料リスト」を確認し、完全に合致するものがある場合のみ、調合せずそれを使う手順を出力。
            【便利調味料リスト】: {premades}
            
            [出力フォーマット]
            [{{"name": "料理名", "intro": "紹介", "ingredients": {{"豚肉": 200}}, "unit_map": {{"豚肉": "g"}}, "seasonings": {{"調味料": ["醤油", "酒"]}}, "steps": [{{"title": "下準備", "desc": "切る"}}], "tips": ["ポイント"]}}]
            """
            res = generate_via_gemini(f"条件: {target_days * 2}品提案。2人前。\n" + req_prompt, sys_prompt)
            if res:
                st.session_state.draft_plan = res
                st.session_state.selected_order = []

    if st.session_state.draft_plan:
        st.divider()
        st.markdown(f"**💡 採用するレシピを {target_days} つ選択**")
        for i, recipe in enumerate(st.session_state.draft_plan):
            is_selected = (i in st.session_state.selected_order)
            badge = f"<span class='day-badge'>DAY {st.session_state.selected_order.index(i) + 1}</span>" if is_selected else ""
            cols = st.columns([1, 8])
            with cols[0]:
                if st.checkbox(" ", value=is_selected, key=f"sel_{i}"):
                    if i not in st.session_state.selected_order: st.session_state.selected_order.append(i)
                else:
                    if i in st.session_state.selected_order: st.session_state.selected_order.remove(i)
            with cols[1]: st.markdown(f"{badge} **{recipe['name']}**", unsafe_allow_html=True)
                
        if len(st.session_state.selected_order) > 0 and st.button("✅ 確定する", type="primary", use_container_width=True):
            st.session_state.final_plan = [st.session_state.draft_plan[idx] for idx in st.session_state.selected_order]
            refresh_shopping_list() # 確定時に自動計算
            st.session_state.draft_plan = []
            st.session_state.selected_order = []
            st.success("確定しました！「📅 確定」タブへ移動してください。")
            st.rerun()

# ------------------------------------------
# Tab 2: 献立確定 & 評価システム
# ------------------------------------------
with tab_home:
    st.header("📅 確定した献立")
    if not st.session_state.final_plan:
        st.info("「⚙️ 作成」タブからメニューを確定させてください。")
    else:
        inv_total = st.session_state.df_Inventory.groupby("食材名")["残量"].sum().to_dict()
        stocked_seasonings = st.session_state.df_Seasonings[st.session_state.df_Seasonings["在庫あり"] == True]["調味料名"].tolist()
        
        for i, r in enumerate(st.session_state.final_plan):
            with st.expander(f"Day {i+1}: {r['name']}", expanded=(i==0)):
                st.markdown(f"*{r.get('intro', '')}*")
                col_left, col_right = st.columns(2)
                
                with col_left:
                    st.markdown("#### 🔪 材料（2人分）")
                    for ing, amt in r.get("ingredients", {}).items():
                        unit = r.get("unit_map", {}).get(ing, "")
                        if float(inv_total.get(ing, 0)) < float(amt): st.error(f"- {ing}: {amt} {unit} ⚠️不足")
                        else: st.write(f"- {ing}: {amt} {unit}")
                        
                with col_right:
                    st.markdown("#### 🧂 調味料・ストック確認")
                    for cat, items in r.get("seasonings", {}).items():
                        st.markdown(f"**【{cat}】**")
                        for s in items:
                            # 調味料リストと照合。リストに無いものは赤字でアラート
                            is_stocked = any(stocked_s in s for stocked_s in stocked_seasonings)
                            if is_stocked: st.write(f"✔️ {s}")
                            else: st.error(f"❌ {s} (在庫リスト未登録)")
                        
                st.markdown("#### 🍳 作り方")
                for step in r.get("steps", []): st.write(f"**{step.get('title', '')}**: {step.get('desc', '')}")
                
                st.divider()
                st.markdown("**🍽️ 食後の評価・記録**")
                cols_r = st.columns([1, 2])
                with cols_r[0]: rating = st.selectbox("星評価", [5,4,3,2,1], format_func=lambda x: "⭐"*x, key=f"rate_{i}")
                with cols_r[1]:
                    if st.button("👩‍🍳 調理完了 (消費記録と評価を保存)", key=f"btn_{i}", type="primary", use_container_width=True):
                        consume_fifo(r.get("ingredients", {}))
                        # 評価を保存
                        new_rating = pd.DataFrame([{"日時": datetime.datetime.now(), "レシピ名": r['name'], "評価": rating}])
                        st.session_state.df_Ratings = pd.concat([st.session_state.df_Ratings, new_rating], ignore_index=True)
                        save_to_excel()
                        st.success("クラウド在庫の更新とAIへの評価送信が完了しました！")
                        st.rerun()

# ------------------------------------------
# Tab 3: 買出しリスト (厳密再計算連動)
# ------------------------------------------
with tab_shop:
    st.header("🛒 買出しリスト")
    st.caption("※在庫管理表の量と必要量を比較し、不足分のみを自動で計算して表示しています。")
    
    if len(st.session_state.df_ShoppingList) == 0:
        st.info("買い出しが必要な食材はありません。すべて充足しています！")
    else:
        edited_shop = st.data_editor(st.session_state.df_ShoppingList, num_rows="dynamic", use_container_width=True, key="ed_shop",
            column_config={"買出済": st.column_config.CheckboxColumn("買出済", default=False), "目的": st.column_config.TextColumn(disabled=True)})
        
        if st.button("✅ チェック済みの品を在庫へ反映", type="primary", use_container_width=True):
            purchased = edited_shop[edited_shop["買出済"] == True]
            inv_df = st.session_state.df_Inventory
            today = datetime.date.today()
            
            for _, row in purchased.iterrows():
                # 購入ログを記録
                log_transaction(row["食材名"], row.get("カテゴリ", "その他"), "購入", row["必要量"])
                match_idx = inv_df[(inv_df["食材名"] == row["食材名"]) & (inv_df["購入日"] == today)].index
                if not match_idx.empty: inv_df.at[match_idx[0], "残量"] += float(row["必要量"])
                else:
                    new_row = pd.DataFrame([{"食材名": row["食材名"], "カテゴリ": row.get("カテゴリ", "その他"), "残量": float(row["必要量"]), "単位": row["単位"], "購入日": today}])
                    inv_df = pd.concat([inv_df, new_row], ignore_index=True)
                    
            st.session_state.df_Inventory = inv_df
            refresh_shopping_list() # 在庫が増えたので買出しリストを再計算（完了したものは自動で消える）
            st.success("在庫に追加し、リストを再計算しました！")
            st.rerun()

# ------------------------------------------
# Tab 4: 個別消費 
# ------------------------------------------
with tab_consume:
    st.header("🍳 個別消費")
    unique_items = sorted(st.session_state.df_Inventory["食材名"].unique().tolist())
    cols_c = st.columns([1, 1])
    with cols_c[0]: target = st.selectbox("🍎 どの食材を使いましたか？", unique_items if unique_items else ["(なし)"])
    with cols_c[1]: amt = st.number_input("⚖️ 使った量", min_value=0.1, value=1.0, step=0.5)
        
    if st.button("一括で消費を記録する", type="primary", use_container_width=True) and unique_items:
        consume_fifo({target: amt})
        st.success(f"✅ {target} を {amt} 消費しました。")
        st.rerun()

# ------------------------------------------
# Tab 5: マスターデータ管理 (サブタブ構成)
# ------------------------------------------
with tab_manage:
    sub_inv, sub_staple, sub_seasoning, sub_premade = st.tabs(["📦 在庫表", "🥛 定番品", "🧂 調味料", "🍛 便利レトルト"])
    
    with sub_inv:
        st.caption("直接編集すると、買い出しリストも連動して再計算されます。")
        edited_inv = st.data_editor(st.session_state.df_Inventory, num_rows="dynamic", use_container_width=True, key="ed_inv")
        if not edited_inv.equals(st.session_state.df_Inventory):
            st.session_state.df_Inventory = edited_inv
            refresh_shopping_list() # 在庫が直接書き換えられたら即座に再計算

    with sub_staple:
        st.caption("常にストックしておきたい量（目標量）を設定します。")
        edited_staple = st.data_editor(st.session_state.df_Staples, num_rows="dynamic", use_container_width=True, key="ed_sta")
        if not edited_staple.equals(st.session_state.df_Staples):
            st.session_state.df_Staples = edited_staple
            refresh_shopping_list() # 定番が変わったら即座に再計算

    with sub_seasoning:
        st.caption("家にある調味料にチェックを入れます。献立確定時に無いものだけアラートが出ます。")
        edited_season = st.data_editor(st.session_state.df_Seasonings, num_rows="dynamic", use_container_width=True, key="ed_sea")
        if not edited_season.equals(st.session_state.df_Seasonings):
            st.session_state.df_Seasonings = edited_season
            save_to_excel()
            
    with sub_premade:
        st.caption("家にある「〇〇の素」やパスタソースを登録。AIが献立に合うと判断した時だけ提案します。")
        edited_premade = st.data_editor(st.session_state.df_PremadeSauces, num_rows="dynamic", use_container_width=True, key="ed_pre")
        if not edited_premade.equals(st.session_state.df_PremadeSauces):
            st.session_state.df_PremadeSauces = edited_premade
            save_to_excel()

# ------------------------------------------
# Tab 6: 分析ダッシュボード (ドリルダウン)
# ------------------------------------------
with tab_dash:
    st.header("📊 入出庫・消費トレンド")
    log_df = st.session_state.df_TransactionLog
    
    if log_df.empty:
        st.info("まだデータの蓄積がありません。買出しや消費を記録するとグラフが表示されます。")
    else:
        log_df["月日"] = log_df["日時"].dt.strftime('%m/%d')
        
        # 1. ドリルダウンUI
        cols_d = st.columns([1, 1])
        categories = ["すべて"] + list(log_df["カテゴリ"].unique())
        with cols_d[0]: sel_cat = st.selectbox("カテゴリを絞り込む", categories)
        
        filtered_df = log_df if sel_cat == "すべて" else log_df[log_df["カテゴリ"] == sel_cat]
        items = ["すべて"] + list(filtered_df["食材名"].unique())
        with cols_d[1]: sel_item = st.selectbox("食材を指定する", items)
        
        final_df = filtered_df if sel_item == "すべて" else filtered_df[filtered_df["食材名"] == sel_item]
        
        st.write("")
        st.subheader("📈 日別 購入 vs 消費量")
        # グラフ用にデータをピボット変形
        chart_data = final_df.groupby(["月日", "入出庫"])["数量"].sum().unstack(fill_value=0)
        if "購入" not in chart_data: chart_data["購入"] = 0
        if "消費" not in chart_data: chart_data["消費"] = 0
        
        st.bar_chart(chart_data[["購入", "消費"]], color=["#2e7bcf", "#E03C31"])
        
        st.divider()
        st.subheader("🏆 AI 評価履歴")
        st.dataframe(st.session_state.df_Ratings.sort_values(by="日時", ascending=False).head(10), use_container_width=True, hide_index=True)
