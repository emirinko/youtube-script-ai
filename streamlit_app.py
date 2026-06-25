"""杉山先生・台本生成ツール（先生が直接触る画面）

Streamlit 製のシンプルなWeb画面。
1) 「企画を生成する」ボタン → 3つの候補（タイトル＋内容サマリ）を表示
2) 1つ選んで「この企画で台本を生成する」ボタン → テンプレに沿ったExcelを出力
3) Excelをその場でダウンロード。コンプラ点検の結果も表示。

起動:
    streamlit run streamlit_app.py
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
import generate as g  # noqa: E402
import usage  # noqa: E402


st.set_page_config(page_title="杉山先生 台本ジェネレーター", page_icon="🎬", layout="centered")

cfg = g.load_config()

# --- セッション状態の初期化 ---
for key in ("ideas", "result"):
    st.session_state.setdefault(key, None)


# --- APIキー ---
def ensure_api_key() -> bool:
    if os.environ.get("ANTHROPIC_API_KEY"):
        return True
    # クラウド（Streamlit Cloud）の Secrets から読む
    try:
        if "ANTHROPIC_API_KEY" in st.secrets:
            os.environ["ANTHROPIC_API_KEY"] = st.secrets["ANTHROPIC_API_KEY"]
            return True
    except Exception:  # noqa: BLE001
        pass
    with st.sidebar:
        st.subheader("APIキー")
        key = st.text_input("ANTHROPIC_API_KEY", type="password",
                            help="https://console.anthropic.com で発行")
        if key:
            os.environ["ANTHROPIC_API_KEY"] = key
            return True
    st.info("左のサイドバーに APIキー を入力すると使えます。")
    return False


def check_password() -> None:
    """パスワード保護。APP_PASSWORD（環境変数 or Secrets）が設定されていれば要求する。
    未設定ならゲートなし（ローカル開発用）。"""
    expected = os.environ.get("APP_PASSWORD")
    if not expected:
        try:
            expected = st.secrets.get("APP_PASSWORD") if hasattr(st, "secrets") else None
        except Exception:  # noqa: BLE001
            expected = None
    if not expected or st.session_state.get("authed"):
        return
    st.title("🎬 杉山先生 台本ジェネレーター")
    pw = st.text_input("パスワードを入力してください", type="password")
    if pw == expected:
        st.session_state.authed = True
        st.rerun()
    elif pw:
        st.error("パスワードが違います。")
    st.stop()


st.title("🎬 杉山先生 台本ジェネレーター")
st.caption("企画を生成 → 1つ選ぶ → 台本(Excel)を生成。定型文は自動で入ります。出力はAI下書きなので、先生の確認・修正が前提です。")

check_password()

if not ensure_api_key():
    st.stop()


BUDGET_JPY = cfg.get("billing", {}).get("monthly_budget_jpy", 5000)


def pct_of_budget(jpy: float) -> float:
    """金額(JPY)を、今月の枠に対する割合(%)に変換。"""
    return (jpy / BUDGET_JPY * 100) if BUDGET_JPY else 0.0


def render_barometer() -> None:
    """サイドバーに今月のAI利用状況を％で表示（金額は出さない）。"""
    ratio = min(pct_of_budget(usage.month_total_jpy()) / 100, 1.0)
    with st.sidebar:
        st.subheader("📊 今月のAI利用状況")
        st.progress(ratio, text=f"{ratio*100:.0f}%")
        if ratio >= 1.0:
            st.error("今月の利用枠の上限に達しました。担当者にご連絡ください。")
        elif ratio >= 0.8:
            st.warning("今月の利用枠の8割を超えました。")
        st.caption("※今月の利用枠に対するおおよその使用率です。")


render_barometer()


# ============================================================
# STEP 1: 企画を生成
# ============================================================
st.header("STEP 1　企画を生成する")

use_web = st.checkbox(
    "最近の話題・トレンドも反映する（ネット検索。少し時間がかかります）",
    value=True,
    help="X（旧Twitter）やニュースで話題のこと（例：破産者マップ）を拾い、新しい切り口を加えます。",
)

col1, col2 = st.columns([1, 1])
with col1:
    if st.button("💡 企画を3つ生成する", type="primary", use_container_width=True):
        with st.spinner("企画を考えています…" + ("（検索ありで1〜2分）" if use_web else "（15〜20秒ほど）")):
            t0 = usage.now_iso()
            try:
                st.session_state.ideas = g.ideate(cfg, 3, use_web=use_web)
                st.session_state.result = None
                st.toast(f"今回の企画生成：今月の枠の約 {pct_of_budget(usage.since_total_jpy(t0)):.1f}%")
            except Exception as e:  # noqa: BLE001
                st.error(f"生成に失敗しました: {e}")
with col2:
    if st.button("🔄 別の3案を出す", use_container_width=True, disabled=st.session_state.ideas is None):
        with st.spinner("別案を考えています…"):
            t0 = usage.now_iso()
            try:
                st.session_state.ideas = g.ideate(cfg, 3, use_web=use_web)
                st.session_state.result = None
                st.toast(f"今回の企画生成：今月の枠の約 {pct_of_budget(usage.since_total_jpy(t0)):.1f}%")
            except Exception as e:  # noqa: BLE001
                st.error(f"生成に失敗しました: {e}")

ideas = st.session_state.ideas
chosen_theme = None

if ideas:
    st.subheader("候補から1つ選んでください")
    labels = [f"{i+1}. {idea['title']}" for i, idea in enumerate(ideas)]
    pick = st.radio("企画", labels, label_visibility="collapsed")
    idx = labels.index(pick)
    idea = ideas[idx]

    st.markdown(f"### {idea.get('title', '')}")
    st.markdown(f"**内容サマリ**：{idea.get('summary', '-')}")
    st.markdown(f"**🆕 新しい切り口**：{idea.get('fresh_angle', '-')}")

    chosen_theme = st.text_input(
        "テーマ（必要ならここで微調整できます）", value=idea["title"]
    )


# ============================================================
# STEP 2: 台本を生成
# ============================================================
st.divider()
st.header("STEP 2　台本（Excel）を生成する")

manual_theme = st.text_input(
    "（企画を使わず、自分でテーマを入れて作ることもできます）",
    placeholder="例：任意整理をすると家族にバレるのか",
)

theme = manual_theme.strip() or (chosen_theme.strip() if chosen_theme else "")

cite_web = st.checkbox(
    "官公庁などの出典をネット検索して参考記事に入れる（信頼性UP・時間がかかります）",
    value=True,
    help="法務省・裁判所・金融庁・消費者庁・日弁連・官報などの実在資料を検索し、Excelの参考記事欄に資料名とURLを入れます。公開前に必ず内容をご確認ください。",
)

if st.button("📝 この内容で台本を生成する", type="primary", disabled=not theme,
             use_container_width=True):
    with st.spinner(f"「{theme}」の台本を作成中…（{'出典を検索しながらなので3〜6分' if cite_web else '1〜3分'}ほど）"):
        t0 = usage.now_iso()
        try:
            parts = g.generate_parts(cfg, theme, use_web=cite_web)
            rows = g.assemble_rows(cfg, parts.get("theme_label") or theme, parts)
            path = g.output_path(parts["title"])
            g.write_excel(path, cfg, parts, rows)
            body = g.assembled_text(rows)
            total = sum(len(r.text) for r in rows if not r.spacer)
            minutes = round(total / cfg["script_spec"]["chars_per_minute"], 1)
            ng = g.local_ng_scan(cfg, body)
            review = g.llm_review(cfg, body)
            st.session_state.result = {
                "parts": parts, "path": str(path), "rows_text": body,
                "total": total, "minutes": minutes, "ng": ng, "review": review,
                "cost": usage.since_total_jpy(t0),
            }
        except Exception as e:  # noqa: BLE001
            st.error(f"生成に失敗しました: {e}")

res = st.session_state.result
if res:
    parts = res["parts"]
    st.success("台本ができました！")
    st.markdown(f"**タイトル**：{parts['title']}")
    st.markdown(f"**サムネ案**：{parts.get('thumbnail','')}")
    st.markdown(f"**総文字数**：{res['total']} 文字（約 {res['minutes']} 分）")
    st.markdown(f"**この台本生成の利用**：今月の枠の約 {pct_of_budget(res.get('cost', 0)):.1f}%")

    path = Path(res["path"])
    st.download_button(
        "⬇️ Excel台本をダウンロード",
        data=path.read_bytes(),
        file_name=path.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    # --- コンプラ点検 ---
    st.subheader("コンプラ点検")
    if res["ng"]:
        st.warning("NGワード検知（要修正）: " + "、".join(res["ng"]))
    else:
        st.write("✓ NGワードの機械検知：なし")
    issues = res["review"]["issues"]
    if not issues:
        st.write("✓ AIレビュー：重大な問題は見つかりませんでした")
    else:
        st.warning(f"AIレビュー：{len(issues)} 件の指摘")
        for n, it in enumerate(issues, 1):
            with st.expander(f"[{n}] 重要度{it['severity']}／{it['law']}"):
                st.markdown(f"**原文**：{it['quote']}")
                st.markdown(f"**問題**：{it['problem']}")
                st.markdown(f"**修正案**：{it['suggestion']}")
    st.caption(f"総評：{res['review']['overall']}")

    # --- 台本プレビュー ---
    with st.expander("台本の中身をプレビューする"):
        st.text(res["rows_text"])

st.divider()
st.caption("※参考記事は、AIがネット検索で入れた出典候補です。弁護士チャンネルのため、公開前に必ずリンク先と内容（数字・制度・年月など）をご確認ください。検索しない設定のときは空欄になります。")
