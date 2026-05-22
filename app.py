#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
荐读小助手 — Web 交互版
Flask + Deepseek API
用法: python app.py
"""
import base64
import json
import os
import re
import sys
import time
import uuid
from io import BytesIO

from flask import Flask, render_template, request, jsonify, session, send_file

import pandas as pd
from openai import OpenAI

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import jiandu_deepseek as jd

app = Flask(__name__)
app.secret_key = os.urandom(24)

sessions = {}
EXCEL_PATH = "参考咨询书库馆藏清单.xlsx"
LOGO_PATH = "馆标黑版.png"
_excel_cache = None  # Cache Excel in memory for fast reloads


def _get_df():
    global _excel_cache
    if _excel_cache is None:
        _excel_cache = jd.load_excel(EXCEL_PATH)
    return _excel_cache


def _robust_expand_keywords(client, theme, max_retries=3):
    """Expand keywords with retry logic and graceful fallback."""
    last_err = None
    for attempt in range(max_retries):
        try:
            kws = jd.expand_keywords(client, theme)
            if kws and len(kws) >= 3:
                return kws
            # If returned too few, treat as failure and retry
            if attempt < max_retries - 1:
                print(f"  Keywords too few ({len(kws) if kws else 0}), retrying...")
                time.sleep(1.5)
                continue
        except Exception as e:
            last_err = e
            print(f"  Keyword attempt {attempt+1} failed: {e}")
            if attempt < max_retries - 1:
                time.sleep(2)

    # Fallback: use theme words + split into n-grams
    theme_clean = re.sub(r"[^\w一-鿿]", "", theme)
    fallback = list(dict.fromkeys(re.split(r"[,，\s]+", theme)))
    # Add theme characters as individual keywords for broader matching
    if len(theme_clean) >= 2:
        fallback.append(theme_clean)  # full concatenated theme
        for i in range(len(theme_clean) - 1):
            bigram = theme_clean[i:i+2]
            if bigram not in fallback:
                fallback.append(bigram)
    print(f"  Using fallback keywords ({len(fallback)}): {fallback[:10]}")
    return fallback


def _robust_generate(client, system, user, max_tokens=4096, temperature=0.7, max_retries=2):
    """Call LLM with retry on transient failures."""
    last_err = None
    for attempt in range(max_retries):
        try:
            return jd._call_llm(client, system, user, max_tokens, temperature)
        except Exception as e:
            last_err = e
            if attempt < max_retries - 1:
                time.sleep(2)
    raise last_err

# ── Style presets ──────────────────────────────────────

STYLE_PRESETS = {
    "elegant": {
        "name": "典雅中国风",
        "desc": "暖棕、赭石、衬线字体，适合文化历史类主题",
        "instruction": "Elegant Chinese style: warm browns and ochre tones, serif fonts (Noto Serif SC), "
                       "traditional decorative elements like thin borders and subtle patterns. "
                       "Rich, scholarly but warm atmosphere.",
    },
    "minimal": {
        "name": "现代简约",
        "desc": "留白、浅灰、无衬线，干净利落",
        "instruction": "Modern minimalist: abundant white space, light grey accents, clean sans-serif fonts. "
                       "Thin lines, no excessive decoration. Airy, refined, contemporary feel.",
    },
    "fresh": {
        "name": "清雅书香",
        "desc": "淡绿、蓝灰、舒适阅读感",
        "instruction": "Fresh and airy scholarly style: soft sage greens and slate blues, "
                       "comfortable reading spacing, subtle botanical or natural motifs. "
                       "Calm, inviting, like a quiet reading room.",
    },
    "ink": {
        "name": "水墨意境",
        "desc": "黑白灰、留白、水墨韵味",
        "instruction": "Ink wash painting aesthetic: black, white, and grey palette with generous negative space. "
                       "Bold calligraphic titles, subtle ink-like gradients, minimalist brush-stroke decorative elements. "
                       "Zen-like tranquility and artistic sophistication.",
    },
}


def _logo_uri():
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return "data:image/png;base64," + base64.b64encode(f.read()).decode("ascii")
    return ""


def _html_from_article(client, article_text, theme, style_key):
    """Generate HTML preview from article text with given style."""
    style = STYLE_PRESETS.get(style_key, STYLE_PRESETS["elegant"])
    logo_uri = _logo_uri()

    # Build style-aware system prompt
    system = jd.HTML_SYSTEM + (
        f"\n\nStyle directive: {style['instruction']}\n"
        "Apply this style strictly for this design."
    )
    user = (
        f"Topic: {theme}\n"
        f"Column Header: {jd.HEADER_TEXT}\n"
        f"Slogan: {jd.FOOTER_TEXT}\n"
        f"Library: {jd.LIBRARY_NAME}\n"
        f"Logo: use <img id=\"library-logo\" src=\"LOGO_DATA_URI\"> at top\n\n"
        f"=== ARTICLE (use this exact text) ===\n\n{article_text}\n\n"
        "Generate HTML now."
    )
    html = jd._call_llm(client, system, user, max_tokens=24576, temperature=0.5)
    if logo_uri and "LOGO_DATA_URI" in html:
        html = html.replace("LOGO_DATA_URI", logo_uri)
    html = html.strip()
    for pfx in ["```html", "```"]:
        if html.startswith(pfx):
            html = html[len(pfx):]
    for sfx in ["```"]:
        if html.endswith(sfx):
            html = html[:-len(sfx)]
    html = html.strip()
    if not html.startswith("<!"):
        html = f"<!DOCTYPE html>\n{html}"
    return html


def get_state():
    sid = session.get("sid")
    if not sid or sid not in sessions:
        sid = str(uuid.uuid4())[:8]
        session["sid"] = sid
        sessions[sid] = {
            "step": "theme", "candidates": [], "selected": [],
            "article": "", "theme": "", "html_preview": "",
            "style": "elegant",
        }
    return sessions[sid]


# ── Routes ─────────────────────────────────────────────

@app.route("/")
def index():
    get_state()
    return render_template("index.html", styles=STYLE_PRESETS)


@app.route("/api/search", methods=["POST"])
def search_books():
    state = get_state()
    data = request.get_json()
    theme = data.get("theme", "").strip()
    if not theme:
        return jsonify({"error": "请输入主题"}), 400

    state["theme"] = theme
    state["step"] = "select"

    df = _get_df()
    client = jd._build_client()

    # Robust keyword expansion with retries
    all_kws = _robust_expand_keywords(client, theme)
    user_kws = [k for k in re.split(r"[,,\s]+", theme) if k]
    all_kws = list(dict.fromkeys(user_kws + all_kws))

    # Search
    candidates = jd.rough_select(df, all_kws, top_n=100)

    # If zero results, try broader search with just the theme as substring
    if not candidates and len(all_kws) > 1:
        print(f"  Zero results with {len(all_kws)} keywords, trying broader match...")
        candidates = jd.rough_select(df, all_kws[:3], top_n=100)

    state["candidates"] = candidates

    return jsonify({
        "keywords": all_kws[:20],
        "total": len(candidates),
        "books": [{"idx": i, "title": c["title"], "author": c["author"],
                    "call_no": c["call_no"]}
                  for i, c in enumerate(candidates)],
    })


@app.route("/api/refresh", methods=["POST"])
def refresh_search():
    """Re-run search with the stored theme (stays on select page)."""
    state = get_state()
    theme = state.get("theme", "")
    if not theme:
        return jsonify({"error": "No previous theme"}), 400

    df = _get_df()
    client = jd._build_client()
    all_kws = _robust_expand_keywords(client, theme)
    user_kws = [k for k in re.split(r"[,,\s]+", theme) if k]
    all_kws = list(dict.fromkeys(user_kws + all_kws))

    candidates = jd.rough_select(df, all_kws, top_n=100)
    if not candidates and len(all_kws) > 1:
        candidates = jd.rough_select(df, all_kws[:3], top_n=100)

    state["candidates"] = candidates
    return jsonify({
        "keywords": all_kws[:20],
        "total": len(candidates),
        "books": [{"idx": i, "title": c["title"], "author": c["author"],
                    "call_no": c["call_no"]}
                  for i, c in enumerate(candidates)],
    })


@app.route("/api/curate", methods=["POST"])
def curate_books():
    state = get_state()
    data = request.get_json()
    selected_indices = data.get("indices", [])
    max_books = min(data.get("max", 8), 8)
    candidates = state["candidates"]
    if not candidates:
        return jsonify({"error": "No candidates"}), 400

    if selected_indices:
        books = [candidates[i] for i in selected_indices if i < len(candidates)]
        seen = set()
        unique = []
        for b in books:
            if b["title"] not in seen:
                seen.add(b["title"])
                unique.append(b)
        state["selected"] = unique[:max_books]
    else:
        client = jd._build_client()
        try:
            books = jd.ai_curate_books(client, candidates, state["theme"], max_books)
        except Exception:
            books = candidates[:max_books]
        state["selected"] = books

    return jsonify({
        "selected": [{"idx": i, "title": b["title"], "author": b["author"],
                       "call_no": b["call_no"]}
                      for i, b in enumerate(state["selected"])],
    })


@app.route("/api/generate", methods=["POST"])
def generate_article():
    state = get_state()
    data = request.get_json()
    selected_indices = data.get("indices", [])
    max_books = min(data.get("max", 8), 8)
    candidates = state["candidates"]

    if selected_indices:
        books = []
        seen = set()
        for i in selected_indices:
            if i < len(candidates):
                c = candidates[i]
                if c["title"] not in seen:
                    seen.add(c["title"])
                    books.append(c)
        state["selected"] = books[:max_books]

    if not state["selected"]:
        return jsonify({"error": "请先选择书籍"}), 400

    client = jd._build_client()
    state["step"] = "preview"

    draft = jd.generate_recommendations(client, state["selected"], state["theme"])
    time.sleep(0.3)
    final = jd.polish_review(client, draft, state["theme"])
    state["article"] = final
    state["feedback"] = ""

    style_key = data.get("style", "elegant")
    state["style"] = style_key
    html_preview = _html_from_article(client, final, state["theme"], style_key)
    state["html_preview"] = html_preview

    return jsonify({"html": html_preview, "article": final, "status": "ok"})


@app.route("/api/style", methods=["POST"])
def change_style():
    """Regenerate HTML with a different style (same article text)."""
    state = get_state()
    data = request.get_json()
    style_key = data.get("style", "elegant")

    if not state.get("article"):
        return jsonify({"error": "请先生成荐读稿"}), 400

    state["style"] = style_key
    client = jd._build_client()
    html = _html_from_article(client, state["article"], state["theme"], style_key)
    state["html_preview"] = html
    return jsonify({"html": html, "status": "ok"})


@app.route("/api/revise", methods=["POST"])
def revise_article():
    """AI rewrites based on feedback."""
    state = get_state()
    data = request.get_json()
    feedback = data.get("feedback", "").strip()

    if not feedback:
        return jsonify({"error": "请输入修改意见"}), 400

    client = jd._build_client()
    revise_system = (
        "You are an editor. Revise the article based on the feedback provided. "
        "Preserve the structure: 卷首语 -> 推荐书籍 -> 结语. "
        "Preserve ALL metadata fields (作者, 索书号, 馆藏位置). "
        "Output ONLY the revised full text. Do NOT explain changes."
    )
    user = f"Feedback from editor: {feedback}\n\n=== Current article ===\n\n{state['article']}"
    revised = jd._call_llm(client, revise_system, user, max_tokens=16384, temperature=0.5)
    state["article"] = revised
    state["feedback"] = feedback

    html = _html_from_article(client, revised, state["theme"], state.get("style", "elegant"))
    state["html_preview"] = html
    return jsonify({"html": html, "article": revised, "status": "ok"})


@app.route("/api/edit", methods=["POST"])
def edit_article():
    """Apply manually edited article text and regenerate HTML."""
    state = get_state()
    data = request.get_json()
    new_text = data.get("text", "").strip()

    if not new_text:
        return jsonify({"error": "文本不能为空"}), 400

    state["article"] = new_text
    client = jd._build_client()
    html = _html_from_article(client, new_text, state["theme"], state.get("style", "elegant"))
    state["html_preview"] = html
    return jsonify({"html": html, "status": "ok"})


@app.route("/api/export/<fmt>")
def export_file(fmt):
    state = get_state()
    article = state.get("article", "")
    theme = state.get("theme", "荐读")
    safe = re.sub(r"[^\w一-鿿]+", "_", theme).strip("_")
    if not article:
        return "No article generated", 400

    buf = BytesIO()
    if fmt == "md":
        buf.write(article.encode("utf-8"))
        buf.seek(0)
        return send_file(buf, mimetype="text/markdown",
                         as_attachment=True, download_name=f"jiandu_{safe}.md")
    elif fmt == "docx":
        tmp = f"output/tmp_{safe}.docx"
        os.makedirs("output", exist_ok=True)
        jd.article_to_docx(article, theme, tmp, LOGO_PATH if os.path.exists(LOGO_PATH) else None)
        with open(tmp, "rb") as f:
            buf.write(f.read())
        buf.seek(0)
        os.remove(tmp)
        return send_file(buf, mimetype="application/vnd...wordprocessingml.document",
                         as_attachment=True, download_name=f"jiandu_{safe}.docx")
    elif fmt == "html":
        html = state.get("html_preview", article)
        buf.write(html.encode("utf-8"))
        buf.seek(0)
        return send_file(buf, mimetype="text/html",
                         as_attachment=True, download_name=f"jiandu_{safe}.html")
    return "Unknown format", 400


@app.route("/api/logo")
def get_logo():
    uri = _logo_uri()
    return uri if uri else ("", 404)


@app.route("/api/reset", methods=["POST"])
def reset_session():
    state = get_state()
    state.update({"step": "theme", "candidates": [], "selected": [],
                   "article": "", "html_preview": "", "feedback": ""})
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    print("=" * 50)
    print("  荐读小助手 Web 版 启动中...")
    print("  浏览器打开: http://127.0.0.1:5000")
    print("=" * 50)
    app.run(debug=False, host="127.0.0.1", port=5000)
