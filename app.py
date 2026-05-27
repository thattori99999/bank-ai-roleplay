# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import google.generativeai as genai
from google.api_core import exceptions  # Rate Limit(429) エラーを確実に捕捉するため
from docx import Document
from PyPDF2 import PdfReader
from pptx import Presentation
import io
import configparser
import os
import signal
import re
import time  # リトライ待機（スリープ）処理のため
import requests  # Supabase REST API用
import json

# --- 1. APIキーの設定 (APIKEY.ini または クラウドのSecretsからハイブリッド取得) ---
# ※既存のAPI取得ロジックの構造・変数名を1行も崩さずに、クラウド安全対策を内包させています
def load_api_key():
    # 1. まずローカルの APIKEY.ini を探す
    config = configparser.ConfigParser()
    file_path = 'APIKEY.ini'
    if os.path.exists(file_path):
        try:
            config.read(file_path, encoding='utf-8-sig')
            return config.get('GEMINI', 'API_KEY')
        except:
            pass
            
    # 2. もしローカルにファイルがなければ、Streamlit Cloudの「Secrets」から安全に取得する
    try:
        if "GEMINI" in st.secrets and "API_KEY" in st.secrets["GEMINI"]:
            return st.secrets["GEMINI"]["API_KEY"]
        elif "API_KEY" in st.secrets:
            return st.secrets["API_KEY"]
    except:
        return None
        
    return None

INI_KEY = load_api_key()
EMBEDDED_API_KEY = INI_KEY

# --- Supabase 連携用関数（詳細デバッグ版） ---
def get_supabase_config():
    try:
        if "SUPABASE" in st.secrets:
            return st.secrets["SUPABASE"].get("URL"), st.secrets["SUPABASE"].get("KEY")
    except:
        pass
    return None, None

def save_log_to_supabase(username, persona_name, messages, report=None):
    url, key = get_supabase_config()
    if not url or not key:
        st.session_state["db_error"] = "StreamlitのSecretsに[SUPABASE]設定が見つかりません。"
        return False
    
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal"
    }
    
    data = {
        "username": username,
        "persona_name": persona_name,
        "messages": messages,
        "report": report
    }
    
    try:
        # 3秒タイムアウトで通信を実行
        res = requests.post(f"{url}/rest/v1/roleplay_logs", headers=headers, json=data, timeout=3)
        if res.status_code in [200, 201]:
            if "db_error" in st.session_state:
                del st.session_state["db_error"]
            return True
        else:
            # サーバーから返ってきたエラーメッセージを保持
            st.session_state["db_error"] = f"ステータスコード: {res.status_code} | 詳細: {res.text}"
            return False
    except Exception as e:
        st.session_state["db_error"] = f"通信例外エラー: {str(e)}"
        return False

def fetch_logs_from_supabase():
    url, key = get_supabase_config()
    if not url or not key:
        return []
    
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json"
    }
    
    try:
        res = requests.get(f"{url}/rest/v1/roleplay_logs?order=created_at.desc", headers=headers, timeout=3)
        if res.status_code == 200:
            return res.json()
    except Exception as e:
        pass
    return []

# --- 2. 各ファイル抽出関数 ---
def extract_from_docx(file):
    doc = Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_from_pdf(file):
    reader = PdfReader(file)
    return "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])

def extract_from_pptx(file):
    prs = Presentation(file)
    text_runs = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text_runs.append(shape.text)
    return "\n".join(text_runs)

def extract_from_excel(file):
    all_sheets = pd.read_excel(file
