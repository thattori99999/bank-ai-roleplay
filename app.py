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

# --- Supabase 連携用関数（タイムアウト安全対策版） ---
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
        # タイムアウトを3秒に制限し、応答なしによる全体フリーズを防ぐ
        res = requests.post(f"{url}/rest/v1/roleplay_logs", headers=headers, json=data, timeout=3)
        return res.status_code in [200, 201]
    except Exception as e:
        # 通信エラーやタイムアウトが発生しても、アプリを止めずにFalseを返してスルーする
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
        # タイムアウトを3秒に制限し、管理者画面の無限ローディングを防ぐ
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
    all_sheets = pd.read_excel(file, sheet_name=None)
    text_data = []
    for sheet_name, df in all_sheets.items():
        text_data.append(f"--- シート名: {sheet_name} ---\n{df.to_string(index=False)}")
    return "\n".join(text_data)

# --- Excel出力用の整形ロジック ---
def create_excel_download(text):
    output = io.BytesIO()
    lines = text.split('\n')
    table_data = []
    
    for line in lines:
        if '|' in line:
            cells = [c.strip() for c in line.split('|') if c.strip()]
            if cells and not all(c == '-' or c.startswith('---') for c in cells):
                table_data.append(cells)
                
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if table_data:
            df = pd.DataFrame(table_data)
            df.to_excel(writer, index=False, header=False, sheet_name='評価結果')
            workbook = writer.book
            worksheet = writer.sheets['評価結果']
            border_fmt = workbook.add_format({'border': 1, 'text_wrap': True, 'valign': 'top'})
            for row_num in range(len(table_data)):
                worksheet.set_row(row_num, None, border_fmt)
            worksheet.set_column(0, 5, 30)
        else:
            df = pd.DataFrame([text.split('\n')])
            df.to_excel(writer, index=False, header=False, sheet_name='評価結果')
            
    return output.getvalue()

# --- 404エラーを回避しつつ、利用可能なモデル名を安全に取得する関数 ---
def get_safe_model_name(api_key):
    try:
        genai.configure(api_key=api_key)
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        
        has_flash = any('gemini-1.5-flash' in m for m in available_models)
        target_raw = 'gemini-1.5-flash' if has_flash else (available_models[0] if available_models else 'gemini-1.5-flash')
        
        safe_name = target_raw.replace('models/', '')
        return safe_name
    except:
        return 'gemini-1.5-flash'

# --- ペルソナから最初の挨拶を動的に生成する関数 ---
def generate_initial_greeting(persona, api_key):
    target_model = get_safe_model_name(api_key)
    for attempt in range(5):
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(target_model)
