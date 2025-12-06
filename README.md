# 📊 AI PPT Style Redesigner (AI 簡報風格重塑系統)

> **AIOT 課程 HW 5-3 作業成果**
> 利用 Python 自動化技術，將傳統 PPT 內容提取並進行「版型重構」，一鍵生成多種設計風格。

## 🚀 線上展示 (Live Demo)

點擊下方連結立即體驗：
**[👉 https://aiot-hw5-3.streamlit.app/](https://aiot-hw5-3.streamlit.app/)**

---

## 💡 專案簡介 (Project Overview)

本專案旨在解決「簡報美化耗時」的問題。不同於傳統的 AI 生成圖片工具，本系統採用 **Rule-Based Design (規則導向設計)**，利用 `python-pptx` 函式庫對投影片進行物件級別的操控。

### ✨ 功能 (Features)

* **1. 現代側邊欄風格 (Modern Sidebar)**
    * 特點：雜誌排版風格，自動繪製左側視覺引導色塊，標題與內文自動對齊右側。
* **2. 賽博龐克風格 (Cyberpunk HUD)**
    * 特點：深色模式 (Dark Mode)，使用等寬字體 (Consolas) 與霓虹綠配色，並繪製發光裝飾線條，適合技術 Demo。
* **3. 優雅襯線風格 (Elegant Serif)**
    * 特點：大量留白與置中對齊，自動計算座標繪製標題下方的金色裝飾線 (Accent Line)。