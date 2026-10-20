# 🛒 ShopIntel AI

> **Next-Generation Real-Time Multi-Platform Product Comparison & Shopping Intelligence System**

[![FastAPI](https://img.shields.io/badge/Backend-FastAPI-009688?logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com)
[![Python 3.12](https://img.shields.io/badge/Python-3.12-3776AB?logo=python&logoColor=white)](https://www.python.org)
[![Supabase](https://img.shields.io/badge/Database-Supabase-3ECF8E?logo=supabase&logoColor=white)](https://supabase.com)
[![GitHub Pages](https://img.shields.io/badge/Frontend-GitHub%20Pages-222222?logo=githubpages&logoColor=white)](https://pages.github.com)

---

## 🌟 Key Features

- **Real-Time Parallel Scraping**: Live product extraction across **Amazon.in**, **Flipkart**, **Myntra**, **Puma.in**, and **The Souled Store** using asynchronous HTTP and 6-tier fallback heuristics.
- **MCDA Scoring Engine**: Multi-Criteria Decision Analysis evaluating:
  - 🏷️ Price Advantage (35%)
  - ⭐ Rating & Sentiment (25%)
  - 🚚 Delivery Speed (15%)
  - 📦 Stock Availability (15%)
  - 💬 Review Volume Confidence (10%)
- **Persistent Wishlist & Cloud Sync**: Powered by **Supabase PostgreSQL & Realtime**. Save items on one device and see them instantly sync across your phone, tablet, and PC.
- **Zero-Dependency SPA**: High-performance frontend with SVG animated radial score rings, interactive particle canvas, dark/light theme engine, and fully responsive layout.
- **Anti-Bot & Compression Resilience**: Handles Brotli/gzip streams, Cloudflare Turnstile verification, and platform hydration states.

---

## 🏛️ System Architecture

```
                       ┌───────────────────────────────┐
                       │    FRONTEND (GitHub Pages)    │
                       │          index.html           │
                       │   Supabase JS + MCDA Engine   │
                       └───────────────┬───────────────┘
                                       │
                     ┌─────────────────┴─────────────────┐
                     ▼                                   ▼
      ┌─────────────────────────────┐     ┌─────────────────────────────┐
      │      BACKEND (Render)       │     │     DATABASE (Supabase)     │
      │    FastAPI + BeautifulSoup  │     │   PostgreSQL + Realtime     │
      │   Cascade Scraper & MCDA    │     │   Auth & Multi-Device Sync  │
      └──────────────┬──────────────┘     └─────────────────────────────┘
                     │
     ┌───────────────┼───────────────┬───────────────┐
     ▼               ▼               ▼               ▼
  Amazon.in       Flipkart        Myntra          Puma.in
```

---

## 🚀 Local Development Setup

### 1. Backend (FastAPI)
```bash
cd Server/backend
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```
API Documentation will be live at: `http://127.0.0.1:8000/docs`

### 2. Frontend
Open `index.html` directly in your browser or run:
```bash
python -m http.server 3000
```
Navigate to: `http://localhost:3000`

---

## 👥 Authors
- **Soumya Das**
- **Sandip Guria**
- **Sisir Mondal**
- **Dharitri**

*Department of Computer Science & Engineering*
