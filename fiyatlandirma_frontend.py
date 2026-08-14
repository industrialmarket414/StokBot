"""Fiyatlandırma karar ekranı için lokal web paneli.

Çalıştırma:
    python fiyatlandirma_frontend.py

Tarayıcı:
    http://127.0.0.1:8787

Bu panel ek paket gerektirmez. Python standart kütüphanesi + projede zaten kullanılan
pandas/openpyxl ile çalışır.
"""

from __future__ import annotations

import json
import os
import pickle
import subprocess
import sys
import threading
import time
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
CIKTI_DIR = BASE_DIR / "Cikti"
AYARLAR_DIR = BASE_DIR / "Ayarlar"
FIYAT_ONAY_DOSYASI = CIKTI_DIR / "fiyat_onay_listesi.xlsx"
FIYATLANDIRMA_ESLESME_DOSYASI = CIKTI_DIR / "fiyatlandirma_eslesme_listesi.xlsx"
KATEGORI_KAR_DOSYASI = AYARLAR_DIR / "kategori_kar_oranlari.xlsx"
TSOFT_DOSYASI = BASE_DIR / "Tsoft" / "tsoft_secili_alanlar.xlsx"
KAR_YEDEK_DOSYASI = Path.home() / "OneDrive" / "Desktop" / "karyedek.xlsx"
FRONTEND_CACHE_DOSYASI = CIKTI_DIR / "_frontend_fiyat_cache.pkl"
FRONTEND_AKTARIM_DOSYASI = CIKTI_DIR / "_frontend_tsoft_aktarim.json"
HOST = "127.0.0.1"
PORT = 8787
APP_VERSION = "2026-08-13-frontend-aktarim-kontrol-v28"
WATCHED_FILES = [
    Path(__file__).resolve(),
    BASE_DIR / "fiyatlandirma_raporu.py",
    BASE_DIR / "tsoft_fiyat_aktar.py",
]

ISLEM_KILIDI = threading.Lock()
SON_ISLEM = {"calisiyor": False, "baslik": "", "cikti": "", "hata": "", "zaman": ""}
VERI_KILIDI = threading.Lock()
VERI_CACHE = {
    "fiyat_mtime": None,
    "kategori_mtime": None,
    "tsoft_mtime": None,
    "eslesme_mtime": None,
    "fiyat_df": None,
    "kategori_df": None,
}


HTML = r"""
<!doctype html>
<html lang="tr">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>EntesStokBot | Fiyatlandırma Paneli</title>
  <style>
    :root {
      --bg: #f4f7fb;
      --panel: #ffffff;
      --ink: #172033;
      --muted: #667085;
      --line: #d9e2f3;
      --blue: #1f4e79;
      --green: #2e7d32;
      --red: #b42318;
      --yellow: #b7791f;
      --soft-blue: #eaf2fb;
      --soft-green: #e8f5e9;
      --soft-red: #ffebee;
      --soft-yellow: #fff8e1;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", Arial, sans-serif;
      color: var(--ink);
      background: var(--bg);
    }
    header {
      background: linear-gradient(135deg, #17365d, #1f4e79);
      color: white;
      padding: 22px 28px;
      box-shadow: 0 2px 14px rgba(23, 54, 93, .22);
    }
    header h1 { margin: 0; font-size: 24px; letter-spacing: .2px; }
    header p { margin: 6px 0 0; color: rgba(255,255,255,.82); }
    main { padding: 22px 28px 36px; }
    .grid { display: grid; gap: 16px; }
    .cards { grid-template-columns: repeat(6, minmax(150px, 1fr)); }
    .card, .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 14px;
      box-shadow: 0 2px 9px rgba(16, 24, 40, .05);
    }
    .card { padding: 15px 16px; }
    .card .label { color: var(--muted); font-size: 12px; }
    .card .value { font-size: 24px; font-weight: 700; margin-top: 5px; }
    .panel { margin-top: 16px; overflow: hidden; }
    .panel-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 14px 16px;
      border-bottom: 1px solid var(--line);
      background: #fbfdff;
    }
    .panel-head h2 { margin: 0; font-size: 17px; color: var(--blue); }
    .panel-body { padding: 14px 16px; }
    .toolbar {
      display: grid;
      grid-template-columns: 1.2fr .9fr .9fr .8fr .9fr auto auto;
      gap: 10px;
      align-items: end;
    }
    label { display: block; color: var(--muted); font-size: 12px; margin-bottom: 5px; }
    input, select {
      width: 100%;
      border: 1px solid #cfd8ea;
      border-radius: 10px;
      padding: 9px 10px;
      background: white;
      color: var(--ink);
      outline: none;
    }
    input:focus, select:focus { border-color: #4f83c2; box-shadow: 0 0 0 3px rgba(79, 131, 194, .16); }
    button {
      border: 0;
      border-radius: 10px;
      padding: 10px 14px;
      font-weight: 700;
      cursor: pointer;
      background: var(--blue);
      color: white;
      white-space: nowrap;
    }
    button.secondary { background: #eef3f8; color: #17365d; border: 1px solid #ccd8e6; }
    button.green { background: var(--green); }
    button.red { background: var(--red); }
    button:disabled { opacity: .55; cursor: not-allowed; }
    .select-toggle {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 0;
      width: 42px;
      height: 42px;
      padding: 0;
      background: #eef3f8;
      color: #17365d;
      border: 1px solid #ccd8e6;
    }
    .select-square {
      width: 18px;
      height: 18px;
      border: 2px solid #8aa4c0;
      border-radius: 5px;
      background: white;
      box-sizing: border-box;
      display: inline-block;
      position: relative;
      flex: 0 0 auto;
    }
    .select-toggle.active .select-square,
    .row-select.active {
      background: #1e5f9f;
      border-color: #1e5f9f;
    }
    .select-toggle.active .select-square::after,
    .row-select.active::after {
      content: "✓";
      color: white;
      position: absolute;
      left: 3px;
      top: -3px;
      font-size: 16px;
      font-weight: 900;
    }
    .row-select {
      width: 18px;
      height: 18px;
      border: 2px solid #8aa4c0;
      border-radius: 5px;
      background: white;
      position: relative;
      cursor: pointer;
      padding: 0;
      display: inline-block;
    }
    tr.selected td { background: #eef6ff; }
    th.select-col, td.select-col { width: 42px; text-align: center; }
    .status {
      white-space: pre-wrap;
      background: #0b1220;
      color: #d7e5ff;
      padding: 12px;
      border-radius: 10px;
      max-height: 260px;
      overflow: auto;
      font-family: Consolas, monospace;
      font-size: 12px;
    }
    .table-wrap { overflow-y: auto; overflow-x: hidden; max-height: 68vh; border: 1px solid var(--line); border-radius: 12px; }
    table { width: 100%; table-layout: fixed; border-collapse: collapse; font-size: 13px; background: white; }
    th, td {
      border-bottom: 1px solid #edf1f7;
      padding: 7px 6px;
      vertical-align: middle;
      white-space: normal;
      overflow-wrap: normal;
      word-break: normal;
    }
    th {
      position: sticky;
      top: 0;
      background: #17365d;
      color: white;
      z-index: 1;
      text-align: center;
      white-space: normal;
      vertical-align: middle;
      line-height: 1.08;
      font-size: 10px;
      padding: 7px 4px;
      overflow-wrap: normal;
      word-break: keep-all;
    }
    .column-filter {
      width: 100%;
      min-width: 0;
      margin-top: 6px;
      padding: 4px 5px;
      border: 1px solid #9fb6cf;
      border-radius: 6px;
      background: white;
      color: #17365d;
      font-size: 10px;
      font-weight: 700;
    }
    td.num { text-align: center; font-variant-numeric: tabular-nums; overflow-wrap: normal; word-break: normal; }
    td.col-marka,
    td.col-urun-web-servis-kodu,
    td.col-stok,
    td.col-onay,
    td.col-aktarim-durumu,
    td.col-rakip-stok-durumu {
      text-align: center;
      vertical-align: middle;
    }
    th.col-urun, td.col-urun { width: 260px; }
    th.col-t-soft-liste, td.col-t-soft-liste { width: 120px; text-align: center; }
    th.col-en-iyi-alis, td.col-en-iyi-alis { width: 155px; }
    th.col-alis-ortalamasi, td.col-alis-ortalamasi { width: 145px; }
    th.col-en-uygun-rakip, td.col-en-uygun-rakip { width: 190px; }
    th.col-fiyatlandirma, td.col-fiyatlandirma { width: 210px; }
    th.col-karar, td.col-karar { width: 170px; }
    .cell-main { font-weight: 800; color: #0f2e52; }
    .cell-muted { color: #5d6f86; font-size: 12px; margin-top: 2px; }
    .cell-lines { line-height: 1.35; }
    .cell-lines div + div { margin-top: 2px; }
    tr:hover td { background: #f7fbff; }
    .badge {
      display: inline-block;
      border-radius: 999px;
      padding: 4px 8px;
      font-weight: 700;
      font-size: 12px;
      white-space: nowrap;
    }
    .ok { background: var(--soft-green); color: var(--green); }
    .bad { background: var(--soft-red); color: var(--red); }
    .wait { background: var(--soft-yellow); color: var(--yellow); }
    .info { background: var(--soft-blue); color: var(--blue); }
    .muted { color: var(--muted); }
    .split {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
    }
    .category-table td input { padding: 6px 8px; min-width: 90px; }
    @media (max-width: 1100px) {
      .cards { grid-template-columns: repeat(2, 1fr); }
      .toolbar { grid-template-columns: 1fr 1fr; }
      .split { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <header>
    <h1>Fiyatlandırma Paneli</h1>
    <p>B2B alış + rakip fiyat + kategori kâr oranı ile otomatik iskonto/onay ekranı</p>
    <p style="font-size:12px;margin-top:4px;opacity:.75;">Sürüm: __APP_VERSION__</p>
  </header>
  <main>
    <section class="panel">
      <div class="panel-head">
        <h2>Fiyat Onay Listesi</h2>
        <button class="secondary" onclick="loadPricing()">Yenile</button>
      </div>
      <div class="panel-body">
        <div class="toolbar">
          <div>
            <label>Arama</label>
            <input id="q" placeholder="Kod, marka, kategori ara..." oninput="resetPricingPageDelayed()" />
          </div>
          <div>
            <label>Marka</label>
            <select id="brand" onchange="resetPricingPage()"></select>
          </div>
          <div>
            <label>Kategori</label>
            <select id="category" onchange="resetPricingPage()"></select>
          </div>
          <div>
            <label>Göster</label>
            <div style="display:flex;gap:8px;align-items:center;">
              <select id="pageSize" onchange="resetPricingPage()">
                <option value="100">100</option>
                <option value="250">250</option>
                <option value="500" selected>500</option>
                <option value="1000">1000</option>
                <option value="all">Tümü</option>
              </select>
              <button class="select-toggle" id="allSelectBtn" onclick="toggleAllVisible()" title="Ekrandaki satırların tümünü seç">
                <span class="select-square"></span>
              </button>
            </div>
          </div>
          <button class="secondary" onclick="clearFilters()">Temizle</button>
          <button onclick="loadPricing()">Filtrele</button>
        </div>
        <div class="toolbar" style="grid-template-columns: auto auto auto 1fr; align-items:center;">
          <button class="secondary" id="prevPage" onclick="changePricingPage(-1)">← Geri</button>
          <button class="secondary" id="nextPage" onclick="changePricingPage(1)">İleri →</button>
          <span class="muted" id="pageInfo"></span>
          <p class="muted" id="count" style="text-align:right;margin:0;"></p>
        </div>
        <div class="table-wrap">
          <table>
            <thead id="priceHead"></thead>
            <tbody id="priceRows"></tbody>
          </table>
        </div>
      </div>
    </section>

    <section class="grid cards" id="cards"></section>

    <section class="panel">
      <div class="panel-head">
        <h2>İşlemler</h2>
        <span class="muted" id="paths"></span>
      </div>
      <div class="panel-body split">
        <div>
          <button onclick="runAction('rapor')">Fiyat Raporunu Oluştur</button>
          <button class="secondary" onclick="runAction('kontrol')">T-Soft Aktarım Kontrol</button>
          <button class="red" onclick="confirmSend()">T-Soft'a Gönder</button>
          <p class="muted">Gönderim butonu yalnızca <b>ONAY=EVET</b> ve durumu <b>AKTARIMA HAZIR</b> veya <b>ZAM YAPILMALI</b> olan satırları gönderir.</p>
        </div>
        <div>
          <div class="status" id="job">Henüz işlem çalışmadı.</div>
        </div>
      </div>
    </section>

    <section class="panel">
      <div class="panel-head">
        <h2>Kategori Kâr Oranları</h2>
        <div>
          <button class="secondary" onclick="loadCategories()">Yenile</button>
          <button class="green" onclick="saveCategories()">Kaydet</button>
        </div>
      </div>
      <div class="panel-body">
        <div class="table-wrap" style="max-height: 360px;">
          <table class="category-table">
            <thead><tr><th>Kategori</th><th>Kar Orani</th></tr></thead>
            <tbody id="categoryRows"></tbody>
          </table>
        </div>
      </div>
    </section>

  </main>

<script>
const fmtNum = (v) => {
  if (v === null || v === undefined || v === "") return "";
  const n = Number(v);
  if (!Number.isFinite(n)) return String(v);
  return n.toLocaleString("tr-TR", { maximumFractionDigits: 2 });
};

let pricingPage = 1;
let pricingTimer = null;
let selectedRows = new Set();
let currentRows = [];
let currentColumns = [];
let activityFilter = "";
let decisionFilter = "";
let decisionFilterOptions = [];
let appliedMarginFilter = "";
let appliedMarginFilterOptions = [];
let competitorFilter = "";
let competitorFilterOptions = [];

function rowKey(row) {
  return String(row.__key || row.ProductId || row["Ürün Web Servis Kodu"] || "");
}

function colClass(c) {
  return "col-" + String(c)
    .toLocaleLowerCase("tr-TR")
    .replaceAll("ı", "i")
    .replaceAll("ğ", "g")
    .replaceAll("ü", "u")
    .replaceAll("ş", "s")
    .replaceAll("ö", "o")
    .replaceAll("ç", "c")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function headerText(c) {
  return String(c)
    .replaceAll(" / ", " ")
    .replaceAll("/", " ")
    .replaceAll("-", " ")
    .split(/\s+/)
    .filter(Boolean)
    .map(kelime => kelime.charAt(0).toLocaleUpperCase("tr-TR") + kelime.slice(1).toLocaleLowerCase("tr-TR"))
    .join("<br>");
}

function titleCase(value) {
  return String(value || "")
    .toLocaleLowerCase("tr-TR")
    .split(/\s+/)
    .filter(Boolean)
    .map(kelime => kelime.charAt(0).toLocaleUpperCase("tr-TR") + kelime.slice(1))
    .join(" ");
}

function displayValue(column, value) {
  if (value === null || value === undefined) return "";
  if (["EN UYGUN RAKİP FİRMA", "AKTARIM DURUMU", "RAKİP STOK DURUMU", "ONAY", "Aktiflik"].includes(column)) {
    return titleCase(value);
  }
  return String(value);
}

function renderMultiline(value) {
  return String(value || "")
    .split("\n")
    .filter(line => line.trim() !== "")
    .map((line, idx) => `<div class="${idx === 0 ? "cell-main" : "cell-muted"}">${line}</div>`)
    .join("");
}

function renderDecision(value) {
  const lines = String(value || "").split("\n").filter(Boolean);
  return `<div class="cell-lines">${lines.map(line => badge(titleCase(line))).join("<div style='height:4px'></div>")}</div>`;
}

function resetPricingPage() {
  pricingPage = 1;
  selectedRows.clear();
  loadPricing();
}

function resetPricingPageDelayed() {
  pricingPage = 1;
  selectedRows.clear();
  clearTimeout(pricingTimer);
  pricingTimer = setTimeout(loadPricing, 350);
}

function changePricingPage(delta) {
  pricingPage = Math.max(1, pricingPage + delta);
  loadPricing();
}

function updateSelectButton() {
  const btn = document.getElementById("allSelectBtn");
  if (!btn) return;
  const keys = currentRows.map(rowKey).filter(Boolean);
  const visibleSelected = keys.filter(k => selectedRows.has(k)).length;
  const allSelected = keys.length > 0 && visibleSelected === keys.length;
  btn.classList.toggle("active", allSelected);
  btn.title = allSelected
    ? `${visibleSelected} satır seçildi; seçimi kaldır`
    : "Ekrandaki satırların tümünü seç";
}

function toggleAllVisible() {
  const keys = currentRows.map(rowKey).filter(Boolean);
  const allSelected = keys.length > 0 && keys.every(k => selectedRows.has(k));
  keys.forEach(k => {
    if (allSelected) selectedRows.delete(k);
    else selectedRows.add(k);
  });
  renderPricingTable();
}

function toggleRowSelection(key) {
  if (!key) return;
  if (selectedRows.has(key)) selectedRows.delete(key);
  else selectedRows.add(key);
  renderPricingTable();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function setDecisionFilter(value) {
  decisionFilter = value;
  resetPricingPage();
}

function setActivityFilter(value) {
  activityFilter = value;
  resetPricingPage();
}

function activityHeader() {
  const options = [
    `<option value="">Tümü</option>`,
    `<option value="AKTİF"${activityFilter === "AKTİF" ? " selected" : ""}>Aktif</option>`,
    `<option value="PASİF"${activityFilter === "PASİF" ? " selected" : ""}>Pasif</option>`,
  ].join("");
  return `${headerText("Aktiflik")}<select class="column-filter" aria-label="Aktiflik filtresi" onchange="setActivityFilter(this.value)">${options}</select>`;
}

function decisionHeader() {
  const options = [`<option value="">Tümü</option>`]
    .concat(decisionFilterOptions.map(value => {
      const safe = escapeHtml(value);
      const selected = value === decisionFilter ? " selected" : "";
      return `<option value="${safe}"${selected}>${safe}</option>`;
    }))
    .join("");
  return `${headerText("Karar")}<select class="column-filter" aria-label="Karar filtresi" onchange="setDecisionFilter(this.value)">${options}</select>`;
}

function setAppliedMarginFilter(value) {
  appliedMarginFilter = value;
  resetPricingPage();
}

function pricingHeader() {
  const options = [`<option value="">Tümü</option>`]
    .concat(appliedMarginFilterOptions.map(value => {
      const raw = String(value);
      const safe = escapeHtml(raw);
      const selected = raw === appliedMarginFilter ? " selected" : "";
      return `<option value="${safe}"${selected}>%${escapeHtml(fmtNum(value))}</option>`;
    }))
    .join("");
  return `${headerText("Fiyatlandırma")}<select class="column-filter" aria-label="Uygulanan kâr filtresi" onchange="setAppliedMarginFilter(this.value)">${options}</select>`;
}

function setCompetitorFilter(value) {
  competitorFilter = value;
  resetPricingPage();
}

function competitorHeader() {
  const options = [
    `<option value="">Tümü</option>`,
    `<option value="__NONE__"${competitorFilter === "__NONE__" ? " selected" : ""}>Rakip Yok</option>`,
  ].concat(competitorFilterOptions.map(value => {
    const safe = escapeHtml(value);
    const selected = value === competitorFilter ? " selected" : "";
    return `<option value="${safe}"${selected}>${safe}</option>`;
  })).join("");
  return `${headerText("En Uygun Rakip")}<select class="column-filter" aria-label="En uygun rakip filtresi" onchange="setCompetitorFilter(this.value)">${options}</select>`;
}

function renderPricingTable() {
  document.getElementById("priceHead").innerHTML =
    `<tr><th class="select-col">Seç</th>${currentColumns.map(c => `<th class="${colClass(c)}">${c === "Aktiflik" ? activityHeader() : c === "Karar" ? decisionHeader() : c === "Fiyatlandırma" ? pricingHeader() : c === "En Uygun Rakip" ? competitorHeader() : headerText(c)}</th>`).join("")}</tr>`;
  const numeric = new Set(["Stok","ORTALAMA ALIŞ","ORTALAMA STOK","ORTALAMA İSKONTO","EN UYGUN SATIŞ FİYATI","EN UYGUN İSKONTO","KATEGORİ KÂR ORANI","SATIŞ FİYATI","REEL KÂR ORANI","HESAPLANAN İSKONTO","SEÇİLEN İSKONTO"]);
  document.getElementById("priceRows").innerHTML = currentRows.map(row => {
    const key = rowKey(row);
    const active = selectedRows.has(key);
    return `<tr class="${active ? "selected" : ""}"><td class="select-col"><button class="row-select ${active ? "active" : ""}" onclick="toggleRowSelection('${String(key).replaceAll("'","\\'")}')"></button></td>${currentColumns.map(c => {
      const v = row[c] ?? "";
      const shown = displayValue(c, v);
      const cls = colClass(c);
      if (c === "Karar") return `<td class="${cls}">${renderDecision(v)}</td>`;
      if (["Ürün", "T-Soft Liste", "En İyi Alış", "Alış Ortalaması", "En Uygun Rakip", "Fiyatlandırma"].includes(c)) return `<td class="${cls}"><div class="cell-lines">${renderMultiline(v)}</div></td>`;
      if (["ONAY","AKTARIM DURUMU","RAKİP STOK DURUMU","Aktiflik"].includes(c)) return `<td class="${cls}">${badge(shown)}</td>`;
      if (numeric.has(c)) return `<td class="num ${cls}">${fmtNum(v)}</td>`;
      return `<td class="${cls}">${shown}</td>`;
    }).join("")}</tr>`;
  }).join("");
  updateSelectButton();
}

      const badge = (value) => {
        const v = String(value || "");
        const norm = v.toLocaleUpperCase("tr-TR");
        let cls = "info";
        let label = v;
        if (norm === "AKTİF" || norm === "AKTIF") {
          cls = "ok";
          label = "● AKTİF";
        } else if (norm === "PASİF" || norm === "PASIF") {
          cls = "bad";
          label = "● PASİF";
        } else if (norm.includes("HAZIR") || norm === "EVET" || norm === "VAR") cls = "ok";
        else if (norm.includes("YOK") || norm.includes("DÜŞÜK") || norm.includes("HATALI") || norm === "HAYIR") cls = "bad";
        else if (norm.includes("BEKLİYOR") || norm.includes("ZAM")) cls = "wait";
        return `<span class="badge ${cls}">${label || ""}</span>`;
      };

async function api(path, options={}) {
  const res = await fetch(path, options);
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "İstek başarısız");
  return data;
}

async function loadSummary() {
  try {
    const data = await api("/api/summary");
    document.getElementById("paths").textContent = data.updated ? `Son rapor: ${data.updated}` : "";
    const cards = [
      ["Toplam Ürün", data.total],
      ["Aktarıma Hazır", data.ready],
      ["Onay EVET", data.yes],
      ["Onay HAYIR", data.no],
      ["Pasif Ürün", data.passive],
      ["Kategori Kâr Eksik", data.missingMargin],
    ];
    document.getElementById("cards").innerHTML = cards.map(([l,v]) =>
      `<div class="card"><div class="label">${l}</div><div class="value">${fmtNum(v)}</div></div>`
    ).join("");
  } catch (e) {
    document.getElementById("cards").innerHTML = `<div class="card"><div class="label">Durum</div><div class="value">Rapor yok</div></div>`;
  }
}

async function loadFilters() {
  try {
    const data = await api("/api/filters");
    const fill = (id, values) => {
      const el = document.getElementById(id);
      const current = el.value;
      el.innerHTML = `<option value="">Tümü</option>` + values.map(v => `<option value="${String(v).replaceAll('"','&quot;')}">${v}</option>`).join("");
      el.value = current;
    };
    fill("brand", data.brands);
    fill("category", data.categories);
    appliedMarginFilterOptions = data.appliedMargins || [];
    competitorFilterOptions = data.competitors || [];
    decisionFilterOptions = ["EVET", "HAYIR", "ZAM YAPILMALI"];
  } catch (e) {
    document.getElementById("count").textContent = "Filtre listeleri yüklenemedi: " + e.message;
  }
}

async function loadPricing() {
  try {
    document.getElementById("count").textContent = "Fiyat listesi yükleniyor...";
    document.getElementById("pageInfo").textContent = "";
    const limit = document.getElementById("pageSize").value || "500";
    const params = new URLSearchParams({
      q: document.getElementById("q").value,
      brand: document.getElementById("brand").value,
      category: document.getElementById("category").value,
      activity: activityFilter,
      decision: decisionFilter,
      applied_margin: appliedMarginFilter,
      competitor: competitorFilter,
      limit,
      page: String(pricingPage),
    });
    const data = await api(`/api/pricing?${params.toString()}`);
    pricingPage = data.page || pricingPage;
    const start = data.filtered ? data.offset + 1 : 0;
    const end = data.offset + data.rows.length;
    document.getElementById("count").textContent = `${data.filtered} ürün bulundu. ${start}-${end} arası gösteriliyor.`;
    document.getElementById("pageInfo").textContent = `Sayfa ${data.page} / ${data.totalPages}`;
    document.getElementById("prevPage").disabled = data.page <= 1;
    document.getElementById("nextPage").disabled = data.page >= data.totalPages;
    currentRows = data.rows || [];
    currentColumns = data.columns || [];
    renderPricingTable();
  } catch (e) {
    document.getElementById("count").textContent = "Fiyat listesi yüklenemedi: " + e.message;
    document.getElementById("pageInfo").textContent = "";
  }
}

async function loadCategories() {
  const data = await api("/api/categories");
  document.getElementById("categoryRows").innerHTML = data.rows.map((r, i) =>
    `<tr><td>${r.Kategori}</td><td><input data-i="${i}" value="${r["Kar Orani"] ?? ""}" /></td></tr>`
  ).join("");
  window.categoryData = data.rows;
}

async function saveCategories() {
  const rows = window.categoryData || [];
  document.querySelectorAll("#categoryRows input").forEach(inp => {
    rows[Number(inp.dataset.i)]["Kar Orani"] = inp.value;
  });
  const result = await api("/api/categories", {
    method: "POST",
    headers: {"Content-Type":"application/json"},
    body: JSON.stringify({rows}),
  });
  alert(`Kategori kâr oranları kaydedildi. Dolu kategori: ${result.filled || 0}. Şimdi fiyat raporunu tekrar oluşturmalısın.`);
  await loadCategories();
  await loadSummary().catch(e => console.warn(e));
  await loadPricing().catch(e => console.warn(e));
}

async function runAction(action, payload={}) {
  document.getElementById("job").textContent = "İşlem başlatılıyor...";
  const data = await api(`/api/run?action=${action}`, {
    method:"POST",
    headers:{"Content-Type":"application/json"},
    body:JSON.stringify(payload),
  });
  document.getElementById("job").textContent = data.message;
  pollJob();
}

function confirmSend() {
  const productIds = [...selectedRows];
  if (!productIds.length) {
    alert("T-Soft'a göndermek için en az bir ürün seçmelisiniz.");
    return;
  }
  if (confirm(`${productIds.length} seçili ürünün hesaplanan iskontosu T-Soft'a gönderilecek. Devam edilsin mi?`)) {
    runAction("gonder", {productIds});
  }
}

async function pollJob() {
  const data = await api("/api/job");
  document.getElementById("job").textContent =
    `${data.zaman || ""} ${data.baslik || ""}\n${data.cikti || ""}\n${data.hata || ""}`.trim() || "Henüz işlem yok.";
  if (data.calisiyor) setTimeout(pollJob, 1500);
  else {
    await loadSummary();
    await loadFilters();
    await loadPricing();
  }
}

function clearFilters() {
  document.getElementById("q").value = "";
  document.getElementById("brand").value = "";
  document.getElementById("category").value = "";
  activityFilter = "";
  decisionFilter = "";
  appliedMarginFilter = "";
  competitorFilter = "";
  resetPricingPage();
}

async function init() {
  await loadSummary().catch(e => console.warn(e));
  await loadFilters().catch(e => console.warn(e));
  await loadCategories().catch(e => {
    document.getElementById("categoryRows").innerHTML = `<tr><td colspan="2">Kategori listesi yüklenemedi: ${e.message}</td></tr>`;
  });
  await loadPricing().catch(e => console.warn(e));
  pollJob();
}

init();
</script>
</body>
</html>
"""


def temiz_deger(deger):
    if pd.isna(deger):
        return ""
    if hasattr(deger, "item"):
        deger = deger.item()
    return deger


def sayisal_deger(deger):
    if deger is None or str(deger).strip() == "":
        return None
    if isinstance(deger, (int, float)):
        return float(deger)
    metin = str(deger).strip().replace("%", "").replace(" ", "")
    try:
        if "," in metin:
            return float(metin.replace(".", "").replace(",", "."))
        return float(metin)
    except Exception:
        return None


def oku_fiyat_onay(data_only: bool = True) -> pd.DataFrame:
    if not FIYAT_ONAY_DOSYASI.exists():
        raise FileNotFoundError(f"Fiyat onay dosyasi yok: {FIYAT_ONAY_DOSYASI}")
    fiyat_mtime = FIYAT_ONAY_DOSYASI.stat().st_mtime
    kategori_mtime = (
        KATEGORI_KAR_DOSYASI.stat().st_mtime
        if KATEGORI_KAR_DOSYASI.exists()
        else None
    )
    tsoft_mtime = TSOFT_DOSYASI.stat().st_mtime if TSOFT_DOSYASI.exists() else None
    eslesme_mtime = (
        FIYATLANDIRMA_ESLESME_DOSYASI.stat().st_mtime
        if FIYATLANDIRMA_ESLESME_DOSYASI.exists()
        else None
    )
    with VERI_KILIDI:
        if (
            VERI_CACHE["fiyat_df"] is not None
            and VERI_CACHE["fiyat_mtime"] == fiyat_mtime
            and VERI_CACHE["kategori_mtime"] == kategori_mtime
            and VERI_CACHE["tsoft_mtime"] == tsoft_mtime
            and VERI_CACHE["eslesme_mtime"] == eslesme_mtime
        ):
            return VERI_CACHE["fiyat_df"].copy()
        cache_anahtari = {
            "version": APP_VERSION,
            "fiyat_mtime": fiyat_mtime,
            "kategori_mtime": kategori_mtime,
            "tsoft_mtime": tsoft_mtime,
            "eslesme_mtime": eslesme_mtime,
        }
        try:
            if FRONTEND_CACHE_DOSYASI.exists():
                with FRONTEND_CACHE_DOSYASI.open("rb") as fh:
                    payload = pickle.load(fh)
                if payload.get("anahtar") == cache_anahtari:
                    df = payload.get("df")
                    if isinstance(df, pd.DataFrame):
                        VERI_CACHE["fiyat_mtime"] = fiyat_mtime
                        VERI_CACHE["kategori_mtime"] = kategori_mtime
                        VERI_CACHE["tsoft_mtime"] = tsoft_mtime
                        VERI_CACHE["eslesme_mtime"] = eslesme_mtime
                        VERI_CACHE["fiyat_df"] = df.copy()
                        return df.copy()
        except Exception:
            pass

    URUN_KODU = "\u00dcr\u00fcn Web Servis Kodu"
    MIN_ORT_ALIS = "M\u0130N. ORT. NET ALI\u015e"
    MIN_ORT_ISK = "M\u0130N. ORT. \u0130SKONTO"
    TSOFT_LISTE = "T-SOFT L\u0130STE F\u0130YATI TL"
    TSOFT_SATIS = "T-SOFT \u0130ND\u0130R\u0130ML\u0130 F\u0130YATI TL"
    EN_DUSUK_RAKIP = "EN D\u00dc\u015e\u00dcK RAK\u0130P F\u0130YATI"
    EN_IYI_ALIS = "EN \u0130Y\u0130 ALI\u015e B2B"
    EN_IYI_ALIS_FIYATI = "EN \u0130Y\u0130 ALI\u015e F\u0130YATI"
    ORT_ALIS = "ORTALAMA ALI\u015e"
    ORT_STOK = "ORTALAMA STOK"
    ORT_ISK = "ORTALAMA \u0130SKONTO"
    EN_UYGUN_RAKIP = "EN UYGUN RAK\u0130P F\u0130RMA"
    RAKIP_STOK = "RAK\u0130P STOK DURUMU"
    EN_UYGUN_SATIS = "EN UYGUN SATI\u015e F\u0130YATI"
    EN_UYGUN_ISK = "EN UYGUN \u0130SKONTO"
    KATEGORI_KAR = "KATEGOR\u0130 K\u00c2R ORANI"
    SATIS = "SATI\u015e F\u0130YATI"
    REEL_KAR = "REEL K\u00c2R ORANI"
    HESAP_ISK = "HESAPLANAN \u0130SKONTO"
    SECILEN_ISK = "SE\u00c7\u0130LEN \u0130SKONTO"

    excel = pd.ExcelFile(FIYAT_ONAY_DOSYASI, engine="openpyxl")
    fiyat_sayfasi = next(
        (ad for ad in excel.sheet_names if ad.lower().replace("ı", "i").startswith("fiyatland")),
        "Fiyatlandırma",
    )
    onay_sayfasi = next(
        (ad for ad in excel.sheet_names if ad.lower().replace("ı", "i") == "fiyat onay"),
        "Fiyat Onay",
    )
    fiyat = pd.read_excel(
        FIYAT_ONAY_DOSYASI,
        sheet_name=fiyat_sayfasi,
        header=1,
        dtype=object,
        engine="openpyxl",
    ).fillna("")
    try:
        onay = pd.read_excel(
            FIYAT_ONAY_DOSYASI,
            sheet_name=onay_sayfasi,
            header=1,
            dtype=object,
            engine="openpyxl",
        ).fillna("")
    except Exception:
        onay = pd.DataFrame(index=fiyat.index)

    def sayi(deger):
        if deger is None or str(deger).strip() == "":
            return None
        if isinstance(deger, (int, float)):
            return float(deger)
        metin = str(deger).strip().replace("%", "").replace(" ", "")
        try:
            if "," in metin:
                return float(metin.replace(".", "").replace(",", "."))
            return float(metin)
        except Exception:
            return None

    def kolon_al(kolon, kaynak=None, default=""):
        kaynak = fiyat if kaynak is None else kaynak
        if kolon in kaynak.columns:
            seri = kaynak[kolon]
            return seri.reindex(fiyat.index).fillna("")
        return pd.Series([default] * len(fiyat), index=fiyat.index)

    df = pd.DataFrame(index=fiyat.index)
    for kolon in ["ProductId", "Marka", "Kategori", URUN_KODU]:
        df[kolon] = kolon_al(kolon)
    df["Stok"] = kolon_al("Stok", onay, "")
    df[TSOFT_LISTE] = kolon_al(TSOFT_LISTE)
    df[TSOFT_SATIS] = kolon_al(TSOFT_SATIS)

    def urun_kimligi(deger):
        if deger is None or str(deger).strip() == "":
            return ""
        if isinstance(deger, float) and deger.is_integer():
            return str(int(deger))
        return str(deger).strip()

    kaynak_sayisi_haritasi = {}
    if FIYATLANDIRMA_ESLESME_DOSYASI.exists():
        try:
            eslesme = pd.read_excel(
                FIYATLANDIRMA_ESLESME_DOSYASI,
                dtype=object,
                engine="openpyxl",
            ).fillna("")
            if "ProductId" in eslesme.columns and "Tedarikci" in eslesme.columns:
                eslesme["_urun_kimligi"] = eslesme["ProductId"].map(urun_kimligi)
                eslesme["_tedarikci"] = eslesme["Tedarikci"].astype(str).str.strip()
                gecerli_eslesme = eslesme[
                    eslesme["_urun_kimligi"].ne("")
                    & eslesme["_tedarikci"].ne("")
                ]
                kaynak_sayisi_haritasi = (
                    gecerli_eslesme.groupby("_urun_kimligi")["_tedarikci"]
                    .nunique()
                    .to_dict()
                )
        except Exception:
            kaynak_sayisi_haritasi = {}
    aktiflik = kolon_al("T-SOFT AKTİFLİK", onay, "")
    fiyat_aktiflik = kolon_al("T-SOFT AKTİFLİK", fiyat, "")
    aktiflik = aktiflik.where(aktiflik.astype(str).str.strip() != "", fiyat_aktiflik)
    if TSOFT_DOSYASI.exists():
        try:
            tsoft = pd.read_excel(
                TSOFT_DOSYASI, dtype=object, engine="openpyxl"
            ).fillna("")
            if "Web Servis Kodu" in tsoft.columns and "T-SOFT AKTİFLİK" in tsoft.columns:
                aktiflik_haritasi = {
                    str(kod).strip(): str(durum).strip()
                    for kod, durum in zip(
                        tsoft["Web Servis Kodu"], tsoft["T-SOFT AKTİFLİK"]
                    )
                    if str(kod).strip()
                }
                tsoft_aktiflik = df[URUN_KODU].astype(str).str.strip().map(
                    aktiflik_haritasi
                ).fillna("")
                aktiflik = aktiflik.where(
                    aktiflik.astype(str).str.strip() != "", tsoft_aktiflik
                )
        except Exception:
            pass
    df["T-SOFT AKTİFLİK"] = aktiflik.astype(str).str.strip().replace(
        {"1": "AKTİF", "0": "PASİF", "AKTIF": "AKTİF", "PASIF": "PASİF"}
    )

    b2b_net_kolonlari = [c for c in fiyat.columns if str(c).endswith(" NET ALIŞ KDV HARİÇ")]
    b2b_stok_kolonlari = [c for c in onay.columns if str(c).endswith(" STOK")]
    b2b_liste_kolonlari = [
        c for c in fiyat.columns
        if str(c).endswith(" LİSTE") and f"{str(c).replace(' LİSTE', '')} İSKONTO" in fiyat.columns
    ]
    b2b_firmalari = sorted({
        str(c).replace(" LİSTE", "")
        for c in b2b_liste_kolonlari
    } | {
        str(c).replace(" NET ALIŞ KDV HARİÇ", "")
        for c in b2b_net_kolonlari
    })

    def satir_b2b_fiyatlari(row):
        fiyatlar = []
        gorulen = set()
        urun_kaynak_sayisi = kaynak_sayisi_haritasi.get(
            urun_kimligi(row.get("ProductId", "")),
            0,
        )
        tek_alis_kaynagi = urun_kaynak_sayisi == 1
        for net_kolon in b2b_net_kolonlari:
            firma = str(net_kolon).replace(" NET ALIŞ KDV HARİÇ", "")
            stok = sayi(row.get(f"{firma} STOK", ""))
            net = sayi(row.get(net_kolon, ""))
            iskonto = sayi(row.get(f"{firma} İSKONTO", ""))
            if net is None:
                liste = sayi(row.get(f"{firma} LİSTE", ""))
                if liste is not None and liste > 0 and iskonto is not None:
                    net = liste * (1 - iskonto / 100)
            stok_gecerli = stok is not None and stok > 0
            liste = sayi(row.get(f"{firma} LİSTE", ""))
            iskonto_gecerli = (
                liste is not None
                and liste > 0
                and iskonto is not None
                and iskonto > 0
            )
            sabit_aksiyon_fiyati = (
                firma.endswith(" SABIT")
                and iskonto is None
                and net is not None
                and net > 0
            )
            kaynak_gecerli = (
                iskonto_gecerli and (tek_alis_kaynagi or stok_gecerli)
            ) or (
                stok_gecerli and sabit_aksiyon_fiyati
            )
            if kaynak_gecerli and net is not None and net > 0:
                fiyatlar.append((firma, net, iskonto))
                gorulen.add(firma)
        for liste_kolon in b2b_liste_kolonlari:
            firma = str(liste_kolon).replace(" LİSTE", "")
            if firma in gorulen:
                continue
            stok = sayi(row.get(f"{firma} STOK", ""))
            liste = sayi(row.get(liste_kolon, ""))
            iskonto = sayi(row.get(f"{firma} İSKONTO", ""))
            if liste is None or liste <= 0 or iskonto is None:
                continue
            if not (
                iskonto > 0
                and (tek_alis_kaynagi or (stok is not None and stok > 0))
            ):
                continue
            net = liste * (1 - iskonto / 100)
            if net > 0:
                fiyatlar.append((firma, net, iskonto))
        return fiyatlar

    en_iyi_b2b = []
    en_iyi_b2b_fiyatlari = []
    ortalama_alislar = []
    ortalama_iskontolar = []
    ortalama_stok = []
    for idx, row in fiyat.iterrows():
        en_iyi_firma = ""
        en_iyi_fiyat = ""
        b2b_fiyatlari = satir_b2b_fiyatlari(row)
        if b2b_fiyatlari:
            en_iyi_kaynak = min(b2b_fiyatlari, key=lambda item: item[1])
            en_iyi_firma = en_iyi_kaynak[0]
            en_iyi_fiyat = en_iyi_kaynak[1]
            netler = [item[1] for item in b2b_fiyatlari if item[1] is not None]
            isklar = [item[2] for item in b2b_fiyatlari if item[2] is not None]
            ortalama_alislar.append(sum(netler) / len(netler) if netler else "")
            ortalama_iskontolar.append(sum(isklar) / len(isklar) if isklar else "")
        else:
            ortalama_alislar.append("")
            ortalama_iskontolar.append("")
        en_iyi_b2b.append(en_iyi_firma)
        en_iyi_b2b_fiyatlari.append(en_iyi_fiyat)
        stok_degerleri = []
        for firma in b2b_firmalari:
            stok_col = f"{firma} STOK"
            stok = sayi(row.get(stok_col, "")) if stok_col in fiyat.columns else None
            if stok is not None:
                stok_degerleri.append(stok)
        if not stok_degerleri and idx in onay.index:
            for kolon in b2b_stok_kolonlari:
                stok = sayi(onay.at[idx, kolon])
                if stok is not None:
                    stok_degerleri.append(stok)
        ortalama_stok.append(sum(stok_degerleri) / len(stok_degerleri) if stok_degerleri else "")

    df[EN_IYI_ALIS] = en_iyi_b2b
    df[EN_IYI_ALIS_FIYATI] = en_iyi_b2b_fiyatlari
    mevcut_ort_alis = kolon_al(MIN_ORT_ALIS)
    df[ORT_ALIS] = [
        fallback if sayi(fallback) is not None else mevcut
        for mevcut, fallback in zip(mevcut_ort_alis, ortalama_alislar)
    ]
    df[ORT_STOK] = ortalama_stok
    mevcut_ort_isk = kolon_al(MIN_ORT_ISK)
    df[ORT_ISK] = [
        fallback if sayi(fallback) is not None else mevcut
        for mevcut, fallback in zip(mevcut_ort_isk, ortalama_iskontolar)
    ]

    rakip_net_kolonlari = [
        c for c in fiyat.columns
        if str(c).endswith(" NET SATIŞ KDV HARİÇ") and "RAKİP" not in str(c)
    ]
    en_uygun_firmalar = []
    en_uygun_satislar = []
    en_uygun_iskontolar = []
    en_uygun_stok_durumlari = []
    for _, row in fiyat.iterrows():
        stoklu_hedef = None
        stoksuz_hedef = None
        for net_kolon in rakip_net_kolonlari:
            firma = str(net_kolon).replace(" NET SATIŞ KDV HARİÇ", "")
            stok = sayi(row.get(f"{firma} STOK", ""))
            net = sayi(row.get(net_kolon, ""))
            if net is None or net <= 0:
                continue
            isk_col = f"{firma} İSKONTO"
            iskonto = row.get(isk_col, "") if isk_col in fiyat.columns else ""
            aday = (firma, net, iskonto)
            if stoksuz_hedef is None or net < stoksuz_hedef[1]:
                stoksuz_hedef = aday
            if stok is not None and stok > 0:
                if stoklu_hedef is None or net < stoklu_hedef[1]:
                    stoklu_hedef = aday
        secilen_rakip = stoklu_hedef or stoksuz_hedef
        en_uygun_firmalar.append(secilen_rakip[0] if secilen_rakip else "")
        en_uygun_satislar.append(secilen_rakip[1] if secilen_rakip else "")
        en_uygun_iskontolar.append(secilen_rakip[2] if secilen_rakip else "")
        en_uygun_stok_durumlari.append(
            "VAR" if stoklu_hedef else ("YOK" if stoksuz_hedef else "")
        )

    df[EN_UYGUN_RAKIP] = en_uygun_firmalar
    df[RAKIP_STOK] = en_uygun_stok_durumlari
    df[EN_UYGUN_SATIS] = en_uygun_satislar
    df[EN_UYGUN_ISK] = en_uygun_iskontolar

    rapor_kar = kolon_al(KATEGORI_KAR, onay)
    if rapor_kar.astype(str).str.strip().ne("").any():
        df[KATEGORI_KAR] = rapor_kar
    else:
        df[KATEGORI_KAR] = ""
    bos_kar_mask = df[KATEGORI_KAR].astype(str).str.strip().eq("")
    if bos_kar_mask.any():
        kategori_oranlari = {}
        try:
            kategori_df = oku_kategoriler()
            for _, row in kategori_df.iterrows():
                kategori = str(row.get("Kategori", "")).strip()
                oran = row.get("Kar Orani", "")
                if kategori:
                    kategori_oranlari[kategori] = oran
        except Exception:
            pass
        fallback_kar = df["Kategori"].map(kategori_oranlari).fillna("").astype(str)
        df.loc[bos_kar_mask, KATEGORI_KAR] = fallback_kar[bos_kar_mask]

    satislar = []
    reel_karlar = []
    hesaplanan_iskontolar = []
    secilen_iskontolar = []
    onaylar = []
    durumlar = []
    for idx, row in df.iterrows():
        aktiflik = str(row.get("T-SOFT AKTİFLİK", "")).strip().upper()
        alis = sayi(row.get(ORT_ALIS, ""))
        kar = sayi(row.get(KATEGORI_KAR, ""))
        rakip_stoklu = str(row.get(RAKIP_STOK, "")).strip().upper() == "VAR"
        rakip_fiyati = sayi(row.get(EN_UYGUN_SATIS, ""))
        rakip_iskontosu = sayi(row.get(EN_UYGUN_ISK, ""))
        rakip = rakip_fiyati if rakip_stoklu else None
        liste = sayi(fiyat.at[idx, TSOFT_LISTE]) if TSOFT_LISTE in fiyat.columns else None
        if alis is not None and alis > 0 and kar is not None:
            min_satis = alis * (1 + kar / 100)
            satis = rakip if rakip is not None and rakip >= min_satis else min_satis
            reel = (satis / alis - 1) * 100
            hesap_isk = (1 - satis / liste) * 100 if liste is not None and liste > 0 else ""
        else:
            satis = ""
            reel = ""
            hesap_isk = ""
        secilen = hesap_isk
        if aktiflik in {"PASİF", "PASIF", "0"}:
            onay, durum = "HAYIR", "PASİF ÜRÜN"
        elif aktiflik not in {"AKTİF", "AKTIF", "1"}:
            onay, durum = "HAYIR", "AKTİFLİK BİLGİSİ YOK"
        elif alis is None or alis <= 0:
            onay, durum = "HAYIR", "ALIŞ FİYATI YOK"
        elif kar is None:
            onay, durum = "HAYIR", "KATEGORİ KÂR ORANI YOK"
        elif (
            rakip_fiyati is not None
            and rakip_fiyati > 0
            and (rakip_iskontosu is None or rakip_iskontosu <= 0)
        ):
            onay, durum = "HAYIR", "RAKİP İSKONTOSU YOK"
        elif satis == "":
            onay, durum = "HAYIR", "SATIŞ FİYATI YOK"
        elif secilen == "":
            onay, durum = "HAYIR", "İSKONTO BEKLİYOR"
        elif sayi(secilen) is None or sayi(secilen) >= 100:
            onay, durum = "HAYIR", "İSKONTO HATALI"
        elif sayi(secilen) < 0:
            onay, durum = "EVET", "ZAM YAPILMALI"
        elif sayi(reel) is None or sayi(reel) + 0.0001 < kar:
            onay, durum = "HAYIR", "KÂR ORANI DÜŞÜK"
        else:
            onay, durum = "EVET", "AKTARIMA HAZIR"
        satislar.append(satis)
        reel_karlar.append(reel)
        hesaplanan_iskontolar.append(hesap_isk)
        secilen_iskontolar.append(secilen)
        onaylar.append(onay)
        durumlar.append(durum)

    df[SATIS] = satislar
    df[REEL_KAR] = reel_karlar
    df[HESAP_ISK] = hesaplanan_iskontolar
    df[SECILEN_ISK] = secilen_iskontolar
    df["ONAY"] = onaylar
    df["AKTARIM DURUMU"] = durumlar

    with VERI_KILIDI:
        VERI_CACHE["fiyat_mtime"] = fiyat_mtime
        VERI_CACHE["kategori_mtime"] = kategori_mtime
        VERI_CACHE["tsoft_mtime"] = tsoft_mtime
        VERI_CACHE["eslesme_mtime"] = eslesme_mtime
        VERI_CACHE["fiyat_df"] = df.copy()
        try:
            with FRONTEND_CACHE_DOSYASI.open("wb") as fh:
                pickle.dump(
                    {
                        "anahtar": {
                            "version": APP_VERSION,
                            "fiyat_mtime": fiyat_mtime,
                            "kategori_mtime": kategori_mtime,
                            "tsoft_mtime": tsoft_mtime,
                            "eslesme_mtime": eslesme_mtime,
                        },
                        "df": df,
                    },
                    fh,
                    protocol=pickle.HIGHEST_PROTOCOL,
                )
        except Exception:
            pass
    return df

def oku_kategoriler() -> pd.DataFrame:
    kaynak_dosya = (
        KAR_YEDEK_DOSYASI
        if KAR_YEDEK_DOSYASI.exists()
        else FIYAT_ONAY_DOSYASI
        if FIYAT_ONAY_DOSYASI.exists()
        else KATEGORI_KAR_DOSYASI
    )
    if not kaynak_dosya.exists():
        raise FileNotFoundError(f"Kategori kâr dosyası yok: {KATEGORI_KAR_DOSYASI}")
    kategori_mtime = kaynak_dosya.stat().st_mtime
    with VERI_KILIDI:
        if (
            VERI_CACHE["kategori_df"] is not None
            and VERI_CACHE["kategori_mtime"] == kategori_mtime
        ):
            return VERI_CACHE["kategori_df"].copy()
    if KAR_YEDEK_DOSYASI.exists():
        df = pd.read_excel(KAR_YEDEK_DOSYASI, dtype=object, engine="openpyxl").fillna("")
    else:
        try:
            df = pd.read_excel(
                FIYAT_ONAY_DOSYASI,
                sheet_name="Kategori Kar Oranları",
                dtype=object,
                engine="openpyxl",
            ).fillna("")
        except Exception:
            df = pd.read_excel(KATEGORI_KAR_DOSYASI, dtype=object, engine="openpyxl").fillna("")
    if "Kategori" not in df.columns:
        df["Kategori"] = ""
    if "Kar Orani" not in df.columns:
        df["Kar Orani"] = ""
    df = df[["Kategori", "Kar Orani"]]
    with VERI_KILIDI:
        VERI_CACHE["kategori_mtime"] = kategori_mtime
        VERI_CACHE["kategori_df"] = df.copy()
    return df


def df_json_rows(df: pd.DataFrame) -> list[dict]:
    rows = []
    for _, row in df.iterrows():
        kayit = {str(k): temiz_deger(v) for k, v in row.to_dict().items()}
        key = kayit.get("ProductId") or kayit.get("Ürün Web Servis Kodu") or ""
        kayit["__key"] = key
        kayit.pop("ProductId", None)
        rows.append(kayit)
    return rows


def ozet_deger(deger, ek=""):
    if deger is None or pd.isna(deger) or str(deger).strip() == "":
        return "-"
    if isinstance(deger, (int, float)):
        metin = f"{deger:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        if metin.endswith(",00"):
            metin = metin[:-3]
        return f"{metin}{ek}"
    return f"{deger}{ek}"


def fiyat_deger(deger):
    deger_str = ozet_deger(deger)
    if deger_str == "-":
        return deger_str
    return f"{deger_str} TL"


def baslik_yap(deger):
    metin = str(deger or "").strip()
    if not metin:
        return ""
    return " ".join(kelime[:1].upper() + kelime[1:].lower() for kelime in metin.split())


def fiyat_gorunum_df(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df.iterrows():
        marka = row.get("Marka", "")
        kod = row.get("Ürün Web Servis Kodu", "")
        kategori = row.get("Kategori", "")
        urun = "\n".join(x for x in [str(marka).strip(), str(kod).strip(), str(kategori).strip()] if x)

        en_iyi_alis = "\n".join([
            baslik_yap(row.get("EN İYİ ALIŞ B2B", "")) or "-",
            f"Net alış: {fiyat_deger(row.get('EN İYİ ALIŞ FİYATI', ''))}",
        ])
        alis_ortalamasi = "\n".join([
            f"Ort. alış: {fiyat_deger(row.get('ORTALAMA ALIŞ', ''))}",
            f"Ort. isk.: {ozet_deger(row.get('ORTALAMA İSKONTO', ''), '%')}",
            f"Ort. stok: {ozet_deger(row.get('ORTALAMA STOK', ''))}",
        ])
        rakip_ozeti = "\n".join([
            baslik_yap(row.get("EN UYGUN RAKİP FİRMA", "")) or "-",
            f"Stok: {baslik_yap(row.get('RAKİP STOK DURUMU', '')) or '-'}",
            f"Satış: {fiyat_deger(row.get('EN UYGUN SATIŞ FİYATI', ''))}",
            f"İsk.: {ozet_deger(row.get('EN UYGUN İSKONTO', ''), '%')}",
        ])
        secilen_iskonto = sayisal_deger(row.get("SEÇİLEN İSKONTO", ""))
        fiyat_hareketi = (
            f"Zam: {ozet_deger(abs(secilen_iskonto), '%')}"
            if secilen_iskonto is not None and secilen_iskonto < 0
            else f"Seçilen isk.: {ozet_deger(row.get('SEÇİLEN İSKONTO', ''), '%')}"
        )
        fiyatlandirma = "\n".join([
            f"Uygulanan kâr: {ozet_deger(row.get('KATEGORİ KÂR ORANI', ''), '%')}",
            f"Satış: {fiyat_deger(row.get('SATIŞ FİYATI', ''))}",
            f"Reel kâr: {ozet_deger(row.get('REEL KÂR ORANI', ''), '%')}",
            fiyat_hareketi,
        ])
        tsoft_liste_sayi = sayisal_deger(row.get("T-SOFT LİSTE FİYATI TL", ""))
        tsoft_satis_sayi = sayisal_deger(
            row.get("T-SOFT İNDİRİMLİ FİYATI TL", "")
        )
        tsoft_iskonto = None
        if (
            tsoft_liste_sayi is not None
            and tsoft_liste_sayi > 0
            and tsoft_satis_sayi is not None
            and tsoft_satis_sayi >= 0
        ):
            tsoft_iskonto = (1 - tsoft_satis_sayi / tsoft_liste_sayi) * 100
        tsoft_ozeti = "\n".join([
            f"Liste: {fiyat_deger(row.get('T-SOFT LİSTE FİYATI TL', ''))}",
            f"Satış: {fiyat_deger(row.get('T-SOFT İNDİRİMLİ FİYATI TL', ''))}",
            f"İsk.: {ozet_deger(tsoft_iskonto, '%')}",
            f"Stok: {ozet_deger(row.get('Stok', ''))}",
        ])
        rows.append({
            "ProductId": row.get("ProductId", ""),
            "Ürün": urun,
            "T-Soft Liste": tsoft_ozeti,
            "Aktiflik": row.get("T-SOFT AKTİFLİK", ""),
            "En İyi Alış": en_iyi_alis,
            "Alış Ortalaması": alis_ortalamasi,
            "En Uygun Rakip": rakip_ozeti,
            "Fiyatlandırma": fiyatlandirma,
            "Karar": "\n".join([
                str(row.get("ONAY", "")).strip(),
                str(row.get("AKTARIM DURUMU", "")).strip(),
            ]),
        })
    return pd.DataFrame(rows)


def fiyat_df_filtrele(query: dict) -> tuple[pd.DataFrame, int]:
    df = oku_fiyat_onay()
    kolonlar = [
        "ProductId",
        "Marka",
        "Kategori",
        "Ürün Web Servis Kodu",
        "Stok",
        "T-SOFT LİSTE FİYATI TL",
        "T-SOFT İNDİRİMLİ FİYATI TL",
        "T-SOFT AKTİFLİK",
        "EN İYİ ALIŞ B2B",
        "EN İYİ ALIŞ FİYATI",
        "ORTALAMA ALIŞ",
        "ORTALAMA STOK",
        "ORTALAMA İSKONTO",
        "EN UYGUN RAKİP FİRMA",
        "RAKİP STOK DURUMU",
        "EN UYGUN SATIŞ FİYATI",
        "EN UYGUN İSKONTO",
        "KATEGORİ KÂR ORANI",
        "SATIŞ FİYATI",
        "REEL KÂR ORANI",
        "HESAPLANAN İSKONTO",
        "SEÇİLEN İSKONTO",
        "ONAY",
        "AKTARIM DURUMU",
    ]
    kolonlar = [c for c in kolonlar if c in df.columns]
    df = df[kolonlar].copy()

    q = str(query.get("q", [""])[0]).strip().lower()
    brand = str(query.get("brand", [""])[0]).strip()
    category = str(query.get("category", [""])[0]).strip()
    activity = str(query.get("activity", [""])[0]).strip()
    status = str(query.get("status", [""])[0]).strip()
    decision = str(query.get("decision", [""])[0]).strip()
    applied_margin = str(query.get("applied_margin", [""])[0]).strip()
    competitor = str(query.get("competitor", [""])[0]).strip()
    if q:
        arama_kolonlari = [c for c in ["ProductId", "Marka", "Kategori", "Ürün Web Servis Kodu"] if c in df.columns]
        mask = pd.Series(False, index=df.index)
        for col in arama_kolonlari:
            mask = mask | df[col].astype(str).str.lower().str.contains(q, na=False)
        df = df[mask]
    if brand and "Marka" in df.columns:
        df = df[df["Marka"].astype(str) == brand]
    if category and "Kategori" in df.columns:
        df = df[df["Kategori"].astype(str) == category]
    if activity and "T-SOFT AKTİFLİK" in df.columns:
        df = df[df["T-SOFT AKTİFLİK"].astype(str) == activity]
    if status and "AKTARIM DURUMU" in df.columns:
        df = df[df["AKTARIM DURUMU"].astype(str) == status]
    if decision == "ZAM YAPILMALI" and "AKTARIM DURUMU" in df.columns:
        df = df[df["AKTARIM DURUMU"].astype(str).eq("ZAM YAPILMALI")]
    elif decision and "ONAY" in df.columns:
        df = df[df["ONAY"].astype(str).eq(decision)]
    if applied_margin and "KATEGORİ KÂR ORANI" in df.columns:
        secilen_kar = sayisal_deger(applied_margin)
        if secilen_kar is not None:
            kar_sayilari = pd.to_numeric(df["KATEGORİ KÂR ORANI"], errors="coerce")
            df = df[(kar_sayilari - secilen_kar).abs() < 0.0001]
    if competitor and "EN UYGUN RAKİP FİRMA" in df.columns:
        rakip_firmalar = df["EN UYGUN RAKİP FİRMA"].astype(str).str.strip()
        if competitor == "__NONE__":
            df = df[rakip_firmalar.eq("")]
        else:
            df = df[rakip_firmalar.eq(competitor)]

    toplam = len(df)
    limit_ham = str(query.get("limit", ["500"])[0] or "500").strip().lower()
    tumu = limit_ham in {"all", "tümü", "tumu"}
    if tumu:
        limit = max(toplam, 1)
    else:
        try:
            limit = int(limit_ham)
        except Exception:
            limit = 500
        if limit not in {100, 250, 500, 1000}:
            limit = 500
    try:
        page = int(query.get("page", ["1"])[0] or 1)
    except Exception:
        page = 1
    total_pages = max(1, (toplam + limit - 1) // limit)
    page = min(max(1, page), total_pages)
    offset = (page - 1) * limit
    gorunum = fiyat_gorunum_df(df.iloc[offset : offset + limit])
    return gorunum, toplam, page, "all" if tumu else limit, offset, total_pages


def product_id_metni(deger) -> str:
    if deger is None or pd.isna(deger):
        return ""
    if isinstance(deger, float) and deger.is_integer():
        return str(int(deger))
    return str(deger).strip()


def frontend_aktarim_kayitlari(secili_product_ids) -> list[dict]:
    secim = {product_id_metni(deger) for deger in secili_product_ids}
    secim.discard("")
    if not secim:
        raise ValueError("T-Soft aktarımı için ürün seçilmedi")

    df = oku_fiyat_onay()
    if "ProductId" not in df.columns:
        raise ValueError("Frontend verisinde ProductId kolonu yok")
    kimlikler = df["ProductId"].map(product_id_metni)
    secili = df[kimlikler.isin(secim)].copy()
    bulunan = set(kimlikler[kimlikler.isin(secim)].tolist())
    bulunamayan = sorted(secim - bulunan)
    if bulunamayan:
        raise ValueError(
            "Seçili ürün frontend verisinde bulunamadı: "
            + ", ".join(bulunamayan[:10])
        )

    aktiflik = secili.get("T-SOFT AKTİFLİK", "").astype(str).str.strip().str.upper()
    onay = secili.get("ONAY", "").astype(str).str.strip().str.upper()
    durum = secili.get("AKTARIM DURUMU", "").astype(str).str.strip().str.upper()
    uygun = onay.eq("EVET") & durum.isin({"AKTARIMA HAZIR", "ZAM YAPILMALI"}) & aktiflik.isin(
        {"AKTİF", "AKTIF", "1"}
    )
    if not uygun.all():
        kodlar = secili.loc[~uygun, "Ürün Web Servis Kodu"].astype(str).tolist()
        raise ValueError(
            "Seçili ürünlerden bazıları aktarıma hazır değil: "
            + ", ".join(kodlar[:10])
        )

    return [
        {str(kolon): temiz_deger(deger) for kolon, deger in row.items()}
        for row in secili.to_dict(orient="records")
    ]


def frontend_tum_aktarima_hazir_kayitlar() -> list[dict]:
    """Panelde EVET olan tum gecerli satirlari aktarim kontrolune hazirlar."""
    df = oku_fiyat_onay()
    aktiflik = df.get("T-SOFT AKTİFLİK", "").astype(str).str.strip().str.upper()
    onay = df.get("ONAY", "").astype(str).str.strip().str.upper()
    durum = df.get("AKTARIM DURUMU", "").astype(str).str.strip().str.upper()
    uygun = (
        onay.eq("EVET")
        & durum.isin({"AKTARIMA HAZIR", "ZAM YAPILMALI"})
        & aktiflik.isin({"AKTİF", "AKTIF", "1"})
    )
    secili = df[uygun].copy()
    return [
        {str(kolon): temiz_deger(deger) for kolon, deger in row.items()}
        for row in secili.to_dict(orient="records")
    ]


def json_response(handler: BaseHTTPRequestHandler, data, status=200):
    body = json.dumps(data, ensure_ascii=False, default=str).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
    handler.send_header("Pragma", "no-cache")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def html_response(handler: BaseHTTPRequestHandler):
    body = HTML.replace("__APP_VERSION__", APP_VERSION).encode("utf-8")
    handler.send_response(200)
    handler.send_header("Content-Type", "text/html; charset=utf-8")
    handler.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
    handler.send_header("Pragma", "no-cache")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def run_command(baslik: str, komut: list[str]):
    global SON_ISLEM
    with ISLEM_KILIDI:
        if SON_ISLEM.get("calisiyor"):
            return
        SON_ISLEM = {
            "calisiyor": True,
            "baslik": baslik,
            "cikti": "",
            "hata": "",
            "zaman": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
    try:
        proc = subprocess.run(
            komut,
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        SON_ISLEM["cikti"] = proc.stdout
        SON_ISLEM["hata"] = proc.stderr
        if proc.returncode != 0:
            SON_ISLEM["hata"] = (SON_ISLEM["hata"] + f"\nÇıkış kodu: {proc.returncode}").strip()
    except Exception as exc:
        SON_ISLEM["hata"] = str(exc)
    finally:
        SON_ISLEM["calisiyor"] = False


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):  # noqa: A002
        return

    def do_GET(self):  # noqa: N802
        try:
            parsed = urlparse(self.path)
            if parsed.path == "/":
                return html_response(self)
            if parsed.path == "/api/summary":
                df = oku_fiyat_onay()
                total = len(df)
                ready = int(df.get("AKTARIM DURUMU", "").isin(["AKTARIMA HAZIR", "ZAM YAPILMALI"]).sum()) if total else 0
                yes = int((df.get("ONAY", "") == "EVET").sum()) if total else 0
                no = int((df.get("ONAY", "") == "HAYIR").sum()) if total else 0
                passive = int((df.get("T-SOFT AKTİFLİK", "") == "PASİF").sum()) if total else 0
                missing = int((df.get("AKTARIM DURUMU", "") == "KATEGORİ KÂR ORANI YOK").sum()) if total else 0
                updated = datetime.fromtimestamp(FIYAT_ONAY_DOSYASI.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
                return json_response(self, {
                    "total": total,
                    "ready": ready,
                    "yes": yes,
                    "no": no,
                    "passive": passive,
                    "missingMargin": missing,
                    "updated": updated,
                })
            if parsed.path == "/api/filters":
                df = oku_fiyat_onay()
                def vals(col):
                    if col not in df.columns:
                        return []
                    return sorted(v for v in df[col].astype(str).str.strip().unique().tolist() if v)
                uygulanan_karlar = []
                if "KATEGORİ KÂR ORANI" in df.columns:
                    uygulanan_karlar = sorted({
                        float(deger)
                        for ham in df["KATEGORİ KÂR ORANI"].tolist()
                        if (deger := sayisal_deger(ham)) is not None
                    })
                    uygulanan_karlar = [
                        int(deger) if float(deger).is_integer() else deger
                        for deger in uygulanan_karlar
                    ]
                return json_response(self, {
                    "brands": vals("Marka"),
                    "categories": vals("Kategori"),
                    "activities": vals("T-SOFT AKTİFLİK"),
                    "approvals": vals("ONAY"),
                    "statuses": vals("AKTARIM DURUMU"),
                    "appliedMargins": uygulanan_karlar,
                    "competitors": vals("EN UYGUN RAKİP FİRMA"),
                })
            if parsed.path == "/api/pricing":
                df, filtered, page, limit, offset, total_pages = fiyat_df_filtrele(parse_qs(parsed.query))
                gorunen_kolonlar = [c for c in df.columns.tolist() if c != "ProductId"]
                return json_response(self, {
                    "columns": gorunen_kolonlar,
                    "rows": df_json_rows(df),
                    "filtered": filtered,
                    "page": page,
                    "limit": limit,
                    "offset": offset,
                    "totalPages": total_pages,
                })
            if parsed.path == "/api/categories":
                df = oku_kategoriler()
                return json_response(self, {"rows": df_json_rows(df)})
            if parsed.path == "/api/job":
                return json_response(self, SON_ISLEM)
            return json_response(self, {"error": "Bulunamadı"}, 404)
        except Exception as exc:
            return json_response(self, {"error": str(exc)}, 500)

    def do_POST(self):  # noqa: N802
        try:
            parsed = urlparse(self.path)
            if parsed.path == "/api/categories":
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                rows = payload.get("rows", [])
                df = pd.DataFrame(rows)
                if "Kategori" not in df.columns:
                    df["Kategori"] = ""
                if "Kar Orani" not in df.columns:
                    df["Kar Orani"] = ""
                df["Kategori"] = df["Kategori"].astype(str).str.strip()
                df["Kar Orani"] = df["Kar Orani"].astype(str).str.strip()
                df[["Kategori", "Kar Orani"]].to_excel(KAR_YEDEK_DOSYASI, index=False)
                dogrulama = pd.read_excel(KAR_YEDEK_DOSYASI, dtype=object, engine="openpyxl").fillna("")
                dolu_sayi = int(dogrulama.get("Kar Orani", pd.Series(dtype=object)).astype(str).str.strip().ne("").sum())
                try:
                    FRONTEND_CACHE_DOSYASI.unlink()
                except FileNotFoundError:
                    pass
                with VERI_KILIDI:
                    VERI_CACHE["kategori_mtime"] = None
                    VERI_CACHE["kategori_df"] = None
                    VERI_CACHE["fiyat_mtime"] = None
                    VERI_CACHE["fiyat_df"] = None
                return json_response(self, {"ok": True, "filled": dolu_sayi})
            if parsed.path == "/api/run":
                action = parse_qs(parsed.query).get("action", [""])[0]
                length = int(self.headers.get("Content-Length", "0"))
                payload = (
                    json.loads(self.rfile.read(length).decode("utf-8"))
                    if length
                    else {}
                )
                if SON_ISLEM.get("calisiyor"):
                    return json_response(self, {"message": "Zaten çalışan bir işlem var."})
                if action == "rapor":
                    title = "Fiyat raporu oluşturuluyor"
                    cmd = [sys.executable, str(BASE_DIR / "fiyatlandirma_raporu.py")]
                elif action == "kontrol":
                    kayitlar = frontend_tum_aktarima_hazir_kayitlar()
                    if not kayitlar:
                        raise ValueError("Frontend'de aktarıma hazır ürün yok")
                    CIKTI_DIR.mkdir(exist_ok=True)
                    FRONTEND_AKTARIM_DOSYASI.write_text(
                        json.dumps(kayitlar, ensure_ascii=False, default=str),
                        encoding="utf-8",
                    )
                    title = f"T-Soft aktarım kontrolü ({len(kayitlar)} ürün)"
                    cmd = [
                        sys.executable,
                        str(BASE_DIR / "tsoft_fiyat_aktar.py"),
                        "--frontend-json",
                        str(FRONTEND_AKTARIM_DOSYASI),
                    ]
                elif action == "gonder":
                    kayitlar = frontend_aktarim_kayitlari(
                        payload.get("productIds", [])
                    )
                    CIKTI_DIR.mkdir(exist_ok=True)
                    FRONTEND_AKTARIM_DOSYASI.write_text(
                        json.dumps(kayitlar, ensure_ascii=False, default=str),
                        encoding="utf-8",
                    )
                    title = f"T-Soft gerçek iskonto aktarımı ({len(kayitlar)} ürün)"
                    cmd = [
                        sys.executable,
                        str(BASE_DIR / "tsoft_fiyat_aktar.py"),
                        "--frontend-json",
                        str(FRONTEND_AKTARIM_DOSYASI),
                        "--gonder",
                    ]
                else:
                    return json_response(self, {"error": "Geçersiz işlem"}, 400)
                threading.Thread(target=run_command, args=(title, cmd), daemon=True).start()
                return json_response(self, {"message": f"{title} başlatıldı."})
            return json_response(self, {"error": "Bulunamadı"}, 404)
        except Exception as exc:
            return json_response(self, {"error": str(exc)}, 500)


def main():
    if os.environ.get("ENTES_FRONTEND_RELOADED") != "1":
        os.environ["ENTES_FRONTEND_RELOADED"] = "1"

    def auto_reload():
        mtimes = {
            dosya: dosya.stat().st_mtime
            for dosya in WATCHED_FILES
            if dosya.exists()
        }
        while True:
            time.sleep(2)
            for dosya, eski_mtime in list(mtimes.items()):
                if not dosya.exists():
                    continue
                yeni_mtime = dosya.stat().st_mtime
                if yeni_mtime != eski_mtime:
                    print(f"Dosya değişti, panel yeniden başlatılıyor: {dosya.name}")
                    os.execv(sys.executable, [sys.executable, *sys.argv])

    threading.Thread(target=auto_reload, daemon=True).start()

    server = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}"
    print(f"Fiyatlandırma paneli açılıyor: {url}")
    print("Kapatmak için terminalde Ctrl+C")
    print("Auto-reload aktif: frontend/kod dosyaları değişirse panel kendini yeniler.")
    try:
        webbrowser.open(url)
    except Exception:
        pass
    server.serve_forever()


if __name__ == "__main__":
    main()
