#!/usr/bin/env python3
"""
Rok-Popust Agent - проверува рокови на траење и применува скалиран попуст.
Cowork ја повикува main() директно - нема CLI аргументи, нема input().
"""

import json
import os
from datetime import date, datetime

import pandas as pd

# --- Конфигурација (сите вредности што треба да се менуваат се тука) ---
CONFIG = {
    "input_file": os.path.join(os.path.dirname(__file__), "..", "data", "products.xlsx"),
    "tracker_file": os.path.join(os.path.dirname(__file__), "..", "data", "TRACKER.md"),
    "dashboard_file": os.path.join(os.path.dirname(__file__), "..", "data", "dashboard.html"),
    "state_file": os.path.join(os.path.dirname(__file__), "..", "memory", "state.json"),
    "history_file": os.path.join(os.path.dirname(__file__), "..", "data", "history.csv"),
    # скала на попуст: (мин_денови, макс_денови, процент_попуст)
    "discount_tiers": [
        (11, 15, 10),
        (6, 10, 20),
        (0, 5, 40),
    ],
    # бонус попуст ако залихата е голема - производот треба побрзо да се раздаде
    # (мин_залиха, бонус_во_процентни_поени). Се применува само над веќе активиран попуст.
    "stock_boost_tiers": [
        (30, 10),
        (10, 5),
    ],
    # заштита на маржа: цената никогаш не смее да падне под Cost * (1 + min_margin_percent/100)
    "min_margin_percent": 10,
    # максимален вкупен попуст, без разлика на бустови
    "max_discount_percent": 60,
}


# --- Помошни функции ---

def load_products(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Не постои база на производи на патека: {path}")
    df = pd.read_excel(path)
    required_cols = {"Product", "SKU", "Price", "Expiry_Date"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Недостасуваат колони во базата: {missing}")
    return df


def days_to_expiry(expiry_date, today):
    if pd.isna(expiry_date):
        return None
    if isinstance(expiry_date, str):
        expiry_date = pd.to_datetime(expiry_date, errors="coerce")
        if pd.isna(expiry_date):
            return None
    if isinstance(expiry_date, pd.Timestamp):
        expiry_date = expiry_date.date()
    return (expiry_date - today).days


def discount_for_days(days, tiers):
    """Го враќа само базниот статус/тир според рокот (без залиха/маржа) - се користи за
    групирање/боење во dashboard-от."""
    if days is None:
        return None, "❌ невалиден датум"
    if days < 0:
        return 0, "ЗА ПОВЛЕКУВАЊЕ"
    for dmin, dmax, pct in tiers:
        if dmin <= days <= dmax:
            return pct, f"{pct}% попуст"
    return 0, "во ред"


def stock_boost_for(stock, boost_tiers):
    if stock is None or pd.isna(stock):
        return 0
    for min_stock, boost in sorted(boost_tiers, key=lambda t: -t[0]):
        if stock >= min_stock:
            return boost
    return 0


def process(df, config, today):
    tiers = config["discount_tiers"]
    boost_tiers = config["stock_boost_tiers"]
    min_margin = config["min_margin_percent"]
    max_pct = config["max_discount_percent"]
    has_cost = "Cost" in df.columns
    has_stock = "Stock" in df.columns

    days_list, base_pct_list, discount_list, new_price_list = [], [], [], []
    status_list, note_list = [], []

    for _, row in df.iterrows():
        try:
            days = days_to_expiry(row["Expiry_Date"], today)
            base_pct, status = discount_for_days(days, tiers)
            price = row["Price"]
            notes = []

            if base_pct is None:
                # невалиден датум
                days_list.append(None)
                base_pct_list.append(0)
                discount_list.append(0)
                new_price_list.append(price)
                status_list.append(status)
                note_list.append("")
                continue

            if base_pct == 0 or status == "ЗА ПОВЛЕКУВАЊЕ":
                # нема попуст (или е веќе истечено) - залиха/маржа логиката не важи
                days_list.append(days)
                base_pct_list.append(0)
                discount_list.append(0)
                new_price_list.append(price)
                status_list.append(status)
                note_list.append("")
                continue

            # бонус попуст поради голема залиха
            stock = row["Stock"] if has_stock else None
            boost = stock_boost_for(stock, boost_tiers)
            if boost:
                notes.append(f"+{boost}pp залиха ({int(stock)} парчиња)")

            wanted_pct = min(base_pct + boost, max_pct)
            desired_price = price * (1 - wanted_pct / 100)

            # заштита на маржа
            cost = row["Cost"] if has_cost else None
            if has_cost and not pd.isna(cost):
                floor_price = cost * (1 + min_margin / 100)
                if desired_price < floor_price:
                    new_price = round(max(floor_price, 0), 2)
                    notes.append("заштитена маржа")
                else:
                    new_price = round(desired_price, 2)
            else:
                new_price = round(desired_price, 2)

            actual_pct = round((price - new_price) / price * 100, 1) if price else 0

            days_list.append(days)
            base_pct_list.append(base_pct)
            discount_list.append(actual_pct)
            new_price_list.append(new_price)
            status_list.append(status)
            note_list.append("; ".join(notes))
        except Exception as e:
            days_list.append(None)
            base_pct_list.append(0)
            discount_list.append(0)
            new_price_list.append(row.get("Price"))
            status_list.append(f"❌ грешка: {e}")
            note_list.append("")

    df = df.copy()
    df["Days_To_Expiry"] = days_list
    df["Base_Discount_Percent"] = base_pct_list
    df["Discount_Percent"] = discount_list
    df["New_Price"] = new_price_list
    df["Status"] = status_list
    df["Note"] = note_list
    return df


def write_tracker(df, path, today):
    affected = df[df["Discount_Percent"] > 0]
    expired = df[df["Status"] == "ЗА ПОВЛЕКУВАЊЕ"]
    errors = df[df["Status"].astype(str).str.startswith("❌")]
    total_discount_value = ((df["Price"] - df["New_Price"]).clip(lower=0)).sum()

    margin_protected = df[df["Note"].astype(str).str.contains("маржа", na=False)]
    avg_discount = affected["Discount_Percent"].mean() if len(affected) else 0

    lines = ["# Tracker - Rok-Popust Agent", "", f"Датум на проверка: {today.isoformat()}", ""]
    lines.append("| Product | SKU | Дена до истек | Стара цена | Попуст | Нова цена | Статус | Забелешка |")
    lines.append("|---|---|---|---|---|---|---|---|")
    for _, row in df.sort_values("Days_To_Expiry", na_position="last").iterrows():
        lines.append(
            f"| {row['Product']} | {row['SKU']} | {row['Days_To_Expiry']} | "
            f"{row['Price']} | {row['Discount_Percent']}% | {row['New_Price']} | {row['Status']} | {row['Note']} |"
        )
    lines.append("")
    lines.append("## Summary")
    lines.append(f"- Вкупно производи: {len(df)}")
    lines.append(f"- Засегнати со попуст: {len(affected)}")
    lines.append(f"- За повлекување (истечено): {len(expired)}")
    lines.append(f"- Заштитена маржа кај: {len(margin_protected)} производи")
    lines.append(f"- Грешки: {len(errors)}")
    lines.append(f"- Вкупна вредност на попуст: {total_discount_value:.2f}")
    lines.append(f"- Просечен применет попуст: {avg_discount:.1f}%")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return {
        "total": len(df),
        "affected": len(affected),
        "expired": len(expired),
        "errors": len(errors),
        "margin_protected": len(margin_protected),
        "total_discount_value": round(float(total_discount_value), 2),
        "avg_discount_percent": round(float(avg_discount), 1),
    }


STATUS_COLORS = {
    "во ред": ("#0ca30c", "good"),
    "10% попуст": ("#fab219", "warning"),
    "20% попуст": ("#ec835a", "serious"),
    "40% попуст": ("#d03b3b", "critical"),
    "ЗА ПОВЛЕКУВАЊЕ": ("#d03b3b", "critical"),
}


def status_badge(status):
    hex_color, role = STATUS_COLORS.get(status, ("#898781", "muted"))
    icon = {"good": "✓", "warning": "!", "serious": "!", "critical": "✕", "muted": "?"}[role]
    return hex_color, role, icon


def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


LEVEL_ORDER = ["во ред", "10% попуст", "20% попуст", "40% попуст", "ЗА ПОВЛЕКУВАЊЕ"]

CATEGORICAL = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7"]

STATUS_ROLE_VAR = {
    "во ред": "var(--good)",
    "10% попуст": "var(--warning)",
    "20% попуст": "var(--serious)",
    "40% попуст": "var(--critical)",
    "ЗА ПОВЛЕКУВАЊЕ": "var(--critical)",
}


def write_dashboard(df, path, today, summary, history_path=None):
    df = df.copy()
    df["Days_To_Expiry_JS"] = df["Days_To_Expiry"].apply(lambda v: None if pd.isna(v) else int(v))
    df["Expiry_Date_JS"] = df["Expiry_Date"].apply(
        lambda v: pd.to_datetime(v).strftime("%Y-%m-%d") if not pd.isna(v) else None
    )

    products = []
    for _, row in df.iterrows():
        products.append({
            "product": row["Product"],
            "sku": row["SKU"],
            "category": row.get("Category", ""),
            "price": float(row["Price"]),
            "cost": float(row["Cost"]) if "Cost" in df.columns and not pd.isna(row.get("Cost")) else None,
            "stock": int(row["Stock"]) if "Stock" in df.columns and not pd.isna(row.get("Stock")) else None,
            "expiryDate": row["Expiry_Date_JS"],
            "daysToExpiry": row["Days_To_Expiry_JS"],
            "discountPercent": float(row["Discount_Percent"]),
            "newPrice": float(row["New_Price"]),
            "status": row["Status"],
            "note": row.get("Note", "") or "",
        })

    history = []
    if history_path and os.path.exists(history_path):
        try:
            hist_df = pd.read_csv(history_path)
            for _, hrow in hist_df.iterrows():
                history.append({
                    "date": str(hrow["date"]),
                    "timestamp": str(hrow["timestamp"]),
                    "total": int(hrow["total"]),
                    "affected": int(hrow["affected"]),
                    "expired": int(hrow["expired"]),
                    "totalDiscountValue": float(hrow["total_discount_value"]),
                    "avgDiscountPercent": float(hrow.get("avg_discount_percent", 0) or 0),
                })
        except Exception:
            history = []

    data_json = json.dumps({
        "products": products,
        "history": history,
        "lastUpdated": today.isoformat(),
        "levelOrder": LEVEL_ORDER,
        "categorical": CATEGORICAL,
    }, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="mk">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Rok-Popust Dashboard</title>
<style>
  :root {{
    color-scheme: light;
    --surface-1:      #fcfcfb;
    --page-plane:      #f9f9f7;
    --text-primary:   #0b0b0b;
    --text-secondary: #52514e;
    --text-muted:     #898781;
    --gridline:       #e1e0d9;
    --baseline:       #c3c2b7;
    --border:         rgba(11,11,11,0.10);
    --series-1:       #2a78d6;
    --good:           #0ca30c;
    --warning:        #fab219;
    --serious:        #ec835a;
    --critical:       #d03b3b;
  }}
  @media (prefers-color-scheme: dark) {{
    :root:where(:not([data-theme="light"])) {{
      color-scheme: dark;
      --surface-1:      #1a1a19;
      --page-plane:      #0d0d0d;
      --text-primary:   #ffffff;
      --text-secondary: #c3c2b7;
      --text-muted:     #898781;
      --gridline:       #2c2c2a;
      --baseline:       #383835;
      --border:         rgba(255,255,255,0.10);
      --series-1:       #3987e5;
      --good:           #0ca30c;
      --warning:        #fab219;
      --serious:        #ec835a;
      --critical:       #e66767;
    }}
  }}
  :root[data-theme="dark"] {{
    color-scheme: dark;
    --surface-1:      #1a1a19;
    --page-plane:      #0d0d0d;
    --text-primary:   #ffffff;
    --text-secondary: #c3c2b7;
    --text-muted:     #898781;
    --gridline:       #2c2c2a;
    --baseline:       #383835;
    --border:         rgba(255,255,255,0.10);
    --series-1:       #3987e5;
    --good:           #0ca30c;
    --warning:        #fab219;
    --serious:        #ec835a;
    --critical:       #e66767;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0;
    font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
    background: var(--page-plane);
    color: var(--text-primary);
    padding: 32px 24px 48px;
  }}
  .wrap {{ max-width: 1080px; margin: 0 auto; position: relative; }}
  .brand-row {{ display: flex; align-items: center; gap: 10px; margin-bottom: 6px; }}
  .brand-dot {{ width: 10px; height: 10px; border-radius: 3px; background: var(--series-1); flex: none; }}
  .brand-name {{ font-size: 12px; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase; color: var(--text-muted); }}
  header {{
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    margin-bottom: 20px;
    gap: 16px;
    flex-wrap: wrap;
  }}
  h1 {{ font-size: 25px; margin: 0 0 5px; font-weight: 650; letter-spacing: -0.01em; }}
  .subtitle {{ color: var(--text-secondary); font-size: 13.5px; }}
  .theme-toggle {{
    border: 1px solid var(--border);
    background: var(--surface-1);
    color: var(--text-secondary);
    border-radius: 8px;
    padding: 6px 12px;
    font-size: 13px;
    cursor: pointer;
  }}

  /* page nav */
  .page-nav {{ display: flex; gap: 8px; margin-bottom: 18px; border-bottom: 1px solid var(--border); }}
  .page-nav-btn {{
    border: none;
    background: none;
    color: var(--text-secondary);
    padding: 10px 4px;
    margin-right: 14px;
    font-size: 14px;
    font-weight: 600;
    cursor: pointer;
    border-bottom: 2px solid transparent;
  }}
  .page-nav-btn.active {{ color: var(--text-primary); border-bottom-color: var(--series-1); }}

  /* price stickers */
  .sticker-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(190px, 1fr));
    gap: 16px;
  }}
  .sticker-card {{
    position: relative;
    background: var(--surface-1);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 16px 16px 14px;
    overflow: hidden;
  }}
  .sticker-badge {{
    position: absolute;
    top: 12px;
    right: 12px;
    background: var(--critical);
    color: #fff;
    font-size: 12.5px;
    font-weight: 700;
    padding: 4px 10px;
    border-radius: 999px;
  }}
  .sticker-name {{
    font-size: 13px;
    font-weight: 600;
    line-height: 1.35;
    margin: 0 46px 14px 0;
    min-height: 35px;
  }}
  .sticker-old-price {{ font-size: 12.5px; color: var(--text-muted); text-decoration: line-through; }}
  .sticker-new-price {{ font-size: 26px; font-weight: 750; color: var(--critical); line-height: 1.15; }}
  .sticker-new-price .unit {{ font-size: 14px; font-weight: 600; margin-left: 3px; }}
  .sticker-footer {{
    margin-top: 12px;
    padding-top: 10px;
    border-top: 1px dashed var(--border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
  }}
  .sticker-barcode {{
    flex: 1;
    height: 22px;
    background: repeating-linear-gradient(
      90deg,
      var(--text-primary) 0px, var(--text-primary) 2px,
      transparent 2px, transparent 4px,
      var(--text-primary) 4px, var(--text-primary) 5px,
      transparent 5px, transparent 8px,
      var(--text-primary) 8px, var(--text-primary) 11px,
      transparent 11px, transparent 13px
    );
    opacity: 0.85;
  }}
  .sticker-sku {{ font-size: 10.5px; color: var(--text-muted); font-family: ui-monospace, "SF Mono", Consolas, monospace; white-space: nowrap; }}
  .sticker-empty {{ color: var(--text-muted); font-size: 13.5px; padding: 24px 0; text-align: center; }}

  .kpi-row {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
    gap: 12px;
    margin-bottom: 20px;
  }}
  .tile {{
    background: var(--surface-1);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 16px 18px;
    display: flex;
    flex-direction: column;
    gap: 2px;
  }}
  .tile-top {{ display: flex; align-items: center; justify-content: space-between; }}
  .tile .label {{
    font-size: 12px;
    color: var(--text-muted);
    margin-bottom: 6px;
  }}
  .tile .value {{
    font-size: 28px;
    font-weight: 650;
    font-variant-numeric: proportional-nums;
  }}
  .tile .value.critical {{ color: var(--critical); }}
  .tile .foot {{ font-size: 12px; color: var(--text-secondary); margin-top: 4px; display: flex; align-items: center; gap: 6px; }}
  .tile .delta {{ font-weight: 600; }}
  .tile .delta.up {{ color: var(--good); }}
  .tile .delta.down {{ color: var(--critical); }}
  .icon-chip {{
    width: 30px; height: 30px;
    border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    flex: none;
  }}
  .icon-chip svg {{ width: 16px; height: 16px; }}
  .panel {{
    background: var(--surface-1);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 22px 22px 20px;
    margin-bottom: 16px;
  }}
  .panel-head {{ display: flex; align-items: baseline; justify-content: space-between; gap: 12px; flex-wrap: wrap; margin-bottom: 2px; }}
  .panel h2 {{
    font-size: 15.5px;
    margin: 0 0 4px;
    font-weight: 650;
  }}
  .panel .panel-sub {{ font-size: 12px; color: var(--text-muted); margin-bottom: 18px; }}
  .legend-row {{ display: flex; gap: 14px; flex-wrap: wrap; font-size: 11.5px; color: var(--text-secondary); margin-bottom: 14px; }}
  .legend-item {{ display: inline-flex; align-items: center; gap: 6px; }}
  .legend-swatch {{ width: 9px; height: 9px; border-radius: 2px; flex: none; }}
  .tabs {{ display: flex; gap: 8px; margin: 16px 0 14px; flex-wrap: wrap; }}
  .tab-btn {{
    border: 1px solid var(--border);
    background: var(--page-plane);
    color: var(--text-secondary);
    border-radius: 999px;
    padding: 6px 14px;
    font-size: 12.5px;
    font-weight: 600;
    cursor: pointer;
  }}
  .tab-btn.active {{ background: var(--text-primary); color: var(--surface-1); border-color: var(--text-primary); }}
  .table-scroll {{ overflow-x: auto; -webkit-overflow-scrolling: touch; margin: 0 -4px; padding: 0 4px; }}
  table {{ width: 100%; min-width: 640px; border-collapse: collapse; font-size: 13px; }}
  th, td {{ text-align: left; padding: 9px 10px; border-bottom: 1px solid var(--gridline); }}
  th {{ color: var(--text-muted); font-weight: 500; font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.02em; }}
  td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
  .muted {{ color: var(--text-secondary); }}
  tr.urgent td {{ background: color-mix(in srgb, var(--critical) 7%, transparent); }}
  tr td:first-child {{ font-weight: 500; }}
  .badge {{
    display: inline-flex;
    align-items: center;
    gap: 4px;
    font-size: 11px;
    color: var(--critical);
    font-weight: 600;
  }}
  .cat-badge {{
    display: inline-block;
    font-size: 11px;
    font-weight: 600;
    padding: 3px 9px;
    border-radius: 999px;
    color: white;
  }}
  footer {{
    font-size: 12px;
    color: var(--text-muted);
    margin-top: 16px;
    line-height: 1.6;
  }}
  /* ---- responsive (телефон/таблет) ---- */
  @media (max-width: 720px) {{
    body {{ padding: 20px 14px 36px; }}
    h1 {{ font-size: 20px; }}
    .subtitle {{ font-size: 12.5px; }}
    header {{ align-items: center; }}
    .theme-toggle {{ padding: 6px 10px; font-size: 12.5px; }}

    .kpi-row {{ grid-template-columns: repeat(2, 1fr); gap: 10px; }}
    .tile {{ padding: 13px 14px; }}
    .tile .value {{ font-size: 23px; }}

    .panel {{ padding: 16px 14px 14px; border-radius: 12px; }}
    .panel h2 {{ font-size: 14.5px; }}
    .tab-btn {{ padding: 5px 11px; font-size: 12px; }}
  }}

  @media (max-width: 420px) {{
    .kpi-row {{ grid-template-columns: 1fr 1fr; }}
  }}
</style>
</head>
<body>
<div class="wrap">
  <div class="brand-row">
    <span class="brand-dot"></span>
    <span class="brand-name">Rok-Popust Agent</span>
  </div>
  <header>
    <div>
      <h1>Автоматска заштита на маржа пред истек на рок</h1>
      <div class="subtitle" id="subtitle"></div>
    </div>
    <button class="theme-toggle" id="themeToggle">🌓 Смени тема</button>
  </header>

  <div class="page-nav" id="pageNav">
    <button class="page-nav-btn active" data-page="overview">Преглед</button>
    <button class="page-nav-btn" data-page="stickers">Ценовни стикери</button>
  </div>

  <div class="page" id="page-overview">
    <div class="kpi-row" id="kpiRow"></div>

    <div class="panel">
      <h2>Производи</h2>
      <div class="panel-sub">Најблизок рок горе. Процентот во значката е реално применетиот попуст (по бонус за залиха и заштита на маржа).</div>
      <div class="tabs" id="tableTabs">
        <button class="tab-btn active" data-tab="all">Сите</button>
        <button class="tab-btn" data-tab="discount">Со попуст</button>
        <button class="tab-btn" data-tab="expired">За повлекување</button>
      </div>
      <div class="table-scroll">
        <table id="prodTable">
          <thead>
            <tr>
              <th>Производ</th>
              <th>SKU</th>
              <th class="num">Дена до истек</th>
              <th class="num">Залиха</th>
              <th class="num">Стара цена</th>
              <th class="num">Нова цена</th>
              <th></th>
            </tr>
          </thead>
          <tbody id="prodBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <div class="page" id="page-stickers" style="display:none">
    <div class="panel">
      <h2>Ценовни стикери</h2>
      <div class="panel-sub" id="stickersSub">Автоматски генерирани стикери за производите со активен попуст - за лепење на рафт.</div>
      <div class="sticker-grid" id="stickerGrid"></div>
    </div>
  </div>

  <footer id="footerNote"></footer>
</div>

<script>
const DATA = {data_json};

const STATUS_COLOR_LIGHT = {{
  "во ред": "#0ca30c",
  "10% попуст": "#fab219",
  "20% попуст": "#ec835a",
  "40% попуст": "#d03b3b",
  "ЗА ПОВЛЕКУВАЊЕ": "#d03b3b"
}};
function fmt(n) {{
  return n.toLocaleString("mk-MK", {{ minimumFractionDigits: 2, maximumFractionDigits: 2 }});
}}

const products = [...DATA.products].sort((a, b) => {{
  const da = a.daysToExpiry === null ? Infinity : a.daysToExpiry;
  const db = b.daysToExpiry === null ? Infinity : b.daysToExpiry;
  return da - db;
}});

const total = products.length;
const affected = products.filter(p => p.discountPercent > 0).length;
const expired = products.filter(p => p.status === "ЗА ПОВЛЕКУВАЊЕ").length;

// ---- header ----
document.getElementById("subtitle").textContent =
  `${{total}} производи во базата · последно ажурирано ${{new Date(DATA.lastUpdated).toLocaleDateString("mk-MK")}}`;
document.getElementById("footerNote").innerHTML =
  `Извор: data/products.xlsx. Скала на попуст: 15-11 дена = 10%, 10-6 дена = 20%, 5-0 дена = 40%, ` +
  `со бонус за голема залиха и заштита на минимална маржа. ` +
  `Истечени производи се означуваат "ЗА ПОВЛЕКУВАЊЕ" и не се бришат автоматски.`;

const marginProtected = products.filter(p => (p.note || "").includes("маржа")).length;
const hist = DATA.history || [];
const prevRun = hist.length >= 2 ? hist[hist.length - 2] : null;

function deltaHtml(current, previous, higherIsBetter) {{
  if (previous === null || previous === undefined) return "";
  const diff = current - previous;
  if (Math.abs(diff) < 0.005) return `<span class="delta" style="color:var(--text-muted)">· без промена</span>`;
  const up = diff > 0;
  const good = higherIsBetter ? up : !up;
  const arrow = up ? "▲" : "▼";
  const cls = good ? "up" : "down";
  const shown = Number.isInteger(current) && Number.isInteger(previous) ? Math.abs(diff) : fmt(Math.abs(diff));
  return `<span class="delta ${{cls}}">${{arrow}} ${{shown}}</span> <span class="muted">vs минато пуштање</span>`;
}}

// ---- KPI tiles ----
const ICONS = {{
  box: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/></svg>',
  tag: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20.59 13.41L11 3.83A2 2 0 009.59 3H4a1 1 0 00-1 1v5.59a2 2 0 00.59 1.41l9.58 9.59a2 2 0 002.83 0l4.59-4.59a2 2 0 000-2.83z"/><circle cx="7.5" cy="7.5" r="1.5"/></svg>',
  alert: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><path d="M12 9v4"/><path d="M12 17h.01"/></svg>',
  shield: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="M9 12l2 2 4-4"/></svg>',
}};
const kpiRow = document.getElementById("kpiRow");
const kpis = [
  {{ label: "Вкупно производи", value: String(total), foot: "во базата", icon: "box", iconColor: "var(--series-1)" }},
  {{ label: "Со нов попуст", value: String(affected), foot: deltaHtml(affected, prevRun ? prevRun.affected : null, false), icon: "tag", iconColor: "var(--warning)" }},
  {{ label: "За повлекување", value: String(expired), foot: deltaHtml(expired, prevRun ? prevRun.expired : null, false) || (expired ? "потребна проверка" : "нема истечени"), critical: expired > 0, icon: "alert", iconColor: "var(--critical)" }},
  {{ label: "Заштитена маржа кај", value: String(marginProtected), foot: "производи не паднаа под минимална маржа", icon: "shield", iconColor: "var(--good)" }},
];
kpiRow.innerHTML = kpis.map(k => `
  <div class="tile">
    <div class="tile-top">
      <div class="label">${{k.label}}</div>
      <div class="icon-chip" style="background:color-mix(in srgb, ${{k.iconColor}} 16%, transparent); color:${{k.iconColor}}">${{ICONS[k.icon]}}</div>
    </div>
    <div class="value${{k.critical ? " critical" : ""}}">${{k.value}}</div>
    <div class="foot">${{k.foot}}</div>
  </div>
`).join("");

// ---- table (со табови: Сите / Со попуст / За повлекување) ----
const prodBody = document.getElementById("prodBody");
const TABLE_FILTERS = {{
  all: p => true,
  discount: p => p.discountPercent > 0,
  expired: p => p.status === "ЗА ПОВЛЕКУВАЊЕ",
}};

function renderProductRows(tab) {{
  const rows = products.filter(TABLE_FILTERS[tab]);
  if (!rows.length) {{
    prodBody.innerHTML = `<tr><td colspan="7" class="muted" style="text-align:center;padding:20px">Нема производи во оваа категорија</td></tr>`;
    return;
  }}
  prodBody.innerHTML = rows.map(p => {{
    const color = STATUS_COLOR_LIGHT[p.status] || "#898781";
    const urgent = p.status === "ЗА ПОВЛЕКУВАЊЕ" || p.status === "40% попуст";
    const days = p.daysToExpiry === null ? "-" : p.daysToExpiry;
    const stock = p.stock === null ? "-" : p.stock;
    // значката покажува реален пресметан попуст (не тиерот од скалата)
    const badgeText = p.discountPercent > 0 ? `${{p.discountPercent}}% попуст` : p.status;
    return `
      <tr class="${{urgent ? "urgent" : ""}}">
        <td>${{p.product}}</td>
        <td>${{p.sku}}</td>
        <td class="num">${{days}}</td>
        <td class="num">${{stock}}</td>
        <td class="num">${{fmt(p.price)}}</td>
        <td class="num">${{fmt(p.newPrice)}}</td>
        <td><span class="cat-badge" style="background:${{color}}">${{badgeText}}</span></td>
      </tr>`;
  }}).join("");
}}

const tabCounts = {{
  all: products.length,
  discount: products.filter(TABLE_FILTERS.discount).length,
  expired: products.filter(TABLE_FILTERS.expired).length,
}};
document.querySelectorAll(".tab-btn").forEach(btn => {{
  const tab = btn.dataset.tab;
  const baseLabel = btn.textContent;
  btn.textContent = `${{baseLabel}} (${{tabCounts[tab]}})`;
  btn.addEventListener("click", () => {{
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    renderProductRows(tab);
  }});
}});
renderProductRows("all");

// ---- ценовни стикери (визуелно, за pitch) ----
(function renderStickers() {{
  const grid = document.getElementById("stickerGrid");
  const discounted = products.filter(p => p.discountPercent > 0);
  document.getElementById("stickersSub").textContent =
    `Автоматски генерирани стикери за производите со активен попуст - за лепење на рафт (${{discounted.length}} стикери).`;
  if (!discounted.length) {{
    grid.innerHTML = `<div class="sticker-empty">Нема производи со активен попуст во оваа проверка.</div>`;
    return;
  }}
  grid.innerHTML = discounted.map(p => `
    <div class="sticker-card">
      <span class="sticker-badge">-${{p.discountPercent}}%</span>
      <div class="sticker-name">${{p.product}}</div>
      <div class="sticker-old-price">${{fmt(p.price)}} ден.</div>
      <div class="sticker-new-price">${{fmt(p.newPrice)}}<span class="unit">ден.</span></div>
      <div class="sticker-footer">
        <div class="sticker-barcode"></div>
        <div class="sticker-sku">${{p.sku}}</div>
      </div>
    </div>
  `).join("");
}})();

// ---- навигација меѓу страници (Преглед / Ценовни стикери) ----
document.querySelectorAll(".page-nav-btn").forEach(btn => {{
  btn.addEventListener("click", () => {{
    const page = btn.dataset.page;
    document.querySelectorAll(".page-nav-btn").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("page-overview").style.display = page === "overview" ? "" : "none";
    document.getElementById("page-stickers").style.display = page === "stickers" ? "" : "none";
  }});
}});

// ---- theme toggle ----
const toggle = document.getElementById("themeToggle");
toggle.addEventListener("click", () => {{
  const root = document.documentElement;
  const current = root.getAttribute("data-theme");
  root.setAttribute("data-theme", current === "dark" ? "light" : "dark");
}});
</script>
</body>
</html>
"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)


def update_state(state_path, today, summary, expired_products):
    if os.path.exists(state_path):
        with open(state_path, "r", encoding="utf-8") as f:
            state = json.load(f)
    else:
        state = {}
    state["last_run"] = today.isoformat()
    state["stats"] = {
        "total_processed": summary["total"],
        "total_discount_value": summary["total_discount_value"],
        "last_result": f"{summary['affected']} засегнати, {summary['expired']} за повлекување",
    }
    state["pending"] = expired_products
    with open(state_path, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def write_history(history_path, today, summary):
    """Додава нов ред за секое поединечно пуштање - не презапишува претходни редови."""
    row = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "date": today.isoformat(),
        "total": summary["total"],
        "affected": summary["affected"],
        "expired": summary["expired"],
        "margin_protected": summary.get("margin_protected", 0),
        "total_discount_value": summary["total_discount_value"],
        "avg_discount_percent": summary.get("avg_discount_percent", 0),
    }
    file_exists = os.path.exists(history_path)
    df_row = pd.DataFrame([row])
    if file_exists:
        df_row.to_csv(history_path, mode="a", header=False, index=False)
    else:
        df_row.to_csv(history_path, mode="w", header=True, index=False)


# --- Главна функција ---

def main():
    print("Стартувам проверка на рокови...")
    today = date.today()
    try:
        df = load_products(CONFIG["input_file"])
    except Exception as e:
        print(f"Грешка при читање база: {e}")
        return

    df = process(df, CONFIG, today)

    try:
        df.to_excel(CONFIG["input_file"], index=False)
        print(f"Ажурирана база: {CONFIG['input_file']}")
    except Exception as e:
        print(f"Грешка при зачувување база: {e}")

    try:
        summary = write_tracker(df, CONFIG["tracker_file"], today)
        print(f"Извештај: {CONFIG['tracker_file']}")
    except Exception as e:
        print(f"Грешка при пишување tracker: {e}")
        return

    try:
        write_history(CONFIG["history_file"], today, summary)
        print(f"Историја ажурирана: {CONFIG['history_file']}")
    except Exception as e:
        print(f"Грешка при запишување историја: {e}")

    try:
        write_dashboard(df, CONFIG["dashboard_file"], today, summary, CONFIG["history_file"])
        print(f"Dashboard: {CONFIG['dashboard_file']}")
    except Exception as e:
        print(f"Грешка при генерирање dashboard: {e}")

    expired_products = df[df["Status"] == "ЗА ПОВЛЕКУВАЊЕ"]["Product"].tolist()
    try:
        update_state(CONFIG["state_file"], today, summary, expired_products)
        print(f"Меморија ажурирана: {CONFIG['state_file']}")
    except Exception as e:
        print(f"Грешка при ажурирање меморија: {e}")

    print(
        f"Готово. Вкупно: {summary['total']}, засегнати: {summary['affected']}, "
        f"за повлекување: {summary['expired']}, вкупен попуст: {summary['total_discount_value']:.2f}"
    )


if __name__ == "__main__":
    main()
