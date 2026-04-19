from __future__ import annotations

import json
import math
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl import load_workbook

SCRIPT_PATH = Path(__file__).resolve()
DATA_API_DIR = SCRIPT_PATH.parent
REPO_ROOT = DATA_API_DIR.parent
OUT_DIR = DATA_API_DIR / 'out'


def resolve_default_xlsx() -> Path:
    """Find the workbook in the most likely project locations."""
    candidates = [
        REPO_ROOT / 'Cubo_Semanal_Compactado.xlsx',
        DATA_API_DIR / 'Cubo_Semanal_Compactado.xlsx',
        Path.cwd() / 'Cubo_Semanal_Compactado.xlsx',
        Path.cwd() / 'data_api' / 'Cubo_Semanal_Compactado.xlsx',
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


DEFAULT_XLSX = resolve_default_xlsx()

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

def safe_float(v: Any) -> float:
    if v is None or v == '':
        return 0.0
    try:
        return float(v)
    except Exception:
        return 0.0


def safe_int(v: Any) -> int:
    if v is None or v == '':
        return 0
    try:
        return int(round(float(v)))
    except Exception:
        return 0


def safe_str(v: Any) -> str:
    if v is None:
        return ''
    if isinstance(v, datetime):
        return v.isoformat()
    return str(v).strip()


def norm_text(v: Any) -> str:
    s = safe_str(v).upper()
    return ' '.join(s.split())


def key_text(v: Any) -> str:
    return ''.join(ch for ch in norm_text(v) if ch.isalnum())


def is_finite_number(v: Any) -> bool:
    try:
        return math.isfinite(float(v))
    except Exception:
        return False


def month_prev(month_key: str) -> str:
    if not month_key or '-' not in month_key:
        return ''
    y, m = month_key.split('-')
    y_i = int(y)
    m_i = int(m)
    if m_i == 1:
        return f"{y_i - 1}-12"
    return f"{y_i}-{m_i - 1:02d}"


def month_ly(month_key: str) -> str:
    if not month_key or '-' not in month_key:
        return ''
    y, m = month_key.split('-')
    return f"{int(y) - 1}-{m}"


def week_ly(week_key: str) -> str:
    # Expects YYYY-Www
    if not week_key or '-W' not in week_key:
        return ''
    y, w = week_key.split('-W')
    return f"{int(y) - 1}-W{int(w):02d}"


def month_label(month_key: str) -> str:
    names = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
    if not month_key or '-' not in month_key:
        return month_key
    y, m = month_key.split('-')
    idx = max(1, min(12, int(m))) - 1
    return f"{names[idx]} {y}"


def ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def dump_json(path: Path, data: Any) -> None:
    ensure_parent(path)
    with path.open('w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)


def week_sort_key(week_key: str) -> Tuple[int, int]:
    if '-W' not in week_key:
        return (0, 0)
    y, w = week_key.split('-W')
    return int(y), int(w)


def severity_for_amount(v: float) -> str:
    av = abs(v)
    if av >= 500:
        return 'high'
    if av >= 150:
        return 'medium'
    return 'low'


def priority_for_amount(v: float) -> str:
    av = abs(v)
    if av >= 500:
        return 'P1'
    if av >= 150:
        return 'P2'
    if av >= 50:
        return 'P3'
    return 'P4'


@dataclass
class Metrics:
    sales: float = 0.0
    margin: float = 0.0
    qty: float = 0.0
    tx: float = 0.0

    def add(self, sales: float = 0.0, margin: float = 0.0, qty: float = 0.0, tx: float = 0.0) -> None:
        self.sales += safe_float(sales)
        self.margin += safe_float(margin)
        self.qty += safe_float(qty)
        self.tx += safe_float(tx)

    def to_sales_payload(self) -> Dict[str, Any]:
        avg_ticket = self.sales / self.tx if self.tx else 0.0
        margin_pct = (self.margin / self.sales) * 100 if self.sales else 0.0
        return {
            'sales': round(self.sales, 2),
            'margin': round(self.margin, 2),
            'margin_pct': round(margin_pct, 2),
            'qty': round(self.qty, 2),
            'tx': safe_int(self.tx),
            'avg_ticket': round(avg_ticket, 2),
        }


class Exporter:
    def __init__(self, xlsx_path: Path, out_dir: Path):
        self.xlsx_path = xlsx_path
        self.out_dir = out_dir
        self.wb = load_workbook(self.xlsx_path, read_only=True, data_only=True)

        self.store_names: set[str] = set()
        self.months: set[str] = set()
        self.weeks: set[str] = set()
        self.sales_catalogs = {
            'departments': set(),
            'categories': set(),
            'brands': set(),
            'suppliers': set(),
            'descriptions': set(),
            'sbf_services': set(),
        }

        # Sales + delivery aggregates
        self.sales_month_business: dict[tuple, Metrics] = defaultdict(Metrics)
        self.sales_week_business: dict[tuple, Metrics] = defaultdict(Metrics)
        self.sales_month_store: dict[tuple, Metrics] = defaultdict(Metrics)
        self.sales_week_store: dict[tuple, Metrics] = defaultdict(Metrics)
        self.sales_month_dept: dict[tuple, Metrics] = defaultdict(Metrics)
        self.sales_week_dept: dict[tuple, Metrics] = defaultdict(Metrics)
        self.sales_month_cat: dict[tuple, Metrics] = defaultdict(Metrics)
        self.sales_week_cat: dict[tuple, Metrics] = defaultdict(Metrics)
        self.sales_month_brand: dict[tuple, Metrics] = defaultdict(Metrics)
        self.sales_week_brand: dict[tuple, Metrics] = defaultdict(Metrics)
        self.sales_month_desc: dict[tuple, Metrics] = defaultdict(Metrics)
        self.sales_week_desc: dict[tuple, Metrics] = defaultdict(Metrics)
        self.sales_dow: dict[tuple, Metrics] = defaultdict(Metrics)  # (period_type, period_key, store, dow_name)
        self.sales_daypart: dict[tuple, Metrics] = defaultdict(Metrics)

        self.delivery_month_business: dict[tuple, Metrics] = defaultdict(Metrics)
        self.delivery_week_business: dict[tuple, Metrics] = defaultdict(Metrics)
        self.delivery_month_store_platform: dict[tuple, Metrics] = defaultdict(Metrics)
        self.delivery_week_store_platform: dict[tuple, Metrics] = defaultdict(Metrics)
        self.delivery_month_cat: dict[tuple, Metrics] = defaultdict(Metrics)
        self.delivery_week_cat: dict[tuple, Metrics] = defaultdict(Metrics)
        self.delivery_dow: dict[tuple, Metrics] = defaultdict(Metrics)
        self.delivery_daypart: dict[tuple, Metrics] = defaultdict(Metrics)

        self.tx_week: dict[tuple, float] = defaultdict(float)
        self.tx_month: dict[tuple, float] = defaultdict(float)
        self.delivery_tx_week: dict[tuple, float] = defaultdict(float)
        self.delivery_tx_month: dict[tuple, float] = defaultdict(float)
        self.delivery_tx_store_platform_week: dict[tuple, float] = defaultdict(float)
        self.delivery_tx_store_platform_month: dict[tuple, float] = defaultdict(float)
        self.delivery_tx_dow: dict[tuple, float] = defaultdict(float)

        # SBF
        self.sbf_week_business: dict[tuple, int] = defaultdict(int)
        self.sbf_month_business: dict[tuple, int] = defaultdict(int)
        self.sbf_week_service_store: dict[tuple, int] = defaultdict(int)
        self.sbf_month_service_store: dict[tuple, int] = defaultdict(int)

        # Inventory
        self.inv_rows: list[dict[str, Any]] = []
        self.shortage_rows: list[dict[str, Any]] = []
        self.excess_rows: list[dict[str, Any]] = []
        self.transfer_rows: list[dict[str, Any]] = []

        # Shrink
        self.shrink_month_business: dict[tuple, dict[str, float]] = defaultdict(lambda: {'total_sales': 0.0, 'shrink_amount': 0.0})
        self.shrink_month_store: dict[tuple, dict[str, float]] = defaultdict(lambda: {'total_sales': 0.0, 'shrink_amount': 0.0})
        self.shrink_detail_desc: dict[tuple, dict[str, Any]] = defaultdict(lambda: {'total_sales': 0.0, 'shrink_amount': 0.0, 'qty_packs': 0.0})
        self.shrink_detail_supplier: dict[tuple, dict[str, Any]] = defaultdict(lambda: {'total_sales': 0.0, 'shrink_amount': 0.0, 'qty_packs': 0.0})

        # CXC
        self.cxc_detail_rows: list[dict[str, Any]] = []
        self.cxc_summary: dict[tuple, dict[str, float]] = defaultdict(lambda: {'amount': 0.0, 'entry_count': 0.0})

        # PRD0
        self.prd0_detail_rows: list[dict[str, Any]] = []
        self.prd0_summary: dict[tuple, dict[str, float]] = defaultdict(lambda: {'item_count': 0.0, 'lost_sales_3d': 0.0})

        # Innovation
        self.innov_month_summary: dict[tuple, dict[str, float]] = defaultdict(lambda: {'sales': 0.0, 'qty': 0.0})
        self.innov_week_summary: dict[tuple, dict[str, float]] = defaultdict(lambda: {'sales': 0.0, 'qty': 0.0})
        self.innov_by_store: dict[tuple, dict[str, float]] = defaultdict(lambda: {'sales': 0.0, 'qty': 0.0})
        self.innov_by_dow: dict[tuple, dict[str, float]] = defaultdict(lambda: {'sales': 0.0, 'qty': 0.0})
        self.innov_by_daypart: dict[tuple, dict[str, float]] = defaultdict(lambda: {'sales': 0.0, 'qty': 0.0})

        # Signals
        self.sales_signals: list[dict[str, Any]] = []
        self.delivery_signals: list[dict[str, Any]] = []
        self.inventory_signals: list[dict[str, Any]] = []
        self.shrink_signals: list[dict[str, Any]] = []

    # ------------------------------------------------------------------
    # Sheet utilities
    # ------------------------------------------------------------------
    def iter_dicts(self, sheet_name: str) -> Iterable[dict[str, Any]]:
        ws = self.wb[sheet_name]
        rows = ws.iter_rows(values_only=True)
        headers = [safe_str(v) for v in next(rows)]
        for row in rows:
            rec = {headers[i]: row[i] if i < len(row) else None for i in range(len(headers))}
            yield rec

    # ------------------------------------------------------------------
    # Builders
    # ------------------------------------------------------------------
    def build_sales_and_delivery(self) -> None:
        for r in self.iter_dicts('Cubo_Sales_DetailHallazgos'):
            month_key = safe_str(r.get('CalMonth'))
            week_key = safe_str(r.get('YearWeek'))
            store = safe_str(r.get('Store')).upper()
            dept = safe_str(r.get('Department'))
            cat = safe_str(r.get('Category'))
            supplier = safe_str(r.get('Supplier'))
            brand = safe_str(r.get('Brand'))
            desc = safe_str(r.get('Description'))
            daypart = safe_str(r.get('Daypart'))
            dow_name = safe_str(r.get('DOW_Name'))
            is_delivery = bool(r.get('Is_Delivery'))
            platform = safe_str(r.get('Delivery_Channel'))
            qty = safe_float(r.get('Qty Sold'))
            sales = safe_float(r.get('Sales'))
            margin = safe_float(r.get('Total Gross Margin'))

            if not month_key or not week_key or not store:
                continue

            self.store_names.add(store)
            self.months.add(month_key)
            self.weeks.add(week_key)
            if dept:
                self.sales_catalogs['departments'].add(dept)
            if cat:
                self.sales_catalogs['categories'].add(cat)
            if brand:
                self.sales_catalogs['brands'].add(brand)
            if supplier:
                self.sales_catalogs['suppliers'].add(supplier)
            if desc:
                self.sales_catalogs['descriptions'].add(desc)

            # Sales aggs
            self.sales_month_business[(month_key,)].add(sales=sales, margin=margin, qty=qty)
            self.sales_week_business[(week_key,)].add(sales=sales, margin=margin, qty=qty)
            self.sales_month_store[(month_key, store)].add(sales=sales, margin=margin, qty=qty)
            self.sales_week_store[(week_key, store)].add(sales=sales, margin=margin, qty=qty)
            self.sales_month_dept[(month_key, dept)].add(sales=sales, margin=margin, qty=qty)
            self.sales_week_dept[(week_key, dept)].add(sales=sales, margin=margin, qty=qty)
            self.sales_month_cat[(month_key, dept, cat)].add(sales=sales, margin=margin, qty=qty)
            self.sales_week_cat[(week_key, dept, cat)].add(sales=sales, margin=margin, qty=qty)
            self.sales_month_brand[(month_key, dept, cat, brand)].add(sales=sales, margin=margin, qty=qty)
            self.sales_week_brand[(week_key, dept, cat, brand)].add(sales=sales, margin=margin, qty=qty)
            self.sales_month_desc[(month_key, dept, cat, brand, desc, supplier)].add(sales=sales, margin=margin, qty=qty)
            self.sales_week_desc[(week_key, dept, cat, brand, desc, supplier)].add(sales=sales, margin=margin, qty=qty)
            self.sales_dow[('month', month_key, store, dow_name)].add(sales=sales, margin=margin, qty=qty)
            self.sales_dow[('week', week_key, store, dow_name)].add(sales=sales, margin=margin, qty=qty)
            self.sales_daypart[('month', month_key, store, daypart)].add(sales=sales, margin=margin, qty=qty)
            self.sales_daypart[('week', week_key, store, daypart)].add(sales=sales, margin=margin, qty=qty)

            if is_delivery:
                platform = platform or 'Delivery'
                self.delivery_month_business[(month_key,)].add(sales=sales, margin=margin, qty=qty)
                self.delivery_week_business[(week_key,)].add(sales=sales, margin=margin, qty=qty)
                self.delivery_month_store_platform[(month_key, store, platform)].add(sales=sales, margin=margin, qty=qty)
                self.delivery_week_store_platform[(week_key, store, platform)].add(sales=sales, margin=margin, qty=qty)
                self.delivery_month_cat[(month_key, dept, cat)].add(sales=sales, margin=margin, qty=qty)
                self.delivery_week_cat[(week_key, dept, cat)].add(sales=sales, margin=margin, qty=qty)
                self.delivery_dow[('month', month_key, store, dow_name, platform)].add(sales=sales, margin=margin, qty=qty)
                self.delivery_dow[('week', week_key, store, dow_name, platform)].add(sales=sales, margin=margin, qty=qty)
                self.delivery_daypart[('month', month_key, store, daypart, platform)].add(sales=sales, margin=margin, qty=qty)
                self.delivery_daypart[('week', week_key, store, daypart, platform)].add(sales=sales, margin=margin, qty=qty)

    def build_transactions(self) -> None:
        for r in self.iter_dicts('Cubo_TX_Fact'):
            week_key = safe_str(r.get('YearMonth'))
            store = safe_str(r.get('Store')).upper()
            daypart = safe_str(r.get('Daypart'))
            dow_name = safe_str(r.get('DOW_Name'))
            is_delivery = bool(r.get('Is_Delivery'))
            platform = safe_str(r.get('Delivery_Channel')) or 'Delivery'
            tx = safe_float(r.get('Transactions'))
            if not week_key:
                continue
            self.tx_week[(week_key,)] += tx
            self.tx_week[(week_key, store)] += tx
            if is_delivery:
                self.delivery_tx_week[(week_key,)] += tx
                self.delivery_tx_store_platform_week[(week_key, store, platform)] += tx
                self.delivery_tx_dow[('week', week_key, store, dow_name, platform)] += tx
                self.sales_daypart[('week', week_key, store, daypart)].tx += tx
                self.delivery_daypart[('week', week_key, store, daypart, platform)].tx += tx
            else:
                self.sales_daypart[('week', week_key, store, daypart)].tx += tx
            self.sales_dow[('week', week_key, store, dow_name)].tx += tx

        for r in self.iter_dicts('Cubo_TX_Month_Fact'):
            month_key = safe_str(r.get('CalMonth'))
            store = safe_str(r.get('Store')).upper()
            daypart = safe_str(r.get('Daypart'))
            dow_name = safe_str(r.get('DOW_Name'))
            is_delivery = bool(r.get('Is_Delivery'))
            platform = safe_str(r.get('Delivery_Channel')) or 'Delivery'
            tx = safe_float(r.get('Transactions'))
            if not month_key:
                continue
            self.tx_month[(month_key,)] += tx
            self.tx_month[(month_key, store)] += tx
            if is_delivery:
                self.delivery_tx_month[(month_key,)] += tx
                self.delivery_tx_store_platform_month[(month_key, store, platform)] += tx
                self.delivery_tx_dow[('month', month_key, store, dow_name, platform)] += tx
                self.sales_daypart[('month', month_key, store, daypart)].tx += tx
                self.delivery_daypart[('month', month_key, store, daypart, platform)].tx += tx
            else:
                self.sales_daypart[('month', month_key, store, daypart)].tx += tx
            self.sales_dow[('month', month_key, store, dow_name)].tx += tx

    def build_sbf(self) -> None:
        for r in self.iter_dicts('Cubo_SBF_Fact'):
            period_key = safe_str(r.get('YearMonth'))
            store = safe_str(r.get('Store')).upper()
            service = safe_str(r.get('Provider'))
            tx = safe_int(r.get('Transactions'))
            if not period_key:
                continue
            self.sales_catalogs['sbf_services'].add(service)
            if '-W' in period_key:
                self.sbf_week_business[(period_key,)] += tx
                self.sbf_week_service_store[(period_key, service, store)] += tx
            else:
                self.sbf_month_business[(period_key,)] += tx
                self.sbf_month_service_store[(period_key, service, store)] += tx

    def build_inventory(self) -> None:
        by_desc: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for r in self.iter_dicts('Inventario_Detail'):
            store = safe_str(r.get('Store')).upper()
            dept = safe_str(r.get('Department'))
            cat = safe_str(r.get('Category'))
            brand = safe_str(r.get('Brand'))
            supplier = safe_str(r.get('Supplier'))
            desc = safe_str(r.get('Description'))
            stock_units = safe_float(r.get('Quantity'))
            inventory_amount = safe_float(r.get('Amount'))
            doi = safe_float(r.get('Días Inv'))
            target_qty = safe_float(r.get('TargetQty'))
            ideal_qty = safe_float(r.get('IdealQty'))
            unit_cost = safe_float(r.get('Costo Unitario Inv')) or (inventory_amount / stock_units if stock_units else 0.0)
            sales_s4 = safe_float(r.get('Sales S4'))
            qty_s4 = safe_float(r.get('Qty Sold S4'))
            lookup = {
                'store_name': store,
                'store': store,
                'department': dept,
                'category': cat,
                'brand': brand,
                'supplier': supplier,
                'description': desc,
                'stock_units': round(stock_units, 2),
                'inventory_amount': round(inventory_amount, 2),
                'doi': round(doi, 2),
                'target_qty': round(target_qty, 2),
                'ideal_qty': round(ideal_qty, 2),
                'unit_cost': round(unit_cost, 4),
                'sales_s4': round(sales_s4, 2),
                'qty_s4': round(qty_s4, 2),
            }
            self.inv_rows.append(lookup)
            by_desc[desc].append(lookup)

            shortage_units = max(target_qty - stock_units, 0.0)
            excess_units = max(stock_units - target_qty, 0.0)
            if shortage_units > 0:
                shortage_amount = shortage_units * unit_cost
                self.shortage_rows.append({
                    **lookup,
                    'shortage_units': round(shortage_units, 2),
                    'shortage_amount': round(shortage_amount, 2),
                    'severity': severity_for_amount(shortage_amount),
                    'priority': priority_for_amount(shortage_amount),
                })
            if excess_units > 0:
                excess_amount = excess_units * unit_cost
                self.excess_rows.append({
                    **lookup,
                    'excess_units': round(excess_units, 2),
                    'excess_amount': round(excess_amount, 2),
                    'severity': severity_for_amount(excess_amount),
                    'priority': priority_for_amount(excess_amount),
                })

        # transfer candidates by description
        for desc, rows in by_desc.items():
            shortages = [r for r in self.shortage_rows if r['description'] == desc and r['shortage_units'] > 0]
            excesses = [r for r in self.excess_rows if r['description'] == desc and r['excess_units'] > 0]
            shortages.sort(key=lambda x: x['shortage_amount'], reverse=True)
            excesses.sort(key=lambda x: x['excess_amount'], reverse=True)
            for s in shortages:
                need = s['shortage_units']
                if need <= 0:
                    continue
                for e in excesses:
                    available = e.get('excess_units', 0.0)
                    if available <= 0 or e['store_name'] == s['store_name']:
                        continue
                    move = min(need, available)
                    if move <= 0:
                        continue
                    amount = move * (s.get('unit_cost') or e.get('unit_cost') or 0.0)
                    self.transfer_rows.append({
                        'description': desc,
                        'department': s.get('department', ''),
                        'category': s.get('category', ''),
                        'brand': s.get('brand', ''),
                        'supplier': s.get('supplier', ''),
                        'from_store': e['store_name'],
                        'to_store': s['store_name'],
                        'store_name': s['store_name'],
                        'transfer_units': round(move, 2),
                        'transfer_amount': round(amount, 2),
                        'severity': severity_for_amount(amount),
                        'priority': priority_for_amount(amount),
                    })
                    e['excess_units'] = round(available - move, 2)
                    need = round(need - move, 2)
                    if need <= 0:
                        break

    def build_shrink(self) -> None:
        for r in self.iter_dicts('Merma_Detail'):
            store = safe_str(r.get('Store')).upper()
            month_key = safe_str(r.get('CalMonth'))
            supplier = safe_str(r.get('Supplier'))
            category = safe_str(r.get('Category'))
            brand = safe_str(r.get('Brand'))
            desc = safe_str(r.get('Description'))
            total_sales = safe_float(r.get('Total Sales'))
            shrink_amount = safe_float(r.get('Merma'))
            qty_packs = safe_float(r.get('Quantity (packs)'))
            if not month_key:
                continue
            self.shrink_month_business[(month_key,)]['total_sales'] += total_sales
            self.shrink_month_business[(month_key,)]['shrink_amount'] += shrink_amount
            self.shrink_month_store[(month_key, store)]['total_sales'] += total_sales
            self.shrink_month_store[(month_key, store)]['shrink_amount'] += shrink_amount

            d1 = self.shrink_detail_desc[(month_key, store, desc)]
            d1['month_key'] = month_key
            d1['store_name'] = store
            d1['store'] = store
            d1['supplier'] = supplier
            d1['category'] = category
            d1['brand'] = brand
            d1['description'] = desc
            d1['total_sales'] += total_sales
            d1['shrink_amount'] += shrink_amount
            d1['qty_packs'] += qty_packs

            d2 = self.shrink_detail_supplier[(month_key, supplier)]
            d2['month_key'] = month_key
            d2['supplier'] = supplier
            d2['store_name'] = ''
            d2['store'] = ''
            d2['total_sales'] += total_sales
            d2['shrink_amount'] += shrink_amount
            d2['qty_packs'] += qty_packs

    def build_cxc(self) -> None:
        for r in self.iter_dicts('CXC_DETAIL'):
            store = safe_str(r.get('Store')).upper()
            month_key = safe_str(r.get('Month'))
            week_key = safe_str(r.get('YearMonth')) if '-W' in safe_str(r.get('YearMonth')) else ''
            comment = safe_str(r.get('Comment'))
            cashier = safe_str(r.get('Cashier'))
            batch = safe_str(r.get('BatchNumber'))
            ref = safe_str(r.get('ID'))
            amount = safe_float(r.get('Amount'))
            tm = r.get('Time')
            tm_iso = tm.isoformat() if isinstance(tm, datetime) else safe_str(tm)
            row = {
                'store_name': store,
                'store': store,
                'month_key': month_key,
                'week_key': week_key,
                'comment': comment,
                'cashier': cashier,
                'batch_number': batch,
                'reference_id': ref,
                'time': tm_iso,
                'amount': round(amount, 2),
                'entry_count': 1,
            }
            self.cxc_detail_rows.append(row)
            self.cxc_summary[(month_key, store)]['amount'] += amount
            self.cxc_summary[(month_key, store)]['entry_count'] += 1

    def build_prd0(self) -> None:
        for r in self.iter_dicts('PRD_CERO_DETAIL'):
            store = safe_str(r.get('Store')).upper()
            supplier = safe_str(r.get('Supplier'))
            dept = safe_str(r.get('Department'))
            cat = safe_str(r.get('Category'))
            brand = safe_str(r.get('Brand'))
            desc = safe_str(r.get('Description'))
            lost = safe_float(r.get('Lost Sales 3d'))
            item_row = {
                'store_name': store,
                'store': store,
                'supplier': supplier,
                'department': dept,
                'category': cat,
                'brand': brand,
                'description': desc,
                'item_count': 1,
                'lost_sales_3d': round(lost, 2),
                'quantity': round(safe_float(r.get('Quantity')), 2),
                'rop': round(safe_float(r.get('ROP')), 2),
                'qty_s4': round(safe_float(r.get('Qty Sold S4')), 2),
                'sales_s4': round(safe_float(r.get('Sales S4')), 2),
            }
            self.prd0_detail_rows.append(item_row)
            self.prd0_summary[(supplier,)]['item_count'] += 1
            self.prd0_summary[(supplier,)]['lost_sales_3d'] += lost

    def build_innovation(self) -> None:
        for r in self.iter_dicts('Innovation_Combos_Agg'):
            period_type = safe_str(r.get('PeriodType'))
            period_key = safe_str(r.get('PeriodKey'))
            store = safe_str(r.get('Store')).upper()
            combo_label = safe_str(r.get('ComboLabel'))
            qty = safe_float(r.get('Qty Sold'))
            sales = safe_float(r.get('Sales'))
            if period_type == 'month':
                self.innov_month_summary[(period_key, combo_label)]['sales'] += sales
                self.innov_month_summary[(period_key, combo_label)]['qty'] += qty
            elif period_type == 'week':
                self.innov_week_summary[(period_key, combo_label)]['sales'] += sales
                self.innov_week_summary[(period_key, combo_label)]['qty'] += qty
            self.innov_by_store[(period_type, period_key, store)]['sales'] += sales
            self.innov_by_store[(period_type, period_key, store)]['qty'] += qty

        for r in self.iter_dicts('Innovation_Combos_Detail'):
            month_key = safe_str(r.get('CalMonth'))
            week_key = safe_str(r.get('YearWeek'))
            store = safe_str(r.get('Store')).upper()
            combo_label = safe_str(r.get('ComboLabel'))
            dow_name = safe_str(r.get('DOW_Name'))
            daypart = safe_str(r.get('Daypart'))
            qty = safe_float(r.get('Qty Sold'))
            sales = safe_float(r.get('Sales'))
            self.innov_by_dow[('month', month_key, store, dow_name, combo_label)]['sales'] += sales
            self.innov_by_dow[('month', month_key, store, dow_name, combo_label)]['qty'] += qty
            self.innov_by_dow[('week', week_key, store, dow_name, combo_label)]['sales'] += sales
            self.innov_by_dow[('week', week_key, store, dow_name, combo_label)]['qty'] += qty
            self.innov_by_daypart[('month', month_key, store, daypart, combo_label)]['sales'] += sales
            self.innov_by_daypart[('month', month_key, store, daypart, combo_label)]['qty'] += qty
            self.innov_by_daypart[('week', week_key, store, daypart, combo_label)]['sales'] += sales
            self.innov_by_daypart[('week', week_key, store, daypart, combo_label)]['qty'] += qty

    # ------------------------------------------------------------------
    # Post processing / row builders
    # ------------------------------------------------------------------
    def attach_tx(self) -> None:
        for (month_key,), m in self.sales_month_business.items():
            m.tx += self.tx_month[(month_key,)]
        for (week_key,), m in self.sales_week_business.items():
            m.tx += self.tx_week[(week_key,)]
        for (month_key, store), m in self.sales_month_store.items():
            m.tx += self.tx_month[(month_key, store)]
        for (week_key, store), m in self.sales_week_store.items():
            m.tx += self.tx_week[(week_key, store)]
        for (month_key,), m in self.delivery_month_business.items():
            m.tx += self.delivery_tx_month[(month_key,)]
        for (week_key,), m in self.delivery_week_business.items():
            m.tx += self.delivery_tx_week[(week_key,)]
        for (month_key, store, platform), m in self.delivery_month_store_platform.items():
            m.tx += self.delivery_tx_store_platform_month[(month_key, store, platform)]
        for (week_key, store, platform), m in self.delivery_week_store_platform.items():
            m.tx += self.delivery_tx_store_platform_week[(week_key, store, platform)]
        for key, tx in self.delivery_tx_dow.items():
            if key in self.delivery_dow:
                self.delivery_dow[key].tx += tx

    def build_rows_from_metrics(self, aggs: dict, time_field: str, group_fields: list[str]) -> list[dict[str, Any]]:
        rows = []
        for key, metrics in aggs.items():
            payload = metrics.to_sales_payload()
            row = {time_field: key[0]}
            for idx, gf in enumerate(group_fields, start=1):
                row[gf] = key[idx]
            row.update(payload)
            if 'store_name' not in row and 'store' in row:
                row['store_name'] = row['store']
            rows.append(row)
        return rows

    def enrich_compare(self, rows: list[dict[str, Any]], time_field: str, group_fields: list[str], month_mode: bool) -> list[dict[str, Any]]:
        metric_fields = ['sales', 'margin', 'margin_pct', 'qty', 'tx', 'avg_ticket']
        index = {}
        for row in rows:
            key = (safe_str(row.get(time_field)),) + tuple(safe_str(row.get(g)) for g in group_fields)
            index[key] = row
        for row in rows:
            period = safe_str(row.get(time_field))
            ly_period = month_ly(period) if month_mode else week_ly(period)
            lm_period = month_prev(period) if month_mode else ''
            group_vals = tuple(safe_str(row.get(g)) for g in group_fields)
            ly_row = index.get((ly_period,) + group_vals)
            lm_row = index.get((lm_period,) + group_vals) if month_mode else None
            for mf in metric_fields:
                row[f'{mf}_ly'] = round(safe_float(ly_row.get(mf) if ly_row else 0), 2 if mf not in {'tx'} else 0)
                if month_mode:
                    row[f'{mf}_lm'] = round(safe_float(lm_row.get(mf) if lm_row else 0), 2 if mf not in {'tx'} else 0)
        return rows

    def sales_rows(self) -> dict[str, list[dict[str, Any]]]:
        monthly_business = self.enrich_compare(self.build_rows_from_metrics(self.sales_month_business, 'month_key', []), 'month_key', [], True)
        weekly_business = self.enrich_compare(self.build_rows_from_metrics(self.sales_week_business, 'week_key', []), 'week_key', [], False)
        monthly_store = self.enrich_compare(self.build_rows_from_metrics(self.sales_month_store, 'month_key', ['store']), 'month_key', ['store'], True)
        weekly_store = self.enrich_compare(self.build_rows_from_metrics(self.sales_week_store, 'week_key', ['store']), 'week_key', ['store'], False)
        for row in monthly_store + weekly_store:
            row['store_name'] = row.pop('store')
        monthly_dept = self.enrich_compare(self.build_rows_from_metrics(self.sales_month_dept, 'month_key', ['department']), 'month_key', ['department'], True)
        weekly_dept = self.enrich_compare(self.build_rows_from_metrics(self.sales_week_dept, 'week_key', ['department']), 'week_key', ['department'], False)
        monthly_cat = self.enrich_compare(self.build_rows_from_metrics(self.sales_month_cat, 'month_key', ['department', 'category']), 'month_key', ['department', 'category'], True)
        weekly_cat = self.enrich_compare(self.build_rows_from_metrics(self.sales_week_cat, 'week_key', ['department', 'category']), 'week_key', ['department', 'category'], False)
        monthly_brand = self.enrich_compare(self.build_rows_from_metrics(self.sales_month_brand, 'month_key', ['department', 'category', 'brand']), 'month_key', ['department', 'category', 'brand'], True)
        weekly_brand = self.enrich_compare(self.build_rows_from_metrics(self.sales_week_brand, 'week_key', ['department', 'category', 'brand']), 'week_key', ['department', 'category', 'brand'], False)
        monthly_desc = self.enrich_compare(self.build_rows_from_metrics(self.sales_month_desc, 'month_key', ['department', 'category', 'brand', 'description', 'supplier']), 'month_key', ['department', 'category', 'brand', 'description', 'supplier'], True)

        sales_by_dow = []
        for (ptype, pkey, store, dow_name), m in self.sales_dow.items():
            row = {'period_type': ptype, 'period_key': pkey, 'store_name': store, 'store': store, 'dow_name': dow_name}
            row.update(m.to_sales_payload())
            sales_by_dow.append(row)

        sales_by_daypart = []
        for (ptype, pkey, store, daypart), m in self.sales_daypart.items():
            row = {'period_type': ptype, 'period_key': pkey, 'store_name': store, 'store': store, 'daypart': daypart}
            row.update(m.to_sales_payload())
            sales_by_daypart.append(row)

        return {
            'monthly_business': monthly_business,
            'weekly_business': weekly_business,
            'monthly_store': monthly_store,
            'weekly_store': weekly_store,
            'monthly_department': monthly_dept,
            'weekly_department': weekly_dept,
            'monthly_category': monthly_cat,
            'weekly_category': weekly_cat,
            'monthly_brand': monthly_brand,
            'weekly_brand': weekly_brand,
            'monthly_description': monthly_desc,
            'sales_by_dow': sales_by_dow,
            'sales_by_daypart': sales_by_daypart,
            'sales_by_hour': [],
        }

    def delivery_rows(self) -> dict[str, list[dict[str, Any]]]:
        monthly_business = self.enrich_compare(self.build_rows_from_metrics(self.delivery_month_business, 'month_key', []), 'month_key', [], True)
        weekly_business = self.enrich_compare(self.build_rows_from_metrics(self.delivery_week_business, 'week_key', []), 'week_key', [], False)
        monthly_store_platform = self.enrich_compare(self.build_rows_from_metrics(self.delivery_month_store_platform, 'month_key', ['store', 'platform']), 'month_key', ['store', 'platform'], True)
        weekly_store_platform = self.enrich_compare(self.build_rows_from_metrics(self.delivery_week_store_platform, 'week_key', ['store', 'platform']), 'week_key', ['store', 'platform'], False)
        for row in monthly_store_platform + weekly_store_platform:
            row['store_name'] = row.pop('store')
        monthly_category = self.enrich_compare(self.build_rows_from_metrics(self.delivery_month_cat, 'month_key', ['department', 'category']), 'month_key', ['department', 'category'], True)
        weekly_category = self.enrich_compare(self.build_rows_from_metrics(self.delivery_week_cat, 'week_key', ['department', 'category']), 'week_key', ['department', 'category'], False)

        by_dow = []
        for (ptype, pkey, store, dow_name, platform), m in self.delivery_dow.items():
            row = {'period_type': ptype, 'period_key': pkey, 'store_name': store, 'store': store, 'dow_name': dow_name, 'platform': platform}
            row.update(m.to_sales_payload())
            by_dow.append(row)

        by_daypart = []
        for (ptype, pkey, store, daypart, platform), m in self.delivery_daypart.items():
            row = {'period_type': ptype, 'period_key': pkey, 'store_name': store, 'store': store, 'daypart': daypart, 'platform': platform}
            row.update(m.to_sales_payload())
            by_daypart.append(row)

        return {
            'monthly_business': monthly_business,
            'weekly_business': weekly_business,
            'monthly_store_platform': monthly_store_platform,
            'weekly_store_platform': weekly_store_platform,
            'monthly_category': monthly_category,
            'weekly_category': weekly_category,
            'delivery_by_dow': by_dow,
            'delivery_by_daypart': by_daypart,
            'delivery_by_hour': [],
        }

    def sbf_rows(self) -> dict[str, list[dict[str, Any]]]:
        monthly_business = [{'month_key': k[0], 'tx': v} for k, v in sorted(self.sbf_month_business.items())]
        weekly_business = [{'week_key': k[0], 'tx': v} for k, v in sorted(self.sbf_week_business.items())]
        # compare fields
        self._attach_tx_compares(monthly_business, 'month_key', [], month_mode=True)
        self._attach_tx_compares(weekly_business, 'week_key', [], month_mode=False)
        monthly_service_store = [{'month_key': k[0], 'service': k[1], 'store_name': k[2], 'store': k[2], 'tx': v} for k, v in sorted(self.sbf_month_service_store.items())]
        weekly_service_store = [{'week_key': k[0], 'service': k[1], 'store_name': k[2], 'store': k[2], 'tx': v} for k, v in sorted(self.sbf_week_service_store.items())]
        self._attach_tx_compares(monthly_service_store, 'month_key', ['service', 'store_name'], month_mode=True)
        self._attach_tx_compares(weekly_service_store, 'week_key', ['service', 'store_name'], month_mode=False)
        return {
            'monthly_business': monthly_business,
            'weekly_business': weekly_business,
            'monthly_service_store': monthly_service_store,
            'weekly_service_store': weekly_service_store,
        }

    def _attach_tx_compares(self, rows: list[dict[str, Any]], time_field: str, group_fields: list[str], month_mode: bool) -> None:
        idx = {}
        for r in rows:
            idx[(safe_str(r.get(time_field)),) + tuple(safe_str(r.get(g)) for g in group_fields)] = r
        for r in rows:
            pk = safe_str(r.get(time_field))
            ly = month_ly(pk) if month_mode else week_ly(pk)
            lm = month_prev(pk) if month_mode else ''
            gv = tuple(safe_str(r.get(g)) for g in group_fields)
            ly_row = idx.get((ly,) + gv)
            lm_row = idx.get((lm,) + gv) if month_mode else None
            r['tx_ly'] = safe_int(ly_row.get('tx') if ly_row else 0)
            if month_mode:
                r['tx_lm'] = safe_int(lm_row.get('tx') if lm_row else 0)

    def inventory_rows(self) -> dict[str, list[dict[str, Any]]]:
        return {
            'store_sku_snapshot': self.inv_rows,
            'doi_store_sku': self.inv_rows,
            'shortage_alerts': sorted(self.shortage_rows, key=lambda x: x.get('shortage_amount', 0), reverse=True),
            'excess_alerts': sorted(self.excess_rows, key=lambda x: x.get('excess_amount', 0), reverse=True),
            'transfer_candidates': sorted(self.transfer_rows, key=lambda x: x.get('transfer_amount', 0), reverse=True),
        }

    def shrink_rows(self) -> dict[str, list[dict[str, Any]]]:
        monthly_business = []
        for (month_key,), v in sorted(self.shrink_month_business.items()):
            total_sales = v['total_sales']
            shrink_amount = v['shrink_amount']
            monthly_business.append({
                'month_key': month_key,
                'shrink_amount': round(shrink_amount, 2),
                'total_sales': round(total_sales, 2),
                'shrink_pct': round((shrink_amount / total_sales) * 100, 2) if total_sales else 0.0,
            })
        self._attach_shrink_compares(monthly_business, 'month_key', [])

        monthly_store = []
        for (month_key, store), v in sorted(self.shrink_month_store.items()):
            total_sales = v['total_sales']
            shrink_amount = v['shrink_amount']
            monthly_store.append({
                'month_key': month_key,
                'store_name': store,
                'store': store,
                'shrink_amount': round(shrink_amount, 2),
                'total_sales': round(total_sales, 2),
                'shrink_pct': round((shrink_amount / total_sales) * 100, 2) if total_sales else 0.0,
            })
        self._attach_shrink_compares(monthly_store, 'month_key', ['store_name'])

        detail_desc = []
        for (_m, _s, _d), v in self.shrink_detail_desc.items():
            total_sales = v['total_sales']
            shrink_amount = v['shrink_amount']
            detail_desc.append({
                **{k: v[k] for k in ['month_key', 'store_name', 'store', 'supplier', 'category', 'brand', 'description']},
                'total_sales': round(total_sales, 2),
                'shrink_amount': round(shrink_amount, 2),
                'qty_packs': round(v['qty_packs'], 2),
                'shrink_pct': round((shrink_amount / total_sales) * 100, 2) if total_sales else 0.0,
            })
        self._attach_shrink_compares(detail_desc, 'month_key', ['store_name', 'description'])

        detail_supplier = []
        for (_m, _sup), v in self.shrink_detail_supplier.items():
            total_sales = v['total_sales']
            shrink_amount = v['shrink_amount']
            detail_supplier.append({
                'month_key': v['month_key'],
                'supplier': v['supplier'],
                'total_sales': round(total_sales, 2),
                'shrink_amount': round(shrink_amount, 2),
                'qty_packs': round(v['qty_packs'], 2),
                'shrink_pct': round((shrink_amount / total_sales) * 100, 2) if total_sales else 0.0,
            })
        self._attach_shrink_compares(detail_supplier, 'month_key', ['supplier'])

        return {
            'monthly_business': monthly_business,
            'monthly_store': monthly_store,
            'detail_store_description': detail_desc,
            'detail_supplier': detail_supplier,
        }

    def _attach_shrink_compares(self, rows: list[dict[str, Any]], time_field: str, group_fields: list[str]) -> None:
        idx = {}
        for r in rows:
            idx[(safe_str(r.get(time_field)),) + tuple(safe_str(r.get(g)) for g in group_fields)] = r
        for r in rows:
            pk = safe_str(r.get(time_field))
            ly = month_ly(pk)
            lm = month_prev(pk)
            gv = tuple(safe_str(r.get(g)) for g in group_fields)
            ly_row = idx.get((ly,) + gv)
            lm_row = idx.get((lm,) + gv)
            for f in ['shrink_amount', 'shrink_pct', 'total_sales']:
                r[f'{f}_ly'] = round(safe_float(ly_row.get(f) if ly_row else 0), 2)
                r[f'{f}_lm'] = round(safe_float(lm_row.get(f) if lm_row else 0), 2)

    def cxc_rows(self) -> dict[str, list[dict[str, Any]]]:
        summary = []
        for (month_key, store), v in sorted(self.cxc_summary.items()):
            summary.append({
                'month_key': month_key,
                'store_name': store,
                'store': store,
                'amount': round(v['amount'], 2),
                'entry_count': safe_int(v['entry_count']),
            })
        self._attach_cxc_compares(summary)
        return {'summary': summary, 'detail': self.cxc_detail_rows}

    def _attach_cxc_compares(self, rows: list[dict[str, Any]]) -> None:
        idx = {(safe_str(r['month_key']), safe_str(r['store_name'])): r for r in rows}
        for r in rows:
            pk = safe_str(r['month_key'])
            store = safe_str(r['store_name'])
            ly_row = idx.get((month_ly(pk), store))
            lm_row = idx.get((month_prev(pk), store))
            for f in ['amount', 'entry_count']:
                r[f'{f}_ly'] = round(safe_float(ly_row.get(f) if ly_row else 0), 2 if f == 'amount' else 0)
                r[f'{f}_lm'] = round(safe_float(lm_row.get(f) if lm_row else 0), 2 if f == 'amount' else 0)

    def prd0_rows(self) -> dict[str, list[dict[str, Any]]]:
        summary = []
        for (supplier,), v in sorted(self.prd0_summary.items(), key=lambda kv: kv[1]['lost_sales_3d'], reverse=True):
            summary.append({
                'supplier': supplier,
                'item_count': safe_int(v['item_count']),
                'lost_sales_3d': round(v['lost_sales_3d'], 2),
            })
        return {'supplier_summary': summary, 'supplier_detail': self.prd0_detail_rows}

    def innovation_rows(self) -> dict[str, list[dict[str, Any]]]:
        monthly_summary = []
        for (period_key, combo_label), v in sorted(self.innov_month_summary.items()):
            monthly_summary.append({'period_key': period_key, 'period_type': 'month', 'combo_label': combo_label, 'sales': round(v['sales'], 2), 'qty': round(v['qty'], 2)})
        self._attach_innov_compares(monthly_summary, 'month')
        weekly_summary = []
        for (period_key, combo_label), v in sorted(self.innov_week_summary.items(), key=lambda kv: week_sort_key(kv[0][0])):
            weekly_summary.append({'period_key': period_key, 'period_type': 'week', 'combo_label': combo_label, 'sales': round(v['sales'], 2), 'qty': round(v['qty'], 2)})
        self._attach_innov_compares(weekly_summary, 'week')

        by_store = []
        for (ptype, pkey, store), v in sorted(self.innov_by_store.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2])):
            by_store.append({'period_type': ptype, 'period_key': pkey, 'store_name': store, 'store': store, 'sales': round(v['sales'], 2), 'qty': round(v['qty'], 2)})
        self._attach_innov_compares_mixed(by_store, ['store_name'])

        by_dow = []
        for (ptype, pkey, store, dow_name, combo_label), v in sorted(self.innov_by_dow.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2], kv[0][3], kv[0][4])):
            by_dow.append({'period_type': ptype, 'period_key': pkey, 'store_name': store, 'store': store, 'dow_name': dow_name, 'combo_label': combo_label, 'sales': round(v['sales'], 2), 'qty': round(v['qty'], 2)})
        self._attach_innov_compares_mixed(by_dow, ['store_name', 'dow_name', 'combo_label'])

        by_daypart = []
        for (ptype, pkey, store, daypart, combo_label), v in sorted(self.innov_by_daypart.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2], kv[0][3], kv[0][4])):
            by_daypart.append({'period_type': ptype, 'period_key': pkey, 'store_name': store, 'store': store, 'daypart': daypart, 'combo_label': combo_label, 'sales': round(v['sales'], 2), 'qty': round(v['qty'], 2)})
        self._attach_innov_compares_mixed(by_daypart, ['store_name', 'daypart', 'combo_label'])

        return {
            'monthly_summary': monthly_summary,
            'weekly_summary': weekly_summary,
            'by_store': by_store,
            'by_dow': by_dow,
            'by_daypart': by_daypart,
            'by_hour': [],
        }

    def _attach_innov_compares(self, rows: list[dict[str, Any]], mode: str) -> None:
        idx = {(safe_str(r['period_key']), safe_str(r['combo_label'])): r for r in rows}
        for r in rows:
            pk = safe_str(r['period_key'])
            comp = month_ly(pk) if mode == 'month' else week_ly(pk)
            prev = month_prev(pk) if mode == 'month' else ''
            ly_row = idx.get((comp, safe_str(r['combo_label'])))
            lm_row = idx.get((prev, safe_str(r['combo_label']))) if prev else None
            for f in ['sales', 'qty']:
                r[f'{f}_ly'] = round(safe_float(ly_row.get(f) if ly_row else 0), 2)
                if mode == 'month':
                    r[f'{f}_lm'] = round(safe_float(lm_row.get(f) if lm_row else 0), 2)

    def _attach_innov_compares_mixed(self, rows: list[dict[str, Any]], group_fields: list[str]) -> None:
        idx = {(safe_str(r['period_type']), safe_str(r['period_key'])) + tuple(safe_str(r.get(g)) for g in group_fields): r for r in rows}
        for r in rows:
            ptype = safe_str(r['period_type'])
            pk = safe_str(r['period_key'])
            comp = month_ly(pk) if ptype == 'month' else week_ly(pk)
            prev = month_prev(pk) if ptype == 'month' else ''
            gv = tuple(safe_str(r.get(g)) for g in group_fields)
            ly_row = idx.get((ptype, comp) + gv)
            lm_row = idx.get((ptype, prev) + gv) if prev else None
            for f in ['sales', 'qty']:
                r[f'{f}_ly'] = round(safe_float(ly_row.get(f) if ly_row else 0), 2)
                if ptype == 'month':
                    r[f'{f}_lm'] = round(safe_float(lm_row.get(f) if lm_row else 0), 2)

    def build_signals(self, sales_rows: dict[str, list[dict[str, Any]]], delivery_rows: dict[str, list[dict[str, Any]]], inv_rows: dict[str, list[dict[str, Any]]], shrink_rows: dict[str, list[dict[str, Any]]]) -> None:
        # Sales signals from monthly categories and weekly categories
        for src_name in ['monthly_category', 'weekly_category']:
            tf = 'month_key' if src_name.startswith('monthly') else 'week_key'
            for r in sales_rows[src_name]:
                cur = safe_float(r.get('sales'))
                ly = safe_float(r.get('sales_ly'))
                delta = cur - ly
                if abs(delta) < 50:
                    continue
                delta_pct = (delta / ly * 100) if ly else 0.0
                self.sales_signals.append({
                    'period_key': safe_str(r.get(tf)),
                    'metric': 'sales',
                    'scope_key': safe_str(r.get('category')),
                    'scope_label': safe_str(r.get('category')),
                    'department': safe_str(r.get('department')),
                    'delta': round(delta, 2),
                    'delta_pct': round(delta_pct, 2),
                    'current_value': round(cur, 2),
                    'compare_value': round(ly, 2),
                    'severity': severity_for_amount(delta),
                    'priority': priority_for_amount(delta),
                })
        self.sales_signals.sort(key=lambda x: abs(x['delta']), reverse=True)
        self.sales_signals = self.sales_signals[:400]

        for src_name in ['monthly_store_platform', 'weekly_store_platform']:
            tf = 'month_key' if src_name.startswith('monthly') else 'week_key'
            for r in delivery_rows[src_name]:
                cur = safe_float(r.get('sales'))
                ly = safe_float(r.get('sales_ly'))
                delta = cur - ly
                if abs(delta) < 25:
                    continue
                delta_pct = (delta / ly * 100) if ly else 0.0
                self.delivery_signals.append({
                    'period_key': safe_str(r.get(tf)),
                    'metric': 'sales',
                    'scope_key': f"{safe_str(r.get('store_name'))} | {safe_str(r.get('platform'))}",
                    'scope_label': f"{safe_str(r.get('store_name'))} · {safe_str(r.get('platform'))}",
                    'store_name': safe_str(r.get('store_name')),
                    'platform': safe_str(r.get('platform')),
                    'delta': round(delta, 2),
                    'delta_pct': round(delta_pct, 2),
                    'current_value': round(cur, 2),
                    'compare_value': round(ly, 2),
                    'severity': severity_for_amount(delta),
                    'priority': priority_for_amount(delta),
                })
        self.delivery_signals.sort(key=lambda x: abs(x['delta']), reverse=True)
        self.delivery_signals = self.delivery_signals[:300]

        for r in inv_rows['shortage_alerts'][:250]:
            self.inventory_signals.append({
                'period_key': max(self.months) if self.months else '',
                'metric': 'inventory',
                'scope_key': safe_str(r['description']),
                'scope_label': safe_str(r['description']),
                'store_name': safe_str(r['store_name']),
                'delta': round(safe_float(r.get('shortage_amount')), 2),
                'delta_pct': 0,
                'current_value': round(safe_float(r.get('shortage_units')), 2),
                'compare_value': 0,
                'severity': r['severity'],
                'priority': r['priority'],
            })
        self.inventory_signals.sort(key=lambda x: abs(x['delta']), reverse=True)

        for r in shrink_rows['detail_store_description'][:300]:
            amt = safe_float(r.get('shrink_amount'))
            if abs(amt) < 10:
                continue
            self.shrink_signals.append({
                'period_key': safe_str(r.get('month_key')),
                'metric': 'shrink',
                'scope_key': safe_str(r.get('description')),
                'scope_label': safe_str(r.get('description')),
                'store_name': safe_str(r.get('store_name')),
                'delta': round(amt, 2),
                'delta_pct': round(safe_float(r.get('shrink_pct')), 2),
                'current_value': round(amt, 2),
                'compare_value': round(safe_float(r.get('shrink_amount_ly')), 2),
                'severity': severity_for_amount(amt),
                'priority': priority_for_amount(amt),
            })
        self.shrink_signals.sort(key=lambda x: abs(x['delta']), reverse=True)
        self.shrink_signals = self.shrink_signals[:300]

    def meta_rows(self) -> dict[str, Any]:
        months = sorted(self.months)
        weeks = sorted(self.weeks, key=week_sort_key)
        manifest = {
            'generated_at': datetime.utcnow().isoformat() + 'Z',
            'months_available': months,
            'weeks_available': weeks,
            'stores_available': sorted(self.store_names),
            'data_quality': {
                'delivery_by_hour_available': False,
                'innovation_by_hour_available': False,
                'delivery_by_daypart_available': True,
                'innovation_by_daypart_available': True,
            },
        }
        calendars = {
            'months': [
                {
                    'month_key': m,
                    'label': month_label(m),
                    'comparable_ly_month': month_ly(m),
                    'lm_month': month_prev(m),
                }
                for m in months
            ],
            'weeks': [
                {
                    'week_key': w,
                    'label': w,
                    'comparable_ly_week': week_ly(w),
                }
                for w in weeks
            ],
        }
        catalogs = {
            'departments': sorted(self.sales_catalogs['departments']),
            'categories': sorted(self.sales_catalogs['categories']),
            'brands': sorted(self.sales_catalogs['brands']),
            'suppliers': sorted(self.sales_catalogs['suppliers']),
            'descriptions': sorted(self.sales_catalogs['descriptions']),
            'sbf_services': sorted(self.sales_catalogs['sbf_services']),
        }
        stores = [{'store_name': s, 'store': s} for s in sorted(self.store_names)]
        return {'manifest': manifest, 'calendars': calendars, 'catalogs': catalogs, 'stores': stores}

    # ------------------------------------------------------------------
    def write_all(self) -> None:
        print('1/7 Reading sales + delivery…')
        self.build_sales_and_delivery()
        print('2/7 Reading transactions…')
        self.build_transactions()
        self.attach_tx()
        print('3/7 Reading SBF…')
        self.build_sbf()
        print('4/7 Reading inventory…')
        self.build_inventory()
        print('5/7 Reading shrink + CXC + PRD0…')
        self.build_shrink()
        self.build_cxc()
        self.build_prd0()
        print('6/7 Reading innovation…')
        self.build_innovation()

        meta = self.meta_rows()
        sales = self.sales_rows()
        delivery = self.delivery_rows()
        sbf = self.sbf_rows()
        inventory = self.inventory_rows()
        shrink = self.shrink_rows()
        cxc = self.cxc_rows()
        prd0 = self.prd0_rows()
        innovation = self.innovation_rows()
        self.build_signals(sales, delivery, inventory, shrink)

        print('7/7 Writing JSON files…')
        dump_json(self.out_dir / 'meta/manifest.json', meta['manifest'])
        dump_json(self.out_dir / 'meta/calendars.json', meta['calendars'])
        dump_json(self.out_dir / 'meta/catalogs.json', meta['catalogs'])
        dump_json(self.out_dir / 'meta/stores.json', meta['stores'])

        dump_json(self.out_dir / 'ventas/monthly_business.json', sales['monthly_business'])
        dump_json(self.out_dir / 'ventas/weekly_business.json', sales['weekly_business'])
        dump_json(self.out_dir / 'ventas/monthly_store.json', sales['monthly_store'])
        dump_json(self.out_dir / 'ventas/weekly_store.json', sales['weekly_store'])
        dump_json(self.out_dir / 'ventas/monthly_department.json', sales['monthly_department'])
        dump_json(self.out_dir / 'ventas/weekly_department.json', sales['weekly_department'])
        dump_json(self.out_dir / 'ventas/monthly_category.json', sales['monthly_category'])
        dump_json(self.out_dir / 'ventas/weekly_category.json', sales['weekly_category'])
        dump_json(self.out_dir / 'ventas/monthly_brand.json', sales['monthly_brand'])
        dump_json(self.out_dir / 'ventas/weekly_brand.json', sales['weekly_brand'])
        dump_json(self.out_dir / 'ventas/monthly_description.json', sales['monthly_description'])
        dump_json(self.out_dir / 'ventas/sales_by_dow.json', sales['sales_by_dow'])
        dump_json(self.out_dir / 'ventas/sales_by_daypart.json', sales['sales_by_daypart'])
        dump_json(self.out_dir / 'ventas/sales_by_hour.json', sales['sales_by_hour'])

        dump_json(self.out_dir / 'delivery/monthly_business.json', delivery['monthly_business'])
        dump_json(self.out_dir / 'delivery/weekly_business.json', delivery['weekly_business'])
        dump_json(self.out_dir / 'delivery/monthly_store_platform.json', delivery['monthly_store_platform'])
        dump_json(self.out_dir / 'delivery/weekly_store_platform.json', delivery['weekly_store_platform'])
        dump_json(self.out_dir / 'delivery/monthly_category.json', delivery['monthly_category'])
        dump_json(self.out_dir / 'delivery/weekly_category.json', delivery['weekly_category'])
        dump_json(self.out_dir / 'delivery/delivery_by_dow.json', delivery['delivery_by_dow'])
        dump_json(self.out_dir / 'delivery/delivery_by_daypart.json', delivery['delivery_by_daypart'])
        dump_json(self.out_dir / 'delivery/delivery_by_hour.json', delivery['delivery_by_hour'])

        dump_json(self.out_dir / 'inventario/store_sku_snapshot.json', inventory['store_sku_snapshot'])
        dump_json(self.out_dir / 'inventario/doi_store_sku.json', inventory['doi_store_sku'])
        dump_json(self.out_dir / 'inventario/shortage_alerts.json', inventory['shortage_alerts'])
        dump_json(self.out_dir / 'inventario/excess_alerts.json', inventory['excess_alerts'])
        dump_json(self.out_dir / 'inventario/transfer_candidates.json', inventory['transfer_candidates'])

        dump_json(self.out_dir / 'merma/monthly_business.json', shrink['monthly_business'])
        dump_json(self.out_dir / 'merma/monthly_store.json', shrink['monthly_store'])
        dump_json(self.out_dir / 'merma/detail_store_description.json', shrink['detail_store_description'])
        dump_json(self.out_dir / 'merma/detail_supplier.json', shrink['detail_supplier'])

        dump_json(self.out_dir / 'sbf/monthly_business.json', sbf['monthly_business'])
        dump_json(self.out_dir / 'sbf/weekly_business.json', sbf['weekly_business'])
        dump_json(self.out_dir / 'sbf/monthly_service_store.json', sbf['monthly_service_store'])
        dump_json(self.out_dir / 'sbf/weekly_service_store.json', sbf['weekly_service_store'])

        dump_json(self.out_dir / 'cxc/summary.json', cxc['summary'])
        dump_json(self.out_dir / 'cxc/detail.json', cxc['detail'])

        dump_json(self.out_dir / 'prd0/supplier_summary.json', prd0['supplier_summary'])
        dump_json(self.out_dir / 'prd0/supplier_detail.json', prd0['supplier_detail'])

        dump_json(self.out_dir / 'hallazgos/sales_signals.json', self.sales_signals)
        dump_json(self.out_dir / 'hallazgos/delivery_signals.json', self.delivery_signals)
        dump_json(self.out_dir / 'hallazgos/inventory_signals.json', self.inventory_signals)
        dump_json(self.out_dir / 'hallazgos/shrink_signals.json', self.shrink_signals)

        dump_json(self.out_dir / 'innovation/monthly_summary.json', innovation['monthly_summary'])
        dump_json(self.out_dir / 'innovation/weekly_summary.json', innovation['weekly_summary'])
        dump_json(self.out_dir / 'innovation/by_store.json', innovation['by_store'])
        dump_json(self.out_dir / 'innovation/by_dow.json', innovation['by_dow'])
        dump_json(self.out_dir / 'innovation/by_daypart.json', innovation['by_daypart'])
        dump_json(self.out_dir / 'innovation/by_hour.json', innovation['by_hour'])

        print('Done:', self.out_dir)


def main() -> None:
    print(f'Workbook source: {DEFAULT_XLSX}')
    exporter = Exporter(DEFAULT_XLSX, OUT_DIR)
    exporter.write_all()


if __name__ == '__main__':
    main()
