"""Pricing engine: the single place where Offer amounts are computed (ADR 0006).

Usage as a library:
    from pricing import World, compute_offer
    w = World()
    w.price("30.120.01", 2026)          -> unit price of an article in a given year
    w.rate("fitter_a", 2023)            -> hourly rate in a given year
    compute_offer(lines, date="2026-07-07", discount=0.02, skonto=True)

A line is a dict: {code?, text, qty, unit, unit_price?, optional?: bool, client_provision?: bool}
If unit_price is missing it is taken from the Item master for the Offer year.
"""
import csv, datetime as dt
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import yaml

ROOT = Path(__file__).resolve().parents[2]

def D(x):
    return Decimal(str(x))

def round_to(x, step):
    step = D(step)
    return (D(x) / step).quantize(Decimal("1"), rounding=ROUND_HALF_UP) * step

class World:
    def __init__(self):
        self.ops = yaml.safe_load(open(ROOT / "generation/world/operations.yaml", encoding="utf-8"))
        with open(ROOT / "generation/world/item_master.csv", encoding="utf-8") as f:
            self.items = {r["code"]: r for r in csv.DictReader(f)}
        pi = self.ops["price_index"]
        self.index = {fam: dict(zip(pi["years"], vals)) for fam, vals in pi.items() if fam != "years"}

    def factor(self, family, year):
        years = sorted(self.index[family])
        year = min(max(year, years[0]), years[-1])
        return D(self.index[family][year])

    def price(self, code, year):
        it = self.items[code]
        base = D(it["vk"])
        if year == self.ops["base_year"]:
            return base
        if it["family"] == "system_products":
            return round_to(base * self.factor(it["family"], year), 10)
        if it["family"] == "labour" and it["unit"] == "Std":
            return round_to(base * self.factor("labour", year), 1)
        v = base * self.factor(it["family"], year)       # older list prices are round values
        return round_to(v, "0.05" if v < 10 else "0.50" if v < 100 else "5" if v < 1000 else "10")

    def rate(self, category, year):
        base = D(self.ops["rate_categories"][category]["rate"])
        return round_to(base * self.factor("labour", year), 1)

    def vat_rate(self, date):
        d = dt.date.fromisoformat(str(date))
        for v in self.ops["vat"]:
            lo = dt.date.fromisoformat(str(v["from"])); hi = v["to"] and dt.date.fromisoformat(str(v["to"]))
            if lo <= d and (not hi or d <= hi):
                return D(v["rate"])
        raise ValueError(f"no VAT rate for {date}")

def compute_offer(lines, date, discount=0.0, skonto=False, world=None):
    """Return line totals and the totals block exactly as Fronalp prints them."""
    w = world or World()
    year = dt.date.fromisoformat(str(date)).year
    rules = w.ops
    out, gross, optional_sum = [], D(0), D(0)
    for ln in lines:
        up = D(ln["unit_price"]) if ln.get("unit_price") is not None else w.price(ln["code"], year)
        total = round_to(D(ln["qty"]) * up, rules["rounding"]["line_total"])
        if ln.get("client_provision"):
            total = D(0)
        row = dict(ln, unit_price=up, total=total)
        out.append(row)
        if ln.get("optional"):
            optional_sum += total
        else:
            gross += total
    disc = round_to(gross * D(discount), rules["rounding"]["discount"])
    sk = round_to(gross * D(rules["commercial_rules"]["skonto"]["rate"]), rules["rounding"]["discount"]) if skonto else D(0)  # on gross: company habit
    net = gross - disc - sk
    vat_rate = w.vat_rate(date)
    vat = round_to(net * vat_rate, rules["rounding"]["vat"])
    total = round_to(net + vat, rules["rounding"]["grand_total"])
    return {"lines": out, "gross": gross, "discount_rate": D(discount), "discount": disc,
            "skonto_rate": D(rules["commercial_rules"]["skonto"]["rate"]) if skonto else D(0), "skonto": sk,
            "net": net, "vat_rate": vat_rate, "vat": vat, "total": total, "optional_sum": optional_sum,
            "needs_owner_approval": net > D(rules["commercial_rules"]["approval_owner_above_net"])
                                    or D(discount) > D(rules["commercial_rules"]["approval_owner_discount_above"])}

if __name__ == "__main__":   # self test with invented numbers
    w = World()
    lines = [{"code": "10.100.01", "qty": 1}, {"code": "30.120.01", "qty": 52},
             {"code": "20.110.01", "qty": 48, "optional": True}, {"text": "Strom bauseits", "qty": 1, "unit_price": 0, "client_provision": True}]
    r = compute_offer(lines, "2026-05-04", discount=0.02, skonto=True, world=w)
    assert r["gross"] == D("2686.00"), r["gross"]          # 450 + 52 x 43
    assert r["discount"] == D("53.70") and r["skonto"] == D("53.70")   # 53.72 -> 5 Rappen
    assert r["vat"] == D("208.85"), r["vat"]                # 2578.60 x 8.1 % = 208.87 -> 208.85
    assert r["total"] == D("2787.45"), r["total"]
    assert r["optional_sum"] == D("1680.00")
    assert w.vat_rate("2023-12-31") == D("0.077") and w.vat_rate("2024-01-01") == D("0.081")
    assert w.rate("fitter_a", 2026) == 125 and w.rate("fitter_a", 2020) == 115
    print("pricing self test ok;", "A fitter 2018 to 2026:", [int(w.rate("fitter_a", y)) for y in range(2018, 2027)])
