"""Smoke test: build an RFS from three Tally TBs with a minimal config.
usage: python tests/smoke_test.py <latest_tb.xlsx> <previous_tb.xlsx> <earliest_tb.xlsx> <out.xlsx>"""
import sys, os, datetime
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from rfs import tally, builder

src = {"FY26": sys.argv[1], "FY25": sys.argv[2], "FY24": sys.argv[3]}
data = {k: tally.read_tb(v) for k, v in src.items()}
for k, v in data.items():
    print(k, "opening, debit, credit, closing of included ledgers:", tally.tb_check(v))
cfg = {"name": "Sample Company Limited", "cin": "U00000XX0000PLC000000", "authorised_capital": 1000000, "face_value": 10,
       "closing_stock": 0, "years": {"d26": datetime.date(2026, 3, 31), "d25": datetime.date(2025, 3, 31), "d24": datetime.date(2024, 3, 31)}}
print("saved", builder.build(cfg, data, src, sys.argv[4]))
