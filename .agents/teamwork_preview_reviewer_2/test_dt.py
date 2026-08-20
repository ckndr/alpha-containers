import datetime

dt = datetime.datetime(2026, 8, 14, 10, 30)
d = datetime.date(2026, 8, 14)

print("isinstance(dt, datetime.date):", isinstance(dt, datetime.date))
print("isinstance(dt, datetime.datetime):", isinstance(dt, datetime.datetime))
print("type(dt) is datetime.date:", type(dt) is datetime.date)

# Broken snippet:
res_broken = dt if isinstance(dt, datetime.date) else dt.date()
print("Broken snippet result:", repr(res_broken), type(res_broken))

# Fixed snippet:
res_fixed = dt.date() if isinstance(dt, datetime.datetime) else dt if isinstance(dt, datetime.date) else None
print("Fixed snippet result:", repr(res_fixed), type(res_fixed))
