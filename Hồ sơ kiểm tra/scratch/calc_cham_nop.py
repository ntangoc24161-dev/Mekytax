from datetime import date, datetime

d_start = date(2026, 4, 1)
d_end_old = date(2026, 8, 6)
d_end_new = date(2026, 11, 30)

days_old = (d_end_old - d_start).days
days_new = (d_end_new - d_start).days

b62 = 14023597 # TNDN 2025 truy thu

k62_old = b62 * days_old * 0.0003
k62_new = b62 * days_new * 0.0003

print(f"Old days: {days_old}, old K62: {k62_old:.2f} -> round: {round(k62_old)}")
print(f"New days: {days_new}, new K62: {k62_new:.2f} -> round: {round(k62_new)}")
