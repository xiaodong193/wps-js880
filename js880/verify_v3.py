principal = 2700000
annual_rate = 0.041
period_rate = annual_rate / 2  # 2.05%

# 各段每期租金（固定）
R1 = 362745.00
R2 = 88077.38
R3 = 40955.28

# 每段还本总额
p1_total = principal * 0.72  # 1,944,000
p2_total = principal * 0.17  # 459,000
p3_total = principal * 0.11  # 297,000

def calc_irr(cashflows):
    def npv(r):
        return sum(cf / (1 + r) ** t for t, cf in enumerate(cashflows))
    low, high = -0.5, 0.5
    for _ in range(100):
        mid = (low + high) / 2
        if abs(npv(mid)) < 1e-6:
            return mid
        if npv(mid) > 0:
            low = mid
        else:
            high = mid
    return (low + high) / 2

# 计算每期明细
cashflows = [-principal]
balance = principal
total_principal_paid = 0

print("=" * 75)
print("各阶段租金一致方案 - 详细计算")
print("=" * 75)
print(f"{'期数':^6} {'期初本金':>14} {'利息':>12} {'每期租金':>14} {'还本金额':>12} {'期末本金':>14}")
print("-" * 75)

for p in range(1, 21):
    if p <= 6:
        rent = R1
    elif p <= 12:
        rent = R2
    else:
        rent = R3

    interest = balance * period_rate
    principal_paid = rent - interest
    end_balance = balance - principal_paid
    total_principal_paid += principal_paid
    cashflows.append(rent)

    print(f"{p:^6} {balance:>14,.0f} {interest:>12,.2f} {rent:>14,.2f} {principal_paid:>12,.2f} {end_balance:>14,.2f}")
    balance = end_balance

print("-" * 75)

# 验证分段还本
print(f"\n【分段还本验证】")
print(f"1-6期还本合计: {sum(cashflows[1:7]) - sum([(2700000 - (2700000 * 0.72 * i/6) if i>0 else 2700000) * period_rate for i in range(6)]):,.2f}")
# 简化验证
p1_principal = 0
b = 2700000
for p in range(1, 7):
    rent = R1
    interest = b * period_rate
    principal_paid = rent - interest
    p1_principal += principal_paid
    b -= principal_paid

p2_principal = 0
for p in range(7, 13):
    rent = R2
    interest = b * period_rate
    principal_paid = rent - interest
    p2_principal += principal_paid
    b -= principal_paid

p3_principal = 0
for p in range(13, 21):
    rent = R3
    interest = b * period_rate
    principal_paid = rent - interest
    p3_principal += principal_paid
    b -= principal_paid

print(f"1-6期还本合计: {p1_principal:,.2f} (目标: {p1_total:,.0f})")
print(f"7-12期还本合计: {p2_principal:,.2f} (目标: {p2_total:,.0f})")
print(f"13-20期还本合计: {p3_principal:,.2f} (目标: {p3_total:,.0f})")
print(f"总还本: {p1_principal+p2_principal+p3_principal:,.2f} (目标: {principal:,.0f})")

# IRR
irr_rate = calc_irr(cashflows)
annual_irr = (1 + irr_rate) ** 2 - 1

print(f"\n【IRR验证】")
print(f"期 IRR: {irr_rate*100:.6f}%")
print(f"年化 IRR: {annual_irr*100:.6f}%")
print(f"目标 IRR: {annual_rate*100:.2f}%")
print(f"差异: {(annual_irr - annual_rate)*100:.6f}%")