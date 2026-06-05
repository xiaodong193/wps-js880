# 基本参数
principal = 2700000
total_periods = 20
annual_rate = 0.041
period_rate = annual_rate / 2  # 半年 2.05%

# 各段还本比例
# 1-6期: 72% -> 1,944,000 -> 每期 324,000
# 7-12期: 17% -> 459,000 -> 每期 76,500
# 13-20期: 11% -> 297,000 -> 每期 37,125

principal_schedule = []
for p in range(1, 21):
    if p <= 6:
        principal_per_period = principal * 0.72 / 6  # 324,000
    elif p <= 12:
        principal_per_period = principal * 0.17 / 6  # 76,500
    else:
        principal_per_period = principal * 0.11 / 8  # 37,125
    principal_schedule.append(principal_per_period)

def irr(cashflows, rate):
    """计算给定折现率的NPV"""
    return sum(cf / (1 + rate) ** t for t, cf in enumerate(cashflows))

def calculate_irr(cashflows):
    """二分法计算IRR"""
    low = -0.5
    high = 0.5
    for _ in range(100):
        mid = (low + high) / 2
        npv = irr(cashflows, mid)
        if abs(npv) < 1e-6:
            return mid
        if npv > 0:
            low = mid
        else:
            high = mid
    return (low + high) / 2

# 生成现金流
# 期初：本金支出 -principal，然后每期收回租金
cashflows = [-principal]

# 计算每期租金
balance = principal
for p in range(1, 21):
    principal_payment = principal_schedule[p-1]
    interest_payment = balance * period_rate
    rent = principal_payment + interest_payment
    cashflows.append(rent)
    balance -= principal_payment

print("=" * 70)
print("经营性租赁租金计算验证")
print("=" * 70)
print(f"本金总额: {principal:,.0f} 元")
print(f"年利率: {annual_rate*100:.2f}%")
print(f"期利率(半年): {period_rate*100:.4f}%")
print(f"总期数: {total_periods} 期 (每半年一期)")
print()

# 打印每期明细
print(f"{'期数':^6} {'期初本金':>14} {'还本':>12} {'利息':>12} {'租金':>12} {'期末本金':>14}")
print("-" * 72)

balance = principal
period_rents = []
for p in range(1, 21):
    pmt = principal_schedule[p-1]
    interest = balance * period_rate
    rent = pmt + interest
    period_rents.append(rent)
    end_balance = balance - pmt
    print(f"{p:^6} {balance:>14,.0f} {pmt:>12,.0f} {interest:>12,.0f} {rent:>12,.2f} {end_balance:>14,.0f}")
    balance = end_balance

print("-" * 72)
total_principal = sum(principal_schedule)
print(f"还本总额: {total_principal:,.0f} 元")
print(f"租金总额: {sum(period_rents):,.2f} 元")
print()

# 计算IRR
irr_rate = calculate_irr(cashflows)
annual_irr = (1 + irr_rate) ** 2 - 1
print("=" * 70)
print("IRR 验证结果")
print("=" * 70)
print(f"期 IRR: {irr_rate*100:.4f}%")
print(f"年化 IRR: {annual_irr*100:.4f}%")
print(f"目标年化 IRR: {annual_rate*100:.2f}%")
print(f"差异: {(annual_irr - annual_rate)*100:.4f}%")

print()
print("=" * 70)
print("各期租金（IRR ≈ 4.10%）")
print("=" * 70)
print(f"{'期数':^6} {'每期租金':>15}")
print("-" * 25)
for p in range(1, 21):
    print(f"{p:^6} {period_rents[p-1]:>15,.2f}")
print("-" * 25)
print(f"{'合计':^6} {sum(period_rents):>15,.2f}")

# 验证分段还本
print()
print("=" * 70)
print("分段还本验证")
print("=" * 70)
print(f"1-6期还本合计: {sum(principal_schedule[0:6]):,.0f} (应为 {principal*0.72:,.0f})")
print(f"7-12期还本合计: {sum(principal_schedule[6:12]):,.0f} (应为 {principal*0.17:,.0f})")
print(f"13-20期还本合计: {sum(principal_schedule[12:20]):,.0f} (应为 {principal*0.11:,.0f})")