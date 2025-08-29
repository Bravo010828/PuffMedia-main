import pandas as pd

file1 = "file1.xlsx"
file2 = "file2.xlsx"
output = "merged_output.xlsx"

# 读两张表
df1 = pd.read_excel(file1)
df2 = pd.read_excel(file2)

# 保证第一列是名字
df1 = df1.rename(columns={df1.columns[0]: "Name"})
df2 = df2.rename(columns={df2.columns[0]: "Name"})

# 复制一份结果表
merged = df1.copy()

# 普通列：逐行相加（跳过特殊列）
special_cols = ["Avg. order value", "CTR", "Avg. affiliate customers"]
for col in df1.columns:
    if col not in ["Name"] + special_cols:
        merged[col] = pd.to_numeric(df1[col], errors="coerce").fillna(0) + \
                      pd.to_numeric(df2[col], errors="coerce").fillna(0)

# ========== 三个特殊列用公式算 ==========
# 1) AOV = GMV / Items sold
if {"Affiliate GMV", "Items sold"} <= set(merged.columns):
    merged["Avg. order value"] = merged["Affiliate GMV"] / merged["Items sold"]

# 2) CTR = Affiliate orders / Product impressions
if {"Affiliate orders", "Product impressions"} <= set(merged.columns):
    merged["CTR"] = merged["Affiliate orders"] / merged["Product impressions"]

# 3) Avg. affiliate customers = Affiliate orders / Affiliate customers（如果有）
if {"Affiliate orders", "Affiliate customers"} <= set(merged.columns):
    merged["Avg. affiliate customers"] = merged["Affiliate orders"] / merged["Affiliate customers"]
else:
    # 如果缺列，回退到两表原值平均
    if "Avg. affiliate customers" in df1.columns:
        merged["Avg. affiliate customers"] = (
            pd.to_numeric(df1["Avg. affiliate customers"], errors="coerce").fillna(0) +
            pd.to_numeric(df2["Avg. affiliate customers"], errors="coerce").fillna(0)
        ) / 2

# 输出
merged.to_excel(output, index=False)
print(f"✅ 合并完成：{output}")
