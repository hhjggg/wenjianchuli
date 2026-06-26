from openpyxl import load_workbook
import streamlit as st
import pandas as pd
from io import BytesIO

# ========== 网页界面 ==========
st.title("订单SN数据回填工具")
st.subheader("请上传两份Excel文件")

# 文件上传组件
sn_file = st.file_uploader("① 上传SN数据源表", type=["xlsx"])
target_file = st.file_uploader("② 上传待回填目标表", type=["xlsx"])
if sn_file is not None and target_file is not None:
    st.success("文件读取完成，开始处理数据！")

    # 1、读取两份表格，全部强制为字符串，杜绝科学计数法
    df_sn = pd.read_excel(sn_file, dtype=str)
    df_target = pd.read_excel(target_file, dtype=str)

    # 清洗字段首尾空格，防止匹配失效
    df_sn["订单号"] = df_sn["订单号"].str.strip()
    df_sn["SKU"] = df_sn["SKU"].str.strip()
    df_sn["SN"] = df_sn["SN"].str.strip()
    df_target["订单号"] = df_target["订单号"].str.strip()
    df_target["SKU"] = df_target["SKU"].str.strip()

    # 2、构建 订单号+SKU 对应SN的字典
    sn_mapping = {}
    over_three_sn_order = []

    for _, row in df_sn.iterrows():
        order_id = row["订单号"]
        sku = row["SKU"]
        sn = row["SN"]
        key = (order_id, sku)
        if key not in sn_mapping:
            sn_mapping[key] = []
        sn_mapping[key].append(sn)
        # 收集SN超过3个的订单
        if len(sn_mapping[key]) == 3:
            over_three_sn_order.append(order_id)

    # 3、给目标表回填SN
    def get_sn_text(row):
        k = (row["订单号"], row["SKU"])
        if k in sn_mapping:
            return ",".join(sn_mapping[k])
        else:
            return "SN未出，需补齐"

    df_target["SN"] = df_target.apply(get_sn_text, axis=1)

    # 打印匹配信息（网页看不到print，仅本地调试）
    success_pair_list = [k for k, v in sn_mapping.items() if len(v) > 0]
    print("======所有匹配成功的(sku, SN)元组明细======")
    if len(success_pair_list) == 0:
        print("暂无匹配成功的SKU-SN配对数据")
    else:
        for idx, pair in enumerate(success_pair_list, start=1):
            print(f"第{idx}组配对：{pair}")

    # 4、生成下载文件（全程文本，不会变成科学计数）
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_target.to_excel(writer, index=False)
    output.seek(0)

    # 网页下载按钮
    st.download_button(
        label="📥 下载处理完成Excel",
        data=output,
        file_name="回填完成_订单表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("请依次上传两份xlsx格式Excel文件")