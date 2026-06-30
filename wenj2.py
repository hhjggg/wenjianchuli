from openpyxl import load_workbook
import streamlit as st
import pandas as pd
from io import BytesIO

# ========== 网页界面 ==========
st.title("订单SN数据回填工具")
st.subheader("请上传两份Excel文件")

# 文件上传组件
sn_file = st.file_uploader("① 上传售后平台下载数据源表", type=["xlsx"])
target_file = st.file_uploader("② 上传目标表", type=["xlsx"])
if sn_file is not None and target_file is not None:
    st.success("文件读取完成，开始处理数据！")

    # 1、读取两份表格，全部强制为字符串，杜绝科学计数法
    # 替换换行符、删除空格，统一干净列名
    df_sn = pd.read_excel(sn_file, dtype=str).fillna("")
    df_target = pd.read_excel(target_file, dtype=str).fillna("")
    df_sn.columns = [col.replace("\n", "").strip() for col in df_sn.columns]
    df_target.columns = [col.replace("\n", "").strip() for col in df_target.columns]

    # 清洗字段首尾空格，防止匹配失效
    df_sn["三方单号"] = df_sn["三方单号"].str.strip()
    df_sn["sku id"] = df_sn["sku id"].str.strip()
    df_sn["sn"] = df_sn["sn"].str.replace(r"\s+", "", regex=True)
    df_target["京东订单号不要重复除非多单号"] = df_target["京东订单号不要重复除非多单号"].str.replace(r"\s+", "", regex=True)
    
    # 判断是否存在商品名称列，不存在就标记False
    has_goods_col = "商品名称" in df_sn.columns
    if has_goods_col:
        df_sn["商品名称"] = df_sn["商品名称"].str.replace(r"\s+", "", regex=True)


    # 2、构建 订单号+SKU 对应SN的字典
    sn_mapping = {}
    success_pair_list = []

    for _, row in df_sn.iterrows():
        order_id = row["三方单号"]
        sku = row["sku id"]
        sn = row["sn"]
        # 没有商品名称列就赋值为空字符串
        if has_goods_col:
            goods_name = row["商品名称"]
        else:
            goods_name = "kong"
        key = (order_id)
        if key not in sn_mapping:
            sn_mapping[key] = []
        sn_mapping[key].append((sku,goods_name,sn))
 # 3. 新增两列对应原G、H列，初始置空
    # df_target["SN码=内机无SN的备注清楚原因【售后填】"] = None
    # df_target["SN码=外机无SN的备注清楚原因【售后填】"] = None
    
    # 工具函数：直接返回sku+sn配对列表
    def get_sku_sn_pairs(data_list):
        return data_list

    # 单行数据处理核心函数
    def fill_two_sn_cols(row):
        current_key = row["京东订单号不要重复除非多单号"]
        # 订单无匹配数据，返回双空
        if current_key not in sn_mapping:
            return ["", "SN未出，需跟进"]

        data_list = sn_mapping[current_key]
        sn_count = len(data_list)
        # 提取纯SN列表，用于单条/多条计数场景
        sn_only = [item[2] for item in data_list]
        valid_sn_list = [sn for sn in sn_only if sn.strip() != ""]

        if len(valid_sn_list) == 0:
            return ["", "SN未出，需跟进"]

        # 情况3：整个订单仅存在1个有效SN，优先按商品名称匹配内外机（优先级最高）
        if len(valid_sn_list) == 1:
            # 筛选出唯一那条带有有效SN的数据行
            valid_items = next(item for item in data_list if item[2].strip() != "")

            sku_name, goods_name, single_sn = valid_items

            # 规则：商品名称非空 + 包含空调+室内机/室外机，才执行智能匹配
            if goods_name != "" and "空调" in goods_name:
                if "室内机" in goods_name:
                    return ["缺失室外机sn", single_sn]
                elif "室外机" in goods_name:
                    return [single_sn, "缺失室内机sn"]
            # 其余所有情况：商品名称为空 / 字段不全，回归原有旧逻辑（外机空，SN放内机）
            return ["", single_sn]
        
        elif sn_count == 2:
            # 两组纯数字SKU，按数字大小升序排序
            def sort_key(pair):
                sku_raw = str(pair[0]).strip()
                # 空SKU默认数字0，排最前面
                if not sku_raw or not sku_raw.isdigit():
                    return 0
                # SKU纯数字，直接转整数排序
                return int(sku_raw)

            # 按SKU数字升序排序配对
            sorted_pairs = sorted(data_list, key=sort_key)
            # 升序后：靠前=SKU更小(内机)，靠后=SKU更大(外机)
            sn_outer = sorted_pairs[0][2]
            sn_inner = sorted_pairs[1][2]
            return [sn_outer, sn_inner]
        elif sn_count >= 3:
            # 过滤空SN，只统计有效SN
            valid_sn_list = [sn for sn in sn_only if sn.strip() != ""]
            valid_cnt = len(valid_sn_list)
            if valid_cnt == 0:
                # 全部SN为空，返回双空
                return ["", ""]
            else:
                # 存在有效SN，外机列填有效数量，内机空
                return [str(valid_cnt), ""]
        else:
            return ["", ""]

    # 批量回填两列SN
    res = df_target.apply(lambda r: pd.Series(fill_two_sn_cols(r)), axis=1)
    df_target["SN码=外机无SN的备注清楚原因【售后填】"] = res[0]
    df_target["SN码=内机无SN的备注清楚原因【售后填】"] = res[1]
    # 本地调试打印匹配明细
    # print("======所有匹配成功SN明细======")
    # if len(success_pair_list) == 0:
    #     print("暂无匹配成功数据")
    # else:
    #     for idx, pair in enumerate(success_pair_list, start=1):
    #         print(f"第{idx}组配对：{pair}")
    # 4. 生成下载文件，全程文本无科学计数
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
