import numpy as np
import pandas as pd


# 特殊列宽预设
SPECIAL_COLUMNS = {
    4: 20,  # 示例：第5列
    10: 20,  # 示例：第11列
    7: 15  # 新增特殊列
}
ADAPTIVE_CONFIG = {
    'font_name': '微软雅黑',  # 精确匹配Windows系统字体
    'font_size': 11,
    'width_ratio': 1.15,  # 基于字体尺寸的动态比例
    'min_width': 8,  # 最小列宽（对应英文字符数）
    'max_width': 50  # 最大列宽限制
}


def apply_adaptive_width(worksheet, widths):
    for col_idx, raw_width in enumerate(widths):
        adj_width = min(max(raw_width * ADAPTIVE_CONFIG['width_ratio'],
                            ADAPTIVE_CONFIG['min_width']),
                        ADAPTIVE_CONFIG['max_width'])
        worksheet.set_column(col_idx, col_idx, adj_width)


def smart_column_width(df):
    def ch_width_calculator(text):
        char_width = sum(2 if '\u4e00' <= c <= '\u9fff' else 1 for c in str(text))
        return char_width * 1.25 + 3  # 动态比例系数+安全缓冲

    header_widths = df.columns.to_series().apply(ch_width_calculator)
    content_widths = df.astype(str).apply(
        lambda col: col.apply(ch_width_calculator).max()
    )
    return np.maximum(header_widths, content_widths)


def export_excel(df, filename, sheet_name):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        # 初始化工作簿
        workbook = writer.book
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)

        # 创建智能格式
        header_format = workbook.add_format({
            'bold': True,
            'font_name': ADAPTIVE_CONFIG['font_name'],
            'font_size': ADAPTIVE_CONFIG['font_size'],
            'align': 'center',  # 水平居中
            'valign': 'vcenter',  # 垂直居中
            'fg_color': '#D3D3D3',  # 标准灰色（RGB:211,211,211）
            'text_wrap': True,
            'border': 1,
            'border_color': '#808080'  # 增加深灰边框
        })

        # 计算并应用列宽
        worksheet = writer.sheets[sheet_name]
        calculated_widths = smart_column_width(df)
        apply_adaptive_width(worksheet, calculated_widths)

        # 批量写入标题（性能优化版）
        for col_idx, col_name in enumerate(df.columns):
            worksheet.write(0, col_idx, col_name, header_format)

        # 列宽跟踪器初始化
        col_width_tracker = {}
        for col_idx, width in enumerate(calculated_widths):
            worksheet.set_column(col_idx,  col_idx, width)
            col_width_tracker[col_idx] = width

        # 动态调整特殊列（示例）
        special_cols = {4: 20, 10: 25}
        for col, preset_width in special_cols.items():
            current = col_width_tracker.get(col, 0)
            final_width = max(preset_width, current)
            worksheet.set_column(col, col, final_width)
            col_width_tracker[col] = final_width  # 更新跟踪器
