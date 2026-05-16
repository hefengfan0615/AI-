import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
from datetime import datetime
import warnings
import locale

# ---------- 1. 抑制 OpenPyXL 警告 ----------
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

# ---------- 2. 修正 Tkinter 中文报错 ----------
try:
    locale.setlocale(locale.LC_CTYPE, 'Chinese')
except:
    try:
        locale.setlocale(locale.LC_CTYPE, 'zh_CN.UTF-8')
    except:
        pass

# ==================== 内置品项映射表（UPC → 品名） ====================
ITEM_MAP = {
    "0040020393821": "多力黄金三益葵花油5.68L",
    "0040020891217": "多力黄金三益葵花油5L",
    "0040020915837": "多力葵花籽油.4L",
    "0690993101001": "多力葵花籽油.1.8L",
    "0690993101005": "多力葵花籽油.5L",
    "0690993121001": "多力花生油5L",
    "0690993121009": "多力花生油900ml",
    "0690993121007": "多力花生油400ml",
    "0690993121011": "多力花生油4L",
    "0690993121015": "O2O多力花生油5.68L",
    "0690993124111": "多力葵花籽油.900ml",
    "0690993124411": "多力葵花籽油.4L",
    "0690993124517": "多力葵花籽油.5.68L",
    "0690993124711": "多力黄金三益葵花油5L",
    "0690993124717": "多力黄金三益葵籽油700ml",
    "0690993124733": "多力黄金三益葵花油5.68L",
    "0690993126415": "多力浓香菜籽油.5L",
    "0690993126611": "多力特香压榨菜籽油5L",
    "0690993146211": "多力芥花油5L",
    "0690993150211": "多力压榨玉米油5L",
    "0690993182001": "多力橄榄葵花油.5L",
    "0690993170209": "多力黄金三益菜籽油",
    "0040020600713": "多力葵花籽油.900ml",
    "0040025047989": "多力花生油400ml",
    "0040020196260": "多力葵花籽油.1.8L",
    "0040020910552": "多力葵花籽油.5.68L",
    "0040025069833": "多力黄金三益菜籽油",
    "0040020911547": "多力葵花籽油.5L",
    "0040020759839": "多力压榨玉米油5L",
    "0690993121003": "多力花生油1.8L",
    "0040020064290": "多力花生油1.8L",
    "0690993170260": "多力黄金三益玉米油700mL",
    "0690993170293": "多力浓香葵花籽油5L",
    "0040020094453": "多力花生油5L",
    "0040025245655": "多力醇香花生油4.8L",
    "0690993170328": "多力醇香花生油4.8L",
    "0040025215388": "多力食用油大单券.",
    "0040025243032": "多力食用油大单券.",
    "0040025249910": "多力食用油大单券.",
    "0040025256998": "多力食用油大单券.",
    "0240000559772": "多力食用油大单券.",
    "0240000563321": "多力食用油大单券.",
    "0240000563902": "多力食用油大单券.",
    "0240000566020": "多力食用油大单券.",
}

def get_item_name(upc):
    return ITEM_MAP.get(str(upc).strip(), f"未知商品({upc})")

# ============ 显示宽度计算工具 ============
def char_display_width(c):
    """返回字符在等宽终端中的显示宽度（英文=1，中文=2）"""
    code = ord(c)
    # 粗略覆盖常见双宽字符范围
    if (0x1100 <= code <= 0x115F or   # Hangul Jamo
        0x2E80 <= code <= 0xA4CF or   # CJK ... Yi
        0xAC00 <= code <= 0xD7A3 or   # Hangul Syllables
        0xF900 <= code <= 0xFAFF or   # CJK Compatibility Ideographs
        0xFE10 <= code <= 0xFE19 or   # Vertical forms
        0xFE30 <= code <= 0xFE6F or   # CJK Compatibility Forms
        0xFF00 <= code <= 0xFF60 or   # Fullwidth Forms
        0xFFE0 <= code <= 0xFFE6 or   # Fullwidth Signs
        0x20000 <= code <= 0x2FFFD or # SIP
        0x30000 <= code <= 0x3FFFD):  # TIP
        return 2
    return 1

def str_display_width(s):
    return sum(char_display_width(ch) for ch in s)

def fit_cell_by_width(text, target_width):
    """截断文本，使显示宽度不超过target_width，超出部分用'...'代替，保证最终宽度为target_width"""
    if str_display_width(text) <= target_width:
        return text
    if target_width <= 3:
        # 宽度太小，简单截断
        return text[:target_width]
    # 预留3个宽度给...
    result = []
    current_width = 0
    for ch in text:
        w = char_display_width(ch)
        if current_width + w > target_width - 3:
            break
        result.append(ch)
        current_width += w
    return ''.join(result) + '...'

def pad_cell(text, target_width, align='left'):
    """将文本用空格填充到显示宽度target_width，align='left'或'right'"""
    current_width = str_display_width(text)
    if current_width >= target_width:
        return text
    padding = ' ' * (target_width - current_width)
    if align == 'left':
        return text + padding
    else:
        return padding + text

# ==================== 文件解析 ====================
def parse_order_metadata(filepath):
    df_raw = pd.read_excel(filepath, header=None, nrows=50, dtype=str)
    range_dates = {}
    header_row = None
    date_pattern = re.compile(
        r"Time Range (\d+) Is Between\s+(\d{2})-(\d{2})-(\d{4})\s+and\s+(\d{2})-(\d{2})-(\d{4})",
        re.IGNORECASE
    )
    for idx, row in df_raw.iterrows():
        if row.isnull().all():
            continue
        row_str = " ".join([str(v) for v in row if pd.notna(v)])
        match = date_pattern.search(row_str)
        if match:
            rn = int(match.group(1))
            start_month = int(match.group(2))
            start_day = int(match.group(3))
            start_year = int(match.group(4))
            end_month = int(match.group(5))
            end_day = int(match.group(6))
            end_year = int(match.group(7))
            try:
                start_date = datetime(start_year, start_month, start_day)
                end_date = datetime(end_year, end_month, end_day)
                range_dates[rn] = {"start": start_date, "end": end_date}
            except ValueError as e:
                raise ValueError(f"日期解析错误 Range {rn}: {e}")
        if "Item Nbr" in row.values:
            header_row = idx
            break
    if header_row is None:
        raise ValueError("未找到数据表头行（缺少 'Item Nbr' 列）")

    df_full = pd.read_excel(filepath, header=header_row, dtype=str)
    range_qty_cols = {}
    range_sales_cols = {}
    for col in df_full.columns:
        qty_match = re.match(r"Range\s*(\d+)\s*POS\s*Qty", col, re.IGNORECASE)
        sales_match = re.match(r"Range\s*(\d+)\s*POS\s*Sales", col, re.IGNORECASE)
        if qty_match:
            rn = int(qty_match.group(1))
            range_qty_cols[rn] = col
        if sales_match:
            rn = int(sales_match.group(1))
            range_sales_cols[rn] = col

    if not range_qty_cols:
        raise ValueError("未找到 Range N POS Qty 列（销量）")
    if not range_sales_cols:
        raise ValueError("未找到 Range N POS Sales 列（销额）")
    return range_dates, header_row, range_qty_cols, range_sales_cols

def load_combined_data(filepath, header_row, range_qty_cols, range_sales_cols):
    df = pd.read_excel(filepath, header=header_row, dtype=str)
    if "UPC" not in df.columns or "Item Desc 1" not in df.columns:
        raise ValueError("订单文件缺少 UPC 或 Item Desc 1 列")
    df["UPC"] = df["UPC"].astype(str).str.strip()

    # 销量
    qty_cols_used = list(range_qty_cols.values())
    df_qty = df[["UPC", "Item Desc 1"] + qty_cols_used].copy()
    for col in qty_cols_used:
        df_qty[col] = pd.to_numeric(df_qty[col], errors="coerce").fillna(0).astype(int)
    grouped_qty = df_qty.groupby("UPC", as_index=False).agg(
        {**{col: "sum" for col in qty_cols_used}, "Item Desc 1": "first"}
    )
    rename_qty = {col: f"Range{rn}_Qty" for rn, col in range_qty_cols.items()}
    grouped_qty.rename(columns=rename_qty, inplace=True)

    # 销额
    sales_cols_used = list(range_sales_cols.values())
    df_sales = df[["UPC"] + sales_cols_used].copy()
    for col in sales_cols_used:
        df_sales[col] = pd.to_numeric(df_sales[col], errors="coerce").fillna(0.0)
    grouped_sales = df_sales.groupby("UPC", as_index=False).agg({col: "sum" for col in sales_cols_used})
    rename_sales = {col: f"Range{rn}_Sales" for rn, col in range_sales_cols.items()}
    grouped_sales.rename(columns=rename_sales, inplace=True)

    combined = grouped_qty.merge(grouped_sales, on="UPC", how="outer")
    return combined

def calc_growth(current, previous):
    if previous is None or previous == 0:
        return "N/A"
    return f"{(current - previous) / previous * 100:.1f}%"

# ==================== 表格构建（等宽对齐） ====================
def build_table(combined_df, range_numbers, metric_type):
    is_qty = (metric_type == 'Qty')
    col_prefix = "Range{}_" + metric_type
    metric_label = "销量" if is_qty else "销额"

    headers = [
        "品项",
        f"日{metric_label}",
        f"月{metric_label}",
        f"去年同月{metric_label}",
        f"{metric_label}月成长率",
        f"今年至今{metric_label}",
        f"去年同期{metric_label}",
        f"{metric_label}YTD成长率"
    ]

    has_day = 1 in range_numbers
    has_month = 2 in range_numbers
    has_ly_month = 3 in range_numbers
    has_ytd = 4 in range_numbers
    has_ly_ytd = 5 in range_numbers

    rows = []
    total = {rn: 0.0 for rn in range_numbers}

    for _, row in combined_df.iterrows():
        upc = row["UPC"]
        item_name = get_item_name(upc)
        desc = row.get("Item Desc 1", "")
        if pd.notna(desc) and desc and "未知商品" in item_name:
            display_name = f"{item_name} ({desc})"
        else:
            display_name = item_name

        day_val = row.get(col_prefix.format(1), 0) if has_day else 0
        month_val = row.get(col_prefix.format(2), 0) if has_month else 0
        ly_month_val = row.get(col_prefix.format(3), 0) if has_ly_month else 0
        ytd_val = row.get(col_prefix.format(4), 0) if has_ytd else 0
        ly_ytd_val = row.get(col_prefix.format(5), 0) if has_ly_ytd else 0

        total[1] = total.get(1, 0) + day_val
        total[2] = total.get(2, 0) + month_val
        total[3] = total.get(3, 0) + ly_month_val
        total[4] = total.get(4, 0) + ytd_val
        total[5] = total.get(5, 0) + ly_ytd_val

        # 格式化每个单元格字符串
        line = [display_name]
        if has_day:
            line.append(f"{int(day_val):,}" if is_qty else f"{day_val:,.2f}")
        else:
            line.append("N/A")
        if has_month:
            line.append(f"{int(month_val):,}" if is_qty else f"{month_val:,.2f}")
        else:
            line.append("N/A")
        if has_ly_month:
            line.append(f"{int(ly_month_val):,}" if is_qty else f"{ly_month_val:,.2f}")
        else:
            line.append("N/A")
        line.append(calc_growth(month_val, ly_month_val) if has_month and has_ly_month else "N/A")
        if has_ytd:
            line.append(f"{int(ytd_val):,}" if is_qty else f"{ytd_val:,.2f}")
        else:
            line.append("N/A")
        if has_ly_ytd:
            line.append(f"{int(ly_ytd_val):,}" if is_qty else f"{ly_ytd_val:,.2f}")
        else:
            line.append("N/A")
        line.append(calc_growth(ytd_val, ly_ytd_val) if has_ytd and has_ly_ytd else "N/A")
        rows.append(line)

    # 总计行
    total_line = ["【总计】"]
    if has_day:
        t_val = total[1]
        total_line.append(f"{int(t_val):,}" if is_qty else f"{t_val:,.2f}")
    else:
        total_line.append("N/A")
    if has_month:
        t_val = total[2]
        total_line.append(f"{int(t_val):,}" if is_qty else f"{t_val:,.2f}")
    else:
        total_line.append("N/A")
    if has_ly_month:
        t_val = total[3]
        total_line.append(f"{int(t_val):,}" if is_qty else f"{t_val:,.2f}")
    else:
        total_line.append("N/A")
    total_line.append(calc_growth(total.get(2,0), total.get(3,0)) if has_month and has_ly_month else "N/A")
    if has_ytd:
        t_val = total[4]
        total_line.append(f"{int(t_val):,}" if is_qty else f"{t_val:,.2f}")
    else:
        total_line.append("N/A")
    if has_ly_ytd:
        t_val = total[5]
        total_line.append(f"{int(t_val):,}" if is_qty else f"{t_val:,.2f}")
    else:
        total_line.append("N/A")
    total_line.append(calc_growth(total.get(4,0), total.get(5,0)) if has_ytd and has_ly_ytd else "N/A")
    rows.append(total_line)

    # 计算每列最大显示宽度
    col_widths = [0] * len(headers)
    for i, h in enumerate(headers):
        col_widths[i] = max(col_widths[i], str_display_width(h))
    for row in rows:
        for i, cell in enumerate(row):
            col_widths[i] = max(col_widths[i], str_display_width(str(cell)))
    # 限制最大宽度，避免单列过宽
    col_widths = [min(w, 25) for w in col_widths]
    col_widths[0] = min(col_widths[0], 30)  # 品名列可稍宽

    # 构建表格行字符串
    table_lines = []
    # 表头
    header_parts = []
    for i, h in enumerate(headers):
        cell = fit_cell_by_width(h, col_widths[i])
        if i == 0:
            header_parts.append(pad_cell(cell, col_widths[i], 'left'))
        else:
            header_parts.append(pad_cell(cell, col_widths[i], 'right'))
    table_lines.append("  ".join(header_parts))
    table_lines.append("-" * (sum(col_widths) + 2*(len(headers)-1)))

    # 数据行
    for row in rows:
        parts = []
        for i, cell in enumerate(row):
            cell_str = str(cell)
            cell_str = fit_cell_by_width(cell_str, col_widths[i])
            if i == 0:
                parts.append(pad_cell(cell_str, col_widths[i], 'left'))
            else:
                parts.append(pad_cell(cell_str, col_widths[i], 'right'))
        table_lines.append("  ".join(parts))
    return table_lines

def analyze_sales(order_file):
    try:
        range_dates, header_row, range_qty_cols, range_sales_cols = parse_order_metadata(order_file)
        combined_df = load_combined_data(order_file, header_row, range_qty_cols, range_sales_cols)
        range_numbers = sorted(set(list(range_qty_cols.keys()) + list(range_sales_cols.keys())))

        output = []
        output.append("=" * 80)
        output.append(f"订单文件：{order_file}")
        if 1 in range_dates:
            sale_date = range_dates[1]["end"]
            output.append(f"销售日期参考：{sale_date.strftime('%Y年%m月%d日')}")
        output.append("=" * 80)
        output.append("")

        output.append("【销量分析（单位：瓶）】")
        output.append("")
        output.extend(build_table(combined_df, range_numbers, 'Qty'))
        output.append("")
        output.append("")

        output.append("【销额分析（单位：元）】")
        output.append("")
        output.extend(build_table(combined_df, range_numbers, 'Sales'))
        output.append("")
        output.append("")

        output.append("【各时间段说明】")
        for rn in sorted(range_dates.keys()):
            start = range_dates[rn]["start"].strftime("%Y-%m-%d")
            end = range_dates[rn]["end"].strftime("%Y-%m-%d")
            output.append(f"  Range{rn}：{start} 至 {end}")

        return "\n".join(output), None
    except Exception as e:
        return None, str(e)

# ==================== GUI 程序 ====================
class SalesAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("hefengfan沃尔玛销售")
        self.root.geometry("950x750")
        self.root.option_add('*Font', ('微软雅黑', 10))

        self.order_path = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.root, text="订单文件（必选）:", font=("微软雅黑", 10)).pack(pady=(10,0), anchor="w", padx=10)
        frame_order = tk.Frame(self.root)
        frame_order.pack(fill="x", padx=10, pady=5)
        tk.Entry(frame_order, textvariable=self.order_path, width=40, font=("微软雅黑", 10)).pack(side="left", fill="x", expand=True)
        tk.Button(frame_order, text="浏览...", command=self.browse_order).pack(side="right", padx=5)

        self.analyze_btn = tk.Button(self.root, text="开始分析", command=self.analyze, width=20, font=("微软雅黑", 10))
        self.analyze_btn.pack(pady=15)

        result_frame = tk.Frame(self.root)
        result_frame.pack(padx=10, pady=5, fill="both", expand=True)

        # 关键：使用等宽字体 Consolas，保证中英文对齐
        self.result_text = tk.Text(result_frame, wrap=tk.NONE, font=('Consolas', 10))
        v_scroll = tk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=v_scroll.set)
        h_scroll = tk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_text.xview)
        self.result_text.configure(xscrollcommand=h_scroll.set)

        self.result_text.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        result_frame.grid_rowconfigure(0, weight=1)
        result_frame.grid_columnconfigure(0, weight=1)

    def browse_order(self):
        fname = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        if fname:
            self.order_path.set(fname)

    def analyze(self):
        order_file = self.order_path.get().strip()
        if not order_file:
            messagebox.showerror("错误", "请先选择订单文件")
            return

        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "正在分析，请稍候...\n")
        self.root.update()

        result, error = analyze_sales(order_file)
        if error:
            messagebox.showerror("分析失败", error)
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"发生错误：\n{error}")
        else:
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, result)

if __name__ == "__main__":
    root = tk.Tk()
    app = SalesAnalyzerGUI(root)
    root.mainloop()
