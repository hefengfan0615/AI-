import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
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
    """根据UPC返回品名，若找不到则返回原UPC"""
    return ITEM_MAP.get(str(upc).strip(), f"未知商品({upc})")

def parse_order_metadata(filepath):
    """解析订单文件，返回 (range_dates, header_row, range_sales_columns)"""
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
        # 检查该行是否包含列名 "Item Nbr"
        if "Item Nbr" in row.values:
            header_row = idx
            break
    if header_row is None:
        raise ValueError("未找到数据表头行（缺少 'Item Nbr' 列）")

    # 读取完整数据，获取所有 Range N POS Sales 列
    df_full = pd.read_excel(filepath, header=header_row, dtype=str)
    range_sales_cols = {}
    for col in df_full.columns:
        match = re.match(r"Range\s*(\d+)\s*POS\s*Sales", col, re.IGNORECASE)
        if match:
            rn = int(match.group(1))
            range_sales_cols[rn] = col

    if not range_sales_cols:
        raise ValueError("未找到任何 Range N POS Sales 列")
    return range_dates, header_row, range_sales_cols

def load_order_data(filepath, header_row, range_sales_cols):
    """读取订单数据，按UPC聚合各Range销售额"""
    df = pd.read_excel(filepath, header=header_row, dtype=str)

    # 必需列检查
    if "UPC" not in df.columns:
        raise ValueError("订单文件缺少 UPC 列")
    if "Item Desc 1" not in df.columns:
        raise ValueError("订单文件缺少 Item Desc 1 列")

    # 清洗UPC（去除小数点，转为字符串）
    df["UPC"] = df["UPC"].astype(str).str.strip()
    # 保留需要的列
    keep_cols = ["UPC", "Item Desc 1"] + list(range_sales_cols.values())
    df_clean = df[keep_cols].copy()

    # 转换销售额为数值
    for col in range_sales_cols.values():
        df_clean[col] = pd.to_numeric(df_clean[col], errors="coerce").fillna(0)

    # 按UPC分组聚合
    agg_dict = {col: "sum" for col in range_sales_cols.values()}
    agg_dict["Item Desc 1"] = "first"
    grouped = df_clean.groupby("UPC", as_index=False).agg(agg_dict)

    # 重命名列为 RangeX_Sales
    rename_map = {col: f"Range{rn}_Sales" for rn, col in range_sales_cols.items()}
    grouped.rename(columns=rename_map, inplace=True)
    return grouped

def analyze_sales(order_file):
    """主分析函数：返回结果字符串和错误信息"""
    try:
        range_dates, header_row, range_sales_cols = parse_order_metadata(order_file)
        if not range_sales_cols:
            raise ValueError("未找到销售额数据列")

        # 按UPC聚合数据
        sales_df = load_order_data(order_file, header_row, range_sales_cols)

        # 获取所有Range编号并排序
        range_numbers = sorted(range_sales_cols.keys())
        metric_names = {
            0: "日销额",
            1: "月销额",
            2: "去年同月销额",
            3: "今年至今销额",
            4: "去年同期销额"
        }

        # 准备输出行
        result_lines = []
        result_lines.append("=" * 80)
        result_lines.append(f"订单文件：{order_file}")
        result_lines.append("=" * 80)
        result_lines.append("")

        # 显示日期信息（若有Range1日期）
        if 1 in range_dates:
            sale_date = range_dates[1]["end"]
            result_lines.append(f"销售日期参考：{sale_date.strftime('%Y年%m月%d日')}")
        result_lines.append("")

        # 构建表头
        available_metrics = []
        for i, rn in enumerate(range_numbers):
            if i < len(metric_names):
                available_metrics.append(metric_names[i])
        show_ytd = len(range_numbers) >= 4
        show_ly_ytd = len(range_numbers) >= 5

        display_headers = ["品项"]
        display_headers.append(available_metrics[0] if len(available_metrics) > 0 else "日销额")
        display_headers.append(available_metrics[1] if len(available_metrics) > 1 else "月销额")
        display_headers.append(available_metrics[2] if len(available_metrics) > 2 else "去年同月销额")
        if show_ytd:
            display_headers.append("今年至今销额")
        if show_ly_ytd:
            display_headers.append("去年同期销额")

        # 准备数据行
        rows = []
        total_sales = {rn: 0.0 for rn in range_numbers}
        for _, row in sales_df.iterrows():
            upc = row["UPC"]
            item_name = get_item_name(upc)
            desc = row["Item Desc 1"] if pd.notna(row["Item Desc 1"]) else ""
            display_name = f"{item_name}" + (f" ({desc})" if desc and "未知商品" in item_name else "")
            sales_vals = []
            for rn in range_numbers:
                val = row.get(f"Range{rn}_Sales", 0)
                sales_vals.append(val)
                total_sales[rn] += val

            row_data = [display_name]
            row_data.append(f"{sales_vals[0]:,.2f}" if len(sales_vals) > 0 else "N/A")
            row_data.append(f"{sales_vals[1]:,.2f}" if len(sales_vals) > 1 else "N/A")
            row_data.append(f"{sales_vals[2]:,.2f}" if len(sales_vals) > 2 else "N/A")
            if show_ytd and len(sales_vals) > 3:
                row_data.append(f"{sales_vals[3]:,.2f}")
            elif show_ytd:
                row_data.append("N/A")
            if show_ly_ytd and len(sales_vals) > 4:
                row_data.append(f"{sales_vals[4]:,.2f}")
            elif show_ly_ytd:
                row_data.append("N/A")
            rows.append(row_data)

        # 总计行
        total_row = ["【总计】"]
        total_row.append(f"{total_sales.get(range_numbers[0], 0):,.2f}" if len(range_numbers) > 0 else "N/A")
        total_row.append(f"{total_sales.get(range_numbers[1], 0):,.2f}" if len(range_numbers) > 1 else "N/A")
        total_row.append(f"{total_sales.get(range_numbers[2], 0):,.2f}" if len(range_numbers) > 2 else "N/A")
        if show_ytd and len(range_numbers) > 3:
            total_row.append(f"{total_sales.get(range_numbers[3], 0):,.2f}")
        elif show_ytd:
            total_row.append("N/A")
        if show_ly_ytd and len(range_numbers) > 4:
            total_row.append(f"{total_sales.get(range_numbers[4], 0):,.2f}")
        elif show_ly_ytd:
            total_row.append("N/A")
        rows.append(total_row)

        # 计算每列最大宽度（考虑表头和数据）
        col_widths = []
        for i, header in enumerate(display_headers):
            max_len = len(header)
            for row in rows:
                cell = str(row[i])
                max_len = max(max_len, len(cell))
            # 最大列宽限制为30，与之前一致
            col_widths.append(min(max_len, 30))

        # 格式化输出（品名列左对齐，数字列右对齐），超长部分截断加省略号
        def fit_cell(text, width):
            """将文本截断到指定宽度，若超出则末尾加'...'，保证最终长度为width"""
            if len(text) <= width:
                return text
            if width <= 3:
                return text[:width]
            return text[:width-3] + "..."

        # 输出表头
        header_parts = []
        for i, header in enumerate(display_headers):
            cell = fit_cell(header, col_widths[i])
            if i == 0:
                header_parts.append(cell.ljust(col_widths[i]))
            else:
                header_parts.append(cell.rjust(col_widths[i]))
        header_line = "  ".join(header_parts)
        result_lines.append(header_line)
        result_lines.append("-" * len(header_line))

        # 输出数据行
        for row in rows:
            parts = []
            for i, cell in enumerate(row):
                cell_str = str(cell)
                cell_str = fit_cell(cell_str, col_widths[i])
                if i == 0:
                    parts.append(cell_str.ljust(col_widths[i]))
                else:
                    parts.append(cell_str.rjust(col_widths[i]))
            line = "  ".join(parts)
            result_lines.append(line)

        # 附加各Range的时间范围说明
        result_lines.append("")
        result_lines.append("【各时间段说明】")
        for rn in range_numbers:
            if rn in range_dates:
                start = range_dates[rn]["start"].strftime("%Y-%m-%d")
                end = range_dates[rn]["end"].strftime("%Y-%m-%d")
                result_lines.append(f"  Range{rn}：{start} 至 {end}")
            else:
                result_lines.append(f"  Range{rn}：时间范围未知")

        return "\n".join(result_lines), None

    except Exception as e:
        return None, str(e)

# ==================== GUI 程序 ====================
class SalesAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("hefengfan沃尔玛销售")
        self.root.geometry("500x600")
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

        # 使用 Frame 容纳 Text 和两个滚动条
        result_frame = tk.Frame(self.root)
        result_frame.pack(padx=10, pady=5, fill="both", expand=True)

        self.result_text = tk.Text(result_frame, wrap=tk.NONE, font=("微软雅黑", 10))
        # 垂直滚动条
        v_scroll = tk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=v_scroll.set)
        # 水平滚动条
        h_scroll = tk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.result_text.xview)
        self.result_text.configure(xscrollcommand=h_scroll.set)

        # 布局
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
