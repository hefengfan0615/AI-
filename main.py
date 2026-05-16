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

# ==================== 内置门店字典 ====================
BUILTIN_STORE_DICT = {
    4754: {"type": "门店", "province": "福建", "erp": "山姆沃尔玛福建福州鼓楼店"},
    4788: {"type": "门店", "province": "北京", "erp": "山姆沃尔玛北京石景山店"},
    4801: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东广州番禺店"},
    4803: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东深圳龙岗店"},
    4805: {"type": "门店", "province": "北京", "erp": "山姆沃尔玛北京亦庄店"},
    4806: {"type": "门店", "province": "江苏", "erp": "山姆沃尔玛江苏常州新北店"},
    4807: {"type": "门店", "province": "上海", "erp": "山姆沃尔玛上海浦东店"},
    4808: {"type": "门店", "province": "浙江", "erp": "山姆沃尔玛浙江杭州西溪店"},
    4809: {"type": "门店", "province": "辽宁", "erp": "山姆沃尔玛辽宁大连香炉礁店"},
    4814: {"type": "门店", "province": "湖北", "erp": "山姆沃尔玛湖北武汉汉口店"},
    4818: {"type": "门店", "province": "江苏", "erp": "山姆沃尔玛江苏苏州金鸡湖店"},
    4829: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东珠海香洲店"},
    4830: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东深圳龙华店"},
    4832: {"type": "门店", "province": "江西", "erp": "山姆沃尔玛江西南昌红谷滩店"},
    4833: {"type": "门店", "province": "辽宁", "erp": "山姆沃尔玛辽宁沈阳和平店"},
    4834: {"type": "门店", "province": "江苏", "erp": "山姆沃尔玛江苏苏州木渎店"},
    4835: {"type": "门店", "province": "天津", "erp": "山姆沃尔玛天津梅江店"},
    4836: {"type": "门店", "province": "江苏", "erp": "山姆沃尔玛江苏南京雨花台店"},
    4839: {"type": "门店", "province": "四川", "erp": "山姆沃尔玛四川成都华侨城店"},
    4840: {"type": "门店", "province": "湖南", "erp": "山姆沃尔玛湖南长沙雨花店"},
    4841: {"type": "门店", "province": "福建", "erp": "山姆沃尔玛福建厦门自贸区店"},
    4846: {"type": "门店", "province": "北京", "erp": "山姆沃尔玛北京顺义店"},
    4851: {"type": "门店", "province": "浙江", "erp": "山姆沃尔玛浙江宁波江北店"},
    4852: {"type": "门店", "province": "江苏", "erp": "山姆沃尔玛江苏南通崇川店"},
    4865: {"type": "门店", "province": "上海", "erp": "山姆沃尔玛上海青浦店"},
    6058: {"type": "前置仓", "province": "广东", "erp": "山姆沃尔玛广东珠海前置仓"},
    6158: {"type": "前置仓", "province": "广东", "erp": "山姆沃尔玛广东广州前置仓"},
    6258: {"type": "前置仓", "province": "北京", "erp": "山姆沃尔玛广东惠州6258前置仓"},
    6259: {"type": "前置仓", "province": "天津", "erp": "山姆沃尔玛广东惠州6259前置仓"},
    6505: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东深圳福田店"},
    6507: {"type": "门店", "province": "福建", "erp": "山姆沃尔玛福建福州仓山店"},
    6508: {"type": "门店", "province": "重庆", "erp": "山姆沃尔玛重庆九龙坡店"},
    6509: {"type": "门店", "province": "江苏", "erp": "山姆沃尔玛江苏苏州昆山店"},
    6510: {"type": "门店", "province": "江苏", "erp": "山姆沃尔玛江苏南京江北店"},
    6511: {"type": "门店", "province": "四川", "erp": "山姆沃尔玛四川成都店"},
    6512: {"type": "门店", "province": "湖北", "erp": "山姆沃尔玛湖北武汉光谷店"},
    6513: {"type": "门店", "province": "重庆", "erp": "山姆沃尔玛重庆两江新区店"},
    6515: {"type": "门店", "province": "浙江", "erp": "山姆沃尔玛浙江杭州萧山店"},
    6516: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东惠州店"},
    6517: {"type": "门店", "province": "湖南", "erp": "山姆沃尔玛湖南长沙岳麓店"},
    6518: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东广州东店"},
    6519: {"type": "门店", "province": "北京", "erp": "山姆沃尔玛北京大兴店"},
    6523: {"type": "门店", "province": "浙江", "erp": "山姆沃尔玛浙江杭州北店"},
    6526: {"type": "门店", "province": "湖北", "erp": "山姆沃尔玛湖北武汉江岸后湖店"},
    6529: {"type": "门店", "province": "浙江", "erp": "山姆沃尔玛浙江宁波鄞州店"},
    6532: {"type": "门店", "province": "广西", "erp": "山姆沃尔玛广西南宁店"},
    6533: {"type": "门店", "province": "四川", "erp": "山姆沃尔玛四川成都武侯店"},
    6538: {"type": "门店", "province": "上海", "erp": "山姆沃尔玛上海嘉定店"},
    6551: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东深圳宝安店"},
    6557: {"type": "门店", "province": "浙江", "erp": "山姆沃尔玛浙江绍兴店"},
    6558: {"type": "前置仓", "province": "上海", "erp": "山姆沃尔玛上海全渠道履约中心"},
    6559: {"type": "履约中心", "province": "江苏", "erp": "山姆沃尔玛广东惠州履约中心6559"},
    6560: {"type": "门店", "province": "江苏", "erp": "山姆沃尔玛江苏无锡店"},
    6567: {"type": "门店", "province": "湖北", "erp": "山姆沃尔玛湖北武汉汉阳店"},
    6568: {"type": "门店", "province": "上海", "erp": "山姆沃尔玛上海真如店"},
    6569: {"type": "门店", "province": "浙江", "erp": "山姆沃尔玛浙江嘉兴南湖店"},
    6570: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东东莞店"},
    6571: {"type": "门店", "province": "福建", "erp": "山姆沃尔玛福建泉州晋江店"},
    6580: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东深圳前海店"},
    6582: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东佛山顺德店"},
    6583: {"type": "门店", "province": "浙江", "erp": "山姆沃尔玛浙江温州店"},
    6587: {"type": "门店", "province": "上海", "erp": "山姆沃尔玛上海浦东东店"},
    6590: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东广州荔湾店"},
    6599: {"type": "门店", "province": "北京", "erp": "山姆沃尔玛北京昌平店"},
    6600: {"type": "前置仓", "province": "天津", "erp": "山姆沃尔玛广东惠州6600前置仓"},
    6601: {"type": "前置仓", "province": "天津", "erp": "山姆沃尔玛广东惠州6601前置仓"},
    6602: {"type": "前置仓", "province": "广东", "erp": "山姆沃尔玛广东广州6602前置仓（次级）"},
    6660: {"type": "京东VC", "province": "广东", "erp": "山姆沃尔玛广东广州6660京东仓"},
    6661: {"type": "京东VC", "province": "辽宁", "erp": "山姆沃尔玛广东广州6661京东仓"},
    6662: {"type": "京东VC", "province": "湖北", "erp": "山姆沃尔玛广东广州6662京东仓"},
    6663: {"type": "履约中心", "province": "天津", "erp": "山姆沃尔玛北京虚拟店"},
    6664: {"type": "京东VC", "province": "四川", "erp": "山姆沃尔玛广东广州6664京东仓"},
    6665: {"type": "京东VC", "province": "上海", "erp": "山姆沃尔玛广东广州6665京东仓"},
    6668: {"type": "京东VC", "province": "北京", "erp": "山姆沃尔玛广东广州6668京东仓"},
    6682: {"type": "前置仓", "province": "上海", "erp": "山姆沃尔玛上海云仓"},
    6758: {"type": "履约中心", "province": "广东", "erp": "山姆沃尔玛广东深圳电商拣货中心"},
    6852: {"type": "履约中心", "province": "辽宁", "erp": "山姆沃尔玛广东惠州6852履约中心"},
    6855: {"type": "履约中心", "province": "湖南", "erp": "山姆沃尔玛广东惠州6855履约中心"},
    6856: {"type": "履约中心", "province": "四川", "erp": "山姆沃尔玛广东惠州西南履约中心"},
    6858: {"type": "门店", "province": "上海", "erp": "山姆沃尔玛上海宝山店"},
    6865: {"type": "前置仓", "province": "辽宁", "erp": "山姆沃尔玛广东惠州6865前置仓"},
    6866: {"type": "前置仓", "province": "广东", "erp": "山姆沃尔玛广东惠州6866前置仓"},
    6886: {"type": "前置仓", "province": "江西", "erp": "山姆沃尔玛广东惠州6886前置仓"},
    6887: {"type": "前置仓", "province": "辽宁", "erp": "山姆沃尔玛广东惠州6887前置仓"},
    6902: {"type": "门店", "province": "安徽", "erp": "山姆沃尔玛安徽合肥店"},
    6905: {"type": "门店", "province": "江苏", "erp": "山姆沃尔玛江苏扬州店"},
    6908: {"type": "门店", "province": "江苏", "erp": "山姆沃尔玛江苏张家港店"},
    6909: {"type": "门店", "province": "江苏", "erp": "山姆沃尔玛江苏无锡惠山店"},
    6911: {"type": "门店", "province": "广东", "erp": "山姆沃尔玛广东中山石歧店"},
    6912: {"type": "门店", "province": "山东", "erp": "山姆沃尔玛山东青岛店"},
    6917: {"type": "门店", "province": "山东", "erp": "山姆沃尔玛山东济南店"},
    6988: {"type": "门店", "province": "上海", "erp": "山姆沃尔玛上海外高桥店"},
}

def get_store_info(club_nbr):
    info = BUILTIN_STORE_DICT.get(club_nbr)
    if info:
        return (info["type"], info["province"], info["erp"])
    return ("未知", "未知", "未知")

def parse_order_metadata(filepath):
    """解析订单文件，返回 (range_dates, header_row)"""
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
        if row_str.strip().startswith("Item Nbr"):
            header_row = idx
            break
    if header_row is None:
        raise ValueError("未找到数据表头行（Item Nbr）")
    return range_dates, header_row

def load_order_data(filepath, header_row):
    df = pd.read_excel(filepath, header=header_row, dtype=str)
    range_cols = {}
    for col in df.columns:
        match = re.match(r"Range\s*(\d+)\s*Total Units Sold", col, re.IGNORECASE)
        if match:
            rn = int(match.group(1))
            range_cols[rn] = col
    if not range_cols:
        raise ValueError("未找到任何 Range X Total Units Sold 列")
    if "Club Nbr" not in df.columns:
        raise ValueError("订单文件缺少 Club Nbr 列")
    df["Club Nbr"] = df["Club Nbr"].astype(str).str.extract(r"(\d+)")[0].astype(float).astype(int)
    keep_cols = ["Club Nbr"] + list(range_cols.values())
    df_clean = df[keep_cols].copy()
    for col in range_cols.values():
        df_clean[col] = pd.to_numeric(df_clean[col], errors="coerce").fillna(0)
    agg_dict = {col: "sum" for col in range_cols.values()}
    grouped = df_clean.groupby("Club Nbr", as_index=False).agg(agg_dict)
    rename_map = {col: f"Range{rn}_Units" for rn, col in range_cols.items()}
    grouped.rename(columns=rename_map, inplace=True)
    return grouped

def analyze_sales(order_file):
    try:
        range_dates, header_row = parse_order_metadata(order_file)
        if 1 not in range_dates:
            raise ValueError("未找到 Range 1 的时间范围，无法确定销售日期")
        sale_date = range_dates[1]["end"]
        sale_date_str = sale_date.strftime("%Y年%m月%d日")

        order_df = load_order_data(order_file, header_row)

        total_units = {}
        for col in order_df.columns:
            if col.startswith("Range") and col.endswith("_Units"):
                rn = int(col.split("Range")[1].split("_")[0])
                total_units[rn] = order_df[col].sum()

        day_sales = total_units.get(1, 0)

        if day_sales == 0:
            best_store_desc = "无销售数据"
            best_units = 0
        else:
            store_sales = order_df[["Club Nbr", "Range1_Units"]].copy()
            max_row = store_sales.loc[store_sales["Range1_Units"].idxmax()]
            best_club = max_row["Club Nbr"]
            best_units = max_row["Range1_Units"]
            store_type, province, erp = get_store_info(best_club)
            best_store_desc = f"{store_type}（{province}）{erp}"

        def growth_rate(current, previous):
            if previous == 0:
                return float('inf') if current > 0 else 0
            return (current - previous) / previous * 100

        month_growth = growth_rate(total_units.get(2, 0), total_units.get(3, 0))
        year_growth = growth_rate(total_units.get(4, 0), total_units.get(5, 0))

        # 构建结果字符串
        result_lines = []
        result_lines.append("=" * 50)
        result_lines.append(f"订单文件：{order_file}")
        result_lines.append("门店信息：内置字典")
        result_lines.append("=" * 50)
        result_lines.append("")
        result_lines.append(f"销售日期：{sale_date_str}  总销量：{day_sales:,.0f} 瓶")
        result_lines.append("")
        if best_units > 0:
            result_lines.append(f"最高销售门店：{best_store_desc}")
            result_lines.append(f"销售瓶数：{best_units:,.0f} 瓶")
        else:
            result_lines.append("无单日销售数据，无法确定最高销售门店。")
        result_lines.append("")
        result_lines.append(f"较去年同月份销量成长率：{month_growth:+.2f}% （对比 Range2 vs Range3）")
        result_lines.append(f"较去年同年份销量成长率：{year_growth:+.2f}% （对比 Range4 vs Range5）")
        result_lines.append("")
        result_lines.append("【各时间段销量与周期】")
        # 获取所有存在的 Range 编号（从 total_units 的键和 range_dates 的键合并去重）
        all_ranges = sorted(set(total_units.keys()) | set(range_dates.keys()))
        for rn in all_ranges:
            units = total_units.get(rn, 0)
            # 获取时间范围描述
            if rn in range_dates:
                start = range_dates[rn]["start"].strftime("%Y-%m-%d")
                end = range_dates[rn]["end"].strftime("%Y-%m-%d")
                period = f"{start} 至 {end}"
            else:
                period = "时间范围未知"
            result_lines.append(f"  Range{rn}：{units:,.0f} 瓶  （{period}）")
        return "\n".join(result_lines), None
    except Exception as e:
        return None, str(e)

# ==================== GUI 程序 ====================
class SalesAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("hefengfan山姆销售")
        self.root.geometry("400x500")
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

        self.result_text = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, width=80, height=25, font=("微软雅黑", 10))
        self.result_text.pack(padx=10, pady=5, fill="both", expand=True)

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
