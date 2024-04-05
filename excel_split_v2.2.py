import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import pandas as pd

class ExcelProcessor:
    def __init__(self, master):
        self.master = master
        self.master.title("Excel 数据处理器")
        self.create_widgets()
        self.df = None
        self.filter_groups = []  # 存储筛选组和条件

    def create_widgets(self):
        self.load_button = tk.Button(self.master, text="选择Excel文件", command=self.load_excel)
        self.load_button.pack(pady=20)

    def load_excel(self):
        file_path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls")])
        if file_path:
            self.df = pd.read_excel(file_path)
            self.ask_for_split_column()

    def ask_for_split_column(self):
        columns = self.df.columns.tolist()
        column = simpledialog.askstring("选择列", f"选择要拆分的列（可选项：{', '.join(columns)}）")
        if column in columns:
            self.split_data(column)
        else:
            messagebox.showerror("错误", "选择的列无效，请重新选择。")

    def split_data(self, column):
        self.unique_values = self.df[column].unique()
        self.splitted_dfs = {value: self.df[self.df[column] == value] for value in self.unique_values}
        self.ask_for_filters()

    def ask_for_filters(self):
        response = messagebox.askyesno("筛选数据", "是否需要对数据进行进一步的筛选处理？")
        if response:
            self.filter_window = tk.Toplevel(self.master)
            self.filter_window.title("筛选条件")
            self.build_filter_ui()

    def build_filter_ui(self):
        self.filter_frame = ttk.Frame(self.filter_window)
        self.filter_frame.pack(padx=10, pady=10, fill='x', expand=True)

        self.add_filter_group_button = tk.Button(self.filter_frame, text="添加筛选组", command=self.add_filter_group)
        self.add_filter_group_button.pack(side='top', pady=5)

        self.apply_filter_button = tk.Button(self.filter_frame, text="应用筛选", command=self.apply_filters)
        self.apply_filter_button.pack(side='bottom', pady=5)

    def add_filter_group(self):
        group_frame = ttk.LabelFrame(self.filter_frame, text=f"筛选组 {len(self.filter_groups) + 1}")
        group_frame.pack(fill='x', expand=True, pady=5)

        self.filter_groups.append([])  # 添加一个新的筛选组列表

        # 在每个筛选组内部添加“添加筛选条件”的按钮
        add_condition_button = tk.Button(group_frame, text="添加筛选条件",
                                         command=lambda: self.add_filter_row(group_frame))
        add_condition_button.pack(side='bottom', pady=5)

        self.add_filter_row(group_frame)  # 为新筛选组添加第一个筛选条件

    def add_filter_row(self, group_frame):
        row_frame = ttk.Frame(group_frame)
        row_frame.pack(fill='x', expand=True)

        # 逻辑选择器仅对第一个筛选条件之后的条件添加
        if self.filter_groups[-1]:  # 如果当前筛选组中已有条件，则添加逻辑选择器
            logic_cb = ttk.Combobox(row_frame, values=["AND", "OR"], width=5)
            logic_cb.pack(side='left', padx=5)
            logic_cb.set("AND")  # 默认设置为 AND
        else:
            logic_cb = None  # 第一个条件前不显示逻辑选择器

        column_cb = ttk.Combobox(row_frame, values=self.df.columns.tolist())
        column_cb.pack(side='left', padx=5, expand=True)

        operation_cb = ttk.Combobox(row_frame,
                                    values=["=", "!=", ">", "<", ">=", "<=", "包含", "不包含", "非空", "为空"])
        operation_cb.pack(side='left', padx=5, expand=True)

        value_entry = ttk.Entry(row_frame)
        value_entry.pack(side='left', padx=5, expand=True)

        remove_button = tk.Button(row_frame, text="移除",
                                  command=lambda: self.remove_filter_row(row_frame, group_frame))
        remove_button.pack(side='right', padx=5)

        self.filter_groups[-1].append((logic_cb, column_cb, operation_cb, value_entry))

    def remove_filter_row(self, row_frame, group_frame):
        row_frame.destroy()
        for group in self.filter_groups:
            group[:] = [row for row in group if row[1].master != row_frame]

    def apply_filters(self):
        for value, df in self.splitted_dfs.items():
            with pd.ExcelWriter(f"{value}_filtered.xlsx", engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name="原始数据", index=False)

                for i, group in enumerate(self.filter_groups, start=1):
                    result_df = None

                    for j, (logic_cb, column_cb, operation_cb, value_entry) in enumerate(group):
                        column = column_cb.get()
                        operation = operation_cb.get()
                        raw_value = value_entry.get()

                        condition = self.build_condition(df, column, operation, raw_value)

                        if result_df is None:
                            result_df = df[condition]
                        elif logic_cb and logic_cb.get() == "OR":
                            additional_df = df[condition]
                            result_df = pd.concat([result_df, additional_df]).drop_duplicates().reset_index(drop=True)
                        else:
                            result_df = result_df[condition]

                    if result_df is not None:
                        sheet_name = f"筛选组 {i}"
                        result_df.to_excel(writer, sheet_name=sheet_name, index=False)

                writer.close()

        messagebox.showinfo("完成", "筛选完成，并已保存到新的工作簿中。")

    def build_condition(self, df, column, operation, value):
        if operation == "=":
            return df[column] == value
        elif operation == "!=":
            return df[column] != value
        elif operation == ">":
            return df[column] > value
        elif operation == "<":
            return df[column] < value
        elif operation == ">=":
            return df[column] >= value
        elif operation == "<=":
            return df[column] <= value
        elif operation == "包含":
            return df[column].astype(str).str.contains(value)
        elif operation == "不包含":
            return ~df[column].astype(str).str.contains(value)
        elif operation == "非空":
            return df[column].notna()
        elif operation == "为空":
            return df[column].isna()
        else:
            raise ValueError(f"未知的操作符: {operation}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessor(root)
    root.mainloop()
