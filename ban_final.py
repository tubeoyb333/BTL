import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, filedialog
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import tkinter.filedialog
import sympy as sp
import numpy as np
import pandas as pd
import os
from docx import Document

class MathExpression:
    def __init__(self, expression_str):
        self.expression_str = expression_str
        self.x = sp.symbols('x')
        try:
            self.expression = sp.sympify(expression_str)
        except Exception as e:
            raise ValueError(f"Lỗi: {str(e)}")

class MathAppGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Math Functions GUI")

        # Các biến và danh sách lưu trữ kết quả
        self.math_expression = None
        self.functions_from_excel = []
        self.extreme_points = []
        self.results_history = []

        # Tạo giao diện
        self.create_widgets()

    def create_widgets(self):
        # Cài đặt font lớn hơn
        bigger_font = ("Helvetica", 15)

        # Nhập hàm số
        ttk.Label(self.master, text="Nhập hàm số (sử dụng 'x' là biến):", font=bigger_font).grid(row=0, column=0, pady=10, padx=10, sticky=tk.W)
        self.entry_function = ttk.Entry(self.master, width=50)
        self.entry_function.grid(row=0, column=1, pady=10, padx=10, columnspan=2)

        # Nút Nhập Dữ Liệu từ Excel
        ttk.Button(self.master, text="Nhập Dữ Liệu từ Excel", command=self.load_data_from_excel).grid(row=0, column=3, pady=10, padx=10, sticky=tk.W)

        # Danh sách chức năng
        functions = {
            "Tính Đạo Hàm": self.calculate_derivative,
            "Tính Nguyên Hàm": self.calculate_integral,
            "Giải Phương Trình": self.solve_equation,
            "Tính Diện Tích Hàm Số và Đường Thẳng": self.calculate_area,
            "Tìm Điểm Cực Trị": self.find_extreme_points,
            "Tính Diện Tích Giữa Hai Hàm": self.calculate_area_between_functions,
            "Kiểm Tra Tính Liên Tục": self.check_continuity,
            "Lưu Kết Quả": self.save_result
        }

        # Vị trí hàng ban đầu
        row_index = 1

        # Tạo nút cho mỗi chức năng
        for text, command in functions.items():
            ttk.Button(self.master, text=text, command=command, padding=(5, 15), width=50).grid(row=row_index, column=0, pady=0, padx=10, sticky=tk.W)
            row_index += 1

        # Label hiển thị hàm số đang được sửa đổi
        self.current_function_label = ttk.Label(self.master, text="Hàm số đang được sửa đổi:", font=bigger_font)
        self.current_function_label.grid(row=1, column=1, pady=10, padx=10, columnspan=2, sticky=tk.W)

        # Label hiển thị kết quả
        self.result_label = ttk.Label(self.master, text="", wraplength=400, background="lightgray", font=("Arial", 12))
        self.result_label.grid(row=10, column=0, pady=10, padx=10, columnspan=3, rowspan=2, sticky=tk.W + tk.E)

        # Tạo đồ thị
        self.figure = Figure(figsize=(5, 4), dpi=100)
        self.plot = self.figure.add_subplot(1, 1, 1)
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.master)
        self.canvas.get_tk_widget().grid(row=2, column=2, rowspan=8, pady=50, padx=50, columnspan=2)
        self.entry_function.bind("<KeyRelease>", self.update_plot)

    def update_plot(self, event):
        # Cập nhật đồ thị khi nhập hàm số mới
        expression_str = self.entry_function.get()
        self.math_expression = MathExpression(expression_str)
        x_vals = np.linspace(-10, 10, 1000)
        y_vals = [self.math_expression.expression.subs(self.math_expression.x, val) for val in x_vals]
        self.plot.clear()
        self.plot.plot(x_vals, y_vals, label='Hàm số')

        for point in self.extreme_points:
            x_val = point['point']
            y_val = self.math_expression.expression.subs(self.math_expression.x, x_val)
            self.plot.scatter([x_val], [y_val], color='red', label=f'{point["type"]} tại x={x_val}')

        self.plot.set_title("Đồ thị hàm số")
        self.plot.set_xlabel("x")
        self.plot.set_ylabel("f(x)")
        self.plot.legend()
        self.canvas.draw()

        self.current_function_label.config(text=f"Hàm số đang được sửa đổi: {expression_str}")

    def load_data_from_excel(self):
        # Nhập dữ liệu từ tệp tin Excel
        file_path = filedialog.askopenfilename()
        if file_path:
            try:
                data_frame = pd.read_excel(file_path)

                if 'function' in data_frame.columns and not data_frame['function'].empty:
                    self.functions_from_excel = data_frame['function'].tolist()

                    function_combobox = ttk.Combobox(self.master, values=self.functions_from_excel, state="readonly")
                    function_combobox.grid(row=0, column=4, pady=10, padx=10)
                    function_combobox.set("Chọn hàm số")

                    function_combobox.bind("<<ComboboxSelected>>", self.selected_function_from_combobox)

                data_preview = data_frame.head()
                self.show_result(f"Nội dung
