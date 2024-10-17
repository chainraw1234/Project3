import sys
from openpyxl import Workbook, load_workbook
import tkinter as tk
from tkinter import Toplevel
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from matplotlib import rcParams
import numpy as np
import pulldataRowColumn as data
import tkinter.ttk as ttk
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.table import Table

# ตั้งค่า encoding ให้เป็น utf-8
sys.stdout.reconfigure(encoding='utf-8')

# โหลดไฟล์ workbook และ worksheet
wb = load_workbook('Test/LProfitLoss24-Feb-.xlsx')
ws = wb.active

# ตั้งค่าฟอนต์ที่รองรับภาษาไทย เช่น Tahoma
rcParams['font.family'] = 'Tahoma'

# ฟังก์ชันในการสร้าง Pie Chart แบบ Donut Chart
def create_pie_chart(data, labels, title):
    fig, ax = plt.subplots(figsize=(6, 3), subplot_kw=dict(aspect="equal"))

    fig.subplots_adjust(left=0.03, right=0.9)  # ปรับขอบซ้ายและขวา

    # ใช้สี default จาก matplotlib
    colors = plt.get_cmap('tab20').colors  # ใช้ color map ที่มีสีให้เลือกมากมาย

    # ถ้ามี 'ไม่เกี่ยว' ใน labels ให้ทำให้เป็นสีขาว
    if 'ไม่เกี่ยว' in labels:
        idx = labels.index('ไม่เกี่ยว')
        colors = list(colors)  # แปลงเป็น list เพื่อปรับแต่งสี
        colors[idx] = '#ffffff'  # เปลี่ยนสี 'ไม่เกี่ยว' เป็นสีขาว
        
    # สร้าง Donut Chart โดยใช้ width=0.5 เพื่อทำให้เป็น donut chart
    wedges, texts, autotexts = ax.pie(data, wedgeprops=dict(width=1), startangle=-40, colors=colors[:len(labels)],
                                      radius=1, autopct='', pctdistance=0.75)

    for i, p in enumerate(wedges):
        ang = (p.theta2 - p.theta1) / 2. + p.theta1
        y = np.sin(np.deg2rad(ang))
        x = np.cos(np.deg2rad(ang))
        horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]  # จัดการข้อความทางซ้าย/ขวา
        connectionstyle = "angle,angleA=0,angleB={}".format(ang)  # สไตล์การเชื่อมต่อเส้นชี้
        percent = (data[i] / sum(data)) * 100

        if labels[i] == 'ไม่เกี่ยว':
            continue  # ไม่ต้องแสดงเปอร์เซ็นต์สำหรับ 'ไม่เกี่ยว'

        if percent < 1.5:
            # กรณีเปอร์เซ็นต์น้อยกว่า 1.5% ให้มีการแสดงเส้นชี้และข้อความปกติ
            ax.annotate(f'{percent:.1f}%', xy=(x, y), xytext=(1.2 * x, 1.2 * y),
                        horizontalalignment=horizontalalignment,
                        arrowprops=dict(arrowstyle="-", connectionstyle=connectionstyle))
        else:
            # แสดงเปอร์เซ็นต์ที่กึ่งกลาง wedge
            ax.text(x * 0.85, y * 0.85, '{:.1f}%'.format(percent), ha='center', va='center', fontsize=8)

    # เพิ่ม legend สำหรับ labels และแสดง title
    ax.legend(wedges, labels, title="Categories", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1), fontsize=8)
    ax.set_title(title)

    return fig

# ฟังก์ชันในการเพิ่มข้อมูลส่วนที่ขาด
def adjust_for_missing_percentage(data, labels):
    total = sum(data)
    if total < 98:
        missing_percentage = 100 - total
        data.append(missing_percentage)
        labels.append('ไม่เกี่ยว')
    return data, labels

# ข้อมูลตัวอย่างสำหรับ Pie Charts
percent_price_material = list(data.percent_price_material.values())
Category = list(data.percent_price_material.keys())
price_material_values = list(data.price_material.values())

percent_invest = list(data.percent_invest.values())
invest = list(data.percent_invest.keys())
price_invest = list(data.price_invest.values())

percent_total_cost = list(data.percent_total_cost.values())
cost = list(data.percent_total_cost.keys())

percent_sell = list(data.percent_sell.values())
sell = list(data.percent_sell.keys())

profit_value = list(data.profit.values())
profit_key = list(data.profit.keys())

sell_value = list(data.sell.values())
sell_key = list(data.sell.keys())

title_sell = f"Sell {data.month.year}{data.month.strftime('%b')}"
title_profit = f"Profit {data.month.year}{data.month.strftime('%b')}"

# กรอง data1 และ labels1 ให้ไม่รวมค่าที่เป็น 0
filtered_percent_material = []
filtered_Category = []
filtered_price_material_values = []

for i, value in enumerate(percent_price_material):
    if value != 0:  # ถ้า value ไม่เท่ากับ 0
        filtered_percent_material.append(value)  # เก็บค่าใน filtered_data1
        filtered_Category.append(Category[i])  # เก็บ labels ที่สอดคล้องกับค่าใน filtered_labels1

for i, value in enumerate(price_material_values):
    if value != 0:  # ถ้า value ไม่เท่ากับ 0
        filtered_price_material_values.append(value)

filtered_percent_invest = []
filtered_invest = []
filtered_invest_values = []
total_invest = 0

for i, value in enumerate(percent_invest):
    if value != 0:  # ถ้า value ไม่เท่ากับ 0
        filtered_percent_invest.append(value)  # เก็บค่าใน filtered_data1
        filtered_invest.append(invest[i])  # เก็บ labels ที่สอดคล้องกับค่าใน filtered_labels1

for i, value in enumerate(price_invest):
    if value != 0:  # ถ้า value ไม่เท่ากับ 0
        filtered_invest_values.append(value)
        total_invest+= value

filtered_percent_total_cost = []
filtered_cost = []
for i, value in enumerate(percent_total_cost):
    if value != 0:
        filtered_percent_total_cost.append(value)
        filtered_cost.append(cost[i])

filtered_percent_sell = []
filtered_sell = []
total_percent_sell = 0

for i, value in enumerate(percent_sell):
    if value != 0:
        filtered_percent_sell.append(value)
        filtered_sell.append(sell[i])
        total_percent_sell+=value

# ปรับข้อมูลสำหรับ Pie Chart
adjusted_percent_material, adjusted_Category = adjust_for_missing_percentage(filtered_percent_material, filtered_Category)
adjusted_percent_invest, adjusted_invest = adjust_for_missing_percentage(filtered_percent_invest, filtered_invest)
adjusted_percent_total_cost, adjusted_cost = adjust_for_missing_percentage(filtered_percent_total_cost, filtered_cost)
adjusted_percent_sell, adjusted_sell = adjust_for_missing_percentage(filtered_percent_sell, filtered_sell)

# ฟังก์ชันเปิดหน้าต่างใหม่แสดงกราฟและสัดส่วนในรูปแบบตาราง
def open_new_window(data, labels, title):
    new_window = Toplevel(root)
    new_window.title(f"รายละเอียด {title}")

    # สร้าง Pie Chart ในหน้าต่างใหม่
    fig = create_pie_chart(data, labels, title)
    canvas = FigureCanvasTkAgg(fig, master=new_window)
    canvas.draw()
    canvas.get_tk_widget().pack()

    # ตรวจสอบว่าถ้า title เป็น "Material Cost"
    if title == "Material Cost":
        # สร้าง Treeview สำหรับตารางที่มี 3 คอลัมน์
        table = ttk.Treeview(new_window, columns=('Category', 'Value(baht)', 'Value(%)'), show='headings', height=len(data)+1)
        table.pack(pady=10)  # เพิ่มช่องว่างด้านบนของตาราง

        # กำหนดหัวตาราง
        table.heading('Category', text='Category')
        table.heading('Value(baht)', text='Value (baht)')
        table.heading('Value(%)', text='Value (%)')

        # กำหนดขนาดของแต่ละคอลัมน์
        table.column('Category', width=200, anchor='w')
        table.column('Value(baht)', width=150, anchor='e')
        table.column('Value(%)', width=100, anchor='e')

        # เพิ่มข้อมูลลงในตาราง
        for label, value, price_material in zip(labels, data, filtered_price_material_values):
            table.insert('', 'end', values=(label, f'{price_material:.2f} baht', f'{value:.2f} %'))
        
        # เพิ่มแถวสรุป "รวม" ลงในตาราง และใช้แท็ก 'summary'
        table.insert('', 'end', values=('รวม', f'{total_material:.2f} baht', '100 %'), tags=('summary',))

        # กำหนดสีพื้นหลังสำหรับแถว "รวม"
        table.tag_configure('summary', background='lightgray')

    elif title == "Invest Cost":
         # สร้าง Treeview สำหรับตารางที่มี 3 คอลัมน์
        table = ttk.Treeview(new_window, columns=('Category', 'Value(baht)', 'Value(%)'), show='headings', height=len(data)+1)
        table.pack(pady=10)  # เพิ่มช่องว่างด้านบนของตาราง

        # กำหนดหัวตาราง
        table.heading('Category', text='Category')
        table.heading('Value(baht)', text='Value (baht)')
        table.heading('Value(%)', text='Value (%)')

        # กำหนดขนาดของแต่ละคอลัมน์
        table.column('Category', width=200, anchor='w')
        table.column('Value(baht)', width=150, anchor='e')
        table.column('Value(%)', width=100, anchor='e')

        # เพิ่มข้อมูลลงในตาราง
        for label, value, invest_value in zip(labels, data, filtered_invest_values):
            table.insert('', 'end', values=(label, f'{invest_value:.2f} baht', f'{value:.2f} %'))
        
        # เพิ่มแถวสรุป "รวม" ลงในตาราง และใช้แท็ก 'summary'
        table.insert('', 'end', values=('Total Main Invest', f'{total_invest:.2f} baht', '100 %'), tags=('summary',))

        # กำหนดสีพื้นหลังสำหรับแถว "รวม"
        table.tag_configure('summary', background='lightgray')

    elif title == "Sell Portion":
        # สร้าง Treeview สำหรับตารางที่มี 2 คอลัมน์ปกติ
        table = ttk.Treeview(new_window, columns=('Category', 'Value'), show='headings', height=len(data))
        table.pack(pady=10)  # เพิ่มช่องว่างด้านบนของตาราง

        # กำหนดหัวตาราง
        table.heading('Category', text='Category')
        table.heading('Value', text='Value (%)')

        # กำหนดขนาดของแต่ละคอลัมน์
        table.column('Category', width=200, anchor='w')
        table.column('Value', width=100, anchor='e')

        # เพิ่มข้อมูลลงในตาราง
        for label, value in zip(labels, data):
            if label == "ไม่เกี่ยว":
                continue
            else:
                table.insert('', 'end', values=(label, f'{value:.2f} %'))

        if title == "Sell Portion":
            table.insert('', 'end', values=('Total Cost', f'{total_percent_sell:.2f} %'), tags=('summary',))
            table.tag_configure('summary', background='lightgray')

    else:
        # สร้าง Treeview สำหรับตารางที่มี 2 คอลัมน์ปกติ
        table = ttk.Treeview(new_window, columns=('Category', 'Value'), show='headings', height=len(data)+1)
        table.pack(pady=10)  # เพิ่มช่องว่างด้านบนของตาราง

        # กำหนดหัวตาราง
        table.heading('Category', text='Category')
        table.heading('Value', text='Value (%)')

        # กำหนดขนาดของแต่ละคอลัมน์
        table.column('Category', width=200, anchor='w')
        table.column('Value', width=100, anchor='e')

        # เพิ่มข้อมูลลงในตาราง
        for label, value in zip(labels, data):
            table.insert('', 'end', values=(label, f'{value:.2f} %'))

        if title == "Cost Portion":
            table.insert('', 'end', values=('Total Cost','100 %'), tags=('summary',))
            table.tag_configure('summary', background='lightgray')


    # เพิ่มสไตล์ของเส้นตารางให้เหมือน Excel
    style = ttk.Style()
    style.configure("Treeview", rowheight=25)
    style.configure("Treeview.Heading", font=('Arial', 13, 'bold'))
    style.configure("Treeview", font=('Arial', 13), borderwidth=1)
    style.map('Treeview', background=[('selected', '#d9d9d9')])

def create_table_profit(root, data, labels, title):
    # สร้าง Treeview สำหรับตาราง profit
    table = ttk.Treeview(root, columns=('Category', 'Value'), show='headings', height=len(data))
    table.grid(row=1, column=1, padx=10, pady=(50, 10), sticky='n')  # กำหนด pady ให้มีระยะห่างด้านล่าง 50px
    
    # กำหนดหัวตาราง
    table.heading('Category', text=title)
    table.heading('Value', text='Value')

    # กำหนดขนาดของแต่ละคอลัมน์
    table.column('Category', width=450, anchor='w')
    table.column('Value', width=150, anchor='e')

    # เพิ่มข้อมูลลงในตาราง
    for i, (label, value) in enumerate(zip(labels, data)):
        if i < 5:
            table.insert('', 'end', values=(label, f'{value:.2f} baht'))
        elif i == 5:
            table.insert('', 'end', values=(label, f'{value:.2f} %'))
        else:
            table.insert('', 'end', values=(label, f'{value:.2f}'))

    # เพิ่มสไตล์ของเส้นตาราง
    style = ttk.Style()
    style.configure("Treeview", rowheight=30)
    style.configure("Treeview.Heading", font=('Arial', 13, 'bold'))
    style.configure("Treeview", font=('Arial', 13), borderwidth=1)
    style.map('Treeview', background=[('selected', '#d9d9d9')])

def create_table_sell(root, data, labels, title):
    # สร้าง Treeview สำหรับตาราง sell
    table = ttk.Treeview(root, columns=('Category', 'Value'), show='headings', height=len(data))
    table.grid(row=1, column=1, padx=10, pady=(280, 5), sticky='n')  # วางไว้ใต้ตาราง profit และกำหนด pady ด้านบน 50px
    
    # กำหนดหัวตาราง
    table.heading('Category', text=title)
    table.heading('Value', text='Value')

    # กำหนดขนาดของแต่ละคอลัมน์
    table.column('Category', width=450, anchor='w')
    table.column('Value', width=150, anchor='e')

    # เพิ่มข้อมูลลงในตาราง
    for i, (label, value) in enumerate(zip(labels, data)):
        if i == 0:
            table.insert('', 'end', values=(label, f'{value:.2f} day'))
        elif i > 0 and i <= 2:
            table.insert('', 'end', values=(label, f'{value:.2f} baht'))
        elif i == 3:
            table.insert('', 'end', values=(label, f'{value:.2f} คน'))
        else:
            table.insert('', 'end', values=(label, f'{value:.2f}'))

    # เพิ่มสไตล์ของเส้นตาราง
    style = ttk.Style()
    style.configure("Treeview", rowheight=30)
    style.configure("Treeview.Heading", font=('Arial', 13, 'bold'))
    style.configure("Treeview", font=('Arial', 13), borderwidth=1)
    style.map('Treeview', background=[('selected', '#d9d9d9')])

total_material = data.total_price_material

def save_as_pdf():
    pdf_filename = "report_with_tables.pdf"
    with PdfPages(pdf_filename) as pdf:
        # สร้าง Pie Charts และตาราง
        charts_and_tables = {
            "Material Cost": (adjusted_percent_material, adjusted_Category, filtered_price_material_values, "Material Cost"),
            "Invest Cost": (adjusted_percent_invest, adjusted_invest, filtered_invest_values, "Invest Cost"),
            "Cost Portion": (adjusted_percent_total_cost, adjusted_cost, None, "Cost Portion"),
            "Sell Portion": (adjusted_percent_sell, adjusted_sell, None, "Sell Portion")
        }

        for title, (data, labels, values, chart_title) in charts_and_tables.items():
            # สร้าง layout ที่แบ่งหน้าเป็น 2 ส่วน: Pie Chart และ ตาราง
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 6))
            fig.subplots_adjust(left=0.05, right=0.95, wspace=0.4)  # ปรับพื้นที่ระหว่างแผนภูมิและตาราง

            # ใช้สี default จาก matplotlib
            colors = plt.get_cmap('tab20').colors  # ใช้ color map ที่มีสีให้เลือกมากมาย

            # ถ้ามี 'ไม่เกี่ยว' ใน labels ให้ทำให้เป็นสีขาว
            if 'ไม่เกี่ยว' in labels:
                idx = labels.index('ไม่เกี่ยว')
                colors = list(colors)  # แปลงเป็น list เพื่อปรับแต่งสี
                colors[idx] = '#ffffff'  # เปลี่ยนสี 'ไม่เกี่ยว' เป็นสีขาว

            # สร้าง Donut Chart โดยใช้ width=0.5 เพื่อทำให้เป็น donut chart
            wedges, texts, autotexts = ax1.pie(data, wedgeprops=dict(width=1), startangle=-40, colors=colors[:len(labels)],
                                      radius=1, autopct='', pctdistance=0.75)

            for i, p in enumerate(wedges):
                ang = (p.theta2 - p.theta1) / 2. + p.theta1
                y = np.sin(np.deg2rad(ang))
                x = np.cos(np.deg2rad(ang))
                horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]  # จัดการข้อความทางซ้าย/ขวา
                connectionstyle = "angle,angleA=0,angleB={}".format(ang)  # สไตล์การเชื่อมต่อเส้นชี้
                percent = (data[i] / sum(data)) * 100

                if labels[i] == 'ไม่เกี่ยว':
                    continue  # ไม่ต้องแสดงเปอร์เซ็นต์สำหรับ 'ไม่เกี่ยว'

                if percent < 1.5:
                    # กรณีเปอร์เซ็นต์น้อยกว่า 1.5% ให้มีการแสดงเส้นชี้และข้อความปกติ
                    ax1.annotate(f'{percent:.1f}%', xy=(x, y), xytext=(1.2 * x, 1.2 * y),
                                horizontalalignment=horizontalalignment,
                                arrowprops=dict(arrowstyle="-", connectionstyle=connectionstyle))
                else:
                    # แสดงเปอร์เซ็นต์ที่กึ่งกลาง wedge
                    ax1.text(x * 0.85, y * 0.85, '{:.1f}%'.format(percent), ha='center', va='center', fontsize=8)

            # เพิ่ม legend สำหรับ labels และแสดง title
            ax1.legend(wedges, labels, title="Categories", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1), fontsize=8)
            ax1.set_title(title)

            # สร้างตารางที่ ax2
            ax2.set_axis_off()  # ปิดแกนของตาราง
            table = Table(ax2, bbox=[0, 0, 1, 1])
            headers = ["Category", "Value (%)"] if values is None else ["Category", "Value (Baht)", "Value (%)"]
            data_rows = list(zip(labels, data)) if values is None else list(zip(labels, values, data))
            col_widths = [0.6, 0.4] if values is None else [0.4, 0.3, 0.3]
            row_height = 1 / (len(data_rows) + 2)  # เพิ่ม +2 เพื่อรองรับแถวรวม

            # เพิ่มหัวตาราง
            for col, header in enumerate(headers):
                table.add_cell(0, col, col_widths[col], row_height, text=header, loc='center', facecolor='lightgrey')

            # เพิ่มแถว "รวม"
            total_baht = sum(values) if values is not None else None
            total_percent = sum(data)

            # เพิ่มข้อมูลลงในตาราง
            for row, row_data in enumerate(data_rows, start=1):
                for col, cell_data in enumerate(row_data):
                    #print(cell_data)
                    if row_data[0] == 'ไม่เกี่ยว' and cell_data == 'ไม่เกี่ยว':
                        total_percent -= row_data[1]
                        
                        continue
                    elif row_data[0] != 'ไม่เกี่ยว' and cell_data != 'ไม่เกี่ยว':
                        cell_text = f'{cell_data:.2f}' if isinstance(cell_data, float) else str(cell_data)
                        table.add_cell(row, col, col_widths[col], row_height, text=cell_text, loc='center')


            #print(values)

            if values is not None:
                # กรณีมีทั้งค่า Baht และ Percent
                table.add_cell(len(data_rows) + 1, 0, col_widths[0], row_height, text="รวม", loc='center')
                table.add_cell(len(data_rows) + 1, 1, col_widths[1], row_height, text=f'{total_baht:.2f}', loc='center')
                table.add_cell(len(data_rows) + 1, 2, col_widths[2], row_height, text=f'{total_percent:.2f}', loc='center')
            else:
                # กรณีมีเฉพาะ Percent
                table.add_cell(len(data_rows) + 1, 0, col_widths[0], row_height, text="รวม", loc='center')
                table.add_cell(len(data_rows) + 1, 1, col_widths[1], row_height, text=f'{total_percent:.2f}', loc='center')

            ax2.add_table(table)

            pdf.savefig(fig)  # บันทึก Pie Chart และตารางในหน้าเดียวกัน
            plt.close(fig)

        # สร้างหน้าสุดท้ายสำหรับ Profit และ Sell (สามารถจัดการลักษณะเดียวกันได้)
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 6))
        fig.subplots_adjust(left=0.05, right=0.95, hspace=0.3)  # ปรับพื้นที่ระหว่างตาราง
        
        # Profit
        profit_headers = ["Category", "Value (Baht)"]
        profit_data_rows = list(zip(profit_key, profit_value))
        ax1.set_axis_off()  # ปิดแกนของตาราง
        table = Table(ax1, bbox=[0, 0, 1, 1])
        row_height = 1 / (len(profit_data_rows) + 1)

        for col, header in enumerate(profit_headers):
            table.add_cell(0, col, col_widths[col], row_height, text=header, loc='center', facecolor='lightgrey')
        for row, row_data in enumerate(profit_data_rows, start=1):
            for col, cell_data in enumerate(row_data):
                cell_text = f'{cell_data:.2f}' if isinstance(cell_data, float) else str(cell_data)
                table.add_cell(row, col, col_widths[col], row_height, text=cell_text, loc='center')
        ax1.set_title(title_profit)
        ax1.add_table(table)

        # Sell
        sell_headers = ["Category", "Value"]
        sell_data_rows = list(zip(sell_key, sell_value))
        ax2.set_axis_off()
        table = Table(ax2, bbox=[0, 0, 1, 1])
        row_height = 1 / (len(sell_data_rows) + 1)

        for col, header in enumerate(sell_headers):
            table.add_cell(0, col, col_widths[col], row_height, text=header, loc='center', facecolor='lightgrey')
        for row, row_data in enumerate(sell_data_rows, start=1):
            for col, cell_data in enumerate(row_data):
                cell_text = f'{cell_data:.2f}' if isinstance(cell_data, float) else str(cell_data)
                table.add_cell(row, col, col_widths[col], row_height, text=cell_text, loc='center')
        ax2.set_title(title_sell)
        ax2.add_table(table)

        pdf.savefig(fig)
        plt.close(fig)

    print("PDF saved successfully!")



# สร้างหน้าต่าง Tkinter
root = tk.Tk()
root.title("หน้ารวม Donut Chart")

# ทำให้หน้าต่างขยายเต็มหน้าจอแบบ maximized
root.state('zoomed')

# ฟังก์ชันในการปรับความกว้างและความสูงของแต่ละกราฟให้เป็น 1/3 ของความกว้าง และครึ่งหนึ่งของความสูงหน้าจอ
def resize_canvases(root, canvas_list):
    root.update_idletasks()  # อัพเดตข้อมูลขนาดหน้าจอ
    screen_width = root.winfo_width()  # ดึงความกว้างของหน้าต่าง root
    screen_height = root.winfo_height()  # ดึงความสูงของหน้าต่าง root
    canvas_width = screen_width // 3  # กำหนดความกว้างให้เท่ากับ 1/3 ของหน้าต่าง root
    canvas_height = screen_height // 2  # กำหนดความสูงให้เท่ากับครึ่งหนึ่งของหน้าต่าง root

    for canvas in canvas_list:
        canvas.get_tk_widget().config(width=canvas_width, height=canvas_height)

# Donut Chart ที่ 1
fig1 = create_pie_chart(adjusted_percent_material, adjusted_Category, "Material Cost")
canvas1 = FigureCanvasTkAgg(fig1, master=root)
canvas1.draw()
canvas1.get_tk_widget().grid(row=0, column=0)
canvas1.get_tk_widget().bind("<Button-1>", lambda event: open_new_window(adjusted_percent_material, adjusted_Category, "Material Cost"))

# Donut Chart ที่ 2
fig2 = create_pie_chart(adjusted_percent_invest, adjusted_invest, "Invest Cost")
canvas2 = FigureCanvasTkAgg(fig2, master=root)
canvas2.draw()
canvas2.get_tk_widget().grid(row=0, column=1)
canvas2.get_tk_widget().bind("<Button-1>", lambda event: open_new_window(adjusted_percent_invest, adjusted_invest, "Invest Cost"))

# Donut Chart ที่ 3
fig3 = create_pie_chart(adjusted_percent_total_cost, adjusted_cost, "Cost Portion")
canvas3 = FigureCanvasTkAgg(fig3, master=root)
canvas3.draw()
canvas3.get_tk_widget().grid(row=0, column=2)
canvas3.get_tk_widget().bind("<Button-1>", lambda event: open_new_window(adjusted_percent_total_cost, adjusted_cost, "Cost Portion"))

# Donut Chart ที่ 4
fig4 = create_pie_chart(adjusted_percent_sell, adjusted_sell, "Sell Portion")
canvas4 = FigureCanvasTkAgg(fig4, master=root)
canvas4.draw()
canvas4.get_tk_widget().grid(row=1, column=0)
canvas4.get_tk_widget().bind("<Button-1>", lambda event: open_new_window(adjusted_percent_sell, adjusted_sell, "Sell Portion"))

# ตาราง profit ที่ row 1 column 1 (ใต้ Donut Chart ที่ 4)
create_table_profit(root, list(data.profit.values()), list(data.profit.keys()), f"Profit {data.month.year}{data.month.strftime('%b')}")

# ตาราง sell ใต้ ตาราง profit
create_table_sell(root, list(data.sell.values()), list(data.sell.keys()), f"Sell {data.month.year}{data.month.strftime('%b')}")

# เพิ่มปุ่ม "Save as PDF" ที่ row 1 column 2
save_button = tk.Button(root, text="Save as PDF", command=save_as_pdf)
save_button.grid(row=1, column=2, padx=10, pady=(10, 20), sticky='e')

# รันฟังก์ชันปรับขนาดเมื่อหน้าต่าง root ถูกปรับขนาด (รวมตารางในรายการ)
root.bind("<Configure>", lambda event: resize_canvases(root, [canvas1, canvas2, canvas3, canvas4]))

# รันหน้าต่าง Tkinter
root.mainloop()
