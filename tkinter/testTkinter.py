import tkinter as tk
import tkinter.ttk as ttk

def on_confirm():
    # 取得月份
    month = combobox.get()

    # 使用 API 抓取資料
    data = get_data(month)

    # 顯示資料
    label.config(text=data)

root = tk.Tk()
root.title('銷售量報表')
root.geometry('400x200')

# 建立月曆組件
combobox = ttk.Combobox(root)
combobox['values'] = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']
combobox.pack(pady=10)

# 建立確認按鈕
button = tk.Button(root, text='確認', command=on_confirm)
button.pack(pady=10)

# 建立標籤
label = tk.Label(root)
label.pack()

root.mainloop()

# 假設 get_data() 函式會回傳資料
def get_data(month):
    if month == '1':
        return '1月銷售量為 100 萬'
    elif month == '2':
        return '2月銷售量為 200 萬'
    elif month == '3':
        return '3月銷售量為 300 萬'
    else:
        return '其他月份的銷售量尚未統計'
