#!/usr/bin/env python
# coding: utf-8

# In[3]:


import os
from pikepdf import Pdf, Permissions, Encryption
import glob
import tkinter as tk

# 拆分PDF 動作
def PDF_break_pages():
    row6.config(text="")  # 清空警告内容
    
    if row11.get() == "":
        row6.config(text="未輸入文件路徑!")
        return
        
    if not os.path.exists(row11.get()):
        row6.config(text="文件路徑不存在!")
        return
    
    if row21.get() == "":
        row6.config(text="未輸入PDF文件名!")
        return
        
    if not os.path.exists(row21.get() + ".pdf"):
        row6.config(text="PDF文件不存在!")
        return

    if row41.get() == "":
        row6.config(text="未輸入起始頁碼!")
        return
        
    if int(row41.get())<=0:
        row6.config(text="起始頁碼輸入錯誤!")
        return

    if row51.get() == "":
        row6.config(text="未輸入結束頁碼!")
        return
        
    if int(row51.get()) > len(choose_pdf.pages) or int(row51.get()) < start:
        row6.config(text="結束頁碼輸入錯誤!")
        return
    
    path = row11.get()  # 獲取文件路徑
    os.chdir(path)  # 切换工作路徑
    choose_pdf = Pdf.open(row21.get() + ".pdf")  # 打開所選 PDF
    if row31.get():
        save_pdf = row31.get() # 取得新 PDF 名稱
    else:
        i = 1
        while os.path.exists(f"break_{i}.pdf"):
            i += 1
        save_pdf = f"break_{i}"  # 取得新 PDF 名稱
    start = int(row41.get())  # 起始頁碼
    end = int(row51.get())  # 結束頁碼
    if int(row51.get()) <= len(choose_pdf.pages) and int(row51.get()) >= start:
        end = int(row51.get())  # 結束頁碼
    pages = choose_pdf.pages
    output = Pdf.new()
    output.pages.extend(pages[start-1:end])  # 將指定範圍加到新 pdf
    output.save(save_pdf+".pdf")  # 儲存新 pdf
    row6.config(text="拆分成功!")

# 拆分PDF GUI
def PDF_break_pages_GUI():
    # GUI設置
    row11.delete(0, "end")  # 清除內容
    row21.delete(0, "end")  
    row31.delete(0, "end")  
    row41.delete(0, "end")  
    row51.delete(0, "end")  
    
    row1.config(text="文件路徑：")
    row1.grid(row=1, column=0, sticky="w")
    row11.grid(row=1, column=1)
    
    row2.config(text="PDF 文件名(不含.pdf)：")
    row2.grid(row=2, column=0, sticky="w")
    row21.grid(row=2, column=1)
    
    row3.config(text="新 PDF 文件名(非必要)(不含.pdf)：")
    row3.grid(row=3, column=0, sticky="w")
    row31.grid(row=3, column=1)
    
    row4.config(text="起始頁碼：")
    row4.grid(row=4, column=0, sticky="w")
    row41.grid(row=4, column=1)
    
    row5.config(text="結束頁碼：")
    row5.grid(row=5, column=0, sticky="w")
    row51.grid(row=5, column=1)
    
    row6.config(text=" ")
    row6.grid(row=6, column=0, columnspan=2)

    break_button = tk.Button(root, text="執行", command=PDF_break_pages)
    break_button.grid(row=7, column=0, columnspan=2, pady=10)

# 大量合併PDF 動作
def PDF_L_marge():
    row6.config(text="")  # 清空警告内容
    
    if row11.get()== "":
        row6.config(text="未輸入文件路徑!")
        return
        
    if not os.path.exists(row11.get()):
        row6.config(text="文件路徑不存在!")
        return

    if row21.get()== "":
        row6.config(text="未輸入起始 PDF 文件名!")
        return

    if not os.path.exists(row21.get() + ".pdf"):
        row6.config(text="起始 PDF 文件名不存在!")
        return

    try:
        int(row21.get())
    except ValueError:
        row6.config(text="起始 PDF 文件名輸入錯誤!請輸入數字")
        return
        
    if row31.get()== "":
        row6.config(text="未輸入結束 PDF 文件名!")
        return
        
    if not os.path.exists(row31.get() + ".pdf"):
        row6.config(text="結束 PDF 文件名不存在!")
        return

    try:
        int(row31.get())
    except ValueError:
        row6.config(text="結束 PDF 文件名輸入錯誤!請輸入數字")
        return
        
    if int(row31.get()) <= int(row21.get())+1:
        row6.config(text="PDF 文件名順序錯誤!")
        return

    path = row11.get()  # 獲取文件路徑
    os.chdir(path)  # 切换工作路徑
    output = Pdf.new()                   # 建立新的 pdf 物件
    for i in range(int(row21.get()) , int(row31.get())+1):
        pdf = Pdf.open(f'{i}.pdf')        # 讀取第一份 pdf
        output.pages.extend(pdf.pages)   # 添加進pdf
    if row41.get():
        save_pdf = row41.get()
    else:
        i = 1
        while os.path.exists(f"L_marge_{i}.pdf"):
            i += 1
        save_pdf = f"L_marge_{i}"
    output.save(save_pdf+'.pdf')
    row6.config(text="大量合併成功!")

# 大量合併PDF GUI
def PDF_L_marge_GUI():
    # GUI設置
    row11.delete(0, "end")  # 清除內容
    row21.delete(0, "end")  
    row31.delete(0, "end")  
    row41.delete(0, "end")  
    row51.delete(0, "end")  
    
    row1.config(text="文件路徑：")
    row1.grid(row=1, column=0, sticky="w")
    row11.grid(row=1, column=1)
    
    row2.config(text="起始 PDF 文件名(數字)(不含.pdf)：")
    row2.grid(row=2, column=0, sticky="w")
    row21.grid(row=2, column=1)
    
    row3.config(text="結束 PDF 文件名(數字)(不含.pdf)：")
    row3.grid(row=3, column=0, sticky="w")
    row31.grid(row=3, column=1)
    
    row4.config(text="新 PDF 文件名(非必要)(不含.pdf)：")
    row4.grid(row=4, column=0, sticky="w")
    row41.grid(row=4, column=1)
    
    row5.config(text=" ")
    row5.grid(row=5, column=0)
    
    row6.config(text=" ")
    row6.grid(row=6, column=0, columnspan=2)

    break_button = tk.Button(root, text="執行", command=PDF_L_marge)
    break_button.grid(row=7, column=0, columnspan=2, pady=10)

# 合併PDF 動作
def PDF_marge():
    row6.config(text="")  # 清空警告内容
    
    if row11.get() == "":
        row6.config(text="未輸入文件路徑!")
        return
        
    if not os.path.exists(row11.get()):
        row6.config(text="文件路徑不存在!")
        return
        
    if row21.get() == "":
        row6.config(text="未填入合併1 PDF 文件名!")
        return

    if not os.path.exists(row21.get() + ".pdf"):
        row6.config(text="合併1 PDF 文件名不存在!")
        return
    
    if row31.get() == "":
        row6.config(text="未填入合併2 PDF 文件名!")
        return
        
    if not os.path.exists(row31.get() + ".pdf"):
        row6.config(text="合併2 PDF 文件名不存在!")
        return

    path = row11.get()  # 獲取文件路徑
    os.chdir(path)  # 切换工作路徑
    output = Pdf.new()                   # 建立新的 pdf 物件
    pdf = Pdf.open(row21.get()+'.pdf')        # 讀取第一份 pdf
    output.pages.extend(pdf.pages)   # 添加進pdf
    pdf = Pdf.open(row31.get()+'.pdf')        # 讀取第二份 pdf
    output.pages.extend(pdf.pages)   # 添加進pdf
    if row41.get():
        save_pdf = row41.get()
    else:
        i = 1
        while os.path.exists(f"marge_{i}.pdf"):
            i += 1
        save_pdf = f"marge_{i}"
    output.save(save_pdf+'.pdf')
    row6.config(text="合併成功!")

# 合併PDF GUI
def PDF_marge_GUI():
    # GUI設置
    row11.delete(0, "end")  # 清除內容
    row21.delete(0, "end")  
    row31.delete(0, "end")  
    row41.delete(0, "end")  
    row51.delete(0, "end")  
    
    row1.config(text="文件路徑：")
    row1.grid(row=1, column=0, sticky="w")
    row11.grid(row=1, column=1)
    
    row2.config(text="合併1 PDF 文件名(不含.pdf)：")
    row2.grid(row=2, column=0, sticky="w")
    row21.grid(row=2, column=1)
    
    row3.config(text="合併2 PDF 文件名(不含.pdf)：")
    row3.grid(row=3, column=0, sticky="w")
    row31.grid(row=3, column=1)
    
    row4.config(text="新 PDF 文件名(非必要)(不含.pdf)：")
    row4.grid(row=4, column=0, sticky="w")
    row41.grid(row=4, column=1)
    
    row5.config(text=" ")
    row5.grid(row=5, column=0)
    
    row6.config(text=" ")
    row6.grid(row=6, column=0, columnspan=2)

    break_button = tk.Button(root, text="執行", command=PDF_marge)
    break_button.grid(row=7, column=0, columnspan=2, pady=10)

# 加密PDF 動作
def PDF_encryption():
    row6.config(text="")  # 清空警告内容
    
    if row11.get() == "":
        row6.config(text="未輸入文件路徑!")
        return
        
    if not os.path.exists(row11.get()):
        row6.config(text="文件路徑不存在!")
        return
        
    if row21.get() == "":
        row6.config(text="未輸入 PDF 文件名!")
        return
        
    if not os.path.exists(row21.get() + ".pdf"):
        row6.config(text="PDF 文件名錯誤!")
        return
        
    if row31.get() == "":
        row6.config(text="未輸入密碼!")
        return
    
    path = row11.get()  # 獲取文件路徑
    os.chdir(path)  # 切换工作路徑
    pdf = Pdf.open(row21.get() + ".pdf")         # 開啟 pdf
    no_extracting = Permissions(extract=False)
    pdf.save('new.pdf', encryption = Encryption(user=row31.get(), owner=row31.get(), allow=no_extracting))
        
    if row41.get():
        save_pdf = row41.get()
    else:
        i = 1
        while os.path.exists(f"encryption_{i}.pdf"):
            i += 1
        save_pdf = f"encryption_{i}"
        
    no_extracting = Permissions(extract=False)
    pdf.save(save_pdf+'.pdf', encryption = Encryption(user=row31.get(), owner=row31.get(), allow=no_extracting))
    row6.config(text="加密成功!")

# 加密PDF GUI
def PDF_encryption_GUI():
    # GUI設置
    row11.delete(0, "end")  # 清除內容
    row21.delete(0, "end")  
    row31.delete(0, "end")  
    row41.delete(0, "end")  
    row51.delete(0, "end")  
    
    row1.config(text="文件路徑：")
    row1.grid(row=1, column=0, sticky="w")
    row11.grid(row=1, column=1)
    
    row2.config(text="PDF 文件名(不含.pdf)：")
    row2.grid(row=2, column=0, sticky="w")
    row21.grid(row=2, column=1)
    
    row3.config(text="PDF 加密密碼：")
    row3.grid(row=3, column=0, sticky="w")
    row31.grid(row=3, column=1)
    
    row4.config(text="新 PDF 文件名(非必要)(不含.pdf)：")
    row4.grid(row=4, column=0, sticky="w")
    row41.grid(row=4, column=1)
    
    row5.config(text=" ")
    row5.grid(row=5, column=0)
    
    row6.config(text=" ")
    row6.grid(row=6, column=0, columnspan=2)

    break_button = tk.Button(root, text="執行", command=PDF_encryption)
    break_button.grid(row=7, column=0, columnspan=2, pady=10)

# 旋轉PDF 動作
def PDF_rotate():
    row6.config(text="")  # 清空警告内容
    
    if row11.get()== "":
        row6.config(text="未輸入文件路徑!")
        return
        
    if os.path.exists(row11.get())==0:
        row6.config(text="文件路徑不存在!")
        return
        
    if not row21.get():
        row6.config(text="未輸入 PDF 文件名!")
        return
        
    if not os.path.exists(row21.get() + ".pdf"):
        row6.config(text="PDF 文件名錯誤!")
        return
        
    if row31.get()== "":
        row6.config(text="未輸入角度!")
        return
        
    try:
        int(row31.get())
    except ValueError:
        row6.config(text="角度輸入錯誤!請輸入數字")
        return

    if int(row31.get())%90!=0:
        row6.config(text="角度輸入錯誤!需90的倍數")
        return
        
    path = row11.get()  # 獲取文件路徑
    os.chdir(path)  # 切换工作路徑
    pdf = Pdf.open(row21.get() + ".pdf")         # 開啟 pdf
    for page in pdf.pages:
        page.rotate(int(row31.get()), relative:=True)
    if row41.get():
        save_pdf = row41.get()
    else:
        i = 1
        while os.path.exists(f"rotate_{i}.pdf"):
            i += 1
        save_pdf = f"rotate_{i}"
        
    pdf.save(save_pdf+'.pdf')
    row6.config(text="旋轉成功!")

# 旋轉PDF GUI
def PDF_rotate_GUI():
    # GUI設置
    row11.delete(0, "end")  # 清除內容
    row21.delete(0, "end")  
    row31.delete(0, "end")  
    row41.delete(0, "end")  
    row51.delete(0, "end")  
    
    row1.config(text="文件路徑：")
    row1.grid(row=1, column=0, sticky="w")
    row11.grid(row=1, column=1)
    
    row2.config(text="PDF 文件名(不含.pdf)：")
    row2.grid(row=2, column=0, sticky="w")
    row21.grid(row=2, column=1)
    
    row3.config(text="PDF 旋轉角度(順時針)(90倍數)：")
    row3.grid(row=3, column=0, sticky="w")
    row31.grid(row=3, column=1)
    
    row4.config(text="新 PDF 文件名(非必要)(不含.pdf)：")
    row4.grid(row=4, column=0, sticky="w")
    row41.grid(row=4, column=1)
    
    row5.config(text=" ")
    row5.grid(row=5, column=0)
    
    row6.config(text=" ")
    row6.grid(row=6, column=0, columnspan=2)

    break_button = tk.Button(root, text="執行", command=PDF_rotate)
    break_button.grid(row=7, column=0, columnspan=2, pady=10)

# 主視窗
root = tk.Tk()
root.title("PDF 應用")

# 初始化全局
row11 = tk.Entry(root)
row21 = tk.Entry(root)
row31 = tk.Entry(root)
row41 = tk.Entry(root)
row51 = tk.Entry(root)


button_GUI = tk.Button(root, text="PDF 拆分", command=PDF_break_pages_GUI)
button_GUI.grid(row=0, column=0, pady=10, sticky="ew")  # 使用 sticky="ew" 讓按鈕水平填滿單元格

button_GUI = tk.Button(root, text="PDF 大量合併", command=PDF_L_marge_GUI)
button_GUI.grid(row=0, column=1, pady=10, sticky="ew")

button_GUI = tk.Button(root, text="PDF 合併", command=PDF_marge_GUI)
button_GUI.grid(row=0, column=2, pady=10, sticky="ew")

button_GUI = tk.Button(root, text="PDF 加密", command=PDF_encryption_GUI)
button_GUI.grid(row=0, column=3, pady=10, sticky="ew")

button_GUI = tk.Button(root, text="PDF 旋轉", command=PDF_rotate_GUI)
button_GUI.grid(row=0, column=4, pady=10, sticky="ew")

# 設定每一列的權重，使其平均分配
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)
root.columnconfigure(2, weight=1)
root.columnconfigure(3, weight=1)

row1 = tk.Label(root, text="選擇下一步動作")
row1.grid(row=1, column=0)

row2 = tk.Label(root, text=" ")
row2.grid(row=2, column=0)

row3 = tk.Label(root, text=" ")
row3.grid(row=3, column=0)

row4 = tk.Label(root, text=" ")
row4.grid(row=4, column=0)

row5 = tk.Label(root, text=" ")
row5.grid(row=5, column=0)

row6 = tk.Label(root, text=" ")
row6.grid(row=6, column=0, columnspan=2)

root.mainloop()  # 執行主視窗
#用auto-py-to-exe 打包成.exe


# In[ ]:


#用auto-py-to-exe 打包成.exe

