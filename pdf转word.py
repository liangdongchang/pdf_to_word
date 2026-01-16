# -*- coding: utf-8 -*-

"""
@contact:  1
@file: pdf转word.py
@time: 2025/12/24 16:17
@author: LDC
"""
# import os
# from pdf2docx import Converter
# import pytesseract
# from PIL import Image
# import io
# from pdf2image import convert_from_path
# import tkinter as tk
# from tkinter import ttk, filedialog, messagebox
# import threading
#
# class AdvancedPDFConverter:
#     def __init__(self):
#         self.window = tk.Tk()
#         self.window.title("高级PDF转Word工具")
#         self.window.geometry("600x500")
#
#         self.setup_ui()
#
#     def setup_ui(self):
#         # 文件选择区域
#         tk.Label(self.window, text="PDF转Word高级转换器",
#                 font=("微软雅黑", 16)).pack(pady=10)
#
#         # PDF文件
#         tk.Label(self.window, text="PDF文件:").pack(anchor="w", padx=20)
#         self.pdf_path = tk.StringVar()
#         tk.Entry(self.window, textvariable=self.pdf_path, width=60).pack(padx=20, pady=5)
#         tk.Button(self.window, text="选择PDF",
#                  command=self.select_pdf).pack(pady=5)
#
#         # 输出路径
#         tk.Label(self.window, text="输出Word文件:").pack(anchor="w", padx=20)
#         self.output_path = tk.StringVar()
#         tk.Entry(self.window, textvariable=self.output_path, width=60).pack(padx=20, pady=5)
#         tk.Button(self.window, text="选择输出位置",
#                  command=self.select_output).pack(pady=5)
#
#         # 选项区域
#         options_frame = tk.LabelFrame(self.window, text="转换选项", padx=10, pady=10)
#         options_frame.pack(fill="x", padx=20, pady=10)
#
#         # OCR选项
#         self.use_ocr = tk.BooleanVar(value=False)
#         tk.Checkbutton(options_frame, text="使用OCR识别扫描版PDF",
#                       variable=self.use_ocr).pack(anchor="w")
#
#         # 密码保护
#         tk.Label(options_frame, text="PDF密码（如果有）:").pack(anchor="w", pady=(5,0))
#         self.pdf_password = tk.StringVar()
#         tk.Entry(options_frame, textvariable=self.pdf_password,
#                 width=30, show="*").pack(anchor="w")
#
#         # 页码范围
#         tk.Label(options_frame, text="页码范围（如: 1-5）:").pack(anchor="w", pady=(5,0))
#         self.page_range = tk.StringVar()
#         tk.Entry(options_frame, textvariable=self.page_range,
#                 width=30).pack(anchor="w")
#
#         # 进度条
#         self.progress = ttk.Progressbar(self.window, length=400,
#                                        mode='determinate')
#         self.progress.pack(pady=20)
#         self.progress_label = tk.Label(self.window, text="")
#         self.progress_label.pack()
#
#         # 按钮
#         button_frame = tk.Frame(self.window)
#         button_frame.pack(pady=20)
#
#         tk.Button(button_frame, text="开始转换",
#                  command=self.start_conversion,
#                  bg="#4CAF50", fg="white",
#                  width=15).pack(side=tk.LEFT, padx=10)
#
#         tk.Button(button_frame, text="停止",
#                  command=self.stop_conversion,
#                  width=15).pack(side=tk.LEFT, padx=10)
#
#         tk.Button(button_frame, text="清空",
#                  command=self.clear_fields,
#                  width=15).pack(side=tk.LEFT, padx=10)
#
#     def select_pdf(self):
#         file_path = filedialog.askopenfilename(
#             filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
#         )
#         if file_path:
#             self.pdf_path.set(file_path)
#             # 自动生成输出路径
#             output = file_path.rsplit('.', 1)[0] + "_converted.docx"
#             self.output_path.set(output)
#
#     def select_output(self):
#         file_path = filedialog.asksaveasfilename(
#             defaultextension=".docx",
#             filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
#         )
#         if file_path:
#             self.output_path.set(file_path)
#
#     def ocr_pdf_to_word(self, pdf_path, output_path, password=None):
#         """使用OCR处理扫描版PDF"""
#         try:
#             # 将PDF转换为图片
#             images = convert_from_path(pdf_path,
#                                      dpi=300,
#                                      first_page=1,
#                                      last_page=None)
#
#             from docx import Document
#             doc = Document()
#
#             total_pages = len(images)
#             for i, image in enumerate(images, 1):
#                 # 更新进度
#                 progress = i / total_pages
#                 self.update_progress(progress, f"OCR处理中: {i}/{total_pages}")
#
#                 # 使用Tesseract进行OCR
#                 text = pytesseract.image_to_string(image, lang='chi_sim+eng')
#
#                 # 添加到Word
#                 if text.strip():
#                     doc.add_paragraph(text)
#                     doc.add_page_break()
#
#             doc.save(output_path)
#             return True
#
#         except Exception as e:
#             messagebox.showerror("OCR错误", f"OCR处理失败: {str(e)}")
#             return False
#
#     def convert_with_pdf2docx(self, pdf_path, output_path, password=None, page_range=None):
#         """使用pdf2docx转换"""
#         try:
#             cv = Converter(pdf_path)
#
#             # 解析页码范围
#             start, end = 0, None
#             if page_range:
#                 if '-' in page_range:
#                     parts = page_range.split('-')
#                     start = int(parts[0]) - 1 if parts[0] else 0
#                     end = int(parts[1]) if parts[1] else None
#
#             # 转换
#             cv.convert(output_path, start=start, end=end)
#             cv.close()
#             return True
#
#         except Exception as e:
#             messagebox.showerror("转换错误", f"转换失败: {str(e)}")
#             return False
#
#     def update_progress(self, value, text=""):
#         """更新进度"""
#         self.progress['value'] = value * 100
#         self.progress_label.config(text=text)
#         self.window.update()
#
#     def start_conversion(self):
#         """开始转换"""
#         pdf_path = self.pdf_path.get()
#         output_path = self.output_path.get()
#
#         if not pdf_path or not os.path.exists(pdf_path):
#             messagebox.showerror("错误", "请选择有效的PDF文件")
#             return
#
#         if not output_path:
#             messagebox.showerror("错误", "请指定输出路径")
#             return
#
#         # 在新线程中运行转换
#         thread = threading.Thread(target=self.convert_thread,
#                                 args=(pdf_path, output_path))
#         thread.daemon = True
#         thread.start()
#
#     def convert_thread(self, pdf_path, output_path):
#         """转换线程"""
#         try:
#             self.update_progress(0, "开始转换...")
#
#             if self.use_ocr.get():
#                 # OCR模式
#                 success = self.ocr_pdf_to_word(pdf_path, output_path)
#             else:
#                 # 普通模式
#                 password = self.pdf_password.get() or None
#                 page_range = self.page_range.get() or None
#                 success = self.convert_with_pdf2docx(pdf_path, output_path,
#                                                    password, page_range)
#
#             if success:
#                 self.update_progress(1, "转换完成！")
#                 messagebox.showinfo("成功", f"转换完成！\n保存至: {output_path}")
#             else:
#                 self.update_progress(0, "转换失败")
#
#         except Exception as e:
#             messagebox.showerror("错误", f"转换失败: {str(e)}")
#             self.update_progress(0, "转换失败")
#
#     def stop_conversion(self):
#         """停止转换"""
#         # 目前pdf2docx不支持停止，这里只是UI操作
#         self.progress['value'] = 0
#         self.progress_label.config(text="已停止")
#
#     def clear_fields(self):
#         """清空所有字段"""
#         self.pdf_path.set("")
#         self.output_path.set("")
#         self.pdf_password.set("")
#         self.page_range.set("")
#         self.progress['value'] = 0
#         self.progress_label.config(text="")
#
# # 运行高级版本
# if __name__ == "__main__":
#     # 注意：需要先安装Tesseract和poppler
#     app = AdvancedPDFConverter()
#     app.window.mainloop()


"""
先安装：pip install pdf2docx
这个库能更好地保留格式、表格、图片等
"""
from pdf2docx import Converter
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os


class PDFtoWordConverter:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("PDF转Word工具")
        self.window.geometry("500x400")

        # 进度条变量
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="准备就绪")

        self.setup_ui()

    def setup_ui(self):
        # 标题
        tk.Label(self.window, text="PDF转Word转换器",
                 font=("微软雅黑", 16)).pack(pady=20)

        # PDF文件选择
        tk.Label(self.window, text="选择PDF文件:").pack()

        self.pdf_path_var = tk.StringVar()
        tk.Entry(self.window, textvariable=self.pdf_path_var,
                 width=50).pack(pady=5)

        tk.Button(self.window, text="浏览",
                  command=self.browse_pdf).pack(pady=5)

        # Word保存位置
        tk.Label(self.window, text="保存Word文件到:").pack(pady=(10, 0))

        self.word_path_var = tk.StringVar()
        tk.Entry(self.window, textvariable=self.word_path_var,
                 width=50).pack(pady=5)

        tk.Button(self.window, text="浏览",
                  command=self.browse_word_save).pack(pady=5)

        # 转换选项
        self.convert_options_frame = tk.Frame(self.window)
        self.convert_options_frame.pack(pady=10)

        # 转换质量选项
        tk.Label(self.convert_options_frame, text="转换质量:").grid(row=0, column=0, sticky="w")
        self.quality_var = tk.StringVar(value="medium")
        qualities = [("高质量（慢）", "high"), ("中等质量", "medium"), ("快速转换", "fast")]
        for i, (text, value) in enumerate(qualities):
            tk.Radiobutton(self.convert_options_frame, text=text,
                           variable=self.quality_var, value=value).grid(row=0, column=i + 1, padx=5)

        # 进度条
        tk.Label(self.window, text="进度:").pack()
        self.progress = ttk.Progressbar(self.window,
                                        variable=self.progress_var,
                                        maximum=100,
                                        length=300,
                                        mode='determinate')
        self.progress.pack(pady=5)

        # 状态标签
        self.status_label = tk.Label(self.window,
                                     textvariable=self.status_var,
                                     fg="blue")
        self.status_label.pack(pady=5)

        # 按钮
        button_frame = tk.Frame(self.window)
        button_frame.pack(pady=20)

        tk.Button(button_frame, text="开始转换",
                  command=self.start_conversion,
                  bg="green", fg="white",
                  width=15).pack(side=tk.LEFT, padx=10)

        tk.Button(button_frame, text="批量转换",
                  command=self.batch_conversion,
                  width=15).pack(side=tk.LEFT, padx=10)

        tk.Button(button_frame, text="退出",
                  command=self.window.quit,
                  width=15).pack(side=tk.LEFT, padx=10)

    def browse_pdf(self):
        filename = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if filename:
            self.pdf_path_var.set(filename)
            # 自动生成Word文件名
            word_path = filename.rsplit('.', 1)[0] + '.docx'
            self.word_path_var.set(word_path)

    def browse_word_save(self):
        filename = filedialog.asksaveasfilename(
            title="保存Word文件",
            defaultextension=".docx",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
        )
        if filename:
            self.word_path_var.set(filename)

    def update_progress(self, progress):
        """更新进度条"""
        self.progress_var.set(progress * 100)
        self.status_var.set(f"转换中: {int(progress * 100)}%")
        self.window.update()

    def convert_pdf_to_word(self, pdf_path, word_path):
        """执行转换"""
        try:
            self.status_var.set("正在初始化...")
            self.window.update()

            # 创建转换器
            cv = Converter(pdf_path)

            # 根据质量选择转换参数
            quality = self.quality_var.get()
            if quality == "high":
                kwargs = {'debug': False, 'multi_processing': False}
            elif quality == "fast":
                kwargs = {'debug': False, 'multi_processing': True}
            else:  # medium
                kwargs = {'debug': False}

            # 开始转换
            cv.convert(word_path, start=0, end=None, **kwargs)
            cv.close()

            self.status_var.set("转换完成！")
            messagebox.showinfo("成功", f"转换完成！\n保存至: {word_path}")
            self.progress_var.set(100)

            # 是否打开文件
            if messagebox.askyesno("完成", "是否打开转换后的Word文件？"):
                os.startfile(word_path)

            return True

        except Exception as e:
            messagebox.showerror("错误", f"转换失败: {str(e)}")
            self.status_var.set("转换失败")
            return False
        finally:
            self.progress_var.set(0)

    def start_conversion(self):
        """开始转换（单文件）"""
        pdf_path = self.pdf_path_var.get()
        word_path = self.word_path_var.get()

        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("错误", "请选择有效的PDF文件！")
            return

        # 在新线程中运行转换
        thread = threading.Thread(
            target=self.convert_pdf_to_word,
            args=(pdf_path, word_path)
        )
        thread.daemon = True
        thread.start()

    def batch_conversion(self):
        """批量转换"""
        pdf_files = filedialog.askopenfilenames(
            title="选择多个PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )

        if not pdf_files:
            return

        # 选择保存目录
        save_dir = filedialog.askdirectory(title="选择保存目录")
        if not save_dir:
            return

        # 批量转换
        success_count = 0
        for i, pdf_file in enumerate(pdf_files, 1):
            self.status_var.set(f"正在转换 {i}/{len(pdf_files)}: {os.path.basename(pdf_file)}")
            self.progress_var.set((i - 1) / len(pdf_files) * 100)
            self.window.update()

            # 生成输出路径
            filename = os.path.basename(pdf_file).rsplit('.', 1)[0] + '.docx'
            output_path = os.path.join(save_dir, filename)

            try:
                cv = Converter(pdf_file)
                cv.convert(output_path)
                cv.close()
                success_count += 1
            except Exception as e:
                print(f"转换失败 {pdf_file}: {e}")

        self.progress_var.set(100)
        self.status_var.set("批量转换完成")
        messagebox.showinfo("完成",
                            f"批量转换完成！\n成功: {success_count}/{len(pdf_files)} 个文件")


# 运行GUI
if __name__ == "__main__":
    app = PDFtoWordConverter()
    app.window.mainloop()