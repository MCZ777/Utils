import os
import sys
import pandas as pd
from openpyxl import load_workbook
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import tkinterdnd2 as tkdnd
import time

class DragDropGUI(tkdnd.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("Excel文件合并工具")
        self.geometry("800x600")  # 增加窗口尺寸
        self.configure(padx=30, pady=30)  # 增加内边距
        
        # 设置字体 - 使用更美观的字体设置
        self.title_font = ("微软雅黑", 24, "bold")     # 增大标题字体
        self.hint_font = ("微软雅黑", 12)              # 增大提示文字
        self.normal_font = ("微软雅黑", 11)            # 增大普通文本
        
        self.template_path = None
        self.input_dir = None
        self.output_path = None
        self.create_widgets()
        # 设置默认提示文本
        self.save_path_entry.insert(0, "请输入或选择保存位置")
        
    def create_widgets(self):
        # 标题和说明
        title_frame = ttk.Frame(self)
        title_frame.pack(fill=tk.X, pady=(0, 15))
        
        title_label = ttk.Label(title_frame, text="Excel文件合并工具", font=self.title_font)
        title_label.pack()
        
        hint_label = ttk.Label(title_frame, 
                              text="请按顺序选择：1.模板文件 2.输入文件夹 3.保存位置",
                              font=self.hint_font)
        hint_label.pack(pady=(10, 0))
        
        # 创建拖拽区域的主框架
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧拖拽区域
        left_frame = ttk.LabelFrame(main_frame, text="模板文件", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 添加文件名和删除按钮的框架
        template_name_frame = ttk.Frame(left_frame)
        template_name_frame.pack(fill=tk.X)
        
        self.template_name_label = ttk.Label(template_name_frame, text="", anchor="w", font=self.normal_font)
        self.template_name_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.template_clear_button = ttk.Label(template_name_frame, text="×", cursor="hand2", font=self.normal_font)
        self.template_clear_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.template_clear_button.bind('<Button-1>', lambda e: self.clear_template())
        self.template_clear_button.configure(foreground='red')  # 设置×为红色
        self.template_clear_button.pack_forget()  # 初始时隐藏
        
        self.template_label = ttk.Label(left_frame, 
                                      text="请拖入模板文件\n或点击选择",
                                      background='#F0F0F0',
                                      padding=(10, 30),  # 替换 height，使用 padding 来控制高度
                                      justify='center',
                                      font=self.normal_font)
        self.template_label.pack(fill=tk.BOTH, expand=True)
        self.template_label.drop_target_register(tkdnd.DND_FILES)
        self.template_label.dnd_bind('<<Drop>>', self.drop_template)
        self.template_label.bind('<Button-1>', lambda e: self.select_template())
        
        # 右侧拖拽区域
        right_frame = ttk.LabelFrame(main_frame, text="输入文件/文件夹", padding=10)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # 添加文件名和删除按钮的框架
        input_name_frame = ttk.Frame(right_frame)
        input_name_frame.pack(fill=tk.X)
        
        self.input_name_label = ttk.Label(input_name_frame, text="", anchor="w", font=self.normal_font)
        self.input_name_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.input_clear_button = ttk.Label(input_name_frame, text="×", cursor="hand2", font=self.normal_font)
        self.input_clear_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.input_clear_button.bind('<Button-1>', lambda e: self.clear_input())
        self.input_clear_button.configure(foreground='red')  # 设置×为红色
        self.input_clear_button.pack_forget()  # 初始时隐藏
        
        self.input_label = ttk.Label(right_frame,
                                   text="请拖入Excel文件或文件夹\n或点击选择",
                                   background='#F0F0F0',
                                   padding=(10, 30),  # 替换 height，使用 padding 来控制高度
                                   justify='center',
                                   font=self.normal_font)
        self.input_label.pack(fill=tk.BOTH, expand=True)
        self.input_label.drop_target_register(tkdnd.DND_FILES)
        self.input_label.dnd_bind('<<Drop>>', self.drop_input)
        self.input_label.bind('<Button-1>', lambda e: self.select_input_dir())
        
        # 底部区域使用Grid布局
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))
        bottom_frame.grid_columnconfigure(0, weight=1)
        
        # 保存路径显示和选择按钮
        save_frame = ttk.LabelFrame(bottom_frame, text="保存位置", padding=5)
        save_frame.pack(fill=tk.X, pady=(0, 10))
        save_frame.grid_columnconfigure(0, weight=1)
        
        # 将Label改为Entry
        self.save_path_entry = ttk.Entry(save_frame)
        self.save_path_entry.grid(row=0, column=0, sticky='ew', padx=5)
        self.save_path_entry.bind('<Return>', lambda e: self.validate_save_path())  # 按回车时验证路径
        self.save_path_entry.bind('<FocusOut>', lambda e: self.validate_save_path())  # 失去焦点时验证路径
        
        self.save_button = ttk.Button(
            save_frame,
            text="选择位置",
            command=self.select_save_location,
            width=10
        )
        self.save_button.grid(row=0, column=1, padx=5)
        
        # 状态显示和合并按钮
        status_frame = ttk.Frame(bottom_frame)
        status_frame.pack(fill=tk.X)
        status_frame.grid_columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(status_frame, text="请选择模板文件", foreground="red")
        self.status_label.grid(row=0, column=0, sticky='w', padx=5)
        
        self.merge_button = ttk.Button(
            status_frame,
            text="开始合并",
            command=self.start_merge,
            width=20
        )
        self.merge_button.grid(row=0, column=1, padx=5)
    
    def drop_template(self, event):
        file_path = event.data
        file_path = file_path.strip('{}')
        if file_path.lower().endswith('.xlsx'):
            self.template_path = file_path
            self.template_name_label.configure(text=f"已选择: {os.path.basename(self.template_path)}")
            self.template_label.configure(text="")
            self.template_clear_button.pack(side=tk.RIGHT, padx=(5, 0))  # 显示删除按钮
            self.update_status()
        else:
            messagebox.showwarning("警告", "请选择Excel文件(.xlsx)作为模板")
    
    def drop_input(self, event):
        paths = event.data.split('}') if '}' in event.data else [event.data]
        paths = [p.strip('{ ') for p in paths if p.strip('{ ')]  # 清理路径字符串
        
        valid_paths = []
        for path in paths:
            if os.path.isfile(path) and path.lower().endswith('.xlsx'):
                valid_paths.append(path)
            elif os.path.isdir(path):
                valid_paths.append(path)
        
        if valid_paths:
            self.input_dir = valid_paths
            if len(valid_paths) == 1 and os.path.isdir(valid_paths[0]):
                folder_files = [f for f in os.listdir(valid_paths[0]) if f.lower().endswith('.xlsx')]
                if folder_files:
                    self.input_name_label.configure(text=f"已选择文件夹，包含 {len(folder_files)} 个Excel文件")
                else:
                    self.input_name_label.configure(text="已选择文件夹 (文件夹为空)")
            else:
                self.input_name_label.configure(text=f"已选择 {len(valid_paths)} 个Excel文件")
            self.input_label.configure(text="")
            self.input_clear_button.pack(side=tk.RIGHT, padx=(5, 0))  # 显示删除按钮
            self.update_status()
        else:
            messagebox.showwarning("警告", "请选择Excel文件(.xlsx)或文件夹")
    
    def select_template(self):
        path = filedialog.askopenfilename(
            title="选择模板文件",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if path:
            self.template_path = path
            self.template_name_label.configure(text=f"已选择: {os.path.basename(path)}")
            self.template_label.configure(text="")
            self.template_clear_button.pack(side=tk.RIGHT, padx=(5, 0))  # 显示删除按钮
            self.update_status()
    
    def select_input_dir(self):
        # 先尝试选择文件
        paths = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        
        if paths:  # 如果选择了文件
            self.input_dir = list(paths)
            self.update_input_text(f"已选择 {len(paths)} 个Excel文件")
            self.input_label.configure(text="")
            self.update_status()
        else:  # 如果没选文件，尝试选择文件夹
            path = filedialog.askdirectory(title="选择文件夹")
            if path:
                self.input_dir = [path]
                # 获取文件夹中的Excel文件数量
                folder_files = [f for f in os.listdir(path) if f.lower().endswith('.xlsx')]
                if folder_files:
                    self.update_input_text(f"已选择文件夹，包含 {len(folder_files)} 个Excel文件")
                else:
                    self.update_input_text("已选择文件夹 (文件夹为空)")
                self.input_label.configure(text="")
                self.update_status()
    
    def select_save_location(self):
        output_path = filedialog.asksaveasfilename(
            title="选择保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if output_path:
            self.output_path = output_path
            # 更新Entry中的文本
            self.save_path_entry.delete(0, tk.END)
            self.save_path_entry.insert(0, output_path)
            self.update_status()
    
    def update_status(self):
        if not self.template_path:
            self.status_label.configure(text="请选择模板文件", foreground="red")
            self.merge_button.state(['disabled'])
        elif not self.input_dir:
            self.status_label.configure(text="请选择输入文件夹", foreground="red")
            self.merge_button.state(['disabled'])
        elif not self.output_path:
            self.status_label.configure(text="请选择保存位置", foreground="red")
            self.merge_button.state(['disabled'])
        else:
            self.status_label.configure(text="可以开始合并了！", foreground="green")
            self.merge_button.state(['!disabled'])
    
    def start_merge(self):
        if not (self.template_path and self.input_dir):
            return
            
        if not self.output_path:
            self.select_save_location()
            if not self.output_path:
                return
            
        merge_excel_files(self.template_path, self.input_dir, self.output_path)

    def update_template_text(self, text):
        self.template_name_label.configure(text=text)

    def update_input_text(self, text):
        """更新输入文件显示文本"""
        self.input_name_label.configure(text=text)

    def validate_save_path(self):
        """验证手动输入的保存路径"""
        path = self.save_path_entry.get().strip()
        if not path:
            self.output_path = None
            self.update_status()
            return False
            
        # 确保路径以.xlsx结尾
        if not path.lower().endswith('.xlsx'):
            path += '.xlsx'
            self.save_path_entry.delete(0, tk.END)
            self.save_path_entry.insert(0, path)
        
        try:
            # 检查目录是否存在
            directory = os.path.dirname(path)
            if directory and not os.path.exists(directory):
                messagebox.showwarning("警告", "保存路径所在文件夹不存在！")
                return False
                
            self.output_path = path
            self.update_status()
            return True
        except Exception as e:
            messagebox.showwarning("警告", f"无效的保存路径: {str(e)}")
            return False

    def clear_template(self):
        """清除已选择的模板文件"""
        self.template_path = None
        self.template_name_label.configure(text="")
        self.template_label.configure(text="请拖入模板文件\n或点击选择")
        self.template_clear_button.pack_forget()  # 隐藏删除按钮
        self.update_status()

    def clear_input(self):
        """清除已选择的输入文件"""
        self.input_dir = None
        self.input_name_label.configure(text="")
        self.input_label.configure(text="请拖入Excel文件或文件夹\n或点击选择")
        self.input_clear_button.pack_forget()  # 隐藏删除按钮
        self.update_status()

def validate_template(template_path, input_path):
    """验证输入文件是否符合模板格式"""
    try:
        template = pd.read_excel(template_path, nrows=1)
        input_file = pd.read_excel(input_path, nrows=1)
        return list(template.columns) == list(input_file.columns)
    except Exception as e:
        messagebox.showerror("错误", f"验证文件 {input_path} 时出错: {str(e)}")
        return False

def merge_excel_files(template_path, input_paths, output_path):
    """合并Excel文件"""
    try:
        # 获取所有输入文件
        input_files = []
        for path in input_paths:
            if os.path.isfile(path):
                input_files.append(path)
            else:
                folder_files = [os.path.join(path, f) 
                              for f in os.listdir(path) 
                              if f.endswith('.xlsx')]
                input_files.extend(folder_files)
        
        if not input_files:
            messagebox.showwarning("警告", "没有找到Excel文件！")
            return
        
        # 读取模板文件以获取字体样式
        template_wb = load_workbook(template_path)
        template_ws = template_wb.active
        header_font = None
        
        # 获取第一行（表头）的字体属性
        for cell in template_ws[1]:
            if cell.font:
                header_font = {
                    'name': cell.font.name,
                    'size': cell.font.size,
                    'bold': cell.font.bold,
                    'italic': cell.font.italic,
                    'color': cell.font.color
                }
                break
        
        all_data = pd.DataFrame()
        success_count = 0
        failed_files = []  # 记录失败的文件
        
        for file_path in input_files:
            if validate_template(template_path, file_path):
                df = pd.read_excel(file_path)
                df['来源文件'] = os.path.basename(file_path)
                all_data = pd.concat([all_data, df], ignore_index=True)
                success_count += 1
            else:
                failed_files.append(os.path.basename(file_path))  # 添加到失败列表
                messagebox.showwarning("警告", 
                    f"文件 {os.path.basename(file_path)} 不符合模板格式，已跳过")
        
        # 如果有失败的文件，保存失败记录
        if failed_files:
            # 获取输出文件所在目录
            output_dir = os.path.dirname(output_path)
            # 生成失败记录文件名（使用当前时间）
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            fail_log_path = os.path.join(output_dir, f'合并失败记录_{timestamp}.txt')
            
            # 写入失败记录
            with open(fail_log_path, 'w', encoding='utf-8') as f:
                f.write(f"合并失败的文件列表 (共 {len(failed_files)} 个)：\n")
                f.write("="*50 + "\n")
                for i, file_name in enumerate(failed_files, 1):
                    f.write(f"{i}. {file_name}\n")
                f.write("\n失败原因：文件格式与模板不符\n")
                f.write(f"\n记录时间：{time.strftime('%Y-%m-%d %H:%M:%S')}")
        
        if success_count > 0:
            # 保存数据
            all_data.to_excel(output_path, sheet_name='合并数据', index=False)
            
            # 应用字体和对齐方式
            wb = load_workbook(output_path)
            ws = wb.active
            
            if header_font:
                from openpyxl.styles import Font, Alignment
                # 创建表头和内容的字体样式
                header_font_obj = Font(
                    name=header_font['name'],
                    size=header_font['size'],
                    bold=header_font['bold'],
                    italic=header_font['italic'],
                    color=header_font['color']
                )
                
                content_font_obj = Font(
                    name=header_font['name'],
                    size=header_font['size'],
                    bold=False,  # 内容不使用粗体
                    italic=header_font['italic'],
                    color=header_font['color']
                )
                
                # 创建居中对齐样式
                center_alignment = Alignment(
                    horizontal='center',
                    vertical='center'
                )
                
                # 为所有单元格应用字体和对齐样式
                for row in ws.rows:
                    for cell in row:
                        cell.alignment = center_alignment  # 所有单元格居中对齐
                        if cell.row == 1:  # 表头行
                            cell.font = header_font_obj
                        else:  # 内容行
                            cell.font = content_font_obj
                
                # 合并来源文件列中相同的单元格
                last_col = ws.max_column  # 最后一列（来源文件列）
                current_file = None
                start_row = None
                
                for row in range(2, ws.max_row + 1):  # 从第2行开始（跳过表头）
                    cell_value = ws.cell(row=row, column=last_col).value
                    
                    if current_file is None:  # 第一个文件
                        current_file = cell_value
                        start_row = row
                    elif cell_value != current_file:  # 不同文件
                        if row - start_row > 1:  # 如果有多行相同文件
                            ws.merge_cells(
                                start_row=start_row,
                                start_column=last_col,
                                end_row=row - 1,
                                end_column=last_col
                            )
                        current_file = cell_value
                        start_row = row
                
                # 处理最后一组相同文件
                if start_row and row - start_row >= 0:
                    ws.merge_cells(
                        start_row=start_row,
                        start_column=last_col,
                        end_row=ws.max_row,
                        end_column=last_col
                    )
            
            # 保存修改后的文件
            wb.save(output_path)
            
            # 根据是否有失败文件显示不同的成功消息
            if failed_files:
                messagebox.showinfo("部分成功", 
                    f"已成功合并 {success_count} 个文件到 {output_path}\n"
                    f"有 {len(failed_files)} 个文件合并失败\n"
                    f"失败记录已保存到：{fail_log_path}")
            else:
                messagebox.showinfo("成功", f"已成功合并 {success_count} 个文件到 {output_path}")
        else:
            messagebox.showwarning("警告", "没有成功合并任何文件！")
            
    except Exception as e:
        messagebox.showerror("错误", f"合并文件时出错: {str(e)}")

def main():
    app = DragDropGUI()
    app.mainloop()

if __name__ == "__main__":
    main()