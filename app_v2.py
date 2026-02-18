import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import zipfile
from tkinterdnd2 import DND_FILES, TkinterDnD

class ImageExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Office 图片提取工具")
        self.root.geometry("600x450")
        
        # 变量
        self.file_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        
        # UI 布局
        self.create_widgets()
        
    def create_widgets(self):
        # 顶部说明
        header_label = tk.Label(self.root, text="支持拖拽文件到这里 (PPT/Word)", font=("微软雅黑", 12))
        header_label.pack(pady=10)
        
        # 文件选择区域
        input_frame = tk.LabelFrame(self.root, text="第一步: 选择文件 (或是直接拖入)", padx=10, pady=10)
        input_frame.pack(fill="x", padx=10, pady=5)
        
        # 支持拖拽
        self.file_entry = tk.Entry(input_frame, textvariable=self.file_path, width=40)
        self.file_entry.pack(side=tk.LEFT, padx=5, expand=True, fill="x")
        
        # register drop target
        try:
            self.file_entry.drop_target_register(DND_FILES)
            self.file_entry.dnd_bind('<<Drop>>', self.drop_file)
            
            # Allow dropping on the whole window as well
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.drop_file)
        except Exception as e:
            print(f"DnD setup failed: {e}")
            
        tk.Button(input_frame, text="浏览...", command=self.select_file).pack(side=tk.LEFT)
        
        # 输出目录选择区域
        output_frame = tk.LabelFrame(self.root, text="第二步: 选择保存位置 (可选)", padx=10, pady=10)
        output_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Entry(output_frame, textvariable=self.output_dir, width=40).pack(side=tk.LEFT, padx=5, expand=True, fill="x")
        tk.Button(output_frame, text="浏览...", command=self.select_output_dir).pack(side=tk.LEFT)
        
        # 说明文字
        tk.Label(output_frame, text="留空默认保存在源文件同级目录", fg="gray", font=("微软雅黑", 8)).pack(side=tk.LEFT, padx=5)

        # 提取按钮
        self.extract_btn = tk.Button(self.root, text="开始提取图片", command=self.start_extraction, 
                                   bg="#0078D7", fg="black", font=("微软雅黑", 12, "bold"), height=2)
        self.extract_btn.pack(fill="x", padx=20, pady=20)
        
        # 状态栏
        self.status_label = tk.Label(self.root, text="准备就绪", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def drop_file(self, event):
        file_path = event.data
        # Windows/TkinterDnD often returns path in {} if it contains spaces
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
            
        if os.path.isfile(file_path):
            if file_path.lower().endswith(('.pptx', '.docx')):
                self.file_path.set(file_path)
            else:
                messagebox.showwarning("格式错误", "仅支持 .pptx 和 .docx 文件")

    def select_file(self):
        file_selected = filedialog.askopenfilename(
            filetypes=[("Office 文件", "*.pptx;*.docx"), ("PowerPoint", "*.pptx"), ("Word", "*.docx")]
        )
        if file_selected:
            self.file_path.set(file_selected)
            # 默认输出目录设为文件所在目录
            if not self.output_dir.get():
                self.output_dir.set(os.path.dirname(file_selected))

    def select_output_dir(self):
        dir_selected = filedialog.askdirectory()
        if dir_selected:
            self.output_dir.set(dir_selected)

    def log(self, message):
        # 在主线程更新UI
        self.root.after(0, lambda: self.status_label.config(text=message))

    def start_extraction(self):
        source_file = self.file_path.get()
        target_dir = self.output_dir.get()
        
        if not source_file or not os.path.exists(source_file):
            messagebox.showerror("错误", "请选择有效的源文件！")
            return
            
        # 如果没有指定输出目录，默认为源文件所在目录
        if not target_dir:
            target_dir = os.path.dirname(source_file)
            
        # 禁用按钮防止重复点击
        self.extract_btn.config(state=tk.DISABLED, text="正在处理...")
        
        # 在新线程中运行提取任务
        threading.Thread(target=self.run_extraction_task, args=(source_file, target_dir), daemon=True).start()

    def run_extraction_task(self, source_file, target_dir):
        try:
            # 创建存放图片的子文件夹
            filename_base = os.path.splitext(os.path.basename(source_file))[0]
            images_dir = os.path.join(target_dir, f"{filename_base}_images")
            
            if not os.path.exists(images_dir):
                os.makedirs(images_dir)
            
            count = 0
            # 使用通用的基于zip的提取方法，保证提取原图
            count = self.extract_images_from_zip(source_file, images_dir)

            if count > 0:
                self.root.after(0, lambda: messagebox.showinfo("成功", f"提取完成！\n共保存 {count} 张图片到：\n{images_dir}"))
                self.root.after(0, lambda: os.system(f'explorer "{images_dir}"'))
            else:
                self.root.after(0, lambda: messagebox.showinfo("提示", "未在文件中找到图片。"))
                
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"发生错误：{str(e)}"))
        finally:
            self.root.after(0, self.reset_ui)

    def reset_ui(self):
        self.extract_btn.config(state=tk.NORMAL, text="开始提取图片")
        self.log("准备就绪")

    def extract_images_from_zip(self, file_path, output_folder):
        """
        通用的 Zip 提取方法，适用于 .docx 和 .pptx
        直接解压 media 文件夹中的内容，获取原图
        """
        image_count = 0
        self.log(f"正在分析文件结构: {os.path.basename(file_path)}")
        
        try:
            with zipfile.ZipFile(file_path) as z:
                all_files = z.namelist()
                
                # 查找媒体文件
                # Word: word/media/
                # PPT: ppt/media/
                media_files = [f for f in all_files if f.startswith("word/media/") or f.startswith("ppt/media/")]
                
                total_media = len(media_files)
                self.log(f"找到 {total_media} 个潜在媒体文件...")

                for media_file in media_files:
                    # 获取扩展名
                    ext = os.path.splitext(media_file)[1].lower()
                    # 常见图片格式
                    valid_exts = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.wmf', '.emf', '.svg', '.wdp', '.tif']
                    if ext not in valid_exts:
                        continue
                        
                    image_count += 1
                    self.log(f"正在提取第 {image_count} 张图片 ({ext})...")
                    
                    # 读取图片数据
                    image_data = z.read(media_file)
                    
                    # 使用原始文件名（去除路径）
                    original_name = os.path.basename(media_file)
                    name_without_ext, original_ext = os.path.splitext(original_name)
                    
                    # 直接保存原文件，不进行转换
                    target_path = os.path.join(output_folder, original_name)
                    
                    # 避免同名覆盖
                    if os.path.exists(target_path):
                        target_path = os.path.join(output_folder, f"{name_without_ext}_{image_count}{original_ext}")

                    with open(target_path, 'wb') as f:
                        f.write(image_data)
                        
        except Exception as e:
            if "BadZipFile" in str(e) or "zip" in str(e).lower():
                 raise Exception("文件损坏或不是有效的 Office Open XML 文件")
            raise Exception(f"文件读取失败: {e}")
                    
        return image_count

    # 保留旧方法占位（可选删除）
    def extract_from_ppt(self, ppt_path, output_folder):
        return self.extract_images_from_zip(ppt_path, output_folder)

    def extract_from_word(self, docx_path, output_folder):
        return self.extract_images_from_zip(docx_path, output_folder)

if __name__ == "__main__":
    try:
        # 尝试设置DPI感知，防止在高分屏下模糊
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
        
    root = TkinterDnD.Tk()
    app = ImageExtractorApp(root)
    root.mainloop()
