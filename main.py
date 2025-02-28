import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import csv
from datetime import datetime
from pypinyin import pinyin, Style

# 定义帮助文本
help_text = """使用说明：

1. 搜索功能：
   - 在左上方输入框可以按标题、作者或朝代搜索诗词
   - 支持模糊搜索，输入部分内容即可
   - 按回车键或点击搜索按钮进行搜索

2. 诗词列表：
   - 左侧显示所有诗词的列表
   - 点击任意诗词可在右侧查看详细内容

3. 诗词详情：
   - 右侧上方显示诗词原文和拼音
   - 下方标签页包含译文、注释、赏析和作者介绍

4. 编辑功能：
   - 点击"编辑"按钮可以修改诗词内容
   - 修改完成后点击"保存"按钮保存更改
   - 点击"取消"可放弃更改

5. 导入导出：
   - 可以导入其他JSON格式的诗词数据
   - 可以导出当前所有诗词数据为JSON文件"""

# 定义关于文本
about_text = """少儿古诗词学习 v1.0

本软件旨在帮助儿童学习中国古诗词，提供了以下功能：
- 诗词原文与拼音对照
- 译文与注释辅助理解
- 诗词赏析提升鉴赏能力
- 作者介绍了解历史背景

特点：
- 简洁直观的界面设计
- 方便的搜索和编辑功能
- 支持数据导入导出

作者：奋青
联系方式：393283@qq.com

免责声明：
本软件为开源项目，按"原样"提供，不提供任何明示或暗示的保证。作者不对使用本软件造成的任何直接或间接损失负责。用户可以自由使用、修改和分发本软件，但必须遵守相关开源协议的规定。

版权所有 © 2024"""

class PoemApp:
    def __init__(self, root):
        self.root = root
        self.root.title('少儿古诗词学习')
        self.editing = False
        
        # 使用更大的窗口尺寸
        width = 1400
        height = 800
        self.root.geometry(f'{width}x{height}')
        
        # 设置窗口居中
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f'{width}x{height}+{x}+{y}')

        # 创建标题栏
        title_frame = ttk.Frame(root)
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # 系统标题居中显示，使用更大字体
        self.title_label = ttk.Label(title_frame, text='少儿古诗词学习', font=('SimSun', 28, 'bold'))
        self.title_label.pack(side=tk.TOP, pady=10)

        # 标题居中显示
        self.title_label.pack(side=tk.TOP, pady=10, expand=True)

        # 创建帮助和关于按钮，放在右侧
        button_frame = ttk.Frame(title_frame)
        button_frame.pack(side=tk.RIGHT, padx=10)
        
        self.help_btn = ttk.Button(button_frame, text='帮助', command=lambda: messagebox.showinfo('帮助', help_text), 
                                  style='Borderless.TButton', width=4)
        self.help_btn.pack(side=tk.LEFT, padx=1)
        
        self.about_btn = ttk.Button(button_frame, text='关于', command=lambda: messagebox.showinfo('关于', about_text), 
                                   style='Borderless.TButton', width=4)
        self.about_btn.pack(side=tk.LEFT, padx=1)
        
        # 标题居中显示
        self.title_label.pack(side=tk.TOP, pady=10, expand=True)

        # 加载诗词数据
        with open('poems.json', 'r', encoding='utf-8') as f:
            self.poems_data = json.load(f)['poems']

        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)

        # 创建左右分栏
        self.left_frame = ttk.Frame(self.main_frame)
        self.left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        self.right_frame = ttk.Frame(self.main_frame)
        self.right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 创建搜索框架
        self.search_frame = ttk.Frame(self.left_frame)
        self.search_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        # 创建搜索条件输入框
        ttk.Label(self.search_frame, text="标题:").grid(row=0, column=0, padx=5)
        self.title_search = ttk.Entry(self.search_frame, width=15)
        self.title_search.grid(row=0, column=1, padx=5)
        self.title_search.bind('<Return>', lambda e: self.search_poems())

        ttk.Label(self.search_frame, text="作者:").grid(row=0, column=2, padx=5)
        self.author_search = ttk.Entry(self.search_frame, width=10)
        self.author_search.grid(row=0, column=3, padx=5)
        self.author_search.bind('<Return>', lambda e: self.search_poems())

        ttk.Label(self.search_frame, text="朝代:").grid(row=0, column=4, padx=5)
        self.dynasty_search = ttk.Entry(self.search_frame, width=8)
        self.dynasty_search.grid(row=0, column=5, padx=5)
        self.dynasty_search.bind('<Return>', lambda e: self.search_poems())

        self.search_btn = ttk.Button(self.search_frame, text="搜索", command=self.search_poems)
        self.search_btn.grid(row=0, column=6, padx=5)

        # 创建诗词列表的树形视图
        self.poem_tree = ttk.Treeview(self.left_frame, columns=('title', 'author', 'dynasty'), show='headings')
        self.poem_tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 配置left_frame的网格权重，使poem_tree能够自动扩展
        self.left_frame.grid_columnconfigure(0, weight=1)
        self.left_frame.grid_rowconfigure(1, weight=1)

        # 设置列标题
        self.poem_tree.heading('title', text='标题')
        self.poem_tree.heading('author', text='作者')
        self.poem_tree.heading('dynasty', text='朝代')

        # 设置列宽
        self.poem_tree.column('title', width=150)
        self.poem_tree.column('author', width=100)
        self.poem_tree.column('dynasty', width=80)

        # 绑定选择事件
        self.poem_tree.bind('<<TreeviewSelect>>', self.show_poem_details)

        # 添加滚动条
        tree_scroll = ttk.Scrollbar(self.left_frame, orient=tk.VERTICAL, command=self.poem_tree.yview)
        tree_scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.poem_tree.configure(yscrollcommand=tree_scroll.set)

        # 创建右侧内容框架
        self.content_frame = ttk.Frame(self.right_frame)
        self.content_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(1, weight=1)

        # 创建右侧顶部按钮框架
        self.right_top_frame = ttk.Frame(self.content_frame)
        self.right_top_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        self.right_top_frame.grid_columnconfigure(0, weight=1)

        # 创建一个内部框架来容纳按钮，实现居中效果
        button_container = ttk.Frame(self.right_top_frame)
        button_container.grid(row=0, column=0)

        # 添加编辑、保存、取消、关闭、导入导出按钮到按钮容器中
        self.fullscreen_btn = ttk.Button(button_container, text='全屏', command=self.show_fullscreen)
        self.fullscreen_btn.pack(side=tk.LEFT, padx=5)
        self.edit_btn = ttk.Button(button_container, text='编辑', command=self.edit_poem)
        self.edit_btn.pack(side=tk.LEFT, padx=5)
        self.save_btn = ttk.Button(button_container, text='保存', command=self.save_poem, state='disabled')
        self.save_btn.pack(side=tk.LEFT, padx=5)
        self.cancel_btn = ttk.Button(button_container, text='取消', command=self.cancel_edit, state='disabled')
        self.cancel_btn.pack(side=tk.LEFT, padx=5)
        self.import_btn = ttk.Button(button_container, text='导入', command=self.import_poems)
        self.import_btn.pack(side=tk.LEFT, padx=5)
        self.export_btn = ttk.Button(button_container, text='导出', command=self.export_poems)
        self.export_btn.pack(side=tk.LEFT, padx=5)
        self.close_btn = ttk.Button(button_container, text='退出', command=self.close_app)
        self.close_btn.pack(side=tk.LEFT, padx=5)

        # 创建诗词内容文本框（包含标题、拼音和汉字）
        self.poem_content = tk.Text(self.content_frame, wrap=tk.WORD, height=20, font=('SimSun', 14), relief=tk.SUNKEN, borderwidth=1)
        self.poem_content.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 2))
        
        # 添加滚动条
        content_scroll = ttk.Scrollbar(self.content_frame, orient=tk.VERTICAL, command=self.poem_content.yview)
        content_scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.poem_content.configure(yscrollcommand=content_scroll.set)
        
        # 配置content_frame的网格权重
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)
        
        self.poem_content.config(state='disabled')
        
        # 创建诗词详情的框架
        self.detail_frame = ttk.Frame(self.content_frame)
        self.detail_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        self.detail_frame.grid_columnconfigure(0, weight=1)
        self.detail_frame.grid_rowconfigure(0, weight=1)

        # 创建诗词详情的文本框和标签
        info_frame = ttk.Frame(self.detail_frame)
        info_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        info_frame.grid_columnconfigure(0, weight=1)
        info_frame.grid_rowconfigure(0, weight=1)

        # 创建Notebook用于显示不同类型的信息
        self.notebook = ttk.Notebook(info_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        # 创建各个标签页，使用更大的字体
        style = ttk.Style()
        style.configure('Custom.TNotebook.Tab', font=('SimSun', 14, 'bold'))
        style.configure('Borderless.TButton', font=('SimSun', 9))
        self.notebook.configure(style='Custom.TNotebook')

        translation_frame = ttk.Frame(self.notebook, padding=10)
        note_frame = ttk.Frame(self.notebook, padding=10)
        appreciation_frame = ttk.Frame(self.notebook, padding=10)
        author_frame = ttk.Frame(self.notebook, padding=10)

        # 设置每个frame的最小高度
        for frame in [translation_frame, note_frame, appreciation_frame, author_frame]:
            frame.grid_propagate(False)  # 禁止frame根据内容自动调整大小
            frame.grid_rowconfigure(0, weight=1)  # 使frame能够垂直扩展
            frame.grid_columnconfigure(0, weight=1)  # 使frame能够水平扩展
            frame.configure(height=400)

        self.notebook.add(translation_frame, text='译文')
        self.notebook.add(note_frame, text='注释')
        self.notebook.add(appreciation_frame, text='赏析')
        self.notebook.add(author_frame, text='作者介绍')

        # 创建文本框和滚动条，增加行距和间距
        # 创建文本框和滚动条，并配置grid布局
        for frame, text_attr in [
            (translation_frame, 'translation_text'),
            (note_frame, 'note_text'),
            (appreciation_frame, 'appreciation_text'),
            (author_frame, 'author_text')
        ]:
            # 配置frame的grid权重
            frame.grid_columnconfigure(0, weight=1)
            frame.grid_rowconfigure(0, weight=1)
            
            # 创建Text组件
            text_widget = tk.Text(frame, wrap=tk.WORD, height=20, 
                                 font=('SimSun', 12), spacing1=10, spacing2=5, spacing3=10)
            scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text_widget.yview)
            
            # 配置Text和Scrollbar
            text_widget.configure(yscrollcommand=scroll.set)
            text_widget.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
            scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
            
            # 保存Text组件的引用
            setattr(self, text_attr, text_widget)

        # 设置文本框初始状态为禁用
        for text_widget in [self.translation_text, self.note_text, self.appreciation_text, self.author_text]:
            text_widget.config(state='disabled')

        # 配置网格权重
        translation_frame.grid_columnconfigure(0, weight=1)
        translation_frame.grid_rowconfigure(0, weight=1)
        note_frame.grid_columnconfigure(0, weight=1)
        note_frame.grid_rowconfigure(0, weight=1)
        appreciation_frame.grid_columnconfigure(0, weight=1)
        appreciation_frame.grid_rowconfigure(0, weight=1)
        author_frame.grid_columnconfigure(0, weight=1)
        author_frame.grid_rowconfigure(0, weight=1)

    def search_poems(self):
        # 获取搜索条件
        title = self.title_search.get().strip()
        author = self.author_search.get().strip()
        dynasty = self.dynasty_search.get().strip()

        # 根据条件筛选诗词
        filtered_poems = self.poems_data
        if title:
            filtered_poems = [poem for poem in filtered_poems if title.lower() in poem['title'].lower()]
        if author:
            filtered_poems = [poem for poem in filtered_poems if author.lower() in poem['author'].lower()]
        if dynasty:
            filtered_poems = [poem for poem in filtered_poems if dynasty.lower() in poem['dynasty'].lower()]

        # 清空现有搜索结果
        for item in self.poem_tree.get_children():
            self.poem_tree.delete(item)

        # 显示搜索结果
        for poem in filtered_poems:
            self.poem_tree.insert('', 'end', values=(poem['title'], poem['author'], poem['dynasty']))

    def show_poem_details(self, event):
        # 获取选中的诗词
        selection = self.poem_tree.selection()
        if not selection:
            return
            
        item = selection[0]
        values = self.poem_tree.item(item)['values']
        title = values[0]
        
        # 查找诗词数据
        poem = None
        for p in self.poems_data:
            if p['title'] == title:
                poem = p
                break
                
        if not poem:
            return
            
        # 处理诗词内容和拼音
        self.poem_content.config(state='normal')
        self.poem_content.delete('1.0', 'end')
        
        # 配置标题、拼音和汉字的样式
        self.poem_content.tag_configure('title', font=('SimSun', 18, 'bold'), justify='center', spacing1=10, spacing3=10)
        self.poem_content.tag_configure('pinyin', font=('SimSun', 12), justify='center', spacing1=5, spacing3=5, spacing2=10)
        self.poem_content.tag_configure('hanzi', font=('SimSun', 16, 'bold'), justify='center', spacing1=5, spacing3=10, spacing2=10)
        
        # 插入标题和拼音
        if 'title_pinyin' in poem:
            self.poem_content.insert('end', poem.get('title_pinyin', '') + '\n', 'pinyin')
        self.poem_content.insert('end', poem['title'] + '\n\n', 'title')
        
        # 获取诗词内容和拼音
        content_lines = poem.get('content', [])
        pinyin_lines = poem.get('content_pinyin', [])
        
        # 将拼音显示在汉字上方
        if isinstance(content_lines, list) and isinstance(pinyin_lines, list) and len(content_lines) == len(pinyin_lines):
            for content_line, pinyin_line in zip(content_lines, pinyin_lines):
                # 插入拼音行
                self.poem_content.insert('end', pinyin_line + '\n', 'pinyin')
                # 插入汉字行
                self.poem_content.insert('end', content_line + '\n', 'hanzi')
        else:
            # 如果没有拼音或格式不匹配，只显示内容
            if isinstance(content_lines, list):
                for line in content_lines:
                    self.poem_content.insert('end', line + '\n', 'hanzi')
            else:
                self.poem_content.insert('end', str(content_lines), 'hanzi')
        
        self.poem_content.config(state='disabled')
        
        # 显示诗词内容
        for text_widget in [self.translation_text, self.note_text, self.appreciation_text, self.author_text]:
            text_widget.config(state='normal')
            text_widget.delete('1.0', 'end')
            
        self.translation_text.insert('1.0', poem.get('translation', ''))
        self.note_text.insert('1.0', poem.get('note', ''))
        self.appreciation_text.insert('1.0', poem.get('appreciation', ''))
        self.author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 如果不在编辑模式，禁用文本框
        if not self.editing:
            for text_widget in [self.translation_text, self.note_text, self.appreciation_text, self.author_text]:
                text_widget.config(state='disabled')

    def edit_poem(self):
        # 启用编辑模式
        self.editing = True
        self.edit_btn.config(state='disabled')
        self.save_btn.config(state='normal')
        self.cancel_btn.config(state='normal')
        
        # 启用所有文本框的编辑，包括诗词内容
        self.poem_content.config(state='normal')
        for text_widget in [self.translation_text, self.note_text, self.appreciation_text, self.author_text]:
            text_widget.config(state='normal')

    def save_poem(self):
        # 获取当前选中的诗词
        selection = self.poem_tree.selection()
        if not selection:
            return
            
        item = selection[0]
        values = self.poem_tree.item(item)['values']
        title = values[0]
        
        # 获取诗词内容
        content_text = self.poem_content.get('1.0', 'end-1c')
        lines = content_text.split('\n')
        
        # 解析标题、拼音和内容
        title_pinyin = ''
        new_title = ''
        content_lines = []
        content_pinyin_lines = []
        
        # 处理标题部分
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            if not line:  # 跳过空行
                i += 1
                continue
            if not title_pinyin and not new_title:  # 第一个非空行是标题拼音
                title_pinyin = line
                i += 1
            elif not new_title:  # 第二个非空行是标题
                new_title = line
                i += 2  # 跳过标题后的空行
                break
            else:
                break
        
        # 处理内容部分
        while i < len(lines):
            line = lines[i].strip()
            if line:  # 非空行
                if len(content_lines) == len(content_pinyin_lines):  # 当前是拼音行
                    content_pinyin_lines.append(line)
                else:  # 当前是汉字行
                    content_lines.append(line)
            i += 1
        
        # 更新诗词数据
        for poem in self.poems_data:
            if poem['title'] == title:
                poem['title'] = new_title
                poem['title_pinyin'] = title_pinyin
                poem['content'] = content_lines
                poem['content_pinyin'] = content_pinyin_lines
                poem['translation'] = self.translation_text.get('1.0', 'end-1c')
                poem['note'] = self.note_text.get('1.0', 'end-1c')
                poem['appreciation'] = self.appreciation_text.get('1.0', 'end-1c')
                poem['author_intro'] = self.author_text.get('1.0', 'end-1c')
                break
                
        # 保存到文件
        with open('poems.json', 'w', encoding='utf-8') as f:
            json.dump({'poems': self.poems_data}, f, ensure_ascii=False, indent=4)
            
        # 禁用编辑模式
        self.cancel_edit()
        
        # 更新树形视图中的标题
        self.poem_tree.item(item, values=(new_title, values[1], values[2]))
        
        messagebox.showinfo('成功', '保存成功！')

    def cancel_edit(self):
        # 禁用编辑模式
        self.editing = False
        self.edit_btn.config(state='normal')
        self.save_btn.config(state='disabled')
        self.cancel_btn.config(state='disabled')
        
        # 禁用所有文本框的编辑
        for text_widget in [self.translation_text, self.note_text, self.appreciation_text, self.author_text]:
            text_widget.config(state='disabled')
            
        # 重新显示当前诗词
        self.show_poem_details(None)

    def close_app(self):
        # 弹出确认对话框
        if messagebox.askyesno('确认退出', '确定要退出程序吗？'):
            self.root.quit()

    def import_poems(self):
        # 打开文件选择对话框
        file_path = filedialog.askopenfilename(
            title='选择要导入的文件',
            filetypes=[('JSON files', '*.json'), ('CSV files', '*.csv')]
        )
        
        if not file_path:
            return
            
        try:
            if file_path.lower().endswith('.json'):
                # 读取JSON文件
                with open(file_path, 'r', encoding='utf-8') as f:
                    imported_data = json.load(f)
                    
                if 'poems' not in imported_data:
                    raise ValueError('无效的JSON格式')
                    
                new_poems = imported_data['poems']
            else:  # CSV文件
                new_poems = []
                with open(file_path, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        # 处理content和content_pinyin字段，将字符串转换为列表
                        content = row.get('content', '').split('|')
                        content_pinyin = row.get('content_pinyin', '').split('|')
                        
                        poem = {
                            'title': row.get('title', ''),
                            'author': row.get('author', ''),
                            'dynasty': row.get('dynasty', ''),
                            'content': content,
                            'content_pinyin': content_pinyin,
                            'translation': row.get('translation', ''),
                            'note': row.get('note', ''),
                            'appreciation': row.get('appreciation', ''),
                            'author_intro': row.get('author_intro', ''),
                            'title_pinyin': row.get('title_pinyin', '')
                        }
                        new_poems.append(poem)
            
            # 为没有拼音的诗词添加拼音
            for poem in new_poems:
                # 检查并添加标题拼音
                if not poem.get('title_pinyin'):
                    title_py = pinyin(poem['title'], style=Style.TONE)
                    poem['title_pinyin'] = ' '.join([' '.join(p) for p in title_py])
                
                # 检查并添加内容拼音
                if not any(poem.get('content_pinyin', [])):
                    content_pinyin = []
                    for line in poem.get('content', []):
                        line_py = pinyin(line, style=Style.TONE)
                        content_pinyin.append(' '.join([' '.join(p) for p in line_py]))
                    poem['content_pinyin'] = content_pinyin
                        
            # 合并诗词数据
            existing_titles = {poem['title'] for poem in self.poems_data}
            new_poems = [poem for poem in new_poems 
                        if poem['title'] not in existing_titles]
            
            if not new_poems:
                messagebox.showinfo('提示', '没有新的诗词可以导入')
                return
                
            self.poems_data.extend(new_poems)
            
            # 保存到文件
            with open('poems.json', 'w', encoding='utf-8') as f:
                json.dump({'poems': self.poems_data}, f, ensure_ascii=False, indent=4)
                
            # 刷新显示
            self.search_poems()
            messagebox.showinfo('成功', f'成功导入{len(new_poems)}首诗词！')
            
        except Exception as e:
            messagebox.showerror('错误', f'导入失败：{str(e)}')

    def export_poems(self):
        # 打开文件保存对话框
        file_path = filedialog.asksaveasfilename(
            title='选择导出位置',
            defaultextension='.json',
            filetypes=[('JSON files', '*.json'), ('CSV files', '*.csv')]
        )
        
        if not file_path:
            return
            
        try:
            if file_path.lower().endswith('.json'):
                # 保存为JSON文件
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump({'poems': self.poems_data}, f, ensure_ascii=False, indent=4)
            else:  # CSV文件
                with open(file_path, 'w', encoding='utf-8', newline='') as f:
                    # 定义CSV文件的字段
                    fieldnames = ['title', 'author', 'dynasty', 'content', 'content_pinyin',
                                'translation', 'note', 'appreciation', 'author_intro', 'title_pinyin']
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    writer.writeheader()
                    
                    # 写入诗词数据
                    for poem in self.poems_data:
                        # 将列表转换为字符串
                        content = '|'.join(poem.get('content', []))
                        content_pinyin = '|'.join(poem.get('content_pinyin', []))
                        
                        row = {
                            'title': poem.get('title', ''),
                            'author': poem.get('author', ''),
                            'dynasty': poem.get('dynasty', ''),
                            'content': content,
                            'content_pinyin': content_pinyin,
                            'translation': poem.get('translation', ''),
                            'note': poem.get('note', ''),
                            'appreciation': poem.get('appreciation', ''),
                            'author_intro': poem.get('author_intro', ''),
                            'title_pinyin': poem.get('title_pinyin', '')
                        }
                        writer.writerow(row)
                        
            messagebox.showinfo('成功', '导出成功！')
            
        except Exception as e:
            messagebox.showerror('错误', f'导出失败：{str(e)}')

    def show_fullscreen(self):
        # 获取当前选中的诗词
        selection = self.poem_tree.selection()
        if not selection:
            return
            
        # 获取当前选中的诗词数据
        title = self.poem_tree.item(selection[0])['values'][0]
        poem = None
        for p in self.poems_data:
            if p['title'] == title:
                poem = p
                break
                
        if not poem:
            return
            
        # 创建新窗口
        fullscreen_window = tk.Toplevel(self.root)
        fullscreen_window.title('全屏显示')
        
        # 设置全屏
        fullscreen_window.attributes('-fullscreen', True)
        
        # 创建内容框架
        content_frame = ttk.Frame(fullscreen_window, padding=20)
        content_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_rowconfigure(1, weight=1)
        
        # 创建标题标签
        title_label = ttk.Label(content_frame, text=title, font=('SimSun', 36, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # 创建文本框显示内容
        text_widget = tk.Text(content_frame, wrap=tk.WORD, font=('SimSun', 24),
                             relief=tk.FLAT, bg=fullscreen_window.cget('bg'))
        text_widget.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 获取诗词内容
        text_widget.insert('1.0', self.poem_content.get('1.0', 'end-1c'))
        text_widget.config(state='disabled')
        
        # 添加退出全屏按钮
        exit_btn = ttk.Button(content_frame, text='退出全屏 (Esc)',
                            command=fullscreen_window.destroy)
        exit_btn.grid(row=2, column=0, pady=20)
        
        # 绑定Esc键退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        
        # 插入内容
        translation_text.insert('1.0', poem.get('translation', ''))
        note_text.insert('1.0', poem.get('note', ''))
        appreciation_text.insert('1.0', poem.get('appreciation', ''))
        author_text.insert('1.0', poem.get('author_intro', ''))
        
        # 禁用编辑
        for text_widget in [translation_text, note_text, appreciation_text, author_text]:
            text_widget.config(state='disabled')
            text_widget.pack(expand=True, fill=tk.BOTH)
        
        # 绑定Esc键用于退出全屏
        fullscreen_window.bind('<Escape>', lambda e: fullscreen_window.destroy())
        
        # 设置焦点
        fullscreen_window.focus_set()
        
        # 创建Notebook显示其他信息
        notebook = ttk.Notebook(content_frame)
        notebook.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        # 创建各个标签页
        translation_frame = ttk.Frame(notebook, padding=10)
        note_frame = ttk.Frame(notebook, padding=10)
        appreciation_frame = ttk.Frame(notebook, padding=10)
        author_frame = ttk.Frame(notebook, padding=10)
        
        notebook.add(translation_frame, text='译文')
        notebook.add(note_frame, text='注释')
        notebook.add(appreciation_frame, text='赏析')
        notebook.add(author_frame, text='作者介绍')
        
        # 创建文本框
        translation_text = tk.Text(translation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        note_text = tk.Text(note_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        appreciation_text = tk.Text(appreciation_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)
        author_text = tk.Text(author_frame, wrap=tk.WORD, font=('SimSun', 16), spacing1=10, spacing2=5, spacing3=10)