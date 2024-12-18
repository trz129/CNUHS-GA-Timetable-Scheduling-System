# GUI Version 1.0 by trz129©2024
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import pandas as pd
import threading
import itertools
import logging
from datetime import datetime
import collections
import time
from collections import defaultdict

class AP_Scheduler(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AP排课模块-GUI-V1.0")
        self.geometry("850x580")
        self.resizable(False, False)
        self.current_step = 0
        self.steps = [
            "初始页",
            "输入数据",
            "确认数据",
            "分配课时",
            "提交任务",
            "选择方案",
            "处理冲突",
            "生成名单",
            "保存结果"
        ]
        self.step_data = {}
        self.column_selection_vars = {}
        self.G11_student_choice = []
        self.G12_student_choice = []
        self.next_button_visible = False
        self.teacher_name_vars = {}
        self.max_students_vars = {}
        self.assign_class_vars = {}
        self.course_data_list = []
        self.student_frames = {}
        self.setup_logging()
        self.create_widgets()

    def setup_logging(self):
        self.logger = logging.getLogger("AP_Scheduler")
        self.logger.setLevel(logging.DEBUG)
        log_filename = datetime.now().strftime("AP_Scheduler_%Y%m%d_%H%M%S.log")
        fh = logging.FileHandler(log_filename, encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        ch = logging.StreamHandler()
        ch.setLevel(logging.ERROR)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)
        ch.setFormatter(formatter)
        self.logger.addHandler(fh)
        self.logger.addHandler(ch)
        self.logger.info("AP Scheduler 启动。")

    def create_widgets(self):
        self.indicator_frame = tk.Frame(self)
        self.indicator_frame.pack(pady=20)

        self.indicators = []
        self.labels = []
        for i, step in enumerate(self.steps):
            step_frame = tk.Frame(self.indicator_frame)
            step_frame.pack(side='left', padx=10)

            canvas = tk.Canvas(step_frame, width=30, height=30, highlightthickness=0)
            canvas.pack()
            oval = canvas.create_oval(5, 5, 25, 25, fill="gray")
            self.indicators.append(canvas)

            label = tk.Label(step_frame, text=step, wraplength=100, justify='center', font=("Arial", 10))
            label.pack(pady=5)
            self.labels.append(label)

        self.content_frame = tk.Frame(self, bg="white", bd=2, relief='sunken')
        self.content_frame.pack(pady=10, padx=20, fill='both', expand=True)

        self.button_frame = tk.Frame(self)
        self.button_frame.pack(pady=20)

        self.update_ui()

    def update_ui(self):
        for i, canvas in enumerate(self.indicators):
            if i < self.current_step:
                color = "green"
            elif i == self.current_step:
                color = "orange"
            else:
                color = "gray"
            canvas.itemconfig(1, fill=color)

        if self.current_step < len(self.steps):
            step_name = self.steps[self.current_step]
            for widget in self.content_frame.winfo_children():
                widget.destroy()

            if step_name == "初始页":
                self.add_initial_page()
            elif step_name == "输入数据":
                self.add_input_data()
            elif step_name == "确认数据":
                self.add_confirm_data()
            elif step_name == "分配课时":
                self.add_allocate_hours()
            elif step_name == "提交任务":
                self.add_submit_task()
            elif step_name == "选择方案":
                self.add_select_plan()
            elif step_name == "处理冲突":
                self.add_handle_conflicts()
            elif step_name == "生成名单":
                self.add_generate_list()
            elif step_name == "保存结果":
                self.add_save_results()

    def next_step(self):
        if self.current_step < len(self.steps) - 1:
            self.current_step += 1
            self.logger.info(f"进入步骤：{self.steps[self.current_step]}")
            self.update_ui()
        else:
            self.logger.info("所有步骤已完成。")

    def add_initial_page(self):
        title_label = tk.Label(self.content_frame, text="AP分配模块", font=("Arial", 24), bg="white")
        title_label.pack(pady=20)

        guide_text = ("欢迎使用AP排课模块。请按照以下步骤完成排课任务。\n\n"
                      "提示：\n\n请提前进行数据预处理，本模块暂不支持回退\n\n"
                      "当界面内所有任务完成时，程序将自动跳入下一步")
        guide_label = tk.Label(
            self.content_frame,
            text=guide_text,
            font=("Arial", 14),
            bg="white",
            wraplength=900,
            justify='center'
        )
        guide_label.pack(pady=10)

        start_button = tk.Button(
            self.content_frame,
            text="开始",
            command=self.start_initial_step,
            width=15
        )
        start_button.pack(pady=20)

        label = tk.Label(self.content_frame,
                         text="AP module GUI Version 1.0",
                         font=("Arial", 12), bg="white")
        label.pack(pady=5)
        label = tk.Label(self.content_frame,
                         text="MIT license © trz129 2024",
                         font=("Arial", 12), bg="white")
        label.pack(pady=5)

    def start_initial_step(self):
        self.logger.info("用户点击开始按钮，进入输入数据步骤。")
        self.next_step()

    def add_input_data(self):
        title_label = tk.Label(self.content_frame, text="输入数据", font=("Arial", 20), bg="white")
        title_label.pack(pady=10)

        instruction_label = tk.Label(
            self.content_frame,
            text="请分别选择G11-AP选课表格文件和G12-AP选课表格文件：",
            font=("Arial", 12),
            bg="white"
        )
        instruction_label.pack(pady=10)

        g11_frame = tk.Frame(self.content_frame, bg="white")
        g11_frame.pack(pady=5)

        g11_label = tk.Label(g11_frame, text="G11-AP选课表格文件：", font=("Arial", 12), bg="white")
        g11_label.pack(side='left')

        self.g11_file_var = tk.StringVar()
        g11_entry = tk.Entry(g11_frame, textvariable=self.g11_file_var, width=50)
        g11_entry.pack(side='left', padx=5)

        g11_button = tk.Button(g11_frame, text="浏览", command=self.select_g11_file)
        g11_button.pack(side='left')

        g12_frame = tk.Frame(self.content_frame, bg="white")
        g12_frame.pack(pady=5)

        g12_label = tk.Label(g12_frame, text="G12-AP选课表格文件：", font=("Arial", 12), bg="white")
        g12_label.pack(side='left')

        self.g12_file_var = tk.StringVar()
        g12_entry = tk.Entry(g12_frame, textvariable=self.g12_file_var, width=50)
        g12_entry.pack(side='left', padx=5)

        g12_button = tk.Button(g12_frame, text="浏览", command=self.select_g12_file)
        g12_button.pack(side='left')

    def select_g11_file(self):
        filename = filedialog.askopenfilename(
            title="选择G11-AP选课表格文件",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if filename:
            self.g11_file_var.set(filename)
            self.step_data["G11_file"] = filename
            self.logger.info(f"G11文件选择：{filename}")
            self.check_files_selected()

    def select_g12_file(self):
        filename = filedialog.askopenfilename(
            title="选择G12-AP选课表格文件",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if filename:
            self.g12_file_var.set(filename)
            self.step_data["G12_file"] = filename
            self.logger.info(f"G12文件选择：{filename}")
            self.check_files_selected()

    def check_files_selected(self):
        if self.step_data.get("G11_file") and self.step_data.get("G12_file"):
            self.logger.info("G11和G12文件均已选择，进入确认数据步骤。")
            self.next_step()

    def add_confirm_data(self):
        g11 = self.step_data.get("G11_file", "")
        g12 = self.step_data.get("G12_file", "")

        if not g11 or not g12:
            messagebox.showerror("错误", "未选择G11或G12文件。")
            self.logger.error("未选择G11或G12文件。")
            return

        confirm_label = tk.Label(
            self.content_frame,
            text="请确认您选择的文件，并指定AP选课列、姓名列及AP备选列",
            font=("Arial", 14),
            bg="white",
            wraplength=900,
            justify='left'
        )
        confirm_label.pack(pady=10)

        tip_label = tk.Label(
            self.content_frame,
            text="不要忘了填G12",
            font=("Arial", 12),
            bg="white",
            fg="red",
            justify='left'
        )
        tip_label.pack(pady=5, anchor='w')

        self.notebook = ttk.Notebook(self.content_frame)
        self.notebook.pack(expand=True, fill='both')

        self.frame_g11 = tk.Frame(self.notebook)
        self.frame_g12 = tk.Frame(self.notebook)
        self.notebook.add(self.frame_g11, text='G11 AP选课')
        self.notebook.add(self.frame_g12, text='G12 AP选课')

        try:
            df_g11 = pd.read_excel(g11)
            self.step_data['df_g11'] = df_g11
            df_g11_preview = df_g11.head(8)
            self.display_dataframe(self.frame_g11, df_g11_preview)
            self.add_column_selection(self.frame_g11, df_g11, "G11")
            self.logger.info("G11文件已读取并显示。")
        except Exception as e:
            error_message = f"无法读取G11文件：{e}"
            tk.Label(self.frame_g11, text=error_message, fg="red").pack()
            self.logger.error(error_message)

        try:
            df_g12 = pd.read_excel(g12)
            self.step_data['df_g12'] = df_g12
            df_g12_preview = df_g12.head(8)
            self.display_dataframe(self.frame_g12, df_g12_preview)
            self.add_column_selection(self.frame_g12, df_g12, "G12")
            self.logger.info("G12文件已读取并显示。")
        except Exception as e:
            error_message = f"无法读取G12文件：{e}"
            tk.Label(self.frame_g12, text=error_message, fg="red").pack()
            self.logger.error(error_message)

    def display_dataframe(self, parent, dataframe):
        tree = ttk.Treeview(parent, columns=list(dataframe.columns), show='headings', height=8)
        tree.pack(expand=True, fill='both')

        scrollbar_y = ttk.Scrollbar(parent, orient='vertical', command=tree.yview)
        scrollbar_y.pack(side='right', fill='y')
        tree.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(parent, orient='horizontal', command=tree.xview)
        scrollbar_x.pack(side='bottom', fill='x')
        tree.configure(xscrollcommand=scrollbar_x.set)

        for col in dataframe.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor='center')

        for _, row in dataframe.iterrows():
            tree.insert('', 'end', values=list(row))

    def add_column_selection(self, parent, dataframe, table_key):
        selection_frame = tk.Frame(parent, bg="white")
        selection_frame.pack(pady=10, fill='x')

        ap_select_label = tk.Label(
            selection_frame,
            text=f"{table_key} - AP选课列 (多选-不包括微积分):",
            font=("Arial", 12),
            bg="white"
        )
        ap_select_label.pack(anchor='w', padx=5, pady=(10, 5))

        ap_select_frame = tk.Frame(selection_frame, bg="white")
        ap_select_frame.pack(anchor='w', padx=5, pady=5)

        ap_select_vars = []
        for col in dataframe.columns:
            var = tk.IntVar()
            cb = tk.Checkbutton(
                ap_select_frame,
                text=col,
                variable=var,
                bg="white"
            )
            cb.pack(side='left', padx=5, pady=2)
            ap_select_vars.append((col, var))
            var.trace_add("write", self.check_column_selection)

        name_backup_frame = tk.Frame(selection_frame, bg="white")
        name_backup_frame.pack(anchor='w', padx=5, pady=5, fill='x')

        name_column_label = tk.Label(
            name_backup_frame,
            text=f"{table_key} - 中文姓名列:",
            font=("Arial", 12),
            bg="white"
        )
        name_column_label.pack(side='left', padx=(0, 10))

        name_column_var = tk.StringVar()
        name_column_combobox = ttk.Combobox(
            name_backup_frame,
            textvariable=name_column_var,
            values=list(dataframe.columns),
            state='readonly'
        )
        name_column_combobox.pack(side='left', padx=(0, 20))
        name_column_combobox.bind("<<ComboboxSelected>>", self.check_column_selection)

        ap_backup_label = tk.Label(
            name_backup_frame,
            text=f"{table_key} - AP备选列:",
            font=("Arial", 12),
            bg="white"
        )
        ap_backup_label.pack(side='left', padx=(0, 10))

        ap_backup_var = tk.StringVar()
        ap_backup_combobox = ttk.Combobox(
            name_backup_frame,
            textvariable=ap_backup_var,
            values=list(dataframe.columns),
            state='readonly'
        )
        ap_backup_combobox.pack(side='left')
        ap_backup_combobox.bind("<<ComboboxSelected>>", self.check_column_selection)

        self.column_selection_vars[table_key] = {
            "AP_select": ap_select_vars,
            "AP_backup": ap_backup_var,
            "Name_column": name_column_var
        }

    def check_column_selection(self, *args):
        all_selected = True
        selected_columns = {}
        for table_key, selections in self.column_selection_vars.items():
            ap_select = [col for col, var in selections["AP_select"] if var.get() == 1]
            ap_backup = selections["AP_backup"].get()
            name_column = selections["Name_column"].get()

            if not ap_select or not ap_backup or not name_column:
                all_selected = False
                break

            if ap_backup in ap_select:
                all_selected = False
                break
            if name_column in ap_select or name_column == ap_backup:
                all_selected = False
                break

            selected_columns.setdefault(table_key, {})
            selected_columns[table_key]["AP_select"] = ap_select
            selected_columns[table_key]["AP_backup"] = ap_backup
            selected_columns[table_key]["Name_column"] = name_column

        if all_selected:
            self.step_data["column_selection"] = selected_columns
            self.logger.info("所有列选择已完成，开始统计课程选择。")

            course_counts = {}
            for table_key in ['G11', 'G12']:
                df = self.step_data.get(f'df_{table_key.lower()}')
                if df is None:
                    messagebox.showerror("Error", f"无法获取 {table_key} 的数据。")
                    self.logger.error(f"无法获取 {table_key} 的数据。")
                    return
                selections = selected_columns[table_key]
                ap_select_columns = selections['AP_select']

                for col in ap_select_columns:
                    course_list = df[col].dropna().tolist()
                    for course in course_list:
                        course = str(course).strip()
                        if course not in course_counts:
                            course_counts[course] = {'G11': 0, 'G12': 0}
                        course_counts[course][table_key] += 1

            self.step_data['course_counts'] = course_counts
            self.logger.debug(f"课程选择计数：{course_counts}")

            df_g11 = self.step_data.get('df_g11')
            selected_columns_g11 = selected_columns['G11']
            ap_select_columns_g11 = selected_columns_g11['AP_select']
            name_column_g11 = selected_columns_g11['Name_column']
            ap_backup_column_g11 = selected_columns_g11['AP_backup']

            for index, row in df_g11.iterrows():
                name = row[name_column_g11]
                courses = []
                for col in ap_select_columns_g11:
                    val = row[col]
                    if pd.notnull(val):
                        courses.extend([course.strip() for course in str(val).split(',') if course.strip()])
                backup_course = row[ap_backup_column_g11]
                if pd.notnull(backup_course):
                    backup_course = str(backup_course).strip()
                else:
                    backup_course = None
                self.G11_student_choice.append((name, courses, backup_course))

            df_g12 = self.step_data.get('df_g12')
            selected_columns_g12 = selected_columns['G12']
            ap_select_columns_g12 = selected_columns_g12['AP_select']
            name_column_g12 = selected_columns_g12['Name_column']
            ap_backup_column_g12 = selected_columns_g12['AP_backup']

            for index, row in df_g12.iterrows():
                name = row[name_column_g12]
                courses = []
                for col in ap_select_columns_g12:
                    val = row[col]
                    if pd.notnull(val):
                        courses.extend([course.strip() for course in str(val).split(',') if course.strip()])
                backup_course = row[ap_backup_column_g12]
                if pd.notnull(backup_course):
                    backup_course = str(backup_course).strip()
                else:
                    backup_course = None
                self.G12_student_choice.append((name, courses, backup_course))

            self.logger.debug(f"G11学生课程选择：{self.G11_student_choice}")
            self.logger.debug(f"G12学生课程选择：{self.G12_student_choice}")

            self.remove_duplicate_courses()
            self.next_step()

    def remove_duplicate_courses(self):

        def deduplicate_courses(student_choice):
            name, courses, backup_course = student_choice
            seen = set()
            deduplicated_courses = []
            for course in courses:
                if course not in seen:
                    deduplicated_courses.append(course)
                    seen.add(course)
                else:
                    self.logger.debug(f"学生 {name} 选择的课程 '{course}' 重复，已删除后续选择。")
            return (name, deduplicated_courses, backup_course)

        deduped_g11 = []
        for student in self.G11_student_choice:
            deduped_student = deduplicate_courses(student)
            deduped_g11.append(deduped_student)
        self.G11_student_choice = deduped_g11

        deduped_g12 = []
        for student in self.G12_student_choice:
            deduped_student = deduplicate_courses(student)
            deduped_g12.append(deduped_student)
        self.G12_student_choice = deduped_g12

        self.logger.info("所有学生的课程选择已去重。")

    def add_allocate_hours(self):
        left_frame = tk.Frame(self.content_frame, bg="white")
        left_frame.place(relx=0, rely=0, relwidth=0.75, relheight=1)

        right_frame = tk.Frame(self.content_frame, bg="white")
        right_frame.place(relx=0.75, rely=0, relwidth=0.25, relheight=1)

        self.canvas = tk.Canvas(left_frame, bg="white")
        self.canvas.pack(side='left', fill='both', expand=True)

        scrollbar = ttk.Scrollbar(left_frame, orient='vertical', command=self.canvas.yview)
        scrollbar.pack(side='right', fill='y')
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.frame_allocate = tk.Frame(self.canvas, bg="white")
        self.canvas.create_window((0, 0), window=self.frame_allocate, anchor='nw')

        self.course_data_list = []
        course_counts = self.step_data.get('course_counts', {})
        self.logger.debug(f"分配课时步骤中的课程计数：{course_counts}")

        self.prompt_label = tk.Label(
            self.frame_allocate,
            text="请在下方输入每门课程的教师名称和教室大小(计算机~45)：",
            bg="white",
            fg="red",
            font=('Arial', 12)
        )
        self.prompt_label.grid(row=0, column=0, columnspan=6, sticky='w', padx=1, pady=(10, 10))

        headers = ["课程名称", "教师名称", "教室大小", "G11总人数", "G12总人数", "分班数"]
        for col, header in enumerate(headers):
            label = tk.Label(
                self.frame_allocate,
                text=header,
                bg="lightgray",
                font=('Arial', 12, 'bold'),
                borderwidth=1,
                relief='solid'
            )
            label.grid(row=1, column=col, sticky='nsew', padx=1, pady=1)

        self.teacher_name_vars = {}
        self.max_students_vars = {}
        self.assign_class_vars = {}

        for row_num, (course_name, counts) in enumerate(course_counts.items(), start=2):
            g11_count = counts.get('G11', 0)
            g12_count = counts.get('G12', 0)

            teacher_var = tk.StringVar()
            self.teacher_name_vars[course_name] = teacher_var
            max_students_var = tk.StringVar(value='30')
            self.max_students_vars[course_name] = max_students_var
            assign_class_var = tk.StringVar(value='1')
            self.assign_class_vars[course_name] = assign_class_var

            self.course_data_list.append({
                'course_name': course_name,
                'g11_count': g11_count,
                'g12_count': g12_count,
                'teacher_var': teacher_var,
                'max_students_var': max_students_var,
                'assign_class_var': assign_class_var,
                'teacher_entry': None,
                'max_students_entry': None,
                'assign_class_entry': None
            })

        self.render_allocate_hours_table()
        self.frame_allocate.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        self.canvas.bind("<MouseWheel>", lambda event: self._on_mousewheel(event, self.canvas))

        instruction_label = tk.Label(
            right_frame,
            text=("输入教师名称和教室人数后，点击完成分配\n\n"
                  "确认各个课程的分班数后，程序将自动跳入下一步"),
            bg="white",
            font=("Arial", 12),
            wraplength=200,
            justify='center'
        )
        instruction_label.pack(pady=20)

        self.allocate_button = tk.Button(
            right_frame,
            text="完成分配",
            command=self.allocate_hours_complete,
            width=15
        )
        self.allocate_button.pack(pady=20)

    def _on_mousewheel(self, event, canvas):
        try:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception as e:
            self.logger.error(f"鼠标滚轮事件处理错误：{e}")

    def render_allocate_hours_table(self):
        for widget in self.frame_allocate.winfo_children():
            widget.destroy()

        self.prompt_label = tk.Label(
            self.frame_allocate,
            text="请在下方输入每门课程的教师名称和教室大小(计算机~45)：",
            bg="white",
            fg="red",
            font=('Arial', 12)
        )
        self.prompt_label.grid(row=0, column=0, columnspan=6, sticky='w', padx=1, pady=(10, 10))

        headers = ["课程名称", "教师名称", "教室大小", "G11总人数", "G12总人数", "分班数"]
        for col, header in enumerate(headers):
            label = tk.Label(
                self.frame_allocate,
                text=header,
                bg="lightgray",
                font=('Arial', 12, 'bold'),
                borderwidth=1,
                relief='solid'
            )
            label.grid(row=1, column=col, sticky='nsew', padx=1, pady=1)

        for row_num, data in enumerate(self.course_data_list, start=2):
            course_name = data['course_name']
            teacher_var = data['teacher_var']
            max_students_var = data['max_students_var']
            assign_class_var = data['assign_class_var']
            g11_count = data['g11_count']
            g12_count = data['g12_count']

            label_course = tk.Label(
                self.frame_allocate,
                text=course_name,
                bg="white",
                font=("Arial", 12),
                borderwidth=1,
                relief='solid',
                wraplength=180,
                justify='center'
            )
            label_course.grid(row=row_num, column=0, sticky='nsew', padx=1, pady=1)

            entry_teacher = tk.Entry(
                self.frame_allocate,
                textvariable=teacher_var,
                width=15,
                justify='center'
            )
            entry_teacher.grid(row=row_num, column=1, sticky='nsew', padx=1, pady=1)
            data['teacher_entry'] = entry_teacher

            entry_max_students = tk.Entry(
                self.frame_allocate,
                textvariable=max_students_var,
                width=10,
                justify='center'
            )
            entry_max_students.grid(row=row_num, column=2, sticky='nsew', padx=1, pady=1)
            data['max_students_entry'] = entry_max_students

            label_g11 = tk.Label(
                self.frame_allocate,
                text=str(g11_count),
                bg="white",
                font=("Arial", 12),
                borderwidth=1,
                relief='solid'
            )
            label_g11.grid(row=row_num, column=3, sticky='nsew', padx=1, pady=1)

            label_g12 = tk.Label(
                self.frame_allocate,
                text=str(g12_count),
                bg="white",
                font=("Arial", 12),
                borderwidth=1,
                relief='solid'
            )
            label_g12.grid(row=row_num, column=4, sticky='nsew', padx=1, pady=1)

            entry_assign_class = tk.Entry(
                self.frame_allocate,
                textvariable=assign_class_var,
                width=5,
                justify='center',
                state='disabled'
            )
            entry_assign_class.grid(row=row_num, column=5, sticky='nsew', padx=1, pady=1)
            data['assign_class_entry'] = entry_assign_class

        for col in range(len(headers)):
            self.frame_allocate.grid_columnconfigure(col, weight=1)

    def allocate_hours_complete(self):
        allocations = {}
        for data in self.course_data_list:
            course = data['course_name']
            teacher = data['teacher_var'].get().strip()
            max_students_str = data['max_students_var'].get().strip()

            if not teacher:
                messagebox.showerror("错误", f"课程 '{course}' 未分配教师。")
                self.logger.error(f"课程 '{course}' 未分配教师。")
                return
            if not max_students_str.isdigit():
                messagebox.showerror("错误", f"课程 '{course}' 的教室最多人数必须是数字。")
                self.logger.error(f"课程 '{course}' 的教室最多人数必须是数字。")
                return

            max_students = int(max_students_str)
            data['max_students'] = max_students

            allocations[course] = {
                'teacher': teacher,
                'max_students': max_students
            }

        self.step_data['allocations'] = allocations
        self.logger.debug(f"教师分配和教室最大人数：{allocations}")

        for data in self.course_data_list:
            total_students = data['g11_count'] + data['g12_count']
            data['sort_key'] = total_students / data['max_students']

        self.course_data_list.sort(key=lambda x: x['sort_key'], reverse=True)

        for idx, data in enumerate(self.course_data_list):
            if idx < 5:
                data['assign_class_var'].set('2')
            else:
                data['assign_class_var'].set('1')

        for data in self.course_data_list:
            data['teacher_entry'].config(state='disabled')
            data['max_students_entry'].config(state='disabled')

        for data in self.course_data_list:
            data['assign_class_entry'].config(state='normal')

        if hasattr(self, 'prompt_label'):
            self.prompt_label.config(text="请在下方确认或修改分班数：")

        self.allocate_button.config(text="确认分班数", command=self.confirm_assign_class)

    def confirm_assign_class(self):
        for data in self.course_data_list:
            course = data['course_name']
            assign_class_str = data['assign_class_var'].get().strip()

            if not assign_class_str.isdigit():
                messagebox.showerror("错误", f"课程 '{course}' 的分班数必须是数字。")
                self.logger.error(f"课程 '{course}' 的分班数必须是数字。")
                return

            assign_class = int(assign_class_str)
            if assign_class < 1:
                messagebox.showerror("错误", f"课程 '{course}' 的分班数必须至少为1。")
                self.logger.error(f"课程 '{course}' 的分班数必须至少为1。")
                return

            data['assign_class'] = assign_class

        self.step_data['assign_class'] = {
            data['course_name']: data['assign_class']
            for data in self.course_data_list
        }

        for data in self.course_data_list:
            total_students = data['g11_count'] + data['g12_count']
            data['sort_key'] = total_students / data['assign_class']

        self.course_data_list.sort(key=lambda x: x['sort_key'], reverse=True)

        self.render_allocate_hours_table()

        self.next_step()

    def add_submit_task(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        prompt_label = tk.Label(
            self.content_frame,
            text="请确认或修改AP课表结构，并点击提交。",
            font=("Arial", 14),
            bg="white",
            wraplength=900,
            justify='left'
        )
        prompt_label.pack(pady=5)

        container_frame = tk.Frame(self.content_frame, bg="gray")
        container_frame.pack(fill='both', expand=True, padx=10, pady=10)

        internal_frame = tk.Frame(container_frame, bg="white")
        internal_frame.pack(fill='both', expand=True)

        left_frame = tk.Frame(internal_frame, bg="white", bd=1, relief='solid')
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 5), pady=0)

        right_frame = tk.Frame(internal_frame, bg="white", bd=1, relief='solid')
        right_frame.pack(side='right', fill='both', expand=True, padx=(5, 0), pady=0)
        assign_class_dict = self.step_data.get('assign_class', {})
        total_assigned_classes = sum(assign_class_dict.values())
        self.checkbox_vars = []
        self.total_checked_boxes = tk.IntVar(value=0)
        total_assigned_label = tk.Label(
            right_frame, text=f"分班数之和：{total_assigned_classes}", font=("微软雅黑", 12), bg="white"
        )
        total_assigned_label.pack(pady=10)

        self.checked_boxes_label = tk.Label(
            right_frame, text=f"打勾的框数：{self.total_checked_boxes.get()}", font=("微软雅黑", 12), bg="white"
        )
        self.checked_boxes_label.pack(pady=10)
        self.complete_button = tk.Button(
            right_frame,
            text="完成提交",
            command=self.submit_task_complete,
            width=15,
            font=("微软雅黑", 12),
            bg="#2196F3",
            fg="white",
            activebackground="#0b7dda",
            relief='flat',
            state='disabled'
        )
        self.complete_button.pack(pady=20)
        instruction_label = tk.Label(
            right_frame,
            text=("请在左侧确认AP课表结构：\n\n"
                  "- 表格的行表示课程时段\n"
                  "- 表格的列表示同时上课的课程\n"
                  "- 当打勾数等于分班数之和时，才能继续\n"
                  "- 未懂此页直接点击提交即可\n"),
            bg="white",
            font=("微软雅黑", 8),
            wraplength=250,
            justify='left'
        )
        instruction_label.pack(pady=20)
        table_frame = tk.Frame(left_frame, bg="white")
        table_frame.pack(pady=10, padx=10, fill='both', expand=True)
        table_title = tk.Label(
            table_frame,
            text="AP课表结构",
            font=("Arial", 14, "bold"),
            bg="white"
        )
        table_title.grid(row=0, column=0, columnspan=6, pady=(0, 10))
        headers = ["时段", "课程A", "课程B", "课程C", "课程D", "可选年级"]
        for j, header in enumerate(headers):
            label = tk.Label(
                table_frame,
                text=header,
                font=("微软雅黑", 12, "bold"),
                bg="#f2f2f2",
                borderwidth=1,
                relief='groove',
                width=0
            )
            label.grid(row=1, column=j, padx=0, pady=0, sticky='nsew')
        time_periods = ['时段A', '时段B', '时段C', '时段D', '时段E']
        for i in range(2, 7):
            time_label = tk.Label(
                table_frame,
                text=time_periods[i - 2],
                font=("微软雅黑", 10),
                bg="white",
                borderwidth=1,
                relief='groove',
                width=0,
                wraplength=100,
                justify='left'
            )
            time_label.grid(row=i, column=0, padx=0, pady=0, sticky='nsew')
            row_vars = []
            for j in range(1, 5):
                var = tk.IntVar()
                checkbox = tk.Checkbutton(
                    table_frame,
                    variable=var,
                    command=self.update_checked_boxes,
                    bg="white",
                    activebackground="#e6e6e6",
                    relief='flat'
                )
                checkbox.grid(row=i, column=j, padx=0, pady=0, sticky='nsew')
                row_vars.append(var)
            self.checkbox_vars.append(row_vars)
            if i <= 4:
                grade_text = "G11/G12"
                grade_bg = "#ADD8E6"
            else:
                grade_text = "G12"
                grade_bg = "#90EE90"
            grade_label = tk.Label(
                table_frame,
                text=grade_text,
                font=("微软雅黑", 10),
                bg=grade_bg,
                borderwidth=1,
                relief='groove',
                width=0,
                wraplength=100,
                justify='left'
            )
            grade_label.grid(row=i, column=5, padx=0, pady=0, sticky='nsew')
        for j in range(6):
            table_frame.grid_columnconfigure(j, weight=1)
        total_checked_boxes = 0
        order = []
        for i in range(3):
            for j in range(4):
                order.append((i, j))
        for j in range(4):
            for i in range(3, 5):
                order.append((i, j))
        total_assigned_classes = sum(self.step_data.get('assign_class', {}).values())
        for box in order:
            if total_checked_boxes < total_assigned_classes:
                var = self.checkbox_vars[box[0]][box[1]]
                var.set(1)
                total_checked_boxes += 1
            else:
                var = self.checkbox_vars[box[0]][box[1]]
                var.set(0)
        self.update_checked_boxes()
        self.check_proceed_condition()

    def update_checked_boxes(self):
        total_checked = sum(var.get() for row in self.checkbox_vars for var in row)
        self.total_checked_boxes.set(total_checked)
        if hasattr(self, 'checked_boxes_label'):
            self.checked_boxes_label.config(text=f"打勾的框数：{self.total_checked_boxes.get()}")
        self.check_proceed_condition()

    def check_proceed_condition(self):
        total_assigned_classes = sum(self.step_data.get('assign_class', {}).values())
        total_checked = self.total_checked_boxes.get()
        if total_checked == total_assigned_classes and total_checked > 0:
            self.complete_button.config(state='normal')
        else:
            self.complete_button.config(state='disabled')

    def submit_task_complete(self):
        self.logger.info("用户完成提交任务步骤。")
        self.next_step()

    def add_select_plan(self):
        try:
            course_assign_classes = {data['course_name']: data['assign_class'] for data in self.course_data_list}
            course_teachers = {data['course_name']: data['teacher_var'].get() for data in self.course_data_list}
            time_slots = []
            for i, row in enumerate(self.checkbox_vars):
                for j, var in enumerate(row):
                    if var.get() == 1:
                        time_slots.append((i, j))
            class_instances = []
            for course, num_classes in course_assign_classes.items():
                for n in range(1, num_classes + 1):
                    class_instances.append((course, n))
            time_structure = [0] * 5
            for slot in time_slots:
                time_structure[slot[0]] += 1
            for widget in self.content_frame.winfo_children():
                widget.destroy()
            label = tk.Label(self.content_frame, text="正在生成所有可能的排课方案，请稍候...", font=("Arial", 14))
            label.pack(pady=10)
            progress_frame = tk.Frame(self.content_frame)
            progress_frame.pack(pady=10)
            self.progress = ttk.Progressbar(progress_frame, orient='horizontal', mode='indeterminate', length=400)
            self.progress.pack(side='left', padx=5)
            self.progress.start(10)
            self.attempt_label = tk.Label(progress_frame, text="已尝试方案数：0", font=("Arial", 12))
            self.attempt_label.pack(side='left', padx=5)
            threading.Thread(target=self.calculate_schedules, args=(time_structure, class_instances, course_teachers), daemon=True).start()
        except AssertionError as ae:
            self.logger.error(f"断言错误: {ae}")
            messagebox.showerror("断言错误", str(ae))
        except Exception as e:
            self.logger.error(f"calc 函数中出现错误: {e}")
            messagebox.showerror("错误", f"调试过程中出现错误: {e}")

    def calculate_schedules(self, time_structure, class_instances, course_teachers):
        start_time = time.time()
        all_schedules = self.generate_all_possible_schedules(time_structure, class_instances, course_teachers)
        end_time = time.time()
        self.logger.info(f"生成所有可能的排课方案耗时：{end_time - start_time:.2f}秒")
        if not all_schedules:
            self.after(0, self.no_schedule_found)
            return
        self.progress.stop()
        self.after(0, self.calculate_conflicts, all_schedules)

    def generate_all_possible_schedules(self, time_structure, class_instances, course_teachers):
        course_counts = collections.Counter(course_name for course_name, _ in class_instances)
        total_slots = sum(time_structure)
        if total_slots != sum(course_counts.values()):
            self.logger.error("课程数量与时间结构所提供的总时段数不匹配。")
            return []
        all_schedules = []
        self.attempt_counter = 0
        def recursive_schedule(time_period_index, remaining_course_counts, current_schedule):
            if time_period_index == len(time_structure):
                if sum(remaining_course_counts.values()) == 0:
                    schedule_tuple = tuple(set(time_period) for time_period in current_schedule)
                    all_schedules.append(schedule_tuple)
                return
            s = time_structure[time_period_index]
            available_courses = []
            for course, count in remaining_course_counts.items():
                if count > 0:
                    available_courses.extend([course] * count)
            if len(available_courses) < s:
                return
            possible_combinations = set(itertools.combinations(sorted(set(available_courses)), s))
            for combination in possible_combinations:
                teachers_in_combination = set()
                conflict = False
                for course in combination:
                    teacher = course_teachers[course]
                    if teacher in teachers_in_combination:
                        conflict = True
                        break
                    teachers_in_combination.add(teacher)
                if not conflict:
                    new_remaining_course_counts = remaining_course_counts.copy()
                    for c in combination:
                        new_remaining_course_counts[c] -= 1
                    current_schedule.append(combination)
                    self.attempt_counter += 1
                    if self.attempt_counter % 100 == 0:
                        self.after(0, self.update_attempt_label, self.attempt_counter)
                    recursive_schedule(time_period_index + 1, new_remaining_course_counts, current_schedule)
                    current_schedule.pop()
            return
        recursive_schedule(0, course_counts, [])
        return all_schedules

    def update_attempt_label(self, count):
        self.attempt_label.config(text=f"已尝试方案数：{count}")

    def calculate_conflicts(self, all_schedules):
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        label = tk.Label(self.content_frame, text="正在计算每个方案的冲突情况，请稍候...", font=("Arial", 14))
        label.pack(pady=10)
        self.progress = ttk.Progressbar(self.content_frame, orient='horizontal', mode='determinate', length=400)
        self.progress.pack(pady=20)
        self.percent_label = tk.Label(self.content_frame, text="0%", font=("Arial", 12))
        self.percent_label.pack(pady=5)
        threading.Thread(target=self.compute_conflicts_for_schedules, args=(all_schedules,), daemon=True).start()

    def compute_conflicts_for_schedules(self, all_schedules):
        conflict_sets = {}
        unassigned_students = []
        start_time = time.time()
        total_schedules = len(all_schedules)
        for idx, schedule in enumerate(all_schedules):
            conflict_students = self.varify_AP_combination(
                schedule, self.G11_student_choice, self.G12_student_choice
            )
            conflict_students_set = frozenset(conflict_students)
            if conflict_students_set not in conflict_sets:
                conflict_sets[conflict_students_set] = {
                    'schedule': schedule,
                    'num_conflicts': len(conflict_students),
                    'conflict_students': conflict_students
                }
            if idx % 10 == 0:
                elapsed_time = time.time() - start_time
                remaining_time = (elapsed_time / (idx + 1)) * (total_schedules - idx - 1)
                percentage = ((idx + 1) / total_schedules) * 100
                self.after(0, self.update_progress_bar, idx + 1, remaining_time, percentage)
        schedule_results = list(conflict_sets.values())
        schedule_results.sort(key=lambda x: x['num_conflicts'])
        top_results = schedule_results[:14]
        if top_results:
            min_conflicts = top_results[0]['num_conflicts']
            for result in top_results:
                if result['num_conflicts'] == min_conflicts:
                    unassigned_students.extend(result['conflict_students'])
                else:
                    break
        self.step_data['unassigned_students'] = list(set(unassigned_students))
        self.step_data['schedule_results'] = top_results
        self.progress['value'] = 100
        self.percent_label.config(text=f"100%")
        self.after(0, self.display_schedule_results, top_results)

    def update_progress_bar(self, value, remaining_time, percentage):
        self.progress['value'] = percentage
        time_label_text = f"预计剩余时间：{int(remaining_time)} 秒"
        if hasattr(self, 'time_label'):
            self.time_label.config(text=time_label_text)
        else:
            self.time_label = tk.Label(self.content_frame, text=time_label_text, font=("Arial", 12))
            self.time_label.pack(pady=5)
        self.percent_label.config(text=f"{percentage:.2f}%")

    def no_schedule_found(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        label = tk.Label(self.content_frame, text="未找到任何可行的排课方案。", font=("Arial", 14), fg="red")
        label.pack(pady=20)
        back_button = tk.Button(
            self.content_frame,
            text="返回",
            command=self.next_step,
            width=15
        )
        back_button.pack(pady=10)

    def varify_AP_combination(self, schedule, G11_student_choice, G12_student_choice):
        conflict_students = []
        from itertools import permutations, combinations
        for student in G11_student_choice:
            student_name = student[0]
            courses = student[1]
            allowed_time_slots = range(3)
            time_periods = [schedule[i] for i in allowed_time_slots]
            found = False
            if not courses:
                continue
            for assignment in permutations(allowed_time_slots, len(courses)):
                conflict = False
                used_time_slots = set()
                for i, course in enumerate(courses):
                    time_slot = assignment[i]
                    if course in time_periods[time_slot] and time_slot not in used_time_slots:
                        used_time_slots.add(time_slot)
                    else:
                        conflict = True
                        break
                if not conflict:
                    found = True
                    break
            if not found:
                conflict_students.append(student_name)
        for student in G12_student_choice:
            student_name = student[0]
            courses = student[1]
            allowed_time_slots = range(len(schedule))
            time_periods = [schedule[i] for i in allowed_time_slots]
            found = False
            if not courses:
                continue
            for assignment in permutations(allowed_time_slots, len(courses)):
                conflict = False
                used_time_slots = set()
                for i, course in enumerate(courses):
                    time_slot = assignment[i]
                    if course in time_periods[time_slot] and time_slot not in used_time_slots:
                        used_time_slots.add(time_slot)
                    else:
                        conflict = True
                        break
                if not conflict:
                    found = True
                    break
            if not found:
                conflict_students.append(student_name)
        return conflict_students

    def display_schedule_results(self, schedule_results):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        label = tk.Label(self.content_frame, text="请选择一个排课方案：", font=("Arial", 14))
        label.pack(pady=10)

        columns = ('index', 'num_conflicts', 'conflict_students')
        tree = ttk.Treeview(self.content_frame, columns=columns, show='headings', height=10)
        tree.heading('index', text='方案编号')
        tree.heading('num_conflicts', text='冲突人数')
        tree.heading('conflict_students', text='冲突学生')

        tree.column('index', width=75, anchor='center')
        tree.column('num_conflicts', width=75, anchor='center')
        tree.column('conflict_students', width=650, anchor='w')

        for idx, result in enumerate(schedule_results):
            conflict_names = ', '.join(result['conflict_students'])
            tree.insert('', 'end', values=(idx + 1, result['num_conflicts'], conflict_names))

        tree.pack(fill='both', expand=True)
        tree.bind("<Double-1>", lambda event: self.view_schedule_details(tree, schedule_results))

        confirm_button = tk.Button(self.content_frame, text="确认选择", command=lambda: self.confirm_selected_schedule(tree, schedule_results))
        confirm_button.pack(pady=10)

    def confirm_selected_schedule(self, tree, schedule_results):
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("提示", "请先选择一个方案。")
            return
        item = tree.item(selected_item)
        index = int(item['values'][0]) - 1
        result = schedule_results[index]
        self.select_schedule(result)

    def view_schedule_details(self, tree, schedule_results):
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("提示", "请先选择一个方案。")
            return
        item = tree.item(selected_item)
        index = int(item['values'][0]) - 1
        result = schedule_results[index]

        schedule = result['schedule']
        conflict_students = result['conflict_students']

        detail_window = tk.Toplevel(self)
        detail_window.title(f"方案 {index + 1} 详情")
        detail_window.transient(self)
        detail_window.grab_set()
        detail_window.attributes('-topmost', True)

        schedule_label = tk.Label(detail_window, text="排课方案：", font=("Arial", 12))
        schedule_label.pack(pady=5)

        for time_period_index, time_period in enumerate(schedule):
            classes_in_time_period = ', '.join(time_period)
            time_label = tk.Label(
                detail_window,
                text=f"时间段 {time_period_index + 1}: {classes_in_time_period}",
                font=("Arial", 10)
            )
            time_label.pack()

        conflict_label = tk.Label(
            detail_window,
            text=f"冲突学生人数: {len(conflict_students)}",
            font=("Arial", 12)
        )
        conflict_label.pack(pady=5)

        if conflict_students:
            student_label = tk.Label(detail_window, text="无法适应课表的学生：", font=("Arial", 10))
            student_label.pack()
            for name in conflict_students:
                name_label = tk.Label(detail_window, text=name, font=("Arial", 10))
                name_label.pack()
        else:
            no_conflict_label = tk.Label(detail_window, text="所有学生都能适应课表。", font=("Arial", 10))
            no_conflict_label.pack()

        def on_select():
            self.select_schedule(result)
            detail_window.destroy()

        select_button = tk.Button(
            detail_window,
            text="选择此方案",
            command=on_select
        )
        select_button.pack(pady=10)

    def select_schedule(self, result):
        self.step_data['selected_schedule'] = result
        self.logger.info(f"用户选择了方案，冲突人数: {result['num_conflicts']}")
        self.next_step()

    def add_handle_conflicts(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        selected_schedule = self.step_data['selected_schedule']['schedule']
        conflict_students = self.step_data['selected_schedule']['conflict_students']

        all_students = self.G11_student_choice + self.G12_student_choice
        student_choices = {student[0]: {'courses': student[1], 'backup_course': student[2]} for student in all_students}

        self.adjusted_students = []

        for student in all_students:
            student_name = student[0]
            courses = student[1]
            backup_course = student[2]
            if student_name in conflict_students:
                if student in self.G11_student_choice:
                    grade = 'G11'
                else:
                    grade = 'G12'
                adjusted_courses, replaced_indices, adjustment_colors, manual_adjustment_needed = self.adjust_student_courses(
                    courses,
                    backup_course,
                    selected_schedule,
                    student_name,
                    grade
                )

                if replaced_indices or manual_adjustment_needed:
                    self.adjusted_students.append({
                        'name': student_name,
                        'grade': grade,
                        'original_courses': courses,
                        'backup_course': backup_course,
                        'adjusted_courses': adjusted_courses,
                        'replaced_indices': replaced_indices,
                        'adjustment_colors': adjustment_colors,
                        'manual_adjustment_needed': manual_adjustment_needed,
                        'confirmed': False,
                        'buttons': {},
                        'adjusted_courses_labels': []
                    })
            else:
                continue

        if not self.adjusted_students:
            messagebox.showinfo("信息", "没有需要调整的学生。")
            self.logger.info("没有需要调整的学生。")
            self.next_step()
            return

        self.adjusted_students.sort(key=lambda s: not s['manual_adjustment_needed'])

        main_frame = tk.Frame(self.content_frame, bg="white")
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=0, minsize=550)

        left_frame = tk.Frame(main_frame, bg="white")
        left_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 5))
        left_frame.config(width=200)
        left_frame.grid_propagate(False)

        right_frame = tk.Frame(main_frame, bg="white", width=550)
        right_frame.grid(row=0, column=1, sticky='n', padx=(5, 0))
        right_frame.grid_propagate(False)

        left_top_frame = tk.Frame(left_frame, bg="white")
        left_top_frame.pack(fill='x', expand=True, pady=(0, 5))

        schedule_label = tk.Label(left_top_frame, text="已选择的课表", font=("Arial", 16), bg="white", wraplength=180,
                                  justify='left')
        schedule_label.pack(padx=6, pady=6, anchor='w')

        for time_period_index, time_period in enumerate(selected_schedule):
            classes_in_time_period = ', '.join(time_period)
            time_label = tk.Label(
                left_top_frame,
                text=f"时间段 {time_period_index + 1}: {classes_in_time_period}",
                font=("微软雅黑", 8),
                bg="white",
                wraplength=180,
                justify='left'
            )
            time_label.pack(anchor='w', padx=8, pady=2)

        left_bottom_frame = tk.Frame(left_frame, bg="white")
        left_bottom_frame.pack(fill='both', expand=True)

        total_confirm_buttons = len(self.adjusted_students)
        remaining_confirm_buttons = sum(1 for student in self.adjusted_students if not student['confirmed'])

        self.tips_label = tk.Label(
            left_bottom_frame,
            text=(
                f"需要确认的学生：剩余 {remaining_confirm_buttons} 个\n"
                "请为每位学生调整课程并确认。当所有学生确认后将自动进入下一步。"
            ),
            font=("Arial", 10, "bold"),
            bg="white",
            fg="blue",
            wraplength=180,
            justify='left'
        )
        self.tips_label.pack(pady=10, padx=10, anchor='w')

        canvas = tk.Canvas(right_frame, bg="white", width=550, height=500)
        scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        canvas.bind_all("<MouseWheel>", lambda event: self._on_mousewheel(event, canvas))

        for student in self.adjusted_students:
            frame = tk.Frame(scrollable_frame, bg="white", bd=1, relief="solid")
            frame.pack(pady=5, padx=5, fill="x")
            self.student_frames[student['name']] = frame

            top_frame = tk.Frame(frame, bg="white")
            top_frame.pack(anchor='w', padx=5, pady=2, fill='x')

            name_label = tk.Label(top_frame, text=f"学生：{student['name']}", font=("Arial", 12, "bold"),
                                  bg="white", wraplength=150, justify='left')
            name_label.pack(side='left')

            if student['confirmed']:
                status_text = ' (已确认)'
                status_font = ("Arial", 12)
                status_fg = "black"
            elif student['manual_adjustment_needed']:
                status_text = ' (需要手动调整)'
                status_font = ("Arial", 12, "bold")
                status_fg = "red"
            else:
                status_text = ' (自动调整成功)'
                status_font = ("Arial", 12)
                status_fg = "green"

            status_label = tk.Label(top_frame, text=status_text, font=status_font, bg="white", fg=status_fg,
                                    wraplength=150, justify='left')
            status_label.pack(side='left', padx=5)

            button_frame = tk.Frame(top_frame, bg="white")
            button_frame.pack(side='right')

            if student['manual_adjustment_needed']:
                manual_button = tk.Button(
                    button_frame,
                    text="开始更改选课",
                    font=("Arial", 8),
                    command=lambda s=student: self.open_manual_adjustment_window(s)
                )
                manual_button.pack(side='left', padx=2)

                confirm_button = tk.Button(
                    button_frame,
                    text="确认",
                    font=("Arial", 8),
                    command=lambda s=student: self.confirm_student(s)
                )
                confirm_button.pack(side='left', padx=2)

                manual_button.config(state='normal', bg='light blue')
                confirm_button.config(state='disabled', bg='gray')
            else:
                manual_button = tk.Button(
                    button_frame,
                    text="更改选课",
                    font=("Arial", 8),
                    command=lambda s=student: self.open_manual_adjustment_window(s)
                )
                manual_button.pack(side='left', padx=2)

                confirm_button = tk.Button(
                    button_frame,
                    text="确认",
                    font=("Arial", 8),
                    command=lambda s=student: self.confirm_student(s)
                )
                confirm_button.pack(side='left', padx=2)

                manual_button.config(state='normal', bg='light blue')
                confirm_button.config(state='normal', bg='light green')

            student['buttons']['manual'] = manual_button
            student['buttons']['confirm'] = confirm_button

            original_courses_text = ', '.join(student['original_courses'])
            original_label = tk.Label(
                frame,
                text=f"原始选课：{original_courses_text}",
                font=("Arial", 8),
                bg="white",
                wraplength=400,
                justify='left'
            )
            original_label.pack(anchor='w', padx=5, pady=2)

            backup_course_text = student['backup_course'] if student['backup_course'] else "无"
            backup_label = tk.Label(
                frame,
                text=f"备选科目：{backup_course_text}",
                font=("Arial", 8),
                bg="white",
                wraplength=200,
                justify='left'
            )
            backup_label.pack(anchor='w', padx=5, pady=2)

            adjusted_courses_frame = tk.Frame(frame, bg="white")
            adjusted_courses_frame.pack(anchor='w', padx=5, pady=2, fill='x')

            adjusted_label = tk.Label(adjusted_courses_frame, text="调整后选课：", font=("Arial", 8), bg="white",
                                      wraplength=80, justify='left')
            adjusted_label.pack(side="left")

            for idx, course in enumerate(student['adjusted_courses']):
                font_style = ("Arial", 8)
                fg_color = "black"
                if idx in student['replaced_indices']:
                    color = student['adjustment_colors'][idx]
                    fg_color = color
                    font_style = ("Arial", 8, "bold")
                course_label = tk.Label(
                    adjusted_courses_frame,
                    text=course,
                    font=font_style,
                    fg=fg_color,
                    bg="white",
                    wraplength=118,
                    justify='left'
                )
                course_label.pack(side="left", padx=2)
                student['adjusted_courses_labels'].append(course_label)

            student['buttons']['status'] = status_label

        remaining_confirm_buttons = sum(1 for s in self.adjusted_students if not s['confirmed'])
        self.tips_label.config(
            text=(
                f"需要确认的学生：剩余 {remaining_confirm_buttons} 个\n"
                "请为每位学生调整课程并确认。当所有学生确认后将自动进入下一步。"
            )
        )

        if remaining_confirm_buttons == 0:
            self.save_adjusted_choices()
            self.next_step()

    def confirm_student(self, student):
        self.logger.debug(f"确认学生 {student['name']} 的调整。")
        student['confirmed'] = True
        self.logger.info(f"学生 {student['name']} 的调整已确认。")
        self.update_student_frame(student)

    def adjust_student_courses(self, courses, backup_course, schedule, student_name, grade):
        adjusted_courses = courses.copy()
        replaced_indices = []
        adjustment_colors = {}
        manual_adjustment_needed = False

        if backup_course:
            steps = []
            num_courses = len(courses)
            for i in range(num_courses - 1, -1, -1):
                steps.append({'replace_indices': [i], 'color': 'orange'})
            for n in range(2, num_courses + 1):
                indices = list(range(num_courses - n, num_courses))
                steps.append({'replace_indices': indices, 'color': 'red'})

            for step in steps:
                temp_courses = adjusted_courses.copy()
                for idx in step['replace_indices']:
                    temp_courses[idx] = backup_course
                if self.check_student_schedule(temp_courses, schedule, grade):
                    adjusted_courses = temp_courses
                    replaced_indices = step['replace_indices']
                    for idx in replaced_indices:
                        adjustment_colors[idx] = step['color']
                    manual_adjustment_needed = False
                    return adjusted_courses, replaced_indices, adjustment_colors, manual_adjustment_needed

            adjusted_courses = courses
            replaced_indices = []
            adjustment_colors = {}
            manual_adjustment_needed = True
            return adjusted_courses, replaced_indices, adjustment_colors, manual_adjustment_needed
        else:
            adjusted_courses = courses
            replaced_indices = []
            adjustment_colors = {}
            manual_adjustment_needed = False
            return adjusted_courses, replaced_indices, adjustment_colors, manual_adjustment_needed

    def open_manual_adjustment_window(self, student):
        student_name = student['name']
        grade = student['grade']
        manual_window = tk.Toplevel(self)
        manual_window.title(f"手动调整 - {student['name']}")
        manual_window.transient(self)
        manual_window.grab_set()
        manual_window.attributes('-topmost', True)

        instruction_label = tk.Label(
            manual_window,
            text="请选择调整后的选课（可多选）：\n点击保存进行确认",
            font=("Arial", 10),
            wraplength=300,
            justify='left'
        )
        instruction_label.pack(pady=5)

        all_courses = set()
        for data in self.course_data_list:
            all_courses.add(data['course_name'])
        all_courses = sorted(all_courses)

        course_vars = {}
        for course in all_courses:
            var = tk.IntVar()
            checkbox = tk.Checkbutton(
                manual_window,
                text=course,
                variable=var,
                font=("Arial", 8)
            )
            checkbox.pack(anchor='w')
            if course in student['adjusted_courses']:
                var.set(1)
            course_vars[course] = var

        def save_manual_adjustment():
            selected_courses = [course for course, var in course_vars.items() if var.get() == 1]

            if self.check_student_schedule(selected_courses, self.step_data['selected_schedule']['schedule'], grade):
                student['adjusted_courses'] = selected_courses
                student['replaced_indices'] = [
                    idx for idx, (orig, adj) in enumerate(zip(student['original_courses'], selected_courses))
                    if orig != adj
                ]
                student['adjustment_colors'] = {idx: 'blue' for idx in student['replaced_indices']}
                student['manual_adjustment_needed'] = False
                student['confirmed'] = False
                self.logger.info(f"学生 {student['name']} 的选课已手动调整并验证通过。")
                manual_window.destroy()
                self.update_student_frame(student)
            else:
                messagebox.showerror("错误", "调整后的选课存在冲突，请重新选择。")

        save_button = tk.Button(
            manual_window,
            text="保存",
            command=save_manual_adjustment
        )
        save_button.pack(pady=10)

    def update_student_frame(self, student):
        self.logger.debug(f"更新学生 {student['name']} 的按钮状态。")
        frame = self.student_frames.get(student['name'])
        if not frame:
            self.logger.error(f"未找到学生 {student['name']} 的框架。")
            return
        manual_button = student['buttons'].get('manual')
        confirm_button = student['buttons'].get('confirm')
        if not manual_button or not confirm_button:
            self.logger.error(f"未找到学生 {student['name']} 的按钮引用。")
            return
        if student['confirmed']:
            manual_button.config(state='disabled', bg='gray')
            confirm_button.config(state='disabled', bg='gray')
        elif student['manual_adjustment_needed']:
            manual_button.config(text="开始更改选课", state='normal', bg='light blue')
            confirm_button.config(state='disabled', bg='gray')
        else:
            if student['manual_adjustment_needed'] == False and not student['confirmed']:
                manual_button.config(text="更改选课", state='normal', bg='light blue')
                confirm_button.config(state='normal', bg='light green')
            else:
                manual_button.config(state='disabled', bg='gray')
                confirm_button.config(state='disabled', bg='gray')
        for child in frame.winfo_children():
            if isinstance(child, tk.Frame):
                for sub_child in child.winfo_children():
                    if isinstance(sub_child, tk.Label):
                        current_text = sub_child.cget("text")
                        if '已确认' in current_text or '需要手动调整' in current_text or '自动调整' in current_text:
                            if student['confirmed']:
                                sub_child.config(text=' (已确认)', font=("Arial", 12), fg="black")
                            elif student['manual_adjustment_needed']:
                                sub_child.config(text=' (需要手动调整)', font=("Arial", 12, "bold"), fg="red")
                            else:
                                sub_child.config(text=' (手动调整成功)', font=("Arial", 12), fg="green")
        adjusted_courses_frame = None
        for child in frame.winfo_children():
            if isinstance(child, tk.Frame) and "调整后选课" in child.winfo_children()[0].cget("text"):
                adjusted_courses_frame = child
                break
        if adjusted_courses_frame:
            for lbl in student['adjusted_courses_labels']:
                lbl.destroy()
            student['adjusted_courses_labels'].clear()
            for idx, course in enumerate(student['adjusted_courses']):
                font_style = ("Arial", 8)
                fg_color = "black"
                if idx in student['replaced_indices']:
                    color = student['adjustment_colors'][idx]
                    fg_color = color
                    font_style = ("Arial", 8, "bold")
                course_label = tk.Label(
                    adjusted_courses_frame,
                    text=course,
                    font=font_style,
                    fg=fg_color,
                    bg="white",
                    wraplength=118,
                    justify='left'
                )
                course_label.pack(side="left", padx=2)
                student['adjusted_courses_labels'].append(course_label)
        remaining_confirm_buttons = sum(1 for s in self.adjusted_students if not s['confirmed'])
        self.tips_label.config(
            text=(
                f"需要确认的学生：剩余 {remaining_confirm_buttons} 个\n"
                "请为每位学生调整课程并确认。当所有学生确认后将自动进入下一步。"
            )
        )
        if remaining_confirm_buttons == 0:
            self.save_adjusted_choices()
            self.next_step()

    def save_adjusted_choices(self):
        for student in self.adjusted_students:
            name = student['name']
            adjusted_courses = student['adjusted_courses']
            grade = student['grade']
            if grade == 'G11':
                for idx, (s_name, s_courses, s_backup) in enumerate(self.G11_student_choice):
                    if s_name == name:
                        self.G11_student_choice[idx] = (s_name, adjusted_courses, s_backup)
                        break
            elif grade == 'G12':
                for idx, (s_name, s_courses, s_backup) in enumerate(self.G12_student_choice):
                    if s_name == name:
                        self.G12_student_choice[idx] = (s_name, adjusted_courses, s_backup)
                        break
        self.logger.info("已保存所有学生的调整选课。")

    def check_student_schedule(self, courses, schedule, grade):
        from itertools import combinations, permutations
        if grade == 'G11':
            allowed_time_slots = range(3)
        else:
            allowed_time_slots = range(len(schedule))
        for time_slot_combination in combinations(allowed_time_slots, len(courses)):
            for assignment in permutations(time_slot_combination):
                conflict = False
                used_slots = set()
                for i, course in enumerate(courses):
                    time_slot = assignment[i]
                    if course in schedule[time_slot] and time_slot not in used_slots:
                        used_slots.add(time_slot)
                    else:
                        conflict = True
                        break
                if not conflict:
                    return True
        return False

    def add_generate_list(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        label = tk.Label(self.content_frame, text="生成名单中...", font=("Arial", 14))
        label.pack(pady=20)
        threading.Thread(target=self.generate_final_assignment, daemon=True).start()

    def generate_final_assignment(self):
        try:
            selected_schedule = self.step_data.get('selected_schedule', {}).get('schedule', ())
            if not selected_schedule:
                raise ValueError("未选择任何排课方案。")
            assign_class = self.step_data.get('assign_class', {})
            if not assign_class:
                raise ValueError("未分配任何课程的分班数。")
            all_students = self.G11_student_choice + self.G12_student_choice
            course_time_mapping = defaultdict(list)
            for time_slot, courses in enumerate(selected_schedule, start=1):
                for course in courses:
                    course_time_mapping[course].append(time_slot)
            student_assignments = {}
            unassigned_students = []
            course_time_counts = defaultdict(lambda: defaultdict(int))
            for student in all_students:
                name, preferences, backup = student
                is_G11 = student in self.G11_student_choice
                if is_G11:
                    allowed_time_slots = set(range(1, 4))
                else:
                    allowed_time_slots = set(range(1, 6))
                selected_courses = preferences
                valid = True
                for course in selected_courses:
                    available_slots = [slot for slot in course_time_mapping[course] if slot in allowed_time_slots]
                    if not available_slots:
                        self.logger.warning(f"{course} for student {name} is not available in allowed time slots.")
                        valid = False
                        break
                if not valid:
                    unassigned_students.append(name)
                    continue
                possible_slots = []
                for course in selected_courses:
                    available_slots = [slot for slot in course_time_mapping[course] if slot in allowed_time_slots]
                    sorted_slots = sorted(available_slots, key=lambda x: course_time_counts[course][x])
                    possible_slots.append([(course, slot) for slot in sorted_slots])
                assigned = {}
                used_slots = set()
                def backtrack(index):
                    if index == len(selected_courses):
                        return True
                    for course, slot in possible_slots[index]:
                        if slot not in used_slots:
                            assigned[course] = slot
                            used_slots.add(slot)
                            if backtrack(index + 1):
                                return True
                            used_slots.remove(slot)
                            del assigned[course]
                    return False
                if backtrack(0):
                    student_assignments[name] = assigned.copy()
                    for course, slot in assigned.items():
                        course_time_counts[course][slot] += 1
                else:
                    unassigned_students.append(name)
            self.logger.info("学生课程分配结果:")
            for student, courses in student_assignments.items():
                self.logger.info(f"\n学生: {student}")
                for course, slot in courses.items():
                    self.logger.info(f"  - {course} (时间段{slot})")
            if unassigned_students:
                self.logger.info("\n未分配学生名单:")
                for student in unassigned_students:
                    self.logger.info(f"  - {student}")
            else:
                self.logger.info("\n所有学生均已成功分配课程。")
            course_students = defaultdict(list)
            for student, courses in student_assignments.items():
                for course, slot in courses.items():
                    course_students[course].append((student, slot))
            final_assignment = {}
            for course, students in course_students.items():
                for student, slot in students:
                    if slot not in final_assignment:
                        final_assignment[slot] = {}
                    if course not in final_assignment[slot]:
                        final_assignment[slot][course] = []
                    final_assignment[slot][course].append(student)
            final_assignment = dict(final_assignment)
            self.step_data['final_assignment'] = final_assignment
            self.step_data['student_assignments'] = student_assignments
            self.step_data['unassigned_students'] = unassigned_students
            self.logger.info("已生成最终分配方案。")
            if unassigned_students:
                messagebox.showwarning("警告", f"有 {len(unassigned_students)} 位学生未能分配课程。请检查日志以获取详细信息。")
            self.after(0, self.generate_list_complete)
        except Exception as e:
            self.logger.error(f"生成名单过程中出现错误：{e}")
            messagebox.showerror("错误", f"生成名单过程中出现错误：{e}")

    def generate_list_complete(self):
        final_assignment = self.step_data.get('final_assignment', {})
        if not final_assignment:
            messagebox.showerror("错误", "未找到最终分配数据，请在前面步骤中生成该数据后再试。")
            return
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        main_frame = tk.Frame(self.content_frame, bg="white")
        main_frame.pack(fill='both', expand=True)
        top_frame = tk.Frame(main_frame, bg="white", height=int(self.winfo_height() * 0.2))
        top_frame.pack(fill='x', padx=10, pady=10)
        message_label = tk.Label(
            top_frame,
            text="AP名单生成结果",
            font=("Arial", 16, "bold"),
            bg="white",
            fg="green"
        )
        message_label.pack(pady=0)
        confirm_button = tk.Button(
            top_frame,
            text="确认",
            command=self.confirm_final_assignment,
            font=("Arial", 12),
            width=15,
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049"
        )
        confirm_button.pack(pady=2)
        lower_frame = tk.Frame(main_frame, bg="white")
        lower_frame.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        canvas = tk.Canvas(lower_frame, bg="white")
        scrollbar = ttk.Scrollbar(lower_frame, orient='vertical', command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        canvas.bind_all("<MouseWheel>", lambda event: self._on_mousewheel(event, canvas))
        time_slot_colors = ["#FFD700", "#00BFFF", "#9932CC", "#FFA07A", "#ADFF2F"]
        course_color_list = [
            "#FAFAD2", "#F1F8E9", "#E8E8E8", "#D3E9F5", "#F6F9F1",
            "#F9F0F0", "#D0F0C0", "#F2F2F2", "#E4F4F1", "#D9F7E5",
            "#CBE8F1", "#E7D8E9"
        ]
        unique_courses = set()
        for slot in final_assignment.values():
            unique_courses.update(slot.keys())
        unique_courses = sorted(unique_courses)
        course_color_mapping = {course: course_color_list[i % len(course_color_list)] for i, course in
                                enumerate(unique_courses)}
        flattened = []
        time_slots = sorted(final_assignment.keys())
        for ts in time_slots:
            course_dict = final_assignment[ts]
            for course, students in course_dict.items():
                flattened.append((ts, course, students))
        if not flattened:
            messagebox.showinfo("生成名单", "没有学生被分配到任何课程。")
            self.logger.info("没有学生被分配到任何课程。")
            self.next_step()
            return
        for col, (ts, course, students) in enumerate(flattened):
            ts_label = tk.Label(
                scrollable_frame,
                text=f"时段 {ts}",
                bg=time_slot_colors[(ts - 1) % len(time_slot_colors)],
                fg="black",
                font=("Arial", 10, "bold"),
                borderwidth=1,
                relief='solid',
                width=5,
                wraplength=45,
                justify='center'
            )
            ts_label.grid(row=0, column=col, sticky='nsew', padx=1, pady=1)
        for col, (ts, course, students) in enumerate(flattened):
            c_color = course_color_mapping.get(course, "#FFFFFF")
            course_label = tk.Label(
                scrollable_frame,
                text=course,
                bg=c_color,
                fg="black",
                font=("Arial", 7, "bold"),
                borderwidth=1,
                relief='solid',
                width=5,
                wraplength=45,
                justify='center'
            )
            course_label.grid(row=1, column=col, sticky='nsew', padx=1, pady=1)
        for col, (ts, course, students) in enumerate(flattened):
            count_text = f"{len(students)}人"
            count_label = tk.Label(
                scrollable_frame,
                text=count_text,
                bg="white",
                fg="black",
                font=("Arial", 8),
                borderwidth=1,
                relief='solid',
                width=5,
                wraplength=45,
                justify='center'
            )
            count_label.grid(row=2, column=col, sticky='nsew', padx=1, pady=1)
        for col, (ts, course, students) in enumerate(flattened):
            for i, student in enumerate(students):
                student_label = tk.Label(
                    scrollable_frame,
                    text=student,
                    bg=course_color_mapping.get(course, "#FFFFFF"),
                    fg="black",
                    font=("Arial", 8),
                    borderwidth=1,
                    relief='solid',
                    width=5,
                    wraplength=48,
                    justify='center'
                )
                student_label.grid(row=3 + i, column=col, sticky='nsew', padx=1, pady=1)
        for col in range(len(flattened)):
            scrollable_frame.grid_columnconfigure(col, weight=1)
        max_students = max(len(students) for _, _, students in flattened)
        for row in range(3 + max_students):
            scrollable_frame.grid_rowconfigure(row, weight=1)
        self.logger.info("用户完成生成名单步骤。")

    def confirm_final_assignment(self):
        self.logger.info("用户确认了最终分配名单。")
        self.next_step()

    def add_save_results(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        title_label = tk.Label(self.content_frame, text="保存结果", font=("Arial", 20), bg="white")
        title_label.pack(pady=10)
        save_button = tk.Button(
            self.content_frame,
            text="保存结果",
            command=self.save_results_complete,
            width=15
        )
        save_button.pack(pady=10)
        label = tk.Label(self.content_frame,
                         text="提示:excel表中有四个分表",
                         font=("Arial", 12), bg="white")
        label.pack(pady=20)
        label = tk.Label(self.content_frame,
                         text="如果您有任何对本程序的建议或意见，请您与本项目管理者trz129.com@gmail.com 联系",
                         font=("Arial", 12), bg="white")
        label.pack(pady=20)
        label = tk.Label(self.content_frame,
                         text="AP module GUI Version 1.0",
                         font=("Arial", 12), bg="white")
        label.pack(pady=5)
        label = tk.Label(self.content_frame,
                         text="MIT license © trz129 2024",
                         font=("Arial", 12), bg="white")
        label.pack(pady=5)

    def save_results_complete(self):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="保存结果文件"
        )
        if save_path:
            try:
                allocations = self.step_data.get('allocations', {})
                allocations_with_class = []
                for course, details in allocations.items():
                    course_data = next((item for item in self.course_data_list if item["course_name"] == course), None)
                    if course_data:
                        allocations_with_class.append({
                            "课程名称": course,
                            "教师名称": details['teacher'],
                            "教室人数": details['max_students'],
                            "分班数": course_data.get('assign_class')
                        })
                df_allocations = pd.DataFrame(allocations_with_class)
                final_assignment = self.step_data.get('final_assignment', {})
                flattened = []
                time_slots = sorted(final_assignment.keys())
                for ts in time_slots:
                    course_dict = final_assignment[ts]
                    for course, students in course_dict.items():
                        flattened.append({
                            "时间段": ts,
                            "课程名称": course,
                            "学生名单": ", ".join(students),
                            "学生人数": len(students)
                        })
                df_final_assignment = pd.DataFrame(flattened)
                unassigned_students = self.step_data.get('unassigned_students', [])
                df_unassigned = pd.DataFrame({"未分配学生": unassigned_students})
                all_assignments = []
                student_assignments = self.step_data.get('student_assignments', {})
                for student, courses in student_assignments.items():
                    for course, slot in courses.items():
                        all_assignments.append({
                            "学生姓名": student,
                            "课程名称": course,
                            "时间段": slot
                        })
                df_student_assignments = pd.DataFrame(all_assignments)
                with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                    df_allocations.to_excel(writer, sheet_name='课程分配', index=False)
                    df_final_assignment.to_excel(writer, sheet_name='最终排课', index=False)
                    df_unassigned.to_excel(writer, sheet_name='未分配学生', index=False)
                    df_student_assignments.to_excel(writer, sheet_name='学生分配', index=False)
                messagebox.showinfo("保存结果", f"结果已成功保存到 {save_path}")
                self.logger.info(f"分配结果已保存到 {save_path}")
                self.next_step()
            except Exception as e:
                error_message = f"保存结果失败：{e}"
                messagebox.showerror("错误", error_message)
                self.logger.error(error_message)
        else:
            messagebox.showwarning("保存结果", "保存操作已取消。")
            self.logger.warning("用户取消了保存结果操作。")

if __name__ == "__main__":
    app = AP_Scheduler()
    app.mainloop()
