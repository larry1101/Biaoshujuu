"""
Python3!!!!!!!
下好包
先看_ini_data里面最后面几行自己电脑能不能用！
source文件3列：序号，标题，内容
右边点了之后 键盘 ↓下一个字段 ↑上一个字段 →序号一样的下一个产品 回车下一行
"""

import os
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import LoadFileDialog

import jieba
import xlrd
import xlwt


class DataDiv:
    def __init__(self):
        self._ini_data()
        self._ini_widget()

        self._save_data()

    def _ini_data(self):
        jieba.initialize()

        self.PREPARED = 0

        self.DATA_S = None
        self.DATA_D = None

        self.fields = []
        self.target_entrys = []
        self.target_row_cell_values = []

        self.xuhaos = []
        self.titles = []
        self.conten = []

        self.proced_rows = []

        # self.start_xuhao = None

        self.current_row = -1
        self.current_row_cell_index = 0

        self.COUNT = 0

        self.NO_FILE_WARNING = '请选好上面俩文件先'

        # ======================================这里可以设置右侧treeview每行多少列~
        self.lv_col_num = 16
        # ======================================这里可以设置左侧拼合segments的listbox宽度~
        self.width_l_t_r = 32
        # ======================================存储地点~
        self.FILE_SAVE_DIR = r'E:\标数据\标好的'
        if not os.path.exists(self.FILE_SAVE_DIR):
            os.makedirs(self.FILE_SAVE_DIR)

    def _save_data(self):

        if self.PREPARED < 2:
            return

        if self.current_row <= 0:
            return

        if self.COUNT <= 0:
            return

        print('saving...')
        wb = xlwt.Workbook()
        sh = wb.add_sheet('Sheet1')

        r = 0
        c = 0
        for field in self.fields:
            sh.write(r, c, field)
            c += 1
        r += 1
        for row in self.proced_rows:
            c = 0
            for cell in row:
                sh.write(r, c, cell)
                c += 1
            r += 1

        wb.save(self.FILE_SAVE_DIR + os.sep + '%d.xls' % self.xuhaos[self.current_row - 1])

        print('进度已保存')

    def _ini_widget(self):
        self.root_window = Tk()
        self.root_window.title('biaoshuju  右边点了之后 键盘 ↓下一个字段 ↑上一个字段 →序号一样的下一个产品 回车下一行  点左边字段名清除这个item该字段的value')

        # operation panel
        frame_opera = Frame(self.root_window)
        frame_opera.pack(side=LEFT, fill=BOTH)

        # data files selection panel
        frame_data_op = Frame(frame_opera)
        frame_data_op.pack(side=TOP, fill=X)

        Button(frame_data_op,
               text='select source',
               command=self.on_button_pick_s_click
               ).pack(side=TOP, fill=X)
        self.label_file_s = Label(frame_data_op, text='不要headings！直接数据！')
        self.label_file_s.pack(side=TOP, fill=X)

        Button(frame_data_op,
               text='选择 模板',
               command=self.on_button_pick_d_click
               ).pack(side=TOP, fill=X)
        self.label_file_d = Label(frame_data_op, text='这个是模板，到时候存到另一个文件里')
        self.label_file_d.pack(side=TOP, fill=X)

        self.start_xuhao = StringVar()
        self.start_xuhao.set('起始的content序号，选好上面两个文件先')
        self.entry_start_id = Entry(frame_data_op, textvariable=self.start_xuhao)
        self.entry_start_id.pack(side=TOP, fill=X)

        Button(frame_data_op,
               text='START',
               command=self.on_button_start_click,
               height=3,
               bg='#B9C4D5'
               ).pack(side=TOP, fill=X)

        # words combinator panel
        frame_targets = Frame(frame_opera)
        frame_targets.pack(fill=BOTH, expand=True)

        self.l_target_field = Listbox(frame_targets,
                                      width=10,
                                      selectbackground='white',
                                      selectforeground='black'
                                      )
        self.l_target_field.bind('<ButtonRelease-1>', self.on_t_f_clicked)
        self.l_target_field.pack(side=LEFT, fill=BOTH)

        self.l_target_row = Listbox(frame_targets, width=self.width_l_t_r + 1)
        self.l_target_row.pack(side=LEFT, fill=BOTH)

        Button(frame_targets,
               text='Previous field',
               command=self.on_btn_pre_field_clicked,
               height=4
               ).pack(side=TOP, fill=X, pady=5)
        Button(frame_targets,
               text='Next field',
               command=self.on_btn_next_field_clicked,
               height=4
               ).pack(side=TOP, fill=X, pady=5)
        Button(frame_targets,
               text='Next row',
               command=self.on_btn_next_row_clicked,
               height=8,
               bg='#C8ECC8'
               ).pack(side=TOP, fill=X, pady=5)
        Button(frame_targets,
               text='同序号添加item',
               command=self.on_btn_add_item_clicked,
               height=4
               ).pack(side=TOP, fill=X, pady=5)
        Button(frame_targets,
               text='Exit',
               command=self.root_window.destroy,
               bg='#FBE6E6',
               height=3
               ).pack(side=BOTTOM, fill=X, pady=5, padx=5)

        # segments selector panel
        frame_segs = Frame(self.root_window)
        frame_segs.pack(fill=BOTH, expand=True)

        self.tree_s_s = ttk.Treeview(
            frame_segs,
            columns=list(range(self.lv_col_num)),
            show='headings',
            selectmode='browse',
            # 行数in show
            height=35
        )

        for i in range(self.lv_col_num):
            self.tree_s_s.column('%d' % i, width=64, anchor='center')
            self.tree_s_s.heading('%d' % i, text='%d' % i)

        self.tree_s_s.bind("<ButtonRelease-1>", self.on_tree_s_s_click)
        self.tree_s_s.bind("<Key>", self.on_tree_key)
        self.tree_s_s.pack(side=LEFT, fill=BOTH, expand=True)

        tree_vbar = ttk.Scrollbar(frame_segs, orient=VERTICAL, command=self.tree_s_s.yview)
        self.tree_s_s.configure(yscrollcommand=tree_vbar.set)
        tree_vbar.pack(side=RIGHT, fill=Y)

        # open window
        self.root_window.mainloop()

    def on_button_pick_s_click(self):
        fd = LoadFileDialog(self.root_window)  # 创建打开文件对话框
        filename = fd.go()
        if filename is None:
            return

        if os.path.splitext(filename)[1] != '.xlsx' and os.path.splitext(filename)[1] != '.xls':
            print(os.path.splitext(filename)[1], 'Not Excel file, reselect')
            return

        self.label_file_s['text'] = filename
        self.DATA_S = filename  # to do ?????????

        wb = xlrd.open_workbook(filename)
        sh = wb.sheet_by_index(0)
        rows = sh.get_rows()
        for row in rows:
            self.xuhaos.append(int(row[0].value))
            self.titles.append(row[1].value)
            self.conten.append(row[2].value)

        # print(self.xuhaos)
        # print(self.titles)
        # print(self.conten)

        if os.path.exists(self.FILE_SAVE_DIR):
            saved_files_names = [int(os.path.splitext(os.path.split(filename)[1])[0]) for filename in
                                 os.listdir(self.FILE_SAVE_DIR)]
            if saved_files_names.__len__() == 0:
                self.start_xuhao.set(self.xuhaos[0])
            else:
                saved_files_names.sort()
                self.start_xuhao.set(saved_files_names[-1] + 1)

        else:
            self.start_xuhao.set(self.xuhaos[0])

        # self.start_xuhao.set(self.xuhaos[0])

        self.PREPARED += 1

    def on_button_pick_d_click(self):
        fd = LoadFileDialog(self.root_window)  # 创建打开文件对话框
        filename = fd.go()

        if filename is None:
            return

        if os.path.splitext(filename)[1] != '.xlsx' and os.path.splitext(filename)[1] != '.xls':
            print(os.path.splitext(filename)[1], 'Not Excel file, reselect')
            return

        self.label_file_d['text'] = filename
        data_mo = filename

        wb = xlrd.open_workbook(data_mo)
        sh = wb.sheet_by_index(0)
        fields = sh.row(0)

        # to do need?
        # sh = None
        # wb.release_resources()
        # wb = None

        # fields monitor
        self.fields.clear()
        self.target_entrys.clear()
        self.target_row_cell_values.clear()
        self.l_target_field.delete(0, END)

        for field in fields:
            self.fields.append(field.value)
            self.l_target_field.insert(END, field.value)

            # entry
            s = StringVar()
            self.target_row_cell_values.append(s)
            entry = Entry(self.l_target_row, width=self.width_l_t_r, textvariable=s)
            self.target_entrys.append(entry)
            self.l_target_row.insert(END, entry)
            x, y, width, height = self.l_target_row.bbox(END)
            entry.place(x=x, y=y)

        self.PREPARED += 1

    def on_tree_s_s_click(self, event=None):
        if self.tree_s_s.selection() == '':
            return
        item = self.tree_s_s.selection()[0]
        column = self.tree_s_s.identify_column(event.x)
        column = int(column[column.find('#') + 1:]) - 1
        if self.tree_s_s.identify_region(event.x, event.y) == 'cell':
            if column > self.tree_s_s.item(item, "values").__len__() - 1:
                print('click a null')
                return
            seg = self.tree_s_s.item(item, "values")[column]

            self.target_row_cell_values[self.current_row_cell_index].set(
                self.target_row_cell_values[self.current_row_cell_index].get() + seg)

    def on_button_start_click(self):
        if not self.PREPARED:
            print(self.NO_FILE_WARNING)
            return
            # start process

        if self.PREPARED < 2:
            return

        x = int(self.start_xuhao.get())
        self.current_row = -1
        for xuhao in self.xuhaos:
            self.current_row += 1
            if x == xuhao:
                break
        print('start at', self.current_row)

        self.show_content(self.current_row)

    def show_content(self, index):
        seg_xuhao = self.xuhaos[index]
        seg_title = list(jieba.cut(self.titles[index]))
        seg_conte = list(jieba.cut(self.conten[index]))
        # print(seg_title)
        # print(seg_conte)

        items = self.tree_s_s.get_children()
        [self.tree_s_s.delete(item) for item in items]

        for i in range(0, len(seg_title), self.lv_col_num):
            self.tree_s_s.insert('', 'end', values=seg_title[i:i + self.lv_col_num])
        for i in range(0, len(seg_conte), self.lv_col_num):
            self.tree_s_s.insert('', 'end', values=seg_conte[i:i + self.lv_col_num])

        [item.set('') for item in self.target_row_cell_values]
        self.target_row_cell_values[0].set(seg_xuhao)

        self.current_row_cell_index = 1

    def on_btn_next_field_clicked(self):
        if not self.PREPARED:
            print(self.NO_FILE_WARNING)
            return
        self.current_row_cell_index += 1
        if self.current_row_cell_index >= self.fields.__len__():
            self.current_row_cell_index = self.fields.__len__() - 1
        self.sel_target_cell(self.current_row_cell_index)

    def on_btn_pre_field_clicked(self):
        if not self.PREPARED:
            print(self.NO_FILE_WARNING)
            return
        self.current_row_cell_index -= 1
        if self.current_row_cell_index < 0:
            self.current_row_cell_index = 0
        self.sel_target_cell(self.current_row_cell_index)

    def sel_target_cell(self, index):
        self.l_target_row.selection_clear(0, END)
        self.l_target_row.selection_set(self.current_row_cell_index)

    def on_btn_next_row_clicked(self):
        if not self.PREPARED:
            print(self.NO_FILE_WARNING)
            return
        # save
        self.proced_rows.append([cell.get() for cell in self.target_row_cell_values])
        # print(self.proced_rows)

        # next row
        self.current_row_cell_index = 0
        self.current_row += 1

        if self.current_row >= self.xuhaos.__len__():
            print('No more, finish, please exit and notify Binyann later')
            return

        self.show_content(self.current_row)
        self.COUNT += 1

    def on_btn_add_item_clicked(self):
        if not self.PREPARED:
            print(self.NO_FILE_WARNING)
            return
        # save
        self.proced_rows.append([cell.get() for cell in self.target_row_cell_values])

        # same row new item
        self.current_row_cell_index = 1

        [item.set('') for item in self.target_row_cell_values]
        self.target_row_cell_values[0].set(self.xuhaos[self.current_row])

        self.COUNT += 1

    def on_tree_key(self, event):
        # print(event.keysym)
        if not self.PREPARED:
            print(self.NO_FILE_WARNING)
            return

        if event.keysym == 'Return':
            # 在右侧选择完后按下了回车
            self.on_btn_next_row_clicked()
        elif event.keysym == 'Up':
            # 在右侧选择完后按下了↑
            self.on_btn_pre_field_clicked()
        elif event.keysym == 'Down':
            # 在右侧选择完后按下了↓
            self.on_btn_next_field_clicked()
        elif event.keysym == 'Right':
            # 在右侧选择完后按下了→
            self.on_btn_add_item_clicked()

    def on_t_f_clicked(self, event=None):
        sel = self.l_target_field.curselection()
        if sel == ():
            return
        self.target_row_cell_values[sel[0]].set('')


DataDiv()
