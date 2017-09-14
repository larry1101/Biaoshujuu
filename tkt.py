from tkinter import *
from tkinter import ttk

import jieba

# data
doc = "['YAWARA多合一面霜是由韩国株式会社艾丽公司与日本核酸原料研究所LS库弗雷斯（音译）厂家联合开发的产品。\u3000\u3000YAWARA多合一面霜是爽肤水，乳液，保湿霜，补水面膜，基础霜等护肤品合而为一的产品（懒癌患者看这里）。这款面霜绝无添加对羟基苯甲酸酯，界面活性剂，食用合成染料，矿物油等化学成分（特别安全）。这款面霜富含日本核酸研发所LS库弗雷斯（音译）厂家独家研发的核酸原料，这个原料提取于三文鱼卵子的DNA（好高大上是吧），可强力达到肌肤深层发挥抗衰老功效和肌肤再生作用哦~在紫外线辐射极强的春季夏季和秋季使用（好像一年四季都挺强的），具有镇静肌肤和修复紫外线伤害的功效，另外，YAWARA多合一面霜含有的薏米原料可有效达到美白的功效。面霜质地是凝乳质地延展性极好，味道也是淡淡的花香挺好闻的，吸收的特别快，使用后肤色提亮，光泽度有所提升，而且一大罐100ml性价比太高了']"
s = list(jieba.cut(doc))

# UI
root_window = Tk()
root_window.title('biaoshuju')

frame_segs = Frame(root_window)
frame_segs.pack(fill=BOTH, expand=True)

# ListBox
# lb_segs = Listbox(frame_segs)
# lb_segs.pack(fill=BOTH, expand=True)
#
#
# def on_seg_select(event=None):
#     if lb_segs.curselection() == ():
#         return
#     print(s[lb_segs.curselection()[0]])
#
#
# lb_segs.bind("<ButtonRelease-1>", on_seg_select)
#
# for seg in s:
#     lb_segs.insert(END, seg)


# TreeView
bookList=[('aaa',123,','),('bbb',123,'['),('xxx',123,'pp'),('sss',123,'pp'),('ddd',123,'pp')]

col_num=10
tree=ttk.Treeview(
    frame_segs
    ,columns=list(range(col_num))
    ,show='headings'
    ,selectmode='browse'
    ,height=10 # 行数in show
)
for i in range(col_num):
    tree.column('%d'%i, width=100, anchor='center')
    tree.heading('%d'%i, text='%d'%i)

# tree.heading('name',text='name')
# tree.heading('price',text='price')
# tree.heading('3',text='3', anchor='center')

# for _ in map(tree.delete, tree.get_children("")):
#     pass

for item in bookList:
    tree.insert('','end',values=item)

def on_l_w_click(event=None):
    if tree.selection() == '':
        return
    item = tree.selection()[0]
    # rowid = tree.identify_row(event.y)
    column = tree.identify_column(event.x)
    column = int(column[column.find('#')+1:])-1
    # print(tree.item(item, "values"))
    # print(tree.identify_element(event.x,event.y))
    # print(tree.identify_region(event.x,event.y))
    # print(tree.bbox(item,column))
    if tree.identify_region(event.x,event.y) == 'cell':
        if column>tree.item(item, "values").__len__()-1:
            print('click a null')
            return
        print(tree.item(item, "values")[column])

tree.bind("<ButtonRelease-1>", on_l_w_click)
tree.pack(fill=BOTH, expand=True)


root_window.mainloop()
