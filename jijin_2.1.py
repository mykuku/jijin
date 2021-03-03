"""
    新增功能：
        1.自动排序，根据基金的估值自动从涨幅大到小排序
        2.显示上一分钟的跟这一分钟的涨幅比较的上升（或下降）箭头
"""

import requests, openpyxl, base64, json, sys,re
import os, shutil, hmac, hashlib, urllib
from time import sleep
import ast, threading, _tkinter
import tkinter as tk
from tkinter import ttk
import datetime, time
from decimal import *
from PIL import Image, ImageTk


# 老规矩，多线程任务，防止窗口卡死
def my_thread(func, *args):
    t = threading.Thread(target=func, args=args)
    t.setDaemon(True)
    t.start()


# 基金代码以及份额数据输入
def add_jijin():
    # 框架设置
    add_window = tk.Toplevel(jijing_root)
    add_window.title('基金代码添加')
    add_window.geometry('250x100')
    add_window.resizable(0, 0)

    # 添加按钮函数：
    def add_one():
        if number.get() == None or number.get() == '' or fund_share.get() == None or fund_share.get() == '':
            pass
        else:
            if number.get() in jijin_code:
                pass
            else:
                jijin_code.append(number.get())
                tree.insert('', len(jijin_code), values=(number.get()))
                ws[f'A{len(jijin_code)}'] = number.get()
                fund_share_dir[number.get()]=fund_share.get()
                fund_share_list.append(fund_share.get())
                ws[f'B{len(fund_share_list)}'] = fund_share.get()
                wb.save(path)

    # 内容
    tk.Label(add_window, text='基金代码输入：').place(x=10, y=5)
    number = tk.StringVar()
    number_entry = tk.Entry(add_window, textvariable=number, width=17)
    number_entry.place(x=110, y=5)
    # 按钮
    # tk.Label(add_window, text='PS:键盘点击“Delete”\n简易框中的基金号码\n可以删除输入错误\n或者不用的基金号码').grid(row=1, column=0)
    tk.Label(add_window, text='请输入该基金\n你所持有的份额:', font=('', 8)).place(x=10, y=25)
    fund_share = tk.StringVar()
    fund_share_entry = tk.Entry(add_window, textvariable=fund_share, width=17)
    fund_share_entry.place(x=110, y=30)

    tk.Button(add_window, text='添加', width=10, height=1, bg='#24a9ff', command=lambda: my_thread(add_one)).place(x=80,
                                                                                                                 y=55)
    tk.Label(add_window, text='PS:键盘点击“Delete”简易框中的基金号码可以删除输入错误或者不用的基金号码', font=('', 7)).place(x=0, y=85)


# 获取数据按钮执行函数
def jijin_run():
    # 获取数据函数
    def get_data():
        global  fund_dir_last, data_list_sort, gszzl_dir_last,fund_chang
        content = ''
        gszzl_count = 0
        fund_share_count = 0
        headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36"
        }
        print(jijin_code)
        for jijin in jijin_code:
            url = f'http://fundgz.1234567.com.cn/js/{jijin}.js?rt={timestamp}000'
            try:
                response = requests.get(url=url, headers=headers)
            except:
                sleep(60)
                response = requests.get(url=url, headers=headers)
            response.encoding = 'UTF-8'
            # datasplit = response.text.split('{')
            # datasplit = datasplit[1].split('}')
            # datasplit = datasplit[0]
            # datasplit=dict(datasplit)
            # print(datasplit)
            print(response.text)
            data=re.search(r'\(([\s\S]*?)$',response.text).group(1)[:-1]
            #print(data)
            data=data.replace(')','').replace('(','')
            #print(data)
            data = eval(data)

            fund_share = (Decimal(fund_share_dir[data['fundcode']]) * Decimal(data['gszzl']) / 100 * Decimal(
                data['dwjz'])).quantize(Decimal('0.00'))
            # gszzl_dir[data['fundcode']]=data['gszzl']
            # fund_dir[data['fundcode']]=fund_share
            # 'fundcode': '320007', 'name': '诺安成长混合', 'jzrq': '2020-08-27', 'dwjz': '1.7640', 'gsz': '1.7635', 'gszzl': '-0.03', 'gztime': '2020-08-28 15:00'}
            data = [data['fundcode'], data['name'][:4], data['gszzl'], data['gztime'][11:], fund_share]
            # ('320007', '诺安成长', '-0.03', '08-28 15:00', Decimal('-0.02'))
            data_list.append(data)
            data_list_sort = (sorted(data_list, key=lambda i: i[-1], reverse=True))
        #print(data_list_sort)
        for data1 in data_list_sort:
            gszzl_dir[data1[0]] = data1[2]
            fund_dir[data1[0]] = data1[4]
            gszzl_count += Decimal(data1[2])
            fund_share_count += data1[4]
        gszzl_dir['gszzl_count'] = str(gszzl_count)
        fund_dir['fund_share_count'] = fund_share_count

        # print(fund_dir_last)
        # print(fund_dir)
        # print(fund_chang)


        if len(gszzl_dir_last) == 0:
            pass
        else:
            for new in gszzl_dir.keys():
                if float(gszzl_dir[new]) >= 0:
                    gszzl_dir[new] = '+' + gszzl_dir[new]
                else:
                    gszzl_dir[new] = gszzl_dir[new]

        # {'001302': '+3.15', '001300': '+0.69', '003634': '+0.27', '161616': '+0.06', '320007': '+0.92',
        # {'001302': '3.15', '001300': '0.69', '003634': '0.18', '161616': '0.10', '320007': '0.88', '001838': '-0.01',
        # {'001302': Decimal('15.55'), '001300': Decimal('4.29'), '003634': Decimal('1.32'), '161616': Decimal('0.63'),
        # {'001302': '+15.55↗', '001300': '+4.29↗', '003634': '+1.98↗', '161616': '+0.38↗', '320007': '+0.58↗',

        if len(fund_dir_last) == 0:
            fund_chang=fund_dir.copy()
        else:
            for num in fund_dir.keys():
                # print(fund_dir[num])
                # print(fund_dir_last[num])
                if fund_dir[num] - fund_dir_last[num]>= 0:
                    if fund_dir[num] >= 0:
                        fund_chang[num] = '+' + str(fund_dir[num]) + ' ↗ +'+str(fund_dir[num] - fund_dir_last[num])
                    else:
                        fund_chang[num] = str(fund_dir[num]) + ' ↗ +'+str(fund_dir[num] - fund_dir_last[num])
                else:
                    if float(fund_dir[num]) >= 0:
                        fund_chang[num] = '+' + str(fund_dir[num]) + ' ↘ '+str(fund_dir[num] - fund_dir_last[num])
                    else:
                        fund_chang[num] = str(fund_dir[num]) + ' ↘ '+str(fund_dir[num] - fund_dir_last[num])



        for data_item, data2 in enumerate(data_list_sort):
            tree.set(f'I00{hex(data_item + 1)[2:].upper()}', column=columns[0], value=data2[0])
            tree.set(f'I00{hex(data_item + 1)[2:].upper()}', column=columns[1], value=data2[1])
            tree.set(f'I00{hex(data_item + 1)[2:].upper()}', column=columns[2], value=gszzl_dir[data2[0]])
            tree.set(f'I00{hex(data_item + 1)[2:].upper()}', column=columns[3], value=fund_chang[data2[0]])
            tree.set(f'I00{hex(data_item + 1)[2:].upper()}', column=columns[4], value=data2[3])
            con = data2[1] + "：  " + gszzl_dir[data2[0]] + '  ' + str(fund_chang[data2[0]])
            content += con + '\n'
        if len(fund_dir_last) == 0:
            tree.insert('',len(data_list_sort)+1, value=('', '总计：',gszzl_dir['gszzl_count'],fund_chang['fund_share_count'],date_yy[11:-3]))
        tree.set(f'I00{hex(len(data_list_sort) + 1)[2:].upper()}', column=columns[1], value='总计：')
        tree.set(f'I00{hex(len(data_list_sort) + 1)[2:].upper()}', column=columns[2], value=gszzl_dir['gszzl_count'])
        tree.set(f'I00{hex(len(data_list_sort) + 1)[2:].upper()}', column=columns[3], value=fund_chang['fund_share_count'])
        tree.set(f'I00{hex(len(data_list_sort) + 1)[2:].upper()}', column=columns[4], value=date_yy[11:-3])
        gszzl_dir_last = gszzl_dir.copy()
        fund_dir_last = fund_dir.copy()
        data_list.clear()
        if rt == 1:
            url = s_msg()
            content = content + '基金总计: ' + gszzl_dir['gszzl_count'] + '  ' + str(
                fund_chang['fund_share_count']) + '\n' + '\t\t\t\t\t\t\t' + date_yy[5:-3]
            # print(content)
            remind_msg(content=content, url=url)

    # 钉钉另外提醒
    def remind_msg(content, url):
        headers = {'Content-Type': 'application/json;charset=utf-8'}
        data = {
            "msgtype": "text",
            "text": {
                "content": content
            },
            "at": {
                # "atMobiles":['13751759726'],
                "isAtAll": False
            },
        }
        r = requests.post(url, data=json.dumps(data), headers=headers
                          )

        return r.text

    # 钉钉机器人webhook生成url链接
    def s_msg():
        timestamp = str(round(time.time() * 1000))
        secret = 'SECdf9fdef28f5fb98337d0ecb3ea54db9f6a904cd56a802df52429a9fcacc509a4'
        secret_enc = secret.encode('utf-8')
        string_to_sign = '{}\n{}'.format(timestamp, secret)
        string_to_sign_enc = string_to_sign.encode('utf-8')
        hmac_code = hmac.new(secret_enc, string_to_sign_enc, digestmod=hashlib.sha256).digest()
        sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))
        url = f'https://oapi.dingtalk.com/robot/send?access_token=9f35a78a9fc15a9481ab38137eb6c57c468ba1df5d55f11d077048f5e81e9004&timestamp={timestamp}&sign={sign}'

        return url

    rt = 0
    # 循环运行
    while True:

        # 利用时间戳来判定时间断运行
        date_yy = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        timearray = time.strptime(date_yy, '%Y-%m-%d %H:%M:%S')
        timestamp = int(time.mktime(timearray))
        t1 = int(time.mktime(time.strptime(datetime.datetime.now().strftime('%Y-%m-%d 09:30:00'), '%Y-%m-%d %H:%M:%S')))
        t2 = int(time.mktime(time.strptime(datetime.datetime.now().strftime('%Y-%m-%d 11:33:00'), '%Y-%m-%d %H:%M:%S')))
        t3 = int(time.mktime(time.strptime(datetime.datetime.now().strftime('%Y-%m-%d 13:00:00'), '%Y-%m-%d %H:%M:%S')))
        t4 = int(time.mktime(time.strptime(datetime.datetime.now().strftime('%Y-%m-%d 15:03:00'), '%Y-%m-%d %H:%M:%S')))
        if t1 <= timestamp <= t2 or t3 <= timestamp <= t4:
            get_data()
            rt += 1
            sleep(60)
        elif t2 <= timestamp <= t3:
            get_data()
            sleep(900)
        else:
            #rt=1
            get_data()

            quit()

        if rt > 4:
            rt = 0


# 增加双击删除功能：
def del_num(event):
    # 删除行函数
    def deleterows(row_num):
        for row in range(row_num, ws.max_row + 1):
            ws[f'A{row}'].value = ws[f'A{row + 1}'].value
            ws[f'B{row}'].value = ws[f'B{row + 1}'].value
        wb.save(path)

    # print(tree.selection())
    item = tree.selection()[0]
    if len(str(tree.item(item)['values'][0])) >= 6:
        row_n = jijin_code.index(str(tree.item(item)['values'][0])) + 1
    else:
        row_n = jijin_code.index('00' + str(tree.item(item)['values'][0])) + 1
    if len(str(tree.item(item)['values'][0])) >= 6:
        jijin_code.remove(str(tree.item(item)['values'][0]))
    else:
        jijin_code.remove('00' + str(tree.item(item)['values'][0]))
    print(row_n)
    fund_share_list.remove(fund_share_list[row_n - 1])
    tree.delete(item)
    my_thread(deleterows(row_n))


# 双击获取图片详细的功能函数
def get_img(event):
    # 将函数封装加进程可打开多个图片不卡屏
    def img():
        item = tree.selection()[0]
        # print(jijin_code)
        row_n = str(tree.item(item)['values'][0]).rjust(6, '0')
        # print(row_n, item)

        headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36"
        }
        url = f'http://j4.dfcfw.com/charts/pic6/{row_n}.png'
        img_content = requests.get(url, headers).content

        if not os.path.exists('Picture'):
            os.mkdir('Picture')
        with open('Picture/' + row_n + '.png', 'wb') as f:
            f.write(img_content)
        # img = imread('Picture/' + row_n + '.png')
        # imshow(row_n, img)
        # waitKey(0)
        img_root = tk.Toplevel(jijing_root)
        img_root.title(row_n)
        img_root.geometry('420x281')
        img_root.resizable(0, 0)

        paned = tk.PanedWindow(img_root)
        paned.pack()
        path = 'Picture/' + row_n + '.png'
        img = Image.open(path)
        paned.img_file = ImageTk.PhotoImage(img)
        label_img = tk.Label(paned, image=paned.img_file)
        label_img.pack(fill=tk.BOTH, expand=True)
        sleep(5)
        img_root.destroy()

    my_thread(img)


def close():
    # 关闭清理文件
    if os.path.exists('Picture'):
        shutil.rmtree('Picture')
    jijing_root.destroy()


if __name__ == '__main__':
    path = 'jiji.xlsx'
    jijin_code = []
    fund_share_list = []
    fund_share_dir = {}
    gszzl_dir = {}
    gszzl_dir_last = {}
    fund_dir_last = {}
    fund_dir = {}
    fund_chang={}
    data_list=[]

    # 主窗口
    jijing_root = tk.Tk()
    jijing_root.title('基金简易框')
    jijing_root.geometry('360x200')
    jijing_root.resizable(0, 1)

    # 列表框架防止列表变形
    frame = tk.Frame(jijing_root)
    frame.pack(side=tk.TOP, fill=tk.BOTH, expand='yes')

    # 创建列表
    columns = ('num', 'name', 'zxjz', 'gjzf', 'time')
    tree = ttk.Treeview(frame, columns=columns, height=6, selectmode=tk.BROWSE, show='headings')
    # 设置滚动条
    scroll = ttk.Scrollbar(frame)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)
    scroll.config(command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    # 设置列表让其中间对齐
    # tree.column('#0', width=45, anchor='center')
    tree.column('num', width=60, anchor='center')
    tree.column('name', width=60, anchor='center')
    tree.column('zxjz', width=60, anchor='center')
    tree.column('gjzf', width=100, anchor='center')
    tree.column('time', width=60, anchor='center')
    # 设置列表标题文字

    tree.heading('num', text='基金号码')
    tree.heading('name', text='基金名字')
    tree.heading('zxjz', text='估计涨幅')
    tree.heading('gjzf', text='收益预估')
    tree.heading('time', text='更新时间')
    tree.pack(side=tk.LEFT, fill=tk.Y)
    # 增加添加按钮
    button_add = tk.Button(jijing_root, text='添加基金', bg='#24a9ff', fg='white', width=8, height=1,
                           command=lambda: my_thread(add_jijin))
    button_add.pack(side=tk.LEFT, fill=tk.Y, anchor='sw', padx=20, pady=10)
    # 添加脚本运行按钮
    button_run = tk.Button(jijing_root, text='获取数据', bg='#FF0000', fg='white', width=8, height=1,
                           command=lambda: my_thread(jijin_run))
    button_run.pack(side=tk.LEFT, fill=tk.Y, anchor='w', padx=20, pady=10)
    # 添加结束按钮
    tk.Button(jijing_root, text='关闭', width=10, height=1, command=lambda: my_thread(close)).pack(side=tk.LEFT,
                                                                                                 fill=tk.Y,
                                                                                                 anchor='se',
                                                                                                 padx=15, pady=10)

    tree.bind('<Delete>', del_num)
    tree.bind('<Double-Button-1> ', get_img)

    # 读取表格
    try:
        wb = openpyxl.load_workbook(path)
    except FileNotFoundError:
        wb = openpyxl.Workbook()

    ws = wb.active
    sheet = wb['Sheet']
    num_list = list(sheet.rows)
    i = 0

    for num in num_list:
        if num[0].value == None:
            pass
        else:
            jijin_code.append(num[0].value)
            fund_share_list.append(num[1].value)
            fund_share_dir[num[0].value] = num[1].value
        try:
            tree.insert('', i, values=(num[0].value))
        except:
            pass
        i += 1
    # print(jijin_code)
    # print(fund_share_list)
    wb.save(path)

    # 主窗口显示
    jijing_root.mainloop()
