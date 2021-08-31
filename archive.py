import tkinter as tk
from tkinter import ttk
import sqlite3
import pyperclip

import pandas as pd

import os
import platform
import threading
import socket
from datetime import datetime


class AddPopupMenu:
    def copy_selection(self):
        try:
            selection_text = self.selection_get()
        except tk.TclError:
            return
        root.clipboard_clear()
        root.clipboard_append(selection_text)

        pyperclip.copy(selection_text)
        spam = pyperclip.paste()

    def delete_selection(self):
        try:
            self.delete('sel.first', 'sel.last')
        except tk.TclError:
            pass  # Nothing selected

    def cut_selection(self):
        self.copy_selection()
        self.delete_selection()

    def paste_from_clipboard(self):
        try:
            clipboard_text = root.clipboard_get()
        except tk.TclError:
            pass
        else:
            self.delete_selection()
            self.insert(tk.INSERT, clipboard_text)

    def select_all(self):
        self.tag_add(tk.SEL, "1.0", tk.END)
        self.mark_set(tk.INSERT, "1.0")
        self.see(tk.INSERT)

    def show_context_menu(self, event):
        pos_x = self.winfo_rootx() + event.x
        pos_y = self.winfo_rooty() + event.y
        self.popup_menu.tk_popup(pos_x, pos_y)

    def init_menu(self):
        menu = tk.Menu(self, tearoff=False)
        menu.add_command(label="Вырезать", command=self.cut_selection)
        menu.add_command(label="Копировать", command=self.copy_selection)
        menu.add_command(label="Вставить", command=self.paste_from_clipboard)
        menu.add_command(label="Удалить", command=self.delete_selection)
        menu.add_separator()
        menu.add_command(label="Выделить все", command=self.select_all)
        return menu

    def __init__(self, widget_class, *args, **kwargs):
        widget_class.__init__(self, *args, **kwargs)
        self.popup_menu = self.init_menu()
        self.bind("<3>", self.show_context_menu)


class MyText(tk.Text, AddPopupMenu):
    def __init__(self, *args, **kwargs):
        AddPopupMenu.__init__(self, tk.Text, *args, **kwargs)


class MyEntry(tk.Entry, AddPopupMenu):
    def __init__(self, *args, **kwargs):
        AddPopupMenu.__init__(self, tk.Entry, *args, **kwargs)


class Main(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.init_main()
        self.db = db
        self.view_records()

    def init_main(self):
        toolbar = tk.Frame(bg='#d7d8e0', bd=2)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        self.add_img = tk.PhotoImage(file='add.gif')
        btn_open_dialog = tk.Button(toolbar, text='Добавить позицию', command=self.open_dialog, bg='#d7d8e0', bd=0,
                                    compound=tk.TOP, image=self.add_img)
        btn_open_dialog.pack(side=tk.LEFT)

        self.update_img = tk.PhotoImage(file='update.gif')
        btn_edit_dialog = tk.Button(toolbar, text='Редактировать', bg='#d7d8e0', bd=0, image=self.update_img,
                                    compound=tk.TOP, command=self.open_update_dialog)
        btn_edit_dialog.pack(side=tk.LEFT)

        self.delete_img = tk.PhotoImage(file='delete.gif')
        btn_delete = tk.Button(toolbar, text='Удалить позицию', bg='#d7d8e0', bd=0, image=self.delete_img,
                               compound=tk.TOP, command=self.delete_records)
        btn_delete.pack(side=tk.LEFT)

        self.search_img = tk.PhotoImage(file='search.gif')
        btn_search = tk.Button(toolbar, text='Поиск', bg='#d7d8e0', bd=0, image=self.search_img,
                               compound=tk.TOP, command=self.open_search_dialog)
        btn_search.pack(side=tk.LEFT)

        self.refresh_img = tk.PhotoImage(file='refresh.gif')
        btn_refresh = tk.Button(toolbar, text='Обновить', bg='#d7d8e0', bd=0, image=self.refresh_img,
                                compound=tk.TOP, command=self.view_records)
        btn_refresh.pack(side=tk.LEFT)

        btn_flc = tk.Button(toolbar, text='Find', bg='#d7d8e0', bd=0, compound=tk.TOP, command=self.start_file)
        btn_flc.pack(side=tk.RIGHT)

        btn_rel = tk.Button(toolbar, text='Release', bg='#d7d8e0', bd=0, compound=tk.TOP, command=self.release_db)
        btn_rel.pack(side=tk.RIGHT)

        scrollbary = tk.Scrollbar(self, orient='vertical')

        self.tree = ttk.Treeview(self,
                                 columns=('ID', 'Сотрудник', 'Подразделение', 'Кабинет', 'Стац. телефон', 'AD', 'IP'),
                                 height=15, show='headings', yscrollcommand=scrollbary.set)

        scrollbary.config(command=self.tree.yview)
        scrollbary.pack(side=tk.RIGHT, fill="y")

        self.tree.column('ID', width=30, anchor=tk.CENTER)
        self.tree.column('Сотрудник', width=300, anchor=tk.CENTER)
        self.tree.column('Подразделение', width=270, anchor=tk.CENTER)
        self.tree.column('Кабинет', width=60, anchor=tk.CENTER)
        self.tree.column('Стац. телефон', width=90, anchor=tk.CENTER)
        self.tree.column('AD', width=100, anchor=tk.CENTER)
        self.tree.column('IP', width=100, anchor=tk.CENTER)

        self.tree.heading('ID', text='ID')
        self.tree.heading('Сотрудник', text='Сотрудник')
        self.tree.heading('Подразделение', text='Подразделение')
        self.tree.heading('Кабинет', text='Кабинет')
        self.tree.heading('Стац. телефон', text='Стац. телефон')
        self.tree.heading('AD', text='AD')
        self.tree.heading('IP', text='IP')


        self.tree.pack()

        for col in ('ID', 'Сотрудник', 'Подразделение', 'Кабинет', 'Стац. телефон', 'AD', 'IP'):
            self.tree.heading(col, text=col, command=lambda _col=col: self.treeview_sort_column(self.tree, _col, False))

    def release_db(self):
        conn = sqlite3.connect('local_net.db')
        df = pd.read_sql('select * from users', conn)
        df.to_excel(r'result.xlsx', index=False)

    def start_file(self):
        global strin
        strin = []

        def getMyIp():
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)  # Создаем сокет (UDP)
            s.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)  # Настраиваем сокет на BROADCAST вещание.
            s.connect(('<broadcast>', 0))
            return s.getsockname()[0]

        def scan_Ip(ip):
            addr = net + str(ip)
            comm = ping_com + addr
            response = os.popen(comm)
            data = response.readlines()
            name = data[1].split(' ')
            for line in data:
                if 'TTL' in line:
                    response_art = os.popen('arp -a')
                    data_arp = response_art.readlines()
                    for line_arp in data_arp:
                        flag = line_arp.split()

                        if len(flag) > 0 and flag[0] == addr:
                            tmp = (addr + "         Ping Ok" + '            ' + name[3] + '             ' + flag[1])
                            strin.append(tmp)

                    break

        net = getMyIp()
        print('You IP :', net)
        net_split = net.split('.')
        a = '.'
        net = net_split[0] + a + net_split[1] + a + net_split[2] + a
        start_point = int(input("Enter the Starting Number: "))
        end_point = int(input("Enter the Last Number: "))

        oper_sys = platform.system()
        if oper_sys == "Windows":
            ping_com = "ping -n 1 -a "
        else:
            ping_com = "ping -c 1 "

        t1 = datetime.now()
        print("Scanning in Progress:")
        print('IP                 Status               Name                 MAC')
        for ip in range(start_point, end_point):
            if ip == int(net_split[3]):
                continue
            potoc = threading.Thread(target=scan_Ip, args=[ip])
            potoc.start()
        potoc.join()
        t2 = datetime.now()
        total = t2 - t1
        for i in strin:
            print(i)
        print('Find ip :', len(strin))
        print("Scanning completed in: ", total)
        input()

    def treeview_sort_column(self, tv, col, reverse):
        # "сортировка"
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        l.sort(reverse=reverse)

        # rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        # reverse sort next time
        tv.heading(col, command=lambda _col=col: self.treeview_sort_column(tv, _col, not reverse))

    def records(self, MAC, NAME, IP, Domain, Pass_account, MOL, employee, landline_phone, work_phone,
                computer_characteristics, number_inventory, monitor, keyboard, inventory_keyboard, mouse,
                mouse_inventory, AD, local_mail, local_mail_pass, external_mail, external_mail_pass, accounts,
                accounts_pass, FS, FS_pass, subdivision, cabinet, coments):
        self.db.insert_data(MAC, NAME, IP, Domain, Pass_account, MOL, employee, landline_phone, work_phone,
                            computer_characteristics, number_inventory, monitor, keyboard, inventory_keyboard, mouse,
                            mouse_inventory, AD, local_mail, local_mail_pass, external_mail, external_mail_pass,
                            accounts, accounts_pass, FS, FS_pass, subdivision, cabinet, coments)
        self.view_records()

    def update_record(self, ID, MAC, NAME, IP, Domain, Pass_account, MOL, employee, landline_phone, work_phone,
                      computer_characteristics, number_inventory, monitor, keyboard, inventory_keyboard, mouse,
                      mouse_inventory, AD, local_mail, local_mail_pass, external_mail, external_mail_pass,
                      accounts, accounts_pass, FS, FS_pass, subdivision, cabinet, coments):
        self.db.c.execute('''UPDATE users SET ID=?, Сотрудник=?, Подразделение=?, Должность=?, Кабинет=?, 
        Стационарный_телефон=?, Корп_телефон=?, AD =?, Pass_account=?, Локальная_почта=?, Пароль_локал=?, 
        Внешняя_почта=?, Пароль_внешняя=?, FS=?, FS_pass=?, Name=?, IP=?, MAC=?, Domain=?, Учетные_записи=?, 
        Пароль_учетные=?, МОЛ=?, Характеристики_компьютера=?, Инв_номер_компьютера=?, Монитор=?, 
        Инв_номер_монитора=?, Клавиатура=?, Мышь=?, Коментарии=? WHERE ID=?''',
                          (ID, MAC, NAME, IP, Domain, Pass_account, MOL, employee, landline_phone, work_phone,
                           computer_characteristics, number_inventory, monitor, keyboard, inventory_keyboard, mouse,
                           mouse_inventory, AD, local_mail, local_mail_pass, external_mail, external_mail_pass,
                           accounts, accounts_pass, FS, FS_pass, subdivision, cabinet, coments,
                           self.tree.set(self.tree.selection()[0], '#1')))
        self.db.conn.commit()
        self.view_records()

    def view_records(self):
        self.db.c.execute('''SELECT * FROM users''')

        [self.tree.delete(i) for i in self.tree.get_children()]

        [self.tree.insert('', 'end', values=(row[0], row[1], row[2], row[4], row[5], row[7], row[16])) for row in
         self.db.c.fetchall()]


    def delete_records(self):
        for selection_item in self.tree.selection():
            self.db.c.execute('''DELETE FROM users WHERE ID=?''', (self.tree.set(selection_item, '#1'),))
        self.db.conn.commit()
        self.view_records()

    def search_records(self, description):
        description = ('%' + description + '%',)
        self.db.c.execute('''SELECT * FROM users WHERE Сотрудник LIKE ?''', description)
        [self.tree.delete(i) for i in self.tree.get_children()]
        [self.tree.insert('', 'end', values=row) for row in self.db.c.fetchall()]

    def open_dialog(self):
        Child()

    def open_update_dialog(self):
        Update()

    def open_search_dialog(self):
        Search()


class Child(tk.Toplevel):
    def __init__(self):
        super().__init__(root)
        self.init_child()
        self.view = app

    def init_child(self):
        self.title('Создание нового пользователя')
        self.geometry('1220x500+400+300')
        self.resizable(True, True)

        label_ID = tk.Label(self, text='ID :')
        label_ID.place(x=50, y=20)
        label_MAC = tk.Label(self, text='Сотрудник :')
        label_MAC.place(x=50, y=50)
        label_NAME = tk.Label(self, text='Подразделение :')
        label_NAME.place(x=50, y=80)
        label_IP = tk.Label(self, text='Должность :')
        label_IP.place(x=50, y=110)
        label_Domain = tk.Label(self, text='Кабинет :')
        label_Domain.place(x=50, y=140)
        label_Pass_account = tk.Label(self, text='Телефон :')
        label_Pass_account.place(x=50, y=170)
        label_MOL = tk.Label(self, text='Моб. телефон :')
        label_MOL.place(x=50, y=200)
        label_employee = tk.Label(self, text='AD :')
        label_employee.place(x=50, y=230)
        label_landline_phone = tk.Label(self, text='Пароль от AD :')
        label_landline_phone.place(x=50, y=260)
        label_work_phone = tk.Label(self, text='Почта .local :')
        label_work_phone.place(x=50, y=290)
        label_computer_characteristics = tk.Label(self, text='Пароль от .local :')
        label_computer_characteristics.place(x=50, y=320)
        label_number_inventory = tk.Label(self, text='Внешняя почта :')
        label_number_inventory.place(x=50, y=350)
        label_monitor = tk.Label(self, text='Пароль от почты :')
        label_monitor.place(x=50, y=380)
        label_keyboard = tk.Label(self, text='FS :')
        label_keyboard.place(x=50, y=410)
        label_inventory_keyboard = tk.Label(self, text='Пароль от FS :')
        label_inventory_keyboard.place(x=370, y=50)
        label_mouse = tk.Label(self, text='Имя компа :')
        label_mouse.place(x=370, y=80)
        label_mouse_inventory = tk.Label(self, text='IP :')
        label_mouse_inventory.place(x=370, y=110)
        label_AD = tk.Label(self, text='MAC :')
        label_AD.place(x=370, y=140)
        label_local_mail = tk.Label(self, text='Domain :')
        label_local_mail.place(x=370, y=170)
        label_local_mail_pass = tk.Label(self, text='Учетная запись :')
        label_local_mail_pass.place(x=370, y=200)
        label_external_mail = tk.Label(self, text='Пароль от учетной запсии :')
        label_external_mail.place(x=370, y=230)
        label_external_mail_pass = tk.Label(self, text='МОЛ :')
        label_external_mail_pass.place(x=370, y=260)
        label_accounts = tk.Label(self, text='Характеристики :')
        label_accounts.place(x=370, y=290)
        label_accounts_pass = tk.Label(self, text='Инв. № компьютера  :')
        label_accounts_pass.place(x=370, y=320)
        label_FS = tk.Label(self, text='Монитор :')
        label_FS.place(x=370, y=350)
        label_FS_pass = tk.Label(self, text='Инв. № монитора  :')
        label_FS_pass.place(x=370, y=380)
        label_subdivision = tk.Label(self, text='Клавиатура :')
        label_subdivision.place(x=370, y=410)
        label_cabinet = tk.Label(self, text='Мышь :')
        label_cabinet.place(x=735, y=50)
        label_coments = tk.Label(self, text='Коментарии :')
        label_coments.place(x=735, y=80)

        self.entry_ID = MyEntry(self, width=33)
        self.entry_ID.place(x=165, y=20)

        self.entry_MAC = MyEntry(self, width=33)
        self.entry_MAC.place(x=165, y=50)

        self.entry_NAME = MyEntry(self, width=33)
        self.entry_NAME.place(x=165, y=80)

        self.entry_IP = MyEntry(self, width=33)
        self.entry_IP.place(x=165, y=110)

        self.entry_Domain = MyEntry(self, width=33)
        self.entry_Domain.place(x=165, y=140)

        self.entry_Pass_account = MyEntry(self, width=33)
        self.entry_Pass_account.place(x=165, y=170)

        self.entry_MOL = MyEntry(self, width=33)
        self.entry_MOL.place(x=165, y=200)

        self.entry_employee = MyEntry(self, width=33)
        self.entry_employee.place(x=165, y=230)

        self.entry_landline_phone = MyEntry(self, width=33)
        self.entry_landline_phone.place(x=165, y=260)

        self.entry_work_phone = MyEntry(self, width=33)
        self.entry_work_phone.place(x=165, y=290)

        self.entry_computer_characteristics = MyEntry(self, width=33)
        self.entry_computer_characteristics.place(x=165, y=320)

        self.entry_number_inventory = MyEntry(self, width=33)
        self.entry_number_inventory.place(x=165, y=350)

        self.entry_monitor = MyEntry(self, width=33)
        self.entry_monitor.place(x=165, y=380)

        self.entry_keyboard = MyEntry(self, width=33)
        self.entry_keyboard.place(x=165, y=410)

        self.entry_inventory_keyboard = MyEntry(self, width=33)
        self.entry_inventory_keyboard.place(x=530, y=50)

        self.entry_mouse = MyEntry(self, width=33)
        self.entry_mouse.place(x=530, y=80)

        self.entry_mouse_inventory = MyEntry(self, width=33)
        self.entry_mouse_inventory.place(x=530, y=110)

        self.entry_AD = MyEntry(self, width=33)
        self.entry_AD.place(x=530, y=140)

        self.entry_local_mail = MyEntry(self, width=33)
        self.entry_local_mail.place(x=530, y=170)

        self.entry_local_mail_pass = MyEntry(self, width=33)
        self.entry_local_mail_pass.place(x=530, y=200)

        self.entry_external_mail = MyEntry(self, width=33)
        self.entry_external_mail.place(x=530, y=230)

        self.entry_external_mail_pass = MyEntry(self, width=33)
        self.entry_external_mail_pass.place(x=530, y=260)

        self.entry_accounts = MyEntry(self, width=33)
        self.entry_accounts.place(x=530, y=290)

        self.entry_accounts_pass = MyEntry(self, width=33)
        self.entry_accounts_pass.place(x=530, y=320)

        self.entry_FS = MyEntry(self, width=33)
        self.entry_FS.place(x=530, y=350)

        self.entry_FS_pass = MyEntry(self, width=33)
        self.entry_FS_pass.place(x=530, y=380)

        self.entry_subdivision = MyEntry(self, width=33)
        self.entry_subdivision.place(x=530, y=410)

        self.entry_cabinet = MyEntry(self, width=60)
        self.entry_cabinet.place(x=820, y=50)

        self.entry_coments = MyText(self, width=45, height=22)
        self.entry_coments.place(x=820, y=80)

        self.btn_cancel = ttk.Button(self, text='Закрыть', command=self.destroy)
        self.btn_cancel.place(x=1000, y=460)

        self.btn_ok = ttk.Button(self, text='Добавить')
        self.btn_ok.place(x=900, y=460)
        self.btn_ok.bind('<Button-1>', lambda event: self.view.records(self.entry_MAC.get(),
                                                                       self.entry_NAME.get(),
                                                                       self.entry_IP.get(),
                                                                       self.entry_Domain.get(),
                                                                       self.entry_Pass_account.get(),
                                                                       self.entry_MOL.get(),
                                                                       self.entry_employee.get(),
                                                                       self.entry_landline_phone.get(),
                                                                       self.entry_work_phone.get(),
                                                                       self.entry_computer_characteristics.get(),
                                                                       self.entry_number_inventory.get(),
                                                                       self.entry_monitor.get(),
                                                                       self.entry_keyboard.get(),
                                                                       self.entry_inventory_keyboard.get(),
                                                                       self.entry_mouse.get(),
                                                                       self.entry_mouse_inventory.get(),
                                                                       self.entry_AD.get(),
                                                                       self.entry_local_mail.get(),
                                                                       self.entry_local_mail_pass.get(),
                                                                       self.entry_external_mail.get(),
                                                                       self.entry_external_mail_pass.get(),
                                                                       self.entry_accounts.get(),
                                                                       self.entry_accounts_pass.get(),
                                                                       self.entry_FS.get(),
                                                                       self.entry_FS_pass.get(),
                                                                       self.entry_subdivision.get(),
                                                                       self.entry_cabinet.get(),
                                                                       self.entry_coments.get(1.0, 500.0)))

        self.grab_set()
        self.focus_set()


class Update(Child):
    def __init__(self):
        super().__init__()
        self.init_edit()
        self.view = app
        self.db = db
        self.default_data()

    def init_edit(self):
        self.title('Редактировать данные юзера')
        btn_edit = ttk.Button(self, text='Сохранить и закрыть', command=self.destroy)
        btn_edit.place(x=900, y=460)

        btn_edit.bind('<Button-1>', lambda event: self.view.update_record(self.entry_ID.get(),
                                                                          self.entry_MAC.get(),
                                                                          self.entry_NAME.get(),
                                                                          self.entry_IP.get(),
                                                                          self.entry_Domain.get(),
                                                                          self.entry_Pass_account.get(),
                                                                          self.entry_MOL.get(),
                                                                          self.entry_employee.get(),
                                                                          self.entry_landline_phone.get(),
                                                                          self.entry_work_phone.get(),
                                                                          self.entry_computer_characteristics.get(),
                                                                          self.entry_number_inventory.get(),
                                                                          self.entry_monitor.get(),
                                                                          self.entry_keyboard.get(),
                                                                          self.entry_inventory_keyboard.get(),
                                                                          self.entry_mouse.get(),
                                                                          self.entry_mouse_inventory.get(),
                                                                          self.entry_AD.get(),
                                                                          self.entry_local_mail.get(),
                                                                          self.entry_local_mail_pass.get(),
                                                                          self.entry_external_mail.get(),
                                                                          self.entry_external_mail_pass.get(),
                                                                          self.entry_accounts.get(),
                                                                          self.entry_accounts_pass.get(),
                                                                          self.entry_FS.get(),
                                                                          self.entry_FS_pass.get(),
                                                                          self.entry_subdivision.get(),
                                                                          self.entry_cabinet.get(),
                                                                          self.entry_coments.get(1.0, 500.0)))
        self.btn_ok.destroy()
        self.btn_cancel.destroy()

    def default_data(self):
        self.db.c.execute('''SELECT * FROM users WHERE ID=?''',
                          (self.view.tree.set(self.view.tree.selection()[0], '#1'),))
        row = self.db.c.fetchone()

        self.entry_ID.insert(0, row[0])
        self.entry_MAC.insert(0, row[1])
        self.entry_NAME.insert(0, row[2])
        self.entry_IP.insert(0, row[3])
        self.entry_Domain.insert(0, row[4])
        self.entry_Pass_account.insert(0, row[5])
        self.entry_MOL.insert(0, row[6])
        self.entry_employee.insert(0, row[7])
        self.entry_landline_phone.insert(0, row[8])
        self.entry_work_phone.insert(0, row[9])
        self.entry_computer_characteristics.insert(0, row[10])
        self.entry_number_inventory.insert(0, row[11])
        self.entry_monitor.insert(0, row[12])
        self.entry_keyboard.insert(0, row[13])
        self.entry_inventory_keyboard.insert(0, row[14])
        self.entry_mouse.insert(0, row[15])
        self.entry_mouse_inventory.insert(0, row[16])
        self.entry_AD.insert(0, row[17])
        self.entry_local_mail.insert(0, row[18])
        self.entry_local_mail_pass.insert(0, row[19])
        self.entry_external_mail.insert(0, row[20])
        self.entry_external_mail_pass.insert(0, row[21])
        self.entry_accounts.insert(0, row[22])
        self.entry_accounts_pass.insert(0, row[23])
        self.entry_FS.insert(0, row[24])
        self.entry_FS_pass.insert(0, row[25])
        self.entry_subdivision.insert(0, row[26])
        self.entry_cabinet.insert(0, row[27])
        self.entry_coments.insert(1.0, row[28])


        self.entry_ID['state'] = 'disable'


class Search(tk.Toplevel):
    def __init__(self):
        super().__init__()
        self.init_search()
        self.view = app

    def init_search(self):
        self.title('Поиск')
        self.geometry('300x100+400+300')
        self.resizable(False, False)

        label_search = tk.Label(self, text='Поиск')
        label_search.place(x=50, y=20)

        self.entry_search = ttk.Entry(self)
        self.entry_search.place(x=105, y=20, width=150)

        btn_cancel = ttk.Button(self, text='Закрыть', command=self.destroy)
        btn_cancel.place(x=185, y=50)

        btn_search = ttk.Button(self, text='Поиск')
        btn_search.place(x=105, y=50)
        btn_search.bind('<Button-1>', lambda event: self.view.search_records(self.entry_search.get()))
        btn_search.bind('<Button-1>', lambda event: self.destroy(), add='+')


class DB:
    def __init__(self):
        self.conn = sqlite3.connect('local_net.db')
        self.c = self.conn.cursor()
        self.c.execute(
            '''CREATE TABLE IF NOT EXISTS users(
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            Сотрудник TEXT,
            Подразделение TEXT,
            Должность TEXT,
            Кабинет TEXT,
            Стационарный_телефон TEXT,
            Корп_телефон TEXT,
            AD TEXT,
            Pass_account TEXT,
            Локальная_почта TEXT,
            Пароль_локал TEXT,
            Внешняя_почта TEXT,
            Пароль_внешняя TEXT,
            FS TEXT,
            FS_pass TEXT,
            Name TEXT,
            IP TEXT,
            MAC TEXT,
            Domain TEXT,
            Учетные_записи TEXT,
            Пароль_учетные TEXT,
            МОЛ TEXT,
            Характеристики_компьютера TEXT,
            Инв_номер_компьютера TEXT,
            Монитор TEXT,
            Инв_номер_монитора TEXT,
            Клавиатура TEXT,
            Мышь TEXT,
            Коментарии TEXT)''')
        self.conn.commit()

    def insert_data(self, MAC, NAME, IP, Domain, Pass_account, MOL, employee, landline_phone, work_phone,
                    computer_characteristics, number_inventory, monitor, keyboard, inventory_keyboard, mouse,
                    mouse_inventory, AD, local_mail, local_mail_pass, external_mail, external_mail_pass,
                    accounts, accounts_pass, FS, FS_pass, subdivision, cabinet, coments):
        try:
            self.c.execute('''INSERT INTO users VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 
            ?, ?, ?, ?, ?, ?, ?)''',
                           (None, MAC, NAME, IP, Domain, Pass_account, MOL, employee, landline_phone, work_phone,
                            computer_characteristics, number_inventory, monitor, keyboard, inventory_keyboard, mouse,
                            mouse_inventory, AD, local_mail, local_mail_pass, external_mail, external_mail_pass,
                            accounts, accounts_pass, FS, FS_pass, subdivision, cabinet, coments))
            self.conn.commit()
        except sqlite3.IntegrityError:
            root_eror = tk.Tk()
            bt = tk.Button(root_eror, text='Такой уже существует')
            bt.pack()
            root_eror.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    db = DB()
    app = Main(root)
    app.pack()
    root.title("База данных пользователей")
    root.geometry("950x380+300+200")
    root.resizable(False, False)
    root.mainloop()
