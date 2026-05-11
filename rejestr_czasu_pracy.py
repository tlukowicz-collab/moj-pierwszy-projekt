# rejestr_czasu_pracy_PRO_v7.py

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from datetime import datetime
import calendar
import pandas as pd
import pyodbc
import os
import sys
import webbrowser

# --- STAŁE ---
MAIN_COLOR = "#ADD0B3"
CFG_FILE = "config.json"

# --- ŚCIEŻKI ---
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

XLSX_PATH = os.path.join(BASE_DIR, "Pracownicy.xlsx")
MDB_PATH = os.path.join(BASE_DIR, "HXData.mdb")
CFG_PATH = os.path.join(BASE_DIR, CFG_FILE)

# --- POMOC ---
def hhmm(seconds):
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    return f"{h:02d}:{m:02d}"

def calculate_day(times):
    times = sorted(times)
    sec = 0
    for i in range(0, len(times)-1, 2):
        sec += (times[i+1]-times[i]).total_seconds()
    return sec

# Nowa logika statusu na podstawie ostatniego zdarzenia Event
def status_from_events(df_day, live_day=False):

    # brak zdarzeń = NIEOBECNY
    if df_day.empty:
        return "NIEOBECNY"

    df_day = df_day.sort_values("eventtime")

    # DZIEŃ BIEŻĄCY – liczy się ostatnie zdarzenie
    if live_day:
        last_event = df_day.iloc[-1]["Event"]

        if last_event == "Invalid Card":
            return "OBECNY"
        elif last_event == "Entry access":
            return "NIEOBECNY"
        else:
            return "NIEOBECNY"

    # DZIEŃ HISTORYCZNY – miał przynajmniej jedno odbicie
    return "OBECNY"

# --- DANE ---
def load_employees():
    df = pd.read_excel(XLSX_PATH, dtype=str)
    df.columns = df.columns.str.lower().str.strip()
    df["display"] = df["nazwisko"] + " " + df["imię"]
    return df

def load_events():
    if not os.path.exists(MDB_PATH):
        messagebox.showerror("Błąd", f"Nie znaleziono bazy:\n{MDB_PATH}")
        return pd.DataFrame(columns=["CardNo","EventTime","Event","eventtime","date"])
    try:
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={MDB_PATH};'
        )
        conn = pyodbc.connect(conn_str)
        df = pd.read_sql("SELECT CardNo, EventTime, Event FROM VKQCardRecord", conn)
        conn.close()
        df["eventtime"] = pd.to_datetime(df["EventTime"])
        df["date"] = df["eventtime"].dt.date
        return df.sort_values("eventtime")
    except Exception as e:
        messagebox.showerror("Błąd bazy danych", str(e))
        return pd.DataFrame(columns=["CardNo","EventTime","Event","eventtime","date"])

# --- APP ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Rejestr Czasu Pracy PRO v7.0")
        self.root.geometry("1400x850")
        self.root.configure(bg=MAIN_COLOR)

        self.employees = load_employees()
        self.events = load_events()

        self.style = ttk.Style()
        self.style.theme_use("default")
        self.style.configure("Treeview", font=("Segoe UI",10), rowheight=26)
        self.style.configure("Treeview.Heading", font=("Segoe UI",10,"bold"))

        self.start()

    def clear(self):
        for w in self.root.winfo_children():
            w.destroy()

    def tree_with_scroll(self):
        frame = tk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=15, pady=10)

        scroll_y = ttk.Scrollbar(frame, orient="vertical")
        scroll_x = ttk.Scrollbar(frame, orient="horizontal")

        tree = ttk.Treeview(
            frame,
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set
        )

        scroll_y.config(command=tree.yview)
        scroll_x.config(command=tree.xview)

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)

        def treeview_sort_column(tv, col, reverse):
            data_list = [(tv.set(k, col), k) for k in tv.get_children('')]
            try:
                data_list.sort(key=lambda t: float(t[0].replace("%","").replace(":","")), reverse=reverse)
            except ValueError:
                data_list.sort(key=lambda t: t[0], reverse=reverse)
            for index, (val, k) in enumerate(data_list):
                tv.move(k, '', index)
            tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

        tree.sort_column = treeview_sort_column
        return tree

    # === DASHBOARD ===
    def start(self):
        self.clear()
        tk.Label(self.root,text="REJESTR CZASU PRACY – DASHBOARD",
                 font=("Segoe UI",22,"bold"),
                 bg=MAIN_COLOR).pack(pady=15)

        top = tk.Frame(self.root,bg=MAIN_COLOR)
        top.pack()

        tk.Label(top,text="Data:",bg=MAIN_COLOR,font=("Segoe UI",11)).grid(row=0,column=0,padx=5)
        self.cal = DateEntry(top,date_pattern="yyyy-mm-dd")
        self.cal.grid(row=0,column=1,padx=5)

        tk.Button(top,text="Odśwież",command=self.refresh_dashboard,width=15).grid(row=0,column=2,padx=10)

        self.kpi = tk.Label(self.root,font=("Segoe UI",13,"bold"),bg=MAIN_COLOR)
        self.kpi.pack(pady=8)

        self.tree = self.tree_with_scroll()
        self.tree["columns"]=["Dział","Wszyscy","Obecni","Nieobecni","%"]

        for c in self.tree["columns"]:
            self.tree.heading(c, text=c,
                              command=lambda _c=c: self.tree.sort_column(self.tree, _c, False))
            self.tree.column(c,anchor="center", width=130)

        menu = tk.Frame(self.root,bg=MAIN_COLOR)
        menu.pack(pady=10)

        tk.Button(menu,text="Przegląd dnia",width=25,command=self.day_view).grid(row=0,column=0,padx=5)
        tk.Button(menu,text="Przegląd pracownika",width=25,command=self.employee_view).grid(row=0,column=1,padx=5)
        tk.Button(menu,text="Raport miesięczny – WSZYSCY",width=25,command=self.all_month_view).grid(row=0,column=2,padx=5)

        self.footer()
        self.refresh_dashboard()

    def footer(self):
        f=tk.Frame(self.root,bg=MAIN_COLOR)
        f.pack(side="bottom",pady=8)
        tk.Label(f,text="Copyright © 2026 by Tomasz Łukowicz",
                 font=("Segoe UI",10,"bold"),bg=MAIN_COLOR).pack()
        tk.Label(f,text="TLDesign",bg=MAIN_COLOR,font=("Segoe UI",10)).pack()
        email = tk.Label(f, text="tlukowicz.projekt@gmail.com",
                         font=("Segoe UI", 10, "underline"), fg="blue", bg=MAIN_COLOR, cursor="hand2")
        email.pack()
        email.bind("<Button-1>", lambda e: webbrowser.open("mailto:tlukowicz.projekt@gmail.com"))

    def refresh_dashboard(self):
        self.events = load_events()
        for i in self.tree.get_children(): self.tree.delete(i)

        date = self.cal.get_date()
        today = datetime.now().date()
        is_live = (date == today)

        total = len(self.employees)
        present=0
        stats={}

        for _,emp in self.employees.iterrows():
            card=emp["nr karty"]
            dzial=emp.get("dział","Brak")
            df=self.events[(self.events["CardNo"]==card)&(self.events["date"]==date)]
            st = status_from_events(df, live_day=is_live)
            stats.setdefault(dzial,{"all":0,"ok":0})
            stats[dzial]["all"]+=1
            if st=="OBECNY":
                present+=1
                stats[dzial]["ok"]+=1

        absent=total-present
        perc=round((present/total)*100,1) if total else 0
        self.kpi.config(text=f"👥 Obecni: {present}   ❌ Nieobecni: {absent}   📊 Frekwencja: {perc}%")

        for dzial,d in stats.items():
            nieob=d["all"]-d["ok"]
            p=round((d["ok"]/d["all"])*100,1) if d["all"] else 0
            self.tree.insert("", "end", values=[dzial,d["all"],d["ok"],nieob,f"{p}%"])

    # === PRZEGLĄD DNIA ===
    def day_view(self):
        self.clear()
        tk.Button(self.root,text="← Powrót",command=self.start).pack(anchor="w",padx=10,pady=5)

        bar=tk.Frame(self.root,bg=MAIN_COLOR)
        bar.pack()

        tk.Label(bar,text="Dział:",bg=MAIN_COLOR).grid(row=0,column=0,padx=5)
        dzialy=["Wszyscy"]+sorted(self.employees["dział"].dropna().unique())
        self.combo_d=ttk.Combobox(bar,values=dzialy,width=25)
        self.combo_d.set("Wszyscy")
        self.combo_d.grid(row=0,column=1,padx=5)

        tk.Label(bar,text="Data:",bg=MAIN_COLOR).grid(row=0,column=2,padx=5)
        self.cal_d=DateEntry(bar,date_pattern="yyyy-mm-dd")
        self.cal_d.grid(row=0,column=3,padx=5)

        tk.Button(bar,text="Pokaż",command=self.show_day,width=15).grid(row=0,column=4,padx=10)

        self.tree=self.tree_with_scroll()
        self.tree["columns"]=["Dział","Card","Nazwisko","Imię","Czas","Status"]
        for c in self.tree["columns"]:
            self.tree.heading(c, text=c, command=lambda _c=c: self.tree.sort_column(self.tree,_c,False))

        tk.Button(self.root,text="Export Excel",command=self.export_current,width=20).pack(pady=5)

        self.show_day()

    def show_day(self):
        for i in self.tree.get_children(): self.tree.delete(i)

        date=self.cal_d.get_date()
        dzial=self.combo_d.get()

        df=self.events[self.events["date"]==date]

        if dzial!="Wszyscy":
            cards=self.employees[self.employees["dział"]==dzial]["nr karty"]
            df=df[df["CardNo"].isin(cards)]

        df=df.merge(self.employees,left_on="CardNo",right_on="nr karty",how="left")
        df=df.sort_values("eventtime")

        for _,r in df.iterrows():
            if r["Event"]=="Invalid Card":
                status_event="WEJŚCIE"
            elif r["Event"]=="Entry access":
                status_event="WYJŚCIE"
            else:
                status_event=r["Event"]

            self.tree.insert("", "end",
                             values=[r.get("dział",""), r["CardNo"], r.get("nazwisko",""), r.get("imię",""),
                                     r["eventtime"].strftime("%H:%M:%S"), status_event])

    # === PRZEGLĄD PRACOWNIKA / RAPORT MIESIĘCZNY ===
    def employee_view(self):
        self.clear()
        tk.Button(self.root,text="← Powrót",command=self.start).pack(anchor="w",padx=10,pady=5)

        bar=tk.Frame(self.root,bg=MAIN_COLOR)
        bar.pack()

        self.combo_e=ttk.Combobox(bar,values=["Wszyscy"]+list(self.employees["display"]),width=35)
        self.combo_e.set("Wszyscy")
        self.combo_e.grid(row=0,column=0,padx=5)

        tk.Label(bar,text="Miesiąc:",bg=MAIN_COLOR).grid(row=0,column=1,padx=5)
        self.combo_m=ttk.Combobox(bar,values=list(range(1,13)),width=5)
        self.combo_m.set(datetime.now().month)
        self.combo_m.grid(row=0,column=2)

        tk.Label(bar,text="Rok:",bg=MAIN_COLOR).grid(row=0,column=3,padx=5)
        self.combo_y=ttk.Combobox(bar,values=list(range(2023,2031)),width=7)
        self.combo_y.set(datetime.now().year)
        self.combo_y.grid(row=0,column=4)

        tk.Button(bar,text="Pokaż",command=self.show_employee,width=15).grid(row=0,column=5,padx=10)

        self.tree=self.tree_with_scroll()
        tk.Button(self.root,text="Export Excel",command=self.export_current,width=20).pack(pady=5)

    def show_employee(self):
        for i in self.tree.get_children(): self.tree.delete(i)

        month=int(self.combo_m.get())
        year=int(self.combo_y.get())
        days=calendar.monthrange(year,month)[1]
        sel=self.combo_e.get()

        if sel=="Wszyscy":
            cols=["Dział","Nazwisko","Imię","Data","Start","Koniec","Suma","Status"]
            self.tree["columns"]=cols
            for c in cols: self.tree.heading(c,text=c, command=lambda _c=c:self.tree.sort_column(self.tree,_c,False))

            for _,emp in self.employees.iterrows():
                for d in range(1,days+1):
                    date=datetime(year,month,d).date()
                    df=self.events[(self.events["CardNo"]==emp["nr karty"]) & (self.events["date"]==date)]
                    times=list(df["eventtime"])
                    sec=calculate_day(times)
                    st=status_from_events(df, live_day=False)

                    start=times[0].strftime("%H:%M") if times else ""
                    end=times[-1].strftime("%H:%M") if len(times)>1 else start

                    self.tree.insert("", "end",
                        values=[emp.get("dział",""), emp["nazwisko"], emp["imię"], date.strftime("%Y-%m-%d"),
                                start, end, hhmm(sec), st])
        else:
            emp=self.employees[self.employees["display"]==sel].iloc[0]
            cols=["Data","Start","Koniec","Suma","Odbicia","Status"]
            self.tree["columns"]=cols
            for c in cols: self.tree.heading(c,text=c, command=lambda _c=c:self.tree.sort_column(self.tree,_c,False))

            total_sec=0
            ok_days=0

            for d in range(1,days+1):
                date=datetime(year,month,d).date()
                df=self.events[(self.events["CardNo"]==emp["nr karty"]) & (self.events["date"]==date)]
                times=list(df["eventtime"])
                sec=calculate_day(times)
                st=status_from_events(df, live_day=False)

                start=times[0].strftime("%H:%M") if times else ""
                end=times[-1].strftime("%H:%M") if len(times)>1 else start

                if len(times)>=2 and len(times)%2==0:
                    total_sec+=sec
                    ok_days+=1

                self.tree.insert("", "end",
                                 values=[date, start, end, hhmm(sec), len(times), st],
                                 tags=("ok",) if len(times)>=2 and len(times)%2==0 else ("err",))

            self.tree.tag_configure("ok", background="#d4f5d4")
            self.tree.tag_configure("err", background="#f5d4d4")

            avg_sec = total_sec // ok_days if ok_days else 0
            self.tree.insert("", "end", values=["SUMA / ŚREDNIA", "", "", hhmm(total_sec), "", ""])
            self.tree.insert("", "end", values=["ŚREDNIA dziennie (OBECNI)", "", "", hhmm(avg_sec), "", ""])

    # === EXPORT ===
    def export_current(self):
        rows=[]
        cols=self.tree["columns"]
        for iid in self.tree.get_children():
            rows.append(self.tree.item(iid)["values"])
        if not rows:
            messagebox.showwarning("Export","Brak danych")
            return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx")
        if not path: return
        pd.DataFrame(rows,columns=cols).to_excel(path,index=False)
        messagebox.showinfo("Export","Zapisano ✔")

    # === RAPORT WSZYSCY ===
    def all_month_view(self):
        self.employee_view()

# === START ===
if __name__=="__main__":
    root=tk.Tk()
    app=App(root)
    root.mainloop()