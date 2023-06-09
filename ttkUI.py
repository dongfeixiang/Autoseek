import ttkbootstrap as ttk
from ttkbootstrap.dialogs import Dialog,Messagebox,Querybox
from ttkbootstrap.scrolled import ScrolledFrame
import calendar
from datetime import datetime,date,timedelta
from datamodel import *

class DateTile(ttk.Frame):
    '''日期块'''
    def __init__(self, master=None, text:str="", **kargs):
        super().__init__(master=master, padding=2, bootstyle="light", relief="groove")
        self.day = ttk.StringVar(value=text)
        self.Rstate = "U"
        self.Wstate = "U"
        self._setUI()
        self.configState(self.Rstate, self.Wstate)
    def _setUI(self):
        self.recordcheck = ttk.Label(self, text="R")
        self.recordcheck.grid(row=0, column=0)
        self.wtimecheck = ttk.Label(self, text="T")
        self.wtimecheck.grid(row=0, column=1)
        self.label = ttk.Label(self, textvariable=self.day, width=4)
        self.label.grid(row=1, column=0, columnspan=2)
    def configLabel(self, text:str, bootstyle="default"):
        self.day.set(text)
        self.label.config(bootstyle=bootstyle)
    def configState(self, rstate:str, wstate:str):
        style = ttk.Style()
        style.configure("U.TLabel", font=("Arial",5), width=3, relief="groove", background="white", foreground="#ADB5BD", bordercolor="#ADB5BD")
        style.configure("M.TLabel", font=("Arial",5), width=3, background="#A0F0A0")
        style.configure("C.TLabel", font=("Arial",5), width=3, background="#02B875")
        self.recordcheck.config(text=("" if rstate=="C" else "R"), style=f"{rstate}.TLabel")
        self.wtimecheck.config(text=("" if wstate=="C" else "T"), style=f"{wstate}.TLabel")

class TkCalender(ttk.Frame):
    '''自定义日历'''
    def __init__(self, master=None, **kargs):
        super().__init__(master=master, padding=10)
        self.year = ttk.IntVar(self, datetime.today().year)
        self.month = ttk.IntVar(self, datetime.today().month)
        self.Rstatedict = {}
        self.Wstatedict = {}
        self._setUI()        # 设计UI
        self._setEvent()     # 绑定事件
        self.event_generate("<<change>>")
    def _setUI(self):
        self.header = ttk.Frame(self)
        self.preYearBtn = ttk.Button(self.header, text="<<")
        self.nextYearBtn = ttk.Button(self.header, text=">>")
        self.preMonthBtn = ttk.Button(self.header, text="<")
        self.nextMonthBtn = ttk.Button(self.header, text=">")
        self.yearInput = ttk.Entry(self.header, width=6, textvariable=self.year, justify="center", state="disable")
        self.monthInput = ttk.Entry(self.header, width=6, textvariable=self.month, justify="center", state="disable")
        self.preYearBtn.pack(side="left")
        self.preMonthBtn.pack(side="left")
        self.yearInput.pack(side="left")
        ttk.Label(self.header, text="年").pack(side="left")
        self.monthInput.pack(side="left")
        ttk.Label(self.header, text="月").pack(side="left")
        self.nextMonthBtn.pack(side="left")
        self.nextYearBtn.pack(side="left")
        self.header.pack()
        self.body = ttk.Frame(self)
        # 日历UI
        cols = ["日","一","二","三","四","五","六"]
        self.titles = [ttk.Label(self.body, text=cols[i], padding=14) for i in range(7)]
        self.weeks = [[] for i in range(6)]
        for i in range(7):
            for j in range(7):
                if i==0:
                    self.titles[j].grid(row=i, column=j)
                else:
                    self.weeks[i-1].append(DateTile(self.body))
        self.body.pack()
    def _setEvent(self):
        self.preYearBtn.config(command=self._preYear)
        self.nextYearBtn.config(command=self._nextYear)
        self.preMonthBtn.config(command=self._preMonth)
        self.nextMonthBtn.config(command=self._nextMonth)
        self.bind("<<change>>", self._update)
        self.bind("<<state-change>>", self._updatestate)
    def _preYear(self):
        if self.year.get() > 1900:
            self.year.set(self.year.get()-1)
            self.event_generate("<<change>>")
            self.event_generate("<<state-change>>")
    def _nextYear(self):
        if self.year.get() < datetime.today().year:
            self.year.set(self.year.get()+1)
            self.event_generate("<<change>>")
            self.event_generate("<<state-change>>")
    def _preMonth(self):
        firstday = date(self.year.get(), self.month.get(), 1)
        lastday = firstday-timedelta(days=1)
        if lastday.year >= 1900:
            self.year.set(lastday.year)
            self.month.set(lastday.month)
            self.event_generate("<<change>>")
            self.event_generate("<<state-change>>")
    def _nextMonth(self):
        lastday = date(self.year.get(), self.month.get(), 1)
        nextday = lastday+timedelta(days=calendar.monthrange(self.year.get(), self.month.get())[1]+1)
        if nextday.year <= datetime.today().year:
            self.year.set(nextday.year)
            self.month.set(nextday.month)
            self.event_generate("<<change>>")
            self.event_generate("<<state-change>>")
    def _update(self, event):
        cal = calendar.Calendar(6)
        weeks = cal.monthdayscalendar(self.year.get(),self.month.get())
        for i in range(6):
            for j in range(7):
                self.weeks[i][j].configLabel(text="")
                if i < len(weeks) and weeks[i][j]:
                    cur = date(self.year.get(), self.month.get(), weeks[i][j])      # 当前日期
                    if cur == datetime.today().date():
                        self.weeks[i][j].configLabel(text=weeks[i][j], bootstyle="primary-inverse")
                    elif cur.isoweekday()==6 or cur.isoweekday()==7:
                        self.weeks[i][j].configLabel(text=weeks[i][j], bootstyle="danger")
                    else:
                        self.weeks[i][j].configLabel(text=weeks[i][j])
                    self.weeks[i][j].grid(row=i+1, column=j)
                else:
                    self.weeks[i][j].grid_forget()
    def _updatestate(self, event):
        for i in range(6):
            for j in range(7):
                if self.weeks[i][j].day.get():
                    today = date(self.year.get(), self.month.get(), int(self.weeks[i][j].day.get()))
                    today = today.strftime("%Y-%m-%d")      # 当前日期
                    rstate = self.Rstatedict[today] if today in self.Rstatedict else "U"
                    wstate = self.Wstatedict[today] if today in self.Wstatedict else "U"
                    self.weeks[i][j].configState(rstate, wstate)

class UserWindow(Dialog):
    '''用户设置界面'''
    def __init__(self, master=None, user:User=None, buffers:dict={}):
        super().__init__(master, title="用户设置")
        self.user = user        # 私有对象实例
        self.name = ttk.StringVar(value=user.name)
        self.notebookID = ttk.StringVar(value=user.notebookID)
        self.notebookPWD = ttk.StringVar(value=user.notebookPWD)
        self.worktimeID = ttk.StringVar(value=user.worktimeID)
        self.worktimePWD = ttk.StringVar(value=user.worktimePWD)
        self.srcpath = ttk.StringVar(value=user.srcpath)
        self.chrome = ttk.StringVar(value=user.google_path)
        self.buffers = buffers
        self.batch = {
            "proteinA":ttk.StringVar(value=user.batch["proteinA"]),
            "proteinA_AKTA":ttk.StringVar(value=user.batch["proteinA_AKTA"]),
            "Ni":ttk.StringVar(value=user.batch["Ni"]),
            "Ni_AKTA":ttk.StringVar(value=user.batch["Ni_AKTA"])}
        self.switch = ttk.BooleanVar(value=user.switch)
    def create_body(self, master):
        self.body = ttk.Frame(master, padding=10)
        ttk.Label(self.body, text="姓  名").grid(row=0, column=0)
        ttk.Entry(self.body, textvariable=self.name, width=6).grid(row=0, column=1, sticky="W")
        ttk.Label(self.body, text="电子记录账号").grid(row=1, column=0)
        ttk.Entry(self.body, textvariable=self.notebookID).grid(row=1, column=1, columnspan=3, pady=2, sticky="W")
        ttk.Label(self.body, text="电子记录密码").grid(row=2, column=0)
        ttk.Entry(self.body, textvariable=self.notebookPWD, show="*").grid(row=2, column=1, columnspan=3, sticky="W")
        ttk.Label(self.body, text="工时系统账号").grid(row=3, column=0)
        ttk.Entry(self.body, textvariable=self.worktimeID).grid(row=3, column=1, columnspan=3, pady=2, sticky="W")
        ttk.Label(self.body, text="工时系统密码").grid(row=4, column=0)
        ttk.Entry(self.body, textvariable=self.worktimePWD, show="*").grid(row=4, column=1, columnspan=3, sticky="W")
        ttk.Label(self.body, text="数据文件路径").grid(row=5, column=0)
        ttk.Entry(self.body, textvariable=self.srcpath).grid(row=5, column=1, columnspan=3, pady=2, sticky="W")
        ttk.Label(self.body, text="Chrome路径").grid(row=6, column=0)
        ttk.Entry(self.body, textvariable=self.chrome).grid(row=6, column=1, columnspan=3, sticky="W")
        ttk.Label(self.body, text="Buffer批号").grid(row=7, column=0)
        self.bufferbox = ttk.Combobox(self.body, values=list(self.buffers.values()), width=14)
        self.bufferbox.grid(row=7, column=1, columnspan=2, pady=2, sticky="W")
        ttk.Button(self.body, text="+", bootstyle="dark", command=self.addbuffer).grid(row=7, column=3)
        self.batchframe = ttk.Labelframe(self.body, text="柱子批号", padding=5, bootstyle="dark")
        ttk.Label(self.batchframe, text="Protein A重力柱").grid(row=0, column=0, sticky="W")
        ttk.Entry(self.batchframe, textvariable=self.batch["proteinA"], width=15).grid(row=1, column=0, padx=2)
        ttk.Label(self.batchframe, text="Protein A AKTA").grid(row=0, column=1, sticky="W")
        ttk.Entry(self.batchframe, textvariable=self.batch["proteinA_AKTA"], width=15).grid(row=1, column=1, padx=2)
        ttk.Label(self.batchframe, text="Ni2+重力柱").grid(row=2, column=0, sticky="W")
        ttk.Entry(self.batchframe, textvariable=self.batch["Ni"], width=15).grid(row=3, column=0, padx=2)
        ttk.Label(self.batchframe, text="Ni2+ AKTA").grid(row=2, column=1, sticky="W")
        ttk.Entry(self.batchframe, textvariable=self.batch["Ni_AKTA"], width=15).grid(row=3, column=1, padx=2)
        self.batchframe.grid(row=8, column=0, columnspan=4)
        ttk.Label(self.body, text="调  试").grid(row=0, column=2)
        ttk.Checkbutton(self.body, variable=self.switch, bootstyle="dark-round-toggle").grid(row=0, column=3, sticky="W")
        self.body.pack()
    def create_buttonbox(self, master):
        self.footer = ttk.Frame(master, padding=(10,0,10,10))
        self.enterbtn = ttk.Button(self.footer, text="确 定", width=8, bootstyle="dark", command=self.save)
        self.enterbtn.pack(side="left", padx=30)
        self.cancelbtn = ttk.Button(self.footer, text="取 消", width=8, bootstyle="dark", command=self._toplevel.destroy)
        self.cancelbtn.pack(side="left", padx=30)
        self.footer.pack()
    def addbuffer(self):
        newbatch = self.bufferbox.get()
        if newbatch:
            thisweek = datetime.strptime(newbatch[-8:], "%Y%m%d")+timedelta(days=4)
            weeknum = thisweek.strftime("%Y-%W")
            if weeknum not in self.buffers:
                self.buffers[weeknum] = newbatch
            else:
                res = Messagebox.show_question(f"{weeknum}周已存在批号，确定是否修改？", parent=self.master)
                if res == "Yes":
                    self.buffers[weeknum] = newbatch
            self.bufferbox.config(values=list(self.buffers.values()))
            self.bufferbox.set("")
    def save(self):
        batch = {}
        for key in self.batch:
            batch[key] = self.batch[key].get()
        self.user = User(self.name.get(), self.notebookID.get(), self.notebookPWD.get(), self.worktimeID.get(), self.worktimePWD.get(),
        self.srcpath.get(), batch, self.switch.get(), self.chrome.get())
        self._toplevel.destroy()

class MeetingWindow(Dialog):
    def __init__(self, master=None):
        super().__init__(master, "会议设置")
        self.meeting = None
        self.name = ttk.StringVar()
        self.lasttime = ttk.IntVar(value=1)
        self.frequency = ttk.StringVar(value="single")
    def create_body(self, master):
        row1 = ttk.Frame(master, padding=(10,10,10,0))
        ttk.Label(row1, text="会议名称：").grid(row=0, column=0)
        ttk.Entry(row1, textvariable=self.name, width=14).grid(row=0, column=1, columnspan=2)
        row1.pack(fill="x")
        row2 = ttk.Frame(master, padding=(10,3,10,3))
        ttk.Label(row2, text="会议时长：").pack(side="left")
        ttk.Spinbox(row2, from_=1, to=24, textvariable=self.lasttime, width=2).pack(side="left")
        ttk.Label(row2, text="小时").pack(side="left")
        row2.pack(fill="x")
        self.freq = ttk.Labelframe(master, text="周期", padding=10, bootstyle = "dark")
        self.singleRadio= ttk.Radiobutton(self.freq, text="单次", value="single", variable=self.frequency)
        self.singleRadio.pack(side="left")
        self.dailyRadio = ttk.Radiobutton(self.freq, text="每日", value="daily", variable=self.frequency, state="disable")
        self.dailyRadio.pack(side="left", padx=10)
        self.weeklyRadio = ttk.Radiobutton(self.freq, text="每周", value="weekly", variable=self.frequency, state="disable")
        self.weeklyRadio.pack(side="left")
        self.freq.pack(fill="x", padx=10)
    def create_buttonbox(self, master):
        row3 = ttk.Frame(master, padding=10)
        ttk.Button(row3, text="确 定", bootstyle="dark", command=self.save).pack(side="left", padx=20)
        ttk.Button(row3, text="取 消", bootstyle="dark", command=self._toplevel.destroy).pack(side="left", padx=20)
        row3.pack()
    def save(self):
        name = self.name.get() if self.name.get() else "会议"
        self.meeting = Meeting(name=name, lasttime=self.lasttime.get(), frequency=self.frequency.get())
        self._toplevel.destroy()

class PlanLabel(ttk.Label):
    def __init__(self, master=None, plan:Plan=None):
        super().__init__(master=master, padding=3)
        if isinstance(plan, Routine):
            self.config(text="平台事务", width=10, bootstyle="primary-inverse")
        elif isinstance(plan, Meeting):
            self.config(text=f"{plan.name} {plan.lasttime}h", width=10, bootstyle="info-inverse")
        elif isinstance(plan, Project):
            self.config(text=plan.name, width=10, bootstyle="warning-inverse")

class Planlist(ScrolledFrame):
    def __init__(self, master=None, plans:list[Plan]=None, **kargs):
        super().__init__(master=master, autohide=True, **kargs)
        self.plans = plans.copy()
        self._setUI()
    def _setUI(self):
        self.labellist = []
        for plan in self.plans:
            self.labellist.append(PlanLabel(self, plan))
            self.labellist[-1].pack(pady=1)
    def add(self, plan:Plan) -> PlanLabel:
        planlabel = PlanLabel(self, plan)
        self.labellist.append(planlabel)
        self.labellist[-1].pack(pady=1)
        self.plans.append(plan)
        return planlabel
    def delete(self, planlabel:PlanLabel):
        del self.plans[self.labellist.index(planlabel)]
        self.labellist.remove(planlabel)
        planlabel.destroy()

class WTimeWindow(ttk.Toplevel):
    '''手动设置界面'''
    def __init__(self, master=None, date:str="", schedule:Schedule=None, **kargs):
        super().__init__(master=master, title="日程安排",transient=master, resizable=(0, 0), **kargs)
        self.schedule = schedule
        self.projects = schedule.projects
        self.date = ttk.StringVar(self, date)
        self.opentime_h = ttk.StringVar(self, "%02d"%schedule.opentime.hour)
        self.opentime_m = ttk.StringVar(self, "%02d"%schedule.opentime.minute)
        self.closetime_h = ttk.StringVar(self, "%02d"%schedule.closetime.hour)
        self.closetime_m = ttk.StringVar(self, "%02d"%schedule.closetime.minute)
        self._setUI()
        self._setEvent()
        self.grab_set()
        self.wait_window()
    def _setUI(self):
        self.header = ttk.Frame(self, padding=(10,10,10,0))
        ttk.Label(self.header, textvariable=self.date, padding=3, bootstyle="secondary-inverse").grid(row=0, column=0, columnspan=4)
        ttk.Label(self.header, text="上班").grid(row=1, column=0)
        ttk.Spinbox(self.header, from_=0, to=23, format="%02.0f", textvariable=self.opentime_h, width=2).grid(row=1, column=1, pady=3)
        ttk.Label(self.header, text=":").grid(row=1, column=2)
        ttk.Spinbox(self.header, from_=0, to=59, format="%02.0f", textvariable=self.opentime_m, width=2).grid(row=1, column=3)
        ttk.Label(self.header, text="下班").grid(row=2, column=0)
        ttk.Spinbox(self.header, from_=0, to=23, format="%02.0f", textvariable=self.closetime_h, width=2).grid(row=2, column=1)
        ttk.Label(self.header, text=":").grid(row=2, column=2)
        ttk.Spinbox(self.header, from_=0, to=59, format="%02.0f", textvariable=self.closetime_m, width=2).grid(row=2, column=3)
        self.header.pack()
        self.body = ttk.Frame(self, padding=5, relief="groove", borderwidth=1)
        self.panel = Planlist(self.body, self.schedule.plans, width=120, height=200)
        self.panel.pack_propagate(False)
        self.panel.pack(side="left")
        ttk.Separator(self.body, orient="vertical", bootstyle="secondary").pack(side="left", fill="y")
        self.mode = ttk.Frame(self.body, width=60, height=200)
        self.routineIcon = ttk.Label(self.mode, text="平台", width=5, bootstyle="primary-inverse")
        self.routineIcon.pack(padx=2, pady=2)
        self.projectIcon = ttk.Label(self.mode, text="项目", width=5, bootstyle="warning-inverse")
        self.projectIcon.pack(padx=2)
        self.meetingIcon = ttk.Label(self.mode, text="会议", width=5, bootstyle="info-inverse")
        self.meetingIcon.pack(padx=2, pady=2)
        self.mode.pack_propagate(False)
        self.mode.pack(side="left", fill="both")
        self.body.pack(padx=10, pady=10)
        self.footer = ttk.Frame(self, padding=(10,0,10,10))
        self.enterBtn = ttk.Button(self.footer, text="确 定", bootstyle="dark", command=self.save)
        self.enterBtn.pack(side="left", padx=10)
        self.cancelBtn = ttk.Button(self.footer, text="取 消", bootstyle="dark", command=self.destroy)
        self.cancelBtn.pack(side="left", padx=10)
        self.footer.pack()
    def _setEvent(self):
        self.routineIcon.bind("<Button-1>", self.add)
        self.projectIcon.bind("<Button-1>", self.add)
        self.meetingIcon.bind("<Button-1>", self.add)
        for label in self.panel.labellist:
            label.bind("<Button-3>", self.popmenu)
    def add(self, event):
        if event.widget["text"] == "平台":
            if len(self.projects) > 0:
                Messagebox.show_error("存在项目任务，禁止再添加平台事务", title="错误", parent=self)
            else:
                self.panel.add(Routine()).bind("<Button-3>", self.popmenu)
        elif event.widget["text"] == "项目":
            name = Querybox.get_string("项目号：", parent=self)
            if name:
                project = Project(name=name)
                self.projects.append(project)
                self.panel.add(project).bind("<Button-3>", self.popmenu)
        elif event.widget["text"] == "会议":
            window = MeetingWindow(self)
            window.show((event.x_root, event.y_root-70))
            meeting = window.meeting
            if meeting:
                self.panel.add(meeting).bind("<Button-3>", self.popmenu)
    def popmenu(self, e):
        rightmenu = ttk.Menu(self, tearoff=False)
        rightmenu.add_command(label="删除", command=lambda:self.panel.delete(e.widget))
        rightmenu.post(e.x_root,e.y_root)
    def save(self):
        today = datetime.strptime(self.date.get(), "%Y-%m-%d")
        opentime = f"{self.date.get()} {self.opentime_h.get()}:{self.opentime_m.get()}"
        closetime = f"{self.date.get()} {self.closetime_h.get()}:{self.closetime_m.get()}"
        state = "M" if len(self.panel.plans) else "U"
        self.schedule = Schedule(today, datetime.strptime(opentime, "%Y-%m-%d %H:%M"),
            datetime.strptime(closetime, "%Y-%m-%d %H:%M"), self.panel.plans, self.projects, state)
        self.destroy()

class ProgressBox(ttk.Frame):
    def __init__(self, master=None, taskname:str=""):
        super().__init__(master=master)
        self.task = ttk.Label(self, text=taskname)
        self.task.pack(side="left")
        self.pro = ttk.Progressbar(self, mode="determinate", length=250)
        self.pro.pack(side="left")
        self.tag = ttk.Label(self, text="0%")
        self.tag.pack(side="left")
    def init(self):
        self.pro.config(value=0)
        self.tag.config(text="0%")
    def setmax(self, max:int):
        self.pro.config(maximum=max)
    def step(self):
        if self.pro["value"]+1 == self.pro["maximum"]:
            self.pro.config(value=self.pro["maximum"])
        else:
            self.pro.step()
        per = (self.pro["value"]/self.pro["maximum"])*100
        self.tag.config(text="%.0f"%per+"%")

class AutoseekUI(ttk.Window):
    '''主窗口'''
    def __init__(self):
        self.version = "Autoseek 1.0.0-Alpha"
        super().__init__("Autoseek", resizable=(0,0))
        self._setStyle()
        self._setUI()
        self.place_window_center()
    def _setUI(self):
        # 菜单
        mainmenu = ttk.Menu(self)
        self.cmdmenu = ttk.Menu(mainmenu, tearoff=False)
        mainmenu.add_cascade(label="命令", menu=self.cmdmenu)
        self.helpmenu = ttk.Menu(mainmenu, tearoff=False)
        self.helpmenu.add_command(label="关于", command=lambda:Messagebox.show_info(self.version))
        mainmenu.add_cascade(label="帮助", menu=self.helpmenu)
        self.config(menu=mainmenu)
        self.calender = TkCalender(self)
        self.calender.pack()
        self.notebookpro = ProgressBox(self, "notebook")
        self.worktimepro = ProgressBox(self, "worktime")
    def _setStyle(self):
        style = ttk.Style()
        style.configure("TLabel", anchor="center")