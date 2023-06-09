import random
from datamodel import *
from ttkUI import *
from tkinter.filedialog import askopenfile
from ttkbootstrap.dialogs import Messagebox

class Autoseek(AutoseekUI):
    def __init__(self):
        super().__init__()
        # self.up
        # 数据模型
        self.datamodel = DataModel()
        self.configMenu()
        self.configEvent()
    def configMenu(self):
        '''菜单交互'''
        self.cmdmenu.add_command(label="导入", command=self.parseExcel)
        self.cmdmenu.add_command(label="保存", command=self.datamodel.save)
        self.cmdmenu.add_command(label="运行", command=self.runcmd)
        self.cmdmenu.add_command(label="设置", command=self.userset)
        self.cmdmenu.add_separator()
        self.cmdmenu.add_command(label="退出", command=self.quit)
    def configEvent(self):
        '''事件处理'''
        self.calender.bind_class("TLabel", "<Double-Button-1>", self.RTset)
    def configThread(self):
        '''子线程管理'''
        self.check_notebook()
        self.check_worktime()
    def parseExcel(self):
        '''文件导入'''
        filename = askopenfile("rb", filetypes=[("","*.xlsx"),("","*.xls")])
        try:
            data = pd.read_excel(filename)
            datelist = list(set(data["纯化时间"]))
            for today in datelist:
                _today = datetime.strptime(today,"%Y-%m-%d")
                subdata1 = data[data["纯化时间"]==today]
                if (today in self.datamodel.Rstatedict) and (self.datamodel.Rstatedict[today] == "C"):
                    pass    # 已完成
                else:
                    # 生成记录
                    protocollist = list(set(subdata1["Protocol"]))
                    if len(protocollist) > 1:
                        Messagebox.show_error("同一天禁止出现不同的Protocol", "错误")
                    else:
                        protocol = protocollist[0]
                        subdata2 = subdata1.drop(columns = ["纯化方式", "Protocol", "纯化时间"])
                        # # 修改表格标题
                        if protocol in ["离子交换层析操作流程", "分子排阻色谱层析操作流程"]:
                            subdata2.rename(columns={"上样体积(mL)":"上样量(mg)"}, inplace=True)
                        # 重置编号
                        for i in range(len(subdata2)):
                            subdata2.iloc[i,0] = i+1
                        # 设置数字格式
                        subdata2["浓度(mg/mL)"] = subdata2["浓度(mg/mL)"].map(lambda x:("%.2f")%x)
                        subdata2["总量(mg)"] = subdata2["总量(mg)"].map(lambda x:("%.2f")%x)
                        record = Record(_today, protocol, subdata2, state="M")
                        record.save()
                        self.datamodel.recordlist[today] = record
                        self.calender.Rstatedict[today] = "M"
                if (today in self.datamodel.Wstatedict) and (self.datamodel.Wstatedict[today] == "C"):
                    pass    # 已完成
                else:
                    # 生成日程
                    projectlist = list(set(subdata1["项目号"]))
                    projects = [Project(name=name) for name in projectlist]
                    opentime = datetime.strptime(today+" 07:"+"{:0>2d}".format(random.randint(15,30)),"%Y-%m-%d %H:%M")
                    closetime = datetime.strptime(today+" 16:"+"{:0>2d}".format(random.randint(30,45)),"%Y-%m-%d %H:%M")
                    schedule = Schedule(datetime.strptime(today, "%Y-%m-%d"), opentime, closetime, projects, projects, state="M")
                    self.datamodel.schelist[today] = schedule
                    self.calender.Wstatedict[today] = "M"
            self.calender.event_generate("<<state-change>>")
        except Exception as e:
            Messagebox.show_error(e, "错误")
    def userset(self):
        window = UserWindow(self, self.datamodel.user, self.datamodel.buffers)
        window.show()
        self.datamodel.user = window.user
        self.datamodel.buffers = window.buffers
        self.datamodel.user.save()
        with SQLDB("autoseek.db") as db:
            db.update_buffer(self.datamodel.buffers)
    def RTset(self, event):
        '''手动设置'''
        if event.widget["text"] == "T":
            today = date(self.calender.year.get(), self.calender.month.get(), int(event.widget.master.day.get()))
            today = today.strftime("%Y-%m-%d")      # 当前日期
            schedule = Schedule()
            if today not in self.calender.Wstatedict:
                schedule.date = datetime.strptime(today, "%Y-%m-%d")
                schedule.opentime = datetime.strptime(today+" 07:"+"{:0>2d}".format(random.randint(15,30)),"%Y-%m-%d %H:%M")
                schedule.closetime = datetime.strptime(today+" 16:"+"{:0>2d}".format(random.randint(30,45)),"%Y-%m-%d %H:%M")
            elif self.calender.Wstatedict[today] == "M":
                # 获取对应Schedule对象
                schedule = self.datamodel.schelist[today]
            window = WTimeWindow(self, today, schedule, position=(event.x_root, event.y_root-100))
            if window.schedule.state == "M":
                self.calender.Wstatedict[today] = "M"
                self.datamodel.schelist[today] = window.schedule
            elif window.schedule.state == "U":
                if today in self.calender.Wstatedict:
                    del self.calender.Wstatedict[today]
                if today in self.datamodel.schelist:
                    del self.datamodel.schelist[today]
                with SQLDB("autoseek.db") as db:
                    db.delete_schedule(today)
            self.calender.event_generate("<<state-change>>")
    def check_notebook(self):
        while self.datamodel.notebookQueue.qsize() > 0:
            msg = self.datamodel.notebookQueue.get()
            if msg == "no task":
                Messagebox.show_info("没有可执行记录任务")
            elif msg == "timeout":
                Messagebox.show_error("电子记录系统超时", "错误")
            elif msg == "no buffer":
                Messagebox.show_warning("未添加Buffer批号")
            elif msg == "end":
                self.notebookpro.pack_forget()
                self.notebookpro.init()
            else:
                self.calender.Rstatedict[msg] = "C"
                self.calender.event_generate("<<state-change>>")
                self.notebookpro.step()
        self.after(100, self.check_notebook)
    def check_worktime(self):
        while self.datamodel.worktimeQueue.qsize() > 0:
            msg = self.datamodel.worktimeQueue.get()
            if msg == "no task":
                Messagebox.show_info("没有可执行工时任务")
            elif msg == "timeout":
                Messagebox.show_error("工时系统超时", "错误")
            elif "未找到项目号：" in msg:
                Messagebox.show_info(msg)
            elif msg == "end":
                self.worktimepro.pack_forget()
                self.worktimepro.init()
            else:
                self.calender.Wstatedict[msg] = "C"
                self.calender.event_generate("<<state-change>>")
                self.worktimepro.step()
        self.after(100, self.check_worktime)
    def runcmd(self):
        self.datamodel.run()
        self.notebookpro.setmax(len(self.datamodel.recordlist))
        self.notebookpro.pack()
        self.worktimepro.setmax(len(self.datamodel.schelist))
        self.worktimepro.pack()
        self.configThread()
    def run(self):
        self.datamodel.read()
        self.calender.Rstatedict = self.datamodel.Rstatedict
        self.calender.Wstatedict = self.datamodel.Wstatedict
        self.calender.event_generate("<<state-change>>")
        self.mainloop()

if __name__ == "__main__":
    app = Autoseek()
    app.run()