# encoding=utf-8

import os
import pandas as pd
from pandas import DataFrame
import pypinyin
from queue import Queue
from dataclasses import dataclass, field
from datetime import datetime, timedelta
import asyncio
from playwright.async_api import async_playwright, Page
from playwright.sync_api import sync_playwright
from threading import Thread
import sqlite3

@dataclass
class Plan():
    '''单位计划: starttime[开始时间], endtime[结束时间]'''
    starttime: datetime=None
    endtime: datetime=None
    def alloc(self, start:datetime, end:datetime):
        '''分配时间'''
        self.starttime = start
        self.endtime = end
    def __str__(self):
        return self.__class__.__name__
@dataclass
class Routine(Plan):
    '''一般事务'''
    pass
@dataclass
class Project(Plan):
    '''项目: name[项目号]'''
    name: str=""
    def __str__(self):
        return f"{super().__str__()}_{self.name}"
@dataclass
class Meeting(Plan):
    '''会议: name[会议名称], lasttime[时长h], frequency[周期("single"|"daily"|"weekly")]'''
    name: str=""
    lasttime: int=0
    frequency: str=""
    def __str__(self):
        return f"{super().__str__()}_{self.name}_{self.lasttime}_{self.frequency}"

@dataclass
class User():
    '''用户类: name[中文名], notebookID[记录账号], notebookID[记录密码], worktimeID[工时账号], worktimePWD[工时密码]'''
    name: str=""
    notebookID: str=""
    notebookPWD: str=""
    worktimeID: str=""
    worktimePWD: str=""
    srcpath: str=""
    batch: dict=field(default_factory=lambda:{"proteinA":"N/A", "proteinA_AKTA":"N/A", "Ni":"N/A", "Ni_AKTA":"N/A"})
    switch: bool=False
    google_path: str=""
    def eng_name(self) -> str:
        return "".join(pypinyin.lazy_pinyin(self.name, style=pypinyin.Style.FIRST_LETTER)).upper()
    def todb(self):
        return (self.name, self.notebookID, self.notebookPWD, self.worktimeID, self.worktimePWD,
            self.srcpath, self.batch["proteinA"], self.batch["proteinA_AKTA"],
            self.batch["Ni"], self.batch["Ni_AKTA"], self.switch, self.google_path)
    def save(self):
        with SQLDB("autoseek.db") as db:
            db.update_user(self)
@dataclass
class Record():
    '''记录类: date[日期], protocol[SOP], df[数据表], filepath[Excel文件路径], state[状态("U"|"M"|"C")]'''
    date: datetime=None
    protocol: str=""
    df: DataFrame=None
    filepath: str=""
    state: str="U"
    def todb(self) -> tuple:
        return (self.date.strftime("%Y-%m-%d"), self.protocol, self.filepath, self.state)
    def save(self) -> None:
        '''生成Excel文件'''
        self.filepath = "datafile/"+self.date.strftime("%Y%m%d")+self.protocol+".xlsx"
        # 设置表格样式
        with pd.ExcelWriter(self.filepath, engine="xlsxwriter") as writer:
            self.df.to_excel(writer, index=False)
            fmt = writer.book.add_format({
                "font_name": "Times New Roman",
                "valign": "vcenter",
                "align": "center"
                })
            borderfmt = writer.book.add_format({"border": 1})
            # 设置行高
            writer.sheets["Sheet1"].set_default_row(20)
            # 设置列宽
            for i, col in enumerate(self.df):
                series = self.df[col]
                max_len = max((
                series.astype(str).map(len).max(),
                2*len(str(series.name))
                )) + 2
                writer.sheets["Sheet1"].set_column(i, i, max_len, fmt)
            writer.sheets["Sheet1"].set_column("A:B", 8, fmt)
            writer.sheets["Sheet1"].set_column("E:G", 13, fmt)
            writer.sheets["Sheet1"].set_column("I:J", 12, fmt)
            # 设置边框
            writer.sheets["Sheet1"].conditional_format("A1:J1000",{"type":"no_blanks", "format":borderfmt})
    def complete(self) -> None:
        self.protocol = ""
        self.df = None
        os.remove(self.filepath)
        self.filepath = ""
        self.state = "C"
    def buffers(self) -> list[str]:
        s1 = list(self.df["缓冲液成分"])
        s2 = list(set(s1))
        s2.sort(key=s1.index)
        return s2
    def projects(self) -> list[str]:
        s1 = list(self.df["项目号"])
        s2 = list(set(s1))
        s2.sort(key=s1.index)
        return s2
    def conclusion(self) -> str:
        c = ""
        for p in self.projects():
            total = self.df[self.df["项目号"]==p]
            success = 0
            fail = []
            for i in total["总量(mg)"].index:
                if float(total["总量(mg)"][i]) > 0:
                    success = success + 1
                else:
                    fail.append(total["蛋白编号"][i])
            if success == len(total):
                c += f"{p}项目纯化{len(total)}个样品，均获得蛋白\n"
            elif success == 0:
                c += f"{p}项目纯化{len(total)}个样品，均未获得蛋白\n"
            else:
                c += (f"{p}项目纯化{len(total)}个样品，"+"、".join(fail)+f"未获得蛋白，其余{len(total)-len(fail)}个均获得蛋白\n")
        return c[:-1]
@dataclass
class Schedule():
    '''日程: date[日期], opentime[上班时间], closetime[下班时间], plans[计划列表(Routine|Project|Meeting)], state[状态("U"|"M"|"C")]'''
    date: datetime=None
    opentime: datetime=None
    closetime: datetime=None
    plans: list[Plan]=field(default_factory=lambda:[])
    projects: list[Project]=field(default_factory=lambda:[])
    state: str="U"
    def todb(self) -> tuple:
        return (
            self.date.strftime("%Y-%m-%d"),
            self.opentime.strftime("%Y-%m-%d %H:%M"),
            self.closetime.strftime("%Y-%m-%d %H:%M"),
            ",".join([plan.__str__() for plan in self.plans]),
            ",".join([project.name for project in self.projects]),
            self.state
        )
    def dispatch(self):
        '''自动调度计划，根据项目数分配时长'''
        if len(self.projects) > 0:
            starttime = self.opentime
            for plan in self.plans:
                if isinstance(plan, Meeting):
                    starttime = starttime + timedelta(hours=plan.lasttime)
            ave = (self.closetime-starttime)/len(self.projects)
            for i in range(len(self.projects)):
                if i == len(self.projects)-1:
                    self.projects[i].alloc(self.closetime-ave, self.closetime)
                else:
                    self.projects[i].alloc(starttime+ave*i, starttime+ave*(i+1))
    def complete(self):
        self.plans = []
        self.projects = []
        self.state = "C"

class SQLDB:
    def __init__(self, dbpath) -> None:
        self.con = sqlite3.connect(dbpath)
        self.cur = self.con.cursor()
    def __enter__(self):
        return self
    def __exit__(self, type, val, tb):
        self.cur.close()
        self.con.close()
    def get_admin(self) -> User:
        self.cur.execute("SELECT * FROM user WHERE type='admin'")
        u = self.cur.fetchone()
        return User(u[1], u[2], u[3], u[4], u[5], u[6], {
            "proteinA":u[7], "proteinA_AKTA":u[8], "Ni":u[9], "Ni_AKTA":u[10]
        }, u[11], u[12])
    def update_user(self, user:User):
        try:
            self.cur.execute('''
            UPDATE user SET (name,notebookID,notebookPWD,worktimeID,worktimePWD,srcpath,
            proteinA,proteinA_AKTA,Ni,Ni_AKTA,switch,google_path)=(?,?,?,?,?,?,?,?,?,?,?,?) WHERE type='admin'
            ''', user.todb())
            self.con.commit()
        except Exception as e:
            self.con.rollback()
    def get_buffer(self) -> dict:
        self.cur.execute("SELECT weeknum,batch FROM buffer")
        buffers = self.cur.fetchall()
        return {b[0]:b[1] for b in buffers}
    def update_buffer(self, buffers:dict):
        for k,v in buffers.items():
            self.cur.execute("SELECT * FROM buffer WHERE weeknum=?", (k,))
            try:
                if self.cur.fetchall():
                    self.cur.execute("UPDATE buffer SET (weeknum,batch)=(?,?) WHERE weeknum=?", (k,v,k))
                else:
                    self.cur.execute("INSERT INTO buffer(weeknum,batch) VALUES(?,?)", (k,v))
                self.con.commit()
            except Exception as e:
                self.con.rollback()
    def get_rstate(self) -> dict:
        Rstate = {}
        self.cur.execute("SELECT date,state FROM record")
        for s in self.cur.fetchall():
            Rstate[s[0]] = s[1]
        return Rstate
    def get_wstate(self) -> dict:
        Wstate = {}
        self.cur.execute("SELECT date,state FROM schedule")
        for s in self.cur.fetchall():
            Wstate[s[0]] = s[1]
        return Wstate
    def init_recordlist(self) -> dict:
        recordlist = {}
        self.cur.execute("SELECT * FROM record WHERE state='M'")
        for r in self.cur.fetchall():
            df = pd.read_excel(r[3]) if r[3] else None
            recordlist[r[1]] = Record(datetime.strptime(r[1], "%Y-%m-%d"), r[2], df, r[3], r[4])
        return recordlist
    def init_schelist(self) -> dict:
        schelist = {}
        self.cur.execute("SELECT * FROM schedule WHERE state='M'")
        for s in self.cur.fetchall():
            plans = []
            plans_str = s[4].split(",") if s[4] else []
            for p in plans_str:
                pstr = p.split("_")
                if pstr[0] == "Routine":
                    plans.append(Routine())
                elif pstr[0] == "Meeting":
                    plans.append(Meeting(name=pstr[1], lasttime=int(pstr[2]), frequency=pstr[3]))
                elif pstr[0] == "Project":
                    plans.append(Project(name=pstr[1]))
            names = s[5].split(",") if s[5] else []
            schelist[s[1]] = Schedule(
                datetime.strptime(s[1], "%Y-%m-%d"),
                datetime.strptime(s[2], "%Y-%m-%d %H:%M"),
                datetime.strptime(s[3], "%Y-%m-%d %H:%M"),
                plans,
                [Project(name=i) for i in names],
                s[6]
            )
        return schelist
    def update_recordlist(self, recordlist:dict):
        for k,v in recordlist.items():
            self.cur.execute("SELECT * FROM record WHERE date=?", (k,))
            try:
                if self.cur.fetchall():
                    self.cur.execute("UPDATE record SET (date,protocol,filepath,state)=(?,?,?,?) WHERE date=?", v.todb()+(k,))
                else:
                    self.cur.execute("INSERT INTO record(date,protocol,filepath,state) VALUES(?,?,?,?)", v.todb())
                self.con.commit()
            except Exception as e:
                self.con.rollback()
    def update_schelist(self, schelist:dict):
        for k,v in schelist.items():
            self.cur.execute("SELECT * FROM schedule WHERE date=?", (k,))
            try:
                if self.cur.fetchall():
                    self.cur.execute("UPDATE schedule SET (date,opentime,closetime,plans,projects,state)=(?,?,?,?,?,?) WHERE date=?", v.todb()+(k,))
                else:
                    self.cur.execute("INSERT INTO schedule(date,opentime,closetime,plans,projects,state) VALUES(?,?,?,?,?,?)", v.todb())
                self.con.commit()
            except Exception as e:
                self.con.rollback()
    def delete_schedule(self, date:str):
        self.cur.execute("SELECT * FROM schedule WHERE date=?", (date,))
        try:
            if self.cur.fetchall():
                self.cur.execute("DELETE FROM schedule WHERE date=?", (date,))
                self.con.commit()
        except Exception as e:
            self.con.rollback()
class DataModel():
    def __init__(self) -> None:
        self.user = None
        self.buffers = {}
        self.recordlist = {}
        self.schelist = {}
        self.Rstatedict = {}
        self.Wstatedict = {}
        # 消息队列
        self.notebookQueue = Queue()
        self.worktimeQueue = Queue()
    def read(self):
        '''读取数据库'''
        with SQLDB("autoseek.db") as db:
            self.user = db.get_admin()
            self.buffers = db.get_buffer()
            self.recordlist = db.init_recordlist()
            self.schelist = db.init_schelist()
            self.Rstatedict = db.get_rstate()
            self.Wstatedict = db.get_wstate()
    def save(self):
        '''保存'''
        with SQLDB("autoseek.db") as db:
            db.update_recordlist(self.recordlist)
            db.update_schelist(self.schelist)
    def run(self):
        # 测试
        # Thread(name="notebook", target=lambda:asyncio.run(self.notebook()), daemon=True).start()
        if len(self.recordlist):
            Thread(name="notebook", target=lambda:asyncio.run(self.notebook()), daemon=True).start()
        else:
            self.notebookQueue.put("no task")
            self.notebookQueue.put("end")
        if len(self.schelist):
            Thread(name="worktime", target=self.worktime, daemon=True).start()
        else:
            self.worktimeQueue.put("no task")
            self.worktimeQueue.put("end")
    async def notebook(self) -> None:
        '''
        实验记录处理函数
        '''
        async with async_playwright() as playwright:
            browser = await playwright.chromium.launch(
                headless=not self.user.switch,
                executable_path=self.user.google_path
            )
            context = await browser.new_context()
            page = await context.new_page()
            page.set_default_timeout(60000)
            try:
                await page.goto("https://edm.ilabpower.com/")
                # 登录
                await page.fill("[placeholder='请输入手机号/邮箱']", self.user.notebookID)
                await page.fill("[placeholder='请输入密码']", self.user.notebookPWD)
                async with page.expect_navigation():
                    await page.click("//button/*[text()='登录']")
                await page.wait_for_load_state("networkidle")
                # 打开实验数据管理模块
                async with page.expect_popup() as popup_info:
                    # 旧版app入口
                    # await page.click("//a[text()='实验数据管理']")

                    # 新版app入口
                    await page.click("//div[text()='电子实验记录本']")
                page1 = await popup_info.value
                page1.set_default_timeout(60000)
                await page1.wait_for_load_state("networkidle")
                # 关闭公告
                # fo = page1.locator("//*[@id='helpModal']/../.. >> visible=true")
                # await fo.locator("a.ivu-modal-close").click()
                # 循环写实验记录
                tasks = []
                sortlist = sorted(self.recordlist.items())
                for recordtuple in sortlist:
                    record = recordtuple[1]
                    weeknum = record.date.strftime("%Y-%W")
                    if weeknum[-2:] == "00":
                        weeknum = (record.date-timedelta(days=6)).strftime("%Y-%W")
                    if weeknum not in self.buffers:
                        self.notebookQueue.put("no buffer")
                        continue
                    await page1.click("//p[text()='新建实验']")
                    await page1.wait_for_selector("//div[text()='新建实验']/../..", state="attached")
                    panel = page1.locator("//div[text()='新建实验']/../..")     # base panel
                    # 实验标题
                    experiment_title = record.date.strftime("%Y%m%d") + "-" + record.protocol
                    await panel.locator("[placeholder='输入标题(300字符)']").fill(experiment_title)
                    # 选择模板
                    await panel.locator("//label[text()='实验模板']/.. >> //input[@type='text']").click()
                    protocol_list = panel.locator("//label[text()='实验模板']/.. >> li.ivu-select-item")
                    for i in range(await protocol_list.count()):
                        protocolname = await protocol_list.nth(i).text_content()
                        if protocolname[0:len(record.protocol)] == record.protocol:
                            await protocol_list.nth(i).click()
                            break
                    # 选择记录本
                    await panel.locator("//label[text()='记录本']/.. >> //input[@type='text']").click()
                    notebook_list = panel.locator("//label[text()='记录本']/.. >> li.ivu-select-item")
                    for i in range(await notebook_list.count()):
                        notebookname = await notebook_list.nth(i).text_content()
                        notebookname = notebookname[notebookname.find("(")+1:-1]
                        if notebookname == ("蛋白纯化-"+self.user.name):
                            await notebook_list.nth(i).click()
                            break
                    # await page1.goto("https://edm.ilabpower.com/project", wait_until="networkidle")
                    # await page1.wait_for_timeout(1000)
                    # await page1.click("text=全部")
                    # 提交
                    async with page1.expect_popup() as popup_info:
                        # await page1.click("text=NB221862-2")
                        await panel.locator("button:has-text('提交')").click()
                    page2 = await popup_info.value
                    page2.set_default_timeout(60000)
                    tasks.append(asyncio.create_task(self.writeRecord(page2, record)))
                if tasks:
                    await asyncio.wait(tasks)
            except TimeoutError:
                self.notebookQueue.put("timeout")
            finally:
                # 关闭浏览器
                await context.close()
                await browser.close()
                with SQLDB("autoseek.db") as db:
                    db.update_recordlist(self.recordlist)
                    self.recordlist = db.init_recordlist()
                self.notebookQueue.put("end")
    async def writeRecord(self, page:Page, record:Record) -> None:
        '''
        填写记录
        '''
        await page.wait_for_load_state("networkidle")
        # 实验成功
        await page.locator("//div[text()='结论']/.. >> //span[text()='请选择']").click()
        await page.click("text=成功")
        # 关键词
        keywords = record.projects()
        await page.fill("textarea[placeholder='一百字以内']",",".join(keywords))
        # 填写实验人，实验日期
        await page.fill("//*[text()='实验人员']/../.. >> input", self.user.name)
        await page.fill("//*[text()='实验日期']/../.. >> input", record.date.strftime("%Y.%m.%d"))
        # 柱子批号
        await page.wait_for_selector("div.component-title:has-text('耗材') >> .. >> tr.ivu-table-row", state="attached")
        consumable = page.locator("div.component-title:has-text('耗材') >> .. >> tr.ivu-table-row")
        if record.protocol == "Protein A重力柱亲和层析操作流程":
            await consumable.nth(0).locator("//td[3]").click()
            await consumable.nth(0).locator("//td[3] >> input").fill(self.user.batch["proteinA"])
        elif record.protocol == "Protein A AKTA亲和层析操作流程":
            await consumable.nth(0).locator("//td[3]").click()
            await consumable.nth(0).locator("//td[3] >> input").fill(self.user.batch["proteinA_AKTA"])
        elif record.protocol == "Ni2+ Smart Bead重力柱亲和层析操作流程":
            await consumable.nth(0).locator("//td[3]").click()
            await consumable.nth(0).locator("//td[3] >> input").fill(self.user.batch["Ni"])
        elif record.protocol == "Ni2+ Smart Beads AKTA亲和层析操作流程":
            await consumable.nth(0).locator("//td[3]").click()
            await consumable.nth(0).locator("//td[3] >> input").fill(self.user.batch["Ni_AKTA"])
        else:
            pass
        await page.click("div.component-title:has-text('耗材')")
        await page.mouse.wheel(0,1000)
        await page.wait_for_load_state("networkidle")
        # 填写溶液种类、批号、保质期
        await page.wait_for_selector("div.component-title:has-text('溶液配制') >> .. >> tr.ivu-table-row", state="attached")
        buffer_list = page.locator("div.component-title:has-text('溶液配制') >> .. >> tr.ivu-table-row")
        weeknum = record.date.strftime("%Y-%W")
        if weeknum[-2:] == "00":
            weeknum = (record.date-timedelta(days=6)).strftime("%Y-%W")
        buffer_batch = self.buffers[weeknum]       # 常规buffer批号
        shelflife = self.getnextmonth(buffer_batch[-8:])     # 保质期
        buffers = record.buffers()     # 缓冲液列表
        # 添加缓冲液
        if len(buffers) > 1 :
            for i in range(len(buffers)-1):
                await page.click("div.component-title:has-text('溶液配制') >> .. >> text='新增'")
                newrow = buffer_list.nth(await buffer_list.count()-1)
                await newrow.locator("//td[2]").click()
                if record.protocol == "分子排阻色谱层析操作流程":
                    await newrow.locator("//td[2] >> input").fill("Binding Buffer（"+buffers[i+1]+"）")
                elif record.protocol == "离子交换层析操作流程":
                    await newrow.locator("//td[2] >> input").fill("Buffer A（"+buffers[i+1]+"）")
                else:
                    await newrow.locator("//td[2] >> input").fill("置换缓冲液（"+buffers[i+1]+"）")
                await newrow.locator("//td[4]").click()
                await newrow.locator("//td[4] >> input").fill("三优生物")
                if record.protocol == "离子交换层析操作流程":
                    await page.click("div.component-title:has-text('溶液配制') >> .. >> text='新增'")
                    extrarow = buffer_list.nth(await buffer_list.count()-1)
                    await extrarow.locator("//td[2]").click()
                    elution = buffers[i+1].split(",")
                    elution.insert(1,"0.5M NaCl")
                    elution = ",".join(elution)
                    await extrarow.locator("//td[2] >> input").fill("Buffer B（"+elution+"）")
                    await extrarow.locator("//td[4]").click()
                    await extrarow.locator("//td[4] >> input").fill("三优生物")
        for i in range(await buffer_list.count()):
            buffername = await buffer_list.nth(i).locator("//td[2]/div/div").text_content()
            # buffer 种类
            if record.protocol == "分子排阻色谱层析操作流程":
                if buffername == "Binding Buffer":
                    await buffer_list.nth(i).locator("//td[2]").click()
                    await buffer_list.nth(i).locator("//td[2] >> input").fill("Binding Buffer（"+buffers[0]+"）")
            elif record.protocol == "离子交换层析操作流程":
                if buffername == " Buffer A":
                    await buffer_list.nth(i).locator("//td[2]").click()
                    await buffer_list.nth(i).locator("//td[2] >> input").fill("Buffer A（"+buffers[0]+"）")
                elif buffername == "Buffer B":
                    await buffer_list.nth(i).locator("//td[2]").click()
                    elution = buffers[0].split(",")
                    elution.insert(1,"0.5M NaCl")
                    elution = ",".join(elution)
                    await buffer_list.nth(i).locator("//td[2] >> input").fill("Buffer B（"+elution+"）")
            else:
                if buffername == "置换缓冲液":
                    await buffer_list.nth(i).locator("//td[2]").click()
                    await buffer_list.nth(i).locator("//td[2] >> input").fill("置换缓冲液（"+buffers[0]+"）")
            # 批号、保质期
            if (record.protocol == "离子交换层析操作流程") and (("Buffer A" in buffername) or ("Buffer B" in buffername)):
                await buffer_list.nth(i).locator("//td[3]").click()
                await buffer_list.nth(i).locator("//td[3] >> input").fill(self.user.eng_name()+record.date.strftime("%Y%m%d"))     # 批号
                await buffer_list.nth(i).locator("//td[5]").click()
                await buffer_list.nth(i).locator("//td[5] >> input").fill(self.getnextmonth(record.date.strftime("%Y%m%d")))     # 保质期
                await buffer_list.nth(i).locator("//td[3]").click()
            else:
                await buffer_list.nth(i).locator("//td[3]").click()
                await buffer_list.nth(i).locator("//td[3] >> input").fill(buffer_batch)     # 批号
                await buffer_list.nth(i).locator("//td[5]").click()
                await buffer_list.nth(i).locator("//td[5] >> input").fill(shelflife)     # 保质期
                await buffer_list.nth(i).locator("//td[3]").click()
        await page.click("div.component-title:has-text('溶液配制')")
        await page.mouse.wheel(0,1200)
        await page.wait_for_load_state("networkidle")
        # 上传数据文件
        await page.frame_locator("text=实验结果 导入 >> iframe").locator("text=文件").click()
        await page.frame_locator("text=实验结果 导入 >> iframe").locator("//div[@title='导入']").click()
        await page.frame_locator("text=实验结果 导入 >> iframe").locator("text=Excel文件").first.click()
        async with page.expect_file_chooser() as fc_info:
            await page.frame_locator("text=实验结果 导入 >> iframe").locator("//div[@title='导入 Excel 文件']").click()
        file_chooser = await fc_info.value
        await file_chooser.set_files(record.filepath)
        # 填写实验结论
        await page.fill("div.component-title:has-text('实验结论') >> .. >> textarea", record.conclusion())
        # 填写数据路径
        await page.fill("div.component-title:has-text('数据保存路径') >> .. >> textarea", self.user.srcpath)
        # 提交保存
        await page.click("p:has-text('签名提交')")
        await page.click("button:has-text('保存并提交')")
        await page.wait_for_selector("//div[text()='签名提交']", state="attached")
        submitpanel = page.locator("//div[text()='签名提交']/../.. >> visible=true")
        await submitpanel.locator("input[placeholder='请输入密码']").fill(self.user.notebookPWD)
        await submitpanel.locator("div.ivu-select-selection").click()
        teams = submitpanel.locator("div.ivu-select-dropdown > ul.ivu-select-dropdown-list > li")
        for i in range(await teams.count()):
            teamname = await teams.nth(i).locator("//div").text_content()
            if (teamname == "FM1-蛋白纯化"):
                ppers = teams.nth(i).locator("//ul/li/li")
                for j in range(await ppers.count()):
                    ppname = await ppers.nth(j).text_content()
                    if (ppname == "王娇"):
                        await ppers.nth(j).click()
                        break
                break
        await submitpanel.locator("div.ivu-select-selection").click()
        await submitpanel.locator("button:has-text(\"提交\")").click()
        await page.wait_for_load_state("networkidle")
        # Close page
        record.complete()
        await page.close()
        self.notebookQueue.put(record.date.strftime("%Y-%m-%d"))
    def getnextmonth(self, today:str) -> str:
        '''获取30天后日期'''
        _today = datetime.strptime(today, "%Y%m%d")
        nextmonth = _today + timedelta(days=30)
        return nextmonth.strftime("%Y-%m-%d")
    def worktime(self) ->None:
        '''工时处理函数'''
        Notfoundlist = []
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                headless=not self.user.switch,
                executable_path=self.user.google_path
            )
            context = browser.new_context()
            page = context.new_page()
            try:
                page.goto("http://sanyoubio.gnway.vip/seeyon/")
                # 登录
                page.fill("#login_username", self.user.worktimeID)
                page.fill("#login_password1", self.user.worktimePWD)
                page.click("text=登 录")
                sortlist = sorted(self.schelist.items())
                for scheduletuple in sortlist:
                    schedule = scheduletuple[1]
                    schedule.dispatch()
                    # 打开填报单
                    page.click("//div[text()='工时管理']")
                    with page.expect_popup() as popup_info:
                        page.click("//div[text()='工时填报单']")
                    page1 = popup_info.value
                    res = self.writeSchedule(page1, schedule)
                    Notfoundlist += res
            except TimeoutError:
                self.worktimeQueue.put("timeout")
            finally:
                # 关闭浏览器
                context.close()
                browser.close()
                with SQLDB("autoseek.db") as db:
                    db.update_schelist(self.schelist)
                    self.schelist = db.init_schelist()
                if Notfoundlist:
                    self.worktimeQueue.put("未找到项目号："+",".join(Notfoundlist))
                self.worktimeQueue.put("end")
    def writeSchedule(self, page:Page, schedule:Schedule):
        Notfoundlist = []
        page.wait_for_load_state("networkidle")
        frame1 = page.frame("zwIframe")
        # 填写时间
        frame1.fill("//div[text()='填写时间']/.. >> input[maxlength='255']", schedule.opentime.strftime("%Y-%m-%d"))
        # 平台事务
        if len(schedule.projects) == 0:
            frame1.fill("//div[text()='平台事务开始时间']/.. >> input[maxlength='255']", schedule.opentime.strftime("%Y-%m-%d %H:%M"))
            frame1.fill("//div[text()='平台事务结束时间']/.. >> input[maxlength='255']", schedule.closetime.strftime("%Y-%m-%d %H:%M"))
        else:
            if schedule.opentime < schedule.projects[0].starttime:
                frame1.fill("//div[text()='平台事务开始时间']/.. >> input[maxlength='255']", schedule.opentime.strftime("%Y-%m-%d %H:%M"))
                frame1.fill("//div[text()='平台事务结束时间']/.. >> input[maxlength='255']", schedule.projects[0].starttime.strftime("%Y-%m-%d %H:%M"))
            # 项目工时
            frame1.click("//div[text()='平台']/.. >> div[class='icon CAP cap-icon-mingxibiaoxuanzeqi']")
            page.wait_for_timeout(1000)
            page.frame_locator("#RelationPage_main >> iframe").locator("input[type='checkbox']").check()
            page.click("text=确定")
            frame1.wait_for_selector("//div[text()='蛋白纯化']")
            for i in range(len(schedule.projects)-1):
                with page.expect_navigation():
                    frame1.click("text=复制")
            frame1.wait_for_selector("//div[text()='项目号']", state="attached")
            sections = frame1.locator("//div[text()='项目号']/.. ")
            # print("section:%d"%sections.count())
            for i in range(len(schedule.projects)):
                sections.nth(i).locator("div[class='icon CAP cap-icon-mingxibiaoxuanzeqi']").click()
                # 搜索项目
                page.wait_for_selector("iframe[name='layui-layer-iframe%d']"%(2+i), state="attached")
                active_frame = page.frame_locator("iframe[name='layui-layer-iframe%d']" % (2+i))
                page.wait_for_timeout(1000)
                active_frame.locator("//em[@title='项目号']/.. >> input").fill(schedule.projects[i].name)
                active_frame.locator("button:has-text('筛选')").click()
                page.wait_for_timeout(500)
                search_list = active_frame.locator("tr:has(input[type='checkbox'])")
                search_num = search_list.count()
                if search_num == 0:
                    page.click("text=取消")
                    Notfoundlist.append(schedule.projects[i].name)
                else:
                    for j in range(search_num):
                        content = search_list.nth(j).locator("//td[3] >> span").text_content()
                        content = content.strip()
                        if schedule.projects[i].name == content:
                            search_list.nth(j).locator("input[type='checkbox']").check()
                            page.click("text=确定")
                            page.wait_for_timeout(500)
                            # 填写开始结束时间
                            frame1.locator("//div[text()='项目开始时间']/.. >> input[maxlength='255']").nth(i).fill(schedule.projects[i].starttime.strftime("%Y-%m-%d %H:%M"))
                            frame1.locator("//div[text()='项目结束时间']/.. >> input[maxlength='255']").nth(i).fill(schedule.projects[i].endtime.strftime("%Y-%m-%d %H:%M"))
                            break
                        if j == search_num-1:
                            page.click("text=取消")
                            Notfoundlist.append(schedule.projects[i].name)
        # 发送
        frame1.click("text=审批意见:")
        if not Notfoundlist:
            page.click("text=发送")
            page.wait_for_selector("text=确定")
            schedule.complete()
            self.worktimeQueue.put(schedule.opentime.strftime("%Y-%m-%d"))
        page.close()
        return Notfoundlist