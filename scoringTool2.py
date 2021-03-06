import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import tkinter as tk
from tkinter import IntVar, ttk
from tkinter import filedialog as fd
from tkinter import messagebox as msg

class mainWindow():
    def __init__(self, master):
        self.master = master
        self.master.grid_propagate(0)

        self.mainMenu = tk.Menu(self.master)
        self.fMenu = tk.Menu(self.mainMenu, tearoff=0)
        self.fMenu.add_command(label="새 파일", command=self.newFile)
        self.fMenu.add_command(label="파일 열기", command=self.openFile)
        self.fMenu.add_command(label="저장", command=self.saveFile)
        self.fMenu.add_command(label="다른 이름으로 저장", command=self.saveAs)
        self.mainMenu.add_cascade(label="파일", menu=self.fMenu)

        self.fMenu2 = tk.Menu(self.mainMenu, tearoff=0)
        self.fMenu2.add_command(label="편집", command=self.dataInput)
        self.fMenu2.add_command(label="통계", command=self.analyze)
        self.mainMenu.add_cascade(label="데이터", menu=self.fMenu2)

        self.master.config(menu=self.mainMenu)

    def display(self, df):
        self.frameMain = tk.Frame(self.master, borderwidth=1, relief="solid", width=800, height=500)
        self.frameMain.grid_propagate(0)
        self.frameMain.grid(row=0, column=0)

        self.scores = df
        self.treeScrollY = tk.Scrollbar(self.master, orient='vertical')
        self.treeScrollY.grid(row=0, column=1, sticky=tk.NE + tk.SE)
        self.tree = ttk.Treeview(self.frameMain, height=23, yscrollcommand=self.treeScrollY.set)
        self.treeScrollY.config(command=self.tree.yview)

        self.tree.delete(*self.tree.get_children())

        self.tree["column"] = list(df.columns)
        for i in df.columns:
            self.tree.column(i, width=45, minwidth=20)
        self.tree["show"] = "headings"

        for column in self.tree["column"]:
            self.tree.heading(column, text=column)

        df_rows = df.to_numpy().tolist()
        
        for row in df_rows:
            self.tree.insert("", "end", values=row)
        
        self.tree.grid(row=0, column=0)

    def newFile(self):
        newWindow = tk.Toplevel(self.master)
        newWindow.title("클래스 정보 설정")
        newWindow.geometry("500x500+100+100")
        n = newFileWindow(self.master, newWindow)
    
    def openFile(self):
        self.f = fd.askopenfilename(title="파일 선택", filetypes=(("xlsx files", "*.xlsx"), ("All files", "*.*")))
        if self.f:
            try:
                self.f = r"{}".format(self.f)
                self.scores = pd.read_excel(self.f)
                self.scores.index = self.scores.index + 1
            except ValueError:
                msg.showerror("경고", "파일이 선택되지 않았습니다.")
            except FileNotFoundError:
                msg.showerror("경고", "파일이 선택되지 않았습니다.")
        
        self.display(self.scores)

    def saveFile(self):
        try:
            self.scores.to_excel(self.f, index=False)
        except AttributeError:
            self.saveAs()

    def saveAs(self):
        f = fd.asksaveasfilename(title="다른 이름으로 저장하기", filetypes=(("xlsx files", "*.xlsx"), ("All files", "*.*")))
        if f:
            self.f = f + ".xlsx"
            self.scores.to_excel(self.f, index=False)
        else:
            print("파일의 이름을 작성하지 않았습니다.")

    def dataInput(self):
        try:
            inputWindow = tk.Toplevel(self.master)
            inputWindow.title("데이터 입력")
            inputWindow.geometry("250x150+100+100")
            inputWindow.resizable(False, False)
            i = inputDataWindow(self.master, inputWindow, self.scores)
        except AttributeError:
            msg.showerror("경고", "파일을 선택하거나 생성해주세요.")

    def analyze(self):
        try:
            analysisWindow = tk.Toplevel(self.master)
            analysisWindow.geometry("600x400+200+500")
            analysisWindow.title("통계")
            a = analysisDataWindow(self.master, analysisWindow, self.scores)
        except AttributeError:
            msg.showerror("경고", "파일을 선택하거나 생성해주세요.")

class newFileWindow:
    def __init__(self, root, master):
        self.master = master
        self.root = root
        self.master.grab_set()

        self.evList = ["1주차", "2주차", "3주차", "4주차", "5주차", "6주차", "7주차", "8주차", "9주차", "출석"]
        self.evLbl = []

        label = tk.Label(self.master, text="학생 수와 클래스 정보를 설정해주세요.(이름과 학번칸은 자동으로 생성됩니다.)")
        label.grid(row=0, columnspan=3, column=0)

        self.totalStudentLbl = tk.Label(self.master, text="학생 수")
        self.totalStudentLbl.grid(row=1, column=0, sticky='W')

        self.totalStudentEnt = tk.Entry(self.master, width=10)
        self.totalStudentEnt.grid(row=1, column=1, sticky='W')
        
        self.numEvalItemsLbl = tk.Label(self.master, text="평가항목 개수")
        self.numEvalItemsLbl.grid(row=2, column=0, sticky='W')

        self.numEvalItemsEnt = tk.Entry(self.master, width=10)
        self.numEvalItemsEnt.grid(row=2, column=1, sticky='W')

        self.numEvalItemsBtn = tk.Button(self.master, text="평가항목 편집", command=lambda: self.setEvalList())
        self.numEvalItemsBtn.grid(row=2, column=2, sticky='W')

        self.evalItemsLbl = tk.Label(self.master, text="평가항목")
        self.evalItemsLbl.grid(row=3, column=0, sticky='W')
        
        self.createCompleteBtn(13, False)
        self.mkEvalLabel()

    # newFile창에 완료버튼의 위치를 지정해서 만들기
    def createCompleteBtn(self, r, update=True):
        if update:
            self.completeBtn.destroy()

        self.completeBtn = tk.Button(self.master, text="완료", command=lambda: self.mkDf())
        self.completeBtn.grid(row=r, column=1)

    #newFile창에 평가항목 Label들을 표시
    def mkEvalLabel(self):
        self.rowCnt = 3 #완료버튼의 위치를 정해주는 변수

        for i in range(len(self.evList)):
            label = tk.Label(self.master, text=str(i+1)+". "+self.evList[i])
            label.grid(row=i+3, column=1, sticky='W')
            self.evLbl.append(label)
            self.rowCnt += 1

        self.createCompleteBtn(self.rowCnt)

    #입력받은 평가항목의 개수를 바탕으로 구체적인 평가항목를 작성할 수 있는 Entry가 포함된 새로운 창 표시
    def setEvalList(self):
        try:
            if int(self.numEvalItemsEnt.get()) > 15:
                msg.showerror("경고", "평가항목는 15개보다 많이 만들 수 없습니다.")
                return False

            self.setting = tk.Toplevel(self.master)
            self.setting.geometry("300x600+200+200")
            self.setting.title("평가항목 입력")
            self.setting.grab_set()

            name = tk.Label(self.setting, text="평가항목들을 입력해주세요")
            name.grid(row=0, column=1)

            rowCnt = 1
            entList = []
            for i in range(int(self.numEvalItemsEnt.get())):
                lab = tk.Label(self.setting, text=str(i+1)+". ")
                lab.grid(row=1+i, column=0)

                ent = tk.Entry(self.setting)
                ent.grid(row=1+i, column=1)
                entList.append(ent)

                rowCnt += 1

            submit = tk.Button(self.setting, text="완료", command=lambda: self.createEvList(entList))
            submit.grid(row=rowCnt, column=1)
        except ValueError:
            msg.showerror("경고", "평가 항목를 입력하지 않았습니다.")

    #setting 창의 Entry로부터 평가항목의 리스트를 만들고 newFile창에 표시
    def createEvList(self, list):
        self.evList = [l.get() for l in list]
        if len(self.evList) != len(set(self.evList)):
            msg.showerror("경고", "중복되는 평가항목이 있습니다.")
            return False

        for l in self.evLbl:
            l.destroy()
        
        self.mkEvalLabel()
        self.setting.destroy()

    #newFile창의 학생수와 평가항목를 바탕으로 Dataframe을 만들기
    def mkDf(self):
        try:
            t = int(self.totalStudentEnt.get())
            self.df = pd.DataFrame(index=range(1, t+1), columns=["학번", "이름"]+self.evList)

            mainWindow.scores = self.df
            mainWindow.display(self.root, self.df)
            self.master.destroy()

        except ValueError:
            msg.showerror("경고", "학생 수를 입력하지 않았습니다.")
            return False

class inputDataWindow:
    def __init__(self, root, master, data):
        self.root = root
        self.master = master
        self.master.grab_set()
        self.data = data
        self.ev = []
        for i in self.data:
            self.ev.append(i)
        self.ev.remove("학번")
        self.ev.remove("이름")
        self.studentFocus = 1
        self.evFocus = 0

        self.evMinusBtn = tk.Button(self.master, text="<<", command=lambda:[self.updateData(), self.studentPlusMinus(plus=False)])
        self.evMinusBtn.grid(row=0, column=0)
        self.rowLbl = tk.Label(self.master, text="학생 (1/"+str(len(self.data))+")")
        self.rowLbl.grid(row=0, column=1, columnspan=2)
        self.evPlusBtn = tk.Button(self.master, text=">>", command=lambda:[self.updateData(), self.studentPlusMinus(plus=True)])
        self.evPlusBtn.grid(row=0, column=3)

        self.evMinusBtn = tk.Button(self.master, text="<<", command=lambda:[self.updateData(), self.evPlusMinus(plus=False)])
        self.evMinusBtn.grid(row=1, column=0)
        self.columnLbl = tk.Label(self.master, text="평가항목: 1주차(1/"+str(len(self.data.columns)-2)+")")
        self.columnLbl.grid(row=1, column=1, columnspan=2)
        self.evPlusBtn = tk.Button(self.master, text=">>", command=lambda:[self.updateData(), self.evPlusMinus(plus=True)])
        self.evPlusBtn.grid(row=1, column=3)

        self.studentLbl = tk.Label(self.master, text="학번")
        self.studentLbl.grid(row=2, column=1)
        self.studentEnt = tk.Entry(self.master)
        self.studentEnt.insert(0, self.data['학번'][self.studentFocus])
        self.studentEnt.grid(row=2, column=2)

        self.nameLbl = tk.Label(self.master, text="이름")
        self.nameLbl.grid(row=3, column=1)
        self.nameEnt = tk.Entry(self.master)
        self.nameEnt.insert(0, self.data['이름'][self.studentFocus])
        self.nameEnt.grid(row=3, column=2)

        self.scoreLbl = tk.Label(self.master, text="점수")
        self.scoreLbl.grid(row=4, column=1)
        self.scoreEnt = tk.Entry(self.master)
        self.scoreEnt.insert(0, self.data[self.ev[self.evFocus]][self.studentFocus])
        self.scoreEnt.grid(row=4, column=2)

        self.ontopValue = IntVar()
        self.ontopCheck = tk.Checkbutton(self.master, text="모든 창 위에 항상고정", variable=self.ontopValue, command=self.ontop)
        self.ontopCheck.grid(row=5, column=0, columnspan=4, sticky=tk.SW)

        self.saveBtn = tk.Button(self.master, text='저장', command=lambda: [self.updateData(), self.save()])
        self.saveBtn.grid(row=5, column=5)

    def setDataType(self):
        self.data["학번"] = self.data["학번"].astype(str)
        self.data["이름"] = self.data["이름"].astype(str)
        for i in self.ev:
            self.data[i] = pd.to_numeric(self.data[i], errors='coerce', downcast='float')

    def save(self):
        self.setDataType()
        mainWindow.scores = self.data
        mainWindow.display(self.root, self.data)

    def updateData(self):
        self.data.at[self.studentFocus, '학번'] = self.studentEnt.get()
        self.data.at[self.studentFocus, '이름'] = self.nameEnt.get()
        try:
            self.data.at[self.studentFocus, self.ev[self.evFocus]]= self.scoreEnt.get()
        except ValueError:
            msg.showerror("경고", "점수칸에는 실수만 입력가능합니다.")

        self.setDataType()

    def updateEntries(self):
        self.studentEnt.delete(0, 'end')
        self.nameEnt.delete(0, 'end')
        self.scoreEnt.delete(0, 'end')

        self.studentEnt.insert(0, self.data['학번'][self.studentFocus])
        self.nameEnt.insert(0, self.data['이름'][self.studentFocus])
        self.scoreEnt.insert(0, self.data[self.ev[self.evFocus]][self.studentFocus])

        self.rowLbl.config(text="학생 ("+str(self.studentFocus)+"/"+str(len(self.data))+")")
        self.columnLbl.config(text="평가항목: "+str(self.ev[self.evFocus])+"("+str(self.evFocus+1)+"/"+str(len(self.data.columns)-2)+")")

    def studentPlusMinus(self, plus=True):
        if plus:
            self.studentFocus += 1
            if self.studentFocus > len(self.data):
                self.studentFocus = 1
        else:
            self.studentFocus -= 1
            if self.studentFocus < 1:
                self.studentFocus = len(self.data)
        
        self.updateEntries()

    def evPlusMinus(self, plus=True):
        if plus:
            self.evFocus += 1
            if self.evFocus >= len(self.ev):
                self.evFocus = 0
        else:
            self.evFocus -= 1
            if self.evFocus < 0:
                self.evFocus = len(self.ev)-1

        self.updateEntries()

    def ontop(self):
        if self.ontopValue.get() == 1:
            self.master.wm_attributes("-topmost", 1)
        else:
            self.master.wm_attributes("-topmost", 0)

class analysisDataWindow:
    def __init__(self, root, master, data):
        self.root = root
        self.master = master
        self.master.grab_set()
        self.data = data
        self.ev = []
        for i in self.data:
            self.ev.append(i)
        self.ev.remove("학번")
        self.ev.remove("이름")
        self.sum = [0 for i in range(len(self.data))]
        
        self.canvas = tk.Frame(self.master, relief='solid', bd=2, width=400, height=400)
        self.canvas.grid(row=0, column=0, rowspan=15)
        self.canvas.grid_propagate(0)

        self.a = IntVar()
        self.showAvCheck = tk.Checkbutton(self.master, text="평균 표시", command=self.plot, variable=self.a)
        self.showAvCheck.grid(row=0, column=1)

        self.mkCheckEv()

    def mkCheckEv(self):
        rowCnt = 1
        self.v = [0 for i in range(len(self.ev))]
        for i in range(len(self.ev)):
            self.v[i] = IntVar()
            self.checkEv = tk.Checkbutton(self.master, text=self.ev[i], command=lambda: self.dataConverge(), variable=self.v[i])
            self.checkEv.grid(row=rowCnt, column=1)
            rowCnt += 1

    def getAvg(self):
        s = 0
        for i in self.sum:
            s += i
        self.avg = s/len(self.sum)

    def dataConverge(self):
        self.sum = [0 for i in range(len(self.data))]
        for i in range(len(self.ev)):
            if self.v[i].get() == 1:
                self.sum = self.sum + self.data[self.ev[i]]
        self.plot()

    def plot(self):
        self.fig = plt.figure(figsize=(3.8,3.8), dpi=50)
        plt.hist(self.sum, bins=10)
        if self.a.get() == 1:
            self.getAvg()
            plt.axvline(self.avg, color='r')
        canvas = FigureCanvasTkAgg(self.fig, self.canvas)
        canvas.get_tk_widget().grid(row=0, column=0)

def main():
    root = tk.Tk()
    root.title("메인 페이지")
    root.geometry("830x560+380+50")
    app = mainWindow(root)

    root.mainloop()

if __name__ == '__main__':
    main()