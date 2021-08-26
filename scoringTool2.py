import numpy as np
import pandas as pd

import tkinter as tk
from tkinter import IntVar, ttk
from tkinter import filedialog as fd
from tkinter import messagebox as msg

class mainWindow():
    def __init__(self, master):
        self.master = master
        
        self.mainMenu = tk.Menu(self.master)
        self.fMenu = tk.Menu(self.mainMenu, tearoff=0)
        self.fMenu.add_command(label="새 파일", command=self.newFile)
        self.fMenu.add_command(label="파일 열기", command=self.openFile)
        self.fMenu.add_command(label="다른 이름으로 저장", command=self.saveAs)
        self.mainMenu.add_cascade(label="파일", menu=self.fMenu)

        self.fMenu2 = tk.Menu(self.mainMenu, tearoff=0)
        self.fMenu2.add_command(label="데이터 입력", command=self.dataInput)
        self.mainMenu.add_cascade(label="데이터", menu=self.fMenu2)

        self.master.config(menu=self.mainMenu)

    def display(self, df):
        self.frameMain = tk.Frame(self.master, borderwidth=1, relief="solid", width=700, height=500)
        self.frameMain.grid_propagate(0)
        self.frameMain.grid(row=0, column=0)

        self.scores = df
        self.treeScrollY = tk.Scrollbar(self.master, orient='vertical')
        self.treeScrollX = tk.Scrollbar(self.master, orient='horizontal')
        self.treeScrollY.grid(row=0, column=1, sticky=tk.NE + tk.SE)
        self.treeScrollX.grid(row=1, column=0, sticky=tk.SW + tk.SE)
        self.tree = ttk.Treeview(self.frameMain, height=23, xscrollcommand=self.treeScrollX.set, yscrollcommand=self.treeScrollY.set)
        self.treeScrollY.config(command=self.tree.yview)
        self.treeScrollX.config(command=self.tree.xview)

        self.tree.delete(*self.tree.get_children())

        self.tree["column"] = list(df.columns)
        for i in df.columns:
            self.tree.column(i, width=70, minwidth=20)
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
            except ValueError:
                msg.showerror("경고", "파일이 선택되지 않았습니다.")
            except FileNotFoundError:
                msg.showerror("경고", "파일이 선택되지 않았습니다.")
        
        self.display(self.scores)
        

    def saveAs(self):
        f = fd.asksaveasfilename(title="다른 이름으로 저장하기", filetypes=(("xlsx files", "*.xlsx"), ("All files", "*.*")))
        if f:
            self.f = f + ".xlsx"
            self.scores.to_excel(self.f)
        else:
            print("파일의 이름을 작성하지 않았습니다.")

    def dataInput(self):
        try:
            inputWindow = tk.Toplevel(self.master)
            inputWindow.title("데이터 입력")
            inputWindow.geometry("250x150+100+100")
            inputWindow.resizable(False, False)
            i = inputDataWindow(self.master, inputWindow)
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
            if int(self.numEvalItemsEnt.get()) > 20:
                msg.showerror("경고", "평가항목는 20개 이상 만들 수 없습니다.")
                return False

            setting = tk.Toplevel(self.master)
            setting.geometry("300x600+200+200")
            setting.title("평가항목 입력")

            name = tk.Label(setting, text="평가항목들을 입력해주세요")
            name.grid(row=0, column=1)

            rowCnt = 1
            entList = []
            for i in range(int(self.numEvalItemsEnt.get())):
                lab = tk.Label(setting, text=str(i+1)+". ")
                lab.grid(row=1+i, column=0)

                ent = tk.Entry(setting)
                ent.grid(row=1+i, column=1)
                entList.append(ent)

                rowCnt += 1

            submit = tk.Button(setting, text="완료", command=lambda: [self.createEvList(entList), setting.destroy()])
            submit.grid(row=rowCnt, column=1)
        except ValueError:
            msg.showerror("경고", "평가 항목를 입력하지 않았습니다.")

    #setting 창의 Entry로부터 평가항목의 리스트를 만들고 newFile창에 표시
    def createEvList(self, list):
        self.evList = [l.get() for l in list]

        for l in self.evLbl:
            l.destroy()
        
        self.mkEvalLabel()
    #newFile창의 학생수와 평가항목를 바탕으로 Dataframe을 만들기
    def mkDf(self):
        try:
            t = int(self.totalStudentEnt.get())
            self.df = pd.DataFrame(data=np.zeros((t, len(self.evList)+2)), index=range(1, t+1), columns=["학번", "이름"]+self.evList)
            
            mainWindow.scores = self.df
            mainWindow.display(self.root, self.df)
            self.master.destroy()

        except ValueError:
            msg.showerror("경고", "학생 수를 입력하지 않았습니다.")
            return False

class inputDataWindow:
    def __init__(self, root, master):
        self.master = master
        self.root = root
        self.data = mainWindow.scores
        self.ev = self.data

        self.evMinusBtn = tk.Button(self.master, text="<<", command=lambda: self.evPlusMinus(plus=False))
        self.evMinusBtn.grid(row=0, column=0)
        self.rowLbl = tk.Label(self.master, text="학생 (1/"+str(len(self.data))+")")
        self.rowLbl.grid(row=0, column=1, columnspan=2)
        self.evPlusBtn = tk.Button(self.master, text=">>", command=lambda: self.evPlusMinus(plus=True))
        self.evPlusBtn.grid(row=0, column=3)

        self.evMinusBtn = tk.Button(self.master, text="<<", command=lambda: self.evPlusMinus(plus=False))
        self.evMinusBtn.grid(row=1, column=0)
        self.columnLbl = tk.Label(self.master, text="평가항목 (1/"+str(len(self.data.columns)-2)+")")
        self.columnLbl.grid(row=1, column=1, columnspan=2)
        self.evPlusBtn = tk.Button(self.master, text=">>", command=lambda: self.evPlusMinus(plus=True))
        self.evPlusBtn.grid(row=1, column=3)

        self.studentLbl = tk.Label(self.master, text="학번")
        self.studentLbl.grid(row=2, column=1)
        self.studentEnt = tk.Entry(self.master)
        self.studentEnt.grid(row=2, column=2)

        self.nameLbl = tk.Label(self.master, text="이름")
        self.nameLbl.grid(row=3, column=1)
        self.nameEnt = tk.Entry(self.master)
        self.nameEnt.grid(row=3, column=2)

        self.scoreLbl = tk.Label(self.master, text="점수")
        self.scoreLbl.grid(row=4, column=1)
        self.scoreEnt = tk.Entry(self.master)
        self.scoreEnt.grid(row=4, column=2)

        self.ontopValue = IntVar()
        self.ontopCheck = tk.Checkbutton(self.master, text="모든 창 위에 항상고정", variable=self.ontopValue, command=self.ontop)
        self.ontopCheck.grid(row=5, column=0, columnspan=4, sticky=tk.SW)

    def evPlusMinus(self, plus=True):
        if plus:
            print("hey")
        else:
            print("hey")

    def ontop(self):
        if self.ontopValue.get() == 1:
            self.master.wm_attributes("-topmost", 1)
        else:
            self.master.wm_attributes("-topmost", 0)
'''
O.필요한 함수들
    complete) pandas 데이터프레임을 표 형식으로 창에 표시

I.파일관리
    complete) 새 파일 - pandas 데이터프레임 만들기
    complete) 불러오기 - 엑셀파일을 불러와서 pandas 데이터프레임 만들기
    complete) 저장하기 - pandas 데이터프레임 엑셀파일로 만들어서 저장하기

II.점수관리
    1. 채점툴 - 기존에 만든 채점툴 활용
    2. 계산 - 지정한 범위의 학생들의 평균, 최고점, 최저점, 표준편차 등을 보여주는 툴
    3. 시각화툴 - 그래프로 여러가지를 보여주는 툴
'''

def main():
    root = tk.Tk()
    root.title("메인 페이지")
    root.geometry("720x560+100+100")
    app = mainWindow(root)

    tree = ttk.Treeview()

    
    root.mainloop()

if __name__ == '__main__':
    main()