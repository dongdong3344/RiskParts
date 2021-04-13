import tkinter as tk
from tkinter import filedialog
from handlerData import RiskPartsAnalysis
import threading

class MyApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.create_UI()

    def create_UI(self):

        self.geometry('500x300')
        self.title('电子元器件Risk分析')
        self.resizable(0, 0)
        self.configure(bg='#fdfee9')
        self.iconbitmap('analysis.ico')
        select_BOM_btn = tk.Button(self, text='Step1-请选择BOM文件', font=("宋体", 12 ),borderwidth = 2,fg ='black',bg='#9bd4e4',padx=10, width=25,height =3,
                                        command = self.open_BOM_file)
        select_BOM_btn.place(relx=0.5, rely=0.2, anchor=tk.CENTER)
        select_status_btn = tk.Button(self, text ='Step2-请选择元器件状态文件', font=('宋体', 12), borderwidth = 2,fg ='black',bg='#4ce1c3',padx=10, width=25,height = 3,
                                           command = self.open_part_status)

        select_status_btn.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        save_btn = tk.Button(self, text='Step3-保存结果文件', font=('宋体', 12 ), borderwidth = 2,fg ='black',bg='#ff4040',padx=10, width=25,height = 3,
                                      command = lambda :self.thread_it(self.save))



        save_btn.place(relx=0.5, rely=0.8, anchor=tk.CENTER)

        self.BOM_filename = ''
        self.part_status_filename = ''


    def open_BOM_file(self):
        BOM_filename = filedialog.askopenfilename(
            title='请选择BOM文件',
            initialdir='/',
            filetypes=[('BOM文件', ['*.xlsx', '*.xls', '*.xlsm'])])
        self.BOM_filename = BOM_filename

    def open_part_status(self):
        part_status_filename = filedialog.askopenfilename(
            title='请选择元器件状态文件',
            initialdir='/',
            filetypes=[('元器件状态文件', ['*.xlsx', '*.xls', '*.xlsm'])])
        self.part_status_filename = part_status_filename


    def save(self):

        risk_analysis = RiskPartsAnalysis(self.BOM_filename,self.part_status_filename)
        # risk_analysis.main()

    @staticmethod
    def thread_it(func,*args):
        t = threading.Thread(target=func,args = args)
        t.setDaemon(True)
        t.start()
        # t.join()

if __name__ == '__main__':
    app = MyApp()
    app.mainloop()
