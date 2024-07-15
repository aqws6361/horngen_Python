import tkinter as tk

class MyApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.title('test')
        self.geometry('600x300')
        

if __name__ == '__main__':
    app = MyApp()
    app.mainloop()