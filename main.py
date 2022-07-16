import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import SemiAuto_sele, DBD_coltd, DBD_plc, DBD_lp

class SampleApp(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title('DataCrwaler')
        self.geometry('700x500+450+150')
        self.config(bg='#D9D8D7')
        self._frame = None
        self.switch_frame(StartPage)

    def switch_frame(self, frame_class):
        """Destroys current frame and replaces it with a new one."""
        new_frame = frame_class(self)
        if self._frame is not None:
            self._frame.destroy()
        self._frame = new_frame
        self._frame.pack()

  

        
class StartPage(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.config(bg='#D9D8D7')
        tk.Label(self, text="โปรแกรมสำหรับดึงของมูลจากเว็บไซต์", font=("Arial",25)).pack(side="top", fill="x", pady=10, ipadx=10, ipady=10)

        tk.Button(self, text="กรมบัญชีกลาง",font=("Arial",10),
                  command=lambda: master.switch_frame(PageOne)).pack(side="top",  pady=10, ipadx=10, ipady=10)

        tk.Button(self, text="กรมพัฒนาธุรกิจการค้า บริษัทจำกัด",font=("Arial",10),
                  command=lambda: master.switch_frame(PageTwo)).pack(side="top",  pady=10, ipadx=10, ipady=10)

        tk.Button(self, text="กรมพัฒนาธุรกิจการค้า บริษัทจำกัดมหาชน",font=("Arial",10),
                  command=lambda: master.switch_frame(PageThree)).pack(side="top",  pady=10, ipadx=10, ipady=10)

        tk.Button(self, text="กรมพัฒนาธุรกิจการค้า ห้างหุ้นส่วนจำกัด",font=("Arial",10),
                  command=lambda: master.switch_frame(PageFour)).pack(side="top",  pady=10, ipadx=10, ipady=10)

        tk.Button(self, text="ปิด",font=("Arial",10),
                  command=lambda: self.quit()).pack(side='top', ipadx=10, ipady=10)
        
    

class PageOne(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.config(bg='#D9D8D7')
        tk.Label(self, text="กรมบัญชีกลาง", font=("Arial",25)).pack(side="top", fill="x", pady=10)

        tk.Label(self, text="จำนวนโครงการ", font=("Arial",10)).pack(side="top", fill="x", pady=10)
        pno = tk.Entry(self,  width=30) 
        pno.pack(side="top", fill="x", pady=10)

        tk.Label(self, text="วันที่ต้องการจะเริ่มดึงข้อมูล", font=("Arial",10)).pack(side="top", fill="x", pady=10)
        sdate = tk.Entry(self,  width=30) 
        sdate.pack(side="top", fill="x", pady=10)

        tk.Label(self, text="วันสุดท้ายต้องการจะดึงข้อมูล", font=("Arial",10)).pack(side="top", fill="x", pady=10)
        edate = tk.Entry(self,  width=30) 
        edate.pack(side="top", fill="x", pady=10)

        tk.Label(self, text="ชื่อไฟล์ที่ต้องการบันทึก", font=("Arial",10)).pack(side="top", fill="x", pady=10)
        fname = tk.Entry(self,  width=30) 
        fname.pack(side="top", fill="x", pady=10)

        tk.Button(self, text="เริ่มดึงข้อมูล",
                  command=lambda: SemiAuto_sele.govspend(st_date = sdate.get(), en_date=edate.get(), project_number = pno.get(), file_name = fname.get())).pack(side='top', ipadx=10, ipady=10)
        
        tk.Button(self, text="กลับ",
                  command=lambda: master.switch_frame(StartPage)).pack(side='top', ipadx=10, ipady=10)
        
        

       

class PageTwo(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.config(bg='#D9D8D7')
        tk.Label(self, text="บริษัทจำกัด", font=("Arial",25)).pack(side="top", fill="x", pady=10)

        tk.Label(self, text="ชื่อไฟล์ที่ต้องการบันทึก", font=("Arial",10)).pack(side="top", fill="x", pady=10)
        fname = tk.Entry(self,  width=30) 
        fname.pack(side="top", fill="x", pady=10)

        tk.Button(self, text="เริ่มดึงข้อมูล",
                  command=lambda:  DBD_coltd.coltd(file_name = fname.get())).pack(side='top', ipadx=10, ipady=10)
        
        tk.Button(self, text="กลับ",
                  command=lambda: master.switch_frame(StartPage)).pack(side='top', ipadx=10, ipady=10)

class PageThree(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.config(bg='#D9D8D7')
        tk.Label(self, text="บริษัทจำกัดมหาชน", font=("Arial",25)).pack(side="top", fill="x", pady=10)

        tk.Label(self, text="ชื่อไฟล์ที่ต้องการบันทึก", font=("Arial",10)).pack(side="top", fill="x", pady=10)
        fname = tk.Entry(self,  width=30) 
        fname.pack(side="top", fill="x", pady=10)

        tk.Button(self, text="เริ่มดึงข้อมูล",
                  command=lambda:  DBD_plc.plc(file_name = fname.get())).pack(side='top', ipadx=10, ipady=10)
        
        tk.Button(self, text="กลับ",
                  command=lambda: master.switch_frame(StartPage)).pack(side='top', ipadx=10, ipady=10)


class PageFour(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.config(bg='#D9D8D7')
        tk.Label(self, text="ห้างหุ้นส่วนจำกัด", font=("Arial",25)).pack(side="top", fill="x", pady=10)

        tk.Label(self, text="ชื่อไฟล์ที่ต้องการบันทึก", font=("Arial",10)).pack(side="top", fill="x", pady=10)
        fname = tk.Entry(self,  width=30) 
        fname.pack(side="top", fill="x", pady=10)

        tk.Button(self, text="เริ่มดึงข้อมูล",
                  command=lambda:  DBD_lp.lp(file_name = fname.get())).pack(side='top', ipadx=10, ipady=10)
        
        tk.Button(self, text="กลับ",
                  command=lambda: master.switch_frame(StartPage)).pack(side='top', ipadx=10, ipady=10)
                

if __name__ == "__main__":
    app = SampleApp()
    app.mainloop()