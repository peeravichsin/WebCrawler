import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
import SemiAuto_sele, DBD_coltd, DBD_plc, DBD_lp


class SampleApp(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title('DataCrawler')
        self.geometry('800x800+400+0')
        self.config(bg='#EEEED5')
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
        self.config(bg='#EEEED5')

        tk.Label(self, bg='#EEEED5', fg='#000000',text="โปรแกรมสำหรับดึงข้อมูลจากเว็บไซต์", font=("Tahoma",30)).pack(side="top", fill="x", pady=10, ipadx=10, ipady=10)
      
        tk.Button(self, bg='#4682B4', fg='#FFFAFA', text="กรมบัญชีกลาง",font=("Tahoma",20),
                  command=lambda: master.switch_frame(PageOne)).pack(side="top",  fill="x",pady=10, ipadx=10, ipady=10)

        tk.Button(self, bg='#8968CD', fg='#FFFAFA',text="กรมพัฒนาธุรกิจการค้า บริษัทจำกัด",font=("Tahoma",20),
                  command=lambda: master.switch_frame(PageTwo)).pack(side="top",  fill="x",pady=10, ipadx=10, ipady=10)

        tk.Button(self, bg='#8968CD', fg='#FFFAFA',text="กรมพัฒนาธุรกิจการค้า บริษัทจำกัดมหาชน",font=("Tahoma",20),
                  command=lambda: master.switch_frame(PageThree)).pack(side="top",  fill="x",pady=10, ipadx=10, ipady=10)

        tk.Button(self, bg='#8968CD', fg='#FFFAFA',text="กรมพัฒนาธุรกิจการค้า ห้างหุ้นส่วนจำกัด",font=("Tahoma",20),
                  command=lambda: master.switch_frame(PageFour)).pack(side="top",  fill="x",pady=10, ipadx=10, ipady=10)

        tk.Button(self, bg='#CD3333', fg='#000000',text="X ปิด",font=("Tahoma",20),
                  command=lambda: self.quit()).pack(side='top', fill="x",pady=50, ipadx=10, ipady=10)
        
    

class PageOne(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master ,width=1000, height=1000)
        self.config(bg='#EEEED5')

        tk.Label(self, bg='#EEEED5', fg='#000000',text="กรมบัญชีกลาง", font=("Tahoma",30)).pack(side="top", fill="x", pady=10)

        tk.Label(self, bg='#EEEED5', fg='#000000',text="วันที่ต้องการจะเริ่มดึงข้อมูล", font=("Tahoma",18)).pack(side="top", fill="x", pady=10)
        sdate = DateEntry(self,  width=30, font=("Tahoma",20)) 
        sdate.pack(side="top", fill="x", pady=10)

        tk.Label(self, bg='#EEEED5', fg='#000000',text="วันสุดท้ายต้องการจะดึงข้อมูล", font=("Tahoma",18)).pack(side="top", fill="x", pady=10)
        edate = DateEntry(self,  width=30, font=("Tahoma",20)) 
        edate.pack(side="top", fill="x", pady=10)

        tk.Label(self, bg='#EEEED5', fg='#000000',text="ชื่อไฟล์ที่ต้องการบันทึก", font=("Tahoma",18)).pack(side="top", fill="x", pady=10)
        fname = tk.Entry(self,  width=30, font=("Tahoma",20)) 
        fname.pack(side="top", fill="x", pady=10)

        tk.Button(self, bg='#4682B4', fg='#FFFAFA',text="เริ่มดึงข้อมูล", font=("Tahoma",18),
                  command=lambda: SemiAuto_sele.govspend(st_date = sdate.get_date(), en_date=edate.get_date(), file_name = fname.get())).pack(side='top', ipadx=10, ipady=10)
        
        tk.Button(self, text="กลับ", font=("Tahoma",18),
                  command=lambda: master.switch_frame(StartPage)).pack(side='top', pady=30, ipadx=10, ipady=10)
        
        

       

class PageTwo(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.config(bg='#EEEED5')

        tk.Label(self, bg='#EEEED5', fg='#000000',text="บริษัทจำกัด", font=("Tahoma",30)).pack(side="top", fill="x", pady=10)

        tk.Label(self, bg='#EEEED5', fg='#000000',text="ชื่อไฟล์ที่ต้องการบันทึก", font=("Tahoma",20)).pack(side="top", fill="x", pady=10)
        fname = tk.Entry(self, width=30, font=("Tahoma",20)) 
        fname.pack(side="top", fill="x", pady=10)

        tk.Button(self,  bg='#8968CD', fg='#FFFAFA', text="เริ่มดึงข้อมูล", font=("Tahoma",20),
                  command=lambda:  DBD_coltd.coltd(file_name = fname.get())).pack(side='top', pady=10, ipadx=10, ipady=10)
        
        tk.Button(self, text="กลับ", font=("Tahoma",20),
                  command=lambda: master.switch_frame(StartPage)).pack(side='top', pady=10, ipadx=10, ipady=10)

class PageThree(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.config(bg='#EEEED5')

        tk.Label(self, bg='#EEEED5', fg='#000000',text="บริษัทจำกัดมหาชน", font=("Tahoma",30)).pack(side="top", fill="x", pady=10)

        tk.Label(self, bg='#EEEED5', fg='#000000',text="ชื่อไฟล์ที่ต้องการบันทึก", font=("Tahoma",20)).pack(side="top", fill="x", pady=10)
        fname = tk.Entry(self, width=30, font=("Tahoma",20)) 
        fname.pack(side="top", fill="x", pady=10)

        tk.Button(self, bg='#8968CD', fg='#FFFAFA', text="เริ่มดึงข้อมูล", font=("Tahoma",20),
                  command=lambda:  DBD_plc.plc(file_name = fname.get())).pack(side='top', pady=10, ipadx=10, ipady=10)
        
        tk.Button(self, text="กลับ", font=("Tahoma",20),
                  command=lambda: master.switch_frame(StartPage)).pack(side='top', pady=10, ipadx=10, ipady=10)


class PageFour(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.config(bg='#EEEED5')

        tk.Label(self, bg='#EEEED5', fg='#000000',text="ห้างหุ้นส่วนจำกัด", font=("Tahoma",30)).pack(side="top", fill="x", pady=10)

        tk.Label(self, bg='#EEEED5', fg='#000000',text="ชื่อไฟล์ที่ต้องการบันทึก", font=("Tahoma",20)).pack(side="top", fill="x", pady=10)
        fname = tk.Entry(self, width=30, font=("Tahoma",20)) 
        fname.pack(side="top", fill="x", pady=10)

        tk.Button(self, bg='#8968CD', fg='#FFFAFA', text="เริ่มดึงข้อมูล", font=("Tahoma",20),
                  command=lambda:  DBD_lp.lp(file_name = fname.get())).pack(side='top', pady=10, ipadx=10, ipady=10)
        
        tk.Button(self, text="กลับ", font=("Tahoma",20),
                  command=lambda: master.switch_frame(StartPage)).pack(side='top', pady=10, ipadx=10, ipady=10)
                

if __name__ == "__main__":
    app = SampleApp()
    app.mainloop()