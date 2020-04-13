from tkinter import *
from PIL import Image, ImageTk
from win32gui import GetForegroundWindow, ShowWindow, FindWindow, SetWindowLong, GetWindowLong
from win32con import SW_MINIMIZE, WS_EX_LAYERED, WS_EX_TRANSPARENT, GWL_EXSTYLE
import win32api



class Overlay:
    def __init__(self):
        self.root = Tk()
        self.root.title("Custom Overlay")

        self.width = 3840 #self.winfo_screenwidth()
        self.height = 2160 #self.winfo_screenheight()

        #self.root.overrideredirect(1)
        self.root.geometry('%dx%d' % (self.width, self.height))
        self.root.attributes("-alpha", 0.25)

        # Set transparency and click-through
        self.root.config(bg='#000000') 
        self.root.wm_attributes("-topmost", 1)
        self.root.attributes('-transparentcolor', '#000000', '-topmost', 1)

        self.bg = Canvas(self.root, width=self.width, height=self.height, highlightthickness=0)
        frame = ImageTk.PhotoImage(file="Resources/Gifs/Test/Test-0000.gif")
        self.bg.create_image(1920/2, 1080/2, image=frame)

        # Set click-through for canvas
        self.bg.configure(bg="#000000")
        self.bg.pack(fill=BOTH, expand=YES)
        self.setClickthrough()

       
        self.root.mainloop()

    def setClickthrough(self, windowname="Custom Overlay"):
        try:
            hwnd = FindWindow(None, windowname)
            #styles = GetWindowLong(hwnd, GWL_EXSTYLE)
            styles = WS_EX_LAYERED ^ WS_EX_TRANSPARENT
            win32api.SetWindowLong(hwnd, GWL_EXSTYLE, styles)
        except Exception as e:
            print(e)


overlay = Overlay()