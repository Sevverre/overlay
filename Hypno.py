from tkinter import *
from PIL import Image, ImageTk
from win32gui import GetForegroundWindow, ShowWindow, FindWindow, SetWindowLong, GetWindowLong, SetLayeredWindowAttributes
from win32con import SW_MINIMIZE, WS_EX_LAYERED, WS_EX_TRANSPARENT, GWL_EXSTYLE, LWA_ALPHA
import win32api
from glob import iglob, glob


class Overlay:
    def __init__(self, delay, opacity, hypno, pinup, globfile,s_rulename,
                             gifset, c_vid, c_pinup, c_hypno):
        self.root = Tk()

        self.width = self.root.winfo_screenwidth()
        self.height = self.root.winfo_screenheight()
        self.delay = delay
        self.opacity = opacity
        self.playingvideo = False
        self.c_hypno = c_hypno
        self.enable_hypno = hypno
        self.enable_pinup = pinup
        self.c_vid = c_vid
        self.gifset = gifset
        self.s_rulename = s_rulename
        self.bg_img = ""

        self.root.overrideredirect(1)
        self.root.geometry('%dx%d+%d+%d' % (self.width, self.height, 0, 0))
        self.root.attributes("-alpha", 0.25)

        # Set transparency and click-through
        print("Configuring tk")
        self.root.config(bg='#000000') 
        self.root.wm_attributes("-topmost", 1)
        self.root.attributes('-transparentcolor', '#000000', '-topmost', 1)

        print("Configuring bg")
        self.bg = Canvas(self.root, highlightthickness=0)
        self.setClickthrough(self.bg.winfo_id())
        
        # Set click-through for canvas
        self.bg.configure(bg="#000000")
        self.bg.pack()
        
        self.gif_path = "Resources/Background Gif Original/Test.gif"
        self.gif_image_path = "Resources/Gifs/Test/"
        self.runGifBG()
       
        self.root.mainloop()

    def runGifBG(self):
        print("Setting PhotoImages")
        self.imagelist = sorted(glob(self.gif_image_path + "*.gif", recursive=True))
        self.gif_frame_count = len(self.imagelist)

        self.mem_error = False
        try:
            self.gif_frames = [(PhotoImage(file=image)) for image in self.imagelist]
        except Exception as e:
            if str(e) == "not enough free memory for image buffer":
                self.mem_error = True
            else:
                print(e)
                return

        if self.mem_error:
            self.updateUnloadedGif()
        else:
            self.updatePreloadedGif()

    def updatePreloadedGif(self, ind=0):
        if ind < self.gif_frame_count:
            frame = self.gif_frames[ind]
            self.bg.delete(ALL) # avoid excess memory usage
            self.bg.configure(width=frame.width(), height=frame.height())
            self.bg.create_image(frame.width()/2, frame.height()/2, image=frame)
            self.bg.place(relx=0.5, rely=0.5, anchor=CENTER)
            ind += 1
            if (ind == self.gif_frame_count):
                ind = 0
        self.root.after(10, self.updatePreloadedGif, ind)

    def updateUnloadedGif(self, ind=0):
        if ind < self.gif_frame_count:
            self.bg.delete(ALL) # avoid excess memory usage
            if self.bg_img is not NONE:
                del self.bg_img
            self.bg_img = ImageTk.PhotoImage(file=self.imagelist[ind])
            self.bg.configure(width=self.bg_img.width(), height=self.bg_img.height())
            self.bg.create_image(self.width/2, self.height/2, image=self.bg_img)
            ind += 1
            if (ind == self.gif_frame_count):
                ind = 0
        self.root.after(10, self.updateUnloadedGif, ind)


    def setClickthrough(self, hwnd):
        print("setting window properties")
        try:
            styles = GetWindowLong(hwnd, GWL_EXSTYLE)
            styles = WS_EX_LAYERED | WS_EX_TRANSPARENT
            SetWindowLong(hwnd, GWL_EXSTYLE, styles)
            SetLayeredWindowAttributes(hwnd, 0, 255, LWA_ALPHA)
        except Exception as e:
            print(e)

def launch(delay, opacity, hypno, pinup, globfile,s_rulename,
                             gifset, c_vid, c_pinup, c_hypno):
    print("starting hypno")
    overlay = Overlay(delay, opacity, hypno, pinup, globfile,s_rulename,
                             gifset, c_vid, c_pinup, c_hypno)

#test = Overlay(0, 0, 1, 0, 0, 0, 0, 0, 0, 0)
