from tkinter import *
from PIL import Image, ImageTk
from itertools import cycle
from win32gui import GetForegroundWindow, ShowWindow, FindWindow, SetWindowLong, GetWindowLong
from win32con import SW_MINIMIZE, WS_EX_LAYERED, WS_EX_TRANSPARENT, GWL_EXSTYLE
from glob import iglob, glob
from os import path, remove, makedirs
import os
import sys
import win32gui
import win32con

class App(Tk):
    def __init__(self,*args, **kwargs):
        Tk.__init__(self, *args, **kwargs)
        self.enable_hypno = 1

        self.width = 1920 #self.winfo_screenwidth()
        self.height = 1080 #self.winfo_screenheight()
        self.make_background()


        self.hyp_folders = self.GenFolders()
        self.background_list = self.GenBGList()

        self.formatGifs()
        self.gif_path = "Resources/Background Gif Original/Test.gif"
        self.gif_image_path = "Resources/Gifs/Test/"

        self.runGifBG()

    def make_background(self):
        self.bg = Canvas(self, width=self.width, height=self.height, background='white')
        self.bg.pack(fill=BOTH, expand=YES)
        self.geometry('%dx%d' % (self.width, self.height))
        self.title("Test")
        self.overrideredirect(1)
        self.attributes('-alpha', 0.25)
        self.attributes('-transparentcolor', 'white', '-topmost', 1)
        self.config(bg='white') 
        self.wm_attributes("-topmost", 1)
        setClickthrough()
        self.bg_img = ""

        if self.enable_hypno >= 1:
            self.bg.gif_create = self.bg.create_image(self.width/2,self.height/2,image='')
        #else:
        #    self.bg.config(bg=self.TRANS_CLR())

        self.bg.pack(fill=BOTH, expand=YES)

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
            self.bg.create_image(self.width/2, self.height/2, image=frame)
            ind += 1
            if (ind == self.gif_frame_count):
                ind = 0
        self.after(10, self.updatePreloadedGif, ind)

    def updateUnloadedGif(self, ind=0):
        if ind < self.gif_frame_count:
            self.bg.delete(ALL) # avoid excess memory usage
            if self.bg_img is not NONE:
                del self.bg_img
            self.bg_img = ImageTk.PhotoImage(file=self.imagelist[ind])
            self.bg.create_image(self.width/2, self.height/2, image=self.bg_img)
            ind += 1
            if (ind == self.gif_frame_count):
                ind = 0
        self.after(10, self.updateUnloadedGif, ind)

    def GenFolders(self):
        hyp_folders = glob('Resources\\Images/*/', recursive=True)
        if len(hyp_folders) == 0:
            hyp_folders.append('No .png files found')
        else:
            hyp_folders.append('All')
        return hyp_folders

    def ConvertImg(self, folder, DelOld, screenwidth, screenheight):
        def ResizeImg(img,name,screenwidth,screenheight):
            POS = .9 #percent of screen
            ImgW, ImgH = img.size	
            fract = None
            if ImgW > screenwidth -100:
                fract = (ImgW - 100 - 3000 % screenwidth*POS)/ImgW 
            elif ImgH > screenheight -100:
                fract = (ImgW - 100 - 3000 % screenheight*POS)/ImgH 
            if fract:
                w = int(ImgW*fract)
                h = int(ImgH*fract)
                reimg = img.resize((w,h), Image.ANTIALIAS)
                reimg.save(name+'.png', "PNG")
                print('Resizing...')
                return reimg
            return img
    
        filelist = glob(folder+'*.jpg')+glob(folder+'*.png')
        x=0
        for file in filelist:
            try:
                print(x,'of',len(filelist), file)
                name,sep,tail = file.rpartition('.')
                print(file)
                with Image.open(file) as im:
                    im = ResizeImg(im,name,screenwidth,screenheight)
                    im.save(name+'.png', "PNG")
                x+=1
            except OSError as e:
                print('\n', e, '\n','error', file, '\n')
        if DelOld == 1:
            print('clearing old .jpgs and resizing images too large for your screen...')
            for file in filelist:
                if '.jpg' in file:
                    remove(file)

    def GenBGList(self, BGPath='Resources\\Background Gif Original'):
        bglist = []
        Filepath = path.abspath(BGPath)
        
        for item in glob(Filepath+'/*.gif', recursive=True):
            NewPath = item.split('Resources\\Background Gif Original\\')[-1]
            bglist.append(NewPath)
           
        return bglist

    def formatGifs(self):
        print('Formatting Gifs...')
        mywidth,myheight = self.width, self.height
        Filepath = path.abspath('Resources\\Background Gif Original')
        self.ExtractFrames(mywidth,myheight,Filepath)
        self.background_list = self.GenBGList()
        print('Done.')

    def ExtractFrames(self, mywidth,myheight,Filepath):
        for OGGif in glob(Filepath+'\\*.gif', recursive=True):
            frame = Image.open(OGGif)
            nframes = 0
            namecount = 0
            while frame:
                if 'POV' in path.basename(OGGif):
                    wx = (mywidth-50)/frame.width
                    hx = (myheight-50)/frame.height
                    if wx > hx:
                        Tup = (int(frame.width*hx),int(frame.height*hx))
                    else:
                        Tup = (int(frame.width*wx),int(frame.height*wx))
                    sizeframe = frame.resize(Tup,Image.ANTIALIAS)
                else:
                    sizeframe = frame.resize((mywidth,myheight),Image.ANTIALIAS)

                FileDst = OGGif.replace('Resources\Background Gif Original','Resources\Gifs').replace('.gif','')+'\\'
                if not path.exists(FileDst):
                    makedirs(FileDst)
                cnt = str(namecount)
                if len(str(namecount)) == 1:
                    cnt = '000'+str(namecount)
                elif len(str(namecount)) == 2:
                    cnt = '00'+str(namecount)
                elif len(str(namecount)) == 3:
                    cnt = '0'+str(namecount)
                else:
                    cnt = str(namecount)
                OutFile = '%s/%s-%s.gif'%(FileDst,path.basename(OGGif).replace('.gif','').replace('/',''), cnt)
                sizeframe.save(OutFile,'GIF',quality=95)
                nframes += 1
                namecount = round(namecount+1,1)
                try:
                    frame.seek(nframes)
                except EOFError:
                    break;

    def HandleOSBackground(self, Value):
        try:
            if Value == 1:
                windll.user32.SystemParametersInfoW(20, 0,  path.abspath("Resources\\Desktop Backgrounds\\Healslut Background.png") , 0)
            elif Value == 'Exit':
                windll.user32.SystemParametersInfoW(20, 0,  path.abspath("Resources\\Desktop Backgrounds\\Background.png") , 0)
        except Exception as e:
            print('Desktop Background Setup Failed.')
            HandleError(format_exc(2), e, 'HandleOSBackground', subj='')

def setClickthrough(windowname="Test"):
    try:
        hwnd = FindWindow(None, windowname) # Getting window handle
        # hwnd = root.winfo_id() getting hwnd with Tkinter windows
        # hwnd = root.GetHandle() getting hwnd with wx windows
        lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
        lExStyle |=  WS_EX_TRANSPARENT | WS_EX_LAYERED 
        SetWindowLong(hwnd, GWL_EXSTYLE , lExStyle)
    except Exception as e:
        print(e)

if __name__ == '__main__':
    app = App()
    app.mainloop()