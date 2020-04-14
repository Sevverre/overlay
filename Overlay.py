from tkinter import *
from tkinter.ttk import Notebook
from PIL import Image, ImageTk
from itertools import cycle
from win32gui import GetForegroundWindow, ShowWindow, FindWindow, SetWindowLong, GetWindowLong
from win32con import SW_MINIMIZE, WS_EX_LAYERED, WS_EX_TRANSPARENT, GWL_EXSTYLE
from multiprocessing import freeze_support, Process, Pipe
from threading import Thread
from glob import iglob, glob
from os import path, remove, makedirs
import os
import sys
import win32gui
import win32con
from functools import partial
from ctypes import windll

import Hypno

PREFDICT_PRESET = \
	{
	'hyp_delay':'500',
	'hyp_opacity':'25',
	'hyp_able':0,
	'hyp_pinup':1,
	'display_rules':0,
	'delold':1,
	'hyp_gfile_var':0,
    'actionmenu': 1,
	'background_select_var':0,
	's_rulename':'Overwatch Helpful',
	'UseHSBackground':0
	}

class App(Tk):
    def __init__(self,*args, **kwargs):
        Tk.__init__(self, *args, **kwargs)
        
        self.hyp_folders = self.GenFolders()
        self.background_list = self.GenBGList()
        self.userPrefs = self.GenUserPrefs()
        self.SetupVars(self.background_list, self.hyp_folders, self.userPrefs)
        self.SetupMenu()
        self.SavePref()

    def SetupVars(self,background_list, hyp_folders, prefdict):
        self.p_hypno,           self.c_hypno = Pipe()
        self.p_vid,             self.c_vid = Pipe()
        self.p_pinup,           self.c_pinup = Pipe()

        self.background_list = background_list
        self.background_select = StringVar(self)

        if len(background_list) < int(prefdict['background_select_var']):
                prefdict['background_select_var'] = 0

        self.background_select_var = int(prefdict['background_select_var'])
        self.background_select.set(background_list[self.background_select_var])

        self.width = self.winfo_screenwidth()
        self.height = self.winfo_screenheight()
        self.winWidth = 50
        self.winHeight= 100

        geo = '%dx%d+%d+%d' % (self.winWidth, self.winHeight, self.width-self.winWidth, self.height/2-self.winHeight/2)
        self.overrideredirect(1)
        self.geometry(geo)

        self.hyp_delay = StringVar(self)
        self.hyp_delay.set(prefdict['hyp_delay'])
            
        self.hyp_opacity = StringVar(self)
        self.hyp_opacity.set(prefdict['hyp_opacity'])

        self.hyp_able = IntVar(self)
        self.hyp_able.set(int(prefdict['hyp_able']))

        self.UseActionMenu = IntVar(self.master)
        self.UseActionMenu.set(int(prefdict['actionmenu']))

        self.hyp_pinup = IntVar(self)
        self.hyp_pinup.set(int(prefdict['hyp_pinup']))

        self.delold = IntVar(self)
        self.delold.set(int(prefdict['delold']))

        self.UseHSBackground = IntVar(self)
        self.UseHSBackground.set(prefdict['UseHSBackground'])

        self.Old_UseHSBackground = int(prefdict['UseHSBackground'])
        if self.UseHSBackground.get() == 1:
                self.HandleOSBackground(1)
                    
        self.Alphabet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']

        self.convfolder = StringVar(self)
        self.convfolder.set('                                              ')

        self.s_rulename = StringVar(self)
        self.s_rulename.set(prefdict['s_rulename'])

        self.hyp_gfile = StringVar(self)
        self.hyp_gfile_var = int(prefdict['hyp_gfile_var'])

        self.AllCharList = ['No Images']
        self.hyp_folders = hyp_folders

        try:
            self.hyp_gfile.set(hyp_folders[self.hyp_gfile_var])
        except IndexError:
            self.hyp_gfile.set(hyp_folders[0])
                    
        self.conv_hyp_folders = hyp_folders

        self.textwidth = StringVar()
        self.textwidth.set(self.width)
            
        self.textheight = StringVar()
        self.textheight.set(self.height)

        self.OverlayActive = False
        self.Editting = False
        self.RulesOkay = False
        self.ActionMenuOpen = False
        self.WSActive = False

        self.rulesets = []
        Filepath = path.abspath('Resources/Games')
        for filename in iglob(Filepath+'/*/', recursive=True):
            filename = filename.split(Filepath)[-1].replace('\\','')
            self.rulesets.append(filename)

    def LoadPreDict(self):
        Filepath = path.abspath('Resources/Games/'+self.s_rulename.get()+'/Preferences.txt')
        for i in glob(Filepath):
            self.ResetPrefDict(self.GenUserPrefs(i))

    def GenUserPrefs(self, PrefFilePath='Resources/Preferences.txt'):
        try:
            with open(PrefFilePath, 'r') as f:
                lines = f.read().split('\n')
            prefdict = {}
            for line in lines:
                key, sep, value = line.partition(':')
                prefdict[key] = value
        except FileNotFoundError:
            prefdict = PREFDICT_PRESET

        return prefdict

    def ResetPrefDict(self,prefdict):
        print('resetting vars')
        self.background_select_var = int(prefdict['background_select_var'])
        self.background_select.set(self.background_list[self.background_select_var])
        self.hyp_delay.set(prefdict['hyp_delay'])
        self.hyp_opacity.set(prefdict['hyp_opacity'])
        self.hyp_able.set(int(prefdict['hyp_able']))
        self.hyp_pinup.set(int(prefdict['hyp_pinup']))
        self.UseActionMenu.set(int(prefdict['actionmenu']))
        self.delold.set(int(prefdict['delold']))
        self.UseHSBackground.set(prefdict['UseHSBackground'])		
        self.Old_UseHSBackground = int(prefdict['UseHSBackground'])
        self.s_rulename.set(prefdict['s_rulename'])
        self.hyp_gfile_var = int(prefdict['hyp_gfile_var'])

    def SetupMenu(self):
        def StartMoveMM(event):
            self.MMy = event.y
            self.MMx = event.x
        def StopMoveMM(event):
            self.MMy = None
        def MainMenuOnMotion(event):
            deltax = event.x - self.MMx
            deltay = event.y - self.MMy
            x = self.winfo_x() + deltax
            y = self.winfo_y() + deltay
            self.geometry("+%s+%s" % (x, y))
                    

        self.wm_attributes("-transparentcolor", 'white')
        self.wm_attributes("-topmost", 1)
        self.frame = Frame(self, width=1000, height=1000,#height=270,
                                            borderwidth=2, bg='white', relief=RAISED)
        self.frame.grid(row=0,column=0)
        self.bg = Label(self.frame, bg='gray40', width=5, height=30, anchor=E)
        self.bg.place(x=0, y=20)
        self.bg.grip = Label(self.frame, height=1, bg='Gray80', text="<Move>", font=('Times', 8))
        self.bg.grip.place(x=0,y=0)
        self.bg.grip.bind("<ButtonPress-1>", StartMoveMM)
        self.bg.grip.bind("<ButtonRelease-1>", StopMoveMM)
        self.bg.grip.bind("<B1-Motion>", MainMenuOnMotion)
        self.BtnHypno = Button(self.bg,text="Start",width=5,command=self.LaunchHypno)
        self.BtnHypno.grid(row=1, column=0)
        self.BtnEdit = Button(self.bg, text="Edit",width=5,command=self.EditHypno)
        self.BtnEdit.grid(row=2, column=0)
        self.BtnQuit = Button(self.bg, text="Quit",width=5,command=self.Shutdown)
        self.BtnQuit.grid(row=3, column=0)
        self.config(takefocus=1)
            
    def LaunchHypno(self):
        try:
            if self.Editting == True:
                self.DestroyActions()
            if self.OverlayActive == False:
                self.BtnHypno.config(text='End')
                self.OverlayActive = True
                self.RulesOkay = False
                while self.c_hypno.poll() == True:
                    self.c_hypno.recv()
                                    
                self.delay = int(self.hyp_delay.get())
                self.opacity = int(self.hyp_opacity.get())
                self.hypno = self.hyp_able.get()
                                            
                self.pinup = self.hyp_pinup.get()
                self.globfile = path.abspath(self.hyp_gfile.get())
                self.s_rulename = self.s_rulename.get()
                            
                self.gifset = self.background_select.get().replace('.gif','')

                Hypno.launch(self.delay, self.opacity, self.OverlayActive, self.pinup, self.globfile, self.s_rulename,
                             self.gifset, self.c_vid, self.c_pinup, self.c_hypno)
                            
                self.EstablishRules()
            else:
                self.DestroyActions()
        except Exception as e:
            print(e)

    def EstablishRules(self):
        def GenButtonLines(rulefilename):
                Filepath = path.abspath('Resources/Games/'+rulefilename)
                with open(Filepath, 'r') as f:
                    self.templines = f.readlines()
                    if '.jpg' in self.templines[0].replace('\n',''):
                        icon = self.templines[0].replace('\n','')
                        self.templines.remove(self.templines[0])
                    else:
                        icon = ''

                    return icon, self.templines, cycle(self.templines)	
            
        try:
            print("checking rules")
            if self.RulesOkay == False:
                print("rules ok")
                iterlist = []
              
                print(self.UseActionMenu.get())
                if self.UseActionMenu.get() == 1:
                    print("Setting actions")
                    try:
                        FirstIconFound = False
                        self.IconA, ActionListA, self.ActionCycleA = GenButtonLines(self.s_rulename.get()+'/ButtonA.txt')
                        FirstIconFound = True
                        iterlist.append(ActionListA)
                        self.IconB, ActionListB, self.ActionCycleB = GenButtonLines(self.s_rulename.get()+'/ButtonB.txt')
                        iterlist.append(ActionListB)
                        self.IconC, ActionListC, self.ActionCycleC = GenButtonLines(self.s_rulename.get()+'/ButtonC.txt')
                        iterlist.append(ActionListC)
                        self.IconD, ActionListD, self.ActionCycleD = GenButtonLines(self.s_rulename.get()+'/ButtonD.txt')
                        iterlist.append(ActionListD)
                        self.IconE, ActionListE, self.ActionCycleE = GenButtonLines(self.s_rulename.get()+'/ButtonE.txt')
                        iterlist.append(ActionListE)
                        self.IconF, ActionListF, self.ActionCycleF = GenButtonLines(self.s_rulename.get()+'/ButtonF.txt')
                        iterlist.append(ActionListF)
                        self.IconG, ActionListG, self.ActionCycleG = GenButtonLines(self.s_rulename.get()+'/ButtonG.txt')
                        iterlist.append(ActionListG)
                        self.IconH, ActionListH, self.ActionCycleH = GenButtonLines(self.s_rulename.get()+'/ButtonH.txt')
                        iterlist.append(ActionListH)
                        self.IconI, ActionListI, self.ActionCycleI = GenButtonLines(self.s_rulename.get()+'/ButtonI.txt')
                        iterlist.append(ActionListI)
                        self.IconJ, ActionListJ, self.ActionCycleJ = GenButtonLines(self.s_rulename.get()+'/ButtonJ.txt')
                        iterlist.append(ActionListJ)
                    except FileNotFoundError as e:
                            print(e)
                    except Exception as e:
                            print(e)
                            
                self.RulesOkay = True
                self.BuildActionMenu()
        except Exception as e:
            print(e)

    def EditHypno(self):
        def UpdateEditMenu():
            try:
                if self.Editting == True:
                    if self.gifcanvas:
                        self.gifcanvas.itemconfig(self.gifpreview, image=next(self.GifCycle))
                    
                    if self.c_hypno.poll() == False:
                        self.p_hypno.send(True)
                                            
                    self.after(25, UpdateEditMenu)
            except AttributeError as e:
                print('Gif Preview failed.')
                print('Please configure the Width and Height and press Format Gifs ')
            except Exception as e:
                print(e)

        def CloseEditMenu():
                try:
                    self.Editting = False
                    self.EditMenu.destroy()
                except Exception as e:
                    print(e)

        try:
            if self.OverlayActive == False and self.Editting == False:
                self.Editting = True
                if self.c_hypno.poll() == True:
                    self.c_hypno.recv()
                                    
                self.EditMenu = Toplevel()
                self.EditMenu.title("Settings")
                self.EditMenu.overrideredirect(False)
                width, height = 1200, 500
                x = (self.width  / 2) - (width  / 2)
                y = (self.height / 2) - (height / 2)
                self.EditMenu.geometry('%dx%d+%d+%d'%(width, height, x, y))
                            
                self.note = Notebook(self.EditMenu)
                            
                self.tab1 = Frame(self.note, width=width-100, height=height-100, borderwidth=0, relief=RAISED)
                self.SetupTab1()
                self.note.add(self.tab1, text = "Background and Pinup")
                self.note.place(x=50,y=0)

                            
                # Exit button
                Button(self.EditMenu, text="Save/Close", command=CloseEditMenu).place(x=width-115,y=height-65)
                self.after(25, UpdateEditMenu)
            else:
                if self.OverlayActive:
                    self.after(25, self.EditHypno)
                self.DestroyActions()
        except Exception as e:
            print(e)
            
    def SetupTab1(self):
        def GenGifCycle(event=None):
            self.gifcyclist = []
            i = self.background_select.get().replace('.gif','').split('\\')[-1]
            Filepath = 'Resources/Gifs/'+i+'/*.gif'
            for myimage in sorted(glob(Filepath, recursive=True)):
                self.gifcyclist.append(ImageTk.PhotoImage(Image.open(myimage).resize((250, 250), Image.LANCZOS)))
            self.GifCycle = cycle(self.gifcyclist)
                    
        def HandleImgConvert():
            print('Converting Images...')
            folder = self.convfolder.get()
            if folder == 'All':
                for folder in self.hyp_folders:
                    if not folder == 'All':
                        self.ConvertImg(folder,self.delold.get(), self.width,self.height)
            else:
                self.ConvertImg(folder,self.delold.get(), self.width,self.height)
            print('Done.')
                    
        def FormatGifs():
            print('Formatting Gifs...')
            mywidth,myheight = int(self.textwidth.get()),int(self.textheight.get())
            Filepath = path.abspath('Resources/Background Gif Original')
            self.ExtractFrames(mywidth, myheight, Filepath)
            self.background_list = self.GenBGList()
            print('Done.')
            
        # #Handle Gif# #
        self.handlegif = Canvas(self.tab1, bg='gray50',width=250, height=300)
        self.handlegif.place(x=15,y=5)
            
        OptionMenu(self.handlegif, self.background_select, *self.background_list,command=GenGifCycle).place(x=15,y=7)
            
        self.gifcanvas = Canvas(self.handlegif, width=250, height=250, bg='gray75')
        self.gifcanvas.place(x=0,y=50)
            
        GenGifCycle()
            
        try:
            self.gifpreview = self.gifcanvas.create_image(1,1,image=next(self.GifCycle), anchor=NW)
        except StopIteration:
            Label(self.gifcanvas,width=25,font=('Times',12),text='!Format Gif Before Running',anchor=W,bg='gray75').place(x=0,y=0)

        # #Format gif# #
        self.bggif = Label(self.tab1, bg='gray50',width=300)
        self.bggif.place(x=870,y=240)
            
        Label(self.bggif, width=6, font=('Times', 12),text='Width', anchor=W).grid(row=0,column=0, pady=2)
        Label(self.bggif, width=6, font=('Times', 12),text='Height', anchor=W).grid(row=1,column=0, pady=2)
            
        Entry(self.bggif, width=5, borderwidth=0, font=('Times', 14), textvariable=self.textwidth).grid(row=0,column=1)
        Entry(self.bggif, width=5, borderwidth=0, font=('Times', 14), textvariable=self.textheight).grid(row=1,column=1)
            
        Button(self.bggif, text="Format Gifs", command=FormatGifs).grid(row=1,column=2)

        #  Convert png #
        self.bgConvert = Label(self.tab1, bg='gray50',width=30,height=4)
        self.bgConvert.place(x=500,y=240)
        Button(self.bgConvert, text="Convert jpg to png", command=HandleImgConvert).place(x=1,y=1)
        Checkbutton(self.bgConvert, text="Delete jpgs", variable=self.delold).place(x=138,y=2)
        OptionMenu(self.bgConvert, self.convfolder, *self.conv_hyp_folders).place(x=1,y=35)

        # ############ #
        Message(self.tab1, text='Overlay Opacity (0-100)').place(x=300,y=25)
        Message(self.tab1, text='Cycle Delay').place(x=300,y=75)
        Message(self.tab1, text='Image Folder').place(x=300,y=125)
            
        Entry(self.tab1, textvariable=self.hyp_opacity).place(x=400,y=30)
        OptionMenu(self.tab1, self.hyp_delay, '250', '500', '1000', '1500', '3000').place(x=400,y=75)
        OptionMenu(self.tab1, self.hyp_gfile, *self.hyp_folders).place(x=400,y=125)
            
        Radiobutton(self.tab1, text="None", variable=self.hyp_able,value=0).place(x=300,y=200)
        Radiobutton(self.tab1, text="Hypno Background", variable=self.hyp_able,value=1).place(x=300,y=225)
        Radiobutton(self.tab1, text="Turbo Hypno", variable=self.hyp_able,value=2).place(x=300,y=250)
            
        self.EnablePinups = Checkbutton(self.tab1, text="Enable Pinups", variable=self.hyp_pinup)
        self.EnablePinups.place(x=300,y=275)	# has to be an object
            
        Checkbutton(self.tab1, text="Healslut Desktop Background", variable=self.UseHSBackground).place(x=800,y=50)
        Checkbutton(self.tab1, text="Use ActionMenu", variable=self.UseActionMenu).place(x=800,y=75)

        Message(self.tab1, text='Rule set').place(x=800,y=125)
        Button(self.tab1, text="Load Premade Rules", command=self.LoadPreDict).place(x=900,y=125)
        OptionMenu(self.tab1, self.s_rulename, *self.rulesets).place(x=800,y=150)

    def BuildActionMenu(self):
        def GenButtonImage(filename):
            Filepath = path.abspath('Resources\\Buttonlabels\\'+filename)
            tempphoto = Image.open(Filepath)
            tempphoto = tempphoto.resize((50, 50), Image.LANCZOS)
            return ImageTk.PhotoImage(tempphoto)

        def StartMoveAM(event):
            self.AMx, self.AMy = event.x, event.y

        def StopMoveAM(event):
            self.AMx,self.AMy = None, None

        def ActionMenuOnMotion(event):
            x = self.ActionMenu.winfo_x() + event.x - self.AMx
            y = self.ActionMenu.winfo_y() + event.y - self.AMy
            self.ActionMenu.geometry("+%s+%s" % (x, y))


        if self.ActionMenuOpen == False and self.UseActionMenu.get() == 1:
            Filepath = path.abspath('Resources/Games/'+self.s_rulename.get()+'/ButtonA.txt')
            if path.isfile(Filepath):
                self.ActionMenuOpen = True
                self.ActionMenu = Toplevel(self, bg='white', highlightthickness=0)
                self.ActionMenu.overrideredirect(True)
                self.ActionMenu.wm_attributes("-topmost", 1)
                self.ActionMenu.wm_attributes("-transparentcolor", 'white')
                self.ActionFrame = Frame(self.ActionMenu, width=50, height=50, bg='white', borderwidth=0, relief=RAISED,)
                self.ActionFrame.grid(row=0,column=0)

                # Movement grip
                self.grip = Label(self.ActionFrame, height=1, bg='Gray50', text="< Hold to Move >", font=('Times', 8))
                self.grip.grid(row=0,column=0,columnspan=2)
                self.grip.bind("<ButtonPress-1>", StartMoveAM)
                self.grip.bind("<ButtonRelease-1>", StopMoveAM)
                self.grip.bind("<B1-Motion>", ActionMenuOnMotion)
                            
                x = (self.width -300)
                y = (self.height *.4)
                self.ActionMenu.geometry('%dx%d+%d+%d' % (115, 310, x, y))
                
                try:
                    self.imageA = GenButtonImage(self.IconA)
                    print(self.IconA)
                    self.ButtonA = Button(self.ActionFrame, image=self.imageA, text="A",
                            command=partial(self.HandleCycles,self.ActionCycleA)).grid(row=1,column=0,sticky=W+E+N+S)
                    self.imageB = GenButtonImage(self.IconB)
                    self.ButtonB = Button(self.ActionFrame, image=self.imageB, text="A",
                            command=partial(self.HandleCycles,self.ActionCycleB)).grid(row=1,column=1,sticky=W+E+N+S)
                    self.imageC = GenButtonImage(self.IconC)
                    self.ButtonC = Button(self.ActionFrame, image=self.imageC, text="A",
                            command=partial(self.HandleCycles,self.ActionCycleC)).grid(row=2,column=0,sticky=W+E+N+S)
                    self.imageD = GenButtonImage(self.IconD)
                    self.ButtonD = Button(self.ActionFrame, image=self.imageD, text="A",
                            command=partial(self.HandleCycles,self.ActionCycleD)).grid(row=2,column=1,sticky=W+E+N+S)
                    self.imageE = GenButtonImage(self.IconE)
                    self.ButtonE = Button(self.ActionFrame, image=self.imageE, text="A",
                            command=partial(self.HandleCycles,self.ActionCycleE)).grid(row=3,column=0,sticky=W+E+N+S)
                    self.imageF = GenButtonImage(self.IconF)
                    self.ButtonF = Button(self.ActionFrame, image=self.imageF, text="A",
                            command=partial(self.HandleCycles,self.ActionCycleF)).grid(row=3,column=1,sticky=W+E+N+S)
                    self.imageG = GenButtonImage(self.IconG)
                    self.ButtonG = Button(self.ActionFrame, image=self.imageG, text="A",
                            command=partial(self.HandleCycles,self.ActionCycleG)).grid(row=4,column=0,sticky=W+E+N+S)
                    self.imageH = GenButtonImage(self.IconH)
                    self.ButtonH = Button(self.ActionFrame, image=self.imageH, text="A",
                            command=partial(self.HandleCycles,self.ActionCycleH)).grid(row=4,column=1,sticky=W+E+N+S)
                    self.imageI = GenButtonImage(self.IconI)
                    self.ButtonI = Button(self.ActionFrame, image=self.imageI, text="A",
                            command=partial(self.HandleCycles,self.ActionCycleI)).grid(row=5,column=0,sticky=W+E+N+S)
                    self.imageJ = GenButtonImage(self.IconJ)
                    self.ButtonJ = Button(self.ActionFrame, image=self.imageJ, text="A",
                            command=partial(self.HandleCycles,self.ActionCycleJ)).grid(row=5,column=1,sticky=W+E+N+S)
                except Exception as e:
                    pass

    def HandleCycles(self,mycycle):
        def do_macro(macro):
            if '$playvideo' in macro:
                file = macro.replace('$playvideo ','')
                Filepath = path.abspath('Resources/Video/'+file)
                self.p_vid.send(Filepath)
            if '$pinup' in macro:
                Filepath = path.abspath('Resources/Images/'+macro.replace('$pinup ',''))+'/'
                self.p_pinup.send(Filepath)
        try:
            cyc = next(mycycle)
            line = str(cyc).split(',')
            for macro in line:
                do_macro(macro.replace('\n',''))
        except Exception as e:
            print(e)

    def GenFolders(self):
        hyp_folders = glob('Resources\\Images\\*\\', recursive=True)
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
                windll.user32.SystemParametersInfoW(20, 0,  path.abspath("Resources/Desktop Backgrounds/Overlay Background.png") , 0)
            elif Value == 'Exit':
                windll.user32.SystemParametersInfoW(20, 0,  path.abspath("Resources/Desktop Backgrounds/Background.png") , 0)
        except Exception as e:
            print('Desktop Background Setup Failed.')
            print(e)

    
    def DestroyActions(self,Exit=False):
        if self.Editting or self.OverlayActive or Exit:
            self.Editting = False
            self.ActionMenuOpen = False
            self.WSActive = False
            self.OverlayActive = False
            self.BtnHypno.config(text='Start')
            if self.c_hypno.poll() == False:
                self.p_hypno.send(True)
            try:
                self.ActionMenu.destroy()
            except AttributeError:
                pass
            try:
                self.WSFrame.destroy()
            except AttributeError:
                pass
            try:
                self.EditMenu.destroy()
            except AttributeError:
                pass
    
    def SavePref(self):
        self.hyp_gfile_var = 0
        for item in self.hyp_folders:
            if item == self.hyp_gfile.get():
                break
            self.hyp_gfile_var +=1

        self.background_select_var = 0
        for item in self.background_list:
            if item == self.background_select.get():
                break
            self.background_select_var +=1

        PrefDictList = [
                'hyp_delay:'+str(self.hyp_delay.get()),
                'hyp_opacity:'+str(self.hyp_opacity.get()),
                'hyp_able:'+str(self.hyp_able.get()),
                'hyp_pinup:'+str(self.hyp_pinup.get()),
                'delold:'+str(self.delold.get()),
                'hyp_gfile_var:'+str(self.hyp_gfile_var),
                'actionmenu:'+str(self.UseActionMenu.get()),
                'background_select_var:'+str(self.background_select_var),
                's_rulename:'+str(self.s_rulename.get()),
                'UseHSBackground:'+str(self.UseHSBackground.get())
                ]
        Filepath = path.abspath('Resources\\Preferences.txt')
        with open(Filepath, 'w') as f:
            for line in PrefDictList:
                f.write(line+'\n')	

    def Shutdown(self):				
        try:
            self.SavePref()
            if self.Old_UseHSBackground == 1 or self.UseHSBackground.get() == 1:
                self.HandleOSBackground('Exit')
            self.DestroyActions(True)
            self.quit()
        except Exception as e:
            print(e)

if __name__ == '__main__':
    app = App()
    app.mainloop()
    app.destroy()