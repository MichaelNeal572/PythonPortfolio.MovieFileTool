from tkinter import *
from tkinter import filedialog, ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from PIL import ImageTk, Image
import tkinter
import os
import re
import sys
import configparser
import ctypes
from ctypes import POINTER, Structure, c_wchar, c_int, sizeof, byref
from ctypes.wintypes import BYTE, WORD, DWORD, LPWSTR, LPSTR
import win32api 
import shutil
import win32con, win32api

### METHODS START

def set_folder_icon():
    
    if(wd.target_ico_full!=''):    
        if(os.path.join(wd.target_ico_dir, wd.target_ico_base)!=os.path.join(wd.target_folder, wd.target_ico_base)):
            shutil.copy(wd.target_ico_full, os.path.join(wd.target_folder, wd.target_ico_base))
        
        wd.target_ico_dir=wd.target_folder
        wd.target_ico_full=os.path.join(wd.target_folder, wd.target_ico_base)
        
        SetFolderIcon(wd.target_folder,wd.target_ico_full)
        
        if(os.path.isfile('desktop.ini')):
            os.remove('desktop.ini')
            
        config = configparser.ConfigParser()
        config.optionxform=str
        config['.ShellClassInfo'] = {'IconResource' : f'{wd.target_ico_base},0'}
        config['ViewState'] = {'Mode' : '',
                              'Vid' : '',
                              'FolderType' : 'Videos'}
        with open('desktop.ini', 'w') as configfile:
            config.write(configfile)
            win32api.SetFileAttributes('desktop.ini', win32con.FILE_ATTRIBUTE_HIDDEN)
 

   
            
##widget click handler     
def click(key):
    ##button click handler.
    #button methods
    if (key=="Open folder" or key=="Open..."):
        wd.select_folder()
    elif (key=="Clean file Names"):
        rt.name_cleaner()
    elif(key=="Select .ico File"):
        wd.select_icon_file()
    elif(key=='Rename Icon'):
        rt.rename_icon()
    elif(key=="Set Folder Icon"):
        set_folder_icon()
    elif(key=="Mass Rename"):
        rt.mass_file_rename()
    elif(key=="Tag Icon"):
        rt.tag_icon()
    #menu item methods
    elif(key=="Quit"):
        sys.exit()   
##

###METHODS END

##global variables


folder_option_list = ['Open folder', 'Clean file Names','Rename_label','Rename_Input' , 'Mass Rename'  ]
ico_option_list = ['Select .ico File','Rename_label','Rename_Input','Rename Icon','Tag Icon','Set Folder Icon']

menu_file_list = ['Open...','NULL','Quit']
menu_help_list = ['Help']

movie_filetypes = ['.mkv','.mp4','.avi','.flv',
                   '.mpeg','.ogg','.ogm','.mov']

anime_tags_list = sorted(['Action','Adventure','Cars','Cartoon','Comedy','Dementia','Demons'
                   ,'Drama','Dub' ,'Ecchi' ,'Fantasy' ,'Game' ,'Harem' ,'Historical' 
                   ,'Horror' ,'Josei','Kids' ,'Magic' ,'Martial Arts'  ,'Mecha' ,'Military' 
                    ,'Movie' ,'Music','Mystery' ,'ONA','OVA','Parody','Police'
                    ,'Psychological' ,'Romance' ,'Samurai' ,'School','Sci-Fi','Seinen'  ,'Shoujo' 
                    ,'Shoujo Ai','Shounen','Shounen Ai','Slice of Life','Space','Special','Sports' 
                    ,'Super Power','Supernatural','Thriller','Vampire','Yuri'])

western_tags_list = sorted(['Action','Adult','Adventure','Animation / Anime',
                     'Biopic','Childrens','Comedy','Crime / detective /spy',
                     'Documentary','Drama','Horror','Family',
                     'Fantasy','Historical','Medical','Musical',
                     'Paranormal','Romance','Sport','Science fiction',
                     'Talk Show','Thriller / Suspense','War','Western'])
##

##Classes
class Working_Directory:
    program_directory=''
    target_ico_full=''
    target_ico_base=''
    target_ico_dir=''
    target_folder=''
    
    def __init__(self):
        program_directory = os.getcwd()
        Working_Directory.target_ico_full = ''
        Working_Directory.target_ico_base = ''
        Working_Directory.target_ico_dir = ''
        Working_Directory.target_folder = program_directory
        
    def select_folder(self):
        # show an "Open" dialog box and return the path to the selected folder
        target_folder = filedialog.askdirectory()
        os.chdir(target_folder)
        window.dir_display_pad.folder_display_update()
    
    def select_icon_file(self):
        # show an "Open" dialog box and return the path to the selected file
        Working_Directory.target_ico_full = filedialog.askopenfilename() 
        Working_Directory.target_ico_base = os.path.basename(Working_Directory.target_ico_full)
        Working_Directory.target_ico_dir = os.path.dirname(Working_Directory.target_ico_full)
        if(os.path.splitext(Working_Directory.target_ico_base)[1]!='.ico'):
            messagebox.showwarning(
            "Open file",
            "Please select a .ico file"
            )
            Working_Directory.target_ico_full = ''
            Working_Directory.target_ico_base = ''
            Working_Directory.target_ico_dir = ''
        window.right_pad.update_icon_display(window.right_pad.icon_display_pad)
    
class Renaming_Tool:        
    def name_cleaner(self):
        #searches for everything between brackets and drops it from the file name
        #if the file name extension matches the supported media file types
        for f in os.listdir():
            f_name, f_ext = os.path.splitext(f)
            if f_ext in movie_filetypes:
                regex_temp=r'\s*[\{\[\(]+[^\}\]\)]*[\}\]\)]'
                #matches sets of brackets and their contents making sure not to match 
                #the start of one set of brackets with the end of a seperate set 
                #effectively ignoring the end of the first tag if another end bracket is found
                new_name = re.sub(regex_temp,'',f_name)
                new_name = re.sub(regex_temp,'',f_name)
                new_name = new_name.replace("_"," ")
                new_name = new_name.replace("."," ").strip()
                new_name = f'{new_name}{f_ext}'
                os.rename(f, new_name)
        window.dir_display_pad.folder_display_update()        
    
    def mass_file_rename(self):
        if(wd.target_folder!=''):
            input_name=window.left_pad.rename_file_input.get("1.0",END).strip()
            if input_name[:4:-1][::-1] not in movie_filetypes:
                for f in os.listdir():
                    f_name, f_ext = os.path.splitext(f)
                    ep_num=re.findall(r'\b\d+\b', f_name)
                    if f_ext in movie_filetypes:
                        new_name = f'{input_name} - {ep_num[0]}{f_ext}'
                        os.rename(f, new_name)
            else:
                messagebox.showwarning(
                "Rename",
                "Please enter new name without file extensions"
                )
            window.dir_display_pad.folder_display_update()
        
    def rename_icon(self):
        
        f_name, f_ext = os.path.splitext(wd.target_ico_base)
        if(wd.target_ico_full!=''):
            new_name=window.left_pad.rename_ico_input.get("1.0",END).strip()
            #if an icon previously had tags applied to it they will be preserved
            regex_temp=r'[\{\[\(]+[^\}\]\)]*[\}\]\)]'
            if (re.search(regex_temp, wd.target_ico_base)):
                tags=re.findall(regex_temp, wd.target_ico_base)
                for t in tags:
                    new_name=f'{new_name} {t}'
            new_name = f'{new_name}{f_ext}'
            new_name = os.path.join(wd.target_ico_dir, new_name)
            os.rename(wd.target_ico_full, new_name)
            wd.target_ico_full=new_name
            wd.target_ico_dir=os.path.dirname(wd.target_ico_full)
            wd.target_ico_base=os.path.basename(wd.target_ico_full)
            window.right_pad.update_icon_display(window.right_pad.icon_display_pad)
            window.dir_display_pad.folder_display_update()
        
    def tag_icon(self):
        
        f_name, f_ext = os.path.splitext(wd.target_ico_base)
        if(wd.target_ico_full!=''):
            regex_temp=r'[\{\[\(]+[^\}\]\)]*[\}\]\)]'
            new_name = re.sub(regex_temp,'',f_name).strip()
            tags=[]
            for t in CheckBox.boxes:
                if(t.state.get()):
                   if(t.label not in tags):
                       tags.append(t.label)
            tags=sorted(tags)
            for t in tags:
                new_name=f'{new_name} [{t}]'
            new_name=f'{new_name}{f_ext}'
            new_name =os.path.join(wd.target_ico_dir, new_name)
            os.rename(wd.target_ico_full, new_name)
            wd.target_ico_full=new_name
            wd.target_ico_dir=os.path.dirname(wd.target_ico_full)
            wd.target_ico_base=os.path.basename(wd.target_ico_full)
            window.right_pad.update_icon_display(window.right_pad.icon_display_pad)
            window.dir_display_pad.folder_display_update()
            
        for t in CheckBox.boxes:
            t.state.set(False)
            
class SetFolderIcon(object):
    #complex looking class that was imported for the sole purpose of making windows recognize
    #the new desktop.ini file. the desktop.ini file will be overwritten later however
    #without creating a desktop.ini file in the folder at least once through this method
    #windows will not read the file thus leaving the folder icon unchanged
    def __init__(self, target_folder, target_ico_full, reset=False):
        assert os.path.isdir(target_folder), "folderPath '%s' is not a valid folder"%target_folder
        self.__folderPath = target_folder
        assert os.path.isfile(target_ico_full), "iconPath '%s' does not exist"%target_ico_full
        self.__iconPath = target_ico_full
        assert isinstance(reset, bool), "reset must be boolean"
        self.__reset = reset
        # set icon if system is windows
        if os.name == 'nt':
            try:
                self.__set_icon_on_windows()
            except Exception as e:
                warnings.warn("Unable to set folder icon (%s)"%e)
        elif os.name == 'posix':
            raise Exception('posix system not implemented yet')
        elif os.name == 'mac':
            raise Exception('mac system not implemented yet')
        elif os.name == 'os2':
            raise Exception('os2 system not implemented yet')
        elif os.name == 'ce':
            raise Exception('ce system not implemented yet')
        elif os.name == 'java':
            raise Exception('java system not implemented yet')
        elif os.name == 'riscos':
            raise Exception('riscos system not implemented yet')

    def __set_icon_on_windows(self):
        HICON = c_int
        LPTSTR = LPWSTR
        TCHAR = c_wchar
        MAX_PATH = 260
        FCSM_ICONFILE = 0x00000010
        FCS_FORCEWRITE = 0x00000002
        SHGFI_ICONLOCATION = 0x000001000    

        class GUID(Structure):
            _fields_ = [ ('Data1', DWORD),
                         ('Data2', WORD),
                         ('Data3', WORD),
                         ('Data4', BYTE * 8) ]
        class SHFOLDERCUSTOMSETTINGS(Structure):
            _fields_ = [ ('dwSize', DWORD),
                         ('dwMask', DWORD),
                         ('pvid', POINTER(GUID)),
                         ('pszWebViewTemplate', LPTSTR),
                         ('cchWebViewTemplate', DWORD),
                         ('pszWebViewTemplateVersion', LPTSTR),
                         ('pszInfoTip', LPTSTR),
                         ('cchInfoTip', DWORD),
                         ('pclsid', POINTER(GUID)),
                         ('dwFlags', DWORD),
                         ('pszIconFile', LPTSTR),
                         ('cchIconFile', DWORD),
                         ('iIconIndex', c_int),
                         ('pszLogo', LPTSTR),
                         ('cchLogo', DWORD) ]
        class SHFILEINFO(Structure):
            _fields_ = [ ('hIcon', HICON),
                         ('iIcon', c_int),
                         ('dwAttributes', DWORD),
                         ('szDisplayName', TCHAR * MAX_PATH),
                         ('szTypeName', TCHAR * 80) ]    

        shell32 = ctypes.windll.shell32
        fcs = SHFOLDERCUSTOMSETTINGS()
        fcs.dwSize = sizeof(fcs)
        fcs.dwMask = FCSM_ICONFILE
        fcs.pszIconFile = self.__iconPath
        fcs.cchIconFile = 0
        fcs.iIconIndex = self.__reset 
        hr = shell32.SHGetSetFolderCustomSettings(byref(fcs), self.__folderPath, FCS_FORCEWRITE)
        if hr:
            raise WindowsError(win32api.FormatMessage(hr))

        sfi = SHFILEINFO()
        hr = shell32.SHGetFileInfoW(self.__folderPath, 0, byref(sfi), sizeof(sfi), SHGFI_ICONLOCATION)
        #if hr == 0:
        #    raise WindowsError(win32api.FormatMessage(hr))

        index = shell32.Shell_GetCachedImageIndexW(sfi.szDisplayName, sfi.iIcon, 0)
        shell32.SHUpdateImageW(sfi.szDisplayName, sfi.iIcon, 0, index)

##GUI classes        
class Root(Tk):
    def __init__(self, title):
        super(Root, self).__init__()
        self.title(title)
        self.main_pad = Frame(self, width = 500, height = 25, relief="solid")
        self.main_pad.grid(row=0, column=0 , padx=5, pady=5)
        self.title_pad = Title_Pad(self.main_pad,r=0,c=0)
        ttk.Separator(self.main_pad, orient="horizontal",).grid(row=1, column=0, sticky=(E,W), columnspan=8)
        self.left_pad = Left_Pad(self.main_pad,r=2,c=1)
        ttk.Separator(self.main_pad, orient="vertical").grid(row=2, column=2, sticky=(N,S), padx = 5, rowspan=2)
        self.dir_display_pad= Dir_Display_Pad(self.main_pad,r=2,c=3)
        ttk.Separator(self.main_pad, orient="vertical").grid(row=2, column=4, sticky=(N,S), padx = 5, rowspan=2)
        self.tags_pad = Tags_Pad(self.main_pad,r=2,c=5)
        ttk.Separator(self.main_pad, orient="vertical").grid(row=2, column=6, sticky=(N,S), padx = 5, rowspan=2)
        self.right_pad = Right_Pad(self.main_pad,r=2,c=7)
        
        self.menu_bar = Menu(self)
        self.config(menu=self.menu_bar)
        self.menu_file = Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label='File', menu=self.menu_file)
        for btn_text in menu_file_list:
            if (btn_text!='NULL'):   
                def cmd(x=btn_text):
                    click(x)
                self.menu_file.add_command(label=btn_text, command=cmd)
            else:
                self.menu_file.add_separator()

class Title_Pad:
    def __init__(self, master,r,c):
        self.title_pad = Frame(master, borderwidth=3)
        self.title_pad.grid(row=r, column=c, columnspan=8, sticky=(N), padx=10)
        Label(self.title_pad, font = ('helvetica', 21,'bold'), text='Movie File Tool').grid(row=0, column=0, columnspan=2)
        
        Label(self.title_pad, text="Current Directory:").grid(row=2, column=0, sticky=W, pady=10)
        self.folder_dir_display = Text(self.title_pad, height=1, width=125)
        self.folder_dir_display.grid(row=2,column=1, sticky=W)
        self.folder_dir_display.configure(state="disabled")
        self.folder_dir_display_update()
    
    def folder_dir_display_update(self):
        self.folder_dir_display.configure(state="normal")
        self.folder_dir_display.delete(1.0,END)
        self.folder_dir_display.insert(END,f'{wd.target_folder}')
        self.folder_dir_display.configure(state="disabled")
    
    
class Left_Pad:
    def __init__(self, master,r,c):
        self.left_pad = LabelFrame(master, text="Options", padx=2, pady=2)
        self.left_pad.grid(row=r, column=c, columnspan=1, padx=5, pady=5, sticky=(N,W,S))
        
        self.folder_options_pad = LabelFrame(self.left_pad, text="Folder:", padx=2, pady=2)
        self.folder_options_pad.grid(row=0, column=0, padx=5, pady=5, sticky=(N,W,S))
        r,c=0,0
        for btn_text in folder_option_list:
            if(btn_text=='Rename_Input'):
                self.rename_file_input = Text(self.folder_options_pad, height=2, width=22)
                self.rename_file_input.grid(row=r, column=c ,pady = 5, padx = 5)
            elif(btn_text=='Rename_label'):
                Label(self.folder_options_pad, text="Rename:").grid(row=r, column=c, sticky=W)
            else:
                def cmd(x=btn_text):
                        click(x)
                Button(self.folder_options_pad, text=btn_text, width=25, command=cmd).grid(row=r, column=c ,pady = 5, padx = 5)
            r=r+1
        
        ttk.Separator(self.left_pad, orient="horizontal").grid(row=1, column=0, pady=5, sticky=(E,W), columnspan=2)
        
        
        self.icon_options_pad = LabelFrame(self.left_pad, text="Icon:", padx=2, pady=2)
        self.icon_options_pad.grid(row=2, column=0, padx=5, pady=5, sticky=(N,W,S))
        r,c=0,0
        for btn_text in ico_option_list:
            if(btn_text=='Rename_Input'):
                self.rename_ico_input = Text(self.icon_options_pad, height=2, width=22)
                self.rename_ico_input.grid(row=r, column=c ,pady = 5, padx = 5)
            elif(btn_text=='Rename_label'):
                Label(self.icon_options_pad, text="Rename:").grid(row=r, column=c, sticky=W)
            else:
                def cmd(x=btn_text):
                        click(x)
                Button(self.icon_options_pad, text=btn_text, width=25, command=cmd).grid(row=r, column=c ,pady = 5, padx = 5)
            r=r+1

class Dir_Display_Pad:
    def __init__(self, master,r,c):
        self.dir_display_pad = LabelFrame(master, text="Directory", padx=2, pady=2)
        self.dir_display_pad.grid(row=r, column=c, columnspan=1, padx=5, pady=5, sticky=(N,W,S))
        
        self.folder_display_pad = Frame(self.dir_display_pad, padx=2, pady=2)
        self.folder_display_pad.grid(row=0,column=0)
        
        
        self.folder_display = ScrolledText(self.folder_display_pad, height=27, width=30)
        self.folder_display.grid(row=3,column=0, padx=10, pady=10)
        self.folder_display.configure(state="disabled")
        self.folder_display_update()
    
    def folder_display_update(self):
        global target_folder
        result=''
        for file in os.listdir():
            result = f'{result}{file}\n'
            
        self.folder_display.configure(state="normal")
        self.folder_display.delete(1.0,END)
        self.folder_display.insert(END,f'{result}')
        self.folder_display.configure(state="disabled")

class Tags_Pad:
    def __init__(self, master,r,c):
        self.tags_pad = LabelFrame(master, text="Genres", padx=2, pady=2)
        self.tags_pad.grid(row=r, column=c, padx=5, pady=5, sticky="ns")
        
        self.genre_type_notebook = ttk.Notebook(self.tags_pad)
        self.frame_anime = ttk.Frame(self.genre_type_notebook)
        self.frame_western = ttk.Frame(self.genre_type_notebook)
        self.genre_type_notebook.add(self.frame_anime, text="Anime")
        self.genre_type_notebook.add(self.frame_western, text="Western")
        self.genre_type_notebook.grid(row=0, column=0, sticky='nsew')
    
        r,c=0,0
        for btn_text in anime_tags_list:
            button=CheckBox(self.frame_anime, text=btn_text)  # Replace Checkbutton
            button.grid(row=r, column=c, sticky=W)
            r=r+1
            if(r==16):
                r=0
                c=c+1
                
        r,c=0,0
        for btn_text in western_tags_list:
            button=CheckBox(self.frame_western, text=btn_text)  # Replace Checkbutton
            button.grid(row=r, column=c, sticky=W)
            r=r+1
            if(r==16):
                r=0
                c=c+1
            
class Right_Pad:
    def __init__(self, master,r,c):
        self.right_pad = LabelFrame(master, text="Icon", padx=2, pady=2)
        self.right_pad.grid(row=r, column=c, columnspan=1,padx=5, pady=5, sticky=(N,S,W))
    
        self.icon_display_pad = Frame(self.right_pad, padx=2, pady=2)
        self.icon_display_pad.grid(row=0,column=0)
        
        self.img = ImageTk.PhotoImage(Image.new(mode="1", size=(256,256)))
        self.img_panel = Label(self.icon_display_pad, image = self.img, relief="sunken")
        self.img_panel.grid(row=0, column=0, columnspan=2, sticky=(N))
        
        Label(self.icon_display_pad, text="Filename:").grid(row=1,column=0,sticky=W)
        
        
        self.icon_dir_display = Text(self.icon_display_pad, height=3, width=30)
        self.icon_dir_display.grid(row=2,column=0)
        self.icon_dir_display.configure(state="disabled")
        self.icon_dir_display_update()
        
        ttk.Separator(self.right_pad, orient="horizontal").grid(row=1, column=0, pady=5, sticky=(E,W), columnspan=2)
        
        icon_options_pad = Frame(self.right_pad, padx=2, pady=2)
        icon_options_pad.grid(row=2,column=0)
        
    def icon_dir_display_update(self):
        self.icon_dir_display.configure(state="normal")
        self.icon_dir_display.delete(1.0,END)
        self.icon_dir_display.insert(END,f'{wd.target_ico_base}')
        self.icon_dir_display.configure(state="disabled")
    
    def update_icon_display(self, icon_display_pad): 
        self.img = ImageTk.PhotoImage(Image.new(mode="1", size=(256,256))) 
        if(wd.target_ico_full!=''):
            self.img = ImageTk.PhotoImage(Image.open(wd.target_ico_full).resize((256,256), Image.ANTIALIAS))
        self.img_panel.image=self.img #image must be anchor with this method or image will display blank
        self.img_panel.configure(image=self.img)
        self.icon_dir_display_update()
        
class CheckBox(tkinter.Checkbutton):
    ''' 
    https://stackoverflow.com/questions/50485891/how-to-get-and-save-checkboxes-names-into-a-list-with-tkinter-python-3-6-5
    '''
    boxes = []  # Storage for all checkboxes

    def __init__(self, master=None, **options):
        tkinter.Checkbutton.__init__(self, master, options)  # Subclass checkbutton to keep other methods
        self.boxes.append(self)
        self.state = tkinter.BooleanVar()  # var used to store checkbox state (on/off)
        self.label = self.cget('text')  # store the text for later
        self.configure(variable=self.state)  # set the checkbox to use our var  
##END GUI classes
        

##

if __name__=='__main__':
    wd=Working_Directory()
    rt=Renaming_Tool()
    window=Root('Movie Folder Tool')
    window.mainloop()



