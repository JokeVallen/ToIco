import PIL.Image as im
import tkinter.filedialog as tf
import tkinter as tk
import tkinter.messagebox as tm
import random
import os
from sys import executable as EXECUTABLE
from win32api import SetFileAttributes
from win32con import FILE_ATTRIBUTE_HIDDEN
from subprocess import Popen, PIPE, STDOUT
import Tool.PackTool as pt
from Tool.ico import ico1
import Tool.MyLogger as ml
from ctypes import windll


class ToIco:
    __filePath = ''
    __foderPath = ''
    __cWidth = 16
    __cHeight = 16
    __sizeMode = 0
    __imgType = '.ico'
    __tips = ('生成的图形文件保存地址为所选PNG或JPG文件所在地址', '可以处理单个文件也可以处理多个文件', '处理图形文件夹时仅处理PNG和JPG文件', '当存在同名图形文件时会覆盖已存在的图形文件')
    __site = [('单个图形文件', 0), ('图形文件夹', 1)]
    __type = [('.ico', 0), ('.wmf', 1)]
    __rotation = [('0.5倍', 0), ('原比例', 1), ('1.5倍', 2), ('2倍', 3), ('3倍', 4), ('5倍', 5)]
    __choice = [('尺寸转换', 0), ('缩放比例', 1), ('自定义宽和高', 2)]
    __isDefineCb = True
    __log = ml.IMyLogger()

    def _windowSize(self, p_width, p_height):
        size = '%dx%d+%d+%d' % (
            p_width, p_height, (self.__screenWidth - p_width) / 2, (self.__screenHeight - p_height) / 2)
        return size

    def _createShortcut(self):
        execpath = os.path.dirname(EXECUTABLE)
        batpath = execpath + "\\init.bat"
        if os.path.exists(batpath):
            self.__log.InLog("Create shortcut successful!")
            startcmd = "start " + "shortcut.bat"
            startpath = execpath + "\\shortcut.bat"
            cmd = "cd " + execpath + " && " + startcmd
            pop = Popen(cmd, shell=True, stdout=PIPE, stderr=STDOUT)
            pop.wait()
            if pop.returncode == 0:
                pop.terminate()
                pop.kill()
                SetFileAttributes(startpath, FILE_ATTRIBUTE_HIDDEN)

    def __init__(self):
        self._createShortcut()
        # 告诉操作系统使用程序自身的dpi适配
        windll.shcore.SetProcessDpiAwareness(1)
        # 获取屏幕的缩放因子
        ScaleFactor = windll.shcore.GetScaleFactorForDevice(0)
        self.__cPIPT = pt.IPackTool(False)
        self.__root = tk.Tk()
        self.__root.title('图形转换工具')
        # 获取屏幕宽高
        self.__screenWidth = self.__root.winfo_screenwidth()
        self.__screenHeight = self.__root.winfo_screenheight()
        # 设置窗口大小
        self.__root.geometry(self._windowSize(int(self.__screenWidth * 0.35), int(self.__screenHeight * 0.625)))
        self.__root.wm_minsize(int(self.__screenWidth * 0.31), int(self.__screenHeight * 0.56))
        self.__cPIPT.IcoBase64Decoder('icoTool', ico1)
        self.__cPIPT.IcoSetToTk(self.__root, 'icoTool')
        # 设置程序缩放
        self.__root.tk.call('tk', 'scaling', ScaleFactor / 75)

        self.__v = tk.IntVar()
        self.__r = tk.IntVar()
        self.__c = tk.IntVar()
        self.__t = tk.IntVar()
        self.__check1 = tk.IntVar()
        self.__check2 = tk.IntVar()
        self.__check3 = tk.IntVar()
        self.__check4 = tk.IntVar()
        self.__check5 = tk.IntVar()
        self.__check6 = tk.IntVar()
        self.__sizeCheck = [('16x16', self.__check1), ('32x32', self.__check2), ('48x48', self.__check3),
                            ('64x64', self.__check4), ('128x128', self.__check5), ('256x256', self.__check6)]

        self.__Fr_m = tk.Frame(self.__root, pady='10px')
        self.__Fr_m.pack(expand='yes')

        self.__Fr_1 = tk.Frame(self.__Fr_m)
        self.__Lb_1 = tk.Label(self.__Fr_1, text='请选择图形文件或图形文件夹', font=('幼圆', 9, 'bold'), fg='red')
        self.__Lb_1.pack(expand='yes', pady='5px')
        for name, index in self.__site:
            Rb = tk.Radiobutton(self.__Fr_1, text=name, font=('幼圆', 9, 'bold'), variable=self.__v, value=index,
                                command=lambda: self._ClearPath())
            Rb.pack(expand='yes')
        self.__Bt_1 = tk.Button(self.__Fr_1, text='选择文件或文件夹', font=('幼圆', 9, 'bold'), width=20,
                                command=lambda: self._ChooseMode(self.__v.get()))
        self.__Bt_1.pack(expand='yes', pady='5px')
        self.__Fr_1_1 = tk.Frame(self.__Fr_1)
        self.__Lb_9 = tk.Label(self.__Fr_1_1, text='请选择待转换图形文件的格式', font=('幼圆', 9, 'bold'), fg='red')
        self.__Lb_9.pack(expand='yes', pady='5px')
        for name, index in self.__type:
            Rb = tk.Radiobutton(self.__Fr_1_1, text=name, font=('幼圆', 9, 'bold'), variable=self.__t, value=index,
                                command=lambda: self._ImageTypeChange(self.__t.get()))
            Rb.pack(expand='yes', padx='5px', side='left')
        self.__Fr_1_1.pack(expand='yes', pady='5px')
        self.__Lb_2 = tk.Label(self.__Fr_1, font=('幼圆', 9, 'bold'), foreground='blue', text='')
        self.__Lb_2.pack(expand='yes', pady='5px')
        for name, index in self.__choice:
            Rb = tk.Radiobutton(self.__Fr_1, text=name, font=('幼圆', 9, 'bold'), variable=self.__c, value=index,
                                command=lambda: self._FrameActive(self.__c.get()))
            Rb.pack(expand='yes', padx='5px', side='left')
        self.__Fr_1.pack(expand='yes', pady='5px')

        self.__Fr_2 = tk.Frame(self.__Fr_m)
        self.__Lb_3 = tk.Label(self.__Fr_2, font=('幼圆', 9, 'bold'), text='请选择转换后的图形文件尺寸', fg='red')
        self.__Lb_3.pack(expand='yes', pady='5px')
        for name, vari in self.__sizeCheck:
            Cb = tk.Checkbutton(self.__Fr_2, text=name, font=('幼圆', 9, 'bold'), variable=vari,
                                command=lambda: self._RandomTip())
            Cb.pack(expand='yes')
            if self.__isDefineCb:
                Cb.select()
                self.__isDefineCb = False
        self.__Fr_2.pack(expand='yes')

        self.__Fr_3 = tk.Frame(self.__Fr_m)
        self.__Lb_4 = tk.Label(self.__Fr_3, text='请选择转换后的图形文件缩放比例', font=('幼圆', 9, 'bold'), fg='red')
        self.__Lb_4.pack(expand='yes', pady='5px')
        for name, index in self.__rotation:
            Rb = tk.Radiobutton(self.__Fr_3, text=name, font=('幼圆', 9, 'bold'), variable=self.__r, value=index,
                                command=lambda: self._RandomTip())
            Rb.pack(expand='yes')
        self.__Fr_3.pack_forget()

        self.__Fr_4 = tk.Frame(self.__Fr_m)
        self.__Lb_5 = tk.Label(self.__Fr_4, text='请输入转换后的图形文件宽和高', font=('幼圆', 9, 'bold'), fg='red')
        self.__Lb_5.pack(expand='yes', pady='5px')
        self.__Fr_4_1 = tk.Frame(self.__Fr_4)
        self.__Lb_5 = tk.Label(self.__Fr_4_1, text='Width:', font=('幼圆', 9, 'bold'))
        self.__Lb_5.pack(expand='yes', side='left')
        self.__Et1_R = self.__root.register(self._CallBack)
        self.__Et_1 = tk.Entry(self.__Fr_4_1, validate='focus', validatecommand=(self.__Et1_R, "%P"))
        self.__Et_1.pack(expand='yes', side='left')
        self.__Fr_4_1.pack(expand='yes', pady='5px')
        self.__Fr_4_2 = tk.Frame(self.__Fr_4)
        self.__Lb_6 = tk.Label(self.__Fr_4_2, text='Height:', font=('幼圆', 9, 'bold'))
        self.__Lb_6.pack(expand='yes', side='left')
        self.__Et2_R = self.__root.register(self._CallBack)
        self.__Et_2 = tk.Entry(self.__Fr_4_2, validate='focus', validatecommand=(self.__Et2_R, "%P"))
        self.__Et_2.pack(expand='yes', side='left')
        self.__Fr_4_2.pack(expand='yes', pady='5px')
        self.__Fr_4.pack_forget()
        self.__Lb_7 = tk.Label(self.__Fr_4, text='', fg='red', font=('幼圆', 9, 'bold'))
        self.__Lb_7.pack_forget()
        self.__Bt_1 = tk.Button(self.__root, text='转换', font=('幼圆', 9, 'bold'), width=20, command=lambda: self._Run())
        self.__Bt_1.pack(expand='yes', pady='5px')
        self.__Lb_8 = tk.Label(self.__root, text='Tip:生成的图形转换文件保存地址为所选PNG或JPG文件所在地址', fg='red', font=('幼圆', 9, 'bold'))
        self.__Lb_8.pack(expand='yes', pady='5px', side='bottom')
        self.__root.mainloop()

    def _ClearPath(self):
        self.__filePath = ""
        self.__foderPath = ""
        self.__Lb_2.config(text="")
        self._RandomTip()

    def _Run(self):
        if self.__filePath == '' and self.__foderPath == '':
            tm.showinfo('提示', '请选择文件或文件夹')
            self._RandomTip()
        if self.__c.get() == 2:
            if self._CustomValueCheck() is True:
                self._Transform(self.__filePath, self.__foderPath)
        else:
            if self._ValueCheck() is True:
                self._Transform(self.__filePath, self.__foderPath)

    def _CallBack(self, p_value):
        p_value = None
        self._RandomTip()
        return True

    def _RandomTip(self):
        i = random.randint(0, len(self.__tips) - 1)
        self.__Lb_8.config(text='Tip:' + self.__tips[i])
        self.__root.update()

    def _ChooseMode(self, p_index):
        self._RandomTip()
        if p_index == 0:
            self.__filePath = tf.askopenfilename()
            self.__foderPath = ''
            self.__Lb_2.config(text=self.__filePath)
        if p_index == 1:
            self.__foderPath = tf.askdirectory()
            self.__filePath = ''
            self.__Lb_2.config(text=self.__foderPath)

    def _GetCheckButtonValue(self):
        v_vari = [self.__check1, self.__check2, self.__check3, self.__check4, self.__check5, self.__check6]
        v_values = []
        for i in range(6):
            v_values.append((i, v_vari[i].get()))
        return v_values

    def _ChooseSize(self, img, i):
        if i == 0:
            return self._Resize(img, self._GetCheckButtonValue())
        elif i == 1:
            return self._RotateSize(img, self.__r.get())
        elif i == 2:
            return self._CustomSize(img)

    def _ImageTypeChange(self, i):
        self._RandomTip()
        if i == 0:
            self.__imgType = '.ico'
        else:
            self.__imgType = '.wmf'

    def _FrameActive(self, i):
        self.__sizeMode = i
        self._RandomTip()
        if i == 0:
            self.__Fr_2.pack()
            self.__Fr_3.pack_forget()
            self.__Fr_4.pack_forget()
            self.__root.update()
        elif i == 1:
            self.__Fr_3.pack()
            self.__Fr_2.pack_forget()
            self.__Fr_4.pack_forget()
            self.__root.update()
        elif i == 2:
            self.__Fr_4.pack()
            self.__Fr_2.pack_forget()
            self.__Fr_3.pack_forget()
            self.__root.update()

    # 图形转换
    def _Transform(self, p_filePath, p_foderPath):
        if p_filePath != '':
            a = im.open(p_filePath)
            v_imgs = self._ChooseSize(a, self.__sizeMode)
            v_isNull = False
            if self.__sizeMode == 0 and (v_imgs is None or len(v_imgs) == 0):
                tm.showwarning('Warning', '请选择图形尺寸！')
                self._RandomTip()
                v_isNull = True
            v_name = p_filePath.split('/')[-1].split('.')
            imageName = v_name[0]
            suffixName = v_name[-1]
            if (suffixName == 'png' or suffixName == 'jpg') and not v_isNull:
                v_path = p_filePath.replace(p_filePath.split('/')[-1], "")
                i = 1
                v_imageInfos = []
                if self.__sizeMode == 0:
                    for v_img in v_imgs:
                        v_filePath = v_path + imageName + str(i) + '.bmp'
                        v_imageInfos.append((imageName + str(i), v_filePath))
                        v_img.save(v_filePath)
                        i += 1
                else:
                    v_filePath = v_path + imageName + str(i) + '.bmp'
                    v_imageInfos.append((imageName + str(i), v_filePath))
                    v_imgs.save(v_filePath)
                try:
                    for v_imageName, v_filePath in v_imageInfos:
                        n_filePath = v_path + v_imageName + self.__imgType
                        if not os.path.exists(n_filePath):
                            os.rename(v_filePath, n_filePath)
                        else:
                            os.remove(n_filePath)
                            os.rename(v_filePath, n_filePath)
                except FileNotFoundError:
                    tm.showerror('错误', '文件不存在')
                    self._RandomTip()
                else:
                    self.__Lb_8.config(text='转换成功')
                    self.__Lb_8.pack(pady='5px', side='bottom')
        elif p_foderPath != '':
            files = os.listdir(p_foderPath)
            isSuccess = False
            v_isNull = True
            v_values = self._GetCheckButtonValue()
            for i, v in v_values:
                if v == 1:
                    v_isNull = False
            if not v_isNull:
                for img in files:
                    path = p_foderPath + '/' + img
                    v_name = path.split('/')[-1].split('.')
                    imageName = v_name[0]
                    suffixName = v_name[-1]
                    if suffixName == 'png' or suffixName == 'jpg':
                        a = im.open(path)
                        v_imgs = self._ChooseSize(a, self.__sizeMode)
                        v_path = path.replace(path.split('/')[-1], "")
                        i = 1
                        v_imageInfos = []
                        for v_img in v_imgs:
                            v_filePath = v_path + imageName + str(i) + '.bmp'
                            v_imageInfos.append((imageName + str(i), v_filePath))
                            v_img.save(v_filePath)
                            i += 1
                        try:
                            for v_imageName, v_filePath in v_imageInfos:
                                n_filePath = v_path + v_imageName + self.__imgType
                                if not os.path.exists(n_filePath):
                                    os.rename(v_filePath, n_filePath)
                                else:
                                    os.remove(n_filePath)
                                    os.rename(v_filePath, n_filePath)
                        except FileNotFoundError:
                            tm.showerror('错误', '文件不存在')
                            self._RandomTip()
                            isSuccess = False
                        else:
                            isSuccess = True
            if isSuccess is True:
                self.__Lb_8.config(text='转换成功')
                self.__Lb_8.pack(pady='5px', side='bottom')

    def _Resize(self, img, p_values):
        width, height = img.size
        v_imgs = []
        if width == height:
            for i, value in p_values:
                if value == 1:
                    if i == 0:
                        v_imgs.append(img.resize((16, 16), im.Resampling.LANCZOS))
                    elif i == 1:
                        v_imgs.append(img.resize((32, 32), im.Resampling.LANCZOS))
                    elif i == 2:
                        v_imgs.append(img.resize((48, 48), im.Resampling.LANCZOS))
                    elif i == 3:
                        v_imgs.append(img.resize((64, 64), im.Resampling.LANCZOS))
                    elif i == 4:
                        v_imgs.append(img.resize((128, 128), im.Resampling.LANCZOS))
                    elif i == 5:
                        v_imgs.append(img.resize((256, 256), im.Resampling.LANCZOS))
        elif width > height:
            for i, value in p_values:
                if value == 1:
                    if i == 0:
                        v_height = int((16 / width) * height)
                        v_imgs.append(img.resize((16, v_height), im.Resampling.LANCZOS))
                    elif i == 1:
                        v_height = int((32 / width) * height)
                        v_imgs.append(img.resize((32, v_height), im.Resampling.LANCZOS))
                    elif i == 2:
                        v_height = int((48 / width) * height)
                        v_imgs.append(img.resize((48, v_height), im.Resampling.LANCZOS))
                    elif i == 3:
                        v_height = int((64 / width) * height)
                        v_imgs.append(img.resize((64, v_height), im.Resampling.LANCZOS))
                    elif i == 4:
                        v_height = int((128 / width) * height)
                        v_imgs.append(img.resize((128, v_height), im.Resampling.LANCZOS))
                    elif i == 5:
                        v_height = int((256 / width) * height)
                        v_imgs.append(img.resize((256, v_height), im.Resampling.LANCZOS))
        else:
            for i, value in p_values:
                if value == 1:
                    if i == 0:
                        v_width = int((16 / height) * width)
                        v_imgs.append(img.resize((v_width, 16), im.Resampling.LANCZOS))
                    elif i == 1:
                        v_width = int((32 / height) * width)
                        v_imgs.append(img.resize((v_width, 32), im.Resampling.LANCZOS))
                    elif i == 2:
                        v_width = int((48 / height) * width)
                        v_imgs.append(img.resize((v_width, 48), im.Resampling.LANCZOS))
                    elif i == 3:
                        v_width = int((64 / height) * width)
                        v_imgs.append(img.resize((v_width, 64), im.Resampling.LANCZOS))
                    elif i == 4:
                        v_width = int((128 / height) * width)
                        v_imgs.append(img.resize((v_width, 128), im.Resampling.LANCZOS))
                    elif i == 5:
                        v_width = int((256 / height) * width)
                        v_imgs.append(img.resize((v_width, 256), im.Resampling.LANCZOS))
        return v_imgs

    def _RotateSize(self, img, p_size):
        width, height = img.size
        if p_size == 0:
            v_size = (int(width * 0.5), int(height * 0.5))
            return img.resize(v_size, im.Resampling.LANCZOS)
        elif p_size == 1:
            v_size = (int(width), int(height))
            return img.resize(v_size, im.Resampling.LANCZOS)
        elif p_size == 2:
            v_size = (int(width * 1.5), int(height * 1.5))
            return img.resize(v_size, im.Resampling.LANCZOS)
        elif p_size == 3:
            return img.resize((width * 2, height * 2), im.Resampling.LANCZOS)
        elif p_size == 4:
            return img.resize((width * 3, height * 3), im.Resampling.LANCZOS)
        elif p_size == 5:
            return img.resize((width * 4, height * 4), im.Resampling.LANCZOS)

    def _CustomSize(self, img):
        return img.resize((self.__cWidth, self.__cHeight), im.Resampling.LANCZOS)

    # 选择文件检测
    def _ValueCheck(self):
        v_filePath = self.__filePath
        if v_filePath != '':
            v_fileSuffix = v_filePath.split('/')[-1].split('.')[-1]
            if v_fileSuffix == 'png' or v_fileSuffix == 'jpg':
                return True
            else:
                tm.showwarning('Warning', '请选择PNG或JPG文件！')
                self.__filePath = ''
                self.__Lb_2.config(text='')
                self._RandomTip()
                return False
        else:
            return True

    # 自定义宽高输入检测
    def _CustomValueCheck(self):
        self.__Lb_7.pack_forget()
        if self.__filePath != "" or self.__foderPath != "":
            v_i = 3
            try:
                pwidth = str(self.__Et_1.get()).strip()
                pheight = str(self.__Et_2.get()).strip()
                if pwidth == '' or pheight == '':
                    width = 16
                    height = 16
                    self.__Lb_7.config(text='你输入的数据为空')
                    self.__Lb_7.pack(pady='5px', side='bottom')
                    self._RandomTip()
                    self.__root.update()
                else:
                    v_i = 0
                    width = int(pwidth)
                    widthInfo = (width, True) if width > 0 else (1, False)
                    print(widthInfo[0])
                    if not widthInfo[1]:
                        tm.showwarning("Warning", "Width不能为负数!")
                        return
                    width = widthInfo[0]
                    widthInfo = (width, True) if width <= 3840 else (3840, False)
                    width = widthInfo[0]
                    if not widthInfo[1]:
                        tm.showwarning("Warning", "Width不能大于3840!")
                        return
                    v_i = 1
                    height = int(pheight)
                    heightInfo = (height, True) if height > 0 else (1, False)
                    if not heightInfo[1]:
                        tm.showwarning("Warning", "Height不能为负数!")
                        return
                    height = heightInfo[0]
                    heightInfo = (height, True) if height <= 2048 else (2048, False)
                    if not heightInfo[1]:
                        tm.showwarning("Warning", "Height不能大于2048!")
                        return
                    height = heightInfo[0]
                    v_i = 2
            except ValueError:
                if v_i == 0:
                    self.__Lb_7.config(text='你输入的Width的数据存在类型错误，请输入整数')
                elif v_i == 1:
                    self.__Lb_7.config(text='你输入的Height的数据存在类型错误，请输入整数')
                self.__Lb_7.pack(pady='5px', side='bottom')
                self._RandomTip()
                self.__root.update()
                return False
            else:
                if v_i == 2:
                    self.__Lb_7.pack_forget()
                    self.__Lb_8.config(text='转换成功')
                    self.__Lb_8.pack(pady='5px', side='bottom')
                    self.__cWidth = width
                    self.__cHeight = height
                    return True


class IToIco:
    def __init__(self):
        self.__ti = ToIco()


cTITI = IToIco()
