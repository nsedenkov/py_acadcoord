#!/usr/bin/python
#! coding: utf-8

#===============================================================================
#
#   Экспортирует в Excel координаты объектов Polyline и 2DPolyline из указанного слоя
#   активного чертежа AutoCAD
#   В модели в слое Numbers проставляет номера точек (если слоя нет, он будет создан),
#   а также номера контуров
#
#===============================================================================

from __future__ import division
import sys
from os import path, environ
if path.exists(path.normpath('x:\СИД\~INFO\Python\COMLibs').decode('utf-8').encode('cp1251')):
    sys.path.append(path.normpath('x:\СИД\~INFO\Python\COMLibs').decode('utf-8').encode('cp1251'))
if path.exists(path.normpath('d:\PY\LIB')):
    sys.path.append(path.normpath('d:\PY\LIB'))
from Tkinter import Tk, Button, Frame, LabelFrame, Listbox, Scrollbar, Menu, Menubutton, Radiobutton, Label, Entry, StringVar, IntVar
from Tkinter import DISABLED, NORMAL, END, LEFT
from comtypes.client import *
from comtypes.automation import *
from array import array
import locale
from math import sqrt

class main:

    def __init__(self, master):
        self.master = master
        self.DicObjType = {1:'', 2:'3D полилиния', 3:'', 4:'Дуга', 5:'', 6:'', 7:'Вхождение блока', 8:'Окружность', 9:'Aligned размер', 10:'Угловой размер', 11:'', 12:'', 13:'', 14:'', 15:'', 16:'Эллипс', 17:'Штриховка', 18:'Выноска', 19:'Линия', 20:'', 21:'', 22:'Точка', 23:'2D полилиния', 24:'Полилиния', 25:'', 26:'Маскирующая область', 27:'Луч', 28:'Область', 29:'', 30:'', 31:'Сплайн', 32:'Текст', 33:'', 34:'', 35:'', 36:'Конструктивная линия', 37:'', 38:'', 39:'', 40:''}
        self.title = 'Экспорт координат объектов AutoCAD'
        self.master.title(self.title)
        self.master.geometry('360x290')
        self.master.resizable(False, False)
        self.MenuVar = StringVar()
        self.RBVar = IntVar()
        self.RBVar.set(1)
        self.PLineCrd = []
        self.nprefix = ''
        self.startnumpntfrom = 1
        self.startnumprclfrom = 1
        self.master.LFrame = Frame(self.master, width = 280)
        self.master.LFrame.pack(side = 'left', fill = 'y')
        self.master.RFrame = Frame(self.master)
        self.master.RFrame.pack(side = 'right', fill = 'both')
        self.master.frame1=LabelFrame(self.master.RFrame, height = 80, labelanchor='nw', text='Слои чертежа')
        self.master.frame1.pack(side = 'top', fill = 'x', expand = 0, padx=2)
        self.master.frame3=Frame(self.master.RFrame, height = 80)
        self.master.frame3.pack(side = 'bottom', fill = 'x', expand = 0)
        self.master.frame2=LabelFrame(self.master.RFrame, labelanchor='nw', text='Объекты слоя')
        self.master.frame2.pack(side = 'top', fill = 'both', expand = 1)
        self.master.frame4=LabelFrame(self.master.LFrame, height = 80, width = 220, labelanchor='nw', text='Порядок нумерации')
        self.master.frame4.pack(side = 'top', fill = 'both', expand = 1, padx=2)
        self.master.frame5=Frame(self.master.LFrame, height = 80)
        self.master.frame5.pack(side = 'top', fill = 'x', expand = 0)
        self.master.scrl=Scrollbar(self.master.frame2)
        self.master.entitys = Listbox(self.master.frame2, yscrollcommand=(self.master.scrl, 'set'))
        self.master.layers = self.createMenu(self.master.frame1)
        self.master.layers.pack(side=LEFT, expand = 1, fill='x')
        self.master.btn1 = Button(self.master.frame3, width = 16, height = 1, text = 'Экспорт', state=DISABLED, command=self.btn1_press)
        self.master.btn2 = Button(self.master.frame3, width = 16, height = 1, text = 'Закрыть', command=self.Quit)
        self.master.btn1.pack(anchor = 'n', side = 'left', padx = 2, pady = 2, expand = 1)
        self.master.btn2.pack(anchor = 'n', side = 'right', padx = 2, pady = 2, expand = 1)
        self.master.scrl.configure(command = self.master.entitys.yview)
        self.master.scrl.pack(side = 'right', fill = 'y', pady = 1)
        self.master.entitys = Listbox(self.master.frame2, yscrollcommand=(self.master.scrl, 'set'))
        self.master.entitys.pack(side = 'top', fill = 'both', expand = 1)
        self.master.sf1=Frame(self.master.frame4)
        self.master.sf2=Frame(self.master.frame4)
        self.master.sf3=Frame(self.master.frame4)
        self.master.sf4=Frame(self.master.frame4)
        self.master.sf5=Frame(self.master.frame4)
        self.master.sf6=LabelFrame(self.master.frame5, height = 80, width = 220, labelanchor='nw', text='Префикс нумерации')
        self.master.sf7=LabelFrame(self.master.frame5, height = 80, width = 220, labelanchor='nw', text='Нумерация точек с:')
        self.master.sf8=LabelFrame(self.master.frame5, height = 80, width = 220, labelanchor='nw', text='Нумерация участков с:')
        self.master.sf1.pack(pady = 1, fill = 'x')
        self.master.sf2.pack(pady = 1, fill = 'x')
        self.master.sf3.pack(pady = 1, fill = 'x')
        self.master.sf4.pack(pady = 1, fill = 'x')
        self.master.sf5.pack(pady = 1, fill = 'x')
        self.master.sf6.pack(side = 'bottom', fill = 'both', expand = 1, padx=2, pady = 2)
        self.master.sf7.pack(side = 'bottom', fill = 'x', expand = 0, padx=2, pady = 2)
        self.master.sf8.pack(side = 'bottom', fill = 'x', expand = 0, padx=2, pady = 2)
        self.master.rb1=Radiobutton(self.master.sf1, text = 'По умолчанию', variable=self.RBVar, value=1)
        self.master.rb2=Radiobutton(self.master.sf2, text = 'Север - юг', variable=self.RBVar, value=2)
        self.master.rb3=Radiobutton(self.master.sf3, text = 'Запад - восток', variable=self.RBVar, value=3)
        self.master.rb4=Radiobutton(self.master.sf4, text = 'Юг - север', variable=self.RBVar, value=4)
        self.master.rb5=Radiobutton(self.master.sf5, text = 'Восток - запад', variable=self.RBVar, value=5)
        self.master.rb1.pack(side = 'left', padx = 2, fill = 'x')
        self.master.rb2.pack(side = 'left', padx = 2, fill = 'x')
        self.master.rb3.pack(side = 'left', padx = 2, fill = 'x')
        self.master.rb4.pack(side = 'left', padx = 2, fill = 'x')
        self.master.rb5.pack(side = 'left', padx = 2, fill = 'x')
        self.master.etr1 = Entry(self.master.sf6, width = 18)
        self.master.etr2 = Entry(self.master.sf7, width = 18)
        self.master.etr3 = Entry(self.master.sf8, width = 18)
        self.master.etr1.pack(side = 'left', padx = 2, pady = 2, fill = 'both')
        self.master.etr2.pack(side = 'left', padx = 2, pady = 2, fill = 'both')
        self.master.etr3.pack(side = 'left', padx = 2, pady = 2, fill = 'both')
        self.master.etr2.insert(0, '1')
        self.master.etr3.insert(0, '1')
        locale.setlocale(locale.LC_NUMERIC, 'Russian_Russia')
        self.ConnectACAD()
        self.master.mainloop()
        
    def GetDcmlSep(self):
        return locale.localeconv()['decimal_point']#str(bytes(str(3 / 2))[1])
        
    def XlsCrdString(self, ws, row, n, x, y):
        ws.Cells[row, 1] = str(n)
        ws.Cells[row, 1].BorderAround(1,3,1,1)
        ws.Cells[row, 1].HorizontalAlignment = 3
        ws.Cells[row, 3] = '{0:10.2f}'.format(x)
        ws.Cells[row, 3].BorderAround(1,3,1,1)
        ws.Cells[row, 3].NumberFormat = '0'+self.GetDcmlSep()+'00'
        ws.Cells[row, 2] = '{0:10.2f}'.format(y)
        ws.Cells[row, 2].BorderAround(1,3,1,1)
        ws.Cells[row, 2].NumberFormat = '0'+self.GetDcmlSep()+'00'
        
    def XlsHdrString(self, ws, row):
        st = 'Номер точки'
        ws.Cells[row, 1] = st.decode('utf-8').encode('cp1251')
        ws.Cells[row, 1].BorderAround(1,3,1,1)
        ws.Cells[row, 1].HorizontalAlignment = 3
        st = 'X, м'
        ws.Cells[row, 2] = st.decode('utf-8').encode('cp1251')
        ws.Cells[row, 2].BorderAround(1,3,1,1)
        st = 'Y, м'
        ws.Cells[row, 3] = st.decode('utf-8').encode('cp1251')
        ws.Cells[row, 3].BorderAround(1,3,1,1)
        return row + 1
        
    def ToExcel(self):
        xls = CreateObject("Excel.Application")
        xls.WorkBooks.Add()
        S = xls.WorkBooks[1].WorkSheets[1]
        S.Columns[1].ColumnWidth = 5 + len(self.nprefix)
        S.Columns[2].ColumnWidth = 10
        S.Columns[3].ColumnWidth = 10
        xls.Visible = True
        ROW = 0
        lxy = [] # Сквозная нумерация точек
        pxy = [] # Список точек внутри участка для исключения дублирования строк в ведомости
        i = 1
        for crdlst in self.PLineCrd:
            ROW += 1
            st = 'Участок '
            S.Cells[ROW, 1] = st.decode('utf-8').encode('cp1251') + self.nprefix + str(i + self.startnumprclfrom - 1)
            self.MarkParcel(crdlst, self.nprefix + str(i + self.startnumprclfrom - 1), 'Numbers')
            ROW += 1
            ROW = self.XlsHdrString(S, ROW)
            del pxy[0:len(pxy)]
            j = 1
            for txy in crdlst:
                if lxy.count(txy) == 0:
                    lxy.append(txy)
                    self.MarkPoint(txy, self.nprefix + str(lxy.index(txy)+self.startnumpntfrom), 'Numbers')
                if pxy.count(txy) == 0:
                    self.XlsCrdString(S, ROW, self.nprefix + str(lxy.index(txy)+self.startnumpntfrom), txy[0], txy[1])
                    pxy.append(txy)
                    ROW += 1
                if j == 1:
                    txy1 = txy
                j += 1
            self.XlsCrdString(S, ROW, self.nprefix + str(lxy.index(txy1)+self.startnumpntfrom), txy1[0], txy1[1])
            i += 1

    def Quit(self):
        self.master.destroy()

    def btn1_press(self):
        self.nprefix = self.master.etr1.get()
        self.startnumpntfrom = int(self.master.etr2.get())
        self.startnumprclfrom = int(self.master.etr3.get())
        self.SortPntList()
        self.ToExcel()
        
    def createMenu(self, master = None):
        self.mb = Menubutton(master, indicatoron = 1, anchor = 'w', textvariable = self.MenuVar)
        self.mb.pack(padx = 2, pady = 1)
        self.menu = Menu(self.mb, tearoff = 0, bg = 'white')
        self.mb.configure(menu = self.menu)
        return self.mb

    def ConnectACAD(self):
        self.acad = GetActiveObject("AutoCAD.Application")
        self.dwg = self.acad.ActiveDocument
        self.mspace = self.dwg.ModelSpace
        self.master.title(self.title+' - '+self.dwg.Name.encode('utf-8'))
        for x in xrange(0,self.dwg.Layers.Count) :
            self.menu.add_command(label = self.dwg.Layers[x].Name, command = lambda x = x: self.SetActiveLayer(x))
        self.MenuVar.set(self.menu.entrycget(0, 'label'))
        self.LayerObjects(self.menu.entrycget(0, 'label'))
        try:
            lay = self.dwg.Layers('Numbers')
        except:
            lay = self.dwg.Layers.Add('Numbers')
        lay.Color = 253
        lay.IsPlot = False
        
    def SetActiveLayer(self, idx):
        self.MenuVar.set(self.menu.entrycget(idx, 'label'))
        self.LayerObjects(self.menu.entrycget(idx, 'label'))
        
    def ResetCoord(self):
        del self.PLineCrd[0:len(self.PLineCrd)]
        
    def CollectCoord(self, coords):
        tmplst = []
        crdx = coords[0:len(coords)-1:2]
        crdy = coords[1:len(coords):2]
        for j in xrange(0,len(crdx)):
            txy = (round(crdx[j],2), round(crdy[j],2))
            tmplst.append(txy)
        NP = self.GetNordPnt(tmplst)
        j = tmplst.index(NP)
        lst1 = tmplst[j:len(tmplst)]
        lst2 = tmplst[0:j ]
        del tmplst[0:len(tmplst)]
        tmplst.extend(lst1)
        tmplst.extend(lst2)
        self.PLineCrd.append(tmplst)
        
    def SortPntList(self):
        lxy = []    # Список уникальных пар X Y
        tmplst = [] # Копия основного списка контуров PLineCrd
        looplst = []    # Временный список текущих контуров для итерации
        reslst = [] # Список отсортированных контуров
        fndlst = [] # Временный список для найденных соседних контуров - живет в течение одной итерации
        lidx = []   # Список индексов контуров, найденных в tmplst в ходе итерации
        tmplst.extend(self.PLineCrd)
        for crdlst in tmplst:
            for txy in crdlst:
                if lxy.count(txy) == 0:
                    lxy.append(txy)
        NP = self.GetNordPnt(lxy)
        SP = self.GetSouthPnt(lxy)
        WP = self.GetWestPnt(lxy)
        EP = self.GetEastPnt(lxy)
        if self.RBVar.get() ==  1: # Автоматический выбор направления
            deltax = EP[0] - WP[0]
            deltay = NP[1] - SP[1]
            if deltay >= deltax: # Ориентация север - юг
                for crdlst in tmplst:
                    if NP in crdlst:
                        idx = tmplst.index(crdlst)
            else: # Ориентация запад - восток
                for crdlst in tmplst:
                    if WP in crdlst:
                        idx = tmplst.index(crdlst)
        elif self.RBVar.get() ==  2: # Принудительно от севера
            for crdlst in tmplst:
                if NP in crdlst:
                    idx = tmplst.index(crdlst)
        elif self.RBVar.get() ==  3: # Принудительно от запада
            for crdlst in tmplst:
                if WP in crdlst:
                    idx = tmplst.index(crdlst)
        elif self.RBVar.get() ==  4: # Принудительно от юга
            for crdlst in tmplst:
                if SP in crdlst:
                    idx = tmplst.index(crdlst)
        else: # Принудительно от востока
            for crdlst in tmplst:
                if EP in crdlst:
                    idx = tmplst.index(crdlst)
        # Берем первый контур по найденному индексу
        tmpcrd = tmplst.pop(idx)
        reslst.insert(0,tmpcrd)
        looplst.append(tmpcrd)
        while len(tmplst) > 1:
            del lidx[0:len(lidx)]
            for tmpcrd in looplst:
                idx = -1
                m1 = set(tmpcrd)
                for crdlst in tmplst:
                    m2 = set(crdlst)
                    if len(m1 & m2) > 0: # Если мощность пересечения больше нуля - соседний участок найден
                        idx = tmplst.index(crdlst)
                        lidx.append(idx)
                        continue
            i = 0
            for idx in lidx:
                idx -= i # Сдвиг индекса в случае удаления нескольких записей (на 2 шаге -1 и т.д.)
                i += 1
                tmpcrd = tmplst.pop(idx)
                fndlst.append(tmpcrd)
            if (idx == -1) and (len(lidx) == 0): # если разрыв в трассе
                del lxy[0:len(lxy)]
                for crdlst in tmplst:
                    for txy in crdlst:
                        if lxy.count(txy) == 0:
                            lxy.append(txy)
                # Крайние точки оставшихся участков
                NP = self.GetNordPnt(lxy)
                SP = self.GetSouthPnt(lxy)
                WP = self.GetWestPnt(lxy)
                EP = self.GetEastPnt(lxy)
                idx = 0
                minlen = 100000
                for tmpcrd in looplst:
                    for txy in tmpcrd:
                        nlen = self.Pifagor(txy, NP)
                        slen = self.Pifagor(txy, SP)
                        wlen = self.Pifagor(txy, WP)
                        elen = self.Pifagor(txy, EP)
                        if nlen < minlen:
                            minlen = nlen
                            PP = NP
                        if slen < minlen:
                            minlen = slen
                            PP = SP
                        if wlen < minlen:
                            minlen = wlen
                            PP = WP
                        if elen < minlen:
                            minlen = elen
                            PP = EP
                for crdlst in tmplst:
                    if PP in crdlst: # если ближайшая точка принадлежит контуру - ближайший участок найден
                        idx = tmplst.index(crdlst)
                tmpcrd = tmplst.pop(idx)
                fndlst.append(tmpcrd)
            del looplst[0:len(looplst)]
            looplst.extend(fndlst)
            del fndlst[0:len(fndlst)]
            reslst.extend(looplst)
        reslst.extend(tmplst)
        self.ResetCoord()
        self.PLineCrd.extend(reslst)
        
    def Pifagor(self, t1, t2): # на вход - два кортежа вида (x, y), на выходе float - расстояние между точками
        return sqrt((t1[0] - t2[0]) * (t1[0] - t2[0]) + (t1[1] - t2[1]) * (t1[1] - t2[1]))
        
    def GetNordPnt(self, crdlst):
        # crdlst - список кортежей вида (X, Y)
        rxy = crdlst[0]
        for txy in crdlst:
            if txy[1] > rxy[1]:
                rxy = txy
        return rxy

    def GetSouthPnt(self, crdlst):
        rxy = crdlst[0]
        for txy in crdlst:
            if txy[1] < rxy[1]:
                rxy = txy
        return rxy

    def GetEastPnt(self, crdlst):
        rxy = crdlst[0]
        for txy in crdlst:
            if txy[0] > rxy[0]:
                rxy = txy
        return rxy

    def GetWestPnt(self, crdlst):
        rxy = crdlst[0]
        for txy in crdlst:
            if txy[0] < rxy[0]:
                rxy = txy
        return rxy
        
    def LayerObjects(self, LName):
        self.master.entitys.delete(0, END)
        self.master.btn1.configure(state = DISABLED)
        self.ResetCoord()
        ocnt = [0]*41
        for entity in self.mspace:
            if entity.Layer == LName:
                ocnt[entity.EntityType] += 1
                if entity.EntityType in (23,24):
                    self.CollectCoord(entity.Coordinates)
        for oc in xrange(1,len(ocnt)):
            if ocnt[oc] > 0:
                self.master.entitys.insert(END, self.DicObjType[oc]+' ('+str(oc)+')'+' - '+str(ocnt[oc]))
        if self.master.entitys.size() > 0:
            self.master.btn1.configure(state = NORMAL)
            
    def MarkPoint(self, xy, num, lay):
        
        def point(*args):
            lst = [0.]*3
            if len(args) < 3:
                lst[0:2] = [float(x) for x in args[0:2]]
            else:
                lst = [float(x) for x in args[0:3]]
            return VARIANT(array("d",lst))
        
        p1 = point(xy[0], xy[1])
        ent = self.mspace.AddCircle(p1, 0.1)
        ent.Layer = lay
        p1 = point(xy[0]+0.3, xy[1]+0.3)
        ent = self.mspace.AddText(num, p1, 1)
        ent.Layer = lay
        
    def MarkParcel(self, lxy, num, lay):

        def point(*args):
            lst = [0.]*3
            if len(args) < 3:
                lst[0:2] = [float(x) for x in args[0:2]]
            else:
                lst = [float(x) for x in args[0:3]]
            return VARIANT(array("d",lst))
        
        NP = self.GetNordPnt(lxy)
        SP = self.GetSouthPnt(lxy)
        WP = self.GetWestPnt(lxy)
        EP = self.GetEastPnt(lxy)
        x = (WP[0] + EP[0]) // 2
        y = (NP[1] + SP[1]) // 2
        p1 = point(x,y)
        ent = self.mspace.AddText(num, p1, 2)
        ent.Layer = lay
    
root=Tk()
main(root)
