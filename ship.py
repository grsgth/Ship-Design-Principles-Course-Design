# -*- coding: utf-8 -*-
"""
Created on Fri Jan  5 18:04:49 2018

@author: Cheshire
"""

import xlwings as xw
class ship(object):
    def __init__(self,
                 Lo = 154.8,
                 #母型船船长Lo
                 Bo = 21.6, 
                 #母型船船宽Bo
                 do = 8.8,    
                 #母型船吃水do
                 Do = 12.2,   
                 #母型船型深Do
                 Deltao = 23433,
                 #母型船排水量Deltao
                 WHo = 3636.506,
                 #母型船钢料重量WHo
                 WOo = 581.2793,
                 #母型船舾装重量WOo
                 WMo = 382.2151,
                 #母型船机电重量WMo
                 Lppo = 146,
                 Vo = 12,
                 #母型船航速Vo
                 Po = 3552,
                 #母型船Lpp
                 DW = 19120,
                 DWo = 18833,
                 #设计船载重量DW
                 V = 11.1,
                 #设计船载货量
                 WC = 19120*0.96,
                 dL = 1, dB = 1, dD = 1, dCb = 1,):
                 #设计船航速V       
        self.Lo = Lo
        self.Bo = Bo
        self.do = do
        self.Do = Do
        self.Deltao = Deltao
        self.WHo = WHo
        self.WOo = WOo
        self.WMo = WMo
        self.Lppo = Lppo
        self.DW = DW
        self.DWo = DWo
        self.nDW = 0.7666 + 0.1304 * (DW / 1e5) - 0.0775 * (DW / 1e5)**2 + 0.1294 * (DW  / 1e5)**3 - 0.1441 * (DW / 1e5)**4 + 0.0469 * (DW / 1e5)**5
        #载重量系数nDW
        self.Delta = DW / self.nDW
        xw.Book(name).sheets('主尺度初').range('E3').value=self.nDW
        xw.Book(name).sheets('主尺度初').range('E4').value=self.Delta
        #设计船排水量Delta
        self.V = V
        self.Vo = Vo
        self.Po = Po
        self.WC = WC
        self.dL = dL
        self.dB = dB
        self.dD = dD
        self.dCb = dCb
    
    
    ###第一次预估主尺度
    def zcd(self,Print = False):
        L = self.Lo * (self.Delta / self.Deltao)**(1/3.0) * self.dL
        #设计船船长
        Lpp = self.Lppo * (self.Delta / self.Deltao)**(1/3.0) * self.dL
        #设计船Lpp
        B = self.Bo * (self.Delta / self.Deltao)**(1/3.0) * self.dB
        #设计船Lpp
        d = self.do * (self.Delta / self.Deltao)**(1/3.0)   
        #设计船吃水d
        D = self.Do * self.Lo * self.Bo * self.DW / (L * B * self.DWo) * self.dD
        #设计船型深D
        Cb = 1.08 - 1.68 * (11 * 0.5144 / (9.8 * L )**0.5) * self.dCb
        #设计船方形系数Cb
        Co = self.Deltao**(2/3.0)*self.Vo**3/self.Po
        #母型船海军系数
        P = self.Delta**(2/3.0) * self.V**3 / Co
        xw.Book(name).sheets('主尺度初').range('E8').value=L
        xw.Book(name).sheets('主尺度初').range('E9').value=Lpp
        xw.Book(name).sheets('主尺度初').range('E11').value=B
        xw.Book(name).sheets('主尺度初').range('E10').value=d
        xw.Book(name).sheets('主尺度初').range('E13').value=D
        xw.Book(name).sheets('主尺度初').range('E12').value=Cb
        xw.Book(name).sheets('主尺度初').range('E14').value=P
        #设计船功率预估
        if Print ==True:
            print('L:%.2f B:%.2f d:%.2f D:%.2f Lpp:%.2f Cb:%.2f P:%.2f'%( L,B,d,D,Lpp,Cb,P))
        return L,B,d,D,Lpp,Cb,P
    
    
    ###迭代选取合适主尺度
    def W(self,choice,Print = True):
        L,B,d,D,Lpp,Cb,P = self.zcd()
        CH1 = self.WHo / (self.Lo * (self.Bo + self.Do))
        WH = CH1 * L *(B + D)
        #钢料重量估算
        WO = self.WOo / (self.Lo * self.Bo) * L * B
        #舾装重量估算
        CMo = self.WMo / (self.Po / 0.7355)**0.5
        #母型船中横剖面系数CMo
        WM = CMo * (P/0.7355)**0.5
        #设计船机电重量WM
        Delta = 1.025 * 1.006 * L *B *d *Cb
        dDW = self.DW -(Delta - (WH + WO + WM))
        dW = Delta - (WH + WO + WM)
        LW = WH+WO+WM
        xw.Book(name).sheets('重力与浮力平衡').range('B2').value=WH
        xw.Book(name).sheets('重力与浮力平衡').range('B3').value=WO
        xw.Book(name).sheets('重力与浮力平衡').range('D6').value=WM
        xw.Book(name).sheets('重力与浮力平衡').range('B10:N10').value=[Delta,L,Lpp,d,B,Cb,D,WH,WO,WM,LW,dW,dDW]
        if Print ==True:
            print('dDW:%.2f Delta:%.2f L:%.2f Lpp:%.2f B:%.2f D:%.2f WH:%.2f WO:%.2f WM:%.2f'%(dDW,Delta,L,Lpp,B,D,WH,WO,WM))
        n = 10
        if choice == "L":
            while (dDW<-50) | (dDW>100) :
                n = n+1
                Delta = Delta+dDW
                #改变船长
                L = self.Lo * (Delta / self.Deltao)**(1/3.0)
                Lpp = self.Lppo * (Delta / self.Deltao)**(1/3.0)
                WH = CH1 * L *(B + D)
                WO = self.WOo / (self.Lo * self.Bo) * L * B
                dDW = self.DW -(Delta - (WH + WO + WM))
                dW = Delta - (WH + WO + WM)
                LW = WH+WO+WM
                xw.Book(name).sheets('重力与浮力平衡').range('B'+str(n)+':N'+str(n)).value=[Delta,L,Lpp,d,B,Cb,D,WH,WO,WM,LW,dW,dDW]
                if Print ==True:
                    print('dDW:%.2f Delta:%.2f L:%.2f Lpp:%.2f WH:%.2f WO:%.2f WM:%.2f'%(dDW,Delta,L,Lpp,WH,WO,WM))
        elif choice == "B":
            while (dDW<-50) | (dDW>100):
                n = n+1
                Delta = Delta+dDW
                #改变船宽
                B = self.Bo * (Delta / self.Deltao)**(1/3.0)
                WH = CH1 * L *(B + D)
                WO = self.WOo / (self.Lo * self.Bo) * L * B
                dDW = self.DW -(Delta - (WH + WO + WM))
                dW = Delta - (WH + WO + WM)
                LW = WH+WO+WM
                xw.Book(name).sheets('重力与浮力平衡').range('B'+str(n)+':N'+str(n)).value=[Delta,L,Lpp,d,B,Cb,D,WH,WO,WM,LW,dW,dDW]
                if Print ==True:
                    print('dDW:%.2f Delta:%.2f B:%.2f WH:%.2f WO:%.2f WM:%.2f'%(dDW,Delta,B,WH,WO,WM))
        elif choice == "D":
            while (dDW<-100) | (dDW>100):
                n = n+1
                Delta = Delta+dDW
                #改变船吃水
                D = self.Do * self.Lo * self.Bo * Delta / (L * B * self.Deltao)
                WH = CH1 * L *(B + D)
                WO = self.WOo / (self.Lo * self.Bo) * L * B
                dDW = self.DW -(Delta - (WH + WO + WM))
                dW = Delta - (WH + WO + WM)
                LW = WH+WO+WM
                xw.Book(name).sheets('重力与浮力平衡').range('B'+str(n)+':N'+str(n)).value=[Delta,L,Lpp,d,B,Cb,D,WH,WO,WM,LW,dW,dDW]
                if Print ==True:
                   print('dDW:%.2f Delta:%.2f D:%.2f WH:%.2f WO:%.2f WM:%.2f'%(dDW,Delta,D,WH,WO,WM))
        else:
            print("error:请输入L,B,D")
        return WH,WO,WM,L,B,D,Lpp,Delta
    
    def Vjh( self , choi="L"):
        _,_,d,_,_,Cb,P = self.zcd()
        _,_,_,L,B,D,Lpp,_ = self.W(choice = choi,Print = False)
        uc = xw.Book(name).sheets('舱容校核').range('B3').value
        kc = xw.Book(name).sheets('舱容校核').range('B4').value
        VC = self.WC * uc/kc
        xw.Book(name).sheets('舱容校核').range('B5').value = VC
        ###VC
        kB = xw.Book(name).sheets('舱容校核').range('B10').value
        VB = kB * self.DW
        xw.Book(name).sheets('舱容校核').range('B11').value = VB
        ###VB
        KM = xw.Book(name).sheets('舱容校核').range('B14').value
        LM = xw.Book(name).sheets('舱容校核').range('B15').value
        hDM = xw.Book(name).sheets('舱容校核').range('B16').value
        VM = KM * LM * B * (D - hDM)
        xw.Book(name).sheets('舱容校核').range('B17').value = VM
        ###VM
        C1 = xw.Book(name).sheets('舱容校核').range('B20').value
        SM = xw.Book(name).sheets('舱容校核').range('B21').value
        xw.Book(name).sheets('舱容校核').range('B22').value = B / 50
        CBD = Cb + (1-Cb)*(D - d)/(C1 * d)
        VH = CBD * Lpp * B *(D + 0.7 * B/50 +SM )  
        xw.Book(name).sheets('舱容校核').range('B23').value = VH
        ###VH
        VTC = 0.57 * Lpp * B * D
        xw.Book(name).sheets('舱容校核').range('B6').value = VTC
        ###VTC
        dVH = VH-VC-VB-VM
        xw.Book(name).sheets('舱容校核').range('B7').value = VTC - VC
        xw.Book(name).sheets('舱容校核').range('B7').value = VTC - VC
        xw.Book(name).sheets('舱容校核').range('B24').value = VC+VB+VM
        xw.Book(name).sheets('舱容校核').range('B25').value = dVH
        xw.Book(name).sheets('舱容校核').range('B26').value = str(dVH / VH*100) + '%'
        
        print('VC:%.2f VTC:%.2f'%(VC,VTC))
        print('VC:%.2f VB:%.2f VM:%.2f VH:%.2f dVH:%.2f %.2f %%'%(VC,VB,VM,VH,dVH,dVH / VH*100))
        
    def cbxn(self,choi = "L",zgHo = 7.3):
        CE = zgHo / self.Do
        WH,WO,WM,L,B,D,Lpp,Delta = self.W(choice = choi,Print = False)
        zgH = CE * D
        zgO = 1.02 * D
        zgM = 0.55 * D
        zgE = (WH * zgH + WO * zgO + WM * zgM)/(WH + WO + WM)
        
        _,_,d,_,Lpp,Cb,P = self.zcd()
        Co = self.Deltao**(2/3.0)*self.Vo**3/self.Po
        
        Vsd = (Co * P /Delta**(2/3.0))**(1/3.0)
        #M = a1 * d + a2 * B**2 / d - a3 * D
        #f = 0.75 * B / (GM)**0.5
        return zgE,Vsd

print("输入文件名：")
name = 'w1.xlsm'#input()   
sht=xw.Book(name).sheets('主尺度初').range('B2:B22').value
Lo = sht[4]
Bo = sht[6]
do = sht[7]
Do = sht[8]
Deltao = sht[3]
WHo = sht[12]
WOo = sht[10]
WMo = sht[11]
Lppo = sht[5]
Vo = sht[15]
Po = sht[14]
DW = sht[19]
DWo = sht[1]
V = sht[20]
WC = DW*0.95
ship1 = ship(Lo,Bo,do,Do,Deltao,WHo,WOo,WMo,Lppo,Vo ,Po ,DW,DWo,V ,WC ,dD = 1.01)
ship1.zcd(Print =True)
gb = xw.Book(name).sheets('重力与浮力平衡').range('B8').value
ship1.W(gb)
ship1.Vjh(gb)
ship1.cbxn(gb)