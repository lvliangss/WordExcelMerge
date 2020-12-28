# -*- coding:utf-8 -*-
import os;
import xlrd;
import xlwt;
import wx;
import sys;
import traceback;
from datetime import datetime, date;
from xlrd import xldate_as_tuple;
import win32com.client as win32;

def ShowErrorMsg(msg):
        dlg=wx.MessageDialog(None,msg,"Error/Warning/Info Message",wx.OK|wx.ICON_QUESTION)
        if dlg.ShowModal()==wx.ID_OK:
                pass
        dlg.Destroy()

def add_MBV(event):
        try:
                dlg=wx.FileDialog(win, u"请选择EXCEL文件", wildcard="xls files (*.*)|*.*|(*.xls)|*.xls|(*.xlsx)|*.xlsx",style=wx.FD_MULTIPLE);
                if dlg.ShowModal() == wx.ID_OK:
                        dlg_MBV_Add = wx.ProgressDialog("", "稍安勿躁，正在合并文件......", 100, style=wx.PD_AUTO_HIDE)
                        MBV_new=dlg.GetPaths();                      
                        progress = 10.0
                        dlg_MBV_Add.Update(progress);
                        xls=XlsMerge(MBV_new);
                        xls.clear();
                        xls.merge();                                 
                        dlg_MBV_Add.Update(100);
                        dirname, filename = os.path.split(os.path.abspath(sys.argv[0]))
                        ##dirname.decode('gbk')将str转化成unicode否则无法显示中文
                        ShowErrorMsg(u'合并完毕！请查看'+dirname.decode('gbk')+u'\\汇总表.xls')
        except:
                dlg_MBV_Add.Update(100);
##                print traceback.format_exc()
                ShowErrorMsg(traceback.format_exc() + "please close and restart software");
def MergePureXls(event):
        try:
                dlg=wx.FileDialog(win, u"请选择EXCEL文件", wildcard="xls files (*.*)|*.*|(*.xls)|*.xls|(*.xlsx)|*.xlsx",style=wx.FD_MULTIPLE);
                if dlg.ShowModal() == wx.ID_OK:
                        dlg_MBV_Add = wx.ProgressDialog("", "稍安勿躁，正在合并文件......", 100, style=wx.PD_AUTO_HIDE)
                        MBV_new=dlg.GetPaths();                      
                        progress = 10.0
                        dlg_MBV_Add.Update(progress);
                        xls=XlsMerge(MBV_new);
                        xls.clear();
                        xls.mergePure();                                 
                        dlg_MBV_Add.Update(100);
                        dirname, filename = os.path.split(os.path.abspath(sys.argv[0]))
                        ##dirname.decode('gbk')将str转化成unicode否则无法显示中文
                        ShowErrorMsg(u'合并完毕！请查看'+dirname.decode('gbk')+u'\\汇总表.xls')
        except:
                dlg_MBV_Add.Update(100);
##                print traceback.format_exc()
                ShowErrorMsg(traceback.format_exc() + "please close and restart software");
                
def mergeWord(event):
        # 合并后格式会乱，故暂时不用
        try:
                dlg=wx.FileDialog(win, u"请选择word文件", wildcard="doc files (*.*)|*.*|(*.doc)|*.doc|(*.docx)|*.docx",style=wx.FD_MULTIPLE);
                if dlg.ShowModal() == wx.ID_OK:
                        dlg_Word_Add = wx.ProgressDialog("", "稍安勿躁，正在合并文件......", 100, style=wx.PD_AUTO_HIDE)
                        Word_new=dlg.GetPaths();                        
                        word = win32.gencache.EnsureDispatch('Word.Application')
                        word.Visible = False;                        
                        files = Word_new;                        
                        #新建合并后的文档
                        output = word.Documents.Add()
                        progress = 10.0;
                        for file in files:
                            output.Application.Selection.InsertFile(file)#拼接文档
                            if progress<95:
                                    dlg_Word_Add.Update(progress);
                        #获取合并后文档的内容
                        doc = output.Range(output.Content.Start, output.Content.End);
                        dirname, filename = os.path.split(os.path.abspath(sys.argv[0]))
                        output.SaveAs(dirname.decode('gbk')+u'\\合并后.docx') #保存
                        output.Close();
                        
                                                         
                        dlg_Word_Add.Update(100);
                        
                        ##dirname.decode('gbk')将str转化成unicode否则无法显示中文
                        ShowErrorMsg(u'合并完毕！请查看'+dirname.decode('gbk')+u'\\合并后.docx')
        except:
                dlg_Word_Add.Update(100);
##                print traceback.format_exc()
                ShowErrorMsg(traceback.format_exc() + "please close and restart software");

def mergeWord1(event):
        # 格式不会乱
        try:
                dlg=wx.FileDialog(win, u"请选择word文件", wildcard="doc files (*.*)|*.*|(*.doc)|*.doc|(*.docx)|*.docx",style=wx.FD_MULTIPLE);
                if dlg.ShowModal() == wx.ID_OK:
                        dlg_Word_Add = wx.ProgressDialog("", "稍安勿躁，正在合并文件......", 100, style=wx.PD_AUTO_HIDE)
                        Word_new=dlg.GetPaths();                        
                        word = win32.gencache.EnsureDispatch('Word.Application')
                        word.Visible = False;                        
                        files = Word_new;                        
                        #新建合并后的文档
                        output = word.Documents.Add()
                        progress = 10.0;
                        for file in files:
                                temp_document = word.Documents.Open(file)
                                word.Selection.WholeStory();#全选
                                word.Selection.Copy() 
                                temp_document.Close()
                                # output.Range(output.Content.Start, output.Content.End)
                                output.Range()                                
                                # word.Selection.Delete()                               
                                word.Selection.Paste() 
                                # output.Range(output.Content.Start, output.Content.End).InsertAfter('\n')                           
                                # output.Range().InsertAfter('\n')
                                if progress<95:
                                        dlg_Word_Add.Update(progress);
                        #获取合并后文档的内容
                        # doc = output.Range(output.Content.Start, output.Content.End);
                        dirname, filename = os.path.split(os.path.abspath(sys.argv[0]))
                        output.SaveAs(dirname.decode('gbk')+u'\\合并后.docx') #保存
                        output.Close();
                        word.Quit();
                                                         
                        dlg_Word_Add.Update(100);
                        
                        ##dirname.decode('gbk')将str转化成unicode否则无法显示中文
                        ShowErrorMsg(u'合并完毕！请查看'+dirname.decode('gbk')+u'\\合并后.docx')
        except:
                dlg_Word_Add.Update(100);
##                print traceback.format_exc()
                ShowErrorMsg(traceback.format_exc() + "please close and restart software");

class XlsMerge:
        def __init__(self, mergeList=[]):
                self.mergeList=mergeList;
        def clear(self):
                dirName= os.path.dirname(self.mergeList[0]);
                desti= dirName+u'\\汇总表.xls'
                try:
                        os.remove(desti);
                except:
                        pass;
        def toDate(self,value):
                ##*将tuple转化成integer，value为从excel里读出来的值
                dateValue=xldate_as_tuple(value, 0);
                dateValueYMD=datetime(*dateValue[0:3]);
                formatDate=dateValueYMD.strftime('%Y-%d-%m');
                return formatDate;
        def merge(self):
                firstExcel=True;
                dataExcelNew=[];
                for xls in self.mergeList:
                        if xls.endswith('.xls') or xls.endswith('.xlsx'):
                                excel=xlrd.open_workbook(xls);
                                sheet_M = excel.sheet_by_index(0);##默认只读第1个表
                                numRow = sheet_M.nrows;
                                numColum = sheet_M.ncols;
                                for i in range(numRow):
                                        if i==0 and firstExcel==False:
                                                continue;
                                        else:
                                                data=sheet_M.row_values(i);
                                                for j in range(numColum):
                                                        if sheet_M.cell(i,j).ctype==3:                                                                
                                                                data[j]=self.toDate(sheet_M.cell_value(i,j));                                                
                                                dataExcelNew.append(data);
                                firstExcel=False;
                newWorkbook=xlwt.Workbook(encoding = 'utf-8');
                worksheet = newWorkbook.add_sheet(u'总表');
                i=0;
                for eachData in dataExcelNew:
                        j=0;
                        for eachElemnet in eachData:
                                worksheet.write(i, j, eachElemnet);
                                j=j+1;
                        i=i+1;
                newWorkbook.save(u'汇总表.xls');

        def mergePure(self):
                firstExcel=True;
                dataExcelNew=[];
                for xls in self.mergeList:
                        if xls.endswith('.xls') or xls.endswith('.xlsx'):
                                excel=xlrd.open_workbook(xls);
                                sheet_M = excel.sheet_by_index(0);##默认只读第1个表
                                numRow = sheet_M.nrows;
                                numColum = sheet_M.ncols;
                                for i in range(numRow):       
                                        data=sheet_M.row_values(i);
                                        for j in range(numColum):
                                                if sheet_M.cell(i,j).ctype==3:                                                                
                                                        data[j]=self.toDate(sheet_M.cell_value(i,j));                                                
                                        dataExcelNew.append(data);
                                firstExcel=False;
                newWorkbook=xlwt.Workbook(encoding = 'utf-8');
                worksheet = newWorkbook.add_sheet(u'总表');
                i=0;
                for eachData in dataExcelNew:
                        j=0;
                        for eachElemnet in eachData:
                                worksheet.write(i, j, eachElemnet);
                                j=j+1;
                        i=i+1;
                newWorkbook.save(u'汇总表.xls');


license=False;      
app = wx.App()

frame= wx.Frame(None,title=u'EXCEL合并',pos=(400,200), size=(350,300))#实例化窗口
win=wx.Panel(frame)

StartButton= wx.Button(win,label=u'请选择要合并的表格(去重合并)',pos=(95,25),size=(180,40))
MergeButton= wx.Button(win,label=u'请选择要合并的表格(单纯合并)',pos=(95,80),size=(180,40))
wordButton= wx.Button(win,label=u'请选择要合并的word',pos=(95,135),size=(180,40))
StartButton.Bind(wx.EVT_BUTTON,add_MBV)
MergeButton.Bind(wx.EVT_BUTTON,MergePureXls)
wordButton.Bind(wx.EVT_BUTTON,mergeWord1)
        

frame.Show(True)
app.MainLoop()          


        
