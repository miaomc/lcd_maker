# -*- coding: cp936 -*-
import tkinter as tk
from tkinter import ttk
import excel
from datetime import datetime


making_times = 0


def main():
    # 模板文件名称
    originFileName = 'LCD_template.xlsx'

    # 数据文件名次
    dataFileName = 'data.xlsx'

    # 工具标题
    mainframe = tk.Tk()
    mainframe.title(u'一键LCD配单生成工具V1.0')

    # 内容
    """
    XiangMuMingCheng = u'项目名称'
    Hang = u'行'
    Lie = u'列'
    DAPING = u'大屏型号'
    ChuLiQi = u'处理器'
    XianCai = u'线材长度'
    ZhiJia = u'支架'
    YunFei = u'运费'
    AnZhuangFei = u'安装费'
    """

    # 输入参数:argument
    argumentDict = {'XIANGMUMINGCHENG':u'项目名称',
                   'HANG':u'行',
                   'LIE':u'列',
                   }

    # 表格行项:row
    contentFileDict = excel.getFromExcel(dataFileName)
    rowDict = {'DAPING':u'大屏型号',
              'CHULIQI':u'处理器',
              'XIANCAI':u'线材长度',
              'ZHIJIA':u'支架',
              'YUNFEI':u'运费',
              'ANZHUANGFEI':u'安装费',
              }

    # 生成输入的界面
    lst = list(argumentDict.keys())+list(rowDict.keys())
       
    itemDict = {} # 组件数据
    
    for n,i in enumerate(lst):
        if i in argumentDict.keys():
            label = tk.Label(mainframe,text=argumentDict[i]+' :')
        else:
            label = tk.Label(mainframe,text=rowDict[i]+' :')
        label.grid(row=n,sticky=tk.E) #靠东
        if i == "XIANGMUMINGCHENG":
            content = tk.Entry(mainframe, width=48)
        else:
            content = ttk.Combobox(mainframe,width=45)
            if i in ["HANG","LIE"]:
                content['value'] = [ind for ind in range(1,15)]
            else:
                content['value'] = contentFileDict[i]
            content.current(0)
        itemDict[i] = content
        content.grid(row=n,column=1)    

    # 获取信息
    if not itemDict['XIANGMUMINGCHENG'].get():
        itemDict['XIANGMUMINGCHENG'].insert(0,'液晶拼接屏清单-宇视'+datetime.now().strftime('%Y%m%d%H%M%S'))
        # print(itemDict['XIANGMUMINGCHENG'].get())
    
    
    
    # 处理函数
    def do_replace():
        
        #计数显示
        global making_times
        making_times += 1        
        label = tk.Label(mainframe,text='    making..').grid(row=len(lst)*2+2,sticky=tk.W) 

        # 获取界面数据的字典
        keyDict = {}
        for key,value in itemDict.items():
            keyDict[key]= value.get()

        # 获取完整data数据的字典
        detailDict = excel.getDetail(dataFileName, keyDict)

        # 将detailDict的value-本身也是一个dict,全部拆解到一个平级的新dict中去用来替换
        replaceDict = {}
        for value in detailDict.values():
            replaceDict.update(value)

        # 新增部分key用来做特殊替换
        replaceDict["SHULIANG"] = int(keyDict["HANG"]) * int(keyDict["LIE"])
        
        # 生成Excel文件
        newFileName = itemDict['XIANGMUMINGCHENG'].get()+'.xlsx'
        excel.copyExcel(originFileName, newFileName, replaceDict)
        
        
        #for n,i in enumerate(lst):
        #    replaceDir[i] = entry_list[n].get()#!!!!!!h01768.decode(sys.stdin.encoding)
        
        label = tk.Label(mainframe,text='      '+str(making_times)+'  ok.     ').grid(row=len(lst)*2+2,sticky=tk.W)#靠右
        # print('ok')

    # 显示次数和点击按钮
    label = tk.Label(mainframe,text='      0          ').grid(row=len(lst)*2+2,sticky=tk.W)#靠右
    tk.Button(mainframe,text=u'生成大屏清单',width=15,height=2,command=do_replace).grid(row=len(lst)*2+2, column=1)

    mainframe.mainloop()


    
    """
    replaceLst = ['XIANGMUMINGCHENG',
                   'HANG',
                   'LIE',
                   'DAPING',
                   'CHULIQI',
                   'XIANCAI',
                   'ZHIJIA',
                   'YUNFEI',
                   'ANZHUANGFEI',
                   ]
    itemLst = ['MINGCHENG','CANSHU',"XINGHAO","DANJIA"]
        
                    
    contentKeys = 0
    

    
    contentFileDict = excel.getFromExcel('data.xlsx')    
    lst = replaceDir.keys()
    # 获取需要替换的子选项
    for n,i in enumerate(lst):
        label = tk.Label(mainframe,text=replaceDir[i]+':')
        label.grid(row=n,sticky=tk.W) #靠东
        if i == "XIANGMUMINGCHENG":
            content = tk.Entry(mainframe, width=35)
        else:
            content = ttk.Combobox(mainframe,width=35)
            if i in ["HANG","LIE"]:
                content['value'] = [ind for ind in range(1,15)]
            else:
                content['value'] = contentFileDict[i]
            content.current(0)
        content.grid(row=n,column=1)
    """

    """
    def on_combobox_select(event):
    # 获取Combobox当前选中的值
    selected_value = combobox.get()
    print("选中的值:", selected_value)

    # 绑定事件处理函数到Combobox的选中事件
    combobox.bind("<<ComboboxSelected>>", on_combobox_select)
    """

    """
    for i in replaceDir.keys():
        label = tk.Label(mainframe, text=i)
        label.grid(sticky=tk.E, row=0)
    
    label = tk.Label(mainframe, text=XiangMuMingCheng)
    label.grid(sticky=tk.E, row=0)
    entry = tk.Entry(mainframe, width=45)
    entry.grid(column=1, row=0)
    
    label = tk.Label(mainframe, text=Hang)
    label.grid(sticky=tk.E, row=1)
    combobox = ttk.Combobox(mainframe, values=[1,2,3,4,5])
    combobox.grid(column=1, row=1)
    """




    

if __name__=='__main__':
    main()
