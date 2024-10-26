# -*- coding: cp936 -*-
import tkinter as tk
from tkinter import ttk
import excel
from datetime import datetime


making_times = 0


def main():
    # ģ���ļ�����
    originFileName = 'LCD_template.xlsx'

    # �����ļ�����
    dataFileName = 'data.xlsx'

    # ���߱���
    mainframe = tk.Tk()
    mainframe.title(u'һ��LCD�䵥���ɹ���V1.0')

    # ����
    """
    XiangMuMingCheng = u'��Ŀ����'
    Hang = u'��'
    Lie = u'��'
    DAPING = u'�����ͺ�'
    ChuLiQi = u'������'
    XianCai = u'�߲ĳ���'
    ZhiJia = u'֧��'
    YunFei = u'�˷�'
    AnZhuangFei = u'��װ��'
    """

    # �������:argument
    argumentDict = {'XIANGMUMINGCHENG':u'��Ŀ����',
                   'HANG':u'��',
                   'LIE':u'��',
                   }

    # �������:row
    contentFileDict = excel.getFromExcel(dataFileName)
    rowDict = {'DAPING':u'�����ͺ�',
              'CHULIQI':u'������',
              'XIANCAI':u'�߲ĳ���',
              'ZHIJIA':u'֧��',
              'YUNFEI':u'�˷�',
              'ANZHUANGFEI':u'��װ��',
              }

    # ��������Ľ���
    lst = list(argumentDict.keys())+list(rowDict.keys())
       
    itemDict = {} # �������
    
    for n,i in enumerate(lst):
        if i in argumentDict.keys():
            label = tk.Label(mainframe,text=argumentDict[i]+' :')
        else:
            label = tk.Label(mainframe,text=rowDict[i]+' :')
        label.grid(row=n,sticky=tk.E) #����
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

    # ��ȡ��Ϣ
    if not itemDict['XIANGMUMINGCHENG'].get():
        itemDict['XIANGMUMINGCHENG'].insert(0,'Һ��ƴ�����嵥-����'+datetime.now().strftime('%Y%m%d%H%M%S'))
        # print(itemDict['XIANGMUMINGCHENG'].get())
    
    
    
    # ������
    def do_replace():
        
        #������ʾ
        global making_times
        making_times += 1        
        label = tk.Label(mainframe,text='    making..').grid(row=len(lst)*2+2,sticky=tk.W) 

        # ��ȡ�������ݵ��ֵ�
        keyDict = {}
        for key,value in itemDict.items():
            keyDict[key]= value.get()

        # ��ȡ����data���ݵ��ֵ�
        detailDict = excel.getDetail(dataFileName, keyDict)

        # ��detailDict��value-����Ҳ��һ��dict,ȫ����⵽һ��ƽ������dict��ȥ�����滻
        replaceDict = {}
        for value in detailDict.values():
            replaceDict.update(value)

        # ��������key�����������滻
        replaceDict["SHULIANG"] = int(keyDict["HANG"]) * int(keyDict["LIE"])
        
        # ����Excel�ļ�
        newFileName = itemDict['XIANGMUMINGCHENG'].get()+'.xlsx'
        excel.copyExcel(originFileName, newFileName, replaceDict)
        
        
        #for n,i in enumerate(lst):
        #    replaceDir[i] = entry_list[n].get()#!!!!!!h01768.decode(sys.stdin.encoding)
        
        label = tk.Label(mainframe,text='      '+str(making_times)+'  ok.     ').grid(row=len(lst)*2+2,sticky=tk.W)#����
        # print('ok')

    # ��ʾ�����͵����ť
    label = tk.Label(mainframe,text='      0          ').grid(row=len(lst)*2+2,sticky=tk.W)#����
    tk.Button(mainframe,text=u'���ɴ����嵥',width=15,height=2,command=do_replace).grid(row=len(lst)*2+2, column=1)

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
    # ��ȡ��Ҫ�滻����ѡ��
    for n,i in enumerate(lst):
        label = tk.Label(mainframe,text=replaceDir[i]+':')
        label.grid(row=n,sticky=tk.W) #����
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
    # ��ȡCombobox��ǰѡ�е�ֵ
    selected_value = combobox.get()
    print("ѡ�е�ֵ:", selected_value)

    # ���¼���������Combobox��ѡ���¼�
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
