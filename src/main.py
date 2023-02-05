import sys
import os
import comtypes.client
import shutil
# 设置路径

srcpath = 'D:/BOOK/csbook'
despath = 'D:/BOOK/out'


def convert(inpath: str, outpath: str):
    # 创建PDF
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 2
    print(inpath.replace('/','\\'))
    slides = powerpoint.Presentations.Open(inpath.replace('/','\\'))
    # 保存PDF
    print(outpath.split('.')[0]+'.pdf')
    slides.SaveAs((outpath.split('.')[0]+'.pdf').replace('/','\\'), 32)
    slides.Close()
   

#path 是相对路径
def dfs(spath: str, dpath: str, path: str):
    for fn in os.listdir(spath+path):
        print(fn,os.path.isdir(path+'/'+fn))
        if os.path.isdir(spath+path+'/'+fn):
            dfs(spath, dpath, path+'/'+fn)
            pass
        elif fn.split('.')[-1] == 'pptx' or fn.split('.')[-1] == 'ppt':
            print('isppt')
            if not os.path.exists(dpath+path+'/'):
                os.makedirs(dpath+path+'/')
            convert(spath+path+'/'+fn, dpath+path+'/'+fn)
            pass
        elif fn.split('.')[-1] == 'pdf':
            print('ispdf')
            if not os.path.exists(dpath+path+'/'):
                os.makedirs(dpath+path+'/')
            shutil.copyfile(spath+path+'/'+fn, dpath+path+'/'+fn)
            
        else:
            print('bcdmm!')
    pass


if __name__ == '__main__':
    dfs(srcpath,despath,'')
    pass
