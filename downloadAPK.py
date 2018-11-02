import requests,sys,bs4,re,xlrd,xlwt,os
from xlutils.copy import copy

wbR=xlrd.open_workbook("TOP15应用.xls",formatting_info=True)
wsR=wbR.sheets()[0]

wbW=copy(wbR)
wsW=wbW.get_sheet(0)


proxies = {}

#requests.packages.urllib3.disable_warnings()

#for m in range(1,4):
def Download(m):
    apkName=wsR.cell(m,1).value
    apkVersion=wsR.cell(m,2).value
    
    url="http://sj.qq.com/myapp/detail.htm?apkName="+''.join(apkName)
    res=requests.get(url,proxies=proxies,verify=False)
    res.raise_for_status()

    soup=bs4.BeautifulSoup(res.text)
    linklist=soup.select('div.det-ins-btn-box a')
    apkurl=linklist[1].get('data-apkurl')

    pattern=re.compile(u"(?<=_).*?(?=.apk)")
    apkVersionNew=pattern.findall(apkurl)[0]

    if apkVersionNew==apkVersion:
        print("%s No updata"%apkName)
        wsW.write(m,3,'N')
    else:
        resource=requests.get(apkurl,proxies=proxies,stream=True,verify=False)
        print("%s is loading..."%apkName)
        with open("D:\\STU\\stu_WEB\\apk\\%s_%s.apk"%(apkName,apkVersionNew),mode='wb') as fh:
            for chunk in resource.iter_content(chunk_size=512):
                fh.write(chunk)
        wsW.write(m,3,'Y')
        wsW.write(m,2,apkVersionNew)
        FileNameOld=apkName+'_'+apkVersion+'.apk'
        #os.remove(os.path.join('D:\\STU\\stu_WEB\\apk',FileNameOld))
        
'''
for m in range(1,wsR.nrows):
    Download(m)
    
wbW.save('APPlist.xls')
print("Done...")
'''

def test():
    for m in range(1, wsR.nrows):
        Download(m)
    wbW.save('APPlist.xls')
    print("Done...")

from timeit import Timer

if __name__ =='__main__':
    t=Timer("test()","from __main__ import test")
    print(t.timeit(1))
