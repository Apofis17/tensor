 # -- coding: utf-8 --
 
from ClassXLS import XLS
import math


xls = XLS()
analysis = xls.analysis()
s = list(analysis.items())
s.sort(key=lambda item: item[1])
for i in s:
    the = ['не', ''][i[1]>0]
    print '%s %s отработал %s'%(i[0].encode('utf8'), the, int(math.fabs(i[1])))