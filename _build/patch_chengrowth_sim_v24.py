# ver2.3 → ver2.4：アジェンダv11に合わせ、通知メッセ有(SIM1)と登録0.5%(SIM2)のシートを外し、SIM3だけにする
import sys
from openpyxl import load_workbook
from openpyxl.workbook.properties import CalcProperties
SRC,OUT=sys.argv[1],sys.argv[2]
wb=load_workbook(SRC)
for n in ('SIM1_通知メッセ有','SIM1_比較','SIM2_通知メッセ無','SIM2_比較'):
    del wb[n]
wb.active=0
wb.calculation=CalcProperties(fullCalcOnLoad=True)
wb.save(OUT)
