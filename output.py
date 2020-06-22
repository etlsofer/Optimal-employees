import pandas as pd
from pandas import ExcelWriter
import xlsxwriter



class WriteArrangement:
    arrangement = []
    def __init__(self, result: list):
        self.arrangement = result


    def write(self):
        hedear = ["שבוע","יום","תאריך"]
        region = []
        days = [self.arrangement[0].date]
        date = self.arrangement[0].date
        for worker in self.arrangement:
            if worker.date is not date:
                break
            region += [worker.region]

        for worker in self.arrangement:
            if worker.date is not date:
                days += [worker.date]
                date = worker.date

        hedear += region
        sheet1 = [i+1 for i in range(len(days))]
        sheet2 = ["ד" for i in range(len(days))]
        sheet3 = days
        sheet4 = []
        sheet5 = []
        sheet6 = []
        sheet7 = []
        sheet8 = []
        sheet9 =[]
        sheet = [sheet1, sheet2, sheet3, sheet4, sheet5, sheet6, sheet7, sheet8, sheet9]
        for sh in range(3,len(sheet)):
            for day in range(len(days)):
                for item in self.arrangement:
                    if item.date == days[day] and item.region == hedear[sh]:
                        sheet[sh] += [item.worker]
                        break
        hedear.reverse()
        sheet.reverse()
        df = pd.DataFrame({hedear[i]: sheet[i] for i in range(len(sheet))})
        writer = ExcelWriter('res/res.xlsx')
        df.to_excel(writer, index=False,inf_rep=8)
        writer.save()


        # hiding all empty columns and rows usung copy to other file with "xlsxwriter"
        workbook = xlsxwriter.Workbook('res/res.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.add_table(0,0,5,8,{'autofilter': False})
        data = pd.read_excel('res/res.xlsx', header=0)
        i=1
        j=0
        for item in hedear:
            worksheet.write(0, j, item)
            j = (j + 1) % len(hedear)
        for row in data.values:
            for item in row:
                worksheet.write(i,j,item)
                j =(j+1)%len(row)
            i = i+1
        worksheet.set_default_row(hide_unused_rows=True)
        worksheet.set_column('J:XFD', None, None, {'hidden': True})
        workbook.close()


