import io
import unittest
from datetime import date
from openpyxl import Workbook, load_workbook
from core import count_procurement, electricity, build_report, formula_number


def blob(wb):
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def procure(rows, duplicate=False):
    wb = Workbook()
    ws = wb.active
    ws.append(['采购项目名称','中标日期','参与批次','投标供应商'])
    for row in rows: ws.append(row)
    if duplicate:
        other = wb.create_sheet('重复明细')
        for row in ws.values: other.append(row)
    return blob(wb)


def templates():
    wb = Workbook()
    s = wb.active
    s.title = '服务双碳目标贡献力-9月'
    for addr, value in {'B9':'全年投标人的投标家次数（次）','B12':'全年投标人的投标包次数量（次）',
        'B21':'总建筑面积（㎡）','B23':'其中：外购使用电力（MWh）','B29':'汽油（t）',
        'C9':"='9月采购、油、电'!K9",'C12':71,'C23':132.1931,'C29':.4809459,
        'C5':'=C6+C15','C6':'=C9*0.0000144+C12*0.00000003','C15':0}.items(): s[addr] = value
    a = wb.create_sheet('9月采购、油、电')
    a['B9'] = 9
    a['K9'] = '=52+7'
    a['E9'] = .4809459
    a['I9'] = 132.1931
    last = Workbook()
    l = last.active
    l['A2'] = date(2025,10,31)
    for row,label,value in [(9,'全年投标人的投标家次数（次）',108),
        (12,'全年投标人的投标包次数量（次）',108),(21,'总建筑面积（㎡）',1066.11),
        (22,'用电量（MWh）',118.6782),(27,'汽油（t）',1.1614)]:
        l.cell(row,2).value=label; l.cell(row,3).value=value
    return blob(wb),blob(last)


class CoreTests(unittest.TestCase):
    def test_61_companies_are_69_times_in_three_batches(self):
        a=[f'A{i:02}' for i in range(27)]
        b=a[:3]+[f'B{i:02}' for i in range(23)]
        c=a[3:7]+b[3:4]+[f'C{i:02}' for i in range(11)]
        self.assertEqual(len(set(a+b+c)),61)
        rows=[]
        for batch,names,extra in [('批4',a,21),('批5',b,9),('标2',c,8)]:
            rows.extend([[batch+'包1',date(2026,10,15),batch,'\n'.join(names)],
                         [batch+'包2',date(2026,10,15),batch,'\n'.join(names[:extra])]])
        p=count_procurement(procure(rows),2026,10)
        self.assertEqual((p.company_times,p.package_times),(69,107))

    def test_cross_batch_count(self):
        rows=[['包1',date(2026,10,15),'批1','甲\n乙'],['包2',date(2026,10,15),'批1','甲'],
              ['包3',date(2026,10,24),'批2','甲\n丙'],['历史',date(2026,9,23),'批0','丁']]
        p=count_procurement(procure(rows,True),2026,10)
        self.assertEqual((p.company_times,p.package_times),(4,5))

    def test_month_and_pending(self):
        p=count_procurement(procure([['待定','TBD','批1','甲'],['历史',date(2026,9,1),'批0','乙']]),2026,10)
        self.assertEqual((p.company_times,p.package_times,p.pending),(0,0,1))

    def test_batch_required(self):
        with self.assertRaises(ValueError):
            count_procurement(procure([['包1',date(2026,10,1),None,'甲']]),2026,10)

    def test_same_package_conflict(self):
        with self.assertRaises(ValueError):
            count_procurement(procure([['包1',date(2026,10,1),'批1','甲'],['包1',date(2026,10,1),'批1','乙']]),2026,10)

    def test_meter_readings_not_cached_formula(self):
        wb=Workbook();s=wb.active
        for meter,old,new,factor in [('AL4(C11)',10,12.9,'/'),('AL5(C11)',20,64.8,30),('AL6(C11)',100,601.4,30)]:
            s.append([meter,old,new,factor,'=1+1'])
        self.assertAlmostEqual(electricity(blob(wb))[0],16388.9)

    def test_all_meters_required(self):
        wb=Workbook();wb.active.append(['AL4(C11)',1,2,'/'])
        with self.assertRaises(ValueError): electricity(blob(wb))

    def test_report_and_missing_procurement(self):
        prev,last=templates()
        result,summary,warnings=build_report(prev,last,2026,10,date(2026,11,3),0,16388.9,None)
        self.assertEqual(summary['累计家次'],59)
        self.assertAlmostEqual(summary['累计外购电MWh'],148.582)
        self.assertEqual(summary['本月家次'],None)
        self.assertTrue(warnings)
        wb=load_workbook(io.BytesIO(result))
        self.assertEqual(wb.worksheets[0]['D29'].value,1.1614)
        self.assertEqual(wb.worksheets[1]['B10'].value,10)
        self.assertEqual(formula_number(wb,wb.worksheets[0],'C9'),59)
        self.assertNotIn('9月采购',wb.worksheets[0]['C9'].value)

    def test_generated_report_reusable_without_caches(self):
        prev,last=templates()
        p=count_procurement(procure([['包1',date(2026,10,1),'批1','甲\n乙']]),2026,10)
        data,summary,_=build_report(prev,last,2026,10,date(2026,11,3),0,1000,p)
        wb=load_workbook(io.BytesIO(data))
        self.assertEqual(formula_number(wb,wb.worksheets[0],'C9'),61)
        self.assertEqual(formula_number(wb,wb.worksheets[0],'C12'),73)
        self.assertAlmostEqual(formula_number(wb,wb.worksheets[0],'C23'),133.1931)


if __name__=='__main__': unittest.main()
