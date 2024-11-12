from read_proses import ReadFormulaXlsx as rfx




test  = rfx('./Madela_Template/10001.xlsx')
test.findFormula()
print(test.fromulaList)
test.endReadProses()