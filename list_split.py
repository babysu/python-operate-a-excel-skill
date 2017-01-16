
from optparse import OptionParser
import sys
import xlrd,xlwt

STRING = '部门'
TITLE_ROW = 0
def split_file(filename):
	workbook = xlrd.open_workbook(filename)#
	sheet = workbook.sheet_by_index(1) #通过index选择你需要分割的那个sheet	
	Title=sheet.row_values(TITLE_ROW) #
	#print(Title)
	index = Title.index(STRING)#选择所需要的那一列数据
	#print(index)
	all= sheet.col_values(index)
	department = list(set(all))
	department.remove(STRING) #删除Title这一个元素得到的是所有的部门了
	#print(department)
	wb_result=xlwt.Workbook()
	for sub_dt in department:
		row_i =0 
		sheet_subdt=wb_result.add_sheet(sub_dt,cell_overwrite_ok=True)
		for j in range(sheet.ncols):
			sheet_subdt.write(row_i,j,sheet.row_values(TITLE_ROW)[j])
		row_i=row_i+1
		for i in range(1,sheet.nrows): #第1行是Titile，从第2行开始
			if sheet.row_values(i)[index] == sub_dt:
				for j in range(sheet.ncols):
					sheet_subdt.write(row_i,j,sheet.row_values(i)[j])
				row_i=row_i+1
			
	wb_result.save('result-split.xls')
		
			
	 
def main():
	parser = OptionParser(description="Query the stock's value.", usage="%prog [-f]", version="%prog 1.0")
	parser.add_option('-f', '--filename', dest='filename',
                      help="the filename that you need to split.")
	options, args = parser.parse_args(args=sys.argv[1:])
	stock = split_file(options.filename)
	
	
if __name__=='__main__':
	main()