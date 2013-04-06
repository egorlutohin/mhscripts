import win32com.client
import os.path
import sys
import glob

cells_wnum_s2 = [ # ячейки с цифрами на втором листе
'BP26', 'CB26', 'CN26', 'DA26', 'DM26', 'DY26', 'EN26',
'BP28', 'CB28', 'CN28', 'DA28', 'DM28',
'BP32', 'CB32', 'CN32', 'DA32', 'DM32',
'BP37', 'CB37', 'CN37', 'DA37', 'DM37',
'BP41', 'CB41', 'CN41', 'DA41',
'BP42', 'CB42', 'CN42', 'DA42',
'BP44', 'CB44', 'CN44', 'DA44', 'DM44',
'BP46', 'CB46', 'CN46', 'DA46', 'DM46',
'BP53', 'CB53', 'CN53', 'DA53', 'DM53', 'DY53', 'EN53',
'BP55', 'CB55', 'CN55', 'DA55', 'DM55', 'DY55', 'EN55',
'BP56', 'CB56', 'CN56', 'DA56', 'DM56', 'DY56', 'EN56',
'BP58', 'CB58', 'CN58', 'DA58', 'DM58', 'DY58', 'EN58',
'BP60', 'CB60', 'CN60', 'DA60', 'DM60', 'DY60', 'EN60',
'BP61', 'CB61', 'CN61', 'DA61', 'DM61', 'DY61', 'EN61',
'BP66', 'CB66', 'CN66', 'DA66', 'DM66', 'DY66',
'BP68', 'CB68', 'CN68', 'DA68', 'DM68',
'BP69', 'CB69', 'CN69', 'DA69', 'DM69',
'BP73', 'CB73', 'CN73', 'DA73',
'BP75', 'CB75', 'CN75', 'DA75',
'BP80', 'CB80', 'CN80', 'DA80', 'DM80',
'BP82', 'CB82', 'CN82', 'DA82', 'DM82',
'BP85', 'CB85', 'CN85', 'DA85', 'DM85',
'BP88', 'CB88', 'CN88', 'DA88', 'DM88',
'BP90', 'CB90', 'CN90', 'DA90', 'DM90',
'BP93', 'CB93', 'CN93', 'DA93', 'DM93',
'BP95', 'CB95', 'CN95', 'DA95', 'DM95',
'CN99', 'DA99',
'BP101', 'CB101', 'CN101', 'DA101', 'DM101',
'BP105', 'CB105', 'CN105', 'DA105', 'DM105',
'BP107', 'CB107', 'CN107', 'DA107', 'DM107',
'BP108', 'CB108', 'CN108', 'DA108', 'DM108',
'BP111', 'CB111', 'CN111', 'DA111', 'DM111', 'DY111', 'EN111',
'BP115', 'CB115', 'CN115', 'DA115', 'DM115', 
'BP119', 'CB119', 'CN119', 'DA119', 'DM119',
'BP125', 'CB125', 'CN125', 'DA125', 'DM125',
'BP131', 'CB131', 'CN131', 'DA131',
'BP134', 'CB134', 'CN134', 'DA134',
'BP137', 'CB137', 'CN137', 'DA137', 'DM137',
'BP139', 'CB139', 'CN139', 'DA139', 'DM139',
'BP143', 'CB143', 'CN143', 'DA143', 'DM143', 'DY143', 'EN143',
'BP149', 'CB149', 'CN149', 'DA149', 'DM149', 'DY149', 'EN149',
'BP151', 'CB151', 'CN151', 'DA151', 'DM151', 'DY151', 'EN151',
'BP152', 'CB152', 'CN152', 'DA152', 'DM152', 'DY152', 'EN152',
'BP154', 'CB154', 'CN154', 'DA154', 'DM154', 'DY154', 'EN154']

cells_wnum_s3 = [ # ячейки с цифрами на третьем листе
'BS9', 'CJ9', 'DA9', 'DR9', 'EI9',
'BS10', 'CJ10', 'DA10', 'DR10', 'EI10',
'BS11', 'CJ11', 'DA11', 'DR11', 'EI11',
'BS13', 'CJ13', 'DA13', 'DR13', 'EI13',
'BS14', 'CJ14', 'DA14', 'DR14', 'EI14',
'BS15', 'CJ15', 'DA15', 'DR15', 'EI15',
'BS16', 'CJ16', 'DA16', 'DR16', 'EI16',
'DA17', 'DR17',
'BS18', 'CJ18', 'DA18', 'DR18', 'EI18',
'BB20',
'BB21', 'BS21', 'CJ21', 'DA21', 'DR21', 'EI21',
'BB22', 'BS22', 'CJ22', 'DA22', 'DR22', 'EI22',
'BB23', 'BS23', 'CJ23', 'DA23', 'DR23', 'EI23'
]


def parse_float(v):
	if type(v) == float:
		return v
	elif type(v) == str:
		v = str(v)
		v = v.strip().replace(',', '.')
		if v is '':
			v = 0
		else:
			v = float(v)
		return v
	elif type(v) == int:
		v = int(v)
		return v
	elif v == None:
		return 0


pwd  = unicode(os.path.dirname(__file__), 'cp1251')

def prints(s): # print to console
	print s.encode('cp866')
#~ import win32api, win32con
#~ win32api.MessageBox(0, "Question", "Title", win32con.MB_YESNO)	

try:
	RESULT_FILE = unicode(sys.argv[1], 'cp1251')
except:
	RESULT_FILE = u"Результат.xls"
	
if not os.path.exists(RESULT_FILE):
	prints(u"Результирующего файла не существует, непонятно куда складывать все данные")
	sys.exit(0)

excel_files = set(glob.iglob(u'*.xls')) - set((RESULT_FILE,))

excel = win32com.client.Dispatch("Excel.Application")
excel.visible = True
#~ excel2 = win32com.client.DispatchEx("Excel.Application")

#~ excel1.visible = True
#~ excel2.visible = True

full_path = os.path.join(pwd, RESULT_FILE)
for wb in excel.workbooks:
	if wb.fullname == full_path:
		to_wb = wb
		#~ excel1 = excel
		break;
else:
	if len(excel.workbooks) == 0:
		excel1 = excel
	else:
		excel1 = win32com.client.DispatchEx("Excel.Application")
		excel1.visible = True
	to_wb = excel1.workbooks.Open(full_path)
	



for i in excel_files:
	full_path = os.path.join(pwd, i)
	for wb in excel.workbooks:
		if wb.fullname == full_path:
			from_wb = wb
			#~ excel2 = excel
			break;
	else:
		excel2 = win32com.client.DispatchEx("Excel.Application")
		excel2.visible = True
		from_wb = excel2.workbooks.Open(os.path.join(pwd, i))
	
	# Для второго листа
	to_cs = to_wb.worksheets[1] #curent sheet
	from_cs = from_wb.worksheets[1]
	READ_ERROR2 = []
	WRITE_ERROR2 = []
	for c in cells_wnum_s2:
		try:
			from_v = parse_float(from_cs.Range(c).value)
		except:
			READ_ERROR2.append(c)
			continue
			
		try:
			to_v = parse_float(to_cs.Range(c).value)
			to_cs.Range(c).value = to_v + from_v
		except:
			WRITE_ERROR2.append(c)
			continue
			
	prints(u"Второй лист файла \"%s\" добавлен к файлу \"%s\"" % (i, RESULT_FILE))
	prints(u"Ошибки чтения - %d шт, ячейки: %s" % (len(READ_ERROR2), ', '.join(READ_ERROR2)))
	prints(u"Ошибки записи - %d шт, ячейки: %s" % (len(WRITE_ERROR2), ', '.join(WRITE_ERROR2)))
	print ''
	
	# Для третьего листа
	to_cs = to_wb.worksheets[2] #curent sheet
	from_cs = from_wb.worksheets[2]
	READ_ERROR3 = []
	WRITE_ERROR3 = []
	for c in cells_wnum_s3:
		try:
			from_v = parse_float(from_cs.Range(c).value)
		except:
			READ_ERROR3.append(c)
			continue
			
		try:
			to_v = parse_float(to_cs.Range(c).value)
			to_cs.Range(c).value = to_v + from_v
		except:
			WRITE_ERROR3.append(c)
			continue
			
	prints(u"Третий лист файла \"%s\" добавлен к файлу \"%s\"" % (i, RESULT_FILE))
	prints(u"Ошибки чтения - %d шт, ячейки: %s" % (len(READ_ERROR3), ', '.join(READ_ERROR3)))
	prints(u"Ошибки записи - %d шт, ячейки: %s" % (len(WRITE_ERROR3), ', '.join(WRITE_ERROR3)))
	print ''
	
	
	print(u"Для окончания нажмите enter")
	raw_input()
	break;
	
	#~ cs.Range('a12').value 
	#~ cs.Cells(row, col).value
#~ excel.quit()

