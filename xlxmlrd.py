# transpose and compile excel .xml format

# parse xml
# compose a list of sheets
# for each sheet
# 	compose a list of tables
# 	for each row in table
# 		compose a list of cells.

import xml.etree.ElementTree as ET

ns={'ss':"urn:schemas-microsoft-com:office:spreadsheet"}


class xlxmlrd:
	def open_workbook(file_name):
		global ns
		ET.register_namespace('ss', ns['ss'])
		tree = ET.parse(file_name)
		# print(tree)
		workbook = Workbook(tree.getroot())
		return workbook

class Cell:
	'''
	<ss:Cell><ss:Data ss:Type="String">Total Weight: 1137.00 g (40.11 oz-wt.)</ss:Data></ss:Cell>
	'''
	value = None

	def __init__(self, _data):
		global ns
		if _data == None:
			self.value = ""
			self.type = ""
		else:
			self.value = _data[0].text
			self.type = _data[0].attrib['{%s}Type' % ns['ss']]

class Worksheet:
	'''
	<ss:Worksheet ss:Name="page 1">
	<ss:Table>
		<ss:Row>
			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String">Sunday</ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
			<ss:Cell><ss:Data ss:Type="String">13 March, 2019</ss:Data></ss:Cell>
		</ss:Row>
	</ss:Table>
	</ss:Worksheet>

	'''
	node = None
	name = ""

	def __init__(self, _node):
		global ns
		# print(_node.attrib)
		# get the table node
		self.node = _node[0]
		# get the sheet name from attribute dictionary
		self.name = _node.attrib['{%s}Name' % ns['ss']]


	def cell(self, _row, _col):
		# print(_row, _col)
		# print("row: %s, col: %s" %(_row, _col))
		# print(self.node)
		try:
			cell_data = self.node[_row][_col]
		except IndexError:
			return Cell(None)
		cell = Cell(cell_data)
		return cell


class Workbook:
	'''
	<ss:Workbook>
	<ss:Worksheet ss:Name="page 1">
	<ss:Table>
		<ss:Row>
			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String">Sunday</ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
 			<ss:Cell><ss:Data ss:Type="String"></ss:Data></ss:Cell>
			<ss:Cell><ss:Data ss:Type="String">13 March, 2019</ss:Data></ss:Cell>
		</ss:Row>
	</ss:Table>
	</ss:Worksheet>
	</ss:Workbook>
	'''
	node = None

	def __init__(self, _node):
		self.node = _node

	def sheet_by_index(self, _index):
		return Worksheet(self.node[_index])
