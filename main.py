import os; os.system('cls')

from docx2pdf import convert
from docx import Document
from docx.shared import Mm
from docx.enum.section import WD_ORIENTATION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_table_borders(table, color='000000', size=4, outer=True):
	"""
	Apply borders to a whole table.
	color = hex RGB (e.g. "FF0000" for red)
	size = border size in 1/8 pt (e.g. 8 = 1pt)
	"""
	tblPr = table._element.tblPr
	# for el in tblPr.findall(qn("w:tblBorders")):
	# 	tblPr.remove(el)
	# for el in tblPr.findall(qn("w:tblStyle")):
	# 	tblPr.remove(el)

	tblBorders = OxmlElement('w:tblBorders')
	for border_name in (*(['top', 'left', 'bottom', 'right'] if outer else []), 'insideH', 'insideV'):
		border = OxmlElement(f'w:{border_name}')
		border.set(qn('w:val'), 'single')
		border.set(qn('w:sz'), str(size))
		border.set(qn('w:space'), '0')
		border.set(qn('w:color'), color)
		tblBorders.append(border)
	tblPr.append(tblBorders)

def remove_cell_borders(cell):
	tcPr = cell._element.tcPr
	tcBorders = tcPr.find(qn('w:tcBorders'))
	if tcBorders is not None:
		tcPr.remove(tcBorders)
	tcBorders = OxmlElement('w:tcBorders')
	for side in ('top', 'left', 'bottom', 'right'):
		border = OxmlElement(f'w:{side}')
		border.set(qn('w:val'), 'nil')
		tcBorders.append(border)
	tcPr.append(tcBorders)


def remove_cell_margins(cell):
	"""Remove all cell padding"""
	tcPr = cell._element.tcPr
	tcMar = tcPr.find(qn('w:tcMar'))
	if tcMar is None:
		tcMar = OxmlElement('w:tcMar')
		tcPr.append(tcMar)
	for side in ['top', 'start', 'bottom', 'end']:
		el = tcMar.find(qn(f'w:{side}'))
		if el is None:
			el = OxmlElement(f'w:{side}')
			tcMar.append(el)
		el.set(qn('w:w'), '0')  # zero twips
		el.set(qn('w:type'), 'dxa')

def set_cell_background(cell, color="FFFF00"):
	"""
	Set the background color of a cell.
	color = RGB hex string, e.g. "FF0000" for red
	"""
	tc = cell._element
	tcPr = tc.get_or_add_tcPr()
	shd = tcPr.find(qn('w:shd'))
	if shd is None:
		shd = OxmlElement('w:shd')
		tcPr.append(shd)
	shd.set(qn('w:val'), 'clear')
	shd.set(qn('w:color'), 'auto')
	shd.set(qn('w:fill'), color)  # this is the background fill

def zero_paragraph_spacing(cell):
	"""Remove spacing before and after all paragraphs in a cell."""
	for paragraph in cell.paragraphs:
		paragraph.paragraph_format.space_before = 0
		paragraph.paragraph_format.space_after = 0
		paragraph.paragraph_format.line_spacing = 1  # optional

def normalize_cell(cell):
	p = cell.paragraphs[0]
	cell._tc.remove(p._p)

	remove_cell_margins(cell)
	zero_paragraph_spacing(cell)

doc = Document()

section = doc.sections[0]
section.orientation = WD_ORIENTATION.LANDSCAPE
section.page_width = Mm(297)
section.page_height = Mm(210)

margin = Mm(9)

section.top_margin = round(margin * 0.8)
section.bottom_margin = round(margin * 0.8)
section.left_margin = round(margin)
section.right_margin = round(margin * 1.2)
middle_margin = (margin * .9, margin * .4)

left_half_width = round(
	(section.page_width / 2)
	- section.left_margin
	- (middle_margin[0] / 2)
)
right_half_width = round(
	(section.page_width / 2)
	- section.right_margin
	- (middle_margin[1] / 2)
)
a5table = doc.add_table(rows=1, cols=3)
a5table.autofit = False
a5table.allow_autofit = False
a5table.rows[0].height = section.page_height - section.top_margin - section.bottom_margin

a5table.cell(0, 0).width = left_half_width
a5table.cell(0, 1).width = sum(middle_margin)
a5table.cell(0, 2).width = right_half_width

remove_cell_borders(a5table.cell(0, 1))
normalize_cell(a5table.cell(0, 0))
# normalize_cell(a5table.cell(0, 2))

set_table_borders(a5table, color='000000', size=4)

data = [
	'''<b>EXAMPLE FULL BLOCK HEADER</b><br>Example text blah blah blah...<br>More text idk''',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><br>Example text 2 blah blah blah...<br>More text idk v2''',
	'', ''
]
data_table = a5table.cell(0, 0).add_table(rows=len(data), cols=1)

for txt, row in zip(data, data_table.rows):
	row.cells[0].width = left_half_width
	# row.cells[0].vertical_alignment = 1  # center
	row.cells[0].text = txt

set_cell_background(data_table.cell(0, 0), '00FF00')

set_table_borders(data_table, color='FF0000', size=4)
# data_table.autofit = False
# data_table.allow_autofit = False

doc.save('out.docx')
convert('out.docx', 'out.pdf')