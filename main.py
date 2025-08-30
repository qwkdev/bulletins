import os; os.system('cls')

from docx2pdf import convert
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.section import WD_ORIENTATION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

GLOBAL_FONT: str = 'Calibri'

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

def set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
	"""Sets all cell padding (in twips)"""
	tcPr = cell._element.tcPr
	tcMar = tcPr.find(qn('w:tcMar'))
	if tcMar is None:
		tcMar = OxmlElement('w:tcMar')
		tcPr.append(tcMar)
	for size, side in zip([top, start, bottom, end], ['top', 'start', 'bottom', 'end']):
		el = tcMar.find(qn(f'w:{side}'))
		if el is None:
			el = OxmlElement(f'w:{side}')
			tcMar.append(el)
		el.set(qn('w:w'), str(size))
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

def normalize_cell(cell, margins=True, paragraph_spacing=True):
	p = cell.paragraphs[0]
	cell._tc.remove(p._p)

	if margins: set_cell_margins(cell, 0, 0, 0, 0)
	if paragraph_spacing: zero_paragraph_spacing(cell)

def normalize_p(p, size, top=0, bottom=0):
	p.paragraph_format.space_before = Pt(top)
	p.paragraph_format.space_after = Pt(bottom)
	for run in p.runs:
		run.font.size = Pt(size)
		run.font.name = GLOBAL_FONT

def remove_blank_p(cell):
    paragraphs = list(cell.paragraphs)
    for p in paragraphs:
        if not p.text.strip():
            if len(cell.paragraphs) > 1:
                p._element.getparent().remove(p._element)

def parseText(obj, raw_text, size):
	text = raw_text.replace('<br><ul>', '<ul>').replace('</ul><br>', '</ul>')
	p = obj.add_paragraph()
	normalize_p(p, size)

	ctx, intag, txt = [], None, ''
	for c in text:
		if c == '<':
			intag = ''
		elif c == '>' and intag:
			if txt:
				run = p.add_run(txt)
				run.font.size = Pt(size)
				run.font.name = GLOBAL_FONT
				txt = ''
				for tag in ctx:
					if tag == 'b':
						run.bold = True
					elif tag == 'i':
						run.italic = True
					elif tag == 'u':
						run.underline = True
					elif tag == 's':
						run.font.superscript = True
			
			if 'ul' in intag or (ctx and ctx[-1] == 'ul'):
				if intag == '/ul':
					ctx = ctx[:-1]
					p = obj.add_paragraph(txt)
					normalize_p(p, size)
				elif intag == 'ul' or intag == 'br':
					if intag == 'ul':
						ctx.append('ul')
					p = obj.add_paragraph(txt, style="List Bullet")
					normalize_p(p, size)
					fmt = p.paragraph_format
					n = 16
					m = 1.5

					fmt.left_indent = Pt(n*m)
					fmt.first_line_indent = -Pt(n)

					
			else:
				if intag.startswith('/') and intag[1:] in ctx:
					ctx = ctx[::-1]
					ctx.remove(intag[1:])
					ctx = ctx[::-1]
				elif intag == 'br':
					# run.add_break()
					
					p = obj.add_paragraph(txt)
					normalize_p(p, size)
				else:
					ctx.append(intag)

			intag = None
		else:
			if intag is not None:
				intag += c
			else:
				txt += c
	
	run = p.add_run(txt)
	run.font.size = Pt(size)
	run.font.name = GLOBAL_FONT
	for tag in ctx:
		if tag == 'b':
			run.bold = True
		elif tag == 'i':
			run.italic = True
		elif tag == 'u':
			run.underline = True
		elif tag == 's':
			run.font.superscript = True

	remove_blank_p(obj)
			
doc = Document()

style = doc.styles['Normal']
style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

section = doc.sections[0]
section.orientation = WD_ORIENTATION.LANDSCAPE
section.page_width = Mm(297)
section.page_height = Mm(210)

margin = Mm(9)

section.top_margin = round(margin * 0.8)
section.bottom_margin = round(margin * 0.8)
section.left_margin = round(margin * 0.8)
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

# ’

data = [
	'''
<b>MASS TIMES IN OUR PASTORAL AREA</b><br>
Mass times are changing in our pastoral area from the <b>12<s>th</s> July</b> they will be:<br>
<ul>
St Joseph's <b>Church</b> - 4.30 pm Saturday Vigil (Confessions at 4 pm) and 11.30 am Sunday<br>
St <i>Andrew's</i> Church - 6 pm <s>Saturday</s> Vigil and 10 am Sunday<br>
</ul>
This change has become necessary due to Father John's illness and due to a lack of available 
priests in the archdiocese. This change has been approved by the Archbishop and the Dean, 
and will continue for the foreseeable future. We appreciate your understanding here.<br>
Changes for weekday masses in both parishes will also be announced in due course.
''', '''
<b>SECOND COLLECTION</b> The next second collection will be the 28/29<s>th</s> June for Peter's Pence.
''', '''
<b>ST NICHOLAS' PRIMARY 1 WELCOME EVENT</b> - <i>Sunday 22<s>nd</s> June from 1 pm to 3 pm</i><br>
In St Andrew's Church Hall. Open to all families to drop in for activities, refreshments and
meet the P6 buddies. Pre-loved uniforms are available. For children starting in August 2025.
''', '''
<b>APOSTOLIC NUNCIO, H.E. ARCHBISHOP MIGUEL<br>MAURY BUENDÍA VISIT TO GLASGOW</b><br>
<b>Sunday 22<s>nd</s> June:</b> Preside at the 12 noon Mass in Saint Andrew's Cathedral.<br>
<b>Sunday 22<s>nd</s> June:</b> Blessed Sacrament Procession in Croy, beginning 3.45 pm at Holy Cross
Church, then Eucharistic Procession through Village at 4 pm, return to Church for Benediction
at 5.15 pm.<br>
<b>Monday 23<s>rd</s> June:</b> Celebrate 1 pm Mass in Saint Andrew's Cathedral.
The Nuncio's will also visit Barlinnie Prison, Glasgow University and Glasgow Cathedral (meet
and pray with other church leaders). He will also celebrate Mass in the Carmelite Monastery
in Dumbarton, meet with Archdiocesan agencies (Evangelisation, Youth and SCIAF).
''', '''
<b>ABBA'S VINEYARD SACRED HEART PRAYER EVENING</b> - <i>Saturday 28<s>th</s> June from 5-9 pm</i><br>
For young adults aged 18-35. Gather in an evening for the Sacred Heart. Includes the
opportunity for confession, mass and dinner. All are welcome to join at any point. Address -<br>
Bl John Duns Scotus, 270 Ballater Street, Glasgow, G5 0YT. Organised by Abba's Vineyard.
For more information and a timetable search @abbasvineyard on social media or email:
abbasvineyard@gmail.com.
''', '''
<b>NICAEA 2025 - 1700<s>TH</s> ANNIVERSARY OF NICAEA</b> - <i>Sunday 22<s>nd</s> June at 3 pm</i><br>
Glasgow Churches Together invites you to Nicaea in St Andrew's Cathedral, Clyde Street.
Commemorating the legacy of faith and unity. Celebrate 1700 years since the First Council of
Nicaea, a cornerstone of Christian history. Experience a service filled with prayer, reflection
and sacred music. Witness the unity and enduring significance of the Nicene Creed. Be part
of a celebration of Nicaea's enduring legacy. Deepen your understanding of the Council of
Nicaea and its impact on spiritual traditions.
''', '''
<b>THANK YOU</b> Frances Gillian Millerick would like to say a very big thank you to those very kind
parishioners who came to her aid when she took unwell at Saturday night Mass and stayed
until the ambulance arrived. She is home now and feeling so much better.
''',


	'''<b>EXAMPLE FULL BLOCK HEADER</b><br>Example text blah blah blah...<br>More text idk''',
	'''a b c d E a b c d E a b c d E a b c d E a b c d E a b c d E a b c d E a b c d E a b c d E a b c d E'''
	'', '',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><ul>Example bullet 1<br>Example bullet 2<br>Example bullet 3</ul>More text idk v2''',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><br>Example text 2 blah blah blah...<br>More text idk v2''',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><br>Example text 2 blah blah blah...<br>More text idk v2''',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><br>Example text 2 blah blah blah...<br>More text idk v2''',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><br>Example text 2 blah blah blah...<br>More text idk v2''',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><br>Example text 2 blah blah blah...<br>More text idk v2''',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><br>Example text 2 blah blah blah...<br>More text idk v2''',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><br>Example text 2 blah blah blah...<br>More text idk v2''',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><br>Example text 2 blah blah blah...<br>More text idk v2''',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><br>Example text 2 blah blah blah...<br>More text idk v2''',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><br>Example text 2 blah blah blah...<br>More text idk v2''',
	'''<b>EXAMPLE FULL BLOCK HEADER 2</b><br>Example text 2 blah blah blah...<br>More text idk v2'''
]

data = [i.replace('\n', '') for i in data]

data_table = a5table.cell(0, 0).add_table(rows=len(data), cols=1)
data_table.autofit = True
data_table.allow_autofit = True

for txt, row in zip(data, data_table.rows):
	normalize_cell(row.cells[0], margins=False)
	set_cell_margins(row.cells[0], 150, 80, 100, 40)
	row.cells[0].width = left_half_width
	# row.cells[0].vertical_alignment = 1  # center
	parseText(row.cells[0], txt, 10)

# set_cell_background(data_table.cell(0, 0), '00FF00')

set_table_borders(data_table, color='000000', size=4, outer=False)
# data_table.autofit = False
# data_table.allow_autofit = False

doc.save('out.docx')
convert('out.docx', 'out.pdf')