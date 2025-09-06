import requests as rq
from bs4 import BeautifulSoup
from copy import deepcopy as dc
import json
from datetime import datetime

def split_list(l: list, sep) -> list[list]:
	result, temp = [], []
	for i in l:
		if i == sep:
			result.append(temp)
			temp = []
		else:
			temp.append(i)
	result.append(temp)

	return result

DATE = '20250907'

rqg = rq.get(f'https://universalis.com/{DATE}/mass.htm')
soup = BeautifulSoup(rqg.text, 'html.parser')
soup = soup.select_one('#innertexst')

section_types = {
	'First reading': 'reading1',
	'Responsorial Psalm': 'psalm',
	'Second reading': 'reading2',
	'Gospel Acclamation': 'acclamation',
	'Gospel': 'gospel',
	'Or:': 'alt'
}

sections, tmp = [], {}
for ele in soup.children:
	if ele.name is None: continue
	if ele.name == 'table':
		if tmp:
			sections.append(tmp)
		tmp = {'type': None, 'alt': False, 'ref': None, 'title': None, 'lines': [], 'text': []}
		
		section_type = ele.select_one('th[align="left"]').text.strip()
		reference = ele.select_one('th[align="right"]').text.strip()

		if section_types[section_type] == 'alt':
			tmp['type'] = sections[-1]['type']
			tmp['alt'] = True
		else:
			tmp['type'] = section_types[section_type]
		tmp['ref'] = reference
	elif ele.name == 'h4':
		tmp['title'] = ele.text.strip()
	elif ele.name == 'div':
		cl = ele.get('class', [])
		if any(c in cl for c in ('v', 'vi', 'p', 'pi')):
			if not tmp.get('lines'):
				tmp['lines'] = []
			tmp['lines'].append(ele.text.strip())

sections.append(tmp)

for section in sections:
	if section.get('type') == 'psalm':
		section['text'] = [section['lines'][0], *[i for i in split_list(section['lines'], section['lines'][0]) if i]]
	elif section.get('type') == 'acclamation':
		section['text'] = section['lines'][1:-1]
	else:
		section['text'] = dc(section['lines'])
	del section['lines']

final = {
	'success': True,
	'date': DATE,
	# 'timestamp': datetime.now().isoformat(),
	'readings': sections
}

with open('readings2.json', 'w', encoding='utf-8') as f:
	json.dump(final, f, indent=4)

# for section in sections:
# 	print(section)


