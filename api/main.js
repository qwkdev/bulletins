const axios = require("axios");
const cheerio = require("cheerio");

function splitArray(arr, sep) {
	let result = [];
	let temp = [];
	arr.forEach(i => {
		if (i === sep) {
			result.push(temp);
			temp = [];
		} else {
			temp.push(i);
		}
	})
	result.push(temp);
	
	return result
}

const section_types = {
	'First reading': 'reading1',
	'Responsorial Psalm': 'psalm',
	'Second reading': 'reading2',
	'Gospel Acclamation': 'acclamation',
	'Gospel': 'gospel',
	'Or:': 'alt'
}

async function scrape(DATE) {
	const resp = await axios.get(`https://universalis.com/${DATE}/mass.htm`);
	const $ = cheerio.load(resp.data);
	const parent = $('#innertexst').first();

	let sections = [];
	let temp = {};
	parent.contents().each((_, ele) => {
		if (!ele.tagName) return;
		if (ele.tagName === 'table') {
			if (Object.keys(temp).length) sections.push(temp);
			temp = {
				type: null,
				alt: false,
				ref: null,
				title: null,
				lines: [],
				text: []
			}

			let section_type = $(ele).find('th[align="left"]').text().trim();
			let reference = $(ele).find('th[align="right"]').text().trim();

			if (section_types[section_type] === 'alt') {
				temp.type = sections.at(-1).type;
				temp.alt = true;
			} else {
				temp.type = section_types[section_type];
			}
			temp.ref = reference
		} else if (ele.tagName === 'h4') {
			temp.title = $(ele).text().trim();
		} else if (ele.tagName === 'div') {
			let classList = $(ele).attr('class')?.split(/\s+/) || [];
			if (['v', 'vi', 'p', 'pi'].some(c => classList.includes(c))) {
				if (!temp.lines) {
					temp.lines = [];
				}
				temp.lines.push($(ele).text().trim());
			}
		}
	});

	sections.push(temp);

	for (let section of sections) {
		if (section.type === 'psalm') {
			section.text = [
				section.lines[0],
				...splitArray(section.lines, section.lines[0]).filter(i => i.length)
			];
		} else if (section.type === 'acclamation') {
			section.text = section.lines.slice(1, -1);
		} else {
			section.text = structuredClone(section.lines);
		}

		delete section.lines;
	}

	return {
		success: true,
		date: DATE,
		readings: sections
	}
}

async function test() {
	const resp = await scrape('20250921');
	console.log(resp.readings);
	console.log(resp.readings[1].text);
}
test();