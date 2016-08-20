
__author__ = 'gsn'


import xlrd
import collections

PATH = './input/keywords_table_chevet.xlsx'
SHEET_RESULT = 'Cleaning'
SHEET_INPUT = 'Input'
# WORDS = ['table', 'bureau', 'bibliotheque', 'buffet', 'commode', 'table basse', 'console', 'meuble tv', 'meuble tele', 'meuble television', 'dressing', 'placard', 'banc', 'chaise', 'meuble hifi', 'rangement', 'table de nuit', 'table de chevet', 'etagere', 'table d\'appoint', 'tabouret']
WORDS = ['design', 'verre', 'bois', 'extensible', 'scandinave', 'rallonge', 'industrielle', 'mesure', 'television', 'contreplaque', 'laque', 'tv', 'beton', 'ronde', 'acier', 'vintage', 'style', 'massif', 'metal', 'contemporain']

class ParseKeywords():

	dict = {}
	words = {}

	def __init__(self, path, sheet_input, sheet_result, words_list):
		self.path = path
		self.sheet_input = sheet_input
		self.sheet_result = sheet_result
		self.words_list = words_list

	def get_sheet(self, sheet):
		wb = xlrd.open_workbook(self.path)
		sh = wb.sheet_by_name(sheet)
		return sh

	def get_total_dict_volume(self):
		self.parse_sheet_result()
		total_volume = 0
		for keyword in self.dict:
			total_volume += self.dict[keyword]["volume"]
		return total_volume


	def parse_sheet_result(self):
		sh = self.get_sheet(self.sheet_result)
		for row in range(1, sh.nrows):
			keyword = sh.cell_value(row, 1)
			requests = sh.cell_value(row, 3)
			compet = sh.cell_value(row, 4)
			bid = sh.cell_value(row, 4)
			seed = sh.cell_value(row, 0)
			self.dict[keyword] = {}
			self.dict[keyword]["volume"] = requests
			self.dict[keyword]["competition"] = compet
			self.dict[keyword]["bid"] = bid
			self.dict[keyword]["type"] = seed

	def counting_words(self):
		self.parse_sheet_result()
		all_words = []
		for keyword in self.dict:
			list = keyword.split(' ')
			all_words += list
		cnt = collections.Counter(all_words)
		for w in cnt:
			print('%s \t %i' % (w, cnt[w]))


	def occurence(self, word):
		self.words[word] = []
		for keyword in self.dict:
			if word in keyword:
				self.words[word].append(keyword)

	def get_volume_from_word(self, word):
		volume = 0
		for w in self.words[word]:
			volume += int(self.dict[w]["volume"])
		return volume

	def print_keyword_weight(self):
		self.parse_sheet_result()
		total_volume = self.get_total_dict_volume()
		print ('%s \t %s \t %s' % ('Seed', 'Percent', 'Volume'))
		for word in self.words_list:
			self.occurence(word)
		for word in self.words:
			volume = self.get_volume_from_word(word)
			percent = volume/total_volume
			print ('%s \t %f \t %i' % (word, percent, volume))


if __name__ == "__main__":
	keywords = ParseKeywords(PATH, SHEET_INPUT, SHEET_RESULT, WORDS)
	# keywords.counting_words()
	keywords.print_keyword_weight()