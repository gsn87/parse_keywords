
__author__ = 'gsn'


import xlrd
import collections

PATH = './input/keywords.xlsx'
SHEET_RESULT = 'Cleaning'
SHEET_INPUT = 'Input'
WORDS = ['table', 'bureau', 'bibliotheque', 'buffet', 'commode', 'table basse', 'console', 'meuble tv', 'meuble tele', 'meuble television', 'dressing', 'placard', 'banc', 'chaise', 'meuble hifi', 'rangement', 'table de nuit', 'table de chevet', 'etagere', 'table d\'appoint', 'tabouret']


wb = xlrd.open_workbook('./input/keywords.xlsx')

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
		total_keywords = len(self.dict)
		print ('%s \t %s \t %s' % ('Seed', 'Percent', 'Volume'))
		for word in self.words_list:
			self.occurence(word)
		for word in self.words:
			percent = len(self.words[word])*1./total_keywords
			volume = self.get_volume_from_word(word)
			print ('%s \t %f \t %i' % (word, percent, volume))


if __name__ == "__main__":
	keywords = ParseKeywords(PATH, SHEET_INPUT, SHEET_RESULT, WORDS)
	keywords.print_keyword_weight()