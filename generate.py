from table import SpreadSheet


sheet = SpreadSheet()
for i in range(2):
	sheet.createSheet(f"{i}")

	table = sheet.sheets[i].table

	# Set a 2D range of values
	table[0:100][0:100].value = [
	[f"({x*10}, {y})" for x in range(100)]
	 for y in range(100)]

for s in sheet.sheets:
	s.table.clean()

with open("test.html", "w") as f:
	f.write(sheet.serialize())
import os
os.startfile("test.html")
#import time; time.sleep(5)
#os.remove("test.html")