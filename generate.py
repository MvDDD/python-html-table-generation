from table import SpreadSheet, Server
from time import sleep

sheet = SpreadSheet()
server = Server(sheet, port=8080)
sheet.createSheet("sheet1")
table = sheet.sheets[0].table
server.start()
server.update_shortcut("shortcut.htm")
server.open_in_browser()

for x,y in [(x,y) for y in range(10) for x in range(10)]:
	print(f"({x}, {y})")
	table[x][y].value = f"({x}, {y})"

def forever(start=0):
	while 1:
		yield start
		start += 1

for i in forever(1):
	sleep(1)
	print(i)

	for x,y in [(x,y) for y in range(10) for x in range(10)]:
		table[x][y].value = i
		table[x][y].dirty = True

	server.update()

server.stop()