from table import SpreadSheet, Server
from time import sleep

sheet = SpreadSheet()
server = Server(sheet, port=8000)
sheet.createSheet("sheet1")
table = sheet.sheets[0].table
table[10][10].value = 10
server.start()
server.open_in_browser()
sleep(1)
table[1][1].value = 100
server.update()
try:
# Keep the main thread alive while the server runs
	import time
	while True:
		time.sleep(1)
except KeyboardInterrupt:
	print("Stopping server...")
	server.stop()