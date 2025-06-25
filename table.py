from typing import Any, List, Tuple, Iterator, Union
import html

import http.server, socketserver, threading, webbrowser, json, os
from urllib.parse import urlparse
import asyncio, websockets
import atexit

class Server:
	def __init__(self, spreadsheet, port=80):
		self.sheet = spreadsheet
		self.port = port
		self.clients = set()
		self.scroll_pos = (0, 0)
		self.inc_file = None
		self.http_thread = None
		self.ws_thread = None
		self.loop = asyncio.new_event_loop()
		self.should_stop = False

		atexit.register(self.stop)

	def update_shortcut(self, inc_filename):
		self.inc_file = inc_filename

	def open_in_browser(self):
		import socket
		ip = socket.gethostbyname(socket.gethostname())
		webbrowser.open(f"http://{ip}:{self.port}")

	def start(self):
		self._start_http_server()
		self._start_websocket_server()

	def _websocket_script(self):
		return f"""
	<script>
	let ws = new WebSocket(`ws://${{location.hostname}}:{self.port + 1}`);
	ws.onmessage = msg => {{
		let data = JSON.parse(msg.data);
		if (data.type === "update") {{
			data.cells.forEach(cell => {{
				let el = document.querySelector(`td[x="${{cell.x}}"][y="${{cell.y}}"]`);
				if (el) {{
					el.textContent = cell.value;
					el.style.background = cell.style.bg;
					el.style.color = cell.style.color;
				}}
			}});
		}} else if (data.type === "reload") {{
			location.reload();
		}} else if (data.type === "scroll") {{
			document.querySelectorAll(".TBC").forEach(e => {{
			e.scrollLeft = data.x;
			e.scrollTop = data.y;
			}});
		}}
	}};
	</script>
	"""

	def _start_http_server(self):
		class Handler(http.server.BaseHTTPRequestHandler):
			def do_GET(self):
				if self.path == '/' or self.path.startswith('/index.html'):
					# Serialize fresh on each request
					html = self.server_instance.sheet.serialize()
					if self.server_instance.inc_file and os.path.exists(self.server_instance.inc_file):
						with open(self.server_instance.inc_file, 'r', encoding='utf8') as f:
							html += f.read()
					html += self.server_instance._websocket_script()
	
					self.send_response(200)
					self.send_header("Content-type", "text/html")
					self.end_headers()
					self.wfile.write(html.encode("utf8"))
				else:
					self.send_error(404, "File not found")
	
			def log_message(self, format, *args):
				return  # silence default logging
	
		# We need a way for Handler to access 'self', so assign a reference
		Handler.server_instance = self
	
		self.http_thread = threading.Thread(
			target=lambda: socketserver.TCPServer(("0.0.0.0", self.port), Handler).serve_forever(),
			daemon=True
		)
		self.http_thread.start()


	def _start_websocket_server(self):
		async def ws_handler(websocket):
			self.clients.add(websocket)
			try:
				await websocket.send(json.dumps({
					"type": "full",
					"html": self.sheet.serialize(),
					"scroll": self.scroll_pos
				}))
				while True:
					msg = await websocket.recv()
					# handle messages here if needed
			except:
				pass
			finally:
				self.clients.remove(websocket)
	
		async def run_ws():
			async with websockets.serve(ws_handler, "0.0.0.0", self.port + 1):
				await asyncio.Future()  # run forever
	
		self.ws_thread = threading.Thread(target=self.loop.run_until_complete, args=(run_ws(),), daemon=True)
		self.ws_thread.start()


	def update(self):
		updates = []
		for x, y, cell in self.sheet.sheets[0].table[:][:].superRange:
			if getattr(cell, "dirty", False):
				updates.append({
					"x": x, "y": y,
					"value": cell.value,
					"style": {
						"bg": cell.style.background,
						"color": cell.style.color
					}
				})
				cell.dirty = False

		if updates:
			message = json.dumps({"type": "update", "cells": updates})
			asyncio.run_coroutine_threadsafe(self._broadcast(message), self.loop)

	def setClientScroll(self, x, y):
		self.scroll_pos = (x, y)
		message = json.dumps({"type": "scroll", "x": x, "y": y})
		asyncio.run_coroutine_threadsafe(self._broadcast(message), self.loop)

	def reload(self):
		asyncio.run_coroutine_threadsafe(self._broadcast(json.dumps({"type": "reload"})), self.loop)

	def stop(self):
		self.should_stop = True
		for task in asyncio.all_tasks(loop=self.loop):
			task.cancel()

	async def _broadcast(self, msg):
		dead = []
		for client in self.clients:
			try:
				await client.send(msg)
			except:
				dead.append(client)
		for c in dead:
			self.clients.remove(c)


class SpreadSheet:
	class Style:
		class Font:
			pass
		class Border:
			pass
	class Cell:
		pass
	class Formula:
		pass
	class Table:
		pass
	pass

class SpreadSheet:
	class Style:
		class Font:
			def __init__(self, size:float, family:str, modifiers:str):
				self.size = size
				self.family = family
				self.modifiers = modifiers
			def __repr__(self):
				return f"{self.size} {self.family} {self.modifiers}"
			def clone(self):
				return SpreadSheet.Style.Font(self.size, self.family, self.modifiers)
			def __str__(self):
				return f"{self.size} {self.family} {self.modifiers}"
		class Border:
			def __init__(self, dict):
				self.left = "1px solid #aaa"
				self.right = "1px solid #aaa"
				self.top = "1px solid #aaa"
				self.bottom = "1px solid #aaa"
				for k,v in dict.items():
					self.__setattr__(k, v)
			def __repr__(self):
				return f"Border({self.left},{self.right},{self.top},{self.bottom})"
			def clone(self):
				return SpreadSheet.Style.Border({
					"left": self.left,
					"right": self.right,
					"top": self.top,
					"bottom": self.bottom
				})
		def __init__(self, border=None, font=None, fill=None):
			self.border = border if isinstance(border, SpreadSheet.Style.Border) else SpreadSheet.Style.Border({
				"left":"1px solid #aaa",
				"right":"1px solid #aaa",
				"top":"1px solid #aaa",
				"bottom":"1px solid #aaa",
				**(border if isinstance(border, dict) else {})
			})
			self.font = font if isinstance(font, SpreadSheet.Style.Font)  else SpreadSheet.Style.Font(14.0, "calibri", "monospace")
			self.background = "#ffffff"
			self.color = "#000"
		def clone(self):
			cloned_border = self.border.clone()
			cloned_font = self.font.clone()
			new_style = SpreadSheet.Style(cloned_border, cloned_font, self.background)
			new_style.color = self.color
			return new_style

	class Cell:
		def __init__(self, value: Any = None, style: SpreadSheet.Style = None):
			self._value = value
			self._style = style if style is not None else SpreadSheet.Style()
			self._dirty = True
			self.formula = None
		@property
		def dirty(self):
			if self.formula is None:
				return self._dirty
			return True
		@dirty.setter
		def dirty(self, val):
			if self.formula is not None:
				self._dirty = True
			self._dirty = val
		@property
		def value(self):
			if self.formula is not None:
				return self.formula()
			return self._value
		@value.setter
		def value(self, val):
			if self.formula is not None:
				self.formula = None
			self._value = val
		def __repr__(self):
			return f"C({self.value}, S({self.style.border}, {self.style.color}, {self.style.background}, {self.style.font}))"
		def clone(self):
			# For value, shallow copy is likely fine; deep copy if needed
			newCell = SpreadSheet.Cell(self.value, self.style.clone())
			if self.formula:
				newCell.formula = self.formula
			return 
	class Formula:
		add=0
		sub=1
		mult=2
		div=3
		def __init__(self, operation, a, b):
			self.op = operation
			self.a = a
			self.b = b
		def __call__(self):
			a = self.a
			b = self.b
			if isinstance(self.a, SpreadSheet.Formula):
				a = a()
			elif isinstance(self.a, SpreadSheet.Cell):
				a = a.value

			if isinstance(self.b, SpreadSheet.Formula):
				b = b()
			elif isinstance(self.b, SpreadSheet.Cell):
				b = b.value

			match(self.op):
				case SpreadSheet.Formula.add:
					return a + b
				case SpreadSheet.Formula.sub:
					return a - b
				case SpreadSheet.Formula.mult:
					return a * b
				case SpreadSheet.Formula.div:
					return a / b
	class Table:
		def __init__(self, width: int, height: int):
			self.data = [[SpreadSheet.Cell() for _ in range(height)] for _ in range(width)]
			self.width = width
			self.height = height
			self.server = None
	
		def __getitem__(self, x):
			if isinstance(x, int):
				x = slice(x, x + 1)
			elif not isinstance(x, slice):
				raise TypeError(f"Unsupported index type: {type(x)}")
			return SpreadSheet.TableColumnProxy(self, x)

		def __repr__(self):
			data = []
			for y in range(self.height):
				row = []
				for x in range(self.width):
					row.append(repr(self.data[x][y]))
				data.append("[" + ", ".join(row) + "]")
			return "[\n\t" + ",\n\t".join(data) + "\n]"
		def serialize(self):
			class Node:
				def __init__(self, type, contents=None):
					self.type = type
					self.contents = contents if contents is not None else []
				def append(self, *items):
					self.contents.extend(items)
				def __str__(self):
					if isinstance(self.contents, list):
						inner = " ".join(map(str, self.contents))
					else:
						inner = str(self.contents)
					return f"({self.type} {inner})"
			tree = Node("table")
			row = Node("tr")
			for x in range(self.width+1):
				row.append(Node("td", Node("string", f"'{number_to_excel_col(x)}'")))
			tree.append(row)
			for y in range(self.height):
				row = Node("tr")
				row.append(
					Node("--row", str(y)),
					Node("td", Node("number", str(y)))
				)
				for x in range(self.width):
					cell = self.data[x][y]
					item = Node("td")
					style = Node("style")
					style.append(
						Node("border-left",   "'" + cell.style.border.left	 .replace("'", "\\'") + "'"),
						Node("border-right",  "'" + cell.style.border.right	 .replace("'", "\\'") + "'"),
						Node("border-top",    "'" + cell.style.border.top	 .replace("'", "\\'") + "'"),
						Node("border-bottom", "'" + cell.style.border.bottom .replace("'", "\\'") + "'"),
						Node("background",    "'" + cell.style.background	 .replace("'", "\\'") + "'"),
						Node("color",         "'" + cell.style.color		 .replace("'", "\\'") + "'"),
						Node("font-family",   "'" + repr(cell.style.font)	 .replace("'", "\\'") + "'"),
					)
					item.append(style, Node("--col",str(x)))
					if isinstance(cell.value, str):
						escaped = cell.value.replace("'", "\\'")
						value = Node("string", f"'{escaped}'")
					elif isinstance(cell.value, (int, float)):
						value = Node("number", str(cell.value))
					elif cell.value is None:
						value = Node("None")
					else:
						escaped = str(cell.value).replace("'", "\\'")
						value = Node("string", f"'{escaped}'")
		
					item.append(value)
					row.append(item)
				tree.append(row)
		
			return str(tree)


		def _expand_to_include(self, x: int, y: int):
			# Expand columns if needed
			if x >= self.width:
				for _ in range(self.width, x + 1):
					self.data.append([SpreadSheet.Cell() for _ in range(self.height)])
				self.width = x + 1
			
			# Expand rows in each column if needed
			if y >= self.height:
				for col in self.data:
					 col.extend(SpreadSheet.Cell() for _ in range(self.height, y + 1))
				self.height = y + 1
			if self.server is not None:
				self.server.reload()

	
	
		def clean(self):
			min_x, max_x = self.width, -1
			min_y, max_y = self.height, -1
	
			# Find bounding box of all non-empty cells
			for x in range(self.width):
				for y in range(self.height):
					if self.data[x][y].value is not None:
						if x < min_x: min_x = x
						if x > max_x: max_x = x
						if y < min_y: min_y = y
						if y > max_y: max_y = y
	
			# If no non-empty cells found, clear all
			if max_x == -1 or max_y == -1:
				# Clear everything
				self.data = []
				self.width = 0
				self.height = 0
				return
	
			new_width = max_x - min_x + 1
			new_height = max_y - min_y + 1
	
			# Slice data to new bounding rectangle
			new_data = []
			for x in range(min_x, max_x + 1):
				col = self.data[x][min_y:max_y + 1]
				new_data.append(col)
	
			self.data = new_data
			self.width = new_width
			self.height = new_height
			if self.server is not None:
				self.server.reload()
	
		def __iter__(self) -> Iterator[Tuple[int, int, SpreadSheet.Cell]]:
			return self[:][:].superRange

		def clone(self):
			new_table = SpreadSheet.Table(self.width, self.height)
			for x in range(self.width):
				for y in range(self.height):
					new_table.data[x][y] = self.data[x][y].clone()
			return new_table

		def __repr__(self):
			return f"Table({self.width}, {self.height})"
	
	class TableColumnProxy:
		def __init__(self, table: SpreadSheet.Table, x_slice: Union[int, slice]):
			self.table = table
			if isinstance(x_slice, int):
				self.x_slice = slice(x_slice, x_slice + 1)
			else:
				self.x_slice = x_slice
	
		def __getitem__(self, y_slice: Union[int, slice]):
			return SpreadSheet.TableRange(self.table, self.x_slice, y_slice)


		def __setitem__(self, y_index: Union[int, slice], cell_value: SpreadSheet.Cell):
			if isinstance(y_index, int):
				x = self.x_slice.start
				y = y_index
				self.table._expand_to_include(x, y)
				if not isinstance(cell_value, SpreadSheet.Cell):
					raise ValueError("Assigned value must be a SpreadSheet.Cell")
				# 🛠️ No clone here — just assign reference
				self.table.data[x][y] = cell_value
			else:
				raise NotImplementedError("Only integer index assignment is supported")


		def __repr__(self):
			x = str(self.x_slice.start), str(self.x_slice.stop), str(self.x_slice.step)
			return f"TableColumn({":".join(x)})"
	
	class TableRange:
		def __init__(self, table: SpreadSheet.Table, x_slice: Union[int, slice], y_slice: Union[int, slice]):
			self.table = table
	
			# Normalize int to slice
			if isinstance(x_slice, int):
				x_slice = slice(x_slice, x_slice + 1)
			if isinstance(y_slice, int):
				y_slice = slice(y_slice, y_slice + 1)
			
			self.x_slice = x_slice.start or 0, x_slice.stop or table.width,  x_slice.step or 1
			self.y_slice = y_slice.start or 0, y_slice.stop or table.height, y_slice.step or 1

			self.table._expand_to_include(self.x_slice[1], self.y_slice[1])
			
		def __iter__(self):
			for y in range(*self.y_slice):
				row = []
				for x in range(*self.x_slice):
					row.append(self.table.data[x][y])
				yield row
	
		def __getattr__(self, attr):
			# Check if attribute exists on any cell and is complex (has sub-attributes)
			has_subattrs = False
			for x in range(*self.x_slice):
				for y in range(*self.y_slice):
					cell = self.table.data[x][y]
					if hasattr(cell, attr):
						val = getattr(cell, attr)
						# Check if val itself has attributes (is a complex object, not basic type)
						if hasattr(val, '__dict__') or hasattr(val, '__slots__'):
							has_subattrs = True
							break
				if has_subattrs:
					break
		
			if has_subattrs:
				return SpreadSheet.RecursiveAccessor(self, [attr])
			else:
				# Return attribute values as 2D list
				return [[getattr(self.table.data[x][y], attr)
						 for y in range(*self.y_slice)]
						 for x in range(*self.x_slice)]
	
		def __setattr__(self, prop, new_values: Any):
			if prop in ('table', 'x_slice', 'y_slice'):
				object.__setattr__(self, prop, new_values)
				return
			
			# Check if attribute is complex on any cell
			has_subattrs = False
			self.table._expand_to_include(self.x_slice[1], self.y_slice[1])
			for x in range(*self.x_slice):
				for y in range(*self.y_slice):
					cell = self.table.data[x][y]
					if hasattr(cell, prop):
						val = getattr(cell, prop)
						if hasattr(val, '__dict__') or hasattr(val, '__slots__'):
							has_subattrs = True
							break
				if has_subattrs:
					break
			
			if has_subattrs:
				# Delegate setting to RecursiveAccessor to handle nested attributes properly
				rec = SpreadSheet.RecursiveAccessor(self, [prop])
				rec.__setattr__(prop, new_values)
			else:
				# Simple direct set of attribute on each cell
				if isinstance(new_values, list):
					for dx, x in enumerate(range(*self.x_slice)):
						for dy, y in enumerate(range(*self.y_slice)):
							setattr(self.table.data[x][y], prop, new_values[dx][dy])
							self.table.data[x][y].dirty = True
				else:
					for x in range(*self.x_slice):
						for y in range(*self.y_slice):
							setattr(self.table.data[x][y], prop, new_values)
							self.table.data[x][y].dirty = True

	
		@property
		def superRange(self) -> Iterator[Tuple[int, int, SpreadSheet.Cell]]:
			for x in range(*self.x_slice):
				for y in range(*self.y_slice):
					yield (x, y, self.table.data[x][y])
	
		@property
		def border(self):
			return SpreadSheet.TableBorderAccessor(self, include_edges=False)
	
		@property
		def borderRange(self):
			return SpreadSheet.TableBorderAccessor(self, include_edges=True)

		def __repr__(self):
			x = str(self.x_slice[0]), str(self.x_slice[1]), str(self.x_slice[2])
			y = str(self.y_slice[0]), str(self.y_slice[1]), str(self.y_slice[2])
			return f"TableRange({':'.join(x)}, {':'.join(y)})"
	
	class RecursiveAccessor:
		def __init__(self, table_range, attr_path=None):
			object.__setattr__(self, 'table_range', table_range)
			object.__setattr__(self, 'attr_path', attr_path or [])
	
		def __getattr__(self, name):
			# Return new RecursiveAccessor with extended attribute path
			return RecursiveAccessor(self.table_range, self.attr_path + [name])
	
		def __setattr__(self, name, value):
			# Internal attributes set normally
			if name in ('table_range', 'attr_path'):
				object.__setattr__(self, name, value)
				return
			full_path = self.attr_path + [name]
	
			for x in range(*self.table_range.x_slice):
				for y in range(*self.table_range.y_slice):
					cell = self.table_range.table.data[x][y]
					target = cell
					# Traverse all but last attribute in the path
					for attr in full_path[:-1]:
						target = getattr(target, attr)
					setattr(target, full_path[-1], value)
					cell.dirty = True
	
		def __getattribute__(self, name):
			if name in ('table_range', 'attr_path', '__class__', '__dict__', '__weakref__', '__setattr__', '__getattr__', '__getattribute__'):
				return object.__getattribute__(self, name)
	
			full_path = object.__getattribute__(self, 'attr_path') + [name]
			table_range = object.__getattribute__(self, 'table_range')
	
			# Gather values from all cells for the full attribute path
			values = []
			for x in range(*table_range.x_slice):
				col = []
				for y in range(*table_range.y_slice):
					cell = table_range.table.data[x][y]
					target = cell
					for attr in full_path:
						target = getattr(target, attr)
					col.append(target)
				values.append(col)
			return values
	class Sheet:
		def __init__(self, name, table: SpreadSheet.Table = None, server = None):
			self.name = name
			self.server = server
			self.table = table if isinstance(table, SpreadSheet.Table) else SpreadSheet.Table(0, 0)
			self.table.server = self.server
	def __init__(self):
		self.sheets = []
		self.server = None
	def createSheet(self, name:str, table : SpreadSheet.Table = None):
		self.sheets.append(SpreadSheet.Sheet(name, table, self.server))
		if self.server is not None:
			self.server.reload()
	def serialize(self):
		def freeze_style(style):
			return (
				style.border.left,
				style.border.right,
				style.border.top,
				style.border.bottom,
				style.background,
				style.color,
				style.font.size,
				style.font.family,
				style.font.modifiers,
			)

		def style_to_css(style_key):
			bl, br, bt, bb, bg, color, fsize, ffam, fmod = style_key
			return (
				f"background:{bg};"
				f"color:{color};"
				f"border-left:{bl};"
				f"border-right:{br};"
				f"border-top:{bt};"
				f"border-bottom:{bb};"
				f"font-family:{ffam};"
				f"font-size:{fsize}px;"
				f"font-style:{fmod};"
			)
		def number_to_excel_col(n):
			result = ""
			while n > 0:
				n -= 1  # Adjust because Excel columns are 1-based but modulo is 0-based
				result = chr((n % 26) + ord('A')) + result
				n //= 26
			return result
		global_styles = {}   # style_key -> 'S{num}'
		next_global_id = 1

		all_tables_html = []
		all_tables_local_styles = []  # [(table_index, {style_key: SSclass})]
		all_tables_local_classes_map = []  # to hold cell-to-class mappings per table

		# First pass: assign global styles to known global pool, assign local styles otherwise
		for table_index, sheet in enumerate(self.sheets, 1):
			table = sheet.table

			local_styles = {}  # style_key -> 'SS{num}'
			next_local_id = 1

			# We'll store the classes assigned to each cell for possible replacement later
			cell_classes = [[None]*table.height for _ in range(table.width)]

			for y in range(table.height):
				for x in range(table.width):
					cell = table.data[x][y]
					skey = freeze_style(cell.style)

					if skey in global_styles:
						cls = global_styles[skey]
					else:
						if skey not in local_styles:
							cls = f"SS{next_local_id}"
							local_styles[skey] = cls
							next_local_id += 1
						else:
							cls = local_styles[skey]
					cell_classes[x][y] = cls

			all_tables_local_styles.append((table_index, local_styles))
			all_tables_local_classes_map.append(cell_classes)

		# Second pass: check each table's local styles against global styles
		# If local style matches a global style, replace all occurrences of that SS class with the S class
		for idx, (table_index, local_styles) in enumerate(all_tables_local_styles):
			# Reverse map from local class -> style_key
			local_class_to_style = {v: k for k, v in local_styles.items()}

			cell_classes = all_tables_local_classes_map[idx]

			# For each local style in the table
			for skey, ss_class in list(local_styles.items()):
				if skey in global_styles:
					# Replace local ss_class with global s_class in all cells for this table
					s_class = global_styles[skey]
					# Replace in cell_classes
					for x in range(len(cell_classes)):
						for y in range(len(cell_classes[0])):
							if cell_classes[x][y] == ss_class:
								cell_classes[x][y] = s_class
					# Remove this local style, as it is now global
					del local_styles[skey]

				else:
					# This is a new unique style, so add to global_styles now to avoid duplicates in other tables later
					global_styles[skey] = f"S{next_global_id}"
					# Replace ss_class with the new global S class in all cells
					new_global_class = global_styles[skey]
					for x in range(len(cell_classes)):
						for y in range(len(cell_classes[0])):
							if cell_classes[x][y] == ss_class:
								cell_classes[x][y] = new_global_class
					del local_styles[skey]
					next_global_id += 1

			# Update the map after replacements
			all_tables_local_classes_map[idx] = cell_classes

		# Now generate final CSS and HTML with updated classes

		# Compose global CSS
		global_css = (
			"body {margin: 0; display: flex; flex-direction: row; height: 100vh; }"
			".TBCC {min-width: 0; display: flex; flex-direction: row; margin: 10px; }"
			".TBC {overflow:auto;margin:10px;scrollbar-width:none;-ms-overflow-style:none;}"
			".TBC::-webkit-scrollbar{display:none;}"
			"table {border-collapse:collapse;position:relative;overflow:clip;}"
			"thead th{position:sticky;top:0;background:#eee;z-index:5;border-right:1px solid #aaa;padding:4px 8px;}"
			"thead th::after{content:\"\";position:absolute;left:0;bottom:0;height:3px;width:103%;background:#aaa;z-index:-1;}"
			"thead th:first-child{left:0;z-index:10;background:#eee;position:sticky;top:0;left:0;}"
			"thead th:first-child::before{content:\"\";position:absolute;top:0;right:0;width:3px;height:100%;background:#aaa;z-index:1;}"
			"tbody th{position:sticky;left:0;z-index:4;background:#eee;min-width:40px;text-align:center;border-bottom:1px solid #aaa;padding:4px 8px;}"
			"tbody th::before{content:\"\";position:absolute;top:0;right:0;width:3px;height:100%;background:#aaa;z-index:1;}"
			"tbody td{white-space:nowrap;border-bottom:1px solid #ccc;padding:4px 8px;}"
		)
		for skey, cls in sorted(global_styles.items(), key=lambda i: int(i[1][1:])):
			global_css += f".{cls} {{{style_to_css(skey)}}}\n"

		# Compose per-table local CSS (now should be empty or minimal)
		local_css_blocks = []
		for table_index, local_styles in all_tables_local_styles:
			css_block = ""
			for skey, cls in local_styles.items():
				css_block += f".TBC.{table_index} .{cls} {{{style_to_css(skey)}}}\n"
			local_css_blocks.append(css_block)

		# Compose tables HTML
		all_tables_html = []
		for idx, sheet in enumerate(self.sheets):
			table = sheet.table
			table_index = idx + 1
			cell_classes = all_tables_local_classes_map[idx]
		
			rows_html = []
		
			# Header row (empty top-left + column letters)
			header_row = ['<th></th>'] + [
				f'<th>{number_to_excel_col(x+1)}</th>' for x in range(table.width)
			]
			rows_html.append("<thead>\n\t<tr>" + "".join(header_row) + "</tr>\n\t</thead>\n\t<tbody>")
		
			# Data rows with row numbers
			for y in range(table.height):
				row_cells = [f'<th>{y + 1}</th>']
				for x in range(table.width):
					cell = table.data[x][y]
					cls = cell_classes[x][y]
					val = cell.value if cell.value is not None else ""
					row_cells.append(f'<td class="{cls}" x="{x}" y="{y}">{html.escape(str(val))}</td>')
				rows_html.append("<tr>" + "".join(row_cells) + "</tr>")
			all_tables_html.append(f'<div class="TBCC"><div class="TBC {table_index}"><table>\n' + "\n".join(rows_html) + "\n\t</tbody>\n</table></div></div>")

		style_tag = f"<style>\n{global_css}</style>\n"
		for local_css in local_css_blocks:
			if local_css.strip():
				style_tag += f"<style>\n{local_css}</style>\n"

		return style_tag + "\n".join(all_tables_html) + '<script>document.addEventListener("DOMContentLoaded",()=>{requestIdleCallback(()=>{let e=document.querySelectorAll(".TBC"),l=!1;e.forEach(r=>{r.addEventListener("scroll",()=>{if(l)return;l=!0;let o=r.scrollLeft,t=r.scrollTop;e.forEach(e=>{e!==r&&(e.scrollLeft=o,e.scrollTop=t)}),requestAnimationFrame(()=>{l=!1})})})})});</script>'
'''
(()=>{
	const elements = document.querySelectorAll('.THC');
	let isSyncing = false;
	
	elements.forEach(el => {
		el.addEventListener('scroll', () => {
			if (isSyncing) return;
	
			isSyncing = true;
			const sl = el.scrollLeft;
			const st = el.scrollTop;
	
			elements.forEach(other => {
				if (other !== el) {
					other.scrollLeft = sl;
					other.scrollTop = st;
				}
			});
		requestAnimationFrame(() => { isSyncing = false; });
		});
	});
})()
'''