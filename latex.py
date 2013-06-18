command = "\\"
linebreak = "\n"
parameter_open = "{"
parameter_close = "}"
keyword_begin = "begin"
keyword_end = "end"
keyword_itemize = "itemize"
keyword_frame = "frame"
keyword_frametitle = "frametitle"
keyword_item = "item"

def indent(depth, spacer = "\t"):
	return depth*spacer

def begin():
	return command+keyword_begin

def item(text = ""):
	return command+keyword_item+" "+text + linebreak

def end():
	return command+keyword_end

def begin_frame():
	return begin() + parameter_open + keyword_frame + parameter_close + linebreak

def end_frame():
	return  end() + parameter_open + keyword_frame + parameter_close + linebreak

def begin_itemize():
	return  begin() + parameter_open + keyword_itemize+ parameter_close + linebreak

def end_itemize():
	return  end() + parameter_open + keyword_itemize + parameter_close + linebreak

def frametitle(frametitle):
	return command + keyword_frametitle + parameter_open + frametitle + parameter_close + linebreak

class Item:
	pass

class Itemize:

	def __init__(self, item):
		self.items = []
		add_item(self, item)	
	
	def __init__(self):
		self.items = []

	def add_item(self, item):
		self.items.append(item)

	def add_items(self, items):
		self.items.append(items)

	
	def as_string(self, depth = 0):
		s = indent(depth)+ begin_itemize() #"\\begin{itemize}\n"
		for i in self.items:
			#print(type(i), isinstance(i, Itemize))
			if isinstance(i, Itemize):
				depth += 1
				s += i.as_string(depth)
				depth -= 1
			if type(i) == str:
				s+=indent(depth+1) + item(i) #"\\item "+i+"\n"
		s+=indent(depth) + end_itemize() #"\\end{itemize}\n"
		return s;



class BeamerFrame:
	#def __init__(self, title, items):
		#self.title = title
		#self.items = items

	def __init__(self, title):
		self.frame_title = title
		self.items = Itemize()
	
	#def add_item(self, items):
		#items.append(item)
	
	def get_items(self):
		return self.items 

	def as_string(self):
		s = begin_frame() # "\\begin{frame}\n"
		s += indent(1)+frametitle(self.frame_title) #"\\frametitle{%s}\n"%self.frame_title
		s += self.get_items().as_string(1)
		s += end_frame() # "\\end{frame}"
		return s


class LatexBeamer:
	def __init__(self, title, frames):
		self.title = title
		self.frames = frames

	def __init__(self, title):
		self.title = title
		self.frames = []
	
	def as_string(self):
		s = ""
		for f in self.frames:
			s += f.as_string()
		return s

	def write_to_file(self, filename):
		s = self.as_string()
		tex_file = open(filename, "w")
		tex_file.write(s)
		tex_file.close()
	
	def add_frame(self, frame):
		self.frames.append(frame)

	def get_frames(self):
		return self.frames


#testing
latex = LatexBeamer("Title");

frame = BeamerFrame("FrameTitle")
frame.get_items().add_item("Hallo")
frame.get_items().add_item("Welt")

#i = frame.get_items()
#frame.get_items().add_items(i)

items = Itemize()
items.add_item("asd")
items.add_item("asdf")

frame.get_items().add_items(items)


latex.add_frame(frame)
latex.write_to_file("test.tex")





