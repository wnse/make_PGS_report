from tkinter import *
from tkinter import scrolledtext
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
import logging
from excel2docx import excel2docx
import os


class TextHandler(logging.Handler):
    """This class allows you to log to a Tkinter Text or ScrolledText widget"""
    def __init__(self, text):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text.configure(state='normal')
            self.text.insert(END, msg + '\n')
            self.text.configure(state='disabled')
            # Autoscroll to the bottom
            self.text.yview(END)
        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)

def start():
	try:
		excel2docx(config_label['text'], input_label['text'], template_label['text'], outdir_label['text'], picdir_label['text'])
	except Exception as e:
		logger.error(e)


def get_dir(path, type='file'):
	if type=='file':
		file_path = askopenfilename()
		path.configure(text=file_path)   
	else:
		dir_path = askdirectory()
		path.configure(text=dir_path)   

def pack_button_label(default_text, row, type='file'):
	temp_label = Label(root, text=default_text)
	temp_label.grid(row=row, column=1)
	Button(root, text=default_text, command=lambda: get_dir(temp_label, type=type)).grid(row=row, column=0)
	return temp_label

# logger = logging.getLogger()
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
bin_dir = os.path.split(os.path.realpath(__file__))[0]

root = Tk()
root.title('excel2docx')

#### 图标
from icon import Icon
import base64
with open('tmp.ico','wb') as tmp:
    tmp.write(base64.b64decode(Icon().img))
root.iconbitmap('tmp.ico')
os.remove('tmp.ico')

help_info = (f'这个模块的作用是:将excel文件的信息填入到docx文件中各表格中,形成docx的报告')
help_label = Label(root, text=f'{help_info:>1}', justify='left')
help_label.grid(row=0, columnspan=2)

template_label = pack_button_label('选择模板文件', row=1, type='file')
config_label = pack_button_label('选择配置文件', row=2, type='file')
input_label = pack_button_label('选择输入文件', row=3, type='file')
outdir_label = pack_button_label('选择输出文件夹', row=4, type='directory')
picdir_label = pack_button_label('选择图片文件夹', row=5, type='directory')

Button(root, text='开始', command=start).grid(row=7, column=0)
Button(root, text='退出', command=root.quit).grid(row=7, column=1)
st = scrolledtext.ScrolledText(root, state='disabled')
st.configure(font='TkFixedFont')
st.grid(columnspan=2)
text_handler = TextHandler(st)
logger = logging.getLogger()
logger.addHandler(text_handler)
# logger.debug('debug message')
logger.info('log message:')
# logger.error('error message')



root.mainloop()

