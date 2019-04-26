# coding=utf-8
from Tkinter import *
from tkFileDialog import askopenfilename
import subprocess
import os
import time


def click():
    tex1 = textentry1.get().encode('utf-8')
    tex2 = textentry2.get().encode('utf-8')
    if tex1 and tex2:
        homepath = os.path.expanduser(os.getenv('USERPROFILE'))
        out.configure(state='normal')
        out.delete(0.0, END)
        try:
            process = subprocess.check_output('python kvu.py -o "%s" -n "%s"' % (tex1, tex2),
                                              shell=True,
                                              stderr=subprocess.STDOUT)
        except subprocess.CalledProcessError as e:
            rcode = e.returncode
        else:
            rcode = 0
        with open(os.path.join(homepath, 'kvu.result'), 'r+') as result:
            output = result.readline()
        out.insert(END, "%s" % (e.output if rcode == 1 else output))
        out.configure(state='disabled')


def browsefunc1():
    filename = askopenfilename()
    textentry1.configure(state='normal')
    textentry1.delete(0, END)
    textentry1.insert(END, filename)
    textentry1.configure(state='disabled')


def browsefunc2():
    filename = askopenfilename()
    textentry2.configure(state='normal')
    textentry2.delete(0, END)
    textentry2.insert(END, filename)
    textentry2.configure(state='disabled')


module = ['jdcal-1.4', 'et_xmlfile-1.0.1', 'openpyxl-2.6.2']

try:
    import jdcal, et_xmlfile, openpyxl
except ImportError:
    process = subprocess.Popen('cd C:\\ & dir /P Python27 /AD', shell=True, stdout=subprocess.PIPE)
    for i in process.communicate()[0].strip().split():
        if 'C:\\' in i:
            python_folder = i

    subprocess.call('xcopy %s %s /E /H /I & xcopy %s %s /E /H /I & xcopy %s %s /E /H /I' %
                    (module[0], os.path.join(python_folder, 'Lib\\', module[0]),
                     module[1], os.path.join(python_folder, 'Lib\\', module[1]),
                     module[2], os.path.join(python_folder, 'Lib\\', module[2])), shell=True)

    subprocess.call('cd %s & python setup.py install' % (os.path.join(python_folder, 'Lib\\%s' % module[0])),
                    shell=True)
    subprocess.call('cd %s & python setup.py install' % (os.path.join(python_folder, 'Lib\\%s' % module[1])),
                    shell=True)
    subprocess.call('cd %s & python setup.py install' % (os.path.join(python_folder, 'Lib\\%s' % module[2])),
                    shell=True)


root = Tk()
root.title("Обновление реестра")

Label(root, text='Cтарая форма реестра: ') .grid(row=1, column=0, sticky=W)

textentry1 = Entry(root, width=40, bg="white", state='disabled')
textentry1.grid(row=1, column=1, sticky=W)

browsebutton1 = Button(root, text="Обзор...", command=browsefunc1)
browsebutton1.grid(row=1, column=2, sticky=W)

Label(root, text='Новая форма реестра: ') .grid(row=2, column=0, sticky=W)

textentry2 = Entry(root, width=40, bg="white", state='disabled')
textentry2.grid(row=2, column=1, sticky=W)

browsebutton2 = Button(root, text="Обзор...", command=browsefunc2)
browsebutton2.grid(row=2, column=2, sticky=W)

btn = Button(root, text='Старт', width=6, command=click)
btn.grid(row=3, column=1, sticky=N)

out = Text(root, width=27, height=5, wrap=WORD, font="11", state="disabled")
out.grid(row=4, column=1, sticky=W)

root.mainloop()