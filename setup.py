from distutils.core import setup
import py2exe

setup(
    windows=['script.py'],
    zipfile=None,
    data_files=[(".", ["kvu.py"])],
    options={
        'py2exe': {
            'includes': [
                'Tkinter',
                'tkFileDialog'
            ],
            'compressed': 1
        }
    }
)