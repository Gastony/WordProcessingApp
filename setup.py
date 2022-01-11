from distutils.core import setup # Need this to handle modules
import math # We have to import all modules used in our progra
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot
from mainwindow import Ui_MainWindow
import sys
from model import Model
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import *
import docx
from docx import *
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from number_converter import numerals
import datetime
import math
import os
import re
from num2words import num2words
from mainwindow import *
from model import *
from number_converter import *
from page_numbers import *
from PyQt5 import sip
import sip
import PyQt5.sip
from dateutil.parser import parse
from dateparser.search import search_dates
from docx.shared import RGBColor
from docx.enum.style import WD_STYLE_TYPE


from cx_Freeze import setup, Executable


setup(name = "reandurllib" ,
      version = "0.1" ,
      description = "" ,
      executables = [Executable("App.py")])