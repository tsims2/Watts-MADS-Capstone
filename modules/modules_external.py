color_palette = [('#FF5400'),  # 0 orange
                    ('#FFC907'),  # 1 yellow
                    ('#072B60'),  # 2 royal blue
                    ('#00C9AZ'),  # 3 green
                    ('#59C8E3'), # 4 sky blue
                    ('#020244'), # 5 navy
                    ('#EBF2F5'), # 6 light blue
                    ('#363C49'), # 7 gray
                    ('#0078A6'), # 8 Teal
                    ('#12B34C'), # 9 Green
                    ('#316130'), # 10 Dark Green
                    ('#613031'), # 11 Dark Red
                    ]  

import sys
import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import re
import time
from datetime import datetime, timedelta, date, time
from dateutil.relativedelta import relativedelta
import calendar
import math
from prettytable import PrettyTable
import glob
import traceback

pd.set_option('future.no_silent_downcasting', True) 

import folium
from folium.plugins import MarkerCluster
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderUnavailable, GeocoderServiceError
import webbrowser
from PyQt5.QtWidgets import QLabel, QProgressBar

import seaborn as sns
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtChart import (
    QChart, 
    QChartView, 
    QBarSet, 
    QBarSeries, 
    QBarCategoryAxis, 
    QValueAxis, 
    QLineSeries,
)
from PyQt5.QtCore import (
    Qt, 
    QTimer, 
    QUrl, 
    QEventLoop, 
    QSortFilterProxyModel, 
    QThreadPool, 
    QRunnable, 
    pyqtSignal, 
    QObject,
    QRectF,
    pyqtSlot,
)
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QTextEdit,
    QLabel,
    QLineEdit,
    QScrollArea,
    QCheckBox,
    QButtonGroup,
    QFileDialog,
    QMessageBox,
    QTreeWidgetItem,
    QComboBox,
    QTableWidgetItem,
    QHeaderView,
    QTreeWidget,
    QCalendarWidget,
    QStyledItemDelegate,
    QStyleOptionButton, 
    QStyle,
    QStatusBar,
    QTableWidget,  
    QGraphicsScene, 
    QGraphicsView, 
    QGraphicsEllipseItem,
    QStackedWidget,
    QStackedLayout,
    QSizePolicy,
    QFrame,
    QGraphicsTextItem,
    QGraphicsRectItem,
    QTreeView,
    QAbstractItemView,
    
)
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineProfile
from PyQt5.QtGui import (
    QFont, 
    QColor, 
    QPixmap, 
    QImage,
    QIcon, 
    QTextCursor, 
    QTextCharFormat, 
    QPainter,
    QBrush,
    QPen,
    QTransform,
    QFontMetrics,
    
)

import traceback

import mplcursors

import tkinter as tk
from tkinter import (
    Button,
    Entry,
    Label,
    Tk,
    ttk,
    Frame,
    filedialog,
    Text,
    Scrollbar,
    font,
    messagebox,
)
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter.messagebox import showinfo

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import WebDriverException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service as ChromeService

from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import PyPDF2
from PyPDF2 import PdfWriter, PdfReader
from dateutil.relativedelta import relativedelta

import warnings

# Suppress all UserWarnings from openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties

import urllib.parse

import cx_Oracle

import fnmatch

import win32com.client as win32
from exchangelib import Credentials, Account, Configuration, Mailbox, EWSDateTime
import pytz
import csv
import io


from dateutil import parser

import requests
import shutil
from pathlib import Path

from bs4 import BeautifulSoup

import numpy as np

from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, Frame, PageBreak
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.textlabels import Label
#from reportlab.graphics.charts.linecharts import LineChart as LineChart2
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.colors import black, white, HexColor
from reportlab.pdfgen import canvas
import matplotlib.dates as mdates
import os
from PIL import Image
import string

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from calendar import month_abbr
from calendar import monthrange

public_sans_font_path = r"C:\Users\Tyler.Sims\OneDrive - CPower\Documents\CPower\ISONE\ISONEdb\Branding\Public_Sans\PublicSans-VariableFont_wght.ttf"
public_sans_font = TTFont('Public Sans', public_sans_font_path)
pdfmetrics.registerFont(public_sans_font)

public_sans_font_path = r"C:\Users\Tyler.Sims\OneDrive - CPower\Documents\CPower\ISONE\ISONEdb\Branding\Public_Sans\static\PublicSans-Bold.ttf"
public_sans_font = TTFont('Public Sans-Bold', public_sans_font_path)
pdfmetrics.registerFont(public_sans_font)

libre_franklin_font_path = r"C:\Users\Tyler.Sims\OneDrive - CPower\Documents\CPower\ISONE\ISONEdb\Branding\Libre_Franklin\LibreFranklin-VariableFont_wght.ttf"
libre_franklin_font = TTFont('Libre Franklin', libre_franklin_font_path)
pdfmetrics.registerFont(libre_franklin_font)

libre_franklin_font_path = r"C:\Users\Tyler.Sims\OneDrive - CPower\Documents\CPower\ISONE\ISONEdb\Branding\Libre_Franklin\static\LibreFranklin-Medium.ttf"
libre_franklin_font = TTFont('Libre Franklin-Medium', libre_franklin_font_path)
pdfmetrics.registerFont(libre_franklin_font)
