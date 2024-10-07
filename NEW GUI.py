import pandas as pd
import warnings
warnings.filterwarnings("ignore")
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTabWidget, QTextEdit, QGridLayout, QFrame, QSizePolicy, QSpacerItem,QDialog
from PyQt5.QtGui import QPixmap, QFont, QIcon
import threading
import time
import serial
import serial.tools.list_ports
from PyQt5.QtCore import QThread, pyqtSignal,Qt, QTimer
import pyqtgraph as pg
import numpy as np
import matplotlib.pyplot as plt
import os
from datetime import datetime

AllData = ['Slot1','Millis','FET_ON_OFF','CFET ON/OFF','FET TEMP1','DSG_VOLT','SC Current',
    'DSG_Current','CHG_VOLT','CHG_Current','DSG Time','CHG Time','DSG Charge','CHG Charge',
    'Cell1','Cell2','Cell3','Cell4','Cell5','Cell6','Cell7','Cell8','Cell9','Cell10','Cell11',
    'Cell12','Cell13','Cell14','Cell Delta Volt','Sum-of-cells','DSG Power','DSG Energy','CHG Power',
    'CHG Energy','Min CV','BAL_ON_OFF','TS1','TS2','TS3','TS4','TS5','TS6','TS7','TS8','TS9',
    'TS10','TS11','TS12','FET Temp Front','BAT + ve Temp','BAT - ve Temp','Pack + ve Temp','TS0_FLT',
    'TS13_FLT','FET_TEMP_REAR','DSG INA','BAL_RES_TEMP','HUM','IMON','Hydrogen','FG_CELL_VOLT',
    'FG_PACK_VOLT','FG_AVG_CURN','SOC','MAX_TTE','MAX_TTF','REPORTED_CAP','TS0_FLT1','IR','Cycles','DS_CAP',
    'FSTAT','VFSOC','CURR','CAP_NOM','REP_CAP.1','AGE','AGE_FCAST','QH','ICHGTERM','DQACC','DPACC','QRESIDUAL','MIXCAP']
GlobalDictionary = {}
SerialConnectionCycler = None
SerialConnectionVCU = None
SerialConnectionQRScanner = None
PortAssignedToCycler = None
BatteryPackName = None
PackType = None
StoppageCurrentDuetoPowerCut = None
StopConditions = {'DischargingCurrent_':[5], 'DischargingSOC_':[1],'Rest1':[120], 'Charging': [50], 'Rest2':[900]}
FeedParms = {'HR':[154, 42.1, 15000, 115, 57.7, 15000], 'LR':[106, 42.1, 15000, 80, 57.7, 15000]}
########################################
FaultDetectionResults = pd.DataFrame()
########################################
StatusMessages = []
ErrorMessages = []

ConnectionCompleted = None
ScanCompleted= None
StartButtonClicked = None
StartTaskEndsHere = None

ThresholdForOvershoot = 40

FolderForSavingAllFiles = None

CyclingResultsVariable = None

RelativePathHR = "HRRequirementSheet.csv"
RelativePathLR = "LRRequirementSheet.csv"

ResetButtonClicked = None

AllThreads = []

ThresholdTimeToInitializeDisplay = 5

AllowStartTempClause = None

CyclingStatus = None

Rounding_Valv_Check = 4

################################################
ConnectButtonAblingDisabling = 'Able'
ScanButtonAblingDisabling = 'Disable'
StartButtonAblingDisabling = 'Disable'
EmergencyButtonAblingDisabling = 'Disable'
ResetButtonAblingDisabling = 'Disable'
################################################

class MainWindow(QMainWindow):
    
    def __init__(self):
        super().__init__()

        global StatusMessages
        global AllThreads

        self.setWindowTitle("Sample PyQt5 GUI")
        self.setWindowIcon(QIcon("T0-9iuQk_400x400.png"))
        self.setGeometry(QApplication.desktop().screenGeometry())
        self.ImageName = "UV_LOGO.png"

        # Main layout
        self.MainLayout = QVBoxLayout()

        #Image
        self.Image()

        #Buttons
        self.Buttons()

        #Labels and PackInfo
        self.LabelsAndPackInfo()

        #Tabs
        self.Tabs()

        #Errors
        self.StatusErrorResult()

        #ButtonAbleDisable
        self.ButtonAbleDisableThread = ButtonAbleDisable()
        self.ButtonAbleDisableThread.ConnectButtonStatus.connect(self.ButtonConnectFunction)
        self.ButtonAbleDisableThread.ScanButtonStatus.connect(self.ButtonScanFunction)
        self.ButtonAbleDisableThread.StartButtonStatus.connect(self.ButtonStartFunction)
        self.ButtonAbleDisableThread.ResetButtonStatus.connect(self.ButtonResetFunction)
        self.ButtonAbleDisableThread.start()
        AllThreads.append(self.ButtonAbleDisableThread)

        #Status Thread
        self.StatusThread = StatusMessageUpdaterClass()
        self.StatusThread.new_message_signal.connect(self.UpdateStatusMessageBox)
        self.StatusThread.start()
        AllThreads.append(self.StatusThread)

        #Errors Thread
        self.ErrorThread = ErrorMessageUpdaterClass()
        self.ErrorThread.new_message_signal.connect(self.UpdateErrorMessageBox)
        self.ErrorThread.start()
        AllThreads.append(self.ErrorThread)

        #DisplayBMSData
        self.BMSDataDisplayingThread = DisplayBMSData()
        self.BMSDataDisplayingThread.Status.connect(self.UpdateBMSParms)
        self.BMSDataDisplayingThread.start()
        AllThreads.append(self.BMSDataDisplayingThread)

        #DisplayStatusData
        self.StatusDataDisplayinThread = DisplayStatusData()
        self.StatusDataDisplayinThread.Status.connect(self.UpdateStatusParms)
        self.StatusDataDisplayinThread.start()
        AllThreads.append(self.StatusDataDisplayinThread)

        #DataLogging
        self.DataLoggingThread = DataLoggingVCU()
        self.DataLoggingThread.start()
        AllThreads.append(self.DataLoggingThread)

        #DisplayPackInfo
        self.PackInfoDisplayThread = DisplayPackInfo()
        self.PackInfoDisplayThread.Status.connect(self.UpdatePackInfoStatus)
        self.PackInfoDisplayThread.start()
        AllThreads.append(self.PackInfoDisplayThread)

        #DisplayGraph
        self.GraphUpdateThread = DisplayGraph()
        self.GraphUpdateThread.Status.connect(self.UpdatePlot)
        self.GraphUpdateThread.start()
        AllThreads.append(self.GraphUpdateThread)

        #CreatingStorageFolders
        self.StorageFolderCreationThread = CreatingStorageFolders()
        self.StorageFolderCreationThread.start()
        AllThreads.append(self.StorageFolderCreationThread)

        #FinalPassFailDisplay
        self.FinalPassFailDisplayThread = FinalPassFailDisplay()
        self.FinalPassFailDisplayThread.Status.connect(self.UpdateFinalPassFailDisplay)
        self.FinalPassFailDisplayThread.start()
        AllThreads.append(self.FinalPassFailDisplayThread)

        # Set the central widget of the window
        CentralWidget = QWidget()
        CentralWidget.setLayout(self.MainLayout)
        self.setCentralWidget(CentralWidget)

        #Appening StatusMessages
        StatusMessages.append('Ready for Connection')

    def Image(self):
        pixmap = QPixmap(self.ImageName)
        pixmap = pixmap.scaled(500, 400, Qt.KeepAspectRatio, Qt.SmoothTransformation)  # Adjust size as needed
        image_label = QLabel()
        image_label.setPixmap(pixmap)
        image_label.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        spacer = QSpacerItem(20, 100, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.MainLayout.addWidget(image_label)

    def Buttons(self):
        self.ButtonsLayout = QHBoxLayout()
        self.button1 = QPushButton("Connect")
        self.button2 = QPushButton("QRScanner")
        self.button3 = QPushButton("Start")
        self.button4 = QPushButton("Emergency Stop")
        self.button5 = QPushButton("Reset")

         # Set button sizes individually
        self.button1.setFixedSize(350, 40)  # Width: 120, Height: 50
        self.button2.setFixedSize(350, 40)  # Width: 130, Height: 50
        self.button3.setFixedSize(350, 40)  # Width: 140, Height: 50
        self.button4.setFixedSize(350, 40)  # Width: 150, Height: 50
        self.button5.setFixedSize(350, 40)  # Width: 160, Height: 50

        #Styling Buttons
        button1_style = ('''
        QPushButton {background-color: #A3DAFF; color: black; font-size: 16px; font-family: Arial; border-radius: 10px; border: 2px solid #5FBFFF; padding: 10px;}
        QPushButton:hover {background-color: #C2E6FF;}
        QPushButton:pressed {background-color: #7DB8E3; border: 2px solid #7DB8E3; padding-left: 12px; padding-top: 12px;}
        QPushButton:disabled {background-color: #E0E0E0; color: #A0A0A0; border: 1px solid #A0A0A0;}
        ''')
        button2_style = ('''
        QPushButton {background-color: #FFD4A3; color: black; font-size: 16px; font-family: Arial; border-radius: 10px; border: 2px solid #FFAB5F; padding: 10px;}
        QPushButton:hover {background-color: #FFE3C2;}
        QPushButton:pressed {background-color: #E3A07D; border: 2px solid #E3A07D; padding-left: 12px; padding-top: 12px;}
        QPushButton:disabled {background-color: #E0E0E0; color: #A0A0A0; border: 1px solid #A0A0A0;}
        ''')
        button3_style = ('''
        QPushButton {background-color: #96FF8A; color: black; font-size: 16px; font-family: Arial; border-radius: 10px; border: 2px solid #4CFF38; padding: 10px;}
        QPushButton:hover {background-color: #B6FFB2;}
        QPushButton:pressed {background-color: #75D657; border: 2px solid #75D657; padding-left: 12px; padding-top: 12px;}
        QPushButton:disabled {/* Your disabled style here */background-color: #E0E0E0; /* Light gray background for disabled */color: #A0A0A0; /* Light gray text color for disabled *//* You can also adjust other properties to make it visually distinct */border: 1px solid #A0A0A0; /* Light gray border for disabled *//* Add any other styling you desire for disabled buttons */}
        ''')
        button4_style = ('''
        QPushButton {background-color: #FFA3A3; color: black; font-size: 16px; font-family: Arial; border-radius: 10px; border: 2px solid #FF5F5F; padding: 10px;}
        QPushButton:hover {background-color: #FFC2C2;}
        QPushButton:pressed {background-color: #E37D7D; border: 2px solid #E37D7D; padding-left: 12px; padding-top: 12px;}
        QPushButton:disabled {background-color: #E0E0E0; color: #A0A0A0; border: 1px solid #A0A0A0;}
        ''')
        button5_style = ('''
        QPushButton {background-color: #D3B5FF; color: black; font-size: 16px; font-family: Arial; border-radius: 10px; border: 2px solid #A99FFF; padding: 10px;}
        QPushButton:hover {background-color: #E0C8FF;}
        QPushButton:pressed {background-color: #B0A0FF; border: 2px solid #B0A0FF; padding-left: 12px; padding-top: 12px;}
        QPushButton:disabled {background-color: #E0E0E0; color: #A0A0A0; border: 1px solid #A0A0A0;}
        ''')

        self.button1.setStyleSheet(button1_style)
        self.button2.setStyleSheet(button2_style)
        self.button3.setStyleSheet(button3_style)
        self.button4.setStyleSheet(button4_style)
        self.button5.setStyleSheet(button5_style)

        self.ButtonsLayout.addWidget(self.button1)
        self.ButtonsLayout.addWidget(self.button2)
        self.ButtonsLayout.addWidget(self.button3)
        self.ButtonsLayout.addWidget(self.button4)
        self.ButtonsLayout.addWidget(self.button5)

        self.button1.clicked.connect(self.ConnectTaskStart)
        self.button2.clicked.connect(self.QRScannerTaskStart)
        self.button3.clicked.connect(self.StartTaskStart)
        self.button5.clicked.connect(self.ResetTaskStart)

        self.button4.setEnabled(False)

        self.MainLayout.addLayout(self.ButtonsLayout)

    def LabelsAndPackInfo(self):
        self.LabelsLayout = QHBoxLayout()

        def create_frame(label_text):
            frame = QFrame()
            frame.setFrameShape(QFrame.Box)
            frame.setLineWidth(1)
            frame.setFixedSize(180, 80)
            frame.setStyleSheet("""
                QFrame {
                    border: 1px solid black;
                    border-radius: 15px;
                    background-color: white;
                }
            """)

            frame_layout = QVBoxLayout()
            frame.setLayout(frame_layout)

            label1 = QLabel(label_text, frame)
            label1.setAlignment(Qt.AlignCenter)
            label1.setStyleSheet("font-weight: bold;")
            label1.setStyleSheet("font-size: 18px;")
            frame_layout.addWidget(label1)

            label2 = QLabel('None', frame)
            label2.setAlignment(Qt.AlignCenter)
            label2.setStyleSheet("font-size: 16px;")
            frame_layout.addWidget(label2)

            return frame, label2

        labels = [
            'DSG Current (A)',
            'CHG Current (A)',
            'FET Status',
            'Voltage (V)',
            'SoC (%)',
            'Capacity (Ah)',
            'Max Temp (°C)',
            'Timestamp (sec)'
        ]
        self.StatusLabels = []
        for text in labels:
            frame, label2 = create_frame(text)
            self.LabelsLayout.addWidget(frame)
            self.StatusLabels.append(label2)

        #Pack Info
        layout = QVBoxLayout()
        frame = QFrame()
        frame.setFrameShape(QFrame.Box)
        frame.setLineWidth(1)

        headingLabel = QLabel("Battery Pack Info")
        headingLabel.setAlignment(Qt.AlignCenter)

        self.BatInfo = QLabel("None", frame)
        self.BatInfo.setAlignment(Qt.AlignCenter)
        self.BatInfo.setObjectName("batInfoLabel")

        layout.addWidget(headingLabel)
        layout.addWidget(self.BatInfo)
        frame.setLayout(layout)

        frame.setFixedSize(300, 80)
        frame.setStyleSheet("""
            QFrame {
                border: 1px solid black;
                border-radius: 15px;
                background-color: white;
            }
            QLabel {
                font-weight: bold;
            }
            QLabel#batInfoLabel {
                color: blue;
            }
        """)
        self.LabelsLayout.addWidget(frame)
        self.MainLayout.addLayout(self.LabelsLayout)
    
    def Tabs(self):

        global AllData

        # Tab widget
        TabWidget = QTabWidget()

        tab1 = QWidget()
        Tab1Layout = QVBoxLayout()

        # Create a plot widget
        self.plotWidget = pg.PlotWidget()
        Tab1Layout.addWidget(self.plotWidget)

        # Enable multiple y-axes
        self.plotWidget.showAxis('right')
        self.plotWidget.scene().sigMouseClicked.connect(self.mouseClicked)

        # Create another ViewBox for the second y-axis
        self.vb2 = pg.ViewBox()
        self.plotWidget.scene().addItem(self.vb2)
        self.plotWidget.getAxis('right').linkToView(self.vb2)
        self.vb2.setXLink(self.plotWidget)
        
        # Connect the view of vb2 with plotWidget
        def updateViews():
            self.vb2.setGeometry(self.plotWidget.getViewBox().sceneBoundingRect())
            self.vb2.linkedViewChanged(self.plotWidget.getViewBox(), self.vb2.XAxis)
        
        self.plotWidget.getViewBox().sigResized.connect(updateViews)

        # Initialize empty data
        self.x = []
        self.y1 = []
        self.y2 = []

        # Plot initial empty data
        self.curve1 = self.plotWidget.plot(self.x, self.y1, pen=pg.mkPen(color='r', width=2), name="Voltage(V)")
        self.curve2 = pg.PlotCurveItem(self.x, self.y2, pen=pg.mkPen(color='b', width=2, style=Qt.DashLine), name="Current(A)")
        self.vb2.addItem(self.curve2)

        # Add grid lines
        self.plotWidget.showGrid(x=True, y=True, alpha=0.3)

        # Add axis labels
        self.plotWidget.setLabel('left', 'Voltage(V)', color='black')
        self.plotWidget.setLabel('right', 'Current(A)', color='black')
        self.plotWidget.setLabel('bottom', 'Time', color='black')

        # Add a title
        self.plotWidget.setTitle('Voltage-Current vs Time', color='black', size='20pt')

        # Add a legend
        self.legend = self.plotWidget.addLegend()
        self.legend.addItem(self.curve1, 'Voltage (V)')
        self.legend.addItem(self.curve2, 'Current (A)')
        self.legend.setParentItem(self.plotWidget.getPlotItem())
        self.legend.anchor((0.025, 0.05), (0.025,0.05), offset=(10, 10))

        tab1.setLayout(Tab1Layout)
        TabWidget.addTab(tab1, "Graph")

        tab2 = QWidget()
        Tab2Layout = QGridLayout()

        self.AllLabels = []  # List to store the label references

        # Add 84 labels to tab2
        for i in range(84):
            frame = QFrame()
            frame.setFrameShape(QFrame.Box)
            frame.setLineWidth(1)

            # Create a vertical box layout for the frame
            vbox_layout = QVBoxLayout()
            vbox_layout.setAlignment(Qt.AlignCenter)
            vbox_layout.setSpacing(20)

            # Create two labels
            label_top = QLabel(AllData[i], frame)
            label_top.setAlignment(Qt.AlignCenter)
            label_top.setFont(QFont(label_top.font().family(), label_top.font().pointSize(), QFont.Bold))
            
            label_bottom = QLabel('None', frame)
            label_bottom.setAlignment(Qt.AlignCenter)

            font____ = label_bottom.font()
            font____.setPointSize(font____.pointSize() + 2)  # Increase font size by 2 points (adjust as needed)
            font____.setBold(True)
            label_bottom.setFont(font____)

            # Add labels to the vertical box layout
            vbox_layout.addWidget(label_top)
            vbox_layout.addWidget(label_bottom)

            # Set the layout for the frame
            frame.setLayout(vbox_layout)

            Tab2Layout.addWidget(frame, i // 14, i % 14)  # 14 columns in a row
            self.AllLabels.append(label_bottom)
        
        tab2.setLayout(Tab2Layout)
        TabWidget.addTab(tab2, "BMS Data")
        TabWidget.setStyleSheet("""
        QTabBar::tab {
            height: 15px; 
            width: 100px; 
            font-size: 12px; 
            padding: 10px;
        }
        QTabBar::tab:selected {
            background: #dcdcdc;
        }
        """)
        self.MainLayout.addWidget(TabWidget)
    
    def UpdatePlot(self, new_x, new_y1, new_y2):
        # Update data
        self.x.append(new_x)
        self.y1.append(new_y1)
        self.y2.append(new_y2)
        self.curve1.setData(self.x, self.y1)
        self.curve2.setData(self.x, self.y2)

    def mouseClicked(self, event):
        pass

    def StatusErrorResult(self):
        # Status display with heading
        statusLabel = QLabel("Status")
        self.StatusDisplay = QTextEdit()
        
        # Error display with heading
        errorLabel = QLabel("Errors")
        self.ErrorDisplay = QTextEdit()

        # Create a horizontal layout
        hLayout = QHBoxLayout()

        # Add labels and text edits to the horizontal layout
        vLayout2 = QVBoxLayout()
        vLayout2.addWidget(statusLabel)
        vLayout2.addWidget(self.StatusDisplay)
        
        vLayout3 = QVBoxLayout()
        vLayout3.addWidget(errorLabel)
        vLayout3.addWidget(self.ErrorDisplay)

        #Cycling Results
        layout = QVBoxLayout()
        frame = QFrame()
        frame.setFrameShape(QFrame.Box)
        frame.setLineWidth(1)
        headingLabel = QLabel("RESULTS")
        headingLabel.setAlignment(Qt.AlignCenter)
        self.CyclingResult = QLabel("None", frame)
        self.CyclingResult.setAlignment(Qt.AlignCenter)
        layout.addWidget(headingLabel)
        layout.addWidget(self.CyclingResult)
        frame.setLayout(layout)
        frame.setFixedSize(300, 160)
        frame.setStyleSheet("""
            QFrame {
                border: 1px solid black;
                border-radius: 15px;
                background-color: white;
            }
            QLabel {
                font-weight: bold;
                font-size: 12px;
            }
        """)
        
        hLayout.addLayout(vLayout2)
        hLayout.addLayout(vLayout3)
        hLayout.addWidget(frame)

        # Add the horizontal layout to the main layout
        self.MainLayout.addLayout(hLayout)

    def UpdateErrorMessageBox(self, message):
        self.ErrorDisplay.append(message)

    def UpdateStatusMessageBox(self, message):
        self.StatusDisplay.append(message)

    def ConnectTaskStart(self):
        self.ConnectButtonThread = ConnectButton()
        self.ConnectButtonThread.start()

    def QRScannerTaskStart(self):
        global AllThreads
        self.ScanQRThread = ScanQR()
        self.ScanQRThread.start()
        AllThreads.append(self.ScanQRThread)
        
    def StartTaskStart(self):
        global AllThreads
        global GlobalDictionary
        global AllowStartTempClause

        MaxTempBeforeStart = max([GlobalDictionary['TS1'][-1], GlobalDictionary['TS2'][-1], GlobalDictionary['TS3'][-1], GlobalDictionary['TS4'][-1], GlobalDictionary['TS5'][-1], GlobalDictionary['TS6'][-1], GlobalDictionary['TS7'][-1], GlobalDictionary['TS8'][-1], GlobalDictionary['TS9'][-1], GlobalDictionary['TS10'][-1], GlobalDictionary['TS11'][-1], GlobalDictionary['TS12'][-1], GlobalDictionary['TS13_FLT'][-1], GlobalDictionary['TS0_FLT'][-1]])
        if (MaxTempBeforeStart <= 25) or (AllowStartTempClause == 1):
            self.StartButtonThread = StartButton()
            self.StartButtonThread.start()
        else:
            popup = PopupDialog()
            popup.exec_()

    def ResetTaskStart(self):
        self.ResetButtonFunctionsThread = ResetButtonFunctions()
        self.ResetButtonFunctionsThread.Status.connect(self.InitializeDisplays)
        self.ResetButtonFunctionsThread.Status2.connect(self.CallAllThreadsAfterReset)
        self.ResetButtonFunctionsThread.start()

    def UpdateBMSParms(self, Parmss):
        for iterateee in range(len(self.AllLabels)):
            self.AllLabels[iterateee].setText(str(Parmss[iterateee]))

    def UpdateStatusParms(self, DSGCurrentStat, CHGCurrentStat, FETStat, VoltageStat, SoCStat, ReportedCapStat, MaxTemp,Timestamps):
        self.StatusLabels[0].setText(str(DSGCurrentStat))
        self.StatusLabels[1].setText(str(CHGCurrentStat))
        self.StatusLabels[2].setText(str(FETStat))
        self.StatusLabels[3].setText(str(VoltageStat))
        self.StatusLabels[4].setText(str(SoCStat))
        self.StatusLabels[5].setText(str(ReportedCapStat))
        self.StatusLabels[6].setText(str(MaxTemp))
        self.StatusLabels[7].setText(str(Timestamps))

    def UpdatePackInfoStatus(self, PackInfoParam):
        self.BatInfo.setText(str(PackInfoParam))
        
    def UpdateFinalPassFailDisplay(self, StringValv):
        self.CyclingResult.setText(str(StringValv))

    def InitializeDisplays(self, Valvvv):
        #Initialize graphs
        self.x = []
        self.y1 = []
        self.y2 = []
        self.curve1.setData(self.x, self.y1)
        self.curve2.setData(self.x, self.y2)

        #Initialize Status Data
        for statuslabelsiterate in range(len(self.StatusLabels)):
            self.StatusLabels[statuslabelsiterate].setText(str('None'))

        #Initialize BMS Data
        for BMSlabelsiterate in range(len(self.AllLabels)):
            self.AllLabels[BMSlabelsiterate].setText(str('None'))

        #Initialize Pack Number
        self.BatInfo.setText(str('None'))

        #Inititalize Results
        self.CyclingResult.setText(str('None'))

        #Clear Status Display
        self.StatusDisplay.clear()

        #Clear Error Display
        self.ErrorDisplay.clear()

    def CallAllThreadsAfterReset(self):
        global StatusMessages
        global ConnectButtonAblingDisabling
        global ScanButtonAblingDisabling
        global StartButtonAblingDisabling
        global EmergencyButtonAblingDisabling
        global ResetButtonAblingDisabling

        #ButtonAbleDisable
        self.ButtonAbleDisableThread = ButtonAbleDisable()
        self.ButtonAbleDisableThread.ConnectButtonStatus.connect(self.ButtonConnectFunction)
        self.ButtonAbleDisableThread.ScanButtonStatus.connect(self.ButtonScanFunction)
        self.ButtonAbleDisableThread.StartButtonStatus.connect(self.ButtonStartFunction)
        self.ButtonAbleDisableThread.ResetButtonStatus.connect(self.ButtonResetFunction)
        self.ButtonAbleDisableThread.start()
        AllThreads.append(self.ButtonAbleDisableThread)

        #Status Thread
        self.StatusThread = StatusMessageUpdaterClass()
        self.StatusThread.new_message_signal.connect(self.UpdateStatusMessageBox)
        self.StatusThread.start()
        AllThreads.append(self.StatusThread)

        #Errors Thread
        self.ErrorThread = ErrorMessageUpdaterClass()
        self.ErrorThread.new_message_signal.connect(self.UpdateErrorMessageBox)
        self.ErrorThread.start()
        AllThreads.append(self.ErrorThread)

        #DisplayBMSData
        self.BMSDataDisplayingThread = DisplayBMSData()
        self.BMSDataDisplayingThread.Status.connect(self.UpdateBMSParms)
        self.BMSDataDisplayingThread.start()
        AllThreads.append(self.BMSDataDisplayingThread)

        #DisplayStatusData
        self.StatusDataDisplayinThread = DisplayStatusData()
        self.StatusDataDisplayinThread.Status.connect(self.UpdateStatusParms)
        self.StatusDataDisplayinThread.start()
        AllThreads.append(self.StatusDataDisplayinThread)

        #DataLogging
        self.DataLoggingThread = DataLoggingVCU()
        self.DataLoggingThread.start()
        AllThreads.append(self.DataLoggingThread)

        #DisplayPackInfo
        self.PackInfoDisplayThread = DisplayPackInfo()
        self.PackInfoDisplayThread.Status.connect(self.UpdatePackInfoStatus)
        self.PackInfoDisplayThread.start()
        AllThreads.append(self.PackInfoDisplayThread)

        #DisplayGraph
        self.GraphUpdateThread = DisplayGraph()
        self.GraphUpdateThread.Status.connect(self.UpdatePlot)
        self.GraphUpdateThread.start()
        AllThreads.append(self.GraphUpdateThread)

        #CreatingStorageFolders
        self.StorageFolderCreationThread = CreatingStorageFolders()
        self.StorageFolderCreationThread.start()
        AllThreads.append(self.StorageFolderCreationThread)

        #FinalPassFailDisplay
        self.FinalPassFailDisplayThread = FinalPassFailDisplay()
        self.FinalPassFailDisplayThread.Status.connect(self.UpdateFinalPassFailDisplay)
        self.FinalPassFailDisplayThread.start()
        AllThreads.append(self.FinalPassFailDisplayThread)

        ConnectButtonAblingDisabling = 'Able'
        ScanButtonAblingDisabling = 'Disable'
        StartButtonAblingDisabling = 'Disable'
        EmergencyButtonAblingDisabling = 'Disable'
        ResetButtonAblingDisabling = 'Disable'

        #Appening StatusMessages
        StatusMessages.append('Ready for Connection')

    def ButtonConnectFunction(self, value):
        if value == 1:
            if not self.button1.isEnabled():
                self.button1.setEnabled(True)
        elif value == 0:
            if self.button1.isEnabled():
                self.button1.setEnabled(False)
    
    def ButtonScanFunction(self, value):
        if value == 1:
            if not self.button2.isEnabled():
                self.button2.setEnabled(True)
        elif value == 0:
            if self.button2.isEnabled():
                self.button2.setEnabled(False)
    
    def ButtonStartFunction(self, value):
        if value == 1:
            if not self.button3.isEnabled():
                self.button3.setEnabled(True)
        elif value == 0:
            if self.button3.isEnabled():
                self.button3.setEnabled(False)
    
    def ButtonResetFunction(self, value):
        if value == 1:
            if not self.button5.isEnabled():
                self.button5.setEnabled(True)
        elif value == 0:
            if self.button5.isEnabled():
                self.button5.setEnabled(False)

    def closeEvent(self, event):
        global SerialConnectionCycler
        if SerialConnectionCycler:
            #Initialize Cycler
            SerialConnectionCycler.write(b"SYST:LOCK ON")
            #Turn off cycler
            for trying in range(0,10):
                SerialConnectionCycler.write(b"OUTP OFF")
                time.sleep(0.1)
                SerialConnectionCycler.write(b"OUTPUT?")
                if SerialConnectionCycler.readline().decode().strip() == 'OFF':
                    time.sleep(0.1)
                    break
        os._exit(0)

class PopupDialog(QDialog):
    
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("PopUp")
        self.resize(300, 200)
        
        # Create label
        self.label = QLabel("Battery Temp > 25DegreeC!", self)
        self.label.setFont(QFont("Arial", 16, QFont.Bold))
        self.label.setStyleSheet("color: #2c3e50; text-align: center; padding: 10px;")
        
        # Create buttons
        self.button1 = QPushButton("Proceed", self)
        self.button2 = QPushButton("Wait", self)
        
        # Connect buttons to functions
        self.button1.clicked.connect(self.button1_clicked)
        self.button2.clicked.connect(self.button2_clicked)
        
        # Layout for buttons
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.button1)
        button_layout.addWidget(self.button2)
        
        # Main layout
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.label)
        main_layout.addLayout(button_layout)
        
        self.setLayout(main_layout)
    
    def button1_clicked(self):
        global AllowStartTempClause
        global StatusMessages
        AllowStartTempClause = 1
        StatusMessages.append('Click Start Button to start the Cycling')
        self.accept()

    def button2_clicked(self):
        self.reject()

class StatusMessageUpdaterClass(QThread):
    
    new_message_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()

    def run(self):
        global StatusMessages
        global ResetButtonClicked
        while ResetButtonClicked != 'Clicked':
            if len(StatusMessages) > 0:
                self.new_message_signal.emit(StatusMessages[0])
                StatusMessages = StatusMessages[1:]
            time.sleep(0.01)

class ErrorMessageUpdaterClass(QThread):
    
    new_message_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()

    def run(self):
        global ErrorMessages
        global ResetButtonClicked
        while ResetButtonClicked != 'Clicked':
            if len(ErrorMessages) > 0:
                self.new_message_signal.emit(ErrorMessages[0])
                ErrorMessages = ErrorMessages[1:]
            time.sleep(0.01)

class ConnectButton(threading.Thread):
    
    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
        global SerialConnectionCycler
        global SerialConnectionVCU
        global SerialConnectionQRScanner
        global PortAssignedToCycler
        global ConnectionCompleted
        global StatusMessages
        global ConnectButtonAblingDisabling
        global ScanButtonAblingDisabling
        global ResetButtonAblingDisabling

        ConnectButtonAblingDisabling = 'Disable'

        available_ports = [port.device for port in serial.tools.list_ports.comports()]

        StatusMessages.append('Connection in Progress...')

        #QR Scanner
        scannerPort = self.listports()
        if scannerPort:
            SerialConnectionQRScanner = serial.Serial(port=scannerPort,baudrate = 115200,parity=serial.PARITY_NONE,stopbits=serial.STOPBITS_ONE,bytesize=serial.EIGHTBITS,timeout=1)
            available_ports = [pr for pr in available_ports if pr != scannerPort]
        if SerialConnectionQRScanner != None: 
            StatusMessages.append('Success : Connection with QR Scanner')
        else:
            StatusMessages.append('Fail : Connection with QR Scanner')
        #Cycler
        for port in serial.tools.list_ports.comports():
            if port.device in available_ports:
                if 'PSB' in port.description:
                    SerialConnectionCycler = serial.Serial(port.device, baudrate=9600, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE, bytesize=serial.EIGHTBITS, timeout=1)
                    available_ports = [pr for pr in available_ports if pr != port.device]
                    PortAssignedToCycler = port.device
                    break
        if SerialConnectionCycler != None:
            StatusMessages.append('Success : Connection with Cycler')
        else:
            StatusMessages.append('Fail : Connection with Cycler')
        #VCU
        for port in serial.tools.list_ports.comports():
            if port.device in available_ports:
                if 'USB' in port.description:
                    SerialConnectionVCU = serial.Serial(port.device, baudrate=115200, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE, bytesize=serial.EIGHTBITS, timeout=1)
                    available_ports = [pr for pr in available_ports if pr != port.device]
                    break
        if SerialConnectionVCU != None:
            StatusMessages.append('Success : Connection with VCU')
        else:
            StatusMessages.append('Fail : Connection with VCU')
        #Status Messages
        if (SerialConnectionQRScanner != None) and (SerialConnectionVCU != None) and (SerialConnectionCycler != None):
            StatusMessages.append('Success : Connection')
            StatusMessages.append('Started BMS Data Logging')
            ConnectionCompleted = 'Success'
            ScanButtonAblingDisabling = 'Able'
            ResetButtonAblingDisabling = 'Able'
            StatusMessages.append('Ready for Scanning Battery Pack')
        else:
            if SerialConnectionQRScanner:
                SerialConnectionQRScanner.close()
            if SerialConnectionCycler:
                SerialConnectionCycler.close()
            if SerialConnectionVCU:
                SerialConnectionVCU.close()
            ConnectButtonAblingDisabling = 'Able'
            StatusMessages.append('Fail : Connection | Check Connections and Connect Again')
        ######################################################################

    def is_micropython_usb_device(self, port):

        """Checks a USB device to see if it looks like a MicroPython device.
        """
        if type(port).__name__ == 'Device':
            # Assume its a pyudev.device.Device
            if ('ID_BUS' not in port or port['ID_BUS'] != 'usb' or
                'SUBSYSTEM' not in port or port['SUBSYSTEM'] != 'tty'):
                return False
            usb_id = 'usb vid:pid={}:{}'.format(port['ID_VENDOR_ID'], port['ID_MODEL_ID'])
        else:
            # Assume its a port from serial.tools.list_ports.comports()
            usb_id = port[2].lower()
        # We don't check the last digit of the PID since there are 3 possible
        # values.
        if usb_id.startswith('usb vid:de=f055:980'):
            return True
        # Check for Teensy VID:PID
        if usb_id.startswith('usb vid:pid=16c0:0483'):
            return True
        return False

    def listports(self):
        detected = False
        scannerPort = ""
        for port in serial.tools.list_ports.comports():
            detected = True
            if port.vid:
                micropythonPort = ''
                if self.is_micropython_usb_device(port):
                    micropythonPort = ' *'
                if (port.vid, port.pid) == (9969, 22096):
                    scannerPort = str(port.device) + str(micropythonPort)
                    break
                elif (port.vid, port.pid) == (9168, 3178):
                    scannerPort = str(port.device) + str(micropythonPort)
                    break
        return scannerPort

class DataLoggingVCU(threading.Thread):

    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):

        global GlobalDictionary
        global SerialConnectionVCU
        global ResetButtonClicked
        global ConnectionCompleted
        global CyclingStatus

        SLFaultArray = ['STAT_DSG_FET_STATUS_FLAG','STAT_CHG_FET_STATUS_FLAG','STAT_BAL_TIMER_STATUS_FLAG','STAT_BAL_ACT_STATUS_FLAG','STAT_LTC2946_DSG_ALERT_FLAG','STAT_LTC2946_CHG_ALERT_FLAG','STAT_PWR_MODE_CHARGE','STAT_PWR_MODE_CHARGE_NX','STAT_BMS_UNECOVERABLE_FAILURE','STAT_UV_THR_FLAG','STAT_OV_THR_FLAG','STAT_LTC6812_WDT_SET_FLAG','STAT_BATTERY_TEMP_OVER_MIN_THRESHOLD','STAT_BATTERY_TEMP_OVER_MAX_THRESHOLD','STAT_BATTERY_TEMP_TOO_LOW','STAT_LTC6812_SAFETY_TIMER_FLAG','STAT_BALANCER_ABORT_FLAG','STAT_BALANCER_RESET_FLAG','STAT_BALANCING_COMPLETE_FLAG','STAT_LTC6812_PEC_ERROR','STAT_UV_OV_THR_FOR_TURN_ON','STAT_ECC_ERM_ERR_FLAG','STAT_DSG_INA302_ALERT1','STAT_DSG_INA302_ALERT2','STAT_MOSFET_OVER_TMP_ALERT','STAT_FET_FRONT_OVER_TMP_ALERT','STAT_BAT_PLUS_OVER_TMP_ALERT','STAT_BAT_MINUS_OVER_TMP_ALERT','STAT_PACK_PLUS_MCPCB_OVER_TMP_ALERT','STAT_REL_HUMIDITY_OVERVALUE_ALERT','STAT_DSG_FUSE_BLOWN_ALERT','STAT_CHG_FUSE_BLOWN_ALERT']
        SHValvArray = ['STAT_FET_TURN_ON_FAILURE','STAT_FET_TURN_OFF_FAILURE','STAT_BAL_RES_OVER_TEMPERATURE','STAT_LTC2946_COMM_FAILURE','STAT_HW_UV_SHUTDOWN','STAT_HW_OV_SHUTDOWN','STAT_HW_OVER_TMP_SHUTDOWN','STAT_LTC7103_PGOOD','STAT_SYS_BOOT_FAILURE','STAT_CAN_MSG_SIG_ERR','STAT_FG_I2C_BUS_RECOVERY_EXEC','STAT_FG_MEAS_ABORT','STAT_BAT_TAMPER_DETECTED','STAT_TMP_THR_FOR_TURN_ON','STAT_FET_FRONT_OVER_TMP_WARN','STAT_BAT_PLUS_OVER_TMP_WARN','STAT_BAT_MINUS_OVER_TMP_WARN','STAT_PACK_PLUS_OVER_TMP_WARN','STAT_FG_MEAS_ERROR','STAT_PM_CHG_CURRENT_LIMIT_UPDATE','STAT_H2_SNS_ALERT','STAT_THRM_RUNAWAY_ALRT_V','STAT_THRM_RUNAWAY_ALRT_T','STAT_THRM_RUNAWAY_ALRT_H','STAT_PRE_DISCHARGE_STRESSED','STAT_FG_SETTINGS_UPDATE','STAT_BIT_UNUSED16','STAT_BIT_UNUSED17','STAT_BIT_UNUSED18','STAT_BIT_UNUSED19','STAT_BIT_UNUSED20','MAX_BMS_STATUS_FLAGS']
        NotErrorsArr = []#['STAT_DSG_FET_STATUS_FLAG', 'STAT_CHG_FET_STATUS_FLAG', 'STAT_BAL_TIMER_STATUS_FLAG', 'STAT_BAL_ACT_STATUS_FLAG','STAT_PWR_MODE_CHARGE','STAT_PWR_MODE_CHARGE_NX','STAT_LTC6812_WDT_SET_FLAG', 'STAT_LTC6812_SAFETY_TIMER_FLAG', 'STAT_BALANCER_ABORT_FLAG','STAT_BALANCER_RESET_FLAG','STAT_BALANCING_COMPLETE_FLAG','STAT_LTC6812_PEC_ERROR','STAT_ECC_ERM_ERR_FLAG', 'STAT_LTC2946_COMM_FAILURE', 'STAT_LTC7103_PGOOD','STAT_SYS_BOOT_FAILURE','STAT_CAN_MSG_SIG_ERR','STAT_FG_I2C_BUS_RECOVERY_EXEC','STAT_FG_MEAS_ABORT,STAT_TMP_THR_FOR_TURN_ON','STAT_FET_FRONT_OVER_TMP_WARN','STAT_BAT_PLUS_OVER_TMP_WARN,STAT_FG_SETTINGS_UPDATE','STAT_BAT_MINUS_OVER_TMP_WARN','STAT_PACK_PLUS_OVER_TMP_WARN','STAT_FG_MEAS_ERROR','STAT_PM_CHG_CURRENT_LIMIT_UPDATE','STAT_H2_SNS_ALERT']
        AlreadyFetchedFaults = []

        while ResetButtonClicked != 'Clicked':
            if ConnectionCompleted == 'Success':
                while True:
                    time.sleep(0.01)
                    try:
                        string_data_res = SerialConnectionVCU.read(size = 2000).decode('utf-8')
                        #Errors
                        index_Valv_string = string_data_res.find('I,0,SL | SH:')
                        ErrorHexValue = string_data_res[index_Valv_string:][:string_data_res[index_Valv_string:].find('\n')][string_data_res[index_Valv_string:][:string_data_res[index_Valv_string:].find('\n')].find(':'):][2:]
                        SLValv = bin(int(ErrorHexValue[:ErrorHexValue.find(' | ')], 16))[2:].zfill(32)[::-1]
                        SHValv = bin(int(ErrorHexValue[ErrorHexValue.find('| '):][2:], 16))[2:].zfill(32)[::-1]
                        SLValvIndex = [i for i in range(len(SLValv)) if SLValv.startswith('1', i)]
                        SHValvindex = [i for i in range(len(SHValv)) if SHValv.startswith('1', i)]
                        if len(SLValvIndex) > 0:
                            for TargetIndex_ in SLValvIndex:
                                if (SLFaultArray[TargetIndex_] not in NotErrorsArr) and (SLFaultArray[TargetIndex_] not in AlreadyFetchedFaults):
                                    AlreadyFetchedFaults.append(SLFaultArray[TargetIndex_])
                                    Time_Now = datetime.now()
                                    Time_Now_PRINT = Time_Now.strftime("%Y-%m-%d %H-%M")
                                    ErrorMessages.append(f'Error : {SLFaultArray[TargetIndex_]} ({Time_Now_PRINT})')
                        if len(SHValvindex) > 0:
                            for TargetIndex_ in SHValvindex:
                                if (SHValvArray[TargetIndex_] not in NotErrorsArr) and (SHValvArray[TargetIndex_] not in AlreadyFetchedFaults):
                                    AlreadyFetchedFaults.append(SHValvArray[TargetIndex_])
                                    Time_Now = datetime.now()
                                    Time_Now_PRINT = Time_Now.strftime("%Y-%m-%d %H-%M")
                                    ErrorMessages.append(f'Error : {SHValvArray[TargetIndex_]} ({Time_Now_PRINT})')
                        #Other Data
                        if '*' not in string_data_res:
                            if '(' and ')' and '[' and ']' and'{' and '}' in string_data_res:
                                #Initializing
                                string_curved_bracket = ''
                                string_flower_bracket = ''
                                string_squared_bracket = ''
                                #Curved brackets
                                curverd_bracket_starting_indexes = [i for i, char in enumerate(string_data_res) if char == '(']
                                curverd_bracket_ending_indexes = [i for i, char in enumerate(string_data_res) if char == ')']
                                #Flower brackets
                                flowerd_bracket_starting_indexes = [i for i, char in enumerate(string_data_res) if char == '{']
                                flowerd_bracket_ending_indexes = [i for i, char in enumerate(string_data_res) if char == '}']
                                #Square brackets
                                squared_bracket_starting_indexes = [i for i, char in enumerate(string_data_res) if char == '[']
                                squared_bracket_ending_indexes = [i for i, char in enumerate(string_data_res) if char == ']']
                                #Extracting the latest Curved bracket string
                                for i in range(0, len(curverd_bracket_starting_indexes)):
                                    if curverd_bracket_ending_indexes[0] > curverd_bracket_starting_indexes[i]:
                                        string_curved_bracket = string_data_res[curverd_bracket_starting_indexes[i]+1:curverd_bracket_ending_indexes[0]]
                                #Extracting the latest Flower bracket string
                                for i in range(0, len(flowerd_bracket_starting_indexes)):
                                    if flowerd_bracket_ending_indexes[0] > flowerd_bracket_starting_indexes[i]:
                                        string_flower_bracket = string_data_res[flowerd_bracket_starting_indexes[i]+1:flowerd_bracket_ending_indexes[0]]
                                #Extracting the latest Squared bracket string
                                for i in range(0, len(squared_bracket_starting_indexes)):
                                    if squared_bracket_ending_indexes[0] > squared_bracket_starting_indexes[i]:
                                        string_squared_bracket = string_data_res[squared_bracket_starting_indexes[i]+1:squared_bracket_ending_indexes[0]]
                                if string_curved_bracket and string_flower_bracket and string_squared_bracket != '':
                                    #Extracting the data
                                    array_string_flower_bracket = string_flower_bracket.split(",")
                                    array_string_squared_bracket = string_squared_bracket.split(",")
                                    array_string_curved_bracket = string_curved_bracket.split(",")
                                    if len(array_string_flower_bracket) == 20 and len(array_string_squared_bracket) == 36 and len(array_string_curved_bracket) == 32:
                                        MILLIS2 = array_string_flower_bracket[1]
                                        TS1 = float(array_string_flower_bracket[2])
                                        TS2 = float(array_string_flower_bracket[3])
                                        TS3 = float(array_string_flower_bracket[4])
                                        TS4 = float(array_string_flower_bracket[5])
                                        TS5 = float(array_string_flower_bracket[6])
                                        TS6 = float(array_string_flower_bracket[7])
                                        TS7 = float(array_string_flower_bracket[8])
                                        TS8 = float(array_string_flower_bracket[9])
                                        TS9 = float(array_string_flower_bracket[10])
                                        TS10 = float(array_string_flower_bracket[11])
                                        TS11 = float(array_string_flower_bracket[12])
                                        TS12 = float(array_string_flower_bracket[13])
                                        FET_Temp_Front = float(array_string_flower_bracket[14])
                                        BAT_POS_TEMP = float(array_string_flower_bracket[15])
                                        BAT_NEG_TEMP = float(array_string_flower_bracket[16])
                                        PACK_POS_TEMP = float(array_string_flower_bracket[17])
                                        TS0_FLT = float(array_string_flower_bracket[18])
                                        TS13_FLT = float(array_string_flower_bracket[19])

                                        SLOT1 = float(array_string_squared_bracket[0])
                                        MILLIS = float(array_string_squared_bracket[1])
                                        FET_ON_OFF = float(array_string_squared_bracket[2])
                                        CFET_ON_OFF = float(array_string_squared_bracket[3])
                                        FET_TEMP1 = float(array_string_squared_bracket[4])
                                        DSG_VOLT = float(array_string_squared_bracket[5])
                                        SC_Current = float(array_string_squared_bracket[6])
                                        DSG_Current = float(array_string_squared_bracket[7])
                                        CHG_VOLT = float(array_string_squared_bracket[8])
                                        CHG_Current = float(array_string_squared_bracket[9])
                                        DSG_Time = float(array_string_squared_bracket[10])
                                        CHG_Time = float(array_string_squared_bracket[11])
                                        DSG_Charge = float(array_string_squared_bracket[12])
                                        CHG_Charge = float(array_string_squared_bracket[13])
                                        Cell1 = float(array_string_squared_bracket[14])
                                        Cell2 = float(array_string_squared_bracket[15])
                                        Cell3 = float(array_string_squared_bracket[16])
                                        Cell4 = float(array_string_squared_bracket[17])
                                        Cell5 = float(array_string_squared_bracket[18])
                                        Cell6 = float(array_string_squared_bracket[19])
                                        Cell7 = float(array_string_squared_bracket[20])
                                        Cell8 = float(array_string_squared_bracket[21])
                                        Cell9 = float(array_string_squared_bracket[22])
                                        Cell10 = float(array_string_squared_bracket[23])
                                        Cell11 = float(array_string_squared_bracket[24])
                                        Cell12 = float(array_string_squared_bracket[25])
                                        Cell13 = float(array_string_squared_bracket[26])
                                        Cell14 = float(array_string_squared_bracket[27])
                                        Cell_Delta_Volt = float(array_string_squared_bracket[28])
                                        Sum_of_cells = float(array_string_squared_bracket[29])
                                        DSG_Power = float(array_string_squared_bracket[30])
                                        DSG_Energy = float(array_string_squared_bracket[31])
                                        CHG_Power = float(array_string_squared_bracket[32])
                                        CHG_Energy = float(array_string_squared_bracket[33])
                                        Min_CV = float(array_string_squared_bracket[34])
                                        BAL_ON_OFF = float(array_string_squared_bracket[35])

                                        MILLIS3 = float(array_string_curved_bracket[1])
                                        FET_TEMP_REAR = float(array_string_curved_bracket[2])
                                        DSG_INA = float(array_string_curved_bracket[3])
                                        BAL_RES_TEMP = float(array_string_curved_bracket[4])
                                        HUM = float(array_string_curved_bracket[5])
                                        IMON = float(array_string_curved_bracket[6])
                                        Hydrogen = float(array_string_curved_bracket[7])
                                        FG_CELL_VOLT = float(array_string_curved_bracket[8])
                                        FG_PACK_VOLT = float(array_string_curved_bracket[9])
                                        FG_AVG_CURN = float(array_string_curved_bracket[10])
                                        SOC = float(array_string_curved_bracket[11])
                                        MAX_TTE = float(array_string_curved_bracket[12])
                                        MAX_TTF = float(array_string_curved_bracket[13])
                                        REPORTED_CAP = float(array_string_curved_bracket[14])
                                        TS0_FLT1 = float(array_string_curved_bracket[15])
                                        IR = float(array_string_curved_bracket[16])
                                        Cycles = float(array_string_curved_bracket[17])
                                        DS_CAP = float(array_string_curved_bracket[18])
                                        FSTAT = float(array_string_curved_bracket[19])
                                        VFSOC = float(array_string_curved_bracket[20])
                                        CURR = float(array_string_curved_bracket[21])
                                        CAP_NOM = float(array_string_curved_bracket[22])
                                        REP_CAP_1 = float(array_string_curved_bracket[23])
                                        AGE = float(array_string_curved_bracket[24])
                                        AGE_FCAST = float(array_string_curved_bracket[25])
                                        QH = float(array_string_curved_bracket[26])
                                        ICHGTERM = float(array_string_curved_bracket[27])
                                        DQACC = float(array_string_curved_bracket[28])
                                        DPACC = float(array_string_curved_bracket[29])
                                        QRESIDUAL = float(array_string_curved_bracket[30])
                                        MIXCAP = float(array_string_curved_bracket[31])

                                        #Creating a dictionary
                                        DICTIONARY = {'Cycling Status':[CyclingStatus],'Slot1':[SLOT1],'Millis':[MILLIS],'FET_ON_OFF':[FET_ON_OFF],'CFET ON/OFF':[CFET_ON_OFF],'FET TEMP1':[FET_TEMP1],'DSG_VOLT':[DSG_VOLT],'SC Current':[SC_Current],'DSG_Current':[DSG_Current],'CHG_VOLT':[CHG_VOLT],'CHG_Current':[CHG_Current],'DSG Time':[DSG_Time],'CHG Time':[CHG_Time],'DSG Charge':[DSG_Charge],'CHG Charge':[CHG_Charge],'Cell1':[Cell1],'Cell2':[Cell2],'Cell3':[Cell3],'Cell4':[Cell4],'Cell5':[Cell5],'Cell6':[Cell6],'Cell7':[Cell7],'Cell8':[Cell8],'Cell9':[Cell9],'Cell10':[Cell10],'Cell11':[Cell11],'Cell12':[Cell12],'Cell13':[Cell13],'Cell14':[Cell14],'Cell Delta Volt':[Cell_Delta_Volt],'Sum-of-cells':[Sum_of_cells],'DSG Power':[DSG_Power],'DSG Energy':[DSG_Energy],'CHG Power':[CHG_Power],'CHG Energy':[CHG_Energy],'Min CV':[Min_CV],'BAL_ON_OFF':[BAL_ON_OFF],'Millis2':[MILLIS2],'TS1':[TS1],'TS2':[TS2],'TS3':[TS3],'TS4':[TS4],'TS5':[TS5],'TS6':[TS6],'TS7':[TS7],'TS8':[TS8],'TS9':[TS9],'TS10':[TS10],'TS11':[TS11],'TS12':[TS12],'FET Temp Front':[FET_Temp_Front],'BAT + ve Temp':[BAT_POS_TEMP],'BAT - ve Temp':[BAT_NEG_TEMP],'Pack + ve Temp':[PACK_POS_TEMP],'TS0_FLT':[TS0_FLT],'TS13_FLT':[TS13_FLT],'Millis3':[MILLIS3],'FET_TEMP_REAR':[FET_TEMP_REAR],'DSG INA':[DSG_INA],'BAL_RES_TEMP':[BAL_RES_TEMP],'HUM':[HUM],'IMON':[IMON],'Hydrogen':[Hydrogen],'FG_CELL_VOLT':[FG_CELL_VOLT],'FG_PACK_VOLT':[FG_PACK_VOLT],'FG_AVG_CURN':[FG_AVG_CURN],'SOC':[SOC],'MAX_TTE':[MAX_TTE],'MAX_TTF':[MAX_TTF],'REPORTED_CAP':[REPORTED_CAP],'TS0_FLT1':[TS0_FLT1],'IR':[IR],'Cycles':[Cycles],'DS_CAP':[DS_CAP],'FSTAT':[FSTAT],'VFSOC':[VFSOC],'CURR':[CURR],'CAP_NOM':[CAP_NOM],'REP_CAP.1':[REP_CAP_1],'AGE':[AGE],'AGE_FCAST':[AGE_FCAST],'QH':[QH],'ICHGTERM':[ICHGTERM],'DQACC':[DQACC],'DPACC':[DPACC],'QRESIDUAL':[QRESIDUAL],'MIXCAP':[MIXCAP]}                        
                                        for key, value in DICTIONARY.items():
                                            if key in GlobalDictionary:
                                                GlobalDictionary[key].extend(value)
                                            else:
                                                GlobalDictionary[key] = value
                                        break
                    except:
                        pass
            time.sleep(0.01)

class ScanQR(threading.Thread):

    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
        global SerialConnectionQRScanner
        global BatteryPackName
        global PackType
        global ScanCompleted
        global StatusMessages
        global ResetButtonClicked
        global ScanButtonAblingDisabling
        global StartButtonAblingDisabling

        ScanButtonAblingDisabling = 'Disable'

        LocalVariableValidation = False
        while ResetButtonClicked != 'Clicked':
            if LocalVariableValidation == False:
                StatusMessages.append('Scanning....')
                LocalVariableValidation = True
            line = SerialConnectionQRScanner.readline()
            line = line.decode("utf-8")
            if len(line) > 0:
                line = line.rstrip()
                BatteryPackName = line
                PackVarient = line.partition('-')[0][-2]
                if PackVarient == '2':
                    PackType = 'HR'
                elif PackVarient == '3':
                    PackType = 'LR'
                ScanCompleted = 'Success'
                StatusMessages.append('Scanning Completed. Battery Pack Info Displayed')
                StatusMessages.append('Ready for Cycling')
                StartButtonAblingDisabling = 'Able'
                break
            time.sleep(0.01)

class CreatingStorageFolders(threading.Thread):

    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
        global ScanCompleted
        global BatteryPackName
        global FolderForSavingAllFiles
        global ResetButtonClicked

        while ResetButtonClicked != 'Clicked':
            if ScanCompleted == 'Success':
                FolderName = "Battery Folder"
                os.makedirs(FolderName, exist_ok=True)
                BatteryPackFilesPath = os.path.join(FolderName, BatteryPackName.replace(':','-'))
                os.makedirs(BatteryPackFilesPath, exist_ok=True)
                now = datetime.now()
                folder_name = now.strftime("%Y-%m-%d %H-%M")
                FolderForSavingAllFiles = os.path.join(BatteryPackFilesPath, folder_name)
                os.makedirs(FolderForSavingAllFiles, exist_ok=True)
                break
            time.sleep(1)

class DisplayStatusData(QThread):

    Status = pyqtSignal(float, float, float, float, float, float, float, float)

    def __init__(self):
        super().__init__()
        self.timer = QTimer()
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.update_timer)
        self.timer.start()
        self.current_time = 0

    def run(self):
        global GlobalDictionary
        global ConnectionCompleted
        global ResetButtonClicked
        global Rounding_Valv_Check

        while ResetButtonClicked != 'Clicked':
            if (ConnectionCompleted == 'Success') and (len(GlobalDictionary) != 0):
                self.Status.emit(
                    round(GlobalDictionary['DSG_Current'][-1], Rounding_Valv_Check),
                    round(GlobalDictionary['CHG_Current'][-1], Rounding_Valv_Check),
                    round(GlobalDictionary['FET_ON_OFF'][-1], Rounding_Valv_Check),
                    round(GlobalDictionary['Sum-of-cells'][-1], Rounding_Valv_Check),
                    round(GlobalDictionary['SOC'][-1], Rounding_Valv_Check),
                    round(GlobalDictionary['REPORTED_CAP'][-1], Rounding_Valv_Check),
                    round(max([GlobalDictionary['TS1'][-1], GlobalDictionary['TS2'][-1], GlobalDictionary['TS3'][-1], GlobalDictionary['TS4'][-1], GlobalDictionary['TS5'][-1], GlobalDictionary['TS6'][-1], GlobalDictionary['TS7'][-1], GlobalDictionary['TS8'][-1], GlobalDictionary['TS9'][-1], GlobalDictionary['TS10'][-1], GlobalDictionary['TS11'][-1], GlobalDictionary['TS12'][-1], GlobalDictionary['TS13_FLT'][-1], GlobalDictionary['TS0_FLT'][-1]]), Rounding_Valv_Check),
                    self.current_time
                )
            time.sleep(0.1)
    
    def update_timer(self):
        self.current_time += 1

class DisplayBMSData(QThread):

    Status = pyqtSignal(list)

    def __init__(self):
        super().__init__()

    def run(self):

        global AllData
        global GlobalDictionary
        global ConnectionCompleted
        global ResetButtonClicked
        global Rounding_Valv_Check

        while ResetButtonClicked != 'Clicked':
            if (ConnectionCompleted == 'Success') and (len(GlobalDictionary) != 0):
                Arr = [round(GlobalDictionary[element][-1], Rounding_Valv_Check) for element in AllData]
                self.Status.emit(Arr)
            time.sleep(0.1)

class DisplayPackInfo(QThread):

    Status = pyqtSignal(str)

    def __init__(self):
        super().__init__()

    def run(self):
        global BatteryPackName
        global ScanCompleted
        global ResetButtonClicked

        while ResetButtonClicked != 'Clicked':
            if ScanCompleted == 'Success':
                if BatteryPackName != None:
                    self.Status.emit(BatteryPackName) 
                    break
            time.sleep(0.1)

class StartButton(threading.Thread):

    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):

        global PackType
        global FeedParms
        global SerialConnectionCycler
        global SerialConnectionVCU
        global GlobalDictionary
        global StopConditions
        global StoppageCurrentDuetoPowerCut
        global PortAssignedToCycler
        global StartButtonClicked
        global StartTaskEndsHere
        global StatusMessages
        global FolderForSavingAllFiles
        global ThresholdForTempBeforeStart
        global CyclingStatus
        global AllThreads
        global StartButtonAblingDisabling
        global ResetButtonAblingDisabling

        StartButtonAblingDisabling = 'Disable'
        ResetButtonAblingDisabling = 'Disable'

        if GlobalDictionary['FET_ON_OFF'][-1] == 1:
            StatusMessages.append('Started : Cycling')

            #Extracting DSG and CHG parms
            DSGParms = FeedParms[PackType][:3]
            CHGParms = FeedParms[PackType][3:]

            #Updating variable
            StartButtonClicked = 'Success'

            #DISCHARGING
            ##########################################################################################################################
            if (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                StatusMessages.append(f'Started : Discharging ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
                #Extracting Parms
                DsgCurrent = DSGParms[0]
                DsgVoltage = DSGParms[1]
                DsgPower = DSGParms[2]
                #Initialize Cycler
                SerialConnectionCycler.write(b"SYST:LOCK ON")
                time.sleep(0.1)
                #Setting parms
                Parms = f'SINK:CURR {DsgCurrent};VOLT {DsgVoltage};SINK:POW {DsgPower}'
                ByteParms = bytes(Parms, 'utf-8')
                SerialConnectionCycler.write(ByteParms)
                time.sleep(0.1)
                #Turing cycler ON
                while True:
                    SerialConnectionCycler.write(b"OUTP ON")
                    time.sleep(0.1)
                    SerialConnectionCycler.write(b"OUTPUT?")
                    if SerialConnectionCycler.readline().decode().strip() == 'ON':
                        time.sleep(0.1)
                        break
            while (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                if GlobalDictionary['DSG_Current'][-1] > 150:
                    CyclingStatus = 'Discharging'
                    break
                time.sleep(0.001)
            #Discharging Loop
            while (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                time.sleep(0.01)
                #Reading POWER STATUS FROM CYCLER
                SerialConnectionCycler.write(b"STAT:QUES:COND?")
                time.sleep(0.01)
                power_resistor_ques = SerialConnectionCycler.readline().decode().strip()
                power_status = int(bin(int(power_resistor_ques))[2:].zfill(16)[2])
                if (GlobalDictionary['DSG_Current'][-1] <= StopConditions['DischargingCurrent_'][0]) and (GlobalDictionary['DSG_Current'][-1] > 1) and (power_status == 0) and (GlobalDictionary['SOC'][-1] <= StopConditions['DischargingSOC_'][0]) and (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                    #Turning the cycler off
                    while True:
                        SerialConnectionCycler.write(b"OUTP OFF")
                        time.sleep(0.1)
                        SerialConnectionCycler.write(b"OUTPUT?")
                        if SerialConnectionCycler.readline().decode().strip() == 'OFF':
                            time.sleep(0.1)
                            break
                    break
                if power_status == 1:
                    StatusMessages.append(f'Power Cut ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
                    #Saving Previous Current
                    for FindingValvs in range(0,len(GlobalDictionary['DSG_Current'][::-1])):
                        if GlobalDictionary['DSG_Current'][::-1][FindingValvs] > 1:
                            StoppageCurrentDuetoPowerCut = GlobalDictionary['DSG_Current'][::-1][FindingValvs]
                            break
                    #Closing Cycler port
                    SerialConnectionCycler.close()
                    #Power Cut functions
                    time.sleep(20)
                    #Try to make connection with cycler
                    while True:
                        #Restoring Serial Connection Aand Power Status
                        try:
                            SerialConnectionCycler = serial.Serial(port=PortAssignedToCycler, baudrate = 9600, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE, bytesize=serial.EIGHTBITS, timeout=1)
                            SerialConnectionCycler.write(b"STAT:QUES:COND?")
                            PowerCutStatus = SerialConnectionCycler.readline().decode().strip()
                            PowerStatus = int(bin(int(PowerCutStatus))[2:].zfill(16)[2])
                            if PowerStatus == 0:
                                break
                            else:
                                SerialConnectionCycler.close()
                        except serial.SerialException:
                            time.sleep(0.1)
                    #Intializing cycler
                    SerialConnectionCycler.write(b"SYST:LOCK ON")
                    time.sleep(0.1)
                    #SETTING UP ALL PARAMETERS FOR DISCHARGING PHASE again after the power restoration
                    Parms = f'SINK:CURR {StoppageCurrentDuetoPowerCut};VOLT {DsgVoltage};SINK:POW {DsgPower}'
                    ByteParms = bytes(Parms, 'utf-8')
                    SerialConnectionCycler.write(ByteParms)
                    time.sleep(0.1)
                    #Turing Cycler 'ON'
                    while True:
                        SerialConnectionCycler.write(b"OUTP ON")
                        time.sleep(0.1)
                        SerialConnectionCycler.write(b"OUTPUT?")
                        if SerialConnectionCycler.readline().decode().strip() == 'ON':
                            time.sleep(0.1)
                            break
                    StatusMessages.append(f'Power Cut Restored ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
            if (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                StatusMessages.append(f'Completed : Discharging ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
            ##########################################################################################################################
            
            #REST
            ##########################################################################################################################
            if (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                StatusMessages.append(f'Started : Rest 1 ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
            #Turning the cycler off
            while True:
                SerialConnectionCycler.write(b"OUTP OFF")
                time.sleep(0.1)
                SerialConnectionCycler.write(b"OUTPUT?")
                if SerialConnectionCycler.readline().decode().strip() == 'OFF':
                    time.sleep(0.1)
                    break
            while (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                if (GlobalDictionary['DSG_Current'][-1] < 1) and (GlobalDictionary['CHG_Current'][-1] < 1):
                    CyclingStatus = 'Rest 1'
                    break
                time.sleep(0.001)
            #Initializing parms for Power cut
            PowerCutSignal = 0
            #Starting Rest Period and collecting data
            for i in range(StopConditions['Rest1'][0]):
                if (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                    time.sleep(1)
                    if PowerCutSignal == 0:
                        #Checking Power Status
                        SerialConnectionCycler.write(b"STAT:QUES:COND?")
                        PowerCutStatus = SerialConnectionCycler.readline().decode().strip()
                        PowerStatus = int(bin(int(PowerCutStatus))[2:].zfill(16)[2])
                        if PowerStatus == 1:
                            SerialConnectionCycler.close()
                            PowerCutSignal = PowerCutSignal + 1
                            StatusMessages.append(f'Power Cut ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
                else:
                    break
            #Waiting till Power returns
            if (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                if PowerCutSignal != 0:
                    while True:
                        try:
                            SerialConnectionCycler = serial.Serial(port=PortAssignedToCycler, baudrate = 9600, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE, bytesize=serial.EIGHTBITS, timeout=1)
                            SerialConnectionCycler.write(b"STAT:QUES:COND?")
                            PowerCutStatus = SerialConnectionCycler.readline().decode().strip()
                            if PowerCutStatus:
                                PowerStatus = int(bin(int(PowerCutStatus))[2:].zfill(16)[2])
                                if PowerStatus == 0:
                                    StatusMessages.append(f'Power Cut Restored ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
                                    break
                            SerialConnectionCycler.close()
                        except serial.SerialException:
                            time.sleep(0.1)
                StatusMessages.append(f'Completed : Rest 1 ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
            ##########################################################################################################################
            
            #CHARGING
            ##########################################################################################################################
            if (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                StatusMessages.append(f'Started : Charging ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
                #Parms
                ChgCurrent = CHGParms[0]
                ChgVoltage = CHGParms[1]
                ChgPower = CHGParms[2]
                #Initialize Cycler
                SerialConnectionCycler.write(b"SYST:LOCK ON")
                time.sleep(0.1) 
                #SETTING UP ALL PARAMETERS FOR CHARGING PHASE
                Parms = f'SOUR:CURRENT {ChgCurrent}A;VOLT {ChgVoltage};SOUR:POWER {ChgPower/1000}kW'
                ByteParms = bytes(Parms, 'utf-8')
                SerialConnectionCycler.write(ByteParms)
                time.sleep(0.1)
                #Turing Cycler 'ON'
                while True:
                    SerialConnectionCycler.write(b"OUTP ON")
                    time.sleep(0.1)
                    SerialConnectionCycler.write(b"OUTPUT?")
                    if SerialConnectionCycler.readline().decode().strip() == 'ON':
                        time.sleep(0.1)
                        break
            while (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                if GlobalDictionary['CHG_Current'][-1] > 100:
                    CyclingStatus = 'Charging'
                    break
                time.sleep(0.001)
            #Charging Loop
            while (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                time.sleep(0.01)
                #Reading POWER STATUS FROM CYCLER
                SerialConnectionCycler.write(b"STAT:QUES:COND?")
                time.sleep(0.01)
                power_resistor_ques = SerialConnectionCycler.readline().decode().strip()
                power_status = int(bin(int(power_resistor_ques))[2:].zfill(16)[2])
                if GlobalDictionary['SOC'][-1] >= StopConditions['Charging'][0]:
                    #Turning the cycler off
                    while True:
                        SerialConnectionCycler.write(b"OUTP OFF")
                        time.sleep(0.1)
                        SerialConnectionCycler.write(b"OUTPUT?")
                        if SerialConnectionCycler.readline().decode().strip() == 'OFF':
                            time.sleep(0.1)
                            break
                    break
                if power_status == 1:
                    StatusMessages.append(f'Power Cut ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
                    #Saving Previous Current
                    for FindingValvs in range(0,len(GlobalDictionary['CHG_Current'][::-1])):
                        if GlobalDictionary['CHG_Current'][::-1][FindingValvs] > 1:
                            StoppageCurrentDuetoPowerCut = GlobalDictionary['CHG_Current'][::-1][FindingValvs]
                            break
                    #Closing Cycler port
                    SerialConnectionCycler.close()
                    #Power Cut functions
                    time.sleep(20)
                    #Try to make connection with cycler
                    while True:
                        #Restoring Serial Connection Aand Power Status
                        try:
                            SerialConnectionCycler = serial.Serial(port=PortAssignedToCycler, baudrate = 9600, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE, bytesize=serial.EIGHTBITS, timeout=1)
                            SerialConnectionCycler.write(b"STAT:QUES:COND?")
                            PowerCutStatus = SerialConnectionCycler.readline().decode().strip()
                            PowerStatus = int(bin(int(PowerCutStatus))[2:].zfill(16)[2])
                            if PowerStatus == 0:
                                break
                            else:
                                SerialConnectionCycler.close()
                        except serial.SerialException:
                            time.sleep(0.1)
                    #Intializing cycler
                    SerialConnectionCycler.write(b"SYST:LOCK ON")
                    time.sleep(0.1)
                    #SETTING UP ALL PARAMETERS FOR DISCHARGING PHASE again after the power restoration
                    Parms = f'SOUR:CURRENT {StoppageCurrentDuetoPowerCut}A;VOLT {ChgVoltage};SOUR:POWER {ChgPower/1000}kW'
                    ByteParms = bytes(Parms, 'utf-8')
                    SerialConnectionCycler.write(ByteParms)
                    time.sleep(0.1)
                    #Turing Cycler 'ON'
                    while True:
                        SerialConnectionCycler.write(b"OUTP ON")
                        time.sleep(0.1)
                        SerialConnectionCycler.write(b"OUTPUT?")
                        if SerialConnectionCycler.readline().decode().strip() == 'ON':
                            time.sleep(0.1)
                            break
                    StatusMessages.append(f'Power Cut Restored ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
            if (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                StatusMessages.append(f'Completed : Charging ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
            ##########################################################################################################################

            #REST
            ##########################################################################################################################
            if (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                StatusMessages.append(f'Started : Rest 2 ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
            #Turning the cycler off
            while True:
                SerialConnectionCycler.write(b"OUTP OFF")
                time.sleep(0.1)
                SerialConnectionCycler.write(b"OUTPUT?")
                if SerialConnectionCycler.readline().decode().strip() == 'OFF':
                    time.sleep(0.1)
                    break
            while (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                if (GlobalDictionary['DSG_Current'][-1] < 1) and (GlobalDictionary['CHG_Current'][-1] < 1):
                    CyclingStatus = 'Rest 2'
                    break
                time.sleep(0.001)
            #Initializing parms for Power cut
            PowerCutSignal = 0
            #Starting Rest Period and collecting data
            for i in range(StopConditions['Rest2'][0]):
                if (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                    time.sleep(1)
                    if PowerCutSignal == 0:
                        #Checking Power Status
                        SerialConnectionCycler.write(b"STAT:QUES:COND?")
                        PowerCutStatus = SerialConnectionCycler.readline().decode().strip()
                        PowerStatus = int(bin(int(PowerCutStatus))[2:].zfill(16)[2])
                        if PowerStatus == 1:
                            SerialConnectionCycler.close()
                            PowerCutSignal = PowerCutSignal + 1
                            StatusMessages.append(f'Power Cut ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
                else:
                    break
            #Waiting till Power returns
            if (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                if PowerCutSignal != 0:
                    while True:
                        try:
                            SerialConnectionCycler = serial.Serial(port=PortAssignedToCycler, baudrate = 9600, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE, bytesize=serial.EIGHTBITS, timeout=1)
                            SerialConnectionCycler.write(b"STAT:QUES:COND?")
                            PowerCutStatus = SerialConnectionCycler.readline().decode().strip()
                            if PowerCutStatus:
                                PowerStatus = int(bin(int(PowerCutStatus))[2:].zfill(16)[2])
                                if PowerStatus == 0:
                                    StatusMessages.append(f'Power Cut Restored ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
                                    break
                            SerialConnectionCycler.close()
                        except:
                            time.sleep(0.1)
                StatusMessages.append(f'Completed : Rest 2 ({datetime.now().strftime("%Y-%m-%d %H:%M")})')
            ##########################################################################################################################
            CyclingStatus = 'None'
            if (GlobalDictionary['FET_ON_OFF'][-1] == 1):
                StartTaskEndsHere = 'Success'
                #EoLAnalysis
                self.EoLAnalysisThread = EoLAnalysis()
                self.EoLAnalysisThread.start()
                AllThreads.append(self.EoLAnalysisThread)
            else:
                self.PreEoLAnalysisThread = FETOFFEoLAnalysis()
                self.PreEoLAnalysisThread.start()
                AllThreads.append(self.PreEoLAnalysisThread)
        else:
            StartButtonAblingDisabling = 'Able'
            ResetButtonAblingDisabling = 'Able'
            StatusMessages.append('FET Status : 0. Turn ON the FET and Click START button again!')

class DisplayGraph(QThread):

    Status = pyqtSignal(int, float, float)

    def __init__(self):
        super().__init__()

    def run(self):
        global GlobalDictionary
        global StartButtonClicked
        global ResetButtonClicked

        x = 0
        while ResetButtonClicked != 'Clicked':
            if StartButtonClicked == 'Success':
                x = x + 1
                new_x = int(x)
                new_y1 = float(GlobalDictionary['Sum-of-cells'][-1])
                new_y2 = float(GlobalDictionary['DSG_Current'][-1] + GlobalDictionary['CHG_Current'][-1])
                self.Status.emit(new_x, new_y1, new_y2)
            time.sleep(0.5)

class EoLAnalysis(threading.Thread):

    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):

        global GlobalDictionary
        global FaultDetectionResults
        global StatusMessages
        global PackType
        global BatteryPackName
        global ThresholdForOvershoot
        global FolderForSavingAllFiles
        global RelativePathHR
        global RelativePathLR
        global CyclingResultsVariable
        global ResetButtonAblingDisabling

        try:
            #Updating Status
            StatusMessages.append('Started : EoL Analysis')
            #Converting GlobalDictionary to pandas dataframe
            LocalGlobalDataFrame = pd.DataFrame.from_dict(GlobalDictionary)
            #Fetching Usable 
            UsableLocalGlobalDataFrame = LocalGlobalDataFrame.iloc[LocalGlobalDataFrame[LocalGlobalDataFrame['Cycling Status'] == 'Discharging'].index[0]:LocalGlobalDataFrame[LocalGlobalDataFrame['Cycling Status'] == 'Rest 2'].iloc[-1:].index[0]+1].reset_index(drop = True)
            #Creating DSGData, RST1Data, CHGData, RST2Data
            DSGData = UsableLocalGlobalDataFrame[UsableLocalGlobalDataFrame['Cycling Status'] == 'Discharging'].reset_index(drop = True)
            RST1Data = UsableLocalGlobalDataFrame[UsableLocalGlobalDataFrame['Cycling Status'] == 'Rest 1'].reset_index(drop = True)
            CHGData = UsableLocalGlobalDataFrame[UsableLocalGlobalDataFrame['Cycling Status'] == 'Charging'].reset_index(drop = True)
            RST2Data = UsableLocalGlobalDataFrame[UsableLocalGlobalDataFrame['Cycling Status'] == 'Rest 2'].reset_index(drop = True)
            #Saving file
            particulate = ['TS1','TS2','TS3','TS4','TS5','TS6','TS7','TS8','TS9','TS10','TS11','TS12','FET Temp Front','BAT + ve Temp','BAT - ve Temp','Pack + ve Temp','TS0_FLT','TS13_FLT','FET_TEMP_REAR']
            for particle in particulate:
                dTdtArray_Paticle = UsableLocalGlobalDataFrame[particle].diff()/UsableLocalGlobalDataFrame['Millis'].diff()
                UsableLocalGlobalDataFrame[f'd({particle})/dt'] = dTdtArray_Paticle
            UsableLocalGlobalDataFrame.to_csv(os.path.join(FolderForSavingAllFiles, 'CyclingData.csv'), index = False)

            #EoL Algorithms
            #####################################################################################################################################################################
            #Solder Issue
            SolderIssueDataframe = UsableLocalGlobalDataFrame[['Cell1', 'Cell2', 'Cell3', 'Cell4', 'Cell5','Cell6', 'Cell7', 'Cell8','Cell9', 'Cell10','Cell11', 'Cell12', 'Cell13', 'Cell14', 'DSG_Current', 'CHG_Current']]
            SolderIssueSignal = self.SolderIssueDetection(SolderIssueDataframe)
            if SolderIssueSignal >= 1:
                DictToSave = {
                    'Parameters':['Solder Issue'],
                    'Result':['Fail']
                }
                FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            else:
                DictToSave = {
                    'Parameters':['Solder Issue'],
                    'Result':['Pass']
                }
                FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            #Weld Issue Charging
            WeldIssueDataFrame = UsableLocalGlobalDataFrame[['Cell1', 'Cell2', 'Cell3', 'Cell4', 'Cell5','Cell6', 'Cell7', 'Cell8','Cell9', 'Cell10','Cell11', 'Cell12', 'Cell13', 'Cell14', 'DSG_Current', 'CHG_Current', 'SOC']]
            WeldIssueSignal = self.WeldIssueDetection(WeldIssueDataFrame)
            if WeldIssueSignal >= 1:
                DictToSave = {
                    'Parameters':['Weld Issue'],
                    'Result':['Fail']
                }
                FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            else:
                DictToSave = {
                    'Parameters':['Weld Issue'],
                    'Result':['Pass']
                }
                FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            #Temperature Fluctuation Issue 
            TemperatureFluctuationDataFrame = UsableLocalGlobalDataFrame[['Millis','TS1','TS2','TS3','TS4','TS5','TS6','TS7','TS8','TS9','TS10','TS11','TS12','TS0_FLT','TS13_FLT']]
            TemperatureFluctuationSignal = self.TemperatureFluctuationDetection(TemperatureFluctuationDataFrame)
            if TemperatureFluctuationSignal >= 1:
                DictToSave = {
                    'Parameters':['Temperature Fluctuation'],
                    'Result':['Fail']
                }
                FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            else:
                DictToSave = {
                    'Parameters':['Temperature Fluctuation'],
                    'Result':['Pass']
                }
                FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            #Thermister Open Issue
            ThermisterOpenDataFrame = UsableLocalGlobalDataFrame[['TS1','TS2','TS3','TS4','TS5','TS6','TS7','TS8','TS9','TS10','TS11','TS12','TS0_FLT','TS13_FLT']]
            ThermisterOpenSignal = self.ThermisterOpenIssueDetection(ThermisterOpenDataFrame)
            if ThermisterOpenSignal >= 1:
                DictToSave = {
                    'Parameters':['Thermister Open'],
                    'Result':['Fail']
                }
                FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            else:
                DictToSave = {
                    'Parameters':['Thermister Open'],
                    'Result':['Pass']
                }
                FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            #Delta Temperature Issue
            DeltaTempDataFrame = UsableLocalGlobalDataFrame[['TS1','TS2','TS3','TS4','TS5','TS6','TS7','TS8','TS9','TS10','TS11','TS12','TS0_FLT','TS13_FLT']]
            DeltaTempSignal = self.DeltaTemperatureIssueDetection(DeltaTempDataFrame)
            if DeltaTempSignal >= 1:
                DictToSave = {
                    'Parameters':['Delta Temperature'],
                    'Result':['Fail']
                }
                FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            else:
                DictToSave = {
                    'Parameters':['Delta Temperature'],
                    'Result':['Pass']
                }
                FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            #####################################################################################################################################################################

            #EoL Summary Sheet
            #####################################################################################################################################################################
            #Checking LimitsData
            if PackType == 'HR':
                LimitsData = pd.read_csv(RelativePathHR)
            elif PackType == 'LR':
                LimitsData = pd.read_csv(RelativePathLR)
            #concatinating DSG, RST1 and CHG, RST2
            DSGRST1Data = pd.concat([DSGData, RST1Data], ignore_index = True).reset_index(drop = True)
            CHGRST2Data = pd.concat([CHGData, RST2Data], ignore_index = True).reset_index(drop = True) 
            #Logic
            for SingleParm in np.array(LimitsData['Parameters']):
                if SingleParm in ['TS1', 'TS2', 'TS3', 'TS4', 'TS5', 'TS6', 'TS7', 'TS8', 'TS9','TS10', 'TS11', 'TS12', 'TS0_FLT', 'TS13_FLT', 'FET Temp Front','BAT + ve Temp', 'BAT - ve Temp', 'Pack + ve Temp','FET_TEMP_REAR', 'BAL_RES_TEMP']:
                    Min_Limit = LimitsData[LimitsData['Parameters'] == SingleParm]['Min'][LimitsData[LimitsData['Parameters'] == SingleParm].index[0]]
                    Max_DSG_Limit = LimitsData[LimitsData['Parameters'] == SingleParm]['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == SingleParm].index[0]]
                    Max_CHG_Limit = LimitsData[LimitsData['Parameters'] == SingleParm]['Charging Phase & 15 min rest'][LimitsData[LimitsData['Parameters'] == SingleParm].index[0]]
                    CheckArrDSGRST1 = []
                    for LimitCross in np.array(DSGRST1Data[SingleParm]):
                        if (LimitCross < Min_Limit) or (LimitCross > Max_DSG_Limit):
                            CheckArrDSGRST1.append('Fail')
                            break
                        else:
                            CheckArrDSGRST1.append('Pass')
                    CheckArrCHGRST2 = []
                    for LimitCross in np.array(CHGRST2Data[SingleParm]):
                        if (LimitCross < Min_Limit) or (LimitCross > Max_DSG_Limit):
                            CheckArrCHGRST2.append('Fail')
                            break
                        else:
                            CheckArrCHGRST2.append('Pass')
                    TotalArrIS = CheckArrDSGRST1 + CheckArrCHGRST2
                    if 'Fail' in TotalArrIS:
                        DictToSave = {
                            'Parameters':[SingleParm],
                            'Result':['Fail']
                        }
                        FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
                    else:
                        DictToSave = {
                            'Parameters':[SingleParm],
                            'Result':['Pass']
                        }
                        FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
                elif SingleParm in ['dTS1/dt', 'dTS2/dt', 'dTS3/dt','dTS4/dt', 'dTS5/dt', 'dTS6/dt', 'dTS7/dt', 'dTS8/dt', 'dTS9/dt','dTS10/dt', 'dTS11/dt', 'dTS12/dt', 'dTS0_FLT/dt', 'dTS13_FLT/dt','d(FET Temp Front)/dt', 'd(BAT + ve Temp)/dt','d(BAT - ve Temp)/dt', 'd(Pack + ve Temp)/dt','d(FET_TEMP_REAR)/dt', 'd(BAL_RES_TEMP)/dt']:
                    Min_Limit = LimitsData[LimitsData['Parameters'] == SingleParm]['Min'][LimitsData[LimitsData['Parameters'] == SingleParm].index[0]]
                    Max_DSG_Limit = LimitsData[LimitsData['Parameters'] == SingleParm]['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == SingleParm].index[0]]
                    if SingleParm == 'dTS1/dt':
                        TurnDownParm = 'TS1'
                    if SingleParm == 'dTS2/dt':
                        TurnDownParm = 'TS2'
                    if SingleParm == 'dTS3/dt':
                        TurnDownParm = 'TS3'
                    if SingleParm == 'dTS4/dt':
                        TurnDownParm = 'TS4'
                    if SingleParm == 'dTS5/dt':
                        TurnDownParm = 'TS5'
                    if SingleParm == 'dTS6/dt':
                        TurnDownParm = 'TS6'
                    if SingleParm == 'dTS7/dt':
                        TurnDownParm = 'TS7'
                    if SingleParm == 'dTS8/dt':
                        TurnDownParm = 'TS8'
                    if SingleParm == 'dTS9/dt':
                        TurnDownParm = 'TS9'
                    if SingleParm == 'dTS10/dt':
                        TurnDownParm = 'TS10'
                    if SingleParm == 'dTS11/dt':
                        TurnDownParm = 'TS11'
                    if SingleParm == 'dTS12/dt':
                        TurnDownParm = 'TS12'
                    if SingleParm == 'dTS0_FLT/dt':
                        TurnDownParm = 'TS0_FLT'
                    if SingleParm == 'dTS13_FLT/dt':
                        TurnDownParm = 'TS13_FLT'
                    if SingleParm == 'd(FET Temp Front)/dt':
                        TurnDownParm = 'FET Temp Front'
                    if SingleParm == 'd(BAT + ve Temp)/dt':
                        TurnDownParm = 'BAT + ve Temp'
                    if SingleParm == 'd(BAT - ve Temp)/dt':
                        TurnDownParm = 'BAT + ve Temp'
                    if SingleParm == 'd(Pack + ve Temp)/dt':
                        TurnDownParm = 'Pack + ve Temp'
                    if SingleParm == 'd(FET_TEMP_REAR)/dt':
                        TurnDownParm = 'FET_TEMP_REAR'
                    if SingleParm == 'd(BAL_RES_TEMP)/dt':
                        TurnDownParm = 'BAL_RES_TEMP'
                    dTArray = np.diff(np.array(UsableLocalGlobalDataFrame[TurnDownParm]))
                    dtArray = np.diff(np.array(UsableLocalGlobalDataFrame['Millis']))
                    dTdtArray = dTArray/dtArray
                    TotalCheckDone = []
                    for LimitCross in dTdtArray:
                        if (LimitCross < Min_Limit) or (LimitCross > Max_DSG_Limit):
                            TotalCheckDone.append('Fail')
                            break
                        else:
                            TotalCheckDone.append('Pass')
                    if 'Fail' in TotalCheckDone:
                        DictToSave = {
                            'Parameters':[SingleParm],
                            'Result':['Fail']
                        }
                        FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
                    else:
                        DictToSave = {
                            'Parameters':[SingleParm],
                            'Result':['Pass']
                        }
                        FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
                elif SingleParm == 'SoC':
                    Min_Limit = LimitsData[LimitsData['Parameters'] == SingleParm]['Min'][LimitsData[LimitsData['Parameters'] == SingleParm].index[0]]
                    Max_DSG_Limit = LimitsData[LimitsData['Parameters'] == SingleParm]['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == SingleParm].index[0]]
                    CheckArrDSGRST1 = []
                    for LimitCross in np.array(RST2Data['SOC']):
                        if (LimitCross < Min_Limit) or (LimitCross > Max_DSG_Limit):
                            CheckArrDSGRST1.append('Fail')
                            break
                        else:
                            CheckArrDSGRST1.append('Pass')
                    if 'Fail' in CheckArrDSGRST1:
                        DictToSave = {
                            'Parameters':[SingleParm],
                            'Result':['Fail']
                        }
                        FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
                    else:
                        DictToSave = {
                            'Parameters':[SingleParm],
                            'Result':['Pass']
                        }
                        FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
                elif SingleParm == 'Delta Voltage':
                    MinValvDeltaVoltage = LimitsData[LimitsData['Parameters'] == 'Delta Voltage']['Min'][LimitsData[LimitsData['Parameters'] == 'Delta Voltage'].index[0]]
                    DSGRSTPeriod = LimitsData[LimitsData['Parameters'] == 'Delta Voltage']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == 'Delta Voltage'].index[0]]
                    CHGRSTPeriod = LimitsData[LimitsData['Parameters'] == 'Delta Voltage']['Charging Phase & 15 min rest'][LimitsData[LimitsData['Parameters'] == 'Delta Voltage'].index[0]]
                    CHGStartOvershoot = LimitsData[LimitsData['Parameters'] == 'Delta Voltage']['chg_start_overshoot'][LimitsData[LimitsData['Parameters'] == 'Delta Voltage'].index[0]]
                    CHGEndOvershoot = LimitsData[LimitsData['Parameters'] == 'Delta Voltage']['chg_end_overshoot'][LimitsData[LimitsData['Parameters'] == 'Delta Voltage'].index[0]]
                    PostCHGIdle = LimitsData[LimitsData['Parameters'] == 'Delta Voltage']['post_chg_idle'][LimitsData[LimitsData['Parameters'] == 'Delta Voltage'].index[0]]
                    LowerBoundDeltaVoltage = [MinValvDeltaVoltage]*len(UsableLocalGlobalDataFrame['Cell Delta Volt'])
                    FirstBoundDSGRSTPeriod = (
                        [DSGRSTPeriod] * (len(DSGData['SOC']) + len(RST1Data['SOC']) - ThresholdForOvershoot) + 
                        [CHGStartOvershoot] * (2*ThresholdForOvershoot) +
                        [CHGRSTPeriod] * (len(CHGData['SOC']) - (2*ThresholdForOvershoot)) + 
                        [CHGEndOvershoot] * (2*ThresholdForOvershoot) +
                        [PostCHGIdle] * (len(RST2Data['SOC']) - (2*ThresholdForOvershoot))
                    )
                    DeltaVoltageData_ = np.array(UsableLocalGlobalDataFrame[['Cell1', 'Cell2', 'Cell3',
                                'Cell4', 'Cell5', 'Cell6', 'Cell7', 'Cell8', 'Cell9', 'Cell10',
                                'Cell11', 'Cell12', 'Cell13', 'Cell14']].max(axis = 1) - UsableLocalGlobalDataFrame[['Cell1', 'Cell2', 'Cell3',
                                'Cell4', 'Cell5', 'Cell6', 'Cell7', 'Cell8', 'Cell9', 'Cell10',
                                'Cell11', 'Cell12', 'Cell13', 'Cell14']].min(axis = 1))
                    MyArray = []
                    for value, min_val, max_val in zip(DeltaVoltageData_, LowerBoundDeltaVoltage, FirstBoundDSGRSTPeriod):
                        if (value < min_val) or (value > max_val):
                            MyArray.append('Fail')
                        else:
                            MyArray.append('Pass')
                    if 'Fail' in MyArray:
                        DictToSave = {
                            'Parameters':[SingleParm],
                            'Result':['Fail']
                        }
                        FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
                    else:
                        DictToSave = {
                            'Parameters':[SingleParm],
                            'Result':['Pass']
                        }
                        FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
                elif SingleParm == 'delta_temperature':
                    Min_Limit = LimitsData[LimitsData['Parameters'] == SingleParm]['Min'][LimitsData[LimitsData['Parameters'] == SingleParm].index[0]]
                    Max_DSG_Limit = LimitsData[LimitsData['Parameters'] == SingleParm]['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == SingleParm].index[0]]
                    DeltaTemperatureData = np.array(UsableLocalGlobalDataFrame[['TS1', 'TS2', 'TS3', 'TS4', 'TS5',
                            'TS6', 'TS7', 'TS8', 'TS9', 'TS10', 'TS11', 'TS12','TS0_FLT',
                            'TS13_FLT']].max(axis = 1) - UsableLocalGlobalDataFrame[['TS1', 'TS2', 'TS3', 'TS4', 'TS5',
                            'TS6', 'TS7', 'TS8', 'TS9', 'TS10', 'TS11', 'TS12','TS0_FLT',
                            'TS13_FLT']].min(axis = 1))
                    CheckArrDSGRST1 = []
                    for iterate_Valvs in np.array(DeltaTemperatureData):
                        if (iterate_Valvs < Min_Limit) or (iterate_Valvs > Max_DSG_Limit):
                            CheckArrDSGRST1.append('Fail')
                            break
                        else:
                            CheckArrDSGRST1.append('Pass')
                    if 'Fail' in CheckArrDSGRST1:
                        DictToSave = {
                            'Parameters':[SingleParm],
                            'Result':['Fail']
                        }
                        FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
                    else:
                        DictToSave = {
                            'Parameters':[SingleParm],
                            'Result':['Pass']
                        }
                        FaultDetectionResults = pd.concat([FaultDetectionResults, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            
            def _color_red_or_green(val):
                color = 'red' if val == 'Fail' else 'green'
                return 'color: %s' % color
            try:
                # Apply the color formatting to the DataFrame
                styled_data = FaultDetectionResults.style.applymap(_color_red_or_green, subset='Result')
                styled_data.to_excel(os.path.join(FolderForSavingAllFiles, 'Summary.xlsx'), index = False)
            except:
                FaultDetectionResults.to_excel(os.path.join(FolderForSavingAllFiles, 'Summary.xlsx'), index = False)
            #####################################################################################################################################################################
            
            #EoL Plots
            #####################################################################################################################################################################
            #Plotting TS1, TS2, TS3, TS4
            fig = plt.figure(figsize=(20,15))
            moment = 1
            for SensorNum in range(1, 5):
                #Primary subplot
                #LimitsData Setting
                MinValv = LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}']['Min'][LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}'].index]
                DSGRESTMaxValv = LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}'].index]
                CHGRESTMaxValv = LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}']['Charging Phase & 15 min rest'][LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}'].index]
                #Setting Upper and Lower Limit
                UpperDSGRSTLimitArray = np.full((len(DSGData[f'TS{SensorNum}']) + len(RST1Data[f'TS{SensorNum}'])), DSGRESTMaxValv)
                UpperCHGRSTLimitArray = np.full((len(CHGData[f'TS{SensorNum}']) + len(RST2Data[f'TS{SensorNum}'])), CHGRESTMaxValv)
                LowerLimitArray = np.full(len(DSGData[f'TS{SensorNum}']) + len(RST1Data[f'TS{SensorNum}']) + len(CHGData[f'TS{SensorNum}']) + len(RST2Data[f'TS{SensorNum}']), MinValv)
                #Plot
                ax1 = fig.add_subplot(2,2,moment)
                moment = moment + 1
                ax1.plot(UsableLocalGlobalDataFrame[f'TS{SensorNum}'],color='g')
                ax1.plot(np.concatenate((UpperDSGRSTLimitArray, UpperCHGRSTLimitArray)),linestyle='dashed')
                ax1.plot(LowerLimitArray,linestyle='dashed')
                ax1.set_ylabel('Temperature (°C)')
                ax1.set_title(f'TS{SensorNum}')
                ax1.legend((f'TS{SensorNum}',f'TS{SensorNum}_Min',f'TS{SensorNum}_Max'),loc=2)
                ax1.set_ylim(0,60)
                ax1.grid()
                #Secondary subplot
                #Calculating dT/dt
                dTArray = np.diff(np.array(UsableLocalGlobalDataFrame[f'TS{SensorNum}']))
                dtArray = np.diff(np.array(UsableLocalGlobalDataFrame['Millis']))
                dTdtArray = dTArray/dtArray
                #Finding Limits
                MinValvSec = LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt']['Min'][LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt'].index[0]]
                MaxValvSec = LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt'].index[0]]
                ax1_twin = ax1.twinx()
                ax1_twin.plot(dTdtArray, color='m')
                ax1_twin.axhline(MaxValvSec,linestyle='dashed', linewidth=2,color='orange')
                ax1_twin.axhline(MinValvSec,linestyle='dashed', linewidth=2)
                ax1_twin.set_ylabel('dT/dt')
                ax1_twin.legend((f'dTS{SensorNum}/dt','dT/dt upperlimit','dT/dt lowerlimit'),loc=9)
                ax1_twin.set_ylim(-0.05,0.5)

            plt.suptitle(BatteryPackName)
            plt.savefig(os.path.join(FolderForSavingAllFiles, 'Cell Temperatures plot 1.png'))

            #Plotting TS5, TS6, TS7, TS8
            fig = plt.figure(figsize=(20,15))
            moment = 1
            for SensorNum in range(5, 9):
                #Primary subplot
                #LimitsData Setting
                MinValv = LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}']['Min'][LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}'].index]
                DSGRESTMaxValv = LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}'].index]
                CHGRESTMaxValv = LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}']['Charging Phase & 15 min rest'][LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}'].index]
                #Setting Upper and Lower Limit
                UpperDSGRSTLimitArray = np.full((len(DSGData[f'TS{SensorNum}']) + len(RST1Data[f'TS{SensorNum}'])), DSGRESTMaxValv)
                UpperCHGRSTLimitArray = np.full((len(CHGData[f'TS{SensorNum}']) + len(RST2Data[f'TS{SensorNum}'])), CHGRESTMaxValv)
                LowerLimitArray = np.full(len(DSGData[f'TS{SensorNum}']) + len(RST1Data[f'TS{SensorNum}']) + len(CHGData[f'TS{SensorNum}']) + len(RST2Data[f'TS{SensorNum}']), MinValv)
                #Plot
                ax1 = fig.add_subplot(2,2,moment)
                moment = moment + 1
                ax1.plot(UsableLocalGlobalDataFrame[f'TS{SensorNum}'],color='g')
                ax1.plot(np.concatenate((UpperDSGRSTLimitArray, UpperCHGRSTLimitArray)),linestyle='dashed')
                ax1.plot(LowerLimitArray,linestyle='dashed')
                ax1.set_ylabel('Temperature (°C)')
                ax1.set_title(f'TS{SensorNum}')
                ax1.legend((f'TS{SensorNum}',f'TS{SensorNum}_Min',f'TS{SensorNum}_Max'),loc=2)
                ax1.set_ylim(0,60)
                ax1.grid()
                #Secondary subplot
                #Calculating dT/dt
                dTArray = np.diff(np.array(UsableLocalGlobalDataFrame[f'TS{SensorNum}']))
                dtArray = np.diff(np.array(UsableLocalGlobalDataFrame['Millis']))
                dTdtArray = dTArray/dtArray
                #Finding Limits
                MinValvSec = LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt']['Min'][LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt'].index[0]]
                MaxValvSec = LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt'].index[0]]
                ax1_twin = ax1.twinx()
                ax1_twin.plot(dTdtArray, color='m')
                ax1_twin.axhline(MaxValvSec,linestyle='dashed', linewidth=2,color='orange')
                ax1_twin.axhline(MinValvSec,linestyle='dashed', linewidth=2)
                ax1_twin.set_ylabel('dT/dt')
                ax1_twin.legend((f'dTS{SensorNum}/dt','dT/dt upperlimit','dT/dt lowerlimit'),loc=9)
                ax1_twin.set_ylim(-0.05,0.5)

            plt.suptitle(BatteryPackName)
            plt.savefig(os.path.join(FolderForSavingAllFiles, 'Cell Temperatures plot 2.png'))

            #Plotting TS9, TS10, TS11, TS12
            fig = plt.figure(figsize=(20,15))
            moment = 1
            for SensorNum in range(9, 13):
                #Primary subplot
                #LimitsData Setting
                MinValv = LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}']['Min'][LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}'].index]
                DSGRESTMaxValv = LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}'].index]
                CHGRESTMaxValv = LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}']['Charging Phase & 15 min rest'][LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}'].index]
                #Setting Upper and Lower Limit
                UpperDSGRSTLimitArray = np.full((len(DSGData[f'TS{SensorNum}']) + len(RST1Data[f'TS{SensorNum}'])), DSGRESTMaxValv)
                UpperCHGRSTLimitArray = np.full((len(CHGData[f'TS{SensorNum}']) + len(RST2Data[f'TS{SensorNum}'])), CHGRESTMaxValv)
                LowerLimitArray = np.full(len(DSGData[f'TS{SensorNum}']) + len(RST1Data[f'TS{SensorNum}']) + len(CHGData[f'TS{SensorNum}']) + len(RST2Data[f'TS{SensorNum}']), MinValv)
                #Plot
                ax1 = fig.add_subplot(2,2,moment)
                moment = moment + 1
                ax1.plot(UsableLocalGlobalDataFrame[f'TS{SensorNum}'],color='g')
                ax1.plot(np.concatenate((UpperDSGRSTLimitArray, UpperCHGRSTLimitArray)),linestyle='dashed')
                ax1.plot(LowerLimitArray,linestyle='dashed')
                ax1.set_ylabel('Temperature (°C)')
                ax1.set_title(f'TS{SensorNum}')
                ax1.legend((f'TS{SensorNum}',f'TS{SensorNum}_Min',f'TS{SensorNum}_Max'),loc=2)
                ax1.set_ylim(0,60)
                ax1.grid()
                #Secondary subplot
                #Calculating dT/dt
                dTArray = np.diff(np.array(UsableLocalGlobalDataFrame[f'TS{SensorNum}']))
                dtArray = np.diff(np.array(UsableLocalGlobalDataFrame['Millis']))
                dTdtArray = dTArray/dtArray
                #Finding Limits
                MinValvSec = LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt']['Min'][LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt'].index[0]]
                MaxValvSec = LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt'].index[0]]
                ax1_twin = ax1.twinx()
                ax1_twin.plot(dTdtArray, color='m')
                ax1_twin.axhline(MaxValvSec,linestyle='dashed', linewidth=2,color='orange')
                ax1_twin.axhline(MinValvSec,linestyle='dashed', linewidth=2)
                ax1_twin.set_ylabel('dT/dt')
                ax1_twin.legend((f'dTS{SensorNum}/dt','dT/dt upperlimit','dT/dt lowerlimit'),loc=9)
                ax1_twin.set_ylim(-0.05,0.5)

            plt.suptitle(BatteryPackName)
            plt.savefig(os.path.join(FolderForSavingAllFiles, 'Cell Temperatures plot 3.png'))

            #Plotting TS0_FLT, TS13_FLT
            fig = plt.figure(figsize=(20,15))
            moment = 1
            for SensorNum in ['0_FLT', '13_FLT']:
                #Primary subplot
                #LimitsData Setting
                MinValv = LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}']['Min'][LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}'].index]
                DSGRESTMaxValv = LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}'].index]
                CHGRESTMaxValv = LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}']['Charging Phase & 15 min rest'][LimitsData[LimitsData['Parameters'] == f'TS{SensorNum}'].index]
                #Setting Upper and Lower Limit
                UpperDSGRSTLimitArray = np.full((len(DSGData[f'TS{SensorNum}']) + len(RST1Data[f'TS{SensorNum}'])), DSGRESTMaxValv)
                UpperCHGRSTLimitArray = np.full((len(CHGData[f'TS{SensorNum}']) + len(RST2Data[f'TS{SensorNum}'])), CHGRESTMaxValv)
                LowerLimitArray = np.full(len(DSGData[f'TS{SensorNum}']) + len(RST1Data[f'TS{SensorNum}']) + len(CHGData[f'TS{SensorNum}']) + len(RST2Data[f'TS{SensorNum}']), MinValv)
                #Plot
                ax1 = fig.add_subplot(2,2,moment)
                moment = moment + 1
                ax1.plot(UsableLocalGlobalDataFrame[f'TS{SensorNum}'],color='g')
                ax1.plot(np.concatenate((UpperDSGRSTLimitArray, UpperCHGRSTLimitArray)),linestyle='dashed')
                ax1.plot(LowerLimitArray,linestyle='dashed')
                ax1.set_ylabel('Temperature (°C)')
                ax1.set_title(f'TS{SensorNum}')
                ax1.legend((f'TS{SensorNum}',f'TS{SensorNum}_Min',f'TS{SensorNum}_Max'),loc=2)
                ax1.set_ylim(0,60)
                ax1.grid()
                #Secondary subplot
                #Calculating dT/dt
                dTArray = np.diff(np.array(UsableLocalGlobalDataFrame[f'TS{SensorNum}']))
                dtArray = np.diff(np.array(UsableLocalGlobalDataFrame['Millis']))
                dTdtArray = dTArray/dtArray
                #Finding Limits
                MinValvSec = LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt']['Min'][LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt'].index[0]]
                MaxValvSec = LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == f'dTS{SensorNum}/dt'].index[0]]
                ax1_twin = ax1.twinx()
                ax1_twin.plot(dTdtArray, color='m')
                ax1_twin.axhline(MaxValvSec,linestyle='dashed', linewidth=2,color='orange')
                ax1_twin.axhline(MinValvSec,linestyle='dashed', linewidth=2)
                ax1_twin.set_ylabel('dT/dt')
                ax1_twin.legend((f'dTS{SensorNum}/dt','dT/dt upperlimit','dT/dt lowerlimit'),loc=9)
                ax1_twin.set_ylim(-0.05,0.5)

            plt.suptitle(BatteryPackName)
            plt.savefig(os.path.join(FolderForSavingAllFiles, 'Cell Temperatures plot 4.png'))

            #Discharge Current
            Discharge_current_plot = plt.figure(figsize=(10,6))
            plt.plot(UsableLocalGlobalDataFrame['DSG_Current'], label = 'DSG Current')
            plt.legend()
            plt.grid()
            plt.title('Discharge Current')
            plt.ylabel('Current (A)')

            plt.savefig(os.path.join(FolderForSavingAllFiles, 'Discharge_current_plot.png'))

            #Charge Current
            Charge_current_plot=plt.figure(figsize=(10,6))
            plt.plot(UsableLocalGlobalDataFrame['CHG_Current'], label = 'CHG Current')
            plt.legend()
            plt.grid()
            plt.title('Charge Current')
            plt.ylabel('Current (A)')

            plt.savefig(os.path.join(FolderForSavingAllFiles, 'Charge_current_plot.png'))

            #Cell1-Cell14
            Cell_voltage_plot=plt.figure(figsize=(10,6))
            for CellNumber in range(1, 15):
                plt.plot(UsableLocalGlobalDataFrame[f'Cell{CellNumber}'], label = f'Cell{CellNumber}')
            plt.legend()
            plt.grid()
            plt.title('Cell Voltages')
            plt.ylabel('Voltage (V)')

            plt.savefig(os.path.join(FolderForSavingAllFiles, 'Cell_voltage_plot.png'))

            #FET Temp Front, BAT + ve Temp, BAT - ve Temp, Pack + ve Temp
            fig = plt.figure(figsize=(20,15))
            moment = 1
            for SensorNum in ['FET Temp Front', 'BAT + ve Temp', 'BAT - ve Temp', 'Pack + ve Temp']:
                #Primary subplot
                #LimitsData Setting
                MinValv = LimitsData[LimitsData['Parameters'] == SensorNum]['Min'][LimitsData[LimitsData['Parameters'] == SensorNum].index]
                DSGRESTMaxValv = LimitsData[LimitsData['Parameters'] == SensorNum]['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == SensorNum].index]
                CHGRESTMaxValv = LimitsData[LimitsData['Parameters'] == SensorNum]['Charging Phase & 15 min rest'][LimitsData[LimitsData['Parameters'] == SensorNum].index]
                #Setting Upper and Lower Limit
                UpperDSGRSTLimitArray = np.full((len(DSGData[SensorNum]) + len(RST1Data[SensorNum])), DSGRESTMaxValv)
                UpperCHGRSTLimitArray = np.full((len(CHGData[SensorNum]) + len(RST2Data[SensorNum])), CHGRESTMaxValv)
                LowerLimitArray = np.full(len(DSGData[SensorNum]) + len(RST1Data[SensorNum]) + len(CHGData[SensorNum]) + len(RST2Data[SensorNum]), MinValv)
                #Plot
                ax1 = fig.add_subplot(2,2,moment)
                moment = moment + 1
                ax1.plot(UsableLocalGlobalDataFrame[SensorNum],color='g')
                ax1.plot(np.concatenate((UpperDSGRSTLimitArray, UpperCHGRSTLimitArray)),linestyle='dashed')
                ax1.plot(LowerLimitArray,linestyle='dashed')
                ax1.set_ylabel('Temperature (°C)')
                ax1.set_title(SensorNum)
                ax1.legend((SensorNum,f'{SensorNum}_Min',f'{SensorNum}_Max'),loc=2)
                ax1.set_ylim(0,60)
                ax1.grid()
                #Secondary subplot
                #Calculating dT/dt
                dTArray = np.diff(np.array(UsableLocalGlobalDataFrame[SensorNum]))
                dtArray = np.diff(np.array(UsableLocalGlobalDataFrame['Millis']))
                dTdtArray = dTArray/dtArray
                #Finding Limits
                MinValvSec = LimitsData[LimitsData['Parameters'] == f'd({SensorNum})/dt']['Min'][LimitsData[LimitsData['Parameters'] == f'd({SensorNum})/dt'].index[0]]
                MaxValvSec = LimitsData[LimitsData['Parameters'] == f'd({SensorNum})/dt']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == f'd({SensorNum})/dt'].index[0]]
                ax1_twin = ax1.twinx()
                ax1_twin.plot(dTdtArray, color='m')
                ax1_twin.axhline(MaxValvSec,linestyle='dashed', linewidth=2,color='orange')
                ax1_twin.axhline(MinValvSec,linestyle='dashed', linewidth=2)
                ax1_twin.set_ylabel('dT/dt')
                ax1_twin.legend((f'd({SensorNum})/dt','dT/dt upperlimit','dT/dt lowerlimit'),loc=9)
                ax1_twin.set_ylim(-0.05,0.5)

            plt.suptitle(BatteryPackName)

            plt.savefig(os.path.join(FolderForSavingAllFiles, 'BMS Temperatures plot 1.png'))

            #FET_TEMP_REAR, BAL_RES_TEMP
            fig = plt.figure(figsize=(20,15))
            moment = 1
            for SensorNum in ['FET_TEMP_REAR', 'BAL_RES_TEMP']:
                #Primary subplot
                #LimitsData Setting
                MinValv = LimitsData[LimitsData['Parameters'] == SensorNum]['Min'][LimitsData[LimitsData['Parameters'] == SensorNum].index]
                DSGRESTMaxValv = LimitsData[LimitsData['Parameters'] == SensorNum]['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == SensorNum].index]
                CHGRESTMaxValv = LimitsData[LimitsData['Parameters'] == SensorNum]['Charging Phase & 15 min rest'][LimitsData[LimitsData['Parameters'] == SensorNum].index]
                #Setting Upper and Lower Limit
                UpperDSGRSTLimitArray = np.full((len(DSGData[SensorNum]) + len(RST1Data[SensorNum])), DSGRESTMaxValv)
                UpperCHGRSTLimitArray = np.full((len(CHGData[SensorNum]) + len(RST2Data[SensorNum])), CHGRESTMaxValv)
                LowerLimitArray = np.full(len(DSGData[SensorNum]) + len(RST1Data[SensorNum]) + len(CHGData[SensorNum]) + len(RST2Data[SensorNum]), MinValv)
                #Plot
                ax1 = fig.add_subplot(2,2,moment)
                moment = moment + 1
                ax1.plot(UsableLocalGlobalDataFrame[SensorNum],color='g')
                ax1.plot(np.concatenate((UpperDSGRSTLimitArray, UpperCHGRSTLimitArray)),linestyle='dashed')
                ax1.plot(LowerLimitArray,linestyle='dashed')
                ax1.set_ylabel('Temperature (°C)')
                ax1.set_title(SensorNum)
                ax1.legend((SensorNum,f'{SensorNum}_Min',f'{SensorNum}_Max'),loc=2)
                ax1.set_ylim(0,60)
                ax1.grid()
                #Secondary subplot
                #Calculating dT/dt
                dTArray = np.diff(np.array(UsableLocalGlobalDataFrame[SensorNum]))
                dtArray = np.diff(np.array(UsableLocalGlobalDataFrame['Millis']))
                dTdtArray = dTArray/dtArray
                #Finding Limits
                MinValvSec = LimitsData[LimitsData['Parameters'] == f'd({SensorNum})/dt']['Min'][LimitsData[LimitsData['Parameters'] == f'd({SensorNum})/dt'].index[0]]
                MaxValvSec = LimitsData[LimitsData['Parameters'] == f'd({SensorNum})/dt']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == f'd({SensorNum})/dt'].index[0]]
                ax1_twin = ax1.twinx()
                ax1_twin.plot(dTdtArray, color='m')
                ax1_twin.axhline(MaxValvSec,linestyle='dashed', linewidth=2,color='orange')
                ax1_twin.axhline(MinValvSec,linestyle='dashed', linewidth=2)
                ax1_twin.set_ylabel(f'd({SensorNum})/dt')
                ax1_twin.legend((f'd({SensorNum})/dt','dT/dt upperlimit','dT/dt lowerlimit'),loc=9)
                ax1_twin.set_ylim(-0.09,0.6)

            plt.suptitle(BatteryPackName)

            plt.savefig(os.path.join(FolderForSavingAllFiles, 'BMS Temperatures plot 2.png'))

            #Delta Temperature
            def fetch_and_calculate(dictionary, keys):
                dictionary = pd.DataFrame.to_dict(dictionary)
                selected_dict = {key: dictionary[key] for key in keys}            
                final_array = []
                for i in range(len(list(selected_dict.values())[0])):
                    values = [selected_dict[key][i] for key in selected_dict]
                    max_val = max(values)
                    min_val = min(values)
                    final_array.append(max_val - min_val)
                return final_array
            DeltaTempArr = fetch_and_calculate(UsableLocalGlobalDataFrame, ['TS1','TS2','TS3','TS4','TS5','TS6','TS7','TS8','TS9','TS10','TS11','TS12','TS0_FLT','TS13_FLT'])
            plt.figure(figsize=(10,6))
            plt.plot(DeltaTempArr,color='g')
            plt.hlines(LimitsData[LimitsData['Parameters'] == 'delta_temperature']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == 'delta_temperature'].index],xmin=0,xmax=len(UsableLocalGlobalDataFrame),linestyle='dashed', linewidth=2)
            plt.legend(('delta_temperature','max_delta_temperature','min_delta_temperature'))
            plt.grid()
            plt.title('Delta Temperature')

            plt.savefig(os.path.join(FolderForSavingAllFiles, 'Cell_Delta_temp_plot.png'))

            #Delta Voltage
            MinValvDeltaVoltage = LimitsData[LimitsData['Parameters'] == 'Delta Voltage']['Min'][LimitsData[LimitsData['Parameters'] == 'Delta Voltage'].index[0]]
            DSGRSTPeriod = LimitsData[LimitsData['Parameters'] == 'Delta Voltage']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == 'Delta Voltage'].index[0]]
            CHGRSTPeriod = LimitsData[LimitsData['Parameters'] == 'Delta Voltage']['Charging Phase & 15 min rest'][LimitsData[LimitsData['Parameters'] == 'Delta Voltage'].index[0]]
            CHGStartOvershoot = LimitsData[LimitsData['Parameters'] == 'Delta Voltage']['chg_start_overshoot'][LimitsData[LimitsData['Parameters'] == 'Delta Voltage'].index[0]]
            CHGEndOvershoot = LimitsData[LimitsData['Parameters'] == 'Delta Voltage']['chg_end_overshoot'][LimitsData[LimitsData['Parameters'] == 'Delta Voltage'].index[0]]
            PostCHGIdle = LimitsData[LimitsData['Parameters'] == 'Delta Voltage']['post_chg_idle'][LimitsData[LimitsData['Parameters'] == 'Delta Voltage'].index[0]]

            LowerBoundDeltaVoltage = []
            for bounding__ in range(len(UsableLocalGlobalDataFrame['Cell Delta Volt'])):
                LowerBoundDeltaVoltage.append(MinValvDeltaVoltage)
            FirstBoundDSGRSTPeriod = []
            series__1 = len(DSGData['SOC']) + len(RST1Data['SOC']) - ThresholdForOvershoot
            for bounding__ in range(series__1):
                FirstBoundDSGRSTPeriod.append(DSGRSTPeriod)
            series__2 = 2*ThresholdForOvershoot
            for bounding__ in range(series__2):
                FirstBoundDSGRSTPeriod.append(CHGStartOvershoot)
            series__3 = len(CHGData['SOC']) - (2*ThresholdForOvershoot)
            for bounding__ in range(series__3):
                FirstBoundDSGRSTPeriod.append(CHGRSTPeriod)
            series__4 = 2*ThresholdForOvershoot
            for bounding__ in range(series__4):
                FirstBoundDSGRSTPeriod.append(CHGEndOvershoot)
            series__5 = len(RST2Data['SOC']) - (2*ThresholdForOvershoot)
            for bounding__ in range(series__5):
                FirstBoundDSGRSTPeriod.append(PostCHGIdle)

            Cell_delta_volt_plot=plt.figure(figsize=(10,6))
            plt.plot(LowerBoundDeltaVoltage,linestyle='dashed')
            plt.plot(FirstBoundDSGRSTPeriod,linestyle='dashed')
            plt.plot(UsableLocalGlobalDataFrame['Cell Delta Volt'])
            plt.legend(('Cell_Delta_Volt_Min','Cell_Delta_Volt_Max','Cell_Delta_Volt_observed'))
            plt.grid()
            plt.title('Cell_delta_voltage')
            plt.suptitle('Cell_delta_voltage')
            plt.ylabel('Voltage (V)')

            plt.savefig(os.path.join(FolderForSavingAllFiles, 'Cell_delta_volt_plot.png'))

            #SOC
            LowerBoundSoC = []
            for bounding__ in range(len(DSGData['SOC'])):
                LowerBoundSoC.append(None)
            for bounding__ in range(len(RST1Data['SOC'])):
                LowerBoundSoC.append(None)
            for bounding__ in range(len(CHGData['SOC'])):
                LowerBoundSoC.append(None)
            for bounding__ in range(len(RST2Data['SOC'])):
                LowerBoundSoC.append(LimitsData[LimitsData['Parameters'] == 'SoC']['Min'][LimitsData[LimitsData['Parameters'] == 'SoC'].index[0]])
            
            UpperBoundSoC = []
            for bounding__ in range(len(DSGData['SOC'])):
                UpperBoundSoC.append(None)
            for bounding__ in range(len(RST1Data['SOC'])):
                UpperBoundSoC.append(None)
            for bounding__ in range(len(CHGData['SOC'])):
                UpperBoundSoC.append(None)
            for bounding__ in range(len(RST2Data['SOC'])):
                UpperBoundSoC.append(LimitsData[LimitsData['Parameters'] == 'SoC']['Dis-Charging Phase & 2mins rest'][LimitsData[LimitsData['Parameters'] == 'SoC'].index[0]])

            SOC_plot=plt.figure(figsize=(10,6))
            plt.plot(LowerBoundSoC, linestyle='dashed')
            plt.plot(UpperBoundSoC, linestyle='dashed')
            plt.plot(UsableLocalGlobalDataFrame['SOC'])
            plt.legend(('SOC_Min','SOC_Max','SOC_observed'))
            plt.grid()
            plt.title('SOC')
            plt.ylabel('SOC (%)')

            plt.savefig(os.path.join(FolderForSavingAllFiles, 'SOC_plot.png'))
            #####################################################################################################################################################################
            #Updating the status
            StatusMessages.append('Passed : EoL Analysis')
            #Updating Result
            if 'Fail' in FaultDetectionResults['Result'].to_numpy():
                CyclingResultsVariable = 'Fail'
            else:
                CyclingResultsVariable = 'Pass'
            StatusMessages.append('Completed : Cycling')
            ResetButtonAblingDisabling = 'Able'
        except:
            ResetButtonAblingDisabling = 'Able'
            StatusMessages.append('Failed : EoL Analysis')

    def ConsecutiveSequence(self, index_list, Threshold):
        sequences = []
        current_sequence = []
        for i in index_list:
            if len(current_sequence) == 0 or i == current_sequence[-1] + 1:
                current_sequence.append(i)
            else:
                if len(current_sequence) >= Threshold:
                    sequences.append(current_sequence)
                current_sequence = [i]
        if len(current_sequence) >= Threshold:
            sequences.append(current_sequence)
        return sequences
    
    def moving_average(data, window_size):
        return data.rolling(window=window_size).mean()

    def SolderIssueDetection(self, data):
        #Signal
        Signal = 0
        #Initializing Parameters
        Threshold = 15
        NeglectFirstRows = 5
        NeglectLastRows = 5
        CellDVThreshold = 0.01
        Distance = 0.01
        #Fetching all Rest periods
        Rest_period_data = data[(data['DSG_Current'] <= 1) & (data['CHG_Current'] <= 1)]
        sequences = self.ConsecutiveSequence(index_list = Rest_period_data.index, Threshold = Threshold)
        result_dfs = [Rest_period_data.loc[seq] for seq in sequences] 
        #Ensure there are Rest Periods in dataframe
        if len(result_dfs) > 0:
            #Iterate through all Rest Periods
            for RestPeriod in range(0,len(result_dfs)):
                df = result_dfs[RestPeriod]
                #Taking required data for algorithm
                df = df.iloc[NeglectFirstRows:].reset_index(drop = True).iloc[:-NeglectLastRows].reset_index(drop = True)[['Cell1', 'Cell2', 'Cell3', 'Cell4', 'Cell5','Cell6', 'Cell7', 'Cell8','Cell9', 'Cell10','Cell11', 'Cell12', 'Cell13', 'Cell14']]
                #Calculate CellDV
                MAX = df.max(axis = 1)
                MIN = df.min(axis = 1)
                CellDV = np.array(MAX - MIN)
                #Condition for CellDV
                if max(CellDV) >= CellDVThreshold:
                    #Mean of all datapoints
                    CentralTendency = [df[f'Cell{avg}'].mean() for avg in range(1, 15)]
                    #Checking if MAX and MIN are adjacent to each other or not
                    if abs(CentralTendency.index(max(CentralTendency)) - CentralTendency.index(min(CentralTendency))) == 1:
                        Q1 = np.percentile(CentralTendency, 25)
                        Q3 = np.percentile(CentralTendency, 75)
                        UpperOutlierLimit = Q3 + Distance
                        LowerOutlierLimit = Q1 - Distance
                        if (max(CentralTendency) > UpperOutlierLimit) and (min(CentralTendency) < LowerOutlierLimit):
                            Signal = Signal + 1
                            break
        return Signal

    def WeldIssueDetection(self, data):
        #Signal
        Signal = 0
        #Initializing Parameters
        Threshold = 50
        valv = 0.02
        SoCCheck = 20
        NeglectFirstRows = 20
        NeglectLastRows = 10
        #Fetching all Rest periods
        Rest_period_data = data[(data['DSG_Current'] <= 1) & (data['CHG_Current'] <= 1)]
        sequences = self.ConsecutiveSequence(index_list = Rest_period_data.index, Threshold = Threshold)
        result_dfs = [Rest_period_data.loc[seq] for seq in sequences] 
        #Ensure there are Rest Periods in dataframe
        if len(result_dfs) > 0:
            #Iterate through all Rest Periods
            for RestPeriod in range(0,len(result_dfs)):
                df = result_dfs[RestPeriod]
                FilteredData = df.iloc[NeglectFirstRows:].reset_index(drop = True).iloc[:-NeglectLastRows].reset_index(drop = True)
                if len(FilteredData) > 1:
                    #Fetching first SOC value
                    SOC = df['SOC'][df.index[0]]
                    #Calculate CellDV
                    OnlyCells = FilteredData[['Cell1', 'Cell2', 'Cell3', 'Cell4', 'Cell5','Cell6', 'Cell7', 'Cell8','Cell9', 'Cell10','Cell11', 'Cell12', 'Cell13', 'Cell14']]
                    MAX = OnlyCells.max(axis = 1)
                    MIN = OnlyCells.min(axis = 1)
                    CellDV = np.array(MAX - MIN)
                    if SOC <= SoCCheck:
                        if min(CellDV) >= valv:
                            Signal = Signal + 1
        return Signal

    def TemperatureFluctuationDetection(self, data):
        #data initialization
        TS_FILE_DF = data
        #Signal Initialization
        Signal = 0
        SensorWithIssue = None
        #Sensor Names
        TempArrays = ['TS1', 'TS2', 'TS3', 'TS4', 'TS5', 'TS6', 'TS7', 'TS8', 'TS9', 'TS10', 'TS11', 'TS12', 'TS0_FLT', 'TS13_FLT']
        # TempArrays = ['ts1', 'ts2', 'ts3', 'ts4', 'ts5', 'ts6','ts7', 'ts8', 'ts9', 'ts10', 'ts11', 'ts12', 'ts0_flt', 'ts13_flt']
        #Parameters
        ThresholdValv1 = 0.0011
        ThresholdValv2 = 0.0025
        WindowThreshold = 20
        #FeatureFetching
        TS_FILE_DF = TS_FILE_DF[TempArrays]
        #Mean Centering
        TS_FILE_DF = TS_FILE_DF - TS_FILE_DF.mean()
        #Moving average filtered New data
        NewArr = np.array([moving_average(data = TS_FILE_DF[i], window_size = WindowThreshold) for i in TempArrays])
        #OldArr
        OldArr = np.array([np.array(TS_FILE_DF[i]) for i in TempArrays])
        #Old vs New Diff
        DiffArr = OldArr-NewArr
        #Varience Storage
        VarianceStorage = [np.var(DiffArr[i][WindowThreshold-1:]) for i in range(0,len(DiffArr))]
        #Threshold
        for i in range(0,len(VarianceStorage)):
            if i <= 11: #This is the threshold given for temperature sensors 'ts1', 'ts2', 'ts3', 'ts4', 'ts5','ts6', 'ts7', 'ts8', 'ts9', 'ts10', 'ts11', 'ts12'
                if VarianceStorage[i] > ThresholdValv1:
                    SensorWithIssue = f"TS{i+1}"
                    Signal = Signal + 1
                    break
            else:       #This is the threshold given for temperature sensors 'ts0_flt','ts13_flt'
                if VarianceStorage[i] > ThresholdValv2:
                    if i == 12:
                        SensorWithIssue = "TS0_FLT"
                        Signal = Signal + 1
                        break
                    elif i == 13:
                        SensorWithIssue = "TS13_FLT"
                        Signal = Signal + 1
                        break
        return Signal

    def ThermisterOpenIssueDetection(self, data):
        #Signal
        Signal = 0
        #Logic
        for i in range(0,14):
            if i == 0:
                if data['TS0_FLT'].max() >= 600:
                    for y in data.index:
                        if data['TS0_FLT'][y] >= 600:
                            starting_index = y
                            break
                    req_arr = np.array(data['TS0_FLT'][starting_index:])
                    if all(element >= 600 for element in req_arr):
                        Signal = Signal + 1
                        break

            elif i == 13:
                if data['TS13_FLT'].max() >= 600:
                    for y in data.index:
                        if data['TS13_FLT'][y] >= 600:
                            starting_index = y
                            break
                    req_arr = np.array(data['TS13_FLT'][starting_index:])
                    if all(element >= 600 for element in req_arr):
                        Signal = Signal + 1
                        break

            else:
                if data[f'TS{i}'].max() >= 600:
                    for y in data.index:
                        if data[f'TS{i}'][y] >= 600:
                            starting_index = y
                            break
                    req_arr = np.array(data[f'TS{i}'][starting_index:])
                    if all(element >= 600 for element in req_arr):
                        Signal = Signal + 1
                        break
        return Signal

    def DeltaTemperatureIssueDetection(self, data):
        #Signal
        Signal = 0
        #Calculating the Cell DV
        max_values = np.array(data[['TS1', 'TS2', 'TS3', 'TS4', 'TS5', 'TS6', 'TS7','TS8', 'TS9', 'TS10', 'TS11', 'TS12','TS0_FLT', 'TS13_FLT']].max(axis=1))
        min_values = np.array(data[['TS1', 'TS2', 'TS3', 'TS4', 'TS5', 'TS6', 'TS7','TS8', 'TS9', 'TS10', 'TS11', 'TS12','TS0_FLT', 'TS13_FLT']].min(axis=1))

        #Calculating Delta Temperature
        delta_temp = max_values - min_values

        #Criterion for Delta Temperature
        if len(delta_temp) > 10:
            Counter = 0
            for iterate in delta_temp:
                if iterate >= 6:
                    Counter = Counter + 1
            if Counter >= 0.7*len(delta_temp):
                Signal = Signal + 1
        return Signal

class FinalPassFailDisplay(QThread):

    Status = pyqtSignal(str)

    def __init__(self):
        super().__init__()

    def run(self):
        global CyclingResultsVariable
        global StatusMessages

        while ResetButtonClicked != 'Clicked':
            if (CyclingResultsVariable == 'Pass') or (CyclingResultsVariable == 'Fail'):

                if CyclingResultsVariable == 'Pass':
                    self.Status.emit('Pass')
                else:
                    self.Status.emit('Fail')
                
                break
            time.sleep(1)

class ResetButtonFunctions(QThread):

    Status = pyqtSignal(int)
    Status2 = pyqtSignal(int)

    def __init__(self):
        super().__init__()

    def run(self):

        global AllData
        global GlobalDictionary
        global SerialConnectionCycler
        global SerialConnectionVCU
        global SerialConnectionQRScanner
        global PortAssignedToCycler
        global BatteryPackName
        global PackType
        global StoppageCurrentDuetoPowerCut
        global StopConditions
        global FeedParms
        global FaultDetectionResults
        global StatusMessages
        global ErrorMessages
        global ConnectionCompleted
        global ScanCompleted
        global StartButtonClicked
        global StartTaskEndsHere
        global ThresholdForOvershoot
        global FolderForSavingAllFiles
        global CyclingResultsVariable
        global RelativePathHR
        global RelativePathLR
        global ResetButtonClicked
        global AllThreads
        global ThresholdTimeToInitializeDisplay
        global AllowStartTempClause
        global CyclingStatus
        global ConnectButtonAblingDisabling
        global ScanButtonAblingDisabling
        global StartButtonAblingDisabling
        global EmergencyButtonAblingDisabling
        global ResetButtonAblingDisabling

        ResetButtonAblingDisabling = 'Disable'
        EmergencyButtonAblingDisabling = 'Disable'
        StartButtonAblingDisabling = 'Disable'
        ScanButtonAblingDisabling = 'Disable'
        ConnectButtonAblingDisabling = 'Disable'

        time.sleep(2)

        #Closing All Threads
        ResetButtonClicked = 'Clicked'

        #Checking if all threads are closed or not
        for ThreadNumber in AllThreads:
            if isinstance(ThreadNumber, threading.Thread):
                while True:
                    if not ThreadNumber.is_alive():
                        break
                    time.sleep(0.001)
            elif isinstance(ThreadNumber, QThread):
                while True:
                    if not ThreadNumber.isRunning():
                        break
                    time.sleep(0.001)

        #Initialize Cycler
        SerialConnectionCycler.write(b"SYST:LOCK ON")
        time.sleep(0.1)
        #Turing off Cycler
        while True:
            SerialConnectionCycler.write(b"OUTP OFF")
            time.sleep(0.1)
            SerialConnectionCycler.write(b"OUTPUT?")
            if SerialConnectionCycler.readline().decode().strip() == 'OFF':
                time.sleep(0.1)
                break

        #Close all Serial Connections
        if SerialConnectionCycler != None:
            SerialConnectionCycler.close()
        if SerialConnectionVCU != None:
            SerialConnectionVCU.close()
        if SerialConnectionQRScanner != None:
            SerialConnectionQRScanner.close()

        #Initialize all parmeters
        AllData = ['Slot1','Millis','FET_ON_OFF','CFET ON/OFF','FET TEMP1','DSG_VOLT','SC Current',
            'DSG_Current','CHG_VOLT','CHG_Current','DSG Time','CHG Time','DSG Charge','CHG Charge',
            'Cell1','Cell2','Cell3','Cell4','Cell5','Cell6','Cell7','Cell8','Cell9','Cell10','Cell11',
            'Cell12','Cell13','Cell14','Cell Delta Volt','Sum-of-cells','DSG Power','DSG Energy','CHG Power',
            'CHG Energy','Min CV','BAL_ON_OFF','TS1','TS2','TS3','TS4','TS5','TS6','TS7','TS8','TS9',
            'TS10','TS11','TS12','FET Temp Front','BAT + ve Temp','BAT - ve Temp','Pack + ve Temp','TS0_FLT',
            'TS13_FLT','FET_TEMP_REAR','DSG INA','BAL_RES_TEMP','HUM','IMON','Hydrogen','FG_CELL_VOLT',
            'FG_PACK_VOLT','FG_AVG_CURN','SOC','MAX_TTE','MAX_TTF','REPORTED_CAP','TS0_FLT1','IR','Cycles','DS_CAP',
            'FSTAT','VFSOC','CURR','CAP_NOM','REP_CAP.1','AGE','AGE_FCAST','QH','ICHGTERM','DQACC','DPACC','QRESIDUAL','MIXCAP']
        GlobalDictionary = {}
        SerialConnectionCycler = None
        SerialConnectionVCU = None
        SerialConnectionQRScanner = None
        PortAssignedToCycler = None
        BatteryPackName = None
        PackType = None
        StoppageCurrentDuetoPowerCut = None
        FaultDetectionResults = pd.DataFrame()
        StatusMessages = []
        ErrorMessages = []
        ConnectionCompleted = None
        ScanCompleted= None
        StartButtonClicked = None
        StartTaskEndsHere = None
        FolderForSavingAllFiles = None
        CyclingResultsVariable = None
        ResetButtonClicked = None
        AllThreads = []
        AllowStartTempClause = None
        CyclingStatus = None

        #Status Emitting to initialize all displays, Clear Status
        self.Status.emit(1)

        time.sleep(ThresholdTimeToInitializeDisplay)

        #Recalling all Threads
        self.Status2.emit(2)

class ButtonAbleDisable(QThread):
    
    ConnectButtonStatus = pyqtSignal(int)
    ScanButtonStatus = pyqtSignal(int)
    StartButtonStatus = pyqtSignal(int)
    ResetButtonStatus = pyqtSignal(int)

    def __init__(self):
        super().__init__()
    
    def run(self):
        global ConnectButtonAblingDisabling
        global ScanButtonAblingDisabling
        global StartButtonAblingDisabling
        global ResetButtonAblingDisabling

        while ResetButtonClicked != 'Clicked':
            if ConnectButtonAblingDisabling == 'Able':
                self.ConnectButtonStatus.emit(1)
            else:
                self.ConnectButtonStatus.emit(0)
            if ScanButtonAblingDisabling == 'Able':
                self.ScanButtonStatus.emit(1)
            else:
                self.ScanButtonStatus.emit(0)
            if StartButtonAblingDisabling == 'Able':
                self.StartButtonStatus.emit(1)
            else:
                self.StartButtonStatus.emit(0)
            if ResetButtonAblingDisabling == 'Able':
                self.ResetButtonStatus.emit(1)
            else:
                self.ResetButtonStatus.emit(0)
            time.sleep(0.01)

class FETOFFEoLAnalysis(threading.Thread):

    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):

        global StatusMessages
        global GlobalDictionary
        global FolderForSavingAllFiles
        global SerialConnectionCycler
        global CyclingResultsVariable
        global ResetButtonAblingDisabling
        global ResetButtonClicked

        #Updating Status
        StatusMessages.append('FET Turned OFF!')

        #Turn off cycler
        while True:
            SerialConnectionCycler.write(b"OUTP OFF")
            time.sleep(0.1)
            SerialConnectionCycler.write(b"OUTPUT?")
            if SerialConnectionCycler.readline().decode().strip() == 'OFF':
                time.sleep(0.1)
                break
        
        #Enabling Reset Button
        ResetButtonAblingDisabling = 'Able'
        
        #Updating Status
        StatusMessages.append('Started : FET OFF EoL Analysis | It will take 150 seconds. Click Reset if EoL Analysis file not needed.')

        #Time sleep
        for range_ in range(120):
            if ResetButtonClicked != 'Clicked':
                time.sleep(1)
            else:
                break

        if ResetButtonClicked != 'Clicked':
            #Data
            UsableLocalGlobalDataFrame = pd.DataFrame.from_dict(GlobalDictionary)
            StorageDf = pd.DataFrame()

            #EoL Algorithms
            #####################################################################################################################################################################
            #Solder Issue
            SolderIssueDataframe = UsableLocalGlobalDataFrame[['Cell1', 'Cell2', 'Cell3', 'Cell4', 'Cell5','Cell6', 'Cell7', 'Cell8','Cell9', 'Cell10','Cell11', 'Cell12', 'Cell13', 'Cell14', 'DSG_Current', 'CHG_Current']]
            SolderIssueSignal = self.SolderIssueDetection(SolderIssueDataframe)
            if SolderIssueSignal >= 1:
                DictToSave = {
                    'Parameters':['Solder Issue'],
                    'Result':['Fail']
                }
                StorageDf = pd.concat([StorageDf, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            else:
                DictToSave = {
                    'Parameters':['Solder Issue'],
                    'Result':['Pass']
                }
                StorageDf = pd.concat([StorageDf, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            #Weld Issue Charging
            WeldIssueDataFrame = UsableLocalGlobalDataFrame[['Cell1', 'Cell2', 'Cell3', 'Cell4', 'Cell5','Cell6', 'Cell7', 'Cell8','Cell9', 'Cell10','Cell11', 'Cell12', 'Cell13', 'Cell14', 'DSG_Current', 'CHG_Current', 'SOC']]
            WeldIssueSignal = self.WeldIssueDetection(WeldIssueDataFrame)
            if WeldIssueSignal >= 1:
                DictToSave = {
                    'Parameters':['Weld Issue'],
                    'Result':['Fail']
                }
                StorageDf = pd.concat([StorageDf, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            else:
                DictToSave = {
                    'Parameters':['Weld Issue'],
                    'Result':['Pass']
                }
                StorageDf = pd.concat([StorageDf, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            #Temperature Fluctuation Issue 
            TemperatureFluctuationDataFrame = UsableLocalGlobalDataFrame[['Millis','TS1','TS2','TS3','TS4','TS5','TS6','TS7','TS8','TS9','TS10','TS11','TS12','TS0_FLT','TS13_FLT']]
            TemperatureFluctuationSignal = self.TemperatureFluctuationDetection(TemperatureFluctuationDataFrame)
            if TemperatureFluctuationSignal >= 1:
                DictToSave = {
                    'Parameters':['Temperature Fluctuation'],
                    'Result':['Fail']
                }
                StorageDf = pd.concat([StorageDf, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            else:
                DictToSave = {
                    'Parameters':['Temperature Fluctuation'],
                    'Result':['Pass']
                }
                StorageDf = pd.concat([StorageDf, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            #Thermister Open Issue
            ThermisterOpenDataFrame = UsableLocalGlobalDataFrame[['TS1','TS2','TS3','TS4','TS5','TS6','TS7','TS8','TS9','TS10','TS11','TS12','TS0_FLT','TS13_FLT']]
            ThermisterOpenSignal = self.ThermisterOpenIssueDetection(ThermisterOpenDataFrame)
            if ThermisterOpenSignal >= 1:
                DictToSave = {
                    'Parameters':['Thermister Open'],
                    'Result':['Fail']
                }
                StorageDf = pd.concat([StorageDf, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            else:
                DictToSave = {
                    'Parameters':['Thermister Open'],
                    'Result':['Pass']
                }
                StorageDf = pd.concat([StorageDf, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            #Delta Temperature Issue
            DeltaTempDataFrame = UsableLocalGlobalDataFrame[['TS1','TS2','TS3','TS4','TS5','TS6','TS7','TS8','TS9','TS10','TS11','TS12','TS0_FLT','TS13_FLT']]
            DeltaTempSignal = self.DeltaTemperatureIssueDetection(DeltaTempDataFrame)
            if DeltaTempSignal >= 1:
                DictToSave = {
                    'Parameters':['Delta Temperature'],
                    'Result':['Fail']
                }
                StorageDf = pd.concat([StorageDf, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            else:
                DictToSave = {
                    'Parameters':['Delta Temperature'],
                    'Result':['Pass']
                }
                StorageDf = pd.concat([StorageDf, pd.DataFrame.from_dict(DictToSave)], ignore_index = True).reset_index(drop = True)
            #####################################################################################################################################################################
            
            #Saving data
            particulate = ['TS1','TS2','TS3','TS4','TS5','TS6','TS7','TS8','TS9','TS10','TS11','TS12','FET Temp Front','BAT + ve Temp','BAT - ve Temp','Pack + ve Temp','TS0_FLT','TS13_FLT','FET_TEMP_REAR']
            for particle in particulate:
                dTdtArray_Paticle = UsableLocalGlobalDataFrame[particle].diff()/UsableLocalGlobalDataFrame['Millis'].diff()
                UsableLocalGlobalDataFrame[f'd({particle})/dt'] = dTdtArray_Paticle
            UsableLocalGlobalDataFrame.to_csv(os.path.join(FolderForSavingAllFiles, 'PreCyclerData.csv'), index = False)
            
            def _color_red_or_green(val):
                color = 'red' if val == 'Fail' else 'green'
                return 'color: %s' % color
            try:
                # Apply the color formatting to the DataFrame
                styled_data = StorageDf.style.applymap(_color_red_or_green, subset='Result')
                styled_data.to_excel(os.path.join(FolderForSavingAllFiles, 'PreSummarySheet.xlsx'), index = False)
            except:
                StorageDf.to_excel(os.path.join(FolderForSavingAllFiles, 'PreSummarySheet.xlsx'), index = False)                

            #Update Status
            StatusMessages.append('Completed : FET OFF EoL Analysis.')

        #Update variable
        CyclingResultsVariable = 'Fail'
        StatusMessages.append('Click RESET to START new Cycling')

    def ConsecutiveSequence(self, index_list, Threshold):
        sequences = []
        current_sequence = []
        for i in index_list:
            if len(current_sequence) == 0 or i == current_sequence[-1] + 1:
                current_sequence.append(i)
            else:
                if len(current_sequence) >= Threshold:
                    sequences.append(current_sequence)
                current_sequence = [i]
        if len(current_sequence) >= Threshold:
            sequences.append(current_sequence)
        return sequences

    def moving_average(data, window_size):
        return data.rolling(window=window_size).mean()

    def SolderIssueDetection(self, data):
        #Signal
        Signal = 0
        #Initializing Parameters
        Threshold = 15
        NeglectFirstRows = 5
        NeglectLastRows = 5
        CellDVThreshold = 0.01
        Distance = 0.01
        #Fetching all Rest periods
        Rest_period_data = data[(data['DSG_Current'] <= 1) & (data['CHG_Current'] <= 1)]
        sequences = self.ConsecutiveSequence(index_list = Rest_period_data.index, Threshold = Threshold)
        result_dfs = [Rest_period_data.loc[seq] for seq in sequences] 
        #Ensure there are Rest Periods in dataframe
        if len(result_dfs) > 0:
            #Iterate through all Rest Periods
            for RestPeriod in range(0,len(result_dfs)):
                df = result_dfs[RestPeriod]
                #Taking required data for algorithm
                df = df.iloc[NeglectFirstRows:].reset_index(drop = True).iloc[:-NeglectLastRows].reset_index(drop = True)[['Cell1', 'Cell2', 'Cell3', 'Cell4', 'Cell5','Cell6', 'Cell7', 'Cell8','Cell9', 'Cell10','Cell11', 'Cell12', 'Cell13', 'Cell14']]
                #Calculate CellDV
                MAX = df.max(axis = 1)
                MIN = df.min(axis = 1)
                CellDV = np.array(MAX - MIN)
                #Condition for CellDV
                if max(CellDV) >= CellDVThreshold:
                    #Mean of all datapoints
                    CentralTendency = [df[f'Cell{avg}'].mean() for avg in range(1, 15)]
                    #Checking if MAX and MIN are adjacent to each other or not
                    if abs(CentralTendency.index(max(CentralTendency)) - CentralTendency.index(min(CentralTendency))) == 1:
                        Q1 = np.percentile(CentralTendency, 25)
                        Q3 = np.percentile(CentralTendency, 75)
                        UpperOutlierLimit = Q3 + Distance
                        LowerOutlierLimit = Q1 - Distance
                        if (max(CentralTendency) > UpperOutlierLimit) and (min(CentralTendency) < LowerOutlierLimit):
                            Signal = Signal + 1
                            break
        return Signal

    def WeldIssueDetection(self, data):
        #Signal
        Signal = 0
        #Initializing Parameters
        Threshold = 50
        valv = 0.02
        SoCCheck = 20
        NeglectFirstRows = 20
        NeglectLastRows = 10
        #Fetching all Rest periods
        Rest_period_data = data[(data['DSG_Current'] <= 1) & (data['CHG_Current'] <= 1)]
        sequences = self.ConsecutiveSequence(index_list = Rest_period_data.index, Threshold = Threshold)
        result_dfs = [Rest_period_data.loc[seq] for seq in sequences] 
        #Ensure there are Rest Periods in dataframe
        if len(result_dfs) > 0:
            #Iterate through all Rest Periods
            for RestPeriod in range(0,len(result_dfs)):
                df = result_dfs[RestPeriod]
                FilteredData = df.iloc[NeglectFirstRows:].reset_index(drop = True).iloc[:-NeglectLastRows].reset_index(drop = True)
                if len(FilteredData) > 1:
                    #Fetching first SOC value
                    SOC = df['SOC'][df.index[0]]
                    #Calculate CellDV
                    OnlyCells = FilteredData[['Cell1', 'Cell2', 'Cell3', 'Cell4', 'Cell5','Cell6', 'Cell7', 'Cell8','Cell9', 'Cell10','Cell11', 'Cell12', 'Cell13', 'Cell14']]
                    MAX = OnlyCells.max(axis = 1)
                    MIN = OnlyCells.min(axis = 1)
                    CellDV = np.array(MAX - MIN)
                    if SOC <= SoCCheck:
                        if min(CellDV) >= valv:
                            Signal = Signal + 1
        return Signal

    def TemperatureFluctuationDetection(self, data):
        #data initialization
        TS_FILE_DF = data
        #Signal Initialization
        Signal = 0
        SensorWithIssue = None
        #Sensor Names
        TempArrays = ['TS1', 'TS2', 'TS3', 'TS4', 'TS5', 'TS6', 'TS7', 'TS8', 'TS9', 'TS10', 'TS11', 'TS12', 'TS0_FLT', 'TS13_FLT']
        # TempArrays = ['ts1', 'ts2', 'ts3', 'ts4', 'ts5', 'ts6','ts7', 'ts8', 'ts9', 'ts10', 'ts11', 'ts12', 'ts0_flt', 'ts13_flt']
        #Parameters
        ThresholdValv1 = 0.0011
        ThresholdValv2 = 0.0025
        WindowThreshold = 20
        #FeatureFetching
        TS_FILE_DF = TS_FILE_DF[TempArrays]
        #Mean Centering
        TS_FILE_DF = TS_FILE_DF - TS_FILE_DF.mean()
        #Moving average filtered New data
        NewArr = np.array([moving_average(data = TS_FILE_DF[i], window_size = WindowThreshold) for i in TempArrays])
        #OldArr
        OldArr = np.array([np.array(TS_FILE_DF[i]) for i in TempArrays])
        #Old vs New Diff
        DiffArr = OldArr-NewArr
        #Varience Storage
        VarianceStorage = [np.var(DiffArr[i][WindowThreshold-1:]) for i in range(0,len(DiffArr))]
        #Threshold
        for i in range(0,len(VarianceStorage)):
            if i <= 11: #This is the threshold given for temperature sensors 'ts1', 'ts2', 'ts3', 'ts4', 'ts5','ts6', 'ts7', 'ts8', 'ts9', 'ts10', 'ts11', 'ts12'
                if VarianceStorage[i] > ThresholdValv1:
                    SensorWithIssue = f"TS{i+1}"
                    Signal = Signal + 1
                    break
            else:       #This is the threshold given for temperature sensors 'ts0_flt','ts13_flt'
                if VarianceStorage[i] > ThresholdValv2:
                    if i == 12:
                        SensorWithIssue = "TS0_FLT"
                        Signal = Signal + 1
                        break
                    elif i == 13:
                        SensorWithIssue = "TS13_FLT"
                        Signal = Signal + 1
                        break
        return Signal

    def ThermisterOpenIssueDetection(self, data):
        #Signal
        Signal = 0
        #Logic
        for i in range(0,14):
            if i == 0:
                if data['TS0_FLT'].max() >= 600:
                    for y in data.index:
                        if data['TS0_FLT'][y] >= 600:
                            starting_index = y
                            break
                    req_arr = np.array(data['TS0_FLT'][starting_index:])
                    if all(element >= 600 for element in req_arr):
                        Signal = Signal + 1
                        break

            elif i == 13:
                if data['TS13_FLT'].max() >= 600:
                    for y in data.index:
                        if data['TS13_FLT'][y] >= 600:
                            starting_index = y
                            break
                    req_arr = np.array(data['TS13_FLT'][starting_index:])
                    if all(element >= 600 for element in req_arr):
                        Signal = Signal + 1
                        break

            else:
                if data[f'TS{i}'].max() >= 600:
                    for y in data.index:
                        if data[f'TS{i}'][y] >= 600:
                            starting_index = y
                            break
                    req_arr = np.array(data[f'TS{i}'][starting_index:])
                    if all(element >= 600 for element in req_arr):
                        Signal = Signal + 1
                        break
        return Signal

    def DeltaTemperatureIssueDetection(self, data):
        #Signal
        Signal = 0
        #Calculating the Cell DV
        max_values = np.array(data[['TS1', 'TS2', 'TS3', 'TS4', 'TS5', 'TS6', 'TS7','TS8', 'TS9', 'TS10', 'TS11', 'TS12','TS0_FLT', 'TS13_FLT']].max(axis=1))
        min_values = np.array(data[['TS1', 'TS2', 'TS3', 'TS4', 'TS5', 'TS6', 'TS7','TS8', 'TS9', 'TS10', 'TS11', 'TS12','TS0_FLT', 'TS13_FLT']].min(axis=1))

        #Calculating Delta Temperature
        delta_temp = max_values - min_values

        #Criterion for Delta Temperature
        if len(delta_temp) > 10:
            Counter = 0
            for iterate in delta_temp:
                if iterate >= 6:
                    Counter = Counter + 1
            if Counter >= 0.7*len(delta_temp):
                Signal = Signal + 1
        return Signal

if __name__ == "__main__":
    pg.setConfigOption('background', 'w')
    pg.setConfigOption('foreground', 'k')
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
