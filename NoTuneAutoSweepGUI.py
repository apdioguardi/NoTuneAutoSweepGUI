__author__ = 'adioguardi@gmail.com (Adam P. Dioguardi)'
__version__ = 1.0

## To Do:
## actually replace testy with the no-tune autosweep function
## after abort, can't restart... abort bool must be True still
## (even after clear)
##
## !ahoy = search string for things that can be removed when all is working

#***********************************************************#
#                Import Various libraries                   #
#***********************************************************#
import sys
import time
import datetime
import math
import os
import os.path
import smtplib
from email.mime.text import MIMEText

#import win32com.client # !ahoy

from PySide.QtUiTools import QUiLoader
from PySide import QtGui, QtCore
from PySide.QtGui import QIntValidator, QDoubleValidator

import matplotlib
matplotlib.use('Qt4Agg')
matplotlib.rcParams['backend.qt4']='PySide'
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt4agg import FigureCanvasQTAgg as FigureCanvas

#import pythoncom #!ahoy
#pythoncom.CoInitialize() #!ahoy

#***********************************************************#
#                Global Variables                           #
#***********************************************************#
start_time = 0 #!ahoy
start_freq = 1 #!ahoy
end_freq = 2 #!ahoy
step_freq = 0.1 #!ahoy
email_notification = 0 #!ahoy
to_addresses = "adioguardi@gmail.com" #!ahoy
subject = "AutoSweep has finished!" #!ahoy
message = "Someone analyze me! -The Data" #!ahoy
total_scan_time = 0 #!ahoy
elapsed_scan_time = 0 #!ahoy
sweep_finished_bool = False
graphX = []
graphY = []
run_sweep_bool = False
error_caught_bool = False
elapse_bool = False
abort_bool = False
error_message = ""
integration_file = ""

class UpdateStatsThread(QtCore.QThread):
    #***********************************************************#
    #                Initialize the class                       #
    #***********************************************************#
    def __init__(self, parent=None):
        super(UpdateStatsThread, self).__init__(parent)
        self._running = False
    #***********************************************************#
    #    Define what function to call and timer interval        #
    #***********************************************************#
    def run(self):
        self._running = True
        while self._running:
            self.doWork()
            self.msleep(1000)
    #***********************************************************#
    #    Main function, writes to machines, writes to file      #
    #***********************************************************#          
    def doWork(self):
        global run_sweep_bool, error_caught_bool, error_message
        if run_sweep_bool:
            # pythoncom.CoInitialize() #!ahoy
            run_sweep_bool = False
            try:
                testy() #!ahoy testy() --> NoTuneAutoSweep()
                return 0
            except Exception, err:
                error_caught_bool = True
                error_message = err
                sys.stderr.write('THIS IS THE ERROR: %s\n' % str(error_message)) #!ahoy
                return 1

#***********************************************************#
#         Class to update GUI independent of loop           #
#***********************************************************#
class UpdateGuiWithStats(QtCore.QObject):
    def startWorker(self):
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.doWork)
        self.timer.start(10)

#***********************************************************#
#               Class for GUI                               #
#***********************************************************#
class NoTuneAutoSweepGUI(QtGui.QMainWindow):
    def __init__(self):
        super(NoTuneAutoSweepGUI,self).__init__()
        loader = QUiLoader()
        UIfile = QtCore.QFile("NoTuneAutoSweepGUI.ui")
        UIfile.open(QtCore.QFile.ReadOnly)
        self.myWidget = loader.load(UIfile,self)
        UIfile.close()
        self.initUI()
        #***********************************************************#
        #                Create New Thread                          #
        #***********************************************************#
        self.statsThread = UpdateStatsThread()
        self.statsThread.start(QtCore.QThread.TimeCriticalPriority)

    def initUI(self):
        dock = QtGui.QDockWidget("",self)
        dock.setFeatures(dock.NoDockWidgetFeatures)
        posValidator = QDoubleValidator(0.0,500.0,10,self)
        boolValidator = QIntValidator(0, 1, self)

        #***********************************************************#
        #                Connect Signals and Slots for Widgets      #
        #***********************************************************#

        self.myWidget.StartButton.clicked.connect(self.startButtonClicked)
        self.myWidget.ClearButton.clicked.connect(self.clearButtonClicked)
        self.myWidget.AbortButton.clicked.connect(self.abortButtonClicked)

        #***********************************************************#
        #                Connect Text Fields to update              #
        #***********************************************************#

        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.updateT)
        self.timer.start(1000)

        self.myWidget.EmailNotificationTxt.setValidator(boolValidator)

        self.myWidget.StartFreqTxt.editingFinished.connect(self.onChanged)
        self.myWidget.EndFreqTxt.editingFinished.connect(self.onChanged)
        self.myWidget.StepSizeTxt.editingFinished.connect(self.onChanged)
        self.myWidget.EmailNotificationTxt.editingFinished.connect(self.onChanged)
        self.myWidget.SubjectTxt.editingFinished.connect(self.onChanged)
        self.myWidget.MessageTxt.editingFinished.connect(self.onChanged)


        self.myWidget.StartFreqTxt.setText(str(start_freq))
        self.myWidget.EndFreqTxt.setText(str(end_freq))
        self.myWidget.StepSizeTxt.setText(str(step_freq))
        self.myWidget.EmailNotificationTxt.setText(str(email_notification))
        self.myWidget.ToAddressesTxt.setText(str(to_addresses))
        self.myWidget.SubjectTxt.setText(str(subject))
        self.myWidget.MessageTxt.setText(str(message))

        if not total_scan_time == 0:
            self.myWidget.progressBar.setValue(int((elapsed_scan_time/total_scan_time)*100))
        else:
            self.myWidget.progressBar.setValue(int(0))

        self.myWidget.TotalScanTimeTxt.setText("")
        self.myWidget.ElapsedScanTimeTxt.setText("")

        self.figureData = Figure()
        self.figureData.set_size_inches(7.5,4.2)
        self.canvasData = FigureCanvas(self.figureData)
        self.canvasData.setParent(self.myWidget.GraphData)
        self.axesData = self.figureData.add_subplot(111)
        self.axesData.set_xlabel("Frequency (MHz)")
        self.axesData.set_ylabel("Echo Integral (Arb. Units)")
        y_formatter = matplotlib.ticker.ScalarFormatter(useOffset=False)
        self.axesData.yaxis.set_major_formatter(y_formatter)
        self.axesData.xaxis.set_major_formatter(y_formatter)
        self.canvasData.draw()

        self.exitAction = QtGui.QAction('&Exit', self)
        self.exitAction.triggered.connect(self.close)
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')
        fileMenu.addAction(self.exitAction)

        self.setCentralWidget(self.myWidget)
        self.setWindowTitle('No-Tune Auto Sweep GUI')
        self.setGeometry(100,100,620,645)
        self.show()
    #***********************************************************#
    #                Clear Entered Fields                       #
    #***********************************************************#
    def clearButtonClicked(self):
        global start_time, start_freq, end_freq, step_freq
        global email_notification, to_addresses, subject, message
        global elapsed_scan_time, total_scan_time, sweep_finished_bool
        global graphX, graphY, run_sweep_bool, error_caught_bool
        global elapse_bool, abort_Bool, error_message, integration_file

        start_time = 0 #!ahoy
        start_freq = 2 #!ahoy
        end_freq = 3 #!ahoy
        step_freq = 0.01 #!ahoy
        email_notification = 0 #!ahoy
        to_addresses = "adioguardi@gmail.com" #!ahoy
        subject = "AutoSweep has finished!" #!ahoy
        message = "Someone analyze me! -The Data" #!ahoy
        total_scan_time = 0 #!ahoy
        elapsed_scan_time = 0 #!ahoy
        sweep_finished_bool = False
        graphX = []
        graphY = []
        run_sweep_bool = False
        error_caught_bool = False
        elapse_bool = False
        abort_bool = False
        error_message = ""
        integration_file = ""

        self.myWidget.StartFreqTxt.setText("")
        self.myWidget.EndFreqTxt.setText("")
        self.myWidget.StepSizeTxt.setText("")

        self.myWidget.EmailNotificationTxt.setText("")
        self.myWidget.ToAddressesTxt.setText("")
        self.myWidget.SubjectTxt.setText("")
        self.myWidget.MessageTxt.setText("")

        self.myWidget.TotalScanTimeTxt.setText("")
        self.myWidget.ElapsedScanTimeTxt.setText("")


    #***********************************************************#
    #         Abort The Sweep                                   #
    #***********************************************************#
    def abortButtonClicked(self):
        global abort_bool

        reply = QtGui.QMessageBox.question(self, 'Message', "Are you sure you want to abort?", QtGui.QMessageBox.Yes | 
                    QtGui.QMessageBox.No, QtGui.QMessageBox.No)
        if reply == QtGui.QMessageBox.Yes:
            abort_bool = True

    #***********************************************************#
    #               Change the global variables upon input      #
    #***********************************************************#        
    def onChanged(self):
        global start_freq, end_freq, step_freq, email_notification
        global to_addresses, subject, message

        if not(self.sender().text() == ""):
            string = str(self.sender().objectName() + ' = ' + '"' + self.sender().text()) + '"'
            exec string in globals()

    #***********************************************************#
    #               Overload closeEvent                         #
    #***********************************************************#
    #overloaded function to exit the GUI without throwing and error
    def closeEvent(self,event):
        #exit the app without throwing an error
        reply = QtGui.QMessageBox.question(self, 'Message',"Do you want to exit", QtGui.QMessageBox.Yes |
                    QtGui.QMessageBox.No, QtGui.QMessageBox.No)
        if reply == QtGui.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
    #***********************************************************#
    #               Timer update                                #
    #***********************************************************#   
    def updateT(self):
        global error_caught_bool, total_scan_time, elapsed_scan_time
        global run_sweep_bool, sweep_finished_bool, elapse_bool, abort_bool
        global start_freq, end_freq, step_freq, email_notification
        global to_addresses, subject, message, start_time

        if elapse_bool and not abort_bool:
            elapsed_scan_time = time.time() - start_time

            TotalTimeString = str(datetime.timedelta(seconds=int(total_scan_time)))
            self.myWidget.TotalScanTimeTxt.setText(TotalTimeString)

            ElapsedTimeString = str(datetime.timedelta(seconds=int(elapsed_scan_time)))
            self.myWidget.ElapsedScanTimeTxt.setText(ElapsedTimeString)
            if not total_scan_time == 0:
                self.myWidget.progressBar.setValue(int((elapsed_scan_time/total_scan_time)*100))

        if sweep_finished_bool and not abort_bool:
            FinishedTimeString = str(datetime.timedelta(seconds=int(elapsed_scan_time)))
            self.myWidget.TotalScanTimeTxt.setText(FinishedTimeString)
            self.myWidget.progressBar.setValue(int(100)) #!ahoy this need to be able to clear this.

        if abort_bool:
            self.myWidget.progressBar.setValue(int(0))
            self.myWidget.ElapsedScanTimeTxt.setText("Aborted")
            self.myWidget.TotalScanTimeTxt.setText("")

        if error_caught_bool:
            error_caught_bool = False
            en = email_notifier()
            if email_notification == 1:
                en.sendemail(to_addresses, "Error caught!!!", "Get in here and see whats wrong! - the Data,\n\n\n," + str(error_message))
            QtGui.QMessageBox.question(self, 'Message',"Your sweep failed :/ I also sent you an email.", QtGui.QMessageBox.Ok)
            print error_message ##########################!ahoy

        absolute_value = abs((start_freq - end_freq)/step_freq)
        if not absolute_value == 0:
            upper = max(start_freq, end_freq)
            lower = min(start_freq, end_freq)
            self.axesData.set_xlim(lower,upper)
        self.axesData.plot(graphX,graphY, 'bo-', linewidth = .5)
        self.canvasData.draw()

    #***********************************************************#
    #              Get Values from the GUI                      #
    #***********************************************************#
    #gets the entered values so that we can set parameters
    def startButtonClicked(self):
        global start_freq, end_freq, step_freq, email_notification
        global to_addresses, subject, message, run_sweep_bool
        global total_scan_time, elapsed_scan_time

        print "start clicked" #!ahoy
        errorFlag = 0
        warning = ""
        if self.myWidget.StartFreqTxt.text() == "":
            errorFlag = 1
            warning += "Please enter a start frequnecy.\n"
        else:
            start_freq = float(self.myWidget.StartFreqTxt.text())
        if self.myWidget.EndFreqTxt.text() == "":
            errorFlag = 1
            warning += "Please enter an end frequnecy.\n"
        else:
            end_freq = float(self.myWidget.EndFreqTxt.text())
        if self.myWidget.StepSizeTxt.text() == "":
            errorFlag = 1
            warning += "Please enter a step size.\n"
        else:
            step = float(self.myWidget.StepSizeTxt.text())
        if self.myWidget.EmailNotificationTxt.text() == "":
            errorFlag = 1
            warning += "Please enter a 1 (yes) or 0 (no) for email notification.\n"
        else:
            email_notification = int(self.myWidget.EmailNotificationTxt.text())
        if self.myWidget.ToAddressesTxt.text() == "" and email_notification == 1:
            errorFlag = 1
            warning += "Please enter an email address or addresses as a comma separated list.\n"
        else:
            string = self.myWidget.ToAddressesTxt.text()
            string = string.split(',')
            to_addresses = string
        subject = self.myWidget.SubjectTxt.text()
        message = self.myWidget.MessageTxt.text()

        if errorFlag == 1:
            self.ErrorCaughtWithParams(warning)
        else:
            run_sweep_bool = True

    #***********************************************************#
    #               Error Handling params                       #
    #***********************************************************#
    def ErrorCaughtWithParams(self,error_message):
        reply = QtGui.QMessageBox.question(self, 'Message', error_message, QtGui.QMessageBox.Ok)
        if reply == QtGui.QMessageBox.Ok:
            return 1
        else:
            return 0

#******************************************************#
#                No-Tune AutoSweep Script              #
#******************************************************#
#def NoTuneAutoSweep():
#    global start_freq, end_freq, step_freq, email_notification, to_addresses, subject, message
#    global run_sweep_bool, elapsed_scan_time, total_scan_time, sweep_finished_bool, abort_bool, start_time
#
#    NTNMR = win32com.client.Dispatch("NTNMR.Application")
#    # get the open file/path
#    open_file = NTNMR.GetActiveDocPath
#    # snag the path
#    path_name = os.path.dirname(open_file)
#    path_name_out = path_name + "\\stack"
#    if not os.path.exists(path_name_out):
#        os.makedirs(path_name_out)
#    #temp value is used to calc total_scan_time
#    OneScanTime = 0
#    # step through freq in NTNMR
#    numberfreqs = int(2 + (end_freq - start_freq)/step_freq)
#    integration_file = open(path_name + "\\IntegrationResults.txt", "w")
#    integration_file.write("RealWave, ImagWave, MagWave, FreqWave\n")
#
#    #iterate through the correct number of steps
#    for i in range(1, numberfreqs):
#        if abort_bool:
#            abort_bool = False
#            return
#        #if this is the first iteration start a clock to estimate total_scan_time
#        if i == 1:
#            start_time = time.time()
#
#        NTNMR.SetNMRParameter("Observe Freq.", start_freq + (i-1)*step_freq )
#        NTNMR.ZG
#        check = False
#        #check to see if we are done
#        while not check:
#            time.sleep(1)
#            check = NTNMR.CheckAcquisition
#        #check to see if the run was canceled by the user
#
#        #perfom calculations and make sure that the file has the correct name
#        iterfreqfloat = float(start_freq + (i - 1)*step_freq)
#        iterfreqstr = "%07.3f" % iterfreqfloat
#        file_name_out = path_name_out + "\\notuneAS_" + iterfreqstr + "MHz" + ".tnt"
#        NTNMR.SaveAs(file_name_out)
#        #get parameters in order to integrate the magnitude for dynamic graphing
#        intStart = NTNMR.GetCursorPosition
#        intEnd = NTNMR.Get1DSelectionEnd
#        nmrdata = NTNMR.GetData
#        #initialize the real and imaginary parts of the data
#        realTotal = 0
#        imagTotal = 0
#        #iterate through to find the real and imag part of the data from TNMR
#        for j in range(intStart, intEnd*2 - 1, 2):
#            realTotal += nmrdata[j]/NTNMR.GetNMRParameter("Scans 1D")
#            imagTotal += nmrdata[j+1]/NTNMR.GetNMRParameter("Scans 1D")
#        #calcualte the mag_total
#        magTotal = math.sqrt(pow(realTotal,2) + pow(imagTotal,2))
#        #add the values to the data to be graphed
#        x_int = float((start_freq + (numberfreqs-1) * step_freq)/1000)
#        graphY.append(magTotal)
#        graphX.append(x_int)
#        integration_file.write("%s, %s, %s, %s\n" %(realTotal, imagTotal, magTotal, x_int))
#        #finally close the file since we are done
#        NTNMR.CloseFile(file_name_out)
#
#        #if this is the first iteration calculate total_scan_time
#        if i == 1:
#            OneScanTime = time.time() - start_time
#            total_scan_time = temp*(numberfreqs-2)
#
#    sweep_finished_bool = True
#    en = email_notifier()
#    if email_notification == 1:
#        en.sendemail(to_addresses, subject, message)
#    integration_file.close()

#***********************************************************#
#                Class for Emailing                         #
#***********************************************************#

class email_notifier():
    def __init__(self):
        self.fromaddress = '"Felix Bloch" <relaxedfelix@gmail.com>'
        self.username = 'relaxedfelix'
        self.password = 'alertshirazpython'
        self.server = smtplib.SMTP('smtp.gmail.com:587')

    def sendemail(self, to_addresses, subject, message):
        COMMASPACE = ', '
        msg = MIMEText(message)
        msg['Subject'] = subject
        msg['From'] = self.fromaddress
        msg['To'] = COMMASPACE.join(to_addresses)
        self.server.starttls()
        self.server.login(self.username, self.password)
        self.server.sendmail(self.fromaddress, to_addresses, msg.as_string())
        self.server.quit()

#**********************************#
#                testy             # !ahoy
#**********************************#
def testy():
    global elapse_bool, run_sweep_bool, start_time, abort_bool
    global total_scan_time, elapsed_scan_time, sweep_finished_bool
    global start_freq, end_freq, step_freq, email_notification, to_addresses, subject, message

    elapse_bool = True
    numberfreqs = int(2 + (end_freq - start_freq)/step_freq)
    for i in range(1, numberfreqs):
        if i == 1:
            start_time = time.time()
        if abort_bool:
            print "aborting"
            return
        else:
            print "  i=" + str(i)
            time.sleep(1.5)
        if i == 1:
            OneScanTime = time.time() - start_time
            total_scan_time = OneScanTime*(numberfreqs - 2)
    elapse_bool = False
    sweep_finished_bool = True
    en = email_notifier()
    if email_notification == 1:
        en.sendemail(to_addresses, subject, message)
    total_scan_time = 0
    print "sweep_finished_bool = " + str(sweep_finished_bool)
    print "elapsed_scan_time = " + str(elapsed_scan_time)
    print "total_scan_time = " + str(total_scan_time)
    print "abort_bool = " + str(abort_bool)
    print "start_time = " + str(start_time)
    print "run_sweep_bool = " + str(run_sweep_bool)
    print "elapse_bool = " + str(elapse_bool)
    elapse_bool = False
#***********************************************************#
#                    Execute Main                           #
#***********************************************************#
def main():
    app = QtGui.QApplication(sys.argv)
    ex = NoTuneAutoSweepGUI()
    app.exec_()

if __name__ == '__main__':
    main()
