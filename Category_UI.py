# -*- coding: utf-8 -*-
'''
This a simple desktop UI fro Charter category check. This UI have one input area , one output areas and one select
button , one process button and one clean button. The select button will fork a window selector where you can pick
your .xlsx file which needs to be process and the same path will show you on input area. Once you selected the file
you need to press process button which will take the .xlsx file as input and filter the duplicates out of the. Once
duplicates are terminated it will generate another .xlsx file at the same location where your .exe is present. It
will then open the log files linked in input files and fetch all the categories associate with the titles.Then all
the catagories will check in vosia DB . Once all the check completed it will provide you the list of catagories needs to create
along with the package id and title name for each provider/distributor .

It will also create a category_{date}.log which will have all the catagories which are found in DB and not found in DB.

NOTE:

Please download "instantclient-basiclite-nt-12.2.0.1.0.zip" from web and unzip it. Then place 'instantclient_12_2'
directory under 'C:\\oracle' . So the final path will be "C:\\oracle\\instantclient_12_2". This is a library use to
connect oracle server through python.

TODO:
-----
1>Directly download the file from the website.
2>Convert xls file to xlsx internally and then execute.

COMMAND:
--------
Category_UI.exe

python: 2.7
Date: 08/26/2018
Name: JPD
'''


from PyQt4 import QtGui
import sys, os, openpyxl
import requests, re, datetime, logging

log_file = 'category_' + str(datetime.datetime.now().strftime('%m%d%Y')) + ".log"
logging.basicConfig(level=logging.DEBUG,filename=log_file,format='%(asctime)s - %(name)s - %(levelname)s - %('
                                                                'message)s',filemode='w')
logger = logging.getLogger(__name__)

session = requests.Session()
session.auth = ('d3noc_charter','g8tk33pp3r')
session.post('http://charter.vodera.bydeluxe.com')

os.environ['PATH'] += ';' + "C:\\oracle\\instantclient_12_2"
logger.info('Extending PATH with path {}'.format("C:\oracle\instantclient_12_2"))

import cx_Oracle



#handler = logging.FileHandler('log_file')
#handler.setLevel(logging.INFO)
#formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
#handler.setFormatter(formatter)
#logger.addHandler(handler)


con = cx_Oracle.connect('VTMS/h4mm3r@10.42.231.206:1521/CHARTER')
cursor = con.cursor()

exclude = ['Cannot reschedule an already scheduled package.']
reporting = ['Cannot schedule invalid versions package.','Attempting to delete a package which does not exist.','Unable to create new package version: Package is already marked for deletion.']
toCheck = [r'The import was rejected in strict mode with the following error codes: Unable to create new schedule '
           r'entry: '
           'Provider NameProvider IDRating and Category fields are required for automatic scheduling.Unknown <Category> '
           'in Title asset.',r'/error/vtms/services/import-assets/cl-11/title/ec-9',r'Unknown <Category> in Title '
                                                                                    'asset.',
           'Unknown <Category> in Title asset.Unknown <Category> in Title asset.',
           r'The import was rejected in strict mode with the following error codes: Unable to create new schedule entry: '
           r'Provider NameProvider IDRating and Category fields are required for automatic scheduling.Unknown <Category> '
           r'in Title asset.Unknown <Category> in Title asset.']
metadIssue = ['Unable to create new schedule entry: Provider NameProvider IDRating and Category fields are required '
              'for automatic scheduling.','WorkflowException: Unable to update WorkflowInstance.']

category_check = {}
reporting_check = {}


class Window(QtGui.QMainWindow):
    def __init__(self):
        super(Window, self).__init__()
        self.setGeometry(50, 50, 700, 500)
        self.setWindowTitle('Category')
        self.setWindowIcon(QtGui.QIcon('favicon.png'))

        self.inBox = QtGui.QTextEdit(self)
        self.inBox.setGeometry(10, 10, 500, 30)

        self.brw = QtGui.QPushButton(self)
        self.brw.setText("Browse")
        self.brw.resize(self.brw.sizeHint())
        self.brw.move(520, 15)
        self.brw.clicked.connect(self.onBrowse)

        self.outBox = QtGui.QTextEdit(self)
        self.outBox.setGeometry(10, 50, 500, 400)

        self.sub = QtGui.QPushButton(self)
        self.sub.setText("Process")
        self.sub.resize(self.brw.sizeHint())
        self.sub.move(600, 15)
        self.sub.clicked.connect(self.onProcess)


        self.clean = QtGui.QPushButton(self)
        self.clean.setText("Clean")
        self.clean.resize(self.brw.sizeHint())
        self.clean.move(520, 50)
        self.clean.clicked.connect(self.outBox.clear)


        self.show()


    def onBrowse(self):

        self.file_path = QtGui.QFileDialog.getOpenFileName(parent=self, caption='Open file')

        if self.file_path:
            self.inBox.setText(self.file_path)


    def onProcess(self):

        try:
            inFile = self.file_path
            logger.info('File to process {}'.format(inFile))
            self.inBox.clear()

        except IOError, e:
            QtGui.QMessageBox.question(self, 'Want to Process!!', 'Please make sure you provided proper file '
                                                                           'path!!', QtGui.QMessageBox.Ok)
            logger.error('Failed to open the file  {} '.format(e))
        except AttributeError, e:
            self.outBox.append('Please select a file using Browse button!!')
            logger.error('AttributeError:Please select a file using Browse button {}!!'.format(e))
            self.inBox.clear()

        try:
            self.setWindowTitle('Category Processing...')
            logger.info('Processing start')
            wb = openpyxl.load_workbook(filename=str(inFile))
        except openpyxl.utils.exceptions.InvalidFileException:
            self.outBox.append('Please provide a .xlsx file for processing!!')
            logger.warning('User provided other than .xlsx file :{} '.format(inFile))
            logger.critical('Exiting!!')
            #sys.exit(0)
        except IOError:
            self.outBox.append('File is not available !!!')
            logger.error('User provided a file {} which is not present or not able to open'.format(inFile))
            logger.critical('Exiting!!')
            #sys.exit(1)


        sheet = wb.worksheets[0]
        header = ['Log File Name', 'Event Date', 'Title', 'Package Asset ID', 'Version', 'Package Description',
                  'Title Asset ID', 'Billing ID', 'Distributor', 'Error Message', 'Provider ID', 'Start', 'End', 'PCT']
        out_path = "SHEIngestErrors-{}.{}".format(datetime.datetime.now().strftime('%m%d%Y-%H%M%S'), 'xlsx')
        book = openpyxl.Workbook()
        bsheet = book.active
        #col_count = sheet.max_column
        row_count = sheet.max_row

        data = {}
        for rows in range(6, row_count + 1):
            link = sheet.cell(row=rows, column=1).hyperlink.target
            if sheet.cell(row=rows, column=4) in data:
                if data[sheet.cell(row=rows, column=4).value][1] < sheet.cell(row=rows, column=2):
                    data[sheet.cell(row=rows, column=4).value] = [link, sheet.cell(row=rows, column=2).value,
                                                                  sheet.cell(row=rows, column=3).value,
                                                                  sheet.cell(row=rows, column=4).value,
                                                                  sheet.cell(row=rows, column=5).value,
                                                                  sheet.cell(row=rows, column=6).value,
                                                                  sheet.cell(row=rows, column=7).value,
                                                                  sheet.cell(row=rows, column=8).value,
                                                                  sheet.cell(row=rows, column=9).value,
                                                                  sheet.cell(row=rows, column=10).value,
                                                                  sheet.cell(row=rows, column=11).value,
                                                                  sheet.cell(row=rows, column=12).value,
                                                                  sheet.cell(row=rows, column=13).value,
                                                                  sheet.cell(row=rows, column=14).value]
            else:
                data[sheet.cell(row=rows, column=4).value] = [link, sheet.cell(row=rows, column=2).value,
                                                              sheet.cell(row=rows, column=3).value,
                                                              sheet.cell(row=rows, column=4).value,
                                                              sheet.cell(row=rows, column=5).value,
                                                              sheet.cell(row=rows, column=6).value,
                                                              sheet.cell(row=rows, column=7).value,
                                                              sheet.cell(row=rows, column=8).value,
                                                              sheet.cell(row=rows, column=9).value,
                                                              sheet.cell(row=rows, column=10).value,
                                                              sheet.cell(row=rows, column=11).value,
                                                              sheet.cell(row=rows, column=12).value,
                                                              sheet.cell(row=rows, column=13).value,
                                                              sheet.cell(row=rows, column=14).value]
        wrow = 2

        for x in range(len(header)):
            bsheet.cell(row=1, column=x + 1).value = header[x]
            bsheet.cell(row=1, column=x + 1).font = openpyxl.styles.Font(bold=True, color='337AFF')

        for keys in data:
            link = data[keys][0]
            name = data[keys][0].split("=")[1]
            bsheet.cell(row=wrow, column=1).value = '=HYPERLINK("' + link + '","' + name + '")'
            bsheet.cell(row=wrow, column=1).style = 'Hyperlink'
            bsheet.cell(row=wrow, column=2).value = data[keys][1]
            bsheet.cell(row=wrow, column=3).value = data[keys][2]
            bsheet.cell(row=wrow, column=4).value = data[keys][3]
            bsheet.cell(row=wrow, column=5).value = data[keys][4]
            bsheet.cell(row=wrow, column=6).value = data[keys][5]
            bsheet.cell(row=wrow, column=7).value = data[keys][6]
            bsheet.cell(row=wrow, column=8).value = data[keys][7]
            bsheet.cell(row=wrow, column=9).value = data[keys][8]
            bsheet.cell(row=wrow, column=10).value = data[keys][9]
            bsheet.cell(row=wrow, column=10).alignment = openpyxl.styles.Alignment(wrapText=True)
            bsheet.cell(row=wrow, column=11).value = data[keys][10]
            bsheet.cell(row=wrow, column=12).value = data[keys][10]
            bsheet.cell(row=wrow, column=13).value = data[keys][10]
            bsheet.cell(row=wrow, column=14).value = data[keys][10]
            wrow += 1

        book.save(out_path)
        self.outBox.append('Processed file generated {}'.format(out_path))
        logger.info('New file created at mentioned location: {}'.format(out_path))
        notinDb = {}
        categoryDict = {}
        meta_error = {}
        logger.info('Fatching catagories from each title from log!!')
        for d in data:
            if data[d][9] in exclude:
                continue
            elif data[d][9] in reporting:
                if data[d][9] in reporting_check:
                    reporting_check[data[d][9]].append([data[d][2],data[d][3],data[d][8]])
                else:
                    reporting_check[data[d][9]] = [[data[d][2],data[d][3],data[d][8]]]
            elif data[d][9] in toCheck:
                for line in session.get(data[d][0]).text.splitlines():
                    if re.search(r'Name="Category"',line):
                        category = (line.split("=")[3].split("\"")[1]).replace('&amp;','&')
                try:
                    category = str(category)
                except UnicodeEncodeError:
                    self.outBox.append(category + ' '+  data[d][3] +' '+ data[d][8])
                    continue
                if data[d][8] in categoryDict:
                    if category in categoryDict[data[d][8]].keys():
                        categoryDict[data[d][8]][category].append([data[d][3],data[d][2]])
                    else:
                        categoryDict[data[d][8]][category] = [[data[d][3],data[d][2]]]
                else:
                    categoryDict[data[d][8]] = {category:[[data[d][3],data[d][2]]]}
            elif data[d][9] in metadIssue:
                if data[d][8] in meta_error:
                    if data[d][9] in meta_error[data[d][8]]:
                        meta_error[data[d][8]][data[d][9]].append([str(data[d][3]),str(data[d][2])])
                    else:
                        meta_error[data[d][8]][data[d][9]] = [[str(data[d][3]),str(data[d][2])]]
                else:
                    meta_error[data[d][8]] = {data[d][9] : [[str(data[d][3]),str(data[d][2])]]}
        for dictributor in categoryDict:
            for category in categoryDict[dictributor]:
                logger.info("select * from CATEGORIES where NAME IN {}".format(category))
                catagoriesDB_check = cursor.execute("select * from CATEGORIES where NAME IN :cat ",
                                                   {'cat': category})
                dbExtract = catagoriesDB_check.fetchall()
                if dbExtract:
                    for d in dbExtract:
                        category_id = d[0]
                        logger.info('select * from SERVICE_CLIENT_CATEGORIES where CATEGORY_ID = {}'.format(
                            category_id))
                        srv_clt_chk = cursor.execute("select * from SERVICE_CLIENT_CATEGORIES where CATEGORY_ID = :cid",
                                                     {'cid' : str(category_id)})
                        if srv_clt_chk.fetchall():
                            logger.info("select * from CATEGORY_PLATFORMS "
                                        "where CATEGORY_ID = {} "
                                        "and PLATFORM_ID = 'CHARTER-CL11'".format(category_id))
                            ctg_pltf_chk = cursor.execute("select * from CATEGORY_PLATFORMS "
                                                          "where CATEGORY_ID = :cid and PLATFORM_ID = "
                                                          "'CHARTER-CL11'", {'cid':str(category_id)})
                            ctg_pltf_data =  ctg_pltf_chk.fetchone()
                            if ctg_pltf_data:
                                logger.info(category + " <--- found in DB")
                            else:
                                logger.info("Not found in CATEGORY_PLATFORMS ---> {}".format(category))
                        else:
                            logger.info("Not found in SERVICE_CLIENT_CATEGORIES ---> {}".format(category))
                else:
                    if dictributor in notinDb:
                        notinDb[dictributor].append(category)
                    else:
                        notinDb[dictributor]= [category]
        self.outBox.append('Catagories having issue!!')
        for dist in notinDb:
            pack = {}
            self.outBox.append('')
            self.outBox.append(dist)
            self.outBox.append('Category')
            self.outBox.append(' ')
            for cat in notinDb[dist]:
                self.outBox.append(cat)
            self.outBox.append(' ')
            self.outBox.append('Assets')
            for c in xrange(len(notinDb[dist])):
                cat = notinDb[dist][c]
                for d in  categoryDict[dist][cat]:
                    pack[d[0]] = d[1]
            for k in pack.keys():
                self.outBox.append(k + "\t" + pack[k])
        self.setWindowTitle('Category Completed...')

        for dist in meta_error:
            self.outBox.append('')
            self.outBox.append(dist + ":")
            self.outBox.append('')
            for d in meta_error[dist]:
                self.outBox.append('Error : {}'.format(d))
                self.outBox.append('')
                for l in meta_error[dist][d]:
                    self.outBox.append(l[0] +'\t' + l[1])

        for reportErr in reporting_check:
            self.outBox.append('')
            self.outBox.append('Assets needs to report :')
            self.outBox.append('Error: {}'.format(reportErr))
            self.outBox.append('')
            self.outBox.append('Assets')
            for l in reporting_check[reportErr]:
                self.outBox.append('{} \t {}'.format(l[1],l[0]))
        logger.info('Category processing Completed!!!')


def run():
    app = QtGui.QApplication(sys.argv)
    GUI = Window()
    sys.exit(app.exec_())

run()
