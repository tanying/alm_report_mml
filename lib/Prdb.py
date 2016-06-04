#!/usr/bin/python

# coding=gbk
import codecs

import re
import sqlite3
import types
from Utils import *
import MySQLdb
import datetime, time
import email
import pyExcelerator

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from UserInfo import *
from email.mime.base import MIMEBase
from email import encoders
from email.mime.image import MIMEImage
from time import strftime, localtime
from Integrity import IntegrityClient

class PrStatic(UserInfo):
    def __init__(self,conf):
        self.now = datetime.datetime.today().date()
        self.toList = []
        self.ogList = []
        self.conf = conf.dumpConf()

    def addTo(self, mailAddr):
        self.toList.append(mailAddr)

    def initUserInfo(self):
        for line in file('%s/alm_report_teammember/conf/platform.conf' % getToolPath()):
            line = line.split('=')
            self.ogList.append({line[0].strip() : line[1][1:].strip()})

    def createPRSheet(self, worksheetOg, items):
        worksheetOg.write(0, 0, 'ID')
        worksheetOg.write(0, 1, 'Type')
        worksheetOg.write(0, 2, 'Project')
        worksheetOg.write(0, 3, 'Branch')
        worksheetOg.write(0, 4, 'Priority')
        worksheetOg.write(0, 5, 'Assigner')
        worksheetOg.write(0, 6, 'State')
        worksheetOg.write(0, 7, 'Summary')
        worksheetOg.write(0, 8, 'deadline')
        worksheetOg.write(0, 9, 'regression')
        worksheetOg.write(0, 10, 'comment_from_cea')
        worksheetOg.write(0, 11, 'Team')
        worksheetOg.write(0, 12, 'Function')
        lineNum = 1

        for item in items:
            try:
                # branch = item.branch
                # if self.projectName == branch:
                worksheetOg.write(lineNum, 3, item.branch)
                # else:
                #     continue
            except AttributeError:
                continue
            except AssertionError:
                branches = item.branch
                for branch in branches:
                    if self.projectName == branch:
                        worksheetOg.write(lineNum, 3, item.branch)

            worksheetOg.write(lineNum, 0, item.id)
            worksheetOg.write(lineNum, 2, item.project)
            worksheetOg.write(lineNum, 1, item.type)
            worksheetOg.write(lineNum, 4, item.priority)
            worksheetOg.write(lineNum, 5, item.assigned_user.email)
            worksheetOg.write(lineNum, 6, item.state)
            worksheetOg.write(lineNum, 7, item.summary)

            if 'Deadline' in item.keys():
                worksheetOg.write(lineNum, 8, item['Deadline'][1].value)

            if 'Regression' in item.keys():
                worksheetOg.write(lineNum, 9, item['Regression'][1].value)

            if 'All CEA comments' in item.keys():
                worksheetOg.write(lineNum, 10, item['All CEA comments'][1].value)

#            worksheetOg.write(lineNum, 9, item.comment_from_cea)
            worksheetOg.write(lineNum, 12, item.function)

            files = []
            files = os.listdir('%s/alm_report_mml/conf/TeamMember' % getToolPath())
            haveTeam = False
            for fileName in files:
                for line in file('%s/alm_report_mml/conf/TeamMember/%s' % (getToolPath(), fileName)):
                    if line.strip() == '':
                        break
                    line = line.split(':')
                    if cmp((line[1].strip()).lower(), (item.assigned_user.email).lower()) == 0:
                        worksheetOg.write(lineNum, 11, (str(fileName)).upper().decode("utf-8"))
                        haveTeam = True
                        break
            if haveTeam == False:
                worksheetOg.write(lineNum, 11, '#NA')

            lineNum += 1

    def genSheetName(self, projectName):
        r_index = projectName.rfind("/")
        sheetName = projectName[r_index+1:].replace(" ", "").strip('"')
        return sheetName

    def sendPRStatic(self, days_value):
        wsdl = "http://alm.tclcom.com:7001/webservices/10/2/Integrity/?wsdl"
        intclient = IntegrityClient(wsdl=wsdl,
                                    credential_username="scm_tools",
                                    credential_password="SCM_TOOLS123!")

        workbook = pyExcelerator.Workbook()

        fields=["ID","Type","Project","Priority","Assigned User","State","Summary","Deadline","Regression","All CEA comments","Function"]
        allProjects = self.conf['integrity_project']
        print allProjects #tanya_log
        ProjectList = self.conf['integrity_project'].split(',')
        ProjectList.sort()

        # worksheetOg = workbook.add_sheet("MML_all_projects")
        # query = ('((field[Type]="Defect","Task") and ' \
        #          '(field[Project]=%s) and '\
        #          '(field[Assigned User]="ying.tan","tyin","wenhua.tu","sichao.hu","yuanxing.tan","v-nj-nie.lei","jianying.zhang","xuan.zhou"))')%(allProjects)
        # items = intclient.getItemsByCustomQuery(fields=fields, query=query)
        # self.createPRSheet(worksheetOg, items)

        for projectName in ProjectList:
            print projectName #tanya_log
            sheetName = self.genSheetName(projectName)
            worksheetOg = workbook.add_sheet(sheetName)
            query = ('((field[Type]="Defect","Task") and ' \
                     '(field[Project]=%s) and '\
                     '(field[Assigned User]="ying.tan","tyin","wenhua.tu","sichao.hu","yuanxing.tan","v-nj-nie.lei","jianying.zhang","xuan.zhou"))')%(projectName)
#                      '(field[Project]="/TCT/GApp/CameraL","/TCT/MTK MT6755M/X1 PLUS","/TCT/QCT MSM8952/Idol4","/TCT/QCT MSM8976/Idol4 S","/TCT/QCT MSM8976/Idol4 S VF") and ' \
            items = intclient.getItemsByCustomQuery(fields=fields, query=query)
            self.createPRSheet(worksheetOg, items)

            if days_value != 0:#query date setting
                now = datetime.datetime.today().date()
                report_days = now.strftime('%Y-%m-%d')
                date_days = datetime.timedelta(days=days_value)
                last = now - date_days
                report_days = '(' + last.strftime('%Y-%m-%d') + '~' + now.strftime('%Y-%m-%d') + ')'

                datetime_format = '%b %d, %Y %I:%M:%S %p'
                date_days = datetime.timedelta(days=days_value)
                date = datetime.datetime.today().date() - date_days
                start_str = date.strftime(datetime_format)
                end_str = now.strftime(datetime_format)
                print "start time: ",start_str
                print "end time: ",end_str

                worksheetOg = workbook.add_sheet(sheetName+ "_" + report_days)
                query = ('((field[Type]="Defect","Task") and ' \
                          '(field[Project]= %s) and ' \
                          '(field[Assigned User]="ying.tan","tyin","wenhua.tu","sichao.hu","yuanxing.tan","v-nj-nie.lei","jianying.zhang","xuan.zhou") and' \
                          '(histdate[State] was changed between time %s and %s))')%(projectName, start_str, end_str)
                items = intclient.getItemsByCustomQuery(fields=fields, query=query)
                self.createPRSheet(worksheetOg, items)
        print "ongoing pr export from alm done!"

        workbook = workbook.save('/local/reportpr/pr.xls')
        print "sendPRStatic done!"

