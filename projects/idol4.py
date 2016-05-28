#!/usr/bin/python
from Config import *
from UserInfo import *
from Prdb import *
from dataCollection import *

class project(UserInfo, Config):
    def __init__(self, conf):
        self.conf = conf
        self.full_name = self.conf.getConf('full_name', 'full name for PrStatic')
        self.integrity_project = self.conf.getConf('integrity_project', 'Integrity Project')

    def run(self, conf):
        prdb = PrStatic(self.full_name)
        teamCollection = dataCollection(self.full_name)
        if self.getConf('mailto', 'Receiver of this mail <all|self>') == 'all':
            teamCollection.addTo('<%s>'%(self.conf.getConf('mail_list','mail list')))
        else:
            teamCollection.addTo('"\'%s\'" <%s>' % (self.getFullName(), self.getMail()))

        cc = self.conf.getConf('cc_list','cc list')
        cc_list = cc.split(',')
        for cc_mail in cc_list:
            clean_mail = cc_mail.strip()
            if clean_mail:
                teamCollection.addCc('<%s>'%(clean_mail))

        days = conf.getConf('days', 'Days')

        teamCollection.deleteNATeamMember('%s/alm_report_teammember/conf/TeamMember/#NA' % getToolPath())
        prdb.sendPRStatic(self.full_name, self.integrity_project, int(days))
        teamCollection.addNATeamMember('%s/alm_report_teammember/conf/TeamMember/#NA' % getToolPath())
        teamCollection.start(self.full_name, int(days))
