#!/usr/bin/python

import os
import sys
import re
from Utils import *
from Config import *
from UserInfo import *
from Prdb import *
from dataCollection import *

def main():
    conf = Config();
    conf.addFromArg(sys.argv[1:])

    user = conf.getConf('user', 'User name')

    env = UserInfo()
    env.initUserInfo(user)

    conf.loadConfigFromFile(os.path.dirname(__file__)+'/conf/mml.conf')

    # if projname != '__init__' and os.path.isfile(os.path.dirname(__file__)+"/projects/"+projname+".py"):
    #     projMod = __import__('projects.'+projname, globals(), locals(), '*')
    # else: 
    #     from projects import BaseProject as projMod

    # proj = projMod.project(conf)
    try:
        run(conf)
    except SystemExit:
        raise

def run(conf):
        prdb = PrStatic(conf)
        projectCollection = dataCollection()
        # if conf.getConf('mailto', 'Receiver of this mail <all|self>') == 'all':
        #     projectCollection.addTo('<%s>'%(conf.getConf('mail_list','mail list')))
        # else:
        #     projectCollection.addTo('"\'%s\'" <%s>' % (self.getFullName(), self.getMail()))

        cc = conf.getConf('cc_list','cc list')
        cc_list = cc.split(',')
        for cc_mail in cc_list:
            clean_mail = cc_mail.strip()
            if clean_mail:
                projectCollection.addCc('<%s>'%(clean_mail))

        days = conf.getConf('days', 'Days')

        projectCollection.deleteNATeamMember('%s/alm_report_teammember/conf/TeamMember/#NA' % getToolPath())
        #prdb.sendPRStatic(int(days))#Export the bugs from ALM.
        projectCollection.addNATeamMember('%s/alm_report_teammember/conf/TeamMember/#NA' % getToolPath())
        projectCollection.start(conf, int(days))

if __name__ == '__main__':
    main()
