#!/usr/bin/python

import os
import sys
import re
from Utils import *
from Config import *
from UserInfo import *

def main():
    conf = Config();
    conf.addFromArg(sys.argv[1:])

    user = conf.getConf('user', 'User name')

    env = UserInfo()
    env.initUserInfo(user)

    projname = conf.getConf('project', 'Project name')
    conf.loadConfigFromFile(os.path.dirname(__file__)+'/conf/'+projname+'.conf')

    if projname != '__init__' and os.path.isfile(os.path.dirname(__file__)+"/projects/"+projname+".py"):
        projMod = __import__('projects.'+projname, globals(), locals(), '*')
    else: 
        from projects import BaseProject as projMod

    proj = projMod.project(conf)
    try:
        proj.run(conf)
    except SystemExit:
        raise

if __name__ == '__main__':
    main()
