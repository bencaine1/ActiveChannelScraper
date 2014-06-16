# -*- coding: utf-8 -*-
"""
Created on Fri Jun 13 10:01:28 2014

@author: bcaine
"""

import pyodbc
import os
import re

class ActiveChannelDataPt:
    def __init__(self, system, channel, filename):
        self.system = system
        self.channel = channel
        self.filename = filename
    def __str__(self):
        s = 'System: ' + self.system + '\n'
        s += 'Channel: ' + self.channel + '\n'
        s += 'Filename: ' + self.filename + '\n'
        return s

# connect to db
cnxn_str = """
Driver={SQL Server Native Client 11.0};
Server=172.16.111.235\SQLEXPRESS;
Database=CellTestData;
UID=sa;
PWD=Welcome!;
"""
cnxn = pyodbc.connect(cnxn_str)
cnxn.autocommit = True
cursor = cnxn.cursor()

paths = {
'\\\\maccor-a\\Maccor\\System\\MACCOR-A\\Active':'Maccor-A',
'\\\\maccor-b\\Maccor\\System\\MACCOR-B\\Active':'Maccor-B',
'\\\\maccor-m\\Maccor\\System\\MACCOR-M\\Active':'Maccor-M',
'\\\\maccor-n\\Maccor\\System\\MACCOR-N\\Active':'Maccor-N',
'\\\\maccor-o\\Maccor\\System\\MACCOR-O\\Active':'Maccor-O'
}

data = []
for p in paths.keys():
    for f in os.listdir(p):
        if os.path.isfile(os.path.join(p,f)):
            system, channel, filename = ('',)*3
            filename, sep, channel = f.rpartition('.')
            try:
                channel = int(channel)
            except:
                print 'non-integer extension for file ', f
#            m = re.search('\.(?P<number>[0-9]{3})', f)
#            if m:
#                channel = m.group('number')
            system = paths[p] # e.g. 'Maccor-A'
            a = ActiveChannelDataPt(system, channel, filename)
            data.append(a)

#for a in data:
#    print a
    
# Delete old records
cursor.execute("""
DELETE FROM ActiveChannelLog;
""")

for a in data:
    cursor.execute("""
    insert into ActiveChannelLog (System, Channel, Filename)
    values (?,?,?);
    """, a.system, a.channel, a.filename)
