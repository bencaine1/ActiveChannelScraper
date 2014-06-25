# -*- coding: utf-8 -*-
"""
Created on Fri Jun 13 10:01:28 2014

@author: bcaine
"""

import pyodbc
import os
import re
import sys
#import pywin32

class ActiveChannelDataPt:
    def __init__(self, system, channel, filename, status):
        self.system = system
        self.channel = channel
        self.filename = filename
        self.status = status
    def __str__(self):
        s = 'System: ' + self.system + '\n'
        s += 'Channel: ' + str(self.channel) + '\n'
        s += 'Filename: ' + self.filename + '\n'
        s += 'Status: ' + self.status + '\n'
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

eq_cnxn_str = """
Driver={SQL Server Native Client 11.0};
Server=172.16.111.235\SQLEXPRESS;
Database=EquipmentInfo;
UID=sa;
PWD=Welcome!;
"""
eq_cnxn = pyodbc.connect(eq_cnxn_str)
eq_cnxn.autocommit = True
eq_cursor = eq_cnxn.cursor()

paths = {}
status_paths = {}

res = eq_cursor.execute("""
select tester_nm from MaccorList
where active_val = 'y'
""").fetchall()

for row in res:
    path = '\\\\' + row[0].rstrip(' ') + '.24m.local\\Maccor\\System\\' + row[0].rstrip(' ') + '\\Active'
    paths[path] = row[0].rstrip(' ')
    status_paths[row[0].rstrip(' ')] = '\\\\' + row[0].rstrip(' ') + '.24m.local\\Maccor\\RemoteCtrl'

data = []
for p in paths.keys():
    for f in os.listdir(p):
        if os.path.isfile(os.path.join(p,f)):
            system, channel, filename, status = ('',)*4
            filename, sep, channel = f.rpartition('.')
            try:
                channel = int(channel)
            except:
                print 'non-integer extension for file ', f
            system = paths[p] # e.g. 'Maccor-A'
            for sf in os.listdir(status_paths[paths[p]]):
                s_filename, s_sep, s_channel = sf.rpartition('.')
                if s_channel == str(channel).zfill(4):
                    if 'complete' in s_filename.lower():
                        status = 'complete'
                    elif 'running' in s_filename.lower():
                        status = 'running'
            a = ActiveChannelDataPt(system, channel, filename, status)
            data.append(a)

#for p in status_paths:
    

#for a in data:
#    print a
    
# Delete old records
cursor.execute("""
DELETE FROM ActiveChannelLog;
""")

# Populate table
sys.stdout.write('Populating table')
for a in data:
    sys.stdout.write('.')
    cursor.execute("""
    insert into ActiveChannelLog (System, Channel, Filename, Status)
    values (?,?,?,?);
    """, a.system, a.channel, a.filename, a.status)
print

#close up shop
cursor.close()
del cursor
cnxn.close()
eq_cursor.close()
del eq_cursor
eq_cnxn.close()