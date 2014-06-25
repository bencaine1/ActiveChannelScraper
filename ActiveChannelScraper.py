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
res = eq_cursor.execute("""
select tester_nm from MaccorList
where active_val = 'y'
""").fetchall()

for row in res:
    path = '\\\\' + row[0].rstrip(' ') + '.24m.local\\Maccor\\System\\' + row[0].rstrip(' ') + '\\Active'
    paths[path] = row[0].rstrip(' ')


#paths = {
#r'\\maccor-a.24m.local\Maccor\System\MACCOR-A\Active':'Maccor-A',
#r'\\maccor-b.24m.local\Maccor\System\MACCOR-B\Active':'Maccor-B',
#r'\\maccor-m.24m.local\Maccor\System\MACCOR-M\Active':'Maccor-M',
#r'\\maccor-n.24m.local\Maccor\System\MACCOR-N\Active':'Maccor-N',
#r'\\maccor-o.24m.local\Maccor\System\MACCOR-O\Active':'Maccor-O'
#}

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

# Populate table
sys.stdout.write('Populating table')
for a in data:
    sys.stdout.write('.')
    cursor.execute("""
    insert into ActiveChannelLog (System, Channel, Filename)
    values (?,?,?);
    """, a.system, a.channel, a.filename)
print

#close up shop
cursor.close()
del cursor
cnxn.close()
eq_cursor.close()
del eq_cursor
eq_cnxn.close()