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
'\\\\maccor-a\\Maccor\\System\\MACCOR-A\\Active':'Maccor-a',
'\\\\maccor-b\\Maccor\\System\\MACCOR-B\\Active':'Maccor-b',
'\\\\maccor-m\\Maccor\\System\\MACCOR-M\\Active':'Maccor-m',
'\\\\maccor-n\\Maccor\\System\\MACCOR-N\\Active':'Maccor-n',
'\\\\maccor-o\\Maccor\\System\\MACCOR-O\\Active':'Maccor-o'
}

data = []
for p in paths.keys():
    for f in os.listdir(p):
        if os.path.isfile(os.path.join(p,f)):            
            system, channel, filename = ('',)*3
            filename = f
            m = re.search('\.([0-9]){3}', f)
            if m:
                channel = m.group()
            system = paths[p] # e.g. 'Maccor-a'
            a = ActiveChannelDataPt(system, channel, filename)
            data.append(a)

#for a in data:
#    print a
    
for a in data:
    cursor.execute("""
    merge ActiveChannelLog as T
    using (select ?,?,?) as S (System, Channel, Filename)
    on S.Filename = T.Filename
    when not matched then insert(System, Channel, Filename)
    values (S.System, S.Channel, S.Filename);
    """, a.system, a.channel, a.filename)
