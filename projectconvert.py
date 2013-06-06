#!/usr/bin/python

import xlrd, time, csv

timeformat =  "%m/%d/%y %I:%M %p"

class sched_item ():
    def __init__ (self, name='', start='', end=''):
        self.name = name
        self.start = start
        self.end  = end
        self.level = 1
        self.src_id = 0
        self.src_dep = 0
        self.dst_id = ''
        self.dep = None
        self.color = "grey1"

    def sheet_init (self, sheet, row):
        self.name = sheet.cell_value(row,3)
        try:
            self.src_id  = int(sheet.cell_value(row,0))
        except ValueError:
            self.src_id  = 0
        try:
            self.src_dep = int(sheet.cell_value(row,7))
        except ValueError:
            self.src_dep = 0
        self.start   = reformat_time(sheet.cell_value(row,5))
        self.end     = reformat_time(sheet.cell_value(row,6))
        self.level   = int(sheet.cell_value(row,8))

    def set_dst_id (self,dst_id):
        self.dst_id = dst_id

    def get_csv (self):
        if self.dep == None:
            dep = ''
        else:
            dep = self.dep.dst_id
        return [self.name,self.start,self.end,self.dst_id,dep,self.color]

def reformat_time (str_in):
    try:
        ts = time.strptime(str_in, timeformat)
        rv = time.strftime("%m/%d/%y", ts)
    except:
        rv = str_in
    return rv

def open_book (filename, outname):
    book = xlrd.open_workbook (filename)
    ofh = open (outname, 'w')
    writer = csv.writer (ofh)
    sheet = book.sheet_by_index(0)
    plevel = 0
    hier = [1]
    src_taskid = {}
    dst_taskid = {}
    sitems = []
    colors = ["blue1","blue2","grey1","green1","purple1","red1","orange1"]
    curcolor = 0

    writer.writerow(["Name/Title","Start Date","End Date","WBS #","Precessors","Task Color"])
    writer.writerow(["Hardware Schedule","3/1/13","12/31/14","1","","grey1"])

    for row in range(1,sheet.nrows):
        s_item = sched_item()
        s_item.sheet_init (sheet, row)
        #taskname = sheet.cell_value(row,3)
        #duration = sheet.cell_value(row,4)
        #src_tasknum  = sheet.cell_value(row,0)
        #src_dep      = sheet.cell_value(row,7)
        src_taskid[s_item.src_id] = s_item
        sitems.append (s_item)
        try:
            start    = reformat_time(sheet.cell_value(row,5))
            end      = reformat_time(sheet.cell_value(row,6))
            level    = int(sheet.cell_value(row,8))
            if (level > plevel):
                hier.append (1)
            elif (level < plevel):
                hier = hier[:level+1]
                hier[-1] += 1
                curcolor = (curcolor+1) % len(colors)
            else:
                hier[-1] += 1
            s_item.color = colors[curcolor]

            hier_val = '.'.join(map(str,hier))
            s_item.set_dst_id (hier_val)
            dst_taskid[hier_val] = s_item
            #if start == '': start = '1/1/13'
            #if end == ''  : end = '1/1/13'
            #writer.writerow([taskname,start,end,hier_val])
            plevel = level
        except:
            pass

    # cross-reference all tasks to build dependencies
    for s_item in sitems:
        if s_item.src_dep != 0:
            s_item.dep = src_taskid[s_item.src_dep]

    for s_item in sitems:
        print repr(s_item.get_csv())
        writer.writerow (s_item.get_csv())

    ofh.close()

open_book("develop-schedule.xlsx", "develop-schedule.csv")

