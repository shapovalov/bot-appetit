#!/usr/bin/env python
# -*- coding: utf-8 -*- 

import sys
import pickle
import openpyxl
import arteclunch

def make_person_string(order, name, verbose=False):
    reply = [u"Here is the order for Mr/Ms "+name+u":"]
    
    for item in order:
        reply.append(item[0])
        if verbose and item[1]:
            reply.append("* " + item[1])

    return reply

if __name__=="__main__":
    if len(sys.argv) < 2:
        print "Incorrect format: pass at least the xls and pickle filename."
        sys.exit()

    fname = sys.argv[1]
    fname_out = sys.argv[2]
    wb = openpyxl.load_workbook(filename = fname, read_only=True)
    ws = wb.active

    splits = arteclunch.split_days(ws)
    persons = arteclunch.enumerate_persons(ws)

    db = {day : {} for day in splits}

    for name, col in persons:
    	for day in splits:
    		order = arteclunch.get_order(ws, splits, day, col)
    		reply = make_person_string(order, name.strip(), verbose=True)
    		db[day][name.strip()] = reply

    with open(fname_out, 'w') as f:
    	pickle.dump(db, f)
