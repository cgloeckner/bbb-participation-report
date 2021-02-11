#!/usr/bin/python3 
# -*- coding: utf-8 -*-

import os, pathlib
import xlsxwriter


class Conference(object):
    def __init__(self):
        self.title = None
        self.names = set()
        self.date  = None

    def loadFromFile(self, fname):
        with open(fname, 'r') as h:
            content = h.read()
        # parse meta
        line1 = content.split('\n')[0]
        self.title = line1.split('Konferenz ')[1].split(' um')[0]
        self.date = list()
        for d in line1.split('um ')[1].split(':')[0].split('.'):
            self.date.append(int(d))
        self.date.reverse()

        # parse participation
        names = set()
        region = content.split('Sortiert nach Nachname:\n')[1]
        for name in region.split('\n'):
            if name == '':
                continue
            last_name = name.split(' ')[-1]
            first_name = ' '.join(name.split(' ')[:-1])
            self.names.add((last_name, first_name))


# ---------------------------------------------------------------------

class Course(object):
    def __init__(self, name):
        self.name        = name
        self.conferences = list()

    def loadFromDirectory(self, root):
        p = pathlib.Path(root)
        for fname in os.listdir(root):
            c = Conference()
            c.loadFromFile(p / fname)
            self.conferences.append(c)

    def getAllNames(self):
        names = set()
        for c in self.conferences:
            names = names.union(c.names)
        l = list(names)
        l.sort(key=lambda elem: elem[0])
        return l

    def makeSheet(self, doc):
        sheet = doc.add_worksheet(self.name)
        sheet.set_column(0, 0, 30)
        sheet.write(0, 0, 'Schüler')
        sheet.set_column(1, 1 + len(self.conferences), 5)

        self.conferences.sort(key=lambda c: c.date)
        for i, c in enumerate(self.conferences):
            sheet.write(0, i+2, '{0}.{1}'.format(c.date[2], c.date[1]))
        
        for i, name in enumerate(self.getAllNames()):
            sheet.write(i+2, 0, '{0}, {1}'.format(name[0], name[1]))
            for j, c in enumerate(self.conferences):
                if name in c.names:
                    sheet.write(i+2, j+2, 'X')


# ---------------------------------------------------------------------        

class Analysis(object):
    def __init__(self):
        self.courses = list()

    def loadFromDirectory(self, root):
        p = pathlib.Path(root)
        for fname in os.listdir(root):
            if fname in ['__pycache__']:
                continue
            if os.path.isdir(p / fname): 
                c = Course(fname)
                c.loadFromDirectory(p / fname)
                self.courses.append(c)
        
    def saveToFile(self, fname):
        doc = xlsxwriter.Workbook(fname)
        for c in self.courses:
            c.makeSheet(doc)
        doc.close()


# ---------------------------------------------------------------------        

def main():
    a = Analysis()
    a.loadFromDirectory('.')
    a.saveToFile('out.xlsx')


if __name__ == '__main__':
    main()
