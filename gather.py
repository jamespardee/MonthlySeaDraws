#!/usr/bin/env python
# -*- coding: utf-8 -*-


# Mac Dir Paths
# IN '/Users/jrp/Google Drive/Comp Prog/Python2/sealink_email'
# OUT '/Users/jrp/desktop/test'

import os as os
import shutil as shutil


class FileCollector:
    '''Copy from in_dir a type of fize(suffix) in a directory, group in order \
    by size and zip folders into out_dir'''
    def __init__(self, suffix='', size=3500000, in_dir='H:\draw_py\sea_files',
                 out_dir='H:\draw_out'):
        self.suffix = suffix
        self.size = size
        self.in_dir = in_dir
        self.out_dir = out_dir

    def locate(self):
        '''
        Search directory for files ending in "suffix".
        will begin  at in_dir and resursively go into all paths
        below that point.
        '''
        dic = {}
        for root, directory, files in os.walk(self.in_dir):
            for f in files:
                if self.suffix in f:
                    na = os.path.join(root, f).split('\\')[-1]
                    pa = os.path.join(root, f)
                    new_dir = os.path.join(root, f).split('\\')[-1][0:4].lower()
                    si = os.stat(os.path.join(root, f)).st_size
                    dic[na] = pa, new_dir, si
                else:
                    continue
            return dic

    def make_dir(self, dic):
        '''Make directorys to group the files with the same ####'s'''
        dir_made = []
        for pa, new_dir, si in dic.values():
            try:
                os.mkdir(os.path.join(self.out_dir, new_dir))
                dir_made.insert(10000, new_dir)
            except:
                continue
        return dir_made

    def sized(self, dic, dir_made):
        ''' Copy files into their correct folder & return the size assocaited
        with each folder'''
        dir_sm = []
        sm = 0
        for i in dir_made:
            for pa, new_dir, si in dic.values():
                if new_dir == i:
                    shutil.copy(pa, os.path.join(self.out_dir, i))
                    sm += si
                else:
                    continue
            dir_sm.append([i, sm])
            sm = 0
        return sorted(dir_sm)

    def sized_two(self, dir_sm):
        siz = 0
        temp = []
        t = []
        gdir = []
        for f, s in dir_sm:
            siz += s
            try:
                if siz < self.size:
                    t.insert(100, f)
                    continue
                elif siz > self.size:
                    t.insert(100, f)
                    temp.insert(100, t)
                    gdir.insert(100, t[0] + '-' + t[-1])
                    os.mkdir(os.path.join(self.out_dir, t[0] + '-' + t[-1]))
                    t = []
                    siz = 0
            except IndexError:
                temp.insert(100, t)
                gdir.insert(100, t[0] + '-' + t[-1])
                os.mkdir(os.path.join(self.out_dir, t[0] + '-' + t[-1]))
        return gdir, temp

    def cpy_to(self, gdir, temp):
        dex = 0
        for grp in temp:
            out = gdir[dex]
            for single in grp:
                fi = os.listdir(os.path.join(self.out_dir, single))
                for i in fi:
                    shutil.copy(os.path.join(self.out_dir, single, i),
                                os.path.join(self.out_dir, out))
                shutil.rmtree(os.path.join(self.out_dir, single))
            dex += 1


sea = FileCollector(suffix='.pdf')
loc = sea.locate()
mk_dir = sea.make_dir(loc)
si = sea.sized(loc, mk_dir)
gdir, temp = sea.sized_two(si)
sea.cpy_to(gdir, temp)
