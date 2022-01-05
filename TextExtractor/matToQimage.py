# -*- coding: utf-8 -*-
"""
Created on Thu Jul 19 12:14:32 2018

@author: Sunil
"""

import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QImage, qRgb
import numpy as np

COLORTABLE=[]
for i in range(256): COLORTABLE.append(QtGui.qRgb(i,i,i))

def toQImage(im, copy=False):
    if im is None:
        return QImage()

    if im.dtype == np.uint8:
        if len(im.shape) == 2:
            #print('imside image')
            #print(im)
            qim = QImage(im.data, im.shape[1], im.shape[0], im.strides[0], QImage.Format_Indexed8)
            qim.setColorTable(COLORTABLE)
            return qim.copy() if copy else qim

        elif len(im.shape) == 3:
            if im.shape[2] == 3:
                qim = QImage(im.data, im.shape[1], im.shape[0], im.strides[0], QImage.Format_RGB888);
                return qim.copy() if copy else qim
            elif im.shape[2] == 4:
                qim = QImage(im.data, im.shape[1], im.shape[0], im.strides[0], QImage.Format_ARGB32);
                return qim.copy() if copy else qim