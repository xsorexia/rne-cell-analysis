import zipfile
import roifile
import numpy as np
import csv
import openpyxl

class Cell:
    def __init__(self):
        self.x = 0
        self.y = 0
        self.area = 0
        self.intensity = 0
        self.type = "none"
    
    def init(self, type, coords):
        centroid = cellCentroid(coords)
        self.x, self.y = centroid[0], centroid[1]
        self.area = cellArea(coords)
        self.type = type


def cellCentroid(coords):
    x = coords[:, 0]
    y = coords[:, 1]
    
    # 다각형의 centroid 공식 :
    cx = abs(np.sum((x + np.roll(x, 1)) * (x * np.roll(y, 1) - y * np.roll(x, 1)))) / (6 * area)
    cy = abs(np.sum((y + np.roll(y, 1)) * (x * np.roll(y, 1) - y * np.roll(x, 1)))) / (6 * area)
    geometric_centroid = np.array([cx, cy])
    
    return geometric_centroid


def cellArea(coords):
    x = coords[:, 0]
    y = coords[:, 1]
    
    # 신발끈 공식을 이용한 다각형의 넓이 계산 :
    area = 0.5 * np.abs(np.dot(x, np.roll(y, 1)) - np.dot(y, np.roll(x, 1)))

    return area

def extractCells(type, zipPath):
    cellList = []

    with zipfile.ZipFile(zipPath, 'r') as zipRef:
        for fileName in zipRef.namelist():
            if fileName.endswith('.roi'):
                with zipRef.open(fileName) as roiFile:
                    roiData = roiFile.read()
                    roi = roifile.ImagejRoi.frombytes(roiData)

                    if roi.roitype == roifile.ROI_TYPE.FREEHAND:
                        coords = np.array(roi.coordinates())
                        cell = Cell()
                        cell.init(type, coords)

                        cellList.append(cell)
    return cellList


