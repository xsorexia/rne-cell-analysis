import zipfile
import roifile
import numpy as np
import csv

headerList = ["4T1_Raw_1_1", "4T1_Raw_1_2", "4T1_Raw_1_4", "Raw_4T1_1_2", "Raw_4T1_1_4"]
hourList = ["0h", "24h", "72h"]
folderCount = [[3, 3, 2], [3, 3, 3], [1, 3, 3], [3, 2, 2], [3, 3, 2]]

class Cell:
    def __init__(self):
        self.x = 0
        self.y = 0
        self.area = 0
        self.intensity = 0
        self.type = "none"
        self.label = "default"
    
    def __str__(self):
        return f"[Cell {self.type}] {self.label} \nIntensity | {self.intensity}\nCentroid | ({self.x}, {self.y})\n"
    
    def init(self, type, coords):
        self.area = cellArea(coords)
        centroid = cellCentroid(coords, self.area)
        self.x, self.y = centroid[0], centroid[1]
        self.type = type
    
    def setIntensity(self, intensity):
        self.intensity = intensity
    
    def setLabel(self, label):
        self.label = label

    def intensityPerArea(self):
        if (float(self.area) != 0):
            return float(self.intensity) / float(self.area)
        return 0


def cellCentroid(coords, area):
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

def cellExtractRoi(type, zipPath, roiName):
    roiFileName = roiName + ".roi"

    with zipfile.ZipFile(zipPath, 'r') as zipRef:
        for fileName in zipRef.namelist():
            if fileName.endswith('.roi') and fileName == roiFileName:
                with zipRef.open(fileName) as roiFile:
                    roiData = roiFile.read()
                    roi = roifile.ImagejRoi.frombytes(roiData)

                    if roi.roitype == roifile.ROI_TYPE.FREEHAND:
                        coords = np.array(roi.coordinates())
                        cell = Cell()
                        cell.init(type, coords)

                        return cell
    return False


def cellDist(cellA, cellB):
    scaleBar = 1/2.08 # 광학망원경 20배율, 2.08 pixel / 1 micrometer --> 1/2.08 micrometer / 1 pixel
    return (((cellA.x - cellB.x) ** 2 + (cellA.y - cellB.y) ** 2) ** 0.5) * scaleBar





def extractCells(header, hours, folder):
    zip_path_4T1 = "data/" + header + "/" + hours + "/" + folder + "/" + "4T1.zip"
    zip_path_Raw = "data/" + header + "/" + hours + "/" + folder + "/" + "Raw.zip"
    xl_path_4T1 = "data/" + header + "/" + hours + "/" + folder + "/" + "4T1.csv"
    xl_path_Raw = "data/" + header + "/" + hours + "/" + folder + "/" + "Raw.csv"

    cellListRaw = []
    cellList4T1 = []
    
    try:
        csvRaw = open(xl_path_Raw, 'r')
        dataRaw = csv.DictReader(csvRaw)

        csv4T1 = open(xl_path_4T1, 'r')
        data4T1 = csv.DictReader(csv4T1)
    except:
        print("파일 위치 지정 오류")
    
    for row in dataRaw:
        roiName = row['Label'].split(":")[1]
        cellRaw = cellExtractRoi('raw', zip_path_Raw, roiName)
        cellRaw.setIntensity(row['Mean'])
        cellRaw.setLabel(roiName)
        cellListRaw.append(cellRaw)
    
    for row in data4T1:
        roiName = row['Label'].split(":")[1]
        cell4T1 = cellExtractRoi('raw', zip_path_4T1, roiName)
        cell4T1.setIntensity(row['Mean'])
        cell4T1.setLabel(roiName)
        cellList4T1.append(cell4T1)

    return [cellListRaw, cellList4T1]

