import sys, random

import pymongo
import dns
from PyQt5.QtWidgets import (QApplication, QMainWindow)
from PyQt5.QtChart import QChart, QChartView, QValueAxis, QBarCategoryAxis, QBarSet, QBarSeries
from PyQt5.Qt import Qt
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QPainter


class Window(QMainWindow):
    def __init__(self):

        super().__init__()
        self.tern = ''
        self.typed = ''
        self.province = ''
        self.setWindowTitle("recommended")
        self.resize(799, 476)

    def setwin(self):
        Window.__init__(self)

    def setaug(self, tern, typed, provice):
        self.tern = tern
        self.typed = typed
        self.province = provice
        self.Ui_grafh()

    def Ui_grafh(self):
        connectMongodb = 'mongodb+srv://Kornwara:2tIbDeNWUF7Yjswi@cluster0.ezb5a.mongodb.net/<dbname>?retryWrites=true&w=majority'
        with pymongo.MongoClient(connectMongodb)as conn:  # add comment
            db = conn.get_database('VacationPlace')
            have = True
            where = {'$and': [{'ภาค': self.tern}, {self.typed: {'$exists': have}}]}
            cursor = db.Province.find(where)
            data=[]
            score=[]
            avg=0
            count=[]
            place=[]
            for i in cursor:
                avg = 0

                for j in i[self.typed]:
                    if "Comment" in j.keys():
                        ecount = 0
                        edata = []
                        place.append(j['name'])
                        for k in j['Comment']:
                            edata.append(k["Score"])
                            ecount += 1
                        data.append(edata)
                        count.append(ecount)
        avg =[]
        avgs=[]
        for i in range(data.__len__()):
            score1=0
            for j in data[i]:
                score1+=j
            score1/=float(count[i])
            avg.append(score1.__round__(2))
        top5avgs=[]
        avgs = sorted(avg, reverse=True)
        for i in range(avg.__len__()):# หาค่าอินเด็ก
            for j in range(avg.__len__()):
                if avgs[i]==avg[j]:
                    top5avgs.append(j)
                    avg[j]=0
                    break

        set0 = QBarSet(place[top5avgs[0]])
        set1 = QBarSet(place[top5avgs[1]])
        set2 = QBarSet(place[top5avgs[2]])
        set3 = QBarSet(place[top5avgs[3]])
        set4 = QBarSet(place[top5avgs[4]])

        set0.append(avgs[0])
        set1.append(avgs[1])
        set2.append(avgs[2])
        set3.append(avgs[3])
        set4.append(avgs[4])

        series = QBarSeries()
        series.append(set0)
        series.append(set1)
        series.append(set2)
        series.append(set3)
        series.append(set4)

        chart = QChart()
        chart.addSeries(series)
        chart.setTitle('Recommened Place')
        chart.setAnimationOptions(QChart.SeriesAnimations)


        axisY = QValueAxis()
        axisY.setRange(0, 5)

        chart.addAxis(axisY, Qt.AlignLeft)

        chart.legend().setVisible(True)
        chart.legend().setAlignment(Qt.AlignBottom)
        chartView = QChartView(chart)
        self.setCentralWidget(chartView)


if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = Window()
    window.show()

    sys.exit(app.exec_())
