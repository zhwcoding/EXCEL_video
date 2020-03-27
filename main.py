# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import QWidget, QApplication, QMessageBox, QFileDialog, QProgressBar, QDialog, QSizePolicy, QVBoxLayout, QLabel, QHBoxLayout, QPushButton
from PyQt5.QtCore import Qt
from threading import Thread, Semaphore
import cv2
import pandas
import sys
import os
import time

from Ui_main import Ui_Form


class MainWindow(QWidget, Ui_Form):
    '''create mainwindow'''
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)

        self.videoPath = ''
        self.outPath = ''
        self.videoLenth = 0
        self.framesEverySeconds = 0
        self.frames_num = 0
        self.h_pixel = 0
        self.v_pixel = 0
        self.step = 0
        self.shrink = 0
        self.begin = 0
        self.end = 0
        
        self.img_num = 0

        self.thread1 = Thread(target=self.extract_write)

        self.pushButton_selectVideoPath.clicked.connect(self.select_videoPath)
        self.pushButton_selectOutPath.clicked.connect(self.select_outPath)
        self.pushButton_STARTWORK.clicked.connect(self.start_work)
        self.pushButton_cancel.clicked.connect(self.cancel)
    

    # “开始转换”按钮触发
    def start_work(self):
        self.step = self.spinBox_step.value()
        self.shrink = self.spinBox_shrink.value()
        self.begin = self.spinBox_start.value()
        self.end = self.spinBox_end.value()
        self.img_num = (self.end - self.begin) // (self.step + 1) + 1
        # 只有指定了路径才执行操作
        if self.videoPath and self.outPath:
            self.pushButton_STARTWORK.setEnabled(False)
            self.pushButton_selectOutPath.setEnabled(False)
            self.pushButton_selectVideoPath.setEnabled(False)
            self.spinBox_end.setEnabled(False)
            self.spinBox_start.setEnabled(False)
            self.spinBox_shrink.setEnabled(False)
            self.spinBox_step.setEnabled(False)
            self.frame.setVisible(True)
            time.sleep(1)
            self.thread1.start()
        else:
            QMessageBox.information(self, '！！！', '请先指定两个路径')


    def cancel(self):
        self.thread1._delete()
        self.thread1 = Thread(target=self.extract_write)
        self.progressBar.setValue(0)
        self.pushButton_STARTWORK.setEnabled(True)
        self.pushButton_selectOutPath.setEnabled(True)
        self.pushButton_selectVideoPath.setEnabled(True)
        self.spinBox_end.setEnabled(True)
        self.spinBox_start.setEnabled(True)
        self.spinBox_shrink.setEnabled(True)
        self.spinBox_step.setEnabled(True)
        self.frame.setVisible(False)


    # 选取视频路径, 并显示视频的基本信息
    def select_videoPath(self):
        fileDialog = QFileDialog(self, filter="mp4文件 (*.mp4)")
        fileDialog.setWindowModality(Qt.WindowModal)
        fileDialog.exec_()
        try:
            self.videoPath = fileDialog.selectedFiles()[0]
        except:
            pass
        video = cv2.VideoCapture()
        if not video.open(self.videoPath):
             QMessageBox.information(self, '！！！', '无法读取视频')
             return
        self.frames_num = video.get(7)
        self.h_pixel = video.get(3)
        self.v_pixel = video.get(4)
        self.framesEverySeconds = video.get(5)
        self.videoLenth = self.frames_num/self.framesEverySeconds

        self.spinBox_end.setMaximum(self.frames_num)
        self.spinBox_start.setMaximum(self.frames_num)
        self.spinBox_step.setMaximum(self.frames_num-1)
        self.label_videoPath.setText(self.videoPath)
        self.label_framesEverySeconds.setText(str(int(self.framesEverySeconds)))
        self.label_h_pixel.setText(str(int(self.h_pixel)))
        self.label_v_pixel.setText(str(int(self.v_pixel)))
        self.label_videoLenth.setText(str(int(self.videoLenth)))
        

    def select_outPath(self):
        self.outPath = QFileDialog().getSaveFileName(self, 'save file', filter='(*.xlsx)')[0]
        self.label_excelPath.setText(self.outPath)


    # 抽取视频中的帧，并存入EXCEL中
    def extract_write(self):
        video = cv2.VideoCapture()
        video.open(self.videoPath)
        step = self.step + 1
        index = 0
        count = self.begin
        writer = pandas.ExcelWriter(self.outPath)
        b_data = pandas.DataFrame()
        g_data = pandas.DataFrame()
        r_data = pandas.DataFrame()
        video_property = pandas.DataFrame(['水平像素值：', int(self.h_pixel*self.shrink),
        '竖直像素值：', int(self.v_pixel*self.shrink),
        '向右平移列数：', 5,
        '向下平移行数：', 0,
        '总帧数：', self.img_num,
        '每两帧间隔(秒)：', 0.1,
        '进度(%)：',0])
        while True:
            time.sleep(0.005)
            _, frame = video.read()
            if frame is None:
                break
            if (count-self.begin) % step == 0:
                index += 1
                res_min = cv2.resize(frame,None,fx=self.shrink,fy=self.shrink,interpolation=cv2.INTER_AREA)
                b1, g1, r1 = cv2.split(res_min)
                b_data1 = pandas.DataFrame(b1)
                g_data1 = pandas.DataFrame(g1)
                r_data1 = pandas.DataFrame(r1)
                b_data = pandas.concat([b_data, b_data1], axis=0)
                g_data = pandas.concat([g_data, g_data1], axis=0)
                r_data = pandas.concat([r_data, r_data1], axis=0)
                self.progressBar.setValue((index/self.img_num)*100-1)
            if (count == self.end):
                break
            count += 1
        video.release()
        video_property.to_excel(writer, 'video')
        b_data = b_data//9*9
        g_data = g_data//9*9
        r_data = r_data//9*9        # 此操作是因为EXCEL中总单元格格式数有限制
        b_data.to_excel(writer, 'b')
        g_data.to_excel(writer, 'g')
        r_data.to_excel(writer, 'r')
        writer.save()
        writer.close()
        self.progressBar.setValue(100)
        self.pushButton_cancel.setEnabled(False)
        time.sleep(2)
        self.progressBar.setValue(0)
        self.thread1 = Thread(target=self.extract_write)
        self.pushButton_STARTWORK.setEnabled(True)
        self.pushButton_selectOutPath.setEnabled(True)
        self.pushButton_selectVideoPath.setEnabled(True)
        self.spinBox_end.setEnabled(True)
        self.spinBox_start.setEnabled(True)
        self.spinBox_shrink.setEnabled(True)
        self.spinBox_step.setEnabled(True)
        self.frame.setVisible(False)




if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = MainWindow()
    main.setWindowTitle('视频转EXCEL数据')
    main.show()
    sys.exit(app.exec_())    
    