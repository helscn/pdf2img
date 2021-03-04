#!/usr/bin/python3
# -*- coding: utf-8 -*-


import os
import sys
import fitz
from io import BytesIO
from PIL import Image, ImageQt

from PyQt5.uic import loadUi
from PyQt5.QtCore import Qt, QCoreApplication, pyqtSignal, QThread
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox


class ApplicationWindow(QMainWindow):
    def __init__(self):
        super(ApplicationWindow, self).__init__()
        loadUi('mainWindow.ui', self)
        self.__pdfPath = None             # 打开的PDF文件路径
        self.__pdfDoc = None              # 打开的PDF文档对象
        self.__minHeight = 0              # PDF页面最大高度
        self.__minWidth = 0               # PDF页面最小高度
        self.__totalPage = 0              # PDF总页数
        self.__currentPage = 0            # 当前打开的PDF页面
        self.initUi()
        self.initSignal()

    def initUi(self):
        self.statusBar.show()
        self.progressBar.hide()
        self.refreshOption()

    def initSignal(self):
        self.btnLoadPDF.clicked.connect(self.loadPDF)
        self.btnSaveImage.clicked.connect(self.savePDF)
        self.btnLastPage.clicked.connect(self.showLastPage)
        self.btnNextPage.clicked.connect(self.showNextPage)
        self.spinClipTop.valueChanged.connect(self.showPreview)
        self.spinClipBottom.valueChanged.connect(self.showPreview)
        self.spinClipLeft.valueChanged.connect(self.showPreview)
        self.spinClipRight.valueChanged.connect(self.showPreview)
        self.spinScaleFactor.valueChanged.connect(self.showPreview)

    def disableSignal(self):
        self.btnLoadPDF.clicked.disconnect()
        self.btnSaveImage.clicked.disconnect()
        self.btnLastPage.clicked.disconnect()
        self.btnNextPage.clicked.disconnect()
        self.spinClipTop.valueChanged.disconnect()
        self.spinClipBottom.valueChanged.disconnect()
        self.spinClipLeft.valueChanged.disconnect()
        self.spinClipRight.valueChanged.disconnect()
        self.spinScaleFactor.valueChanged.disconnect()

    def setEnable(self):
        self.btnLoadPDF.setDisabled(False)
        self.btnSaveImage.setDisabled(False)
        self.btnLastPage.setDisabled(False)
        self.btnNextPage.setDisabled(False)
        self.spinClipTop.setDisabled(False)
        self.spinClipBottom.setDisabled(False)
        self.spinClipLeft.setDisabled(False)
        self.spinClipRight.setDisabled(False)
        self.spinPageStart.setDisabled(False)
        self.spinPageEnd.setDisabled(False)
        self.spinScaleFactor.setDisabled(False)

    def setDisabled(self):
        self.btnLoadPDF.setDisabled(True)
        self.btnSaveImage.setDisabled(True)
        self.btnLastPage.setDisabled(True)
        self.btnNextPage.setDisabled(True)
        self.spinClipTop.setDisabled(True)
        self.spinClipBottom.setDisabled(True)
        self.spinClipLeft.setDisabled(True)
        self.spinClipRight.setDisabled(True)
        self.spinPageStart.setDisabled(True)
        self.spinPageEnd.setDisabled(True)
        self.spinScaleFactor.setDisabled(True)

    def refreshOption(self):
        """根据PDF文件中的页面信息刷新各组件的设定值"""
        self.spinClipTop.setMaximum(int(self.__minHeight/2))
        self.spinClipBottom.setMaximum(int(self.__minHeight/2))
        self.spinClipLeft.setMaximum(int(self.__minWidth/2))
        self.spinClipRight.setMaximum(int(self.__minWidth/2))
        if self.__totalPage == 0:
            minimum = 0
        else:
            minimum = 1

        self.spinPageEnd.setMinimum(minimum)
        self.spinPageEnd.setMaximum(self.__totalPage)
        self.spinPageEnd.setValue(self.__totalPage)

        self.spinPageStart.setMinimum(minimum)
        self.spinPageStart.setMaximum(self.spinPageEnd.value())

    def getPageImage(self, page, scale=1.0, clipTop=0.0, clipBottom=0.0, clipLeft=0.0, clipRight=0.0):
        """将PDF的Page对象转换为对应的PIL图片对象，可以设置缩放比例及裁剪大小"""
        # 设置PDF缩放因子
        matrix = fitz.Matrix(scale, scale)

        # 裁剪PDF页面
        rect = page.rect
        tl = rect.tl
        br = rect.br
        tl.x += clipLeft
        tl.y += clipTop
        br.x -= clipRight
        br.y -= clipBottom
        clip = fitz.Rect(tl, br)
        png = page.getPixmap(matrix=matrix, clip=clip).getPNGData()
        pixmap = Image.open(BytesIO(png))
        return pixmap

    def getClipOptions(self):
        """获取UI界面中的图像输出参数"""
        options = {
            'scale': self.spinScaleFactor.value(),
            'clipTop': self.spinClipTop.value(),
            'clipBottom': self.spinClipBottom.value(),
            'clipLeft': self.spinClipLeft.value(),
            'clipRight': self.spinClipRight.value(),
        }
        return options

    def showPreview(self):
        """显示当前页面的缩略图"""
        self.lblPageCount.setText('第 {} 页，共 {} 页'.format(
            self.__currentPage, self.__totalPage))
        if self.__totalPage > 0 and self.__currentPage > 0:
            self.btnLoadPDF.setText('')
            options = self.getClipOptions()
            pixmap = self.getPageImage(
                self.__pdfDoc[self.__currentPage-1], **options)
            icon = QIcon(ImageQt.toqpixmap(pixmap))
            self.btnLoadPDF.setIcon(icon)
            self.lblPageSize.setText(
                '宽度：{} / 高度：{}'.format(pixmap.width, pixmap.height))
        else:
            self.btnLoadPDF.setIcon(QIcon())
            self.btnLoadPDF.setText('点击载入PDF文件...')
            self.lblPageSize.setText('宽度：0 / 高度：0')

    def showLastPage(self):
        """显示上一页PDF页面预览"""
        if self.__currentPage > 1:
            self.__currentPage -= 1
            self.showPreview()

    def showNextPage(self):
        """显示下一页PDF页面预览"""
        if self.__currentPage < self.__totalPage:
            self.__currentPage += 1
            self.showPreview()

    def loadPDF(self):
        """打开需要转换的PDF文件"""
        filePath, fileType = QFileDialog.getOpenFileName(
            parent=self, caption="选择PDF文件", filter="PDF Files (*.pdf)")  # 设置文件扩展名过滤注意用双分号间隔
        if filePath and os.path.isfile(filePath):
            try:
                self.setDisabled()
                self.__pdfPath = filePath
                self.__pdfDoc = fitz.open(filePath)
                self.__totalPage = self.__pdfDoc.pageCount
                minWidth = 0
                minHeight = 0
                for i in range(self.__totalPage):
                    br = self.__pdfDoc[i].rect.br
                    if minWidth == 0 or br.x < minWidth:
                        minWidth = br.x
                    if minHeight == 0 or br.y > minHeight:
                        minHeight = br.y
                self.__minWidth = int(minWidth)
                self.__minHeight = int(minHeight)
                self.__currentPage = 1
            except Exception as e:
                QMessageBox.warning(
                    self, "错误", "无法读取指定的PDF文件！\n\n".format(e.args[0]), QMessageBox.Yes)
                self.__pdfDoc = None
                self.__totalPage = 0
                self.__minWidth = 0
                self.__minHeight = 0
                self.__currentPage = 0
            self.refreshOption()
            self.showPreview()
            self.setEnable()

    def savePDF(self):
        """将PDF文件输出为图片"""
        if self.__pdfDoc is None:
            QMessageBox.warning(self, "错误", "请先打开需要转换的PDF文件！", QMessageBox.Yes)
            return None
        filePath, fileType = QFileDialog.getSaveFileName(
            parent=self, caption="图片另存为", filter="PNG Files (*.png);;JPG Files (*.jpg);;All Files (*.*)")  # 设置文件扩展名过滤注意用双分号间隔

        if filePath:
            self.setDisabled()
            self.showProgress(0)
            self.progressBar.show()
            start = self.spinPageStart.value()
            end = self.spinPageEnd.value()
            step = 0
            count = (end-start+1)*2+2
            posX = posY = maxWidth = maxHeight = 0

            options = self.getClipOptions()

            for i in range(start-1, end):
                # 分析合并后PDF页面的总大小
                img = self.getPageImage(
                    self.__pdfDoc[i], **options)
                if img.width > maxWidth:
                    maxWidth = img.width
                maxHeight += img.height
                step += 1
                self.showProgress(step, count)

            target = Image.new('RGB', (maxWidth, maxHeight))
            step += 1
            self.showProgress(step, count)

            for i in range(start-1, end):
                # 拼接PDF图片
                img = self.getPageImage(
                    self.__pdfDoc[i], **options)
                posX = int((maxWidth-img.width)/2)
                target.paste(img, (posX, posY))
                posY += img.height
                step += 1
                self.showProgress(step, count)
            try:
                target.save(filePath)
            except Exception as e:
                QMessageBox.warning(
                    self, "错误", "输出图片拼接文件出错，无法保存图片！\n\n{}".format(e.args[0]), QMessageBox.Yes)
                return None
            step += 1
            self.showProgress(step, count)
            self.progressBar.hide()
            QMessageBox.information(self, "完成", "拼接图片已保存完毕。", QMessageBox.Ok)
            self.setEnable()

    def showProgress(self, current, total=100):
        """更新进度条的当前完成进度"""
        self.progressBar.setValue(current/total*100)
        QApplication.processEvents()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    appWindow = ApplicationWindow()
    appWindow.show()
    sys.exit(app.exec_())
