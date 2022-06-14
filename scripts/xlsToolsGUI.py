#!/usr/bin/python3
# -*-  coding:utf-8 -*-

import sys
import argparse
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from xlsTools import Converter, Logger


class UILogger(Logger):
    def __init__(self, box):
        super().__init__()
        self.box = box

    def info(self, msg):
        self.box.append("[INFO] {}".format(msg))

    def error(self, msg):
        self.box.append("[ERROR] {}".format(msg))

class mainWindow():
    def __init__(self):
        self.inputDir = "./xls"
        self.outputDir = "./output"
        self.exportType = "all"
        self.isforce = False

    def onInputDialogClicked(self):
        self.inputDir = QFileDialog.getExistingDirectory(None, "选取文件", "./")

    def onOutputDialogClicked(self):
        self.outputDir = QFileDialog.getExistingDirectory(None, "选取文件", "./")

    def onExportTypeClicked(self, box):
        self.allTypeBox.setChecked(False)
        self.jsonTypeBox.setChecked(False)
        self.luaTypeBox.setChecked(False)
        box.setChecked(True)
        self.exportType = box.text()

    def onForceClicked(self, box):
        self.forceFalseBox.setChecked(False)
        self.forceTrueBox.setChecked(False)
        box.setChecked(True)
        self.isforce = box.text() == "是"

    def do(self):
        args = argparse.Namespace()
        args.input_dir = self.inputDir
        args.output_dir = self.outputDir
        args.type = self.exportType
        args.force = self.isforce
        self.progressText.clear()
        logger = UILogger(self.progressText)
        self.converter = Converter(args, UILogger(self.progressText))
        self.converter.convertAll()

    def MainLoop(self):
        app = QApplication(sys.argv)

        widget = QWidget()
        widget.resize(680, 430)
        widget.setWindowTitle("转表工具")

        inputDirLayout = QHBoxLayout()
        inputDirLabel = QLabel(widget)
        inputDirLabel.setText("excel目录:")
        inputDirLabel.setAlignment(Qt.AlignCenter)
        inputDirLine = QLineEdit(widget)
        inputDirLine.setText(self.inputDir)
        #inputDirLine.setFocusPolicy(Qt.NoFocus)
        inputDirButton = QPushButton("打开文件夹")
        inputDirButton.clicked.connect(self.onInputDialogClicked)
        inputDirLayout.addWidget(inputDirLabel)
        inputDirLayout.addWidget(inputDirLine)
        inputDirLayout.addWidget(inputDirButton)

        outputDirLayout = QHBoxLayout()
        outputDirLabel = QLabel(widget)
        outputDirLabel.setText("输出目录:")
        outputDirLabel.setAlignment(Qt.AlignCenter)
        outputDirLine = QLineEdit(widget)
        outputDirLine.setText(self.outputDir)
        #outputDirLine.setFocusPolicy(Qt.NoFocus)
        outputDirButton = QPushButton("打开文件夹")
        outputDirButton.clicked.connect(self.onOutputDialogClicked)
        outputDirLayout.addWidget(outputDirLabel)
        outputDirLayout.addWidget(outputDirLine)
        outputDirLayout.addWidget(outputDirButton)

        exportGroupBox = QGroupBox("导出类型")
        exportGroupBox.setFlat(False)
        self.allTypeBox = QCheckBox("all")
        self.allTypeBox.setChecked(self.exportType == "all")
        self.allTypeBox.clicked.connect(lambda:self.onExportTypeClicked(self.allTypeBox))
        self.jsonTypeBox = QCheckBox("json")
        self.jsonTypeBox.setChecked(self.exportType == "json")
        self.jsonTypeBox.clicked.connect(lambda:self.onExportTypeClicked(self.jsonTypeBox))
        self.luaTypeBox = QCheckBox("lua")
        self.luaTypeBox.setChecked(self.exportType == "lua")
        self.luaTypeBox.clicked.connect(lambda:self.onExportTypeClicked(self.luaTypeBox))
        exportTypeLayout = QHBoxLayout()
        exportTypeLayout.addWidget(self.allTypeBox)
        exportTypeLayout.addWidget(self.jsonTypeBox)
        exportTypeLayout.addWidget(self.luaTypeBox)
        exportGroupBox.setLayout(exportTypeLayout)

        forceGroupBox = QGroupBox("是否强制导出所有表格")
        forceGroupBox.setFlat(False)
        self.forceTrueBox = QCheckBox("是")
        self.forceTrueBox.setChecked(self.isforce)
        self.forceTrueBox.clicked.connect(lambda:self.onForceClicked(self.forceTrueBox))
        self.forceFalseBox = QCheckBox("否")
        self.forceFalseBox.setChecked(not self.isforce)
        self.forceFalseBox.clicked.connect(lambda:self.onForceClicked(self.forceFalseBox))
        forceLayout = QHBoxLayout()
        forceLayout.addWidget(self.forceTrueBox)
        forceLayout.addWidget(self.forceFalseBox)
        forceGroupBox.setLayout(forceLayout)

        self.progressText = QTextEdit()
        self.progressText.setReadOnly(True)

        doLayout = QHBoxLayout()
        doButton = QPushButton("执行")
        doButton.clicked.connect(self.do)
        doButton.setStyleSheet("background-color:rgb(0, 105, 205)");
        doLayout.addStretch()
        doLayout.addWidget(doButton)
        doLayout.addStretch()

        mainLayout = QVBoxLayout()
        mainLayout.addLayout(inputDirLayout)
        mainLayout.addLayout(outputDirLayout)
        mainLayout.addWidget(exportGroupBox)
        mainLayout.addWidget(forceGroupBox)
        mainLayout.addWidget(self.progressText)
        mainLayout.addWidget(doButton)
        mainLayout.addStretch()
        widget.setLayout(mainLayout)

        widget.show()
        sys.exit(app.exec_())

if __name__ == "__main__":
    window = mainWindow()
    window.MainLoop()
