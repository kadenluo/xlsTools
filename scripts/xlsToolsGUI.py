#!/usr/bin/python3
# -*-  coding:utf-8 -*-

import sys
import os
import json
import argparse
from datetime import datetime
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
        self.box.append("{} [INFO] {}".format(datetime.now().strftime('%Y-%m-%d %H:%M:%S'), msg))

    def error(self, msg):
        self.box.append("{} [ERROR] {}".format(datetime.now().strftime('%Y-%m-%d %H:%M:%S'), msg))

class mainWindow():
    def __init__(self, cfgfile):
        self.isForce = False
        self.inputDir = "./xls"
        self.outputDir = "./output"
        self.exportType = "server"
        self.outputType = "lua"
        if os.path.exists(cfgfile):
            with open(cfgfile) as f:
                cfg = json.load(f)
                self.inputDir = cfg["inputDir"]
                self.outputDir = cfg["outputDir"]
                self.exportType = cfg["exportType"]
                self.outputType = cfg["outputType"]
                self.isForce = cfg["isForce"]

    def __str__(self):
        return json.dumps(self.__dict__)

    def onInputDialogClicked(self):
        self.inputDir = QFileDialog.getExistingDirectory(None, "选取文件", "./")
        self.inputDirLine.setText(self.inputDir)

    def onOutputDialogClicked(self):
        self.outputDir = QFileDialog.getExistingDirectory(None, "选取文件", "./")
        self.outputDirLine.setText(self.outputDir)

    def onOutputTypeClicked(self, box):
        self.allTypeBox.setChecked(False)
        self.jsonTypeBox.setChecked(False)
        self.luaTypeBox.setChecked(False)
        box.setChecked(True)
        self.outputType = box.text()

    def onExportTypeClicked(self, box):
        self.clientTypeBox.setChecked(False)
        self.serverTypeBox.setChecked(False)
        box.setChecked(True)
        self.exportType = box.text()

    def onForceClicked(self, box):
        self.forceFalseBox.setChecked(False)
        self.forceTrueBox.setChecked(False)
        box.setChecked(True)
        self.isForce = box.text() == "是"

    def do(self):
        args = argparse.Namespace()
        args.input_dir = self.inputDir
        args.output_dir = self.outputDir
        args.type = self.outputType
        args.force = self.isForce
        args.export = self.exportType
        self.progressText.clear()
        self.converter = Converter(args, self.logger)
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
        self.inputDirLine = QLineEdit(widget)
        self.inputDirLine.setText(self.inputDir)
        #self.inputDirLine.setFocusPolicy(Qt.NoFocus)
        inputDirButton = QPushButton("打开文件夹")
        inputDirButton.clicked.connect(self.onInputDialogClicked)
        inputDirLayout.addWidget(inputDirLabel)
        inputDirLayout.addWidget(self.inputDirLine)
        inputDirLayout.addWidget(inputDirButton)

        outputDirLayout = QHBoxLayout()
        outputDirLabel = QLabel(widget)
        outputDirLabel.setText("输出目录:")
        outputDirLabel.setAlignment(Qt.AlignCenter)
        self.outputDirLine = QLineEdit(widget)
        self.outputDirLine.setText(self.outputDir)
        #self.outputDirLine.setFocusPolicy(Qt.NoFocus)
        outputDirButton = QPushButton("打开文件夹")
        outputDirButton.clicked.connect(self.onOutputDialogClicked)
        outputDirLayout.addWidget(outputDirLabel)
        outputDirLayout.addWidget(self.outputDirLine)
        outputDirLayout.addWidget(outputDirButton)

        exportGroupBox = QGroupBox("表格类型")
        exportGroupBox.setFlat(False)
        self.clientTypeBox = QCheckBox("client")
        self.clientTypeBox.setChecked(self.exportType == "client")
        self.clientTypeBox.clicked.connect(lambda:self.onExportTypeClicked(self.clientTypeBox))
        self.serverTypeBox = QCheckBox("server")
        self.serverTypeBox.setChecked(self.exportType == "server")
        self.serverTypeBox.clicked.connect(lambda:self.onExportTypeClicked(self.serverTypeBox))
        exportTypeLayout = QHBoxLayout()
        exportTypeLayout.addWidget(self.clientTypeBox)
        exportTypeLayout.addWidget(self.serverTypeBox)
        exportGroupBox.setLayout(exportTypeLayout)

        outputGroupBox = QGroupBox("输出类型")
        outputGroupBox.setFlat(False)
        self.allTypeBox = QCheckBox("all")
        self.allTypeBox.setChecked(self.outputType == "all")
        self.allTypeBox.clicked.connect(lambda:self.onOutputTypeClicked(self.allTypeBox))
        self.jsonTypeBox = QCheckBox("json")
        self.jsonTypeBox.setChecked(self.outputType == "json")
        self.jsonTypeBox.clicked.connect(lambda:self.onOutputTypeClicked(self.jsonTypeBox))
        self.luaTypeBox = QCheckBox("lua")
        self.luaTypeBox.setChecked(self.outputType == "lua")
        self.luaTypeBox.clicked.connect(lambda:self.onOutputTypeClicked(self.luaTypeBox))
        outputTypeLayout = QHBoxLayout()
        outputTypeLayout.addWidget(self.allTypeBox)
        outputTypeLayout.addWidget(self.jsonTypeBox)
        outputTypeLayout.addWidget(self.luaTypeBox)
        outputGroupBox.setLayout(outputTypeLayout)

        forceGroupBox = QGroupBox("是否强制导出所有表格")
        forceGroupBox.setFlat(False)
        self.forceTrueBox = QCheckBox("是")
        self.forceTrueBox.setChecked(self.isForce)
        self.forceTrueBox.clicked.connect(lambda:self.onForceClicked(self.forceTrueBox))
        self.forceFalseBox = QCheckBox("否")
        self.forceFalseBox.setChecked(not self.isForce)
        self.forceFalseBox.clicked.connect(lambda:self.onForceClicked(self.forceFalseBox))
        forceLayout = QHBoxLayout()
        forceLayout.addWidget(self.forceTrueBox)
        forceLayout.addWidget(self.forceFalseBox)
        forceGroupBox.setLayout(forceLayout)

        self.progressText = QTextEdit()
        self.progressText.setReadOnly(True)
        self.logger = UILogger(self.progressText)

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
        mainLayout.addWidget(outputGroupBox)
        mainLayout.addWidget(forceGroupBox)
        mainLayout.addWidget(self.progressText)
        mainLayout.addWidget(doButton)
        mainLayout.addStretch()
        widget.setLayout(mainLayout)

        widget.show()
        sys.exit(app.exec_())

if __name__ == "__main__":
    window = mainWindow("./config.json")
    window.MainLoop()
