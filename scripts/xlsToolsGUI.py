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

    def info(self, pattern, *args):
        self.box.append("{} [INFO] {}".format(datetime.now().strftime('%Y-%m-%d %H:%M:%S'), pattern.format(*args)))

    def error(self, pattern, *args):
        self.box.append("{} [ERROR] {}".format(datetime.now().strftime('%Y-%m-%d %H:%M:%S'), pattern.format(*args)))

class mainWindow():
    def __init__(self, cfgfile):
        self.isForce = False
        self.inputDir = "./xls"
        self.clientOutputDir = "./output/client"
        self.serverOutputDir = "./output/server"
        self.outputType = "lua"
        self.excludeFiles = []
        if os.path.exists(cfgfile):
            with open(cfgfile, 'r') as f:
                cfg = json.load(f)
                self.inputDir = "inputDir" in cfg and cfg["inputDir"] or self.inputDir
                self.clientOutputDir = "clientOutputDir" in cfg and cfg["clientOutputDir"] or self.clientOutputDir
                self.serverOutputDir = "serverOutputDir" in cfg and cfg["serverOutputDir"] or self.serverOutputDir
                self.outputType = "outputType" in cfg and cfg["outputType"] or self.outputType
                self.isForce = "isForce" in cfg and cfg["isForce"] or self.isForce
                self.excludeFiles = "excludeFiles" in cfg and cfg["excludeFiles"] or self.excludeFiles

    def __str__(self):
        return json.dumps(self.__dict__)

    def onInputDialogClicked(self):
        self.inputDir = QFileDialog.getExistingDirectory(None, "选取文件", "./")
        self.inputDirLine.setText(self.inputDir)

    def onClientOutputDialogClicked(self):
        self.clientOutputDir = QFileDialog.getExistingDirectory(None, "选取文件", "./")
        self.clientOutputDirLine.setText(self.clientOutputDir)

    def onServerOutputDialogClicked(self):
        self.serverOutputDir = QFileDialog.getExistingDirectory(None, "选取文件", "./")
        self.serverOutputDirLine.setText(self.serverOutputDir)

    def onOutputTypeClicked(self, box):
        self.allTypeBox.setChecked(False)
        self.jsonTypeBox.setChecked(False)
        self.luaTypeBox.setChecked(False)
        box.setChecked(True)
        self.outputType = box.text()

    def onForceClicked(self, box):
        self.forceFalseBox.setChecked(False)
        self.forceTrueBox.setChecked(False)
        box.setChecked(True)
        self.isForce = box.text() == "是"

    def do(self):
        files = self.excludeFilesLine.text()
        self.excludeFiles = files.split(",")

        args = argparse.Namespace()
        args.input_dir = self.inputDir
        args.client_output_dir = self.clientOutputDir
        args.server_output_dir = self.serverOutputDir
        args.type = self.outputType
        args.force = self.isForce
        args.exclude_files = self.excludeFiles
        self.progressText.clear()
        self.converter = Converter(args, self.logger)
        self.converter.convertAll()

    def MainLoop(self):
        app = QApplication(sys.argv)

        widget = QWidget()
        widget.resize(960, 720)
        widget.setWindowTitle("转表工具")

        inputDirLayout = QHBoxLayout()
        inputDirLabel = QLabel(widget)
        inputDirLabel.setText("excel目录：")
        inputDirLabel.setAlignment(Qt.AlignCenter)
        self.inputDirLine = QLineEdit(widget)
        self.inputDirLine.setText(self.inputDir)
        inputDirButton = QPushButton("打开文件夹")
        inputDirButton.clicked.connect(self.onInputDialogClicked)
        inputDirLayout.addWidget(inputDirLabel, stretch=1)
        inputDirLayout.addWidget(self.inputDirLine, stretch=8)
        inputDirLayout.addWidget(inputDirButton, stretch=1)

        excludeFilesLayout = QHBoxLayout()
        excludeFilesLabel = QLabel(widget)
        excludeFilesLabel.setText("排除文件：")
        excludeFilesLabel.setAlignment(Qt.AlignCenter)
        self.excludeFilesLine = QLineEdit(widget)
        self.excludeFilesLine.setText(",".join(self.excludeFiles))
        excludeFilesLayout.addWidget(excludeFilesLabel, stretch=1)
        excludeFilesLayout.addWidget(self.excludeFilesLine, stretch=9)
        #excludeFilesLayout.addStretch(1)

        clientOutputDirLayout = QHBoxLayout()
        clientOutputDirLabel = QLabel(widget)
        clientOutputDirLabel.setText("client输出目录：")
        clientOutputDirLabel.setAlignment(Qt.AlignCenter)
        self.clientOutputDirLine = QLineEdit(widget)
        self.clientOutputDirLine.setText(self.clientOutputDir)
        clientOutputDirButton = QPushButton("打开文件夹")
        clientOutputDirButton.clicked.connect(self.onClientOutputDialogClicked)
        clientOutputDirLayout.addWidget(clientOutputDirLabel, stretch=1)
        clientOutputDirLayout.addWidget(self.clientOutputDirLine, stretch=8)
        clientOutputDirLayout.addWidget(clientOutputDirButton, stretch=1)

        serverOutputDirLayout = QHBoxLayout()
        serverOutputDirLabel = QLabel(widget)
        serverOutputDirLabel.setText("server输出目录：")
        serverOutputDirLabel.setAlignment(Qt.AlignCenter)
        self.serverOutputDirLine = QLineEdit(widget)
        self.serverOutputDirLine.setText(self.serverOutputDir)
        serverOutputDirButton = QPushButton("打开文件夹")
        serverOutputDirButton.clicked.connect(self.onServerOutputDialogClicked)
        serverOutputDirLayout.addWidget(serverOutputDirLabel, stretch=1)
        serverOutputDirLayout.addWidget(self.serverOutputDirLine, stretch=8)
        serverOutputDirLayout.addWidget(serverOutputDirButton, stretch=1)

        outputGroupBox = QGroupBox("导出类型")
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
        doButton = QPushButton("开始")
        doButton.clicked.connect(self.do)
        doButton.setStyleSheet("background-color:rgb(14, 137, 205);color:white;border-radius:8px;font-family:Microsoft Yahei;font-size:20pt");
        doLayout.addStretch(4)
        doLayout.addWidget(doButton, stretch=2)
        doLayout.addStretch(4)

        mainLayout = QVBoxLayout()
        mainLayout.addLayout(inputDirLayout)
        mainLayout.addLayout(excludeFilesLayout)
        mainLayout.addLayout(clientOutputDirLayout)
        mainLayout.addLayout(serverOutputDirLayout)
        mainLayout.addWidget(outputGroupBox)
        mainLayout.addWidget(forceGroupBox)
        mainLayout.addWidget(self.progressText, stretch=50)
        mainLayout.addLayout(doLayout, stretch=1)
        mainLayout.addStretch(1)
        widget.setLayout(mainLayout)

        widget.show()
        sys.exit(app.exec_())

if __name__ == "__main__":
    window = mainWindow("./config.json")
    window.MainLoop()
