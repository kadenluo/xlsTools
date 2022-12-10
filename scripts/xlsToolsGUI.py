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

    def critical(self, pattern, *args):
        QMessageBox.critical(self.box, "错误", pattern.format(*args), QMessageBox.Yes)
class mainWindow():
    def __init__(self, cfgfile):
        self.isForce = False
        self.inputDir = "./xls"
        self.clientOutputDir = "./output/client"
        self.serverOutputDir = "./output/server"
        self.clientOutputType = "lua"
        self.serverOutputType = "lua"
        self.excludeFiles = []
        if os.path.exists(cfgfile):
            with open(cfgfile, 'r') as f:
                cfg = json.load(f)
                self.inputDir = "inputDir" in cfg and cfg["inputDir"] or self.inputDir
                self.clientOutputDir = "clientOutputDir" in cfg and cfg["clientOutputDir"] or self.clientOutputDir
                self.serverOutputDir = "serverOutputDir" in cfg and cfg["serverOutputDir"] or self.serverOutputDir
                self.clientOutputType = "clientOutputType" in cfg and cfg["clientOutputType"] or self.clientOutputType
                self.serverOutputType = "serverOutputType" in cfg and cfg["serverOutputType"] or self.serverOutputType
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

    def onClientOutputTypeClicked(self, box):
        self.clientAllTypeBox.setChecked(False)
        self.clientJsonTypeBox.setChecked(False)
        self.clientLuaTypeBox.setChecked(False)
        box.setChecked(True)
        self.clientOutputType = box.text()

    def onServerOutputTypeClicked(self, box):
        self.serverAllTypeBox.setChecked(False)
        self.serverJsonTypeBox.setChecked(False)
        self.serverLuaTypeBox.setChecked(False)
        box.setChecked(True)
        self.serverOutputType = box.text()

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
        args.exclude_files = self.excludeFiles
        args.force = self.isForce

        args.client_type = self.clientOutputType
        args.client_output_dir = self.clientOutputDir

        args.server_type = self.serverOutputType
        args.server_output_dir = self.serverOutputDir

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

        #======================client==================
        clientOutputGroupBox = QGroupBox("client")
        clientOutputGroupBox.setFlat(False)
        self.clientAllTypeBox = QCheckBox("all")
        self.clientAllTypeBox.setChecked(self.clientOutputType == "all")
        self.clientAllTypeBox.clicked.connect(lambda:self.onClientOutputTypeClicked(self.clientAllTypeBox))
        self.clientJsonTypeBox = QCheckBox("json")
        self.clientJsonTypeBox.setChecked(self.clientOutputType == "json")
        self.clientJsonTypeBox.clicked.connect(lambda:self.onClientOutputTypeClicked(self.clientJsonTypeBox))
        self.clientLuaTypeBox = QCheckBox("lua")
        self.clientLuaTypeBox.setChecked(self.clientOutputType == "lua")
        self.clientLuaTypeBox.clicked.connect(lambda:self.onClientOutputTypeClicked(self.clientLuaTypeBox))
        clientTypeLabel = QLabel(widget)
        clientTypeLabel.setText("导出类型：")
        clientTypeLabel.setAlignment(Qt.AlignCenter)
        clientOutputTypeLayout = QHBoxLayout()
        clientOutputTypeLayout.addWidget(clientTypeLabel, stretch=1)
        clientOutputTypeLayout.addWidget(self.clientAllTypeBox, stretch=1)
        clientOutputTypeLayout.addWidget(self.clientJsonTypeBox, stretch=1)
        clientOutputTypeLayout.addWidget(self.clientLuaTypeBox, stretch=1)
        clientOutputTypeLayout.addStretch(6)

        clientOutputDirLayout = QHBoxLayout()
        clientOutputDirLabel = QLabel(widget)
        clientOutputDirLabel.setText("输出目录：")
        clientOutputDirLabel.setAlignment(Qt.AlignCenter)
        self.clientOutputDirLine = QLineEdit(widget)
        self.clientOutputDirLine.setText(self.clientOutputDir)
        clientOutputDirButton = QPushButton("打开文件夹")
        clientOutputDirButton.clicked.connect(self.onClientOutputDialogClicked)
        clientOutputDirLayout.addWidget(clientOutputDirLabel, stretch=1)
        clientOutputDirLayout.addWidget(self.clientOutputDirLine, stretch=8)
        clientOutputDirLayout.addWidget(clientOutputDirButton, stretch=1)

        clientLayout = QVBoxLayout()
        clientLayout.addLayout(clientOutputTypeLayout)
        clientLayout.addLayout(clientOutputDirLayout)
        clientOutputGroupBox.setLayout(clientLayout)
        #======================end client==================

        #======================server==================
        serverOutputGroupBox = QGroupBox("server")
        serverOutputGroupBox.setFlat(False)
        self.serverAllTypeBox = QCheckBox("all")
        self.serverAllTypeBox.setChecked(self.clientOutputType == "all")
        self.serverAllTypeBox.clicked.connect(lambda:self.onServerOutputTypeClicked(self.serverAllTypeBox))
        self.serverJsonTypeBox = QCheckBox("json")
        self.serverJsonTypeBox.setChecked(self.clientOutputType == "json")
        self.serverJsonTypeBox.clicked.connect(lambda:self.onServerOutputTypeClicked(self.serverJsonTypeBox))
        self.serverLuaTypeBox = QCheckBox("lua")
        self.serverLuaTypeBox.setChecked(self.clientOutputType == "lua")
        self.serverLuaTypeBox.clicked.connect(lambda:self.onServerOutputTypeClicked(self.serverLuaTypeBox))
        serverTypeLabel = QLabel(widget)
        serverTypeLabel.setText("导出类型：")
        serverTypeLabel.setAlignment(Qt.AlignCenter)
        serverOutputTypeLayout = QHBoxLayout()
        serverOutputTypeLayout.addWidget(serverTypeLabel, stretch=1)
        serverOutputTypeLayout.addWidget(self.serverAllTypeBox, stretch=1)
        serverOutputTypeLayout.addWidget(self.serverJsonTypeBox, stretch=1)
        serverOutputTypeLayout.addWidget(self.serverLuaTypeBox, stretch=1)
        serverOutputTypeLayout.addStretch(6)

        serverOutputDirLayout = QHBoxLayout()
        serverOutputDirLabel = QLabel(widget)
        serverOutputDirLabel.setText("输出目录：")
        serverOutputDirLabel.setAlignment(Qt.AlignCenter)
        self.serverOutputDirLine = QLineEdit(widget)
        self.serverOutputDirLine.setText(self.serverOutputDir)
        serverOutputDirButton = QPushButton("打开文件夹")
        serverOutputDirButton.clicked.connect(self.onServerOutputDialogClicked)
        serverOutputDirLayout.addWidget(serverOutputDirLabel, stretch=1)
        serverOutputDirLayout.addWidget(self.serverOutputDirLine, stretch=8)
        serverOutputDirLayout.addWidget(serverOutputDirButton, stretch=1)

        serverLayout = QVBoxLayout()
        serverLayout.addLayout(serverOutputTypeLayout)
        serverLayout.addLayout(serverOutputDirLayout)
        serverOutputGroupBox.setLayout(serverLayout)
        #======================end server==================

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
        mainLayout.addWidget(clientOutputGroupBox)
        mainLayout.addWidget(serverOutputGroupBox)
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
