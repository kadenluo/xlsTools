#!/usr/bin/python3
# -*-  coding:utf-8 -*-

import sys
import os
import json
import xlrd
import os
import argparse
import traceback
import subprocess
from datetime import datetime
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from xlsTools import convertFiles, getAllFiles, getModifiedFiles
from logger import Logger
from tkinter import messagebox

class FileState:
    Normal = 0  # 000000 正常在svn管理下的最新的文件
    RemoteLocked = 1  # 000001 云端锁定态
    LocalLocked = 2  # 000010 本地锁定态
    Locked = 3  # 000011 已锁定 state and Locked == True
    LocalMod = 4  # 000100 本地有修改需提交
    RemoteMod = 8  # 001000 远程有修改需要更新
    Conflicked = 12  # 001100 冲突 state and Conflicked == Conflicked
    UnVersioned = 16  # 010000 未提交到库
    Error = 32  # 100000 错误状态

class mainWindow(QWidget):
    _config = argparse.ArgumentParser()
    def __init__(self, cfgfile):
        super(mainWindow, self).__init__()
        default_config = {
            "input_dir":  "./xls",
            "client_type": "lua",
            "client_output_dir": "./output/client",
            "server_output_dir": "./output/server",
            "server_type": "lua",
            "exclude_files": [".svn", ".git"],
        }
        if os.path.exists(cfgfile):
            with open(cfgfile, 'r') as f:
                default_config.update(json.load(f))
        self._config = argparse.Namespace(**default_config)

        self.resize(720, 960)
        self.setWindowTitle("转表工具")
        self.setWindowIcon(QIcon("app.ico"))

        # init logger
        loggerBox = QTextEdit()
        loggerBox.setReadOnly(True)
        self.logger = Logger(loggerBox)

        # init ui
        mainLayout = QVBoxLayout()
        self.setLayout(mainLayout)

        # 设置
        label = QLabel("设置") 
        mainLayout.addWidget(label)

        #======================client==================
        clientOutputGroupBox = QGroupBox("client")
        clientOutputGroupBox.setFlat(False)
        self.clientJsonTypeBox = QCheckBox("json")
        self.clientJsonTypeBox.setChecked(self._config.client_type == "json")
        self.clientJsonTypeBox.clicked.connect(lambda:self.onClientOutputTypeClicked(self.clientJsonTypeBox))
        self.clientLuaTypeBox = QCheckBox("lua")
        self.clientLuaTypeBox.setChecked(self._config.client_type == "lua")
        self.clientLuaTypeBox.clicked.connect(lambda:self.onClientOutputTypeClicked(self.clientLuaTypeBox))
        clientTypeLabel = QLabel(self)
        clientTypeLabel.setText("导出类型：")
        clientTypeLabel.setAlignment(Qt.AlignCenter)
        clientOutputTypeLayout = QHBoxLayout()
        clientOutputTypeLayout.addWidget(clientTypeLabel, stretch=1)
        clientOutputTypeLayout.addWidget(self.clientJsonTypeBox, stretch=1)
        clientOutputTypeLayout.addWidget(self.clientLuaTypeBox, stretch=1)
        clientOutputTypeLayout.addStretch(6)

        clientOutputDirLayout = QHBoxLayout()
        clientOutputDirLabel = QLabel(self)
        clientOutputDirLabel.setText("输出目录：")
        clientOutputDirLabel.setAlignment(Qt.AlignCenter)
        self.clientOutputDirLine = QLineEdit(self)
        self.clientOutputDirLine.setText(self._config.client_output_dir)
        clientOutputDirButton = QPushButton("打开文件夹")
        clientOutputDirButton.clicked.connect(self.onClientOutputDialogClicked)
        clientOutputDirLayout.addWidget(clientOutputDirLabel, stretch=1)
        clientOutputDirLayout.addWidget(self.clientOutputDirLine, stretch=8)
        clientOutputDirLayout.addWidget(clientOutputDirButton, stretch=1)

        clientLayout = QVBoxLayout()
        clientLayout.addLayout(clientOutputTypeLayout)
        clientLayout.addLayout(clientOutputDirLayout)
        clientOutputGroupBox.setLayout(clientLayout)

        mainLayout.addWidget(clientOutputGroupBox)
        #======================end client==================
        
        #======================server==================
        serverOutputGroupBox = QGroupBox("server")
        serverOutputGroupBox.setFlat(False)
        self.serverJsonTypeBox = QCheckBox("json")
        self.serverJsonTypeBox.setChecked(self._config.server_type == "json")
        self.serverJsonTypeBox.clicked.connect(lambda:self.onServerOutputTypeClicked(self.serverJsonTypeBox))
        self.serverLuaTypeBox = QCheckBox("lua")
        self.serverLuaTypeBox.setChecked(self._config.server_type == "lua")
        self.serverLuaTypeBox.clicked.connect(lambda:self.onServerOutputTypeClicked(self.serverLuaTypeBox))
        serverTypeLabel = QLabel(self)
        serverTypeLabel.setText("导出类型：")
        serverTypeLabel.setAlignment(Qt.AlignCenter)
        serverOutputTypeLayout = QHBoxLayout()
        serverOutputTypeLayout.addWidget(serverTypeLabel, stretch=1)
        serverOutputTypeLayout.addWidget(self.serverJsonTypeBox, stretch=1)
        serverOutputTypeLayout.addWidget(self.serverLuaTypeBox, stretch=1)
        serverOutputTypeLayout.addStretch(6)

        serverOutputDirLayout = QHBoxLayout()
        serverOutputDirLabel = QLabel(self)
        serverOutputDirLabel.setText("输出目录：")
        serverOutputDirLabel.setAlignment(Qt.AlignCenter)
        self.serverOutputDirLine = QLineEdit(self)
        self.serverOutputDirLine.setText(self._config.server_output_dir)
        serverOutputDirButton = QPushButton("打开文件夹")
        serverOutputDirButton.clicked.connect(self.onServerOutputDialogClicked)
        serverOutputDirLayout.addWidget(serverOutputDirLabel, stretch=1)
        serverOutputDirLayout.addWidget(self.serverOutputDirLine, stretch=8)
        serverOutputDirLayout.addWidget(serverOutputDirButton, stretch=1)

        serverLayout = QVBoxLayout()
        serverLayout.addLayout(serverOutputTypeLayout)
        serverLayout.addLayout(serverOutputDirLayout)
        serverOutputGroupBox.setLayout(serverLayout)

        mainLayout.addWidget(serverOutputGroupBox)
        #======================end server==================

        # ui top
        topLayout = QVBoxLayout()

        self.recentListWidget = QListWidget()
        self.recentListWidget.itemDoubleClicked.connect(self.onListItemClicked)
        scrollBar = QScrollBar()
        self.recentListWidget.addScrollBarWidget(scrollBar, Qt.AlignLeft)
        topLayout.addWidget(self.recentListWidget)
        mainLayout.addLayout(topLayout, stretch = 30)

        # ui center
        centerLayout = QVBoxLayout()

        label = QLabel("通用功能") 
        centerLayout.addWidget(label)

        buttonLayout = QHBoxLayout()

        refreshAllButton = QPushButton("刷新") 
        buttonLayout.addWidget(refreshAllButton)
        refreshAllButton.setFixedWidth(60)
        refreshAllButton.clicked.connect(self.refreshFiles)

        convertAllButton = QPushButton("转全部") 
        buttonLayout.addWidget(convertAllButton)
        convertAllButton.setFixedWidth(60)
        convertAllButton.clicked.connect(self.onConvertAllFileClicked)

        commitAllButton = QPushButton("提交全部") 
        buttonLayout.addWidget(commitAllButton)
        commitAllButton.setFixedWidth(60)
        commitAllButton.clicked.connect(self.onCommitAllFileClicked)

        clearLogButton = QPushButton("清Log") 
        buttonLayout.addWidget(clearLogButton)
        clearLogButton.setFixedWidth(50)
        clearLogButton.clicked.connect(self.onCleanLogClicked)
        centerLayout.addLayout(buttonLayout)

        mainLayout.addLayout(centerLayout, stretch=10)

        # ui bottom
        bottomLayout = QVBoxLayout()
        self.allListWidget = QListWidget()
        scrollBarAll = QScrollBar()
        self.allListWidget.addScrollBarWidget(scrollBarAll, Qt.AlignLeft)
        bottomLayout.addWidget(self.allListWidget)
        mainLayout.addLayout(bottomLayout, stretch=40)

        ##messgebox
        mainLayout.addWidget(loggerBox, stretch=30)

        self.refreshFiles()


    def __str__(self):
        return json.dumps(self.__dict__)

    def getCorrelativeFileNames(self, filepath):
        names = []
        wb = xlrd.open_workbook(filepath)
        for sheet in wb.sheets():
            name = sheet.name.lower()
            if name.startswith("~"):
                return False
            names.append(name)

        return names

    def getCommitFilePathsByFile(self, filepath):
        paths = []
        paths.append(filepath)
        exportNames = self.getCorrelativeFileNames(filepath)
        for export in exportNames:
            paths.append("{}/{}.{}".format(self._config.client_output_dir, export, self._config.client_type))
            paths.append("{}/{}.{}".format(self._config.server_output_dir, export, self._config.server_type))
        return paths
    

    def onListItemClicked(self, item):
        fileName = item.data(1)
        self.logger.info("open xls file:" + str(fileName)) 
        if(fileName == None):
            return

        cmd = "start excel \"{}\"".format(item.data(1))
        # cmd = "start excel \"%s\\%s\"" %(self._config.input_dir, item.data(1))
        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=True
        )
        
    def resetListWidget(self, listWidget, files):
        listWidget.clear()
        for fileName in files:
            item = QListWidgetItem()
            item.setSizeHint(QSize(400, 50))
            item.setData(1, fileName)
            widget = self.getListItemWidget(fileName)
            listWidget.addItem(item)
            listWidget.setItemWidget(item, widget)
        listWidget.itemDoubleClicked.connect(self.onListItemClicked)
    
    def onOpenFileDirClicked(self, fileName):
        self.logger.info("open file:" + fileName)
        path = os.path.abspath(self._config.input_dir)
        cmd = "explorer.exe \"%s\"" %(path)
        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            shell=True
        )
    
    def onCommitFileClicked(self, fileName):
        self.logger.info("commit file:" + fileName)
        paths = self.getCommitFilePathsByFile(fileName)
        abspaths = []
        for path in paths:
             abspaths.append(os.path.abspath(path))
        pathsStr = "*".join(abspaths)

        cmd = "TortoiseGitProc.exe /command:commit /path:\"%s\"" % pathsStr
        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            shell=True
        )
    
    def onRevertFileClicked(self, fileName):
        self.logger.info("revert file:" + fileName)
        paths = self.getCommitFilePathsByFile(fileName)
        abspaths = []
        for path in paths:
             abspaths.append(os.path.abspath(path))
        pathsStr = "*".join(abspaths)

        cmd = "TortoiseGitProc.exe /command:revert /path:\"%s\"" % pathsStr
        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            shell=True
        )
        self.refreshFiles()

    def getListItemWidget(self, fileName):
        widget = QWidget()
        layout_main = QHBoxLayout()

        label = QLabel(fileName) 
        ##设置颜色，根据是否修改，区分红色和绿色
        state = None
        # if fileName in self.fileStatusDict:
        #     state = self.fileStatusDict[fileName]["state"]

        if state == FileState.LocalMod or state == FileState.UnVersioned:
            label.setStyleSheet("background-color: red")
        elif state == FileState.Conflicked:
            label.setStyleSheet("background-color: yellow")
        elif state == FileState.RemoteLocked or state == FileState.LocalLocked:
            label.setStyleSheet("background-color: gray")
        else:
            label.setStyleSheet("background-color: lightgreen") 
        
        layout_main.addWidget(label)

        convertButton = QPushButton("转")
        convertButton.clicked.connect(lambda: self.onConvertFileClicked(fileName))
        layout_main.addWidget(convertButton)
        convertButton.setFixedWidth(40)

        openDirButton = QPushButton("夹")
        openDirButton.clicked.connect(lambda: self.onOpenFileDirClicked(fileName))
        layout_main.addWidget(openDirButton)
        openDirButton.setFixedWidth(40)

        commitButton = QPushButton("交")
        commitButton.clicked.connect(lambda: self.onCommitFileClicked(fileName))
        layout_main.addWidget(commitButton)
        commitButton.setFixedWidth(40)

        revertButton = QPushButton("退")
        revertButton.clicked.connect(lambda: self.onRevertFileClicked(fileName))
        layout_main.addWidget(revertButton)
        revertButton.setFixedWidth(40)

        # if state == FileState.RemoteLocked or state == FileState.LocalLocked:
        #     unlockButton = QPushButton("解锁")
        #     unlockButton.clicked.connect(lambda: self.onUnlockFileClicked(fileName))
        #     layout_main.addWidget(unlockButton)
        #     unlockButton.setFixedWidth(40)
        # else:
        #     lockButton = QPushButton("锁")
        #     lockButton.clicked.connect(lambda: self.onLockFileClicked(fileName))
        #     layout_main.addWidget(lockButton)
        #     lockButton.setFixedWidth(40)
        
        widget.setLayout(layout_main)
        return widget
    
    def refreshFiles(self):
        self.resetListWidget(self.allListWidget, getAllFiles(self._config.input_dir))
        self.resetListWidget(self.recentListWidget, getModifiedFiles(self._config.input_dir))

    def convertFiles(self, files):
        cfg = self._config
        return convertFiles(files, cfg.client_type, cfg.client_output_dir, cfg.server_type, cfg.server_output_dir)

    def onConvertFileClicked(self, fileName):
        self.logger.info("convert file:" + fileName)
        ret = convertFiles([fileName])
        self.refreshFiles()
        # if ret:
        #     QMessageBox.information(self, "Message", "转出表成功")
        # else:
        #     QMessageBox.critical(self, "Error", "转出表失败")

    def onConvertAllFileClicked(self):
        self.logger.info("covert all file")
        ret = self.convertFiles(getAllFiles(self._config.input_dir))
        self.refreshFiles()
        # if ret:
        #     QMessageBox.information(self, "Message", "转出所有表成功")
        # else:
        #     QMessageBox.critical(self, "Error", "转出所有表失败")
    
    def onCommitAllFileClicked(self):
        self.logger.info("commit all file")
        paths = [
            self._config.input_dir,
            self._config.client_output_dir,
            self._config.server_output_dir,
        ]
        abspaths = []
        for path in paths:
             abspaths.append(os.path.abspath(path))
        pathsStr = "*".join(abspaths)

        cmd = "TortoiseGitProc.exe /command:commit /path:\"%s\"" % pathsStr
        p = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            shell=True
        )

    def onCleanLogClicked(self):
        self.logger.clear()

    def onClientOutputTypeClicked(self, box):
        self.clientJsonTypeBox.setChecked(False)
        self.clientLuaTypeBox.setChecked(False)
        box.setChecked(True)
        self._config.client_type = box.text()
    
    def onServerOutputTypeClicked(self, box):
        self.serverJsonTypeBox.setChecked(False)
        self.serverLuaTypeBox.setChecked(False)
        box.setChecked(True)
        self._config.server_type = box.text()
    
    def onClientOutputDialogClicked(self):
        self._config.client_output_dir = QFileDialog.getExistingDirectory(None, "选取文件", "./")
        self.clientOutputDirLine.setText(self._config.client_output_dir)
    
    def onServerOutputDialogClicked(self):
        self._config.server_output_dir = QFileDialog.getExistingDirectory(None, "选取文件", "./")
        self.serverOutputDirLine.setText(self._config.server_output_dir)

if __name__ == "__main__":
    QCoreApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    try:
        window = mainWindow("./config.json")
        window.show()
    except Exception as ex:
        messagebox.showerror(title="错误", message=traceback.format_exc())
    sys.exit(app.exec_())
    