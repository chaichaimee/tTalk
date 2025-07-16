# -*- coding: utf-8 -*-
# clipboard.py
# Clipboard monitoring module for tTalk
# Copyright (C) 2025 Chai Chaimee
# Licensed under GNU General Public License. See COPYING.txt for details.

import wx
import os
from time import sleep
from datetime import datetime
import addonHandler

addonHandler.initTranslation()

textContent = ""
tempContent = ""

class ClipboardMonitor(object):
    def getClipboard(self):
        clipboard = wx.Clipboard.Get()
        max_retries = 10
        retry_delay = 0.1

        for i in range(max_retries):
            try:
                clipboard.Open()
                break
            except Exception:
                sleep(retry_delay)
                retry_delay *= 2
        else:
            return None

        try:
            if clipboard.IsSupported(wx.DataFormat(wx.DF_FILENAME)):
                file_data = wx.FileDataObject()
                clipboard.GetData(file_data)
                return file_data.GetFilenames()
            elif clipboard.IsSupported(wx.DataFormat(wx.DF_TEXT)):
                text_data = wx.TextDataObject()
                clipboard.GetData(text_data)
                return text_data.GetText()
            return None
        finally:
            if clipboard.IsOpened():
                clipboard.Close()

    def validClipboardData(self):
        data = self.getClipboard()
        if data is None:
            return 0, None

        if isinstance(data, list):
            if len(data) == 1:
                text = _("file/folder ") + os.path.basename(data[0])
            elif len(data) <= 3:
                names = "; ".join([os.path.basename(p) for p in data])
                text = _("files/folders: ") + names
            else:
                text = str(len(data)) + _(" files/folders")
            return 1, text
        elif isinstance(data, str):
            return 2, data[:1024] + "..." if len(data) > 1024 else data
        return 0, None

    def clipboardHasChanged(self):
        global textContent, tempContent
        current_data = self.getClipboard()
        timestamp = datetime.now().strftime("%H:%M:%S.%f")

        if isinstance(current_data, list):
            tempContent = "_FILES_" + ";".join(sorted(current_data)) + timestamp
        elif isinstance(current_data, str):
            tempContent = "_TEXT_" + current_data + timestamp
        else:
            tempContent = "_UNKNOWN_" + timestamp

        if tempContent != textContent:
            textContent = tempContent
            return True
        return False
