# -*- coding: utf-8 -*-
# tTalk
# Copyright (C) 2025 Chai Chaimee
# Licensed under GNU General Public License. See COPYING.txt for details.

import globalPluginHandler
import keyboardHandler
import speech
import addonHandler
import inputCore
import api
import logHandler
import tones
import treeInterceptorHandler
import textInfos
import controlTypes
import wx
import os
import re
import time
import winUser
import gui
from . import clipboard
import browseMode

addonHandler.initTranslation()

if hasattr(controlTypes, 'State'):
    controlTypes.STATE_SELECTED = controlTypes.State.SELECTED
    controlTypes.STATE_READONLY = controlTypes.State.READONLY

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = _("tTalk")

    def __init__(self):
        super().__init__()
        self.language = "th"
        # Keep clipboard monitor for future use
        self.clipboard_monitor = clipboard.ClipboardMonitor() 

        self.messages = {
            "en": {
                "control+c": "copy",
                "control+v": "paste",
                "control+x": "cut",
                "control+z": "undo",
                "control+a": "select all",
                "control+s": "save",
                "control+shift+c": "copy as path",
                "langSwitched": "Soundtrack",
                "noSelection": "No selection",
                "noFileSelected": "No file selected",
                "fileCopied": "copy file",  # Fixed typo in key
                "clipboardError": "Error accessing clipboard",
                "pastedText": "Pasted text",  # Reserved
                "pastedFile": "Pasted file"   # Reserved
            },
            "th": {
                "control+c": "คัดลอก",
                "control+v": "วาง",
                "control+x": "ตัด",
                "control+z": "ย้อนกลับ",
                "control+a": "เลือกทั้งหมด",
                "control+s": "บันทึก",
                "control+shift+c": "คัดลอกเส้นทาง",
                "langSwitched": "พากย์ไทย",
                "noSelection": "ไม่มีข้อความที่เลือก",
                "noFileSelected": "ไม่มีไฟล์ที่เลือก",
                "fileCopied": "คัดลอกไฟล์",  # Fixed key to match English
                "clipboardError": "ข้อผิดพลาดในการเข้าถึงคลิปบอร์ด",
                "pastedText": "วางข้อความ",   # Reserved
                "pastedFile": "วางไฟล์"       # Reserved
            }
        }

        gestures = {
            "kb:NVDA+alt+t": "toggleLanguage",
            "kb:control+c": "announceCopy",
            "kb:control+v": "announcePaste",
            "kb:control+x": "announceCut",
            "kb:control+z": "announceUndo",
            "kb:control+a": "announceSelectAll",
            "kb:control+s": "announceSave",
            "kb:control+shift+c": "announceCopyAsPath"
        }
        for gesture, script in gestures.items():
            try:
                self.bindGesture(gesture, script)
                logHandler.log.info(f"tTalk: Successfully bound {gesture}")
            except Exception as e:
                speech.speakMessage(f"การผูกคีย์ {gesture} ล้มเหลว")
                logHandler.log.error(f"tTalk: Failed to bind {gesture}: {str(e)}")
                tones.beep(200, 50)

    def _get_selected_text_robust(self, obj_param):
        """
        Retrieves selected text using makeTextInfo or Ctrl+C fallback
        """
        logHandler.log.info("tTalk: _get_selected_text_robust started.")
        current_obj = obj_param
        selected_text = None

        try:
            target_obj_for_text = None
            if hasattr(current_obj, 'treeInterceptor') and isinstance(current_obj.treeInterceptor, browseMode.BrowseModeDocumentTreeInterceptor):
                target_obj_for_text = current_obj.treeInterceptor
                logHandler.log.info("tTalk: Using treeInterceptor for text info.")
            elif hasattr(current_obj, 'makeTextInfo'):
                target_obj_for_text = current_obj
            
            if target_obj_for_text:
                try:
                    info = target_obj_for_text.makeTextInfo(textInfos.POSITION_SELECTION)
                    if info and not info.isCollapsed:
                        selected_text = info.clipboardText
                        if selected_text:
                            logHandler.log.info(f"tTalk: Retrieved text via makeTextInfo")
                            return selected_text.replace('\r\n', '\n').replace('\r', '\n').strip()
                except (RuntimeError, NotImplementedError) as e:
                    logHandler.log.warning(f"tTalk: makeTextInfo failed: {str(e)}")
            else:
                logHandler.log.info("tTalk: No makeTextInfo available")

        except Exception as e_info:
            logHandler.log.error(f"tTalk: makeTextInfo error: {str(e_info)}")

        logHandler.log.info("tTalk: Attempting Ctrl+C fallback")
        original_clipboard_data = ""
        try:
            with winUser.openClipboard(gui.mainFrame.Handle):
                original_clipboard_data = winUser.getClipboardData(winUser.CF_UNICODETEXT) or ""
                winUser.emptyClipboard()
            
            keyboardHandler.injectKey("control+c")
            time.sleep(0.1) 
            
            with winUser.openClipboard(gui.mainFrame.Handle):
                clipboard_text = winUser.getClipboardData(winUser.CF_UNICODETEXT) or ""
            
            if clipboard_text:
                selected_text = clipboard_text
                logHandler.log.info("tTalk: Retrieved text via Ctrl+C fallback")
                return selected_text.replace('\r\n', '\n').replace('\r', '\n').strip()

        except Exception as e_fallback:
            logHandler.log.error(f"tTalk: Ctrl+C fallback failed: {str(e_fallback)}")
        finally:
            try:
                with winUser.openClipboard(gui.mainFrame.Handle):
                    winUser.emptyClipboard()
                    if original_clipboard_data:
                        winUser.setClipboardData(winUser.CF_UNICODETEXT, original_clipboard_data)
            except Exception as e_restore:
                logHandler.log.error(f"tTalk: Clipboard restore failed: {str(e_restore)}")

        logHandler.log.info("tTalk: No selected text found")
        return None

    def script_toggleLanguage(self, gesture):
        self.language = "th" if self.language == "en" else "en"
        try:
            speech.speakMessage(self.messages[self.language]["langSwitched"])
            logHandler.log.info(f"tTalk: Language switched to {self.language}")
        except Exception as e:
            logHandler.log.error(f"tTalk: Language switch error: {str(e)}")
            tones.beep(200, 50)

    def script_announceCopy(self, gesture):
        try:
            obj = api.getFocusObject()
            app_name = obj.appModule.appName if obj.appModule else ""

            # Browser/text field handling
            if app_name in ("chrome", "firefox", "edge", "msedge", "opera", "safari") or \
               hasattr(obj, 'makeTextInfo'):
                selectedText = self._get_selected_text_robust(obj)
                if selectedText:
                    api.copyToClip(selectedText)
                    speech.speakMessage(self.messages[self.language]["control+c"])
                    return

            # File Explorer handling
            gesture.send()
            if app_name == "explorer":
                speech.speakMessage(self.messages[self.language]["fileCopied"])
            else:
                speech.speakMessage(self.messages[self.language]["control+c"])

        except Exception as e:
            logHandler.log.error(f"tTalk: Copy error: {str(e)}")
            speech.speakMessage(self.messages[self.language]["clipboardError"])
            tones.beep(200, 50)

    def script_announceCopyAsPath(self, gesture):
        try:
            gesture.send()
            speech.speakMessage(self.messages[self.language]["control+shift+c"])
        except Exception as e:
            logHandler.log.error(f"tTalk: Copy path error: {str(e)}")
            tones.beep(200, 50)

    def script_announcePaste(self, gesture):
        try:
            gesture.send()
            speech.speakMessage(self.messages[self.language]["control+v"])
        except Exception as e:
            logHandler.log.error(f"tTalk: Paste error: {str(e)}")
            tones.beep(200, 50)

    def script_announceCut(self, gesture):
        try:
            gesture.send()
            time.sleep(0.1)
            speech.speakMessage(self.messages[self.language]["control+x"])
        except Exception as e:
            logHandler.log.error(f"tTalk: Cut error: {str(e)}")
            tones.beep(200, 50)

    def script_announceUndo(self, gesture):
        try:
            gesture.send()
            speech.speakMessage(self.messages[self.language]["control+z"])
        except Exception as e:
            logHandler.log.error(f"tTalk: Undo error: {str(e)}")
            tones.beep(200, 50)

    def script_announceSelectAll(self, gesture):
        try:
            gesture.send()
            speech.speakMessage(self.messages[self.language]["control+a"])
        except Exception as e:
            logHandler.log.error(f"tTalk: Select all error: {str(e)}")
            tones.beep(200, 50)

    def script_announceSave(self, gesture):
        try:
            gesture.send()
            speech.speakMessage(self.messages[self.language]["control+s"])
        except Exception as e:
            logHandler.log.error(f"tTalk: Save error: {str(e)}")
            tones.beep(200, 50)

    # Updated docstrings
    script_toggleLanguage.__doc__ = _("Toggles language")
    script_announceCopy.__doc__ = _("copy")
    script_announceCopyAsPath.__doc__ = _("copy as path")
    script_announceCut.__doc__ = _("cut")
    script_announcePaste.__doc__ = _("paste")
    script_announceSave.__doc__ = _("save")
    script_announceSelectAll.__doc__ = _("select all")
    script_announceUndo.__doc__ = _("undo")
