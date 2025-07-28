from PyQt5 import QtGui, QtCore, QtWidgets
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QMainWindow, 
    QCheckBox, QTextEdit, QDialog, QSplitter, QGridLayout,
    QFileDialog, QLineEdit, QMessageBox, QComboBox, QProgressBar,
    QHBoxLayout, QProgressDialog,QTextBrowser, QTreeWidget, QTreeWidgetItem,
    QScrollArea, QGroupBox, QLayout, QSpacerItem, 
    QSizePolicy, QProxyStyle, QStyleFactory, QFrame)
from PyQt5.QtGui import (
    QIcon, QPixmap, QCursor, QFont, 
    QColor, QPalette, QTextDocument, QTextOption,  
    QTextCharFormat, QFontDatabase, QTextFrameFormat, QTextBlockFormat, 
    QTextListFormat, QTextFrame, QSyntaxHighlighter, QFontMetrics,
    QTextBlock, QTextList, QTextCursor,)
from PyQt5.QtWinExtras import QtWin
from PyQt5.QtCore import (
    Qt, QDateTime, QThread, pyqtSignal, QTimer, QObject, QEvent,
    QPropertyAnimation, QEasingCurve, QRegExp, QLocale,)
import win32con
import win32gui
import sys
import ctypes
import re
import os
import requests
import shutil
import base64
import urllib.request
import time
import gematriapy
import os.path
import traceback
import certifi
import win32api
import ssl
import requests.adapters
import logging
import json
import subprocess
from pyluach import dates
from datetime import datetime
from pathlib import Path
from pyluach import gematria
from bs4 import BeautifulSoup
from functools import partial
from packaging import version
from ctypes import wintypes
from logging import info
from urllib3.util.ssl_ import create_urllib3_context
from bidi.algorithm import get_display

# גרסת התוכנה - משתנה גלובלי מרכזי
VERSION = "3.5.0"

# מזהה ייחודי לאפליקציה
myappid = 'MIT.LEARN_PYQT.dictatootzaria'

# מחרוזת Base64 של האייקון
icon_base64 = "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"

class EnterKeyFilter(QObject):
    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress and event.key() in (Qt.Key_Return, Qt.Key_Enter):
            if isinstance(obj, QPushButton) and obj.hasFocus():
                obj.click()
                return True
            elif isinstance(obj, QComboBox) and obj.hasFocus():
                # נפעיל את האירוע של בחירת הפריט הנוכחי
                if not obj.view().isVisible():
                    obj.showPopup()
                    return True
            elif isinstance(obj, QCheckBox) and obj.hasFocus():
                # נשנה את מצב התיבת הסימון
                obj.toggle()
                return True                
        return super().eventFilter(obj, event)
    
# ==========================================
# Script 1: יצירת כותרות לאוצריא
# ==========================================

class CreateHeadersOtZria(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("יצירת כותרות לאוצריא")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setGeometry(100, 100, 1150, 400)
        self.setContentsMargins(30, 30, 30, 30)  # הוספת שוליים
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
       
        # נתיב קובץ
        file_layout = QHBoxLayout()
        browse_button = QPushButton("עיון")
        browse_button.clicked.connect(self.browse_file)
        browse_button.setStyleSheet("font-size: 20px;")
        self.file_entry = QLineEdit()
        file_label = QLabel("נתיב קובץ:")
        file_label.setStyleSheet("font-size: 20px;")
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_entry)
        file_layout.addWidget(browse_button)
        layout.addLayout(file_layout)

        # מילה לחיפוש
        search_layout = QHBoxLayout()
        search_label = QLabel("מילה לחפש:")
        search_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        search_label.setStyleSheet("font-size: 26px;")
        self.level_var = QComboBox()
        self.level_var.setStyleSheet("font-size: 26px;")
        self.level_var.setFixedSize(150, 40)  # רוחב: 150 פיקסלים, גובה: 40 פיקסלים
        search_choices = ["דף", "עמוד", "פרק", "פסוק", "שאלה", "סימן", "סעיף", "הלכה", "הלכות", "סק", "ענף"]
        self.level_var.addItems(search_choices)
        self.level_var.setEditable(True)  # מאפשר להקליד
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.level_var, alignment=Qt.AlignLeft | Qt.AlignVCenter)
       
        layout.addLayout(search_layout)

        # הסבר למשתמש
        explanation = QLabel(
            "בתיבת 'מילה לחפש' יש לבחור או להקליד את המילה בה אנו רוצים שתתחיל הכותרת.\nלדוג': פרק/פסוק/סימן/סעיף/הלכה/שאלה/עמוד/סק/ענף\n\nשים לב!\nאין להקליד רווח אחרי המילה, וכן אין להקליד את התו גרש (') או גרשיים (\"), וכן אין להקליד יותר ממילה אחת\n"
        )
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setStyleSheet("font-size: 20px;")
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # מספר סימן מקסימלי
        end_layout = QHBoxLayout()
        end_label = QLabel("מספר סימן מקסימלי:")
        self.end_var = QComboBox()
        self.end_var.addItems([str(i) for i in range(1, 1000)])
        self.end_var.setCurrentText("999")
        self.end_var.setFixedWidth(68)
        end_layout.addWidget(end_label)
        end_layout.addWidget(self.end_var)
        layout.addLayout(end_layout)

        # רמת כותרת
        heading_layout = QHBoxLayout()
        self.heading_label = QLabel("רמת כותרת:")
        self.heading_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.heading_label.setStyleSheet("font-size: 26px;")
        self.heading_level_var = QComboBox()
        self.heading_level_var.addItems([str(i) for i in range(2, 7)])
        self.heading_level_var.setCurrentText("2")
        self.heading_level_var.setStyleSheet("font-size: 26px;")
        self.heading_level_var.setFixedWidth(50)
        heading_layout.addWidget(self.heading_label)
        heading_layout.addWidget(self.heading_level_var, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        layout.addLayout(heading_layout)
           
        # כפתור הפעלה
        run_button = QPushButton("הפעל")
        run_button.clicked.connect(self.run_script)
        run_button.setFixedHeight(40)
        run_button.setStyleSheet("font-size: 25px;")
        layout.addWidget(run_button)

        self.setLayout(layout)

    def browse_file(self):
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getOpenFileName(
            self, "בחר קובץ", "", "קבצי טקסט (*.txt);;כל הפורמטים (*.*)", options=options
        )
        if filename:
            self.file_entry.setText(filename)

    def show_custom_message(self, title, message_parts, window_size=("560x330")):
        msg = QMessageBox(self)
        msg.setWindowTitle(title)

        # בניית ההודעה עם גודל גופן שונה
        full_message = ""
        for part in message_parts:
            if len(part) == 3 and part[2] == "bold":
                full_message += f"<b><span style='font-size:{part[1]}pt'>{part[0]}</span></b><br>"
            else:
                full_message += f"<span style='font-size:{part[1]}pt'>{part[0]}</span><br>"

        msg.setTextFormat(Qt.RichText)
        msg.setText(full_message)
        msg.exec_()

    def ot(self, text, end):
        remove = ["<b>", "</b>", "<big>", "</big>", ":", '"', ",", ";", "[", "]", "(", ")", "'", ".", "״", "‚", "”", "’"]
        aa = ["ק", "ר", "ש", "ת", "תק", "תר", "תש", "תת", "תתק", "יה", "יו", "קיה", "קיו", "ריה", "ריו", "שיה", "שיו", "תיה", "תיו", "תקיה", "תקיו", "תריה", "תריו", "תשיה", "תשיו", "תתיה", "תתיו", "תתקיה", "תתקיו"]
        bb = ["ם", "ן", "ץ", "ף", "ך"]
        cc = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "ששי", "שביעי", "שמיני", "תשיעי", "עשירי", "יוד", "למד", "נון", "דש", "חי", "טל", "שדמ", "ער", "שדם", "תשדם", "תשדמ", "ערב", "ערה", "עדר", "רחצ"]
        append_list = []
        for i in aa:
            for ot_sofit in bb:
                append_list.append(i + ot_sofit)

        for tage in remove:
            text = text.replace(tage, "")
        withaute_gershayim = [gematria._num_to_str(i, thousands=False, withgershayim=False) for i in range(1, end)] + bb + cc + append_list
        return text in withaute_gershayim

    def strip_html_tags(self, text):
        html_tags = ["<b>", "</b>", "<big>", "</big>", ":", '"', ",", ";", "[", "]", "(", ")", "'", "״", ".", "‚", "”", "’"]
        for tag in html_tags:
            text = text.replace(tag, "")
        return text

    def main(self, book_file, finde, end, level_num):
        found = False
        count_headings = 0
        finde_cleaned = self.strip_html_tags(finde).strip()
        with open(book_file, "r", encoding="utf-8") as file_input:
            content = file_input.read().splitlines()
            all_lines = content[0:2]
            for line in content[2:]:
                words = line.split()
                try:
                    if self.strip_html_tags(words[0]) == finde and self.ot(words[1], end):
                        found = True
                        count_headings += 1
                        heading_line = f"<h{level_num}>{self.strip_html_tags(words[0])} {self.strip_html_tags(words[1])}</h{level_num}>"
                        all_lines.append(heading_line)
                        if words[2:]:
                            fix_2 = " ".join(words[2:])
                            all_lines.append(fix_2)
                    else:
                        all_lines.append(line)
                except IndexError:
                    all_lines.append(line)
        join_lines = "\n".join(all_lines)
        with open(book_file, "w", encoding="utf-8") as autpoot:
            autpoot.write(join_lines)

        return found, count_headings

    def run_script(self):
        book_file = self.file_entry.text()
        finde = self.level_var.currentText()
        try:
            end = int(self.end_var.currentText())
            level_num = int(self.heading_level_var.currentText())
        except ValueError:
            self.show_custom_message(
                "קלט לא תקין",
                [("אנא הזן 'מספר סימן מקסימלי' ו'רמת כותרת' תקינים", 12)],
                "250x150"
            )
            return

        if not book_file or not finde:
            self.show_custom_message(
                "קלט לא תקין",
                [("אנא מלא את כל השדות", 12)],
                "250x80"
            )
            return

        try:
            found, count_headings = self.main(book_file, finde, end + 1, level_num)
            if found and count_headings > 0:
                detailed_message = [
                    ("<div style='text-align: center;'>התוכנה רצה בהצלחה!</div>", 12),
                    (f"<div style='text-align: center;'>נוצרו {count_headings} כותרות</div>", 15, "bold"),
                    ("<div style='text-align: center;'>כעת פתח את הספר בתוכנת 'אוצריא', והשינויים ישתקפו ב'ניווט' שבתפריט הצידי.</div>", 11),
                    ("<div style='text-align: center;'>אם ישנם טעויות או תיקונים, פתח את הספר בעורך טקסט, כגון פנקס רשימות, וורד, ++notepad או vs code, ותקן את הדרוש תיקון</div>", 11),
                    ("<div style='text-align: center;'>שים לב! אם הספר כבר פתוח ב'אוצריא', יש לסגור אותו ולפתוח אותו שוב, אך אין צורך להפעיל את התוכנה מחדש</div>", 9),
                ]
                self.show_custom_message("!מזל טוב", detailed_message, "560x310")
            else:
                self.show_custom_message("!שים לב", [("לא נמצא מה להחליף", 12)], "250x80")
        except Exception as e:
            self.show_custom_message("שגיאה", [("אירעה שגיאה: " + str(e), 12)], "250x150")

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
   
# ==========================================
# Script 2: יצירת כותרות לאותיות בודדות
# ==========================================

# מסתיר חלק מהתפריט ב QLineEdit
class AdvancedProtectedLineEdit(QLineEdit):
    def __init__(self, hidden_prefix="", visible_text="", parent=None):
        super().__init__(parent)
        self.hidden_prefix = hidden_prefix
        self.setText(visible_text)
        
    def get_full_text(self):
        """מחזיר את הטקסט המלא כולל החלק המוסתר"""
        return self.hidden_prefix + self.text()
    
    def get_visible_text(self):
        """מחזיר רק את החלק הנראה"""
        return self.text()
"""
לכתוב 
"""
    
class CreateSingleLetterHeaders(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("יצירת כותרות לאותיות בודדות")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setGeometry(100, 100, 1130, 350)
        self.setContentsMargins(30, 30, 30, 30)  # הוספת שוליים
        self.setLayoutDirection(Qt.RightToLeft)  # הגדרת כיוון ימין לשמאל
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # נתיב קובץ
        file_layout = QHBoxLayout()
        browse_button = QPushButton("עיון")
        browse_button.setStyleSheet("font-size: 20px;")
        browse_button.clicked.connect(self.browse_file)
        self.file_entry = QLineEdit()
        file_label = QLabel("נתיב קובץ:")
        file_label.setStyleSheet("font-size: 20px;")
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_entry)
        file_layout.addWidget(browse_button)
        layout.addLayout(file_layout)

        only_label = QLabel("חפש רק אם קיים:")
        only_label.setStyleSheet("font-weight: bold;")

        # תו בתחילת האות ותו בסוף האות
        start_char_label = QLabel("תו בתחילת האות:")
        self.start_var = QComboBox()
        #self.start_var.setLayoutDirection(Qt.RightToLeft)  # הגדרת כיוון כללי
        self.start_var.addItems(["", "(", "["])
        #self.start_var.setStyleSheet("text-align: right;")  # מוודא שהטקסט ייושר לימין
        
        end_char_label = QLabel("תו/ים בסוף האות:")
        self.finde_var = QComboBox()
        self.finde_var.addItems(['', '.', ',', "'", "',",
            "'.", ']', ')', "']", "')", "].",
            ").", "],", "),", "'),", "').", "'],", "']."]) 
        
        # תיבת סימון לחיפוש עם תווי הדגשה בלבד
        self.bold_var = QCheckBox("לחפש אותיות מודגשות בלבד")
        self.bold_var.setChecked(True)
        
        regex_layout = QHBoxLayout()
        regex_layout.addWidget(self.bold_var)
        regex_layout.addStretch()
        regex_layout.addStretch()
        regex_layout.addStretch()
        regex_layout.addWidget(only_label)
        regex_layout.addStretch()
        regex_layout.addWidget(start_char_label)
        regex_layout.addWidget(self.start_var)
        regex_layout.addStretch()
        regex_layout.addWidget(end_char_label)
        regex_layout.addWidget(self.finde_var)
        layout.addLayout(regex_layout)
       
        # הסבר למשתמש
        explanation = QLabel(
            "שים לב! הבחירה בברירת מחדל [האפשרות הראשונה הריקה], משמעותה סימון כל סוגי התווים שלפני ואחרי הכותרות"
        )
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setStyleSheet("font-size: 18px;")
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # רמת כותרת
        heading_layout = QHBoxLayout()
        heading_label = QLabel("רמת כותרת:")
        heading_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        heading_label.setStyleSheet("font-size: 30px;")
        self.level_var = QComboBox()
        self.level_var.setFixedWidth(45)
        self.level_var.setStyleSheet("font-size: 30px;")
        self.level_var.addItems([str(i) for i in range(2, 7)])
        self.level_var.setCurrentText("3")
        heading_layout.addWidget(heading_label)
        heading_layout.addWidget(self.level_var, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        layout.addLayout(heading_layout)

        # התעלם מהתווים
        ignore_remove_layout = QHBoxLayout()
        ignore_label = QLabel("התעלם מהתווים הבאים:")
        ignore_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.ignore_entry = QLineEdit()
        self.ignore_entry.setText('<big> </big> <i> </i> </small> </small> <span> </span> <br> </br> <p> </p>')
        self.ignore_entry.setStyleSheet("font-size: 20px;")

        # הסרת תווים
        remove_label = QLabel("הסר את התווים הבאים מהכותרות שיווצרו:")
        self.remove_entry = QLineEdit()
        self.remove_entry.setText(', : " \' . ( ) [ ] { }')
        self.remove_entry.setStyleSheet("font-size: 20px;")

        # hidden_chars = '<b> </b> <big> </big> <i> </i> </small> </small> <span> </span> <br> </br> <p> </p>'
        # self.remove_entry = AdvancedProtectedLineEdit(
        #     hidden_prefix=hidden_chars,
        #     visible_text='( ) [ ] { } : \' ’ " ” ״ . , ;')  # העבר את הטקסט כאן במקום setText
        # # self.remove_entry.setText(', : " \' . ( ) [ ] { }')
        # # self.remove_entry.setFixedWidth(180)
        # self.remove_entry.setStyleSheet("font-size: 20px;")

        explanation3 = QLabel("נועד כדי ליצור כותרות 'נקיות', בלי תווים נוספים")

        # ignore_layout.addWidget(ignore_label)
        # ignore_layout.addWidget(self.ignore_entry)
        # ignore_layout.addStretch()
        ignore_remove_layout.addWidget(remove_label)
        ignore_remove_layout.addWidget(self.remove_entry)  # , alignment=Qt.AlignLeft)
        ignore_remove_layout.addStretch()
        ignore_remove_layout.addWidget(explanation3)
        layout.addLayout(ignore_remove_layout)

        # הסבר למשתמש
        explanation2 = QLabel(
            "באם הנך רוצה להשאיר סימנים (כגון סימני סוגריים) "
            "סביב הכותרות, או להשאיר גרש וגרשיים בתוך הכותרות, מחק תווים אלו משורה זו"
        )
        explanation2.setStyleSheet("font-size: 18px;")
        explanation2.setWordWrap(True)
        explanation2.setAlignment(Qt.AlignCenter)
        layout.addWidget(explanation2)

        # מספר סימן מקסימלי
        end_layout = QHBoxLayout()
        end_label = QLabel("מספר סימן מקסימלי:")
        self.end_var = QComboBox()
        self.end_var.setFixedWidth(68)
        self.end_var.addItems([str(i) for i in range(1, 1000)])
        self.end_var.setCurrentText("999")
        end_layout.addWidget(end_label)
        end_layout.addWidget(self.end_var)
        layout.addLayout(end_layout)

        # כאן להוסיף לבל של הסבר שכדאי לגבות את הקובץ לפני הפעלת תוכנה זו
        self.word_count_label = QLabel("מומלץ מאוד ליצור גיבוי של הספר לפני הפעלת תוכנה זו!!\n")
        self.word_count_label.setStyleSheet("font-size: 20px; font-weight: bold;")
        self.word_count_label.setWordWrap(True)
        self.word_count_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.word_count_label)

        # כפתור הפעלה
        run_button = QPushButton("הפעל")
        run_button.clicked.connect(self.run_script)
        run_button.setFixedHeight(40)
        run_button.setStyleSheet("font-size: 25px;")
        layout.addWidget(run_button)

        self.setLayout(layout)

    def browse_file(self):
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getOpenFileName(
            self, "בחר קובץ", "", "קבצי טקסט (*.txt);;כל הפורמטים (*.*)", options=options
        )
        if filename:
            self.file_entry.setText(filename)

    def run_script(self):
        book_file = self.file_entry.text()
        finde = self.finde_var.currentText()
        remove = ["<b>", "</b>", "<big>", "</big>", "<i>", "</i>", "</small>", "</small>", "<span>", "</span>", "<br>", "</br>", "<p>", "</p>"] + self.remove_entry.text().split()
        ignore = self.ignore_entry.text().split()
        start = self.start_var.currentText()
        is_bold_checked = self.bold_var.isChecked()
        
        if is_bold_checked:
            finde += "</b>"
            start = "<b>" + start
        else:
            ignore += ["<b>", "</b>"]

        try:
            end = int(self.end_var.currentText())
            level_num = int(self.level_var.currentText())
        except ValueError:
            QMessageBox.critical(self, "קלט לא תקין", "אנא הזן 'מספר סימן מקסימלי' ו'רמת כותרת' תקינים")
            return

        if not book_file:
            QMessageBox.critical(self, "קלט לא תקין", "אנא בחר קובץ")
            return

        try:
            self.main(book_file, finde, end + 1, level_num, ignore, start, remove)
            QMessageBox.information(self, "!מזל טוב", "התוכנה רצה בהצלחה!")
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"אירעה שגיאה: {e}")

    def ot(self, text, end):
        remove = ["<b>", "</b>", "<big>", "</big>", "<i>", "</i>", "</small>", "</small>", "<span>", "</span>", "<br>", "</br>", "<p>", "</p>", ":", '"', ",", ";", "[", "]", "(", ")", "{", "}", "'", ".", "״", "”", "’"]
        aa = ["ק", "ר", "ש", "ת", "תק", "תר", "תש", "תת", "תתק", "יו", "קיה", "קיו"]
        bb = ["ם", "ן", "ץ", "ף", "ך"]
        cc = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "ששי", "שביעי", "שמיני", "תשיעי", "עשירי", "חי", "יוד", "למד", "נון", "טל", "דש", "שדמ", "ער", "שדם", "תשדם", "תשדמ", "ערה", "ערב", "עדר", "רחצ"]
        append_list = []
        for i in aa:
            for ot_sofit in bb:
                append_list.append(i + ot_sofit)

        for tage in remove:
            text = text.replace(tage, "")
        withaute_gershayim = [gematria._num_to_str(i, thousands=False, withgershayim=False) for i in range(1, end)] + bb + cc + append_list
        return text in withaute_gershayim

    def strip_html_tags(self, text, ignore=None):
        if ignore is None:
            ignore = []
        for tag in ignore:
            text = text.replace(tag, "")
        return text

    def main(self, book_file, finde, end, level_num, ignore, start, remove):
        with open(book_file, "r", encoding="utf-8") as file_input:
            content = file_input.read().splitlines()
            all_lines = content[0:1]
            for line in content[1:]:
                words = line.split()
                try:
                    if self.strip_html_tags(words[0], ignore).endswith(finde) and self.ot(words[0], end) and self.strip_html_tags(words[0], ignore).startswith(start):
                        heading_line = f"<h{level_num}>{self.strip_html_tags(words[0], remove)}</h{level_num}>"
                        all_lines.append(heading_line)
                        if words[1:]:
                            fix_2 = " ".join(words[1:])
                            all_lines.append(fix_2)
                    else:
                        all_lines.append(line)
                except IndexError:
                    all_lines.append(line)
        join_lines = "\n".join(all_lines)
        with open(book_file, "w", encoding="utf-8") as autpoot:
            autpoot.write(join_lines)

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
   
# ==========================================
# Script 3: הוספת מספר עמוד בכותרת הדף
# ==========================================
class AddPageNumberToHeading(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("הוספת מספר עמוד בכותרת הדף")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setLayoutDirection(Qt.RightToLeft)
        self.setGeometry(100, 100, 1100, 400)
        self.setContentsMargins(40, 40, 40, 40)  # הוספת שוליים

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # הסבר למשתמש
        explanation = QLabel(
            "התוכנה מחליפה בקובץ בכל מקום שיש כותרת 'דף' ובתחילת שורה הבאה כתוב: ע\"א או ע\"ב, כגון:\n\n"
            "<h2>דף ב</h2>\nע\"א [טקסט כלשהו]\n\n"
            "הפעלת התוכנה תעדכן את הכותרת ל:\n\n"
            "<h2>דף ב.</h2>\n[טקסט כלשהו]\n"
        )
        explanation.setStyleSheet("font-size: 20px;")
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # נתיב קובץ
        file_layout = QHBoxLayout()
        file_label = QLabel("נתיב קובץ:")
        file_label.setStyleSheet("font-size: 20px;")
        self.file_entry = QLineEdit()
        browse_button = QPushButton("עיון")
        browse_button.setStyleSheet("font-size: 20px;")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_entry)
        file_layout.addWidget(browse_button)
        self.file_entry.setStyleSheet("font-size: 20px;")
        layout.addLayout(file_layout)

        # סוג ההחלפה
        heading_layout = QHBoxLayout()
        replacement_label = QLabel("בחר את סוג ההחלפה:")
        replacement_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        replacement_label.setStyleSheet("font-size: 20px;")
        self.replace_option = QComboBox()
        self.replace_option.setStyleSheet("font-size: 20px;")
        self.replace_option.addItems(["נקודה ונקודותיים", "ע\"א וע\"ב"])
        self.replace_option.setFixedWidth(180)
        heading_layout.addWidget(replacement_label)
        heading_layout.addWidget(self.replace_option, alignment=Qt.AlignLeft | Qt.AlignVCenter)
       
        layout.addLayout(heading_layout)

        # דוגמאות
        example1 = QLabel("לדוגמא:\nדף ב.   דף ב:   דף ג.   דף ג: דף ד. דף ד:\nוכן הלאה")
        example1.setAlignment(Qt.AlignCenter)
        example1.setStyleSheet("font-size: 16px;")
        example1.setWordWrap(True)
        layout.addWidget(example1)

        example2 = QLabel("או:\nדף ב ע\"א   דף ב ע\"ב   דף ג ע\"א   דף ג ע\"ב\nוכן הלאה")
        example2.setAlignment(Qt.AlignCenter)
        example2.setStyleSheet("font-size: 16px;")
        example2.setWordWrap(True)
        layout.addWidget(example2)

        # כפתור הפעלה
        run_button = QPushButton("בצע החלפה")
        run_button.clicked.connect(self.run_script)
        run_button.setFixedHeight(40)
        run_button.setStyleSheet("font-size: 25px;")
        layout.addWidget(run_button)

        self.setLayout(layout)

    def browse_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "בחר קובץ טקסט", "", "קבצי טקסט (*.txt);;HTML קבצי (*.html)", options=options
        )
        if file_path:
            self.file_entry.setText(file_path)
            self.sender().setText("קובץ נבחר, המשך לבחירת סוג ההחלפה")

    def process_file(self, filename, replace_with):
        try:
            with open(filename, 'r', encoding='utf-8') as file:
                content = file.readlines()
        except FileNotFoundError:
            QMessageBox.critical(self, "קלט לא תקין", "הקובץ לא נמצא")
            return
        except UnicodeDecodeError:
            QMessageBox.critical(self, "קלט לא תקין", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return
        except Exception as e:
            QMessageBox.critical(self, "קלט לא תקין", f"שגיאה בפתיחת קובץ: {e}")
            return

        changes_made = False
        updated_content = []
        i = 0

        while i < len(content):
            line = content[i]
            match = re.match(r'<h([2-9])>(דף \S+)</h\1>', line)
            if match:
                level = match.group(1)
                title = match.group(2)
                next_line_index = i + 1
                if next_line_index < len(content):
                    next_line = content[next_line_index].strip()

                    pattern = r'(<[a-z]+>)?(ע["\']+?[א-ב]|עמוד [א-ב])[.,:()\[\]\'"״׳]?(</[a-z]+>)?\s?'
                    match_next_line = re.match(pattern, next_line)

                    if match_next_line:
                        changes_made = True

                        if replace_with == 'נקודה ונקודותיים':
                            if "א" in match_next_line.group(2):
                                new_title = f'<h{level}>{title.rstrip(".")}.</h{level}>\n'
                            else:
                                new_title = f'<h{level}>{title.rstrip(".")}:</h{level}>\n'
                        elif replace_with == 'ע\"א וע\"ב':
                            suffix = "ע\"א" if "א" in match_next_line.group(2) else "ע\"ב"
                            new_title = f'<h{level}>{title.rstrip(".")} {suffix}</h{level}>\n'

                        updated_content.append(new_title)

                        modified_next_line = re.sub(pattern, '', next_line, count=1).strip()
                        if modified_next_line != '':
                            updated_content.append(modified_next_line + '\n')

                        i += 1
                    else:
                        updated_content.append(line)
                else:
                    updated_content.append(line)
            else:
                updated_content.append(line)
            i += 1

        if changes_made:
            with open(filename, 'w', encoding='utf-8') as file:
                file.writelines(updated_content)
            QMessageBox.information(self, "!מזל טוב", "ההחלפה הושלמה בהצלחה!")
        else:
            msg = QMessageBox(self)
            msg.setWindowTitle("!שים לב")
            msg.setText("אין מה להחליף בקובץ זה")
            QTimer.singleShot(2500, msg.close)  # סוגר את ההודעה לאחר 2500 מילי־שניות (2.5 שניות)
            msg.show()
            return

    def run_script(self):
        file_path = self.file_entry.text()
        if file_path:
            replace_with = self.replace_option.currentText()
            self.process_file(file_path, replace_with)
        else:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)            
            msg.setWindowTitle("קלט לא תקין")
            msg.setText("נא לבחור קובץ תחילה")
            QTimer.singleShot(1000, msg.close)  # סוגר את ההודעה לאחר 1000 מילי־שניות (1 שניות)
            msg.show()
            return
  
    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
   
# ==========================================
# Script 4: שינוי רמת כותרת
# ==========================================
class ChangeHeadingLevel(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("שינוי רמת כותרת")
        self.setGeometry(100, 100, 950, 300)
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setLayoutDirection(Qt.RightToLeft)
        self.setContentsMargins(30, 30, 30, 30)  # הוספת שוליים
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # נתיב קובץ
        file_layout = QHBoxLayout()
        browse_button = QPushButton("עיון")
        browse_button.setStyleSheet("font-size: 20px;")
        browse_button.clicked.connect(self.browse_file)
        self.file_entry = QLineEdit()
        file_label = QLabel("נתיב קובץ:")
        file_label.setStyleSheet("font-size: 20px;")
        self.file_entry.setStyleSheet("font-size: 20px;")
        file_layout.addWidget(file_label) 
        file_layout.addWidget(self.file_entry)
        file_layout.addWidget(browse_button)
        layout.addLayout(file_layout)

        # רמת כותרת נוכחית
        current_level_layout = QHBoxLayout()
        current_level_label = QLabel("רמת כותרת נוכחית: (לדוגמא: 2)")
        current_level_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        current_level_label.setStyleSheet("font-size: 20px;")
        self.current_level_var = QComboBox()
        self.current_level_var.setStyleSheet("font-size: 20px;")
        self.current_level_var.setFixedWidth(50)
        self.current_level_var.addItems([str(i) for i in range(1, 10)])
        current_level_layout.addWidget(current_level_label)
        current_level_layout.addWidget(self.current_level_var, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        layout.addLayout(current_level_layout)

        # רמת כותרת חדשה
        new_level_layout = QHBoxLayout()
        new_level_label = QLabel("רמת כותרת חדשה: (לדוגמא: 3)")
        new_level_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        new_level_label.setStyleSheet("font-size: 20px;")
        self.new_level_var = QComboBox()
        self.new_level_var.setStyleSheet("font-size: 20px;")
        self.new_level_var.setFixedWidth(50)
        self.new_level_var.addItems([str(i) for i in range(1, 10)])
        new_level_layout.addWidget(new_level_label)
        new_level_layout.addWidget(self.new_level_var, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        layout.addLayout(new_level_layout)

        # כפתור הפעלה
        run_button = QPushButton("שנה רמת כותרת")
        run_button.clicked.connect(self.run_script)
        run_button.setFixedHeight(40)
        run_button.setStyleSheet("font-size: 25px;")
        layout.addWidget(run_button)

        self.setLayout(layout)

    def browse_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "בחר קובץ טקסט", "", "קבצי טקסט (*.txt);;HTML קבצי (*.html)", options=options
        )
        if file_path:
            self.file_entry.setText(file_path)

    def change_heading_level_func(self, file_path, current_level, new_level):
        file_path = self.file_entry.text()
        if not file_path:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)            
            msg.setWindowTitle("קלט לא תקין")
            msg.setText("נא לבחור קובץ תחילה")
            QTimer.singleShot(1000, msg.close)  # סוגר את ההודעה לאחר 1000 מילי־שניות (1 שניות)
            msg.show()
            return
        # בדיקת סוג הקובץ לפי סיומת
        if Path(file_path).suffix.lower() != '.txt':
            QMessageBox.critical(self, "קלט לא תקין", "סוג הקובץ אינו נתמך\nבחר קובץ טקסט [בסיומת TXT.]")
            return
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()

                current_tag = f"h{current_level}"
                new_tag = f"h{new_level}"
                updated_content = re.sub(f"<{current_tag}>(.*?)</{current_tag}>", f"<{new_tag}>\\1</{new_tag}>", content, flags=re.DOTALL)

                if content == updated_content:
                    msg = QMessageBox(self)
                    msg.setWindowTitle("!שים לב")
                    msg.setText("אין מה להחליף בקובץ זה")
                    QTimer.singleShot(2500, msg.close)  # סוגר את ההודעה לאחר 2500 מילי־שניות (2.5 שניות)
                    msg.show()
                    return
                else:
                    with open(file_path, 'w', encoding='utf-8') as file:
                        file.write(updated_content)
                    QMessageBox.information(self, "!מזל טוב", "רמות הכותרות עודכנו בהצלחה!")
        
        except FileNotFoundError:
            QMessageBox.critical(self, "קלט לא תקין", "הקובץ לא נמצא")
            return
        except UnicodeDecodeError:
            QMessageBox.critical(self, "קלט לא תקין", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"ארעה שגיאה במהלך העיבוד: {str(e)}")    

    def run_script(self):
        file_path = self.file_entry.text()
        current_level = self.current_level_var.currentText()
        new_level = self.new_level_var.currentText()

        if not current_level.isdigit() or not new_level.isdigit():
            QMessageBox.critical(self, "קלט לא תקין", "אנא הכנס רמות מספריות חוקיות")
            return

        self.change_heading_level_func(file_path, current_level, new_level)

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
   
# ==========================================
# Script 5: הדגשת מילה ראשונה וניקוד בסוף קטע
# ==========================================
class EmphasizeAndPunctuate(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("הדגשה וניקוד")
        self.setGeometry(100, 100, 900, 300)
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setLayoutDirection(Qt.RightToLeft)
        self.setContentsMargins(40, 40, 40, 40)  # הוספת שוליים
        self.setStyleSheet("font-size: 20px;")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # נתיב קובץ
        file_layout = QHBoxLayout()
        browse_button = QPushButton("עיון")
        browse_button.setStyleSheet("font-size: 20px;")
        browse_button.clicked.connect(self.select_file)
        self.file_path_entry = QLineEdit()
        file_label = QLabel("נתיב קובץ:")
        file_label.setStyleSheet("font-size: 20px;")
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_path_entry)
        file_layout.addWidget(browse_button)
        layout.addLayout(file_layout)
       
        # בחירה להוספת נקודה או נקודותיים
        ending_layout = QHBoxLayout()  # שינוי ל-QHBoxLayout
        ending_label = QLabel("בחר פעולה לסוף קטע:")
        ending_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        ending_label.setStyleSheet("font-size: 20px;")
        self.ending_var = QComboBox()
        self.ending_var.setStyleSheet("font-size: 20px;")
        self.ending_var.addItems(["הוסף נקודותיים", "הוסף נקודה", "ללא שינוי"])
        self.ending_var.setFixedWidth(170)
        ending_layout.addWidget(ending_label)
        ending_layout.addWidget(self.ending_var, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        layout.addLayout(ending_layout)

        # הדגשת תחילת קטע
        self.emphasize_var = QCheckBox("הדגש את תחילת הקטעים")
        self.emphasize_var.setStyleSheet("font-size: 20px;")
        layout.addWidget(self.emphasize_var)

        # כאן להוסיף לבל של הסבר שזה עובד רק אם יש מעל עשר מילים בשורה
        self.word_count_label = QLabel("התוכנה פועלת רק על שורות שיש בהם עשר מילים ומעלה")
        self.word_count_label.setStyleSheet("font-size: 20px; font-weight: bold;")
        self.word_count_label.setWordWrap(True)
        self.word_count_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.word_count_label)

        # כפתור הפעלה
        run_button = QPushButton("הפעל")
        run_button.clicked.connect(self.run_processing)
        run_button.setFixedHeight(40)
        run_button.setStyleSheet("font-size: 25px;")
        layout.addWidget(run_button)

        self.setLayout(layout)

    def select_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "בחר קובץ טקסט", "", "קבצי טקסט (*.txt);;כל הפורמטים (*.*)", options=options
        )
        if file_path:
            self.file_path_entry.setText(file_path)

    def process_file(self, file_path, add_ending, emphasize_start):
        # בדיקת סוג הקובץ לפי סיומת
        if Path(file_path).suffix.lower() != '.txt':
            QMessageBox.critical(self, "קלט לא תקין", "סוג הקובץ אינו נתמך\nבחר קובץ טקסט [בסיומת TXT.]")
            return
        
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()

            changed = False
            for i in range(len(lines)):
                line = lines[i].rstrip()
                words = line.split()

                # בדיקה אם יש יותר מעשר מילים ושאין סימן כותרת בהתחלה
                if len(words) > 10 and not any(line.startswith(f'<h{n}>') for n in range(2, 10)):
                    # הסרת רווחים ותווים מיותרים בסוף השורה לפני בדיקה
                    stripped_line = line.rstrip(" .,;:!?)</small></big></b>")  # מסיר תווים מיותרים מסוף השורה

                    # מחיקת רווחים לפני נקודה או נקודותיים קיימים בסוף השורה
                    if line.endswith(('.', ':')):
                        line = line.rstrip()  # הסרת רווחים מיותרים לפני הסימן

                    # הוספת נקודה או נקודותיים בסוף השורה
                    if add_ending != "ללא הוספת סימן":
                        if line.endswith(','):
                            line = line.rstrip()  # הסרת רווחים מיותרים לפני הוספת הסימן
                            if add_ending == "הוסף נקודה":
                                lines[i] = line[:-1] + '.\n'
                            elif add_ending == "הוסף נקודותיים":
                                lines[i] = line[:-1] + ':\n'
                            changed = True
                        elif not line.endswith(('.', ':', '!', '?')) and not any(line.endswith(tag) for tag in ['</small>', '</big>', '</b>']):
                            line = line.rstrip()  # הסרת רווחים מיותרים לפני הוספת הסימן
                            if add_ending == "הוסף נקודה":
                                lines[i] = line.rstrip() + '.\n'
                            elif add_ending == "הוסף נקודותיים":
                                lines[i] = line.rstrip() + ':\n'
                            changed = True

                    # הדגשת המילה הראשונה אם אין סימנים מיוחדים
                    if emphasize_start:
                        first_word = words[0]
                        if not any(tag in first_word for tag in ['<b>', '<small>', '<big>', '<h2>', '<h3>', '<h4>', '<h5>', '<h6>']):
                            if not (first_word.startswith('<') and first_word.endswith('>')):
                                lines[i] = '<b>' + first_word + '</b> ' + ' '.join(words[1:]) + '\n'
                                changed = True

            if changed:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.writelines(lines)
                QMessageBox.information(self, "!מזל טוב", "השינויים נשמרו בהצלחה")
            else:
                msg = QMessageBox(self)
                msg.setWindowTitle("!שים לב")
                msg.setText("אין מה להחליף בקובץ זה")
                QTimer.singleShot(2500, msg.close)  # סוגר את ההודעה לאחר 2500 מילי־שניות (2.5 שניות)
                msg.show()
                return
	
        except FileNotFoundError:
            QMessageBox.critical(self, "קלט לא תקין", "הקובץ לא נמצא")
            return
        except UnicodeDecodeError:
            QMessageBox.critical(self, "קלט לא תקין", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בעיבוד הקובץ: {str(e)}")

    def run_processing(self):
        selected_file_path = self.file_path_entry.text()
        if selected_file_path:
            add_ending = self.ending_var.currentText()
            emphasize_start = self.emphasize_var.isChecked()
            self.process_file(selected_file_path, add_ending, emphasize_start)
        else:
            QMessageBox.warning(self, "קלט לא תקין", "אנא בחר קובץ תחילה")

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
   
# ==========================================
# Script 6: יצירת כותרות לעמוד ב
# ==========================================
class CreatePageBHeaders(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("'יצירת כותרות ל'עמוד ב")
        self.setGeometry(100, 100, 800, 350)
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setLayoutDirection(Qt.RightToLeft)
        self.setContentsMargins(40, 40, 40, 40)  # הוספת שוליים
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # הסבר למשתמש
        explanation = QLabel(
            "התוכנה יוצרת כותרת בכל מקום בקובץ שכתוב בתחילת שורה -\n"
            "'עמוד ב', או 'ע\"ב'.\n"
            "באם כתוב את המילה 'שם' לפני המילים הנ\"ל, המילה 'שם' נמחקת\n"
            "ובאם כתוב את המילה 'גמרא' לפני המילים 'עמוד ב' או 'ע\"ב'\n"
            "התוכנה תעביר את המילה 'גמרא' לתחילת השורה שאחרי הכותרת"
        )
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setStyleSheet("font-size: 18px;")
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # נתיב קובץ
        file_layout = QHBoxLayout()
        file_label = QLabel("נתיב קובץ:")
        file_label.setStyleSheet("font-size: 20px;")
        self.file_entry = QLineEdit()
        browse_button = QPushButton("עיון")
        browse_button.setStyleSheet("font-size: 20px;")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_entry)
        file_layout.addWidget(browse_button)
        layout.addLayout(file_layout)

        # רמת כותרת
        header_layout = QHBoxLayout()
        header_label = QLabel("בחר רמת כותרת:")
        header_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        header_label.setStyleSheet("font-size: 20px;")
        self.header_var = QComboBox()
        self.header_var.setFixedWidth(50)
        self.header_var.setStyleSheet("font-size: 20px;")
        self.header_var.addItems([str(i) for i in range(2, 7)])
        self.header_var.setCurrentText("3")
        header_layout.addWidget(header_label)
        header_layout.addWidget(self.header_var, alignment=Qt.AlignLeft | Qt.AlignVCenter)
       
        layout.addLayout(header_layout)

        # כפתור הפעלה
        run_button = QPushButton("הפעל")
        run_button.clicked.connect(self.run_script)
        run_button.setFixedHeight(40)
        run_button.setStyleSheet("font-size: 25px;")
        layout.addWidget(run_button)

        self.setLayout(layout)

    def browse_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "בחר קובץ טקסט", "", "קבצי טקסט (*.txt);;כל הפורמטים (*.*)", options=options
        )
        if file_path:
            self.file_entry.setText(file_path)

    def build_tag_agnostic_pattern(self, word, optional_end_chars="['\"’]*"):
        ANY_TAGS_SPACES = r'(?:<[^>]+>\s*)*'
        pattern = ''
        for char in word:
            pattern += ANY_TAGS_SPACES + re.escape(char)
        pattern += ANY_TAGS_SPACES
        if optional_end_chars:
            pattern += optional_end_chars + ANY_TAGS_SPACES
        return pattern

    def strip_and_replace(self, text, header_level, counter):
        ANY_TAGS_SPACES = r'(?:<[^>]+>\s*)*'
        NON_WORD = r'(?:[^\w<>]|$)'  # תו שאינו אות או סוף השורה
        pattern = r"^\s*" + ANY_TAGS_SPACES

        # אופציונלי: המילה 'שם'
        shem_pattern = self.build_tag_agnostic_pattern('שם', optional_end_chars='')
        pattern += r"(?P<shem>" + shem_pattern + r"\s*)?"

        # אופציונלי: 'גמרא' וגרסאותיה
        gmarah_variants = ['גמרא', 'בגמרא', "גמ'", "בגמ'"]
        gmarah_patterns = [self.build_tag_agnostic_pattern(word, optional_end_chars='') for word in gmarah_variants]
        gmarah_pattern = r"(?P<gmarah>" + '|'.join(gmarah_patterns) + r")\s*"
        pattern += r"(?:" + gmarah_pattern + ")?"

        # 'עמוד ב' או 'ע"ב'
        ab_variants = ['עמוד ב', 'ע"ב', "ע''ב", "ע'ב"]
        ab_patterns = [r"(?<!\w)" + self.build_tag_agnostic_pattern(word) + r"(?!\w)" for word in ab_variants]
        ab_pattern = r"(?P<ab>" + '|'.join(ab_patterns) + r")"

        pattern += ab_pattern
        pattern += NON_WORD  # לוודא שאין תו אות לאחר מכן

        # שאר השורה
        pattern += r"(?P<rest>.*)"

        match_pattern = re.compile(pattern, re.IGNORECASE | re.UNICODE)

        def replace_function(match):
            header = f"<h{header_level}>עמוד ב</h{header_level}>"
            rest_of_line = match.group('rest').lstrip()

            gmarah_text = match.group('gmarah')
            if gmarah_text:
                # הסרת תגים ורווחים
                gmarah_text = re.sub(ANY_TAGS_SPACES, '', gmarah_text).strip()

            counter[0] += 1  # עדכון המונה

            # בניית הטקסט החדש
            if gmarah_text:
                return f"{header}\n{gmarah_text} {rest_of_line}\n" if rest_of_line else f"{header}\n{gmarah_text}\n"
            else:
                return f"{header}\n{rest_of_line}\n" if rest_of_line else f"{header}\n"

        # אם כבר יש תג כותרת, לא משנים
        if re.search(r"<h\d>.*?</h\d>", text, re.IGNORECASE):
            return text

        # ביצוע ההחלפה
        new_text = match_pattern.sub(replace_function, text)

        # הסרת שורות ריקות כפולות
        new_text = re.sub(r'\n\s*\n', '\n', new_text)

        return new_text

    def process_file(self, file_path, header_level):
        # בדיקת סוג הקובץ לפי סיומת
        if Path(file_path).suffix.lower() != '.txt':
            QMessageBox.critical(self, "קלט לא תקין", "סוג הקובץ אינו נתמך\nבחר קובץ טקסט [בסיומת TXT.]")
            return
        
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()

            new_lines = []
            counter = [0]  # מונה כותרות

            for line in lines:
                new_line = self.strip_and_replace(line, header_level, counter)
                new_lines.append(new_line)

            with open(file_path, 'w', encoding='utf-8') as file:
                file.writelines(new_lines)

            # הצגת הודעה מתאימה לפי כמות הכותרות שנוצרו
            if counter[0] == 0:
                msg = QMessageBox(self)
                msg.setWindowTitle("!שים לב")
                msg.setText("אין מה להחליף בקובץ זה")
                QTimer.singleShot(2500, msg.close)  # סוגר את ההודעה לאחר 2500 מילי־שניות (2.5 שניות)
                msg.show()
                return
            else:
                QMessageBox.information(self, "!מזל טוב", f"נוספו {counter[0]} כותרות לקובץ!")
        
        except FileNotFoundError:
            QMessageBox.critical(self, "קלט לא תקין", "הקובץ לא נמצא")
            return
        except UnicodeDecodeError:
            QMessageBox.critical(self, "קלט לא תקין", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return        
        except Exception as e:
            QMessageBox.critical(self, "קלט לא תקין", f"שגיאה בפתיחת קובץ: {e}")
            return

    def run_script(self):
        file_path = self.file_entry.text()
        try:
            header_level = int(self.header_var.currentText())  # המרה למספר שלם
        except ValueError:
            QMessageBox.warning(self, "שגיאה", "רמת הכותרת צריכה להיות מספר בין 2 ל-6")
            return

        if not file_path:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)            
            msg.setWindowTitle("קלט לא תקין")
            msg.setText("נא לבחור קובץ תחילה")
            QTimer.singleShot(1000, msg.close)  # סוגר את ההודעה לאחר 1000 מילי־שניות (1 שניות)
            msg.show()
            return

        if header_level < 2 or header_level > 6:
            QMessageBox.warning(self, "שגיאה", "בחר רמת כותרת בין 2 ל-6")
            return

        self.process_file(file_path, header_level)

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
   
# ==========================================
# Script 7: החלפת כותרות לעמוד ב
# ==========================================
class ReplacePageBHeaders(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("'החלפת כותרות ל'עמוד ב")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setLayoutDirection(Qt.RightToLeft)
        self.setGeometry(100, 100, 1000, 350)
        self.setContentsMargins(40, 40, 40, 40)  # הוספת שוליים
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # הסבר למשתמש
        explanation = QLabel(
            "שים לב!\nהתוכנה פועלת רק אם הדפים והעמודים הוגדרו כבר ככותרות\n[לא משנה באיזה רמת כותרת]\nכגון:  <h3>עמוד ב</h3> או: <h2>עמוד ב</h2> וכן הלאה\n\nזהירות!\nבדוק היטב שלא פספסת שום כותרת של 'דף' לפני שאתה מריץ תוכנה זו\nכי במקרה של פספוס, הכותרת 'עמוד ב' שאחרי הפספוס תהפך לכותרת שגויה\n"
        )
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setStyleSheet("font-size: 18px;")
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # נתיב קובץ
        file_layout = QHBoxLayout()
        file_label = QLabel("נתיב קובץ:")
        file_label.setStyleSheet("font-size: 20px;")
        self.file_entry = QLineEdit()
        browse_button = QPushButton("עיון")
        browse_button.setStyleSheet("font-size: 20px;")
        browse_button.clicked.connect(self.choose_file)
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_entry)
        file_layout.addWidget(browse_button)
        layout.addLayout(file_layout)

        # סוג ההחלפה
        replace_layout = QHBoxLayout()
        replace_label = QLabel("בחר את סוג ההחלפה:")
        replace_label.setStyleSheet("font-size: 20px;")
        replace_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.replace_type = QComboBox()
        self.replace_type.setFixedWidth(130)
        self.replace_type.setStyleSheet("font-size: 20px;")
        self.replace_type.addItems(["נקודותיים", "ע\"ב"])
        replace_layout.addWidget(replace_label)
        replace_layout.addWidget(self.replace_type, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        layout.addLayout(replace_layout)

        # דוגמאות
        example1 = QLabel("לדוגמא:\nדף ב:   דף ג:   דף ד:   דף ה:\nוכן הלאה")
        example1.setAlignment(Qt.AlignCenter)
        example1.setStyleSheet("font-size: 16px;")
        example1.setWordWrap(True)
        layout.addWidget(example1)

        example2 = QLabel("או:\nדף ב ע\"ב   דף ג ע\"ב   דף ד ע\"ב   דף ה ע\"ב\nוכן הלאה")
        example2.setAlignment(Qt.AlignCenter)
        example2.setStyleSheet("font-size: 16px;")
        example2.setWordWrap(True)
        layout.addWidget(example2)

        # כפתור הפעלה
        run_button = QPushButton("בצע החלפה")
        run_button.clicked.connect(self.run_script)
        run_button.setFixedHeight(40)
        run_button.setStyleSheet("font-size: 25px;")
        layout.addWidget(run_button)

        self.setLayout(layout)

    def choose_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "בחר קובץ טקסט", "", "קבצי טקסט (*.txt)", options=options
        )
        if file_path:
            self.file_entry.setText(file_path)
            self.sender().setText("קובץ נבחר, המשך לבחירת סוג ההחלפה")

    def update_file(self, replace_type):
        file_path = self.file_entry.text()
        if not file_path:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)            
            msg.setWindowTitle("קלט לא תקין")
            msg.setText("נא לבחור קובץ תחילה")
            QTimer.singleShot(1000, msg.close)  # סוגר את ההודעה לאחר 1000 מילי־שניות (1 שניות)
            msg.show()
            return
        # בדיקת סוג הקובץ לפי סיומת
        if Path(file_path).suffix.lower() != '.txt':
            QMessageBox.critical(self, "קלט לא תקין", "סוג הקובץ אינו נתמך\nבחר קובץ טקסט [בסיומת TXT.]")
            return      

        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()

            previous_title = ""
            previous_level = ""
            replacements_made = 0  # ספירת כמות ההחלפות

            def replace_match(match):
                nonlocal previous_title, previous_level, replacements_made
                level = match.group(1)
                title = match.group(2)

                # בדיקה אם הכותרת היא "דף"
                if re.match(r"דף \S+\.?", title):
                    previous_title = title.strip()
                    previous_level = level
                    return match.group(0)

                # בדיקה אם הכותרת היא "עמוד ב"
                elif title == "עמוד ב":
                    replacements_made += 1  # הוחלפה כותרת
                    if replace_type == "נקודותיים":
                        return f'<h{previous_level}>{previous_title.rstrip(".")}:</h{previous_level}>'
                    elif replace_type == "ע\"ב":
                        # הסרת "ע\"א" או "עמוד א" מהכותרת הקודמת אם קיימים
                        modified_title = re.sub(r'( ע\"א| עמוד א)$', '', previous_title)
                        return f'<h{previous_level}>{modified_title.rstrip(".")} ע\"ב</h{previous_level}>'

                # אם זה לא אחד המקרים למעלה, נשאיר את הכותרת כפי שהיא
                return match.group(0)

            content = re.sub(r'<h([1-9])>(.*?)</h\1>', replace_match, content)

            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(content)

            if replacements_made == 0:
                msg = QMessageBox(self)
                msg.setWindowTitle("!שים לב")
                msg.setText("אין מה להחליף בקובץ זה")
                QTimer.singleShot(2500, msg.close)  # סוגר את ההודעה לאחר 2500 מילי־שניות (2.5 שניות)
                msg.show()
                return
            else:
                QMessageBox.information(self, "!מזל טוב", f"הקובץ עודכן בהצלחה!\n\nבוצעו {replacements_made} החלפות")
        
        except FileNotFoundError:
            QMessageBox.critical(self, "קלט לא תקין", "הקובץ לא נמצא")
            return
        except UnicodeDecodeError:
            QMessageBox.critical(self, "קלט לא תקין", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return
        except Exception as e:
            QMessageBox.critical(self, "קלט לא תקין", f"שגיאה בפתיחת קובץ: {e}")
            return

    def run_script(self):
        replace_type = self.replace_type.currentText()
        self.update_file(replace_type)

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)
   
# ==========================================
# Script 8: בדיקת שגיאות בכותרות
# ==========================================

def create_labeled_widget(label_text, widget):
    container = QWidget()
    v_layout = QVBoxLayout()
    v_layout.setContentsMargins(0, 0, 0, 0)  # מסיר את כל המרווחים סביב ה-layout
    v_layout.setSpacing(2)  # מגדיר מרווח קטן בין הווידג'טים (ניתן להתאים לערך הרצוי)
    label = QLabel(label_text)
    label.setStyleSheet("font-size: 18px;")
    v_layout.addWidget(label)
    v_layout.addWidget(widget)
    container.setLayout(v_layout)
    return container

def create_labeled_widget_2(label_text, widget, extra_label):
    container = QWidget()
    v_layout = QVBoxLayout()
    v_layout.setContentsMargins(0, 0, 0, 0)  # מסיר את כל המרווחים סביב ה-layout
    v_layout.setSpacing(2)  # מגדיר מרווח קטן בין הווידג'טים (ניתן להתאים לערך הרצוי)
    label = QLabel(label_text)
    label.setStyleSheet("font-size: 18px;")
    v_layout.addWidget(label)
    v_layout.addWidget(widget)
    v_layout.addWidget(extra_label)
    container.setLayout(v_layout)
    return container

# ------------------ מחלקה ראשונה: בדיקת שגיאות בכותרות ------------------ #
class בדיקת_שגיאות_בכותרות(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("בדיקת שגיאות בכותרות")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # תווים בתחילת וסוף הכותרת
        regex_layout = QHBoxLayout()
        re_start_label = QLabel("תו/ים בתחילת הכותרת:")
        self.re_start_entry = QLineEdit()
        self.re_start_entry.setLayoutDirection(Qt.RightToLeft)
        # הגדר RLM כברירת מחדל
        self.re_start_entry.setText('\u200F')
        def maintain_rtl():
            text = self.re_start_entry.text()
            if not text.startswith('\u200F'):
                cursor_pos = self.re_start_entry.cursorPosition()
                self.re_start_entry.setText('\u200F' + text)
                self.re_start_entry.setCursorPosition(cursor_pos + 1)
        self.re_start_entry.textChanged.connect(maintain_rtl)

        re_end_label = QLabel("תו/ים בסוף הכותרת:")
        self.re_end_entry = QLineEdit()
        self.re_end_entry.setLayoutDirection(Qt.RightToLeft)
        # הגדר RLM כברירת מחדל
        self.re_end_entry.setText('\u200F')
        def maintain_rtl():
            text = self.re_end_entry.text()
            if not text.startswith('\u200F'):
                cursor_pos = self.re_end_entry.cursorPosition()
                self.re_end_entry.setText('\u200F' + text)
                self.re_end_entry.setCursorPosition(cursor_pos + 1)
        self.re_end_entry.textChanged.connect(maintain_rtl)
        
        self.gershayim_var = QCheckBox("הגדר גרש/ים כתקינים")

        # יצירת label לרמות חסרות
        self.missing_levels_label = QLabel("")
        self.missing_levels_label.setStyleSheet("font-size: 14px;")
        self.missing_levels_label.hide()  # מוסתר בהתחלה

        # יצירת QTextEdit והגדרותיהם
        self.unmatched_regex_text = QTextEdit()
        self.unmatched_regex_text.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.unmatched_regex_text.setReadOnly(True)
        
        self.unmatched_tags_text = QTextEdit()
        self.unmatched_tags_text.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.unmatched_tags_text.setReadOnly(True)
        self.unmatched_tags_text.setStyleSheet("margin-bottom: 20px;")
        
        # עטיפת כל ווידג'ט במכולה עם תווית מעליו
        regex_container = create_labeled_widget(
            "\nכותרות שיש בהן תווים מיותרים (חוץ ממה שנכתב בתיבות הבחירה למטה)\nכגון: גרש/ים, פסיק, נקודה, נקודותיים. או רווח לפני הכותרת או לאחרי'.",
            self.unmatched_regex_text)
        
        tags_container = create_labeled_widget_2("כותרות שאינן לפי הסדר", self.unmatched_tags_text, self.missing_levels_label)

        # הוספת המכולות ל־QSplitter אנכי
        v_splitter = QSplitter(Qt.Vertical)
        v_splitter.setHandleWidth(1)               # מגדיר רוחב לפס הגרירה כדי שיהיה ברור
        v_splitter.setStyleSheet("""
            QSplitter::handle:vertical {
                background-color: gray;
                height: 0.01px;
                width: 100%;
            }
        """)
        v_splitter.addWidget(tags_container)
        v_splitter.addWidget(self.missing_levels_label)
        v_splitter.addWidget(regex_container)

        # הוספת ה־splitter ל-layout הראשי
        layout.addWidget(v_splitter)

        # הוספת הרכיבים
        regex_layout.addWidget(re_start_label)
        regex_layout.addWidget(self.re_start_entry)
        regex_layout.addWidget(re_end_label)
        regex_layout.addWidget(self.re_end_entry)
        regex_layout.addWidget(self.gershayim_var)
        layout.addLayout(regex_layout)

        self.setLayout(layout)

    def update_missing_levels_label(self, missing_levels):
        if not missing_levels:
            self.missing_levels_label.setText("")
            self.missing_levels_label.hide()
        else:
            levels_str = ", ".join(map(str, missing_levels))
            if len(missing_levels) == 1:
                text = f"מידע: אין בקובץ כותרת ברמה {levels_str}"
            else:
                text = f"מידע: אין בקובץ כותרות ברמות - {levels_str}"
            self.missing_levels_label.setText(text)
            self.missing_levels_label.show()

    def load_file_and_process(self, file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                html_content = file.read()
        except Exception as e:
            return

        re_start = self.re_start_entry.text()
        re_end = self.re_end_entry.text()
        gershayim = self.gershayim_var.isChecked()

        unmatched_regex, unmatched_tags, missing_levels = self.process_html(html_content, re_start, re_end, gershayim)
        self.unmatched_regex_text.setPlainText("\n".join(unmatched_regex))
        self.unmatched_tags_text.setPlainText("\n".join(unmatched_tags))
        
        # עדכון ה-label של הרמות החסרות
        self.update_missing_levels_label(missing_levels)


    def process_html(self, html_content, re_start, re_end, gershayim):
        soup = BeautifulSoup(html_content, 'html.parser')

        # קומפילציה של תבנית Regex לפי קלט המשתמש
        if re_start and re_end:
            pattern = re.compile(f"^[{re.escape(re_start)}]*[א-ת]([א-ת \-]*[א-ת])?[{re.escape(re_end)}]*$")
        elif re_start:
            pattern = re.compile(f"^[{re.escape(re_start)}]*[א-ת]([א-ת \-]*[א-ת])?$")
        elif re_end:
            pattern = re.compile(f"^[א-ת]([א-ת \-]*[א-ת])?[{re.escape(re_end)}]*$")
        else:
            pattern = re.compile(r"^[א-ת]([א-ת \-]*[א-ת])?$")

        unmatched_regex = []
        unmatched_tags = []
        missing_levels = []  # רשימה חדשה לרמות חסרות

        # נעבור על תגי כותרות h2 עד h6
        for i in range(2, 7):
            tags = soup.find_all(f"h{i}")

            # בדיקה אם נמצאו תגים
            if not tags:
                missing_levels.append(i)  # הוספה לרשימת הרמות החסרות
                continue

            # עיבוד כל התגים למעט האחרון
            for index in range(len(tags) - 1):
                current_tag = tags[index].string or ""
                next_tag = tags[index + 1].string or ""

                # וידוא שהמחרוזות של התגים אינן ריקות
                if not current_tag or not next_tag:
                    continue

                # בהנחה שהפיצול מבוצע על רווח כדי לקבל את הכותרות
                current_heading_parts = current_tag.split()
                next_heading_parts = next_tag.split()

                if len(current_heading_parts) > 1:
                    current_heading = current_heading_parts[1]
                else:
                    current_heading = current_tag

                if len(next_heading_parts) > 1:
                    next_heading = next_heading_parts[1]
                else:
                    next_heading = next_tag

                # בדיקה אם התג הנוכחי תואם את התבנית
                # **שינוי חשוב**: אם גרשיים מופעל ויש גרשיים בכותרת, לא נוסיף לרשימת unmatched_regex
                if not re.match(pattern, current_tag):
                    if gershayim and ("'" in current_tag or '"' in current_tag):
                        # לא נוסיף לרשימת unmatched_regex
                        pass
                    else:
                        unmatched_regex.append(current_tag)

                # בדיקה עבור תנאי גרשיים - תמיד כאילו gershayim=False
                if "'" in current_heading or '"' in current_heading:
                    unmatched_tags.append(current_heading)

                # בדיקה אם הכותרות הן ברצף
                if not gematriapy.to_number(current_heading) + 1 == gematriapy.to_number(next_heading):
                    unmatched_tags.append(f"כותרת נוכחית - {current_tag} || כותרת הבאה - {next_tag}")

            # עיבוד התג האחרון
            last_tag = tags[-1].string or ""
            if last_tag and not re.match(pattern, last_tag):
                # **שינוי חשוב**: אם גרשיים מופעל ויש גרשיים בכותרת, לא נוסיף לרשימת unmatched_regex
                if gershayim and ("'" in last_tag or '"' in last_tag):
                    # לא נוסיף לרשימת unmatched_regex
                    pass
                else:
                    unmatched_regex.append(last_tag)

            last_heading_parts = last_tag.split()
            if len(last_heading_parts) > 1:
                last_heading = last_heading_parts[1]
            else:
                last_heading = last_tag

            # הקוד של unmatched_tags - תמיד כאילו gershayim=False
            if "'" in last_heading or '"' in last_heading:
                unmatched_tags.append(last_heading)

        return unmatched_regex, unmatched_tags, missing_levels

# ------------------ מחלקה שנייה: בדיקת שגיאות בעיצוב (תגים וכו') ------------------ #
class בדיקת_שגיאות_בתגים(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("בודק שגיאות בעיצוב")
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        # יצירת תיבות טקסט והגדרותיהם
        self.opening_without_closing = QTextEdit()
        self.opening_without_closing.setLayoutDirection(Qt.RightToLeft)
        self.opening_without_closing.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.opening_without_closing.setReadOnly(True)

        self.closing_without_opening = QTextEdit()
        self.closing_without_opening.setLayoutDirection(Qt.RightToLeft)
        self.closing_without_opening.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.closing_without_opening.setReadOnly(True)

        self.heading_errors = QTextEdit()
        self.heading_errors.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.heading_errors.setReadOnly(True)

        # עטיפת כל ווידג'ט במכולה עם תווית
        opening_container = create_labeled_widget("תגים פותחים ללא תגים סוגרים", self.opening_without_closing)
        closing_container = create_labeled_widget("תגים סוגרים ללא תגים פותחים", self.closing_without_opening)
        heading_container = create_labeled_widget("טקסט שאינו חלק מכותרת, שנמצא באותה שורה עם הכותרת", self.heading_errors)

        # יצירת QSplitter אנכי
        v_splitter_tags = QSplitter(Qt.Vertical)
        v_splitter_tags.setHandleWidth(10)
        v_splitter_tags.addWidget(opening_container)
        v_splitter_tags.addWidget(closing_container)
        v_splitter_tags.addWidget(heading_container)

        # הוספת QSplitter ל-layout הראשי
        main_layout.addWidget(v_splitter_tags)

        self.setLayout(main_layout) 

    def load_file_and_check(self, file_path):
        LRM = '\u202B'
        # ניקוי תוצאות קודמות
        self.opening_without_closing.clear()
        self.closing_without_opening.clear()
        self.heading_errors.clear()

        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                lines = file.readlines()
        except Exception as e:
            return

        open_tags = ["b", "big", "i", "small", "h2", "h3", "h4", "h5", "h6"]
        opening_without_closing_list = []
        closing_without_opening_list = []
        heading_errors_list = []

        for line_number, line in enumerate(lines, start=1):
            # מציאת כל התגים הפותחים והסוגרים עם המיקום שלהם
            all_tags = []
            for match in re.finditer(r'<(/?\w+)>', line):
                tag = match.group(1)
                position = match.start()
                if tag.startswith('/'):
                    all_tags.append(('close', tag[1:], position))
                else:
                    all_tags.append(('open', tag, position))
            
            # מיון לפי מיקום בשורה
            all_tags.sort(key=lambda x: x[2])
            
            # מעקב אחר תגים פתוחים לפי סדר
            open_stack = []
            
            for tag_type, tag_name, position in all_tags:
                if tag_type == 'open':
                    open_stack.append(tag_name)
                else:  # tag_type == 'close'
                    # חיפוש התג בסטאק (מהסוף להתחלה)
                    found = False
                    for i in range(len(open_stack) - 1, -1, -1):
                        if open_stack[i] == tag_name:
                            open_stack.pop(i)
                            found = True
                            break
                    
                    if not found:
                        # תג סוגר ללא פותח מתאים
                        closing_without_opening_list.append(
                            f"שורה {line_number}: {LRM}</{tag_name}>{LRM} || {LRM}{line.strip()}{LRM}"
                        )
            
            # כל התגים שנותרו בסטאק הם פותחים ללא סגירה
            for tag in open_stack:
                opening_without_closing_list.append(
                    f"שורה {line_number}: {LRM}<{tag}>{LRM} || {LRM}{line.strip()}{LRM}"
                )

            # בדיקה לכותרת המכילה טקסט נוסף
            for tag in ["h2", "h3", "h4", "h5", "h6"]:
                heading_pattern = rf'<{tag}>.*?</{tag}>'
                heading_match = re.search(heading_pattern, line)
                if heading_match:
                    start, end = heading_match.span()
                    before = line[:start].strip()
                    after = line[end:].strip()
                    if before or after:
                        heading_errors_list.append(f"שורה {line_number}: {LRM}{line.strip()}{LRM}")

        # הצגת תוצאות
        if opening_without_closing_list:
            self.opening_without_closing.setPlainText("\n".join(opening_without_closing_list))
        else:
            self.opening_without_closing.setPlainText("לא נמצאו שגיאות")

        if closing_without_opening_list:
            self.closing_without_opening.setPlainText("\n".join(closing_without_opening_list))
        else:
            self.closing_without_opening.setPlainText("לא נמצאו שגיאות")
        # הצגת שגיאות בכותרות
        if heading_errors_list:
            self.heading_errors.setPlainText("\n".join(heading_errors_list))
        else:
            self.heading_errors.setPlainText("לא נמצאו שגיאות")

# ------------------ חלון משולב שמאחד את שתי המחלקות ------------------ #
class CheckHeadingErrorsOriginal(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("בודק כותרות + בודק תגים")
        self.setWindowIcon(self.get_app_icon())
        self.setGeometry(100, 100, 1300, 800)

        # שני ה־Widgets שלנו
        self.check_headings_widget = בדיקת_שגיאות_בכותרות()
        self.html_tag_checker_widget = בדיקת_שגיאות_בתגים()
        self.check_headings_widget.setLayoutDirection(Qt.RightToLeft)
        self.html_tag_checker_widget.setLayoutDirection(Qt.RightToLeft)
        self.check_headings_widget.resize(800, 400)
        self.html_tag_checker_widget.resize(1200, 900)
        
        # תיבות למעלה: נתיב קובץ וכפתור Browse
        top_layout = QHBoxLayout()
        self.file_path_label = QLabel("נתיב קובץ:")
        self.file_path_label.setStyleSheet("font-size: 18px;")

        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(False)
        self.file_path_edit.returnPressed.connect(self.run_from_line_edit)

        self.browse_button = QPushButton("בחר קובץ")
        self.browse_button.setStyleSheet("font-size: 18px;")
        self.browse_button.setFixedHeight(40)
        self.browse_button.setFixedWidth(280)
        self.browse_button.clicked.connect(self.browse_file)

        self.check_again = QPushButton("בדוק שוב")
        self.check_again.setStyleSheet("font-size: 18px;")        
        self.check_again.clicked.connect(self.run_from_line_edit)

        top_layout.addWidget(self.check_again)
        top_layout.addWidget(self.file_path_label)
        top_layout.addWidget(self.file_path_edit)
        top_layout.addWidget(self.browse_button)
    
        # הפרדה אופקית (splitter) בין שני הרכיבים
        splitter = QSplitter(Qt.Horizontal)
        splitter.setStyleSheet("QSplitter::handle { background-color: gray; }")
        splitter.setHandleWidth(5)               # מגדיר רוחב לפס הגרירה כדי שיהיה ברור
        splitter.setStyleSheet("""
            QSplitter::handle:horizontal {
                width: 5px;
                margin-left: 1.5px;
                margin-right: 1.5px;
                background: gray;
            }
        """)
        splitter.setChildrenCollapsible(False)   # מונע מקיפול אוטומטי של אחד מהווידג'טים
        self.html_tag_checker_widget.setMinimumWidth(10)  # מגדיר רוחב מינימלי
        self.check_headings_widget.setMinimumWidth(10)      # מגדיר רוחב מינימלי        

        # עדכון בתוך בניית ה־ html_container:
        self.html_container_layout = QVBoxLayout()
        self.html_container_layout.setContentsMargins(0, 0, 0, 0)
        self.html_container_layout.addWidget(self.html_tag_checker_widget)
        # אל נוסיף כאן את ה־ pic_count_label
        html_container = QWidget()
        html_container.setLayout(self.html_container_layout)
        

        html_container.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.check_headings_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        html_container.setMaximumHeight(16777215)
        self.check_headings_widget.setMaximumHeight(16777215)
        self.check_headings_widget.resize(800, 400)
        html_container.resize(800, 400)

        self.pic_count_label = QLabel("")
        self.pic_count_label.setStyleSheet("font-size: 18px; color: blue;")

        splitter.addWidget(self.check_headings_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)

        # בניית ה־layout הכללי
        main_layout = QVBoxLayout()
        top_container = QWidget()
        top_container.setLayout(top_layout)
        top_container.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        main_layout.addWidget(top_container)
        splitter.addWidget(html_container)

        main_layout.addWidget(splitter, 1)

        self.setLayout(main_layout)
        self.resize(1700, 900)  # גודל התחלתי

    def process_file(self, file_path):
        if not file_path:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)            
            msg.setWindowTitle("קלט לא תקין")
            msg.setText("נא לבחור קובץ תחילה")
            QTimer.singleShot(1000, msg.close)  # סוגר את ההודעה לאחר 1000 מילי־שניות (1 שניות)
            msg.show()
            return
        # בדיקת סוג הקובץ לפי סיומת
        if Path(file_path).suffix.lower() != '.txt':
            QMessageBox.critical(self, "קלט לא תקין", "סוג הקובץ אינו נתמך\nבחר קובץ טקסט [בסיומת TXT.]")
            return      

        # עדכון הנתיב בתיבת הטקסט (אם לא נעשה כבר)
        self.file_path_edit.setText(file_path)

        # הפעלת הבדיקות בשני ה־widgets
        self.check_headings_widget.load_file_and_process(file_path)
        self.html_tag_checker_widget.load_file_and_check(file_path)

        # קריאת תוכן הקובץ עם טיפול בשגיאות
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
        except FileNotFoundError:
            QMessageBox.critical(self, "קלט לא תקין", "הקובץ לא נמצא")
            return
        except UnicodeDecodeError:
            QMessageBox.critical(self, "קלט לא תקין", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return
        except Exception as e:
            QMessageBox.critical(self, "קלט לא תקין", f"שגיאה בפתיחת קובץ: {e}")
            return

        # בדיקה עבור המחרוזת "ציור בספר"
        patterns = ["ציור בספר", "$$ ציור $$", "$$ציור$$", "$$ ציור", "ציור $$", "ציור$$", "$$ציור"]
        total_count = sum(content.count(pattern) for pattern in patterns)
        if total_count > 0:
            text = (f'שים לב! יש בספר {total_count} ציורים.\n'
                    'חפש בתוך הספר את המילים "ציור בספר", או את התווים: "$$",\n'
                    'הורד את הספר מהיברובוקס, עשה צילום מסך לתמונה,\n'
                    'והמר אותה לטקסט ע"י תוכנה מספר 10')
            self.pic_count_label.setText(text)
            if self.pic_count_label.parent() is None:
                self.html_container_layout.addWidget(self.pic_count_label)
            self.pic_count_label.setVisible(True)
        else:
            self.pic_count_label.setText("")
            if self.pic_count_label.parent() is not None:
                self.html_container_layout.removeWidget(self.pic_count_label)
                self.pic_count_label.setParent(None)

    def browse_file(self):
        """
        בוחרים קובץ, מעדכנים את תיבת הנתיב ומריצים את כל הבדיקות.
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "בחר קובץ", "", "קבצי טקסט (*.txt);;כל הקבצים (*)")
        if file_path:
            self.process_file(file_path)

    def run_from_line_edit(self):
        file_path = self.file_path_edit.text().strip()
        
        if not file_path:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)            
            msg.setWindowTitle("קלט לא תקין")
            msg.setText("נא לבחור קובץ תחילה")
            QTimer.singleShot(1000, msg.close)  # סוגר את ההודעה לאחר 1000 מילי־שניות (1 שניות)
            msg.show()
            return        
        
        if file_path:
            self.process_file(file_path)

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def get_app_icon(self):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(icon_base64))
        return QIcon(pixmap)
   
# ==========================================
# Script 9: בדיקת שגיאות בכותרות מותאם לספרים על השס
# ==========================================

# ------------------ מחלקה ראשונה: בדיקת שגיאות בכותרות ------------------ #
class בדיקת_שגיאות_בכותרות_לשס(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("בדיקת שגיאות בכותרות")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # תווים בתחילת וסוף הכותרת
        regex_layout = QHBoxLayout()

        re_start_label = QLabel("תו/ים בתחילת הכותרת:")
        self.re_start_entry = QLineEdit()
        self.re_start_entry.setLayoutDirection(Qt.RightToLeft)
        # הגדר RLM כברירת מחדל
        self.re_start_entry.setText('\u200F')
        def maintain_rtl():
            text = self.re_start_entry.text()
            if not text.startswith('\u200F'):
                cursor_pos = self.re_start_entry.cursorPosition()
                self.re_start_entry.setText('\u200F' + text)
                self.re_start_entry.setCursorPosition(cursor_pos + 1)
        self.re_start_entry.textChanged.connect(maintain_rtl)        

        re_end_label = QLabel("תו/ים בסוף הכותרת:")
        self.re_end_entry = QLineEdit()
        self.re_end_entry.setLayoutDirection(Qt.RightToLeft)
        # הגדר RLM כברירת מחדל
        self.re_end_entry.setText('\u200F')
        def maintain_rtl():
            text = self.re_end_entry.text()
            if not text.startswith('\u200F'):
                cursor_pos = self.re_end_entry.cursorPosition()
                self.re_end_entry.setText('\u200F' + text)
                self.re_end_entry.setCursorPosition(cursor_pos + 1)
        self.re_end_entry.textChanged.connect(maintain_rtl)
        self.re_end_entry.setText(".:")

        self.gershayim_var = QCheckBox("הגדר גרש/ים כתקינים")

        # יצירת label לרמות חסרות
        self.missing_levels_label = QLabel("")
        self.missing_levels_label.setStyleSheet("font-size: 14px;")
        self.missing_levels_label.hide()  # מוסתר בהתחלה

        # יצירת QTextEdit והגדרותיהם
        self.unmatched_regex_text = QTextEdit()
        self.unmatched_regex_text.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.unmatched_regex_text.setReadOnly(True)

        self.unmatched_tags_text = QTextEdit()
        self.unmatched_tags_text.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.unmatched_tags_text.setReadOnly(True)
        self.unmatched_tags_text.setStyleSheet("margin-bottom: 20px;")

        # עטיפת כל ווידג'ט במכולה עם תווית מעליו
        regex_container = create_labeled_widget(
            "\nכותרות שיש בהן תווים מיותרים (חוץ ממה שנכתב בתיבות הבחירה למטה)\nכגון: גרש/ים, פסיק, נקודה, נקודותיים. או רווח לפני הכותרת או לאחרי'.",
            self.unmatched_regex_text)
            
        tags_container = create_labeled_widget_2("כותרות שאינן לפי הסדר\nהתוכנה מדלגת בבדיקה בכל פעם על כותרת אחת, בגלל הכותרות הכפולות של עמוד ב", self.unmatched_tags_text, self.missing_levels_label)

        # הוספת המכולות ל־QSplitter אנכי
        v_splitter = QSplitter(Qt.Vertical)
        v_splitter.setHandleWidth(1)  # עובי handle לפי בחירתך
        v_splitter.setStyleSheet("""
            QSplitter::handle:vertical {
                background-color: gray;
                height: 0.01px;
                width: 100%;
            }
        """)
        v_splitter.addWidget(tags_container)
        v_splitter.addWidget(self.missing_levels_label)
        v_splitter.addWidget(regex_container)
        
        # הוספת ה־splitter ל-layout הראשי
        layout.addWidget(v_splitter)

        # הוספת הרכיבים
        regex_layout.addWidget(re_start_label)
        regex_layout.addWidget(self.re_start_entry)
        regex_layout.addWidget(re_end_label)
        regex_layout.addWidget(self.re_end_entry)
        regex_layout.addWidget(self.gershayim_var)
        layout.addLayout(regex_layout)

        self.setLayout(layout)

    def update_missing_levels_label(self, missing_levels):
        if not missing_levels:
            self.missing_levels_label.setText("")
            self.missing_levels_label.hide()
        else:
            levels_str = ", ".join(map(str, missing_levels))
            if len(missing_levels) == 1:
                text = f"מידע: אין בקובץ כותרת ברמה {levels_str}"
            else:
                text = f"מידע: אין בקובץ כותרות ברמות - {levels_str}"
            self.missing_levels_label.setText(text)
            self.missing_levels_label.show()

    def load_file_and_process(self, file_path):
        """
        פונקציה זו תחליף את open_file, כך שנקבל ישירות את הנתיב מבחוץ
        ונעבד את תוכן הקובץ בהתאם.
        """
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                html_content = file.read()
        except Exception as e:
            return

        re_start = self.re_start_entry.text()
        re_end = self.re_end_entry.text()
        gershayim = self.gershayim_var.isChecked()

        unmatched_regex, unmatched_tags, missing_levels = self.process_html(html_content, re_start, re_end, gershayim)
        self.unmatched_regex_text.setPlainText("\n".join(unmatched_regex))
        self.unmatched_tags_text.setPlainText("\n".join(unmatched_tags))

        # עדכון ה-label של הרמות החסרות
        self.update_missing_levels_label(missing_levels)        

    def process_html(self, html_content, re_start, re_end, gershayim):
        soup = BeautifulSoup(html_content, 'html.parser')

        # קומפילציה של תבנית Regex לפי קלט המשתמש
        if re_start and re_end:
            pattern = re.compile(f"^[{re.escape(re_start)}]*[א-ת]([א-ת \-]*[א-ת])?[{re.escape(re_end)}]*$")
        elif re_start:
            pattern = re.compile(f"^[{re.escape(re_start)}]*[א-ת]([א-ת \-]*[א-ת])?$")
        elif re_end:
            pattern = re.compile(f"^[א-ת]([א-ת \-]*[א-ת])?[{re.escape(re_end)}]*$")
        else:
            pattern = re.compile(r"^[א-ת]([א-ת \-]*[א-ת])?$")


        unmatched_regex = []
        unmatched_tags = []
        missing_levels = []  # רשימה חדשה לרמות חסרות

        # נעבור על תגי כותרות h2 עד h6
        for i in range(2, 7):
            tags = soup.find_all(f"h{i}")

            # בדיקה אם נמצאו תגים
            if not tags:
                missing_levels.append(i)  # הוספה לרשימת הרמות החסרות
                continue

            # עיבוד כל התגים למעט האחרון
            for index in range(len(tags) - 1):
                current_tag = tags[index].string or ""
                next_tag = tags[index + 1].string or ""

                # וידוא שהמחרוזות של התגים אינן ריקות
                if not current_tag or not next_tag:
                    continue

            # עיבוד כל התגים למעט האחרון
            for index in range(len(tags) - 2):
                current_tag = tags[index].string or ""
                next_tag = tags[index + 2].string or ""

                # וידוא שהמחרוזות של התגים אינן ריקות
                if not current_tag or not next_tag:
                    continue

                # בהנחה שהפיצול מבוצע על רווח כדי לקבל את הכותרות
                current_heading_parts = current_tag.split()
                next_heading_parts = next_tag.split()

                if len(current_heading_parts) > 1:
                    current_heading = current_heading_parts[1]
                else:
                    current_heading = current_tag

                if len(next_heading_parts) > 1:
                    next_heading = next_heading_parts[1]
                else:
                    next_heading = next_tag

                # בדיקה אם התג הנוכחי תואם את התבנית
                # **שינוי חשוב**: אם גרשיים מופעל ויש גרשיים בכותרת, לא נוסיף לרשימת unmatched_regex
                if not re.match(pattern, current_tag):
                    if gershayim and ("'" in current_tag or '"' in current_tag):
                        # לא נוסיף לרשימת unmatched_regex
                        pass
                    else:
                        unmatched_regex.append(current_tag)

                # בדיקה עבור תנאי גרשיים - תמיד כאילו gershayim=False
                if "'" in current_heading or '"' in current_heading:
                    unmatched_tags.append(current_heading)

                # בדיקה אם הכותרות הן ברצף
                if not gematriapy.to_number(current_heading) + 1 == gematriapy.to_number(next_heading):
                    unmatched_tags.append(f"כותרת נוכחית - {current_tag} || כותרת הבאה - {next_tag}")

            # עיבוד התג האחרון
            if len(tags) >= 2:
                last_tages = (tags[-2].string or "", tags[-1].string or "")
            elif len(tags) == 1:
                last_tages = ("", tags[-1].string or "")
            else:
                last_tages = ("", "")

            for last_tag in last_tages:
                if last_tag and not re.match(pattern, last_tag):
                    # **שינוי חשוב**: אם גרשיים מופעל ויש גרשיים בכותרת, לא נוסיף לרשימת unmatched_regex
                    if gershayim and ("'" in last_tag or '"' in last_tag):
                        # לא נוסיף לרשימת unmatched_regex
                        pass
                    else:
                        unmatched_regex.append(last_tag)

            last_heading_parts = last_tages[-1].split()
            if len(last_heading_parts) > 1:
                last_heading = last_heading_parts[1]
            else:
                last_heading = last_tages[-1]

            # הקוד של unmatched_tags - תמיד כאילו gershayim=False
            if "'" in last_heading or '"' in last_heading:
                unmatched_tags.append(last_heading)

        return unmatched_regex, unmatched_tags, missing_levels

# ------------------ חלון משולב שמאחד את שתי המחלקות ------------------ #
class CheckHeadingErrorsCustom(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("בודק כותרות לשס + בודק תגים")
        self.setWindowIcon(self.get_app_icon())

        # שני ה־Widgets שלנו
        self.check_headings_widget = בדיקת_שגיאות_בכותרות_לשס()
        self.html_tag_checker_widget = בדיקת_שגיאות_בתגים()
        self.check_headings_widget.resize(800, 400)
        self.html_tag_checker_widget.resize(1200, 900)
        
        # תיבות למעלה: נתיב קובץ וכפתור Browse
        top_layout = QHBoxLayout()
        self.file_path_label = QLabel("נתיב קובץ:")
        self.file_path_label.setStyleSheet("font-size: 18px;")

        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(False)
        self.file_path_edit.returnPressed.connect(self.run_from_line_edit)

        self.browse_button = QPushButton("בחר קובץ")
        self.browse_button.setStyleSheet("font-size: 18px;")
        self.browse_button.setFixedHeight(40)
        self.browse_button.setFixedWidth(280)
        self.browse_button.clicked.connect(self.browse_file)

        self.check_again = QPushButton("בדוק שוב")
        self.check_again.setStyleSheet("font-size: 18px;")        
        self.check_again.clicked.connect(self.run_from_line_edit)

        top_layout.addWidget(self.check_again)
        top_layout.addWidget(self.file_path_label)
        top_layout.addWidget(self.file_path_edit)
        top_layout.addWidget(self.browse_button)

        # הפרדה אופקית (splitter) בין שני הרכיבים
        splitter = QSplitter(Qt.Horizontal)
        splitter.setStyleSheet("QSplitter::handle { background-color: gray; }")
        splitter.setHandleWidth(5)               # מגדיר רוחב לפס הגרירה כדי שיהיה ברור
        splitter.setStyleSheet("""
            QSplitter::handle:horizontal {
                width: 5px;
                margin-left: 1.5px;
                margin-right: 1.5px;
                background: gray;
            }
        """)
        
        splitter.setChildrenCollapsible(False)   # מונע מקיפול אוטומטי של אחד מהווידג'טים
        self.html_tag_checker_widget.setMinimumWidth(10)  # מגדיר רוחב מינימלי
        self.check_headings_widget.setMinimumWidth(10)      # מגדיר רוחב מינימלי        

        # עדכון בתוך בניית ה־ html_container:
        self.html_container_layout = QVBoxLayout()
        self.html_container_layout.setContentsMargins(0, 0, 0, 0)
        self.html_container_layout.addWidget(self.html_tag_checker_widget)
        # אל נוסיף כאן את ה־ pic_count_label
        html_container = QWidget()
        html_container.setLayout(self.html_container_layout)

        html_container.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.check_headings_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        html_container.setMaximumHeight(16777215)
        self.check_headings_widget.setMaximumHeight(16777215)
        self.check_headings_widget.resize(800, 400)
        html_container.resize(800, 400)

        self.pic_count_label = QLabel("")
        self.pic_count_label.setStyleSheet("font-size: 18px; color: blue;")

        splitter.addWidget(self.check_headings_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)

        # בניית ה־layout הכללי
        main_layout = QVBoxLayout()
        top_container = QWidget()
        top_container.setLayout(top_layout)
        top_container.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        main_layout.addWidget(top_container)
        splitter.addWidget(html_container)
        
        main_layout.addWidget(splitter, 1)

        self.setLayout(main_layout)
        self.resize(1700, 900)  # גודל התחלתי

    def process_file(self, file_path):
        if not file_path:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)            
            msg.setWindowTitle("קלט לא תקין")
            msg.setText("נא לבחור קובץ תחילה")
            QTimer.singleShot(1000, msg.close)  # סוגר את ההודעה לאחר 1000 מילי־שניות (1 שניות)
            msg.show()
            return
        # בדיקת סוג הקובץ לפי סיומת
        if Path(file_path).suffix.lower() != '.txt':
            QMessageBox.critical(self, "קלט לא תקין", "סוג הקובץ אינו נתמך\nבחר קובץ טקסט [בסיומת TXT.]")
            return      

        # עדכון הנתיב בתיבת הטקסט (אם לא נעשה כבר)
        self.file_path_edit.setText(file_path)

        # הפעלת הבדיקות בשני ה־widgets
        self.check_headings_widget.load_file_and_process(file_path)
        self.html_tag_checker_widget.load_file_and_check(file_path)

        # קריאת תוכן הקובץ עם טיפול בשגיאות
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
        except FileNotFoundError:
            QMessageBox.critical(self, "קלט לא תקין", "הקובץ לא נמצא")
            return
        except UnicodeDecodeError:
            QMessageBox.critical(self, "קלט לא תקין", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return
        except Exception as e:
            QMessageBox.critical(self, "קלט לא תקין", f"שגיאה בפתיחת קובץ: {e}")
            return

        # בדיקה עבור המחרוזת "ציור בספר"
        patterns = ["ציור בספר", "$$ ציור $$", "$$ציור$$", "$$ ציור", "ציור $$", "ציור$$", "$$ציור"]
        total_count = sum(content.count(pattern) for pattern in patterns)
        if total_count > 0:
            text = (f'שים לב! יש בספר {total_count} ציורים.\n'
                    'חפש בתוך הספר את המילים "ציור בספר", או את התווים: "$$",\n'
                    'הורד את הספר מהיברובוקס, עשה צילום מסך לתמונה,\n'
                    'והמר אותה לטקסט ע"י תוכנה מספר 10')
            self.pic_count_label.setText(text)
            if self.pic_count_label.parent() is None:
                self.html_container_layout.addWidget(self.pic_count_label)
            self.pic_count_label.setVisible(True)
        else:
            self.pic_count_label.setText("")
            if self.pic_count_label.parent() is not None:
                self.html_container_layout.removeWidget(self.pic_count_label)
                self.pic_count_label.setParent(None)

    def browse_file(self):
        """
        בוחרים קובץ, מעדכנים את תיבת הנתיב ומריצים את כל הבדיקות.
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "בחר קובץ", "", "קבצי טקסט (*.txt);;כל הקבצים (*)")
        if file_path:
            self.process_file(file_path)

    def run_from_line_edit(self):
        file_path = self.file_path_edit.text().strip()
        
        if not file_path:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)            
            msg.setWindowTitle("קלט לא תקין")
            msg.setText("נא לבחור קובץ תחילה")
            QTimer.singleShot(1000, msg.close)  # סוגר את ההודעה לאחר 1000 מילי־שניות (1 שניות)
            msg.show()
            return        
        
        if file_path:
            self.process_file(file_path)

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def get_app_icon(self):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(icon_base64))
        return QIcon(pixmap)

# ==========================================
# Script 10: המרת תמונה לטקסט
# ==========================================

class ImageToHtmlApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("המרת תמונה לטקסט")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setGeometry(100, 100, 800, 550)
        self.setLayoutDirection(Qt.RightToLeft)
        self.setContentsMargins(40, 40, 40, 40)  # הוספת שוליים

        self.layout = QtWidgets.QVBoxLayout(self)
        
        self.information_label = QLabel("לפניך מספר אפשרויות לבחירת התמונה\nבחר אחת מהן")
        self.information_label.setAlignment(Qt.AlignCenter)
        self.information_label.setStyleSheet("font-size: 20px;")
        self.layout.addWidget(self.information_label)

        # יצירת תווית להנחיה
        self.label = QLabel("גרור ושחרר את הקובץ", self)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("border: 2px dashed gray; font-size: 20px; padding: 40px;")
        self.layout.addWidget(self.label)

        self.instruction_label = QtWidgets.QLabel("הדבק נתיב קובץ [או קישור מקוון לתמונה]\nאו הדבק את התמונה (Ctrl+V):")
        self.instruction_label.setAlignment(Qt.AlignCenter)
        self.instruction_label.setStyleSheet("font-size: 20px;")
        self.layout.addWidget(self.instruction_label)

        self.url_edit = QtWidgets.QLineEdit()
        self.url_edit.textChanged.connect(self.on_text_changed)  # מאזין לשינויים בטקסט
        self.url_edit.returnPressed.connect(self.convert_image)
        self.layout.addWidget(self.url_edit)

        self.add_files_button = QPushButton('בחר קובץ דרך סייר הקבצים', self)
        self.add_files_button.clicked.connect(self.open_file_dialog)
        self.add_files_button.setStyleSheet("font-size: 20px;")
        self.layout.addWidget(self.add_files_button)

        self.convert_btn = QtWidgets.QPushButton("המר")
        self.convert_btn.setEnabled(False)
        self.convert_btn.clicked.connect(self.convert_image)
        self.convert_btn.setStyleSheet("font-size: 25px;")
        self.layout.addWidget(self.convert_btn)

        self.nextInFocusChain = QLabel("ההמרה בוצעה בהצלחה!")
        self.nextInFocusChain.setVisible(False)
        self.nextInFocusChain.setAlignment(Qt.AlignCenter)
        self.nextInFocusChain.setStyleSheet("font-size: 25px;")
        self.layout.addWidget(self.nextInFocusChain)
        
        self.copy_btn = QtWidgets.QPushButton("לחץ כאן להעתקת הטקסט")
        self.copy_btn.setEnabled(False)
        self.copy_btn.clicked.connect(self.copy_to_clipboard)
        self.copy_btn.setStyleSheet("font-size: 25px;")
        self.layout.addWidget(self.copy_btn)

        self.cop = QLabel("הטקסט הועתק ללוח, ניתן להדביקו במסמך")
        self.cop.setVisible(False)
        self.cop.setAlignment(Qt.AlignCenter)
        self.cop.setStyleSheet("font-size: 25px;")
        self.layout.addWidget(self.cop)
        
        # כפתורים שיוצגו לאחר ההמרה
        self.convert_new_button = QPushButton('המרת תמונה נוספת', self)
        self.convert_new_button.setVisible(False)
        self.convert_new_button.clicked.connect(self.reset_for_new_convert)
        self.convert_new_button.setStyleSheet("font-size: 20px;")
        self.layout.addWidget(self.convert_new_button)

        self.setAcceptDrops(True)
        self.img_data = None
        
        # משתנה לאחסון נתיבי קבצי תמונה
        self.image_files = []

    def on_text_changed(self):
        text = self.url_edit.text().strip()
        if text.startswith("file:///"):
            text = text[8:]  # הסרת "file:///"
            self.url_edit.setText(text)  # עדכון השדה לאחר התיקון

        if os.path.exists(text):  # בדיקת קובץ מקומי
            self.label.setText("התמונה נטענה בהצלחה!")
            self.convert_btn.setEnabled(True)
        elif text.lower().startswith("http://") or text.lower().startswith("https://"):
            try:
                req = urllib.request.Request(text, method="HEAD")  # שליחה רק של בקשת HEAD לבדיקה
                urllib.request.urlopen(req)
                self.label.setText("הקישור תקין ונטען בהצלחה!")
                self.convert_btn.setEnabled(True)
            except Exception:
                self.label.setText("לא ניתן לפתוח את הקישור, ודא שהוא תקין")
                self.convert_btn.setEnabled(False)
        else:
            self.label.setText("הנתיב שסיפקת אינו קיים")
            self.convert_btn.setEnabled(False)

    def open_file_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(self, "בחר קבצי תמונה", "", 
                                                "קבצי תמונה (*.png;*.jpg;*.jpeg;*.svg;*.tif;*.heic;*.heif;*.ico;*.webp;*.tiff;*.gif;*.bmp)")
        if files:
            supported_extensions = ('.png', '.jpg', '.jpeg', '.svg', '.tif', '.tiff', '.heic', '.heif', '.ico', '.webp', '.gif', '.bmp')
            for file in files:
                if file.lower().endswith(supported_extensions):
                    self.image_files.append(file)
                    with open(file, 'rb') as f:
                        self.img_data = f.read()
                    self.label.setText("התמונה נטענה בהצלחה!")
                    self.convert_btn.setEnabled(True)
                else:
                    self.label.setText("הסיומת של הקובץ אינה נתמכת.\nבחר קובץ אחר")

    # פונקציה שמופעלת כשגוררים קובץ לחלון
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    # פונקציה שמופעלת כשמשחררים את הקבצים בחלון
    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            supported_extensions = ('.png', '.jpg', '.jpeg', '.svg', '.tif', '.tiff', '.heic', '.heif', '.ico', '.webp', '.gif', '.bmp')
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path and os.path.exists(file_path):
                    if file_path.lower().endswith(supported_extensions):
                        with open(file_path, 'rb') as f:
                            self.img_data = f.read()
                        self.image_files.append(file_path)
                        self.label.setText("התמונה נטענה בהצלחה!")
                        self.convert_btn.setEnabled(True)
                    else:
                        self.label.setText("הסיומת של הקובץ אינה נתמכת.\nבחר קובץ אחר")

    def keyPressEvent(self, event):
        if event.matches(QtGui.QKeySequence.Paste):
            clipboard = QtWidgets.QApplication.clipboard()
            mime_data = clipboard.mimeData()
            # בדיקה אם מדובר בתמונה שהועתקה
            if mime_data.hasImage():
                image = clipboard.image()
                if not image.isNull():
                    buffer = QtCore.QBuffer()
                    buffer.open(QtCore.QBuffer.WriteOnly)
                    image.save(buffer, "PNG")
                    self.img_data = buffer.data().data()
                    self.label.setText("התמונה הודבקה בהצלחה!")
                    self.convert_btn.setEnabled(True)
            else:
                text = clipboard.text().strip().strip('"')
                self.url_edit.setText(text)
            event.accept()

    def convert_image(self):
        path = self.url_edit.text().strip().strip('"')
        if path.startswith("file:///"):  # טיפול בפרוטוקול file:///
            path = path[8:]  # הסרת "file:///"

        if not self.img_data and path:
            if path.lower().startswith("http://") or path.lower().startswith("https://"):
                try:
                    with urllib.request.urlopen(path) as resp:
                        self.img_data = resp.read()
                    self.label.setText("הקישור נטען בהצלחה!")
                except Exception as e:
                    QtWidgets.QMessageBox.warning(self, "שגיאה", f"לא ניתן לפתוח את הקישור:\n{e}")
                    return
            elif os.path.exists(path):  # בדיקה אם הקובץ קיים
                with open(path, 'rb') as f:
                    self.img_data = f.read()
                self.label.setText("התמונה נטענה בהצלחה!")
            else:
                QtWidgets.QMessageBox.warning(self, "שגיאה", "לא ניתן לאתר קובץ בנתיב שסיפקת, ודא שהנתיב נכון")
                return

        if not self.img_data:
            QtWidgets.QMessageBox.warning(self, "שגיאה", "לא נמצאה תמונה להמרה")
            return

        # זיהוי סוג הקובץ על בסיס הסיומת
        file_extension = "png"  # ברירת מחדל
        if self.image_files:
            file_extension = os.path.splitext(self.image_files[0])[-1].lstrip(".").lower()
        elif path:
            file_extension = os.path.splitext(path)[-1].lstrip(".").lower()

        encoded = base64.b64encode(self.img_data).decode('utf-8')
        self.html_code = f'<img src="data:image/{file_extension};base64,{encoded}" >'
        self.copy_btn.setEnabled(True)
        self.nextInFocusChain.setVisible(True)

    def copy_to_clipboard(self):
        clipboard = QtWidgets.QApplication.clipboard()
        clipboard.setText(self.html_code)
        self.cop.setVisible(True)
        self.show_post_convert_buttons()

    # פונקציה להצגת כפתורים אחרי ההמרה
    def show_post_convert_buttons(self):
        self.add_files_button.setEnabled(True)
        self.convert_btn.setEnabled(False)
        self.convert_new_button.setVisible(True)

    # פונקציה לאיפוס עבור המרת קבצים חדשים
    def reset_for_new_convert(self):
        self.img_data = None
        self.image_files = []
        self.label.setText("גרור ושחרר קבצי תמונה")
        self.convert_btn.setEnabled(False)
        self.convert_new_button.setVisible(False)
        self.nextInFocusChain.setVisible(False)
        self.copy_btn.setEnabled(False)
        self.cop.setVisible(False)
        self.url_edit.clear()

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)

# ==========================================
# Script 11: תיקון שגיאות נפוצות
# ==========================================

class TextCleanerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setLayoutDirection(Qt.RightToLeft)
        self.setContentsMargins(40, 40, 40, 40)  # הוספת שוליים
        self.setGeometry(100, 100, 600, 500)        

    def initUI(self):
        layout = QVBoxLayout()
        
        file_path_Layout = QHBoxLayout()
        self.file_path = QLineEdit()
        file_Label = QLabel("נתיב קובץ:")
        file_path_Layout.addWidget(file_Label)
        file_path_Layout.addWidget(self.file_path)

        layout.addLayout(file_path_Layout)
        
        self.loadBtn = QPushButton("עיון")
        self.loadBtn.clicked.connect(self.loadFile)
        layout.addWidget(self.loadBtn)
        
        buttonLayout = QHBoxLayout()
        
        self.selectAllBtn = QPushButton("בחר הכל")
        self.selectAllBtn.clicked.connect(self.selectAll)
        buttonLayout.addWidget(self.selectAllBtn)
        
        self.deselectAllBtn = QPushButton("בטל הכל")
        self.deselectAllBtn.clicked.connect(self.deselectAll)
        buttonLayout.addWidget(self.deselectAllBtn)
        
        layout.addLayout(buttonLayout)
        
        self.checkBoxes = {
            "remove_empty_lines": QCheckBox("מחיקת שורות ריקות"),
            "remove_double_spaces": QCheckBox("מחיקת רווחים כפולים"),
            "remove_spaces_before": QCheckBox("\u202Bמחיקת רווחים לפני - . , : ) ]"),
            "remove_spaces_after": QCheckBox("\u202Bמחיקת רווחים וירידות שורה אחרי - [ ("),
            "remove_spaces_around_newlines": QCheckBox("מחיקת רווחים לפני ואחרי אנטר"),
            "replace_double_quotes": QCheckBox("החלפת 2 גרשים בודדים בגרשיים"),
            "normalize_quotes": QCheckBox("המרת גרשיים לא סטנדרטיים לגרשיים רגילים"),
        }

        for checkbox in self.checkBoxes.values():
            checkbox.setChecked(True)     
            layout.addWidget(checkbox)
        
        self.cleanBtn = QPushButton("הרץ כעת")
        self.cleanBtn.clicked.connect(self.cleanText)
        layout.addWidget(self.cleanBtn)
        
        self.undoBtn = QPushButton("בטל שינוי אחרון")
        self.undoBtn.clicked.connect(self.undoChanges)
        layout.addWidget(self.undoBtn)
        
        self.setLayout(layout)
        self.setWindowTitle("תיקון שגיאות נפוצות")
        self.resize(500, 400)
        self.originalText = ""

    def cleanText(self, file_path):
        file_path = self.file_path.text()
 
        if not file_path:
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)            
            msg.setWindowTitle("קלט לא תקין")
            msg.setText("נא לבחור קובץ תחילה")
            QTimer.singleShot(1000, msg.close)  # סוגר את ההודעה לאחר 1000 מילי־שניות (1 שניות)
            msg.show()
            return        
        
        # בדיקת סוג הקובץ לפי סיומת
        if Path(file_path).suffix.lower() != '.txt':

            QMessageBox.critical(self, "קלט לא תקין", "סוג הקובץ אינו נתמך\nבחר קובץ טקסט [בסיומת TXT.]")
            return
        
        try:
            with open(self.file_path.text(), 'r', encoding='utf-8') as file:
                text = file.read()
            
            self.originalText = text
            
            if self.checkBoxes["remove_empty_lines"].isChecked():
                text = re.sub(r'\n\s*\n', '\n', text)
            if self.checkBoxes["remove_double_spaces"].isChecked():
                text = re.sub(r' +', ' ', text)
            if self.checkBoxes["remove_spaces_before"].isChecked():
                text = re.sub(r'[ \t]+([\)\],.:])', r'\1', text)
            if self.checkBoxes["remove_spaces_after"].isChecked():
                text = re.sub(r'(\s|^)([\[\(])(\s+)', r'\1\2', text, flags=re.MULTILINE)
            if self.checkBoxes["remove_spaces_around_newlines"].isChecked():
                text = re.sub(r'\s*\n\s*', '\n', text)
            if self.checkBoxes["replace_double_quotes"].isChecked():
                text = text.replace("''", '"').replace("``", '"').replace("’’", '"')
            if self.checkBoxes["normalize_quotes"].isChecked():
                text = text.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'").replace("„", '"')
            
            text = text.rstrip()  # מחיקת שורה אחרונה אם היא ריקה

            if text == self.originalText:
                msg = QMessageBox(self)
                msg.setWindowTitle("!שים לב")
                msg.setText("אין מה להחליף בקובץ זה")
                QTimer.singleShot(2500, msg.close)  # סוגר את ההודעה לאחר 2500 מילי־שניות (2.5 שניות)
                msg.show()
                return
            else:
                with open(self.file_path.text(), 'w', encoding='utf-8') as file:
                    file.write(text)
                QMessageBox.information(self, "!מזל טוב", "השינויים בוצעו בהצלחה")

        except FileNotFoundError:
            QMessageBox.critical(self, "קלט לא תקין", "הקובץ לא נמצא")
            return
        except UnicodeDecodeError:
            QMessageBox.critical(self, "קלט לא תקין", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בעיבוד הקובץ: {str(e)}")

    def loadFile(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "בחר קובץ טקסט", "", "קבצי טקסט (*.txt);", options=options)
        if fileName:
            self.file_path.setText(fileName)
    
    def selectAll(self):
        for checkbox in self.checkBoxes.values():
            checkbox.setChecked(True)
    
    def deselectAll(self):
        for checkbox in self.checkBoxes.values():
            checkbox.setChecked(False)
    
    def undoChanges(self):
        if self.file_path.text() and self.originalText:
            with open(self.file_path.text(), 'w', encoding='utf-8') as file:
                file.write(self.originalText)

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)

# ==========================================
# Script 12: סנכרון ספרי דיקטה
# ==========================================

# הגדרות גלובליות
BASE_URL = "https://raw.githubusercontent.com/zevisvei/otzaria-library/refs/heads/main/"
BOOKS_FOLDER = ""

# פונקציה לחילוץ התיקייה האחרונה מהלוג
def extract_last_folder_from_log(log_file):
    """חילוץ שם התיקייה האחרונה מקובץ הלוג"""
    if not os.path.exists(log_file):
        return ""
    
    last_folder = ""
    try:
        with open(log_file, "r", encoding="utf-8") as f:
            for line in f:
                # חיפוש הרשומות של בחירת תיקייה
                match = re.search(r"נבחרה תיקייה: (.+)$", line)
                if match:
                    last_folder = match.group(1).strip()
    except Exception:
        # במקרה של שגיאה בפתיחה או קריאת הקובץ, מחזיר מחרוזת ריקה
        return ""
    
    # בדיקה אם התיקייה עדיין קיימת במערכת
    if last_folder and os.path.exists(last_folder):
        return last_folder
    return ""

class SyncWorker(QThread):
    """מחלקה לביצוע פעולות סנכרון ברקע"""
    update_progress = pyqtSignal(str)
    update_progress_bar = pyqtSignal(int, int)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, books_folder):
        super().__init__()
        self.books_folder = books_folder

    def run(self):
        try:
            # בדיקה אם התיקייה קיימת
            if not os.path.exists(self.books_folder):
                self.finished_signal.emit(False, "התיקייה שנבחרה אינה קיימת")
                return

            # קבלת רשימת קבצים מקומית
            list_local_files = []
            for root, _, files in os.walk(self.books_folder):
                for file in files:
                    if not file.endswith(".txt"):
                        continue
                    rel_path = os.path.relpath(os.path.join(root, file), self.books_folder)
                    list_local_files.append(rel_path)
            
            self.update_progress.emit(f"נמצאו {len(list_local_files)} קבצים מקומיים")
            
            # קבלת רשימת קבצים מ-GitHub
            try:
                response = requests.get(BASE_URL + "DictaToOtzaria/ספרים/לא ערוך/list.txt")
                if response.status_code != 200:
                    self.finished_signal.emit(False, f"שגיאה בקבלת רשימת קבצים מגיטהאב: {response.status_code}")
                    return
                list_from_github = response.text.splitlines()
            except Exception as e:
                self.finished_signal.emit(False, f"שגיאה בגישה לגיטהאב: {str(e)}")
                return
                
            self.update_progress.emit(f"נמצאו {len(list_from_github)} קבצים בתיקייה שבגיטהאב")
            
            # התאמת שמות קבצים למערכת ההפעלה
            list_all_per_os = [file.replace("/", os.sep) for file in list_from_github]
            
            # הורדת קבצים חדשים
            files_to_download = [f for f in list_from_github if f.replace("/", os.sep) not in list_local_files]
            self.update_progress.emit(f"מספר קבצים להורדה: {len(files_to_download)}")
            
            for i, file in enumerate(files_to_download):
                self.update_progress_bar.emit(i+1, len(files_to_download))
                file_name_per_os = file.replace("/", os.sep)
                file_path = os.path.join(self.books_folder, file_name_per_os)
                self.update_progress.emit(f"מוריד: {file}")
                
                # יצירת תיקיות נדרשות
                os.makedirs(os.path.dirname(file_path), exist_ok=True)
                
                # הורדת הקובץ
                try:
                    response = requests.get(BASE_URL + f"DictaToOtzaria/ספרים/לא ערוך/אוצריא/{file}")
                    if response.status_code != 200:
                        self.update_progress.emit(f"שגיאה בהורדת הקובץ: {file}")
                        continue
                    
                    with open(file_path, "w", encoding="utf-8") as f:
                        f.write(response.text)
                except Exception as e:
                    self.update_progress.emit(f"שגיאה בכתיבת הקובץ {file}: {str(e)}")
            
            # מחיקת קבצים שאינם בשרת
            files_to_delete = [f for f in list_local_files if f not in list_all_per_os]
            self.update_progress.emit(f"מספר קבצים למחיקה: {len(files_to_delete)}")
            
            for i, file in enumerate(files_to_delete):
                self.update_progress_bar.emit(i+1, len(files_to_delete))
                file_path = os.path.join(self.books_folder, file)
                self.update_progress.emit(f"מוחק: {file}")
                
                try:
                    os.remove(file_path)
                except Exception as e:
                    self.update_progress.emit(f"שגיאה במחיקת הקובץ {file}: {str(e)}")
            
            # מחיקת תיקיות ריקות
            self.update_progress.emit("מנקה תיקיות ריקות...")
            deleted_folders_count = 0
            for root, dirs, _ in os.walk(self.books_folder, topdown=False):
                for folder in dirs:
                    folder_path = os.path.join(root, folder)
                    try:                    
                        if not os.listdir(folder_path):
                            self.update_progress.emit(f"מוחק תיקייה ריקה: {folder_path}")
                            os.rmdir(folder_path)
                            deleted_folders_count += 1
                    except OSError as e:
                         self.update_progress.emit(f"שגיאה במחיקת תיקייה {folder_path}: {str(e)}")
            if deleted_folders_count > 0:
                self.update_progress.emit(f"נמחקו {deleted_folders_count} תיקיות ריקות")

            self.finished_signal.emit(True, "הסנכרון הושלם בהצלחה")
        
        except Exception as e:
            self.update_progress.emit(f"שגיאה קריטית בתהליך הסנכרון: {str(e)}")
            self.update_progress.emit(traceback.format_exc()) # הדפסת Traceback ללוג לדיבוג
            self.finished_signal.emit(False, f"שגיאה בתהליך הסנכרון: {str(e)}")
            self.update_progress_bar.emit(0, 1)  # עדכון פס ההתקדמות למצב סופי
            self.update_progress.emit("הסנכרון נכשל. יש לבדוק את הלוג לפרטים נוספים")

class CompareWorker(QThread):
    """מחלקה לביצוע פעולות השוואה ברקע"""
    update_progress = pyqtSignal(str)
    finished_signal = pyqtSignal(list, list)

    def __init__(self, books_folder):
        super().__init__()
        self.books_folder = books_folder

    def run(self):
        try:
            # קבלת רשימת קבצים מקומית
            list_local_files = []
            if os.path.exists(self.books_folder):
                for root, _, files in os.walk(self.books_folder):
                    for file in files:
                        if not file.endswith(".txt"):
                            continue
                        rel_path = os.path.relpath(os.path.join(root, file), self.books_folder)
                        list_local_files.append(rel_path)
            
            self.update_progress.emit(f"נמצאו {len(list_local_files)} קבצים מקומיים")
            
            # קבלת רשימת קבצים מ-GitHub
            try:
                response = requests.get(BASE_URL + "DictaToOtzaria/ספרים/לא ערוך/list.txt")
                if response.status_code != 200:
                    self.update_progress.emit(f"שגיאה בקבלת רשימת קבצים מגיטהאב: {response.status_code}")
                    return
                list_from_github = response.text.splitlines()
            except Exception as e:
                self.update_progress.emit(f"שגיאה בגישה לגיטהאב: {str(e)}")
                return
                
            self.update_progress.emit(f"נמצאו {len(list_from_github)} קבצים בתיקייה שבגיטהאב")
            
            # התאמת שמות קבצים למערכת ההפעלה
            list_all_per_os = [file.replace("/", os.sep) for file in list_from_github]
            
            # קבצים חדשים להורדה
            files_to_download = [f.replace("/", os.sep) for f in list_from_github if f.replace("/", os.sep) not in list_local_files]
            
            # קבצים מקומיים שאינם בשרת
            files_to_delete = [f for f in list_local_files if f not in list_all_per_os]
            
            self.finished_signal.emit(files_to_download, files_to_delete)
            
        except Exception as e:
            self.update_progress.emit(f"שגיאה בתהליך ההשוואה: {str(e)}")
            self.finished_signal.emit([], [])


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # קבלת תיקיית המשתמש בצורה תואמת למערכת ההפעלה
        user_home = str(Path.home())
        self.log_file = os.path.join(user_home, "sync_of_dicta_log.txt")
        
        # ניסיון לטעון את התיקייה האחרונה מהלוג
        self.books_folder = extract_last_folder_from_log(self.log_file)
        self.setContentsMargins(40, 40, 40, 40)  # הוספת שוליים
        self.setLayoutDirection(Qt.RightToLeft)

        self.init_ui()
        self.init_logger()  # אתחול מערכת הלוג
        self.apply_styles()  # החלת העיצוב

 
    def init_ui(self):
        # הגדרת החלון הראשי
        self.setWindowTitle("סנכרון ספרי דיקטה שבאוצריא")
        self.setMinimumSize(800, 600)
        self.setGeometry(100, 100, 800, 600)
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        
        # יצירת ווידג'ט מרכזי
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # יצירת הפריסה הראשית
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)  # הוספת שוליים
        main_layout.setSpacing(15)  # הוספת ריווח בין אלמנטים

        # תווית לתיקייה נבחרת
        folder_layout = QHBoxLayout()
        folder_layout.setSpacing(10)

        # עדכון תווית התיקייה בהתאם לקובץ שנטען
        if self.books_folder:
            self.folder_label = QLabel(f"תיקייה נבחרת: {self.books_folder}")
        else:
            self.folder_label = QLabel("תיקייה נבחרת: לא נבחרה")
            
        folder_layout.addWidget(self.folder_label)
        
        # כפתור בחירת תיקייה
        self.select_folder_btn = QPushButton("בחר תיקייה")
        self.select_folder_btn.clicked.connect(self.select_folder)
        folder_layout.addWidget(self.select_folder_btn)
        
        main_layout.addLayout(folder_layout)
        
        # כפתורי פעולה
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(15)  # מרווח בין הכפתורים
        
        self.compare_btn = QPushButton("השווה")
        self.compare_btn.clicked.connect(self.compare_files)
        buttons_layout.addWidget(self.compare_btn)
        
        self.sync_btn = QPushButton("סנכרן")
        self.sync_btn.clicked.connect(self.sync_files)
        buttons_layout.addWidget(self.sync_btn)
        
        main_layout.addLayout(buttons_layout)
        main_layout.addSpacing(10)  # מרווח לפני אזור הפלט

        # כותרת לאזור הפלט
        output_label = QLabel("לוג פעולות:")
        main_layout.addWidget(output_label)
        
        # אזור פלט
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        main_layout.addWidget(self.log_output)
        
        # פס התקדמות
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumHeight(20)  # גובה מינימלי
        main_layout.addWidget(self.progress_bar)
        
        # מצב ראשוני של הכפתורים
        self.update_button_states()

    def apply_styles(self):
        """החלת עיצוב מודרני על הממשק"""
        # הגדרת גופן ברירת מחדל
        app = QApplication.instance()
        font = QFont("Segoe UI", 10)
        app.setFont(font)
        
        # עיצוב כללי
        style_sheet = """
        QMainWindow {
            background-color: #f5f5f5;
        }
        
        QWidget {
            font-family: Segoe UI, Arial, Helvetica;
        }
        
        QPushButton {
            background-color: #3498db;
            color: white;
            border: none;
            border-radius: 5px;
            padding: 8px 16px;
            font-weight: bold;
            min-width: 100px;
        }
        
        QPushButton:hover {
            background-color: #2980b9;
        }
        
        QPushButton:disabled {
            background-color: #bdc3c7;
        }
        
        QLabel {
            color: #2c3e50;
            font-weight: bold;
        }
        
        QTextEdit {
            background-color: white;
            border: 1px solid #dcdde1;
            border-radius: 5px;
            padding: 5px;
        }
        
        QProgressBar {
            border-radius: 5px;
            background-color: #ecf0f1;
            text-align: center;
            color: #2c3e50;
        }
        
        QProgressBar::chunk {
            background-color: #2ecc71;
            border-radius: 5px;
        }
        
        QFileDialog {
            background-color: #f5f5f5;
        }
        
        QMessageBox {
            background-color: #f5f5f5;
        }
        
        QMessageBox QPushButton {
            min-width: 80px;
            min-height: 30px;
        }
        """
        
        self.setStyleSheet(style_sheet)
        
        # עיצוב ספציפי לכפתורים
        self.select_folder_btn.setIcon(QIcon.fromTheme("folder-open"))
        self.compare_btn.setIcon(QIcon.fromTheme("view-refresh"))
        self.sync_btn.setIcon(QIcon.fromTheme("emblem-synchronizing"))
        
        # הוספת צללית או אפקט תלת-ממדי לכפתורים
        for btn in [self.select_folder_btn, self.compare_btn, self.sync_btn]:
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #3498db;
                    color: white;
                    border: : 1px solid #888;
                    border-radius: 8px;
                    padding: 10px 20px;
                    font-weight: bold;
                    min-width: 120px;
                }
                QPushButton:hover {
                    background-color: #2980b9;
                }
                QPushButton:pressed {
                    background-color: #1c6ea4;
                    padding-top: 11px;
                    padding-bottom: 9px;
                }
                QPushButton:disabled {
                    background-color: #bdc3c7;
                }
            """)

    def init_logger(self):
        """אתחול מערכת רישום הלוג"""
        # קבלת תיקיית המשתמש בצורה תואמת למערכת ההפעלה
        user_home = str(Path.home())
        self.log_file = os.path.join(user_home, "sync_of_dicta_log.txt")
        self.log(f"לוג יישמר בקובץ: {self.log_file}")

        # הודעה על טעינת תיקייה קודמת
        if self.books_folder:
            self.log(f"נטענה אוטומטית תיקייה אחרונה שסונכרנה: {self.books_folder}")
        else:
            return
    
    def log(self, message):
        """הוספת הודעה לאזור הפלט"""
        self.log_output.append(message)
                
        # כתיבה לקובץ לוג
        timestamp = QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss")
        log_message = f"{timestamp}: {message}"
   
        try:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(log_message + "\n")
        except Exception as e:
            self.log_output.append(f"שגיאה בכתיבה לקובץ לוג: {str(e)}")
    
    def select_folder(self):
        """בחירת תיקיית ספרים מקומית"""
        folder = QFileDialog.getExistingDirectory(self, "בחר תיקיית ספרים")
        if folder:
            self.books_folder = folder
            self.folder_label.setText(f"תיקייה נבחרת: {folder}")
            self.log(f"נבחרה תיקייה: {folder}")
            self.update_button_states()
    
    def update_button_states(self):
        """עדכון מצב הכפתורים בהתאם לבחירת תיקייה"""
        has_folder = bool(self.books_folder)
        self.compare_btn.setEnabled(has_folder)
        self.sync_btn.setEnabled(has_folder)
    
    def compare_files(self):
        """השוואת קבצים בין התיקייה המקומית לשרת"""
        if not self.books_folder:
            QMessageBox.warning(self, "שגיאה", "יש לבחור תיקייה תחילה")
            return
        
        self.log("מתחיל השוואת קבצים...")
        self.compare_btn.setEnabled(False)
        self.sync_btn.setEnabled(False)
        
        # הרצת תהליך ההשוואה בתהליך נפרד
        self.compare_worker = CompareWorker(self.books_folder)
        self.compare_worker.update_progress.connect(self.log)
        self.compare_worker.finished_signal.connect(self.compare_finished)
        self.compare_worker.start()
    
    def compare_finished(self, files_to_download, files_to_delete):
        """טיפול בתוצאות ההשוואה"""
        self.log("\n=== תוצאות ההשוואה ===")
        
        # הצגת קבצים להורדה
        self.log(f"\nקבצים להורדה ({len(files_to_download)}):")
        for file in files_to_download[:10]:  # הצגת 10 הראשונים
            self.log(f"- {file}")
        if len(files_to_download) > 10:
            self.log(f"...ועוד {len(files_to_download) - 10} קבצים נוספים")
        
        # הצגת קבצים למחיקה
        self.log(f"\nקבצים למחיקה ({len(files_to_delete)}):")
        for file in files_to_delete[:10]:  # הצגת 10 הראשונים
            self.log(f"- {file}")
        if len(files_to_delete) > 10:
            self.log(f"...ועוד {len(files_to_delete) - 10} קבצים נוספים")
        
        self.log("\nההשוואה הושלמה")
        self.compare_btn.setEnabled(True)
        self.sync_btn.setEnabled(True)
    
    def sync_files(self):
        """סנכרון קבצים בין התיקייה המקומית לשרת"""
        if not self.books_folder:
            QMessageBox.warning(self, "שגיאה", "יש לבחור תיקייה תחילה")
            return
        
        msg_box = QMessageBox(self)
        msg_box.setLayoutDirection(Qt.RightToLeft)
        msg_box.setContentsMargins(20, 20, 20, 20)  # הוספת שוליים
        msg_box.setWindowTitle("אישור סנכרון")
        msg_box.setText("האם אתה בטוח שברצונך לסנכרן את הספרים? פעולה זו תוריד קבצים חדשים שנוספו לתיקיי', וכן תמחק קבצים שנמחקו ממנה.\nשים לב! פעולה זו אינה הפיכה")
        msg_box.setIcon(QMessageBox.Question)

        # הגדרת הכפתורים בעברית
        yes_button = msg_box.addButton("כן", QMessageBox.YesRole)
        no_button = msg_box.addButton("לא", QMessageBox.NoRole)
        msg_box.setDefaultButton(yes_button)

        yes_button.setStyleSheet("""
            QPushButton {
                background-color: #5dade2; /* צבע רקע מעט שונה */
                color: white;
                border: 2px solid #3498db; /* גבול כחול בולט */
                border-radius: 5px;
                padding: 5px 10px;
                font-weight: bold; /* טקסט מודגש */
                min-width: 80px;
                min-height: 30px;
            }
            QPushButton:hover {
                background-color: #3498db; /* מעט כהה יותר במעבר עכבר */
            }
            QPushButton:pressed {
                background-color: #2980b9; /* עוד יותר כהה בלחיצה */
            }
        """)
        # ודא שהכפתור השני נשאר עם העיצוב הסטנדרטי של QMessageBox
        no_button.setStyleSheet("""
            QPushButton {
                /* אפשר להשאיר ריק כדי שיקבל את ה-CSS מ-QMessageBox QPushButton */
                /* או להגדיר במפורש את הסגנון הרגיל אם צריך */
                background-color: #ebebeb; /* דוגמה לצבע רגיל */
                color: black;
                border: 1px solid #cccccc;
                border-radius: 5px;
                padding: 5px 10px;
                font-weight: normal;
                min-width: 80px;
                min-height: 30px;                 
            }
            QPushButton:hover { background-color: #f5f5f5; }
            QPushButton:pressed { background-color: #dddddd; }
        """)

        msg_box.exec_()

        # בדיקה איזה כפתור נלחץ
        if msg_box.clickedButton() == yes_button:
            self.log("מתחיל סנכרון קבצים...")
            self.compare_btn.setEnabled(False)
            self.sync_btn.setEnabled(False)
            self.progress_bar.setVisible(True)
            
            # הרצת תהליך הסנכרון בתהליך נפרד
            self.sync_worker = SyncWorker(self.books_folder)
            self.sync_worker.update_progress.connect(self.log)
            self.sync_worker.update_progress_bar.connect(self.update_progress)
            self.sync_worker.finished_signal.connect(self.sync_finished)
            self.sync_worker.start()
        else:
            # אם נלחץ על "לא", לא עושים כלום
            self.log("בוטל סנכרון קבצים")
            return

    def update_progress(self, current, total):
        """עדכון פס ההתקדמות"""
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
    
    def sync_finished(self, success, message):
        """טיפול בסיום הסנכרון"""
        self.log(message)
        
        if success:
            msg = QMessageBox(self)  
            msg.setIcon(QMessageBox.Information)        
            msg.setWindowTitle("סנכרון הושלם")
            msg.setText("הסנכרון הושלם בהצלחה")
            QTimer.singleShot(3000, msg.close)  # סוגר את ההודעה לאחר 3000 מילי־שניות (3 שניות)
            msg.show()
        else:
            QMessageBox.warning(self, "שגיאה בסנכרון", message)
        
        self.progress_bar.setVisible(False)
        self.compare_btn.setEnabled(True)
        self.sync_btn.setEnabled(True)

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)


# ==========================================
# Main Menu: תפריט ראשי לבחירת הסקריפטים
# ==========================================
class MainMenu(QWidget):
    def __init__(self):
        super().__init__()
        self.current_version = VERSION
        self._frozen = getattr(sys, 'frozen', False)
        if self._frozen:
            print(f"Running as compiled executable, version: {self.current_version}")
        else:
            print(f"Running in Python environment, version: {self.current_version}")
            
        # הגדרת החלון
        self.setWindowTitle("עריכת ספרי דיקטה עבור אוצריא")
        self.setLayoutDirection(Qt.RightToLeft)
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setContentsMargins(50, 50, 50, 50)  # הוספת שוליים
        self.apply_styles()  # החלת העיצוב

        self.init_ui()

        # הגדרת האייקון לשורת המשימות
        if sys.platform == 'win32':
            QtWin.setCurrentProcessExplicitAppUserModelID(myappid)
        
        # בדיקת עדכונים אוטומטית בהפעלה (בשקט)
        QTimer.singleShot(1000, self.check_for_updates)

    def apply_styles(self):
        """החלת עיצוב מודרני על הממשק"""
        # הגדרת גופן ברירת מחדל
        app = QApplication.instance()
        font = QFont("Segoe UI", 10)
        app.setFont(font)
                
        # עיצוב כללי
        style_sheet = """
        QWidget {
            background-color: #F8F9FA;
        }
        """
        self.setStyleSheet(style_sheet)

    def check_for_updates(self, silent=True):
        """
        בדיקת עדכונים חדשים
        """
        # עדכון הודעת הסטטוס
        self.status_label.setText("בודק עדכונים...")

        self.update_checker = UpdateChecker(self.current_version)
        
        # חיבור הסיגנלים
        self.update_checker.update_available.connect(self.handle_update_available)
        self.update_checker.no_update.connect(self.handle_no_update)
        self.update_checker.error.connect(lambda msg: self.handle_update_error(msg, silent))

        self.update_checker.start()

    def handle_no_update(self, silent=True):
        """טיפול במקרה שאין עדכון"""
        self.status_label.setText("התוכנה מעודכנת")        
        
        if not silent:
            QMessageBox.information(
                self,
                "אין עדכון",
                "אתה משתמש בגרסה העדכנית ביותר"
            )

    # פונקציה להודעות סטטוס זמניות
    def set_temporary_status(self, message, duration_seconds=120):
        """הצגת הודעת סטטוס זמנית שתיעלם אחרי זמן מוגדר"""
        self.status_label.setText(message)
        
        # יצירת טיימר שיוריד את ההודעה
        if hasattr(self, 'status_timer'):
            self.status_timer.stop()
        
        self.status_timer = QTimer()
        self.status_timer.setSingleShot(True)
        self.status_timer.timeout.connect(lambda: self.status_label.setText(""))
        self.status_timer.start(duration_seconds * 1000)  # המרה לאלפיות השנייה

    def handle_update_error(self, error_msg, silent=True):
        """טיפול בשגיאות בתהליך העדכון"""
        self.set_temporary_status("שגיאה בבדיקת העדכונים", 120)  # 2
        
        if not silent:
            QMessageBox.warning(
                self,
                "שגיאה",
                error_msg
            )
        
    def init_ui(self):
        layout = QVBoxLayout()
        self.setGeometry(100, 100, 600, 900)
        label = QLabel("בחר את התוכנה שברצונך להפעיל")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 27px;")
        layout.addWidget(label)

        grid_layout = QGridLayout()

        # רשימת כפתורים עם שמות הפונקציות
        button_info = [
            ("1\n\nיצירת כותרות\nלאוצריא\n(התוכנה הראשית)", self.open_create_headers_otzria),
            ("2\n\nיצירת כותרות\nלאותיות בודדות\n", self.open_create_single_letter_headers),
            ("3\n\nהוספת\nמספר העמוד\nבכותרת הדף", self.open_add_page_number_to_heading),
            ("4\n\nשינוי רמת כותרת\n\n", self.open_change_heading_level),
            ("5\n\nהדגשת\nמילה ראשונה\nוניקוד בסוף קטע", self.open_emphasize_and_punctuate),
            ("6\n\nיצירת כותרות\nלעמוד ב\n", self.open_create_page_b_headers),
            ("7\n\nהחלפת כותרות\nלעמוד ב\n", self.open_replace_page_b_headers),
            ("9\n\nבדיקת שגיאות\n\n", self.open_check_heading_errors_original),
            ("10\n\nבדיקת שגיאות\nלספרים על השס\n", self.open_check_heading_errors_custom),
            ("8\n\nהמרת תמונה\nלטקסט\n", self.open_Image_To_Html_App),
            ("11\n\nתיקון\nשגיאות נפוצות\n", self.open_Text_Cleaner_App),
            ("12\n\nסנכרון\nספרי דיקטה\n", self.open_Main_Window),
        ]
        
        buttons = []
        for i, (text, func) in enumerate(button_info):
            button = QPushButton(text)
            button.setFixedSize(190, 180)  # הגדרת רוחב וגובה שווים (ריבוע)
            button.setStyleSheet('font-size: 20px;')
            button.clicked.connect(func)  # קישור כל כפתור לפונקציה המתאימה

            # הוספת עיצוב של שוליים מעוגלים לכל כפתור
            button.setStyleSheet("""
                QPushButton {
                    border: 1px solid #888;
                    border-radius: 30px;
                    padding: 10px;
                    margin: 5;
                    background-color: #eaeaea;
                    color: black;
                    font-family: "Segoe UI", Arial;
                    font-size: 10pt;
                }
                QPushButton:hover {
                    background-color: #b7b5b5;
                }
            """)
            buttons.append(button)

        # מיקום הלחצנים בתוך ה- Grid
        grid_layout.addWidget(buttons[0], 0, 0)  # שורה 1, טור 1
        grid_layout.addWidget(buttons[1], 0, 1)  # שורה 1, טור 2
        grid_layout.addWidget(buttons[2], 0, 2)  # שורה 1, טור 3
        grid_layout.addWidget(buttons[3], 0, 3)  # שורה 2, טור 1
        grid_layout.addWidget(buttons[4], 1, 0)  # שורה 2, טור 2
        grid_layout.addWidget(buttons[5], 1, 1)  # שורה 2, טור 3
        grid_layout.addWidget(buttons[6], 1, 2)  # שורה 3, טור 1
        grid_layout.addWidget(buttons[7], 2, 0)  # שורה 3, טור 2
        grid_layout.addWidget(buttons[8], 2, 1)  # שורה 3, טור 3
        grid_layout.addWidget(buttons[9], 1, 3)  # שורה 4, טור 1
        grid_layout.addWidget(buttons[10], 2, 2)  # שורה 4, טור 2
        grid_layout.addWidget(buttons[11], 2, 3)  # שורה 4, טור 3
        
        layout.addLayout(grid_layout)

        # יצירת layout אופקי עבור הכפתורים הקטנים
        buttons_layout = QHBoxLayout()
        
        # כפתור "אודות התוכנה" (הכי ימני)
        about_button = QPushButton("i")
        about_button.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                font-size: 12pt;
                border-radius: 30px;
                background-color: #eaeaea;
                border: 1px solid #888;
            }
            QPushButton:hover {
                background-color: #b7b5b5;
            }                                   
        """)
        about_button.setCursor(QCursor(Qt.PointingHandCursor))
        about_button.clicked.connect(self.open_about_dialog)
        about_button.setFixedSize(40, 40)
        about_button.setToolTip("אודות התוכנה")

        # כפתור "עדכונים" (שמאלה מ"אודות")
        update_button = QPushButton("⭳")
        update_button.setStyleSheet("""
            QPushButton {
                font-weight: bold;
                font-size: 12pt;
                border-radius: 30px;
                background-color: #eaeaea;
                border: 1px solid #888;
            }
            QPushButton:hover {
                background-color: #b7b5b5;
            }                                   
        """)
        update_button.setCursor(QCursor(Qt.PointingHandCursor))
        update_button.clicked.connect(lambda: self.check_for_updates(silent=False))
        update_button.setFixedSize(40, 40)
        update_button.setToolTip("בדיקת עדכונים")
        # עיצוב הכפתורים הקטנים
        buttons_layout.setContentsMargins(0, 20, 20, 0)  # ביטול שוליים
        
        buttons_layout.addWidget(about_button)
        buttons_layout.addWidget(update_button)
        
        # הוספת spacer לדחיפת הכפתורים לימין
        buttons_layout.addStretch()

        layout.addLayout(buttons_layout)

        # שורת הסטטוס עבור מצב העדכונים
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("""
            color: #666666;
            font-size: 24px;
            padding: 0px 0px;
            background-color: transparent;
            border-radius: 10px;
        """)     
        layout.addWidget(self.status_label)

        # הגדרת ה-Layout של החלון
        self.setLayout(layout)

    def open_about_dialog(self):
        """פתיחת חלון 'אודות'"""
        dialog = AboutDialog(self)
        dialog.exec_()

    def open_create_headers_otzria(self):
        self.create_headers_window = CreateHeadersOtZria()
        self.create_headers_window.show()

    def open_create_single_letter_headers(self):
        self.create_single_letter_headers_window = CreateSingleLetterHeaders()
        self.create_single_letter_headers_window.show()

    def open_add_page_number_to_heading(self):
        self.add_page_number_window = AddPageNumberToHeading()
        self.add_page_number_window.show()

    def open_change_heading_level(self):
        self.change_heading_level_window = ChangeHeadingLevel()
        self.change_heading_level_window.show()

    def open_emphasize_and_punctuate(self):
        self.emphasize_and_punctuate_window = EmphasizeAndPunctuate()
        self.emphasize_and_punctuate_window.show()

    def open_create_page_b_headers(self):
        self.create_page_b_headers_window = CreatePageBHeaders()
        self.create_page_b_headers_window.show()

    def open_replace_page_b_headers(self):
        self.replace_page_b_headers_window = ReplacePageBHeaders()
        self.replace_page_b_headers_window.show()

    def open_check_heading_errors_original(self):
        self.check_heading_errors_original_window = CheckHeadingErrorsOriginal()
        self.check_heading_errors_original_window.show()

    def open_check_heading_errors_custom(self):
        self.check_heading_errors_custom_window = CheckHeadingErrorsCustom()
        self.check_heading_errors_custom_window.show()
        
    def open_Image_To_Html_App(self):
        self.Image_To_Html_App_window = ImageToHtmlApp()
        self.Image_To_Html_App_window.show()

    def open_Text_Cleaner_App(self):
        self.Text_Cleaner_App_window = TextCleanerApp()
        self.Text_Cleaner_App_window.show()

    def open_Main_Window(self):
        self.Main_Window_window = MainWindow()
        self.Main_Window_window.show()
   
    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)


    def get_default_downloads_folder(self):
        """מחזיר את נתיב תיקיית ההורדות הברירת מחדל של המשתמש"""
        try:
            import winreg
            # קריאה מהרישום עבור תיקיית ההורדות
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                               r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as key:
                downloads_path = winreg.QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")[0]
                if os.path.exists(downloads_path):
                    return downloads_path
        except:
            pass
        
        # אם לא הצליח לקרוא מהרישום, נסה נתיבים רגילים
        possible_paths = [
            os.path.join(os.path.expanduser("~"), "Downloads"),
            os.path.join(os.path.expanduser("~"), "הורדות"),
            "C:\\Users\\{}\\Downloads".format(os.getenv("USERNAME", "")),
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        
        # אם שום נתיב לא נמצא, החזר את התיקייה הנוכחית
        return os.getcwd()

# ========================================== #
# עדכונים והתקנות
# ========================================== #

    def unblock_downloaded_file(self, file_path):
        """הסרת חסימה אינטרנטית מקובץ שהורד (אם קיימת)"""
        try:
            zone_file = file_path + ":Zone.Identifier"
            if os.path.exists(zone_file):
                os.remove(zone_file)
                print(f"הוסרה חסימה אינטרנטית מהקובץ: {file_path}")
            else:
                print("לא נמצאה חסימה אינטרנטית בקובץ")
        except Exception as e:
            print(f"שגיאה בהסרת חסימה אינטרנטית: {e}")
            # לא עוצרים את התהליך אם יש שגיאה
      
    def check_for_update_ready(self):
        """בדיקה אם העדכון מוכן להתקנה"""
        current_dir = os.path.dirname(sys.executable)
        marker_file = os.path.join(current_dir, "update_ready.txt")
        
        if os.path.exists(marker_file):
            try:
                with open(marker_file, "r", encoding="utf-8") as f:
                    new_version = f.readline().strip()
                
                # מחיקת קובץ הסימון
                os.remove(marker_file)
                
                # הפעלת העדכון
                temp_exe = os.path.join(current_dir, f'new_version_{new_version}.exe')
                current_exe = sys.executable
                
                if os.path.exists(temp_exe):
                    # הסרת חסימה אינטרנטית מהקובץ החדש
                    self.unblock_downloaded_file(temp_exe)                    
                    try:
                        # שחרור הקובץ הנוכחי מהזיכרון                   
                        # שליחת הודעת סגירה לכל החלונות של התוכנה
                        def enum_windows_callback(hwnd, _):
                            if win32gui.IsWindowVisible(hwnd):
                                t, w = win32gui.GetWindowText(hwnd), win32gui.GetClassName(hwnd)
                                if "עריכת ספרי דיקטה" in t: 
                                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                        
                        win32gui.EnumWindows(enum_windows_callback, None)
                        
                        # המתנה קצרה לסגירת החלונות
                        time.sleep(1)
                        
                        # העתקת הקובץ החדש
                        shutil.copy2(temp_exe, current_exe)
                        os.remove(temp_exe)
                        
                        # הפעלה מחדש של התוכנה
                        os.startfile(current_exe)
                        
                        # סגירה מסודרת
                        QApplication.quit()
                        
                    except Exception as e:
                        print(f"שגיאה בהחלפת הקובץ: {e}")
                        
            except Exception as e:
                print(f"שגיאה בהתקנת העדכון: {e}")

    def handle_update_available(self, download_url, new_version, published_at, release_body, file_size=None):
        """טיפול בעידכון"""
        self.status_label.setText("עדכון זמין")

        # יצירת חלון דיאלוג מותאם אישית
        dialog = QDialog(self)
        dialog.setLayoutDirection(Qt.RightToLeft)  # תמיכה בעברית
        dialog.setWindowTitle("נמצא עדכון")
        
        # הגדרת גודל החלון (רוחב, גובה)
        dialog.resize(700, 600)  # ניתן לשנות לפי הצורך
        
        # יצירת layout עיקרי
        main_layout = QVBoxLayout(dialog)
        
        # === חלק עליון קבוע (לא נגלל) ===
        top_widget = QTextEdit()
        top_widget.setReadOnly(True) # מונע עריכה
        top_widget.setLayoutDirection(Qt.RightToLeft)
        top_widget.setMaximumHeight(205)  # גובה קבוע
        top_widget.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)  # ללא פס גלילה
        top_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        top_widget.setFrameStyle(0)  # מבטל את המסגרת
        
        # עיצוב התאריך בעברית
        formatted_date = ""
        if published_at:
            try:
                # המרת מספרים לעברית
                hebrew_numbers = {
                    1: "א", 2: "ב'", 3: "ג'", 4: "ד'", 5: "ה'", 6: "ו'", 7: "ז'", 8: "ח'", 9: "ט'", 10: "י'",
                    11: 'י"א', 12: 'י"ב', 13: 'י"ג', 14: 'י"ד', 15: 'ט"ו', 16: 'ט"ז', 17: 'י"ז', 18: 'י"ח', 19: 'י"ט', 20: "כ'",
                    21: 'כ"א', 22: 'כ"ב', 23: 'כ"ג', 24: 'כ"ד', 25: 'כ"ה', 26: 'כ"ו', 27: 'כ"ז', 28: 'כ"ח', 29: 'כ"ט', 30: "ל'"
                }   
                # שמות חודשים עבריים
                hebrew_months = {
                    'Tishrei': 'תשרי', 'Marcheshvan': 'חשון', 'Kislev': 'כסלו', 'Tevet': 'טבת',
                    'Shevat': 'שבט', 'Adar': 'אדר', 'Adar I': 'אדר א', 'Adar II': 'אדר ב',
                    'Nissan': 'ניסן', 'Iyar': 'אייר', 'Sivan': 'סיון', 'Tammuz': 'תמוז',
                    'Av': 'אב', 'Elul': 'אלול'
                }
                def convert_hebrew_year(year):
                    """המרת שנה עברית למספרים עבריים"""
                    # המרה פשוטה לשנים נפוצות (תש"פ-תת"ק בערך)
                    year_mapping = {
                        5785: 'תשפ"ה', 5786: 'תשפ"ו', 5787: 'תשפ"ז', 5788: 'תשפ"ח', 5789: 'תשפ"ט', 5790: 'תש"צ',
                        5791: 'תשצ"א', 5792: 'תשצ"ב', 5793: 'תשצ"ג', 5794: 'תשצ"ד', 5795: 'תשצ"ה',
                        5796: 'תשצ"ו', 5797: 'תשצ"ז', 5798: 'תשצ"ח', 5799: 'תשצ"ט', 5800: 'ת"ת'
                    }
                    return year_mapping.get(year, str(year))
                
                # המרת התאריך ל- datetime
                dt = datetime.fromisoformat(published_at.replace('Z', '+00:00'))
                hebrew_date = dates.GregorianDate(dt.year, dt.month, dt.day).to_heb()
                
                # קבלת המידע מהתאריך העברי
                hebrew_day = hebrew_date.day
                hebrew_month_eng = hebrew_date.month_name()  # שם החודש באנגלית
                hebrew_year = hebrew_date.year
                
                # המרה לעברית
                hebrew_day_str = hebrew_numbers.get(hebrew_day, str(hebrew_day))
                hebrew_month_heb = hebrew_months.get(hebrew_month_eng, hebrew_month_eng)
                hebrew_year_str = convert_hebrew_year(hebrew_year)
                
                # בניית המחרוזת
                time_str = dt.strftime("%H:%M")
                formatted_date = f"{hebrew_day_str} {hebrew_month_heb} {hebrew_year_str}, בשעה {time_str}"
                
            except Exception as e:
                print(f"שגיאה בהמרת תאריך עברי: {e}")
                traceback.print_exc()
                
                # fallback לתאריך רגיל
                try:
                    dt = datetime.fromisoformat(published_at.replace('Z', '+00:00'))
                    formatted_date = dt.strftime("%d/%m/%Y %H:%M")
                except:
                    formatted_date = published_at        
        
        # תוכן החלק העליון
        top_message = f"<div style='text-align: left; direction: rtl; font-family: Segoe UI, Arial;'>"
        top_message += f"<h2 style='color: #2E86AB; text-align: center; margin: 5px 0;'>נמצאה גרסה חדשה!</h2>"
        top_message += f"<p style='margin: 3px 0;'><b>גרסה:</b> {new_version}</p>"
        
        if formatted_date:
            top_message += f"<p><b>תאריך פרסום:</b> {formatted_date}</p>"  

        if file_size:
            # המרה לMB עם עיגול לשתי ספרות אחרי הנקודה
            size_mb = file_size / (1024 * 1024)
            top_message += f"<p><p style='margin: 3px 0;'><b>גודל הקובץ:</b> {size_mb:.2f} MB</p></p>"
            top_message += f"<h3 style='color: #A23B72; text-align: center; margin-bottom: 15px;'>פרטי העדכון:</h3>"
        
        top_message += f"</div>"
        top_widget.setHtml(top_message)

        # === חלק אמצעי (נגלל) ===
        middle_widget = QTextEdit()
        middle_widget.setReadOnly(True)
        middle_widget.setLayoutDirection(Qt.RightToLeft)
        
        # עיצוב פס הגלילה
        middle_widget.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: white;
            }
            QScrollBar:vertical {
                background: #f0f0f0;
                width: 8px;
                border-radius: 4px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                border-radius: 4px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a0a0a0;
            }
            QScrollBar::handle:vertical:pressed {
                background: #808080;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
                border: none;
                background: none;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
        """)
                
        # תוכן החלק האמצעי (רק פרטי העדכון)
        middle_message = f"<div style='text-align: left; direction: rtl; font-family: Segoe UI, Arial;;'>"        
        if release_body:
            middle_message += f"<div style='background-color: #f5f5f5; padding: 15px; border-radius: 5px; line-height: 1.5;'>{release_body}</div>"
        else:
            middle_message += f"<p style='text-align: center; color: #666;'>אין פרטים נוספים על העדכון</p>"
        middle_message += f"</div>"
        middle_widget.setHtml(middle_message)

        
        # === חלק תחתון קבוע (לא נגלל) ===
        bottom_widget = QTextEdit()
        bottom_widget.setReadOnly(True)
        bottom_widget.setLayoutDirection(Qt.RightToLeft)
        bottom_widget.setMaximumHeight(50)  # גובה קבוע
        bottom_widget.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        bottom_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        bottom_widget.setFrameStyle(0)  # מבטל את המסגרת

        # תוכן החלק התחתון
        bottom_message = f"<div style='text-align: left; direction: rtl; font-family: Segoe UI, Arial;;'>"
        bottom_message += f"<h3 style='text-align: center; color: #F18F01; margin: 10px 0;'>האם ברצונך לעדכן כעת?</h3>"
        bottom_message += f"</div>"
        bottom_widget.setHtml(bottom_message)
        
        # הוספת כל החלקים ל-layout
        main_layout.addWidget(top_widget)
        main_layout.addWidget(middle_widget)  # רק החלק הזה יגלל
        main_layout.addWidget(bottom_widget)
        
        # יצירת layout לכפתורים
        button_layout = QHBoxLayout()
        
        # יצירת layout לכפתורים
        button_layout = QHBoxLayout()
        
        # יצירת כפתורים
        yes_button = QPushButton("כן")
        no_button = QPushButton("לא")

        # עיצוב הכפתורים
        yes_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 18px;
                border-radius: 5px;
                font-weight: bold;
                min-width: 80px;
                min-height: 30px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3e8e41;
            }    
        """)
        
        no_button.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 18px;
                border-radius: 5px;
                font-weight: bold;
                min-width: 80px;
                min-height: 30px;                                
            }
            QPushButton:hover {
                background-color: #da190b;
            }
            QPushButton:pressed {
                background-color: #c62828;
            }
        """)
        
        # הוספת הכפתורים ל-layout (כן משמאל, לא מימין עבור עברית)
        button_layout.addStretch()  # מוסיף רווח
        button_layout.addWidget(yes_button)
        button_layout.addWidget(no_button)
        button_layout.addStretch()  # מוסיף רווח
        
        # הוספת layout הכפתורים ל-layout הראשי
        main_layout.addLayout(button_layout)
        
        # חיבור הכפתורים לפונקציות
        yes_button.clicked.connect(dialog.accept)
        no_button.clicked.connect(dialog.reject)
        
        # הגדרת כפתור ברירת המחדל
        yes_button.setDefault(True)
        yes_button.setFocus()
        
        # הרצת החלון
        result = dialog.exec_()

    # בדיקה איזה כפתור נלחץ
        if result == QDialog.Accepted:
            self.download_and_install_update(download_url, new_version)
        else:
            # אם המשתמש בחר שלא לעדכן
            self.status_label.setText("עדכון זמין")
            print("העדכון נדחה")
                        
    def download_and_install_update(self, download_url, new_version):
        """הורדת והתקנת העדכון"""
        print(f"מתחיל תהליך הורדת עדכון: {download_url}")
        
        # עדכון סטטוס
        self.status_label.setText("מוריד את העדכון האחרון...")
        
        # יצירת חלון העדכון המעוצב
        self.update_window = QMainWindow(self)
        self.update_window.setWindowTitle("הורדת עדכון")
        self.update_window.setFixedWidth(600)
        self.update_window.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint)
        self.update_window.setLayoutDirection(Qt.RightToLeft)

        # מיקום החלון במרכז החלון ההורה
        parent_center = self.mapToGlobal(self.rect().center())
        self.update_window.move(
            parent_center.x() - self.update_window.width() // 2,
            parent_center.y() - 150 // 2
        )

        # יצירת הממשק
        central_widget = QWidget()
        self.update_window.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # כותרת ראשית
        self.main_status = QLabel("מתכונן להורדת העדכון...")
        self.main_status.setStyleSheet("""
            QLabel {
                color: #1a365d;
                font-family: "Segoe UI", Arial;
                font-size: 16px;
                font-weight: bold;
                padding: 10px;
            }
        """)
        self.main_status.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.main_status)

        # פירוט המשימה הנוכחית
        self.detail_status = QLabel("מתחבר לשרת...")
        self.detail_status.setStyleSheet("""
            QLabel {
                color: #666666;
                font-family: "Segoe UI", Arial;
                font-size: 12px;
                padding: 5px;
            }
        """)
        self.detail_status.setAlignment(Qt.AlignCenter)
        self.detail_status.setWordWrap(True)
        layout.addWidget(self.detail_status)

        # סרגל התקדמות
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #2b4c7e;
                border-radius: 15px;
                padding: 5px;
                text-align: center;
                background-color: white;
                height: 30px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 13px;
            }
        """)
        layout.addWidget(self.progress_bar)

        central_widget.setLayout(layout)
        self.update_window.show()
        
        # הפעלת thread להורדת העדכון
        self.downloader = UpdateDownloader(download_url, new_version)
        self.downloader.progress.connect(self.update_progress)
        self.downloader.status.connect(self.update_status)
        self.downloader.finished.connect(self.download_finished)
        self.downloader.error.connect(self.download_error)
        self.downloader.start()

    def update_progress(self, percentage, downloaded_mb, total_mb):
        """עדכון התקדמות ההורדה"""
        self.progress_bar.setValue(percentage)
        if total_mb > 0:
            self.detail_status.setText(f"מוריד: {downloaded_mb:.1f}MB מתוך {total_mb:.1f}MB")
        else:
            self.detail_status.setText(f"הורד: {downloaded_mb:.1f}MB")

    def update_status(self, status):
        """עדכון סטטוס ההורדה"""
        self.main_status.setText(status)

    def download_finished(self, file_path):
        """סיום הורדה מוצלח"""
        self.update_window.close()
        self.status_label.setText("סוגר את התוכנה, ופותח מחדש את העדכון...")
        
        QMessageBox.information(
            self,
            "הורדה הושלמה",
            f"העדכון הורד בהצלחה לתיקיית ההורדות!\nבמיקום:\n {file_path}\n\nהתוכנה תופעל מחדש לאחר השלמת ההתקנה."
        )
        
        # הפעלת ההתקנה
        try:
            if sys.platform == 'win32':
                os.startfile(file_path)
                # סגירת התוכנה הנוכחית
                QTimer.singleShot(2000, QApplication.quit)
        except Exception as e:
            QMessageBox.critical(
                self,
                "שגיאה בהפעלת ההתקנה",
                f"העדכון הורד אך לא ניתן להפעיל את ההתקנה:\n{str(e)}\n\nאנא הפעל את הקובץ ידנית: {file_path}"
            )

    def download_error(self, error_msg):
        """שגיאה בהורדה"""
        self.update_window.close()
        self.status_label.setText("שגיאה בהורדת העדכון")
        
        QMessageBox.critical(
            self,
            "שגיאה בהורדת העדכון",
            f"לא ניתן להוריד את העדכון:\n{error_msg}"
        )

class AboutDialog(QDialog):
    """חלון 'אודות'"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("אודות התוכנה")
        self.setLayoutDirection(Qt.RightToLeft)
        self.setContentsMargins(50, 50, 50, 50)  # הוספת שוליים
        self.current_version = VERSION
        layout = QVBoxLayout()

        title_label = QLabel("עריכת ספרי דיקטה עבור 'אוצריא'")
        title_label.setStyleSheet("font-size: 14pt; font-weight: bold;")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        version_label = QLabel(f"גירסה: {self.current_version}")
        version_label.setStyleSheet("font-size: 10pt;")
        version_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(version_label)

        date_label = QLabel("תאריך: י סיון תשפה")
        date_label.setStyleSheet("font-size: 10pt;")
        date_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(date_label)

        dev_label = QLabel("נכתב על ידי 'מתנדבי אוצריא', להצלחת לומדי התורה הקדושה")
        dev_label.setStyleSheet("font-size: 10pt;")
        dev_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(dev_label)

        # קישור ל-GitHub
        github_label = QLabel('ניתן להוריד את הגירסא האחרונה, וכן קובץ הדרכה, בקישור הבא: <a href="https://github.com/YOSEFTT/EditingDictaBooks/releases">GitHub</a>')
        github_label.setStyleSheet("font-size: 10pt;")
        github_label.setOpenExternalLinks(True)
        github_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(github_label)

        # קישור ל-מתמחים.טופ
        mitmachimtop_label = QLabel('או כאן: <a href="https://mitmachim.top/topic/77509/%D7%94%D7%A1%D7%91%D7%A8-%D7%94%D7%95%D7%A1%D7%A4%D7%AA-%D7%95%D7%98%D7%99%D7%A4%D7%95%D7%9C-%D7%91%D7%A1%D7%A4%D7%A8%D7%99%D7%9D-%D7%9C-%D7%90%D7%95%D7%A6%D7%A8%D7%99%D7%90-%D7%9B%D7%A2%D7%AA-%D7%96%D7%94-%D7%A7%D7%9C">מתמחים.טופ</a>')
        mitmachimtop_label.setStyleSheet("font-size: 10pt;")
        mitmachimtop_label.setOpenExternalLinks(True)
        mitmachimtop_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(mitmachimtop_label)

        # קישור ל-דרייב
        drive_label = QLabel('או בדרייב: <a href="http://did.li/dicta-library">כאן</a> או <a href="https://drive.google.com/open?id=1Sqj3sCyYJnVB60KAZxgmDNO7x52qoVm6&usp=drive_fs">כאן</a>')
        drive_label.setStyleSheet("font-size: 10pt;")
        drive_label.setOpenExternalLinks(True)
        drive_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(drive_label)

        info_label = QLabel("אפשר לבקש את התוכנה\nוכן לקבל תמיכה וסיוע בכל הקשור לתוכנה זו ולתוכנת 'אוצריא'\nבמייל הבא:")
        info_label.setStyleSheet("font-size: 10pt;")
        info_label.setAlignment(Qt.AlignCenter)
        
        gmail_label = QLabel('<a href="https://mail.google.com/mail/u/0/?view=cm&fs=1&to=otzaria.1%40gmail.com%E2%80%AC">otzaria.1@gmail.com</a>')
        gmail_label.setStyleSheet("font-size: 10pt;")
        gmail_label.setOpenExternalLinks(True)
        gmail_label.setAlignment(Qt.AlignCenter)
        
        layout.addWidget(info_label)
        layout.addWidget(gmail_label)

        self.setLayout(layout)

# ==========================================
#  update
# ==========================================
class UpdateDownloader(QThread):
    """Thread להורדת העדכון"""
    progress = pyqtSignal(int, float, float)  # percentage, downloaded_mb, total_mb
    status = pyqtSignal(str)
    finished = pyqtSignal(str)  # file_path
    error = pyqtSignal(str)

    def __init__(self, download_url, version):
        super().__init__()
        self.download_url = download_url
        self.version = version
        self.update_checker = UpdateChecker("0.0.0")  # רק בשביל get_with_ssl_fallback

    def run(self):
        try:
            self.status.emit("מתחבר לשרת...")
            
            # קביעת נתיב הקובץ - שימוש בתיקיית ההורדות הברירת מחדל
            downloads_dir = self.get_default_downloads_folder()
            
            filename = f"עריכת ספרי דיקטה {self.version}.exe"
            file_path = os.path.join(downloads_dir, filename)
            
            self.status.emit("מוריד את העדכון...")
            
            # הורדת הקובץ
            response = self.update_checker.get_with_ssl_fallback(self.download_url, stream=True)
            response.raise_for_status()
            
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            block_size = 8192
            
            with open(file_path, 'wb') as f:
                for data in response.iter_content(block_size):
                    if data:
                        downloaded += len(data)
                        f.write(data)
                        
                        # חישוב התקדמות
                        if total_size > 0:
                            percentage = int((downloaded / total_size) * 100)
                        else:
                            percentage = 0
                        
                        downloaded_mb = downloaded / (1024 * 1024)
                        total_mb = total_size / (1024 * 1024) if total_size > 0 else 0
                        
                        self.progress.emit(percentage, downloaded_mb, total_mb)
            
            self.status.emit("ההורדה הושלמה בהצלחה!")
            self.finished.emit(file_path)
            
        except Exception as e:
            self.error.emit(f"שגיאה בהורדת העדכון: {str(e)}")

    def get_default_downloads_folder(self):
        """מחזיר את נתיב תיקיית ההורדות הברירת מחדל של המשתמש"""
        try:
            import winreg
            # קריאה מהרישום עבור תיקיית ההורדות
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                               r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as key:
                downloads_path = winreg.QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")[0]
                if os.path.exists(downloads_path):
                    return downloads_path
        except:
            pass
        
        # אם לא הצליח לקרוא מהרישום, נסה נתיבים רגילים
        possible_paths = [
            os.path.join(os.path.expanduser("~"), "Downloads"),
            os.path.join(os.path.expanduser("~"), "הורדות"),
            "C:\\Users\\{}\\Downloads".format(os.getenv("USERNAME", "")),
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        
        # אם שום נתיב לא נמצא, החזר את התיקייה הנוכחית
        return os.getcwd()

class UpdateChecker(QThread):
    update_available = pyqtSignal(str, str, str, str, int) # download_url, version, name, published_at, body
    no_update = pyqtSignal()  
    error = pyqtSignal(str)  

    def __init__(self, current_version, parent=None):
        super().__init__(parent)
        self.current_version = current_version
        self.headers = {
            'Accept': 'application/vnd.github.v3+json',
            'User-Agent': 'EditingDictaBooks-UpdateChecker'
        }

    def get_with_ssl_fallback(self, url, **kwargs):
        """
        מנסה להתחבר עם תעודות SSL שונות בסדר עדיפות
        """
        ssl_options = []
        
        # אפשרות 1: תעודת נטפרי אם קיימת
        netfree_cert = self._find_netfree_cert()
        if netfree_cert:
            ssl_options.append(('NetFree Certificate', netfree_cert))
        
        # אפשרות 2: תעודות מערכת ברירת מחדל
        ssl_options.append(('System Default', True))
        
        # אפשרות 3: תעודות certifi
        try:
            ssl_options.append(('Certifi Bundle', certifi.where()))
        except:
            pass
        
        # אפשרות 4: ללא אימות SSL (כמוצא אחרון)
        ssl_options.append(('No SSL Verification', False))
        
        for option_name, ssl_setting in ssl_options:
            try:
                print(f"מנסה התחברות עם: {option_name}")
                response = requests.get(url, verify=ssl_setting, timeout=30, **kwargs)
                print(f"התחברות הצליחה עם: {option_name}")
                return response
            except requests.exceptions.SSLError as e:
                print(f"שגיאת SSL עם {option_name}: {str(e)}")
                continue
            except Exception as e:
                print(f"שגיאה כללית עם {option_name}: {str(e)}")
                continue
        
        # אם שום דרך לא עבדה
        raise requests.exceptions.ConnectionError("לא ניתן להתחבר לשרת בשום דרך")

    def _find_netfree_cert(self):
        """
        מחפש תעודת נטפרי במיקומים אפשריים
        """
        possible_paths = [
            r"C:\ProgramData\NetFree\CA\netfree-ca-list.crt",
            r"C:\NetFree\netfree-ca-list.crt",
            os.path.join(os.getcwd(), "netfree-ca-list.crt"),
            os.path.join(os.path.dirname(sys.executable), "netfree-ca-list.crt")
        ]
        
        for path in possible_paths:
            if os.path.exists(path) and os.path.getsize(path) > 0:
                print(f"נמצאה תעודת נטפרי ב: {path}")
                return path
        
        return None

    def run(self):
        try:
            api_url = "https://api.github.com/repos/YOSEFTT/EditingDictaBooks/releases/latest"
            
            print("מנסה להתחבר לשרת GitHub...")
            print(f"URL: {api_url}")
            
            # שימוש בפונקציה החדשה עם נסיונות SSL מרובים
            response = self.get_with_ssl_fallback(api_url, headers=self.headers)
            response.raise_for_status()
            
            latest_release = response.json()
            latest_version = latest_release['tag_name'].replace('v', '')
            published_at = latest_release.get('published_at', '')
            release_body = latest_release.get('body', '')

            print(f"גרסה נוכחית: {self.current_version}")
            print(f"גרסה אחרונה: {latest_version}")
            
            if self._compare_versions(latest_version, self.current_version):
                download_url = None
                # חיפוש גודל הקובץ
                file_size = None
                for asset in latest_release['assets']:
                    if asset['name'].lower().endswith('.exe'):
                        download_url = asset['browser_download_url']
                        file_size = asset.get('size', 0)  # גודל בבייטים
                        break
                
                if download_url:
                    self.update_available.emit(download_url, latest_version, published_at, release_body, file_size)
                else:
                    self.error.emit("נמצאה גרסה חדשה, אך לא נמצא קובץ הורדה מתאים")
            else:
                print("אין גרסה חדשה")
                self.no_update.emit()
                
        except requests.exceptions.ConnectionError as e:
            print(f"Connection Error: {str(e)}")
            self.error.emit("בעיית חיבור לשרת GitHub\nבדוק את חיבור האינטרנט שלך")
            
        except Exception as e:
            print(f"General Error: {str(e)}")
            self.error.emit(f"שגיאה כללית: {str(e)}")

    def _compare_versions(self, latest_version, current_version):
        """
        השוואת גרסאות
        """
        try:
            latest_version = latest_version.upper().strip('V')
            current_version = current_version.upper().strip('V')
            
            latest_parts = [int(x) for x in latest_version.split('.')]
            current_parts = [int(x) for x in current_version.split('.')]
            
            while len(latest_parts) < 3:
                latest_parts.append(0)  # תיקון: הוספת מספר ולא מחרוזת
            while len(current_parts) < 3:
                current_parts.append(0)  # תיקון: הוספת מספר ולא מחרוזת
            
            for i in range(3):
                if latest_parts[i] > current_parts[i]:
                    return True
                elif latest_parts[i] < current_parts[i]:
                    return False
                
            return False  # הגרסאות זהות
        
        except Exception as e:
            print(f"שגיאה בהשוואת גרסאות: {str(e)}")
            print(f"גרסה אחרונה: {latest_version}, גרסה נוכחית: {current_version}")
            return False

# ==========================================
# Main Application
# ==========================================
def main():
    if sys.platform == 'win32':
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.RightToLeft)    
    enter_filter = EnterKeyFilter()
    app.installEventFilter(enter_filter)
    main_menu = MainMenu()
    main_menu.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
