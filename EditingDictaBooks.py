from logging import info
import sys
import ctypes
from PyQt5 import QtGui, QtCore, QtWidgets
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QLabel,
    QLayout, QFileDialog, QLineEdit, QMessageBox, QComboBox, QHBoxLayout,
    QCheckBox, QTextEdit, QDialog, QFrame, QSplitter, QGridLayout, QSpacerItem, QSizePolicy
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap, QCursor
from PyQt5.QtWinExtras import QtWin
from pyluach import gematria
from bs4 import BeautifulSoup
import gematriapy
import re
import os
import base64
import urllib.request

# מזהה ייחודי לאפליקציה
myappid = 'MIT.LEARN_PYQT.dictatootzaria'

# מחרוזת Base64 של האייקון (החלף את זה עם המחרוזת שתקבל אחרי המרת הקובץ שלך ל־Base64)
icon_base64 = "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAYAAADimHc4AAAGP0lEQVR4Ae2dfUgUaRjAn2r9wkzs3D4OsbhWjjKk7EuCLY9NrETJzfujQzgNBBEEIQS9/EOM5PCz68Ly44gsNOjYXVg/2NDo6o8iWg/ME7/WVtJ1XTm7BPUky9tnQLG7dnd2dnee9eb9waDs7DzPu/ObnX3nnZlnZMADpVIZLpfLUywWyzejo6MHbDZbtP3lED7LMpwjczYzNTX1q+np6R+ePn36HbAV7hMcCkhJSSnQ6/U/2v8NErE9kuM/Avbv3x8YEBDQ0N7e/j1Fg6TGJwJw5QcFBf1qNBpTqRokNT4RYN/yf2ErX1xWBSQnJxcaDIZMysZIEU6AWq3+WqPRXKVujBThBLx+/brE/ieAuC2SRHbq1Kkvurq6vqVuiFSRRUREpAHr65MhGx8fT6RuhJSRmUymA9SNkDIym832JWUD9u7diweAYD8AFC3nwsIC9PT0YOdDtJyOwF6Qy09eU1MDCoXC4fyqqip48uSJW4k3b94MN2/ehKSkJLeW8xbLy8tw//59KCwshKWlJUExXK2XsrIyePnypdMYTkdDVzhx4gQcOnTI4fzW1lY+YVbZtGkTt8yRI0fcWs6bbNiwAS5cuAAhISGQm5srKIar9dLQ0OAyBi8B3iYzM5N05a/l3Llz0NLS4vY32FuQCEhLS6NI6xCUICkBO3fupEjrkOjoaLLcJALwB9Cf2LhxI1luEgGOaGpqgvLycp/Fb2xsBJVK5bP4QvArAe/fv4f5+Xmfxf/w4YPPYgvFrwRIESaAGCaAGCaAGCaAGCaAGCaAGCaAGCaAGCaAGCaAGL8SsG/fPu5kja+IioryWWyh+JWAkydPcpOUYOcDgLY9JAIGBwedXk0gNkNDQ2S5SQTU19fDmTNnSM9ErYDnIG7fvk2Wn0TAixcvoKKiAoqKiijSr/Lx40e4fPmy9L4ByLVr12BsbAwKCgogJiZG1G8Drvj+/n6orKwEg8EgWt7PQdoL0mq13CSTySAwMFC0vIuLi35zetIvuqF4aaDQywPXO34hQMowAcQwAcQwAcQwAcSQCjh8+DAWBSHLj13g3t5esvwIqYCjR49CaWkpWX6z2SxtAQwmgBwmgBgmgBgmgBhSAXizNA4JU/Hq1Suy3CuQCnj+/Dk3SRm2CyKGCSCGCSCGCSCGCSCGVACWCNi9ezdZ/oGBAbDZbGT5EVIB2dnZpKOhWVlZcOfOHbL8CNsFEcMEEMMEEMMEEMMEEMMEEEMqYHZ2FiwWC1l+rB9KDamA2tpabpIybBdEDBNADC8Brm5mELPusz8RFhbmcQxeAt6+fet0fmxsrMcNWW8EBwfDrl27nL5nbm7OZRxeAkwmk9P5GRkZUFxcLKm7XPAzu/rmj4yMuIzDSwDe1ZiXl+dwPg4po4ArV67wCbfu2bp1K1y96vyZR1NTU/DmzRuXsXgJwDsJ8c5CZ3cy4rAy1uO/d+8en5DrFrlcDm1tbS7LHXd0dPCKx0uA1WrlLuU+f/68w/egnObmZq4kvE6n43ZbnuyS8AbqZ8+eCV5+LUqlkitX7wnh4eGQkJAAFy9ehB07djh9L5Y+uHXrFq+4vLuhuIVjlXGs/e8I/JBnz57lJk/BM1Xbt2/3OA7S2dkJoaGhXonFB71ez+22+cBbQF9fH9TV1UF+fr7ghkkBHN64dOkS7/e7dSCGpQVw696zZ4/bDZMKuPL59H5WcEsAFtbGgkqPHz+W7MGXM+7evcs9F8cd3B6KwGs58QkYDx48gC1btri7+P8W7PXk5OS4vZygsaCHDx9CXFwcVFdXQ3p6ul+UnaEC+/vYQcESPAIKP80IHozDSid4NBgfHw8lJSX4/Hmu6IZUmJiYgBs3bsD169cFP/PAfiwx7PEaw2v81Wo1bNu2jfuLNd/wR9rT3dPMzIynTVtleHiYe1yVJ+CA5OTkJBiNRq5biw/9wYNTT1AoFD1e22Sx344HH3wPQMTk4MGD1E34LJGRkY+ks8/wP2bNZnMnE0CESqX6ubu7e44JIMD+e/nHu3fvuMdFMQHi89fx48fTdTod13ViAsRl3t5TVGs0muGVF5gAkbB3g2dOnz6NK/+3ta8zASJg3+f/fuzYsQytVjv673lMgA+xb/V/JiYm/mS1Wiv1ev3fn3vPP+R95FTm9cojAAAAAElFTkSuQmCC="

# ==========================================
# Script 1: יצירת כותרות לאוצריא
# ==========================================

class CreateHeadersOtZria(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("יצירת כותרות לאוצריא")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setGeometry(100, 100, 600, 300)
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
        file_layout.addWidget(browse_button)
        file_layout.addWidget(self.file_entry)
        file_layout.addWidget(file_label)
        layout.addLayout(file_layout)

        # מילה לחיפוש
        search_layout = QHBoxLayout()
        search_label = QLabel("מילה לחפש:")
        search_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        search_label.setStyleSheet("font-size: 20px;")
        self.level_var = QComboBox()
        self.level_var.setLayoutDirection(Qt.RightToLeft)
        self.level_var.setStyleSheet("font-size: 20px;")
        self.level_var.setFixedSize(150, 40)  # רוחב: 150 פיקסלים, גובה: 40 פיקסלים
        self.level_var.setLayoutDirection(Qt.RightToLeft)
        search_choices = ["דף", "עמוד", "פרק", "פסוק", "שאלה", "סימן", "סעיף", "הלכה", "הלכות", "סק"]
        self.level_var.addItems(search_choices)
        self.level_var.setEditable(True)  # מאפשר להקליד
        search_layout.addWidget(self.level_var)
        search_layout.addWidget(search_label)
       
        layout.addLayout(search_layout)

        # הסבר למשתמש
        explanation = QLabel(
            "בתיבת 'מילה לחפש' יש לבחור או להקליד את המילה בה אנו רוצים שתתחיל הכותרת.\nלדוג': פרק/פסוק/סימן/סעיף/הלכה/שאלה/עמוד/סק\n\nשים לב!\nאין להקליד רווח אחרי המילה, וכן אין להקליד את התו גרש (') או גרשיים (\"), וכן אין להקליד יותר ממילה אחת\n"
        )
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setStyleSheet("font-size: 21px;")
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # מספר סימן מקסימלי
        end_layout = QHBoxLayout()
        end_label = QLabel("מספר סימן מקסימלי:")
        self.end_var = QComboBox()
        self.end_var.addItems([str(i) for i in range(1, 1000)])
        self.end_var.setCurrentText("999")
        self.end_var.setFixedWidth(65)
        end_layout.addWidget(self.end_var)
        end_layout.addWidget(end_label)
        layout.addLayout(end_layout)

        # רמת כותרת
        heading_layout = QHBoxLayout()
        self.heading_label = QLabel("רמת כותרת:")
        self.heading_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.heading_label.setStyleSheet("font-size: 20px;")
        self.heading_level_var = QComboBox()
        self.heading_level_var.setStyleSheet("font-size: 20px;")
        self.heading_level_var.addItems([str(i) for i in range(2, 7)])
        self.heading_level_var.setCurrentText("2")
        self.heading_level_var.setStyleSheet("font-size: 20px;")
        self.heading_level_var.setFixedWidth(50)
        heading_layout.addWidget(self.heading_level_var, alignment=Qt.AlignRight)
        heading_layout.addWidget(self.heading_label)
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
        msg.setIcon(QMessageBox.Information)

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
class CreateSingleLetterHeaders(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("יצירת כותרות לאותיות בודדות")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setGeometry(100, 100, 650, 300)
        self.init_ui()

    def init_ui(self):
        self.setLayoutDirection(Qt.RightToLeft)  # הגדרת כיוון ימין לשמאל
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

        # תו בתחילת האות ותו בסוף האות
        start_char_label = QLabel("תו בתחילת האות:")
        self.start_var = QComboBox()
        self.start_var.setLayoutDirection(Qt.RightToLeft)  # הגדרת כיוון כללי
        self.start_var.addItems(["", "(", "["])
        self.start_var.setStyleSheet("text-align: right;")  # מוודא שהטקסט ייושר לימין
        end_char_label = QLabel("     תו/ים בסוף האות:")
        self.finde_var = QComboBox()
        self.finde_var.addItems(['', '.', ',', "'", "',", "'.", ']', ')', "']", "')", "].", ").", "],", "),", "'),", "').", "'],", "']."])
               
        regex_layout = QHBoxLayout()
        regex_layout.addWidget(start_char_label)
        regex_layout.addWidget(self.start_var)
        regex_layout.addWidget(end_char_label)
        regex_layout.addWidget(self.finde_var)
        layout.addLayout(regex_layout)
       
        # הסבר למשתמש
        explanation = QLabel(
            "שים לב!\nהבחירה בברירת מחדל [השורה הריקה], משמעותה סימון כל האפשרויות."
        )
        explanation.setAlignment(Qt.AlignCenter)
        explanation.setStyleSheet("font-size: 18px;")
        explanation.setWordWrap(True)
        layout.addWidget(explanation)

        # רמת כותרת
        heading_layout = QHBoxLayout()
        heading_label = QLabel("רמת כותרת:")
        heading_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        heading_label.setStyleSheet("font-size: 20px;")
        self.level_var = QComboBox()
        self.level_var.setFixedWidth(50)
        self.level_var.setStyleSheet("font-size: 20px;")
        self.level_var.addItems([str(i) for i in range(2, 7)])
        self.level_var.setCurrentText("3")
        heading_layout.addWidget(heading_label)
        heading_layout.addWidget(self.level_var, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        
        layout.addLayout(heading_layout)

        # תיבת סימון לחיפוש עם תווי הדגשה בלבד
        self.bold_var = QCheckBox("לחפש עם תווי הדגשה בלבד")
        self.bold_var.setChecked(True)
        layout.addWidget(self.bold_var)

        # התעלם מהתווים
        ignore_layout = QHBoxLayout()
        ignore_label = QLabel("התעלם מהתווים הבאים:")
        ignore_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.ignore_entry = QLineEdit()
        self.ignore_entry.setText('<big> </big> " ')
        ignore_layout.addWidget(ignore_label)
        ignore_layout.addWidget(self.ignore_entry)
        layout.addLayout(ignore_layout)

        # הסרת תווים
        remove_layout = QHBoxLayout()
        remove_label = QLabel("הסר את התווים הבאים:")
        remove_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.remove_entry = QLineEdit()
        self.remove_entry.setText('<b> </b> <big> </big> , : " \' . ( ) [ ] { }')
        remove_layout.addWidget(remove_label)
        remove_layout.addWidget(self.remove_entry)
        layout.addLayout(remove_layout)

        # מספר סימן מקסימלי
        end_layout = QHBoxLayout()
        end_label = QLabel("מספר סימן מקסימלי:")
        self.end_var = QComboBox()
        self.end_var.setFixedWidth(65)
        self.end_var.addItems([str(i) for i in range(1, 1000)])
        self.end_var.setCurrentText("999")
        end_layout.addWidget(end_label)
        end_layout.addWidget(self.end_var)
        layout.addLayout(end_layout)

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
        remove = ["<b>", "</b>"] + self.remove_entry.text().split()
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
        remove = ["<b>", "</b>", "<big>", "</big>", ":", '"', ",", ";", "[", "]", "(", ")", "'", ".", "״", "‚", "”", "’"]
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
        file_layout.addWidget(browse_button)
        file_layout.addWidget(self.file_entry)
        file_layout.addWidget(file_label)
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
        heading_layout.addWidget(self.replace_option, alignment=Qt.AlignRight)
        heading_layout.addWidget(replacement_label)
       
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

        updated_content = []
        changes_made = False

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

                        modified_next_line = re.sub(pattern, '', next_line).strip()
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
            QMessageBox.information(self, "!שים לב", "אין מה להחליף בקובץ זה")

    def run_script(self):
        file_path = self.file_entry.text()
        if file_path:
            replace_with = self.replace_option.currentText()
            self.process_file(file_path, replace_with)
        else:
            QMessageBox.warning(self, "קלט לא תקין", "אנא בחר קובץ או הזן נתיב")
  
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
        self.setGeometry(100, 100, 550, 250)
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
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
        file_layout.addWidget(browse_button)
        file_layout.addWidget(self.file_entry)
        file_layout.addWidget(file_label)
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
        current_level_layout.addWidget(self.current_level_var, alignment=Qt.AlignRight)
        current_level_layout.addWidget(current_level_label)
        layout.addLayout(current_level_layout)

        # רמת כותרת חדשה
        new_level_layout = QHBoxLayout()
        new_level_label = QLabel("רמת כותרת חדשה: (לדוגמא: 3)")
        new_level_label.setStyleSheet("font-size: 20px;")
        self.new_level_var = QComboBox()
        self.new_level_var.setStyleSheet("font-size: 20px;")
        self.current_level_var.setFixedWidth(50)
        self.new_level_var.addItems([str(i) for i in range(1, 10)])
        new_level_layout.addWidget(self.new_level_var, alignment=Qt.AlignRight)
        new_level_layout.addWidget(new_level_label)
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
            QMessageBox.critical(self, "קלט לא תקין", "לא נבחר קובץ!")
            return
        # בדיקת סוג הקובץ לפי סיומת
        if not file_path.lower().endswith('.txt'):
            QMessageBox.critical(self, "קלט לא תקין", "סוג הקובץ אינו נתמך\nבחר קובץ טקסט [בסיומת TXT.]")
            return
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()

                current_tag = f"h{current_level}"
                new_tag = f"h{new_level}"
                updated_content = re.sub(f"<{current_tag}>(.*?)</{current_tag}>", f"<{new_tag}>\\1</{new_tag}>", content, flags=re.DOTALL)

                if content == updated_content:
                    QMessageBox.information(self, "!שים לב", "לא נמצא מה להחליף")
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
        self.setGeometry(100, 100, 550, 250)
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
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
        file_layout.addWidget(browse_button)
        file_layout.addWidget(self.file_path_entry)
        file_layout.addWidget(file_label)
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
        ending_layout.addWidget(self.ending_var, alignment=Qt.AlignRight)  # הוספת תיבת הבחירה קודם
        ending_layout.addWidget(ending_label)  # הוספת התווית אחר כך

        layout.addLayout(ending_layout)

        # הדגשת תחילת קטע
        self.emphasize_var = QCheckBox("הדגש את תחילת הקטעים")
        self.emphasize_var.setStyleSheet("font-size: 20px;")
        layout.addWidget(self.emphasize_var)

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
        if not file_path.lower().endswith('.txt'):
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
                QMessageBox.information(self, "!שים לב", "אין מה לשנות בקובץ זה")

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
        self.setGeometry(100, 100, 550, 300)
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
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
        file_layout.addWidget(browse_button)
        file_layout.addWidget(self.file_entry)
        file_layout.addWidget(file_label)
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
        header_layout.addWidget(self.header_var, alignment=Qt.AlignRight)
        header_layout.addWidget(header_label)
       
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
        if not file_path.lower().endswith('.txt'):
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
                QMessageBox.information(self, "!שים לב", "לא נמצא מה להחליף")
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
            QMessageBox.warning(self, "קלט לא תקין", "בחר קובץ תחילה")
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
        file_layout.addWidget(browse_button)
        file_layout.addWidget(self.file_entry)
        file_layout.addWidget(file_label)
        layout.addLayout(file_layout)

        # סוג ההחלפה
        replace_layout = QHBoxLayout()
        replace_label = QLabel("בחר את סוג ההחלפה:")
        replace_label.setStyleSheet("font-size: 20px;")
        replace_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.replace_type = QComboBox()
        self.replace_type.setFixedWidth(140)
        self.replace_type.setStyleSheet("font-size: 20px;")
        self.replace_type.addItems(["נקודותיים", "ע\"ב"])
        replace_layout.addWidget(self.replace_type, alignment=Qt.AlignRight)
        replace_layout.addWidget(replace_label)
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
            QMessageBox.critical(self, "קלט לא תקין", "לא נבחר קובץ!")
            return
        # בדיקת סוג הקובץ לפי סיומת
        if not file_path.lower().endswith('.txt'):
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
                QMessageBox.information(self, "!שים לב", "לא נמצא מה להחליף")
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
        re_start_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.re_start_entry = QLineEdit()
        self.re_start_entry.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        re_end_label = QLabel("תו/ים בסוף הכותרת:")

        self.re_end_entry = QLineEdit()
        self.re_end_entry.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.gershayim_var = QCheckBox("כולל גרשיים")

        # הוספת הרכיבים
        regex_layout.addWidget(self.gershayim_var)
        regex_layout.addWidget(self.re_end_entry)
        regex_layout.addWidget(re_end_label)
        regex_layout.addWidget(self.re_start_entry)
        regex_layout.addWidget(re_start_label)
        layout.addLayout(regex_layout)

        # יצירת QTextEdit והגדרותיהם
        self.unmatched_regex_text = QTextEdit()
        self.unmatched_regex_text.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.unmatched_regex_text.setReadOnly(True)

        self.unmatched_tags_text = QTextEdit()
        self.unmatched_tags_text.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.unmatched_tags_text.setReadOnly(True)

        # עטיפת כל ווידג'ט במכולה עם תווית מעליו
        regex_container = create_labeled_widget(
            "פירוט הכותרות שיש בהן תווים מיותרים (חוץ ממה שנכתב בתיבות הבחירה למעלה)\nאם יש רווח לפני או אחרי הכותרת, זה גם יוצג כשגיאה",
            self.unmatched_regex_text
        )
        tags_container = create_labeled_widget("פירוט הכותרות שאינן לפי הסדר", self.unmatched_tags_text)

        # הוספת המכולות ל־QSplitter אנכי
        v_splitter = QSplitter(Qt.Vertical)
        v_splitter.setHandleWidth(10)  # עובי handle לפי בחירתך
        v_splitter.addWidget(regex_container)
        v_splitter.addWidget(tags_container)

        # הוספת ה־splitter ל-layout הראשי
        layout.addWidget(v_splitter)

        self.setLayout(layout)

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

        unmatched_regex, unmatched_tags = self.process_html(html_content, re_start, re_end, gershayim)
        self.unmatched_regex_text.setPlainText("\n".join(unmatched_regex))
        self.unmatched_tags_text.setPlainText("\n".join(unmatched_tags))

    def process_html(self, html_content, re_start, re_end, gershayim):
        """
        לוגיקת העיבוד המקורית. כמו בסקריפט הראשוני, רק בלי הדיאלוגים של בחירת קובץ.
        """
        soup = BeautifulSoup(html_content, 'html.parser')

        # קומפילציה של תבנית Regex לפי קלט המשתמש
        if re_start and re_end:
            pattern = re.compile(f"^{re_start}.+[{re_end}]$")
        elif re_start:
            pattern = re.compile(f"^{re_start}.+['א-ת]$")
        elif re_end:
            pattern = re.compile(f"^[א-ת].+[{re_end}]$")
        else:
            pattern = re.compile(r"^[א-ת].+[א-ת']$")

        unmatched_regex = []
        unmatched_tags = []

        # נעבור על תגי כותרות h2 עד h6
        for i in range(2, 7):
            tags = soup.find_all(f"h{i}")

            # בדיקה אם נמצאו תגים
            if not tags:
                unmatched_tags.append(f"מידע: אין בקובץ כותרות ברמה {i}")
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
                if not re.match(pattern, current_tag):
                    unmatched_regex.append(current_tag)

                # בדיקה עבור תנאי גרשיים
                if gershayim:
                    if gematriapy.to_number(current_heading) <= 9:
                        if "'" not in current_heading:
                            unmatched_tags.append(current_heading)
                    else:
                        if '"' not in current_heading:
                            unmatched_tags.append(current_heading)
                else:
                    if "'" in current_heading or '"' in current_heading:
                        unmatched_tags.append(current_heading)

                # בדיקה אם הכותרות הן ברצף
                if not gematriapy.to_number(current_heading) + 1 == gematriapy.to_number(next_heading):
                    unmatched_tags.append(f"כותרת נוכחית - {current_tag}, כותרת הבאה - {next_tag}")

            # עיבוד התג האחרון
            last_tag = tags[-1].string or ""
            if last_tag and not re.match(pattern, last_tag):
                unmatched_regex.append(last_tag)

            last_heading_parts = last_tag.split()
            if len(last_heading_parts) > 1:
                last_heading = last_heading_parts[1]
            else:
                last_heading = last_tag

            if gershayim:
                if gematriapy.to_number(last_heading) <= 9:
                    if "'" not in last_heading:
                        unmatched_tags.append(last_heading)
                else:
                    if '"' not in last_heading:
                        unmatched_tags.append(last_heading)
            else:
                if "'" in last_heading or '"' in last_heading:
                    unmatched_tags.append(last_heading)

        return unmatched_regex, unmatched_tags

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
        self.opening_without_closing.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.opening_without_closing.setReadOnly(True)

        self.closing_without_opening = QTextEdit()
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
        """
        פונקציה זו תחליף את select_file מהסקריפט המקורי.
        תקבל נתיב קובץ ותבצע את כל הבדיקות.
        """
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
            # מציאת כל התגים הפותחים והסוגרים
            tags_in_line = re.findall(r'<(/?\w+)>', line)
            stack = []

            for tag in tags_in_line:
                if not tag.startswith('/'):  # תג פותח
                    stack.append(tag)
                else:  # תג סוגר
                    if stack and stack[-1] == tag[1:]:  # תג תואם במחסנית
                        stack.pop()
                    else:  # תג סוגר בלי פתיחה תואמת
                        closing_without_opening_list.append(
                            f"שורה {line_number}: </{tag[1:]}> || {line.strip()}"
                        )

            # לאחר מעבר על כל התגים בשורה, כל מה שנשאר במחסנית הוא תגים פותחים ללא סגירה
            for unclosed_tag in stack:
                opening_without_closing_list.append(
                    f"שורה {line_number}: <{unclosed_tag}> || {line.strip()}"
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
                        heading_errors_list.append(f"שורה {line_number}: {line.strip()}")

        # הצגת תוצאות
        if opening_without_closing_list:
            self.opening_without_closing.setPlainText("\n".join(opening_without_closing_list))
        else:
            self.opening_without_closing.setPlainText("לא נמצאו שגיאות")

        if closing_without_opening_list:
            self.closing_without_opening.setPlainText("\n".join(closing_without_opening_list))
        else:
            self.closing_without_opening.setPlainText("לא נמצאו שגיאות")

        if heading_errors_list:
            self.heading_errors.setPlainText("\n".join(heading_errors_list))
        else:
            self.heading_errors.setPlainText("לא נמצאו שגיאות")

# ------------------ חלון משולב שמאחד את שתי המחלקות ------------------ #
class CheckHeadingErrorsOriginal(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("בודק כותרות + בודק תגים ביחד")
        self.setWindowIcon(self.get_app_icon())

        # שני ה־Widgets שלנו
        self.check_headings_widget = בדיקת_שגיאות_בכותרות()
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

        top_layout.addWidget(self.browse_button)
        top_layout.addWidget(self.file_path_edit)
        top_layout.addWidget(self.file_path_label)

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
        splitter.addWidget(html_container)

        html_container.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.check_headings_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        html_container.setMaximumHeight(16777215)
        self.check_headings_widget.setMaximumHeight(16777215)
        self.check_headings_widget.resize(1000, 400)
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

        main_layout.addWidget(splitter, 1)

        self.setLayout(main_layout)
        self.resize(1250, 700)  # גודל התחלתי

    def process_file(self, file_path):
        if not file_path:
            QMessageBox.critical(self, "קלט לא תקין", "לא נבחר קובץ!")
            return
        # בדיקת סוג הקובץ לפי סיומת
        if not file_path.lower().endswith('.txt'):
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
        count = content.count("ציור בספר")
        if count > 0:
            text = (f'שים לב! יש בספר {count} ציורים.\n'
                    'חפש בתוך הספר את המילים "ציור בספר",\n'
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
        re_start_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.re_start_entry = QLineEdit()
        self.re_start_entry.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        re_end_label = QLabel("תו/ים בסוף הכותרת:")

        self.re_end_entry = QLineEdit()
        self.re_end_entry.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.re_end_entry.setText('. :')
        self.gershayim_var = QCheckBox("כולל גרשיים")

        # הוספת הרכיבים
        regex_layout.addWidget(self.gershayim_var)
        regex_layout.addWidget(self.re_end_entry)
        regex_layout.addWidget(re_end_label)
        regex_layout.addWidget(self.re_start_entry)
        regex_layout.addWidget(re_start_label)
        layout.addLayout(regex_layout)

        # יצירת QTextEdit והגדרותיהם
        self.unmatched_regex_text = QTextEdit()
        self.unmatched_regex_text.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.unmatched_regex_text.setReadOnly(True)

        self.unmatched_tags_text = QTextEdit()
        self.unmatched_tags_text.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.unmatched_tags_text.setReadOnly(True)

        # עטיפת כל ווידג'ט במכולה עם תווית מעליו
        regex_container = create_labeled_widget(
            "פירוט הכותרות שיש בהן תווים מיותרים (חוץ ממה שנכתב בתיבות הבחירה למעלה)\nאם יש רווח לפני או אחרי הכותרת, זה גם יוצג כשגיאה",
            self.unmatched_regex_text
        )
        tags_container = create_labeled_widget("פירוט הכותרות שאינן לפי הסדר\nהתוכנה מדלגת בבדיקה בכל פעם על כותרת אחת, בגלל הכותרות הכפולות לעמוד ב", self.unmatched_tags_text)

        # הוספת המכולות ל־QSplitter אנכי
        v_splitter = QSplitter(Qt.Vertical)
        v_splitter.setHandleWidth(10)  # עובי handle לפי בחירתך
        v_splitter.addWidget(regex_container)
        v_splitter.addWidget(tags_container)

        # הוספת ה־splitter ל-layout הראשי
        layout.addWidget(v_splitter)

        self.setLayout(layout)

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

        unmatched_regex, unmatched_tags = self.process_html(html_content, re_start, re_end, gershayim)
        self.unmatched_regex_text.setPlainText("\n".join(unmatched_regex))
        self.unmatched_tags_text.setPlainText("\n".join(unmatched_tags))

    def process_html(self, html_content, re_start, re_end, gershayim):
        """
        לוגיקת העיבוד המקורית. כמו בסקריפט הראשוני, רק בלי הדיאלוגים של בחירת קובץ.
        """
        soup = BeautifulSoup(html_content, 'html.parser')

        # קומפילציה של תבנית Regex לפי קלט המשתמש
        if re_start and re_end:
            pattern = re.compile(f"^{re_start}.+[{re_end}]$")
        elif re_start:
            pattern = re.compile(f"^{re_start}.+['א-ת]$")
        elif re_end:
            pattern = re.compile(f"^[א-ת].+[{re_end}]$")
        else:
            pattern = re.compile(r"^[א-ת].+[א-ת']$")

        unmatched_regex = []
        unmatched_tags = []

        # נעבור על תגי כותרות h2 עד h6
        for i in range(2, 7):
            tags = soup.find_all(f"h{i}")

            # בדיקה אם נמצאו תגים
            if not tags:
                unmatched_tags.append(f"מידע: אין בקובץ כותרות ברמה {i}")
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
                if not re.match(pattern, current_tag):
                    unmatched_regex.append(current_tag)

                # בדיקה עבור תנאי גרשיים
                if gershayim:
                    if gematriapy.to_number(current_heading) <= 9:
                        if "'" not in current_heading:
                            unmatched_tags.append(current_heading)
                    else:
                        if '"' not in current_heading:
                            unmatched_tags.append(current_heading)
                else:
                    if "'" in current_heading or '"' in current_heading:
                        unmatched_tags.append(current_heading)

                # בדיקה אם הכותרות הן ברצף
                if not gematriapy.to_number(current_heading) + 1 == gematriapy.to_number(next_heading):
                    unmatched_tags.append(f"כותרת נוכחית - {current_tag}, כותרת הבאה - {next_tag}")

            # עיבוד התג האחרון
            last_tages = (tags[-2].string or "", tags[-1].string or "")
            for last_tag in last_tages:
                if last_tag and not re.match(pattern, last_tag):
                    unmatched_regex.append(last_tag)

            last_heading_parts = last_tag.split()
            if len(last_heading_parts) > 1:
                last_heading = last_heading_parts[1]
            else:
                last_heading = last_tag

            if gershayim:
                if gematriapy.to_number(last_heading) <= 9:
                    if "'" not in last_heading:
                        unmatched_tags.append(last_heading)
                else:
                    if '"' not in last_heading:
                        unmatched_tags.append(last_heading)
            else:
                if "'" in last_heading or '"' in last_heading:
                    unmatched_tags.append(last_heading)

        return unmatched_regex, unmatched_tags

# ------------------ מחלקה שנייה: בדיקת שגיאות בעיצוב (תגים וכו') ------------------ #
class בדיקת_שגיאות_בתגים_לשס(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("בודק שגיאות בעיצוב")
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()


        # יצירת תיבות טקסט והגדרותיהם
        self.opening_without_closing = QTextEdit()
        self.opening_without_closing.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.opening_without_closing.setReadOnly(True)

        self.closing_without_opening = QTextEdit()
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
        """
        פונקציה זו תחליף את select_file מהסקריפט המקורי.
        תקבל נתיב קובץ ותבצע את כל הבדיקות.
        """
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
            # מציאת כל התגים הפותחים והסוגרים
            tags_in_line = re.findall(r'<(/?\w+)>', line)
            stack = []

            for tag in tags_in_line:
                if not tag.startswith('/'):  # תג פותח
                    stack.append(tag)
                else:  # תג סוגר
                    if stack and stack[-1] == tag[1:]:  # תג תואם במחסנית
                        stack.pop()
                    else:  # תג סוגר בלי פתיחה תואמת
                        closing_without_opening_list.append(
                            f"שורה {line_number}: </{tag[1:]}> || {line.strip()}"
                        )

            # לאחר מעבר על כל התגים בשורה, כל מה שנשאר במחסנית הוא תגים פותחים ללא סגירה
            for unclosed_tag in stack:
                opening_without_closing_list.append(
                    f"שורה {line_number}: <{unclosed_tag}> || {line.strip()}"
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
                        heading_errors_list.append(f"שורה {line_number}: {line.strip()}")

        # הצגת תוצאות
        if opening_without_closing_list:
            self.opening_without_closing.setPlainText("\n".join(opening_without_closing_list))
        else:
            self.opening_without_closing.setPlainText("לא נמצאו שגיאות")

        if closing_without_opening_list:
            self.closing_without_opening.setPlainText("\n".join(closing_without_opening_list))
        else:
            self.closing_without_opening.setPlainText("לא נמצאו שגיאות")

        if heading_errors_list:
            self.heading_errors.setPlainText("\n".join(heading_errors_list))
        else:
            self.heading_errors.setPlainText("לא נמצאו שגיאות")

# ------------------ חלון משולב שמאחד את שתי המחלקות ------------------ #
class CheckHeadingErrorsCustom(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("בודק כותרות + בודק תגים ביחד")
        self.setWindowIcon(self.get_app_icon())

        # שני ה־Widgets שלנו
        self.check_headings_widget = בדיקת_שגיאות_בכותרות_לשס()
        self.html_tag_checker_widget = בדיקת_שגיאות_בתגים_לשס()
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

        top_layout.addWidget(self.browse_button)
        top_layout.addWidget(self.file_path_edit)
        top_layout.addWidget(self.file_path_label)

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
        splitter.addWidget(html_container)

        html_container.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.check_headings_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        html_container.setMaximumHeight(16777215)
        self.check_headings_widget.setMaximumHeight(16777215)
        self.check_headings_widget.resize(1000, 400)
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

        main_layout.addWidget(splitter, 1)

        self.setLayout(main_layout)
        self.resize(1250, 700)  # גודל התחלתי

    def process_file(self, file_path):
        if not file_path:
            QMessageBox.critical(self, "קלט לא תקין", "לא נבחר קובץ!")
            return
        # בדיקת סוג הקובץ לפי סיומת
        if not file_path.lower().endswith('.txt'):
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
        count = content.count("ציור בספר")
        if count > 0:
            text = (f'שים לב! יש בספר {count} ציורים.\n'
                    'חפש בתוך הספר את המילים "ציור בספר",\n'
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
        self.setGeometry(100, 100, 500, 500)

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

    def initUI(self):
        layout = QVBoxLayout()
        
        filePathLayout = QHBoxLayout()
        self.filePath = QLineEdit()
        fileLabel = QLabel("נתיב קובץ:")
        filePathLayout.addWidget(fileLabel)
        filePathLayout.addWidget(self.filePath)

        layout.addLayout(filePathLayout)
        
        self.loadBtn = QPushButton("טען קובץ")
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
            "remove_spaces_after": QCheckBox("\u202Bמחיקת רווחים אחרי - [ ("),
            "remove_spaces_around_newlines": QCheckBox("מחיקת רווחים לפני ואחרי אנטר"),
            "replace_double_quotes": QCheckBox("החלפת 2 גרשים בודדים בגרשיים"),
            "normalize_quotes": QCheckBox("המרת גרשיים מוזרים לגרשיים רגילים"),
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

    def cleanText(self, filePath):
        filePath = self.filePath.text()
 
        if not filePath:
            QMessageBox.critical(self, "קלט לא תקין", "לא נבחר קובץ!")
            return
        
        # בדיקת סוג הקובץ לפי סיומת
        if not self.filePath.text().endswith('.txt'):
            QMessageBox.critical(self, "קלט לא תקין", "סוג הקובץ אינו נתמך\nבחר קובץ טקסט [בסיומת TXT.]")
            return   
        
        try:
            with open(self.filePath.text(), 'r', encoding='utf-8') as file:
                text = file.read()
            
            self.originalText = text
            
            if self.checkBoxes["remove_empty_lines"].isChecked():
                text = re.sub(r'\n\s*\n', '\n', text)
            if self.checkBoxes["remove_double_spaces"].isChecked():
                text = re.sub(r' +', ' ', text)
            if self.checkBoxes["remove_spaces_before"].isChecked():
                text = re.sub(r'\s+([\)\],.:])', r'\1', text)
            if self.checkBoxes["remove_spaces_after"].isChecked():
                text = re.sub(r'([\[\(])\s+', r'\1', text)
            if self.checkBoxes["remove_spaces_around_newlines"].isChecked():
                text = re.sub(r'\s*\n\s*', '\n', text)
            if self.checkBoxes["replace_double_quotes"].isChecked():
                text = text.replace("''", '"').replace("``", '"').replace("’’", '"')
            if self.checkBoxes["normalize_quotes"].isChecked():
                text = text.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'").replace("„", '"')
            
            text = text.rstrip()  # מחיקת שורה אחרונה אם היא ריקה

            if text == self.originalText:
                QMessageBox.information(self, "שינויי טקסט", "אין מה להחליף בקובץ זה.")
            else:
                with open(self.filePath.text(), 'w', encoding='utf-8') as file:
                    file.write(text)
                QMessageBox.information(self, "שינויי טקסט", "השינויים בוצעו בהצלחה.")

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
            self.filePath.setText(fileName)
    
    def selectAll(self):
        for checkbox in self.checkBoxes.values():
            checkbox.setChecked(True)
    
    def deselectAll(self):
        for checkbox in self.checkBoxes.values():
            checkbox.setChecked(False)
    
    def undoChanges(self):
        if self.filePath.text() and self.originalText:
            with open(self.filePath.text(), 'w', encoding='utf-8') as file:
                file.write(self.originalText)

    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)

# ==========================================
# Script 12: נקודותיים ורווח
# ==========================================

class ReplaceColonsAndSpaces(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("נקודותיים ורווח")
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.setGeometry(100, 100, 550, 150)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        label = QLabel("תוכנה זו מחליפה את התווים - נקודותיים ורווח, בנקודותיים ואנטר\nתוכנה זו כבר לא אקטואלית למי שמשתמש בגירסה 4.4 ואילך של ספרי דיקטה")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 20px;")
        layout.addWidget(label)

        # נתיב קובץ
        file_layout = QHBoxLayout()
        browse_button = QPushButton("עיון")
        browse_button.setStyleSheet("font-size: 20px;")
        browse_button.clicked.connect(self.select_file)
        self.file_path_entry = QLineEdit()
        file_label = QLabel("נתיב קובץ:")
        file_label.setStyleSheet("font-size: 20px;")
        file_layout.addWidget(browse_button)
        file_layout.addWidget(self.file_path_entry)
        file_layout.addWidget(file_label)
        layout.addLayout(file_layout)

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

    def run_processing(self):
        file_path = self.file_path_entry.text()
        if not file_path:
            QMessageBox.warning(self, "קלט לא תקין", "לא נבחר קובץ!")
            return
        # בדיקת סוג הקובץ לפי סיומת
        if not file_path.lower().endswith('.txt'):
            QMessageBox.critical(self, "קלט לא תקין", "סוג הקובץ אינו נתמך\nבחר קובץ טקסט [בסיומת TXT.]")
            return      

        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()

            # שלב 1: החלפת רצפים של רווחים באנטר בלבד
            new_content = re.sub(r' {1,5}\n', '\n', content)

            # שלב 2: החלפת נקודותיים ורווח בנקודותיים ואנטר
            new_content = re.sub(r':\s', ':\n', new_content)

            if content == new_content:
                QMessageBox.information(self, "!שים לב", "לא נמצא מה להחליף")
            else:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(new_content)
                QMessageBox.information(self, "!מזל טוב", "ההחלפה הושלמה בהצלחה!")
        
        except FileNotFoundError:
            QMessageBox.critical(self, "קלט לא תקין", "הקובץ לא נמצא")
            return
        except UnicodeDecodeError:
            QMessageBox.critical(self, "קלט לא תקין", "קידוד הקובץ אינו נתמך. יש להשתמש בקידוד UTF-8.")
            return   
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"ארעה שגיאה במהלך העיבוד: {str(e)}")

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
       
        # הגדרת החלון
        self.setWindowTitle("עריכת ספרי דיקטה עבור אוצריא")
        self.setLayoutDirection(Qt.RightToLeft)
        self.setWindowIcon(self.load_icon_from_base64(icon_base64))
        self.init_ui()

        # הגדרת האייקון לשורת המשימות
        if sys.platform == 'win32':
            QtWin.setCurrentProcessExplicitAppUserModelID(myappid)

    def init_ui(self):
        layout = QVBoxLayout()

        label = QLabel("בחר את התוכנה שברצונך להפעיל")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 27px;")
        layout.addWidget(label)

        grid_layout = QGridLayout()

        # רשימת כפתורים עם שמות הפונקציות
        button_info = [
            ("1\n\nיצירת כותרות\nלאוצריא\nהתוכנה הראשית", self.open_create_headers_otzria),
            ("2\n\nיצירת כותרות\nלאותיות בודדות\n", self.open_create_single_letter_headers),
            ("3\n\nהוספת\nמספר עמוד\nבכותרת הדף", self.open_add_page_number_to_heading),
            ("4\n\nשינוי רמת כותרת\n\n", self.open_change_heading_level),
            ("5\n\nהדגשת\nמילה ראשונה\nוניקוד בסוף קטע", self.open_emphasize_and_punctuate),
            ("6\n\nיצירת כותרות\nלעמוד ב\n", self.open_create_page_b_headers),
            ("7\n\nהחלפת כותרות\nלעמוד ב\n", self.open_replace_page_b_headers),
            ("8\n\nבדיקת שגיאות\n\n", self.open_check_heading_errors_original),
            ("9\n\nבדיקת שגיאות\nלספרים על השס\n", self.open_check_heading_errors_custom),
            ("10\n\nהמרת תמונה\nלטקסט\n", self.open_Image_To_Html_App),
            ("11\n\nתיקון\nשגיאות נפוצות\n", self.open_Text_Cleaner_App),
            ("12\n\nנקודותיים ורווח\n\n", self.open_replace_colons_and_spaces),
        ]
        
        buttons = []
        for i, (text, func) in enumerate(button_info):
            button = QPushButton(text)
            button.setFixedSize(170, 150)  # הגדרת רוחב וגובה שווים (ריבוע)
            button.setStyleSheet('font-size: 20px;')
            button.clicked.connect(func)  # קישור כל כפתור לפונקציה המתאימה

            # הוספת עיצוב של שוליים מעוגלים לכל כפתור
            button.setStyleSheet("""
                QPushButton {
                    border-radius: 30px;
                    padding: 10px;
                    margin: 5;
                    background-color: #eaeaea;
                    color: black;
                    font-weight: bold;
                    font-family: "Segoe UI", Arial;
                    font-size: 8.5pt;
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
        grid_layout.addWidget(buttons[3], 1, 0)  # שורה 2, טור 1
        grid_layout.addWidget(buttons[4], 1, 1)  # שורה 2, טור 2
        grid_layout.addWidget(buttons[5], 1, 2)  # שורה 2, טור 3
        grid_layout.addWidget(buttons[6], 2, 0)  # שורה 3, טור 1
        grid_layout.addWidget(buttons[7], 2, 1)  # שורה 3, טור 2
        grid_layout.addWidget(buttons[8], 2, 2)  # שורה 3, טור 3
        grid_layout.addWidget(buttons[9], 3, 0)  # שורה 4, טור 1
        grid_layout.addWidget(buttons[10], 3, 1)  # שורה 4, טור 2
        grid_layout.addWidget(buttons[11], 3, 2)  # שורה 4, טור 3

        # יצירת Layout מסוג VBox עבור כפתור "אודות התוכנה"
        main_layout = QVBoxLayout()

        # הוספת כפתור "אודות התוכנה"
        about_button = QPushButton("i")
        about_button.setStyleSheet("font-weight: bold; font-size: 12pt;")
        about_button.setCursor(QCursor(Qt.PointingHandCursor))
        about_button.clicked.connect(self.open_about_dialog)
        about_button.setFixedSize(40, 40)
        main_layout.addLayout(grid_layout)  # הוספת ה-Grid Layout לתוך ה-VBox
        main_layout.addWidget(about_button)  # הוספת הכפתור לתחתית

        # הגדרת ה- Layout של החלון
        self.setLayout(main_layout)

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

    def open_replace_colons_and_spaces(self):
        self.replace_colons_and_spaces_window = ReplaceColonsAndSpaces()
        self.replace_colons_and_spaces_window.show()
   
    # פונקציה לטעינת אייקון ממחרוזת Base64
    def load_icon_from_base64(self, base64_string):
        pixmap = QPixmap()
        pixmap.loadFromData(base64.b64decode(base64_string))
        return QIcon(pixmap)

class AboutDialog(QDialog):
    """חלון 'אודות'"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("אודות התוכנה")
        self.setLayoutDirection(Qt.RightToLeft)

        layout = QVBoxLayout()

        title_label = QLabel("עריכת ספרי דיקטה עבור 'אוצריא'")
        title_label.setStyleSheet("font-size: 14pt; font-weight: bold;")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        version_label = QLabel("גירסה: v3.2")
        version_label.setStyleSheet("font-size: 10pt;")
        version_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(version_label)

        date_label = QLabel("תאריך: כט שבט תשפה")
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
        drive_label = QLabel('או בדרייב: <a href="http://did.li/otzaria-">כאן</a> או <a href="https://drive.google.com/open?id=1KEKudpCJUiK6Y0Eg44PD6cmbRsee1nRO&usp=drive_fs">כאן</a>')
        drive_label.setStyleSheet("font-size: 10pt;")
        drive_label.setOpenExternalLinks(True)
        drive_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(drive_label)

        info_label = QLabel("אפשר לבקש את התוכנה\nוכן להירשם לקבלת עדכון במייל כשיוצא עדכון לתוכנות אלו\nוכן לקבל תמיכה וסיוע בכל הקשור לתוכנה זו ולתוכנת 'אוצריא'\nבמייל הבא:")
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
# Main Application
# ==========================================
def main():
    if sys.platform == 'win32':
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    main_menu = MainMenu()
    main_menu.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
