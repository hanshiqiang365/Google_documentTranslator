#author: hanshiqiang365 （微信公众号）
import os,sys,time
import traceback,subprocess
import googletrans

from docx import Document
from googletrans import Translator

from PyQt5.QtCore import QThread, QSize
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QIcon

g_log = None
g_trans = Translator(service_urls=['translate.google.cn',])

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def translate_buff(buff_para, buff_text, src, dest):
    joiner = '\n'
    tt = joiner.join(buff_text)
    msg = '\t正在翻译：共有 {} 个字'.format(len(tt))
    print(msg)
    if g_log:
        g_log.show.emit(msg)
    try:
        tr = g_trans.translate(tt, dest=dest, src=src)
    except:
        traceback.print_exc()
        msg = '\t<b>Google翻译出现异常，请稍后再试</b>'
        print(msg)
        if g_log:
            g_log.show.emit(msg)
        return
    print(tr.text)
    tr = tr.text.split(joiner)
    print(f'buff_para:{len(buff_para)}, buff_text:{len(buff_text)}, translated para:{len(tr)}')
    for i, t in enumerate(tr):
        para = buff_para[i]
        para.text += '\n' + t

def translate_docx(fn, src, dest):
    doc = Document(fn)
    buff_para = []
    buff_text = []
    buff_len = 0
    max_len = 4900
    for para in doc.paragraphs:
        text = para.text.replace('\n', '').strip()
        if not text: continue
        text_len = len(text.encode('utf8'))
        if buff_len + text_len < max_len:
            buff_para.append(para)
            buff_text.append(text)
            buff_len += text_len
            continue
        translate_buff(buff_para, buff_text, src, dest)
        msg = '休眠 10 秒钟再进行下一次翻译'
        print(msg)
        if g_log:
            g_log.show.emit(msg)
        time.sleep(10)
        buff_para = [para]
        buff_text = [text]
        buff_len = text_len
    if buff_para:
        translate_buff(buff_para, buff_text, src, dest)

    # save
    n_dir = os.path.dirname(fn)
    n_file = os.path.basename(fn)
    to_save = os.path.join(n_dir, 'translated-'+n_file)
    doc.save(to_save)
    return to_save

def translate(fn, src, dest):
    post = fn.split('.')[-1].lower()
    if post == 'docx':
        return translate_docx(fn, src, dest)
    msg = '不支持的文档格式: {}'.format(post)
    print(msg)
    if g_log:
        g_log.show.emit(msg)


class LogHandler(QtCore.QObject):
    show = QtCore.pyqtSignal(str)

class TranslateTask(QThread):
    done = QtCore.pyqtSignal(str)

    def set_attr(self, src, dst, fn):
        self.src = src
        self.dst = dst
        self.fn = fn

    def run(self):
        to_save = translate(self.fn, self.src, self.dst) ####hann
        msg = '翻译成功，保存为：<b>{}</b>'.format(to_save)
        self.done.emit(msg)


class Window(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(Window, self).__init__(parent)

        languages = googletrans.LANGUAGES.copy()
        languages['zh-cn'] = '中文(简体)'
        languages['zh-tw'] = '中文(繁体)'
        self.langcodes = dict(map(reversed, languages.items()))
        self.langlist = [k.capitalize() for k in self.langcodes.keys()]

        self.browseButton = self.createButton("&浏览...", self.browse)
        self.transButton = self.createButton("&翻译", self.translate)

        self.lang_srcComboBox = self.createComboBox()
        srcIndex = self.langlist.index('English')
        self.lang_srcComboBox.setCurrentIndex(srcIndex)
        self.lang_dstComboBox = self.createComboBox()
        dstIndex = self.langlist.index('中文(简体)')
        self.lang_dstComboBox.setCurrentIndex(dstIndex)
        self.fileComboBox = self.createComboBox('file')

        srcLabel = QtWidgets.QLabel("文档语言:")
        dstLabel = QtWidgets.QLabel("目标语言:")
        docLabel = QtWidgets.QLabel("选择文档:")
        self.filesFoundLabel = QtWidgets.QLabel()

        self.logPlainText = QtWidgets.QPlainTextEdit()
        self.logPlainText.setReadOnly(True)


        buttonsLayout = QtWidgets.QHBoxLayout()
        buttonsLayout.addStretch()
        buttonsLayout.addWidget(self.transButton)

        mainLayout = QtWidgets.QGridLayout()
        mainLayout.addWidget(srcLabel, 0, 0)
        mainLayout.addWidget(self.lang_srcComboBox, 0, 1, 1, 2)
        mainLayout.addWidget(dstLabel, 1, 0)
        mainLayout.addWidget(self.lang_dstComboBox, 1, 1, 1, 2)
        mainLayout.addWidget(docLabel, 2, 0)
        mainLayout.addWidget(self.fileComboBox, 2, 1)
        mainLayout.addWidget(self.browseButton, 2, 2)
        mainLayout.addWidget(self.logPlainText, 3, 0, 1, 3)
        mainLayout.addWidget(self.filesFoundLabel, 4, 0)
        mainLayout.addLayout(buttonsLayout, 5, 0, 1, 3)
        self.setLayout(mainLayout)

        app_icon = QIcon()
        icon_path = resource_path('google_translator_icon.png')
        app_icon.addFile(icon_path, QSize(16, 16))
        app_icon.addFile(icon_path, QSize(24, 24))
        app_icon.addFile(icon_path, QSize(32, 32))
        app_icon.addFile(icon_path, QSize(48, 48))
        app_icon.addFile(icon_path, QSize(256, 256))
        self.setWindowIcon(app_icon)

        self.setWindowTitle("Google文档翻译——Developed by hanshiqiang365（微信公众号）")
        self.resize(800, 600)

        self.logger = LogHandler()
        self.logger.show.connect(self.onLog)
        g_log = self.logger
        self.task = TranslateTask()
        self.task.done.connect(self.onLog)

    def init_lang(self):
        self.lang_srcComboBox.currentIndexChanged('english')
        self.lang_dstComboBox.currentIndexChanged('中文(简体)')

    def browse(self):
        sfile, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "选择文件",
            QtCore.QDir.currentPath(),
            "Docx (*.docx)",
        )
        print(sfile)

        if sfile:
            self.logPlainText.clear()
            msg = '选择了文件: <b>{}</b>'.format(sfile)
            self.logger.show.emit(msg)
            if self.fileComboBox.findText(sfile) == -1:
                self.fileComboBox.addItem(sfile)

            self.fileComboBox.setCurrentIndex(self.fileComboBox.findText(sfile))

    def onLog(self, msg):
        self.logPlainText.appendHtml(msg)

    def translate(self):
        lang_src_select = self.lang_srcComboBox.currentText()
        lang_dst_select = self.lang_dstComboBox.currentText()
        fileName = self.fileComboBox.currentText()
        if not fileName:
            self.logger.show.emit('请先选择要翻译的文档')
            return
        lang_src = self.langcodes.get(lang_src_select.lower())
        lang_dst = self.langcodes.get(lang_dst_select.lower())
        if not lang_src:
            self.logger.show.emit('无效的源语言:{}'.format(lang_src_select))
            return
        if not lang_dst:
            self.logger.show.emit('无效的目标语言:{}'.format(lang_dst_select))
            return
        print(lang_src, lang_dst, fileName)
        self.logger.show.emit('开始翻译文档：{}'.format(fileName))
        msg = '从 <b>{}</b> 到 <b>{}</b>'.format(lang_src_select, lang_dst_select)
        self.logger.show.emit(msg)
        self.task.set_attr(lang_src, lang_dst, fileName)
        self.task.start()

    def createButton(self, text, member):
        button = QtWidgets.QPushButton(text)
        button.clicked.connect(member)
        return button

    def createComboBox(self, btype=''):
        comboBox = QtWidgets.QComboBox(self)
        if btype != 'file':
            comboBox.setEditable(True)
            comboBox.addItems(self.langlist)
        comboBox.setSizePolicy(QtWidgets.QSizePolicy.Expanding,
                QtWidgets.QSizePolicy.Preferred)
        return comboBox

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())
