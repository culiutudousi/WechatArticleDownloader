# -*- coding: utf-8 -*-

import sys
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (QWidget, QLabel, QLineEdit, QApplication, QRadioButton, QPushButton, QMessageBox,
                             QButtonGroup, QHBoxLayout, QVBoxLayout, QCheckBox, QSplitter, QFrame, QStyleFactory,
                             QPlainTextEdit, QSizePolicy)
from wechat_article_downloader import (get_article_links_from_wechat_history_list_window_ui, download_from_file_ui,
                                       download_from_file_ui_that_links_to_one_docx, download_one_link_ui)


class Redirect:

    def __init__(self):
        self.content = ''

    def write(self, str):
        self.content += str

    def flush(self):
        self.content = ""


redirect = Redirect()
sys.stdout = redirect


class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.initSettingUI()

        self.setGeometry(800, 100, 800, 800)
        self.setWindowTitle('微信文章下载器')
        self.setWindowIcon(QIcon('./res/download_icon.png'))
        self.setFixedSize(self.width(), self.height())
        self.show()

    def initSettingUI(self):
        self.read_mode_frame = QFrame(self)
        self.read_mode_frame.setFrameShape(QFrame.StyledPanel)
        self.read_mode_widget = ReadModeWidget(self.read_mode_frame)
        self.read_mode_widget.line_edit_enter_press_signal.connect(self.download)

        self.write_mode_frame = QFrame(self)
        self.write_mode_frame.setFrameShape(QFrame.StyledPanel)
        self.write_mode_widget = WriteModeWidget(self.write_mode_frame)

        self.run_frame = QFrame(self)
        self.run_frame.setFrameShape(QFrame.StyledPanel)
        self.run_widget = RunWidget(self.run_frame)
        self.run_widget.go_button.setMinimumWidth(360)
        self.run_widget.go_button.clicked.connect(self.download)
        self.run_widget.finish_signal.connect(lambda b=True: self.read_mode_widget.setLinkInputLineEnabled(b))
        self.run_widget.finish_signal.connect(self.read_mode_widget.clearLink)
        # self.run_widget.control_mouse_and_keyboard_start_signal.connect(lambda b=False: self.setVisible(b))
        # self.run_widget.control_mouse_and_keyboard_end_signal.connect(lambda b=True: self.setVisible(b))

        self.read_mode_frame.setMinimumHeight(350)
        self.run_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self.setting_box = QHBoxLayout()
        self.setting_box.addWidget(self.read_mode_frame)
        self.setting_box.addWidget(self.write_mode_frame)

        self.main_box = QVBoxLayout()
        self.main_box.addLayout(self.setting_box)
        self.main_box.addWidget(self.run_frame)

        self.setLayout(self.main_box)

    def download(self):
        read_mode, get_links_from_wechat_mode, link = self.read_mode_widget.getMode()
        write, writr_mode = self.write_mode_widget.getMode()
        self.read_mode_widget.setLinkInputLineEnabled(False)
        self.run_widget.download(read_mode, get_links_from_wechat_mode, write, writr_mode, link)


class ReadModeWidget(QWidget):

    line_edit_enter_press_signal = pyqtSignal(str)

    def __init__(self, parent=None):
        super(ReadModeWidget, self).__init__(parent)
        self.initUI()

    def initUI(self):
        self.read_mode_title = QLabel('读入设置')
        self.link_input_button = QRadioButton('')
        self.link_input_line = QLineEdit(self)
        self.get_links_from_wechat_button = QRadioButton('读取微信聊天记录')
        self.get_links_from_wechat_overwrite_button = QRadioButton('覆盖links.txt')
        self.get_links_from_wechat_append_button = QRadioButton('添加至links.txt')
        self.get_links_from_file_button = QRadioButton('读links.txt')

        self.link_input_box = QHBoxLayout()
        self.link_input_box.addWidget(self.link_input_button)
        self.link_input_box.addWidget(self.link_input_line)

        self.get_links_from_wechat_mode_core_box = QVBoxLayout()
        self.get_links_from_wechat_mode_core_box.addWidget(self.get_links_from_wechat_overwrite_button)
        self.get_links_from_wechat_mode_core_box.addWidget(self.get_links_from_wechat_append_button)
        self.get_links_from_wechat_mode_surface_box = QHBoxLayout()
        self.get_links_from_wechat_mode_surface_box.addSpacing(25)
        self.get_links_from_wechat_mode_surface_box.addLayout(self.get_links_from_wechat_mode_core_box)
        self.get_links_from_wechat_mode_box = QVBoxLayout()
        self.get_links_from_wechat_mode_box.addWidget(self.get_links_from_wechat_button)
        self.get_links_from_wechat_mode_box.addLayout(self.get_links_from_wechat_mode_surface_box)

        self.read_mode_box = QVBoxLayout()
        self.read_mode_title.setAlignment(Qt.AlignCenter)
        self.read_mode_box.addWidget(self.read_mode_title)
        self.read_mode_box.addLayout(self.link_input_box)
        self.read_mode_box.addLayout(self.get_links_from_wechat_mode_box)
        self.read_mode_box.addWidget(self.get_links_from_file_button)

        self.link = ''
        self.link_input_line.textChanged[str].connect(self.onLinkChanged)
        # self.link_input_line.editingFinished.connect(self.onLinkFinished)
        self.link_input_line.returnPressed.connect(self.onLineEditEnterPressed)

        self.read_mode = 'link'
        self.link_input_button.toggle()
        self.read_mode_button_group = QButtonGroup()
        self.read_mode_button_group.addButton(self.link_input_button)
        self.read_mode_button_group.addButton(self.get_links_from_wechat_button)
        self.read_mode_button_group.addButton(self.get_links_from_file_button)
        self.read_mode_button_group.buttonClicked.connect(self.onReadModeButtonClicked)

        self.get_links_from_wechat_mode = 'overwrite'
        self.get_links_from_wechat_overwrite_button.toggle()
        self.setGetLinksFromWechatModeButtonGroupEnabled(False)
        self.get_links_from_wechat_mode_button_group = QButtonGroup()
        self.get_links_from_wechat_mode_button_group.addButton(self.get_links_from_wechat_overwrite_button)
        self.get_links_from_wechat_mode_button_group.addButton(self.get_links_from_wechat_append_button)
        self.get_links_from_wechat_mode_button_group.buttonClicked.connect(self.onGetLinksFromWechatModeButtonClicked)

        self.setLayout(self.read_mode_box)

    def onLinkChanged(self, text):
        self.link = text

    def onLineEditEnterPressed(self):
        self.line_edit_enter_press_signal.emit(self.link)
        self.clearLink()

    def clearLink(self):
        self.link_input_line.setText('')

    def onReadModeButtonClicked(self):
        if self.read_mode_button_group.checkedButton() is self.link_input_button:
            self.read_mode = 'link'
            self.setLinkInputLineEnabled(True)
            self.setGetLinksFromWechatModeButtonGroupEnabled(False)
        elif self.read_mode_button_group.checkedButton() is self.get_links_from_wechat_button:
            self.read_mode = 'copy'
            self.setLinkInputLineEnabled(False)
            self.setGetLinksFromWechatModeButtonGroupEnabled(True)
        elif self.read_mode_button_group.checkedButton() is self.get_links_from_file_button:
            self.read_mode = 'file'
            self.setLinkInputLineEnabled(False)
            self.setGetLinksFromWechatModeButtonGroupEnabled(False)

    def onGetLinksFromWechatModeButtonClicked(self):
        if self.get_links_from_wechat_mode_button_group.checkedButton() is self.get_links_from_wechat_overwrite_button:
            self.get_links_from_wechat_mode = 'overwrite'
        elif self.get_links_from_wechat_mode_button_group.checkedButton() is self.get_links_from_wechat_append_button:
            self.get_links_from_wechat_mode = 'append'

    def setLinkInputLineEnabled(self, bool):
        self.link_input_line.setEnabled(bool)

    def setGetLinksFromWechatModeButtonGroupEnabled(self, bool):
        self.get_links_from_wechat_overwrite_button.setEnabled(bool)
        self.get_links_from_wechat_append_button.setEnabled(bool)

    def getMode(self):
        return self.read_mode, self.get_links_from_wechat_mode, self.link


class WriteModeWidget(QWidget):

    def __init__(self, parent=None):
        super(WriteModeWidget, self).__init__(parent)
        self.initUI()

    def initUI(self):
        self.writr_mode_title = QLabel('导出设置')
        self.write_button = QCheckBox('导出docx')
        self.write_every_link_a_file = QRadioButton('每个链接对应一个文件')
        self.write_links_to_one_file = QRadioButton('保存为一个文件')

        self.write_mode_core_box = QVBoxLayout()
        self.write_mode_core_box.addWidget(self.write_every_link_a_file)
        self.write_mode_core_box.addWidget(self.write_links_to_one_file)
        self.write_mode_surface_box = QHBoxLayout()
        self.write_mode_surface_box.addSpacing(25)
        self.write_mode_surface_box.addLayout(self.write_mode_core_box)

        self.writr_mode_box = QVBoxLayout()
        self.writr_mode_title.setAlignment(Qt.AlignCenter)
        self.writr_mode_box.addWidget(self.writr_mode_title)
        self.writr_mode_box.addWidget(self.write_button)
        self.writr_mode_box.addLayout(self.write_mode_surface_box)

        self.write = True
        self.write_button.toggle()
        self.writr_mode = 'multi'
        self.write_every_link_a_file.toggle()
        self.write_mode_button_group = QButtonGroup()
        self.write_mode_button_group.addButton(self.write_every_link_a_file)
        self.write_mode_button_group.addButton(self.write_links_to_one_file)

        self.write_button.stateChanged.connect(self.onWriteButtonClicked)
        self.write_mode_button_group.buttonClicked.connect(self.onWriteModeButtonClicked)

        self.setLayout(self.writr_mode_box)

    def onWriteButtonClicked(self, state):
        if state == Qt.Checked:
            self.write = True
            self.setWriteModeButtonGroupEnabled(True)
        else:
            self.write = False
            self.setWriteModeButtonGroupEnabled(False)

    def onWriteModeButtonClicked(self):
        if self.write_mode_button_group.checkedButton() is self.write_every_link_a_file:
            self.writr_mode = 'multi'
        elif self.write_mode_button_group.checkedButton() is self.write_links_to_one_file:
            self.writr_mode = 'one'

    def setWriteModeButtonGroupEnabled(self, bool):
        self.write_every_link_a_file.setEnabled(bool)
        self.write_links_to_one_file.setEnabled(bool)

    def getMode(self):
        return self.write, self.writr_mode


class RunWidget(QWidget):

    finish_signal = pyqtSignal()
    control_mouse_and_keyboard_start_signal = pyqtSignal()
    control_mouse_and_keyboard_end_signal = pyqtSignal()

    def __init__(self, parent=None):
        super(RunWidget, self).__init__(parent)
        self.initUI()

    def initUI(self):
        self.go_button = QPushButton('下载')
        self.go_button.setMinimumHeight(35)
        self.state_show = QPlainTextEdit()
        self.state_show.setReadOnly(True)

        self.run_box = QVBoxLayout(self)
        self.run_box.addWidget(self.go_button)
        self.run_box.addWidget(self.state_show)
        self.run_box.setAlignment(Qt.AlignHCenter | Qt.AlignTop)

        self.state_show.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setLayout(self.run_box)

    def appendText(self, text):
        if text.strip():
            self.state_show.appendPlainText(text.strip())

    def setGoButtonEnabeld(self, bool):
        self.go_button.setEnabled(bool)

    def download(self, read_mode, get_links_from_wechat_mode, write, write_mode, link=''):
        self.dl = Download(read_mode, get_links_from_wechat_mode, write, write_mode, link)
        self.dl.print_signal.connect(self.appendText)
        self.dl.finish_signal.connect(lambda b=True: self.setGoButtonEnabeld(b))
        self.dl.finish_signal.connect(self.finish_signal.emit)
        self.dl.control_mouse_and_keyboard_start_signal.connect(self.control_mouse_and_keyboard_start_signal.emit)
        self.dl.control_mouse_and_keyboard_end_signal.connect(self.control_mouse_and_keyboard_end_signal.emit)
        self.setGoButtonEnabeld(False)
        self.dl.start()


class Download(QThread):

    print_signal = pyqtSignal(str)
    finish_signal = pyqtSignal()
    control_mouse_and_keyboard_start_signal = pyqtSignal()
    control_mouse_and_keyboard_end_signal = pyqtSignal()

    def __init__(self, read_mode, get_links_from_wechat_mode, write, write_mode, link=''):
        super().__init__()
        self.read_mode, self.get_links_from_wechat_mode, self.write, self.write_mode, self.link = \
            read_mode, get_links_from_wechat_mode, write, write_mode, link
        redirect.write = self.print_signal.emit

    def run(self):
        file_name = 'links.txt'
        if self.read_mode == 'link':
            if self.link:
                download_one_link_ui(self.link)
        else:
            if self.read_mode == 'copy':
                self.control_mouse_and_keyboard_start_signal.emit()
                if self.get_links_from_wechat_mode == 'overwrite':
                    get_article_links_from_wechat_history_list_window_ui(file_name, append_or_overwrite='overwrite')
                elif self.get_links_from_wechat_mode == 'append':
                    get_article_links_from_wechat_history_list_window_ui(file_name, append_or_overwrite='append')
                self.control_mouse_and_keyboard_end_signal.emit()
            elif self.read_mode == 'file':
                pass
            if self.write:
                if self.write_mode == 'multi':
                    download_from_file_ui(file_name)
                elif self.write_mode == 'one':
                    download_from_file_ui_that_links_to_one_docx(file_name)
        self.finish_signal.emit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())
