from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QTableWidget, QTableWidgetItem,
    QPushButton, QVBoxLayout, QWidget, QProgressBar, QTextEdit, QCheckBox, QLabel, QHBoxLayout, QLineEdit, QGroupBox, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import os, glob

from replacers.text_replacer import replace_in_text_file
from replacers.word_replacer import replace_in_word, replace_in_word_doc
from replacers.excel_replacer import replace_in_excel, replace_in_excel_xls
from replacers.ppt_replacer import replace_in_ppt, replace_in_ppt_ppt
from replacers.filename_replacer import replace_filename

class ReplaceThread(QThread):
    progress_signal = pyqtSignal(int)
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, files, replacements, options):
        super().__init__()
        self.files = files
        self.replacements = replacements
        self.options = options
        self._pause = False
        self._stop = False

    def run(self):
        total = len(self.files)
        for idx, file in enumerate(self.files, 1):
            if self._stop:
                self.log_signal.emit("已停止处理。")
                break
            while self._pause:
                self.msleep(100)
            ext = os.path.splitext(file)[1].lower()
            try:
                if ext in ('.txt', '.html', '.htm'):
                    replace_in_text_file(file, self.replacements, self.options['fullword'])
                elif ext == '.docx':
                    replace_in_word(
                        file, self.replacements,
                        wildcard=self.options['wildcard'],
                        fullwidth=self.options['fullwidth'],
                        halfwidth=self.options['halfwidth']
                    )
                elif ext == '.doc':
                    replace_in_word_doc(
                        file, self.replacements,
                        wildcard=self.options['wildcard']
                    )
                elif ext == '.xlsx':
                    replace_in_excel(file, self.replacements)
                elif ext == '.xls':
                    replace_in_excel_xls(file, self.replacements)
                elif ext == '.pptx':
                    replace_in_ppt(file, self.replacements)
                elif ext == '.ppt':
                    replace_in_ppt_ppt(file, self.replacements)
                if self.options['filename']:
                    new_file = replace_filename(file, self.replacements)
                    if new_file != file:
                        self.log_signal.emit(f"文件名已改为：{os.path.basename(new_file)}")
                self.log_signal.emit(f"处理完成：{file}")
            except Exception as e:
                self.log_signal.emit(f"处理失败：{file}，原因：{e}")
            self.progress_signal.emit(idx)
        self.finished_signal.emit()

    def pause(self):
        self._pause = True
    def resume(self):
        self._pause = False
    def stop(self):
        self._stop = True

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("批量内容与文件名替换工具")
        self.resize(900, 700)
        self.setStyleSheet("""
            QWidget { font-size: 14px; }
            QPushButton { min-width: 90px; min-height: 32px; font-weight: bold; }
            QProgressBar { height: 22px; }
            QLineEdit, QTextEdit { background: #f8f8f8; }
            QGroupBox { font-size: 15px; font-weight: bold; }
        """)

        # 文件选择区
        file_group = QGroupBox("文件/文件夹选择")
        file_layout = QHBoxLayout()
        self.file_path = QLineEdit()
        self.file_btn = QPushButton("选择")
        self.file_btn.clicked.connect(self.select_files)
        file_layout.addWidget(self.file_path)
        file_layout.addWidget(self.file_btn)
        file_group.setLayout(file_layout)

        # 替换规则区
        rule_group = QGroupBox("替换规则（每行：原内容,新内容）")
        rule_layout = QVBoxLayout()
        self.rule_edit = QTextEdit()
        self.rule_edit.setPlaceholderText("如：foo,bar\nhello,world")
        rule_layout.addWidget(self.rule_edit)
        rule_group.setLayout(rule_layout)

        # 选项区
        opt_group = QGroupBox("选项")
        opt_layout = QHBoxLayout()
        self.fullword_cb = QCheckBox("全字匹配")
        self.filename_cb = QCheckBox("同步替换文件名")
        self.word_wildcard_cb = QCheckBox("Word通配符")
        self.word_fullwidth_cb = QCheckBox("Word转全角")
        self.word_halfwidth_cb = QCheckBox("Word转半角")
        for cb in [self.fullword_cb, self.filename_cb, self.word_wildcard_cb, self.word_fullwidth_cb, self.word_halfwidth_cb]:
            opt_layout.addWidget(cb)
        opt_group.setLayout(opt_layout)

        # 操作按钮区
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始替换")
        self.pause_btn = QPushButton("暂停")
        self.resume_btn = QPushButton("继续")
        self.stop_btn = QPushButton("停止")
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.pause_btn)
        btn_layout.addWidget(self.resume_btn)
        btn_layout.addWidget(self.stop_btn)

        # 进度条
        self.progress = QProgressBar()
        self.progress.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        # 日志区
        log_group = QGroupBox("日志输出")
        log_layout = QVBoxLayout()
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        log_layout.addWidget(self.log_edit)
        log_group.setLayout(log_layout)

        # 主布局
        main_layout = QVBoxLayout()
        main_layout.addWidget(file_group)
        main_layout.addWidget(rule_group)
        main_layout.addWidget(opt_group)
        main_layout.addLayout(btn_layout)
        main_layout.addWidget(self.progress)
        # 这里不再添加任何空白控件，直接添加日志区
        main_layout.addWidget(log_group)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        self.replace_thread = None
        self.start_btn.clicked.connect(self.start_replace)
        self.pause_btn.clicked.connect(self.pause_replace)
        self.resume_btn.clicked.connect(self.resume_replace)
        self.stop_btn.clicked.connect(self.stop_replace)

    def select_files(self):
        dlg = QFileDialog(self)
        dlg.setFileMode(QFileDialog.ExistingFiles)
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setWindowTitle("选择文件或文件夹")
        dlg.setNameFilter("所有文件 (*.*)")
        # 增加文件夹选择按钮
        btns = dlg.findChildren(QPushButton)
        folder_btn = None
        for btn in btns:
            if btn.text() in ("&Open", "打开(&O)", "Open"):
                folder_btn = QPushButton("选择文件夹", dlg)
                btn.parent().layout().addWidget(folder_btn)
                break
        def choose_folder():
            folder = QFileDialog.getExistingDirectory(self, "选择文件夹", "")
            if folder:
                self.file_path.setText(folder)
                self.files = []
                for ext in ('*.txt', '*.html', '*.htm', '*.docx', '*.doc', '*.xlsx', '*.xls', '*.pptx', '*.ppt'):
                    self.files.extend(glob.glob(os.path.join(folder, '**', ext), recursive=True))
                self.log(f"已选文件数：{len(self.files)}")
            dlg.reject()
        if folder_btn:
            folder_btn.clicked.connect(choose_folder)
        if dlg.exec_():
            files = dlg.selectedFiles()
            self.file_path.setText(';'.join(files))
            self.files = list(files)
            self.log(f"已选文件数：{len(self.files)}")

    def parse_rules(self):
        rules = []
        for line in self.rule_edit.toPlainText().splitlines():
            if ',' in line:
                old, new = line.split(',', 1)
                rules.append((old.strip(), new.strip()))
        return rules

    def start_replace(self):
        try:
            self.replacements = self.parse_rules()
            if not hasattr(self, 'files') or not self.files or not self.replacements:
                self.log("请先选择文件并输入替换规则。")
                return
            options = {
                'fullword': self.fullword_cb.isChecked(),
                'filename': self.filename_cb.isChecked(),
                'wildcard': self.word_wildcard_cb.isChecked(),
                'fullwidth': self.word_fullwidth_cb.isChecked(),
                'halfwidth': self.word_halfwidth_cb.isChecked()
            }
            self.progress.setMaximum(len(self.files))
            self.replace_thread = ReplaceThread(self.files, self.replacements, options)
            self.replace_thread.progress_signal.connect(self.progress.setValue)
            self.replace_thread.log_signal.connect(self.log)
            self.replace_thread.finished_signal.connect(lambda: self.log("全部处理完成。"))
            self.replace_thread.start()
        except Exception as e:
            import traceback
            self.log(f"发生异常: {e}\n{traceback.format_exc()}")

    def pause_replace(self):
        if self.replace_thread:
            self.replace_thread.pause()
            self.log("已暂停。")
    def resume_replace(self):
        if self.replace_thread:
            self.replace_thread.resume()
            self.log("已继续。")
    def stop_replace(self):
        if self.replace_thread:
            self.replace_thread.stop()
            self.log("已请求停止。")

    def log(self, msg):
        self.log_edit.append(msg) 