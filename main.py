import os
import shutil
import sys
import traceback
import warnings

from PyQt6.QtCore import Qt, pyqtSignal, QThread
from PyQt6.QtGui import QIcon, QPainter, QPixmap
from PyQt6.QtWidgets import QApplication, QPushButton, QMainWindow, QLabel, QVBoxLayout, QFileDialog, QHBoxLayout, \
    QDialog
from qt_material import apply_stylesheet

from dataTo2700table import dataTo2700table
from fea_json import feature_json_all
from images.UImain import Ui_MainWindow
from PlatformTable import output_template_all
from deviceInfo_And_tupusetting import device_info, tupuSetting_V2, tupuSetting_V3
warnings.filterwarnings('ignore')


class Worker1(QThread):
    finished = pyqtSignal(str)

    def __init__(self, data_all_edit, output_path_1):
        super().__init__()
        self.data_all_edit = data_all_edit
        self.my_deftable = "后台文件/my_def_对应注释.xlsx"
        self.output_path_1 = output_path_1

    def run(self):
        try:
            output_file_True = output_template_all(self.data_all_edit, self.my_deftable, self.output_path_1, False)
            if output_file_True:
                self.finished.emit("文件中设备档案和输入参数中设备不对应，具体参看TXT文件")
            else:
                self.finished.emit(f"平台导入表文件已保存到：\n{self.output_path_1}")
        except PermissionError:
            self.finished.emit("访问权限限制，请关闭相关文件")
        except Exception as e:
            print(traceback.format_exc())
            self.finished.emit("平台导入表输出失败，请检查data_all文件格式")


class Worker2(QThread):
    finished = pyqtSignal(str)

    def __init__(self, data_all_edit, output_path):
        super().__init__()
        self.data_all_edit = data_all_edit
        self.output_path = output_path
        self.my_deftable = "后台文件/my_def_对应注释.xlsx"
        self.ZZ = "中转文件/平台导入表.xlsx"

    def run(self):
        try:
            output_template_all(self.data_all_edit, self.my_deftable, self.ZZ)
            device_info(self.ZZ, self.my_deftable, self.output_path)
            if os.path.exists(self.ZZ):
                os.remove(self.ZZ)
            self.finished.emit(f"device_info文件已保存：\n{self.output_path}")
        except PermissionError:
            self.finished.emit("访问权限限制，请关闭相关文件")
        except Exception as e:
            print(traceback.format_exc())
            self.finished.emit("device_info文件输出失败，请检查data_all文件格式")


class Worker3(QThread):
    finished = pyqtSignal(str)

    def __init__(self, data_all_edit, output_path):
        super().__init__()
        self.data_all_edit = data_all_edit
        self.output_path = output_path
        self.my_deftable = "后台文件/my_def_对应注释.xlsx"
        self.ZZ = "中转文件/平台导入表.xlsx"

    def run(self):
        try:
            output_template_all(self.data_all_edit, self.my_deftable, self.ZZ)
            tupuSetting_V2(self.ZZ, self.output_path)
            self.finished.emit(f"tupuSetting文件已保存：\n{self.output_path}")
        except PermissionError:
            self.finished.emit("访问权限限制，请关闭相关文件")
        except Exception as e:
            print(traceback.format_exc())
            self.finished.emit("tupuSetting(V2.0)文件输出失败，请检查data_all文件格式")


class Worker4(QThread):
    finished = pyqtSignal(str)

    def __init__(self, inputFile, outputFile):
        super().__init__()
        self.inputFile = inputFile
        self.outputFile = outputFile
        self.zz = "中转文件/DW-导入表.xlsx"

    def run(self):
        try:
            dataTo2700table(self.inputFile, self.zz)
            if_err = feature_json_all(self.zz, self.outputFile)
            if if_err:
                print(if_err)
                self.finished.emit("Json文件输出失败")
            else:
                self.finished.emit(f"Json文件保存到：\n{self.outputFile}")
        except PermissionError:
            self.finished.emit("访问权限限制，请关闭相关文件")
        except Exception as e:
            print(traceback.format_exc())
            self.finished.emit("Json文件输出失败，请检查DW-导入表文件格式")


class Worker5(QThread):
    finished = pyqtSignal(str)

    def __init__(self, inputFile, outputFile):
        super().__init__()
        self.inputFile = inputFile
        self.outputFile = outputFile

    def run(self):
        try:
            dataTo2700table(self.inputFile, self.outputFile)
            self.finished.emit(f"DW-导入表保存到：\n{self.outputFile}")
        except PermissionError:
            self.finished.emit("访问权限限制，请关闭相关文件")
        except KeyError as e:
            self.finished.emit(f"文件缺少列：{e}")
        except Exception as e:
            print(traceback.format_exc())
            self.finished.emit("DW-导入表输出失败，请检查DW-导入表文件格式")


class Worker6(QThread):
    finished = pyqtSignal(str)

    def __init__(self, data_all_edit, output_path):
        super().__init__()
        self.data_all_edit = data_all_edit
        self.output_path = output_path
        self.my_deftable = "后台文件/my_def_对应注释.xlsx"
        self.ZZ = "中转文件/平台导入表.xlsx"

    def run(self):
        try:
            output_template_all(self.data_all_edit, self.my_deftable, self.ZZ)
            tupuSetting_V3(self.ZZ, self.output_path)
            self.finished.emit(f"tupuSetting文件已保存：\n{self.output_path}")
        except PermissionError:
            self.finished.emit("访问权限限制，请关闭相关文件")
        except Exception as e:
            print(traceback.format_exc())
            self.finished.emit("tupuSetting(V3.0)文件输出失败，请检查data_all文件格式")


class CustomMessageBox(QDialog):
    def __init__(self, message_text, parent=None):
        super().__init__(parent)
        self.style = """
        QDialog{
        border-image: url("images/dia_bg.png");
        }
        QLabel{
        color: whitesmoke;
        font-family: "Microsoft YaHei UI";
        font-size: 30px;
        }
        QPushButton {
        color: white;
        background-color: #005ABA;
        padding: -20px 1px;
        position: center;
        font-size: 25px;
        min-width: 80px;
        min-height: 20px;
        max-width: 80px;
        max-height: 20px;
        }
        QPushButton:hover {
        border: 1px solid #F5F6FA;
        background-color: #004B9B;
        }
        """
        self.ok_button = None
        self.no_button = None
        self.btn_layout = None
        self.resize(350, 250)
        self.initUI(message_text)

    def initUI(self, message_text):
        self.setWindowTitle("提示")
        # 创建布局
        layout = QVBoxLayout()

        # 添加文本标签
        label = QLabel(message_text)
        label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        layout.addStretch(1)
        layout.addWidget(label)
        layout.addStretch(1)

        # 添加自定义按钮
        self.btn_layout = QHBoxLayout()
        self.ok_button = QPushButton("确定")
        self.no_button = QPushButton("取消")
        self.no_button.clicked.connect(self.reject)
        self.ok_button.clicked.connect(self.accept)  # 关闭对话框

        self.btn_layout.addStretch(1)
        self.btn_layout.addWidget(self.no_button)
        self.btn_layout.addStretch(1)
        self.btn_layout.addWidget(self.ok_button)
        self.btn_layout.addStretch(1)
        layout.addLayout(self.btn_layout)
        layout.addStretch(1)

        self.setLayout(layout)
        self.setStyleSheet(self.style)


class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowIcon(QIcon('images/logo.ico'))
        self.setWindowTitle('凯奥斯特征解析工具V3.0（内部版本）')
        self.excel_name = ''
        self.output_path = ''
        self.data_all_edit = ''
        self.sensor_record_edit = ''
        self.platform_import_edit = ''
        self.my_deftable = '后台文件/my_def_对应注释.xlsx'
        self.init_ui()

    def init_ui(self):
        self.top_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.top_label.setProperty('class', 'label')
        self.top_label.setStyleSheet("QLabel{border-image: url(\"%s\");}" % 'images/top.png')
        self.pushButton_1.clicked.connect(self.load_file)
        upload_pix = QLabel()
        pixmap = QPixmap('images/icn_上传.png')
        # Scale the pixmap to a desired size
        desired_width = 100  # Change this to your desired width
        desired_height = 80  # Change this to your desired height
        scaled_pixmap = pixmap.scaled(desired_width, desired_height, Qt.AspectRatioMode.KeepAspectRatio,
                                      Qt.TransformationMode.SmoothTransformation)
        # Set the scaled pixmap to the label
        upload_pix.setPixmap(scaled_pixmap)
        upload_pix.setAlignment(Qt.AlignmentFlag.AlignCenter)
        upload_zi = QLabel("上传data_all")
        upload_zi.setProperty('class', 'btn_label')
        upload_zi.setAlignment(Qt.AlignmentFlag.AlignCenter)
        upload_layout = QVBoxLayout()
        upload_layout.addWidget(upload_pix)
        upload_layout.addWidget(upload_zi)
        self.pushButton_1.setLayout(upload_layout)
        self.pushButton_1.setProperty('class', 'big_button')
        self.pushButton_1.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        self.pushButton_2.clicked.connect(self.download_tmp)
        download_pix = QLabel()
        pixmap = QPixmap('images/icn_download.png')
        # Scale the pixmap to a desired size
        desired_width = 100  # Change this to your desired width
        desired_height = 80  # Change this to your desired height
        scaled_pixmap = pixmap.scaled(desired_width, desired_height, Qt.AspectRatioMode.KeepAspectRatio,
                                      Qt.TransformationMode.SmoothTransformation)
        # Set the scaled pixmap to the label
        download_pix.setPixmap(scaled_pixmap)
        download_pix.setAlignment(Qt.AlignmentFlag.AlignCenter)
        download_zi = QLabel("下载模板")
        download_zi.setProperty('class', 'btn_label')
        download_zi.setAlignment(Qt.AlignmentFlag.AlignCenter)
        download_layout = QVBoxLayout()
        download_layout.addWidget(download_pix)
        download_layout.addWidget(download_zi)
        self.pushButton_2.setLayout(download_layout)
        self.pushButton_2.setProperty('class', 'big_button')
        self.pushButton_2.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        self.pushButton_3.clicked.connect(self.json_img)
        json_pix = QLabel()
        pixmap = QPixmap('images/icn_download.png')
        # Scale the pixmap to a desired size
        desired_width = 100  # Change this to your desired width
        desired_height = 80  # Change this to your desired height
        scaled_pixmap = pixmap.scaled(desired_width, desired_height, Qt.AspectRatioMode.KeepAspectRatio,
                                      Qt.TransformationMode.SmoothTransformation)
        json_pix.setPixmap(scaled_pixmap)
        json_pix.setAlignment(Qt.AlignmentFlag.AlignCenter)
        json_zi = QLabel("输出\nJson文件")
        json_zi.setProperty('class', 'btn_label')
        json_zi.setAlignment(Qt.AlignmentFlag.AlignCenter)
        json_layout = QVBoxLayout()
        json_layout.addWidget(json_pix)
        json_layout.addWidget(json_zi)
        self.pushButton_3.setLayout(json_layout)
        self.pushButton_3.setProperty('class', 'big_button')
        self.pushButton_3.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        self.pushButton_4.clicked.connect(self.predict_img)
        platformTable_pix = QLabel()
        pixmap = QPixmap('images/icn_download.png')
        # Scale the pixmap to a desired size
        desired_width = 100  # Change this to your desired width
        desired_height = 80  # Change this to your desired height
        scaled_pixmap = pixmap.scaled(desired_width, desired_height, Qt.AspectRatioMode.KeepAspectRatio,
                                      Qt.TransformationMode.SmoothTransformation)
        platformTable_pix.setPixmap(scaled_pixmap)
        platformTable_pix.setAlignment(Qt.AlignmentFlag.AlignCenter)
        platformTable_zi = QLabel("输出\n平台导入表")
        platformTable_zi.setProperty('class', 'btn_label')
        platformTable_zi.setAlignment(Qt.AlignmentFlag.AlignCenter)
        platformTable_layout = QVBoxLayout()
        platformTable_layout.addWidget(platformTable_pix)
        platformTable_layout.addWidget(platformTable_zi)
        self.pushButton_4.setLayout(platformTable_layout)
        self.pushButton_4.setProperty('class', 'big_button')
        self.pushButton_4.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        self.pushButton_5.clicked.connect(self.device_img)
        deviceInfo_pix = QLabel()
        pixmap = QPixmap('images/icn_download.png')
        # Scale the pixmap to a desired size
        desired_width = 100  # Change this to your desired width
        desired_height = 80  # Change this to your desired height
        scaled_pixmap = pixmap.scaled(desired_width, desired_height, Qt.AspectRatioMode.KeepAspectRatio,
                                      Qt.TransformationMode.SmoothTransformation)
        deviceInfo_pix.setPixmap(scaled_pixmap)
        deviceInfo_pix.setAlignment(Qt.AlignmentFlag.AlignCenter)
        deviceInfo_zi = QLabel("输出\ndeviceInfo")
        deviceInfo_zi.setProperty('class', 'btn_label')
        deviceInfo_zi.setAlignment(Qt.AlignmentFlag.AlignCenter)
        deviceInfo_layout = QVBoxLayout()
        deviceInfo_layout.addWidget(deviceInfo_pix)
        deviceInfo_layout.addWidget(deviceInfo_zi)
        self.pushButton_5.setLayout(deviceInfo_layout)
        self.pushButton_5.setProperty('class', 'big_button')
        self.pushButton_5.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        self.pushButton_6.clicked.connect(self.tupuset_img_V2)
        tupuSetting_pix = QLabel()
        pixmap = QPixmap('images/icn_download.png')
        # Scale the pixmap to a desired size
        desired_width = 100  # Change this to your desired width
        desired_height = 80  # Change this to your desired height
        scaled_pixmap = pixmap.scaled(desired_width, desired_height, Qt.AspectRatioMode.KeepAspectRatio,
                                      Qt.TransformationMode.SmoothTransformation)
        tupuSetting_pix.setPixmap(scaled_pixmap)
        tupuSetting_pix.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tupuSetting_zi = QLabel("输出V2.1\ntupuSet")
        tupuSetting_zi.setProperty('class', 'btn_label')
        tupuSetting_zi.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tupuSetting_layout = QVBoxLayout()
        tupuSetting_layout.addWidget(tupuSetting_pix)
        tupuSetting_layout.addWidget(tupuSetting_zi)
        self.pushButton_6.setLayout(tupuSetting_layout)
        self.pushButton_6.setProperty('class', 'big_button')
        self.pushButton_6.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        self.pushButton_7.clicked.connect(self.dat2700_img)
        dat2700_pix = QLabel()
        pixmap = QPixmap('images/icn_download.png')
        # Scale the pixmap to a desired size
        desired_width = 100  # Change this to your desired width
        desired_height = 80  # Change this to your desired height
        scaled_pixmap = pixmap.scaled(desired_width, desired_height, Qt.AspectRatioMode.KeepAspectRatio,
                                      Qt.TransformationMode.SmoothTransformation)
        dat2700_pix.setPixmap(scaled_pixmap)
        dat2700_pix.setAlignment(Qt.AlignmentFlag.AlignCenter)
        dat2700_zi = QLabel("输出\nDW-导入表")
        dat2700_zi.setProperty('class', 'btn_label')
        dat2700_zi.setAlignment(Qt.AlignmentFlag.AlignCenter)
        dat2700_layout = QVBoxLayout()
        dat2700_layout.addWidget(dat2700_pix)
        dat2700_layout.addWidget(dat2700_zi)
        self.pushButton_7.setLayout(dat2700_layout)
        self.pushButton_7.setProperty('class', 'big_button')
        self.pushButton_7.setFocusPolicy(Qt.FocusPolicy.NoFocus)

        self.pushButton_8.clicked.connect(self.tupuset_img_V3)
        tupuSetting_pix = QLabel()
        pixmap = QPixmap('images/icn_download.png')
        # Scale the pixmap to a desired size
        desired_width = 100  # Change this to your desired width
        desired_height = 80  # Change this to your desired height
        scaled_pixmap = pixmap.scaled(desired_width, desired_height, Qt.AspectRatioMode.KeepAspectRatio,
                                      Qt.TransformationMode.SmoothTransformation)
        tupuSetting_pix.setPixmap(scaled_pixmap)
        tupuSetting_pix.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tupuSetting_zi = QLabel("输出V3.0\ntupuSet")
        tupuSetting_zi.setProperty('class', 'btn_label')
        tupuSetting_zi.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tupuSetting_layout = QVBoxLayout()
        tupuSetting_layout.addWidget(tupuSetting_pix)
        tupuSetting_layout.addWidget(tupuSetting_zi)
        self.pushButton_8.setLayout(tupuSetting_layout)
        self.pushButton_8.setProperty('class', 'big_button')
        self.pushButton_8.setFocusPolicy(Qt.FocusPolicy.NoFocus)

    def load_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "excel files(*.xlsx)")
        if file_name:
            self.data_all_edit = file_name
            self.label.setText("导入文件成功")
        else:
            self.label.setText("未导入文件")

    def update_text_edit(self, file_path, label_text):
        if label_text == "data_all文件：":
            self.data_all_edit = file_path
        elif label_text == "DW-导入表：":
            self.sensor_record_edit = file_path
        if file_path != '':
            self.label.setText("导入文件成功")
        else:
            self.label.setText("未导入文件")

    def dat2700_img(self):
        self.label.setText("")
        if self.data_all_edit == '':
            self.label.setText("未读取到data_all文件，请先导入数据")
            return
        output_openfile_name = QFileDialog.getSaveFileName(self, "设置保存路径", "./DW导入表.xlsx", "excel files(*.xlsx)")
        self.output_path_7 = output_openfile_name[0]
        if self.output_path_7 == '':
            return
        self.label.setText("文件正在处理，请稍等")
        self.worker = Worker5(self.data_all_edit, self.output_path_7)
        self.worker.finished.connect(self.on_task_finished)
        self.worker.start()

    def json_img(self):
        self.label.setText("")
        if self.data_all_edit == '':
            self.label.setText("未读取到data_all文件，请先导入数据")
            return
        output_openfile_name = QFileDialog.getExistingDirectory(self, "设置保存文件夹", "./")
        self.output_path_4 = output_openfile_name
        if self.output_path_4 == '':
            return
        self.label.setText("文件正在处理，请稍等")
        self.worker = Worker4(self.data_all_edit, self.output_path_4)
        self.worker.finished.connect(self.on_task_finished)
        self.worker.start()

    def tupuset_img_V2(self):
        self.label.setText("")
        if self.data_all_edit == '':
            self.label.setText("未读取到data_all文件，请先导入数据")
            return
        if self.my_deftable == '':
            self.label.setText("未读取到key值对应表数据，请查看后台文件")
            return
        output_openfile_name = QFileDialog.getSaveFileName(self, "设置保存路径", "./tupusettingV2.xlsx", "excel files(*.xlsx)")
        self.output_path_3 = output_openfile_name[0]
        if self.output_path_3 == '':
            return
        self.label.setText("文件正在处理，请稍等")
        self.worker = Worker3(self.data_all_edit, self.output_path_3)
        self.worker.finished.connect(self.on_task_finished)
        self.worker.start()

    def tupuset_img_V3(self):
        self.label.setText("")
        if self.data_all_edit == '':
            self.label.setText("未读取到data_all文件，请先导入数据")
            return
        if self.my_deftable == '':
            self.label.setText("未读取到key值对应表数据，请查看后台文件")
            return
        output_openfile_name = QFileDialog.getSaveFileName(self, "设置保存路径", "./tupusettingV3.xlsx", "excel files(*.xlsx)")
        self.output_path_3 = output_openfile_name[0]
        if self.output_path_3 == '':
            return
        self.label.setText("文件正在处理，请稍等")
        self.worker = Worker6(self.data_all_edit, self.output_path_3)
        self.worker.finished.connect(self.on_task_finished)
        self.worker.start()

    def device_img(self):
        self.label.setText("")
        if self.data_all_edit == '':
            self.label.setText("未读取到data_all文件，请先导入数据")
            return
        if self.my_deftable == '':
            self.label.setText("未读取到key值对应表数据，请查看后台数据")
            return
        output_openfile_name = QFileDialog.getSaveFileName(self, "设置保存路径", "./deviceInfo.xlsx", "excel files(*.xlsx)")
        self.output_path_2 = output_openfile_name[0]
        if self.output_path_2 == '':
            return
        self.label.setText("文件正在处理，请稍等")
        self.worker = Worker2(self.data_all_edit, self.output_path_2)
        self.worker.finished.connect(self.on_task_finished)
        self.worker.start()

    def predict_img(self):
        self.label.setText("")
        if self.data_all_edit == '':
            self.label.setText("未读取到data_all文件，请先导入数据")
            return

        output_openfile_name = QFileDialog.getSaveFileName(self, "设置保存路径", "./平台导入表.xlsx", "excel files(*.xlsx)")
        self.output_path_1 = output_openfile_name[0]
        if self.output_path_1 == '':
            return
        self.label.setText("文件正在处理，请稍等")
        self.worker = Worker1(self.data_all_edit, self.output_path_1)
        self.worker.finished.connect(self.on_task_finished)
        self.worker.start()

    def download_tmp(self):
        output_openfile_name = QFileDialog.getExistingDirectory(self, "设置保存文件夹", "./")
        if not output_openfile_name:
            self.label.setText("未选择文件夹")
            return
        try:
            tmp_folder_path = f'{output_openfile_name}/data_all(模板).xlsx'
            shutil.copy('后台文件/data_all(模板).xlsx', tmp_folder_path)
            self.label.setText(f"data_all文件模板已经保存到：\n{tmp_folder_path}")
        except Exception as e:
            self.label.setText("保存模板错误，错误信息：\n" + str(e))

    def on_task_finished(self, message):
        self.label.setText(message)

    def closeEvent(self, event):
        box_reply = CustomMessageBox("是否要退出程序？")
        box_reply.setWindowTitle('退出')
        result = box_reply.exec()
        if result == QDialog.DialogCode.Accepted:
            event.accept()
        else:
            event.ignore()

    def paintEvent(self, event):
        painter = QPainter(self)
        pixmap = QPixmap("images/bg.png")
        scaled_pixmap = pixmap.scaled(self.size(), Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                                      Qt.TransformationMode.SmoothTransformation)
        painter.drawPixmap(self.rect(), scaled_pixmap)
        super().paintEvent(event)

    def showEvent(self, event):
        super().showEvent(event)
        self.raise_()  # 确保窗口在最上面
        self.activateWindow()  # 激活窗口


if __name__ == "__main__":
    app = QApplication(sys.argv)
    extra = {
        'density_scale': '13'
    }
    apply_stylesheet(app, theme="dark_blue.xml", extra=extra, css_file='images/custom.css')
    x = MyMainWindow()
    x.show()
    sys.exit(app.exec())
