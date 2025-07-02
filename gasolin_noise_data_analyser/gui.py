import sys
import time
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel,
    QVBoxLayout, QHBoxLayout, QProgressBar, QSplashScreen
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QPixmap

import pandas as pd
import matplotlib.pyplot as plt
import os
from pathlib import Path 
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment

plt.rcParams['legend.fontsize'] = 5

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def plot_acceleration_graph(ax, time_list, acc_list, start_point):
    ax.plot(time_list[start_point-24:], acc_list[start_point-24:], label='accelerations', color='blue', linewidth=0.5)
    ax.axvline(x=time_list[start_point], color='red', linestyle='--', label='peak point', linewidth=0.5)
    ax.set_xlabel("Time(s)", fontsize=6)
    ax.set_ylabel("ACC_X(G)", fontsize=6)
    ax.set_title("Acceleration Graph", fontsize=6)
    ax.tick_params(axis='both', labelsize=5)
    ax.legend()

def plot_rear_center_noise(ax, time_list, noise_list, start_point, before_one_min_val=54, after_one_min_val=34):
    ax.plot(time_list[start_point-24:], noise_list[start_point-24:], label='rear seat center noise(dB)', color='blue', linewidth=0.5)
    ax.axvline(x=time_list[start_point], color='red', linestyle='--', label='peak point', linewidth=0.5)
    ax.hlines(y=before_one_min_val, xmin = time_list[start_point], xmax = time_list[start_point]+1, colors = 'purple', linestyle = '-', linewidth =0.5)
    ax.hlines(y=after_one_min_val, xmin = time_list[start_point]+1, xmax = time_list[-1], colors = 'purple', linestyle = '-', linewidth =0.5)
    ax.vlines(x=time_list[start_point]+1, ymin=34, ymax=54, colors = 'purple', linestyle = '-', label = 'reference line', linewidth =0.5)
    ax.set_xlabel("Time(s)", fontsize=6)
    ax.set_ylabel("Rear SEAT_CENTER(dB-A)", fontsize=6)
    ax.set_title("Rear Seat Center Noise Graph", fontsize=6)
    ax.tick_params(axis='both', labelsize=5)
    ax.legend()

def conclusion_writer(csv_path, peak, peak_after_one_min, before_one_min_val=54, after_one_min_val=34):
    csv_path = Path(csv_path)
    folder= csv_path.parent
    xlsx_path = folder / "result.xlsx"

    wb = load_workbook(xlsx_path)
    ws = wb["conclusion"]

    stem = csv_path.stem
    parts = stem.split("_")

    if 'FL' in parts:
        fuel_key = 'FL'
        acc_vol = parts[5] if len(parts) > 5 else 'Unknown'
    else:
        if len(parts) > 5:
            fuel_key = f'{parts[4]}/{parts[5]}'
            acc_vol = parts[6]
        else:
            raise ValueError(f"fuel type recognition error:{stem}")

    judge = "G" if (peak < before_one_min_val and peak_after_one_min < after_one_min_val) else "N"

    acc_to_row = {
        '02G': 2,
        '03G': 5
    }

    fuel_to_offset = {
        'FL': 0,     # FULL
        '7/8': 1,     # 7/8
        '3/4': 2      # 3/4
    }

    base_row = acc_to_row.get(acc_vol.upper())
    offset = fuel_to_offset.get(fuel_key.upper())

    if base_row is None or offset is None:
        raise ValueError(f"error: acc={acc_vol}, fuel={fuel_key}")

    writing_row = base_row + offset

    ws.cell(row=writing_row, column=3, value=round(peak,1))
    ws.cell(row=writing_row, column=4, value=before_one_min_val)
    ws.cell(row=writing_row, column=5, value=round(peak_after_one_min,1))
    ws.cell(row=writing_row, column=6, value=after_one_min_val)
    ws.cell(row=writing_row, column=7, value=judge)

    wb.save(xlsx_path)

def generate_graph_from_csv(csv_path, save_name=None, save=False):

    df = pd.read_csv(csv_path)

    expected_columns = ["Time(s)", "ACC_X(G)", "REAR SEAT_CENTER(dB-A)"]
    for col in expected_columns:
        if col not in df.columns:
            raise ValueError(f"losing columns:{col},please check the .csv file")

    time_list = df["Time(s)"].tolist()
    acc_list = df["ACC_X(G)"].tolist()
    noise_list = df["REAR SEAT_CENTER(dB-A)"].tolist()

    delta_acc_list = [acc_list[i] - acc_list[i - 1] for i in range(1, len(acc_list))]
    start_point = delta_acc_list.index(max(delta_acc_list))

    peak_point = acc_list.index(max(acc_list[start_point:]))

    print(f"peak point index = {peak_point}, Time = {time_list[peak_point]}")

    fig, axs = plt.subplots(2, 1, figsize=(5, 4))
    fig.suptitle(save_name, fontsize = 7)
    plot_acceleration_graph(axs[0], time_list, acc_list, peak_point)
    plot_rear_center_noise(axs[1], time_list, noise_list, peak_point)

    after_one_min_point = peak_point+25
    conclusion_writer(csv_path, peak=max(noise_list[peak_point:peak_point+24]), peak_after_one_min=max(noise_list[after_one_min_point:]))

    plt.tight_layout()

    if save == True:
        folder = os.path.dirname(csv_path)
        save_folder = os.path.join(folder, "figures")
        os.makedirs(save_folder, exist_ok=True)
        fig_name = f'{save_name}.png'
        save_path = os.path.join(save_folder, fig_name)
        fig.savefig(save_path, dpi=300)

    plt.close(fig)

def xlsx_init(path):
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "conclusion"
    ws2 = wb.create_sheet(title="result")

    headers1 = ["加速度", "容量", "車内音 ~1s (dB-A)", "目標値(dB-A)", "車内音 1s~ (dB-A)", "目標値(dB-A)", "判定"]
    headers2 = ["加速度", "Full", "7/8", "3/4"]

    ws1.append(headers1)
    ws2.append(headers2)

    ws1['A2'] = "0.2G"
    ws1['A5'] = "0.3G"
    ws1.merge_cells('A2:A4')
    ws1.merge_cells('A5:A7')

    ws1['B2'] = "FULL"
    ws1['B3'] = "7/8"
    ws1['B4'] = "3/4"
    ws1['B5'] = "FULL"
    ws1['B6'] = "7/8"
    ws1['B7'] = "3/4"

    for col in ['C', 'E']:
        ws1.column_dimensions[col].width = 17
    for col in ['D', 'F']:
        ws1.column_dimensions[col].width = 13

    ws2.row_dimensions[2].height = 230
    ws2.row_dimensions[3].height = 230

    for col in ['B', 'C', 'D']:
        ws2.column_dimensions[col].width = 51

    ws2['A2'] = "0.2G"
    ws2['A3'] = "0.3G"

    save_path = f'{path}/result.xlsx'
    wb.save(save_path)

def fig_writer(path):
    path = Path(path)

    xlsx_path = path / "result.xlsx"
    img_folder_path = path / "figures"

    if not xlsx_path.exists():
        raise FileNotFoundError(f"未找到 Excel 文件：{xlsx_path}")
    if not img_folder_path.exists():
        raise FileNotFoundError(f"未找到图片文件夹：{img_folder_path}")

    wb = load_workbook(xlsx_path)
    ws = wb["result"]

    acc_to_row = {
        '02G': 2,
        '03G': 3
    }
    fuel_to_col = {
        'FL': 'B',
        '7/8': 'C',
        '3/4': 'D'
    }

    for img_path in img_folder_path.iterdir():
        if not img_path.is_file() or img_path.suffix.lower() != ".png":
            continue

        img = Image(img_path)
        img.width = 370
        img.height = 300

        stem = img_path.stem
        parts = stem.split("_")

        if 'FL' in parts:
            fuel_key = 'FL'
            acc_vol = parts[5] if len(parts) > 5 else 'Unknown'
        else:
            if len(parts) > 5:
                fuel_key = f'{parts[4]}/{parts[5]}'
                acc_vol = parts[6]
            else:
                raise ValueError(f"fuel type recognition error: {stem}")

        import_row = acc_to_row.get(acc_vol.upper())
        import_col = fuel_to_col.get(fuel_key.upper())

        if import_row is None or import_col is None:
            raise ValueError(f"could not find import position, acc={acc_vol}, fuel={fuel_key}, 文件名={stem}")

        import_position = f"{import_col}{import_row}"
        ws.add_image(img, import_position)

    # 一次性保存
    wb.save(xlsx_path)

def processer(path, progress_callback):
    folder_path = Path(path)
    xlsx_init(path)

    csv_files = [f for f in folder_path.iterdir() if f.is_file() and f.suffix.lower() == ".csv"]
    total = len(csv_files)

    if total == 0:
        raise FileNotFoundError("could not find csv file")

    for i, file_path in enumerate(csv_files, 1):
        try:
            stem = file_path.stem
            generate_graph_from_csv(file_path, save_name=stem, save=True)
        except Exception as e:
            print(f"csv process failure: {file_path.name}:{e}")

        progress = int(i / total * 90)  
        progress_callback.emit(progress)

    # 图像写入 Excel
    try:
        fig_writer(path)
        progress_callback.emit(100)
    except Exception as e:
        print(f"figure import failure: {e}")
        progress_callback.emit(100)

class WorkerThread(QThread):
    progress = pyqtSignal(int)

    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path

    def run(self):
        processer(self.folder_path, self.progress)

# main gui
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("揺動音自動処理ツール")
        self.setAcceptDrops(True)
        self.setFixedSize(420, 250)

        self.folder_path = None

        self.label = QLabel(".CSVを導入してください")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setStyleSheet("border: 2px dashed gray; padding: 40px;")

        self.start_button = QPushButton("処理開始")
        self.start_button.clicked.connect(self.start_processing)

        self.quit_button = QPushButton("終了")
        self.quit_button.clicked.connect(self.close)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)

        # 布局
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.progress_bar)

        buttons = QHBoxLayout()
        buttons.addWidget(self.start_button)
        buttons.addWidget(self.quit_button)

        layout.addLayout(buttons)
        self.setLayout(layout)

    # 拖拽进入事件
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    # 拖拽释放事件
    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            folder = Path(urls[0].toLocalFile())
            if folder.is_dir():
                self.folder_path = folder
                self.label.setText(f"✅ フォルダ読み込み完了")
            else:
                self.label.setText("⚠️ フォルダをインポートしてください")

    # 启动主线程处理
    def start_processing(self):
        if not self.folder_path:
            self.label.setText("⚠️ フォルダ無効")
            return

        self.progress_bar.setValue(0)
        self.worker = WorkerThread(self.folder_path)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.finished.connect(lambda: self.label.setText("✅ 処理完了"))
        self.label.setText("⏳ 処理中")
        self.worker.start()

# 启动入口
if __name__ == "__main__":
    app = QApplication(sys.argv)
    splash_pix = QPixmap(resource_path("splash.png"))
    splash = QSplashScreen(splash_pix, Qt.WindowType.WindowStaysOnTopHint)
    splash.show()
    app.processEvents()
    window = MainWindow()
    window.show()
    splash.finish(window)
    sys.exit(app.exec())
