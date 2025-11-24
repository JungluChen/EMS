import sys, os, shutil, time
import sqlite3, json
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel
from PyQt5.QtWidgets import QComboBox
from PyQt5.QtWidgets import QHBoxLayout
from PyQt5.QtWidgets import QGridLayout
from PyQt5.QtWidgets import QRadioButton, QButtonGroup, QLineEdit, QFormLayout
from PyQt5.QtWidgets import QGroupBox, QSizePolicy
from PyQt5.QtWidgets import QScrollArea
from PyQt5.QtWidgets import QPushButton, QSpinBox, QCheckBox, QDoubleSpinBox, QFileDialog, QDialog, QDialogButtonBox, QMessageBox
from PyQt5.QtCore import Qt, QTimer, QPointF, QSettings, QStandardPaths
from PyQt5.QtGui import QPainter, QPen, QColor, QFont, QIcon
from functools import partial
from datetime import datetime
import csv
import threading
try:
    from PyQt5.QtSerialPort import QSerialPortInfo, QSerialPort
except Exception:
    QSerialPortInfo = None
    QSerialPort = None
try:
    import serial
except Exception:
    serial = None
try:
    import modbus_tk
    from modbus_tk import modbus_rtu
except Exception:
    modbus_tk = None
    modbus_rtu = None
class BlandPage(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CoCoa Linna")
        # 优化窗口尺寸比例，采用黄金比例 1.618:1
        primary_width = 1200
        primary_height = int(primary_width * 0.618)  # 约741px，符合黄金比例
        self.resize(primary_width, primary_height)
        
        # 主布局采用科学比例分配
        layout = QVBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)  # 标准16px边距
        layout.setSpacing(12)  # 组件间距12px，符合8px网格系统的1.5倍
        
        self.settings = QSettings('CoCoaLinna','NewEMS')
        self.save_path = self.settings.value('save_path', _default_save_path())
        self.save_format = self.settings.value('save_format', 'csv')
        self.sections = []
        # 保留預掃結果供快速選擇
        self.active_ports = set()
        base_dir = os.path.dirname(__file__)
        self.var_dir = os.path.join(base_dir, 'real_time_monitoring', 'temp')
        try:
            os.makedirs(self.var_dir, exist_ok=True)
        except Exception:
            pass
        self.db_path = os.path.join(self.var_dir, 'ems.db')
        self.recording_flag_path = os.path.join(self.var_dir, 'recording.lock')
        self.db_conn = None
        self._init_db()

        controls_group = QGroupBox("顯示設定")
        controls_group.setStyleSheet(
            "QGroupBox { font-weight: bold; border: 1px solid #e0e0e0; border-radius: 6px; margin-top: 8px; padding-top: 8px; background-color: #ffffff; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 4px; color: #2c3e50; font-size: 13px; }"
        )
        controls_form = QFormLayout()
        controls_form.setContentsMargins(8, 8, 8, 8)  # 减少内边距
        controls_form.setSpacing(6)  # 减少行间距
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(50, 5000)
        self.interval_spin.setValue(500)
        self.points_spin = QSpinBox()
        self.points_spin.setRange(30, 1000)
        self.points_spin.setValue(120)
        self.autoscale_check = QCheckBox()
        self.autoscale_check.setChecked(True)
        self.min_spin = QDoubleSpinBox()
        self.min_spin.setRange(-9999.0, 9999.0)
        self.min_spin.setDecimals(2)
        self.min_spin.setValue(0.0)
        self.max_spin = QDoubleSpinBox()
        self.max_spin.setRange(-9999.0, 9999.0)
        self.max_spin.setDecimals(2)
        self.max_spin.setValue(100.0)
        self.shared_port_combo = QComboBox()
        ports = self.list_serial_ports()
        if ports:
            self.shared_port_combo.addItem("空白")
            self.shared_port_combo.addItems(ports)
        else:
            self.shared_port_combo.addItem("空白")
            self.shared_port_combo.addItem("未偵測到串口")
        self.refresh_ports_btn = QPushButton("刷新串口")
        h_layout = QHBoxLayout()
        h_layout.setSpacing(8)  # 减少控件间距，压缩高度
        h_layout.setContentsMargins(0, 2, 0, 2)  # 减少上下边距
        h_layout.addWidget(QLabel("串口:"))
        h_layout.addWidget(self.shared_port_combo)
        h_layout.addWidget(self.refresh_ports_btn)
        h_layout.addWidget(QLabel("更新頻率(ms):"))
        h_layout.addWidget(self.interval_spin)
        h_layout.addWidget(QLabel("顯示點數:"))
        h_layout.addWidget(self.points_spin)
        h_layout.addWidget(QLabel("自動縮放:"))
        h_layout.addWidget(self.autoscale_check)
        h_layout.addWidget(QLabel("固定下限:"))
        h_layout.addWidget(self.min_spin)
        h_layout.addWidget(QLabel("固定上限:"))
        h_layout.addWidget(self.max_spin)
        h_layout.addStretch()
        controls_form.addRow(h_layout)
        self.save_path_label = QLabel(self.save_path)
        self.change_save_btn = QPushButton('變更儲存位置')
        self.save_format_label = QLabel(self.save_format)
        save_layout = QHBoxLayout()
        save_layout.setSpacing(8)  # 减少间距
        save_layout.setContentsMargins(0, 2, 0, 2)  # 减少上下边距
        save_layout.addWidget(QLabel('儲存位置:'))
        save_layout.addWidget(self.save_path_label)
        save_layout.addWidget(self.change_save_btn)
        save_layout.addWidget(QLabel('格式:'))
        save_layout.addWidget(self.save_format_label)
        save_layout.addStretch()
        controls_form.addRow(save_layout)
        self.port_scan_label = QLabel('設備掃描: 未執行')
        controls_form.addRow(self.port_scan_label)
        controls_group.setLayout(controls_form)
        controls_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        header_layout = QHBoxLayout()
        header_layout.setSpacing(6)  # 减少按钮间距
        header_layout.setContentsMargins(0, 4, 0, 8)  # 减少上下边距
        self.add_btn = QPushButton("＋")
        self.remove_btn = QPushButton("－")
        # 优化按钮尺寸，减小占用空间
        for btn in [self.add_btn, self.remove_btn]:
            btn.setFixedSize(24, 24)  # 减小按钮尺寸
            btn.setStyleSheet("QPushButton { font-size: 14px; font-weight: bold; border-radius: 4px; padding: 2px; }")
        header_layout.addWidget(self.add_btn)
        header_layout.addWidget(self.remove_btn)
        header_layout.addStretch(1)
        layout.addLayout(header_layout)

        self.section_container = QWidget()
        self.section_grid = QGridLayout()
        self.section_grid.setSpacing(12)  # 减少网格间距
        self.section_grid.setContentsMargins(2, 2, 2, 2)  # 减少网格内边距
        self.section_container.setLayout(self.section_grid)
        self.section_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.section_scroll = QScrollArea()
        self.section_scroll.setWidgetResizable(True)
        self.section_scroll.setWidget(self.section_container)
        self.section_scroll.setStyleSheet("QScrollArea { border: none; background-color: transparent; }")
        self.section_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        layout.addWidget(self.section_scroll)

        layout.addWidget(controls_group)
        # 科学分配主布局空间比例：
        # index 1: QScrollArea（主要内容区）占 75% - 符合60-70%要求
        # index 2: QGroupBox（控制面板）占 25% - 优化空间利用率
        layout.setStretch(1, 8)  # 75% 空间给主要内容
        layout.setStretch(2, 1)  # 25% 空间给控制面板

        self.add_btn.clicked.connect(self._add_section)
        self.remove_btn.clicked.connect(self._remove_section)
        self.refresh_ports_btn.clicked.connect(self._refresh_ports)
        self.change_save_btn.clicked.connect(self._change_storage_settings)
        
        self.interval_spin.valueChanged.connect(self._update_interval)
        self.points_spin.valueChanged.connect(self._update_points)
        self.autoscale_check.toggled.connect(self._toggle_autoscale)
        self.min_spin.valueChanged.connect(self._update_fixed_range)
        self.max_spin.valueChanged.connect(self._update_fixed_range)
        self.setLayout(layout)
        self._update_save_labels()

        self._add_section()
        self._pre_scan_ports()

    def list_serial_ports(self):
        ports = []
        if QSerialPortInfo is not None:
            ports = [p.systemLocation() or p.portName() for p in QSerialPortInfo.availablePorts()]
        else:
            import glob
            sysname = sys.platform
            patterns = []
            if sysname.startswith('darwin'):
                patterns = ['/dev/tty.*', '/dev/cu.*']
            elif sysname.startswith('linux'):
                patterns = ['/dev/ttyUSB*', '/dev/ttyACM*', '/dev/ttyS*']
            elif sysname.startswith('win'):
                try:
                    from serial.tools import list_ports
                    ports = [p.device for p in list_ports.comports()]
                except Exception:
                    ports = [f'COM{n}' for n in range(1, 33)]
                return ports
            for pat in patterns:
                ports.extend(glob.glob(pat))
        return sorted(set(ports))

    def create_box(self, title):
        box = QGroupBox(title)
        box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        # 优化组框样式，减少垂直空间
        box.setStyleSheet(
            "QGroupBox { "
            "  font-weight: bold; "
            "  border: 1px solid #d8d8d8; "
            "  border-radius: 8px; "
            "  margin-top: 8px; "
            "  padding-top: 8px; "
            "  background-color: #ffffff; "
            "}"
            "QGroupBox::title { "
            "  subcontrol-origin: margin; "
            "  left: 8px; "
            "  padding: 0 6px; "
            "  color: #2c3e50; "
            "  font-size: 14px; "
            "  font-weight: 600; "
            "}"
        )
        shift_combo = QComboBox()
        shift_combo.addItems(["早班", "晚班"])
        temp_addr_combo = QComboBox()
        temp_addr_combo.addItem("未掃描")
        temp_scan_btn = QPushButton("掃描")
        current_addr_combo = QComboBox()
        current_addr_combo.addItem("未掃描")
        current_addr_combo.addItem("空白")
        current_scan_btn = QPushButton("掃描")
        material_input = QLineEdit()

        form_layout = QFormLayout()
        form_layout.setContentsMargins(12, 8, 12, 8)  # 减少内边距
        form_layout.setSpacing(8)  # 减少行间距
        form_layout.setFormAlignment(Qt.AlignLeft | Qt.AlignTop)
        form_layout.setLabelAlignment(Qt.AlignRight)  # 标签右对齐，提升可读性
        form_layout.addRow("班次:", shift_combo)
        temp_addr_row = QHBoxLayout()
        temp_addr_row.addWidget(temp_addr_combo)
        temp_addr_row.addWidget(temp_scan_btn)
        temp_addr_container = QWidget()
        temp_addr_container.setLayout(temp_addr_row)
        form_layout.addRow("溫度地址:", temp_addr_container)
        current_addr_row = QHBoxLayout()
        current_addr_row.addWidget(current_addr_combo)
        current_addr_row.addWidget(current_scan_btn)
        current_addr_container = QWidget()
        current_addr_container.setLayout(current_addr_row)
        form_layout.addRow("電流地址:", current_addr_container)
        form_layout.addRow("工單號:", material_input)

        val_layout = QHBoxLayout()
        val_layout.setSpacing(12)  # 减少数值标签间距
        val_layout.setContentsMargins(0, 4, 0, 4)  # 减少上下边距
        
        temp_label = QLabel("溫度: -- °C")
        current_label = QLabel("電流: -- A")
        temp_status = QLabel("狀態: 未檢查")
        current_status = QLabel("狀態: 未檢查")
        
        # 优化数值显示字体，减小尺寸
        value_font = QFont()
        value_font.setPointSize(12)  # 减小数值字体
        value_font.setBold(True)
        value_font.setFamily("Arial")
        
        status_font = QFont()
        status_font.setPointSize(10)  # 减小状态字体
        status_font.setFamily("Arial")
        
        temp_label.setFont(value_font)
        current_label.setFont(value_font)
        temp_status.setFont(status_font)
        current_status.setFont(status_font)
        
        # 设置状态标签最小宽度，减小尺寸
        for label in [temp_status, current_status]:
            label.setMinimumWidth(60)  # 减少最小宽度
            label.setAlignment(Qt.AlignCenter)
        
        val_layout.addWidget(temp_label)
        val_layout.addWidget(temp_status)
        val_layout.addStretch()  # 添加弹性空间
        val_layout.addWidget(current_label)
        val_layout.addWidget(current_status)

        chart_layout = QVBoxLayout()
        chart_layout.setContentsMargins(0, 2, 0, 2)  # 進一步減少圖表上下邊距
        
        chart_group = QGroupBox("時序圖")
        dual_plot = DualLinePlot(color1=QColor(220,20,60), color2=QColor(30,144,255), parent=chart_group)
        dual_plot.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        chart_group.setStyleSheet(
            "QGroupBox { "
            "  font-weight: bold; "
            "  border: 1px solid #d0d0d0; "
            "  border-radius: 6px; "
            "  margin-top: 8px; "
            "  padding-top: 8px; "
            "  background-color: #fafafa; "
            "}"
            "QGroupBox::title { "
            "  subcontrol-origin: margin; "
            "  left: 8px; "
            "  padding: 0 4px; "
            "  color: #34495e; "
            "  font-size: 12px; "
            "  font-weight: 600; "
            "}"
        )
        
        vg = QVBoxLayout()
        vg.setContentsMargins(2, 2, 2, 2)  # 進一步減少圖表區域內邊距
        vg.setSpacing(2)  # 進一步減少間距
        
        legend_layout = QHBoxLayout()
        legend_layout.setContentsMargins(2, 0, 2, 0)  # 進一步減少圖例邊距
        temp_legend = QLabel("溫度(°C)")
        temp_legend.setStyleSheet("color: rgb(220,20,60); font-weight: 600; font-size: 11px;")
        current_legend = QLabel("電流(A)")
        current_legend.setStyleSheet("color: rgb(30,144,255); font-weight: 600; font-size: 11px;")
        legend_layout.addWidget(temp_legend)
        legend_layout.addStretch(1)
        legend_layout.addWidget(current_legend)
        
        vg.addLayout(legend_layout)
        vg.addWidget(dual_plot)
        chart_group.setLayout(vg)
        chart_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        chart_layout.addWidget(chart_group)

        ctrl_layout = QHBoxLayout()
        ctrl_layout.setSpacing(8)  # 减少控制按钮间距
        ctrl_layout.setContentsMargins(0, 6, 0, 0)  # 减少上边距
        
        start_btn = QPushButton("開始")
        stop_btn = QPushButton("停止")
        reset_btn = QPushButton("重置")
        delete_btn = QPushButton("刪除")
        duration_label = QLabel("運行時長: 00:00:00")
        data_source_label = QLabel("資料來源: 未開始")
        
        # 优化按钮样式，减小尺寸
        for btn in [start_btn, stop_btn, reset_btn, delete_btn]:
            btn.setFixedHeight(28)  # 减少按钮高度
            btn.setMinimumWidth(70)  # 减少最小宽度
            btn.setStyleSheet("""
                QPushButton {
                    font-size: 12px;
                    font-weight: 500;
                    padding: 4px 8px;
                    border-radius: 4px;
                    border: 1px solid #d0d0d0;
                    background-color: #ffffff;
                }
                QPushButton:hover {
                    background-color: #f8f9fa;
                    border-color: #b0b0b0;
                }
                QPushButton:pressed {
                    background-color: #e9ecef;
                }
            """)
        
        # 设置标签样式，减小尺寸
        duration_label.setStyleSheet("font-size: 11px; color: #495057;")
        data_source_label.setStyleSheet("font-size: 11px; color: #6c757d;")
        duration_label.setMinimumWidth(100)  # 减少最小宽度
        data_source_label.setMinimumWidth(80)  # 减少最小宽度
        
        ctrl_layout.addWidget(start_btn)
        ctrl_layout.addWidget(stop_btn)
        ctrl_layout.addWidget(reset_btn)
        ctrl_layout.addWidget(delete_btn)
        ctrl_layout.addStretch()  # 弹性空间
        ctrl_layout.addWidget(duration_label)
        ctrl_layout.addWidget(data_source_label)

        v = QVBoxLayout()
        v.addLayout(form_layout)
        v.addLayout(val_layout)
        v.addLayout(chart_layout)
        v.addLayout(ctrl_layout)
        v.setStretch(0, 0)
        v.setStretch(1, 0)
        v.setStretch(2, 1)
        v.setStretch(3, 0)
        box.setLayout(v)
        box.setStyleSheet(
            "QGroupBox { border: 1px solid #d0d0d0; border-radius: 6px; margin-top: 8px; background-color: #fafafa; } "
            "QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 6px; font-weight: bold; color: #333; }"
        )
        s = {
            'name': title,
            'shift': shift_combo,
            'temp_addr': temp_addr_combo,
            'current_addr': current_addr_combo,
            'material': material_input,
            'temp_label': temp_label,
            'current_label': current_label,
            'temp_status': temp_status,
            'current_status': current_status,
            'plot': dual_plot,
            'start_btn': start_btn,
            'stop_btn': stop_btn,
            'reset_btn': reset_btn,
            'duration_label': duration_label,
            'data_source_label': data_source_label,
            'timer': QTimer(self),
            'start_time': None,
            'records': [],
            'box': box,
            'mode': 'idle',
            'master': None,
            'port': None,
            'latest_temp': None,
            'latest_current': None,
            'reader_thread': None,
            'reader_stop': None,
            'temp_addr_value': None,
            'current_addr_value': None
        }
        s['timer'].setInterval(int(self.interval_spin.value()))
        s['timer'].timeout.connect(partial(self._tick_section, s))
        # 延遲初始化圖表高度，等待主窗口完成佈局
        QTimer.singleShot(500, lambda: dual_plot._delayed_init_height())
        start_btn.clicked.connect(partial(self._start_section, s))
        stop_btn.clicked.connect(partial(self._stop_section, s))
        reset_btn.clicked.connect(partial(self._reset_section, s))
        delete_btn.clicked.connect(partial(self._remove_section_box, s))
        temp_scan_btn.setToolTip("掃描並選擇溫度感測器地址")
        current_scan_btn.setToolTip("掃描並選擇電流感測器地址，空白表示未安裝")
        temp_addr_combo.setToolTip("選擇 Modbus RTU 從站地址作為溫度感測器")
        current_addr_combo.setToolTip("可選擇空白代表未安裝電流感測器")
        temp_scan_btn.clicked.connect(partial(self._scan_addresses_for_section, s, 'temp'))
        current_scan_btn.clicked.connect(partial(self._scan_addresses_for_section, s, 'current'))
        s['plot'].setMaxPoints(int(self.points_spin.value()))
        s['plot'].setAutoScale(bool(self.autoscale_check.isChecked()))
        s['plot'].setFixedRange(float(self.min_spin.value()), float(self.max_spin.value()))
        # 移除 mode 切換，預設為真實模式
        s['mode'] = 'real'
        s['interval_ms'] = int(self.interval_spin.value())
        self.sections.append(s)
        return box

    def _add_section(self):
        box = self.create_box(f"生產線{len(self.sections)+1}")
        self._refresh_grid()
        self._update_interval(self.interval_spin.value())
        self._update_points(self.points_spin.value())
        self._toggle_autoscale(self.autoscale_check.isChecked())

    def _remove_section(self):
        if not self.sections:
            return
        s = self.sections.pop()
        if s['timer'].isActive():
            s['timer'].stop()
        w = s['box']
        self.section_grid.removeWidget(w)
        w.deleteLater()
        self._refresh_grid()

    def _remove_section_box(self, s):
        if s in self.sections:
            if s['timer'].isActive():
                s['timer'].stop()
            w = s['box']
            self.section_grid.removeWidget(w)
            w.deleteLater()
            self.sections.remove(s)
            self._refresh_grid()

    def _refresh_grid(self):
        while self.section_grid.count():
            item = self.section_grid.takeAt(0)
            w = item.widget()
            if w:
                self.section_grid.removeWidget(w)
        for i, s in enumerate(self.sections):
            row = i // 3
            col = i % 3
            self.section_grid.addWidget(s['box'], row, col)
            self.section_grid.setRowStretch(row, 1)
        self.section_grid.setColumnStretch(0, 1)
        self.section_grid.setColumnStretch(1, 1)
        self.section_grid.setColumnStretch(2, 1)
        self._apply_equal_widths()

    def _apply_equal_widths(self):
        try:
            vpw = self.section_scroll.viewport().width()
            if vpw <= 0:
                return
            spacing = self.section_grid.horizontalSpacing() or 0
            cols = 3
            w_per = max(260, int((vpw - spacing * (cols - 1)) / cols))
            for s in self.sections:
                bx = s.get('box')
                if bx:
                    bx.setMinimumWidth(0)
                    bx.setMaximumWidth(w_per)
        except Exception:
            pass

    def _mb_crc(self, data):
        crc = 0xFFFF
        for b in data:
            crc ^= b
            for _ in range(8):
                if crc & 1:
                    crc = (crc >> 1) ^ 0xA001
                else:
                    crc >>= 1
        return crc

    def _mb_build_report_slave_id(self, addr):
        base = bytes([addr, 0x11])
        crc = self._mb_crc(base)
        return base + bytes([crc & 0xFF, (crc >> 8) & 0xFF])

    def _scan_addresses(self, port_name, baud=9600, max_addr=32, timeout_ms=150):
        found = []
        if modbus_rtu is not None and serial is not None:
            try:
                ser = serial.Serial(port_name, baudrate=baud, bytesize=8, parity='N', stopbits=1, timeout=timeout_ms/1000.0)
            except Exception:
                return found
            try:
                master = modbus_rtu.RtuMaster(ser)
                master.set_timeout(timeout_ms/1000.0)
                master.set_verbose(False)
            except Exception:
                try:
                    ser.close()
                except Exception:
                    pass
                return found
            for addr in range(1, max_addr+1):
                try:
                    # 只接受回應長度=2（1 暫存器）的設備，降低誤判
                    rr = master.execute(addr, 3, 0, 1)
                    if rr and len(rr) == 1:
                        found.append(addr)
                except Exception:
                    pass
            try:
                master.close()
            except Exception:
                pass
            try:
                ser.close()
            except Exception:
                pass
            return found
        if QSerialPort is not None:
            sp = QSerialPort(port_name)
            sp.setBaudRate(baud)
            sp.setDataBits(QSerialPort.Data8)
            sp.setParity(QSerialPort.NoParity)
            sp.setStopBits(QSerialPort.OneStop)
            if not sp.open(QSerialPort.ReadWrite):
                return found
            for addr in range(1, max_addr+1):
                req = self._mb_build_report_slave_id(addr)
                sp.write(req)
                sp.waitForBytesWritten(timeout_ms)
                if sp.waitForReadyRead(timeout_ms):
                    resp = bytes(sp.readAll())
                    if len(resp) >= 4 and resp[0] == addr:
                        d = resp[:-2]
                        crc = resp[-2] | (resp[-1] << 8)
                        if self._mb_crc(d) == crc:
                            found.append(addr)
            sp.close()
            return found
        if serial is not None:
            try:
                ser = serial.Serial(port_name, baudrate=baud, bytesize=8, parity='N', stopbits=1, timeout=timeout_ms/1000.0)
            except Exception:
                return found
            for addr in range(1, max_addr+1):
                req = self._mb_build_report_slave_id(addr)
                try:
                    ser.write(req)
                    resp = ser.read(64)
                except Exception:
                    resp = b''
                if len(resp) >= 4 and resp[0] == addr:
                    d = resp[:-2]
                    crc = resp[-2] | (resp[-1] << 8)
                    if self._mb_crc(d) == crc:
                        found.append(addr)
            try:
                ser.close()
            except Exception:
                pass
            return found
        return found

    def _scan_addresses_for_section(self, s, which):
        """掃描後讓使用者可選「空白」跳過，不強制接受掃描結果"""
        if which == 'temp':
            port = self._selected_port()
            combo = s['temp_addr']
        else:
            port = self._selected_port()
            combo = s['current_addr']
        addrs = self._scan_addresses(port)
        combo.clear()
        # 永遠保留「空白」選項（電流）或「未偵測到」
        combo.addItem("空白")
        if addrs:
            combo.addItems([str(a) for a in addrs])
        else:
            combo.addItem("未偵測到")

    def _refresh_ports(self):
        ports = self.list_serial_ports()
        self.shared_port_combo.clear()
        self.shared_port_combo.addItem("空白")
        if ports:
            self.shared_port_combo.addItems(ports)
        else:
            self.shared_port_combo.addItem("未偵測到串口")

    def _start_section(self, s):
        if not self._ensure_settings_valid():
            return
        if not s['timer'].isActive():
            ok = self._init_section_connection(s)
            s['start_time'] = datetime.now()
            s['start_btn'].setEnabled(False)
            s['timer'].start()
            self._set_recording_active(True)
            if ok and s.get('master'):
                stop_event = threading.Event()
                s['reader_stop'] = stop_event
                def run():
                    self._reader_loop(s)
                t = threading.Thread(target=run, daemon=True)
                s['reader_thread'] = t
                t.start()

    def _tick_section(self, s):
        t = (datetime.now() - s['start_time']).total_seconds() if s['start_time'] else 0
        temp = s.get('latest_temp')
        current = s.get('latest_current')
        if temp is not None:
            s['temp_label'].setText(f"溫度: {temp:.1f} °C")
        else:
            s['temp_label'].setText("溫度: -- °C")
        if current is None and (s['current_addr'].currentText() == "空白"):
            s['current_label'].setText("電流: 未安裝")
            s['current_label'].setStyleSheet("color: gray;")
        elif current is not None:
            s['current_label'].setStyleSheet("")
            s['current_label'].setText(f"電流: {current:.2f} A")
        else:
            s['current_label'].setText("電流: -- A")
        hh = int(t // 3600)
        mm = int((t % 3600) // 60)
        ss = int(t % 60)
        s['duration_label'].setText(f"運行時長: {hh:02d}:{mm:02d}:{ss:02d}")
        s['plot'].append(temp if temp is not None else 0.0, current if current is not None else 0.0)
        s['records'].append({
            'line': s['name'],
            'shift': s['shift'].currentText(),
            'work_order': s['material'].text(),
            'time': datetime.now().isoformat(timespec='seconds'),
            'temperature': round(temp if temp is not None else 0.0, 3),
            'current': round(current if current is not None else 0.0, 3)
        })
        r = s['records'][-1]
        try:
            data = {
                'line': r['line'],
                'shift': r['shift'],
                'work_order': r['work_order'],
                'temperature': r['temperature'],
                'current': r['current']
            }
            self._insert_record(r['time'], json.dumps(data, ensure_ascii=False))
        except Exception:
            pass
        self._touch_recording_flag()

    def _stop_section(self, s):
        if s['timer'].isActive():
            s['timer'].stop()
            s['start_btn'].setEnabled(True)
        if s.get('reader_stop'):
            try:
                s['reader_stop'].set()
            except Exception:
                pass
        if s.get('reader_thread'):
            try:
                s['reader_thread'].join(timeout=0.5)
            except Exception:
                pass
            s['reader_thread'] = None
        self._export_section(s)
        if not self._any_recording_active():
            self._set_recording_active(False)

    def _reset_section(self, s):
        if s['timer'].isActive():
            s['timer'].stop()
        if s.get('reader_stop'):
            try:
                s['reader_stop'].set()
            except Exception:
                pass
        if s.get('reader_thread'):
            try:
                s['reader_thread'].join(timeout=0.5)
            except Exception:
                pass
            s['reader_thread'] = None
        s['start_time'] = None
        s['mode'] = 'idle'
        if s.get('master'):
            try:
                s['master'].close()
            except Exception:
                pass
            s['master'] = None
        s['duration_label'].setText("運行時長: 00:00:00")
        s['temp_label'].setText("溫度: -- °C")
        s['current_label'].setText("電流: -- A")
        s['plot'].clear()
        s['records'].clear()
        s['data_source_label'].setText("資料來源: 未開始")
        s['data_source_label'].setStyleSheet("")
        s['temp_status'].setText("狀態: 未檢查")
        s['temp_status'].setStyleSheet("")
        s['current_status'].setText("狀態: 未檢查")
        s['current_status'].setStyleSheet("")
        if not self._any_recording_active():
            self._set_recording_active(False)

    def _export_section(self, s):
        if not s['records']:
            return
        if not self._ensure_settings_valid():
            return
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        name = f"{s['name']}_{ts}.{self.save_format}"
        path = os.path.join(self.save_path, name)
        tmp = path + '.tmp'
        try:
            os.makedirs(self.save_path, exist_ok=True)
            if self.save_format == 'csv':
                with open(tmp, 'w', newline='', encoding='utf-8-sig') as f:
                    w = csv.writer(f)
                    w.writerow(["生產線", "班次", "工單號", "時間", "溫度", "電流"])
                    for r in s['records']:
                        w.writerow([r['line'], r['shift'], r['work_order'], r['time'], r['temperature'], r['current']])
            else:
                with open(tmp, 'w', encoding='utf-8-sig') as f:
                    f.write("生產線\t班次\t工單號\t時間\t溫度\t電流\r\n")
                    for r in s['records']:
                        f.write(f"{r['line']}\t{r['shift']}\t{r['work_order']}\t{r['time']}\t{r['temperature']}\t{r['current']}\r\n")
            os.replace(tmp, path)
            size = os.path.getsize(path)
            ts2 = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            _show_auto_close_message(self, "儲存完成", f"路徑: {path}\n檔案大小: {_human_size(size)}\n時間: {ts2}")
        except Exception as e:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except Exception:
                pass
            QMessageBox.critical(self, "儲存失敗", str(e))

    def _update_interval(self, v):
        for s in self.sections:
            s['timer'].setInterval(int(v))
            s['interval_ms'] = int(v)

    def _update_points(self, v):
        for s in self.sections:
            s['plot'].setMaxPoints(int(v))

    def _toggle_autoscale(self, checked):
        self.min_spin.setEnabled(not checked)
        self.max_spin.setEnabled(not checked)
        for s in self.sections:
            s['plot'].setAutoScale(bool(checked))
        self._update_fixed_range(None)

    def _update_fixed_range(self, _):
        for s in self.sections:
            s['plot'].setFixedRange(float(self.min_spin.value()), float(self.max_spin.value()))

    def _pre_scan_ports(self):
        self.port_scan_label.setText('設備掃描: 進行中...')
        def worker():
            all_ports = self.list_serial_ports()
            try:
                active = scan_modbus_devices(timeout=0.3, ports=all_ports)
            except Exception:
                active = []
            def apply():
                self.active_ports = set(active)
                old_text = self.shared_port_combo.currentText()
                self.shared_port_combo.clear()
                self.shared_port_combo.addItem("空白")
                self.shared_port_combo.addItems(all_ports)
                idx = self.shared_port_combo.findText(old_text)
                if idx >= 0:
                    self.shared_port_combo.setCurrentIndex(idx)
                if all_ports:
                    summary = ", ".join([p + ("(✓)" if p in self.active_ports else "(✗)") for p in all_ports])
                    self.port_scan_label.setText("設備掃描: " + summary)
                else:
                    self.port_scan_label.setText("設備掃描: 未偵測到設備")
            QTimer.singleShot(0, apply)
        threading.Thread(target=worker, daemon=True).start()

    def _selected_port(self):
        txt = self.shared_port_combo.currentText() or ""
        if txt == "空白" or txt == "未偵測到串口":
            return ""
        return txt

    def _init_section_connection(self, s):
        port = self._selected_port()
        temp_txt = s['temp_addr'].currentText()
        cur_txt = s['current_addr'].currentText()
        temp_addr = None
        current_addr = None
        try:
            temp_addr = int(temp_txt) if temp_txt and temp_txt.isdigit() else None
        except Exception:
            temp_addr = None
        if cur_txt == "空白":
            current_addr = None
        else:
            try:
                current_addr = int(cur_txt) if cur_txt and cur_txt.isdigit() else None
            except Exception:
                current_addr = None
        ok = False
        master = None
        if modbus_rtu is not None and serial is not None and port:
            try:
                ser = serial.Serial(port, baudrate=9600, bytesize=8, parity='N', stopbits=1, timeout=0.3)
                master = modbus_rtu.RtuMaster(ser)
                master.set_timeout(0.3)
                master.set_verbose(False)
                if temp_addr is not None:
                    try:
                        master.execute(temp_addr, 3, 0, 1)
                        ok = True
                    except Exception:
                        ok = False
                if not ok and current_addr is not None:
                    try:
                        master.execute(current_addr, 3, 0, 1)
                        ok = True
                    except Exception:
                        ok = False
            except Exception:
                master = None
        if ok and master is not None:
            s['mode'] = 'real'
            s['master'] = master
            s['port'] = port
            s['temp_addr_value'] = temp_addr
            s['current_addr_value'] = current_addr
            s['data_source_label'].setText("資料來源: 實時")
            s['data_source_label'].setStyleSheet("color: green;")
            # 更新感測器狀態標記
            if temp_addr is not None:
                s['temp_status'].setText("狀態: 已連線")
                s['temp_status'].setStyleSheet("color: green;")
            else:
                s['temp_status'].setText("狀態: 未設定")
                s['temp_status'].setStyleSheet("color: orange;")
            if current_addr is None and cur_txt == "空白":
                s['current_status'].setText("狀態: 未安裝")
                s['current_status'].setStyleSheet("color: gray;")
            elif current_addr is not None:
                s['current_status'].setText("狀態: 已連線")
                s['current_status'].setStyleSheet("color: green;")
            else:
                s['current_status'].setText("狀態: 未設定")
                s['current_status'].setStyleSheet("color: orange;")
        else:
            if master:
                try:
                    master.close()
                except Exception:
                    pass
            # 不再提供模擬 fallback，直接標示離線
            s['master'] = None
            s['port'] = None
            s['data_source_label'].setText("資料來源: 離線")
            s['data_source_label'].setStyleSheet("color: red;")
            s['temp_status'].setText("狀態: 離線")
            s['temp_status'].setStyleSheet("color: red;")
            s['current_status'].setText("狀態: " + ("未安裝" if cur_txt == "空白" else "離線"))
            s['current_status'].setStyleSheet("color: red;")
        return ok

    def _read_section_data(self, s, t):
        """只讀真實設備，失敗即回傳 (None, None)，不再模擬"""
        master = s.get('master')
        if not master:
            return None, None
        temp = None
        current = None
        temp_txt = s['temp_addr'].currentText()
        cur_txt = s['current_addr'].currentText()
        try:
            temp_addr = int(temp_txt) if temp_txt and temp_txt.isdigit() else None
        except Exception:
            temp_addr = None
        try:
            current_addr = int(cur_txt) if cur_txt and cur_txt.isdigit() else None
        except Exception:
            current_addr = None
        try:
            if temp_addr is not None:
                r = master.execute(temp_addr, 3, 0, 1)
                v = r[0] if isinstance(r, (list, tuple)) and r else 0
                temp = float(v) / 10.0
            if current_addr is not None:
                r2 = master.execute(current_addr, 3, 0, 1)
                v2 = r2[0] if isinstance(r2, (list, tuple)) and r2 else 0
                current = float(v2)
            return temp, current
        except Exception:
            # 連線異常，不回傳模擬值
            return None, None

    def _reader_loop(self, s):
        master = s.get('master')
        temp_addr = s.get('temp_addr_value')
        current_addr = s.get('current_addr_value')
        ev = s.get('reader_stop')
        while master and ev and not ev.is_set():
            temp = None
            current = None
            try:
                if temp_addr is not None:
                    r = master.execute(temp_addr, 3, 0, 1)
                    v = r[0] if isinstance(r, (list, tuple)) and r else 0
                    temp = float(v) / 10.0
                if current_addr is not None:
                    r2 = master.execute(current_addr, 3, 0, 1)
                    v2 = r2[0] if isinstance(r2, (list, tuple)) and r2 else 0
                    current = float(v2)
            except Exception:
                temp = None
                current = None
            s['latest_temp'] = temp
            s['latest_current'] = current
            try:
                time.sleep(max(0.05, float(s.get('interval_ms', 500))/1000.0))
            except Exception:
                time.sleep(0.2)

    def _update_save_labels(self):
        self.save_path_label.setText(self.save_path)
        self.save_format_label.setText(self.save_format)

    def _change_storage_settings(self):
        dlg = StorageSettingsDialog(self)
        dlg.prefill_from_settings()
        if dlg.exec_() == QDialog.Accepted:
            s = QSettings('CoCoaLinna','NewEMS')
            self.save_path = s.value('save_path', self.save_path)
            self.save_format = s.value('save_format', self.save_format)
            self._update_save_labels()

    def _ensure_settings_valid(self):
        if not self.save_path or not _is_writable_dir(self.save_path):
            QMessageBox.warning(self, '儲存設定', '請先設定有效的儲存位置')
            self._change_storage_settings()
            return bool(self.save_path and _is_writable_dir(self.save_path))
        du = shutil.disk_usage(self.save_path)
        if du.free < 1024*1024:
            QMessageBox.warning(self, '磁碟空間不足', '可用空間不足，請變更儲存位置')
            self._change_storage_settings()
            return du.free >= 1024*1024
        return True
    
    def resizeEvent(self, event):
        """處理窗口大小改變事件，更新所有圖表高度"""
        super().resizeEvent(event)
        # 延遲更新，等待佈局完成
        QTimer.singleShot(100, self._update_all_chart_heights)
    
    def _update_all_chart_heights(self):
        """更新所有圖表高度"""
        # 遍歷所有生產線區域，更新圖表高度
        for s in self.sections:
            chart_widget = s.get('plot')
            if chart_widget and hasattr(chart_widget, '_update_height'):
                chart_widget._update_height()
        self._apply_equal_widths()

    def _init_db(self):
        try:
            self.db_conn = sqlite3.connect(self.db_path)
            _ensure_db(self.db_conn)
        except Exception:
            self.db_conn = None

    def _insert_record(self, ts, data):
        try:
            if not self.db_conn:
                self._init_db()
            if not self.db_conn:
                return
            cur = self.db_conn.cursor()
            cur.execute('INSERT INTO records (ts, data) VALUES (?, ?)', (ts, data))
            self.db_conn.commit()
        except Exception:
            pass

    def closeEvent(self, event):
        if self._ensure_settings_valid():
            saved_any = False
            for s in self.sections:
                if s['records']:
                    self._export_section(s)
                    saved_any = True
            if not saved_any:
                event.accept()
                return
        try:
            if self.db_conn:
                self.db_conn.close()
        except Exception:
            pass
        try:
            if os.path.exists(self.recording_flag_path):
                os.remove(self.recording_flag_path)
        except Exception:
            pass
        try:
            self._cleanup_realtime_temp()
        except Exception:
            pass
        event.accept()

    def _touch_recording_flag(self):
        try:
            with open(self.recording_flag_path, 'w', encoding='utf-8') as f:
                f.write('1')
        except Exception:
            pass

    def _any_recording_active(self):
        try:
            for s in self.sections:
                if s['timer'].isActive():
                    return True
            return False
        except Exception:
            return False

    def _set_recording_active(self, active):
        try:
            if active:
                self._touch_recording_flag()
            else:
                if os.path.exists(self.recording_flag_path):
                    os.remove(self.recording_flag_path)
        except Exception:
            pass

    def _cleanup_realtime_temp(self):
        try:
            for name in os.listdir(self.var_dir):
                p = os.path.join(self.var_dir, name)
                try:
                    if os.path.isfile(p):
                        os.remove(p)
                except Exception:
                    pass
        except Exception:
            pass

class LinePlot(QWidget):
    def __init__(self, color=QColor(0,0,0), parent=None):
        super().__init__(parent)
        self.color = color
        self.data = []
        self.max_points = 120
        self.auto_scale = True
        self.fixed_min = 0.0
        self.fixed_max = 100.0
        self.show_grid = True
    
    def _delayed_init_height(self):
        """延遲初始化高度（在對象創建後調用）"""
        self._update_height()
    
    def _update_height(self):
        """根據主窗口高度動態計算圖表高度"""
        try:
            # 獲取頂級窗口部件
            top_widget = self.window()
            
            if top_widget and hasattr(top_widget, 'height'):
                # 獲取主窗口高度並計算比例
                main_height = top_widget.height()
                # 使用比例計算：約佔主窗口高度的 20-22%
                chart_height = max(200, int(main_height * 0.3))
                self.setMinimumHeight(chart_height)
            elif self.parent():
                # 後備方案：使用直接父容器
                parent_height = self.parent().height()
                chart_height = max(200, int(parent_height * 0.42))
                self.setMinimumHeight(chart_height)
            else:
                self.setMinimumHeight(200)  # 默認高度
        except Exception as e:
            print(f"高度計算錯誤: {e}")
            self.setMinimumHeight(200)  # 錯誤時的默認值
    
    def resizeEvent(self, event):
        """處理窗口大小改變事件"""
        super().resizeEvent(event)
        self._update_height()

    def append(self, v):
        self.data.append(float(v))
        if len(self.data) > self.max_points:
            self.data.pop(0)
        self.update()

    def clear(self):
        self.data = []
        self.update()

    def setMaxPoints(self, n):
        self.max_points = int(n)
        if len(self.data) > self.max_points:
            self.data = self.data[-self.max_points:]
        self.update()

    def setAutoScale(self, flag):
        self.auto_scale = bool(flag)
        self.update()
    
    def setFixedRange(self, vmin, vmax):
        self.fixed_min = float(vmin)
        self.fixed_max = float(vmax)
        # 立即重繪，讓參數調整即時反映
        self.update()

    def setShowGrid(self, flag):
        self.show_grid = bool(flag)
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)  # 启用抗锯齿
        p.fillRect(self.rect(), QColor(250,250,250))  # 更浅的背景色
        
        pen_axis = QPen(QColor(200,200,200))  # 更淡的轴线颜色
        pen_axis.setWidth(1)
        p.setPen(pen_axis)
        
        r = self.rect()
        m = 8  # 减少边距，压缩空间
        chart_rect = r.adjusted(m, m, -m, -m)
        p.drawRect(chart_rect)
        
        if not self.data:
            return
            
        vals = self.data
        x0 = chart_rect.left()
        y0 = chart_rect.top()
        w = chart_rect.width()
        h = chart_rect.height()
        # 設置Y軸尺度為0-100
        vmin = 0.0
        vmax = 100.0
        scale_x = w / max(1, len(vals)-1)
        pen = QPen(self.color)
        pen.setWidth(2)  # 线条宽度
        pen.setCapStyle(Qt.RoundCap)  # 圆角线帽
        pen.setJoinStyle(Qt.RoundJoin)  # 圆角连接
        p.setPen(pen)
        
        pts = []
        for i, v in enumerate(vals):
            x = x0 + i * scale_x
            y = y0 + h - (v - vmin) * h / (vmax - vmin)
            pts.append(QPointF(x, y))
        
        # 绘制平滑曲线
        if len(pts) > 1:
            p.drawPolyline(pts)

        if self.show_grid:
            gpen = QPen(QColor(230,230,230))  # 更淡的网格线
            gpen.setStyle(Qt.DotLine)
            p.setPen(gpen)
            
            # 绘制水平网格线
            for i in range(1, 5):
                yy = y0 + i * h / 5.0
                p.drawLine(QPointF(x0, yy), QPointF(x0 + w, yy))
            
            # 绘制垂直网格线
            for i in range(1, 8):
                xx = x0 + i * w / 8.0
                p.drawLine(QPointF(xx, y0), QPointF(xx, y0 + h))
            
            # 绘制数值标签
            label_pen = QPen(QColor(100,100,100))
            p.setPen(label_pen)
            label_font = QFont("Arial", 9)
            p.setFont(label_font)
            
            p.drawText(x0 + 6, y0 + 16, f"{vmax:.1f}")
            p.drawText(x0 + 6, y0 + h - 4, f"{vmin:.1f}")

class DualLinePlot(QWidget):
    def __init__(self, color1=QColor(220,20,60), color2=QColor(30,144,255), parent=None):
        super().__init__(parent)
        self.color1 = color1
        self.color2 = color2
        self.data1 = []
        self.data2 = []
        self.max_points = 120
        self.auto_scale = True
        self.fixed_min = 0.0
        self.fixed_max = 100.0
        self.show_grid = True
        # 延遲初始化高度，等待所有方法定義完成
    
    def append(self, v1, v2):
        self.data1.append(float(v1))
        self.data2.append(float(v2))
        if len(self.data1) > self.max_points:
            self.data1.pop(0)
            self.data2.pop(0)
        self.update()

    def clear(self):
        self.data1 = []
        self.data2 = []
        self.update()
    
    def _update_height(self):
        """根據主窗口高度動態計算圖表高度"""
        try:
            # 獲取頂級窗口部件
            top_widget = self.window()
            
            if top_widget and hasattr(top_widget, 'height'):
                # 獲取主窗口高度並計算比例
                main_height = top_widget.height()
                # 使用比例計算：約佔主窗口高度的 20-22%
                chart_height = max(200, int(main_height * 0.21))
                self.setMinimumHeight(chart_height)
            elif self.parent():
                # 後備方案：使用直接父容器
                parent_height = self.parent().height()
                chart_height = max(200, int(parent_height * 0.42))
                self.setMinimumHeight(chart_height)
            else:
                self.setMinimumHeight(200)  # 默認高度
        except Exception as e:
            print(f"高度計算錯誤: {e}")
            self.setMinimumHeight(200)  # 錯誤時的默認值
    
    def resizeEvent(self, event):
        """處理窗口大小改變事件"""
        super().resizeEvent(event)
        self._update_height()
    
    def _delayed_init_height(self):
        """延遲初始化高度（在對象創建後調用）"""
        self._update_height()

    def setMaxPoints(self, n):
        self.max_points = int(n)
        if len(self.data1) > self.max_points:
            self.data1 = self.data1[-self.max_points:]
            self.data2 = self.data2[-self.max_points:]
        self.update()

    def setAutoScale(self, flag):
        self.auto_scale = bool(flag)
        self.update()

    def setFixedRange(self, vmin, vmax):
        self.fixed_min = float(vmin)
        self.fixed_max = float(vmax)
        self.update()

    def setShowGrid(self, flag):
        self.show_grid = bool(flag)
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)  # 启用抗锯齿
        p.fillRect(self.rect(), QColor(250,250,250))  # 更浅的背景色
        
        pen_axis = QPen(QColor(200,200,200))  # 更淡的轴线颜色
        pen_axis.setWidth(1)
        p.setPen(pen_axis)
        
        r = self.rect()
        m = 8  # 减少边距，压缩空间
        chart_rect = r.adjusted(m, m, -m, -m)
        p.drawRect(chart_rect)
        
        if not self.data1:
            return
            
        x0 = chart_rect.left()
        y0 = chart_rect.top()
        w = chart_rect.width()
        h = chart_rect.height()
        vals1 = self.data1
        vals2 = self.data2 if self.data2 else [0]*len(vals1)
        # 單 Y 軸：合併兩組資料取極值
        all_vals = vals1 + vals2
        if self.auto_scale:
            if all_vals:
                vmin = min(all_vals)
                vmax = max(all_vals)
                delta = (vmax - vmin) * 0.05 or 1.0
                vmin -= delta
                vmax += delta
            else:
                vmin, vmax = 0.0, 100.0
        else:
            vmin, vmax = self.fixed_min, self.fixed_max
        # 兩條線共用同一尺度
        vmin1 = vmin2 = vmin
        vmax1 = vmax2 = vmax
        scale_x = w / max(1, len(vals1)-1)

        pen1 = QPen(self.color1); pen1.setWidth(2)
        pen1.setCapStyle(Qt.RoundCap)
        pen1.setJoinStyle(Qt.RoundJoin)
        
        pen2 = QPen(self.color2); pen2.setWidth(2)
        pen2.setCapStyle(Qt.RoundCap)
        pen2.setJoinStyle(Qt.RoundJoin)

        # draw grid
        if self.show_grid:
            gpen = QPen(QColor(230,230,230)); gpen.setStyle(Qt.DotLine)
            p.setPen(gpen)
            for i in range(1,5):
                yy = y0 + i * h / 5.0
                p.drawLine(QPointF(x0, yy), QPointF(x0 + w, yy))
            for i in range(1,8):
                xx = x0 + i * w / 8.0
                p.drawLine(QPointF(xx, y0), QPointF(xx, y0 + h))

        # 设置标签字体
        label_font = QFont("Arial", 9)
        p.setFont(label_font)
        label_pen = QPen(QColor(100,100,100))
        p.setPen(label_pen)

        # 單 Y 軸標籤（左側）
        p.drawText(x0 + 6, y0 + 16, f"{vmax1:.1f}")
        p.drawText(x0 + 6, y0 + h - 4, f"{vmin1:.1f}")

        # 繪製溫度線
        p.setPen(pen1)
        pts = []
        for i, v in enumerate(vals1):
            x = x0 + i * scale_x
            y = y0 + h - (v - vmin1) * h / (vmax1 - vmin1)
            pts.append(QPointF(x, y))
        if len(pts) > 1:
            p.drawPolyline(pts)

        # 繪製電流線（共用同一 Y 軸）
        p.setPen(pen2)
        pts2 = []
        for i, v in enumerate(vals2):
            x = x0 + i * scale_x
            y = y0 + h - (v - vmin2) * h / (vmax2 - vmin2)
            pts2.append(QPointF(x, y))
        if len(pts2) > 1:
            p.drawPolyline(pts2)

        
def _is_writable_dir(path):
    try:
        if not os.path.isdir(path):
            return False
        tf = os.path.join(path, f'.__test_write_{int(time.time()*1000)}')
        with open(tf, 'w', encoding='utf-8') as f:
            f.write('1')
        os.remove(tf)
        return True
    except Exception:
        return False

def _human_size(n):
    n = float(n)
    for u in ['B','KB','MB','GB','TB']:
        if n < 1024:
            return f'{n:.1f} {u}'
        n /= 1024
    return f'{n:.1f} PB'

def _default_save_path():
    p = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
    if not p:
        p = os.path.expanduser('~')
    return p

def _show_auto_close_message(parent, title, text, msec=3000):
    msg = QMessageBox(parent)
    msg.setWindowTitle(title)
    msg.setText(text)
    msg.setStandardButtons(QMessageBox.Ok)
    btn = msg.button(QMessageBox.Ok)
    btn.setText('關閉')
    QTimer.singleShot(msec, msg.accept)
    msg.exec_()

def _ensure_db(conn):
    try:
        cur = conn.cursor()
        cur.execute('CREATE TABLE IF NOT EXISTS records (id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT NOT NULL, data TEXT NOT NULL)')
        conn.commit()
    except Exception:
        pass


class StorageSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('儲存位置設定')
        v = QVBoxLayout()
        form = QFormLayout()
        self.path_label = QLabel()
        self.browse_btn = QPushButton('選擇...')
        hb = QHBoxLayout(); hb.addWidget(self.path_label); hb.addWidget(self.browse_btn)
        cont = QWidget(); cont.setLayout(hb)
        self.format_combo = QComboBox(); self.format_combo.addItems(['csv','txt'])
        form.addRow('儲存路徑', cont)
        form.addRow('儲存格式', self.format_combo)
        v.addLayout(form)
        self.buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        v.addWidget(self.buttons)
        self.setLayout(v)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        self.browse_btn.clicked.connect(self._browse)

    def prefill_from_settings(self):
        s = QSettings('CoCoaLinna','NewEMS')
        path = s.value('save_path', _default_save_path())
        fmt = s.value('save_format', 'csv')
        self.path_label.setText(path)
        self.format_combo.setCurrentText(fmt)

    def _browse(self):
        d = QFileDialog.getExistingDirectory(self, '選擇儲存資料夾', self.path_label.text() or _default_save_path())
        if d:
            self.path_label.setText(d)

    def accept(self):
        path = (self.path_label.text() or '').strip()
        fmt = self.format_combo.currentText()
        if not path:
            QMessageBox.warning(self, '設定無效', '請選擇儲存路徑')
            return
        if not _is_writable_dir(path):
            QMessageBox.critical(self, '不可寫入', '選擇的路徑不可寫入')
            return
        du = shutil.disk_usage(path)
        if du.free < 1024*1024:
            QMessageBox.warning(self, '磁碟空間不足', '可用空間不足，請選擇其他位置')
            return
        s = QSettings('CoCoaLinna','NewEMS')
        s.setValue('save_path', path)
        s.setValue('save_format', fmt)
        super().accept()

def scan_modbus_devices(baud=9600, timeout=1, ports=None, max_addr=32):
    active_ports = []
    print("開始掃描 Modbus 設備...")

    def probe(port):
        try:
            if modbus_rtu is not None and serial is not None:
                try:
                    ser = serial.Serial(port, baudrate=baud, bytesize=8, parity='N', stopbits=1, timeout=timeout)
                except Exception:
                    return
                try:
                    master = modbus_rtu.RtuMaster(ser)
                    master.set_timeout(timeout)
                    master.set_verbose(False)
                except Exception:
                    try:
                        ser.close()
                    except Exception:
                        pass
                    return
                ok = False
                for addr in range(1, max_addr+1):
                    try:
                        master.execute(addr, 3, 0, 1)
                        ok = True
                        break
                    except Exception:
                        pass
                try:
                    master.close()
                except Exception:
                    pass
                try:
                    ser.close()
                except Exception:
                    pass
                if ok:
                    active_ports.append(port)
                    print(f"[+] 發現 Modbus 設備: {port}")
        except Exception:
            pass

    if ports is None:
        ports_list = []
        if QSerialPortInfo is not None:
            ports_list = [p.systemLocation() or p.portName() for p in QSerialPortInfo.availablePorts()]
        else:
            import glob
            sysname = sys.platform
            patterns = []
            if sysname.startswith('darwin'):
                patterns = ['/dev/tty.*', '/dev/cu.*']
            elif sysname.startswith('linux'):
                patterns = ['/dev/ttyUSB*', '/dev/ttyACM*', '/dev/ttyS*']
            elif sysname.startswith('win'):
                ports_list = [f'COM{n}' for n in range(1, 33)]
            if not ports_list:
                for pat in patterns:
                    ports_list.extend(glob.glob(pat))
    else:
        ports_list = list(ports)

    threads = []
    for port in ports_list:
        t = threading.Thread(target=probe, args=(port,))
        threads.append(t)
        t.start()

    for t in threads:
        t.join()

    print("掃描完成。")
    return active_ports


def _icon_path():
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(__file__)
    p = os.path.join(base, 'cocoa-linna.ico')
    if os.path.exists(p):
        return p
    p2 = os.path.join(os.path.dirname(base), 'cocoa-linna.ico')
    return p2

if __name__ == "__main__":
    app = QApplication(sys.argv)
    dlg = StorageSettingsDialog()
    dlg.prefill_from_settings()
    if dlg.exec_() != QDialog.Accepted:
        sys.exit(0)
    window = BlandPage()
    ic = QIcon(_icon_path())
    app.setWindowIcon(ic)
    window.setWindowIcon(ic)
    window.show()
    sys.exit(app.exec_())
