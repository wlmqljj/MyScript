import subprocess
from PyQt5.QtWidgets import QApplication, QWidget, QTableWidget, QTableWidgetItem, QPushButton, QVBoxLayout, QHeaderView
from PyQt5.QtGui import QResizeEvent

class AdbDeviceTable(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('ADB Toolkit')
        self.devices = []
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['Serial Number', 'Product', 'Transport ID'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.cellClicked.connect(self.on_cell_click)
        self.refresh_button = QPushButton('Refresh')
        self.refresh_button.clicked.connect(self.refresh_devices)
        self.shell_button = QPushButton('ADB Activate Shizuku')
        self.shell_button.clicked.connect(self.execute_adb_shell)
        self.scrcpy_button = QPushButton('Scrcpy')
        self.scrcpy_button.clicked.connect(self.execute_scrcpy)
        layout = QVBoxLayout()
        layout.addWidget(self.table)
        layout.addWidget(self.refresh_button)
        layout.addWidget(self.shell_button)
        layout.addWidget(self.scrcpy_button)
        self.setLayout(layout)
        self.refresh_devices()

    def refresh_devices(self):
        self.devices = self._get_connected_devices()
        self.table.setRowCount(len(self.devices))
        for i, device in enumerate(self.devices):
            items = [QTableWidgetItem(device['serial']), QTableWidgetItem(device['product']), QTableWidgetItem(device['transport_id'])]
            for j, item in enumerate(items):
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(i, j, item)

    def on_cell_click(self, row, column):
        self.selected_device = self.devices[row]

    def execute_adb_shell(self):
        if hasattr(self, 'selected_device'):
            serial = self.selected_device['serial']
            subprocess.call(['powershell.exe', f'adb -s {serial} shell sh /storage/emulated/0/Android/data/moe.shizuku.privileged.api/start.sh'])

    def execute_scrcpy(self):
        if hasattr(self, 'selected_device'):
            serial = self.selected_device['serial']
            subprocess.call(['powershell.exe', f'./scrcpy -s {serial}'])

    @staticmethod
    def _get_connected_devices():
        output = subprocess.check_output(['adb', 'devices', '-l']).decode()
        devices = []
        for line in output.splitlines()[1:]:
            s = line.strip().split()
            if len(s) < 4:
                continue
            devices.append({
                'serial': s[0],
                'product': s[3],
                'transport_id': s[5].split(':')[1] if len(s) > 4 else '',
            })
        return devices

if __name__ == '__main__':
    import sys
    from PyQt5.QtCore import Qt

    app = QApplication(sys.argv)
    table = AdbDeviceTable()
    table.show()
    sys.exit(app.exec_())
