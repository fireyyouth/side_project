import sys
import sqlite3
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QTabWidget,
    QTableWidget, QTableWidgetItem, QHBoxLayout, QLineEdit, QFormLayout,
    QDialog, QDialogButtonBox, QLabel, QSpinBox, QMessageBox, QDateTimeEdit, QComboBox, QHeaderView, QFrame
)
from PySide6.QtGui import QDoubleValidator
from PySide6.QtCore import Qt
from PySide6.QtCore import QDateTime
from collections import defaultdict
from functools import partial
import decimal

ITEMS_PER_PAGE = 10

conn = sqlite3.connect("ledger.db")
cursor = conn.cursor()

def init_db():
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS person (
            name TEXT NOT NULL,
            UNIQUE(name)
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS project (
            name TEXT NOT NULL,
            UNIQUE(name)
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS transfer (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            time TEXT NOT NULL,
            person TEXT,
            project TEXT,
            kind TEXT,
            amount TEXT
        )
    """)
    conn.commit()

def get_person():
    cursor.execute("SELECT name FROM person")
    return [item[0] for item in cursor.fetchall()]

def get_project():
    cursor.execute("SELECT name FROM project")
    return [item[0] for item in cursor.fetchall()]

def add_person(name):
    if name:
        cursor.execute("INSERT INTO person (name) VALUES (?)", (name,))
        conn.commit()

def add_project(name):
    if name:
        cursor.execute("INSERT INTO project (name) VALUES (?)", (name,))
        conn.commit()

def delete_person(person):
    cursor.execute("DELETE FROM person WHERE name=?", (person,))
    conn.commit()

def delete_project(project):
    cursor.execute("DELETE FROM project WHERE name=?", (project,))
    conn.commit()

class BalanceError(Exception):
    pass

def kind_sign(kind):
    return 1 if kind == '入账' else -1

def pre_check_balance(person, project, diff):
    cursor.execute("SELECT amount, kind from transfer WHERE person=? AND project=?", (person, project))
    balance = decimal.Decimal(0)
    for (amount, kind) in cursor.fetchall():
        balance += kind_sign(kind) * decimal.Decimal(amount)
    print('pre check balance', balance, diff, balance + diff)
    if balance + diff < 0:
        raise BalanceError(f'{person} 在 {project} 上的余额会变成 {balance + diff}')

def post_check_balance(person, project):
    cursor.execute("SELECT amount, kind from transfer WHERE person=? AND project=?", (person, project))
    balance = decimal.Decimal(0)
    for (amount, kind) in cursor.fetchall():
        balance += kind_sign(kind) * decimal.Decimal(amount)
    print('post check balance', balance)
    if balance < 0:
        raise BalanceError(f'{person} 在 {project} 上的余额会变成 {balance}')

def add_transfer(time, person, project, kind, amount):
    pre_check_balance(person, project, kind_sign(kind) * decimal.Decimal(amount))
    cursor.execute(
        "INSERT INTO transfer (time, person, project, kind, amount) VALUES (?, ?, ?, ?, ?)",
        (time, person, project, kind, amount))
    conn.commit()

def delete_transfer(id_):
    person, project, kind, amount = cursor.execute("SELECT person, project, kind, amount FROM transfer WHERE id = ?", (id_,)).fetchall()[0]
    pre_check_balance(person, project, -kind_sign(kind) * decimal.Decimal(amount))
    cursor.execute("DELETE FROM transfer WHERE id=?", (id_,))
    conn.commit()

def update_transfer(id_, time, person, project, kind, amount):
    old_person, old_project = cursor.execute("SELECT person, project FROM transfer WHERE id = ?", (id_,)).fetchall()[0]
    try:
        cursor.execute('BEGIN')
        cursor.execute("""
            UPDATE transfer
            SET time = ?, person = ?, project = ?, kind = ?, amount = ?
            WHERE id = ?
        """, (time, person, project, kind, amount, id_))
        for p in [old_person, person]:
            for j in [old_project, project]:
                post_check_balance(p, j)
            
        conn.commit()
    except Exception:
        conn.rollback()
        raise

def get_transfer():
    stmt = "SELECT id, time, person, project, kind, amount from transfer ORDER BY time DESC"
    print('sql', stmt)
    cursor.execute(stmt)
    return cursor.fetchall()

def filter_transfer(person, project, kind):
    stmt = 'SELECT id, time, person, project, kind, amount FROM transfer WHERE 1 = 1'
    if person:
        stmt += f' AND person = "{person}"'
    if project:
        stmt += f' AND project = "{project}"'
    if kind:
        stmt += f' AND kind = "{kind}"'
    stmt += ' ORDER BY time DESC'
    print('sql', stmt)
    cursor.execute(stmt)
    return cursor.fetchall()

    

def create_person_combo():
    combo = QComboBox()
    for name in get_person():
        combo.addItem(name)
    return combo

def create_project_combo():
    combo = QComboBox()
    for name in get_project():
        combo.addItem(name)
    return combo

def create_kind_combo(need_blank=False):
    combo = QComboBox()
    if need_blank:
        combo.addItem('')
    combo.addItems(["入账", "出账"])
    return combo

class EditDialog(QDialog):
    def __init__(self, id_, time, person, project, kind, amount):
        super().__init__()

        self.id_ = id_

        self.setWindowTitle("编辑")

        self.time_edit = QDateTimeEdit()
        self.amount_edit = QLineEdit()
        self.amount_edit.setValidator(QDoubleValidator(0.0, float('inf'), 2))
        self.person_combo = create_person_combo()
        self.project_combo = create_project_combo()
        self.kind_combo = create_kind_combo()

        self.time_edit.setDateTime(QDateTime.fromString(time, "yyyy-MM-dd HH:mm:ss"))
        self.amount_edit.setText(str(amount))
        self.person_combo.setCurrentText(person)
        self.project_combo.setCurrentText(project)
        self.kind_combo.setCurrentText(kind)

        form_layout = QVBoxLayout()
        form_layout.addWidget(QLabel("时间:"))
        form_layout.addWidget(self.time_edit)
        form_layout.addWidget(QLabel("人员:"))
        form_layout.addWidget(self.person_combo)
        form_layout.addWidget(QLabel("项目:"))
        form_layout.addWidget(self.project_combo)
        form_layout.addWidget(QLabel("类型:"))
        form_layout.addWidget(self.kind_combo)
        form_layout.addWidget(QLabel("金额:"))
        form_layout.addWidget(self.amount_edit)

        button_box = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.handle_save)
        button_box.rejected.connect(self.reject)

        layout = QVBoxLayout()
        layout.addLayout(form_layout)
        layout.addWidget(button_box)
        self.setLayout(layout)

    def handle_save(self):
        time = self.time_edit.dateTime().toString("yyyy-MM-dd HH:mm:ss")
        person = self.person_combo.currentText()
        project = self.project_combo.currentText()
        kind = self.kind_combo.currentText()
        amount = self.amount_edit.text()
        try:
            update_transfer(self.id_, time, person, project, kind, amount)
        except BalanceError as e:
            QMessageBox.warning(self, "错误", str(e))
            return
        self.accept()

class PersonTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        form = QFormLayout()
        self.person_name_input = QLineEdit()
        form.addRow("姓名:", self.person_name_input)
        add_btn = QPushButton("添加")
        add_btn.clicked.connect(self.handle_add)
        layout.addLayout(form)
        layout.addWidget(add_btn)
        self.person_table = QTableWidget()
        self.person_table.setColumnCount(2)
        self.person_table.setHorizontalHeaderLabels(["姓名", "操作"])
        self.person_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.person_table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.person_table)
        self.setLayout(layout)
        self.load()

    def load(self):
        rows = get_person()
        self.person_table.setRowCount(len(rows))
        for row, name in enumerate(rows):
            self.person_table.setItem(row, 0, QTableWidgetItem(name))
            btn = QPushButton("删除")
            btn.clicked.connect(partial(self.handle_delete, name))
            self.person_table.setCellWidget(row, 1, btn)

    def handle_delete(self, name):
        delete_person(name)
        self.load()

    def handle_add(self):
        name = self.person_name_input.text()
        if name:
            add_person(name)
            self.load()

class ProjectTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        form = QFormLayout()
        self.project_name_input = QLineEdit()
        form.addRow("项目名:", self.project_name_input)
        add_btn = QPushButton("添加")

        add_btn.clicked.connect(self.handle_add)
        layout.addLayout(form)
        layout.addWidget(add_btn)
        self.project_table = QTableWidget()
        self.project_table.setColumnCount(2)
        self.project_table.setHorizontalHeaderLabels(["项目名", "操作"])
        self.project_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.project_table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.project_table)
        self.setLayout(layout)
        self.load()

    def load(self):
        rows = get_project()
        self.project_table.setRowCount(len(rows))
        for row, name in enumerate(rows):
            self.project_table.setItem(row, 0, QTableWidgetItem(name))
            btn = QPushButton("删除")
            btn.clicked.connect(partial(self.handle_delete, name))
            self.project_table.setCellWidget(row, 1, btn)

    def handle_add(self):
        name = self.project_name_input.text()
        if name:
            add_project(name)
            self.load()

    def handle_delete(self, name):
        delete_project(name)
        self.load()

class TransferTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        form = QFormLayout()

        self.time_input = QDateTimeEdit(QDateTime.currentDateTime())
        self.time_input.setCalendarPopup(True)
        self.person_input = QComboBox()
        self.project_input = QComboBox()
        self.kind_input = create_kind_combo()
        self.amount_input = QLineEdit()
        self.amount_input.setValidator(QDoubleValidator(0.0, float('inf'), 2))

        form.addRow("时间:", self.time_input)
        form.addRow("人员:", self.person_input)
        form.addRow("项目:", self.project_input)
        form.addRow("类型:", self.kind_input)
        form.addRow("金额:", self.amount_input)

        add_btn = QPushButton("添加转账")

        def handle_add():
            time = self.time_input.dateTime().toString("yyyy-MM-dd HH:mm:ss")
            person = self.person_input.currentText()
            project = self.project_input.currentText()
            kind = self.kind_input.currentText()
            amount = self.amount_input.text()
            try:
                add_transfer(time, person, project, kind, amount)
            except BalanceError as e:
                QMessageBox.warning(self, "错误", str(e))
                return
            self.load()
        add_btn.clicked.connect(handle_add)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)

        filter_bar = QHBoxLayout()
        self.person_filter = QLineEdit()
        self.project_filter = QLineEdit()
        self.kind_filter = create_kind_combo(True)

        filter_btn = QPushButton("过滤")

        self.filters = ("", "", "")
        self.filters_label = QLabel()
        def handle_filter():
            self.filters = (self.person_filter.text(), self.project_filter.text(), self.kind_filter.currentText())
            self.load()
        filter_btn.clicked.connect(handle_filter)

        filter_bar.addWidget(QLabel("人员:"))
        filter_bar.addWidget(self.person_filter)
        filter_bar.addWidget(QLabel("项目:"))
        filter_bar.addWidget(self.project_filter)
        filter_bar.addWidget(QLabel("类型:"))
        filter_bar.addWidget(self.kind_filter)
        filter_bar.addWidget(filter_btn)

        layout.addLayout(form)
        layout.addWidget(add_btn)
        layout.addWidget(line)
        layout.addLayout(filter_bar)
        layout.addWidget(self.filters_label)

        self.transfer_table = QTableWidget()
        self.transfer_table.setColumnCount(6)
        self.transfer_table.setHorizontalHeaderLabels(["时间", "人员", "项目", "类型", "金额", "操作"])
        self.transfer_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    
        layout.addWidget(self.transfer_table)
        self.setLayout(layout)
        self.load()

    def load_combo(self):
        self.person_input.clear()
        for name in get_person():
            self.person_input.addItem(name)
        self.project_input.clear()
        for name in get_project():
            self.project_input.addItem(name)

    def load(self):
        self.load_combo()

        TIME_CELL = 0
        PERSON_CELL = 1
        PROJECT_CELL = 2
        KIND_CELL = 3
        AMOUNT_CELL = 4
        ACTION_CELL = 5

        rows = filter_transfer(*self.filters)

        if self.filters != ("", "", ""):
            self.filters_label.setText(f"{len(rows)} 条结果, 过滤条件: {','.join([x for x in self.filters if x])}")
        else:
            self.filters_label.setText(f"{len(rows)} 条结果, 未过滤")

        self.transfer_table.setRowCount(len(rows))

        for row, (id_, time, person, project, kind, amount) in enumerate(rows):
            action_cell_widget = QWidget()
            layout = QHBoxLayout(action_cell_widget)

            edit_btn = QPushButton('编辑')
            delete_btn = QPushButton("删除")

            for btn in [edit_btn, delete_btn]:
                btn.setFixedHeight(20)

            edit_btn.clicked.connect(partial(self.handle_edit, id_, time, person, project, kind, amount))
            delete_btn.clicked.connect(partial(self.handle_delete, id_))

            layout.addWidget(edit_btn)
            layout.addWidget(delete_btn)

            self.transfer_table.setItem(row, TIME_CELL, QTableWidgetItem(time))
            self.transfer_table.setItem(row, PERSON_CELL, QTableWidgetItem(person or ""))
            self.transfer_table.setItem(row, PROJECT_CELL, QTableWidgetItem(project or ""))
            self.transfer_table.setItem(row, KIND_CELL, QTableWidgetItem(kind or ""))
            self.transfer_table.setItem(row, AMOUNT_CELL, QTableWidgetItem(str(amount)))
            self.transfer_table.setCellWidget(row, ACTION_CELL, action_cell_widget)

            self.transfer_table.setEditTriggers(QTableWidget.NoEditTriggers)

    def handle_edit(self, id_, time, person, project, kind, amount):
        dialog = EditDialog(id_, time, person, project, kind, amount)
        if dialog.exec() == QDialog.Accepted:
            self.load()

    def handle_delete(self, id_):
        try:
            delete_transfer(id_)
        except BalanceError as e:
            QMessageBox.warning(self, "错误", str(e))
            return
        self.load()

class SummaryTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        self.summary_table = QTableWidget()
        self.summary_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.summary_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.summary_table)
        self.setLayout(layout)

    def load(self):
        transfer_list = get_transfer()
        person_summary = defaultdict(lambda: defaultdict(decimal.Decimal))
        project_summary = defaultdict(decimal.Decimal)
        summary = defaultdict(decimal.Decimal)
        for row, (id_, time, person, project, kind, amount) in enumerate(transfer_list):
            amount = decimal.Decimal(amount)
            summary[(person, kind, project)] += amount
            person_summary[person][kind] += amount
            project_summary[project] += kind_sign(kind) * amount

        self.summary_table.clear()
        self.summary_table.clearSpans()

        project_order = sorted(project_summary.keys())
        person_order = sorted(person_summary.keys())

        horizontal_headers = project_order + ['人员统计', '人员合计']
        vertial_headers = []
        for person in person_order:
            vertial_headers.append(f'{person} 入')
            vertial_headers.append(f'出')
        vertial_headers.append('项目合计')

        self.summary_table.setRowCount(len(vertial_headers))
        self.summary_table.setColumnCount(len(horizontal_headers))
        self.summary_table.setVerticalHeaderLabels(vertial_headers)
        self.summary_table.setHorizontalHeaderLabels(horizontal_headers)            

        for col in range(self.summary_table.rowCount()):
            self.summary_table.verticalHeaderItem(col).setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)


        for col, project in enumerate(project_order):
            for x, person in enumerate(person_order):
                self.summary_table.setItem(x * 2, col, QTableWidgetItem(str(summary[(person, "入账", project)])))
                self.summary_table.setItem(x * 2 + 1, col, QTableWidgetItem(str(summary[(person, "出账", project)])))

        for x, person in enumerate(person_order):
            income = person_summary[person]["入账"]
            spend = person_summary[person]["出账"]
            self.summary_table.setItem(x * 2, len(project_summary), QTableWidgetItem(str(income)))
            self.summary_table.setItem(x * 2 + 1, len(project_summary), QTableWidgetItem(str(spend)))
            self.summary_table.setItem(x * 2, len(project_summary) + 1, QTableWidgetItem(str(income - spend)))
            self.summary_table.setSpan(x * 2, len(project_summary) + 1, 2, 1)
            print('set span', x * 2, len(project_summary) + 1, 2, 1)

        for col, project in enumerate(project_order):
            self.summary_table.setItem(2 * len(person_summary), col, QTableWidgetItem(str(project_summary[project])))

        self.summary_table.setItem(len(vertial_headers) - 1, len(project_summary), QTableWidgetItem(str(sum(project_summary.values()))))
        self.summary_table.setSpan(len(vertial_headers) - 1, len(project_summary), 1, 2)
        print('set span', len(vertial_headers) - 1, len(project_summary), 1, 2)

class LedgerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setFixedSize(1000, 800)
        self.setWindowTitle("账本应用")

        layout = QVBoxLayout()

        self.tabs = QTabWidget()
        self.tabs.addTab(PersonTab(), "人员")
        self.tabs.addTab(ProjectTab(), "项目")
        self.tabs.addTab(TransferTab(), "流水")
        self.tabs.addTab(SummaryTab(), "统计")

        self.tabs.currentChanged.connect(self.on_tab_changed)

        layout.addWidget(self.tabs)
        self.setLayout(layout)

    def on_tab_changed(self, index):
        print('tab change', self.tabs.tabText(index))
        self.tabs.widget(index).load()

if __name__ == "__main__":
    init_db()
    app = QApplication(sys.argv)
    window = LedgerApp()
    window.show()
    sys.exit(app.exec())
