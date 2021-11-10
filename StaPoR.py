import os
import sys
import matplotlib.pyplot as plt
import openpyxl
import sqlite3
import csv
import shutil
from PyQt5 import uic
from os import listdir, getcwd
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QApplication, QMainWindow, QDialog, QTableWidgetItem, QWidget
from PyQt5.QtWidgets import QLabel, QFileDialog, QMessageBox, QInputDialog


class MainForm(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('main_window.ui', self)
        for i in listdir('db'):
            self.comboBox_3.addItem(str(i).split('.')[0])
        self.new_table.clicked.connect(self.create_table)
        self.show_progress.clicked.connect(self.create_progress)
        self.pushButton_7.clicked.connect(self.stat_for_work)
        self.pushButton_3.clicked.connect(self.open_table)
        self.pushButton.clicked.connect(self.open_without_interface)
        self.pushButton_4.clicked.connect(self.del_table)
        self.pushButton_5.clicked.connect(self.person_work)
        self.pushButton_2.clicked.connect(self.add_person)
        self.pushButton_6.clicked.connect(self.perenesti_db)
        self.comboBox_3.currentIndexChanged.connect(self.change_class)
        self.pushButton_8.clicked.connect(self.del_person)
        self.pushButton_9.clicked.connect(self.export_in_csv)
        self.pushButton_10.clicked.connect(self.close)
        self.pushButton_11.clicked.connect(self.del_class)
        if self.comboBox_3.itemText(0):
            self.change_class()

    def del_class(self):
        if self.comboBox_3.itemText(0) and not self.comboBox_3.itemText(1):
            res = QMessageBox.warning(self, self.windowTitle(),
                                      ("У вас имеется только один класс\n"
                                       "Вы действительно хотите его удалить?"),
                                      QMessageBox.Yes | QMessageBox.No)
            if res == QMessageBox.No:
                return
        name, ok_pressed = QInputDialog.getText(self, "Удалить класс",
                                            "Введите его название")
        if ok_pressed:
            path = getcwd()
            c = 0
            fl = False
            for i in listdir('db'):
                if i == name + '.db':
                    fl = True
                    break
                c += 1
            if not fl:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText("Error")
                msg.setInformativeText('Такой класс не существует')
                msg.setWindowTitle("Error")
                msg.exec_()
                return
            self.comboBox_3.removeItem(c)
            os.remove(f'{path}/db/{name}.db')

    def export_in_csv(self):
        name = QFileDialog.getSaveFileName(self, 'Save File')
        if not name[0]:
            return
        path = name[0] + '.csv'
        x = self.data.cursor().execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
        x = [i[0] for i in x]
        del x[x.index('journal')]
        del x[x.index('pupils')]
        x.append('Журнал')
        spisok = tuple(x)
        name, ok_pressed = QInputDialog.getItem(
            self, "Выберите таблицу", "Какую таблицу экспортировать?",
            spisok, 1, False)
        if ok_pressed:
            if name == 'Журнал':
                name = 'journal'
            header = self.data.cursor().execute(f'pragma table_info({name})').fetchall()
            header = [i[1] for i in header]
            del header[0]
            header.insert(0, '№')
            content = self.data.cursor().execute(f'SELECT * FROM {name}').fetchall()
            content = [list(i) for i in content]
            pupils = self.data.cursor().execute(f'SELECT title FROM pupils').fetchall()
            pupils = [i[0] for i in pupils]
            with open(path, 'w', newline='', encoding='utf-8') as file:
                filewriter = csv.writer(file, delimiter=';')
                filewriter.writerow(header)
                for i in range(len(content)):
                    row = content[i]
                    row.insert(2, pupils[row[1] - 1])
                    del row[1]
                    filewriter.writerow(row)
                file.close()

    def del_person(self):
        name, ok_pressed = QInputDialog.getText(self, "Удалить ученика",
                                                 "Введите имя ученика")
        if ok_pressed:
            id = self.data.cursor().execute("SELECT id FROM pupils WHERE title = ?", (name, )).fetchall()
            if not id:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText("Error")
                msg.setInformativeText('Такого ученика не существует')
                msg.setWindowTitle("Error")
                msg.exec_()
                return
            id = id[0][0]
            self.comboBox.removeItem(id - 1)
            self.data.cursor().execute('DELETE FROM pupils WHERE id = ?', (id, ))
            self.data.commit()
            self.data.cursor().execute("""
            UPDATE pupils
            SET id = id - 1
            WHERE id > 3""")
            self.data.commit()
            self.data.cursor().execute('DELETE FROM journal WHERE id = ?', (id, ))
            self.data.commit()
            self.data.cursor().execute("""
                        UPDATE journal
                        SET ФИО = ФИО - 1
                        WHERE id > 3""")
            self.data.commit()
            self.data.cursor().execute("""
                        UPDATE journal
                        SET id = id - 1
                        WHERE id > 3""")
            self.data.commit()
            self.main_table()

    def change_class(self):
        self.data = sqlite3.connect(f'db/{self.comboBox_3.currentText()}.db')
        names = self.data.cursor().execute("SELECT title FROM pupils").fetchall()
        self.comboBox.clear()
        for i, name in enumerate(names):
            self.comboBox.addItem(name[0])
        x = self.data.cursor().execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
        x = [i[0] for i in x]
        del x[x.index('journal')]
        del x[x.index('pupils')]
        self.comboBox_2.clear()
        for i in x:
            self.comboBox_2.addItem(i)
        self.main_table()

    def perenesti_db(self):
        fname = QFileDialog.getOpenFileName(
            self, 'Выбрать таблицу', '',
            'Таблица (*.db);')[0]
        if not fname:
            return
        name = fname.split('/')[-1]
        path = getcwd()
        shutil.move(fname, f'{path}\db\{name}')
        self.data = sqlite3.connect(f'{path}/db/{name}')
        x = self.data.cursor().execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
        x = [i[0] for i in x]
        if 'journal' not in x:
            self.data.cursor().execute(f"""CREATE TABLE journal (
                id INT PRIMARY KEY,
                ФИО INT REFERENCES pupils (id))""")
            self.data.commit()
            ids = self.data.cursor().execute('SELECT id FROM pupils').fetchall()
            for i in [_[0] for _ in ids]:
                self.data.cursor().execute(f'INSERT INTO journal VALUES({i}, {i})')
                self.data.commit()
        self.comboBox_3.addItem(name.split('.')[0])

    def add_person(self):
        name, ok_pressed = QInputDialog.getText(self, "Добавить ученика",
                                                 "Как его зовут?")
        n = len(self.data.cursor().execute("SELECT * FROM journal").fetchall())
        if ok_pressed:
            self.data.cursor().execute(f'INSERT INTO pupils(title) VALUES("{name}")')
            self.data.commit()
            self.comboBox.addItem(name)
            works = self.data.cursor().execute(f'PRAGMA table_info(journal)').fetchall()[2:]
            works = [-1 for _ in works]
            name = tuple([n + 1, n + 1] + works)
            stroka = "INSERT INTO journal VALUES("
            stroka += ', '.join('?' for _ in range(len(name)))
            stroka += ')'
            self.data.cursor().execute(stroka, name)
            self.data.commit()
            self.main_table()

    def person_work(self):
        try:
            name = self.comboBox.currentText()
            name = self.data.cursor().execute(f"SELECT id FROM pupils WHERE title = '{name}'").fetchall()[0][0]
            y = self.data.cursor().execute(f'PRAGMA table_info({self.comboBox_2.currentText()})').fetchall()
            x = self.data.cursor().execute(f"SELECT * FROM {self.comboBox_2.currentText()}"
                                           f"   WHERE ФИО = {name}").fetchall()
            n = 3
            y = [i[1] for i in y]
            x = [i for i in x[0]]
            if 'Вариант' in y:
                n += 1
            y = y[n:]
            x = x[n:]
            x_and_y = dict()
            for i in range(len(y)):
                if x[i] not in x_and_y.keys():
                    x_and_y[x[i]] = [y[i]]
                else:
                    x_and_y[x[i]] += [y[i]]
            x = x_and_y.keys()
            y = x_and_y.values()
            sup_x = [len(i) for i in y]
            sup_y = [', '.join(_ for _ in __) for __ in y]
            fig, ax = plt.subplots()
            ax.pie(sup_x, labels=sup_y)
            plt.savefig('pict/pie1.jpg')
        except IndexError:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Ученик не присутствовал на данной работе')
            msg.setWindowTitle("Error")
            msg.exec_()
            return
        self.form1 = ShowStat(list(y), list(x), self.comboBox.currentText(), 'pie1.jpg')
        self.form1.show()

    def del_table(self):
        extra, ok_pressed = QInputDialog.getText(self, "Удаление работы",
                                                "Введите название работы")
        extra = '⠀'.join(i for i in extra.split(' '))
        if ok_pressed:
            names_col = self.data.cursor().execute(f'PRAGMA table_info(journal)').fetchall()
            names_col = [i[1] for i in names_col]
            if extra not in names_col:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText("Error")
                msg.setInformativeText('Такой работы не существует')
                msg.setWindowTitle("Error")
                msg.exec_()
                return
            self.comboBox_2.removeItem(names_col.index(extra) - 2)
            del names_col[names_col.index(extra)]
            query = 'SELECT '
            query += ','.join(_ for _ in names_col)
            query += ' FROM journal'
            content = self.data.cursor().execute(query).fetchall()
            sqlite_create_table_query = f"""CREATE TABLE journal (
                    id INT PRIMARY KEY,
                    ФИО INT REFERENCES pupils (id)"""
            for i in names_col[2:]:
                sqlite_create_table_query += f",\n\t{i} INT"
            sqlite_create_table_query += '\n);'
            self.data.cursor().execute(f'DROP TABLE IF EXISTS {extra}')
            self.data.cursor().execute(f'DROP TABLE IF EXISTS journal')
            self.data.cursor().execute(sqlite_create_table_query)
            self.data.commit()
            stroka = "INSERT INTO journal VALUES("
            stroka += ', '.join('?' for _ in range(len(names_col)))
            stroka += ')'
            for i in content:
                self.data.cursor().execute(stroka, i)
                self.data.commit()
            self.main_table()

    def open_without_interface(self):
        try:
            fname = QFileDialog.getOpenFileName(
                self, 'Выбрать таблицу', '',
                'Таблица (*.xlsx);')[0]
            if not fname:
                return
            self.form1 = AddTable(fname)
            self.form1.progress.hide()
            self.form1.save_table()
        except openpyxl.utils.exceptions.InvalidFileException:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Неверный формат')
            msg.setWindowTitle("Error")
            self.data.cursor().execute(f'DROP TABLE IF EXISTS {self.form1.lineEdit.text()}')
            msg.exec_()
        except ValueError:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Такая таблица уже существует')
            msg.setWindowTitle("Error")
            msg.exec_()
        except Exception:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Неверно заполнена таблица')
            msg.setWindowTitle("Error")
            self.data.cursor().execute(f'DROP TABLE IF EXISTS {self.form1.lineEdit.text()}')
            msg.exec_()

    def open_table(self):
        fname = QFileDialog.getOpenFileName(
            self, 'Выбрать таблицу', '',
            'Таблица (*.xlsx);')[0]
        if not fname:
            return
        self.form1 = AddTable(fname)
        self.form1.show()

    def main_table(self):
        header = self.data.cursor().execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
        header = [i[0] for i in header]
        del header[header.index('journal')]
        del header[header.index('pupils')]
        header.insert(0, 'ФИО')
        grade = self.data.cursor().execute("SELECT * FROM journal").fetchall()
        name = self.data.cursor().execute("SELECT title from pupils").fetchall()
        self.tableWidget.clear()
        self.tableWidget.setColumnCount(len(header))
        self.tableWidget.setRowCount(0)
        self.tableWidget.setHorizontalHeaderLabels(header)
        for i, row in enumerate(grade):
            self.tableWidget.setRowCount(self.tableWidget.rowCount() + 1)
            for j, elem in enumerate(row[1:]):
                if j == 0:
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(name[i][0])))
                else:
                    self.tableWidget.setItem(i, j,  QTableWidgetItem(str(elem)))

    def stat_for_work(self):
        name = self.comboBox_2.currentText()
        if not name:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Не выбрана работа')
            msg.setWindowTitle("Error")
            msg.exec_()
            return
        nums = self.data.cursor().execute(f"SELECT {name} FROM journal").fetchall()
        x_and_y = dict()
        for i in list(set(nums)):
            if i not in x_and_y:
                x_and_y[i[0]] = 1
            else:
                x_and_y[i[0]] += 1
        x = x_and_y.keys()
        y = x_and_y.values()
        fig, ax = plt.subplots()
        ax.pie(y, labels=x)
        plt.savefig('pict/pie.jpg')
        self.form1 = ShowStat(list(x), list(y), name, 'pie.jpg')
        self.form1.show()

    def create_progress(self):
        name = self.comboBox.currentText()
        if not name:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Не выбран ученик')
            msg.setWindowTitle("Error")
            msg.exec_()
            return
        name = self.data.cursor().execute(f"SELECT id FROM pupils WHERE title = '{name}'").fetchall()
        x = self.data.cursor().execute("pragma table_info(journal)").fetchall()
        x = [i[1] for i in x[2:]]
        y = self.data.cursor().execute("SELECT * FROM journal WHERE ФИО = ?", (name[0][0],)).fetchall()
        y = list(y[0])[2:]
        fig, ax = plt.subplots()
        ax.plot(x, y)
        plt.savefig('pict/figure.jpg')
        self.form1 = ShowStat(x, y, self.comboBox.currentText(), 'figure.jpg')
        self.form1.show()

    def create_table(self):
        self.form1 = NewTable('Создание таблицы')
        self.form1.show()
        self.main_table()

    def closeEvent(self, event):
        self.data.close()


class AddTable(QWidget):
    def __init__(self, name):
        super().__init__()
        self.workbook = openpyxl.open(name)
        uic.loadUi('for_table.ui', self)
        self.label_2.hide()
        self.label_3.hide()
        self.label_4.hide()
        self.lineEdit_2.hide()
        self.lineEdit_3.hide()
        self.plainTextEdit.hide()
        self.pushButton_3.hide()
        self.setWindowTitle('Изменение открытой таблицы')
        self.lineEdit.setText(str(self.workbook.sheetnames[0]))
        self.pushButton.clicked.connect(self.save_table)
        self.pushButton_2.clicked.connect(self.close)
        self.show_table()

    def close(self):
        self.hide()

    def show_table(self):
        worksheet = self.workbook.worksheets[0]
        self.header = [i.value for i in list(worksheet.rows)[0]]
        self.tableWidget.clear()
        self.tableWidget.setColumnCount(len(self.header))
        if 'ФИО' not in self.header or 'Оценка' not in self.header:
            if self.progress.isHidden():
                raise IndexError
            else:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Critical)
                msg.setText("Error")
                msg.setInformativeText('Неверно заполнена таблица')
                msg.setWindowTitle("Error")
                msg.exec_()
                self.hide()
                return
        self.tableWidget.setRowCount(0)
        self.tableWidget.setHorizontalHeaderLabels(self.header)
        for i in range(worksheet.max_row - 1):
            self.tableWidget.setRowCount(self.tableWidget.rowCount() + 1)
            row = [i.value for i in list(worksheet.rows)[i + 1]]
            for j, elem in enumerate(row):
                self.tableWidget.setItem(i, j, QTableWidgetItem(str(elem)))

    def save_table(self):
        try:
            sqlite_connection = sqlite3.connect(f'db/{form.comboBox_3.currentText()}.db')
            extra = sqlite_connection.cursor().execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
            extra = [i[0] for i in extra]
            if self.lineEdit.text() in extra:
                raise ValueError
            sqlite_create_table_query = f"""CREATE TABLE {self.lineEdit.text()} (
        id INT PRIMARY KEY,
        ФИО INT REFERENCES pupils (id)"""
            for i in self.header[1:]:
                stroka = 'ᅠ'.join(_ for _ in i.split(' '))
                sqlite_create_table_query += f",\n\t{stroka} INT"
            sqlite_create_table_query += '\n);'
            cursor = sqlite_connection.cursor()
            cursor.execute(sqlite_create_table_query)
            sqlite_connection.commit()
            for i in range(self.tableWidget.rowCount()):
                sqlite_connection.commit()
                data = [i + 1]
                for j in range(len(self.header)):
                    if j == 0:
                        name = self.tableWidget.item(i, 0).text()
                        name = cursor.execute(f"SELECT id FROM pupils WHERE title = '{name}'").fetchall()
                        if not name:
                            continue
                        data.append(name[0][0])
                    else:
                        if not name:
                            continue
                        data.append(int(self.tableWidget.item(i, j).text()))
                if len(data) > 1:
                    stroka = "INSERT INTO "
                    stroka += self.lineEdit.text() + " VALUES(?, "
                    stroka += ', '.join('?' for _ in range(len(self.header)))
                    stroka += ')'
                    cursor.execute(stroka, data)
                    sqlite_connection.commit()
        except openpyxl.utils.exceptions.InvalidFileException:
            if self.progress.isHidden():
                raise openpyxl.utils.exceptions.InvalidFileException
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Неверный формат')
            msg.setWindowTitle("Error")
            sqlite_connection.cursor().execute(f'DROP TABLE IF EXISTS {self.lineEdit.text()}')
            msg.exec_()
        except sqlite3.OperationalError:
            if self.progress.isHidden():
                raise sqlite3.OperationalError
            self.progress.setText('Ошибка: в заполнении таблицы')
            self.progress.setStyleSheet('color: red')
            self.progress.adjustSize()
            sqlite_connection.cursor().execute(f'DROP TABLE IF EXISTS {self.lineEdit.text()}')
            return
        except ValueError:
            if self.progress.isHidden():
                raise ValueError
            self.progress.setText('Ошибка: такая таблица уже существует')
            self.progress.setStyleSheet('color: red')
            self.progress.adjustSize()
            sqlite_connection.cursor().execute(f'DROP TABLE IF EXISTS {self.lineEdit.text()}')
            return

        if ' ' in self.lineEdit.text():
            self.self.lineEdit.setText('⠀'.join(i for i in self.lineEdit.text().split()))
        cursor.execute(f"ALTER TABLE journal ADD COLUMN {self.lineEdit.text()} INT NOT NULL DEFAULT (-1);")
        sqlite_connection.commit()
        stolb = self.lineEdit.text()
        for i in range(self.tableWidget.rowCount()):
            ocenka = int(self.tableWidget.item(i, 1).text())
            name = self.tableWidget.item(i, 0).text()
            name = cursor.execute(f"SELECT id FROM pupils WHERE title = '{name}'").fetchall()
            if not name:
                continue
            cursor.execute(f"""UPDATE journal
                                SET {stolb} = ?
                                WHERE ФИО = ?""", (ocenka, name[0][0]))
            sqlite_connection.commit()
        sqlite_connection.close()
        self.progress.setText("Таблица успешно сохранена")
        self.progress.setStyleSheet('color: green')
        self.progress.adjustSize()
        form.comboBox_2.addItem(self.lineEdit.text())
        form.main_table()


class ShowStat(QWidget):
    def __init__(self, *args):
        super().__init__()
        self.initUI(args)

    def initUI(self, args):
        works = args[0]
        num = args[1]
        pict = args[-1]
        self.setGeometry(300, 300, 700, 600)
        self.label = QLabel(self)
        self.label.resize(640, 480)
        self.label.move(30, 20)
        self.label.setPixmap(QPixmap(f'pict/{pict}'))
        if args[-1] == 'figure.jpg':
            self.setWindowTitle(f'Успеваемость: {args[2]}')
            line = [(works[i], num[i]) for i in range(len(works))]
            num = [i for i in filter(lambda x: x >= 0, num)]
            bad = min(num)
            good = max(num)
            worst = [i[0] for i in filter(lambda x: x[1] == bad, line)]
            best = [i[0] for i in filter(lambda x: x[1] == good, line)]
            self.worst = QLabel(self)
            stroka = '\n'.join(i for i in worst)
            if len(worst) > 1:
                self.worst.setText(f"Хуже всего написаны работы:\n {stroka}")
            else:
                self.worst.setText(f"Хуже всего написана работа:\n {stroka}")
            self.worst.adjustSize()
            self.worst.move(20, 510)
            stroka = '\n'.join(i for i in best)
            self.best = QLabel(self)
            if len(worst) > 1:
                self.best.setText(f"Лучше всего написаны работы: \n {stroka}")
            else:
                self.best.setText(f"Лучше всего написана работа: \n {stroka}")
            self.best.adjustSize()
            self.best.move(500, 510)
        elif args[-1] == 'pie.jpg':
            self.setWindowTitle(f'Статисктика по {args[2]}')
            sqlite_connection = sqlite3.connect(f'db/{form.comboBox_3.currentText()}.db')
            x, y = 20, 510
            for i in range(len(works)):
                self.label1 = QLabel(self)
                self.label1.move(x, y)
                data = sqlite_connection.cursor().execute(f"SELECT title from pupils WHERE id IN "
                                                   f"(SELECT ФИО FROM journal WHERE {args[2]} = {works[i]})").fetchall()
                data = '\n'.join(i[0] for i in data)
                self.label1.setText(f"{works[i]}:\n {data}")
                self.label1.adjustSize()
                x += self.label1.width() + 20
            sqlite_connection.close()
        elif args[-1] == 'pie1.jpg':
            self.setWindowTitle(f'Баллы за задания в работе {form.comboBox_2.currentText()}: {args[2]}')
            x, y = 20, 510
            for i in range(len(works)):
                self.label1 = QLabel(self)
                self.label1.move(x, y)
                string = '\n'.join(_ for _ in works[i])
                self.label1.setText(f"{num[i]}:\n{string}")
                self.label1.adjustSize()
                self.label1.show()
                x += self.label1.width() + 20


class NewTable(QDialog):
    def __init__(self, *args):
        super().__init__()
        uic.loadUi('for_table.ui', self)
        self.setWindowTitle(args[0])
        self.pushButton_3.clicked.connect(self.show_table)
        self.pushButton.clicked.connect(self.save_table)
        self.pushButton_2.clicked.connect(self.close)

    def close(self):
        self.hide()

    def show_table(self):
        self.tableWidget.clear()
        self.header = ['ФИО', 'Оценка']
        self.db = sqlite3.connect(f'db/{form.comboBox_3.currentText()}.db')
        if self.lineEdit_2.text().isdigit():
            self.var = 0 if self.lineEdit_2.text() == '0' or self.lineEdit_2.text() == '1' else 1
        else:
            self.progress.setText('Ошибка: неправильно указано количество вариантов')
            self.progress.setStyleSheet('color: red')
            self.progress.adjustSize()
            return
        if self.lineEdit_3.text().isdigit():
            self.num = int(self.lineEdit_3.text())
        else:
            self.progress.setText('Ошибка: неправильно указано количество заданий')
            self.progress.setStyleSheet('color: red')
            self.progress.adjustSize()
            return
        if self.var:
            self.header.append('Вариант')
        names = self.db.cursor().execute("SELECT title FROM pupils").fetchall()
        self.row = 2 + self.var + self.num
        zadaniya = []
        if not self.plainTextEdit.toPlainText():
            for i in range(1, self.num + 1):
                zadaniya.append(f'Задание⠀{i}')
        else:
            for i in self.plainTextEdit.toPlainText().split(','):
                zadaniya.append(i)
        if len(zadaniya) != self.num:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setText("Внимание")
            msg.setInformativeText('Количество именовонных заданий не соответствует количстыу заданий')
            msg.setWindowTitle("ВНИМАНИЕ")
            msg.exec_()
            self.tableWidget.clear()
            return
        self.header += zadaniya
        self.tableWidget.setColumnCount(self.row)
        self.tableWidget.setRowCount(0)
        self.tableWidget.setHorizontalHeaderLabels(self.header)
        for i, row in enumerate(names):
            self.tableWidget.setRowCount(self.tableWidget.rowCount() + 1)
            for j, elem in enumerate(row):
                if j == 0:
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(elem)))
        self.db.close()

    def save_table(self):
        if ' ' in self.lineEdit.text():
            self.lineEdit.setText('⠀'.join(i for i in self.lineEdit.text().split()))
        sqlite_connection = sqlite3.connect(f'db/{form.comboBox_3.currentText()}.db')
        try:
            extra = sqlite_connection.cursor().execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
            extra = [i[0] for i in extra]
            if self.lineEdit.text() in extra:
                raise IndexError
            sqlite_create_table_query = f"""CREATE TABLE {self.lineEdit.text()} (
        id INT PRIMARY KEY,
        ФИО INT REFERENCES pupils (id),"""
            for i in self.header[1:-1]:
                sqlite_create_table_query += f"\n\t{i} INT,"
            sqlite_create_table_query += f'\n\t{self.header[-1]} INT'
            sqlite_create_table_query += '\n);'
            cursor = sqlite_connection.cursor()
            cursor.execute(sqlite_create_table_query)
            sqlite_connection.commit()
            for i in range(self.tableWidget.rowCount()):
                sqlite_connection.commit()
                data = [i + 1]
                for j in range(len(self.header)):
                    if j == 0:
                        data.append(i + 1)
                    else:
                        data.append(int(self.tableWidget.item(i, j).text()))
                stroka = "INSERT INTO "
                stroka += self.lineEdit.text() + " VALUES(?, "
                stroka += ', '.join('?' for _ in range(len(self.header)))
                stroka += ')'
                cursor.execute(stroka, data)
                sqlite_connection.commit()
        except sqlite3.OperationalError:
            self.progress.setText('Ошибка: ошибка в заполнении таблицы')
            self.progress.setStyleSheet('color: red')
            self.progress.adjustSize()
            sqlite_connection.cursor().execute(f'DROP TABLE IF EXISTS {self.lineEdit.text()}')
            return
        except ValueError:
            self.progress.setText('Ошибка: ошибка в заполнении таблицы')
            self.progress.setStyleSheet('color: red')
            self.progress.adjustSize()
            sqlite_connection.cursor().execute(f'DROP TABLE IF EXISTS {self.lineEdit.text()}')
            return
        except IndexError:
            self.progress.setText('Ошибка: такая таблица уже существует')
            self.progress.setStyleSheet('color: red')
            self.progress.adjustSize()
            sqlite_connection.cursor().execute(f'DROP TABLE IF EXISTS {self.lineEdit.text()}')
            return

        x = cursor.execute(f"PRAGMA table_info(journal)").fetchall()
        x = [i[1] for i in x]
        if self.lineEdit.text() not in x:
            cursor.execute(f"ALTER TABLE journal ADD COLUMN {self.lineEdit.text()} INT NOT NULL DEFAULT (-1);")
            sqlite_connection.commit()
            cursor = sqlite_connection.cursor()
            for i in range(self.tableWidget.rowCount()):
                name = self.lineEdit.text()
                n = int(self.tableWidget.item(i, 1).text())
                cursor.execute(f"""UPDATE journal
                                    SET {name} = ?
                                    WHERE id = ?""", (n, i + 1))
                sqlite_connection.commit()
            sqlite_connection.close()
            self.progress.setText("Таблица успешно сохранена")
            self.progress.setStyleSheet('color: green')
            self.progress.adjustSize()
            form.comboBox_2.addItem(self.lineEdit.text())
            form.main_table()


def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    form = MainForm()
    form.show()
    sys.excepthook = except_hook
    sys.exit(app.exec())