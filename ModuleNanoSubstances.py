import sys
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QMessageBox, QErrorMessage
from math import ceil
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSize
from PyQt5.QtWidgets import qApp, QInputDialog, QFileDialog

import Calculations
import Interface_nanosubstances
import Instructions
import CreateReportExcel

about_info = "Правообладатель:\n"\
             "Федеральное государственное бюджетное образовательное учреждение"\
             "Высшего образования «Донской государственный технический университет»\n\n"\
             "Программа предназначена для автоматизации и упрощения процесса предварительной "\
             "оценки безопасности продуктов, содержащих нановещества, с целью определения "\
             "необходимого объёма и сложности дальнейших экспериментальных исследований.\n\n"\
             "Использование программы в исследовательских целях даёт возможность пользователю\n"\
             "   1) рассчитывать степень потенциальной опасности исследуемого нановещества на основе "\
             "составления генеральной определительной таблицы, оценивающей основные физико-химические "\
             "свойства данного вещества и признаки, отражающие его взаимодействие с биологическими "\
             "объектами с точки зрения проявления им токсичных свойств;\n"\
             "   2) выполнять анализ безопасности применения продуктов нанотехнологий, на основе "\
             "расчета степени потенциальной опасности входящих в них наноматериалов;\n"\
             "   3) выполнять анализ целесообразности применения стекол с селективным магнетронным "\
             "покрытием, основываясь на сочетании оценки безопасности, входящих в покрытие нанокомпонентов, "\
             "и функционального эффекта от нанесения таких покрытий."

root_app: bool = False

class InstructionsClass(QtWidgets.QWidget, Instructions.Ui_Form_instructions):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

class ClassNanoSubstances(QtWidgets.QMainWindow, Interface_nanosubstances.Ui_Window_nanosubstances):
    signal_close = QtCore.pyqtSignal(list)
    name_material_chek = False
    name_material = "Название не задано"

    def __init__(self, parent=None):
        super().__init__(parent, QtCore.Qt.Window)
        self.setupUi(self)

        self.Icon = QIcon("logo\\icon.ico")
        self.setWindowIcon(self.Icon)

        self.f_grade = 0.0
        self.incomp_ratio = 1.0
        self.f_grade_text = "Не удалось произвести анализ"
        self.incomp_ratio_text = ["-------", "-------"]

        self.createComboInTable_block1()
        self.createComboInTable_block2()
        self.createComboInTable_block3()
        self.createComboInTable_block4()
        self.createComboInTable_block5()
        self.createComboInTable_block6()
        self.kolv_cikl = 0
        self.kontrol_create_data = False

        self.defoult_rangs = [
            [1, 1],
            [1, 2, 1, 3, 3, 4, 5, 2],
            [1, 3, 2],
            [2, 1, 3, 1],
            [4, 2, 3, 1, 1, 1],
            [1, 1, 2, 3]
        ]

        self.Calc = QtCore.QThread()
        self.report = QtCore.QThread()
        self.but_reset_def_rangs.clicked.connect(self.ResetRengs)
        self.but_start.clicked.connect(self.analiz_start)
        self.action_about.triggered.connect(self.AboutClicked)
        self.action_help.triggered.connect(self.HelpClicked)
        if root_app:
            self.action_quit.triggered.connect(qApp.quit)
        else:
            self.action_quit.triggered.connect(self.close)
        self.action_Excel.triggered.connect(self.createExcelReport)

    def analiz_start(self):
        self.kolv_cikl = 0
        self.progressBar.setValue(self.kolv_cikl)

        if self.Calc.isRunning():
            print("Поток с вычислениями уже запущен!!!")
            return

        self.kontrol_create_data = False
        states_block1 = [self.block1_box1.currentIndex(),
                         self.block1_box2.currentIndex()]
        states_block2 = [self.block2_box1.currentIndex(),
                         self.block2_box2.currentIndex(),
                         self.block2_box3.currentIndex(),
                         self.block2_box4.currentIndex(),
                         self.block2_box5.currentIndex(),
                         self.block2_box6.currentIndex(),
                         self.block2_box7.currentIndex(),
                         self.block2_box8.currentIndex()]
        states_block3 = [self.block3_box1.currentIndex(),
                         self.block3_box2.currentIndex(),
                         self.block3_box3.currentIndex()]
        states_block4 = [self.block4_box1.currentIndex(),
                         self.block4_box2.currentIndex(),
                         self.block4_box3.currentIndex(),
                         self.block4_box4.currentIndex()]
        states_block5 = [self.block5_box1.currentIndex(),
                         self.block5_box2.currentIndex(),
                         self.block5_box3.currentIndex(),
                         self.block5_box4.currentIndex(),
                         self.block5_box5.currentIndex(),
                         self.block5_box6.currentIndex()]
        states_block6 = [self.block6_box1.currentIndex(),
                         self.block6_box2.currentIndex(),
                         self.block6_box3.currentIndex(),
                         self.block6_box4.currentIndex()]

        self.kontrol_rangs = True
        rangs_block1 = []
        for i in range(2):
            rang = int(self.table_block_1.item(i, 1).text())
            if rang < 1 or rang > 5:
                self.MsgErrRang("Физические характеристики")
                self.kontrol_rangs = False
                break
            else:
                rangs_block1.append(rang)

        rangs_block2 = []
        for i in range(8):
            rang = int(self.table_block_2.item(i, 1).text())
            if rang < 1 or rang > 5:
                self.MsgErrRang("Физико-химические свойства")
                self.kontrol_rangs = False
                break
            else:
                rangs_block2.append(rang)

        rangs_block3 = []
        for i in range(3):
            rang = int(self.table_block_3.item(i, 1).text())
            if rang < 1 or rang > 5:
                self.MsgErrRang("Молекулярно-биологические свойства")
                self.kontrol_rangs = False
                break
            else:
                rangs_block3.append(rang)

        rangs_block4 = []
        for i in range(4):
            rang = int(self.table_block_4.item(i, 1).text())
            if rang < 1 or rang > 5:
                self.MsgErrRang("Цитологические свойства")
                self.kontrol_rangs = False
                break
            else:
                rangs_block4.append(rang)

        rangs_block5 = []
        for i in range(6):
            rang = int(self.table_block_5.item(i, 1).text())
            if rang < 1 or rang > 5:
                self.MsgErrRang("Физиологические свойства")
                self.kontrol_rangs = False
                break
            else:
                rangs_block5.append(rang)

        rangs_block6 = []
        for i in range(4):
            rang = int(self.table_block_6.item(i, 1).text())
            if rang < 1 or rang > 5:
                self.MsgErrRang("Физиологические свойства")
                self.kontrol_rangs = False
                break
            else:
                rangs_block6.append(rang)

        if self.kontrol_rangs == False:
            return

        states = [states_block1,
                  states_block2,
                  states_block3,
                  states_block4,
                  states_block5,
                  states_block6]

        states_text = [[self.block1_box1.currentText(),
                        self.block1_box2.currentText()],
                       [self.block2_box1.currentText(),
                        self.block2_box2.currentText(),
                        self.block2_box3.currentText(),
                        self.block2_box4.currentText(),
                        self.block2_box5.currentText(),
                        self.block2_box6.currentText(),
                        self.block2_box7.currentText(),
                        self.block2_box8.currentText()],
                       [self.block3_box1.currentText(),
                        self.block3_box2.currentText(),
                        self.block3_box3.currentText()],
                       [self.block4_box1.currentText(),
                        self.block4_box2.currentText(),
                        self.block4_box3.currentText(),
                        self.block4_box4.currentText()],
                       [self.block5_box1.currentText(),
                        self.block5_box2.currentText(),
                        self.block5_box3.currentText(),
                        self.block5_box4.currentText(),
                        self.block5_box5.currentText(),
                        self.block5_box6.currentText()],
                       [self.block6_box1.currentText(),
                        self.block6_box2.currentText(),
                        self.block6_box3.currentText(),
                        self.block6_box4.currentText()]
                       ]

        rangs = [rangs_block1,
                 rangs_block2,
                 rangs_block3,
                 rangs_block4,
                 rangs_block5,
                 rangs_block6]

        self.Calc = Calculations.Calc(states, rangs, states_text)
        self.Calc.reportProgress.connect(self.progress)
        self.Calc.reportValues.connect(self.OutputValues)
        self.Calc.reportEnd.connect(self.EndSys)
        self.Calc.reportPrivateRating.connect(self.OutputPrivateRating)
        self.Calc.reportFinalGrade.connect(self.OutputFinalGrade)
        self.Calc.reportIncompRatio.connect(self.OutputIncompRatio)
        self.Calc.reportFinalData.connect(self.createDataExcel)
        self.Calc.start()

    def createComboInTable_block1(self):
        combo_atribut1 = ["Преобладают частицы\nменее 5 нм",
                          "Преобладают частицы 5—50 нм",
                          "Преобладают частицы\n50—100 нм",
                          "Преобладают частицы > 100 нм, но есть\nсущественная фракция < 100 нм",
                          "Преобладают частицы > 100 нм, содержание\nменьших частиц несущественно",
                          "Не определённый"]
        combo_atribut2 = ["Частицы крайне несферичны (формфактор > 100)",
                          "Частицы высоко несферичны (10—100)",
                          "Форма частиц близка к сферической (1—10)",
                          "Не определённый"]

        self.block1_box1 = QtWidgets.QComboBox()
        self.block1_box2 = QtWidgets.QComboBox()

        for j in combo_atribut1:
            self.block1_box1.addItem(j)
        for j in combo_atribut2:
            self.block1_box2.addItem(j)

        self.block1_box1.setCurrentIndex(len(combo_atribut1) - 1)
        self.block1_box2.setCurrentIndex(len(combo_atribut2) - 1)
        self.table_block_1.setCellWidget(0, 0, self.block1_box1)
        self.table_block_1.setCellWidget(1, 0, self.block1_box2)

        self.table_block_1.setColumnWidth(0, 395)
        self.table_block_1.setColumnWidth(1, 70)
        self.table_block_1.setColumnWidth(2, 130)
        self.table_block_1.setColumnWidth(3, 140)

    def createComboInTable_block2(self):
        combo_atribut1 = ["Нерастворимы",
                          "Растворимы",
                          "Не определённый"]
        combo_atribut2 = ["Нерастворимы",
                          "Малорастворимы",
                          "Растворимы",
                          "Не определённый"]
        combo_atribut3 = ["Положительный",
                          "Отрицательный",
                          "Не заряжены",
                          "Не определённый"]
        combo_atribut4 = ["Высокая",
                          "Низкая",
                          "Не определённый"]
        combo_atribut5 = ["Гидрофобны",
                          "Гидрофильны",
                          "Не определённый"]
        combo_atribut6 = ["Выявлена",
                          "Выявлена в условиях освещения",
                          "Не выявлена",
                          "Не определённый"]

        self.block2_box1 = QtWidgets.QComboBox()
        self.block2_box2 = QtWidgets.QComboBox()
        self.block2_box3 = QtWidgets.QComboBox()
        self.block2_box4 = QtWidgets.QComboBox()
        self.block2_box5 = QtWidgets.QComboBox()
        self.block2_box6 = QtWidgets.QComboBox()
        self.block2_box7 = QtWidgets.QComboBox()
        self.block2_box8 = QtWidgets.QComboBox()

        for j in combo_atribut1:
            self.block2_box1.addItem(j)
        for j in combo_atribut2:
            self.block2_box2.addItem(j)
        for j in combo_atribut3:
            self.block2_box3.addItem(j)
        for j in combo_atribut4:
            self.block2_box4.addItem(j)
            self.block2_box5.addItem(j)
            self.block2_box7.addItem(j)
        for j in combo_atribut5:
            self.block2_box6.addItem(j)
        for j in combo_atribut6:
            self.block2_box8.addItem(j)

        self.block2_box1.setCurrentIndex(len(combo_atribut1) - 1)
        self.block2_box2.setCurrentIndex(len(combo_atribut2) - 1)
        self.block2_box3.setCurrentIndex(len(combo_atribut3) - 1)
        self.block2_box4.setCurrentIndex(len(combo_atribut4) - 1)
        self.block2_box5.setCurrentIndex(len(combo_atribut4) - 1)
        self.block2_box6.setCurrentIndex(len(combo_atribut5) - 1)
        self.block2_box7.setCurrentIndex(len(combo_atribut4) - 1)
        self.block2_box8.setCurrentIndex(len(combo_atribut6) - 1)

        self.table_block_2.setCellWidget(0, 0, self.block2_box1)
        self.table_block_2.setCellWidget(1, 0, self.block2_box2)
        self.table_block_2.setCellWidget(2, 0, self.block2_box3)
        self.table_block_2.setCellWidget(3, 0, self.block2_box4)
        self.table_block_2.setCellWidget(4, 0, self.block2_box5)
        self.table_block_2.setCellWidget(5, 0, self.block2_box6)
        self.table_block_2.setCellWidget(6, 0, self.block2_box7)
        self.table_block_2.setCellWidget(7, 0, self.block2_box8)

        self.table_block_2.setColumnWidth(0, 290)
        self.table_block_2.setColumnWidth(1, 104)
        self.table_block_2.setColumnWidth(2, 140)
        self.table_block_2.setColumnWidth(3, 140)

    def createComboInTable_block3(self):
        combo_atribut1 = ["Выявлено",
                          "Не выявлено",
                          "Не определённый"]

        self.block3_box1 = QtWidgets.QComboBox()
        self.block3_box2 = QtWidgets.QComboBox()
        self.block3_box3 = QtWidgets.QComboBox()

        for j in combo_atribut1:
            self.block3_box1.addItem(j)
            self.block3_box2.addItem(j)
            self.block3_box3.addItem(j)

        self.block3_box1.setCurrentIndex(len(combo_atribut1) - 1)
        self.block3_box2.setCurrentIndex(len(combo_atribut1) - 1)
        self.block3_box3.setCurrentIndex(len(combo_atribut1) - 1)
        self.table_block_3.setCellWidget(0, 0, self.block3_box1)
        self.table_block_3.setCellWidget(1, 0, self.block3_box2)
        self.table_block_3.setCellWidget(2, 0, self.block3_box3)

        self.table_block_3.setColumnWidth(0, 350)
        self.table_block_3.setColumnWidth(1, 126)
        self.table_block_3.setColumnWidth(2, 200)
        self.table_block_3.setColumnWidth(3, 140)

    def createComboInTable_block4(self):
        combo_atribut1 = ["Накапливается в органеллах и цитозоле",
                          "Накапливается только в органеллах",
                          "Накапливается только в цитозоле",
                          "Не определённый"]
        combo_atribut2 = ["Выявлена",
                          "Не выявлена",
                          "Не определённый"]
        combo_atribut3 = ["Вызывает летальные изменения в\nнормальных клетках",
                          "Вызывает стойкие нелетальные морфологические\nизменения в нормальных клетках",
                          "Вызывает летальные изменения в\nтрансформированных клетках",
                          "Вызывает обратимые морфологические\nизменения в клетках различных видов",
                          "Отсутствие влияния",
                          "Не определённый"]

        self.block4_box1 = QtWidgets.QComboBox()
        self.block4_box2 = QtWidgets.QComboBox()
        self.block4_box3 = QtWidgets.QComboBox()
        self.block4_box4 = QtWidgets.QComboBox()

        for j in combo_atribut1:
            self.block4_box1.addItem(j)
        for j in combo_atribut2:
            self.block4_box2.addItem(j)
            self.block4_box3.addItem(j)
        for j in combo_atribut3:
            self.block4_box4.addItem(j)

        self.block4_box1.setCurrentIndex(len(combo_atribut1) - 1)
        self.block4_box2.setCurrentIndex(len(combo_atribut2) - 1)
        self.block4_box3.setCurrentIndex(len(combo_atribut2) - 1)
        self.block4_box4.setCurrentIndex(len(combo_atribut3) - 1)
        self.table_block_4.setCellWidget(0, 0, self.block4_box1)
        self.table_block_4.setCellWidget(1, 0, self.block4_box2)
        self.table_block_4.setCellWidget(2, 0, self.block4_box3)
        self.table_block_4.setCellWidget(3, 0, self.block4_box4)

        self.table_block_4.setColumnWidth(0, 417)
        self.table_block_4.setColumnWidth(1, 70)
        self.table_block_4.setColumnWidth(2, 140)
        self.table_block_4.setColumnWidth(3, 140)

    def createComboInTable_block5(self):
        combo_atribut1 = ["Выявлено",
                          "Не выявлено",
                          "Не определённый"]
        combo_atribut2 = ["Накапливается во многих\nорганах и тканях",
                          "Накапливается в отдельных\nорганах и тканях",
                          "Накапливается в одном органе",
                          "Накопление не выявлено",
                          "Не определённый"]
        combo_atribut3 = ["Доказано",
                          "Не доказано",
                          "Не определённый"]
        combo_atribut4 = ["1—2 класс опасности",
                          "3 класс опасности",
                          "4 класс опасности",
                          "Не определённый"]
        combo_atribut5 = ["Токсично для человека и\nтеплокровных животных",
                          "Токсично для холоднокровных\nпозвоночных",
                          "Токсично для беспозвоночных",
                          "Токсично для растений\nи (или) прокариот",
                          "Токсичность не выявлена",
                          "Не определённый"]

        self.block5_box1 = QtWidgets.QComboBox()
        self.block5_box2 = QtWidgets.QComboBox()
        self.block5_box3 = QtWidgets.QComboBox()
        self.block5_box4 = QtWidgets.QComboBox()
        self.block5_box5 = QtWidgets.QComboBox()
        self.block5_box6 = QtWidgets.QComboBox()

        for j in combo_atribut1:
            self.block5_box1.addItem(j)
            self.block5_box6.addItem(j)
        for j in combo_atribut2:
            self.block5_box2.addItem(j)
        for j in combo_atribut3:
            self.block5_box3.addItem(j)
        for j in combo_atribut4:
            self.block5_box4.addItem(j)
        for j in combo_atribut5:
            self.block5_box5.addItem(j)

        self.block5_box1.setCurrentIndex(len(combo_atribut1) - 1)
        self.block5_box2.setCurrentIndex(len(combo_atribut2) - 1)
        self.block5_box3.setCurrentIndex(len(combo_atribut3) - 1)
        self.block5_box4.setCurrentIndex(len(combo_atribut4) - 1)
        self.block5_box5.setCurrentIndex(len(combo_atribut5) - 1)
        self.block5_box6.setCurrentIndex(len(combo_atribut1) - 1)

        self.table_block_5.setCellWidget(0, 0, self.block5_box1)
        self.table_block_5.setCellWidget(1, 0, self.block5_box2)
        self.table_block_5.setCellWidget(2, 0, self.block5_box3)
        self.table_block_5.setCellWidget(3, 0, self.block5_box4)
        self.table_block_5.setCellWidget(4, 0, self.block5_box5)
        self.table_block_5.setCellWidget(5, 0, self.block5_box6)

        self.table_block_5.setColumnWidth(0, 295)
        self.table_block_5.setColumnWidth(1, 70)
        self.table_block_5.setColumnWidth(2, 140)
        self.table_block_5.setColumnWidth(3, 140)

        self.table_block_5.setRowHeight(0, 35)
        self.table_block_5.setRowHeight(1, 50)
        self.table_block_5.setRowHeight(2, 50)
        self.table_block_5.setRowHeight(3, 35)
        self.table_block_5.setRowHeight(4, 50)
        self.table_block_5.setRowHeight(5, 100)

    def createComboInTable_block6(self):
        combo_atribut1 = ["Крупнотоннажный промышленный\nпродукт (> 1000 т)",
                          "Массово выпускаемый\nпродукт (1 — 1000 т)",
                          "Продукт, выпускаемый в ограниченных\nколичествах (1 кг — 1 т)",
                          "Продукт, выпускаемый в малых\nколичествах (менее 1 кг)",
                          "Продукт в настоящее\nвремя не производится",
                          "Не определённый"]
        combo_atribut2 = ["Население в масштабе страны",
                          "Потребители продукции",
                          "Персонал массового производства",
                          "Персонал в лабораторных масштабах",
                          "Персонал в лабораторных масштабах",
                          "Не определённый"]
        combo_atribut3 = ["Сельскохозяйственные животные\nи культурные растения",
                          "Массовые виды диких животных, дикорастущих\nрастений и свободноживущих микроорганизмов",
                          "Малочисленные компоненты биоценоза,\nбезразличные для хозяйственной деятельности",
                          "Растения и животные – вредители\nсельскохозяйственных культур",
                          "Накопление не выявлено",
                          "Не определённый"]
        combo_atribut4 = ["Данные имеются",
                          "Данные не имеются",
                          "Не определённый"]

        self.block6_box1 = QtWidgets.QComboBox()
        self.block6_box2 = QtWidgets.QComboBox()
        self.block6_box3 = QtWidgets.QComboBox()
        self.block6_box4 = QtWidgets.QComboBox()

        for j in combo_atribut1:
            self.block6_box1.addItem(j)
        for j in combo_atribut2:
            self.block6_box2.addItem(j)
        for j in combo_atribut3:
            self.block6_box3.addItem(j)
        for j in combo_atribut4:
            self.block6_box4.addItem(j)

        self.block6_box1.setCurrentIndex(len(combo_atribut1) - 1)
        self.block6_box2.setCurrentIndex(len(combo_atribut2) - 1)
        self.block6_box3.setCurrentIndex(len(combo_atribut3) - 1)
        self.block6_box4.setCurrentIndex(len(combo_atribut4) - 1)

        self.table_block_6.setCellWidget(0, 0, self.block6_box1)
        self.table_block_6.setCellWidget(1, 0, self.block6_box2)
        self.table_block_6.setCellWidget(2, 0, self.block6_box3)
        self.table_block_6.setCellWidget(3, 0, self.block6_box4)

        self.table_block_6.setColumnWidth(0, 427)
        self.table_block_6.setColumnWidth(1, 70)
        self.table_block_6.setColumnWidth(2, 140)
        self.table_block_6.setColumnWidth(3, 140)

    def progress(self, value):
        self.kolv_cikl += value
        self.progressBar.setValue(self.kolv_cikl)

    def MsgErrRang(self, block):
        QMessageBox.question(self, 'Ошибка ввода данных', "Вы ввели неверные значения для ранга\n"
                                                          "в блоке: " + block + "!\n"
                                                                                "Значение ранга не должно быть меньше 1 и не больше 5.\n"
                                                                                "Повторите ввод рейтинга!!!",
                             QMessageBox.Ok)

    def OutputValues(self, values, col):
        for i in range(len(values[0])):
            item = self.table_block_1.item(i, col)
            item.setText(str(values[0][i]))

        for i in range(len(values[1])):
            item = self.table_block_2.item(i, col)
            item.setText(str(values[1][i]))

        for i in range(len(values[2])):
            item = self.table_block_3.item(i, col)
            item.setText(str(values[2][i]))

        for i in range(len(values[3])):
            item = self.table_block_4.item(i, col)
            item.setText(str(values[3][i]))

        for i in range(len(values[4])):
            item = self.table_block_5.item(i, col)
            item.setText(str(values[4][i]))

        for i in range(len(values[5])):
            item = self.table_block_6.item(i, col)
            item.setText(str(values[5][i]))

    def EndSys(self, var):
        print(self.kolv_cikl)

    def ResetRengs(self):
        self.OutputValues(self.defoult_rangs, 1)

    def OutputPrivateRating(self, PR_mass, tab_indx):
        if tab_indx == 1:
            self.out_block1.setText(str(ceil(PR_mass * 1000) / 1000))
        elif tab_indx == 2:
            self.out_block2.setText(str(ceil(PR_mass * 1000) / 1000))
        elif tab_indx == 3:
            self.out_block3.setText(str(ceil(PR_mass * 1000) / 1000))
        elif tab_indx == 4:
            self.out_block4.setText(str(ceil(PR_mass * 1000) / 1000))
        elif tab_indx == 5:
            self.out_block5.setText(str(ceil(PR_mass * 1000) / 1000))
        elif tab_indx == 6:
            self.out_block6.setText(str(ceil(PR_mass * 1000) / 1000))

    def OutputFinalGrade(self, FGrade):
        self.f_grade = ceil(FGrade * 1000) / 1000
        self.out_kolv.setText(str(self.f_grade))
        if 0.441 <= FGrade <= 1.11:
            self.f_grade_text = "Низкая степень потенциальной опасности"
        elif 1.111 <= FGrade <= 1.779:
            self.f_grade_text = "Средняя степень потенциальной опасности"
        elif 1.78 <= FGrade <= 2.449:
            self.f_grade_text = "Высокая степень потенциальной опасности"
        else:
            self.f_grade_text = "Ошибка в процессе анализа"
        self.out_kachestvo.setText("<center>" + self.f_grade_text + "</center>")

    def OutputIncompRatio(self, IRatio):
        self.incomp_ratio = ceil(IRatio * 1000) / 1000
        self.out_koef_nepoln.setText(str(self.incomp_ratio))
        self.incomp_ratio_text = []
        if 0 <= IRatio <= 0.25:
            self.incomp_ratio_text.append("Оценка является достоверной")
            self.incomp_ratio_text.append("Данных достаточно для определения\n"
                                             "степени опасности наноматериала")
        elif 0.251 <= IRatio <= 0.75:
            self.incomp_ratio_text.append("Оценка является сомнительной")
            self.incomp_ratio_text.append("Часть важных параметров, которые\n"
                                             "характеризуют степень опасности наноматериала\n"
                                             "в изученных источниках не исследованы")
        elif 0.751 <= IRatio <= 1.0:
            self.incomp_ratio_text.append("Оценка является недостоверной")
            self.incomp_ratio_text.append("Имеющиеся в наличии данные крайне\n"
                                             "недостаточны для выводов о степени опастности\n"
                                             "изучаемого наноматериала, необходим более\n"
                                             "расширенный поиск источников или проведение\n"
                                             "эксперементальных исследований искомых параметров")
        else:
            self.incomp_ratio_text.append("Ошибка в определении достоверности")
            self.incomp_ratio_text.append("Пожалуста введите параметры снова и\n"
                                             "повторите анализ данных\n")

        self.out_dostovernost.setText("<center>" + self.incomp_ratio_text[0] + "</center>")
        self.out_dostovernost.setToolTip(self.incomp_ratio_text[1])

    def AboutClicked(self):
        self.Icon.actualSize(QSize(150, 150), QIcon.Normal, QIcon.On)
        QMessageBox.about(self, "Программа ECO-NANO-TEST", about_info)

    def createDataExcel(self, report_struc):
        self.DataStrucReport = CreateReportExcel.DataStructure(report_struc.states,
                                                               report_struc.states_text,
                                                               report_struc.rangs,
                                                               report_struc.fi_values,
                                                               report_struc.p_rating,
                                                               report_struc.f_grade,
                                                               report_struc.incomp_ratio,
                                                               self.f_grade_text,
                                                               self.incomp_ratio_text)
        self.kontrol_create_data = True

    def createExcelReport(self):
        if not self.kontrol_create_data:
            QMessageBox.question(self, 'Ошибка сохранения',
                                 "Сохранение не удалось, так\n"
                                 "как не был проведён анализ данных!",
                                 QMessageBox.Ok)
            return

        text = None
        ok = False
        if not self.name_material_chek:
            text, ok = QInputDialog.getText(self, 'Ввод названия',
                                            'Введите название исследуемого наноматериала:')

        if ok and len(str(text)) > 0:
            self.name_material = str(text)
            self.name_material_chek = True
        else:
            error_message = QtWidgets.QErrorMessage(self)
            error_message.setWindowTitle("Ошибка сохранения")
            error_message.showMessage("Сохранение не удалось, так как не было\n"
                                      "введено название исследуемного материала!")
            return

        file_path = QFileDialog.getSaveFileName(self, "Сохранить файл", "", "Электронная таблица (*.xlsx)")[0]

        self.report = CreateReportExcel.ExcelReport(self.DataStrucReport, self.name_material, file_path)
        self.report.reportErrSave.connect(self.ErrSave)
        self.report.start()

    def ErrSave(self):
        QMessageBox.question(self, 'Ошибка сохранения',
                             "Сохранение не удалось, проверьте не\n"
                             "открыт ли файл в другой программе",
                             QMessageBox.Ok)

    def HelpClicked(self):
        self.window_help = InstructionsClass()
        self.window_help.show()

    def closeEvent(self, event):
        if root_app:
            return

        text = None
        ok = False
        if not self.name_material_chek:
            text, ok = QInputDialog.getText(self, 'Ввод названия',
                                            'Введите название исследуемого наноматериала:')

        if ok and len(str(text)) > 0:
            self.name_material = str(text)

        struktura_data = [
            self.name_material,
            self.f_grade,
            self.f_grade_text,
            self.incomp_ratio,
            self.incomp_ratio_text
        ]
        self.signal_close.emit(struktura_data)


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = ClassNanoSubstances()
    window.show()
    app.exec_()


if __name__ == "__main__":
    root_app = True
    main()
