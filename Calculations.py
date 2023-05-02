from math import pow, sqrt
from PyQt5 import QtWidgets, QtCore
from time import sleep
from CreateReportExcel import DataStructure


class Calc(QtCore.QThread):
    Rating_max = 4
    sleep_time = 0.01
    reportProgress = QtCore.pyqtSignal(int)
    reportValues = QtCore.pyqtSignal(list, int)
    reportEnd = QtCore.pyqtSignal(bool)
    reportPrivateRating = QtCore.pyqtSignal(float, int)
    reportFinalGrade = QtCore.pyqtSignal(float)
    reportIncompRatio = QtCore.pyqtSignal(float)
    reportFinalData = QtCore.pyqtSignal(DataStructure)

    posbl_state_ests = [
        [[4, 3, 2, 1, 0, 2],
         [4, 3, 2, 3]],

        [[4, 0, 2],
         [4, 2, 0, 2],
         [4, 3, 2, 3],
         [4, 2, 3],
         [4, 0, 2],
         [4, 2, 3],
         [4, 2, 3],
         [4, 2, 0, 2]],

        [[4, 1, 2.5],
         [4, 1, 2.5],
         [4, 1, 2.5]],

        [[4, 3, 1, 2.5],
         [4, 0, 2],
         [4, 0, 2],
         [4, 3, 2, 1, 0, 2]],

        [[4, 0, 2],
         [4, 3, 2, 1, 2.5],
         [4, 2, 3],
         [4, 2, 0, 2],
         [4, 3, 2, 1, 0, 2],
         [4, 1, 2.5]],

        [[4, 3, 2, 1, 0, 2],
         [4, 3, 2, 1, 0, 2],
         [4, 3, 2, 1, 0, 2],
         [4, 2, 3]]
    ]

    def __init__(self, states, rangs, states_text):
        super().__init__()
        self.states = states
        self.rangs = rangs                          # Ранг
        self.states_text = states_text

    def run(self) -> None:
        self.states_sys = self.Determ_grades()      # Оценка в баллах для установленного состояния признака
        self.fi_values = self.WeighingFunc_fi()     # Значение «взвешивающей функции»
        self.p_rating = self.PropertyBlock()        # Частная опастность по функциональному блоку
        self.f_grade = self.RatingD()               # Финальная оценка опастности вещества
        self.incomp_ratio = self.IncompRatio()      # Коэфициент достоверности результатов

        structure = DataStructure(self.states_sys,
                                  self.states_text,
                                  self.rangs,
                                  self.fi_values,
                                  self.p_rating,
                                  self.f_grade,
                                  self.incomp_ratio,
                                  "", "")
        self.reportFinalData.emit(structure)
        self.reportEnd.emit(True)

    # Определение по зараннее подготовленой таблице оценки состояния признаков
    def Determ_grades(self):
        states_sys = []
        for i in range(len(self.posbl_state_ests)):
            states_block = []
            for j in range(len(self.states[i])):
                index = self.states[i][j]
                value = self.posbl_state_ests[i][j][index]
                states_block.append(value)
                self.reportProgress.emit(1)
                sleep(self.sleep_time)
            states_sys.append(states_block)
        self.reportValues.emit(states_sys, 3)
        return states_sys

    # Определение значений взвешифающей функции для каждого ранга каждого признака
    def WeighingFunc_fi(self):
        fi_mass = []
        for i in range(len(self.rangs)):
            fi = []
            for j in range(len(self.rangs[i])):
                rang = self.rangs[i][j]
                if rang == 1:
                    fi.append(2)
                else:
                    fi.append(rang / (pow(2, rang - 1)))
                self.reportProgress.emit(1)
                sleep(self.sleep_time)
            fi_mass.append(fi)

        self.reportValues.emit(fi_mass, 2)
        return fi_mass

    # Вычисление колличественной оценки опастности по каждому блоку признаков
    def PropertyBlock(self):
        private_rating = []
        for i in range(len(self.states_sys)):
            sum_R_fi = 0
            sum_Rmax_fi = 0
            for j in range(len(self.states_sys[i])):
                sum_R_fi += (self.states_sys[i][j] * self.fi_values[i][j])
                sum_Rmax_fi += (self.Rating_max * self.fi_values[i][j])
                self.reportProgress.emit(1)
                sleep(self.sleep_time)
            D = sum_R_fi / sum_Rmax_fi
            private_rating.append(D)
            self.reportPrivateRating.emit(D, i+1)
        return private_rating

    # Вычисление финальной колличественной оценки опастности наноматериала
    def RatingD(self):
        final_grade = 0
        for Dk in self.p_rating:
            final_grade += pow(Dk, 2)
            self.reportProgress.emit(1)
            sleep(self.sleep_time)
        final_grade = sqrt(final_grade)
        self.reportFinalGrade.emit(final_grade)
        return final_grade

    # Вычисление коэфициента неполноты анализа
    def IncompRatio(self):
        sum_fi = 0
        for fi_block in self.fi_values:
            for fi_elem in fi_block:
                sum_fi += fi_elem
                self.reportProgress.emit(1)
                sleep(self.sleep_time)

        sum_u_fi = 0
        for i in range(len(self.states_text)):
            for j in range(len(self.states_text[i])):
                if self.states_text[i][j] == 'Не определённый':
                    sum_u_fi += self.fi_values[i][j]
                self.reportProgress.emit(1)
                sleep(self.sleep_time)

        incomp_ratio = sum_u_fi / sum_fi
        self.reportIncompRatio.emit(incomp_ratio)
        return incomp_ratio