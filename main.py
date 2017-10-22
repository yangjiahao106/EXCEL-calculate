# python3: calculate self.average
import openpyxl
import re


class Calculator(object):
    """:type file_name : a string about the path and name you want open
        :type name_scope : a tuple content start cell of the name column and end cell of the name column
                          e.g. ('E6', 'E35')
        :type credit_scope : a tuple content start cell of the credit rand and end cell of the credit rand
                             e.g. ('F4','BJ4')
        :type score_scope : a tuple content the first score of the first student and the last score of the
                            last student e.g. ('F6','BJ36')
    """
    def __init__(self, file_name, name_scope, credit_scope, score_scope):
        self.__file_name = file_name
        self.__name_scope = name_scope
        self.__credit_scope = credit_scope
        self.__score_scope = score_scope
        self.__all_credit = []
        self.__all_stu_names = []
        self.__all_stu_credit = []
        self.__all_stu_scores = []
        self.__average = []

    def start(self):
        """start to calculate and print the average"""
        sheet = self.__get_sheet()
        self.__get_all_stu_names(sheet)
        self.__get_all_credit(sheet)
        self.__get_student_score(sheet)
        self.__get_average()
        self.out_put()

    def __get_sheet(self):
        try:
            wb = openpyxl.load_workbook(self.__file_name)
            sheet = wb.get_active_sheet()
            return sheet
        except FileNotFoundError:
            print('Open file error.')

    def __get_all_stu_names(self, sheet):
        cols = sheet[self.__name_scope[0]:self.__name_scope[1]]
        for col in cols:
            self.__all_stu_names.append(col[0].value.strip())

    def __get_all_credit(self, sheet):
        line1 = sheet[self.__credit_scope[0]:self.__credit_scope[1]]  # return a tuple with one element
        for i in line1[0]:
            self.__all_credit.append(eval(i.value))

    @staticmethod
    def get_evaluate(score):
        """change grade to score"""
        if re.search(r"\d+", score):
            a = re.search(r"\d+", score).group()
            return eval(a)
        elif score.strip() == '优秀':
            return 95
        elif score.strip() == '良好':
            return 85
        elif score.strip() == '中等':
            return 75
        elif score.strip() == '及格':
            return 65
        else:
            return -1

    def __get_student_score(self, sheet):
        """get every student's score and credit"""
        lines = sheet[self.__score_scope[0]:self.__score_scope[1]]
        for student in lines:
            score = []
            credit = []
            i = -1

            for c in student:
                i += 1
                cla = self.get_evaluate(c.value)
                if cla != -1:
                    credit.append(self.__all_credit[i])
                    score.append(cla)

            self.__all_stu_credit.append(credit)
            self.__all_stu_scores.append(score)

    def __get_average(self):
        for i in range(len(self.__all_stu_names)):
            score_sum = 0
            cre_sum = sum(self.__all_stu_credit[i])
            for j in range(len(self.__all_stu_scores[i])):
                score_sum += self.__all_stu_scores[i][j] * self.__all_stu_credit[i][j]
            self.__average.append(score_sum / cre_sum)

    def out_put(self):
        """print the average of all the student"""
        name_ave = []
        for i in range(len(self.__all_stu_names)):
            name_ave.append([self.__all_stu_names[i], self.__average[i]])
        name_ave.sort(key=lambda x: x[1], reverse=True)
        print('人数： ', len(self.__all_stu_names))
        for i in range(len(name_ave)):
            print('({:0>2}) {}\t: {:.4f}分'.format(i + 1, name_ave[i][0], name_ave[i][1]))

    def save_to_file(self):
        """save the average and all score to the file"""
        with open('.\date.txt', 'w', encoding='utf-8') as f:
            for i in range(len(self.__all_stu_names)):
                date = "({}){}:\t{}\n{}\n{}\n\n".format(i + 1, self.__all_stu_names[i], self.__average[i],
                                                        self.__all_stu_credit[i], self.__all_stu_scores[i])
                f.write(date)


def main():
    file_name = '1405.xlsx'
    name_scope = ('E6', 'E36')
    credit_scope = ('F4', 'BJ4')
    score_scope = ('F6', 'BJ36')
    calculator = Calculator(file_name, name_scope, credit_scope, score_scope)
    calculator.start()
    calculator.save_to_file()

if __name__ == '__main__':
    main()
