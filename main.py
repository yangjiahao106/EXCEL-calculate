# python3: calculate average
import openpyxl
import re


all_credit = []
all_stu_names = []
all_stu_credit = []
all_stu_scores = []
average = []
mins = []


def get_sheet():
    wb = openpyxl.load_workbook('1405.xlsx')
    sheet = wb.get_sheet_by_name(wb.get_sheet_names()[0])
    return sheet


def get_all_stu_names(sheet):
    cols = sheet['E6':'E36']
    for col in cols:
        all_stu_names.append(col[0].value.strip())


def get_all_credit(sheet):
    line1 = sheet['F4':'BJ4']  # return a tuple with one element
    for i in line1[0]:
        all_credit.append(eval(i.value))


def get_evaluate(score):
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


def get_student_score(sheet):
    """get every student's score and credit"""
    lines = sheet['F6':'BJ36']
    for student in lines:
        score = []
        credit = []
        i = -1

        for c in student:
            i += 1
            cla = get_evaluate(c.value)
            if cla != -1:
                credit.append(all_credit[i])
                score.append(cla)

        all_stu_credit.append(credit)
        all_stu_scores.append(score)


def get_average():
    for i in range(len(all_stu_names)):
        score_sum = 0
        cre_sum = sum(all_stu_credit[i])
        for j in range(len(all_stu_scores[i])):
            score_sum += all_stu_scores[i][j] * all_stu_credit[i][j]
        average.append(score_sum/cre_sum)
                    

def out_put():
    name_ave = []
    for i in range(len(all_stu_names)):
        name_ave.append([all_stu_names[i], average[i]])
    name_ave.sort(key=lambda x: x[1], reverse=True)
    print('人数： ', len(all_stu_names))
    for i in range(len(name_ave)):
        print('({:0>2}) {}\t: {:.4f}分'.format(i+1, name_ave[i][0], name_ave[i][1]))


def save_raw():
    with open('.\date.txt', 'w', encoding='utf-8') as f:
        for i in range(len(all_stu_scores)):
            date = "({}){}:\t{}\n{}\n{}\n\n".format(i+1, all_stu_names[i], average[i],
                                                    all_stu_credit[i], all_stu_scores[i])
            f.write(date)


def getmin():
    for ss in all_stu_scores:
        mins.append(min(ss))


def main():
    sheet = get_sheet()

    get_all_stu_names(sheet)
    get_all_credit(sheet)
    get_student_score(sheet)
    get_average()
    getmin()

    save_raw()
    out_put()

main()
pass