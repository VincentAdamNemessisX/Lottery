import xlwings as xs

students_list = []
teachers_list = []


def update_student_list(sheet_sec, range_start, range_end, range_cls):
    for i in range(range_start, range_end + 1):
        student_dict = {'no.': int(sheet_sec.range('A' + str(i)).value), 'name': sheet_sec.range('B' + str(i)).value,
                        'class': range_cls, 'fix': 'false'}
        students_list.append(student_dict)


def update_teacher_list(sheet_sec, range_start, range_end, range_cls):
    for i in range(range_start, range_end + 1):
        teacher_dict = {'no.': int(sheet_sec.range('B' + str(i)).value), 'name': sheet_sec.range('C' + str(i)).value,
                        'class': range_cls, 'fix': 'false'}
        teachers_list.append(teacher_dict)


def get_section_1_grade_1(sheet_sec):
    update_student_list(sheet_sec, 3, 40, '一年级一班')
    update_student_list(sheet_sec, 46, 81, '一年级二班')
    update_student_list(sheet_sec, 87, 126, '一年级三班')
    update_student_list(sheet_sec, 132, 172, '一年级四班')
    update_student_list(sheet_sec, 178, 217, '一年级五班')
    return students_list


def get_section_1_grade_2(sheet_sec):
    update_student_list(sheet_sec, 3, 46, '二年级一班')
    update_student_list(sheet_sec, 52, 97, '二年级二班')
    update_student_list(sheet_sec, 103, 147, '二年级三班')
    return students_list


def get_section_1_grade_3(sheet_sec):
    update_student_list(sheet_sec, 3, 44, '三年级一班')
    update_student_list(sheet_sec, 50, 90, '三年级二班')
    update_student_list(sheet_sec, 96, 137, '三年级三班')
    update_student_list(sheet_sec, 143, 182, '三年级四班')
    update_student_list(sheet_sec, 188, 231, '三年级五班')
    update_student_list(sheet_sec, 237, 277, '三年级六班')
    update_student_list(sheet_sec, 283, 324, '三年级七班')
    update_student_list(sheet_sec, 330, 372, '三年级八班')
    return students_list


def get_section_2_grade_4(sheet_sec):
    update_student_list(sheet_sec, 3, 46, '四年级一班')
    update_student_list(sheet_sec, 52, 95, '四年级二班')
    update_student_list(sheet_sec, 101, 144, '四年级三班')
    update_student_list(sheet_sec, 150, 193, '四年级四班')
    update_student_list(sheet_sec, 199, 241, '四年级五班')
    update_student_list(sheet_sec, 247, 289, '四年级六班')
    update_student_list(sheet_sec, 295, 341, '四年级七班')
    update_student_list(sheet_sec, 348, 393, '四年级八班')
    update_student_list(sheet_sec, 399, 442, '四年级九班')
    update_student_list(sheet_sec, 448, 492, '四年级十班')
    return students_list


def get_section_2_grade_5(sheet_sec):
    update_student_list(sheet_sec, 3, 47, '五年级一班')
    update_student_list(sheet_sec, 53, 97, '五年级二班')
    update_student_list(sheet_sec, 103, 147, '五年级三班')
    update_student_list(sheet_sec, 153, 198, '五年级四班')
    update_student_list(sheet_sec, 204, 250, '五年级五班')
    update_student_list(sheet_sec, 256, 299, '五年级六班')
    update_student_list(sheet_sec, 305, 350, '五年级七班')
    update_student_list(sheet_sec, 356, 400, '五年级八班')
    update_student_list(sheet_sec, 406, 450, '五年级九班')
    update_student_list(sheet_sec, 456, 500, '五年级十班')
    return students_list


def get_section_3_grade_6(sheet_sec):
    update_student_list(sheet_sec, 3, 48, '六年级一班')
    update_student_list(sheet_sec, 54, 100, '六年级二班')
    update_student_list(sheet_sec, 106, 152, '六年级三班')
    update_student_list(sheet_sec, 158, 202, '六年级四班')
    update_student_list(sheet_sec, 208, 252, '六年级五班')
    update_student_list(sheet_sec, 258, 303, '六年级六班')
    update_student_list(sheet_sec, 309, 352, '六年级七班')
    update_student_list(sheet_sec, 358, 403, '六年级八班')
    update_student_list(sheet_sec, 409, 453, '六年级九班')
    update_student_list(sheet_sec, 459, 502, '六年级十班')
    update_student_list(sheet_sec, 508, 553, '六年级十一班')
    update_student_list(sheet_sec, 559, 603, '六年级十二班')
    return students_list


def get_section_3_grade_7(sheet_sec):
    update_student_list(sheet_sec, 3, 62, '七年级一班')
    update_student_list(sheet_sec, 68, 116, '七年级二班')
    update_student_list(sheet_sec, 122, 167, '七年级三班')
    update_student_list(sheet_sec, 173, 221, '七年级四班')
    update_student_list(sheet_sec, 227, 275, '七年级五班')
    return students_list


def get_section_3_grade_8(sheet_sec):
    update_student_list(sheet_sec, 3, 55, '八年级一班')
    update_student_list(sheet_sec, 61, 106, '八年级二班')
    update_student_list(sheet_sec, 112, 154, '八年级三班')
    update_student_list(sheet_sec, 160, 204, '八年级四班')
    return students_list


def get_section_1_data(bok):
    get_section_1_grade_1(bok.sheets[1])
    get_section_1_grade_2(bok.sheets[2])
    get_section_1_grade_3(bok.sheets[3])
    # print(len(students_list))
    param_define = "let member = "
    students_list_str = str(students_list).replace("'true'", "true")
    with open('../js/member_section_1.js', 'w', encoding='utf-8') as f:
        f.write(param_define + students_list_str)


def get_section_2_data(bok):
    get_section_2_grade_4(bok.sheets[4])
    get_section_2_grade_5(bok.sheets[5])
    # print(len(students_list))
    param_define = "let member = "
    students_list_str = str(students_list).replace("'true'", "true")
    with open('../js/member_section_2.js', 'w', encoding='utf-8') as f:
        f.write(param_define + students_list_str)


def get_section_3_data(bok):
    get_section_3_grade_6(bok.sheets[6])
    get_section_3_grade_7(bok.sheets[7])
    get_section_3_grade_8(bok.sheets[8])
    # print(len(students_list))
    param_define = "let member = "
    students_list_str = str(students_list).replace("'true'", "true")
    with open('../js/member_section_3.js', 'w', encoding='utf-8') as f:
        f.write(param_define + students_list_str)


def get_section_4_data(bok):
    update_teacher_list(bok.sheets[10], 1, 153, '教师部')
    param_define = "let member = "
    teachers_list_str = str(teachers_list)
    with open('../js/member_section_4.js', 'w', encoding='utf-8') as f:
        f.write(param_define + teachers_list_str)


def get_all_section_data(bok):
    get_section_1_grade_1(bok.sheets[1])
    get_section_1_grade_2(bok.sheets[2])
    get_section_1_grade_3(bok.sheets[3])
    get_section_2_grade_4(bok.sheets[4])
    get_section_2_grade_5(bok.sheets[5])
    get_section_3_grade_6(bok.sheets[6])
    get_section_3_grade_7(bok.sheets[7])
    get_section_3_grade_8(bok.sheets[8])
    update_teacher_list(bok.sheets[10], 1, 153, '教师部')
    param_define = "let member = "
    students_list_str = str(students_list).replace("'true'", "true")
    students_list_str = students_list_str.replace("]", ",")
    teachers_list_str = str(teachers_list).replace("'true'", "true")
    teachers_list_str = teachers_list_str.replace("[", "")
    with open('../js/member.js', 'w', encoding='utf-8') as f:
        f.write(param_define + students_list_str + teachers_list_str)


if __name__ == '__main__':
    with xs.App(visible=False) as app:
        book = app.books.open('../data/all_students_data.xlsx')
        students_list.clear()
        get_section_1_data(book)
        students_list.clear()
        get_section_2_data(book)
        students_list.clear()
        get_section_3_data(book)
        students_list.clear()
        get_section_4_data(book)
        teachers_list.clear()
        get_all_section_data(book)
