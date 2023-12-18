import xlwings as xs

students_list = []


def update_student_list(sheet_sec, range_start, range_end, range_cls):
    for i in range(range_start, range_end + 1):
        student_dict = {'no.': int(sheet_sec.range('A' + str(i)).value), 'name': sheet_sec.range('C' + str(i)).value,
                        'class': range_cls, 'fix': 'false'}
        students_list.append(student_dict)


def get_section_1_grade_1(sheet_sec):
    student_class = sheet_sec.range('A2').value.replace(' ', '')
    if student_class == '一年级一班':
        update_student_list(sheet_sec, 4, 41, student_class)
    student_class = sheet_sec.range('A46').value.replace(' ', '')
    if student_class == '一年级二班':
        update_student_list(sheet_sec, 48, 87, student_class)
    student_class = sheet_sec.range('A91').value.replace(' ', '')
    if student_class == '一年级三班':
        update_student_list(sheet_sec, 93, 131, student_class)
    student_class = sheet_sec.range('A136').value.replace(' ', '')
    if student_class == '一年级四班':
        update_student_list(sheet_sec, 138, 180, student_class)
    student_class = sheet_sec.range('A185').value.replace(' ', '')
    if student_class == '一年级五班':
        update_student_list(sheet_sec, 187, 222, student_class)
    return students_list


def get_section_1_grade_2(sheet_sec):
    student_class = sheet_sec.range('A2').value.replace(' ', '')
    if student_class == '二年级一班':
        update_student_list(sheet_sec, 4, 47, student_class)
    student_class = sheet_sec.range('A51').value.replace(' ', '')
    if student_class == '二年级二班':
        update_student_list(sheet_sec, 53, 94, student_class)
    student_class = sheet_sec.range('A98').value.replace(' ', '')
    if student_class == '二年级三班':
        update_student_list(sheet_sec, 100, 141, student_class)


def get_section_1_grade_3(sheet_sec):
    student_class = sheet_sec.range('A2').value.replace(' ', '')
    if student_class == '三年级一班':
        update_student_list(sheet_sec, 4, 45, student_class)
    student_class = sheet_sec.range('A49').value.replace(' ', '')
    if student_class == '三年级二班':
        update_student_list(sheet_sec, 51, 92, student_class)
    student_class = sheet_sec.range('A97').value.replace(' ', '')
    if student_class == '三年级三班':
        update_student_list(sheet_sec, 99, 141, student_class)
    student_class = sheet_sec.range('A145').value.replace(' ', '')
    if student_class == '三年级四班':
        update_student_list(sheet_sec, 147, 188, student_class)
    student_class = sheet_sec.range('A192').value.replace(' ', '')
    if student_class == '三年级五班':
        update_student_list(sheet_sec, 194, 239, student_class)
    student_class = sheet_sec.range('A243').value.replace(' ', '')
    if student_class == '三年级六班':
        update_student_list(sheet_sec, 245, 285, student_class)
    student_class = sheet_sec.range('A290').value.replace(' ', '')
    if student_class == '三年级七班':
        update_student_list(sheet_sec, 292, 333, student_class)
    student_class = sheet_sec.range('A337').value.replace(' ', '')
    if student_class == '三年级八班':
        update_student_list(sheet_sec, 339, 380, student_class)


def get_section_2_grade_4(sheet_sec):
    student_class = sheet_sec.range('A2').value.replace(' ', '')
    if student_class == '四年级一班':
        update_student_list(sheet_sec, 4, 48, student_class)
    student_class = sheet_sec.range('A53').value.replace(' ', '')
    if student_class == '四年级二班':
        update_student_list(sheet_sec, 55, 99, student_class)
    student_class = sheet_sec.range('A103').value.replace(' ', '')
    if student_class == '四年级三班':
        update_student_list(sheet_sec, 105, 150, student_class)
    student_class = sheet_sec.range('A154').value.replace(' ', '')
    if student_class == '四年级四班':
        update_student_list(sheet_sec, 156, 198, student_class)
    student_class = sheet_sec.range('A202').value.replace(' ', '')
    if student_class == '四年级五班':
        update_student_list(sheet_sec, 204, 250, student_class)
    student_class = sheet_sec.range('A255').value.replace(' ', '')
    if student_class == '四年级六班':
        update_student_list(sheet_sec, 257, 300, student_class)
    student_class = sheet_sec.range('A305').value.replace(' ', '')
    if student_class == '四年级七班':
        update_student_list(sheet_sec, 307, 351, student_class)
    student_class = sheet_sec.range('A355').value.replace(' ', '')
    if student_class == '四年级八班':
        update_student_list(sheet_sec, 357, 400, student_class)
    student_class = sheet_sec.range('A405').value.replace(' ', '')
    if student_class == '四年级九班':
        update_student_list(sheet_sec, 407, 449, student_class)
    student_class = sheet_sec.range('A454').value.replace(' ', '')
    if student_class == '四年级十班':
        update_student_list(sheet_sec, 456, 498, student_class)


def get_section_2_grade_5(sheet_sec):
    student_class = sheet_sec.range('A2').value.replace(' ', '')
    if student_class == '五年级一班':
        update_student_list(sheet_sec, 4, 46, student_class)
    student_class = sheet_sec.range('A51').value.replace(' ', '')
    if student_class == '五年级二班':
        update_student_list(sheet_sec, 53, 97, student_class)
    student_class = sheet_sec.range('A101').value.replace(' ', '')
    if student_class == '五年级三班':
        update_student_list(sheet_sec, 103, 148, student_class)
    student_class = sheet_sec.range('A153').value.replace(' ', '')
    if student_class == '五年级四班':
        update_student_list(sheet_sec, 155, 200, student_class)
    student_class = sheet_sec.range('A204').value.replace(' ', '')
    if student_class == '五年级五班':
        update_student_list(sheet_sec, 206, 253, student_class)
    student_class = sheet_sec.range('A258').value.replace(' ', '')
    if student_class == '五年级六班':
        update_student_list(sheet_sec, 260, 303, student_class)
    student_class = sheet_sec.range('A308').value.replace(' ', '')
    if student_class == '五年级七班':
        update_student_list(sheet_sec, 310, 354, student_class)
    student_class = sheet_sec.range('A358').value.replace(' ', '')
    if student_class == '五年级八班':
        update_student_list(sheet_sec, 360, 405, student_class)
    student_class = sheet_sec.range('A410').value.replace(' ', '')
    if student_class == '五年级九班':
        update_student_list(sheet_sec, 412, 455, student_class)
    student_class = sheet_sec.range('A460').value.replace(' ', '')
    if student_class == '五年级十班':
        update_student_list(sheet_sec, 462, 508, student_class)


def get_section_3_grade_6(sheet_sec):
    student_class = sheet_sec.range('A2').value.replace(' ', '')
    if student_class == '六年级一班':
        update_student_list(sheet_sec, 4, 49, student_class)
    student_class = sheet_sec.range('A54').value.replace(' ', '')
    if student_class == '六年级二班':
        update_student_list(sheet_sec, 56, 102, student_class)
    student_class = sheet_sec.range('A107').value.replace(' ', '')
    if student_class == '六年级三班':
        update_student_list(sheet_sec, 109, 155, student_class)
    student_class = sheet_sec.range('A160').value.replace(' ', '')
    if student_class == '六年级四班':
        update_student_list(sheet_sec, 162, 206, student_class)
    student_class = sheet_sec.range('A211').value.replace(' ', '')
    if student_class == '六年级五班':
        update_student_list(sheet_sec, 213, 258, student_class)
    student_class = sheet_sec.range('A262').value.replace(' ', '')
    if student_class == '六年级六班':
        update_student_list(sheet_sec, 264, 309, student_class)
    student_class = sheet_sec.range('A314').value.replace(' ', '')
    if student_class == '六年级七班':
        update_student_list(sheet_sec, 316, 360, student_class)
    student_class = sheet_sec.range('A365').value.replace(' ', '')
    if student_class == '六年级八班':
        update_student_list(sheet_sec, 367, 412, student_class)
    student_class = sheet_sec.range('A417').value.replace(' ', '')
    if student_class == '六年级九班':
        update_student_list(sheet_sec, 419, 464, student_class)
    student_class = sheet_sec.range('A470').value.replace(' ', '')
    if student_class == '六年级十班':
        update_student_list(sheet_sec, 472, 514, student_class)
    student_class = sheet_sec.range('A518').value.replace(' ', '')
    if student_class == '六年级十一班':
        update_student_list(sheet_sec, 520, 566, student_class)
    student_class = sheet_sec.range('A570').value.replace(' ', '')
    if student_class == '六年级十二班':
        update_student_list(sheet_sec, 572, 616, student_class)


def get_section_3_grade_7(sheet_sec):
    student_class = sheet_sec.range('A2').value.replace(' ', '')
    if student_class == '七年级一班':
        update_student_list(sheet_sec, 4, 60, student_class)
    student_class = sheet_sec.range('A65').value.replace(' ', '')
    if student_class == '七年级二班':
        update_student_list(sheet_sec, 67, 116, student_class)
    student_class = sheet_sec.range('A120').value.replace(' ', '')
    if student_class == '七年级三班':
        update_student_list(sheet_sec, 122, 169, student_class)
    student_class = sheet_sec.range('A173').value.replace(' ', '')
    if student_class == '七年级四班':
        update_student_list(sheet_sec, 175, 225, student_class)
    student_class = sheet_sec.range('A230').value.replace(' ', '')
    if student_class == '七年级五班':
        update_student_list(sheet_sec, 232, 277, student_class)


def get_section_3_grade_8(sheet_sec):
    student_class = sheet_sec.range('A2').value.replace(' ', '')
    if student_class == '八年级一班':
        update_student_list(sheet_sec, 4, 55, student_class)
    student_class = sheet_sec.range('A60').value.replace(' ', '')
    if student_class == '八年级二班':
        update_student_list(sheet_sec, 62, 109, student_class)
    student_class = sheet_sec.range('A114').value.replace(' ', '')
    if student_class == '八年级三班':
        update_student_list(sheet_sec, 116, 161, student_class)
    student_class = sheet_sec.range('A166').value.replace(' ', '')
    if student_class == '八年级四班':
        update_student_list(sheet_sec, 168, 213, student_class)


def get_section_1_data(bok):
    get_section_1_grade_1(bok.sheets[1])
    get_section_1_grade_2(bok.sheets[2])
    get_section_1_grade_3(bok.sheets[3])
    # print(len(students_list))
    param_define = "let member = "
    students_list_str = str(students_list).replace("'true'", "true")
    with open('../js/member_section_1.js', 'w') as f:
        f.write(param_define + students_list_str)


def get_section_2_data(bok):
    get_section_2_grade_4(bok.sheets[4])
    get_section_2_grade_5(bok.sheets[5])
    # print(len(students_list))
    param_define = "let member = "
    students_list_str = str(students_list).replace("'true'", "true")
    with open('../js/member_section_2.js', 'w') as f:
        f.write(param_define + students_list_str)


def get_section_3_data(bok):
    get_section_3_grade_6(bok.sheets[6])
    get_section_3_grade_7(bok.sheets[7])
    get_section_3_grade_8(bok.sheets[8])
    # get_section_3_grade_9(bok.sheets[9])
    # print(len(students_list))
    param_define = "let member = "
    students_list_str = str(students_list).replace("'true'", "true")
    with open('../js/member_section_3.js', 'w') as f:
        f.write(param_define + students_list_str)


def get_all_section_data(bok):
    get_section_1_grade_1(bok.sheets[1])
    get_section_1_grade_2(bok.sheets[2])
    get_section_1_grade_3(bok.sheets[3])
    get_section_2_grade_4(bok.sheets[4])
    get_section_2_grade_5(bok.sheets[5])
    get_section_3_grade_6(bok.sheets[6])
    get_section_3_grade_7(bok.sheets[7])
    get_section_3_grade_8(bok.sheets[8])
    # get_section_3_grade_9(bok.sheets[9])
    # print(len(students_list))
    param_define = "let member = "
    students_list_str = str(students_list).replace("'true'", "true")
    with open('../js/member.js', 'w') as f:
        f.write(param_define + students_list_str)


if __name__ == '__main__':
    with xs.App(visible=False) as app:
        book = app.books.open('all_students_data.xlsx')
        students_list.clear()
        get_section_1_data(book)
        students_list.clear()
        get_section_2_data(book)
        students_list.clear()
        get_section_3_data(book)
        students_list.clear()
        get_all_section_data(book)
