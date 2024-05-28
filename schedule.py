import json
import re
from openpyxl import Workbook

wanted_courses = ["40240354", "10721011", "20240082", "30240063",
                  "40240513", "70612833", "30240382", "40260262",
                  "10260062", "10691572", "84760053", "30240292"]

options = {}
n = 0
data = {}
result = []
tables = []


def get_day_weeks(time_range):
    match = re.match(r"(\d+-\d+)\((.*?)周\)", time_range)
    options_list = match.group(2).split(",")

    week_list = []
    for elem in options_list:
        if elem == "全":
            week_list = [i for i in range(1, 17)]
        elif "-" in elem:
            week_range = elem.split("-")
            week_list += [i for i in range(int(week_range[0]), int(week_range[1])+1)]
        else:
            week_list += [int(elem)]
    return match.group(1), week_list


def get_time_dict(time):
    time_ranges = re.split(r",(?![^(]*\))", time)
    timetable = {}
    for time_range in time_ranges:
        time_range = time_range.strip()
        (day, weeks) = get_day_weeks(time_range)
        timetable[day] = weeks
    return timetable


def has_conflict(course1, course2):
    # 可能的格式：
    # 5-3(2,6,10,14-15周)
    # 4-6(1-11周)
    # 2-4(全周),2-3(全周)
    time1 = get_time_dict(course1["上课时间"])
    time2 = get_time_dict(course2["上课时间"])

    for key in time1.keys():
        if key in time2:
            if set(time1[key]) & set(time2[key]):
                return True
    return False


def convert_table(selected_courses):
    current_timetable = {}
    for selected_course in selected_courses:
        days = get_time_dict(selected_course["上课时间"]).keys()
        for day in days:
            if day not in current_timetable:
                current_timetable[day] = []
            current_timetable[day].append({
                "课程名": selected_course["课程名"],
                "课程号": selected_course["课程号"],
                "课序号": selected_course["课序号"],
                "主讲教师": selected_course["主讲教师"],
            })
    return current_timetable


def save_printable_timetable(timetable):
    current_table = [[None for _ in range(8)] for _ in range(7)]
    for time_slot, courses in timetable.items():
        col, row = map(int, time_slot.split("-"))
        current_table[row][col] = f"{courses[0]["课程名"]}({courses[0]["主讲教师"]})"

        for row in range(0, 7):
            current_table[row][0] = f"第{row}大节"
        for col in range(0, 8):
            current_table[0][col] = f"星期{col}"
        current_table[0][0] = ""
    tables.append(current_table)


def dfs(step, selected_courses):
    if step == n:
        current_timetable = convert_table(selected_courses)
        result.append(current_timetable)
        save_printable_timetable(current_timetable)
        return

    for current_course in options[wanted_courses[step]].values():
        flag = True
        for selected_course in selected_courses:
            if has_conflict(current_course, selected_course):
                flag = False
                break
        if flag:
            dfs(step + 1, selected_courses + [current_course])


if __name__ == "__main__":
    with open("courses.json", "r", encoding="utf-8") as f:
        data = json.load(f)
    for course in wanted_courses:
        options[course] = data[course]
    n = len(options)
    dfs(0, [])
    with open("timetable.json", "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=4)

    # 创建 Workbook 对象
    workbook = Workbook()
    # 删除默认的 "Sheet" 工作表
    workbook.remove(workbook.active)

    # 逐个写入 sheet
    for i, data in enumerate(tables, start=1):
        worksheet = workbook.create_sheet(f"方案{i}")
        for row in data:
            worksheet.append(row)

    # 保存 Excel 文件
    workbook.save("课程表 - 所有方案.xlsx")
