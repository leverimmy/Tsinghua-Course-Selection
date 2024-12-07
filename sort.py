import json

# 读取 JSON 数据
with open("courses_2024_12_07_20_00_00.json", "r", encoding="utf-8") as f:
    course_data = json.load(f)

# 计算每个课程的 报名总人数/可选容量 比值并存储
course_ratios = {}
for course_id, course_info in course_data.items():
    for section_id, section_info in course_info.items():
        total_enrollment = section_info["报名总人数"]
        ratio = total_enrollment / section_info["可选容量"]
        course_ratios[(course_id, section_id)] = ratio

# 对比值进行排序
sorted_courses = sorted(course_ratios.items(), key=lambda x: x[1], reverse=True)

# 输出排序结果
print("排序结果:")
for (course_id, section_id), ratio in sorted_courses:
    print(f"课程号: {course_id}, 课序号: {section_id}, 比值: {ratio:.3f}")
