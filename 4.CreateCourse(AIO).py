import pandas as pd
from datetime import datetime
import random

# 所有講師在同一堂課中

# === 檔案設定 ===
course_file = "./Course/course.xlsx"
teacher_file = "./Course/teacher.xlsx"
output_file = "final_course_schedule.xlsx"

# === 日期範圍設定 ===（實際只會選一天）
start_date = datetime(2025, 8, 1)
end_date = datetime(2025, 8, 31)

# 固定使用的課程與時段
fixed_period = "晨間第一節 (06:10-07:00)"
fixed_date = random.choice(pd.date_range(start_date, end_date).to_list())
roc_date = datetime(2024, 8, 8)  # 固定一天
# roc_date = f"{fixed_date.year - 1911}/{fixed_date.month:02d}/{fixed_date.day:02d}"
classroom = "6202"  # 你可以改成其他教室名稱

# === 讀取資料 ===
df_course = pd.read_excel(course_file)
df_teacher = pd.read_excel(teacher_file)

# === 只取一堂課（可依需要指定課程代碼或名稱）===
selected_course = random.choice(df_course.to_dict('records'))   # 預設選第一筆
# 或你可使用條件選課：df_course[df_course['COURSE_CODE'] == 'F00010'].iloc[0]

# === 產生所有講師的課表資料 ===
rows = []

# 取得老師與助教
teacher_list = df_teacher.to_dict('records')
main_teacher = teacher_list[0]  # 第一位當老師
assistants = teacher_list[1:]  # 其餘當助教

rows.append({
    "課程代碼": selected_course['COURSE_CODE'],
    "課程名稱": selected_course['COURSE_NAME'],
    "上課日期": roc_date,
    "教室": classroom,
    "開始上課節次": fixed_period,
    "課程時數": 1,
    "講師身分證字號前6碼": main_teacher["ID_NO"][:6],
    "講師姓名": main_teacher["CHT_NAME"],
    "助教身分證字號前6碼": "",
    "助教姓名": "",
    "併班": ""
})

for _, teacher in df_teacher.iterrows():
    teacher_id = teacher["ID_NO"][:6]
    teacher_name = teacher["CHT_NAME"]

    rows.append({
        "課程代碼": selected_course['COURSE_CODE'],
        "課程名稱": selected_course['COURSE_NAME'],
        "上課日期": roc_date,
        "教室": classroom,
        "開始上課節次": fixed_period,
        "課程時數": 1,
        "講師身分證字號前6碼": teacher_id,
        "講師姓名": teacher_name,
        "助教身分證字號前6碼": "",
        "助教姓名": "",
        "併班": ""
    })

# === 輸出成 Excel ===
df_result = pd.DataFrame(rows)
df_result.to_excel(output_file, index=False)

print(f"已成功產出單一堂課的排程，共{len(df_result)}筆，輸出檔名為：{output_file}")
