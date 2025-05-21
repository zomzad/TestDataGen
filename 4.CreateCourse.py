import pandas as pd
import random
from datetime import datetime, timedelta

# === 參數設定 ===
course_file = "./Course/course.xlsx"
teacher_file = "./data/taiwan_female_100.csv"
output_file = "./Course/final_course_schedule.xlsx"

start_date = datetime(2024, 8, 1)
end_date = datetime(2024, 8, 31)

classrooms = ["6202", "6204"]
periods = [
    "晨間第一節 (06:10-07:00)", "晨間第二節 (07:10-08:00)", "第一節 (08:10-09:00)", "第二節 (09:10-10:00)", "第三節 (10:10-11:00)", "第四節 (11:10-12:00)",
    "中午第一節 (12:00-12:50)", "中午第二節 (13:00-13:50)"
    "第五節 (14:00-14:40)", "第六節 (15:00-15:50)", "第七節 (16:00-16:50)", "第八節 (17:00-17:50)",
    "夜間第一節 (18:00-18:50)", "夜間第二節 (19:00-19:50)", "夜間第三節 (20:00-20:50)", "夜間第四節 (21:00-21:50)"]

# === 讀取檔案 ===
df_course = pd.read_excel(course_file)
df_teacher = pd.read_csv(teacher_file)

# === 產出資料 ===
rows = []
date_range = pd.date_range(start_date, end_date).to_list()

for _, row in df_course.iterrows():
    course_code = row['COURSE_CODE']
    course_name = row['COURSE_NAME']
    course_date = random.choice(date_range)
    roc_date = f"{course_date.year - 1911}/{course_date.month:02d}/{course_date.day:02d}"
    classroom = random.choice(classrooms)
    period = random.choice(periods)
    teacher_pair = df_teacher.sample(2)

    for _, teacher in teacher_pair.iterrows():
        teacher_id = teacher["身分證字號"][:6]
        teacher_name = teacher["姓名"]

        rows.append({
            "課程代碼": course_code,
            "課程名稱": course_name,
            "上課日期": roc_date,
            "教室": classroom,
            "開始上課節次": period,
            "課程時數": 1,
            "講師身分證字號前6碼": teacher_id,
            "講師姓名": teacher_name,
            "助教身分證字號前6碼": "",
            "助教姓名": "",
            "併班": ""
        })

# === 輸出到 Excel ===
df_result = pd.DataFrame(rows)
df_result.to_excel(output_file, index=False)

print(f"已成功產出排課表: {output_file}")
