import csv
import random
from datetime import datetime, timedelta
from faker import Faker
import os
import pandas as pd

fake = Faker("zh_TW")

# 身分證英文碼與其對應代碼
id_prefix_map = {
    "A": 10, "B": 11, "C": 12, "D": 13, "E": 14, "F": 15,
    "G": 16, "H": 17, "J": 18, "K": 19, "L": 20, "M": 21,
    "N": 22, "P": 23, "Q": 24, "R": 25, "S": 26, "T": 27,
    "U": 28, "V": 29, "X": 30, "Y": 31, "W": 32, "Z": 33, "I": 34, "O": 35
}
id_prefix_list = list(id_prefix_map.keys())

# 驗證碼計算公式


def generate_valid_taiwan_id(gender: str):
    while True:
        prefix = random.choice(id_prefix_list)
        code = id_prefix_map[prefix]
        d1, d2 = divmod(code, 10)
        digits = [d1, d2]
        digits.append(2 if gender == "F" else 1)
        digits += [random.randint(0, 9) for _ in range(6)]

        # 計算驗證碼
        weights = [1, 9, 8, 7, 6, 5, 4, 3, 2, 1]
        total = sum(d * w for d, w in zip(digits, weights))
        checksum = (10 - (total % 10)) % 10
        id_number = f"{prefix}{''.join(map(str, digits[1:]))}{checksum}"
        return id_number

# 隨機生成中文三字姓名


def generate_name():
    family_names = ["陳", "林", "黃", "張", "李", "王", "吳", "劉", "蔡", "楊"]
    given_names = ["怡君", "慧琳", "雅婷", "淑芬", "玉珍", "麗華", "佳蓉", "雅惠", "秀琴", "靜怡",
                   "怡慧", "欣怡", "芳儀", "婉君", "佩珊", "佳慧", "怡君", "欣惠", "君潔", "品如"]
    return random.choice(family_names) + random.choice(given_names)

# 隨機生成手機號碼


def generate_phone():
    return "09" + "".join([str(random.randint(0, 9)) for _ in range(8)])

# 隨機產生民國年生日


def generate_birth():
    start_date = datetime(1950, 1, 1)
    end_date = datetime(2000, 12, 31)
    birth_date = fake.date_between(start_date=start_date, end_date=end_date)
    roc_year = birth_date.year - 1911
    return f"{roc_year}/{birth_date.month}/{birth_date.day}"


# 產生100筆資料
data = []
for i in range(1, 101):
    name = generate_name()
    gender = "女"
    id_number = generate_valid_taiwan_id(gender)
    birth = generate_birth()
    email = f"user{i:03}@gmail.com"
    phone = generate_phone()
    data.append([name, id_number, gender, birth, email, phone])

# 產生CSV檔案
file_path = os.path.join("data", "Ttaiwan_female_100.csv")
with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(["姓名", "身分證字號", "性別", "出生年月日", "Email", "手機"])
    writer.writerows(data)

file_path

# 產出成 Excel
df = pd.DataFrame(data, columns=["姓名", "身分證字號",
                  "性別", "出生年月日", "電子郵件信箱", "行動電話"])
output_path = os.path.join("data", "Ttaiwan_female_100.xlsx")
df.to_excel(output_path, index=False, engine="openpyxl")
