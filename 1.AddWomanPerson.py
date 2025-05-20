import random
from datetime import datetime
from faker import Faker
import os
import pandas as pd
import string

fake = Faker("zh_TW")

# 想產生的資料筆數
num_records = 300

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
        digits.append(2 if gender == "女" else 1)
        digits += [random.randint(0, 9) for _ in range(6)]

        # 計算驗證碼
        weights = [1, 9, 8, 7, 6, 5, 4, 3, 2, 1]
        total = sum(d * w for d, w in zip(digits, weights))
        checksum = (10 - (total % 10)) % 10
        id_number = f"{prefix}{''.join(map(str, digits[1:]))}{checksum}"
        return id_number

# 隨機生成中文三字姓名


def generate_name():
    family_names = ["翁", "白", "韋", "朱", "阮", "雷", "趙", "黃", "丘", "蕭", "盧", "謝", "段", "甘", "華",
                    "關", "胡", "柯", "林", "易", "洪", "毛", "殷", "包", "顧", "樊", "姜", "熊", "石", "佘", "姚", "全", "李"]
    given_names = ["芝訓", "泳如", "恩慧", "楚盈", "得梅", "雯昕", "映凱", "鏡涵", "玲鈴", "予婕", "典鳳", "夏梅", "鬱珍", "詩酉", "雨春", "路瑤", "姝懿", "自若", "柏穎", "佳悅",
                   "子茜", "穎嘉", "子淇", "詠寧", "頤庭", "湘喻", "薇穎", "聖心", "欣琳", "謹恩", "柔緗", "柳沄", "家恩", "丞麗", "禮萱", "怡鈞", "穎賢", "清怡", "映傑", "偉淇", "薇芩", "嘉湞"]
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
for i in range(1, num_records+1):
    now = datetime.now().strftime("%Y%m%d%H%M%S")  # 年月日時分秒
    rand_str = ''.join(random.choices(
        string.ascii_lowercase + string.digits, k=5))  # 5碼亂數

    name = generate_name()
    gender = "女"
    id_number = generate_valid_taiwan_id(gender)
    birth = generate_birth()
    email = f"user_{now}_{rand_str}@gmail.com"
    phone = generate_phone()
    data.append([name, id_number, gender, birth, email, phone])

# 產出成 Excel
df = pd.DataFrame(data, columns=["姓名", "身分證字號",
                  "性別", "出生年月日", "電子郵件信箱", "行動電話"])
output_path = os.path.join("Person", "female.xlsx")
df.to_excel(output_path, index=False, engine="openpyxl")
