import random
from datetime import datetime
from faker import Faker
import os
import pandas as pd
import string

fake = Faker("zh_TW")

# 想產生的資料筆數
num_records = 600

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
        # 隨機選擇一個英文字母前綴
        prefix = random.choice(id_prefix_list)
        # 取得對應的數字代碼
        code = id_prefix_map[prefix]
        # 將代碼拆成兩個數字
        d1, d2 = divmod(code, 10)

        # 性別碼，男性為1，女性為2
        gender_code = 2 if gender == "女" else 1

        # 產生後7碼的隨機數字（不含檢查碼）
        remaining_digits = [random.randint(0, 9) for _ in range(7)]

        # 計算檢查碼
        # 使用正確的權重：第一碼英文字母拆成兩個數字乘以權重 1, 9，性別碼乘以8，後面依序乘以7,6,5,4,3,2,1
        sum_val = d1 * 1 + d2 * 9 + gender_code * 8

        # 加上後7碼乘以對應權重
        weights = [7, 6, 5, 4, 3, 2, 1]
        for i in range(7):
            sum_val += remaining_digits[i] * weights[i]

        # 計算檢查碼：總和除以10的餘數，若為0則檢查碼為0，否則為10減去餘數
        checksum = (10 - (sum_val % 10)) % 10

        # 組合身分證字號：英文字母 + 性別碼 + 後7碼 + 檢查碼
        id_number = f"{prefix}{gender_code}{''.join(map(str, remaining_digits))}{checksum}"

        return id_number

# 隨機生成中文三字姓名


def generate_name():
    family_names = ["鍾", "林", "張", "段", "丘", "華", "殷", "曹", "杜", "武", "鄭", "向", "章", "周", "温", "郭", "李", "康", "邱", "洪", "葉", "蘇", "陶", "方", "莊", "辛", "余", "蔡", "秦", "雷", "徐", "吳", "熊", "王", "全", "涂",
                    "簡", "辜", "錢", "袁", "史", "馮", "田", "利", "何", "謝", "朱", "孫", "藍", "高", "石", "梁", "傅", "包", "唐", "汪", "董", "尹", "彭", "葛", "麥", "黃", "孟", "江", "崔", "賀", "丁", "蔣", "楊", "蕭", "賴", "樊", "侯", "孔"]
    given_names = ["誠鑰", "傅傑", "柏志", "錦江", "嘉志", "震寰", "奕軒", "宗鑫", "衛傑", "榮鵬", "瑋澤", "傲霜", "奕勝", "餘樂", "海彬", "光京", "藝燦", "其丁", "元兵", "大鵬", "鎧瑞", "曦楠", "昱熹", "兆琨", "逸希", "業超", "晉燊", "倬鳴", "安博", "顯灝", "浩泓", "文焯", "宙羲", "紀瑋", "玉樺", "景程", "皚晅", "嘉彥", "賢昊", "祐鍚", "懿洛", "立珩", "豊翔", "學杭", "弘昱", "丞孝", "虹岳", "璉傑", "吉佃", "家鋒", "耀聰",
                   "佳作", "培倫", "佑柏", "育烜", "毓國", "少懷", "秉聿", "政橒", "力申", "冠慶", "煒文", "仲崴", "文奕", "道霖", "焯行", "子泰", "丞韜", "文攸", "積穗", "言名", "泓達", "學知", "也昍", "諾城", "籽竣", "宣楠", "振鴻", "嘉譽", "文烜", "鳴晧", "睿林", "安榮", "晏而", "謹煒", "京華", "明均", "文仙", "榮琛", "啓哲", "柯言", "驁馳", "健康", "樑歡", "弘瑞", "泓迪", "志鋮", "誠越", "書陽", "以丹", "子建", "子荻", "祥榮", "競搏", "縣"]
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
    now = datetime.now().strftime("%Y%m%d%H%M%S")  # 年月日時分秒
    rand_str = ''.join(random.choices(
        string.ascii_lowercase + string.digits, k=5))  # 5碼亂數

    name = generate_name()
    gender = "男"
    id_number = generate_valid_taiwan_id(gender)
    birth = generate_birth()
    email = f"user_{now}_{rand_str}@gmail.com"
    phone = generate_phone()
    data.append([name, id_number, gender, birth, email, phone])

# 產出成 Excel
df = pd.DataFrame(data, columns=["姓名", "身分證字號",
                  "性別", "出生年月日", "電子郵件信箱", "行動電話"])
output_path = os.path.join("Person", "male.xlsx")
df.to_excel(output_path, index=False, engine="openpyxl")
