import pandas as pd
from faker import Faker
import random
import os

# 初始化 Faker（設定 locale 為台灣）
fake = Faker('zh_TW')

# 讀取原本的 CSV
input_path = os.path.join("Person", "female.xlsx")
df = pd.read_excel(input_path)

# 準備新欄位
new_data = []

for _, row in df.iterrows():
    name = row["姓名"]
    gender = row["性別"]
    birthday = row["出生年月日"]  # 應該是民國年格式
    id_number = row["身分證字號"]
    phone = "0" + str(row["行動電話"])
    email = row["電子郵件信箱"]

    zip_code = fake.postcode()
    address = fake.address().replace("\n", "")
    bank_code = "7000021"
    bank_name = "中華郵政股份有限公司郵政存簿儲金"
    account_number = ''.join(random.choices(
        "0123456789", k=random.randint(8, 14)))
    org = "臺北市政府消防局"

    new_row = [
        org, name, gender, birthday, id_number, phone, email,
        zip_code, address, zip_code, address,
        f"{bank_code}　{bank_name}",
        account_number, name  # 匯款戶名與姓名相同
    ]

    # 加入空白欄位（從住宅電話開始）
    new_row += [""] * 8

    new_data.append(new_row)

# 欄位名稱
columns = [
    "組織", "講師姓名\n<必填>", "性別\n<必填>", "出生日期\n民國年/月/日\n例：100/01/01\n<必填>",
    "身分證字號\n註：若為外國籍，請帶入護照號碼。若非中華民國國民身分證字號，系統將自動判斷為外國籍身份。\n<必填>",
    "行動電話\n格式：0911222333\n<必填>", "電子郵件信箱\n<必填>",
    "戶籍郵遞區號", "戶籍地址\n例：縣市鄉鎮詳細地址\n<必填>",
    "通訊郵遞區號", "通訊地址\n例：縣市鄉鎮詳細地址",
    "銀行與分行代碼", "帳戶帳號\n*不含銀行代碼、帳號不須加\"-\"、帳號最多14碼（中國信託帳號僅有12碼）",
    "匯款帳戶戶名\n*請填寫本人帳戶",
    "住宅電話", "辦公室電話", "身高", "體重", "血型", "緊急聯絡人", "聯絡人電話", "聯絡人關係"
]

# 建立 DataFrame 並匯出為 Excel
output_df = pd.DataFrame(new_data, columns=columns)
output_df.to_excel("./Teacher/講師清單.xlsx", index=False)
print("Excel 檔案已產生：講師清單.xlsx")
