import pandas as pd
import requests
from datetime import datetime, timedelta, timezone
import openpyxl

excel_path = "Data.xlsx"

# 1. Durum sekmesini oku (DistrictCode'u string olarak al!)
df_durum = pd.read_excel(excel_path, sheet_name="durum", dtype={"DistrictCode": str})

# 2. EndDate = yarın
tomorrow = (datetime.now(timezone.utc) + timedelta(days=1)).strftime('%Y-%m-%dT00:00:00.0Z')

# 3. Excel dosyasını yükle
wb = openpyxl.load_workbook(excel_path)

for index, row in df_durum.iterrows():
    district_code = row['DistrictCode']
    start_date = pd.to_datetime(row['StartDate']).strftime('%Y-%m-%dT00:00:00.0Z')

    # API URL'si
    url = (
        f"https://api.opet.com.tr/api/fuelprices/prices/archive"
        f"?DistrictCode={district_code}&StartDate={start_date}"
        f"&EndDate={tomorrow}&IncludeAllProducts=true"
    )

    # 4. API çağrısı
    response = requests.get(url)
    if response.status_code != 200:
        print(f"[❌] {district_code} için API hatası: {response.status_code}")
        continue

    # 5. Gelen veriyi işle
    try:
        data = response.json()
    except ValueError:
        print(f"[❌] {district_code} için JSON verisi alınamadı.")
        continue

    records = []

    for day_item in data:
        for price in day_item.get("prices", []):
            if price.get("productName") == "Motorin UltraForce":
                records.append({
                    "priceDate": price.get("priceDate"),
                    "productName": price.get("productName"),
                    "amount": price.get("amount")
                })

    if not records:
        print(f"[⚠️] {district_code} için uygun fiyat verisi bulunamadı.")
        continue

    df_new = pd.DataFrame(records)

    # 6. Tarih dönüştürme ve zaman dilimini silme
    df_new["priceDate"] = pd.to_datetime(df_new["priceDate"], errors="coerce").dt.tz_localize(None)

    sheet_name = str(district_code)

    try:
        df_old = pd.read_excel(excel_path, sheet_name=sheet_name)
        if df_old.empty:
            df_combined = df_new
        else:
            df_combined = pd.concat([df_old, df_new], ignore_index=True).drop_duplicates()
    except:
        df_combined = df_new

    # 7. Sayfaya yaz
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_combined.to_excel(writer, sheet_name=sheet_name, index=False)

    # 8. StartDate hücresini güncelle
    valid_dates = df_combined["priceDate"].dropna()
    if not valid_dates.empty:
        latest_date = valid_dates.max().strftime('%Y-%m-%d')
        df_durum.at[index, "StartDate"] = latest_date
    else:
        print(f"[⚠️] {district_code} için geçerli tarih bulunamadı.")

# 9. Durum sekmesini güncelle
with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_durum.to_excel(writer, sheet_name="durum", index=False)

print("✅ Tüm veriler başarıyla işlendi ve güncellendi.")
