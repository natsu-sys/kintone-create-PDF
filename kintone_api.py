import requests
import json
import pandas as pd


class KintoneAPI:
    def __init__(self, domain, app_id, api_token):
        self.domain = domain.rstrip("/")
        self.app_id = app_id
        self.api_token = api_token

        # Header สำหรับ GET (ดึงข้อมูล)
        self.get_headers = {
            "X-Cybozu-API-Token": self.api_token
        }

        # Header สำหรับ POST/PUT (ถ้าต้องใช้อนาคต)
        self.json_headers = {
            "X-Cybozu-API-Token": self.api_token,
            "Content-Type": "application/json"
        }

    def fetch_all_records(self, query=""):
        """ดึงข้อมูลทั้งหมดจาก Kintone แบบแบ่งหน้า"""
        all_records = []
        offset = 0
        limit = 500

        while True:
            url = f"{self.domain}/k/v1/records.json"

            # ถ้า user ไม่ส่ง query มาจะได้: limit 500 offset 0
            if query:
                q = f"{query} limit {limit} offset {offset}"
            else:
                q = f"limit {limit} offset {offset}"

            params = {
                "app": self.app_id,
                "query": q
            }

            # GET เท่านั้น
            res = requests.get(url, headers=self.get_headers, params=params)
            data = res.json()

            # debug ถ้ามี error
            if "records" not in data:
                print("======== DEBUG REQUEST ========")
                print("URL:", url)
                print("Params:", params)
                print("Response:", data)
                print("===============================")
                raise Exception("Kintone API Error: " + json.dumps(data, ensure_ascii=False))

            recs = data["records"]
            all_records.extend(recs)

            # ถ้ามีข้อมูลน้อยกว่า limit → หมดแล้ว
            if len(recs) < limit:
                break

            offset += limit

        return all_records

    def to_dataframe(self, records):
        """แปลง Kintone records → pandas DataFrame แบบ flatten"""
        flat = []

        for r in records:
            row = {}
            for key, value in r.items():
                if isinstance(value, dict) and "value" in value:
                    row[key] = value["value"]
                else:
                    row[key] = value
            flat.append(row)

        return pd.DataFrame(flat)