import requests
import random
import openpyxl
from datetime import datetime
from dingtalk_auth import get_cached_token

def get_attendance_records(user_ids, work_date=None):
    if work_date is None:
        work_date = datetime.now().strftime("%Y-%m-%d")
    token = get_cached_token()
    url = "https://oapi.dingtalk.com/attendance/listRecord"
    params = {"access_token": token}
    body = {
        "userIds": user_ids[:50],
        "checkDateFrom": f"{work_date} 00:00:00",
        "checkDateTo": f"{work_date} 23:59:59",
        "isI18n": False
    }
    resp = requests.post(url, params=params, json=body, timeout=10)
    data = resp.json()
    if data.get("errcode") != 0:
        raise RuntimeError(f"获取考勤失败: {data}")
    return data.get("recordresult", [])

def read_real_employees(file_path="mock_ledger.xlsx"):
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb["焊工考试统计表"]
        employees = {}
        for row in ws.iter_rows(min_row=4, values_only=True):
            name = row[2]
            welder_no = row[3]
            process = row[16]
            team = row[24]
            if not name or not welder_no:
                continue
            if welder_no not in employees:
                employees[welder_no] = {
                    "name": str(name).strip(),
                    "userid": str(welder_no).strip(),
                    "work_type": f"焊工({process})" if process else "焊工",
                    "dept_name": team if team else "管道队"
                }
        return list(employees.values())
    except Exception as e:
        print(f"读取 Excel 失败: {e}")
        return []

def get_mock_attendance_records(employees, work_date=None):
    if work_date is None:
        work_date = datetime.now().strftime("%Y-%m-%d")
    records = []
    for emp in employees:
        uid = emp["userid"]
        seed_val = sum(ord(c) for c in uid)
        r = random.Random(seed_val)
        is_abnormal = r.random() < 0.1
        
        in_h = r.choice([7, 8])
        in_m = r.randint(0, 59) if in_h == 7 else r.randint(0, 30)
        in_time = f"{work_date} {in_h:02d}:{in_m:02d}:00"
        
        out_h = r.choice([17, 18])
        out_m = r.randint(30, 59) if out_h == 17 else r.randint(0, 30)
        out_time = f"{work_date} {out_h:02d}:{out_m:02d}:00"

        records.append({"userId": uid, "userCheckTime": in_time, "checkType": "OnDuty", "timeResult": "Late" if is_abnormal else "Normal", "locationResult": "Normal"})
        if not is_abnormal:
            records.append({"userId": uid, "userCheckTime": out_time, "checkType": "OffDuty", "timeResult": "Normal", "locationResult": "Normal"})
    return records

def summarize_attendance(records):
    result = {}
    for r in records:
        uid = r["userId"]
        result.setdefault(uid, {"clock_in": None, "clock_out": None, "status": []})
        t = r["userCheckTime"]
        if r["checkType"] == "OnDuty":
            if result[uid]["clock_in"] is None or t < result[uid]["clock_in"]:
                result[uid]["clock_in"] = t
        elif r["checkType"] == "OffDuty":
            if result[uid]["clock_out"] is None or t > result[uid]["clock_out"]:
                result[uid]["clock_out"] = t
    for uid, info in result.items():
        if not info["clock_in"]: info["status"].append("缺上班卡")
        if not info["clock_out"]: info["status"].append("缺下班卡")
    return result