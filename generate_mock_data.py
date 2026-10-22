import openpyxl
from openpyxl.styles import Font, Alignment
import random

def create_mock_ledger():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "焊工考试统计表"

    # 写入表头（还原你考官给的真实表头结构，但内容做脱敏）
    headers = ["考试日期", "序号", "焊工姓名", "焊工号", "拟考项目", "考试材料", "材料规格", 
               "焊接材料", "焊评编号", "试件编号", "ATF-证书编号", "证书原件", "外观检查结果", 
               "RT检查结果", "", "焊接对象", "焊接方法", "母材类别P-No.", "厚度范围", "管径范围", "焊接位置", "备注"]
    
    ws.merge_cells('A1:V1')
    ws['A1'] = "某海外项目焊工考试台账（Mock Data 脱敏版）"
    ws['A1'].font = Font(name='Microsoft YaHei', size=14, bold=True)
    
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.font = Font(name='Microsoft YaHei', size=12, bold=True)

    # 假数据
    first_names = ["张", "李", "王", "赵", "刘", "陈", "杨", "黄", "周", "吴"]
    last_names = ["伟", "刚", "强", "磊", "洋", "勇", "军", "杰", "涛", "明"]
    processes = ["GTAW", "SMAW", "GTAW+SMAW", "SAW", "FCAW"]
    p_nos = ["1", "8", "10H", "45"]
    teams = ["储罐1队", "储罐2队", "火炬队", "管道队", "电仪"]

    # 生成 50 个假员工
    for i in range(1, 51):
        row = i + 3
        name = random.choice(first_names) + random.choice(last_names)
        welder_id = f"MOCK-{i:03d}"
        process = random.choice(processes)
        p_no = random.choice(p_nos)
        team = random.choice(teams)

        ws.cell(row=row, column=1, value="2024-01-01")
        ws.cell(row=row, column=2, value=i)
        ws.cell(row=row, column=3, value=name)
        ws.cell(row=row, column=4, value=welder_id)
        
        # 对应真实表格的索引：16(Process), 17(P-No), 18(Thickness), 19(Diameter), 20(Position)
        ws.cell(row=row, column=17, value=process)
        ws.cell(row=row, column=18, value=p_no)
        ws.cell(row=row, column=19, value="≤21.94")
        ws.cell(row=row, column=20, value="≥73")
        ws.cell(row=row, column=21, value="All")
        
        ws.cell(row=row, column=25, value=team) # 备注列（部门）

    wb.save("mock_ledger.xlsx")
    print("✅ 假数据已成功生成：mock_ledger.xlsx")

if __name__ == "__main__":
    create_mock_ledger()