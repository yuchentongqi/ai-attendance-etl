# AI-Attendance-ETL (钉钉考勤数据自动化处理工具)

基于 Python 实现的企业考勤数据自动化 ETL（抽取、转换、加载）与复杂 Excel 报表生成工具。本项目从真实的企业考勤数据处理需求出发，提取了核心的业务逻辑，并进行了完全的数据脱敏，做到开箱即用。

## ✨ 核心特性

- **API 鉴权封装**：实现钉钉 `access_token` 的自动获取与本地缓存机制，减少频繁请求，提升系统稳定性。
- **打卡数据聚合**：按员工聚合当天的“最早上班”与“最晚下班”打卡记录，并智能识别“缺卡”异常。
- **多源数据匹配**：支持将考勤人员名单与第三方业务台账 Excel（如焊工考试台账）按“姓名”进行自动关联匹配。
- **企业级 Excel 生成**：基于 `openpyxl`，严格实现微软雅黑字体、精准行高、列宽自适应及表头合并等复杂排版要求。
- **数据透视统计**：按“部门”与“工种”自动生成二维交叉统计表，清晰呈现各部门内不同工种的打卡人数。

## 📂 项目结构

- `dingtalk_auth.py`：封装钉钉鉴权逻辑，获取并缓存 Token。
- `attendance_api.py`：考勤接口调用及模拟打卡流水生成器。
- `generate_excel.py`：提取数据并生成符合企业排版要求的考勤报表。
- `statistics.py`：读取报表数据，进行多维度交叉汇总统计。
- `generate_mock_data.py`：脱敏测试数据生成脚本，用于生成 50 名虚构焊工的台账。

## 🚀 快速开始

### 1. 克隆项目并安装依赖
```bash
git clone https://github.com/yuchentongqi/ai-attendance-etl.git
cd ai-attendance-etl
pip install -r requirements.txt
