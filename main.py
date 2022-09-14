import requests
from io import BytesIO
import pandas as pd

User_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:69.0) Gecko/20100101 firefox/69.0"

pocket = ["연결 재무상태표", "연결 손익계산서", "연결 포괄손익계산서"]

def download_excel(company, period, rcp_no, dcm_no, sheets=pocket):
    url = "https://dart.fss.or.kr/pdf/download/excel.do?rcp_no={}&dcm_no={}&lang=ko".format(rcp_no, dcm_no)

    resp = requests.get(url, headers={"user-agent": User_agent})
    table = BytesIO(resp.content)
    
    for sheet in sheets:
        data = pd.read_excel(table, sheet_name=sheet, skiprows=5)
        data.to_csv("_".join([company, period, sheet])+".csv", encoding="euc-kr")

df = pd.read_csv('pocket.txt')
for period, rcp_no, dcm_no in zip(df["period"].values, df["rcp_no"].values, df["dcm_no"].values):
    download_excel("삼성전자", str(period), str(rcp_no), str(dcm_no))