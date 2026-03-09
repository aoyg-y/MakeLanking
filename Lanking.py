from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from dotenv import load_dotenv
import re
import pandas as pd
import os 
import time
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

load_dotenv()

class MakeLanking():
    def __init__(self):
        self.email = os.getenv("EMAIL")
        self.password = os.getenv("PASSWORD")
        self.json_key = os.getenv("JSONKEY")
        #print(self.password)
        #print(self.email)
        #print(self.json_key)
        #男子1,女子2
        #各種目の陸マガのランキングのページのurlを指定するのに必要
        self.eid_dict = {1:{"100":10110,
                        "200":10210,
                        "400":10310,
                        "800":10410,
                        "1500":10510,
                        "5000":10910,
                        "10000":11010,
                        "half":11160,
                        "110H":11810,
                        "400H":11910,
                        "3000sc":12110,
                        "400R":12410,
                        "1600R":12610,
                        "HJ":13010,
                        "PV":13110,
                        "LJ":13210,
                        "TJ":13310,
                        "SP":13410,
                        "DT":13510,
                        "HT":13610,
                        "JT":13710,
                        "5000W":12250,
                        "10000W":12270,
                        "Dec":14210},
                    2:{"100":10110,
                        "200":10210,
                        "400":10310,
                        "800":10410,
                        "1500":10510,
                        "3000":10810,
                        "5000":10910,
                        "10000":11010,
                        "half":11160,
                        "110H":11900,
                        "400H":21910,
                        "3000sc":22110,
                        "400R":12410,
                        "1600R":12610,
                        "HJ":13010,
                        "PV":13110,
                        "LJ":13210,
                        "TJ":13310,
                        "SP":13430,
                        "DT":13530,
                        "HT":13630,
                        "JT":23710,
                        "5000W":12250,
                        "10000W":12270,
                        "Hep":24250}}
    
    def make_result_df(self,url):
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        driver.get(url)

        mail = driver.find_element(By.ID,"mlmail_address")
        mail.send_keys(self.email)

        password = driver.find_element(By.ID,"mlpassword")
        password.send_keys(self.password)

        buttons = driver.find_elements(By.NAME,"btn_type")
        buttons[1].click()
        time.sleep(2)


        #htmlを取得してbeautifulsoupで解析
        html = driver.page_source.encode('utf-8')
        soup = BeautifulSoup(html, "html.parser")
        #順位の親要素から選手の情報の入ったtrタグを取得
        ranks = soup.find_all("td",class_="rank")
        
        results_list = []
        for rank in ranks:
            ls_ = []
            result = rank.parent
            record = result.find("td",class_="record").text.replace("\n","").replace("\t","")
            player = result.find("td",class_="player").text.replace("\n","").replace("\t","").replace("\u3000","")
            univ = result.find_all("input",id=re.compile("belonging"))[1]["value"]
            date = result.find("td",class_="date").text.replace("\n","").replace("\t","")
            try:
                #風速の無い種目についてはwindをNoneとする
                wind = result.find("td",class_="wind").text.replace("\n","").replace("\t","")
            except AttributeError:
                wind = None
            ls_ += [player,univ,record,wind,date]
            results_list.append(ls_)
        columns = ["選手名","大学","記録","風","日時"]
            
        return(pd.DataFrame(results_list,columns=columns))
    
    #使っていない
    def make_csv(self,year,sex,eid,filename):
        #cd直下にfilenameと一致するディレクトリを作っておく必要あり
        url =  f'https://rikumaga.com/records/ranking?pref_id_list=&area_id_list=&year%5Byear%5D={year}&y=nendo&room=1&bg=2&g=3&peid=&eid[]={self.eid_dict[sex][eid]}&g={sex}&rr=1000'
        path = os.getcwd()
        df = self.select_univ(url,sex)
        filepath = os.path.join(path, filename, f"{sex}_{year}_{eid}_ranking.csv")
        df.to_csv(filepath, encoding="shift-jis")

    def select_sex_eid_dict(self,competition):
        #competition

        #1京カレ
        if competition==1:
            univs = {1:["立命大","京産大","同志社大","京大","びわこ成蹊スポーツ大","明治国際医療大","京都教大","仏教大","龍谷大","びわこ学院大","京都府立大","滋賀大","京都橘大","京都工繊大","京都府立医大",],
                     2:["立命大","京産大","同志社大","京大","びわこ成蹊スポーツ大","明治国際医療大","京都教大","仏教大","龍谷大","びわこ学院大","京都府立大","滋賀大","京都橘大","京都工繊大","京都府立医大","京都女大","京都光華女大","同志社女大"]}
            eids = {1:["100","200","400","800","1500","5000","10000","110H","400H","3000sc","400R","1600R","HJ","PV","LJ","TJ","SP","DT","HT","JT"],
                    2:["100","200","400","800","1500","5000","10000","110H","400H","3000sc","400R","1600R","HJ","PV","LJ","TJ","SP","DT","HT","JT"]}
        
        #2関カレ    
        elif competition==2:
            univs = {1:["関大","立命大","関学大","京産大","大体大","同志社大","京大","近大","びわこ成蹊スポーツ大","大阪教大","摂南大","天理大"],
                     2:["関大","立命大","関学大","京産大","大体大","同志社大","京大","近大","びわこ成蹊スポーツ大","大阪教大","摂南大","天理大","仏教大","びわこ学院大","京都府立大","滋賀大","園田学園女大","甲南大","京都教大","大阪学院大","大阪成蹊大","大阪国際大","京都橘大","京都工繊大","京都府立医大","京都光華女大","同志社女大","東大阪大","武庫川女大","明治国際医療大","関西福祉大","関西外語大","大阪経大","京都女子大","大阪公立大","神戸大"]}
            eids = {1:["100","200","400","800","1500","5000","10000","half","110H","400H","3000sc","10000W","400R","1600R","HJ","PV","LJ","TJ","SP","DT","HT","JT","Dec"],
                    2:["100","200","400","800","1500","5000","10000","110H","400H","3000sc","10000W","400R","1600R","HJ","PV","LJ","TJ","SP","DT","HT","JT","Hep"]}

        #3七大戦    
        elif competition==3:
            univs = {1:["京大","名古屋大","九州大","東大","東北大","北大","阪大"],
                     2:["京大","名古屋大","九州大","東大","東北大","北大","阪大"]}
            eids = {1:["100","200","400","800","1500","5000","110H","400H","3000sc","5000W","400R","1600R","HJ","PV","LJ","TJ","SP","DT","HT","JT"],
                    2:["100","400","800","3000","110H","400R","HJ","LJ","SP","JT"]}
        
        #4同志社戦
        elif competition==4:
            univs = {1:["京大","同志社大"]}
            eids = {1:["100","200","400","800","1500","5000","110H","400H","400R","1600R","HJ","PV","LJ","TJ","SP","DT","HT","JT"]}

        #5東大戦
        elif competition==5:
            univs = {1:["京大","東大"],
                     2:["京大","東大"]}
            eids = {1:["100","200","400","800","1500","5000","110H","400H","5000W","400R","1600R","HJ","PV","LJ","TJ","SP","DT","HT","JT"],
                    2:["100","400","800","3000","400R","LJ","SP"]}
        
        #test用はここをいじる
        elif competition=="test":
            univs = {1:["関大","立命大","関学大","京産大","大体大","同志社大","京大","近大","びわこ成蹊スポーツ大","大阪教大","摂南大","天理大"]}
            eids = {1:["100","200"]}
        
        
        return univs, eids

    def make_lanking_spreadsheet(self,competition,year,ss_name):

        #dir_path = os.path.dirname(__file__) # 作業フォルダの取得
        #gc = gspread.oauth(credentials_filename=os.path.join(dir_path, "client_secret_MakeLanking.json"))
        #wb = gc.create(str(ss_name))

        #service accountでのapi認証に変更
        # https://note.com/kohaku935/n/nf69f13012eb8
        # 認証情報の設定
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        dir_path = os.path.dirname(os.path.abspath('__file__'))
        creds = ServiceAccountCredentials.from_json_keyfile_name(os.path.join(dir_path,self.json_key),scope)

        # クライアントの作成
        gs = gspread.authorize(creds)
        spreadsheet_list = gs.openall()
        ss_names = [ss.title for ss in spreadsheet_list]
        #もし同じ名前のシートが存在したら上書き、なければ作成
        if ss_name in ss_names:
            wb = gs.open(ss_name)
            # すべてのワークシートを削除（再作成準備）
            sheets = wb.worksheets()
            for sheet in sheets[1:]:
                wb.del_worksheet(sheet)
            ws = wb.sheet1
            ws.update_title("summery")
            ws.clear()
        else:
            wb = gs.create(str(ss_name), "1P6RvGgPk-6pTPsTH1aW91X7UxSvj0jWh")
        
        univs, eids = self.select_sex_eid_dict(competition)
                
        sheet_num = 1
        
        for sex in eids.keys():
            point_df = pd.DataFrame()
            try:
                for eid in eids[sex]:
                    sheet_name = f"{'男子' if sex==1 else '女子'}{eid}"

                    url =  f'https://rikumaga.com/records/ranking?pref_id_list=&area_id_list=&year%5Byear%5D={year}&y=nendo&room=1&bg=2&g=3&peid=&eid[]={self.eid_dict[sex][eid]}&g={sex}&rr=1000'
                    df = self.make_result_df(url)
                    df = df[df["大学"].isin(univs[sex])]
                    df.reset_index(inplace=True,drop=True)
                    df.index = df.index + 1
                    print(sheet_name,len(df),"人")
                    #得点計算
                    eid_point_df = df[df.index<=8]#順位は場合による
                    eid_point_df["得点"] = 9 - eid_point_df.index
                    point_df = pd.concat([point_df,eid_point_df])

                    # 既存シート確認
                    try:
                        old_ws = wb.worksheet(sheet_name)
                        wb.del_worksheet(old_ws)
                    except gspread.exceptions.WorksheetNotFound:
                        pass
                    ws = wb.add_worksheet(sheet_name, rows=500, cols=10)
                    set_with_dataframe(ws,df,include_index=True)

            #女子がない時
            except KeyError:
                continue
            school_scores = point_df.groupby('大学')['得点'].sum().reset_index()
            school_scores = school_scores.sort_values(by='得点', ascending=False).reset_index()
            school_scores.index = school_scores.index + 1
            school_scores.drop("index",axis=1,inplace=True)
            ws = wb.sheet1
            ws.append_row([f"{'男子' if sex==1 else '女子'}"])
            ws.append_rows(school_scores.reset_index().values.tolist())


ML=MakeLanking()
ML.make_lanking_spreadsheet(1,2025,"2026京都インカレ")
