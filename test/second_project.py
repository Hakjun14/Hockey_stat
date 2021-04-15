### code that rearranges the Integrated data by school and player
encoding = 'utf8'

import pandas as pd
import numpy as np
from glob import glob
import os
import natsort   # need to install natsor
import openpyxl

## idea : name을 적용시킨 intergrated data를 load, match to each sheet
## 1.player data 2.data by team 두 개의 function define -> 3.goalie data after finish 1,2


#print(entry)
def create_team_sheet(team_name, entry):
    #sheet 생성
    sheet = pd.DataFrame(columns=['등번호','이름', '포지션', '경기수', '포인트', 'GWS','골', 
                                    '어시', 'P/GP','G/GP', 'A/GP', '+', '-', '+/-','SOG',
                                    'PPG','PPA', 'SHG', 'SHA', 'GWG', 'GBG', 'GTG', 'ENG', 
                                    'PIM', 'SO횟수', 'SO성공', 'SO성공률', 'Face-off','Face-off win', 'Percentage Faceoff'],
                       index=range(25))

    #sheet에 해당 팀 관련 기본정보 수록하기 ()
    #아니면 게임 기록별로 update에 적용하면 되게해도 될듯?아 엔트리를 대입하면 되겠다.
    n=0
    for i in range(len(entry)):
        if team_name == entry.loc[i,'소속'] :
            sheet.loc[n,'등번호']=entry.loc[i,'등번호']
            sheet.loc[n,'이름']=entry.loc[i,'이름']
            n+=1
    return sheet

def update(file, stat):
    # excel 파일 열고, data.loc[0,'등번호']가 일치하는 시트 제목의 시트에 데이터 업데이트
    
    master = pd.read_excel(stat, sheet_name=None)
    #print(master.keys()) #시트 제목은 이 dict의 (key,value) 중 key에 해당

    data = pd.read_excel(os.path.join('./team_result/',file), header=0, index_col=0)

    for i, a in enumerate(master.keys()):
        #home
        if a == data.loc[0,'등번호']:
            home_team = a
            home_plate = pd.read_excel(stat, sheet_name = a, index_col=0)
            home_plate.iloc[:-3,3:27] = home_plate.iloc[:-3,3:27].fillna(0)
        #print(data.loc[:25,['이름']].values)
            for name in home_plate['이름']:
                if name in data.loc[:25,['이름']].values: 
                    home_plate.loc[home_plate['이름']==name,['포지션']] = data.loc[data['이름']==name,['포지션']].values
                    home_plate.loc[home_plate['이름']==name,['포인트']] += data.loc[data['이름']==name,['포인트']].values
                    home_plate.loc[home_plate['이름']==name,['GWS']] += data.loc[data['이름']==name,['GWS']].values
                    home_plate.loc[home_plate['이름']==name,['골']] += data.loc[data['이름']==name,['골']].values
                    home_plate.loc[home_plate['이름']==name,['어시']] += data.loc[data['이름']==name,['어시']].values
                    home_plate.loc[home_plate['이름']==name,['+']] += data.loc[data['이름']==name,['+']].values
                    home_plate.loc[home_plate['이름']==name,['-']] += data.loc[data['이름']==name,['-']].values
                    home_plate.loc[home_plate['이름']==name,['+/-']] += data.loc[data['이름']==name,['+/-']].values
                    home_plate.loc[home_plate['이름']==name,['SOG']] += data.loc[data['이름']==name,['SOG']].values
                    home_plate.loc[home_plate['이름']==name,['PPG']] += data.loc[data['이름']==name,['PPG']].values
                    home_plate.loc[home_plate['이름']==name,['PPA']] += data.loc[data['이름']==name,['PPA']].values
                    home_plate.loc[home_plate['이름']==name,['SHG']] += data.loc[data['이름']==name,['SHG']].values
                    home_plate.loc[home_plate['이름']==name,['SHA']] += data.loc[data['이름']==name,['SHA']].values
                    home_plate.loc[home_plate['이름']==name,['GWG']] += data.loc[data['이름']==name,['GWG']].values
                    home_plate.loc[home_plate['이름']==name,['GBG']] += data.loc[data['이름']==name,['GBG']].values
                    home_plate.loc[home_plate['이름']==name,['GTG']] += data.loc[data['이름']==name,['GTG']].values
                    home_plate.loc[home_plate['이름']==name,['ENG']] += data.loc[data['이름']==name,['ENG']].values
                    home_plate.loc[home_plate['이름']==name,['PIM']] += data.loc[data['이름']==name,['PIM']].values

                    if data.loc[data['이름']==name,['Y-N']].values == 'Y':
                        home_plate.loc[home_plate['이름']==name,['경기수']] += 1

                    if home_plate.loc[home_plate['이름']==name,['경기수']].values:
                        home_plate.loc[home_plate['이름']==name,['P/GP']] = home_plate.loc[home_plate['이름']==name,['포인트']].values/home_plate.loc[home_plate['이름']==name,['경기수']].values
                        home_plate.loc[home_plate['이름']==name,['G/GP']] = home_plate.loc[home_plate['이름']==name,['골']].values/home_plate.loc[home_plate['이름']==name,['경기수']].values
                        home_plate.loc[home_plate['이름']==name,['A/GP']] = home_plate.loc[home_plate['이름']==name,['어시']].values/home_plate.loc[home_plate['이름']==name,['경기수']].values
                        #print(plate.loc[plate['이름']==name, ['포지션']])



        #away  
        if a == data.loc[30,'등번호']:
            away_team = a
            #print(list(master.values())[i])
            away_plate = pd.read_excel(stat, sheet_name = a, index_col=0)
            away_plate.iloc[:-3,3:27] = away_plate.iloc[:-3,3:27].fillna(0)
  
            for name in away_plate['이름']:
                if name in data.loc[31:57,['이름']].values: 
                    away_plate.loc[away_plate['이름']==name,['포지션']] = data.loc[data['이름']==name,['포지션']].values
                    away_plate.loc[away_plate['이름']==name,['포인트']] += data.loc[data['이름']==name,['포인트']].values
                    away_plate.loc[away_plate['이름']==name,['GWS']] += data.loc[data['이름']==name,['GWS']].values
                    away_plate.loc[away_plate['이름']==name,['골']] += data.loc[data['이름']==name,['골']].values
                    away_plate.loc[away_plate['이름']==name,['어시']] += data.loc[data['이름']==name,['어시']].values
                    away_plate.loc[away_plate['이름']==name,['+']] += data.loc[data['이름']==name,['+']].values
                    away_plate.loc[away_plate['이름']==name,['-']] += data.loc[data['이름']==name,['-']].values
                    away_plate.loc[away_plate['이름']==name,['+/-']] += data.loc[data['이름']==name,['+/-']].values
                    away_plate.loc[away_plate['이름']==name,['SOG']] += data.loc[data['이름']==name,['SOG']].values
                    away_plate.loc[away_plate['이름']==name,['PPG']] += data.loc[data['이름']==name,['PPG']].values
                    away_plate.loc[away_plate['이름']==name,['PPA']] += data.loc[data['이름']==name,['PPA']].values
                    away_plate.loc[away_plate['이름']==name,['SHG']] += data.loc[data['이름']==name,['SHG']].values
                    away_plate.loc[away_plate['이름']==name,['SHA']] += data.loc[data['이름']==name,['SHA']].values
                    away_plate.loc[away_plate['이름']==name,['GWG']] += data.loc[data['이름']==name,['GWG']].values
                    away_plate.loc[away_plate['이름']==name,['GBG']] += data.loc[data['이름']==name,['GBG']].values
                    away_plate.loc[away_plate['이름']==name,['GTG']] += data.loc[data['이름']==name,['GTG']].values
                    away_plate.loc[away_plate['이름']==name,['ENG']] += data.loc[data['이름']==name,['ENG']].values
                    away_plate.loc[away_plate['이름']==name,['PIM']] += data.loc[data['이름']==name,['PIM']].values

                    if data.loc[data['이름']==name,['Y-N']].values == 'Y':
                        away_plate.loc[away_plate['이름']==name,['경기수']] += 1

                    if away_plate.loc[away_plate['이름']==name,['경기수']].values :
                        away_plate.loc[away_plate['이름']==name,['P/GP']] = away_plate.loc[away_plate['이름']==name,['포인트']].values/away_plate.loc[away_plate['이름']==name,['경기수']].values
                        away_plate.loc[away_plate['이름']==name,['G/GP']] = away_plate.loc[away_plate['이름']==name,['골']].values/away_plate.loc[away_plate['이름']==name,['경기수']].values
                        away_plate.loc[away_plate['이름']==name,['A/GP']] = away_plate.loc[away_plate['이름']==name,['어시']].values/away_plate.loc[away_plate['이름']==name,['경기수']].values


    return home_plate, home_team, away_plate, away_team
    
def excel_rewriter(data_source, df_name, sheet_name):
    book = openpyxl.load_workbook(data_source)
    #book.remove(book[sheet_name])
    writer = pd.ExcelWriter(data_source, engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    df_name.to_excel(writer, sheet_name, index=True)
    writer.save()
    #os.rename(data_source, target_file)

def base_file_generation():
    files = os.listdir('./team_result')
    files = natsort.natsorted(files)
    entry = pd.read_csv("entry1.csv",header=0)

    x1 = create_team_sheet('경기고', entry)
    x2 = create_team_sheet('경복고', entry)
    x3 = create_team_sheet('경성고', entry)
    x4 = create_team_sheet('보성고', entry)
    x5 = create_team_sheet('중동고', entry)
    x6 = create_team_sheet('광성고', entry)

    save_dir = './2nd'
    file_nm = "df_final.xlsx"
    xlxs_dir = os.path.join(save_dir, file_nm) 

    with pd.ExcelWriter(xlxs_dir) as writer:
        x1.to_excel(writer, sheet_name = '경기고')
        x2.to_excel(writer, sheet_name = '경복고')
        x3.to_excel(writer, sheet_name = '경성고')
        x4.to_excel(writer, sheet_name = '보성고')
        x5.to_excel(writer, sheet_name = '중동고')
        x6.to_excel(writer, sheet_name = '광성고') 

    return xlxs_dir    

def main():
    master = base_file_generation()
    files = os.listdir('./team_result')
    files = natsort.natsorted(files)
    
    for i, file in enumerate(files):
        print(file)
        h_plate, h_team, a_plate, a_team = update(file, master)
        excel_rewriter(master, h_plate, h_team)
        excel_rewriter(master, a_plate, a_team)

    df_all = pd.read_excel(master, sheet_name = None, index_col=0)
    concatted_df = pd.concat(df_all, ignore_index=True)
    distinct_df = concatted_df.dropna(subset=['이름','포지션'])
    excel_rewriter(master, distinct_df, 'Player')

main()