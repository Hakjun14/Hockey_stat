### code that put team_name and player name into result

import pandas as pd
import numpy as np
from glob import glob
import os
import natsort   # need to install natsor

files = os.listdir('./result')
files = natsort.natsorted(files)

entry = pd.read_csv("entry1.csv",header=0,index_col=0)
lst = pd.read_csv("game_lst.csv",header=0,index_col=0)

def main(input_file, entry, lst, idx):
  result = pd.read_excel(os.path.join('./result/',input_file), header=0,index_col=0) 
  ### put team name
  pd.set_option('display.max_rows',100)
  result.loc[0,'등번호'] = lst.loc['Game'+str(idx),'A']
  #print(result.loc['1','등번호'])
  result.loc[30,'등번호'] = lst.loc['Game'+str(idx),'B']
  #print(result)
  

  # 팀이름에 번호를 합친 키를 엔트리의 함수값과 비교하여 일치하는 선수의 이름 입력
  for i in range(len(result)):
      if i < 30 :
          player_num = lst.loc['Game'+str(idx),'A'] + str(result.loc[i,'등번호'])
          #print(player_num)
          if (entry['함수']==player_num).any():
              result.loc[i,"이름"] = entry.loc[entry['함수']==player_num, '이름'].values
      else:
          player_num = lst.loc['Game'+str(idx),'B'] + str(result.loc[i,'등번호'])
          #print(player_num)
          if (entry['함수']==player_num).any():
              result.loc[i,"이름"] = entry.loc[entry['함수']==player_num, '이름'].values

  save_dir = './team_result/result'+str(idx)+'.xlsx'
  result.to_excel(save_dir)  


for i, file in enumerate(files):
  print(file)
  main(file, entry, lst, i+1)    