#本程式會撈取同資料夾中的csv檔案，抓取lep資料並匯出為excel檔案
#V2, 新增recipe欄位，並模組化，其中parseoneLEPfile可以供日後呼叫方便

# import lepcsvparser as lep
import pandas as pd
import sys
import os
from datetime import datetime
import argparse
import glob
import subprocess
from subprocess import Popen
import streamlit as st

#name_of_wb='LEP_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.csv'
#dir_path = os.path.dirname(os.path.realpath(__file__))
#↑獲取當前資料夾名稱然後存成dir_path變數
#all_file_name = os.listdir(dir_path)
#↑讀取資料夾內所有檔案名稱然後放進all_file_name這個list裡

#path為想要爬取檔案的所在資料夾完整路徑(D:\coding\LEP_offline_parser)，filename為檔案名稱(ex:POLYPR205.Csv)
def parseoneLEPfile(filename):#(path,filename):
    #url=path+'/'+filename #url=完整檔案名稱=檔案路徑+檔案名稱
    #waferid=filename[:filename.index(".")] #抓取檔案名稱，因LEP檔案名稱為LOT+slot，所以直接抓檔案名稱即可作為wafer身分辨別
    #df=pd.read_fwf(url) #讀取檔案，存在名為df的變數中
    #print(filename)
    try:
        waferid=os.path.basename(filename)
        df=pd.read_fwf(filename) #讀取檔案，存在名為df的變數中
        srs=df[list(df.columns)[0]]  #取出df的資料存為series
        df_waferid=pd.DataFrame({'LOT_slot':[waferid]})
        
        recipe_raw=[i for i in srs if i.find('"Recipe Name:",')!=-1] #找到recipe name
        recipe_raw=recipe_raw[0]
        recipe=recipe_raw[recipe_raw.index(',')+1:] #找到,所在的index
        recipe=recipe[1:-1] #去除"的字樣
        df_recipe=pd.DataFrame({'Recipe':[recipe]}) #將recipe資訊存成dataframe

        dia_index=srs.index[srs=='"Dia","DiaB"']+1 #找到"Dia","DiaB"於series中的index number，將index number+1即為資料欄位
        diameter=list(srs[dia_index])[0][:list(srs[dia_index])[0].find(',')] #找到","於string中的index number
        df_diameter=pd.DataFrame({'Diameter':[diameter]})#將diameter完整資訊存成dataframe
        
        diamax_index=srs.index[srs=='"Max, Diff, Dir"']+1 #找到"Dia","DiaB"於series中的index number，將index number+1即為資料欄位
        dia_max=list(srs[diamax_index])[0][:list(srs[diamax_index])[0].find(',')] #找到","於string中的index number

        diamin_index=srs.index[srs=='"Min, Diff, Dir"']+1 #找到"Dia","DiaB"於series中的index number，將index number+1即為資料欄位
        dia_min=list(srs[diamin_index])[0][:list(srs[diamin_index])[0].find(',')] #找到","於string中的index number
        
        df_roundness=pd.DataFrame({'Roundness':[str((float(dia_max)-float(dia_min))*1000)]})#將diameter完整資訊存成dataframe
        
        
        #以下開始處理Notch資訊
        notch_index=srs.index[srs=='"[Notch]"'].tolist()[0] #找到"[Notch]"於series中的index number
        df_notch=srs[notch_index+1:notch_index+3] #取出notch相關的資訊，目前還是一個series
        df_notch=df_notch.str.split(',',expand=True) #辨識"," 將資料拆解並存成dataframe
        df_notch.reset_index(drop=True, inplace=True) 
        df_notch.rename(columns=df_notch.iloc[0],inplace=True) #將最上面一個row設為column name
        df_notch.drop(0,inplace=True) #完成後將最上面一row刪除
        #因column name中帶有""，不美觀，以下將column中的"符號刪除
        df_notch_col=list(df_notch.columns)
        for i in range(len(df_notch.columns)):
            df_notch_col[i]=df_notch_col[i].replace('"','')
        df_notch.columns=df_notch_col
        df_notch.reset_index(drop=True, inplace=True)#在一次reset index, 以便之後要合併dataframe時各dataframe index number一致

        #對edge量測參數和數值做一樣的動作(同上)
        edge_index=srs.index[srs=='"[Edge]"'].tolist()[0]
        df_edge=srs[edge_index+1:edge_index+13]
        df_edge=df_edge.str.split(',',expand=True)
        df_edge.reset_index(drop=True, inplace=True)
        df_edge.rename(columns=df_edge.iloc[0],inplace=True) #將最上面一個row設為column name
        df_edge.drop(0,inplace=True)
        df_edge_col=list(df_edge.columns)
        for i in range(len(df_edge.columns)):
            df_edge_col[i]=df_edge_col[i].replace('"','')
        df_edge.columns=df_edge_col
        df_edge=df_edge[df_edge["Point"]=='"<Ave>"']
        df_edge=df_edge.drop(columns=['No', 'Point'])
        df_edge.reset_index(drop=True, inplace=True)

        df_temp=pd.concat([df_waferid,df_recipe,df_roundness,df_diameter,df_edge,df_notch],axis=1)
        return True,df_temp
    except:
        return False,''

st.set_page_config(page_title="LEP CSV Parser", layout="centered")
st.title("LEP CSV Parser (Streamlit)")
st.caption("Select one or more *.csv files, then click **Parse** to generate lep.csv")

uploaded_files = st.file_uploader(
    "Choose CSV files", type=["csv"], accept_multiple_files=True
)

parse_clicked = st.button("Parse")

import tempfile
from typing import List
import io

def process_files(files: List[io.BytesIO]) -> pd.DataFrame:
    results = []
    for idx, uf in enumerate(files, start=1):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp:
            tmp.write(uf.read())
            tmp_path = tmp.name

        st.write(f"{idx}/{len(files)} :: {os.path.basename(uf.name)}")
        TorF, df_temp = parseoneLEPfile(tmp_path)

        # Clean up the temp file ASAP
        try:
            os.remove(tmp_path)
        except Exception:
            pass

        if TorF and isinstance(df_temp, pd.DataFrame):
            results.append(df_temp)

    if not results:
        return pd.DataFrame()

    df_summary = pd.concat(results, ignore_index=True)

    # Convert each column to numeric where possible; keep text as-is otherwise
    for col in df_summary.columns:
        df_summary[col] = pd.to_numeric(df_summary[col], errors="ignore")

    return df_summary


if parse_clicked:
    if not uploaded_files:
        st.warning("Please upload at least one CSV file.")
    else:
        with st.spinner("Parsing..."):
            df_summary = process_files(uploaded_files)

        if df_summary.empty:
            st.error("Parsing failed. Please check your files.")
        else:
            st.success("Done! Preview and download below.")
            st.dataframe(df_summary, use_container_width=True)

            # Offer a download of lep.csv (same filename as your PyQt flow)
            csv_bytes = df_summary.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="Download lep.csv",
                data=csv_bytes,
                file_name="lep.csv",
                mime="text/csv",
            )


    
