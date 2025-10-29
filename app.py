# 本程式會撈取同資料夾中的csv檔案，抓取lep資料並匯出為excel檔案
# V2, 新增recipe欄位，並模組化，其中parseoneLEPfile可以供日後呼叫方便

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
import tempfile
from typing import List
import io


# name_of_wb='LEP_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.csv'
# dir_path = os.path.dirname(os.path.realpath(__file__))
# ↑獲取當前資料夾名稱然後存成dir_path變數
# all_file_name = os.listdir(dir_path)
# ↑讀取資料夾內所有檔案名稱然後放進all_file_name這個list裡

# path為想要爬取檔案的所在資料夾完整路徑(D:\coding\LEP_offline_parser)，filename為檔案名稱(ex:POLYPR205.Csv)
def parseoneLEPfile(filename) -> tuple[bool, pd.DataFrame | str]:  # (path,filename):
    # url=path+'/'+filename #url=完整檔案名稱=檔案路徑+檔案名稱
    # waferid=filename[:filename.index(".")] #抓取檔案名稱，因LEP檔案名稱為LOT+slot，所以直接抓檔案名稱即可作為wafer身分辨別
    # df=pd.read_fwf(url) #讀取檔案，存在名為df的變數中
    # print(filename)
    try:
        waferid = os.path.basename(filename)
        df = pd.read_fwf(filename)  # 讀取檔案，存在名為df的變數中
        srs = df[list(df.columns)[0]]  # 取出df的資料存為series; dataframe of first column
        df_waferid = pd.DataFrame({'LOT_slot': [waferid]})

        # Recipe
        recipe_matches = [i for i in srs if '"Recipe Name:",' in i] # 找到recipe name; scan all lines and find those containing "Recipe Name:"
        if not recipe_matches:
            recipe = ""
        else:
            recipe_raw = recipe_matches[0]
            recipe = recipe_raw[recipe_raw.index(',') + 1:].strip() # 找到,所在的index; find the first comma, +1 move after the comma, and .strip() remove the leading and trailing whitespace
            recipe = recipe.strip('"') # 去除"的字樣; remove surrounding quotes
        df_recipe = pd.DataFrame({'Recipe': [recipe]}) # 將recipe資訊存成dataframe

        # Diameter
        df_diameter = pd.DataFrame({'Diameter': [None]})
        try:
            dia_index = srs.index[srs == '"Dia","DiaB"'][0] + 1 # 找到"Dia","DiaB"於series中的index number，將index number+1即為資料欄位
            diameter_line = srs.iloc[dia_index]
            diameter = diameter_line.split(',', 1)[0] # 找到","於string中的index number
            df_diameter = pd.DataFrame({'Diameter': [diameter]}) # 將diameter完整資訊存成dataframe
        except Exception:
            pass

        # Roundness
        df_roundness = pd.DataFrame({'Roundness': [None]})
        try:
            max_idx = srs.index[srs == '"Max, Diff, Dir"'][0] + 1 # 找到"Dia","DiaB"於series中的index number，將index number+1即為資料欄位
            min_idx = srs.index[srs == '"Min, Diff, Dir"'][0] + 1 # 找到"Dia","DiaB"於series中的index number，將index number+1即為資料欄位
            dia_max_line = srs.iloc[max_idx]
            dia_min_line = srs.iloc[min_idx]
            dia_max = float(dia_max_line.split(',', 1)[0]) # 找到","於string中的index number
            dia_min = float(dia_min_line.split(',', 1)[0]) # 找到","於string中的index number
            roundness = (dia_max - dia_min) * 1000 # calculate roundness
            df_roundness = pd.DataFrame({'Roundness': [f"{roundness:g}"]})
        except Exception:
            pass

        # Notch
        df_notch = pd.DataFrame({'Notch': [None]})
        try:
            notch_index = srs.index[srs == '"[Notch]"'][0] # 找到"[Notch]"於series中的index number
            df_notch_block = srs[notch_index + 1:notch_index + 3].str.split(',', expand=True) # 取出notch相關的資訊，目前還是一個series
            df_notch_block.reset_index(drop=True, inplace=True)
            df_notch_block.rename(columns=df_notch_block.iloc[0], inplace=True) # 將最上面一個row設為column name
            df_notch_block = df_notch_block.drop(0).reset_index(drop=True) # 完成後將最上面一row刪除
            df_notch_block.columns = [c.replace('"', '') for c in df_notch_block.columns] # 因column name中帶有""，不美觀，以下將column中的"符號刪除
            df_notch = df_notch_block
        except Exception:
            pass

        # Bevel
        df_bevel = pd.DataFrame({'Bevel': [None]})
        try:
            bevel_index = srs.index[srs == '"[Bevel]"'][0] # find "[Bevel]" index.
            df_bevel_block = srs[bevel_index + 1:bevel_index + 3].str.split(',', expand=True) # take two lines after the header and split by commas
            df_bevel_block.reset_index(drop=True, inplace=True)
            df_bevel_block.rename(columns=df_bevel_block.iloc[0], inplace=True) # use first row as column headers
            df_bevel_block = df_bevel_block.drop(0).reset_index(drop=True)
            df_bevel_block.columns = [c.replace('"', '') for c in df_bevel_block.columns] # remove quotes from column names
            df_bevel = df_bevel_block
        except Exception:
            pass

        # Edge
        df_edge = pd.DataFrame({'Edge': [None]})
        try:
            edge_index = srs.index[srs == '"[Edge]"'][0]
            df_edge_block = srs[edge_index + 1:edge_index + 13].str.split(',', expand=True)
            df_edge_block.reset_index(drop=True, inplace=True)
            df_edge_block.rename(columns=df_edge_block.iloc[0], inplace=True)
            df_edge_block = df_edge_block.drop(0).reset_index(drop=True)
            df_edge_block.columns = [c.replace('"', '') for c in df_edge_block.columns]

            if "Point" in df_edge_block.columns: # keep only the row where "Point" == "<Ave>"
                df_edge_block = df_edge_block[df_edge_block["Point"] == '"<Ave>"']
            for c in ['No', 'Point']:
                if c in df_edge_block.columns:
                    df_edge_block = df_edge_block.drop(columns=c)
            df_edge_block = df_edge_block.reset_index(drop=True)

            df_edge = df_edge_block if not df_edge_block.empty else pd.DataFrame()
        except Exception:
            pass

        parts = [df_waferid, df_recipe] # always include waferid and recipe
        for section in [df_roundness, df_diameter, df_edge, df_bevel, df_notch]: # for each section, append only if it has content
            if isinstance(section, pd.DataFrame) and not section.empty:
                parts.append(section)
        df_temp = pd.concat(parts, axis=1) # column-wise concatenation
        return True, df_temp
    except:
        return False, ''


st.set_page_config(page_title="LEP Parser", layout="wide")
st.title("LEP Parser")

uploaded_files = st.file_uploader("Choose CSV files (.csv)", type=["csv"], accept_multiple_files=True)

parse_clicked = st.button("Parse")


def process_files(files: List[io.BytesIO]) -> pd.DataFrame:
    """
    Inputs:
        List of uploaded file-like objects
    Output:
        DataFrame with parsed results
    """
    results = []
    for idx, uf in enumerate(files, start=1):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp:  # delete=False to prevent immediate deletion
            tmp.write(uf.read())  # save uploaded file content to temp file
            tmp_path = tmp.name

        st.write(f"{os.path.basename(uf.name)}")  # display file name
        TorF, df_temp = parseoneLEPfile(tmp_path)

        os.remove(tmp_path)

        # if TorF and isinstance(df_temp, pd.DataFrame):  # proceed if parsing was successful (TorF=True) and df_temp is a DataFrame
        #     original_name = os.path.basename(uf.name)
        #     df_temp["LOT_slot"] = original_name  # new column with original file name
        #     results.append(df_temp)

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
            st.error("Parsing failed. Please check the files.")
        else:
            st.success("Done parsing. Pleaes preview and download below.")
            st.dataframe(df_summary, use_container_width=True)

            csv_bytes = df_summary.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="Download",
                data=csv_bytes,
                file_name="lep.csv",
                mime="text/csv",
            )

            st.caption("The file will be downloaded in the download folder.")
