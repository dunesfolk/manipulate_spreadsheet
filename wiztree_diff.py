import sys
import pandas as pd

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python script.py <file1> <file2> <output>")
        exit

# ヘッダー「ファイル名」にファイルのフルパス、ヘッダー「更新日時」にファイルの更新日時があるカンマ区切りファイルを読み込む。
# argv1にあるファイルで、argv2では更新日時が一致しないか存在しないファイルのリストを、argv3に吐き出す。
file1_path = sys.argv[1]
file2_path = sys.argv[2]
output_path = sys.argv[3]

# WinDirStat/Wiztree互換性維持のため
# CSVファイルの１行目がコメントでヘッダーでは無い場合は１行目をスキップする 
def determine_skip_rows(file_path):
    #１行目にカンマがあれば0を返す、なければ1を返す
    with open(file_path, 'r', encoding='utf-8') as file:
        first_line = file.readline()
        return 1 if first_line.count(',') == 0 else 0
skiprow1 = determine_skip_rows(sys.argv[1])
skiprow2 = determine_skip_rows(sys.argv[2])

# CSVをpandas.DataFrameクラスオブジェクトとして読み込む
df1 = pd.read_csv(file1_path, encoding="utf-8_sig", skiprows=skiprow1)
df2 = pd.read_csv(file2_path, encoding="utf-8_sig", skiprows=skiprow2)

# WinDirStat/Wiztree互換性維持のためファイル名と更新日時のカラム名を取得する
col = {
    "windirstat": {
        "name": "Name",
        "change": "Last Change"
    },
    "wiztree": {
        "name": "ファイル名",
        "change": "更新日時"
    }
}
def determine_df_type(df, df_name):
    if "Name" in df.columns:
        return "windirstat"
    elif "ファイル名" in df.columns:
        return "wiztree"
    else:
        raise ValueError(f"Neither 'Name' nor 'ファイル名' column exists in {df_name}.")

def refactor_date(df):
    dftype = determine_df_type(df, "dataframe")
    # スキップするファイル名のパターンを正規表現で指定する
    pattern2 = (lambda fileext, filename: rf"(?i){fileext}|{filename}")(
        r"\.(apdisk|fseventsd|Spotlight-V100|TemporaryItems|Trashes|lnk|tmp|temp)$",
        r"\\(_.*|\._.*|~\$.*|desktop\.ini|thumbs\.db|\.DS_Store|\$RECYCLE\.BIN)$"
    )
    # データフレームから正規表現に部分一致する行があれば削除する
    df = df[~df[col[dftype]["name"]].astype(str).str.contains(pattern2, regex=True, na=False)]

    # SecureSambaとFileForce固有のディレクトリのプレフィックスを文字列置き換えで削除する
    pattern = r"^(V:\\パブリックフォルダ|Z:\\全社共有)" #正規表現
    replacement = "" #置き換え先（空欄）
    df.loc[:, col[dftype]["name"]] = df[col[dftype]["name"]].astype(str).replace(pattern, replacement, regex=True)

    #WinDirStatの場合、Zulu TimeからGMT+9に変更する
    if dftype == "windirstat":
        # Ensure column is converted to datetime properly
        df.loc[:, col[dftype]["change"]] = pd.to_datetime(df[col[dftype]["change"]], errors='coerce', format="%Y-%m-%dT%H:%M:%SZ")

        # Debug: Check for rows that failed conversion (NaT values)
        print(df[df[col[dftype]["change"]].isna()])  # Prints rows that couldn't be converted
        # Apply timezone conversion **only if dtype is datetime64**
        if pd.api.types.is_datetime64_any_dtype(df[col[dftype]["change"]]):
            df.loc[:, col[dftype]["change"]] = df[col[dftype]["change"]].dt.tz_localize('UTC').dt.tz_convert('Asia/Tokyo')
            df.loc[:, col[dftype]["change"]] = df[col[dftype]["change"]].dt.strftime("%Y/%m/%d %H:%M:%S")
    return df, dftype

df1, df1type = refactor_date(df1)
df2, df2type = refactor_date(df2)

# df2でファイル名から更新日時検索できる辞書（連想配列）を作成する
dict_df2_toLastModified = dict(zip(df2[col[df2type]["name"]], df2[col[df2type]["change"]]))

# df1にヘッダ「df2_更新日時」の列を追加する。
# df1と同じファイル名の行をdf2で探して、df2の更新日時を引っ張ってきて、「df2_更新日時」に書き入れる。
# df1["df2_更新日時"] = df1[col[df1type]["name"]].map(dict_df2_toLastModified)
df1.loc[:, "df2_更新日時"] = df1[col[df1type]["name"]].map(dict_df2_toLastModified)

# df1とdf2とで更新日時が同じであれば、行を削除する
filtered_df1 = df1[df1[col[df1type]["change"]] != df1["df2_更新日時"]]

# df1とdf2で更新日時が一致しない行と、df1には在るがdf2には無いファイルの行が残るので、
# CSVファイルとして書き出す。
filtered_df1.to_csv(output_path, index=False, encoding="utf-8_sig")

