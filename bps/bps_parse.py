from pandas.core.frame import DataFrame
import pandas as pd
import glob

def bps_parse(df) -> DataFrame:
    # Penentuan variabel judul, golongan, jenis dan satuan data
    fields = df.columns[0]
    title = '_'.join(df.columns[1].split('(')[0].split('_')[:-1]).replace('/', '_')
    group = '_'.join(title.split('_')[1:])
    unit = df.columns[1].split('(')[-1][:-1]
    
    # daftar nama kolom awal
    old_col = df.columns[1:]

    # mengeluarkan keterangan data
    df = df.iloc[:-4]
    ls_field = df[fields].dropna().values
    
    # penentuan baris data yang memuat nilai NaN
    nan_rows = df.loc[df.isnull().any(axis=1) == True]
    year = nan_rows.iloc[-1].dropna().astype('int').unique().tolist()
    
    if nan_rows.shape[0] == 1:
        sub_class = [group]
    else: 
        sub_class = nan_rows.dropna(axis=1).iloc[0].to_list()
    
    # mengganti nilai - dengan 0
    df = df.dropna(axis=0)[old_col].replace('-', '0')
    name_of_file = f"./data_target/{title}_{year[0]}-{year[-1]}.csv"
    
    ls_data = []
    for n,_ in enumerate(old_col):
        if (n != len(old_col) - 1) & (n % len(year) == 0):
            m = int(n/len(year))
            df1 = df[old_col[n:n+len(year)]].astype('float')
            df1.columns = year

            df1[fields] = ls_field
            df1['kelompok'] = group
            df1['jenis'] = group if len(sub_class) == 1 else sub_class[m-1]
            df1['satuan'] = unit

            ls_data.append(df1)
    return pd.concat(ls_data).set_index([fields,'kelompok', 'jenis', 'satuan']).to_csv(name_of_file)
    
# data = pd.read_csv('./data_target/Produksi_Palawija_2017-2019.csv')

for f in glob.glob('./data_csv/*csv'):
    data = pd.read_csv(f)
    data.pipe(bps_parse)