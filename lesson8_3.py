import pandas as pd

def main():
    """
    主函數，用於讀取CSV檔案，進行資料清洗和整理，並將結果儲存為CSV和Excel檔案。
    """
    # 讀取名為'上市公司資料.csv'的CSV檔案，並將其儲存為DataFrame物件df。
    df = pd.read_csv('上市公司資料.csv')
    # 移除DataFrame df中包含任何NaN值的列，並將結果儲存為新的DataFrame物件df1。
    df1 = df.dropna()
    # 重新索引DataFrame df1，只保留指定的欄位，並將結果儲存為新的DataFrame物件df2。
    # 保留的欄位包括：'公司代號'、'出表日期'、'公司名稱'、'產業別'、'營業收入-當月營收'、'營業收入-上月營收'。
    df2 = df1.reindex(columns=['公司代號','出表日期','公司名稱','產業別','營業收入-當月營收','營業收入-上月營收'])
    # 重新命名DataFrame df2中的欄位名稱，並將結果儲存為新的DataFrame物件df3。
    # '營業收入-上月營收'被重新命名為'上月營收'，'營業收入-當月營收'被重新命名為'當月營收'。
    df3 = df2.rename(columns={
        '營業收入-上月營收':'上月營收',
        '營業收入-當月營收':'當月營收'
        })
    # 將DataFrame df3儲存為名為'上市公司資料_整理.csv'的CSV檔案，並使用utf-8編碼。
    df3.to_csv('上市公司資料_整理.csv',encoding='utf-8')
    # 將DataFrame df3儲存為名為'上市公司資料_整理.xlsx'的Excel檔案。
    df3.to_excel('上市公司資料_整理.xlsx')
    # 印出'存檔完成'，表示檔案儲存已完成。
    print('存檔完成')

 # 當此腳本被直接執行時，執行main()函數。
if __name__ == '__main__':
    main()
