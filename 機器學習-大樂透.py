import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error
import numpy as np

# 讀入歷史開獎資料
df = pd.read_excel("大樂透開獎號碼-新版.xlsx")

X = pd.DataFrame(df.drop(columns=['開獎日']))
y = pd.DataFrame(df,columns=['頭獎數量','獎號1','獎號2','獎號3','獎號4','獎號5','獎號6','特別號'])
number=np.array([112000024]).reshape(1, -1)

# 將資料分為訓練和測試集
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, shuffle=False)

# 建立隨機森林回歸模型
clf = RandomForestClassifier(n_estimators=100)
#clf = RandomForestRegressor(n_estimators=100, random_state=0)
# 訓練分類器
clf.fit(X_train, y_train)

# 使用分類器預測下一次開獎號碼
prediction = clf.predict(X_test)
#print("Predicted number: ", prediction)

# 計算準確率
score = mean_squared_error(y_test, prediction)
print("預測的準確率：", score)

# 將資料轉成DataFrame
result = pd.DataFrame(prediction)
df1=pd.DataFrame(X_test)
df2=pd.DataFrame(y_test)


# 建立一個ExcelWriter物件
writer = pd.ExcelWriter('大樂透預測號碼.xlsx', engine='openpyxl', mode='w')

# 將DataFrame寫入Excel
result.to_excel(writer, sheet_name='prediction', index=False)
df1.to_excel(writer, sheet_name='X_test', index=False)
df2.to_excel(writer, sheet_name='y_test', index=False)

writer.save()
