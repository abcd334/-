import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split


# 讀入歷史開獎資料
df = pd.read_excel("大樂透歷史開獎號碼.xlsx")

'''
# 把開獎號碼轉為數字
#df["number"] = df["number"].astype(int)
number=[df["號碼1"],df["號碼2"],df["號碼3"],df["號碼4"],df["號碼5"],df["號碼6"],df["特別號"]]

# 定義特徵和標籤
X=pd.DataFrame(df)
print(df)
input()
'''
X = pd.DataFrame(df,columns=['號碼1','號碼2','號碼3','號碼4','號碼5','號碼6','特別號'])
#y = df["number"]
#y=pd.DataFrame(df,columns=['號碼1','號碼2','號碼3','號碼4','號碼5','號碼6','特別號'])

# 將資料分為訓練和測試集
X_train, y_test = train_test_split(X, test_size=0.1, random_state=0)

# 建立隨機森林分類器
clf = RandomForestClassifier(n_estimators=100, random_state=0)

# 訓練分類器
clf.fit(X_train, y_test)

# 使用測試集評估模型效果
accuracy = clf.score(X_train, y_test)
print("Accuracy: ", accuracy)

# 使用分類器預測下一次開獎號碼
prediction = clf.predict(X_train)
print("Predicted number: ", prediction)