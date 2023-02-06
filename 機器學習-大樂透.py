import pandas as pd
import openpyxl
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split

# 讀入歷史開獎資料
df = pd.read_excel("大樂透歷史開獎號碼.xlsx")


# 把開獎號碼轉為數字
#df["number"] = df["number"].astype(int)
number=[df["號碼1"],df["號碼2"]]
print(number)
input()

# 定義特徵和標籤
X = df.drop("號碼1","號碼2", axis=1)
y = df["number"]
print(x)
input()
# 將資料分為訓練和測試集
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, random_state=0)

# 建立隨機森林分類器
clf = RandomForestClassifier(n_estimators=100, random_state=0)

# 訓練分類器
clf.fit(X_train, y_train)

# 使用測試集評估模型效果
accuracy = clf.score(X_test, y_test)
print("Accuracy: ", accuracy)

# 使用分類器預測下一次開獎號碼
prediction = clf.predict([[2023, 2, 5]])
print("Predicted number: ", prediction)