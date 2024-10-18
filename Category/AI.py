import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split

global vectorizer, model

# โหลดและเตรียมโมเดล
data = pd.read_csv("testtrain.csv")

# กำหนดฟีเจอร์และตัวแปรเป้าหมาย
X = data['ชื่อวัตถุดิบ']
y = data['ประเภทวัตถุดิบ']

# แบ่งข้อมูล
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# แปลงข้อมูลข้อความ
vectorizer = TfidfVectorizer()
X_train_vec = vectorizer.fit_transform(X_train)
X_test_vec = vectorizer.transform(X_test)

# สร้างและฝึกโมเดล
model = RandomForestClassifier(n_estimators=100, random_state=42)
model.fit(X_train_vec, y_train)

