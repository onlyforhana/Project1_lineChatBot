import pandas as pd
from neo4j import GraphDatabase

# ตั้งค่า URI และ AUTH สำหรับ Neo4j
URI = "neo4j://localhost:7687"
AUTH = ("neo4j", "theoneandonlyhana")

# ฟังก์ชันสำหรับเชื่อมต่อและเก็บคำถามและคำตอบใน Neo4j
# ลองเพิ่ม logging เพื่อดูว่ามีข้อมูลใดที่ไม่ถูกนำเข้า
def save_to_neo4j(data):
    total_rows = len(data)
    inserted_rows = 0
    try:
        driver = GraphDatabase.driver(URI, auth=AUTH)
        with driver.session() as session:
            for index, row in data.iterrows():
                question = row['Question']
                answer = row['Answer']
                
                result = session.run(
                    '''
                    MERGE (faq:FAQ {question: $question, answer: $answer})
                    ''', 
                    question=question, answer=answer
                )
                inserted_rows += 1
        print(f"Data successfully saved to Neo4j! {inserted_rows}/{total_rows} rows inserted.")
    except Exception as e:
        print(f"Error saving data to Neo4j: {e}")
    finally:
        driver.close()

# อ่านข้อมูลจากไฟล์ Excel
def load_excel(file_path):
    try:
        # อ่านไฟล์ Excel
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

# เส้นทางไฟล์ Excel
excel_file = "C:\\Users\\firhana\\socialai\\proj1\\faxformychatbot.xlsx"

# อ่านข้อมูลจาก Excel และนำไปบันทึกใน Neo4j
data = load_excel(excel_file)
if data is not None:
    save_to_neo4j(data)
print(data)