import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
from datetime import datetime
import csv
from typing import Dict, List, Tuple

class ScoreProcessor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()

    def select_file(self, title: str, file_types: list) -> str:
        """使用彈出視窗選擇檔案"""
        file_path = filedialog.askopenfilename(
            title=title,
            filetypes=file_types
        )
        
        if not file_path:
            messagebox.showwarning("警告", "沒有選擇檔案！")
            return None
            
        return file_path

    def format_time(self, time_str):
        """將時間格式轉換為24小時制，並移除AM/PM"""
        if pd.isna(time_str) or time_str == '':
            return ''
        try:
            dt = pd.to_datetime(time_str)
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        except:
            return time_str

    def parse_chinese_time(self, time_str):
        """解析時間字串為datetime物件"""
        try:
            if pd.isna(time_str):
                return None
            format_str = '%Y-%m-%d %H:%M:%S'
            return datetime.strptime(time_str, format_str)
        except Exception as e:
            print(f"無法解析時間 '{time_str}': {e}")
            return None

    def process_excel_to_csv(self, file_path):
        """處理Excel檔案並轉換為CSV"""
        try:
            df = pd.read_excel(file_path)
            
            # 尋找包含 "題" 字的列名
            answer_columns = [col for col in df.columns if '題' in str(col)]
            
            # 修改欄位名稱
            new_column_names = {col: f'q{i+1}' for i, col in enumerate(answer_columns)}
            df = df.rename(columns=new_column_names)

            # 刪除包含電子郵件地址的欄
            email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
            columns_to_drop = [col for col in df.columns if df[col].astype(str).str.contains(email_pattern).any()]
            if columns_to_drop:
                df = df.drop(columns=columns_to_drop)

            # 刪除所有 'Unnamed' 開頭的欄位
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

            # 處理時間格式
            if '時間戳記' in df.columns:
                df.loc[0, '時間戳記'] = ''
                for idx in df.index[1:]:
                    if pd.notna(df.loc[idx, '時間戳記']):
                        df.loc[idx, '時間戳記'] = self.format_time(df.loc[idx, '時間戳記'])

            # 選擇儲存位置
            save_path = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[('Excel檔案', '*.xlsx')],
                title='儲存處理後的Excel檔案'
            )
            
            if save_path:
                # 儲存 Excel 檔案
                df.to_excel(save_path, index=False)
                
                # 儲存 CSV 檔案
                csv_path = os.path.splitext(save_path)[0] + '.csv'
                df.to_csv(csv_path, index=False, encoding='utf-8-sig', date_format='%Y-%m-%d %H:%M:%S')
                
                return csv_path
            else:
                messagebox.showwarning("警告", "未選擇儲存位置！")
                return None
                
        except Exception as e:
            messagebox.showerror("錯誤", f"處理檔案時發生錯誤：\n{str(e)}")
            return None

    def get_question_count(self, df):
        """自動計算題目數量"""
        q_columns = [col for col in df.columns if str(col).lower().startswith('q')]
        return len(q_columns)

    def find_answer_row(self, df):
        """找出標準答案所在的行"""
        for idx, row in df.iterrows():
            if row['學號'] == 'email' and pd.isna(row['時間戳記']):
                return idx
        return None

    def improved_score_answer(self, student_answer, correct_answer):
        """答案必須完全相同"""
        if pd.isna(student_answer):
            return 0
        if student_answer == 'non' or student_answer == '':
            return 0
        student_answer = str(student_answer).strip()
        correct_answer = str(correct_answer).strip()
        return 1 if student_answer == correct_answer else 0

    def process_score_calculation(self, csv_path, time_unit):
        """處理成績計算"""
        try:
            df = pd.read_csv(csv_path, encoding='utf-8-sig')
            num_questions = self.get_question_count(df)
            
            if num_questions == 0:
                messagebox.showerror("錯誤", "找不到題目欄位 (q1, q2, ...)")
                return None

            # 找出標準答案所在的行
            answer_row_idx = self.find_answer_row(df)
            if answer_row_idx is None:
                messagebox.showerror("錯誤", "找不到標準答案行")
                return None

            # 找到最早的提交時間
            first_submission_time = None
            min_time = None
            for idx, row in df.iterrows():
                if idx != answer_row_idx and pd.notna(row['時間戳記']) and pd.notna(row['學號']):
                    current_time = self.parse_chinese_time(row['時間戳記'])
                    if current_time is not None:
                        if min_time is None or current_time < min_time:
                            min_time = current_time
                            first_submission_time = current_time

            if first_submission_time is None:
                messagebox.showerror("錯誤", "找不到有效的提交時間")
                return None

            # 生成正確答案欄位
            correct_answers = df.iloc[answer_row_idx, 3:3 + num_questions].values
            answer_columns = [f'q{i+1}' for i in range(num_questions)]

            # 計算每個題目的得分
            correct_answer_counts = []
            exclude_idx = answer_row_idx
            for i, col in enumerate(answer_columns):
                scores = pd.Series(index=df.index, dtype='float64')
                valid_rows = df.index[df.index != exclude_idx]
                scores[valid_rows] = df.loc[valid_rows, col].apply(
                    lambda x: self.improved_score_answer(x, correct_answers[i]))
                df[f'{col}得分'] = scores
                correct_answer_counts.append(df[f'{col}得分'])

            # 計算答對題數和考試分數
            df['答對題數'] = sum(correct_answer_counts).fillna(0)
            df['考試分數'] = 0

            for index, row in df.iterrows():
                if index != answer_row_idx and pd.notna(row['學號']):
                    submission_time = self.parse_chinese_time(row['時間戳記'])
                    if submission_time and pd.notna(row['答對題數']):
                        time_difference = (submission_time - first_submission_time).total_seconds() / 60
                        time_units = int(np.ceil(time_difference / time_unit))
                        weight = 1.01 - (time_units * 0.01)
                        score = row['答對題數'] * weight
                        df.at[index, '考試分數'] = score

            # 清除沒有學生的行和標準答案行
            df = df[(df['學號'].notna()) & (df.index != answer_row_idx)]
            final_columns = ['時間戳記', '學號', '答對題數', '考試分數']
            final_df = df[final_columns].copy()

            # 儲存結果
            result_path = csv_path.rsplit('.', 1)[0] + '_results.csv'
            final_df.to_csv(result_path, index=False, encoding='utf-8-sig', na_rep='0', float_format='%.2f')
            
            return result_path
            
        except Exception as e:
            messagebox.showerror("錯誤", f"計算成績時發生錯誤：\n{str(e)}")
            return None

    def read_final_scores(self, file_path: str) -> Dict[str, Tuple[float, float]]:
        """讀取成績CSV檔案"""
        scores = {}
        with open(file_path, 'r', encoding='utf-8-sig') as file:
            reader = csv.DictReader(file)
            for row in reader:
                student_id = row['學號'].lower()
                try:
                    answers_correct = float(row['答對題數'])
                    final_score = float(row['考試分數'])
                    scores[student_id] = (answers_correct, final_score)
                except (ValueError, KeyError) as e:
                    print(f"處理學號 {student_id} 的資料時發生錯誤: {e}")
                    continue
        return scores

    def update_calculus_scores(self, calculus_file: str, final_scores: Dict[str, Tuple[float, float]], date: str):
        """更新微積分成績檔案"""
        with open(calculus_file, 'r', encoding='utf-8-sig') as file:
            reader = csv.reader(file)
            rows = list(reader)

        # 找到資料結束的位置
        data_end = next((i for i, row in enumerate(rows) if row[0].startswith('本校')), len(rows))

        # 在標題列中找到對應日期的欄位
        header = rows[0]
        try:
            date_column = header.index(date)
            answers_column = date_column + 1
        except ValueError:
            messagebox.showerror("錯誤", f"在CSV文件中找不到 '{date}' 列。請確保日期輸入正確。")
            return False

        # 更新成績
        updated_count = 0
        for i in range(1, data_end):
            student_id = rows[i][2].lower()
            if student_id in final_scores:
                rows[i][date_column] = str(final_scores[student_id][1])
                rows[i][answers_column] = str(int(final_scores[student_id][0]))
                updated_count += 1

        try:
            with open(calculus_file, 'w', encoding='utf-8-sig', newline='') as file:
                writer = csv.writer(file)
                writer.writerows(rows)
            return updated_count
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存檔案時發生錯誤：{str(e)}")
            return False

    def process_all(self):
        """執行完整的處理流程"""
        # Step 1: 選擇並處理Excel檔案
        excel_file = self.select_file("選擇Excel檔案", [('Excel檔案', '*.xlsx'), ('所有檔案', '*.*')])
        if not excel_file:
            return
        
        csv_path = self.process_excel_to_csv(excel_file)
        if not csv_path:
            return
            
        # Step 2: 計算成績
        time_unit = simpledialog.askinteger("輸入", "請輸入時間單位（分鐘）:", parent=self.root)
        if time_unit is None:
            messagebox.showwarning("警告", "未輸入時間單位！")
            return
            
        results_path = self.process_score_calculation(csv_path, time_unit)
        if not results_path:
            return
            
        # Step 3: 更新微積分成績檔案
        calculus_file = self.select_file("選擇微積分成績檔案", [('CSV檔案', '*.csv'), ('所有檔案', '*.*')])
        if not calculus_file:
            return
            
        date = simpledialog.askstring("輸入", "請輸入日期 (例如: 10月8號):", parent=self.root)
        if not date:
            messagebox.showwarning("警告", "未輸入日期！")
            return

        try:
            final_scores = self.read_final_scores(results_path)
            if not final_scores:
                messagebox.showerror("錯誤", "無法讀取成績檔案或成績檔案為空")
                return

            updated_count = self.update_calculus_scores(calculus_file, final_scores, date)
            if updated_count:
                messagebox.showinfo("成功", f"已更新 {updated_count} 位學生的成績")
            elif updated_count == 0:
                messagebox.showwarning("警告", "沒有找到需要更新的成績")
                
        except Exception as e:
            messagebox.showerror("錯誤", f"更新成績時發生錯誤：{str(e)}")

def main():
    processor = ScoreProcessor()
    processor.process_all()

if __name__ == "__main__":
    main()
