import pandas as pd

# 엑셀 파일 읽기
df_1 = pd.read_excel('locklock_test.xlsx', engine='openpyxl', sheet_name='Sheet1')
df_2 = pd.read_excel('locklock_test.xlsx', engine='openpyxl', sheet_name='Sheet2')

# 날짜 형식을 변환
df_1['수리일자'] = pd.to_datetime(df_1['수리일자'], format='%Y%m%d').dt.date
df_2['송금일자'] = pd.to_datetime(df_2['송금일자'], format='%Y%m%d').dt.date

# 새로운 결과를 저장할 리스트와 외환송금번호 및 번호를 추적할 집합
results = []
used_forex_numbers = set()
used_numbers = set()

# '무역거래처상호'가 같은 행을 찾고 6개월 이내에 거래된 경우와 입력결제금액과 송금외화금액이 3% 이내로 같은 경우를 찾음
for idx_1, row_1 in df_1.iterrows():
    for idx_2, row_2 in df_2.iterrows():
        if row_1['무역거래처상호'] == row_2['무역거래처상호']:
            date_diff = abs((row_1['수리일자'] - row_2['송금일자']).days)
            if date_diff <= 180:
                money_diff = abs(row_1['입력결제금액'] - row_2['송금외화금액'])
                money_1_threshold = row_1['입력결제금액'] * 0.03
                if money_diff <= money_1_threshold:
                    if row_2['외환송금번호'] not in used_forex_numbers and row_1.name not in used_numbers:
                        results.append({
                            '번호': row_1.name,
                            '무역거래처상호': row_1['무역거래처상호'],
                            '수리일자': row_1['수리일자'],
                            '입력결제금액': row_1['입력결제금액'],                        
                            '송금일자': row_2['송금일자'],                            
                            '송금외화금액': row_2['송금외화금액'],
                            '해외공급자상호': row_1['해외공급자상호'],
                            '수취인계좌번호': row_2['수취인계좌번호'],
                            'B/L번호': row_1['B/L번호'],
                            '외환송금번호': row_2['외환송금번호'],
                            '외환사유코드': row_2['외환사유코드']
                        })
                        used_forex_numbers.add(row_2['외환송금번호'])
                        used_numbers.add(row_1.name)

# 중복되지 않은 행을 찾음
df_1_unique = df_1[~df_1.index.isin(used_numbers)]
df_2_unique = df_2[~df_2['외환송금번호'].isin(used_forex_numbers)]

# 리스트를 데이터프레임으로 변환
result_df = pd.DataFrame(results)

# 결과를 엑셀 파일로 저장
with pd.ExcelWriter('locklock_test_output.xlsx', engine='openpyxl') as writer:
    result_df.to_excel(writer, sheet_name='Matched', index=False)
    df_1_unique.to_excel(writer, sheet_name='Sheet1_Unmatched', index=False)
    df_2_unique.to_excel(writer, sheet_name='Sheet2_Unmatched', index=False)

# Sheet1_Unmatched와 Sheet2_Unmatched 재매칭
new_results = []
used_forex_numbers = set()
used_numbers = set()

# Sheet1_Unmatched 데이터 읽기
df_1_unmatched = pd.read_excel('locklock_test_output.xlsx', engine='openpyxl', sheet_name='Sheet1_Unmatched')
df_2_unmatched = pd.read_excel('locklock_test_output.xlsx', engine='openpyxl', sheet_name='Sheet2_Unmatched')

# 입력결제금액의 인접한 두 행을 더해서 매칭
for i in range(len(df_1_unmatched) - 1):
    row_1 = df_1_unmatched.iloc[i]
    next_row_1 = df_1_unmatched.iloc[i + 1]
    combined_money = row_1['입력결제금액'] + next_row_1['입력결제금액']
    
    for idx_2, row_2 in df_2_unmatched.iterrows():
        if row_1['무역거래처상호'] == row_2['무역거래처상호']:
            money_diff = abs(combined_money - row_2['송금외화금액'])
            if money_diff <= combined_money * 0.03:
                new_results.append({
                    '번호': row_1.name,
                    '무역거래처상호': row_1['무역거래처상호'],
                    '수리일자': row_1['수리일자'],
                    '입력결제금액': row_1['입력결제금액'],
                    '다음입력결제금액': next_row_1['입력결제금액'],
                    '조합된입력결제금액': combined_money,
                    '송금일자': row_2['송금일자'],
                    '송금외화금액': row_2['송금외화금액'],
                    '해외공급자상호': row_1['해외공급자상호'],
                    '수취인계좌번호': row_2['수취인계좌번호'],
                    'B/L번호': row_1['B/L번호'],
                    '외환송금번호': row_2['외환송금번호'],
                    '외환사유코드': row_2['외환사유코드']
                })
                used_forex_numbers.add(row_2['외환송금번호'])
                used_numbers.add(row_1.name)
                used_numbers.add(next_row_1.name)
                break

# Sheet1_Unmatched 데이터의 인접한 세 행을 더해서 매칭
for i in range(len(df_1_unmatched) - 2):
    row_1 = df_1_unmatched.iloc[i]
    next_row_1 = df_1_unmatched.iloc[i + 1]
    next_next_row_1 = df_1_unmatched.iloc[i + 2]
    combined_money = row_1['입력결제금액'] + next_row_1['입력결제금액'] + next_next_row_1['입력결제금액']
    
    for idx_2, row_2 in df_2_unmatched.iterrows():
        if row_1['무역거래처상호'] == row_2['무역거래처상호']:
            money_diff = abs(combined_money - row_2['송금외화금액'])
            if money_diff <= combined_money * 0.03:
                new_results.append({
                    '번호': row_1.name,
                    '무역거래처상호': row_1['무역거래처상호'],
                    '수리일자': row_1['수리일자'],
                    '입력결제금액': row_1['입력결제금액'],
                    '다음입력결제금액': next_row_1['입력결제금액'],
                    '다음다음입력결제금액': next_next_row_1['입력결제금액'],
                    '조합된입력결제금액': combined_money,
                    '송금일자': row_2['송금일자'],
                    '송금외화금액': row_2['송금외화금액'],
                    '해외공급자상호': row_1['해외공급자상호'],
                    '수취인계좌번호': row_2['수취인계좌번호'],
                    'B/L번호': row_1['B/L번호'],
                    '외환송금번호': row_2['외환송금번호'],
                    '외환사유코드': row_2['외환사유코드']
                })
                used_forex_numbers.add(row_2['외환송금번호'])
                used_numbers.add(row_1.name)
                used_numbers.add(next_row_1.name)
                used_numbers.add(next_next_row_1.name)
                break

# 리스트를 데이터프레임으로 변환
new_result_df = pd.DataFrame(new_results)

# 결과를 엑셀 파일로 저장
with pd.ExcelWriter('locklock_test_final_output.xlsx', engine='openpyxl') as writer:
    new_result_df.to_excel(writer, sheet_name='New_Matched', index=False)
    df_1_unmatched[~df_1_unmatched.index.isin(used_numbers)].to_excel(writer, sheet_name='Sheet1_Final_Unmatched', index=False)
    df_2_unmatched[~df_2_unmatched['외환송금번호'].isin(used_forex_numbers)].to_excel(writer, sheet_name='Sheet2_Final_Unmatched', index=False)
