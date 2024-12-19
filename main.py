import pandas as pd
import re

# 금액에서 'KRW'를 제거하고 숫자로 변환하는 함수
def convert_to_float(amount_str):
    try:
        return float(amount_str.replace('KRW', '').replace(',', '').strip())
    except ValueError:
        return 0.0  # 변환 실패 시 0.0 반환

# 거래수량에서 숫자만 추출하는 함수 (한글 및 코인 심볼 제외)
def extract_quantity(quantity_str):
    quantity = re.sub(r'[^\d\.]', '', quantity_str)  # 숫자와 '.'만 남기고 제거
    return float(quantity) if quantity else 0.0

# 엑셀 파일 읽기
def read_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# 수익 계산 함수 (가중평균법 적용)
def calculate_profit(df):
    buy_transactions = {}  # 코인별 매수 정보 (코인명: {수량, 평균 단가})
    total_profit = 0
    withdrawal_account = 0  # 출금 금액
    deposit = 0  # 입금 금액
    withdrawal_binance = 0  # 바이낸스 출금 금액
    
    # 수익을 기록할 새로운 컬럼 추가
    df['수익'] = 0.0

    # 데이터프레임을 아래에서 위로 처리
    df = df.iloc[::-1]  # 데이터프레임을 역순으로 읽음
    
    for index, row in df.iterrows():
        # 금액을 숫자로 변환
        buy_amount = extract_quantity(row['거래금액'])
        commission = extract_quantity(row['수수료'])
        settlement_amount = extract_quantity(row['정산금액'])
        
        # 거래수량에서 숫자만 추출
        quantity = extract_quantity(row['거래수량'])

        # 거래단가에서 숫자만 추출 (문자 제거)
        unit_price = extract_quantity(row['거래단가'])

        # 매수 거래는 buy_transactions에 저장 (가중평균법 적용)
        if row['종류'] == '매수':
            total_profit -= commission
            if row['코인'] in buy_transactions:
                # 기존에 해당 코인이 있으면, 새로운 수량과 단가로 평균 계산
                current_quantity = buy_transactions[row['코인']]['수량']
                current_avg_price = buy_transactions[row['코인']]['단가']
                # 새로운 평균 단가 계산
                total_spent = current_quantity * current_avg_price + quantity * unit_price
                new_avg_price = total_spent / (current_quantity + quantity)
                buy_transactions[row['코인']] = {
                    '수량': current_quantity + quantity,
                    '단가': new_avg_price,
                    '수수료': buy_transactions[row['코인']]['수수료'] + commission
                }
            else:
                # 처음 매수하는 코인인 경우, 새로운 단가와 수량으로 추가
                buy_transactions[row['코인']] = {
                    '수량': quantity,
                    '단가': unit_price,
                    '수수료': commission
                }

        # 매도 거래에서 수익 계산 (가중평균법 적용)
        elif row['종류'] == '매도':
            sell_amount = convert_to_float(row['거래금액'])
            remaining_quantity = quantity  # 매도할 수량 (남은 수량 추적)

            # 해당 코인이 매수 기록에 있는지 확인
            if row['코인'] in buy_transactions:
                # 매도 시 평균 매수 단가를 사용하여 수익 계산
                avg_buy_price = buy_transactions[row['코인']]['단가']
                total_sell_value = remaining_quantity * unit_price
                total_buy_value = remaining_quantity * avg_buy_price
                profit = total_sell_value - total_buy_value - commission
                df.at[index, '수익'] = profit
                total_profit += profit

                # 매도한 만큼 수량 차감
                buy_transactions[row['코인']]['수량'] -= remaining_quantity

                # 매도 후 수량이 0인 코인은 딕셔너리에서 삭제
                if buy_transactions[row['코인']]['수량'] <= 0:
                    del buy_transactions[row['코인']]

        # 입금 처리
        elif row['종류'] == '입금':
            deposit += buy_amount
        
        # 출금 처리 (입금 및 출금 시 'KRW'가 포함되어 있는지 확인)
        elif row['종류'] == '출금':
            if 'KRW' in row['거래수량']:
                withdrawal_account += buy_amount
            else:
                withdrawal_binance += buy_amount

    return df, total_profit, deposit, withdrawal_account, withdrawal_binance, buy_transactions

# 수익 합계 계산
def total_profit(df):
    return df['수익'].sum()

# 엑셀 파일 경로
file_path = 'upbit.xlsx'

# 엑셀 파일 읽기
df = read_excel(file_path)

# 수익 계산
df_with_profit, total, deposit, withdrawal_account, withdrawal_binance, buy_transactions = calculate_profit(df)

# 결과 출력
print("수익 계산 결과:")
print(df_with_profit[['체결시간', '코인', '마켓', '종류', '거래수량', '거래단가', '거래금액', '수수료', '수익']])

# 전체 수익 출력
print(f"\n전체 수익: {total:,.2f} 원")
print(f"\n전체 출금: {withdrawal_account:,.2f} 원")
print(f"\n바이낸스 출금: {withdrawal_binance:,.2f} 원")
print(f"\n전체 입금: {deposit:,.2f} 원")

have_coin_amount= (-withdrawal_binance)
# 남은 매수 코인과 그에 대한 수량 및 평균 단가 출력
print("\n남은 매수 코인과 수량, 평균 단가:")
for coin, data in buy_transactions.items():
    if data['수량'] > 0:
        total_coin_price = data['수량'] * data['단가']
        have_coin_amount += total_coin_price
        print(f"{coin}: 남은 수량 = {data['수량']:.4f}, 평균 단가 = {data['단가']:,.2f} 원, 총 금액 = {total_coin_price:,.2f} 원")

# 총 남은 코인 (원금))
print(f"\n총 남은 코인: {have_coin_amount:,.2f} 원")