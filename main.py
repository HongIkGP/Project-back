from datetime import datetime

import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes

instCpCybos = win32com.client.Dispatch('CpUtil.CpCybos')

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False

    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False

    # 주문 관련 초기화
    if (g_objCpTrade.TradeInit(0) != 0):
        print("주문 초기화 실패")
        return False

    return True


class Cp6033:
    def __init__(self):
        acc = g_objCpTrade.AccountNumber[0]       # 계좌번호
        accFlag = g_objCpTrade.GoodsList(acc, 1)  # 주식상품 구분
        print(acc, accFlag[0])

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, acc)

        self.objRq.SetInputValue(2, 50)
        self.dicflag1 = {ord(' '): '현금',
                         ord('Y'): '융자',
                         ord('D'): '대주',
                         ord('B'): '담보',
                         ord('M'): '매입담보',
                         ord('P'): '플러스론',
                         ord('I'): '자기융자',
                         }

    def rq6033(self, caller):
        # 통신 및 통신 에러 처리
        self.objRq.BlockRequest()
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()

        if rqStatus != 0:
            print("통신상태", rqStatus, rqRet)
            return False

        # 보유 종목 갯수
        cnt = self.objRq.GetHeaderValue(7)
        print("보유 종목 갯수 : ", cnt)

        print('종목코드  종목명  보유수량')

        # 해당 계좌가 보유하고 있는 종목코드, 종목명, 체결잔고수량(몇주) 확인
        for i in range(cnt) :
            code = self.objRq.GetDataValue(12, i)  # 종목코드
            name = self.objRq.GetDataValue(0, i)    # 종목명
            amount = self.objRq.GetDataValue(7, i)  # 체결잔고수량
            test = self.objRq.GetDataValue(4, i)  # 체결잔고수량
            print(name, ": 결제장부단가", test)



            print(code,  name, amount)

    def totalAmount(self, caller):
        # 통신 및 통신 에러 처리
        self.objRq.BlockRequest()
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        if rqStatus != 0:
            print("통신상태", rqStatus, rqRet)
            return False


        print('총평가금액')
        loanAmount = self.objRq.GetHeaderValue(6) # 대출금액
        deposit = self.objRq.GetHeaderValue(9)  # 예수금
        daejuEvalAmount = self.objRq.GetHeaderValue(10) #대주평가금액
        jangoEvalAmount = self.objRq.GetHeaderValue(11) # 잔고평가금액
        daejuAmount = self.objRq.GetHeaderValue(12) # 대주금액

        total = jangoEvalAmount - daejuEvalAmount + deposit - loanAmount + daejuAmount
        print("총평가 : ", total)




print("Cybos Plus 를 시작합니다")

# 연결상태 확인
print(InitPlusCheck())

acc = g_objCpTrade.AccountNumber[0]

# Cp6033 객체 생성
test = Cp6033()


while (True):
    print("0: 프로그램 종료 \n1: 잔액 조회 \n2:종목 조회 \n")
    x = input("번호를 입력하시오: ")
    x = int(x)

    if x == 0:
        print(0)
        break
    if x is 1:
        # 객체의 rq6033 메서드 호출
        print("보유한 종목을 조회합니다.")
        test.rq6033(acc)

        # 보유 금액 가져와야함
    else:
        test.totalAmount(acc)

objTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
initCheck = objTrade.TradeInit(0)
if (initCheck != 0):
    print("주문초기화실패")
    exit()
