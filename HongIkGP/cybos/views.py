import pythoncom
from django.http import JsonResponse, HttpResponse
import win32com.client
from django.views import View

import ctypes

instCpCybos = win32com.client.Dispatch('CpUtil.CpCybos')

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')

instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")

def home(reqeust):
    return HttpResponse("cybos api")
def initPlusCheck(request):
    pythoncom.CoInitialize()

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

class getAccInfo(View):
    def get(self, request):
        pythoncom.CoInitialize()

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, 333016201)

        objTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
        initCheck = objTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문초기화실패")
            exit()

        self.objRq.BlockRequest()
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()

        print("통신상태", rqStatus, rqRet)
        # 통신 상태 확인
        if rqStatus != 0:
            return False

        # 보유 종목 개수
        cnt = self.objRq.GetHeaderValue(7)

        results = {
            'stocks': [],
            'account': []
        }

        print("보유종목 리스트")
        # 해당 계좌가 보유하고 있는 종목코드, 종목명, 체결잔고수량(몇주) 확인
        for i in range(cnt):
            code = self.objRq.GetDataValue(12, i)  # 종목코드
            name = self.objRq.GetDataValue(0, i)  # 종목명
            quantity = self.objRq.GetDataValue(7, i)  # 체결잔고수량
            evalAmount = self.objRq.GetDataValue(9, i)  # 평가금액
            evalGoL = self.objRq.GetDataValue(10, i)  # 평가손익

            print(i, code, name, quantity)

            stock_data = {
                'code': code,
                'name': name,
                'quantity': quantity,
                'evalAmount': evalAmount,
                'evalGoL': evalGoL
            }
            results['stocks'].append(stock_data)

        loanAmount = self.objRq.GetHeaderValue(6)  # 대출금액
        deposit = self.objRq.GetHeaderValue(9)  # 예수금
        daejuEvalAmount = self.objRq.GetHeaderValue(10)  # 대주평가금액
        jangoEvalAmount = self.objRq.GetHeaderValue(11)  # 잔고평가금액
        daejuAmount = self.objRq.GetHeaderValue(12)  # 대주금액

        total = jangoEvalAmount - daejuEvalAmount + deposit - loanAmount + daejuAmount
        results['account'].append({'total': total})
        print("\n잔고조회 : ", total)

        return JsonResponse(results, safe=False)

class getList(View):
    def get(self, request):
        pythoncom.CoInitialize()

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, 333016201)

        objTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
        initCheck = objTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문초기화실패")
            exit()

        self.objRq.BlockRequest()
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()

        print("통신상태", rqStatus, rqRet)
        # 통신 상태 확인
        if rqStatus != 0:
            return False

        # 보유 종목 개수
        cnt = self.objRq.GetHeaderValue(7)

        results = {
            'stocks': []
        }

        print("보유종목 리스트")
        # 해당 계좌가 보유하고 있는 종목코드, 종목명, 체결잔고수량(몇주) 확인
        for i in range(cnt):
            code = self.objRq.GetDataValue(12, i)  # 종목코드
            name = self.objRq.GetDataValue(0, i)  # 종목명
            quantity = self.objRq.GetDataValue(7, i)  # 체결잔고수량
            evalAmount = self.objRq.GetDataValue(9, i)  # 평가금액
            evalGoL = self.objRq.GetDataValue(10, i)  # 평가손익

            print(i, code, name, quantity)

            stock_data = {
                'code': code,
                'name': name,
                'quantity': quantity,
                'evalAmount': evalAmount,
                'evalGoL': evalGoL
            }
            results['stocks'].append(stock_data)


        return JsonResponse(results, safe=False)

class getTotal(View):
    def get(self, request):
        pythoncom.CoInitialize()

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, 333016201)

        objTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
        initCheck = objTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문초기화실패")
            exit()

        self.objRq.BlockRequest()
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()

        print("통신상태", rqStatus, rqRet)
        # 통신 상태 확인
        if rqStatus != 0:
            return False

        results = {
            'account': []
        }

        loanAmount = self.objRq.GetHeaderValue(6)  # 대출금액
        deposit = self.objRq.GetHeaderValue(9)  # 예수금
        daejuEvalAmount = self.objRq.GetHeaderValue(10)  # 대주평가금액
        jangoEvalAmount = self.objRq.GetHeaderValue(11)  # 잔고평가금액
        daejuAmount = self.objRq.GetHeaderValue(12)  # 대주금액

        total = jangoEvalAmount - daejuEvalAmount + deposit - loanAmount + daejuAmount
        results['account'].append({'total': total})
        print("\n잔고조회 : ", total)

        return JsonResponse(results, safe=False)