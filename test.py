
import csv, itertools
from functools import reduce

class aggregateStart():

    
    
    
    
    
    


    def start(self,datePath):
        csv_reader_N = csv.reader(open(datePath, "r", encoding="SJIS"))
        csv_reader_O = csv.reader(open(datePath, "r", encoding="SJIS"))
        #商品コード処理
        itemDict={}
        
        with open("/Users/koui/Desktop/商品コード一覧.txt","r") as f:
            for lines in f.readlines():
                itemStr=lines.strip().split("\t")
                if itemStr[0] not in itemDict:
                    itemDict[itemStr[0]]=[itemStr[2]]
                else:
                    itemDict[itemStr[0]]=itemDict[itemStr[0]]+[itemStr[2]]

        startDate=20170630
        endDate=20170731

        #顧客セット


        def sortCustomer(eventDict,date,csv_reader):
            for order in csv_reader:
                dmEvent=order[0]
                orderEvent=order[1]
                customerCD=order[2]
                orderNo=order[3]
                sendMonth=order[4][0:6]
                orderDate=order[5]
                itemCD=order[6]
                #登録していないコードpass
                if itemCD not in itemDict:
                    continue
                #集計日以外スキップ
                if int(order[4])>date:
                    continue

                if dmEvent not in eventDict:
                    event={}
                else:
                    event=eventDict[dmEvent]
                for item in itemDict[itemCD]:
                    if item not in event:
                        customer = {}
                    else:
                        customer=event[item]
                    if customerCD not in customer:
                        send={}
                    else:
                        send=customer[customerCD]
                    if sendMonth not in send:
                        orderNoList=[]
                    else:
                        orderNoList=send[sendMonth]
                    if orderNo not in orderNoList:
                        orderNoList=orderNoList+[orderNo]
                    send[sendMonth]=orderNoList
                    customer[customerCD]=send
                    event[item]=customer
                    eventDict[dmEvent]=event

        eventDict_N={}
        eventDict_O={}

        sortCustomer(eventDict_O,startDate,csv_reader_O)
        sortCustomer(eventDict_N,endDate,csv_reader_N)
        #顧客ソート完了
        #
        def aggregate(aggregateDict,eventDict):
            for dmEvent in eventDict:
                for item in eventDict[dmEvent]:
                    for customerCD in eventDict[dmEvent][item]:
                        sendMonthLast=max([x for x in eventDict[dmEvent][item][customerCD]])
                        sendMonthFirst=min([x for x in eventDict[dmEvent][item][customerCD]])
                        shoppingTimes=reduce(lambda x,y:x+y,[len(eventDict[dmEvent][item][customerCD][x]) for x in eventDict[dmEvent][item][customerCD]])

                        if shoppingTimes>12:
                            shoppingTimes=12
                        if dmEvent not in aggregateDict:
                            event={}
                        else:
                            event=aggregateDict[dmEvent]
                        if item not in event:
                            customer={}
                        else:
                            customer=event[item]
                        if shoppingTimes not in customer:
                            send={}
                        else:
                            send=customer[shoppingTimes]
                        if  sendMonthLast not in send:
                            send[sendMonthLast] = 1
                        else:
                            send[sendMonthLast] = send[sendMonthLast]+1
                        #初回処理
                        if 0 not in customer:
                            send0={}
                        else:
                            send0=customer[0]
                        if sendMonthFirst not in send0:
                            send0[sendMonthFirst]=1
                        else:
                            send0[sendMonthFirst]=send0[sendMonthFirst]+1

                        #顧客合計処理
                        if 13 not in customer:
                            send13={}
                        else:
                            send13=customer[13]
                        if sendMonthLast not in send13:
                            send13[sendMonthLast]=1
                        else:
                            send13[sendMonthLast]=send13[sendMonthLast]+1


                        customer[13]=send13
                        customer[0]=send0
                        customer[shoppingTimes]=send
                        event[item]=customer
                        aggregateDict[dmEvent]=event

        aggregateDict_N = {}
        aggregateDict_O = {}
        aggregate(aggregateDict_N,eventDict_N)
        aggregate(aggregateDict_O,eventDict_O)

        #集計完了
        #test dict
        def test(Dict):
            for x in Dict:
                for y in Dict[x]:
                    for j in Dict[x][y]:
                        print(x, y, j, Dict[x][y][j])

        #test(aggregateDict_N)
        #出力
        import xlrd,xlwt
        from xlwt import easyxf
        wb=xlwt.Workbook()
        sh=wb.add_sheet("RF表")

        #出力用の年リストを作る
        monthList=[]
        monthEnd=int(20170731/100)
        for x in range(24):
            monthList.append(str(monthEnd))
            monthEnd=monthEnd-1
            if monthEnd%100==0:
                monthEnd=monthEnd-100+12
        print(monthList)


        styleblue = easyxf('pattern: pattern solid, fore_colour light_turquoise;')
        stylelightblue=easyxf('pattern: pattern solid, fore_colour light_turquoise;')





        for dmEvent in [x for x in aggregateDict_N]:
            print(dmEvent)
            eventN=aggregateDict_N[dmEvent]
            if dmEvent in aggregateDict_O:
                eventO = aggregateDict_O[dmEvent]
            else:
                eventO={}
            rowNumber_O = 0
            for item in [x for x in aggregateDict_N[dmEvent]]:
                cellNumber = 0
                rowNumber_N = rowNumber_O + 16
                rowNumber_F = rowNumber_O + 32
                sh.write(rowNumber_O, cellNumber, item)
                cellNumber = cellNumber + 1
                sh.write(rowNumber_O, cellNumber, "年月")
                sh.write(rowNumber_N, cellNumber, "年月")
                sh.write(rowNumber_F, cellNumber, "年月")

                for month in monthList:
                    cellNumber=cellNumber+1
                    sh.write(rowNumber_O, cellNumber, month)
                    sh.write(rowNumber_N, cellNumber, month)
                    sh.write(rowNumber_F, cellNumber, month)
                rowNumber_O = rowNumber_O + 1
                rowNumber_N = rowNumber_N + 1
                rowNumber_F = rowNumber_F + 1

                customerN=eventN[item]
                if item in eventO:
                    customerO=eventO[item]
                else:
                    customerO={}

                for shoppingTimes in range(0,14):

                    if shoppingTimes==12:
                        sh.write(rowNumber_O, 1, str(shoppingTimes)+'~')
                        sh.write(rowNumber_N, 1, str(shoppingTimes)+'~')
                        sh.write(rowNumber_F, 1, str(shoppingTimes)+'~')
                    elif shoppingTimes==13:
                        sh.write(rowNumber_O, 1, "合計")
                        sh.write(rowNumber_N, 1, "合計")
                        sh.write(rowNumber_F, 1, "合計")
                    else:
                        sh.write(rowNumber_O, 1, shoppingTimes)
                        sh.write(rowNumber_N, 1, shoppingTimes)
                        sh.write(rowNumber_F, 1, shoppingTimes)


                    if shoppingTimes in customerN:
                        sendN=customerN[shoppingTimes]
                    else:
                        sendN={}
                    if shoppingTimes in customerO:
                        sendO=customerO[shoppingTimes]
                    else:
                        sendO={}

                    cellNumber = 2

                    for sendMonth in monthList:
                        humen_N, humen_O = 0, 0
                        # Oの出力
                        if sendMonth in sendO:

                            humen_O=sendO[sendMonth]
                        #Nの出力
                        if sendMonth in sendN:

                            humen_N = sendN[sendMonth]

                        if humen_O>0:
                            rate=(humen_O - humen_N)/humen_O
                        else:
                            rate=0
                        if item=="HO" and shoppingTimes==10 and sendMonth=="201608":
                            print(humen_O,humen_N,rate)

                        #パーセンテージの出力
                        if rate>=0.3:
                            if sendMonth in sendO:
                                sh.write(rowNumber_O, cellNumber, humen_O,styleblue)
                            else:
                                sh.write(rowNumber_O, cellNumber, None,styleblue)
                            if sendMonth in sendN:
                                sh.write(rowNumber_N, cellNumber, humen_N,styleblue)
                            else:
                                sh.write(rowNumber_N, cellNumber,  None,styleblue)
                            sh.write(rowNumber_F, cellNumber, format(rate, '.2%'),styleblue)
                        elif rate>=0.2:

                            if sendMonth in sendO:
                                sh.write(rowNumber_O, cellNumber, sendO[sendMonth],stylelightblue)
                            else:
                                sh.write(rowNumber_O, cellNumber, None,stylelightblue)
                            if sendMonth in sendN:
                                sh.write(rowNumber_N, cellNumber, sendN[sendMonth],stylelightblue)
                            else:
                                sh.write(rowNumber_N, cellNumber, None, stylelightblue)
                            sh.write(rowNumber_F, cellNumber, format(rate, '.2%'),stylelightblue)
                        else:
                            if sendMonth in sendO:
                                sh.write(rowNumber_O, cellNumber, sendO[sendMonth])
                            if sendMonth in sendN:
                                sh.write(rowNumber_N, cellNumber, sendN[sendMonth])
                            if rate>0:
                                sh.write(rowNumber_F, cellNumber, format(rate, '.2%'))


                        cellNumber=cellNumber+1

                        #
                    rowNumber_O = rowNumber_O + 1
                    rowNumber_N = rowNumber_N + 1
                    rowNumber_F = rowNumber_F + 1
                rowNumber_O=rowNumber_O+35


            break
        wb.save('/Users/koui/Desktop/example.xls')

if __name__ == '__main__':
        a=aggregateStart()
        a.start()




