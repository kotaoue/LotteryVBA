VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const startRow = 2 '予定が記載されてる最初の行。1行目が日付=ヘッダーの場合は2
Const startColumn = 3 '予定が記載されてる最初の列。１行目が担当者名、2行目が累計抽選回数の場合は3
Const resultColumn = 2 '累計抽選回数が記載される列
Const inChargeCountOfDay = 2 '1日あたりに抽選する担当者数

'マクロを実行する時に選択する関数
Public Sub Main()
    Call InitResults
    
    For columnNo = startColumn To LastColumn()
        Dim targets As String
        targets = PickupTargets(columnNo)
        
        Call Lottery(columnNo, targets)
    Next
End Sub

'累計抽選回数をリセットする
Private Sub InitResults()
    For rowNo = startRow To LastRow()
        Cells(rowNo, resultColumn).Value = 0
    Next
End Sub

'指定日に予定の無いスタッフを抽選対象として抽出する
Private Function PickupTargets(ByVal columnNo As Integer)
    Dim targets As String
    targets = ""
    For rowNo = startRow To LastRow()
        'セルにNGって書かれていない人をカンマ区切りの文字列で抽出していく
        If Cells(rowNo, columnNo).Value <> "NG" Then
            If targets = "" Then
                '一人目だったらそのまま行番号を抽出
                targets = rowNo
            Else
                targets = targets & "," & rowNo
            End If
        End If
    Next
    
    PickupTargets = targets
End Function

'抽選を実施する
Private Sub Lottery(ByVal columnNo As Integer, ByVal targets As String)
    Debug.Print Cells(1, columnNo).Value & "の抽選対象は" & targets
    
    '配列inChargesは(0)のように0から順番にする
    Dim inCharges(inChargeCountOfDay - 1) As String
    Dim lotteryIndex As Integer: lotteryIndex = 0

    '予定が無いスタッフを、累担当計回数の少ない順にソートする
    targets = SortByResults(targets)
    targetArr = Split(targets, ",")
    For target = LBound(targetArr) To UBound(targetArr)
        inCharges(lotteryIndex) = targetArr(target)
        lotteryIndex = lotteryIndex + 1
        
        '日毎の必要担当者数になるまで、累計担当回数の少ない順に選ぶ
        If lotteryIndex >= inChargeCountOfDay Then
            Exit For
        End If
    Next target

    For i = LBound(inCharges) To UBound(inCharges)
        '予定があるスタッフが多い場合、inChargesには空文字が入っている
        If inCharges(i) <> "" Then
            Dim inCharge As Integer: inCharge = inCharges(i)
        
            '抽選結果をシートに書き込む
            Cells(inCharge, columnNo).Value = "担当"
        
            '抽選回数を増やす
            Cells(inCharge, resultColumn).Value = Cells(inCharge, resultColumn).Value + 1
        End If
    Next i
End Sub

'累計抽選回数の少ない順番に担当者をソートする
Private Function SortByResults(ByVal targets As String) As String
    targetArr = Split(targets, ",")
    
    'バブルソート
    For i = UBound(targetArr) To LBound(targetArr) Step -1
        For j = LBound(targetArr) To i - 1
            If Cells(targetArr(j), resultColumn).Value > Cells(targetArr(j + 1), resultColumn).Value Then
                Dim swap As String: swap = targetArr(j)
                targetArr(j) = targetArr(j + 1)
                targetArr(j + 1) = swap
            End If
        Next j
    Next i
    
    Debug.Print "sort from:" & targets & " to:" & Join(targetArr, ",")
    SortByResults = Join(targetArr, ",")
End Function

'最後の行番号を取得する
Private Function LastRow() As Integer
    LastRow = Cells(Rows.count, 1).End(xlUp).Row
End Function

'最後の列番号を取得する
Private Function LastColumn() As Integer
    LastColumn = Cells(1, Columns.count).End(xlToLeft).Column
End Function
