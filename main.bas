Sub main()
    Dim col_cnt As Integer, row_cnt As Integer
    Dim share_cnt As Integer
    Dim mbr_cnt As Integer, shr_cnt As Integer
    Dim member As String, share As String
    Dim asr_offset As Integer
    Dim date_yr As Integer, date_mon As Integer, date_d1 As Integer, date_d2 As Integer
    Dim date_offset_arr() As String, date_itl As Integer
    Dim answer_cnt As Integer, last_cnt As Integer
    Dim i As Integer
    Dim pass_cnt As Integer, lot_cnt As Integer
    Dim today_date As String, today_arr As Variant, fName As String, cPath As String

    col_cnt = 1
    row_cnt = 1
    share_cnt = 1
    answer_cnt = 1
    last_cnt = 1
    pass_cnt = 0
    lot_cnt = 0
    asr_offset = 0
    date_itl = 0
    today_date = Date
    date_yr = 0
    date_mon = 0
    date_d1 = 0
    date_d2 = 0

    Sheets("MemberList").Activate

    'Calculate cnt of columns
    Do While Cells(1, col_cnt).Value <> ""
        col_cnt = col_cnt + 1
    Loop
    col_cnt = col_cnt - 1

    Sheets.Add After := Sheets("MemberList")
    ActiveSheet.Name = "MemberList_Ftd"
    Sheets.Add After := Sheets("ShareList")
    ActiveSheet.Name = "抽獎名單"
    Sheets.Add After := Sheets("抽獎名單")
    ActiveSheet.Name = "中獎名單"

    'Calculate cnt of shared member
    Sheets("ShareList").Activate
    Do While Cells(share_cnt, 5).Value <> ""
        share_cnt = share_cnt + 1
    Loop
    share_cnt = share_cnt - 2

    Sheets("MemberList").Activate
    Rows(1).Select
    Selection.AutoFilter

    'Filter conditions for answering time and last login time
    Do While Cells(1, answer_cnt).Value <> "總答題數"
        answer_cnt = answer_cnt + 1
    Loop
    Do While Cells(1, last_cnt).Value <> "最後登入時間"
        last_cnt = last_cnt + 1
    Loop

    asr_offset = Application.Inputbox(Prompt := "請輸入答題數門檻", Type := 1)
    Do While asr_offset = 0
        asr_offset = Application.Inputbox(Prompt := "此欄位不可為零！"& vbCrlf & _
            "請輸入答題數門檻", Type := 1)
    Loop
    ActiveSheet.Range("A1").AutoFilter Field := answer_cnt, Criteria1 := ">=" & asr_offset

    date_yr = Application.Inputbox(Prompt := "請輸入欲篩選『最後登入日期』之『年份』, e.g. 2020", Type := 1)
    Do While (date_yr = 0) Or (date_yr > 2025) Or (date_yr < 2020)
        date_yr = Application.Inputbox(Prompt := "此欄位必須為2020~2025，且不可為0！"& vbCrlf & _
            "請輸入欲篩選『最後登入日期』之『年份』, e.g. 2020", Type := 1)
    Loop
    date_mon = Application.Inputbox(Prompt := "請輸入欲篩選『最後登入日期』之『月份』, e.g. 5", Type := 1)
    Do While (date_mon = 0) Or (date_mon > 12)
        date_mon = Application.Inputbox(Prompt := "此欄位必須為12以內，且不可為0！"& vbCrlf & _
            "請輸入欲篩選『最後登入日期』之『月份』, e.g. 5", Type := 1)
    Loop
    date_d1 = Application.Inputbox(Prompt := "請輸入欲篩選『最後登入日期』之『起始日』, e.g. 4", Type := 1)
    Do While (date_d1 = 0) Or (date_d1 > 30)
        date_d1 = Application.Inputbox(Prompt := "此欄位必須為31以內，且不可為0！"& vbCrlf & _
            "請輸入欲篩選『最後登入日期』之『起始日』, e.g. 4", Type := 1)
    Loop
    date_d2 = Application.Inputbox(Prompt := "請輸入欲篩選『最後登入日期』之『結束日』, e.g. 17", Type := 1)
    Do While (date_d2 = 0) Or (date_d2 < date_d1) Or (date_d2 > 31)
        date_d2 = Application.Inputbox(Prompt := "此欄位必須比『起始日: "& date_d1 &"』大，且不可為０並於31以內！"& _
            vbCrlf & "請輸入欲篩選『最後登入日期』之『結束日』, e.g. 17", Type := 1)
    Loop

    date_itl = date_d2 - date_d1
    ReDim date_offset_arr(date_itl + 1)
    For i = 0 To date_itl
        If date_mon < 10 Then
            date_offset_arr(i) = CStr(date_yr) & "-0" & CStr(date_mon)
        Else
            date_offset_arr(i) = CStr(date_yr) & "-" & CStr(date_mon)
        End If        
        If date_d1 + i < 10 Then
            date_offset_arr(i) = date_offset_arr(i) & "-0" & CStr(date_d1 + i)
        Else
            date_offset_arr(i) = date_offset_arr(i) & "-" & CStr(date_d1 + i)
        End If
    Next i
    ActiveSheet.Range("A1").AutoFilter Field := last_cnt, Criteria1 := date_offset_arr,  Operator := xlFilterValues

    Sheets("MemberList").UsedRange.Select
    Selection.Copy
    Sheets("MemberList_Ftd").Paste
    Sheets("MemberList_Ftd").Activate
    Cells(1, col_cnt + 1).Value = "是否分享"
    Cells(1, col_cnt + 2).Value = "亂數"
    Rows(1).Select
    Selection.AutoFilter

    'Calculate cnt of rows
    Do While Cells(row_cnt, 1).Value <> ""
        row_cnt = row_cnt + 1
    Loop
    row_cnt = row_cnt - 2

    For i = 2 To row_cnt
        Cells(i, col_cnt + 2).Value = Rnd()
    Next
    
    'Check shared or not
    MsgBox "(1) 本次共有" & share_cnt & "人分享活動" & vbCrlf & "(2) 有" & row_cnt & "人在規定時間內回答應答題數" & _ 
        vbCrlf & "開始搜尋符合上述兩條件之會員"
    oldStatusBar = Application.DisplayStatusBar 
    Application.DisplayStatusBar = True 
    For mbr_cnt = 1 To row_cnt
        member = Sheets("MemberList_Ftd").Cells(1 + mbr_cnt, 5).Value

        For shr_cnt = 1 To share_cnt
            share = Sheets("ShareList").Cells(1 + shr_cnt, 5).Value

            If member = share Then
                Sheets("MemberList_Ftd").Cells(1 + mbr_cnt, col_cnt + 1).Value = 1
            End If
        Next shr_cnt
        Application.StatusBar = "搜尋中...進度" & Format(mbr_cnt / row_cnt * 100, "0.0") & "%"
    Next mbr_cnt

    For i = 2 To row_cnt
        If Cells(i, col_cnt + 1).Value >= 1 Then
            pass_cnt = pass_cnt + 1
        End If
    Next i
    Application.StatusBar = False 
    Application.DisplayStatusBar = oldStatusBar
    MsgBox "結束搜尋，共有" & pass_cnt & "人合乎抽獎資格"
    ActiveSheet.Range("A1").AutoFilter Field := col_cnt + 1, Criteria1 := ">=" & CStr(1)
    Sheets("MemberList_Ftd").UsedRange.Select
    Selection.Copy
    Sheets("抽獎名單").Paste

    'Do lottery
    Sheets("抽獎名單").Activate
    Rows(1).Select
    Selection.AutoFilter
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key := Columns( _ 
        col_cnt + 2), SortOn := xlSortOnValues, Order := xlDescending, DataOption := _
        xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    lot_cnt = Application.Inputbox(Prompt := "本次需要抽出多少人(含候補)？", Type := 1)
    Do While lot_cnt = 0
        lot_cnt = Application.Inputbox(Prompt := "此欄位不可為零！"& vbCrlf & _
            "本次需要抽出多少人(含候補)？", Type := 1)
    Loop
    Rows("1:" & CStr(lot_cnt + 1)).Select
    Selection.Copy
    Sheets("中獎名單").Paste
    Sheets("中獎名單").Activate

    today_arr = Split(today_date, "/")
    If CInt(today_arr(1)) < 10 Then
        fName = "MemberList_" & today_arr(0) & "0" & today_arr(1)
    Else
        fName = "MemberList_" & today_arr(0) &  today_arr(1)
    End If 
    If CInt(today_arr(2)) < 10 Then
        fName = fName & "0" & today_arr(2)
    Else
        fName = fName & today_arr(2) & ".xlsx"
    End If
    cPath = ActiveWorkbook.Path
    ActiveWorkbook.SaveAs Filename := cPath & "/" & fName, FileFormat := xlOpenXMLWorkbook

    If lot_cnt >= pass_cnt Then
        MsgBox "合乎抽獎人數" & pass_cnt & "人，少於欲抽出人數" & lot_cnt & _
            "人，故不進行抽獎" & vbCrlf & "請參考『中獎名單』頁面，並另存新檔於" & _ 
            cPath & vbCrlf & "檔案名稱為" & fName
    Else
        MsgBox lot_cnt & "人已成功抽出，請參考『中獎名單』頁面" & _
            vbCrlf & "並另存新檔於" & cPath & _
            vbCrlf & "檔案名稱為" & fName
    End If
End Sub
