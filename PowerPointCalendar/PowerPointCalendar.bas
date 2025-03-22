Attribute VB_Name = "PowerPointCalendar"
' グローバル変数としてDictionaryを保持
Public GlobalSettings As Object

' 一日分の日付・曜日・祝日・スケジュールを保持する構造体
Type daySchedule
    dateValue As Variant    ' この日の日付
    dayOfWeek As Variant    ' この日の曜日、日曜=1, 土曜=7
    IsHoliday As Boolean    ' 祝日なら True
    Memo As String          ' 日付の横に表示するメモ
    Items As String         ' この日のスケジュール、複数項目は改行(vbCrLf)を付けて連結する
End Type

' 一日分のスケジュールの配列
Public DaySchedules(1 To 31) As daySchedule
'''
''' Sub CreatePowerPointCalendar
'''     メイン関数
'''
Sub CreatePowerPointCalendar()

    ' 変数の定義
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim pptSlide As Object
    Dim pptShape As Object
    Dim savePath As String
    
    ' 設定の読み込み
    LoadSettingsIntoDictionary
    
    ' 予定表の読み込み
    LoadSchedules
    
    ' PowerPoint を起動
    On Error Resume Next
    Set pptApp = GetObject(Class:="PowerPoint.Application")
    If pptApp Is Nothing Then
        Set pptApp = CreateObject(Class:="PowerPoint.Application")
    End If
    On Error GoTo 0

    ' PowerPoint アプリケーションの表示
    pptApp.Visible = True
    
    ' 新規プレゼンテーションを作成
    Set pptPresentation = pptApp.Presentations.Add

    ' ページサイズを設定
    With pptPresentation.PageSetup
        .SlideWidth = GlobalSettings("SlideWidth")
        .SlideHeight = GlobalSettings("SlideHeight")
    End With

    ' 新規スライドを「白紙」に設定
    Set pptSlide = pptPresentation.Slides.Add(1, 12) ' 12 = ppLayoutBlank
    
    ' カレンダーを作成
    AddCalendar pptSlide
    
    ' 保存先のパスを設定 (「ドキュメント」フォルダに保存)
    savePath = Environ("USERPROFILE") & "\Documents\カレンダー " & Format(Now, "yyyymmdd hhnnss") & ".pptx"
    
    ' プレゼンテーションを保存
    pptPresentation.SaveAs savePath

    ' メッセージ表示
    MsgBox "PowerPointファイルを保存しました：" & vbCrLf & savePath, vbInformation

    ' PowerPoint を閉じる（必要なら）
    ' pptPresentation.Close
    ' pptApp.Quit
    
    ' 変数解放
    Set pptShape = Nothing
    Set pptSlide = Nothing
    Set pptPresentation = Nothing
    Set pptApp = Nothing

End Sub
'''
''' Sub LoadSettingsIntoDictionary
'''     「設定」シートから設定値を読み込んで GlobalSettings に保存する
'''
Sub LoadSettingsIntoDictionary()

    ' 変数の定義
    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim key As String, dataType As String, value As Variant
    Dim i As Long
    Dim vKey As Variant
    Dim output As String
    
    ' Dictionaryオブジェクトを作成
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 設定シートを取得
    Set ws = ThisWorkbook.Sheets("設定")
    
    ' 最終行を取得（B列を基準）
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' 2行目から最終行までループ（1行目はヘッダーと仮定）
    For i = 2 To lastRow
        key = Trim(ws.Cells(i, 2).value) ' B列の値をキーとして取得
        dataType = Trim(ws.Cells(i, 3).value) ' C列の値をデータタイプとして取得
        
        ' キーが空でなければ、D列の値を取得してDictionaryに格納
        Set cell = ws.Cells(i, 4)
        If key <> "" Then
            If dataType = "Value" Then
                dict(key) = CSng(cell.value)
            ElseIf dataType = "Color" Then
                dict(key) = cell.Interior.Color
            Else
                dict(key) = Trim(cell.value)
            End If
        End If
    Next i
    
    ' Dictionaryをグローバル変数または後続処理で利用できるように設定
    Set GlobalSettings = dict
    
    ' 読み込み数を確認（デバッグ用）
    Debug.Print "設定データを読み込みました。" & vbCrLf & "項目数: " & dict.Count
    
    ' 一覧を出力（格納したデータを改めて取り出して確認）
    output = ""
    Debug.Print "=== 設定一覧 ==="
    For Each vKey In dict.Keys
        output = output & vKey & " -> " & dict(vKey) & vbCrLf
    Next vKey
    Debug.Print output
    
    ' 変数開放
    Set dict = Nothing
    
End Sub
'''
''' Sub LoadSchedules
'''     「予定表」シートからひと月分の予定を読み込んで DaySchedules に保存する
'''
Sub LoadSchedules()

    ' 変数の定義
    Dim ws As Worksheet
    Dim i As Long
    Dim dateValue As Variant
    
    ' 予定表シートを取得
    Set ws = ThisWorkbook.Sheets("予定表")
    
    ' 1日(2行目)から31日(32行目)までループしてデータを取得、日付がないときはヌルを設定
    For i = 1 To 31
        dateValue = ws.Cells(i + 1, 5)
        If IsDate(dateValue) Then
            DaySchedules(i).dateValue = dateValue
            DaySchedules(i).dayOfWeek = Weekday(dateValue)
            DaySchedules(i).IsHoliday = (ws.Cells(i + 1, 8) = "祝")
            DaySchedules(i).Memo = ws.Cells(i + 1, 9)
            DaySchedules(i).Items = ws.Cells(i + 1, 10) & vbCrLf & _
                                    ws.Cells(i + 1, 11) & vbCrLf & _
                                    ws.Cells(i + 1, 12) & vbCrLf & _
                                    ws.Cells(i + 1, 13)
        Else
            ' 日付ではない、つまり2月の29日以降や4,6,9,11月の31日のときはヌルを設定
            DaySchedules(i).dateValue = Null
        End If
    Next i

End Sub
'''
''' Sub AddCalendar
'''     ひと月分のカレンダーを作成する
'''
'''     pptSlide : Slide オブジェクト
'''
Sub AddCalendar(pptSlide)

    ' 変数の定義
    Dim marginLeft, marginTop As Single
    Dim posLeft, posTop As Single
    Dim boxWidth, weekBoxHeight, dayBoxHeight As Single
    Dim i As Integer
    Dim dayOfWeek As Integer
    
    ' 左と上マージンの初期値の取得
    marginLeft = GlobalSettings("MarginLeft")
    marginTop = GlobalSettings("MarginTop")
    
    ' 曜日ヘッダーと日付の四角のサイズの取得
    boxWidth = GlobalSettings("BoxWidth")
    weekBoxHeight = GlobalSettings("WeekBoxHeight")
    dayBoxHeight = GlobalSettings("DayBoxHeight")
    
    ' 四角の上下左右の隙間
    interval = GlobalSettings("Interval")
    
    ' 曜日ヘッダーの作成
    posTop = marginTop
    posLeft = marginLeft
    Call AddWeekHeader(pptSlide, posLeft, posTop, boxWidth, weekBoxHeight, "日", rgbRed)
    
    posLeft = posLeft + boxWidth + interval
    Call AddWeekHeader(pptSlide, posLeft, posTop, boxWidth, weekBoxHeight, "月", rgbBlack)
    
    posLeft = posLeft + boxWidth + interval
    Call AddWeekHeader(pptSlide, posLeft, posTop, boxWidth, weekBoxHeight, "火", rgbBlack)
    
    posLeft = posLeft + boxWidth + interval
    Call AddWeekHeader(pptSlide, posLeft, posTop, boxWidth, weekBoxHeight, "水", rgbBlack)
    
    posLeft = posLeft + boxWidth + interval
    Call AddWeekHeader(pptSlide, posLeft, posTop, boxWidth, weekBoxHeight, "木", rgbBlack)
    
    posLeft = posLeft + boxWidth + interval
    Call AddWeekHeader(pptSlide, posLeft, posTop, boxWidth, weekBoxHeight, "金", rgbBlack)
    
    posLeft = posLeft + boxWidth + interval
    Call AddWeekHeader(pptSlide, posLeft, posTop, boxWidth, weekBoxHeight, "土", rgbBlue)
    
    ' 各日付の作成、上マージンの初期値を設定
    posTop = posTop + weekBoxHeight + interval

    ' ひと月の31日分をループ
    For i = 1 To 31
        ' 日付が無ければ処理終了
        If IsNull(DaySchedules(i).dateValue) Then GoTo EndOfSub
        
        ' 曜日に応じて左からの位置を計算
        dayOfWeek = DaySchedules(i).dayOfWeek
        posLeft = marginLeft + (boxWidth + interval) * (dayOfWeek - 1)
        Call AddDayBox(pptSlide, posLeft, posTop, boxWidth, dayBoxHeight, i)
        
        ' 週終わりに縦方向のインクリメント
        If 0 = (dayOfWeek Mod 7) Then
            posTop = posTop + dayBoxHeight + interval
        End If
    Next

' この月の処理終了
EndOfSub:

End Sub
'''
''' Sub AddWeekHeader
'''     1つの曜日ヘッダーの角丸四角シェープを作成する
'''
'''     pptSlide : Slide オブジェクト
'''     posLeft : 作成するシェープの水平位置
'''     posTop : 作成するシェープの垂直位置
'''     sizeWidth : 作成するシェープの幅
'''     sizeHeight : 作成するシェープの高さ
'''     textWeek : 曜日を示す文字列
'''     rgbTitle : 文字列の色
'''
Sub AddWeekHeader(pptSlide, posLeft, posTop, sizeWidth, sizeHeight, textWeek, rgbTitle)

    ' 変数の定義
    Dim pptShape As Object

     ' 角丸図形を作成
    Set pptShape = pptSlide.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                            Left:=posLeft, Top:=posTop, Width:=sizeWidth, Height:=sizeHeight)
                                    
    ' 角丸の調整
    pptShape.Adjustments.Item(1) = 0.03
    
    ' 図形の色を設定
    pptShape.Fill.ForeColor.RGB = GlobalSettings("WeekBoxFillColor") ' フィルの色
    pptShape.Line.ForeColor.RGB = GlobalSettings("WeekBoxLineColor") ' 線の色
                
    ' 曜日テキストの設定と調整
    With pptShape.TextFrame
        .TextRange.Text = textWeek
        .MarginBottom = 0
        .marginLeft = 0
        .MarginRight = 0
        .marginTop = 0
        .WordWrap = False
        .VerticalAnchor = msoAnchorTop
        .TextRange.ParagraphFormat.Alignment = 2 ' ppAlignCenter = 2
    End With
    
    ' フォントを「BIZ UDPゴシック」にして色を設定
    With pptShape.TextFrame.TextRange.Font
        .Size = 12
        .Name = "BIZ UDPGothic"
        .Bold = True
        .Color.RGB = rgbTitle
    End With
    
    ' 変数開放
    Set pptShape = Nothing
    
End Sub
'''
''' Sub AddDayBox
'''     1日分の角丸四角シェープを作成する
'''
'''     pptSlide : Slide オブジェクト
'''     posLeft : 作成するシェープの水平位置
'''     posTop : 作成するシェープの垂直位置
'''     sizeWidth : 作成するシェープの幅
'''     sizeHeight : 作成するシェープの高さ
'''     i : 作成対象の日付の DaySchedules 配列のインデックス
'''
Sub AddDayBox(pptSlide, posLeft, posTop, sizeWidth, sizeHeight, i)

    ' 変数の定義
    Dim pptShape As Object
    Dim daySchedule As daySchedule
    Dim textDay As String
    Dim myShape As Variant

    ' 一桁のときは空白を追加
    daySchedule = DaySchedules(i)
    If 10 > i Then
        textDay = " " & i
    Else
        textDay = i
    End If
    
    ' 日曜と祝日の文字を赤く、土曜日を青くする
    If 1 = daySchedule.dayOfWeek Or daySchedule.IsHoliday Then
        colorDay = rgbRed
    ElseIf 7 = daySchedule.dayOfWeek Then
        colorDay = rgbBlue
    Else
        colorDay = rgbBlack
    End If
        
    ' 角丸図形を作成
    Set pptShape = pptSlide.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                            Left:=posLeft, Top:=posTop, Width:=sizeWidth, Height:=sizeHeight)
  
    ' 角丸の調整
    pptShape.Adjustments.Item(1) = 0.03
    
    ' 図形の色を設定
    pptShape.Fill.ForeColor.RGB = GlobalSettings("DayBoxFillColor") ' フィルの色
    pptShape.Line.ForeColor.RGB = GlobalSettings("DayBoxLineColor") ' 線の色
    
    ' 日付の数字の設定と調整
    With pptShape.TextFrame
        .TextRange.Text = textDay
        .MarginBottom = 0
        .marginLeft = 0
        .MarginRight = 0
        .marginTop = 0
        .WordWrap = False
        .VerticalAnchor = msoAnchorTop
        .TextRange.ParagraphFormat.Alignment = 1 ' ppAlignLeft = 1
    End With
    
    ' フォントを「BIZ UDPゴシック」にして色を設定
    With pptShape.TextFrame.TextRange.Font
        .Size = 12
        .Name = "BIZ UDPGothic"
        .Bold = True
        .Color.RGB = colorDay
    End With
    
    ' 日付の横にその日のメモを追加して、フォントを「BIZ UDPゴシック」に指定
    Set textRangeMemo = pptShape.TextFrame.TextRange.InsertAfter(" " & daySchedule.Memo & vbCrLf)
    With textRangeMemo.Font
        .Size = 9
        .Name = "BIZ UDPGothic"
        .NameFarEast = "BIZ UDPGothic"
        .Bold = True
        .Color.RGB = rgbBlack
    End With
    
    ' その日の項目を追加して、フォント「UD デジタル 教科書体 NK-R」と色を設定
    ' NameFarEast にも指定しないとデフォルトフォントになるので注意
    Set textRangeItems = textRangeMemo.InsertAfter(daySchedule.Items)
    With textRangeItems.Font
        .Size = 9
        .Name = "UD Digi Kyokasho NK-R"
        .NameFarEast = "UD デジタル 教科書体 NK-R"
        .Bold = False
        .Color.RGB = rgbBlack
    End With
    
    ' 変数開放
    Set pptShape = Nothing
    
End Sub
