Attribute VB_Name = "Module_Main"
Option Explicit

'画像ダウンロード用API
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
  "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As LongPtr

'エレメント検索Key
Const KEY_ITEM_SHOP = "h5.s-line-clamp-1"
Const KEY_ITEM_NAME = "span.a-size-base-plus.a-color-base.a-text-normal"
Const KEY_ITEM_PRICE = "a-price-whole"
Const KEY_ITEM_IMG = "img.s-image"
Const KEY_ITEM_STAR = "span.a-icon-alt"

Sub MainProc()
    Dim scDriver As New Selenium.ChromeDriver               'Chrome操作用のドライバ
    Dim fso As New FileSystemObject                         'ファイルシステム
    Dim ws As Worksheet                                     '検索結果のワークシート
    Dim oBy As New By                                       '.IsElementPresent用
    Dim sItmeKwy As String                                  '商品検索エレメント
    Dim sImgUrl As String                                   '画像ファイルのURL
    Dim sImgPath As String: sImgPath = "C:\temp\image"      '画像ファイルの保存先
    Dim iRow As Long: iRow = 2                              '検索結果格納行カウント
    Dim i As Long
    
    
    'メニューから検索条件を取得
    With Worksheets("メニュー")
        Dim sAmazonUrl As String: sAmazonUrl = .Range("F3").Value
        Dim sKeyword As String: sKeyword = .Range("B3").Value
        Dim sSortIndex As String: sSortIndex = .Range("C3").Value
        Dim iMaxCount As Long: iMaxCount = .Range("D3").Value + 1   'セルの位置が2行目から開始しているため
    End With
    
    'ダウンロード先のフォルダがあるか存在確認
    If fso.FolderExists(sImgPath) = False Then
        '存在しなければ作成する
        fso.CreateFolder sImgPath
    End If
    'ファイル名は固定（使用後に削除するため）
    sImgPath = sImgPath & "\work_img.jpg"

    '検索結果シートの作成
    Sheets("原紙").Select
    Sheets("原紙").Copy
    ActiveSheet.Name = "検索結果"
    Set ws = Worksheets("検索結果")

    'Webページの操作
    With scDriver
        'ページを取得
        .Get sAmazonUrl
    
        'タイトルにAmazonが含まれていれば商品ページと判定
        If InStr(.Window.Title, "Amazon") = 0 Then
            MsgBox "商品ページの取得ができませんでした。"
            End
        End If
      
        '検索条件にキーワードを入力し、クリック
        .FindElementById("twotabsearchtextbox").SendKeys sKeyword
        .FindElementByClass("nav-right").Click
        
        '並び替えセレクトボックスの選択肢を表示する
        .FindElementById("a-autoid-0-announce").Click
        '選択肢の中から、セルに入力された項目を選択
        For i = 1 To .FindElementsByClass("a-dropdown-item").Count
            With .FindElementsByClass("a-dropdown-link")(i)
                If .Text = sSortIndex Then
                    .Click
                    Exit For
                End If
            End With
        Next i
        
        Do Until iRow >= iMaxCount
            '商品検索結果のページが3～54配列の52件となっているため
            For i = 3 To 54
                sItmeKwy = "//*[@id=""search""]/div[1]/div[2]/div/span[3]/div[2]/div[" & i & "]"
                With .FindElementByXPath(sItmeKwy)
                    '商品名
                    ws.Cells(iRow, 4) = .FindElementByCss(KEY_ITEM_NAME).Text
                    '商品画像添付
                    If .IsElementPresent(oBy.Css(KEY_ITEM_IMG)) Then
                        sImgUrl = .FindElementByCss(KEY_ITEM_IMG).Attribute("src")
                        'ファイルのダウンロード実行
                        If URLDownloadToFile(0, sImgUrl, sImgPath, 0, 0) = 0 Then
                            With ActiveSheet.Pictures.Insert(sImgPath)
                                .Top = Range("B" & iRow).Top
                                .Left = Range("B" & iRow).Left
                            End With
                            '画像ファイルの削除
                            fso.DeleteFile sImgPath
                        End If
                    Else
                        ws.Cells(iRow, 2) = "N/A"
                    End If
                    '出品者情報
                    If .IsElementPresent(oBy.Css(KEY_ITEM_SHOP)) Then
                        ws.Cells(iRow, 3) = .FindElementByCss(KEY_ITEM_SHOP).Text
                    Else
                        ws.Cells(iRow, 3) = "-"
                    End If
                    '価格
                    If .IsElementPresent(oBy.Class(KEY_ITEM_PRICE)) Then
                        ws.Cells(iRow, 5) = .FindElementByClass(KEY_ITEM_PRICE).Text
                    Else
                        ws.Cells(iRow, 5) = "-"
                    End If
                    '商品URL
                    ws.Cells(iRow, 6) = "https://www.amazon.co.jp/dp/" & .Attribute("data-asin")
                    ws.Hyperlinks.Add anchor:=ws.Cells(iRow, 6), Address:=ws.Cells(iRow, 6).Value
                    'If .IsElementPresent(oBy.Css(KEY_ITEM_URL)) Then
                    '    ws.Cells(iRow, 6) = .FindElementByCss(KEY_ITEM_URL).Attribute("href")
                    '    ws.Hyperlinks.Add anchor:=ws.Cells(iRow, 6), Address:=ws.Cells(iRow, 6).Value
                    'End If
                    '評価
                    If .IsElementPresent(oBy.Css(KEY_ITEM_STAR)) Then
                        ws.Cells(iRow, 7) = .FindElementByCss(KEY_ITEM_STAR).Attribute("innerHTML")
                    Else
                        ws.Cells(iRow, 7) = "N/A"
                    End If
                    '要求された件数を満たした場合、処理を終了する
                    If iRow = iMaxCount Then Exit For
                    iRow = iRow + 1
                End With
            Next i
            
            '要求された件数に満たない場合、次のページを表示し、処理を繰り返す
            If iRow < iMaxCount Then
                If .IsElementPresent(oBy.Css("li.a-last")) Then .FindElementByCss("li.a-last").Click
            End If
        Loop
        
        '貼り付けた画像のサイズをセル内に収める
        For i = 1 To ActiveSheet.Shapes.Count
            With ActiveSheet.Shapes(i)
                .IncrementLeft 6
                .IncrementTop 6
                If .Height > 90 Then .Height = 85
                If .Width > 200 Then .Width = 200
            End With
        Next i

        .Quit
    End With
    
    '後始末
    Set ws = Nothing

End Sub
