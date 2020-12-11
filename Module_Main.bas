Attribute VB_Name = "Module_Main"
Option Explicit

'�摜�_�E�����[�h�pAPI
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
  "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As LongPtr

'�G�������g����Key
Const KEY_ITEM_SHOP = "h5.s-line-clamp-1"
Const KEY_ITEM_NAME = "span.a-size-base-plus.a-color-base.a-text-normal"
Const KEY_ITEM_PRICE = "a-price-whole"
Const KEY_ITEM_IMG = "img.s-image"
Const KEY_ITEM_STAR = "span.a-icon-alt"

Sub MainProc()
    Dim scDriver As New Selenium.ChromeDriver               'Chrome����p�̃h���C�o
    Dim fso As New FileSystemObject                         '�t�@�C���V�X�e��
    Dim ws As Worksheet                                     '�������ʂ̃��[�N�V�[�g
    Dim oBy As New By                                       '.IsElementPresent�p
    Dim sItmeKwy As String                                  '���i�����G�������g
    Dim sImgUrl As String                                   '�摜�t�@�C����URL
    Dim sImgPath As String: sImgPath = "C:\temp\image"      '�摜�t�@�C���̕ۑ���
    Dim iRow As Long: iRow = 2                              '�������ʊi�[�s�J�E���g
    Dim i As Long
    
    
    '���j���[���猟���������擾
    With Worksheets("���j���[")
        Dim sAmazonUrl As String: sAmazonUrl = .Range("F3").Value
        Dim sKeyword As String: sKeyword = .Range("B3").Value
        Dim sSortIndex As String: sSortIndex = .Range("C3").Value
        Dim iMaxCount As Long: iMaxCount = .Range("D3").Value + 1   '�Z���̈ʒu��2�s�ڂ���J�n���Ă��邽��
    End With
    
    '�_�E�����[�h��̃t�H���_�����邩���݊m�F
    If fso.FolderExists(sImgPath) = False Then
        '���݂��Ȃ���΍쐬����
        fso.CreateFolder sImgPath
    End If
    '�t�@�C�����͌Œ�i�g�p��ɍ폜���邽�߁j
    sImgPath = sImgPath & "\work_img.jpg"

    '�������ʃV�[�g�̍쐬
    Sheets("����").Select
    Sheets("����").Copy
    ActiveSheet.Name = "��������"
    Set ws = Worksheets("��������")

    'Web�y�[�W�̑���
    With scDriver
        '�y�[�W���擾
        .Get sAmazonUrl
    
        '�^�C�g����Amazon���܂܂�Ă���Ώ��i�y�[�W�Ɣ���
        If InStr(.Window.Title, "Amazon") = 0 Then
            MsgBox "���i�y�[�W�̎擾���ł��܂���ł����B"
            End
        End If
      
        '���������ɃL�[���[�h����͂��A�N���b�N
        .FindElementById("twotabsearchtextbox").SendKeys sKeyword
        .FindElementByClass("nav-right").Click
        
        '���ёւ��Z���N�g�{�b�N�X�̑I������\������
        .FindElementById("a-autoid-0-announce").Click
        '�I�����̒�����A�Z���ɓ��͂��ꂽ���ڂ�I��
        For i = 1 To .FindElementsByClass("a-dropdown-item").Count
            With .FindElementsByClass("a-dropdown-link")(i)
                If .Text = sSortIndex Then
                    .Click
                    Exit For
                End If
            End With
        Next i
        
        Do Until iRow >= iMaxCount
            '���i�������ʂ̃y�[�W��3�`54�z���52���ƂȂ��Ă��邽��
            For i = 3 To 54
                sItmeKwy = "//*[@id=""search""]/div[1]/div[2]/div/span[3]/div[2]/div[" & i & "]"
                With .FindElementByXPath(sItmeKwy)
                    '���i��
                    ws.Cells(iRow, 4) = .FindElementByCss(KEY_ITEM_NAME).Text
                    '���i�摜�Y�t
                    If .IsElementPresent(oBy.Css(KEY_ITEM_IMG)) Then
                        sImgUrl = .FindElementByCss(KEY_ITEM_IMG).Attribute("src")
                        '�t�@�C���̃_�E�����[�h���s
                        If URLDownloadToFile(0, sImgUrl, sImgPath, 0, 0) = 0 Then
                            With ActiveSheet.Pictures.Insert(sImgPath)
                                .Top = Range("B" & iRow).Top
                                .Left = Range("B" & iRow).Left
                            End With
                            '�摜�t�@�C���̍폜
                            fso.DeleteFile sImgPath
                        End If
                    Else
                        ws.Cells(iRow, 2) = "N/A"
                    End If
                    '�o�i�ҏ��
                    If .IsElementPresent(oBy.Css(KEY_ITEM_SHOP)) Then
                        ws.Cells(iRow, 3) = .FindElementByCss(KEY_ITEM_SHOP).Text
                    Else
                        ws.Cells(iRow, 3) = "-"
                    End If
                    '���i
                    If .IsElementPresent(oBy.Class(KEY_ITEM_PRICE)) Then
                        ws.Cells(iRow, 5) = .FindElementByClass(KEY_ITEM_PRICE).Text
                    Else
                        ws.Cells(iRow, 5) = "-"
                    End If
                    '���iURL
                    ws.Cells(iRow, 6) = "https://www.amazon.co.jp/dp/" & .Attribute("data-asin")
                    ws.Hyperlinks.Add anchor:=ws.Cells(iRow, 6), Address:=ws.Cells(iRow, 6).Value
                    'If .IsElementPresent(oBy.Css(KEY_ITEM_URL)) Then
                    '    ws.Cells(iRow, 6) = .FindElementByCss(KEY_ITEM_URL).Attribute("href")
                    '    ws.Hyperlinks.Add anchor:=ws.Cells(iRow, 6), Address:=ws.Cells(iRow, 6).Value
                    'End If
                    '�]��
                    If .IsElementPresent(oBy.Css(KEY_ITEM_STAR)) Then
                        ws.Cells(iRow, 7) = .FindElementByCss(KEY_ITEM_STAR).Attribute("innerHTML")
                    Else
                        ws.Cells(iRow, 7) = "N/A"
                    End If
                    '�v�����ꂽ�����𖞂������ꍇ�A�������I������
                    If iRow = iMaxCount Then Exit For
                    iRow = iRow + 1
                End With
            Next i
            
            '�v�����ꂽ�����ɖ����Ȃ��ꍇ�A���̃y�[�W��\�����A�������J��Ԃ�
            If iRow < iMaxCount Then
                If .IsElementPresent(oBy.Css("li.a-last")) Then .FindElementByCss("li.a-last").Click
            End If
        Loop
        
        '�\��t�����摜�̃T�C�Y���Z�����Ɏ��߂�
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
    
    '��n��
    Set ws = Nothing

End Sub
