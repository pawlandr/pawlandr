Attribute VB_Name = "ZoomImage"
Private Sub auto_open()
 
Application.OnKey "^%{RIGHT}", "EnumImageV2" '������ ������������� �������� ����������� �� ������� ������ Ctrl+Alt+������� ������

Application.OnKey "^%{LEFT}", "ImgScaleAll" '��������� �������� ���� �������� �� ����� ����������� �� ������� ������ Ctrl+Alt+������� �����

ThisWorkbook.OnSheetActivate = "DelImg" '����� ����������� �������� ��� ������������ ������
 
End Sub

Private Sub ImgScaleAll()
    DelImg                  '��������� ��� ����������� ���������
    dblSend = InputBox("������� �������� � ���� ������������� ���������� �����" & Chr(13) & "(����������� �������)" & Chr(13) _
    & "��� ������ �����, ��� ������ ��������", "������� ������� ��� ���� ��������", 0.9)
    On Error Resume Next
    dblSend = CDbl(dblSend)  '����  ������� ���������� �����, ����� �������, �� ��� ����� ������, ����� ��������� ��  ������
    If Err Then
        If MsgBox("�� ����� �������� ��������" & Chr(13) & "������ ���������?", vbYesNo) = vbYes Then
            ImgScaleAll         '���������� �������� �������, ��� �������  ������������ ��������� ����
        Else: Exit Sub          '��� ������  ������������ ��������� ���� - ����� �� �������
        End If
    End If
    Err.Clear
    EnumImageV2 CDbl(dblSend)   '���������� ������  ������������� ��������, �� � ��� �� ����������� �������������, � ������ ������������� ����� �������
End Sub

Private Sub ImgScalePlus()
With ActiveSheet
    For Each ZmImg In .Shapes                                       '�����������  �������� �������� ���� �������� �� �����
      If ZmImg.Name Like "Zoom*" Then                               '���������� ��������, � ������� � �������� ���� Zoom
        strImgName = Mid(ZmImg.Name, 5)                             '���������� ��� �������� ��������
        varData = CDbl(.Shapes(strImgName).AlternativeText) + 0.1   '������������ �������� ���������������� �������� �������� � ������������� �� 10%
        .Shapes(strImgName).AlternativeText = CStr(varData)         '����� �������� ��������������� ������������� �������� ��������
        ZoomImageV3 CStr(strImgName)                                '���������� ������ ZommImageV3
      End If
    Next
End With
End Sub

Private Sub ImgScaleMinus()
With ActiveSheet
    For Each ZmImg In .Shapes                                       '�����������  �������� �������� ���� �������� �� �����
      If ZmImg.Name Like "Zoom*" Then                               '���������� ��������, � ������� � �������� ���� Zoom
        strImgName = Mid(ZmImg.Name, 5)                             '���������� ��� �������� ��������
        varData = CDbl(.Shapes(strImgName).AlternativeText) - 0.1   '������������ �������� ���������������� �������� �������� � ����������� �� 10%
        .Shapes(strImgName).AlternativeText = CStr(varData)         '����� �������� ��������������� ������������� �������� ��������
        ZoomImageV3 CStr(strImgName)                                '���������� ������ ZommImageV3
      End If
    Next
End With
End Sub


Private Sub EnumImageV2(Optional dblSnd As Double)
' ������ ������� ��� �������� � �������� ����� � �������� �� �� �������
' ������� � ������ �������� ���� � ����� ���������� ������ ����� ��������� �������� ������ ZoomImageV3
' ���������� �� ����� ����� ��������� ���� ������: �� ���������� ��������������� ��� �����
' �������,  ������� ������ � ���������� ������  ������������ ����� ������� �� �������� �������  ������� ������� �������������� �� ������
i = 1
    For Each varShtsItm In ActiveWorkbook.Sheets
        For Each varImgItm In varShtsItm.Shapes
            If varImgItm.Name Like "Image_*" Then                       '������������ ��� ������� ������������� ������� �� ��������� �������� ��� ���� ��������
                If dblSnd > 0 Then varImgItm.AlternativeText = dblSnd   '���� ������� ��� ������, �� ��������  ������������ � ����������  �����  ��������
            Else                                                        '���� �������� ����� �� ������������, �� �������� �  ���  � �� ������������� �����
                 varImgItm.Name = "Image_" & i                          '����� ��� � ����� ��������
                 varImgItm.OnAction = "ZoomImageV3"                     '���������� ������� ��������������� ��������
                 varImgItm.AlternativeText = "0,9"                      '������  �  ����������  �����  �������� �� ���������
            End If
        i = i + 1
        Next
    Next

End Sub

Private Sub ZoomImageV3(Optional strImgName As String)
Attribute ZoomImageV3.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim dblWinHeight As Double, dblWinWidth As Double
    Dim dblWinCenterTop As Double, dblWinCenterLeft As Double   '���������� ��� ����������� ���������� ����
    Dim objPict0 As Shape, objPict As Shape                     '����������-������� ��� ������ � ����������
    Dim PictZoom As Double                                      '���������� ���������� ������ ��������  �� �������� ��� ����� ����������������
    
    With ActiveWindow.VisibleRange                                              '��������� ��������� ������� �� ������ �������
        dblWinHeight = WorksheetFunction.Round(.Height, 2)                      '������ ������� ������� �����
        dblWinWidth = WorksheetFunction.Round(.Width, 2)                        '������ ������� ������� �����
        dblWinCenterTop = WorksheetFunction.Round(.Top + dblWinHeight / 2, 2)   '���������� ������ �� ������ ������� ������� �����
        dblWinCenterLeft = WorksheetFunction.Round(.Left + dblWinWidth / 2, 2)  '���������� ����� �� ������ ������� ������� �����
    End With
    
    On Error Resume Next
    Set objPict0 = ActiveSheet.Shapes(Application.Caller)       '��������� ������� ������ �� ��������
    If Err Then
        V = strImgName
        Set objPict0 = ActiveSheet.Shapes(V)
    End If
    Err.Clear
    
    
    
    On Error Resume Next
    DelImg                               '�������� ������� � �������� ����������� ��������, ������ ���������� �������� ESC (������������)
    If Err Then Exit Sub                 '����  �������� �������� ���� ������� ������������������ ���������, �� ��������� ����� �� �������
    Err.Clear
    
    On Error Resume Next
    �ZoomWin = CDbl(objPict0.AlternativeText)    '����������, �������� ����������� ��������������� �������� ������������ ������ ������� ������� ����, �������� ������ �� ��������������� ������ ��������
    If Err Then
        �ZoomWin = 0.9                       '���� � �������������� ������ ������� ������������ ��������, �� ������������� �������� �� ���������
        objPict0.AlternativeText = "0,9"
    End If
    Err.Clear
    
    Set objPict = objPict0.Duplicate        '�������� ����� ��������, ������� ����� �������������
    objPict.Name = "Zoom" & objPict.Name    '���������� � ����� �������� �������� "Zoom"
    objPict.LockAspectRatio = msoTrue       '��������� �������� �������,  ��� ������� ������� ���������� ���������������
    
    If dblWinHeight < dblWinWidth Then      '�������� ���������� ����, ��� ������ ������� ��� ������ ����
        PictZoom = dblWinHeight * �ZoomWin  ' ���� ������ ���� ������ ������, �� �� ������ ������ ������� �������� (������)
    Else
        PictZoom = dblWinWidth * �ZoomWin   ' ���� ������ ���� ������ ������, �� �� ������ ������ ������� �������� (������)
    End If
    
    With objPict                    '�������� � ��������� � � ����������
        If .Height > .Width Then    '�������� ���������� ��������
            .Height = PictZoom      '���� ������ �������� ������ ������, �� �������� �������������� �� ������
        Else
            .Width = PictZoom       '���� ������ �������� ������ ������, �� �������� �������������� �� ������
        End If
        .Top = WorksheetFunction.Round(dblWinCenterTop - (.Height / 2), 2)  '����������� ��������� ������� ������� ��������
        .Left = WorksheetFunction.Round(dblWinCenterLeft - (.Width / 2), 2) '����������� ��������� ����� ������� ��������
    End With
  Application.OnKey "{ESC}", "DelImg"          '���������� ������� ESC ��� �������� ����������� ��������
  Application.OnKey "^%{UP}", "ImgScalePlus" '���������� �������� ��������� �������� ����������� �� ������� ������ Ctrl+Alt+������� �����
  Application.OnKey "^%{DOWN}", "ImgScaleMinus" '���������� �������� ��������� �������� ����������� �� ������� ������ Ctrl+Alt+������� ����
End Sub

Private Sub DelImg()

With ActiveSheet
    For Each ZmImg In .Shapes                       '�����������  �������� �������� ���� �������� �� �����
      If ZmImg.Name Like "Zoom*" Then ZmImg.Delete  '��������� �������� � ���������,  ���������� "Zoom"
    Next
End With
Application.OnKey "{ESC}"                           '���������� ������� ESC ����������� �������
Application.OnKey "^%{UP}"                           '�����  �����������  ������
Application.OnKey "^%{DOWN}"                         '�����  �����������  ������
End Sub
