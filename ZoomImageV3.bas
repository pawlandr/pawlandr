Attribute VB_Name = "ZoomImage"
Private Sub auto_open()
 
Application.OnKey "^%{RIGHT}", "EnumImageV2" 'Запуск перенумерации картинок выполняется по нажатию клавиш Ctrl+Alt+Стрелка вправо

Application.OnKey "^%{LEFT}", "ImgScaleAll" 'Изменение масштаба всех картинок на листе выполняется по нажатию клавиш Ctrl+Alt+Стрелка влево

ThisWorkbook.OnSheetActivate = "DelImg" 'Сброс увеличенных картинок при переключении листов
 
End Sub

Private Sub ImgScaleAll()
    DelImg                  'Удаляются все увеличенные ккартинки
    dblSend = InputBox("Масштаб задается в виде положительной десятичной дроби" & Chr(13) & "(разделитель запятая)" & Chr(13) _
    & "Чем больше цифра, тем больше картинка", "Укажите масштаб для ВСЕХ картинок", 0.9)
    On Error Resume Next
    dblSend = CDbl(dblSend)  'Если  введено правильное число, через запятую, то все пойдёт дальше, иначе сообщение об  ошибке
    If Err Then
        If MsgBox("Вы ввели неверное значение" & Chr(13) & "Хотите повторить?", vbYesNo) = vbYes Then
            ImgScaleAll         'Перезапуск текущего макроса, при желании  пользователя повторить ввод
        Else: Exit Sub          'При отказе  пользователя повторить ввод - выход из макроса
        End If
    End If
    Err.Clear
    EnumImageV2 CDbl(dblSend)   'вызывается макрос  перенумерации картинок, но в нем не выполняется перенумерация, а просто присваивается новый масштаб
End Sub

Private Sub ImgScalePlus()
With ActiveSheet
    For Each ZmImg In .Shapes                                       'выполняется  проверка названий всех картинок на листе
      If ZmImg.Name Like "Zoom*" Then                               'Отбирается картинка, у которой в названии есть Zoom
        strImgName = Mid(ZmImg.Name, 5)                             'Вырезается имя исходной картинки
        varData = CDbl(.Shapes(strImgName).AlternativeText) + 0.1   'Определяется значение масштабированияя исходной картинки и увеличивается на 10%
        .Shapes(strImgName).AlternativeText = CStr(varData)         'Новое значение масштабирования присваивается исходной картинке
        ZoomImageV3 CStr(strImgName)                                'Вызывается макрос ZommImageV3
      End If
    Next
End With
End Sub

Private Sub ImgScaleMinus()
With ActiveSheet
    For Each ZmImg In .Shapes                                       'выполняется  проверка названий всех картинок на листе
      If ZmImg.Name Like "Zoom*" Then                               'Отбирается картинка, у которой в названии есть Zoom
        strImgName = Mid(ZmImg.Name, 5)                             'Вырезается имя исходной картинки
        varData = CDbl(.Shapes(strImgName).AlternativeText) - 0.1   'Определяется значение масштабированияя исходной картинки и уменьшается на 10%
        .Shapes(strImgName).AlternativeText = CStr(varData)         'Новое значение масштабирования присваивается исходной картинке
        ZoomImageV3 CStr(strImgName)                                'Вызывается макрос ZommImageV3
      End If
    Next
End With
End Sub


Private Sub EnumImageV2(Optional dblSnd As Double)
' Макрос находит все картинки в активной книге и нумерует их по порядку
' начиная с левого верхнего угла и после присвоения номера сразу назначает картинке макрос ZoomImageV3
' совершенно не важно когда запускать этот макрос: до назначения масштабирования или после
' масштаб,  который указан в замещающем тексте  представляет собой процент от размеров текущей  рабочей области представленной на экране
i = 1
    For Each varShtsItm In ActiveWorkbook.Sheets
        For Each varImgItm In varShtsItm.Shapes
            If varImgItm.Name Like "Image_*" Then                       'Отрабатывает при запуске пользователем макроса по изменению масштаба для всех картинок
                If dblSnd > 0 Then varImgItm.AlternativeText = dblSnd   'Если масштаб был изменён, то значение  записывается в Замещающий  текст  картинки
            Else                                                        'Если картинка ранее не нумеровалась, то меняется её  имя  и ей присваивается номер
                 varImgItm.Name = "Image_" & i                          'Новые Имя и номер картинки
                 varImgItm.OnAction = "ZoomImageV3"                     'Назначение макроса масштабирующего картинку
                 varImgItm.AlternativeText = "0,9"                      'Запись  в  замещающий  текст  масштаба по умолчанию
            End If
        i = i + 1
        Next
    Next

End Sub

Private Sub ZoomImageV3(Optional strImgName As String)
Attribute ZoomImageV3.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim dblWinHeight As Double, dblWinWidth As Double
    Dim dblWinCenterTop As Double, dblWinCenterLeft As Double   'переменные для определения параметров окна
    Dim objPict0 As Shape, objPict As Shape                     'переменные-объекты для работы с картинками
    Dim PictZoom As Double                                      'Переменная определяет размер картинки  по которому она будет отмасштабирована
    
    With ActiveWindow.VisibleRange                                              'Вычисляем параметры видимой на экране области
        dblWinHeight = WorksheetFunction.Round(.Height, 2)                      'Высота видимой области ячеек
        dblWinWidth = WorksheetFunction.Round(.Width, 2)                        'Ширина видимой области ячеек
        dblWinCenterTop = WorksheetFunction.Round(.Top + dblWinHeight / 2, 2)   'Расстояние сверху до центра видимой области ячеек
        dblWinCenterLeft = WorksheetFunction.Round(.Left + dblWinWidth / 2, 2)  'Расстояние слева до центра видимой области ячеек
    End With
    
    On Error Resume Next
    Set objPict0 = ActiveSheet.Shapes(Application.Caller)       'Обработка нажатия мышкой на картинке
    If Err Then
        V = strImgName
        Set objPict0 = ActiveSheet.Shapes(V)
    End If
    Err.Clear
    
    
    
    On Error Resume Next
    DelImg                               'Проверка наличия и удаление увеличенных рисунков, отмена назначения кнлавиши ESC (подпрограмма)
    If Err Then Exit Sub                 'Если  удаление картинки было вызвано отмасштабированной картинкой, то происходи выход из макроса
    Err.Clear
    
    On Error Resume Next
    сZoomWin = CDbl(objPict0.AlternativeText)    'переменная, задающая коэффициент масштабирования картинки относительно границ рабочей области окна, значение берётся из Альтернативного текста картинки
    If Err Then
        сZoomWin = 0.9                       'Если в Альтернативном тексте введено некорректное значение, то присваивается значение по умолчанию
        objPict0.AlternativeText = "0,9"
    End If
    Err.Clear
    
    Set objPict = objPict0.Duplicate        'Создание копии картинку, которая будет увеличиваться
    objPict.Name = "Zoom" & objPict.Name    'Добавление к новой картинке префикса "Zoom"
    objPict.LockAspectRatio = msoTrue       'Активация свойства рисунка,  при котором размеры изменяются пропорционально
    
    If dblWinHeight < dblWinWidth Then      'Проверка параметров окна, что больше высотиа или ширина окна
        PictZoom = dblWinHeight * сZoomWin  ' Если высота окна меньше ширины, то за основу берётся меньшая величина (высота)
    Else
        PictZoom = dblWinWidth * сZoomWin   ' Если высота окна больше ширины, то за основу берётся меньшая величина (ширина)
    End If
    
    With objPict                    'Работаем с картинкой и её свойствами
        If .Height > .Width Then    'Проверка параметров картинки
            .Height = PictZoom      'Если высота картинки больше ширины, то картинка масштабируется по высоте
        Else
            .Width = PictZoom       'Если высота картинки меньше ширины, то картинка масштабируется по ширине
        End If
        .Top = WorksheetFunction.Round(dblWinCenterTop - (.Height / 2), 2)  'Определение положения верхней границы картинки
        .Left = WorksheetFunction.Round(dblWinCenterLeft - (.Width / 2), 2) 'Определение положения левой границы картинки
    End With
  Application.OnKey "{ESC}", "DelImg"          'Назначение клавиши ESC для удаления увеличенных картинок
  Application.OnKey "^%{UP}", "ImgScalePlus" 'Увеличение масштаба отдельной картинки выполняется по нажатию клавиш Ctrl+Alt+Стрелка вверх
  Application.OnKey "^%{DOWN}", "ImgScaleMinus" 'Уменьшение масштаба отдельной картинки выполняется по нажатию клавиш Ctrl+Alt+Стрелка вниз
End Sub

Private Sub DelImg()

With ActiveSheet
    For Each ZmImg In .Shapes                       'выполняется  проверка названий всех картинок на листе
      If ZmImg.Name Like "Zoom*" Then ZmImg.Delete  'удаляются картинки с названием,  содержащим "Zoom"
    Next
End With
Application.OnKey "{ESC}"                           'Присвоение клавише ESC стандартной функции
Application.OnKey "^%{UP}"                           'Сброс  функционала  клавиш
Application.OnKey "^%{DOWN}"                         'Сброс  функционала  клавиш
End Sub
