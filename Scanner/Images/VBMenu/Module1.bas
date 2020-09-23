Attribute VB_Name = "Module1"
Option Explicit
' §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
' §§                                                                                                                              §§
' §§                            Пример использования графики в стандартном VB меню                                                  §§
' §§                                        Автор: Анатолий Жуков                                                                 §§
' §§                                                                                                                              §§
' §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

' Параметры вывода меню (декларации)
Dim checkAreaRC                 As RECT
Dim pictureAreaRC               As RECT

Public m_Margin                 As Long
Public m_CheckAreaMenuColor     As Long
Public m_PictureAreaMenuColor   As Long
Public m_CaptionAreaMenuColor   As Long
Public m_TextMenuColor          As Long
Public m_SelectTextMenuColor    As Long
Public m_HotKeyTextMenuColor    As Long
Public m_SeparatopColor         As Long
Public m_PictureMaskColor       As Long
Public m_ShadowColor            As Long
Public m_FrameMenuColor         As Long
Public m_FrameMenuBackColor     As Long
Public m_GrayFrameMenuColor     As Long
Public m_GrayFrameMenuBackColor As Long
Public m_LabelRC                As RECT
Public m_LabelBackColor         As Long
Public m_LabelForeColor         As Long

Public m_MaxPictureWidth        As Long
Public MyFont                   As Long
Public OldFont                  As Long
Public hBr                      As Long
Public itemPicture()            As StdPicture
' Флаги DrawPicture
Public Const DP_COLOR = 0
Public Const DP_SHADOW = 1
' Форма...
Public m_Form As Form1
Public Property Get MenuForm() As Form1
     Set MenuForm = m_Form
End Property
Public Property Set MenuForm(ByVal vNewValue As Form1)
     Set m_Form = vNewValue
End Property
Public Function LoWord(LongIn As Long) As Integer
     If (LongIn And &HFFFF&) > &H7FFF Then
          LoWord = (LongIn And &HFFFF&) - &H10000
     Else
          LoWord = LongIn And &HFFFF&
     End If
End Function
Public Function HiWord(LongIn As Long) As Integer
     HiWord = (LongIn And &HFFFF0000) \ &H10000
End Function
' Функция модифицирующая элементы меню в MF_OWNERDRAW
Public Sub CreateOwnerDrawMenu(ByVal hMenu As Long, ByVal MenuID As Long, ByVal ItemData As Long)
     Dim dwordFlag      As Long
     Dim mii            As MENUITEMINFO
     mii.cbSize = Len(mii)
     mii.fMask = MIIM_TYPE
     GetMenuItemInfo hMenu, MenuID, False, mii
     ' Стандартный флаг
     dwordFlag = MF_BYCOMMAND Or MF_OWNERDRAW
     ' Дополнительный флаг  MF_SEPARATOR
     If ((mii.fType And MF_SEPARATOR) = MF_SEPARATOR) Then dwordFlag = dwordFlag Or MF_SEPARATOR
     ' Дополнительный флаг  MF_CHECKED
     If ((GetMenuState(hMenu, MenuID, MF_BYCOMMAND) And MF_CHECKED) = MF_CHECKED) Then dwordFlag = dwordFlag Or MF_CHECKED
     ' Дополнительный флаг  MF_DISABLED
     If ((GetMenuState(hMenu, MenuID, MF_BYCOMMAND) And MF_DISABLED) = MF_DISABLED) Then dwordFlag = dwordFlag Or MF_GRAYED
     ' Модифицируем меню сошласно созданного флага
     Call ModifyMenu(hMenu, MenuID, dwordFlag, MenuID, ItemData)
End Sub
' Проверка не является ли элемент меню IID сепаратором
Public Function IsSeparator(ByVal hwnd As Long, ByVal IID As Integer) As Boolean
     Dim mii As MENUITEMINFO
     mii.cbSize = Len(mii)
     mii.fMask = MIIM_TYPE
     mii.wID = IID
     GetMenuItemInfo GetMenu(hwnd), IID, False, mii
     IsSeparator = ((mii.fType And MFT_SEPARATOR) = MFT_SEPARATOR)
End Function
' Проверка не является ли элемент меню IID очекченым
Public Function IsChecked(ByVal hwnd As Long, ByVal IID As Integer) As Boolean
    If (GetMenuState(GetMenu(hwnd), IID, MF_BYCOMMAND) And MF_CHECKED) = MF_CHECKED Then
        IsChecked = True
    Else
        IsChecked = False
    End If
End Function
' Проверка не является ли элемент меню IID недоступным
Public Function IsGrayed(ByVal hwnd As Long, ByVal IID As Integer) As Boolean
    If (GetMenuState(GetMenu(hwnd), IID, MF_BYCOMMAND) And MF_GRAYED) = MF_GRAYED Then
        IsGrayed = True
    Else
        IsGrayed = False
    End If
End Function
' Прорисовка зоны лейбла
Public Sub DrawLabel(hDc As Long)
    m_LabelRC.Bottom = 1000
    FillRectangle hDc, m_LabelRC, m_LabelBackColor
End Sub
' Новая оконная процедура
Public Function MsgProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    ' Декларации
    Dim tempTextColor       As Long
    Dim MeasureInfo         As MEASUREITEMSTRUCT
    Dim DrawInfo            As DRAWITEMSTRUCT
    Dim mii                 As MENUITEMINFO
    Dim str1                As String
    Dim str2                As String
    ' Декларация строкового буффера
    Dim menuItemString      As String
    ' Создаем буффер
    menuItemString = String(100, " ")
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    '                  Обработка событий
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    ' Событие отвечающее за формирование ширины и высоты каждого элемента меню
    If wMsg = WM_MEASUREITEM Then
        Call CopyMem(MeasureInfo, ByVal lParam, Len(MeasureInfo))
        ' Получаем текст строки меню
        Call GetMenuString(GetMenu(hwnd), MeasureInfo.itemID, menuItemString, 100, MF_BYCOMMAND)
        ' Если есть зона чека
        checkAreaRC.Left = m_LabelRC.Right
        checkAreaRC.Right = checkAreaRC.Left + m_Margin + m_Form.ScaleX(m_Form.CheckImage.Picture.Width) + m_Margin
        ' Если это сепаратор...
        If IsSeparator(hwnd, MeasureInfo.itemID) Then
            MeasureInfo.itemWidth = 20
            MeasureInfo.itemHeight = 3
        Else
            ' Если есть картинка
            If itemPicture(MeasureInfo.itemID).Handle <> 0 Then
                ' Вычисляем ширину
                MeasureInfo.itemWidth = checkAreaRC.Right + m_Margin + m_Form.ScaleX(itemPicture(MeasureInfo.itemID).Width) + m_Margin + m_Form.TextWidth(Trim(menuItemString)) + m_Margin
                ' Вычисляем высоту
                ' Если высота картинки больше высоты текста то высота картинки определяющая...
                If m_Form.ScaleY(itemPicture(MeasureInfo.itemID).Height) >= m_Form.TextHeight(Trim(menuItemString)) Then
                    ' Отступ + высота картинки переведенная из HIMETRIC в шкалу формы + отступ
                    MeasureInfo.itemHeight = m_Margin + m_Form.ScaleY(itemPicture(MeasureInfo.itemID).Height) + m_Margin
                Else
                    ' Отступ + высота текста + отступ
                    MeasureInfo.itemHeight = m_Margin + m_Form.TextHeight(Trim(menuItemString)) + m_Margin
                End If
            Else
                ' Ширина состоит из максимальной ширины картинки + ширина текста
                MeasureInfo.itemWidth = checkAreaRC.Right + m_Margin + m_MaxPictureWidth + m_Margin + m_Form.TextWidth(Trim(menuItemString)) + m_Margin
                ' Отступ + высота текста + отступ
                MeasureInfo.itemHeight = m_Margin + m_Form.TextHeight(Trim(menuItemString)) + m_Margin
            End If
        End If
        Call CopyMem(ByVal lParam, MeasureInfo, Len(MeasureInfo))
        MsgProc = False
        Exit Function
    End If
    
    ' Общие установки для сообщения WM_DRAWITEM параметры
    If wMsg = WM_DRAWITEM And wParam = 0 Then
        Call CopyMem(DrawInfo, ByVal lParam, LenB(DrawInfo))
         ' Если есть зона чека
        checkAreaRC.Left = DrawInfo.rcItem.Left + m_LabelRC.Right
        checkAreaRC.Top = DrawInfo.rcItem.Top
        checkAreaRC.Right = checkAreaRC.Left + m_Margin + m_Form.ScaleX(m_Form.CheckImage.Picture.Width) + m_Margin
        ' Вычисляем зону картинки
        pictureAreaRC.Left = checkAreaRC.Right
        pictureAreaRC.Top = DrawInfo.rcItem.Top
        pictureAreaRC.Right = pictureAreaRC.Left + m_Margin + m_MaxPictureWidth + m_Margin
        ' Если есть картинка
        If itemPicture(DrawInfo.itemID).Handle <> 0 Then
            pictureAreaRC.Bottom = pictureAreaRC.Top + m_Margin + m_Form.ScaleY(itemPicture(DrawInfo.itemID).Height) + m_Margin
        Else
            pictureAreaRC.Bottom = pictureAreaRC.Top + m_Margin + m_Form.TextHeight(Trim(menuItemString)) + m_Margin
        End If
        checkAreaRC.Bottom = pictureAreaRC.Bottom
    End If
    ' ***************************************************************************************************
    If wMsg = WM_DRAWITEM Then
        If wParam = 0 Then
            ' Заряжаем шрифт
            MyFont = SendMessage(hwnd, WM_GETFONT, 0&, 0&)
            OldFont = SelectObject(DrawInfo.hDc, MyFont)
            ' Устанавливаем прозрачность
            Call SetBkMode(DrawInfo.hDc, TRANSPARENT)
            ' Получаем текст строки меню
            Call GetMenuString(GetMenu(hwnd), DrawInfo.itemID, menuItemString, 100, MF_BYCOMMAND)
            '--- MEASURE start --------------------------------------------------------------------------------------
            ' ODA_DRAWENTIRE Прорисовка при measure
            If DrawInfo.itemAction And ODA_DRAWENTIRE Then
                ' Если чек...
                If IsSeparator(hwnd, DrawInfo.itemID) Then
                   ' Закрашиваем зону чека
                    FillRectangle DrawInfo.hDc, checkAreaRC, m_CheckAreaMenuColor
                    ' Закрашиваем зону картинки
                    FillRectangle DrawInfo.hDc, pictureAreaRC, m_PictureAreaMenuColor
                    ' Смещаем прямоугольник в зоне заголовка
                    DrawInfo.rcItem.Top = DrawInfo.rcItem.Top + 1
                    DrawInfo.rcItem.Left = checkAreaRC.Right + m_Margin + m_MaxPictureWidth + m_Margin * 2
                    DrawInfo.rcItem.Bottom = DrawInfo.rcItem.Bottom - 1
                    ' Закрашиваем полосу в зоне заголовка
                    FillRectangle DrawInfo.hDc, DrawInfo.rcItem, m_SeparatopColor
                Else
                    ' Не захватывать зону лейбла
                    DrawInfo.rcItem.Left = m_LabelRC.Right
                    ' Закрашиваем полосу в зоне заголовка
                    FillRectangle DrawInfo.hDc, DrawInfo.rcItem, m_CaptionAreaMenuColor
                   ' Закрашиваем зону чека
                    FillRectangle DrawInfo.hDc, checkAreaRC, m_CheckAreaMenuColor
                    ' Закрашиваем зону картинки
                    FillRectangle DrawInfo.hDc, pictureAreaRC, m_PictureAreaMenuColor
                    ' Если меню Checked
                    If IsChecked(hwnd, DrawInfo.itemID) Then
                        DrawPicture DrawInfo.hDc, 0, 0, checkAreaRC, m_Form.CheckImage.Picture, DP_COLOR
                    End If
                    ' Если есть картинка вставляем ее
                    DrawPicture DrawInfo.hDc, 0, 0, pictureAreaRC, itemPicture(DrawInfo.itemID), DP_COLOR
                    ' Смещаем левый край прямоугольника для вывода текста на ширину картинки и отступы
                    DrawInfo.rcItem.Left = checkAreaRC.Right + m_Margin + m_MaxPictureWidth + m_Margin * 2
                    ' Если недоступно...
                    If IsGrayed(hwnd, DrawInfo.itemID) Then
                        ' Цвет шрифта
                        Call SetTextColor(DrawInfo.hDc, GetSysColor(COLOR_GRAYTEXT))
                    Else
                        ' Цвет шрифта
                        Call SetTextColor(DrawInfo.hDc, m_TextMenuColor)
                    End If
                    ' Если есть горячая клавиша разбиваем строку на две (разделитель Chr(9))
                    If PrsString(Trim(menuItemString), str1, str2) Then
                        ' Первую выводим слева...
                        Call DrawText(DrawInfo.hDc, str1, Len(str1), DrawInfo.rcItem, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER)
                        ' Отступ от правого края...
                        DrawInfo.rcItem.Right = DrawInfo.rcItem.Right - m_Margin
                        ' Ставим цвет HotKey
                        tempTextColor = SetTextColor(DrawInfo.hDc, m_HotKeyTextMenuColor)
                        ' Вторую выводим справа...
                        Call DrawText(DrawInfo.hDc, str2, Len(str2) - 1, DrawInfo.rcItem, DT_SINGLELINE Or DT_RIGHT Or DT_VCENTER)
                        ' Возврат первоначального значения...
                        Call SetTextColor(DrawInfo.hDc, tempTextColor)
                        DrawInfo.rcItem.Right = DrawInfo.rcItem.Right + m_Margin
                    Else
                        ' Пишем текст
                        Call DrawText(DrawInfo.hDc, Trim(menuItemString), Len(Trim(menuItemString)) - 1, DrawInfo.rcItem, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER)
                    End If
                    ' Закрашиваем зону Label
                    DrawLabel DrawInfo.hDc
                End If
                ' Выход
                MsgProc = False
                Exit Function
            End If
            '--- MEASURE end --------------------------------------------------------------------------------------
            ' &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            '--- ODS_SELECTED start ---------------------------------------------------------------------------------
            ' ODS_SELECTED Элемент меню выбран
            If (DrawInfo.itemState And ODS_SELECTED) Then ' And (DrawInfo.itemAction And (ODA_SELECT Or ODA_DRAWENTIRE))) Then
                ' Не захватывать зону лейбла
                DrawInfo.rcItem.Left = m_LabelRC.Right
                ' Временные переменные цветов...
                Dim tempFrameColor As Long, tempFrameBackColor As Long
                ' Если строка меню недоступна...
                If IsGrayed(hwnd, DrawInfo.itemID) Then
                    tempFrameBackColor = m_GrayFrameMenuBackColor
                    tempFrameColor = m_GrayFrameMenuColor
                Else
                    tempFrameBackColor = m_FrameMenuBackColor
                    tempFrameColor = m_FrameMenuColor
                End If
                ' Закрашиваем полосу в зоне заголовка
                FillRectangle DrawInfo.hDc, DrawInfo.rcItem, tempFrameBackColor
                ' Можно по круче с рамкой
                hBr = CreateSolidBrush(tempFrameColor)
                Call FrameRect(DrawInfo.hDc, DrawInfo.rcItem, hBr)
                Call DeleteObject(hBr)
                ' Если меню Checked
                If IsChecked(hwnd, DrawInfo.itemID) Then
                    DrawPicture DrawInfo.hDc, 0, 0, checkAreaRC, m_Form.CheckImage.Picture, DP_COLOR
                End If
                  ' Если есть картинка вставляем ее
                If itemPicture(DrawInfo.itemID).Handle <> 0 Then
                    DrawPicture DrawInfo.hDc, 1, 1, pictureAreaRC, itemPicture(DrawInfo.itemID), DP_SHADOW
                    DrawPicture DrawInfo.hDc, -1, -1, pictureAreaRC, itemPicture(DrawInfo.itemID), DP_COLOR
                End If
                ' Смещаем левый край прямоугольника для вывода текста на ширину картинки и отступы
                DrawInfo.rcItem.Left = checkAreaRC.Right + m_Margin + m_MaxPictureWidth + m_Margin * 2
                ' Если строка меню недоступна
                If IsGrayed(hwnd, DrawInfo.itemID) Then
                    ' Цвет шрифта
                    Call SetTextColor(DrawInfo.hDc, GetSysColor(COLOR_GRAYTEXT))
                Else
                    ' Эффект тени текста
                    Call SetTextColor(DrawInfo.hDc, m_ShadowColor)
                    ' Пишем текст
                    Call OffsetRect(DrawInfo.rcItem, 1, 1)
                    ' Если есть горячая клавиша разбиваем строку на две (разделитель Chr(9))
                    If PrsString(Trim(menuItemString), str1, str2) Then
                        ' Первую выводим слева...
                        Call DrawText(DrawInfo.hDc, str1, Len(str1), DrawInfo.rcItem, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER)
                        ' Отступ от правого края...
                        DrawInfo.rcItem.Right = DrawInfo.rcItem.Right - m_Margin
                        ' Вторую выводим справа...
                        Call DrawText(DrawInfo.hDc, str2, Len(str2) - 1, DrawInfo.rcItem, DT_SINGLELINE Or DT_RIGHT Or DT_VCENTER)
                        ' Возврат первоначального значения...
                        DrawInfo.rcItem.Right = DrawInfo.rcItem.Right + m_Margin
                    Else
                        ' Пишем текст
                        Call DrawText(DrawInfo.hDc, Trim(menuItemString), Len(Trim(menuItemString)) - 1, DrawInfo.rcItem, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER)
                    End If
                        ' Смещаем прямоугольник на единицу
                    Call OffsetRect(DrawInfo.rcItem, -2, -2)
                    Call SetTextColor(DrawInfo.hDc, m_SelectTextMenuColor)
                End If
                ' Пишем текст
                ' Если есть горячая клавиша разбиваем строку на две (разделитель Chr(9))
                If PrsString(Trim(menuItemString), str1, str2) Then
                    ' Первую выводим слева...
                    Call DrawText(DrawInfo.hDc, str1, Len(str1), DrawInfo.rcItem, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER)
                    ' Отступ от правого края...
                    DrawInfo.rcItem.Right = DrawInfo.rcItem.Right - m_Margin
                    ' Ставим цвет HotKey
                    tempTextColor = SetTextColor(DrawInfo.hDc, m_HotKeyTextMenuColor)
                    ' Вторую выводим справа...
                    Call DrawText(DrawInfo.hDc, str2, Len(str2) - 1, DrawInfo.rcItem, DT_SINGLELINE Or DT_RIGHT Or DT_VCENTER)
                    ' Возврат первоначального значения...
                    Call SetTextColor(DrawInfo.hDc, tempTextColor)
                    DrawInfo.rcItem.Right = DrawInfo.rcItem.Right + m_Margin
                Else
                    ' Пишем текст
                    Call DrawText(DrawInfo.hDc, Trim(menuItemString), Len(Trim(menuItemString)) - 1, DrawInfo.rcItem, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER)
                End If
                MsgProc = False
                Exit Function
            End If
            '--- ODS_SELECTED end ---------------------------------------------------------------------------------
            ' &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            '--- ODA_SELECT start ---------------------------------------------------------------------------------
            ' ODS_SELECTED and ODA_SELECT Выход из фокуса элемента меню
            If (Not (DrawInfo.itemState And ODS_SELECTED) And (DrawInfo.itemAction And ODA_SELECT)) Then
                ' Не захватывать зону лейбла
                DrawInfo.rcItem.Left = m_LabelRC.Right
                ' Закрашиваем полосу в зоне заголовка
                FillRectangle DrawInfo.hDc, DrawInfo.rcItem, m_CaptionAreaMenuColor
                 ' Закрашиваем зону чека
                FillRectangle DrawInfo.hDc, checkAreaRC, m_CheckAreaMenuColor
                ' Закрашиваем зону картинки
                FillRectangle DrawInfo.hDc, pictureAreaRC, m_PictureAreaMenuColor
                ' Если меню Checked
                If IsChecked(hwnd, DrawInfo.itemID) Then
                    DrawPicture DrawInfo.hDc, 0, 0, checkAreaRC, m_Form.CheckImage.Picture, DP_COLOR
                End If
                ' Если есть картинка вставляем ее
                DrawPicture DrawInfo.hDc, 0, 0, pictureAreaRC, itemPicture(DrawInfo.itemID), DP_COLOR
                ' Смещаем левый край прямоугольника для вывода текста на ширину картинки и отступы
                DrawInfo.rcItem.Left = checkAreaRC.Right + m_Margin + m_MaxPictureWidth + m_Margin * 2
                ' Если строка меню недоступна
                If IsGrayed(hwnd, DrawInfo.itemID) Then
                    ' Цвет шрифта
                    Call SetTextColor(DrawInfo.hDc, GetSysColor(COLOR_GRAYTEXT))
                Else
                    ' Цвет шрифта
                    Call SetTextColor(DrawInfo.hDc, m_TextMenuColor)
                End If
                ' Если есть горячая клавиша разбиваем строку на две (разделитель Chr(9))
                If PrsString(Trim(menuItemString), str1, str2) Then
                    ' Первую выводим слева...
                    Call DrawText(DrawInfo.hDc, str1, Len(str1), DrawInfo.rcItem, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER)
                    ' Отступ от правого края...
                    DrawInfo.rcItem.Right = DrawInfo.rcItem.Right - m_Margin
                    ' Ставим цвет HotKey
                    tempTextColor = SetTextColor(DrawInfo.hDc, m_HotKeyTextMenuColor)
                    ' Вторую выводим справа...
                    Call DrawText(DrawInfo.hDc, str2, Len(str2) - 1, DrawInfo.rcItem, DT_SINGLELINE Or DT_RIGHT Or DT_VCENTER)
                    ' Возврат первоначального значения...
                    Call SetTextColor(DrawInfo.hDc, tempTextColor)
                    DrawInfo.rcItem.Right = DrawInfo.rcItem.Right + m_Margin
                Else
                    ' Пишем текст
                    Call DrawText(DrawInfo.hDc, Trim(menuItemString), Len(Trim(menuItemString)) - 1, DrawInfo.rcItem, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER)
                End If
                ' Закрашиваем зону Label
                DrawLabel DrawInfo.hDc
                MsgProc = False
                Exit Function
            End If
            '--- ODA_SELECT end   ---------------------------------------------------------------------------------
            Call SelectObject(DrawInfo.hDc, OldFont)
            Call DeleteObject(MyFont)
        End If
        MsgProc = False
        Exit Function
    End If
    ' &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
     If wMsg = WM_MENUSELECT Then
        Debug.Print LoWord(wParam) & " " & HiWord(wParam) ' Номер меню для отладки
         ' Получаем строку
        Call GetMenuString(GetMenu(hwnd), LoWord(wParam), menuItemString, 100, MF_BYCOMMAND)
        ' Если выделяется меню не верхнего уровня
        If HiWord(wParam) <> -32624 And LoWord(wParam) > 0 Then
            ' Выводим картинку на панель подсказок
            Set m_Form.Picture2.Picture = itemPicture(LoWord(wParam))
            ' Выводим наименование выделенной строки меню
            If PrsString(Trim(menuItemString), str1, str2) Then
                m_Form.Label1.Caption = str1 & " HotKey: " & str2
            Else
                m_Form.Label1.Caption = Trim(menuItemString)
            End If
            m_Form.Picture2.Visible = True
         Else
            Set m_Form.Picture2.Picture = LoadPicture()
            m_Form.Label1.Caption = ""
            m_Form.Picture2.Visible = False
        End If
    End If
    
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    MsgProc = CallWindowProc(wlOldProc, hwnd, wMsg, wParam, lParam)
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    
End Function
' Функция деления строки на две строки через разделитель Chr(9). True если строка была разбита...
Public Function PrsString(ByVal soursestring As String, str1 As String, str2 As String) As Boolean
    Dim ls  As Long, ch As Long
    str1 = ""
    str2 = ""
    For ls = 1 To Len(soursestring)
        If Asc(Mid(soursestring, ls, 1)) = 9 Then
            ch = ls
        End If
    Next
    If ch = 0 Then
        str1 = soursestring
        PrsString = False
    Else
        str1 = Mid(soursestring, 1, ch - 1)
        str2 = Mid(soursestring, ch + 1, Len(soursestring))
        PrsString = True
    End If
End Function
' Функция закраски прямоугольника
Public Sub FillRectangle(ByVal hDc As Long, rc As RECT, ColorRC As Long)
    ' Создаем кисть для закраски
    hBr = CreateSolidBrush(ColorRC)
    ' Рисуем приямоугольник и закрашиваем его
    Call FillRect(hDc, rc, hBr)
    ' Удаляем кисть
    Call DeleteObject(hBr)
End Sub
' Функция вывода прозрачной картинки в строку меню
Public Sub DrawPicture(hDc As Long, _
                            popx As Long, _
                            popy As Long, _
                            rcAreaPicture As RECT, _
                            Picture As StdPicture, _
                            flag As Long)
    Dim hBr         As Long
    Dim tempColor   As Long
    Dim X           As Long, Y              As Long
    Dim hDcSource   As Long, hOldMemPicture As Long
    Dim picureRC    As RECT
    
    ' Если нет картинки выходим из процедуры...
    If Picture.Handle = 0 Then Exit Sub
    ' Заполняем прямоугольник картинки
    picureRC.Right = m_Form.ScaleX(Picture.Width)
    picureRC.Bottom = m_Form.ScaleY(Picture.Height)
    
    '// Создаем контекст источник
    hDcSource = CreateCompatibleDC(hDc)
    '// Вставляем в него исходную картинку
    hOldMemPicture = SelectObject(hDcSource, Picture.Handle)
    ' Попиксельно переносим картинку
    For X = 0 To picureRC.Right - 1
        For Y = 0 To picureRC.Bottom - 1
            tempColor = GetPixel(hDcSource, X, Y)
            ' Если цвет текущего пиксела не совпадает с цветом маски
            If m_PictureMaskColor <> tempColor Then
                ' В зависимости от флага переносим картинку по центру rcAreaPicture
                Select Case flag
                    ' Тупо попиксельно переносим цвет
                    Case DP_COLOR
                        Call SetPixel(hDc, popx + X + rcAreaPicture.Left + (rcAreaPicture.Right - rcAreaPicture.Left - picureRC.Right) / 2, popy + Y + rcAreaPicture.Top + (rcAreaPicture.Bottom - rcAreaPicture.Top - picureRC.Bottom) / 2, tempColor)
                    ' Вместо несовпадающих пикселов картинки ставим пиксел с цветом тени
                    Case DP_SHADOW
                        Call SetPixel(hDc, popx + X + rcAreaPicture.Left + (rcAreaPicture.Right - rcAreaPicture.Left - picureRC.Right) / 2, popy + Y + rcAreaPicture.Top + (rcAreaPicture.Bottom - rcAreaPicture.Top - picureRC.Bottom) / 2, m_ShadowColor)
                End Select
            End If
        Next
    Next
    ' Возвращаем картинку
    Call SelectObject(hDcSource, hOldMemPicture)
    ' Удаляем контекст
    Call DeleteDC(hDcSource)

End Sub

