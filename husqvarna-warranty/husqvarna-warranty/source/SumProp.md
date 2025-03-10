' ==========================================================================
' Module Name  : SumProp
' Description  : This module was originally discovered on the internet
'               around 1998, authored by an unknown developer for use in
'               Microsoft Excel. It has been modified and adapted for
'               reports in Microsoft Access.
'
'               The module contains a series of subroutines (both draft
'               and working versions) that convert numeric values
'               (including millions) into words. The adaptation was specifically
'               designed for Ukrainian financial documentation, ensuring that
'               the number-to-text conversion follows the Ukrainian language
'               rules.
'
'               The original author used Russian variable names, which have
'               been preserved in the source code to maintain the module's
'               original structure and functionality.
'
' Developer    : unknown
' Created      : approximately 1998-mm-dd
' Last Updated : 2025-02-19 by Oleh Bondarenko - Added comments for GitHub upload
' ==========================================================================


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         СУММА ПРОПИСЬЮ                       '
'                                                              '
'       Данный модуль определяет две полезные функции:         '
'          Курс (Дата) извлекает из таблицы "Курс" значение    '
'                      курса обмена рубля на доллары США.      '
'          СуммаПрописью (Сумма) выводит сумму прописью для    '
'                      печати в счетах, накладных и пр.        '
'                                                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database   'Использовать функции базы данных при сравнении строк
Option Explicit

Global Сумма As Single, Остаток As Single

Function Десятки(Разряд As Long) As String

    Select Case Разряд
         Case 2
            Десятки = "двадцять "
         Case 3
            Десятки = "тридцять "
         Case 4
            Десятки = "сорок "
         Case 5
            Десятки = "пятдесят "
         Case 6
            Десятки = "шістдесят "
         Case 7
            Десятки = "сімдесят "
         Case 8
            Десятки = "вісімдесят "
         Case 9
            Десятки = "дев'яносто "
    End Select

End Function

Function Единицы(Разряд As Long, Род As String) As String

    Select Case Разряд
        Case 1
            If Род = "Мужской" Then
                Единицы = "одна "
            Else
                Единицы = "одна "
            End If
        Case 2
            If Род = "Мужской" Then
                Единицы = "дві "
            Else
                Единицы = "дві "
            End If
        Case 3
            Единицы = "три "
        Case 4
            Единицы = "чотири "
        Case 5
            Единицы = "п'ять "
        Case 6
            Единицы = "шість "
        Case 7
            Единицы = "сім "
        Case 8
            Единицы = "вісім "
        Case 9
            Единицы = "дев'ять "
        Case 10
            Единицы = "десять "
        Case 11
            Единицы = "одинадцять "
        Case 12
            Единицы = "дванадцять "
        Case 13
            Единицы = "тринадцять "
        Case 14
            Единицы = "чотирнадцять "
        Case 15
            Единицы = "п'ятнадцять "
        Case 16
            Единицы = "шістнадцять "
        Case 17
            Единицы = "сімнадцять "
        Case 18
            Единицы = "вісімнадцять "
        Case 19
            Единицы = "дев'ятнадцять "

    End Select

End Function

Function Миллионы(Разряд As Long) As String

    If Разряд = 1 Then
        Миллионы = "мільйон "
    ElseIf Разряд > 1 And Разряд < 5 Then
        Миллионы = "мільйона "
    Else
        Миллионы = "мільйонів "
    End If

End Function

Function Сотни(Разряд As Long) As String
    
    Select Case Разряд
         Case 1
            Сотни = "сто "
         Case 2
            Сотни = "двісті "
         Case 3
            Сотни = "триста "
         Case 4
            Сотни = "чотириста "
         Case 5
            Сотни = "п'ятсот "
         Case 6
            Сотни = "шістсот "
         Case 7
            Сотни = "сімсот "
         Case 8
            Сотни = "вісімсот "
         Case 9
            Сотни = "дев'ятсот "
    End Select

End Function

Function fnSP(СуммаСчета As Currency) As String
' Параметры:  Используются глобальные параметры
'             Сумма, Остаток и Подпись
' Назначение: Перевод СуммыСчета в строковую константу
' Возвращает: СуммуПрописью

Dim Группа As Single, Разряд As Long, Длина, Копейки As Integer
Dim Пропись As String

    Сумма = Round(СуммаСчета, 2)
    Копейки = Сумма * 100 - (Сумма * 100 \ 100) * 100
    Остаток = Сумма - Копейки / 100
    
    Группа = Остаток \ 1000000
    If Группа <> 0 Then
        Разряд = Группа \ 100
        Пропись = Пропись & Сотни(Разряд)
        Остаток = Остаток - Разряд * 100 * 1000000
        Группа = Группа - Разряд * 100

        If Группа > 19 Then
            Разряд = Группа \ 10
            Пропись = Пропись & Десятки(Разряд)
            Остаток = Остаток - Разряд * 10 * 1000000
            Группа = Группа - Разряд * 10
        End If

        Разряд = Группа
        Пропись = Пропись & Единицы(Разряд, "Мужской")
        Остаток = Остаток - Разряд * 1000000

        Пропись = Пропись & Миллионы(Разряд)
    End If

    Группа = Остаток \ 1000
    If Группа <> 0 Then
        Разряд = Группа \ 100
        Пропись = Пропись & Сотни(Разряд)
        Остаток = Остаток - Разряд * 100 * 1000
        Группа = Группа - Разряд * 100

        If Группа > 19 Then
            Разряд = Группа \ 10
            Пропись = Пропись & Десятки(Разряд)
            Остаток = Остаток - Разряд * 10 * 1000
            Группа = Группа - Разряд * 10
        End If

        Разряд = Группа
        Пропись = Пропись & Единицы(Разряд, "Женский")
        Остаток = Остаток - Разряд * 1000

        Пропись = Пропись & Тысячи(Разряд)
    End If

    Группа = Остаток

    If Группа <> 0 Then
        Разряд = Группа \ 100
        Пропись = Пропись & Сотни(Разряд)
        Остаток = Остаток - Разряд * 100
        Группа = Группа - Разряд * 100

        If Группа > 19 Then
            Разряд = Группа \ 10
            Пропись = Пропись & Десятки(Разряд)
            Остаток = Остаток - Разряд * 10
            Группа = Группа - Разряд * 10
        End If

        Разряд = Группа
        Пропись = Пропись & Единицы(Разряд, "Мужской")
        Остаток = Остаток - Разряд

    End If
    
    Пропись = Пропись & "грн. " & Копейки & " коп."
    
    Длина = Len(Пропись)
    If IsNull(Длина) Then
       Exit Function
    End If

    Пропись = UCase(Mid(Пропись, 1, 1)) & (Mid(Пропись, 2, Длина))
 
    fnSP = Пропись

End Function
    

Function Тысячи(Разряд As Long) As String

    If Разряд = 1 Then
        Тысячи = "тисяча "
    ElseIf Разряд > 1 And Разряд < 5 Then
        Тысячи = "тисячі "
    Else
        Тысячи = "тисяч "
    End If

End Function
Function ЧислоПрописью(СуммаСчета As Currency) As String
' Параметры:  Используются глобальные параметры
'             Сумма, Остаток и Подпись
' Назначение: Перевод СуммыСчета в строковую константу
' Возвращает: СуммуПрописью

Dim Группа As Single, Разряд As Long, Длина, Копейки As Integer
Dim Пропись As String

    Сумма = СуммаСчета
    Копейки = Сумма * 100 - (Сумма * 100 \ 100) * 100
    Остаток = Сумма - Копейки / 100
    
    Группа = Остаток \ 1000000
    If Группа <> 0 Then
        Разряд = Группа \ 100
        Пропись = Пропись & Сотни(Разряд)
        Остаток = Остаток - Разряд * 100 * 1000000
        Группа = Группа - Разряд * 100

        If Группа > 19 Then
            Разряд = Группа \ 10
            Пропись = Пропись & Десятки(Разряд)
            Остаток = Остаток - Разряд * 10 * 1000000
            Группа = Группа - Разряд * 10
        End If

        Разряд = Группа
        Пропись = Пропись & Единицы(Разряд, "Мужской")
        Остаток = Остаток - Разряд * 1000000

        Пропись = Пропись & Миллионы(Разряд)
    End If

    Группа = Остаток \ 1000
    If Группа <> 0 Then
        Разряд = Группа \ 100
        Пропись = Пропись & Сотни(Разряд)
        Остаток = Остаток - Разряд * 100 * 1000
        Группа = Группа - Разряд * 100

        If Группа > 19 Then
            Разряд = Группа \ 10
            Пропись = Пропись & Десятки(Разряд)
            Остаток = Остаток - Разряд * 10 * 1000
            Группа = Группа - Разряд * 10
        End If

        Разряд = Группа
        Пропись = Пропись & Единицы(Разряд, "Женский")
        Остаток = Остаток - Разряд * 1000

        Пропись = Пропись & Тысячи(Разряд)
    End If

    Группа = Остаток

    If Группа <> 0 Then
        Разряд = Группа \ 100
        Пропись = Пропись & Сотни(Разряд)
        Остаток = Остаток - Разряд * 100
        Группа = Группа - Разряд * 100

        If Группа > 19 Then
            Разряд = Группа \ 10
            Пропись = Пропись & Десятки(Разряд)
            Остаток = Остаток - Разряд * 10
            Группа = Группа - Разряд * 10
        End If

        Разряд = Группа
        Пропись = Пропись & Единицы(Разряд, "Мужской")
        Остаток = Остаток - Разряд

    End If
    
    Пропись = Пропись
    
    Длина = Len(Пропись)
    If IsNull(Длина) Then
       Exit Function
    End If

    'Пропись = UCase(Mid(Пропись, 1, 1)) & (Mid(Пропись, 2, Длина))
 
    ЧислоПрописью = Пропись & ""

End Function


Function fnSP2(СуммаСчета As Currency) As String
' Параметры:  Используются глобальные параметры
'             Сумма, Остаток и Подпись
' Назначение: Перевод СуммыСчета в строковую константу
' Возвращает: СуммуПрописью

Dim Группа As Single, Разряд As Long, Длина, Копейки As Integer
Dim Пропись As String

    Сумма = Round(СуммаСчета, 2)
    Копейки = Сумма * 100 - (Сумма * 100 \ 100) * 100
    Остаток = Сумма - Копейки / 100
    
    Группа = Остаток \ 1000000
    If Группа <> 0 Then
        Разряд = Группа \ 100
        Пропись = Пропись & Сотни(Разряд)
        Остаток = Остаток - Разряд * 100 * 1000000
        Группа = Группа - Разряд * 100

        If Группа > 19 Then
            Разряд = Группа \ 10
            Пропись = Пропись & Десятки(Разряд)
            Остаток = Остаток - Разряд * 10 * 1000000
            Группа = Группа - Разряд * 10
        End If

        Разряд = Группа
        Пропись = Пропись & Единицы(Разряд, "Мужской")
        Остаток = Остаток - Разряд * 1000000

        Пропись = Пропись & Миллионы(Разряд)
    End If

    Группа = Остаток \ 1000
    If Группа <> 0 Then
        Разряд = Группа \ 100
        Пропись = Пропись & Сотни(Разряд)
        Остаток = Остаток - Разряд * 100 * 1000
        Группа = Группа - Разряд * 100

        If Группа > 19 Then
            Разряд = Группа \ 10
            Пропись = Пропись & Десятки(Разряд)
            Остаток = Остаток - Разряд * 10 * 1000
            Группа = Группа - Разряд * 10
        End If

        Разряд = Группа
        Пропись = Пропись & Единицы(Разряд, "Женский")
        Остаток = Остаток - Разряд * 1000

        Пропись = Пропись & Тысячи(Разряд)
    End If

    Группа = Остаток

    If Группа <> 0 Then
        Разряд = Группа \ 100
        Пропись = Пропись & Сотни(Разряд)
        Остаток = Остаток - Разряд * 100
        Группа = Группа - Разряд * 100

        If Группа > 19 Then
            Разряд = Группа \ 10
            Пропись = Пропись & Десятки(Разряд)
            Остаток = Остаток - Разряд * 10
            Группа = Группа - Разряд * 10
        End If

        Разряд = Группа
        Пропись = Пропись & Единицы(Разряд, "Мужской")
        Остаток = Остаток - Разряд

    End If
    
    Пропись = Пропись
    
    Длина = Len(Пропись)
    If IsNull(Длина) Then
       Exit Function
    End If

    Пропись = UCase(Mid(Пропись, 1, 1)) & (Mid(Пропись, 2, Длина))
 
    fnSP2 = Пропись

End Function
