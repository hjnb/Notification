Imports System.Reflection
Imports System.Runtime.InteropServices

Public Class Util

    '和暦の記号
    Private Const ERA_MEIJI As String = "M" '明治
    Private Const ERA_TAISYO As String = "T" '大正
    Private Const ERA_SYOWA As String = "S" '昭和
    Private Const ERA_HEISEI As String = "H" '平成
    Private Const ERA_X As String = "R" '令和
    Private Const ERA_X_KANJI As String = "令和"

    ''' <summary>
    ''' コントロールのDoubleBufferedプロパティをTrueにする
    ''' </summary>
    ''' <param name="control">対象のコントロール</param>
    Public Shared Sub EnableDoubleBuffering(control As Control)
        control.GetType().InvokeMember("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, Nothing, control, New Object() {True})
    End Sub

    ''' <summary>
    ''' dgvのセルの値がNullかチェック、Nullの場合空文字を返す
    ''' </summary>
    ''' <param name="dgvCellValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function checkDBNullValue(dgvCellValue As Object) As String
        Return If(IsDBNull(dgvCellValue), "", dgvCellValue)
    End Function

    ''' <summary>
    ''' 西暦(yyyy/MM/dd)を和暦に変換
    ''' </summary>
    ''' <param name="adStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function convADStrToWarekiStr(adStr As String) As String
        If System.Text.RegularExpressions.Regex.IsMatch(adStr, "[12]\d\d\d/\d\d/\d\d") Then
            Dim yearStr As String = adStr.Substring(0, 4)
            Dim monthStr As String = adStr.Substring(5, 2)
            Dim dateStr As String = adStr.Substring(8, 2)

            Dim yearNum As Integer = CInt(yearStr)
            Dim monthNum As Integer = CInt(monthStr)
            Dim dateNum As Integer = CInt(dateStr)

            Dim convEraStr As String = ""
            Dim convYearNum As Integer
            Dim convYearStr As String = ""

            '西暦から和暦への変換処理
            If yearNum >= 2118 Then
                convEraStr = ERA_X
                convYearNum = 99
            ElseIf yearNum >= 2020 Then
                'X２年～
                convEraStr = ERA_X
                convYearNum = yearNum - 2018
            ElseIf yearNum = 2019 Then
                '平成３１年orX1年
                If monthNum <= 4 Then
                    convEraStr = ERA_HEISEI
                    convYearNum = 31
                Else
                    convEraStr = ERA_X
                    convYearNum = 1
                End If
            ElseIf yearNum >= 1990 Then
                '平成２年～３０年
                convEraStr = ERA_HEISEI
                convYearNum = yearNum - 1988
            ElseIf yearNum = 1989 Then
                '平成１年or昭和64年
                If monthNum = 1 AndAlso dateNum <= 7 Then
                    convEraStr = ERA_SYOWA
                    convYearNum = 64
                Else
                    convEraStr = ERA_HEISEI
                    convYearNum = 1
                End If
            ElseIf yearNum >= 1927 Then
                '昭和２年～６３年
                convEraStr = ERA_SYOWA
                convYearNum = yearNum - 1925
            ElseIf yearNum = 1926 Then
                '昭和１年or大正１５年
                If monthNum = 12 AndAlso dateNum >= 25 Then
                    convEraStr = ERA_SYOWA
                    convYearNum = 1
                Else
                    convEraStr = ERA_TAISYO
                    convYearNum = 15
                End If
            ElseIf yearNum >= 1913 Then
                '大正２年～１４年
                convEraStr = ERA_TAISYO
                convYearNum = yearNum - 1911
            ElseIf yearNum = 1912 Then
                '大正１年 or 明治４５年
                If monthNum >= 8 OrElse (monthNum = 7 AndAlso dateNum >= 30) Then
                    convEraStr = ERA_TAISYO
                    convYearNum = 1
                Else
                    convEraStr = ERA_MEIJI
                    convYearNum = 45
                End If
            ElseIf yearNum >= 1900 Then
                '明治３３年～４４年
                convEraStr = ERA_MEIJI
                convYearNum = yearNum - 1867
            ElseIf yearNum < 1900 Then
                '1899年以前は空を返す
                Return ""
            End If

            convYearStr = If(convYearNum < 10, "0" & convYearNum, "" & convYearNum)

            Return convEraStr & convYearStr & "/" & monthStr & "/" & dateStr
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' 和暦を西暦に変換
    ''' </summary>
    ''' <param name="warekiStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function convWarekiStrToADStr(warekiStr As String) As String
        Dim warekiPattenStr As String = ERA_X & ERA_HEISEI & ERA_SYOWA & ERA_TAISYO & ERA_MEIJI
        If System.Text.RegularExpressions.Regex.IsMatch(warekiStr, "[" & warekiPattenStr & "]\d\d/\d\d/\d\d") Then
            Dim eraStr As String = warekiStr.Substring(0, 1)
            Dim yearNum As Integer = CInt(warekiStr.Substring(1, 2))
            Dim adYear As Integer
            If eraStr = ERA_X Then
                adYear = 2018 + yearNum
            ElseIf eraStr = ERA_HEISEI Then
                adYear = 1988 + yearNum
            ElseIf eraStr = ERA_SYOWA Then
                adYear = 1925 + yearNum
            ElseIf eraStr = ERA_TAISYO Then
                adYear = 1911 + yearNum
            ElseIf eraStr = ERA_MEIJI Then
                adYear = 1867 + yearNum
            End If

            Return adYear & warekiStr.Substring(3, 6)
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' 和暦の漢字を取得
    ''' </summary>
    ''' <param name="warekiStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getKanji(warekiStr As String) As String
        Dim warekiPattenStr As String = ERA_X & ERA_HEISEI & ERA_SYOWA & ERA_TAISYO & ERA_MEIJI
        If System.Text.RegularExpressions.Regex.IsMatch(warekiStr, "[" & warekiPattenStr & "]\d\d/\d\d/\d\d") Then
            Dim eraStr As String = warekiStr.Substring(0, 1)
            If eraStr = ERA_X Then
                Return ERA_X_KANJI
            ElseIf eraStr = "H" Then
                Return "平成"
            ElseIf eraStr = "S" Then
                Return "昭和"
            ElseIf eraStr = "T" Then
                Return "大正"
            ElseIf eraStr = "M" Then
                Return "明治"
            Else
                Return ""
            End If
        Else
            Return ""
        End If
    End Function

    <DllImport("KERNEL32.DLL", CharSet:=CharSet.Auto)>
    Public Shared Function GetPrivateProfileString(
        ByVal lpAppName As String,
        ByVal lpKeyName As String, ByVal lpDefault As String,
        ByVal lpReturnedString As System.Text.StringBuilder, ByVal nSize As Integer,
        ByVal lpFileName As String) As Integer
    End Function

    <DllImport("KERNEL32.DLL", CharSet:=CharSet.Auto)>
    Public Shared Function WritePrivateProfileString(
        ByVal lpApplicationName As String,
        ByVal lpKeyName As String,
        ByVal lpString As String,
        ByVal lpFileName As String) As Long
    End Function

    ''' <summary>
    ''' iniファイルから読み込み
    ''' </summary>
    ''' <param name="lpSection">セクション名</param>
    ''' <param name="lpKeyName">読み込むkey名</param>
    ''' <param name="lpFileName">iniファイルパス</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getIniString(ByVal lpSection As String, ByVal lpKeyName As String, ByVal lpFileName As String) As String
        Dim strValue As System.Text.StringBuilder = New System.Text.StringBuilder(1024)

        Dim sLen = GetPrivateProfileString(lpSection, lpKeyName, "", strValue, 1024, lpFileName)
        Dim str As String = strValue.ToString()

        Return str
    End Function

    ''' <summary>
    ''' iniファイルへ書き込み
    ''' </summary>
    ''' <param name="lpSection">セクション名</param>
    ''' <param name="lpKeyName">書き込む対象のkey名</param>
    ''' <param name="lpValue">書き込む値</param>
    ''' <param name="lpFileName">iniファイルパス</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function putIniString(ByVal lpSection As String, lpKeyName As String, ByVal lpValue As String, ByVal lpFileName As String) As Boolean
        If Not System.IO.File.Exists(lpFileName) Then
            Return False
        End If
        Dim result As Long = WritePrivateProfileString(lpSection, lpKeyName, lpValue, lpFileName)
        Return result <> 0
    End Function

    ''' <summary>
    ''' 現在年齢算出
    ''' </summary>
    ''' <param name="birthYmd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function calcAge(birthYmd As String, nowYmd As String) As Integer
        Dim doDate As DateTime = New DateTime(CInt(nowYmd.Split("/")(0)), CInt(nowYmd.Split("/")(1)), CInt(nowYmd.Split("/")(2)))
        Dim birthDate As DateTime = New DateTime(CInt(birthYmd.Split("/")(0)), CInt(birthYmd.Split("/")(1)), CInt(birthYmd.Split("/")(2)))
        Dim age As Integer = doDate.Year - birthDate.Year
        '誕生日がまだ来ていなければ、1引く
        If doDate.Month < birthDate.Month OrElse (doDate.Month = birthDate.Month AndAlso doDate.Day < birthDate.Day) Then
            age -= 1
        End If
        Return age
    End Function
End Class


