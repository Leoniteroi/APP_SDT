Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq

Namespace SDT.Civil

    ''' <summary>
    ''' Repositório simples de códigos em arquivo .ini (seção [DEFAULT], chave Codes=...).
    ''' Observação: grava em pasta do usuário (AppData) para evitar falta de permissão em Program Files.
    ''' </summary>
    Public NotInheritable Class IniCodesRepository

        Private Sub New()
        End Sub

        Private Shared Function GetIniFolder() As String
            Dim appData As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim folder As String = Path.Combine(appData, "SDT", "C3D2024", "ini")
            Directory.CreateDirectory(folder)
            Return folder
        End Function

        Public Shared Function GetIniPath(iniFileName As String) As String
            If String.IsNullOrWhiteSpace(iniFileName) Then Throw New ArgumentException("iniFileName vazio.")
            Dim safeName As String = iniFileName.Trim()

            'garante extensão .ini
            If Not safeName.EndsWith(".ini", StringComparison.OrdinalIgnoreCase) Then
                safeName &= ".ini"
            End If

            Return Path.Combine(GetIniFolder(), safeName)
        End Function

        Public Shared Sub EnsureExists(iniFileName As String, defaultCodes As IEnumerable(Of String))
            Dim iniPath As String = GetIniPath(iniFileName)
            If File.Exists(iniPath) Then Exit Sub

            SaveDefaultCodes(iniPath, defaultCodes)
        End Sub

        Public Shared Function LoadDefaultCodes(iniFileName As String, defaultCodes As IEnumerable(Of String)) As List(Of String)
            EnsureExists(iniFileName, defaultCodes)

            Dim iniPath As String = GetIniPath(iniFileName)
            Dim lines As String() = File.ReadAllLines(iniPath)

            Dim inDefaultSection As Boolean = False
            Dim codesLine As String = ""

            For Each raw As String In lines
                Dim line As String = raw.Trim()

                If line.Length = 0 OrElse line.StartsWith(";"c) OrElse line.StartsWith("#"c) Then
                    Continue For
                End If

                If line.StartsWith("["c) AndAlso line.EndsWith("]"c) Then
                    Dim sec As String = line.Substring(1, line.Length - 2).Trim()
                    inDefaultSection = sec.Equals("DEFAULT", StringComparison.OrdinalIgnoreCase)
                    Continue For
                End If

                If inDefaultSection AndAlso line.StartsWith("Codes=", StringComparison.OrdinalIgnoreCase) Then
                    codesLine = line.Substring("Codes=".Length)
                    Exit For
                End If
            Next

            Return ParseCodes(codesLine)
        End Function

        Private Shared Function ParseCodes(codesLine As String) As List(Of String)
            Dim list As New List(Of String)

            If String.IsNullOrWhiteSpace(codesLine) Then Return list

            For Each part As String In codesLine.Split(","c)
                Dim code As String = part.Trim()
                If code.Length > 0 AndAlso Not list.Contains(code) Then
                    list.Add(code)
                End If
            Next

            Return list
        End Function

        Public Shared Sub SaveDefaultCodes(iniPath As String, codes As IEnumerable(Of String))
            Dim safeList As New List(Of String)(If(codes, Enumerable.Empty(Of String)())) 'sem Object/late binding

            Dim text As String =
                "[DEFAULT]" & Environment.NewLine &
                "Codes=" & String.Join(",", safeList) & Environment.NewLine

            File.WriteAllText(iniPath, text)
        End Sub

    End Class

End Namespace
