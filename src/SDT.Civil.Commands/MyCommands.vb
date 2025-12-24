Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports SDT.Civil
Imports SDT.Core

''' <summary>
''' Comandos SDT: criação de CorridorSurfaces (DATUM/TOP) e edição de link codes via INI.
''' Target: Civil 3D 2024.
''' </summary>
Public Class MyCommands

#Region "DATUM"

    <CommandMethod("SDT_CRIAR_SUPERFICIE_DATUM", CommandFlags.Modal)>
    Public Sub CriarDatum()

        Dim ctx As AcadContext = Nothing
        Dim err As String = Nothing

        If Not AcadContext.TryCreate(ctx, err) Then
            Return
        End If

        Dim spec As New CorridorSurfaceBuilderService.SurfaceSpec With {
            .Suffix = "_DATUM",
            .IniFile = "SDT_DatumSurfaces.ini",
            .DefaultCodes = New String() {"Datum", "DATUM"},
            .OverhangMode = "BottomLinks"
        }

        TransactionRunner.RunWrite(
            ctx.Db,
            Sub(tr As Transaction)
                CorridorSurfaceBuilderService.BuildAll(tr, ctx.CivDoc, ctx.Ed, spec)
            End Sub)

        ctx.Ed.WriteMessage(Environment.NewLine & "[SDT] DATUM concluído.")
    End Sub

    <CommandMethod("SDT_EDITAR_CODES_DATUM", CommandFlags.Modal)>
    Public Sub EditarCodesDatum()
        EditarCodesIni("SDT_DatumSurfaces.ini", New String() {"Datum", "DATUM"}, "DATUM")
    End Sub

#End Region

#Region "TOP"

    <CommandMethod("SDT_CRIAR_SUPERFICIE_TOP", CommandFlags.Modal)>
    Public Sub CriarTop()

        Dim ctx As AcadContext = Nothing
        Dim err As String = Nothing

        If Not AcadContext.TryCreate(ctx, err) Then
            Return
        End If

        Dim spec As New CorridorSurfaceBuilderService.SurfaceSpec With {
            .Suffix = "_TOP",
            .IniFile = "SDT_TopSurfaces.ini",
            .DefaultCodes = New String() {"Top", "TOP"},
            .OverhangMode = "TopLinks"
        }

        TransactionRunner.RunWrite(
            ctx.Db,
            Sub(tr As Transaction)
                CorridorSurfaceBuilderService.BuildAll(tr, ctx.CivDoc, ctx.Ed, spec)
            End Sub)

        ctx.Ed.WriteMessage(Environment.NewLine & "[SDT] TOP concluído.")
    End Sub

    <CommandMethod("SDT_EDITAR_CODES_TOP", CommandFlags.Modal)>
    Public Sub EditarCodesTop()
        EditarCodesIni("SDT_TopSurfaces.ini", New String() {"Top", "TOP"}, "TOP")
    End Sub

#End Region

#Region "INI Editor"

    Private Sub EditarCodesIni(iniFileName As String,
                              defaultCodes As String(),
                              label As String)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        If doc Is Nothing Then Return

        Dim ed As Editor = doc.Editor

        Dim atuais As List(Of String) =
            IniCodesRepository.LoadDefaultCodes(iniFileName, defaultCodes)

        ed.WriteMessage(Environment.NewLine & "==============================")
        ed.WriteMessage(Environment.NewLine & "[SDT] Codes " & label & " DEFAULT atuais:")
        ed.WriteMessage(Environment.NewLine & If(atuais.Count = 0, "(vazio)", String.Join(", ", atuais)))
        ed.WriteMessage(Environment.NewLine & "==============================")

        Dim pso As New PromptStringOptions(
            Environment.NewLine & "Nova lista (separar por vírgula) (Enter = manter): ")
        pso.AllowSpaces = True

        Dim psr As PromptResult = ed.GetString(pso)
        If psr.Status <> PromptStatus.OK OrElse
           String.IsNullOrWhiteSpace(psr.StringResult) Then

            ed.WriteMessage(Environment.NewLine & "[SDT] Nenhuma alteração.")
            Return
        End If

        Dim novos As New List(Of String)
        For Each c As String In psr.StringResult.Split(","c)
            Dim code As String = c.Trim()
            If code <> "" AndAlso Not novos.Contains(code) Then
                novos.Add(code)
            End If
        Next

        If novos.Count = 0 Then
            ed.WriteMessage(Environment.NewLine & "[SDT] Lista inválida.")
            Return
        End If

        Dim iniPath As String = IniCodesRepository.GetIniPath(iniFileName)
        IniCodesRepository.SaveDefaultCodes(iniPath, novos)

        ed.WriteMessage(Environment.NewLine & "[SDT] Codes " & label & " DEFAULT atualizados:")
        ed.WriteMessage(Environment.NewLine & String.Join(", ", novos))

    End Sub

#End Region

End Class
