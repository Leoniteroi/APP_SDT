Option Strict On
Option Explicit On

Imports Autodesk.AutoCAD.EditorInput
Imports System

Namespace SDT.Core

    Public NotInheritable Class Log
        Private Sub New()
        End Sub

        Public Shared Sub Info(ed As Editor, message As String)
            If ed Is Nothing Then Exit Sub
            ed.WriteMessage(Environment.NewLine & "[SDT] " & message)
        End Sub

        Public Shared Sub Warn(ed As Editor, message As String)
            If ed Is Nothing Then Exit Sub
            ed.WriteMessage(Environment.NewLine & "[SDT][WARN] " & message)
        End Sub
    End Class

End Namespace
