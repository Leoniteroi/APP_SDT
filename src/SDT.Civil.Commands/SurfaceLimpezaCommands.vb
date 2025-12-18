Option Strict On
Option Explicit On

Imports System
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports SDT.Core
Imports SDT.Civil

''' <summary>
''' SDT_CRIAR_SUPERFICIES_LIMPEZA
''' </summary>
Public Class SurfaceLimpezaCommands

    <CommandMethod("SDT_CRIAR_SUPERFICIES_LIMPEZA", CommandFlags.Modal)>
    Public Sub SDT_CRIAR_SUPERFICIES_LIMPEZA()

        Dim ctx As AcadContext = Nothing
        Dim err As String = ""

        If Not AcadContext.TryCreate(ctx, err) Then
            Return
        End If

        Dim tnId As ObjectId = PromptSelectTinSurface(ctx.Ed, ctx.CivDoc)
        If tnId.IsNull Then
            ctx.Ed.WriteMessage(Environment.NewLine & "[SDT] Nenhuma superfície selecionada.")
            Return
        End If

        Dim espessuraCm As Double = PromptEspessura(ctx.Ed)
        If espessuraCm <= 0 Then
            ctx.Ed.WriteMessage(Environment.NewLine & "[SDT] Espessura inválida.")
            Return
        End If



        TransactionRunner.RunWrite(
            ctx.Db,
                Sub(tr As Transaction)

                    Dim outerBoundaryId As ObjectId = ObjectId.Null ' por enquanto (ou implemente prompt)
                    Dim limpezaSurfaceName As String =
            "LIMPEZA_" & GetTinName(tr, tnId) & "_" & espessuraCm.ToString("0.##") & "cm"

                    SurfaceLimpezaService.Run(tr, ctx.Db, ctx.CivDoc, ctx.Ed, tnId, espessuraCm, outerBoundaryId, limpezaSurfaceName)
                End Sub)

    End Sub

    Private Function GetTinName(tr As Transaction, tinId As ObjectId) As String
        Dim s As Autodesk.Civil.DatabaseServices.TinSurface =
        TryCast(tr.GetObject(tinId, OpenMode.ForRead), Autodesk.Civil.DatabaseServices.TinSurface)

        If s Is Nothing OrElse String.IsNullOrWhiteSpace(s.Name) Then Return "TN"
        Return s.Name.Trim()
    End Function



    Private Function PromptEspessura(ed As Editor) As Double
        Dim pdo As New PromptDoubleOptions(Environment.NewLine & "Espessura da limpeza (cm): ")
        pdo.AllowNegative = False
        pdo.AllowZero = False
        pdo.DefaultValue = 20.0

        Dim pdr As PromptDoubleResult = ed.GetDouble(pdo)
        If pdr.Status <> PromptStatus.OK Then Return -1
        Return pdr.Value
    End Function

    Private Function PromptSelectTinSurface(ed As Editor,
                                           civDoc As CivilDocument) As ObjectId
        Dim surfaces As ObjectIdCollection = civDoc.GetSurfaceIds()
        If surfaces Is Nothing OrElse surfaces.Count = 0 Then
            ed.WriteMessage(Environment.NewLine & "[SDT] Não há superfícies no desenho.")
            Return ObjectId.Null
        End If

        Using tr As Transaction = ed.Document.Database.TransactionManager.StartTransaction()

            ed.WriteMessage(Environment.NewLine & "[SDT] Superfícies disponíveis (TIN):")
            For Each id As ObjectId In surfaces
                Dim s As Autodesk.Civil.DatabaseServices.TinSurface =
                    TryCast(tr.GetObject(id, OpenMode.ForRead), Autodesk.Civil.DatabaseServices.TinSurface)
                If s IsNot Nothing Then
                    ed.WriteMessage(Environment.NewLine & "  - " & s.Name)
                End If
            Next

            tr.Commit()
        End Using

        Dim pso As New PromptStringOptions(Environment.NewLine & "Digite o nome da superfície TN (exato): ")
        pso.AllowSpaces = True

        Dim psr As PromptResult = ed.GetString(pso)
        If psr.Status <> PromptStatus.OK OrElse String.IsNullOrWhiteSpace(psr.StringResult) Then
            Return ObjectId.Null
        End If

        Dim targetName As String = psr.StringResult.Trim()

        Using tr As Transaction = ed.Document.Database.TransactionManager.StartTransaction()
            For Each id As ObjectId In surfaces
                Dim s As Autodesk.Civil.DatabaseServices.TinSurface =
                    TryCast(tr.GetObject(id, OpenMode.ForRead), Autodesk.Civil.DatabaseServices.TinSurface)

                If s IsNot Nothing AndAlso
                   s.Name.Equals(targetName, StringComparison.OrdinalIgnoreCase) Then
                    Return id
                End If
            Next
        End Using

        Return ObjectId.Null
    End Function

End Class
