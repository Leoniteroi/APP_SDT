Option Strict On
Option Explicit On

Imports System
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices

Namespace SDT.Civil

    ''' <summary>
    ''' Execução em lote:
    ''' 1) recebe TN base (TinSurface) + espessura
    ''' 2) para cada corredor:
    '''    - procura surface "*_DATUM_OUTER"
    '''    - extrai o(s) OuterBoundary(ies) e gera Polyline(s) fechadas 2D temporárias
    '''    - cria a superfície "{Corridor}_LIMPEZA" clonando TN base (TinSurfaceService.CloneTinSurface)
    '''    - aplica raise (espessura) e boundary
    ''' </summary>
    Public NotInheritable Class SurfaceCFTBatchService
        Private Sub New()
        End Sub

        Public Shared Function RunAllCorridors(tr As Transaction,
                                              db As Database,
                                              civDoc As CivilDocument,
                                              ed As Editor,
                                              espessuraCm_CFTcorte As Double,
                                              espessuraCm_CFTaterro As Double,
                                              cft_aterro_Suffix As String,
                                              cft_corte_Suffix As String,
                                              subbase_Suffix As String) As Integer

            If tr Is Nothing OrElse db Is Nothing OrElse civDoc Is Nothing OrElse ed Is Nothing Then Return 0
            If String.IsNullOrWhiteSpace(cft_aterro_Suffix) Then cft_aterro_Suffix = "_ATERRO_CFT"
            If String.IsNullOrWhiteSpace(cft_corte_Suffix) Then cft_corte_Suffix = "_CORTE_CFT"
            If String.IsNullOrWhiteSpace(subbase_Suffix) Then subbase_Suffix = "_SUBBASE"

            Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim ms As BlockTableRecord = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

            Dim created As Integer = 0

            For Each corrId As ObjectId In civDoc.CorridorCollection
                If corrId.IsNull Then Continue For

                Dim corr As Corridor = TryCast(tr.GetObject(corrId, OpenMode.ForRead), Corridor)
                If corr Is Nothing Then Continue For

                ' localizar surface "*_SUBBASE"
                Dim subbaseSurf As ObjectId = TinSurfaceService.FindTinSurfaceIdByName(tr, civDoc, corr.Name & subbase_Suffix)
                If subbaseSurf.IsNull Then
                    ed.WriteMessage(Environment.NewLine & $"[SDT] Corredor '{corr.Name}': sem surface '*{subbase_Suffix}'.")
                    Continue For
                End If

                ' motor: clona TN subbase + raise
                Dim aterrocftId As ObjectId = TinSurfaceService.CloneTinSurface(db, tr, civDoc, subbaseSurf, corr.Name & cft_aterro_Suffix, ed)
                TinSurfaceService.RaiseTinSurface(TryCast(tr.GetObject(aterrocftId, OpenMode.ForWrite), TinSurface), espessuraCm_CFTaterro / 100)

                Dim cortecftId As ObjectId = TinSurfaceService.CloneTinSurface(db, tr, civDoc, subbaseSurf, corr.Name & cft_corte_Suffix, ed)
                TinSurfaceService.RaiseTinSurface(TryCast(tr.GetObject(cortecftId, OpenMode.ForWrite), TinSurface), espessuraCm_CFTcorte / 100)

                If aterrocftId.IsNull Then
                    ed.WriteMessage(Environment.NewLine & $"[SDT] Corredor '{corr.Name}': falha ao criar '{corr.Name & cft_aterro_Suffix}'.")
                    Continue For
                End If

                If cortecftId.IsNull Then
                    ed.WriteMessage(Environment.NewLine & $"[SDT] Corredor '{corr.Name}': falha ao criar '{corr.Name & cft_corte_Suffix}'.")
                    Continue For
                End If

                created += 1
                ed.WriteMessage(Environment.NewLine & $"[SDT] OK: '{corr.Name}' criado (CFT Aterro/Corte).")
            Next

            Return created
        End Function


    End Class

End Namespace
