Option Strict On
Option Explicit On

Imports System
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices

Namespace SDT.Civil

    Public NotInheritable Class SurfaceLimpezaService
        Private Sub New()
        End Sub

        ''' <summary>
        ''' Cria uma "superfície de limpeza" a partir de uma TinSurface base (TN):
        ''' 1) clona (cópia independente) a superfície TN para um novo nome (TinSurfaceService.CloneTinSurface),
        ''' 2) aplica deslocamento vertical (raise),
        ''' 3) aplica boundary outer (Polyline) para limitar a área.
        ''' </summary>
        Public Shared Function Run(tr As Transaction,
                                  db As Database,
                                  civDoc As CivilDocument,
                                  ed As Editor,
                                  tnSurfaceId As ObjectId,
                                  espessuraCm As Double,
                                  outerBoundaryId As ObjectId,
                                  limpezaSurfaceName As String) As ObjectId

            If tr Is Nothing OrElse db Is Nothing OrElse civDoc Is Nothing OrElse ed Is Nothing Then Return ObjectId.Null
            If tnSurfaceId.IsNull Then Return ObjectId.Null
            If String.IsNullOrWhiteSpace(limpezaSurfaceName) Then Return ObjectId.Null

            ' TN -> clone (independente)
            Dim limpezaId As ObjectId = TinSurfaceService.CloneTinSurface(db, tr, civDoc, tnSurfaceId, limpezaSurfaceName, ed)
            If limpezaId.IsNull Then
                ed.WriteMessage(Environment.NewLine & "[SDT] Falha ao clonar superfície TN para '" & limpezaSurfaceName & "'.")
                Return ObjectId.Null
            End If

            Dim limpezaSurf As TinSurface = TryCast(tr.GetObject(limpezaId, OpenMode.ForWrite), TinSurface)
            If limpezaSurf Is Nothing Then Return ObjectId.Null

            ' deslocamento (cm -> m)
            Dim raise As Double = espessuraCm / 100.0
            TinSurfaceService.RaiseTinSurface(limpezaSurf, raise)

            ' boundary outer
            If Not outerBoundaryId.IsNull Then
                Dim ids As New ObjectIdCollection()
                ids.Add(outerBoundaryId)
                TinSurfaceService.ApplyOuterBoundary(limpezaSurf, ids, 0.5, True, ed)
            Else
                ed.WriteMessage(Environment.NewLine & $"[SDT] Aviso: '{limpezaSurfaceName}' criado sem boundary (outerBoundaryId = Null).")
            End If

            Return limpezaId
        End Function

    End Class

End Namespace
