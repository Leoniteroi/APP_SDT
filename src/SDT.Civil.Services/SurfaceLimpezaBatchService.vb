Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports SDT.Drawing
Imports SDT.Geometry

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
    Public NotInheritable Class SurfaceLimpezaBatchService
        Private Sub New()
        End Sub

        Public Shared Function RunAllCorridors(tr As Transaction,
                                              db As Database,
                                              civDoc As CivilDocument,
                                              ed As Editor,
                                              tnBaseSurfaceId As ObjectId,
                                              espessuraCm As Double,
                                              datumSurfaceSuffix As String,
                                              limpezaPrefix As String) As Integer

            If tr Is Nothing OrElse db Is Nothing OrElse civDoc Is Nothing OrElse ed Is Nothing Then Return 0
            If tnBaseSurfaceId.IsNull Then Return 0

            If String.IsNullOrWhiteSpace(datumSurfaceSuffix) Then datumSurfaceSuffix = "_DATUM"
            If limpezaPrefix Is Nothing Then limpezaPrefix = ""

            Const layerOuterTmp As String = "SDT_TMP_BOUNDARY_OUTER"
            LayerService.Ensure(tr, db, layerOuterTmp, 3)

            Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim ms As BlockTableRecord = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

            Dim created As Integer = 0

            For Each corrId As ObjectId In civDoc.CorridorCollection
                If corrId.IsNull Then Continue For

                Dim corr As Corridor = TryCast(tr.GetObject(corrId, OpenMode.ForRead), Corridor)
                If corr Is Nothing Then Continue For

                ' localizar surface "*_DATUM_OUTER"
                Dim datumSurf As CorridorSurface = FindCorridorSurfaceBySuffix(corr, datumSurfaceSuffix)
                If datumSurf Is Nothing Then
                    ed.WriteMessage(Environment.NewLine & $"[SDT] Corredor '{corr.Name}': sem surface '*{datumSurfaceSuffix}'.")
                    Continue For
                End If

                ' extrair outer boundary(s) (gera polylines temporárias)
                Dim boundaryPlIds As ObjectIdCollection = ExtractOuterBoundariesAsPolylines(tr, ms, datumSurf, layerOuterTmp)
                If boundaryPlIds Is Nothing OrElse boundaryPlIds.Count = 0 Then
                    ed.WriteMessage(Environment.NewLine & $"[SDT] Corredor '{corr.Name}': '{datumSurf.Name}' sem OuterBoundary.")
                    Continue For
                End If

                ' usar o MAIOR contorno 2D (contorno externo principal)
                Dim bestBoundaryId As ObjectId = ChooseLargestPolyline2D(tr, boundaryPlIds)
                If bestBoundaryId.IsNull Then
                    For Each pid As ObjectId In boundaryPlIds
                        TinSurfaceService.SafeErase(tr, pid)
                    Next
                    ed.WriteMessage(Environment.NewLine & $"[SDT] Corredor '{corr.Name}': falha ao selecionar boundary.")
                    Continue For
                End If

                ' nome final: "{Corridor}_LIMPEZA" (ou com prefixo)
                Dim limpezaName As String
                If String.IsNullOrWhiteSpace(limpezaPrefix) Then
                    limpezaName = $"{corr.Name}_LIMPEZA"
                Else
                    limpezaName = $"{limpezaPrefix}_{corr.Name}_LIMPEZA"
                End If

                ' motor: clona TN base + raise + boundary
                Dim limpezaId As ObjectId = SurfaceLimpezaService.Run(
                    tr, db, civDoc, ed, tnBaseSurfaceId, espessuraCm, bestBoundaryId, limpezaName)

                ' limpar polylines temporárias
                For Each pid As ObjectId In boundaryPlIds
                    TinSurfaceService.SafeErase(tr, pid)
                Next

                If limpezaId.IsNull Then
                        ed.WriteMessage(Environment.NewLine & $"[SDT] Corredor '{corr.Name}': falha ao criar '{limpezaName}'.")
                        Continue For
                    End If

                    created += 1
                    ed.WriteMessage(Environment.NewLine & $"[SDT] OK: '{limpezaName}' criado (TN + boundary '{datumSurf.Name}').")
                Next

                Return created
        End Function

        Private Shared Function FindCorridorSurfaceBySuffix(corr As Corridor, suffix As String) As CorridorSurface
            If corr Is Nothing OrElse String.IsNullOrWhiteSpace(suffix) Then Return Nothing
            For Each s As CorridorSurface In corr.CorridorSurfaces
                If s Is Nothing OrElse String.IsNullOrWhiteSpace(s.Name) Then Continue For
                If s.Name.Trim().EndsWith(suffix, StringComparison.OrdinalIgnoreCase) Then
                    Return s
                End If
            Next
            Return Nothing
        End Function

        Private Shared Function ExtractOuterBoundariesAsPolylines(tr As Transaction,
                                                                 ms As BlockTableRecord,
                                                                 corSurf As CorridorSurface,
                                                                 layerName As String) As ObjectIdCollection
            Dim res As New ObjectIdCollection()
            If tr Is Nothing OrElse ms Is Nothing OrElse corSurf Is Nothing Then Return res

            Try
                For Each bnd As CorridorSurfaceBoundary In corSurf.Boundaries
                    If bnd Is Nothing Then Continue For
                    If Not IsOuterBoundaryType(bnd.BoundaryType) Then Continue For

                    Dim rings As List(Of Point3dCollection) = BoundaryRingExtractor.GetRings(bnd)
                    If rings Is Nothing OrElse rings.Count = 0 Then Continue For

                    Dim best As Point3dCollection = Nothing
                    Dim bestArea As Double = Double.MinValue
                    For Each r As Point3dCollection In rings
                        Dim a As Double = RingArea2D(r)
                        If a > bestArea Then
                            bestArea = a
                            best = r
                        End If
                    Next

                    If best Is Nothing Then Continue For

                    Dim pl As Polyline = PolylineFactory.CreateClosed2DFromRing(best, layerName)
                    If pl Is Nothing Then Continue For

                    ms.AppendEntity(pl)
                    tr.AddNewlyCreatedDBObject(pl, True)
                    res.Add(pl.ObjectId)
                Next
            Catch
                ' Versões/ambientes onde Boundaries não está disponível: retorna vazio (skip no caller)
            End Try

            Return res
        End Function

        Private Shared Function ChooseLargestPolyline2D(tr As Transaction, ids As ObjectIdCollection) As ObjectId
            If tr Is Nothing OrElse ids Is Nothing OrElse ids.Count = 0 Then Return ObjectId.Null

            Dim bestId As ObjectId = ObjectId.Null
            Dim bestArea As Double = Double.MinValue

            For Each id As ObjectId In ids
                If id.IsNull Then Continue For
                Dim pl As Polyline = TryCast(tr.GetObject(id, OpenMode.ForRead), Polyline)
                If pl Is Nothing Then Continue For

                Dim a As Double = Math.Abs(PolylineArea2D(pl))
                If a > bestArea Then
                    bestArea = a
                    bestId = id
                End If
            Next

            Return bestId
        End Function

        Private Shared Function PolylineArea2D(pl As Polyline) As Double
            If pl Is Nothing OrElse pl.NumberOfVertices < 3 Then Return 0
            Dim area As Double = 0
            Dim n As Integer = pl.NumberOfVertices
            For i As Integer = 0 To n - 1
                Dim p1 As Point2d = pl.GetPoint2dAt(i)
                Dim p2 As Point2d = pl.GetPoint2dAt((i + 1) Mod n)
                area += (p1.X * p2.Y) - (p2.X * p1.Y)
            Next
            Return 0.5 * area
        End Function

        Private Shared Function IsOuterBoundaryType(bt As CorridorSurfaceBoundaryType) As Boolean
            Dim s As String = bt.ToString().ToUpperInvariant()
            Return (s.Contains("OUTER") OrElse s.Contains("OUTSIDE") OrElse s.Contains("EXTERNO"))
        End Function

        Private Shared Function RingArea2D(ring As Point3dCollection) As Double
            If ring Is Nothing OrElse ring.Count < 3 Then Return 0
            Dim area As Double = 0
            For i As Integer = 0 To ring.Count - 2
                Dim p1 As Point3d = ring(i)
                Dim p2 As Point3d = ring(i + 1)
                area += (p1.X * p2.Y) - (p2.X * p1.Y)
            Next
            Dim pLast As Point3d = ring(ring.Count - 1)
            Dim p0 As Point3d = ring(0)
            area += (pLast.X * p0.Y) - (p0.X * pLast.Y)
            Return Math.Abs(area) * 0.5
        End Function

    End Class

End Namespace
