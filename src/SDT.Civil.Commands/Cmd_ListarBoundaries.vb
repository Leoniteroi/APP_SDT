Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports SDT.Drawing
Imports SDT.Geometry

''' <summary>
''' Lista boundaries de Corridors e gera polylines para boundaries externas (outer).
''' </summary>
Public Class Cmd_ListarBoundaries

    <CommandMethod("SDT_LIST_BOUNDARIES_CORRIDOR", CommandFlags.Modal)>
    Public Sub SDT_LIST_BOUNDARIES_CORRIDOR()

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        If doc Is Nothing Then Return

        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim civDoc As CivilDocument = CivilApplication.ActiveDocument
        If civDoc Is Nothing Then
            ed.WriteMessage(Environment.NewLine & "[SDT] CivilDocument não disponível.")
            Return
        End If

        Const layerOuter As String = "SDT_BOUNDARY_OUTER_DATUM"
        Const datumName As String = "DATUM"

        Using tr As Transaction = db.TransactionManager.StartTransaction()

            LayerService.Ensure(tr, db, layerOuter, 3)

            Dim bt As BlockTable =
                CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)

            Dim ms As BlockTableRecord =
                CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

            For Each corrId As ObjectId In civDoc.CorridorCollection

                Dim corr As Corridor =
                    TryCast(tr.GetObject(corrId, OpenMode.ForRead), Corridor)
                If corr Is Nothing Then Continue For

                ed.WriteMessage(Environment.NewLine & Environment.NewLine & "CORRIDOR: " & corr.Name)

                For Each surf As CorridorSurface In corr.CorridorSurfaces

                    ' Para facilitar a depuração com o comando de Limpeza:
                    ' foca apenas na superfície DATUM de cada corredor.
                    If surf Is Nothing OrElse Not surf.Name.Equals(datumName, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    ed.WriteMessage(Environment.NewLine & "  SURFACE: " & surf.Name)

                    For Each bnd As CorridorSurfaceBoundary In surf.Boundaries

                        ed.WriteMessage(Environment.NewLine &
                                        "    Boundary: " & bnd.Name &
                                        " | Tipo: " & bnd.BoundaryType.ToString())

                        If Not IsOuterBoundaryType(bnd.BoundaryType) Then Continue For

                        Dim rings As List(Of Point3dCollection) =
                            BoundaryRingExtractor.GetRings(bnd)

                        If rings Is Nothing OrElse rings.Count = 0 Then
                            ed.WriteMessage(Environment.NewLine & "      Sem geometria.")
                            Continue For
                        End If

                        For Each ring As Point3dCollection In rings
                            Dim pl As Polyline =
                                PolylineFactory.CreateClosed2DFromRing(ring, layerOuter)

                            If pl Is Nothing Then Continue For

                            ms.AppendEntity(pl)
                            tr.AddNewlyCreatedDBObject(pl, True)
                        Next

                    Next
                Next
            Next

            tr.Commit()
        End Using

        ed.WriteMessage(Environment.NewLine &
                        Environment.NewLine &
                        "[SDT] OK - Boundaries listados e polylines criadas no layer " & layerOuter)

    End Sub

    Private Function IsOuterBoundaryType(bt As CorridorSurfaceBoundaryType) As Boolean
        Dim s As String = bt.ToString().ToUpperInvariant()
        Return (s.Contains("OUTER") OrElse s.Contains("OUTSIDE") OrElse s.Contains("EXTERNO"))
    End Function

End Class
