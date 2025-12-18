Option Strict On
Option Explicit On

Imports System
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry

Namespace SDT.Geometry

    Public NotInheritable Class PolylineFactory
        Private Sub New()
        End Sub

        Public Shared Function CreateClosed2DFromRing(ring As Point3dCollection,
                                                     layerName As String,
                                                     Optional closeTolerance As Double = 0.001) As Polyline
            If ring Is Nothing OrElse ring.Count < 2 Then Return Nothing

            Dim pl As New Polyline()
            pl.SetDatabaseDefaults()
            pl.Layer = layerName

            For i As Integer = 0 To ring.Count - 1
                Dim p As Point3d = ring(i)
                pl.AddVertexAt(pl.NumberOfVertices, New Point2d(p.X, p.Y), 0.0, 0.0, 0.0)
            Next

            If ring.Count >= 2 Then
                Dim firstP As Point3d = ring(0)
                Dim lastP As Point3d = ring(ring.Count - 1)

                If Math.Abs(firstP.X - lastP.X) <= closeTolerance AndAlso
                   Math.Abs(firstP.Y - lastP.Y) <= closeTolerance Then
                    If pl.NumberOfVertices > 1 Then
                        pl.RemoveVertexAt(pl.NumberOfVertices - 1)
                    End If
                End If
            End If

            pl.Closed = True
            Return pl
        End Function

    End Class

End Namespace
