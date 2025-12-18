Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports Autodesk.AutoCAD.Geometry

Namespace SDT.Geometry

    Public NotInheritable Class RingMetrics
        Private Sub New()
        End Sub

        Public Shared Function SignedAreaXY(ring As Point3dCollection) As Double
            If ring Is Nothing OrElse ring.Count < 3 Then Return 0.0

            Dim area As Double = 0.0
            Dim n As Integer = ring.Count

            For i As Integer = 0 To n - 1
                Dim p1 As Point3d = ring(i)
                Dim p2 As Point3d = ring((i + 1) Mod n)
                area += (p1.X * p2.Y - p2.X * p1.Y)
            Next

            Return area / 2.0
        End Function

        Public Shared Function LargestByAbsArea(rings As List(Of Point3dCollection)) As Point3dCollection
            If rings Is Nothing OrElse rings.Count = 0 Then Return Nothing

            Dim best As Point3dCollection = Nothing
            Dim bestArea As Double = Double.MinValue

            For Each r As Point3dCollection In rings
                If r Is Nothing OrElse r.Count < 3 Then Continue For
                Dim a As Double = Math.Abs(SignedAreaXY(r))
                If a > bestArea Then
                    bestArea = a
                    best = r
                End If
            Next

            Return best
        End Function

    End Class

End Namespace
