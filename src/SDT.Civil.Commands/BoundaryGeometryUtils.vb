Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports Autodesk.AutoCAD.Geometry
Imports SDT.Geometry

Public NotInheritable Class BoundaryGeometryUtils
    Private Sub New()
    End Sub

    Public Shared Function GetRings(boundary As Object) As List(Of Point3dCollection)
        Return BoundaryRingExtractor.GetRings(boundary)
    End Function
End Class
