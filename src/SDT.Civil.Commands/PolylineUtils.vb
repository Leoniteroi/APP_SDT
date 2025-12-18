Option Strict On
Option Explicit On

Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports SDT.Geometry

Public NotInheritable Class PolylineUtils
    Private Sub New()
    End Sub

    Public Shared Function CreateClosed2DFromRing(ring As Point3dCollection, layerName As String) As Polyline
        Return PolylineFactory.CreateClosed2DFromRing(ring, layerName)
    End Function
End Class
