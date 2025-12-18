Option Strict On
Option Explicit On

Imports System.Collections
Imports System.Collections.Generic
Imports System.Reflection
Imports Autodesk.AutoCAD.Geometry

Namespace SDT.Geometry

    ' Não depende de Autodesk.Civil.* (usa reflection)
    Public NotInheritable Class BoundaryRingExtractor
        Private Sub New()
        End Sub

        Public Shared Function GetRings(bnd As Object) As List(Of Point3dCollection)
            Dim rings As New List(Of Point3dCollection)()
            If bnd Is Nothing Then Return rings

            Dim mi As MethodInfo = bnd.GetType().GetMethod("GetExtentsBoundaries",
                                                          BindingFlags.Instance Or BindingFlags.Public)
            If mi Is Nothing OrElse mi.GetParameters().Length <> 0 Then Return rings

            Dim res As Object = Nothing
            Try
                res = mi.Invoke(bnd, Nothing)
            Catch
                Return rings
            End Try
            If res Is Nothing Then Return rings

            Dim direct As Point3dCollection = TryCast(res, Point3dCollection)
            If direct IsNot Nothing Then
                rings.Add(direct)
                Return rings
            End If

            Dim outer As IEnumerable = TryCast(res, IEnumerable)
            If outer Is Nothing Then Return rings

            Dim foundNested As Boolean = False
            Dim flatFallback As New Point3dCollection()

            For Each item As Object In outer
                If item Is Nothing Then Continue For

                Dim ringP3C As Point3dCollection = TryCast(item, Point3dCollection)
                If ringP3C IsNot Nothing Then
                    rings.Add(ringP3C)
                    foundNested = True
                    Continue For
                End If

                Dim inner As IEnumerable = TryCast(item, IEnumerable)
                If inner IsNot Nothing Then
                    Dim ring As New Point3dCollection()
                    Dim anyPoint As Boolean = False

                    For Each innerItem As Object In inner
                        If TypeOf innerItem Is Point3d Then
                            ring.Add(DirectCast(innerItem, Point3d))
                            anyPoint = True
                        End If
                    Next

                    If anyPoint Then
                        rings.Add(ring)
                        foundNested = True
                    End If
                    Continue For
                End If

                If TypeOf item Is Point3d Then
                    flatFallback.Add(DirectCast(item, Point3d))
                End If
            Next

            If Not foundNested AndAlso flatFallback.Count > 0 Then
                rings.Add(flatFallback)
            End If

            Return rings
        End Function
    End Class

End Namespace
