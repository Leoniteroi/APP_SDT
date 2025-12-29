Option Strict On
Option Explicit On

Imports System
Imports System.Reflection
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

            ' boundary outer (compatível com diferentes assinaturas do TinSurfaceService.ApplyOuterBoundary)
            If Not outerBoundaryId.IsNull Then
                ApplyOuterBoundaryCompat(limpezaSurf, outerBoundaryId, .5, True, ed)
            Else
                ed.WriteMessage(Environment.NewLine & "[SDT] Aviso: '" & limpezaSurfaceName & "' criado sem boundary (outerBoundaryId = Null).")
            End If

            Return limpezaId
        End Function

        ''' <summary>
        ''' Chama TinSurfaceService.ApplyOuterBoundary aceitando:
        ''' - (TinSurface, ObjectId, Double, Boolean, Editor) OU
        ''' - (TinSurface, ObjectIdCollection, Double, Boolean, Editor)
        ''' Usando reflection para evitar erro de compilação por overload diferente entre projetos.
        ''' </summary>
        Private Shared Sub ApplyOuterBoundaryCompat(tin As TinSurface,
                                                   boundaryId As ObjectId,
                                                   midOrdinate As Double,
                                                   useNonDestructiveBreakline As Boolean,
                                                   ed As Editor)

            If tin Is Nothing OrElse boundaryId.IsNull Then Exit Sub
            If midOrdinate <= 0 Then midOrdinate = 1.0

            Try
                Dim tSvc As Type = GetType(TinSurfaceService)


                ' 2) Fallback: overload com ObjectIdCollection (se existir)
                Dim miCol As MethodInfo = tSvc.GetMethod(
                    "ApplyOuterBoundary",
                    BindingFlags.Public Or BindingFlags.Static,
                    Nothing,
                    New Type() {GetType(TinSurface), GetType(ObjectIdCollection), GetType(Double), GetType(Boolean), GetType(Editor)},
                    Nothing)

                If miCol IsNot Nothing Then
                    Dim ids As New ObjectIdCollection()
                    ids.Add(boundaryId)
                    miCol.Invoke(Nothing, New Object() {tin, ids, midOrdinate, useNonDestructiveBreakline, ed})
                    Exit Sub
                End If

                ' 3) Último fallback: procurar qualquer ApplyOuterBoundary com 5 params e tentar montar args
                For Each mi As MethodInfo In tSvc.GetMethods(BindingFlags.Public Or BindingFlags.Static)
                    If mi Is Nothing OrElse mi.Name Is Nothing Then Continue For
                    If Not String.Equals(mi.Name, "ApplyOuterBoundary", StringComparison.OrdinalIgnoreCase) Then Continue For

                    Dim ps As ParameterInfo() = mi.GetParameters()
                    If ps Is Nothing OrElse ps.Length <> 5 Then Continue For
                    If ps(0).ParameterType IsNot GetType(TinSurface) Then Continue For

                    Dim args(4) As Object
                    args(0) = tin

                    If ps(1).ParameterType Is GetType(ObjectId) Then
                        args(1) = boundaryId
                    ElseIf ps(1).ParameterType Is GetType(ObjectIdCollection) Then
                        Dim ids As New ObjectIdCollection()
                        ids.Add(boundaryId)
                        args(1) = ids
                    Else
                        Continue For
                    End If

                    args(2) = midOrdinate
                    args(3) = useNonDestructiveBreakline
                    args(4) = ed

                    mi.Invoke(Nothing, args)
                    Exit Sub
                Next

                If ed IsNot Nothing Then ed.WriteMessage(Environment.NewLine & "[SDT] ApplyOuterBoundaryCompat: overload não encontrado em TinSurfaceService.")
            Catch ex As Exception
                If ed IsNot Nothing Then ed.WriteMessage(Environment.NewLine & "[SDT] ApplyOuterBoundaryCompat falhou: " & ex.Message)
            End Try

        End Sub

    End Class

End Namespace
