Option Strict On
Option Explicit On

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Reflection
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices

Namespace SDT.Civil

    Public NotInheritable Class TinSurfaceService
        Private Sub New()
        End Sub

        Public Shared Function FindTinSurfaceIdByName(tr As Transaction,
                                                      civDoc As CivilDocument,
                                                      name As String) As ObjectId
            If tr Is Nothing OrElse civDoc Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then
                Return ObjectId.Null
            End If

            For Each id As ObjectId In civDoc.GetSurfaceIds()
                If id.IsNull Then Continue For
                Dim s As TinSurface = TryCast(tr.GetObject(id, OpenMode.ForRead), TinSurface)
                If s IsNot Nothing AndAlso String.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase) Then
                    Return id
                End If
            Next

            Return ObjectId.Null
        End Function

        Public Shared Sub SafeErase(tr As Transaction, id As ObjectId)
            If tr Is Nothing OrElse id.IsNull Then Exit Sub
            Try
                Dim dbo As Autodesk.AutoCAD.DatabaseServices.DBObject =
    TryCast(tr.GetObject(id, OpenMode.ForWrite, False), Autodesk.AutoCAD.DatabaseServices.DBObject)


                If dbo IsNot Nothing AndAlso Not dbo.IsErased Then dbo.Erase(True)
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' CLONA uma TinSurface (cópia independente) usando DeepCloneObjects.
        ''' Substitui qualquer abordagem baseada em PasteSurface.
        ''' </summary>
        ''' <remarks>
        ''' - Não cria dependência/ligação entre superfícies.
        ''' - Mais estável para processamento em lote.
        ''' </remarks>
        Public Shared Function CloneTinSurface(db As Database,
                                                  tr As Transaction,
                                                  civDoc As CivilDocument,
                                                  sourceTinId As ObjectId,
                                                  newName As String,
                                                  ed As Editor) As ObjectId
            If db Is Nothing OrElse tr Is Nothing OrElse civDoc Is Nothing Then Return ObjectId.Null
            If sourceTinId.IsNull Then Return ObjectId.Null

            Dim src As TinSurface = TryCast(tr.GetObject(sourceTinId, OpenMode.ForRead), TinSurface)
            If src Is Nothing Then Return ObjectId.Null

            If Not String.IsNullOrWhiteSpace(newName) Then
                Dim existingId As ObjectId = FindTinSurfaceIdByName(tr, civDoc, newName)
                If Not existingId.IsNull Then SafeErase(tr, existingId)
            End If

            Dim ownerId As ObjectId = src.OwnerId
            If ownerId.IsNull Then Return ObjectId.Null

            Dim ids As New ObjectIdCollection()
            ids.Add(sourceTinId)

            Dim idMap As New IdMapping()
            db.DeepCloneObjects(ids, ownerId, idMap, False)

            Dim newId As ObjectId = ObjectId.Null
            For Each pair As IdPair In idMap
                If pair.Key = sourceTinId AndAlso pair.IsCloned Then
                    newId = pair.Value
                    Exit For
                End If
            Next

            If newId.IsNull Then Return ObjectId.Null

            If Not String.IsNullOrWhiteSpace(newName) Then
                Dim cloned As TinSurface = TryCast(tr.GetObject(newId, OpenMode.ForWrite), TinSurface)
                If cloned IsNot Nothing Then cloned.Name = newName
            End If

            Return newId
        End Function

        Public Shared Sub RaiseTinSurface(tin As TinSurface, raise As Double)
            If tin Is Nothing Then Exit Sub
            If Math.Abs(raise) < 0.0000001 Then Exit Sub

            Try
                tin.RaiseSurface(raise)
                Exit Sub
            Catch
            End Try

            ' fallback reflection
            Try
                Dim t As Type = tin.GetType()
                Dim mi As MethodInfo =
                    t.GetMethod("RaiseSurface", BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                                Nothing, New Type() {GetType(Double)}, Nothing)

                If mi Is Nothing Then
                    mi = t.GetMethod("Raise", BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                                     Nothing, New Type() {GetType(Double)}, Nothing)
                End If

                If mi IsNot Nothing Then mi.Invoke(tin, New Object() {raise})
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Retorna True se a superfície aparenta ter triangulação (>=3 vértices).
        ''' (Usado para evitar NoTrianglesInSurface ao aplicar boundaries.)
        ''' </summary>
        Private Shared Function HasTriangulation(tin As TinSurface) As Boolean
            If tin Is Nothing Then Return False
            ' Alguns builds do Civil 3D não expõem GetVertices como membro público;
            ' para manter compatibilidade + Option Strict On, use reflection.
            Try
                Dim t As Type = tin.GetType()
                Dim mi As MethodInfo = t.GetMethod("GetVertices",
                                                  BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                                                  Nothing,
                                                  New Type() {GetType(Boolean)},
                                                  Nothing)

                If mi Is Nothing Then Return False

                Dim res As Object = mi.Invoke(tin, New Object() {False})
                If res Is Nothing Then Return False

                Dim coll As ICollection = TryCast(res, ICollection)
                If coll IsNot Nothing Then
                    Return coll.Count >= 3
                End If

                ' Fallback: se vier como IEnumerable, conta até 3
                Dim en As IEnumerable = TryCast(res, IEnumerable)
                If en Is Nothing Then Return False
                Dim c As Integer = 0
                For Each item As Object In en
                    c += 1
                    If c >= 3 Then Return True
                Next
                Return False
            Catch
                Return False
            End Try
        End Function

        Private Shared Sub TryRebuild(tin As TinSurface, ed As Editor, context As String)
            If tin Is Nothing Then Exit Sub
            Try
                tin.Rebuild()
            Catch ex As Exception
                If ed IsNot Nothing Then ed.WriteMessage(Environment.NewLine & $"[SDT] Rebuild falhou ({context}): {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Aplica boundary OUTER em uma TinSurface usando entidades (Polyline(s)) que definem o contorno.
        ''' 
        ''' Observação:
        ''' - Em Civil 3D 2024 o método comum é:
        '''   BoundariesDefinition.AddBoundaries(ObjectIdCollection, Double, SurfaceBoundaryType, Boolean)
        ''' - Se tentar aplicar boundary em uma TIN sem triângulos, a API lança NoTrianglesInSurface.
        ''' </summary>
        Public Shared Sub ApplyOuterBoundary(tin As TinSurface,
                                             boundaryEntIds As ObjectIdCollection,
                                             midOrdinate As Double,
                                             useNonDestructiveBreakline As Boolean,
                                             ed As Editor)

            If tin Is Nothing Then Exit Sub
            If boundaryEntIds Is Nothing OrElse boundaryEntIds.Count = 0 Then Exit Sub

            ' Garantir triangulação antes de aplicar boundary
            'TryRebuild(tin, ed, "ApplyOuterBoundary-pre")
            'If Not HasTriangulation(tin) Then
            'If ed IsNot Nothing Then ed.WriteMessage(Environment.NewLine & $"[SDT] Boundary ignorado: superfície '{tin.Name}' sem triângulos.")
            'Exit Sub
            'End If

            ' Garantir mid-ordinate válido
            If midOrdinate <= 0 Then midOrdinate = 1.0

            Try
                ' Caminho preferencial: chamada tipada (sem reflection)
                ' Na maioria das versões: SurfaceDefinitionBoundaries.AddBoundaries(ObjectIdCollection, Double, SurfaceBoundaryType, Boolean)
                Dim bd As SurfaceDefinitionBoundaries = tin.BoundariesDefinition
                If bd IsNot Nothing Then
                    Try
                        bd.AddBoundaries(boundaryEntIds, midOrdinate, SurfaceBoundaryType.Outer, useNonDestructiveBreakline)
                        TryRebuild(tin, ed, "ApplyOuterBoundary-post")
                        Return
                    Catch
                        ' Se esta sobrecarga não existir, cai para reflection abaixo
                    End Try
                End If

                ' Fallback por reflection: procurar ESPECIFICAMENTE "AddBoundaries" e sobrecargas comuns
                Dim bdObj As Object = tin.BoundariesDefinition
                If bdObj Is Nothing Then
                    If ed IsNot Nothing Then ed.WriteMessage(Environment.NewLine & "[SDT] BoundariesDefinition = Nothing.")
                    Exit Sub
                End If

                Dim bdType As Type = bdObj.GetType()

                Dim candidates As New List(Of MethodInfo)()

                For Each mi As MethodInfo In bdType.GetMethods(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
                    If mi Is Nothing OrElse mi.Name Is Nothing Then Continue For
                    If Not String.Equals(mi.Name, "AddBoundaries", StringComparison.OrdinalIgnoreCase) Then Continue For

                    Dim ps As ParameterInfo() = mi.GetParameters()
                    If ps Is Nothing OrElse ps.Length < 4 Then Continue For

                    ' p0: ObjectIdCollection
                    If ps(0).ParameterType IsNot GetType(ObjectIdCollection) Then Continue For
                    ' p1: Double
                    If ps(1).ParameterType IsNot GetType(Double) Then Continue For
                    ' p2: enum (SurfaceBoundaryType ou similar)
                    If Not ps(2).ParameterType.IsEnum Then Continue For
                    ' p3: Boolean
                    If ps(3).ParameterType IsNot GetType(Boolean) Then Continue For

                    candidates.Add(mi)
                Next

                If candidates.Count = 0 Then
                    If ed IsNot Nothing Then ed.WriteMessage(Environment.NewLine & "[SDT] Não encontrei AddBoundaries compatível em BoundariesDefinition.")
                    Exit Sub
                End If

                ' Preferir a assinatura com 4 parâmetros; senão pega a menor (mais compatível)
                candidates.Sort(Function(a, b) a.GetParameters().Length.CompareTo(b.GetParameters().Length))
                Dim chosen As MethodInfo = candidates(0)
                Dim psChosen As ParameterInfo() = chosen.GetParameters()

                Dim args(psChosen.Length - 1) As Object
                args(0) = boundaryEntIds
                args(1) = midOrdinate

                Dim enumType As Type = psChosen(2).ParameterType
                Dim outerValue As Object
                Try
                    outerValue = [Enum].Parse(enumType, "Outer", True)
                Catch
                    ' fallback: tenta primeiro item
                    outerValue = [Enum].GetValues(enumType).GetValue(0)
                End Try
                args(2) = outerValue

                args(3) = useNonDestructiveBreakline

                ' parâmetros extras (nome/descrição/flags) -> preencher de forma segura
                For i As Integer = 4 To psChosen.Length - 1
                    Dim pt As Type = psChosen(i).ParameterType
                    If pt Is GetType(String) Then
                        args(i) = "SDT_OUTER"
                    ElseIf pt Is GetType(Boolean) Then
                        args(i) = False
                    ElseIf pt Is GetType(Double) Then
                        args(i) = 0.0R
                    ElseIf pt Is GetType(Integer) Then
                        args(i) = 0
                    Else
                        args(i) = Nothing
                    End If
                Next

                chosen.Invoke(bdObj, args)
                TryRebuild(tin, ed, "ApplyOuterBoundary-post")

            Catch ex As Exception
                If ed IsNot Nothing Then ed.WriteMessage(Environment.NewLine & "[SDT] ApplyOuterBoundary falhou: " & ex.Message)
            End Try

        End Sub

        ' Compatibilidade: versão antiga com 1 entidade
        Public Shared Sub ApplyOuterBoundary(tin As TinSurface,
                                             boundaryPolylineId As ObjectId,
                                             midOrdinate As Double,
                                             keepEntities As Boolean,
                                             ed As Editor)
            Dim ids As New ObjectIdCollection()
            If Not boundaryPolylineId.IsNull Then ids.Add(boundaryPolylineId)
            ApplyOuterBoundary(tin, ids, midOrdinate, False, ed)
        End Sub


    End Class

End Namespace