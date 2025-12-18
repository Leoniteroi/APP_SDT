Option Strict On
Option Explicit On

Imports System
Imports System.Reflection
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices

Namespace SDT.Civil

    Public NotInheritable Class SurfaceLimpezaBatchService
        Private Sub New()
        End Sub

        Public Shared Sub RunAllCorridors(tr As Transaction,
                                          db As Database,
                                          civDoc As CivilDocument,
                                          ed As Editor,
                                          tnBaseSurfaceId As ObjectId,
                                          espessuraCm As Double,
                                          datumSurfaceName As String,
                                          limpezaPrefix As String)

            If tr Is Nothing OrElse db Is Nothing OrElse civDoc Is Nothing OrElse ed Is Nothing Then Exit Sub
            If tnBaseSurfaceId.IsNull Then Exit Sub
            If String.IsNullOrWhiteSpace(datumSurfaceName) Then datumSurfaceName = "DATUM"
            If String.IsNullOrWhiteSpace(limpezaPrefix) Then limpezaPrefix = "LIMPEZA"

            Dim raise As Double = espessuraCm / 100.0

            For Each corrId As ObjectId In civDoc.CorridorCollection
                Dim corr As Corridor = TryCast(tr.GetObject(corrId, OpenMode.ForRead), Corridor)
                If corr Is Nothing Then Continue For

                Dim datumSurf As CorridorSurface = FindCorridorSurfaceByName(corr, datumSurfaceName)
                If datumSurf Is Nothing Then
                    ed.WriteMessage(Environment.NewLine & $"[SDT] Corredor '{corr.Name}': superfície '{datumSurfaceName}' não encontrada.")
                    Continue For
                End If

                ' 1) obter boundary do DATUM daquele corredor
                Dim outerBoundaryId As ObjectId = GetOuterBoundaryPolylineIdFromCorridorSurface(tr, datumSurf, ed)

                If outerBoundaryId.IsNull Then
                    ed.WriteMessage(Environment.NewLine & $"[SDT] Corredor '{corr.Name}': não achei Polyline de boundary do DATUM (precisa extrair).")
                    Continue For
                End If

                ' 2) criar nome da limpeza por corredor
                Dim limpezaName As String = $"{limpezaPrefix}_{corr.Name}"

                ' 3) clonar TN base -> limpeza
                Dim limpezaId As ObjectId = TinSurfaceService.CopyFromTinSurface(db, tr, civDoc, tnBaseSurfaceId, limpezaName, ed)
                If limpezaId.IsNull Then
                    ed.WriteMessage(Environment.NewLine & $"[SDT] Corredor '{corr.Name}': falha ao clonar TN para '{limpezaName}'.")
                    Continue For
                End If

                Dim limpezaSurf As TinSurface = TryCast(tr.GetObject(limpezaId, OpenMode.ForWrite), TinSurface)
                If limpezaSurf Is Nothing Then Continue For

                ' 4) raise
                TinSurfaceService.RaiseTinSurface(limpezaSurf, raise)

                ' 5) aplicar boundary (do DATUM do corredor)
                TinSurfaceService.ApplyOuterBoundary(limpezaSurf, outerBoundaryId, 1.0, True, ed)

                ed.WriteMessage(Environment.NewLine & $"[SDT] OK: '{limpezaName}' criado com boundary do DATUM de '{corr.Name}'.")
            Next

        End Sub

        Private Shared Function FindCorridorSurfaceByName(corr As Corridor, name As String) As CorridorSurface
            If corr Is Nothing OrElse String.IsNullOrWhiteSpace(name) Then Return Nothing
            For Each s As CorridorSurface In corr.CorridorSurfaces
                If s IsNot Nothing AndAlso s.Name.Equals(name, StringComparison.OrdinalIgnoreCase) Then
                    Return s
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Tenta obter o ObjectId da Polyline que foi usada como "Outer Boundary" na surface DATUM do corredor.
        ''' Se a boundary do corredor não tiver entidade associada, retorna Null (aí você precisa extrair).
        ''' </summary>
        Private Shared Function GetOuterBoundaryPolylineIdFromCorridorSurface(tr As Transaction,
                                                                             corSurf As CorridorSurface,
                                                                             ed As Editor) As ObjectId
            If tr Is Nothing OrElse corSurf Is Nothing Then Return ObjectId.Null

            ' --- Tentativa 1: API direta (quando disponível) ---
            Try
                ' Algumas versões expõem .Boundaries (coleção de CorridorSurfaceBoundary)
                Dim bdProp As PropertyInfo = corSurf.GetType().GetProperty("Boundaries", BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
                If bdProp IsNot Nothing Then
                    Dim bdObj As Object = bdProp.GetValue(corSurf, Nothing)
                    Dim enumerable As System.Collections.IEnumerable = TryCast(bdObj, System.Collections.IEnumerable)
                    If enumerable IsNot Nothing Then
                        For Each b As Object In enumerable
                            If b Is Nothing Then Continue For

                            ' boundaryType == Outer ?
                            Dim bt As Object = GetMemberValue(b, "BoundaryType")
                            If bt IsNot Nothing AndAlso bt.ToString().Equals("Outer", StringComparison.OrdinalIgnoreCase) Then
                                Dim entIdObj As Object = GetMemberValue(b, "BoundaryEntityId")
                                If TypeOf entIdObj Is ObjectId Then
                                    Dim entId As ObjectId = CType(entIdObj, ObjectId)
                                    If Not entId.IsNull Then Return entId
                                End If

                                ' fallback: EntityId / PolylineId
                                entIdObj = GetMemberValue(b, "EntityId")
                                If TypeOf entIdObj Is ObjectId Then
                                    Dim entId As ObjectId = CType(entIdObj, ObjectId)
                                    If Not entId.IsNull Then Return entId
                                End If
                            End If
                        Next
                    End If
                End If
            Catch
            End Try

            ' --- Tentativa 2: Reflection procurando qualquer propriedade ObjectId plausível ---
            Try
                Dim t As Type = corSurf.GetType()
                For Each p As PropertyInfo In t.GetProperties(BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
                    If p.PropertyType Is GetType(ObjectId) AndAlso p.Name.IndexOf("Boundary", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        Dim oid As ObjectId = CType(p.GetValue(corSurf, Nothing), ObjectId)
                        If Not oid.IsNull Then Return oid
                    End If
                Next
            Catch
            End Try

            Return ObjectId.Null
        End Function

        Private Shared Function GetMemberValue(obj As Object, memberName As String) As Object
            Dim t As Type = obj.GetType()

            Dim p As PropertyInfo = t.GetProperty(memberName, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
            If p IsNot Nothing Then Return p.GetValue(obj, Nothing)

            Dim f As FieldInfo = t.GetField(memberName, BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
            If f IsNot Nothing Then Return f.GetValue(obj)

            Return Nothing
        End Function

    End Class

End Namespace
