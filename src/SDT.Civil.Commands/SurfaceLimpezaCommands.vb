Option Strict On
Option Explicit On

Imports System
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Win32
Imports SDT.Core
Imports Registry = Autodesk.AutoCAD.Runtime.Registry
Imports RegistryKey = Autodesk.AutoCAD.Runtime.RegistryKey

' OBS: Este arquivo usa WinForms.
' Requisitos no projeto SDT.Civil.Commands:
' - Referência a System.Windows.Forms
' - Referência a System.Drawing
' Se o .vbproj for SDK-style: incluir <UseWindowsForms>true</UseWindowsForms>

Public Class SurfaceLimpezaCommands

    <CommandMethod("SDT_CRIAR_SUPERFICIES_LIMPEZA", CommandFlags.Modal)>
    Public Sub SDT_CRIAR_SUPERFICIES_LIMPEZA()

        Dim ctx As AcadContext = Nothing
        Dim err As String = ""

        If Not AcadContext.TryCreate(ctx, err) Then Return

        Dim espessuraCm As Double = PromptEspessura(ctx.Ed)
        If espessuraCm <= 0 Then
            ctx.Ed.WriteMessage(Environment.NewLine & "[SDT] Espessura inválida.")
            Return
        End If

        TransactionRunner.RunWrite(
            ctx.Db,
            Sub(tr As Transaction)

                Dim tnId As ObjectId = PromptSelectTinSurfaceWithDialog(tr, ctx.CivDoc, ctx.Ed)
                If tnId.IsNull Then
                    ctx.Ed.WriteMessage(Environment.NewLine & "[SDT] Operação cancelada (TN não selecionada).")
                    Return
                End If

                Dim datumSuffix As String = "_DATUM"
                Dim limpezaPrefix As String = ""

                Dim n As Integer = SDT.Civil.SurfaceLimpezaBatchService.RunAllCorridors(
                    tr, ctx.Db, ctx.CivDoc, ctx.Ed, tnId, espessuraCm, datumSuffix, limpezaPrefix)

                ctx.Ed.WriteMessage(Environment.NewLine & "[SDT] Concluído. Superfícies criadas: " & n.ToString() & ".")
            End Sub)

    End Sub

    Private Function PromptEspessura(ed As Editor) As Double
        Dim pdo As New PromptDoubleOptions(Environment.NewLine & "Espessura da limpeza (cm): ")
        pdo.AllowNegative = False
        pdo.AllowZero = False
        pdo.DefaultValue = 20.0

        Dim pdr As PromptDoubleResult = ed.GetDouble(pdo)
        If pdr.Status <> PromptStatus.OK Then Return -1
        Return pdr.Value
    End Function

    ' ==========================================
    ' Caixa de seleção: lista de TinSurfaces (TN)
    ' - Mostra tipo [TIN] no combo
    ' - Filtra apenas TinSurface
    ' - Lembra a última seleção por nome (Registry)
    ' Refatorado: delega coleta de itens e gerenciamento de Registry ao TinSurfacePicker
    ' ==========================================
    Private Function PromptSelectTinSurfaceWithDialog(tr As Transaction, civDoc As CivilDocument, ed As Editor) As ObjectId
        If tr Is Nothing OrElse civDoc Is Nothing Then Return ObjectId.Null

        Dim lastName As String = TinSurfacePicker.ReadLastTinSurfaceName()

        Dim items As System.Collections.Generic.List(Of TinSurfaceItem) =
            TinSurfacePicker.GetTinSurfaceItems(tr, civDoc)

        If items.Count = 0 Then
            If ed IsNot Nothing Then ed.WriteMessage(Environment.NewLine & "[SDT] Nenhuma TinSurface encontrada no desenho.")
            Return ObjectId.Null
        End If

        ' Ordenar por nome sem LINQ (para evitar depender de Option Infer)
        items.Sort(Function(a As TinSurfaceItem, b As TinSurfaceItem) _
                   String.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase))

        Using dlg As New TinSurfacePickerDialog(items, lastName)
            Dim res As System.Windows.Forms.DialogResult = Application.ShowModalDialog(dlg)
            If res <> System.Windows.Forms.DialogResult.OK Then Return ObjectId.Null

            If dlg.SelectedItem IsNot Nothing Then
                TinSurfacePicker.SaveLastTinSurfaceName(dlg.SelectedItem.Name)
            End If

            Return dlg.SelectedId
        End Using
    End Function

End Class

' ----------------------------------------------------
' Modelo de item e utilitário para seleção (testável)
' - Mantém leitura/gravação de Registry e construção da lista
' - UI permanece no dialog, separado
' ----------------------------------------------------
Friend NotInheritable Class TinSurfacePicker

    Private Sub New()
    End Sub

    Private Const REG_PATH As String = "Software\SDT\Civil\SurfaceLimpeza"
    Private Const REG_LAST_TN_NAME As String = "LastTinSurfaceName"

    Public Shared Function GetTinSurfaceItems(tr As Transaction, civDoc As CivilDocument) As System.Collections.Generic.List(Of TinSurfaceItem)
        Dim items As New System.Collections.Generic.List(Of TinSurfaceItem)()

        If tr Is Nothing OrElse civDoc Is Nothing Then Return items

        For Each id As ObjectId In civDoc.GetSurfaceIds()
            If id.IsNull Then Continue For

            Dim civSurface As Autodesk.Civil.DatabaseServices.Surface =
                TryCast(tr.GetObject(id, OpenMode.ForRead), Autodesk.Civil.DatabaseServices.Surface)
            If civSurface Is Nothing Then Continue For

            Dim tin As TinSurface = TryCast(civSurface, TinSurface)
            If tin Is Nothing Then Continue For

            Dim nm As String = If(tin.Name, "").Trim()
            If nm.Length = 0 Then Continue For

            items.Add(New TinSurfaceItem(nm, "TIN", id))
        Next

        Return items
    End Function

    Public Shared Function ReadLastTinSurfaceName() As String
        Try
            Using k As RegistryKey = Registry.CurrentUser.OpenSubKey(REG_PATH, False)
                If k Is Nothing Then Return ""
                Dim v As Object = k.GetValue(REG_LAST_TN_NAME, "")
                Return If(v Is Nothing, "", Convert.ToString(v))
            End Using
        Catch
            Return ""
        End Try
    End Function

    Public Shared Sub SaveLastTinSurfaceName(name As String)
        If String.IsNullOrWhiteSpace(name) Then Exit Sub
        Try
            Using k As RegistryKey = Registry.CurrentUser.CreateSubKey(REG_PATH)
                k.SetValue(REG_LAST_TN_NAME, name.Trim(), RegistryValueKind.String)
            End Using
        Catch
            ' swallow - não crítico
        End Try
    End Sub

End Class

' Item reutilizável (facilita testes unitários)
Friend Class TinSurfaceItem
    Public ReadOnly Property Name As String
    Public ReadOnly Property TypeLabel As String
    Public ReadOnly Property Id As ObjectId

    Public Sub New(name As String, typeLabel As String, id As ObjectId)
        Me.Name = name
        Me.TypeLabel = typeLabel
        Me.Id = id
    End Sub

    Public ReadOnly Property DisplayText As String
        Get
            Return Me.Name & " [" & Me.TypeLabel & "]"
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return DisplayText
    End Function
End Class

' UI: Dialog separado (permanece in-file para simplicidade)
Friend Class TinSurfacePickerDialog
    Inherits System.Windows.Forms.Form

    Private ReadOnly _all As System.Collections.Generic.List(Of TinSurfaceItem)
    Private ReadOnly _txt As System.Windows.Forms.TextBox
    Private ReadOnly _cmb As System.Windows.Forms.ComboBox
    Private ReadOnly _ok As System.Windows.Forms.Button
    Private ReadOnly _cancel As System.Windows.Forms.Button
    Private ReadOnly _lastName As String

    Public Property SelectedId As ObjectId = ObjectId.Null
    Public Property SelectedItem As TinSurfaceItem = Nothing

    Public Sub New(items As System.Collections.Generic.List(Of TinSurfaceItem), lastName As String)
        _all = items
        _lastName = If(lastName, "").Trim()

        Me.Text = "Selecionar superfície TIN (TN)"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.ShowInTaskbar = False
        Me.Width = 560
        Me.Height = 190

        Dim lbl As New System.Windows.Forms.Label()
        lbl.Left = 12 : lbl.Top = 12 : lbl.Width = 520
        lbl.Text = "Filtrar (opcional) e selecione a TinSurface do Terreno Natural:"

        _txt = New System.Windows.Forms.TextBox()
        _txt.Left = 12 : _txt.Top = 35 : _txt.Width = 520
        AddHandler _txt.TextChanged, AddressOf OnFilterChanged

        _cmb = New System.Windows.Forms.ComboBox()
        _cmb.Left = 12 : _cmb.Top = 65 : _cmb.Width = 520
        _cmb.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList

        _ok = New System.Windows.Forms.Button()
        _ok.Left = 356 : _ok.Top = 110 : _ok.Width = 85
        _ok.Text = "OK"

        _cancel = New System.Windows.Forms.Button()
        _cancel.Left = 447 : _cancel.Top = 110 : _cancel.Width = 85
        _cancel.Text = "Cancelar"

        Me.AcceptButton = _ok
        Me.CancelButton = _cancel

        AddHandler _ok.Click, AddressOf OnOk
        AddHandler _cancel.Click, Sub() Me.DialogResult = System.Windows.Forms.DialogResult.Cancel

        Me.Controls.Add(lbl)
        Me.Controls.Add(_txt)
        Me.Controls.Add(_cmb)
        Me.Controls.Add(_ok)
        Me.Controls.Add(_cancel)

        LoadCombo(_all)

        If _lastName.Length > 0 Then
            Dim idx As Integer = FindIndexByName(_lastName)
            If idx >= 0 Then _cmb.SelectedIndex = idx
        End If
    End Sub

    Private Function FindIndexByName(name As String) As Integer
        For i As Integer = 0 To _cmb.Items.Count - 1
            Dim it As TinSurfaceItem = TryCast(_cmb.Items(i), TinSurfaceItem)
            If it IsNot Nothing AndAlso String.Equals(it.Name, name, StringComparison.OrdinalIgnoreCase) Then
                Return i
            End If
        Next
        Return -1
    End Function

    Private Sub LoadCombo(list As System.Collections.Generic.List(Of TinSurfaceItem))
        _cmb.BeginUpdate()
        _cmb.Items.Clear()
        For Each it As TinSurfaceItem In list
            _cmb.Items.Add(it)
        Next
        _cmb.EndUpdate()
        If _cmb.Items.Count > 0 Then _cmb.SelectedIndex = 0
    End Sub

    Private Sub OnFilterChanged(sender As Object, e As EventArgs)
        Dim q As String = If(_txt.Text, "").Trim()

        If q.Length = 0 Then
            LoadCombo(_all)
            Return
        End If

        Dim filtered As New System.Collections.Generic.List(Of TinSurfaceItem)()
        For Each it As TinSurfaceItem In _all
            If it.DisplayText.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0 Then
                filtered.Add(it)
            End If
        Next

        LoadCombo(filtered)
    End Sub

    Private Sub OnOk(sender As Object, e As EventArgs)
        Dim it As TinSurfaceItem = TryCast(_cmb.SelectedItem, TinSurfaceItem)
        If it Is Nothing Then
            System.Windows.Forms.MessageBox.Show(Me, "Selecione uma superfície.", "SDT",
                                                System.Windows.Forms.MessageBoxButtons.OK,
                                                System.Windows.Forms.MessageBoxIcon.Information)
            Return
        End If

        SelectedItem = it
        SelectedId = it.Id
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
    End Sub

End Class