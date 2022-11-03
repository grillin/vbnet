


'Imports System.Collections.Specialized
'Imports System.Collections.ObjectModel
'Imports System.Collections
Imports System.Data
Imports System.IO
Imports System.Configuration

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared


Public Class ven006b
    Inherits System.Windows.Forms.Form

#Region " Código generado por el Diseñador de Windows Forms "

    Public Sub New()
        MyBase.New()

        'El Diseñador de Windows Forms requiere esta llamada.
        InitializeComponent()

        'Agregar cualquier inicialización después de la llamada a InitializeComponent()

    End Sub

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms requiere el siguiente procedimiento
    'Puede modificarse utilizando el Diseñador de Windows Forms. 
    'No lo modifique con el editor de código.
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.DisplayGroupTree = False
        Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.SelectionFormula = ""
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(292, 273)
        Me.CrystalReportViewer1.TabIndex = 5
        Me.CrystalReportViewer1.ViewTimeSelectionFormula = ""
        '
        'ven006b
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Name = "ven006b"
        Me.Text = "Listado de Rutas x Cliente"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

   
    Public empresa As Integer
    Public zona As Integer
    Public dia As Integer

    Dim usuario As String = ConfigurationSettings.AppSettings("usuario")
    Dim clave As String = ConfigurationSettings.AppSettings("clave")


    Private Sub fcotizab_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If dia <> 7 Then
            Call rpt_dias()
        Else
            Call rpt_todos()
        End If

    End Sub

    Sub rpt_dias()

        Dim rpt As New ven006r

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As ParameterValues
        Dim crParameterDiscreteValue As ParameterDiscreteValue

        'rpt.SetDataSource(ds)

        rpt.Refresh()


        rpt.SetDatabaseLogon(usuario, clave)


        crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@empresa")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterDiscreteValue = New ParameterDiscreteValue
        crParameterDiscreteValue.Value = Me.empresa
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue = Nothing
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@zona")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterDiscreteValue = New ParameterDiscreteValue
        crParameterDiscreteValue.Value = Me.zona
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue = Nothing
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@dia")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterDiscreteValue = New ParameterDiscreteValue
        crParameterDiscreteValue.Value = Me.dia
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


        Me.CrystalReportViewer1.ReportSource = rpt

        Me.CrystalReportViewer1.Show()


    End Sub

    Sub rpt_todos()

        Dim rpt As New ven006s

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As ParameterValues
        Dim crParameterDiscreteValue As ParameterDiscreteValue

        'rpt.SetDataSource(ds)

        rpt.Refresh()

        rpt.SetDatabaseLogon(usuario, clave)

        crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@empresa")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterDiscreteValue = New ParameterDiscreteValue
        crParameterDiscreteValue.Value = Me.empresa
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue = Nothing
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@zona")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterDiscreteValue = New ParameterDiscreteValue
        crParameterDiscreteValue.Value = Me.zona
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue = Nothing
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("@dia")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterDiscreteValue = New ParameterDiscreteValue
        crParameterDiscreteValue.Value = Me.dia
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


        Me.CrystalReportViewer1.ReportSource = rpt

        Me.CrystalReportViewer1.Show()


    End Sub
End Class
