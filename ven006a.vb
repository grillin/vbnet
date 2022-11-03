Imports cdata
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.Linq
Imports System.IO
Imports System.Configuration

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class ven006a
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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btn_salir As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CboZona As System.Windows.Forms.ComboBox
    Friend WithEvents cboDia As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboEmpresa As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cboDia = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.CboZona = New System.Windows.Forms.ComboBox
        Me.cboEmpresa = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btn_salir = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(40, 29)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 26)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Generar"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cboDia)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.CboZona)
        Me.GroupBox1.Controls.Add(Me.cboEmpresa)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(32, 32)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(664, 280)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "[ Datos ]"
        '
        'cboDia
        '
        Me.cboDia.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDia.FormattingEnabled = True
        Me.cboDia.Location = New System.Drawing.Point(261, 172)
        Me.cboDia.Name = "cboDia"
        Me.cboDia.Size = New System.Drawing.Size(256, 21)
        Me.cboDia.TabIndex = 25
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(157, 170)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 23)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Dia: "
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CboZona
        '
        Me.CboZona.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboZona.FormattingEnabled = True
        Me.CboZona.Location = New System.Drawing.Point(261, 134)
        Me.CboZona.Name = "CboZona"
        Me.CboZona.Size = New System.Drawing.Size(256, 21)
        Me.CboZona.TabIndex = 23
        '
        'cboEmpresa
        '
        Me.cboEmpresa.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEmpresa.FormattingEnabled = True
        Me.cboEmpresa.Location = New System.Drawing.Point(262, 92)
        Me.cboEmpresa.Name = "cboEmpresa"
        Me.cboEmpresa.Size = New System.Drawing.Size(255, 21)
        Me.cboEmpresa.TabIndex = 22
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(157, 132)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 23)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Zona: "
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(157, 89)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 23)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Empresa : "
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btn_salir)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Location = New System.Drawing.Point(208, 336)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(272, 72)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'btn_salir
        '
        Me.btn_salir.Location = New System.Drawing.Point(160, 29)
        Me.btn_salir.Name = "btn_salir"
        Me.btn_salir.Size = New System.Drawing.Size(75, 26)
        Me.btn_salir.TabIndex = 1
        Me.btn_salir.Text = "Salir"
        '
        'ven006a
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(744, 445)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "ven006a"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Generacion Listado de Rutas x Cliente"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public con As Conexion

    Public scmas As Cls_Scmaster
    Public ftexto As StreamWriter


    Public ds As DataSet
    Public gt As Gtexto


    Public ruta As String
    Public linea As String

    Dim ciudad As Integer = ConfigurationSettings.AppSettings("ciudad")

    Dim oruta As String = ConfigurationSettings.AppSettings("oruta")
    Dim druta As String = ConfigurationSettings.AppSettings("druta")


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        If Trim(Me.cboEmpresa.SelectedValue) = "" Then
            MsgBox("Empresa es requerida", MsgBoxStyle.Information)
            Exit Sub
        End If

        If Trim(Me.CboZona.SelectedValue) = "" Then
            MsgBox("Zona es requerida", MsgBoxStyle.Information)
            Exit Sub
        End If


        If Trim(Me.CboZona.SelectedValue) = "" Then
            MsgBox("Dia es requerido", MsgBoxStyle.Information)
            Exit Sub
        End If

        mostrar_reporte()


    End Sub


    Sub mostrar_reporte()


        Dim frpt As New ven006b


        frpt.empresa = Me.cboEmpresa.SelectedValue
        frpt.zona = Me.CboZona.SelectedValue
        frpt.dia = Me.cboDia.SelectedIndex + 1

        frpt.ShowDialog()


    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Text = "Listado de Rutas x Cliente"

        Call cargar_empresa()
        Call cargar_ruta()
        Call cargar_Dias()

    End Sub

    Function ValidarDatos() As Boolean

        Dim Valor As Boolean = True

        mensaje = ""

        Return Valor

    End Function

    Private Sub btn_salir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_salir.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Me.OpenFileDialog1.ShowDialog()
        'Me.txt_gestion.Text = Me.OpenFileDialog1.FileName

    End Sub

    Sub cargar_ruta()

        Dim con As Conexion
        Dim sql As String

        Dim codini As String
        Dim codfin As String
        Dim rangozona As String = ConfigurationSettings.AppSettings("RangoZona")
        Dim rzonaini As String
        Dim rzonafin As String

        con = New Conexion

        Dim scmas As New Cls_Scmaster

        con.abrir()

        'ciudad = 2

        codini = Trim(ciudad) & "01"
        codfin = Trim(ciudad) & "99"

        rzonaini = Trim(rangozona) & "00"
        rzonafin = Trim(rangozona) & "99"

        sql = "select substring(codigo,1,3) codigo, substring(codigo,1,3)+'-'+ descripcion descripcion from maestro.._tabla_parametros " & _
                          "  where tipo = 'OSVRUTAS' and ( ( codigo between " & codini & " and " & codfin & _
                          "    ) or ( codigo between " & rzonaini & " and " & rzonafin & " ) or codigo = 0 ) order by codigo"

        scmas.cnn = con.cnn
        Dim dtruta As DataTable = scmas.consulta(sql)

        con.cerrar()

        Me.CboZona.DataSource = dtruta


        Me.CboZona.DisplayMember = "descripcion"
        Me.CboZona.ValueMember = "codigo"


        con.cerrar()


    End Sub

    Sub cargar_empresa()

        Dim con As Conexion
        Dim sql As String

        con = New Conexion

        Dim scmas As New Cls_Scmaster

        con.abrir()

        sql = "select codigo-100 codigo, descripcion from maestro.._tabla_parametros where tipo='OMPGRUP7' and codigo in (101,102,103) "

        scmas.cnn = con.cnn
        Dim dtempresa As DataTable = scmas.consulta(sql)

        con.cerrar()

        Me.cboEmpresa.DataSource = dtempresa

        Me.cboEmpresa.DisplayMember = "descripcion"
        Me.cboEmpresa.ValueMember = "codigo"

        con.cerrar()



    End Sub

    Sub cargar_Dias()

        Me.cboDia.Items.Add("Lunes")
        Me.cboDia.Items.Add("Martes")
        Me.cboDia.Items.Add("Miercoles")
        Me.cboDia.Items.Add("Jueves")
        Me.cboDia.Items.Add("Viernes")
        Me.cboDia.Items.Add("Sabado")
        Me.cboDia.Items.Add("Todos")

    End Sub


End Class
