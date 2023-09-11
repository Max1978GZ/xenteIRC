VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form formgenteirc 
   BorderStyle     =   0  'None
   Caption         =   "Agenda de gente del IRC"
   ClientHeight    =   5820
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox edad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5400
      TabIndex        =   23
      Top             =   2400
      Width           =   855
   End
   Begin MSComDlg.CommonDialog dialogo 
      Left            =   6360
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.*"
      DialogTitle     =   "Escoge la imagen del usuario o nick seleccionado "
      FileName        =   "*.*"
      Filter          =   "*.jpg"
      FilterIndex     =   1
   End
   Begin VB.Frame panelfoto 
      Caption         =   "Foto"
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   2175
      Begin VB.Image foto 
         Enabled         =   0   'False
         Height          =   1725
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listado de Nicks"
      Height          =   735
      Left            =   4080
      TabIndex        =   21
      Top             =   0
      Width           =   2535
      Begin VB.ComboBox listadonick 
         Height          =   315
         ItemData        =   "Formgenteirc.frx":0000
         Left            =   120
         List            =   "Formgenteirc.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame panelbusqueda 
      Caption         =   "Busqueda por Nick"
      Height          =   1215
      Left            =   2760
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton btn_buscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   840
         TabIndex        =   20
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox textBox_buscar_nick 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "salir"
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   5280
      Width           =   975
   End
   Begin VB.ComboBox año 
      Height          =   315
      ItemData        =   "Formgenteirc.frx":0004
      Left            =   1680
      List            =   "Formgenteirc.frx":0006
      TabIndex        =   4
      Text            =   "1978"
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox web 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   3120
      Width           =   5055
   End
   Begin VB.TextBox telefono 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   2760
      Width           =   5055
   End
   Begin VB.TextBox ciudad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   2040
      Width           =   5055
   End
   Begin VB.TextBox correo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1680
      Width           =   5055
   End
   Begin VB.TextBox nombre 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   5055
   End
   Begin VB.TextBox nick 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Edad"
      Height          =   195
      Left            =   4800
      TabIndex        =   24
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   120
      Picture         =   "Formgenteirc.frx":0008
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Web"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Año de Nacimiento"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   1350
   End
   Begin VB.Label lCiudad 
      AutoSize        =   -1  'True
      Caption         =   "Ciudad"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lCorreo 
      AutoSize        =   -1  'True
      Caption         =   "Correo"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   465
   End
   Begin VB.Label lnombre 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label lnick 
      AutoSize        =   -1  'True
      Caption         =   "Nick"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   330
   End
   Begin VB.Menu Gestion 
      Caption         =   "&Gestion"
      Begin VB.Menu alta 
         Caption         =   "&Alta"
      End
      Begin VB.Menu Modificar 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu eliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu busqueda 
         Caption         =   "&Busqueda"
      End
      Begin VB.Menu salir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu Acercade 
      Caption         =   "&Acerca de "
   End
End
Attribute VB_Name = "formgenteirc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Nombre : Xente IRC
'Descripción: Gestor de Amigos del IRC con base de datos access 97
'Creador : Máximo Coejo Cores
'email : maximodutty@terra.es
'Web :www.maximocoejo.blogia.com
'Fecha: Marzo 2005 Cambados-Pontevedra-España
Option Explicit
Dim ruta_fotos As String          'guarda la ruta a la carpeta de fotos
Public UsuariosIRC As Database    'guarda la base de datos
Public TablaUsuarios As Recordset 'para recorer la base de datos
Public MODO As String             'Para los 3 modos ("ALTA","BAJA","ACTUALIZAR")

'****** cargamos los valores por defecto en el formulario *****************
Private Sub Form_Load()
  Dim ruta As String 'ruta ala base programa
  Dim años As Integer
  ruta = App.Path 'preparamos el la ruta al programa
  ruta_fotos = ruta & "\fotos" ' preparamos el path del directorio fotos
  Set UsuariosIRC = OpenDatabase(ruta & "\genteIrc.mdb") 'definimos la base de datos a utilizar
  'creamos el puntero para recorrer la tabla y la abrimos
  Set TablaUsuarios = UsuariosIRC.OpenRecordset("usuarios", dbOpenTable)
  MODO = ""
  For años = 1950 To Year(Date) ' cargamos el combo con los años
    año.AddItem (años)
  Next
  Desactivar_textbox
  panelbusqueda.Visible = False
  Actualiza_ComboNicks 'actualiza los nick de la base en el combobox
  Me.Desactivar_MENU_Mod_Eliminar
End Sub
'************************************************************************
'************************* MENU ************************************************
Private Sub alta_Click() 'prepara el form para crearun nuevo registro
 MODO = "ALTA"
 Me.Desactivar_MENU
 Me.listadonick.Clear
 Me.listadonick.Enabled = False
 Me.Vaciar_campos
 Activar_textbox
 Aceptar.Visible = True
 Cancelar.Visible = True
 nick.SetFocus
End Sub
'prepara el formulario para modificar el registro
Private Sub modificar_Click()
 Me.listadonick.Enabled = False
 MODO = "ACTUALIZAR"
 Me.Desactivar_MENU
 Me.Mostrar_botones
 Me.Activar_textbox
 nick.Enabled = False
 nombre.SetFocus
End Sub
Private Sub eliminar_Click() 'prepara el form para borrar el registro
  MODO = "BAJA"
  Me.Desactivar_MENU
  Me.Mostrar_botones
End Sub
Private Sub busqueda_Click() 'muestra el textbox y el boton de busqueda
  If Me.busqueda.Checked = False Then
    panelbusqueda.Visible = True
    textBox_buscar_nick.SetFocus
    Me.busqueda.Checked = True
    Me.btn_buscar.Enabled = False
  Else
    panelbusqueda.Visible = False
    Me.busqueda.Checked = False
  End If
 End Sub
 Private Sub salir_Click() ' salimos de la aplicacion
   TablaUsuarios.Close
   UsuariosIRC.Close
   End
 End Sub
'************************** FIN DEL MENU ******************************
'************** FUNCIONES DE ALTA , BAJA, ACTUALIZACION ***************
Private Sub Aceptar_Click()
Select Case MODO
 Case "ALTA" 'si MODO es igual a alta inserta los campos en la Base de datos
      If Not validar Then
        Exit Sub
      End If
      TablaUsuarios.AddNew ' los añade a la base de datos
      copiarcampos ' copia los campos
      TablaUsuarios.Update  'guardarmos el registro
      Aceptar.Visible = False 'desactiva el boton aceptar
      Me.Activar_MENU_Alta_Busqueda 'activamos el menu de busqueda y de alta
      Me.Vaciar_campos
      Me.listadonick.Enabled = True
      Me.Actualiza_ComboNicks
      Me.Desactivar_textbox
Case "BAJA" 'funciona Si MODO es igual a baja elimina el Campo de la Base de Datos
     listadonick.Clear
     TablaUsuarios.Delete
     Me.Vaciar_campos
     Actualiza_ComboNicks
     Me.Activar_MENU_Alta_Busqueda
 Case "ACTUALIZAR" 'Si MODO es igual Actualizar, modifica los campos de la base de datos
     If Not validar Then
        Exit Sub
     End If
     TablaUsuarios.Edit
     copiarcampos
     TablaUsuarios.Update
     Me.Vaciar_campos
     Actualiza_ComboNicks 'actualiza los nick de la base en el combobox
     Me.Desactivar_textbox
     Me.listadonick.Enabled = True
End Select
Me.Ocultar_Botones
End Sub
'**************** FIN DE LAS FUNCIONES DE ALTA BAJA ACTUALIZACION ******
'++++++ Si pulsamos el boton cancelar +++++++++++++++++
Private Sub Cancelar_Click()
 MODO = ""
 Me.Desactivar_textbox
 Me.listadonick.Enabled = True
 Me.Actualiza_ComboNicks
 Me.Ocultar_Botones
 Me.Desactivar_MENU_Mod_Eliminar
 Me.Activar_MENU_Alta_Busqueda
End Sub
'+++++++++++++++ Comprueba  los campos ++++++++++
Function validar() As Boolean
 If nick.Text = "" Then
   MsgBox "El campo Nick es obligatorio", vbCritical
   nick.SetFocus
   validar = False
  Else
   validar = True
  End If
End Function
'+++++++++++++++Copiar campos a la base de  datos +++++++++++++++++++++++
Function copiarcampos()
  TablaUsuarios.Fields("nick") = nick.Text
  TablaUsuarios.Fields("nombre") = nombre.Text
  TablaUsuarios.Fields("correo") = correo.Text
  TablaUsuarios.Fields("ciudad") = ciudad.Text
  TablaUsuarios.Fields("añonacimiento") = año.Text
  TablaUsuarios.Fields("telefono") = telefono.Text
  TablaUsuarios.Fields("web") = web.Text
  If dialogo.FileName <> "" Then
   TablaUsuarios.Fields("foto") = nick.Text & ".jpg"
  Else
   TablaUsuarios.Fields("foto") = ""
  End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++Modificar un registro++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++ funciona +++Actualiza el listado de Nick del combobox +++
Function Actualiza_ComboNicks()
Dim i As Integer
Dim un As String
listadonick.Clear 'vaciamos el texto del combobox
'ahora lo llenamos con los nicks de la base de datos
TablaUsuarios.Index = "clave"
TablaUsuarios.MoveFirst
For i = 0 To (TablaUsuarios.RecordCount - 1)
  listadonick.AddItem (TablaUsuarios.Fields("Nick"))
  TablaUsuarios.MoveNext
Next
End Function
'+++++ Funciona BUSCAR NICK Y MOSTRARLO  +++++
Private Sub btn_buscar_Click()
   If textBox_buscar_nick.Text <> "" Then
      buscar_nicks textBox_buscar_nick.Text
   End If
End Sub
Function buscar_nicks(nicks As String)
   Me.Vaciar_campos
   TablaUsuarios.Seek "=", nicks
   On Error GoTo fallo ' capturamos las posibles excepciones
    nick.Text = TablaUsuarios.Fields("Nick")
    If TablaUsuarios.Fields("Nombre") <> "" Then nombre.Text = TablaUsuarios.Fields("Nombre")
    If TablaUsuarios.Fields("Correo") <> "" Then correo.Text = TablaUsuarios.Fields("Correo")
    If TablaUsuarios.Fields("Ciudad") <> "" Then ciudad.Text = TablaUsuarios.Fields("Ciudad")
    If TablaUsuarios.Fields("Añonacimiento") <> "" Then año.Text = TablaUsuarios.Fields("Añonacimiento")
    edad.Text = Year(Date) - año.Text
    If TablaUsuarios.Fields("telefono") <> "" Then telefono.Text = TablaUsuarios.Fields("telefono")
    If TablaUsuarios.Fields("web") <> "" Then web.Text = TablaUsuarios.Fields("web")
    Me.Activar_MENU_Mod_Eliminar
   'excepcion de la foto
   On Error GoTo sin_fichero
        Me.foto.Picture = LoadPicture(ruta_fotos & "\" & TablaUsuarios.Fields("foto"))
        Exit Function
sin_fichero:
   Me.foto.Picture = LoadPicture("")
   Exit Function
fallo:
   Me.Desactivar_MENU_Mod_Eliminar
   MsgBox "El Nick no existe"
End Function
'****** FIN DE BUSCAR *******************
'****** MODIFICAR FOTO ******************
Private Sub foto_Click() ' permite modificar la foto
 If nick.Text <> "" Then
  dialogo.ShowOpen
   If dialogo.FileName <> "" And dialogo.FileName <> "*.*" Then
     foto.Picture = LoadPicture(dialogo.FileName)
     foto.Stretch = True
On Error GoTo Fichero_No_exite 'capturamos la excepcion
      If FileLen(ruta_fotos & "\" & nick.Text & ".jpg") = 0 Then ' comprueba si existe el fichero
Fichero_No_exite:       'si no existe copia el fichero
       FileCopy dialogo.FileName, ruta_fotos & "\" & nick.Text & ".jpg" ' copia la imagen ala carpeta fotos
      End If
   End If
 End If
End Sub
'*******************************************
Private Sub listadonick_Click()
    Me.Vaciar_campos
    TablaUsuarios.Seek "=", listadonick.Text
    nick.Text = TablaUsuarios.Fields("Nick")
    Me.Activar_MENU
    If Not IsNull(TablaUsuarios.Fields("Nombre")) Then nombre.Text = TablaUsuarios.Fields("Nombre")
    If Not IsNull(TablaUsuarios.Fields("Correo")) Then correo.Text = TablaUsuarios.Fields("Correo")
    If Not IsNull(TablaUsuarios.Fields("Ciudad")) Then ciudad.Text = TablaUsuarios.Fields("Ciudad")
    If Not IsNull(TablaUsuarios.Fields("Añonacimiento")) Then año.Text = TablaUsuarios.Fields("Añonacimiento")
    edad.Text = Year(Date) - año.Text
    If Not IsNull(TablaUsuarios.Fields("telefono")) Then telefono.Text = TablaUsuarios.Fields("telefono")
    If Not IsNull(TablaUsuarios.Fields("web")) Then web.Text = TablaUsuarios.Fields("web")
    If Not IsNull(TablaUsuarios.Fields("foto")) Then
    On Error GoTo sigue
       foto.Picture = LoadPicture(ruta_fotos & "\" & TablaUsuarios.Fields("foto"))
    Exit Sub
sigue:
     foto.Picture = LoadPicture("")
    End If
End Sub
'********** Activar los campos para modificacion *****************
Function Activar_textbox()
 nick.Enabled = True
 nombre.Enabled = True
 correo.Enabled = True
 ciudad.Enabled = True
 año.Enabled = True
 telefono.Enabled = True
 web.Enabled = True
 foto.Enabled = True
End Function
'********** Desactivar los Objectos para evitar modificarlos *********
Function Desactivar_textbox()
 nick.Enabled = False
 nombre.Enabled = False
 correo.Enabled = False
 ciudad.Enabled = False
 año.Enabled = False
 telefono.Enabled = False
 web.Enabled = False
 foto.Enabled = False
End Function
'************ Poner todos los objectos a vacio *****************************
Function Vaciar_campos()
 nick.Text = ""
 nombre.Text = ""
 correo.Text = ""
 ciudad.Text = ""
 edad.Text = ""
 telefono.Text = ""
 web.Text = "http://"
 foto.Picture = LoadPicture("")
 End Function
 '************** mostrar botones de aceptar y cancelar
 Function Mostrar_botones()
  Aceptar.Visible = True
  Cancelar.Visible = True
 End Function
 '******** ocultar botones
 Function Ocultar_Botones()
  Aceptar.Visible = False
  Cancelar.Visible = False
 End Function
'************** Activar/Desactivar menu Alta y Busqueda ********************************
Function Activar_MENU_Alta_Busqueda()
 Me.alta.Enabled = True
 Me.busqueda.Enabled = True
End Function
Function Desactivar_MENU_Alta_busqueda()
 Me.alta.Enabled = False
 Me.busqueda.Enabled = False
End Function
'************** Activar/Desactivar menu Modificar y Eliminar ********************************
Function Desactivar_MENU_Mod_Eliminar()
 Me.Modificar.Enabled = False
 Me.eliminar.Enabled = False
End Function
Function Activar_MENU_Mod_Eliminar()
 Me.Modificar.Enabled = True
 Me.eliminar.Enabled = True
End Function
'**************** Activar/desactivar TODO el MENU **************************
Function Activar_MENU()
 Me.alta.Enabled = True
 Me.busqueda.Enabled = True
 Me.Modificar.Enabled = True
 Me.eliminar.Enabled = True
End Function
Function Desactivar_MENU()
 Me.alta.Enabled = False
 Me.busqueda.Enabled = False
 Me.Modificar.Enabled = False
 Me.eliminar.Enabled = False
End Function
'+++++++++++++ Salimos de la aplicacion ++++++++++++++++++++++++++++++++++++++
Private Sub Command1_Click()
 TablaUsuarios.Close
 UsuariosIRC.Close
 End
End Sub
'******************* ayuda ************************
Private Sub Acercade_Click()
  Dim a As New frmAbout
  a.Show
End Sub
Private Sub textBox_buscar_nick_Change()
 Me.btn_buscar.Enabled = True
End Sub
