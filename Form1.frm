VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Cambiador de Nombre de Descargas"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9945
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9915
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   5865
      Left            =   1920
      OLEDropMode     =   1  'Manual
      TabIndex        =   11
      Top             =   3210
      Width           =   7755
      Begin VB.ListBox ListLinks 
         Height          =   2205
         ItemData        =   "Form1.frx":08CA
         Left            =   180
         List            =   "Form1.frx":08CC
         OLEDragMode     =   1  'Automatic
         TabIndex        =   15
         Top             =   1860
         Width           =   7305
      End
      Begin VB.ListBox ListNames 
         Appearance      =   0  'Flat
         Height          =   5070
         IntegralHeight  =   0   'False
         ItemData        =   "Form1.frx":08CE
         Left            =   90
         List            =   "Form1.frx":08D0
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   12
         ToolTipText     =   "drag"
         Top             =   180
         Width           =   7575
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   3420
         TabIndex        =   14
         Top             =   5460
         Width           =   945
      End
      Begin VB.TextBox TextRegExpBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "(mother)|(mom)|(madre)"
         ToolTipText     =   "RegExp de búsqueda en la DB"
         Top             =   5460
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5865
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   150
      Width           =   7755
      Begin VB.CheckBox CheckNoRevisarUltimosRenombrados 
         Caption         =   "Check1"
         Height          =   195
         Left            =   5640
         TabIndex        =   30
         ToolTipText     =   "Checkeado => No Revisar Ultimos Renombrados, es para las descargas que repiten el nombre siempre"
         Top             =   240
         Width           =   195
      End
      Begin VB.PictureBox PictureDirectorioDeCambioDeNombre 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   4980
         OLEDropMode     =   1  'Manual
         Picture         =   "Form1.frx":08D2
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   29
         ToolTipText     =   "Directorio donde se cambia el nombre del archivo automáticamente"
         Top             =   210
         Width           =   540
      End
      Begin VB.TextBox TextFilePath 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   28
         Text            =   "FILE_PATH"
         Top             =   5460
         Width           =   7575
      End
      Begin VB.TextBox TextNewName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   27
         Text            =   "NEW_NAME"
         Top             =   5160
         Width           =   7575
      End
      Begin VB.TextBox TextFilePath 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   26
         Text            =   "FILE_PATH"
         Top             =   4740
         Width           =   7575
      End
      Begin VB.TextBox TextNewName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   25
         Text            =   "NEW_NAME"
         Top             =   4440
         Width           =   7575
      End
      Begin VB.TextBox TextContadorDeLinks 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   24
         Text            =   "0"
         ToolTipText     =   "Cantidad de Links agregados"
         Top             =   510
         Width           =   765
      End
      Begin VB.CommandButton cmdCrearIndicesAlfabeticosYMover 
         Caption         =   "a,b...etc"
         Height          =   285
         Left            =   6900
         OLEDropMode     =   1  'Manual
         TabIndex        =   23
         ToolTipText     =   "Crear SubDirs alfabeticos en la DB y Mover Archivos"
         Top             =   510
         Width           =   765
      End
      Begin VB.PictureBox PictureDirectorioDeBaseDeDatos 
         Height          =   540
         Left            =   3840
         OLEDropMode     =   1  'Manual
         Picture         =   "Form1.frx":0D14
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   22
         ToolTipText     =   "Directorio donde se guarda la Base de Datos"
         Top             =   210
         Width           =   540
      End
      Begin VB.PictureBox PictureArchivarLink 
         AutoRedraw      =   -1  'True
         Height          =   540
         Left            =   150
         OLEDropMode     =   1  'Manual
         Picture         =   "Form1.frx":1156
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   21
         ToolTipText     =   "Drag Drop para agregar el URL a la base de URLs"
         Top             =   210
         Width           =   540
      End
      Begin VB.TextBox TextAnteponer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         OLEDropMode     =   1  'Manual
         TabIndex        =   20
         ToolTipText     =   "Nombre que se antepone a cualquier nuevo nombre automáticamente"
         Top             =   210
         Width           =   2685
      End
      Begin VB.PictureBox PictureDirectorioDeDescargas 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   4410
         OLEDropMode     =   1  'Manual
         Picture         =   "Form1.frx":1598
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   19
         ToolTipText     =   "Directorio donde se guarda la descarga y se Revisa para cambiar de nombre"
         Top             =   210
         Width           =   540
      End
      Begin VB.CommandButton cmdTimerRenombrarYVerificarDescargas 
         Caption         =   "Timer"
         Height          =   285
         Left            =   6030
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Inicia Revisión para cambiar de nombre cada 1 segundo"
         Top             =   360
         Width           =   765
      End
      Begin VB.CommandButton cmdRenombrar 
         Caption         =   "----->"
         Height          =   285
         Left            =   6900
         OLEDropMode     =   1  'Manual
         TabIndex        =   17
         ToolTipText     =   "Cambiar de nombre: de el de Arriba reemplaza al de Abajo"
         Top             =   180
         Width           =   765
      End
      Begin VB.CheckBox CheckAnteponer 
         Caption         =   "Check1"
         Height          =   255
         Left            =   3570
         OLEDropMode     =   1  'Manual
         TabIndex        =   16
         ToolTipText     =   "Checkeado abilita anteponer el texto"
         Top             =   210
         Width           =   225
      End
      Begin VB.Timer TimerRenombrarDescargas 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5160
         Top             =   1350
      End
      Begin VB.Timer TimerVerificarDescargas 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   6090
         Top             =   2430
      End
      Begin VB.TextBox TextNewName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   10
         Text            =   "NEW_NAME"
         Top             =   840
         Width           =   7575
      End
      Begin VB.TextBox TextFilePath 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   9
         Text            =   "FILE_PATH"
         Top             =   1140
         Width           =   7575
      End
      Begin VB.TextBox TextFilePath 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   8
         Text            =   "FILE_PATH"
         Top             =   1860
         Width           =   7575
      End
      Begin VB.TextBox TextNewName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Text            =   "NEW_NAME"
         Top             =   1560
         Width           =   7575
      End
      Begin VB.TextBox TextFilePath 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Text            =   "FILE_PATH"
         Top             =   2580
         Width           =   7575
      End
      Begin VB.TextBox TextNewName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         Text            =   "NEW_NAME"
         Top             =   2280
         Width           =   7575
      End
      Begin VB.TextBox TextNewName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         Text            =   "NEW_NAME"
         Top             =   3000
         Width           =   7575
      End
      Begin VB.TextBox TextFilePath 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         Text            =   "FILE_PATH"
         Top             =   3300
         Width           =   7575
      End
      Begin VB.TextBox TextNewName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Text            =   "NEW_NAME"
         Top             =   3720
         Width           =   7575
      End
      Begin VB.TextBox TextFilePath 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Text            =   "FILE_PATH"
         Top             =   4020
         Width           =   7575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOACTIVATE = &H10

Const SWP_SHOWWINDOW = &H40
Const SWP_DRAWFRAME = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4

Private Sub CheckAnteponer_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call OLEDragDropGeneral(Data)
End Sub

Private Sub cmdBuscar_Click()
Dim i As Long
Dim j As Long
Dim miFolders As Folders
Dim miFolder As Folder
Dim miFile As File

On Error GoTo Falla

    VBRE.Pattern = TextRegExpBusqueda.Text
    ListLinks.Clear
    ListNames.Clear
    
    If fso.FolderExists(pathBaseDeDatos) Then
        
        Set miFolder = fso.GetFolder(pathBaseDeDatos)
        For Each miFile In miFolder.Files
            If VBRE.Test(miFile.Name) Then
                If miFile.Attributes <> 36 Then
                    ListLinks.AddItem miFile.Path
                    ListNames.AddItem miFile.Name
                End If
            End If
        Next
        
        Set miFolders = fso.GetFolder(pathBaseDeDatos).SubFolders
        For Each miFolder In miFolders
            For Each miFile In miFolder.Files
                If VBRE.Test(miFile.Name) Then
                    If miFile.Attributes <> 36 Then
                        ListLinks.AddItem miFile.Path
                        ListNames.AddItem miFile.Name
                    End If
                End If
            Next
        Next
        
    End If

    Exit Sub
    
Falla:
    MsgBox "Error en: " & Err.Description, vbOKOnly, "Error"
End Sub

Private Sub cmdConvertirALinks_Click()

End Sub

Private Sub cmdCrearIndicesAlfabeticosYMover_Click()
'crea directorios con los indices alfabéticos y mueve a ellos los links que haya en el directorio
Dim i As Long
Dim losArchivos As Files
Dim elArchivo As File
Dim elIndice As String
        
    VBRE.Pattern = "^."

    If vbYes = MsgBox("Crear sub Indices en: " & pathBaseDeDatos & "?", vbYesNo, "Crear Sub Indices") Then
        For i = 97 To 122
            If Not fso.FolderExists(pathBaseDeDatos & "\" & Chr(i)) Then
                fso.CreateFolder (pathBaseDeDatos & "\" & Chr(i))
            End If
        Next
    End If
    
    If vbYes = MsgBox("Mover Archivos en: " & pathBaseDeDatos & "?", vbYesNo, "Mover Archivos") Then
        Set losArchivos = fso.GetFolder(pathBaseDeDatos).Files
        
        For Each elArchivo In losArchivos
            elIndice = VBRE.Execute(elArchivo.Name)(0) 'el índice del archivo
            If fso.FolderExists(pathBaseDeDatos & "\" & elIndice) Then
                On Error Resume Next
                fso.MoveFile elArchivo.Path, pathBaseDeDatos & "\" & elIndice & "\" & elArchivo.Name
            End If
        Next
        Exit Sub
    End If


    If vbYes = MsgBox("Corregir Archivos en: " & pathBaseDeDatos & "?", vbYesNo, "Mover Archivos") Then
        Set losArchivos = fso.GetFolder(pathBaseDeDatos).Files
        
        For Each elArchivo In losArchivos
            corregirElArchivo elArchivo.Path
        Next
        For i = 97 To 122
            Set losArchivos = fso.GetFolder(pathBaseDeDatos & "\" & Chr(i)).Files
            For Each elArchivo In losArchivos
                corregirElArchivo elArchivo.Path
            Next
        Next
    End If

End Sub


Public Sub corregirElArchivo(filePath As String)
On Error GoTo Falla
Dim miTexto As String
    'corrijo lo que esté mal
    miTexto = fso.GetFile(filePath).OpenAsTextStream(ForReading).ReadAll
    
    VBRE.Pattern = "\.org\.com"
    
    If VBRE.Test(miTexto) Then
        miTexto = VBRE.Replace(miTexto, ".org")
        Call fso.GetFile(filePath).OpenAsTextStream(ForWriting).Write(miTexto)
    End If
    Exit Sub
    
Falla:
    MsgBox "Error al corregir el archivo: " & Err.Description, vbCritical + vbOKOnly, "Error"
    
End Sub



Private Sub cmdCrearIndicesAlfabeticosYMover_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call OLEDragDropGeneral(Data)
End Sub

Private Sub cmdRenombrar_Click()
Dim i As Integer
    For i = 0 To TextNewName.UBound
        Call Renombrar(i)
    Next
End Sub

Private Sub Renombrar(Index As Integer)
Dim i As Integer
Dim nombreFinal As String
Dim nombreBase As String
VBRE.Pattern = "([^\.]+)(\.\w+)$"

    nombreBase = TextNewName(Index).Text
    nombreArchivo = fso.GetFile(TextFilePath(Index).Text).Name
    decinencia = VBRE.Execute(nombreArchivo)(0).SubMatches(1)
    
    
    nombreFinal = nombreBase
    For i = 1 To 100
        If fso.FileExists(pathDescargas & "\" & nombreFinal & decinencia) Then
            nombreFinal = nombreBase & " " & CInt(i) & " "
        Else
            Exit For
        End If
    Next
    
    fso.GetFile(TextFilePath(Index).Text).Name = nombreFinal & decinencia

    DoEvents

End Sub

Private Sub cmdRenombrar_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call OLEDragDropGeneral(Data)
End Sub

Private Sub cmdTimerRenombrarYVerificarDescargas_Click()
    
    If TimerRenombrarDescargas.Enabled Then
        TimerRenombrarDescargas.Enabled = False
        TimerVerificarDescargas.Enabled = False
        cmdTimerRenombrarYVerificarDescargas.BackColor = &H8000000F
    Else
        TimerRenombrarDescargas.Enabled = True
        TimerVerificarDescargas.Enabled = True
        cmdTimerRenombrarYVerificarDescargas.BackColor = vbRed
    End If
    
End Sub

Private Sub cmdTimerRenombrarYVerificarDescargas_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call OLEDragDropGeneral(Data)
End Sub

Private Sub Form_DblClick()
    Form1.Width = Frame1.Width + 80
    Form1.Height = Frame1.Height + 500
End Sub

Private Sub Form_Load()
    bCargando = True
    VBRE.Global = True
    VBRE.IgnoreCase = True
    VBRE.MultiLine = True
    
    Frame1.Top = 0
    Frame1.Left = 0
    Frame2.Top = 0
    Frame2.Left = -20000
    
    ListLinks.Visible = False
    
    Form1.Width = Frame1.Width + 80
    Form1.Height = Frame1.Height + 500
    Form1.Left = CLng(GetSetting("cambiarNombreDescarga", "Load", "Left", Screen.Width - Form1.Width / 2))
    Form1.Top = CLng(GetSetting("cambiarNombreDescarga", "Load", "Top", Screen.Height - Form1.Height / 2))
    
    pathDescargas = GetSetting("cambiarNombreDescarga", "Load", "pathDescargas", "C:\Documents and Settings\Pablo\Mis documentos\Descargas")
    pathBaseDeDatos = GetSetting("cambiarNombreDescarga", "Load", "pathBaseDeDatos", "C:\Documents and Settings\Pablo\Escritorio\Nueva carpeta")
    pathDirectorioDeCambioDeNombre = GetSetting("cambiarNombreDescarga", "Load", "pathDirectorioDeCambioDeNombre", "C:\Documents and Settings\Pablo\Escritorio\Nueva carpeta")
       
    
    If fso.FolderExists(pathBaseDeDatos) Then
        Me.Caption = pathBaseDeDatos
    End If
    
    
    PictureDirectorioDeBaseDeDatos.ToolTipText = "Drag Drop para Directorio de la Base de Datos: " & pathBaseDeDatos
    PictureDirectorioDeDescargas.ToolTipText = "Drag Drop para Directorio de la Descarga: " & pathDescargas
    PictureDirectorioDeCambioDeNombre.ToolTipText = "Drag Drop para Directorio de Cambio de Nobres: " & pathDirectorioDeCambioDeNombre
    
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 100, 100, 100, 100, SWP_DRAWFRAME Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
    
    Dim i As Long
    Dim auxStr As String
    If fso.FileExists(App.Path & "\linksNombres.dat") Then
        Open App.Path & "\linksNombres.dat" For Input As #1
        ListNames.Clear
        i = 0
            Do Until EOF(1)
                Line Input #1, auxStr
                ListNames.AddItem auxStr, i
                i = i + 1
            Loop
        Close #1
    End If
    If fso.FileExists(App.Path & "\linksURLs.dat") Then
        Open App.Path & "\linksURLs.dat" For Input As #1
        ListLinks.Clear
        i = 0
            Do Until EOF(1)
                Line Input #1, auxStr
                ListLinks.AddItem auxStr, i
                i = i + 1
            Loop
        Close #1
    End If
    
    bCargando = False
  
End Sub

Private Sub Form_Resize()
    If bCargando Then Exit Sub
    Frame1.Width = Form1.Width - 80
    Frame1.Height = Form1.Height - 500
    Frame2.Width = Form1.Width - 80
    Frame2.Height = Form1.Height - 500
 
    ListNames.Top = 180
    ListNames.Left = 90
    ListNames.Width = Frame2.Width - 2 * 90
    ListNames.Height = Frame2.Height - TextRegExpBusqueda.Height - 270
    TextRegExpBusqueda.Left = ListNames.Left
    TextRegExpBusqueda.Top = ListNames.Top + ListNames.Height + 90
    cmdBuscar.Top = TextRegExpBusqueda.Top
    cmdBuscar.Left = TextRegExpBusqueda.Left + TextRegExpBusqueda.Width + 180
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SaveSetting("cambiarNombreDescarga", "Load", "Left", Me.Left)
    Call SaveSetting("cambiarNombreDescarga", "Load", "top", Me.Top)

    Call SaveSetting("cambiarNombreDescarga", "Load", "pathDescargas", pathDescargas)
    Call SaveSetting("cambiarNombreDescarga", "Load", "pathBaseDeDatos", pathBaseDeDatos)
    Call SaveSetting("cambiarNombreDescarga", "Load", "pathDirectorioDeCambioDeNombre", pathDirectorioDeCambioDeNombre)

    Dim i As Long
    Dim auxStr As String
    If Not fso.FileExists(App.Path & "\linksNombres.dat") Then
        fso.CreateTextFile (App.Path & "\linksNombres.dat")
    End If
    
    Open App.Path & "\linksNombres.dat" For Output As #1
        For i = 0 To ListNames.ListCount - 1
            Print #1, ListNames.List(i)
        Next
    Close #1

    If Not fso.FileExists(App.Path & "\linksURLs.dat") Then
        fso.CreateTextFile (App.Path & "\linksURLs.dat")
    End If
    
    Open App.Path & "\linksURLs.dat" For Output As #1
        For i = 0 To ListLinks.ListCount - 1
            Print #1, ListLinks.List(i)
        Next
    Close #1



End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call OLEDragDropGeneral(Data)

End Sub


Private Sub Frame1_DblClick()
    Frame1.Left = -20000
    Frame2.Left = 0
End Sub

Private Sub Frame1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call OLEDragDropGeneral(Data)
       
End Sub

Private Sub Frame2_DblClick()
    Frame2.Left = -20000
    Frame1.Left = 0
End Sub

Private Sub Frame2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call OLEDragDropGeneral(Data)
       
End Sub

Private Sub ListNames_Click()
'Pone en el clipboard
Dim elURL As String
Dim elPath As String

    elPath = ListLinks.List(ListNames.ListIndex)
    
    'extaer el URL y ponerlo en el Clipboard
        elURL = fso.OpenTextFile(elPath).ReadAll()
        VBRE.MultiLine = True
        VBRE.Pattern = "URL=(.+)" & vbCrLf '& "IconFile"
        elURL = VBRE.Execute(elURL)(0).SubMatches(0)
        Clipboard.Clear
        Clipboard.SetText elURL
End Sub

Private Sub ListNames_DblClick()
'Saca de las listas, setea atributo e icono, pone en el clipboard
Dim elURL As String
Dim elPath As String
    
    elPath = ListLinks.List(ListNames.ListIndex)
    
    Call cambiarAtributoEIcono(elPath)

    'extaer el URL y ponerlo en el Clipboard
        elURL = fso.OpenTextFile(elPath).ReadAll()
        VBRE.MultiLine = True
        VBRE.Pattern = "URL=(.+)" & vbCrLf & "IconFile"
        elURL = VBRE.Execute(elURL)(0).SubMatches(0)
        
        Clipboard.Clear
        Clipboard.SetText elURL

    'sacar de las listas
    ListLinks.RemoveItem (ListNames.ListIndex)
    ListNames.RemoveItem (ListNames.ListIndex)
    
End Sub

Private Sub ListNames_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PictureArchivarLink_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub ListNames_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Call Data.SetData(ListLinks.List(ListNames.ListIndex), vbCFText)
End Sub

Private Sub PictureArchivarLink_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Si es un link desde el IE se agrega a la lista

    bSiEsUnLink = False
    
    Call OLEDragDropGeneral(Data)
    
'    If bSiEsUnLink Then
'        If noEstaEnLaLista(pathAlUltimoLinkAgregado) Then
'            ListLinks.AddItem pathAlUltimoLinkAgregado
'            ListNames.AddItem nombreDelUltimoLinkAgregado
'        End If
'    End If
    
    Dim indexLista  As Integer
    
    If bSiEsUnLink Then
        
        indexLista = indexEnLaLista(pathAlUltimoLinkAgregado)
        
        If indexLista > -1 Then '=> Ya está en la lista
            ListLinks.RemoveItem (indexLista)   'Lo saco de la lista
            ListNames.RemoveItem (indexLista)
        End If
        
        ListLinks.AddItem pathAlUltimoLinkAgregado 'Lo agrego en último lugar
        ListNames.AddItem nombreDelUltimoLinkAgregado
              
    End If
    
End Sub

Private Function noEstaEnLaLista(elPath As String) As Boolean
Dim i As Integer

    noEstaEnLaLista = True
    
    For i = 0 To ListLinks.ListCount - 1
        If ListLinks.List(i) = elPath Then
            noEstaEnLaLista = False
            Exit Function
        End If
    Next

End Function

Private Function indexEnLaLista(elPath As String) As Integer
Dim i As Integer
    
    indexEnLaLista = -1
     
    For i = 0 To ListLinks.ListCount - 1
        If ListLinks.List(i) = elPath Then
            indexEnLaLista = i
            Exit Function
        End If
    Next

End Function

Private Sub PictureDirectorioDeBaseDeDatos_Click()
    Me.Caption = pathBaseDeDatos
End Sub

Private Sub PictureDirectorioDeBaseDeDatos_DblClick()
    ShellExecute Me.hwnd, "Open", pathBaseDeDatos, vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub PictureDirectorioDeDescargas_Click()
    Me.Caption = pathDescargas
End Sub

Private Sub PictureDirectorioDeDescargas_DblClick()
    ShellExecute Me.hwnd, "Open", pathDescargas, vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub PictureDirectorioDeCambioDeNombre_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Se guarda el path al Directorio donde se están buscando/haciendo las descargas
'se puede arrastrar tanto un directorio como un archivo
Dim miTexto As String
Dim i As Integer
    
    If Not Data.GetFormat(vbCFFiles) Then
        Call OLEDragDropGeneral(Data)
        Exit Sub 'por Ej. si arrastro una imagen u otra cosa rara
    End If

    'For i = 1 To Data.Files.Count
        'dataDragDrop = Data.Files(i)
        dataDragDrop = Data.Files(1)
        If fso.FileExists(dataDragDrop) Then
            pathDirectorioDeCambioDeNombre = fso.GetParentFolderName(dataDragDrop)
            Me.Caption = pathDirectorioDeCambioDeNombre
        End If
        If fso.FolderExists(dataDragDrop) Then
            pathDirectorioDeCambioDeNombre = dataDragDrop
            Me.Caption = pathDirectorioDeCambioDeNombre
        End If
    'Next
    
    PictureDirectorioDeCambioDeNombre.ToolTipText = "Directorio de Cambios de Nombre: " & pathDirectorioDeCambioDeNombre
    
End Sub



Private Sub PictureDirectorioDeDescargas_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Se guarda el path al Directorio donde se están buscando/haciendo las descargas
'se puede arrastrar tanto un directorio como un archivo
Dim miTexto As String
Dim i As Integer
    
    If Not Data.GetFormat(vbCFFiles) Then
        Call OLEDragDropGeneral(Data)
        Exit Sub 'por Ej. si arrastro una imagen u otra cosa rara
    End If

    'For i = 1 To Data.Files.Count
        'dataDragDrop = Data.Files(i)
        dataDragDrop = Data.Files(1)
        If fso.FileExists(dataDragDrop) Then
            pathDescargas = fso.GetParentFolderName(dataDragDrop)
            Me.Caption = pathDescargas
        End If
        If fso.FolderExists(dataDragDrop) Then
            pathDescargas = dataDragDrop
            Me.Caption = pathDescargas
        End If
    'Next
    
    PictureDirectorioDeDescargas.ToolTipText = "Directorio de Descargas: " & pathDescargas
    
End Sub

Private Sub PictureDirectorioDeBaseDeDatos_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Se guarda el path al Directorio donde se están buscando los links
'se puede arrastrar tanto un directorio como un archivo
Dim miTexto As String
Dim i As Integer
    
    If Not Data.GetFormat(vbCFFiles) Then
        Call OLEDragDropGeneral(Data)
        Exit Sub 'por Ej. si arrastro una imagen u otra cosa rara
    End If

    'For i = 1 To Data.Files.Count
        'dataDragDrop = Data.Files(i)
        dataDragDrop = Data.Files(1)
        If fso.FileExists(dataDragDrop) Then
            pathBaseDeDatos = fso.GetParentFolderName(dataDragDrop)
            Me.Caption = pathBaseDeDatos
        End If
        If fso.FolderExists(dataDragDrop) Then
            pathBaseDeDatos = dataDragDrop
            Me.Caption = pathBaseDeDatos
        End If
    'Next
    
    PictureDirectorioDeBaseDeDatos.ToolTipText = "Directorio donde se guarda la Base de Datos: " & pathBaseDeDatos
    
End Sub

Private Sub TextAnteponer_DblClick()
    TextAnteponer.Text = ""
    CheckAnteponer.Value = vbUnchecked
End Sub

Private Sub TextAnteponer_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'En este text arrastro el Path al archivo que voy a renombrar
    Call OLEDragDropGeneral(Data)
End Sub

Private Sub TextContadorDeLinks_DblClick()
    TextContadorDeLinks.Text = "0"
    contadorDeLinks = 0
End Sub

Private Sub TextContadorDeLinks_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call OLEDragDropGeneral(Data)
End Sub

Private Sub TextNewName_DblClick(Index As Integer)
    TextNewName(Index).BackColor = vbWhite
End Sub

Private Sub TextNewName_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'En este text arrastro un Link, el URL o un texto seleccionado desde el Explorador
    If Data.GetFormat(vbCFFiles) Then
        dataDragDrop = Data.Files(1) 'por Ej. si arrastro desde el ListView del WE
    ElseIf Data.GetFormat(vbCFText) Then
        dataDragDrop = Data.GetData(vbCFText) 'por Ej. si arrastro desde el URL del IE o el WE
    Else
        Exit Sub 'por Ej. si arrastro una imagen u otra cosa rara
    End If
    
    VBRE.Pattern = "(https?://)|(\.url$)"
    If VBRE.Test(dataDragDrop) Then 'Es un Link
        Call OLEDragDropGeneral(Data)
        Exit Sub
    End If
    
    dataDragDrop = Replace(dataDragDrop, Chr(9), "")
    dataDragDrop = Replace(dataDragDrop, Chr(10), "")
    dataDragDrop = Replace(dataDragDrop, Chr(11), "")
    dataDragDrop = Replace(dataDragDrop, Chr(12), "")
    dataDragDrop = Replace(dataDragDrop, Chr(13), "")
    
    
    'VBRE.Pattern = ".+/([a-zA-Z0-9_]+)(\.\w+)?" 'Para Link y Texto selecccionado
    VBRE.Pattern = "([a-zA-Z0-9_\s]+)" 'Para Link y Texto selecccionado
    TextNewName(Index).Text = VBRE.Replace(dataDragDrop, "$1")
    
    TextNewName(Index).Text = Replace(TextNewName(Index).Text, "#_tabDownload", "") 'Si era un URL
    
    VBRE.Pattern = "\\|\||\/|\""|\*|:|<|>|\?"
    TextNewName(Index).Text = VBRE.Replace(TextNewName(Index).Text, "_")
    
    TextNewName(Index).Text = Trim(TextNewName(Index).Text)
    
    If TextAnteponer.Text <> "" And CheckAnteponer.Value = vbChecked Then
        TextNewName(Index).Text = TextAnteponer.Text & " " & TextNewName(Index).Text
    End If
    
    TextNewName(Index).BackColor = vbWhite
    
End Sub


Private Sub TextFilePath_DblClick(Index As Integer)
    TextFilePath(Index).Text = ""
    TextNewName(Index).BackColor = vbWhite
End Sub

Private Sub TextFilePath_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'En este text arrastro el Path al archivo que voy a renombrar
    If Data.GetFormat(vbCFFiles) Then
        dataDragDrop = Data.Files(1) 'por Ej. si arrastro desde el ListView del WE
    ElseIf Data.GetFormat(vbCFText) Then
        dataDragDrop = Data.GetData(vbCFText) 'por Ej. si arrastro desde el URL del IE o el WE
    Else
        Exit Sub 'por Ej. si arrastro una imagen u otra cosa rara
    End If
    
    VBRE.Pattern = "(https?://)|(\.url$)"
    If VBRE.Test(dataDragDrop) Then 'Es un Link
        Call OLEDragDropGeneral(Data)
        Exit Sub
    End If
    
    TextFilePath(Index).Text = dataDragDrop
    
End Sub

Private Sub TimerRenombrarDescargas_Timer()
'cambia el nombre del archivo si el fileSize es <> 0
On Error GoTo Falla
Dim Index As Integer
Dim fso1 As FileSystemObject
Static ultimosRenombrados(0 To 4) As String
Dim i As Integer, j As Integer
Set fso1 = New Scripting.FileSystemObject
    
    'verifico que no hayan vuelto a aparecer como files los renombrados vacios
    If Not CheckNoRevisarUltimosRenombrados.Value = vbChecked Then
        For i = 0 To 4
            If fso1.FileExists(ultimosRenombrados(i)) Then
                If fso1.GetFile(ultimosRenombrados(i)).Size = 0 Then
                    fso1.GetFile(ultimosRenombrados(i)).Delete
                End If
            End If
        Next
        
        'verifico que no se hayan vuelto a escribir los renombrados vacios
        For j = 0 To 4
            If TextNewName(j).Text = "" And TextFilePath(j).Text <> "" Then
                For i = 0 To 4
                    If TextFilePath(j).Text = ultimosRenombrados(i) Then
                        TextFilePath(j).Text = ""
                        TextNewName(j).BackColor = vbWhite
                    End If
                Next
            End If
        Next
    End If
    
    'busco los que renombraré
    For Index = 0 To TextNewName.UBound
        If TextNewName(Index).Text <> "" And TextNewName(Index).Text <> "NEW_NAME" Then
            If fso1.FileExists(TextFilePath(Index).Text) Then
                If fso1.GetFile(TextFilePath(Index).Text).Size > 1000 Then
                    
                    Call Renombrar(Index)
                                       
                    'Guardo los últimos renombrados en el arreglo
                    ultimosRenombrados(4) = ultimosRenombrados(3)
                    ultimosRenombrados(3) = ultimosRenombrados(2)
                    ultimosRenombrados(2) = ultimosRenombrados(1)
                    ultimosRenombrados(1) = ultimosRenombrados(0)
                    ultimosRenombrados(0) = TextFilePath(Index).Text
                    
                    TextNewName(Index).Text = ""
                    TextFilePath(Index).Text = ""
                    
                End If
            End If
        End If
    Next
    
    
    
Falla:
End Sub


Private Sub TimerVerificarDescargas_Timer()
'carga el archivo que está descargándose si el fileSize=0
Dim miFolder As Folder
Dim miFiles As Files
Dim miFile As File
Dim i As Integer
Dim j As Integer
Dim bYaEsta As Boolean
    

    If pathDescargas <> "" Then
        Set miFolder = fso.GetFolder(pathDescargas)
        Set miFiles = miFolder.Files
        For Each miFile In miFiles
        
            If miFile.Size = 0 Then
                bYaEsta = False
                For j = 0 To TextNewName.UBound
                    If TextFilePath(j).Text = miFile.Path Then
                        bYaEsta = True
                        Exit For
                    End If
                Next
                If Not bYaEsta Then
                    For j = 0 To TextNewName.UBound
                        If TextFilePath(j).Text = "FILE_PATH" Or TextFilePath(j).Text = "" Then
                            TextFilePath(j).Text = miFile.Path
                            TextNewName(j).BackColor = vbRed
                            Exit For
                        End If
                    Next
                End If
            End If
        Next
    End If
    
End Sub
