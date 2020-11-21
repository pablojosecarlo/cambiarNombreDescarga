Attribute VB_Name = "Module1"
Option Explicit

Public Const SW_SHOWNORMAL As Long = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public dataDragDrop As String  'información dejada caer del Drag and Drop
Public fso As New Scripting.FileSystemObject
Public VBRE As New VBScript_RegExp_55.RegExp
Public decinencia As String
Public nombreArchivo As String
Public pathDescargas As String
Public pathDirectorioDeCambioDeNombre As String
Public pathBaseDeDatos As String
Public contadorDeLinks As Integer

Public bSiEsUnLink As Boolean   'Estas variables son para que cuando se deja caer un link en el icono , además de agregarse a la basede datos, también se agrega a la lista de links selectos
Public pathAlUltimoLinkAgregado As String
Public nombreDelUltimoLinkAgregado As String

Public elMatch As Match
Public losMatch As MatchCollection

Public bCargando As Boolean
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'
'Const HWND_TOPMOST = -1
'Const HWND_NOTOPMOST = -2
'Const SWP_NOACTIVATE = &H10
'
'Const SWP_SHOWWINDOW = &H40
'Const SWP_DRAWFRAME = &H20
'Const SWP_NOMOVE = &H2
'Const SWP_NOSIZE = &H1
'Const SWP_NOZORDER = &H4



Sub ex1()
    'Set FEMM = CreateObject("Femm.ActiveFEMM")

    'Execute some FEMM command
    'Result = FEMM.call2femm("uo")
    
    'Make sheet1 active.  This selects the target sheet.
    'Commenting out this line means that the sheet
    'operated on is the current sheet.
    'Sheets("sheet1").Select
    

    'Display result in the sheet
    'Range("B10:B10").Value = CDbl(Result)

End Sub

Public Sub OLEDragDropGeneral(Data As DataObject)
'puedo arrastrar. . .
'Un Link desde el WE => .url => Se le cambia los atributos, el ícono y se manda el url al clipboard
'Un Link desde el IE => https => Se archiva
'Un texto seleccionado  => Se manda al área de texto a anteponer
'Un archivo de texto con URLs desde el WE => .txt => Se extraen del texto los URLs y se archivan

    If Data.GetFormat(vbCFFiles) Then
        'por Ej. si arrastro desde el ListView del WE o una imagen desde el IE, pues ahí me devuelve el path al archivo de la imagen
        VBRE.Pattern = "\.url$" 'lista de url desde el WE
        If VBRE.Test(Data.Files(1)) Then
            Call CambioDeIconoYTipoDeArchivo(Data)
            
            If Data.Files.Count = 1 Then
               'Se extrae el ulr del archivo y se pone en el clipboard
                Call ExtraerElURL(Data)
            End If
            
        End If
        VBRE.Pattern = "\.txt$" 'lista de archivos desde el WE con conteniendo links, que guarda como urls
        If VBRE.Test(Data.Files(1)) Then
            Call crearURLs(Data)
        End If
    ElseIf Data.GetFormat(vbCFText) Then
        'por Ej. si arrastro desde el URL del IE o el WE o un texto cualquiera seleccionado
        VBRE.Pattern = "https?://"
        If VBRE.Test(Data.GetData(vbCFText)) Then 'Link desde el IE
            Call ArchivarLink(Data.GetData(vbCFText))
            Exit Sub
        Else 'Es un texto
            Form1.TextAnteponer.Text = Data.GetData(vbCFText) 'texto desde el IE
            Form1.CheckAnteponer.Value = vbChecked
        End If
        
    Else
        Exit Sub 'por Ej. si arrastro una imagen u otra cosa rara
    End If
    
End Sub

Public Sub ExtraerElURL(Data As DataObject)
'Se extrae el url del archivo y se pone en el clipboard
Dim i As Integer
Dim elURL As String

    For i = 1 To Data.Files.Count
        dataDragDrop = Data.Files(i)
        If fso.FileExists(dataDragDrop) Then
                    
            elURL = fso.OpenTextFile(dataDragDrop).ReadAll()
            VBRE.MultiLine = True
            VBRE.Pattern = "URL=(.+)" & vbCrLf & "IconFile"
            elURL = VBRE.Execute(elURL)(0).SubMatches(0)
            
            Clipboard.Clear
            Clipboard.SetText elURL
            
        End If
    Next
    
End Sub


Public Sub CambioDeIconoYTipoDeArchivo(Data As DataObject)
'Se cambia el icono del URl y la propiedad de System
Dim i As Integer
    
    For i = 1 To Data.Files.Count
        dataDragDrop = Data.Files(i)
        If fso.FileExists(dataDragDrop) Then
        
            Call cambiarAtributoEIcono(dataDragDrop)
            
        End If
    Next
    
End Sub

Public Sub cambiarAtributoEIcono(laFile As String)
On Error GoTo Falla
Dim miTexto As String
Dim elURL As String
    'Cambio el atributo a SA
    fso.GetFile(laFile).Attributes = 36

    'Cambio el ícono
    miTexto = fso.GetFile(laFile).OpenAsTextStream(ForReading).ReadAll
    
    VBRE.Pattern = "^(URL=https?://.+)" & vbCrLf      'verfico que sea un URL
    If VBRE.Test(miTexto) Then 'saco el nombre del archivo para nombrar al URL
        elURL = VBRE.Execute(miTexto)(0).SubMatches(0)
                
        miTexto = "[InternetShortcut]" & vbCrLf & _
                  elURL & vbCrLf & _
                  "IconFile=C:\WINDOWS\system32\url.dll" & vbCrLf & _
                  "HotKey=0" & vbCrLf & _
                  "IconIndex=4" & vbCrLf & _
                  "IDList=" & vbCrLf
                  
        Call fso.GetFile(laFile).OpenAsTextStream(ForWriting).Write(miTexto)
        fso.GetFile(laFile).Attributes = 36
    End If
    
    Exit Sub
    
Falla:
    MsgBox "Error al cambiar Atributos e icono: " & Err.Description, vbCritical + vbOKOnly, "Error"
    
End Sub

Public Sub crearURLs(Data As DataObject)
'Recibe una lista de archivos con conteniendo links, que guarda como urls
Dim i As Integer
Dim elLink As String
Dim miSTrm As TextStream
Dim VBRE1 As New RegExp
Dim VBRE2 As New RegExp

    If vbNo = MsgBox("Se ha recibido un archivo de texto que se explorará en busca de links para indexar en:" & vbCrLf & pathBaseDeDatos, vbYesNo + vbDefaultButton2 + vbInformation, "Indexacion Masiva") Then Exit Sub
    
    VBRE2.Pattern = "^https?://"
    VBRE2.Global = True
    VBRE2.IgnoreCase = True
    
    VBRE1.Pattern = "(<a href=)|(>)"
    VBRE1.Global = True
    VBRE1.IgnoreCase = True
    
    If Data.GetFormat(vbCFFiles) Then
        'por Ej. si arrastro desde el ListView del WE o una imagen desde el IE, pues ahí me devuelve el path al archivo de la imagen
        For i = 1 To Data.Files.Count
        dataDragDrop = Data.Files(i)
        If fso.FileExists(dataDragDrop) And fso.GetFile(dataDragDrop).Size > 0 Then
            Set miSTrm = fso.GetFile(dataDragDrop).OpenAsTextStream
            Do
                elLink = miSTrm.ReadLine
                elLink = VBRE1.Replace(elLink, "")
                
                If VBRE2.Test(elLink) Then
                    ArchivarLink elLink  'ojo, archivar el link afecta a VBRE
                End If
            Loop While Not miSTrm.AtEndOfStream
        End If
        Next
    End If
    
End Sub

Public Sub ArchivarLink(miURLOriginal As String)
'Si Existe previamente: Se cambia el icono del URL y la propiedad de System si ya existe
'Si no Existe previamente: Se crea el URL
On Error GoTo Falla
Dim miTexto As String
Dim miTextoRepetido As String
'Dim miURLOriginal As String
Dim miURLRepetido As String
Dim miNombreDeURL As String
Dim elIndice As String
Dim pathBaseDeDatosActual
Dim i As Integer
    
    VBRE.Pattern = "https?://"
    If Not VBRE.Test(miURLOriginal) Then
        GoTo Falla
    End If

    'Supongo que el link es del tipo http://mi.servidor.com/micarpeta/algoMas/elLink?elquery

    'elimino el query
    VBRE.Pattern = "\?.+$"
    miNombreDeURL = VBRE.Replace(miURLOriginal, "")
    'verifico si no termina con / (podria ser del tipo /?) y si termina lo saco
    VBRE.Pattern = "/$"
    miNombreDeURL = VBRE.Replace(miNombreDeURL, "")

    'miURLOriginal = Data.GetData(vbCFText) 'traigo el URL
    VBRE.Pattern = "/([^/]+)$" 'dejo lo que haya desde el último /
    miNombreDeURL = VBRE.Execute(miNombreDeURL)(0).SubMatches(0)
    VBRE.Pattern = "_" 'reemplazo los _ por " " OJO, luego debo trimmear por si queda un " " adelante
    miNombreDeURL = VBRE.Replace(miNombreDeURL, " ")
    'VBRE.Pattern = "\?.*$" 'le saco el query String
    'miNombreDeURL = VBRE.Replace(miNombreDeURL, "")
    
    'Si no hay nada => el url era del tipo http://mi.servidor.com/micarpeta/algoMas/elLink/ o http://mi.servidor.com/micarpeta/algoMas/elLink/?elquery
    'por lo tanto reproceso para quedarme con elLink que es la última carpeta: /elLink/
    
'    If miNombreDeURL = "" Then    'reproceso
'        'miURLOriginal = Data.GetData(vbCFText) 'traigo el URL
'        VBRE.Pattern = "/([^/]+)/[^/]*$" 'dejo lo que haya entre el último / y el penúltimo /
'        miNombreDeURL = VBRE.Execute(miURLOriginal)(0).SubMatches(0)
'        VBRE.Pattern = "_" 'reemplazo los _ por " " OJO, luego debo trimmear por si queda un " " adelante
'        miNombreDeURL = VBRE.Replace(miNombreDeURL, " ")
'    End If
    
    miNombreDeURL = Replace(miNombreDeURL, "#mlrelated", "") 'esto es para sacar el #mlrelated que comenzó a poner xHamster a los links inferiores relacionados
    
    If Len(miNombreDeURL) > 47 Then 'si el tamaño es mayor de 47 lo recorto
        VBRE.Pattern = "^.{47}"
        miNombreDeURL = VBRE.Execute(miNombreDeURL)(0) & "..." 'le agrego los tres ... que le dan tamaño final 50
    End If
    
    miNombreDeURL = Trim(miNombreDeURL) & ".url" 'lo trimmeo y le agrego la terminación .url
    
    VBRE.Pattern = "^."
    elIndice = VBRE.Execute(miNombreDeURL)(0) 'el índice de la sub carpeta
    
VerificarURLRepetidos:
    If fso.FolderExists(pathBaseDeDatos & "\" & elIndice) Then 'cambio el path a la base de datos si ya está indexado con el alfabeto
        pathBaseDeDatosActual = pathBaseDeDatos & "\" & elIndice
    Else
        pathBaseDeDatosActual = pathBaseDeDatos
    End If
    
    If fso.FileExists(pathBaseDeDatosActual & "\" & miNombreDeURL) Then 'verifico si existe y si está repetido
        'como si existe, debo ver si esté repetido el url:
        miTextoRepetido = fso.OpenTextFile(pathBaseDeDatosActual & "\" & miNombreDeURL, ForReading).ReadAll()
        VBRE.Pattern = "URL=(.+)" & vbCrLf 'extraigo el url del repetido
        miURLRepetido = VBRE.Execute(miTextoRepetido)(0).SubMatches(0)
        
        If extraerURLSignificativo(miURLRepetido) = extraerURLSignificativo(miURLOriginal) Then
            'Está repetido, termino aqui
            'Debo ver si es ReadOnly y sino agregarlo a los links instantáneos
            If fso.GetFile(pathBaseDeDatosActual & "\" & miNombreDeURL).Attributes <> 36 Then
                bSiEsUnLink = True 'esto es para agregarlo a la lista instantánea si se dejó caer en el icono correspondiente
                pathAlUltimoLinkAgregado = pathBaseDeDatosActual & "\" & miNombreDeURL
                nombreDelUltimoLinkAgregado = miNombreDeURL
            End If
            Exit Sub
        Else
            'no está repetido, cambio el nombre y vuelvo a verificar
            'miNombreDeURL = Replace(miNombreDeURL, ".url", "")
            VBRE.Pattern = "^(.+?)(_\d*)?\.url"
            miNombreDeURL = VBRE.Execute(miNombreDeURL)(0).SubMatches(0)
            miNombreDeURL = miNombreDeURL & "_" & CStr(i) & ".url"
            i = i + 1
            GoTo VerificarURLRepetidos
        End If
    
    End If
    
    'Si estoy aqui es que no está repetido
    'genero el texto interno:
        
    miTexto = "[InternetShortcut]" & vbCrLf & _
              "URL=" & miURLOriginal & vbCrLf & _
              "IconFile=C:\WINDOWS\system32\url.dll" & vbCrLf & _
              "HotKey=0" & vbCrLf & _
              "IconIndex=0" & vbCrLf & _
              "IDList=" & vbCrLf
                  
    Call fso.CreateTextFile(pathBaseDeDatosActual & "\" & miNombreDeURL).Write(miTexto)
    fso.GetFile(pathBaseDeDatosActual & "\" & miNombreDeURL).Attributes = 32
    
    bSiEsUnLink = True 'esto es para agregarlo a la lista instantánea si se dejó caer en el icono correspondiente
    pathAlUltimoLinkAgregado = pathBaseDeDatosActual & "\" & miNombreDeURL
    nombreDelUltimoLinkAgregado = miNombreDeURL
    
    contadorDeLinks = contadorDeLinks + 1
    Form1.TextContadorDeLinks.Text = Str(contadorDeLinks)
    DoEvents
        
    Exit Sub
        
Falla:

    If vbYes = MsgBox("Hay problemas con el archivo: " & vbCrLf & pathBaseDeDatosActual & "\" & miNombreDeURL & vbCrLf & Err.Description & vbCrLf & "¿Desea para la ejecución?", vbCritical + vbYesNo, "PictureArchivarLink_OLEDragDrop") Then
        End
    End If
End Sub

Public Function extraerURLSignificativo(elURL As String) As String
'revisa el url para que solo lo que diferencia dos contenidos distintos sea tomado en cuenta
Dim VBRE1 As New VBScript_RegExp_55.RegExp

    If CBool(InStr(1, elURL, "ashemaletube.com")) Then
        VBRE1.Pattern = "ashemaletube.com/(.+/)(\?.+)?"
        If VBRE1.Test(elURL) Then
            extraerURLSignificativo = VBRE1.Execute(elURL)(0).SubMatches(0)
            Exit Function
        End If
    End If

    If CBool(InStr(1, elURL, "shittytube.com")) Then
        VBRE1.Pattern = "shittytube.com/(.+)"
        If VBRE1.Test(elURL) Then
            extraerURLSignificativo = VBRE1.Execute(elURL)(0).SubMatches(0)
            Exit Function
        End If
    End If
    
    If CBool(InStr(1, elURL, "hclips.com")) Then
        VBRE1.Pattern = "/videos/([^/]+)/"
        If VBRE1.Test(elURL) Then
            extraerURLSignificativo = VBRE1.Execute(elURL)(0).SubMatches(0)
            Exit Function
        End If
    End If

    If CBool(InStr(1, elURL, "luxuretv.com")) Then
        VBRE1.Pattern = "/videos/(.+)\.html"
        If VBRE1.Test(elURL) Then
            extraerURLSignificativo = VBRE1.Execute(elURL)(0).SubMatches(0)
            Exit Function
        End If
    End If

    If CBool(InStr(1, elURL, "txxx.com")) Then
        VBRE1.Pattern = "/videos/(\d+)/"
        If VBRE1.Test(elURL) Then
            extraerURLSignificativo = VBRE1.Execute(elURL)(0).SubMatches(0)
            Exit Function
        End If
    End If

    If CBool(InStr(1, elURL, "videoszoofilia.org")) Then
        VBRE1.Pattern = "/video/(\d+)/"
        If VBRE1.Test(elURL) Then
            extraerURLSignificativo = VBRE1.Execute(elURL)(0).SubMatches(0)
            Exit Function
        End If
    End If

    If CBool(InStr(1, elURL, "pervertslut.com")) Then
        VBRE1.Pattern = "/videos/(\d+)/"
        If VBRE1.Test(elURL) Then
            extraerURLSignificativo = VBRE1.Execute(elURL)(0).SubMatches(0)
            Exit Function
        End If
    End If

    If CBool(InStr(1, elURL, "xhamster.com")) Then
        VBRE1.Pattern = "/videos/([a-zA-Z0-9-]+\-\d+)(#mlrelated)?"
        If VBRE1.Test(elURL) Then
            extraerURLSignificativo = VBRE1.Execute(elURL)(0).SubMatches(0)
            Exit Function
        End If
    End If


    If CBool(InStr(1, elURL, "xvideos.com")) Then
        
        VBRE1.Pattern = "/(video\d+)/\d?"
        If VBRE1.Test(elURL) Then
            extraerURLSignificativo = VBRE1.Execute(elURL)(0).SubMatches(0)
            Exit Function
        End If
        
        VBRE1.Pattern = "/(\d\d\d+)/\d?"
        If VBRE1.Test(elURL) Then
            extraerURLSignificativo = "video" & VBRE1.Execute(elURL)(0).SubMatches(0)
            Exit Function
        End If
        
        MsgBox "No se verifica la condición significativa de un URL de xvideos.com", vbInformation + vbOKOnly, "extraerURLSignificativo"
        extraerURLSignificativo = ""
        
    End If
    
    
    
    
    'si no tengo condiciones devuelvo lo mismo que recibí
    extraerURLSignificativo = elURL
    
End Function

