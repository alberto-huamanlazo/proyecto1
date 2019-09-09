VERSION 5.00
Begin VB.Form frmTestDll 
   Caption         =   "Test Dll"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFolder 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "C:\Temp"
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Proceso"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "frmTestDll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'para saber si existe o no una carpeta
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

'Para capturar la ruta hacia la carpeta temporal
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Funciones creadas en la DLL:
Private Declare Function GetChannels Lib "MSUtil.dll" (ByVal strFileOut As String) As Boolean
Private Declare Function GetTargets Lib "MSUtil.dll" (ByVal strFileOut As String) As Boolean
Private Declare Function GetGenders Lib "MSUtil.dll" (ByVal strFileOut As String) As Boolean
Private Declare Function GetPrograms Lib "MSUtil.dll" (ByVal strChannels As String, ByVal strDateIni As String, ByVal strDateFin As String, ByVal strFileOut As String, ByVal bShowProgress As Boolean) As Boolean
Private Declare Function GetRating Lib "MSUtil.dll" (ByVal strChannels As String, ByVal iTypeRating As Integer, ByVal iTarget As Integer, ByVal iBaseTarget As Integer, ByVal strDateIni As String, ByVal strDateFin As String, ByVal strFileOut As String, ByVal bShowProgress As Boolean) As Boolean

Private Sub cmdProcess_Click()
    '"SEÑAL|CODIGO|DESCRIPCION"
    Dim lExito As Long
    Dim sFileTmp As String 'Ruta hacia el archivo creado por la Dll
    
    'Canales
    lExito = 0
    sFileTmp = txtFolder & "\" & "CanalesTvData8.txt"
    lExito = GetChannels(sFileTmp)      'crea el archivo, desde la DLL, en un ruta temporal.
    If lExito <> 0 Then
        'OK
    End If
    
    'Targets
    lExito = 0
    sFileTmp = txtFolder & "\" & "TargetsTvData8.txt"
    lExito = GetTargets(sFileTmp) 'crea el archivo, desde la DLL, en un ruta temporal.
    If lExito <> 0 Then
        'OK
    End If
    
    'Generos
    lExito = 0
    sFileTmp = txtFolder & "\" & "GenerosTvData8.txt"
    lExito = GetGenders(sFileTmp) 'crea el archivo, desde la DLL, en un ruta temporal.
    If lExito <> 0 Then
        'OK
    End If
    
    '
    lExito = 0
    sFileTmp = txtFolder & "\" & "ProgramasTvData8.txt"
    lExito = GetPrograms("2,4,5,9", "06/07/2011", "06/07/2011", sFileTmp, True) 'crea el archivo, desde la DLL, en un ruta temporal.
    If lExito <> 0 Then
        'OK
    End If
    
    lExito = 0
    'sFileTmp = GetDirectoryTemp & "FileX.tmp"
    sFileTmp = txtFolder & "\" & "RatingTvData8.txt"
    lExito = GetRating("2,4,5,9", 1, 68, 68, "06/07/2011", "06/07/2011", sFileTmp, True) 'crea el archivo, desde la DLL, en un ruta temporal. 1=tipo de rating de bloque de programas, 2=tipo de rating tandas de programas, 3= Ambos tipo de rating de bloques y tandas de programas
    
    If lExito <> 0 Then
        'OK
    End If
    
    MsgBox ("Terminó crear los archivos")
End Sub
