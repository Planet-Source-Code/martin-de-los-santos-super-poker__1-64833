Attribute VB_Name = "Sonidos"
'---------------------------------------------------------------------------------------
'Module/Modulo: Sonidos
'Author/Autor : Mdls
'Purpose      : play sounds with sndPlaySound Lib " WINMM.DLL
'Propósito    : tocar sonidos con sndPlaySound Lib "WINMM.DLL
'---------------------------------------------------------------------------------------

Option Explicit
'Estilo WinXP /Style WinXp
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

' Separa la variable
Public i As Long
' API de soporte de sonido de alto nivel
Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
                              (lpszSoundName As Any, ByVal uFlags As Long) As Long
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
Global Const SND_NODEFAULT = &H2    ' No usar el sonido predeterminado
Global Const SND_MEMORY = &H4    ' lpszSoundName apunta al archivo de memoria

Global SoundBuffer() As Byte
Sub ComienzaTocarSonido(ByVal IdRecurso As String, ByVal grupo As String, Optional ByVal se As Byte)
    SoundBuffer = LoadResData(IdRecurso, grupo)
    If se = 1 Then sndPlaySound SoundBuffer(0), SND_NOSTOP: Exit Sub
    sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End Sub
Sub FinTocarSonido()
    sndPlaySound ByVal vbNullString, 0&
End Sub
Public Function Son(ByVal Arg As String, Optional ByVal gr As String)
    If frm_main.SeSonido = 0 Then Exit Function
    Dim IdRecurso$
    Dim grupo$
    IdRecurso = UCase(Arg)
  '  gr = "WAVSSISTEMAGRAL"
    ComienzaTocarSonido (IdRecurso), gr
End Function
Public Sub Main()
   On Error GoTo Err1
    'solo windows xp /Only in windows xp
    InitCommonControls
    frm_main.Show
    Exit Sub
Err1:
    MsgBox Err.Description
    Resume Next
End Sub




