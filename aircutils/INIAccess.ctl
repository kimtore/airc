VERSION 5.00
Begin VB.UserControl INIAccess 
   Appearance      =   0  'Flat
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2235
   ClipControls    =   0   'False
   Enabled         =   0   'False
   FillStyle       =   3  'Vertical Line
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   PropertyPages   =   "INIAccess.ctx":0000
   ScaleHeight     =   810
   ScaleWidth      =   2235
   ToolboxBitmap   =   "INIAccess.ctx":0011
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INIAccess"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1365
   End
End
Attribute VB_Name = "INIAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "INIAccess"
Dim File As String, Entry As String

Private Function TrimCrLf(ByVal S As String) As String
S = Replace(S, vbCr, "")
S = Replace(S, vbLf, "")
TrimCrLf = Replace(S, Chr(13), "")
End Function

Private Sub UserControl_InitProperties()
File = "WIN.INI"
Entry = "New_Entry"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
File = PropBag.ReadProperty("File", "WIN.INI")
Entry = PropBag.ReadProperty("Entry", "New_Entry")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "File", File, "WIN.INI"
PropBag.WriteProperty "Entry", Entry, "New_Entry"
End Sub

Private Sub UserControl_Resize(): Size Label1.Width + 90, Label1.Height + 90: End Sub

Public Sub INISaveSetting(ByVal Key As String, ByVal Inf As String)
Attribute INISaveSetting.VB_Description = "Saves a setting to the INI file"
WritePrivateProfileString ByVal Entry, ByVal Key, ByVal Inf, ByVal File
End Sub

Public Sub INIDeleteSetting(ByVal Key As String)
Attribute INIDeleteSetting.VB_Description = "Deletes a setting from a ini file"
WritePrivateProfileString Entry, Key, ByVal 0&, File
End Sub

Public Sub INIDeleteEntry()
Attribute INIDeleteEntry.VB_Description = "Deletes a entry from a INI file."
WritePrivateProfileString Entry, ByVal 0&, ByVal 0&, File
End Sub

Public Function INIGetSetting(ByVal Key As String) As String
Attribute INIGetSetting.VB_Description = "Gets a setting from the INI file"
Dim Buf As String
Buf = String(255, 0)
StrLen = GetPrivateProfileString(ByVal Entry, ByVal Key, ByVal (Chr(10) + Chr(13)), Buf, 255, ByVal File)
INIGetSetting = TrimCrLf(Left(Buf, StrLen))
End Function

Public Function INIGetKeyList() As Variant
Attribute INIGetKeyList.VB_Description = "Get a string array with all the items"
Dim Buf As String
Buf = String(25500, 0)
StrLen = GetPrivateProfileString(ByVal Entry, ByVal 0&, "", Buf, 25500, ByVal File)
INIGetKeyList = SplitList(Left(Buf, StrLen))
End Function

Public Sub INIDeCace()
WritePrivateProfileString ByVal 0&, ByVal 0&, ByVal 0&, File
End Sub
 
 ' INIEntry = Entry
Public Property Get INIEntry() As String
Attribute INIEntry.VB_Description = "INI Entry to use for INISettings"
Attribute INIEntry.VB_ProcData.VB_Invoke_Property = "INI_Config;INI Files"
INIEntry = Entry
End Property

Public Property Let INIEntry(ByVal NV As String)
Entry = NV
PropertyChanged "INIEntry"
End Property
 
 ' INIFile = File
Public Property Get INIFile() As String
Attribute INIFile.VB_Description = "INI File to use for INISettings"
Attribute INIFile.VB_ProcData.VB_Invoke_Property = "INI_Config;INI Files"
INIFile = File
End Property

Public Property Let INIFile(ByVal NV As String)
File = NV
PropertyChanged "INIFile"
End Property
