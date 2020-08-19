VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Powered by herryqyh"
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    mystr = ComboBox1.Value
    history = mystr
    'UserForm1.Hide
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    ComboBox1.Value = ""
    mystr = ComboBox1.Value
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ComboBox1.Value = history
    ComboBox1.AddItem "1 6 5 7 10 2 (GB9254 辐射 LO)"
    ComboBox1.AddItem "1 8 2 7 3 4 5 (GB9254 辐射 HI)"
    ComboBox1.AddItem "1 2 7 (GB3434.1 网口骚扰)"
    ComboBox1.AddItem "1 2 8 3 4 (GB3434.1 骚扰功率)"
    ComboBox1.AddItem "1 2 3 4 6 9 (GB3434.1 骚扰电压)"
    
    CommandButton1.SetFocus
End Sub
