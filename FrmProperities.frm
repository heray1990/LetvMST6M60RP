VERSION 5.00
Begin VB.Form FrmProperities 
   Caption         =   "Properities"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3990
   LinkTopic       =   "Form2"
   ScaleHeight     =   5505
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "TV 串口"
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   3735
      Begin VB.ComboBox cmbComID 
         Height          =   315
         ItemData        =   "FrmProperities.frx":0000
         Left            =   1560
         List            =   "FrmProperities.frx":0002
         TabIndex        =   18
         Text            =   "COM1"
         Top             =   320
         Width           =   2000
      End
      Begin VB.ComboBox cmbComBaud 
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Text            =   "9600"
         Top             =   780
         Width           =   2000
      End
      Begin VB.Label Label1 
         Caption         =   "串口："
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   380
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "波特率："
         Height          =   200
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   700
      End
   End
   Begin VB.CommandButton CommandSet 
      Caption         =   "设置"
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame FrameProperty 
      Caption         =   "Monitor 属性预设"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox TextSwVer 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   3080
         Width           =   2000
      End
      Begin VB.ComboBox ComboProduct 
         Height          =   315
         ItemData        =   "FrmProperities.frx":0004
         Left            =   1560
         List            =   "FrmProperities.frx":0006
         TabIndex        =   6
         Text            =   "Product Model"
         Top             =   320
         Width           =   2000
      End
      Begin VB.ComboBox ComboBacklight 
         Height          =   315
         ItemData        =   "FrmProperities.frx":0008
         Left            =   1560
         List            =   "FrmProperities.frx":000A
         TabIndex        =   5
         Text            =   "Backlight Type"
         Top             =   780
         Width           =   2000
      End
      Begin VB.ComboBox ComboBoard 
         Height          =   315
         ItemData        =   "FrmProperities.frx":000C
         Left            =   1560
         List            =   "FrmProperities.frx":000E
         TabIndex        =   4
         Text            =   "Board Model"
         Top             =   1240
         Width           =   2000
      End
      Begin VB.ComboBox ComboHwVer 
         Height          =   315
         ItemData        =   "FrmProperities.frx":0010
         Left            =   1560
         List            =   "FrmProperities.frx":0012
         TabIndex        =   3
         Text            =   "Hardware Version"
         Top             =   1680
         Width           =   2000
      End
      Begin VB.ComboBox Combo2D3D 
         Height          =   315
         ItemData        =   "FrmProperities.frx":0014
         Left            =   1560
         List            =   "FrmProperities.frx":0016
         TabIndex        =   2
         Text            =   "2D/3D"
         Top             =   2160
         Width           =   2000
      End
      Begin VB.ComboBox ComboPanel 
         Height          =   315
         ItemData        =   "FrmProperities.frx":0018
         Left            =   1560
         List            =   "FrmProperities.frx":001A
         TabIndex        =   1
         Text            =   "Panel Model"
         Top             =   2620
         Width           =   2000
      End
      Begin VB.Label Label7 
         Caption         =   "软件版本："
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3140
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "产品类型："
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   380
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "背光类型："
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "制版阶段："
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1300
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "硬件版本："
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1760
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "2D/3D 类型："
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "屏型号："
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2680
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmProperities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim arrProductModel
Dim arrBacklightType
Dim arrBoradModel
Dim arrHwVer
Dim arrDimension
Dim arrPanelModel

Private Sub CommandSet_Click()
On Error GoTo ErrExit
    Dim i As Integer
    Dim clsSaveConfigData As ProjectConfig
    
    Set clsSaveConfigData = New ProjectConfig

    gintTVComID = Val(Replace(cmbComID.Text, "COM", ""))
    glngTVComBaud = Val(cmbComBaud.Text)
    clsSaveConfigData.ComBaud = CStr(glngTVComBaud)
    clsSaveConfigData.ComID = gintTVComID

    For i = 0 To 10
        If Trim(ComboProduct.Text) = Trim(arrProductModel(i)) Then
            clsSaveConfigData.ProductModel = i
            gintProductModel = i
            Exit For
        End If
    Next i
    For i = 0 To 1
        If Trim(ComboBacklight.Text) = Trim(arrBacklightType(i)) Then
            clsSaveConfigData.BacklightType = i + 1
            gintBacklightType = i + 1
            Exit For
        End If
    Next i
    For i = 0 To 7
        If Trim(ComboBoard.Text) = Trim(arrBoradModel(i)) Then
            clsSaveConfigData.BoardModel = i + 1
            gintBoardModel = i + 1
            Exit For
        End If
    Next i
    For i = 0 To 4
        If Trim(ComboHwVer.Text) = Trim(arrHwVer(i)) Then
            clsSaveConfigData.HardwareVersion = i + 1
            gintHardwareVersion = i + 1
            Exit For
        End If
    Next i
    For i = 0 To 1
        If Trim(Combo2D3D.Text) = Trim(arrDimension(i)) Then
            clsSaveConfigData.Dimension = i + 1
            gint2D3DModel = i + 1
            Exit For
        End If
    Next i
    For i = 0 To 9
        If Trim(ComboPanel.Text) = Trim(arrPanelModel(i)) Then
            clsSaveConfigData.PanelModel = i + 1
            gintPanelModel = i + 1
            Exit For
        End If
    Next i
    clsSaveConfigData.SoftwareVersion = Trim(TextSwVer.Text)
    gstrSoftwareVersion = Trim(TextSwVer.Text)

    clsSaveConfigData.SaveConfigData
    Set clsSaveConfigData = Nothing

    Unload Me
    Form1.SubInitComPort
    Form1.Show
    Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub Form_Load()
    Dim i As Integer

    cmbComID.Text = "COM" & CStr(gintTVComID)
    cmbComBaud.Text = CStr(glngTVComBaud)

    For i = 1 To 20
        cmbComID.AddItem "COM" & i
    Next i

    cmbComBaud.AddItem "9600"
    cmbComBaud.AddItem "19200"
    cmbComBaud.AddItem "38400"
    cmbComBaud.AddItem "57600"
    cmbComBaud.AddItem "115200"

    arrProductModel = Array("UNKNOWN", "Max4_70", "Max4_65C", _
                            "Max4_55B", "Max4_65B", "Max4_75B", _
                            "Max4_70S", "Max4_75S", "Max5_55_938", _
                            "Max4_X70", "Max5_65_938")
    arrBacklightType = Array("PWM", "Local Dimming")
    arrBoradModel = Array("EVT", "EVT2", "EVT3", _
                        "DVT", "DVT2", "DVT3", _
                        "PVT", "MP")
    arrHwVer = Array("H1000", "H2000", "H3000", "H5000", "H6000")
    arrDimension = Array("2D", "3D")
    arrPanelModel = Array("X4_70_2D", "X4_70_3D", "X3_55_120", _
                            "X3_55_60", "X4_65_Curve", "X4_55_Blade", _
                            "X4_70S", "X4_75S", "X4_55_938", "X4_65_938")

    For i = 0 To 10
        ComboProduct.AddItem arrProductModel(i)
    Next i
    For i = 0 To 1
        ComboBacklight.AddItem arrBacklightType(i)
    Next i
    For i = 0 To 7
        ComboBoard.AddItem arrBoradModel(i)
    Next i
    For i = 0 To 4
        ComboHwVer.AddItem arrHwVer(i)
    Next i
    For i = 0 To 1
        Combo2D3D.AddItem arrDimension(i)
    Next i
    For i = 0 To 9
        ComboPanel.AddItem arrPanelModel(i)
    Next i
    
    If gintBacklightType < 1 Then
        gintBacklightType = 1
    End If
    If gintBoardModel < 1 Then
        gintBoardModel = 1
    End If
    If gintHardwareVersion < 1 Then
        gintHardwareVersion = 1
    End If
    If gint2D3DModel < 1 Then
        gint2D3DModel = 1
    End If
    If gintPanelModel < 1 Then
        gintPanelModel = 1
    End If
    ComboProduct.Text = arrProductModel(gintProductModel)
    ComboBacklight.Text = arrBacklightType(gintBacklightType - 1)
    ComboBoard.Text = arrBoradModel(gintBoardModel - 1)
    ComboHwVer.Text = arrHwVer(gintHardwareVersion - 1)
    Combo2D3D.Text = arrDimension(gint2D3DModel - 1)
    ComboPanel.Text = arrPanelModel(gintPanelModel - 1)
    TextSwVer.Text = gstrSoftwareVersion
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
