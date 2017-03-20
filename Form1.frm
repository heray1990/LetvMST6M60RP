VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   Caption         =   "Letv MST6M60 属性比对工具"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CommandRead 
      Caption         =   "读属性"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   17
      Top             =   3840
      Width           =   1470
   End
   Begin VB.Frame Frame1 
      Caption         =   "Monitor 属性信息"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7280
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "产品类型"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "背光类型"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   1
         Left            =   3650
         TabIndex        =   14
         Top             =   240
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "硬件版本"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   3
         Left            =   3650
         TabIndex        =   13
         Top             =   1060
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "制版阶段"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1060
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "屏型号"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   5
         Left            =   3650
         TabIndex        =   11
         Top             =   1880
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2D/3D 类型"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1880
         Width           =   3500
      End
      Begin VB.Label lbTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "软件版本"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   410
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   2700
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   1
         Left            =   3650
         TabIndex        =   7
         Top             =   660
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1480
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   3
         Left            =   3650
         TabIndex        =   5
         Top             =   1480
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   2300
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   5
         Left            =   3650
         TabIndex        =   3
         Top             =   2300
         Width           =   3500
      End
      Begin VB.Label lbTVInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   350
         Index           =   6
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   3500
      End
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   3840
   End
   Begin VB.TextBox tbLogInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4305
      Left            =   7440
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   2865
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1920
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Menu MenuItemSetting 
      Caption         =   "Setting"
      Begin VB.Menu MenuItemProperities 
         Caption         =   "Properities"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isReadingSwVer As Boolean
Dim arrProductModel
Dim arrBacklightType
Dim arrBoradModel
Dim arrHwVer
Dim arrDimension
Dim arrPanelModel

Private Sub CommandRead_Click()
'On Error GoTo ErrExit
    Dim i, j As Integer
    Dim arrHint
    
    i = 0
    j = 0
    arrHint = Array("Read Product Model:", "Read Backlight Type:", _
                    "Read Board Model:", "Read Hardware Version:", _
                    "Read 2D/3D mode:", "Read Panel Model:", _
                    "Read Software Version")
    InitBeforeRunning
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If

    For i = 1 To itemNumOfTvInfo
        ClearComBuf
        Log_Info CStr(arrHint(i - 1))
        GetProperty Int(i)
        DelayMS 500
        Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
            
        If isCmdDataRecv = False Then
            If j > cmdResendTimes Then
                j = 0
                Log_Info "Cannot read the property. Please do the Letv Reset!!!"
                MsgBox "Please do the Letv Reset!"
                GoTo FAIL
            Else
                j = j + 1
                i = i - 1
            End If
        Else
            j = 0
        End If
    Next i
    
    ClearComBuf
    Log_Info CStr(arrHint(6))
    isReadingSwVer = True
    GetSwVer
    DelayMS 500
    Call DelaySWithCmdFlag(cmdReceiveWaitS, isCmdDataRecv)
            
    If isCmdDataRecv = False Then
        If j > cmdResendTimes Then
            j = 0
            Log_Info "Cannot read the property. Please do the Letv Reset!!!"
            MsgBox "Please do the Letv Reset!"
            GoTo FAIL
        Else
            j = j + 1
            i = i - 1
        End If
    Else
        j = 0
    End If
    
    isReadingSwVer = False
    
    BurningMode 0
    RebootMonitor
    Exit Sub
FAIL:
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    SubInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrExit
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

    End
Exit Sub

ErrExit:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub MenuItemProperities_Click()
    FrmProperities.Show
End Sub

Private Sub MSComm1_OnComm()
    Select Case MSComm1.CommEvent
        Case comEvReceive
            DelayMS 100
            Call DataReceive
        'Case comEvSend
        Case Else
    End Select
End Sub

Private Sub SubInit()
    Dim clsConfigData As ProjectConfig

    Set clsConfigData = New ProjectConfig
    clsConfigData.LoadConfigData
    
    glngTVComBaud = clsConfigData.ComBaud
    gintTVComID = clsConfigData.ComID
    SubInitComPort

    gintProductModel = clsConfigData.ProductModel
    gintBacklightType = clsConfigData.BacklightType
    gintBoardModel = clsConfigData.BoardModel
    gintHardwareVersion = clsConfigData.HardwareVersion
    gint2D3DModel = clsConfigData.Dimension
    gintPanelModel = clsConfigData.PanelModel
    gstrSoftwareVersion = clsConfigData.SoftwareVersion

    Set clsConfigData = Nothing
    
    isCmdDataRecv = False
End Sub

Public Sub SubInitComPort()
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

    MSComm1.CommPort = gintTVComID
    MSComm1.Settings = glngTVComBaud & ",N,8,1"
    MSComm1.InputLen = 0

    MSComm1.InBufferCount = 0
    MSComm1.OutBufferCount = 0
    MSComm1.InputMode = comInputModeBinary

    MSComm1.NullDiscard = False
    MSComm1.DTREnable = False
    MSComm1.EOFEnable = False
    MSComm1.RTSEnable = False
    MSComm1.SThreshold = 1
    MSComm1.RThreshold = 1
    MSComm1.InBufferSize = 1024
    MSComm1.OutBufferSize = 512
End Sub

Private Sub InitBeforeRunning()
    isCmdDataRecv = False
    isReadingSwVer = False
    
    For i = 0 To itemNumOfTvInfo
        lbTVInfo(i).Caption = strNoRecvData
        lbTVInfo(i).BackColor = &HFFFFFF
    Next i

    Log_Clear
    'tbLogInfo.ForeColor = &H80000008
End Sub

Private Sub ClearComBuf()
    MSComm1.InBufferCount = 0
    MSComm1.OutBufferCount = 0
End Sub

Private Sub DataReceive()
    Dim ReceiveArr() As Byte
    Dim receiveData As String
    Dim strMagic As String
    Dim Counter As Integer
    Dim tmpCounter As Integer
    Dim i As Integer
    Dim j As Integer
    Dim tmp As Integer
    Dim firstByteOfDataIdx As Integer
    
    firstByteOfDataIdx = -1
    Counter = MSComm1.InBufferCount

    If (Counter > 0) Then
        receiveData = ""
        ReceiveArr = MSComm1.Input

        For i = 0 To (Counter - 1)
            If i < (Counter - 1) Then
                If Not isReadingSwVer Then
                    If (ReceiveArr(i) = &HA3) And (i + 2) = (Counter - 1) Then
                        tmp = 0
                        For j = 0 To 1
                            tmp = tmp + ReceiveArr(j + i)
                        Next j
                        
                        tmp = &HFF - tmp And &HFF
                        
                        If tmp = ReceiveArr(i + 2) Then
                            firstByteOfDataIdx = i
                        End If
                    End If
                Else
                    If (ReceiveArr(i) = &HA6) And (i + 5) = (Counter - 1) Then
                        tmp = 0
                        For j = 0 To 4
                            tmp = tmp + ReceiveArr(j + i)
                        Next j
                        
                        tmp = &HFF - tmp And &HFF
                        
                        If tmp = ReceiveArr(i + 5) Then
                            firstByteOfDataIdx = i
                        End If
                    End If
                End If
            End If
        Next i

        If firstByteOfDataIdx > 0 Then
            If Not isReadingSwVer Then
                tmpCounter = firstByteOfDataIdx + 2
            Else
                tmpCounter = firstByteOfDataIdx + 5
            End If

            For i = firstByteOfDataIdx To tmpCounter
                If (ReceiveArr(i) < 16) Then
                    receiveData = receiveData + "0" + Hex(ReceiveArr(i)) + Space(1)
                Else
                    receiveData = receiveData + Hex(ReceiveArr(i)) + Space(1)
                End If
            Next i

            Log_Info receiveData
            receiveData = ""
            If Not isReadingSwVer Then
                receiveData = CStr(ReceiveArr(firstByteOfDataIdx + 1))
            Else
                For i = tmpCounter - 1 To firstByteOfDataIdx + 1 Step -1
                    If i = firstByteOfDataIdx + 1 Then
                        strMagic = ""
                    Else
                        strMagic = "."
                    End If
                    
                    If (ReceiveArr(i) < 16) Then
                        receiveData = receiveData + "0" + Hex(ReceiveArr(i)) + strMagic
                    Else
                        receiveData = receiveData + Hex(ReceiveArr(i)) + strMagic
                    End If
                Next i
            End If
            
            infoCompare cmdIdentifyNum, receiveData
        End If
    End If
End Sub

Private Sub infoCompare(cmdIdx As Integer, recvData As String)
    Dim i As Integer

    isCmdDataRecv = True

    If cmdIdx = 0 Then
        arrProductModel = Array("UNKNOWN", "Max4_70", "Max4_65C", _
                            "Max4_55B", "Max4_65B", "Max4_75B", _
                            "Max4_70S", "Max4_75S", "Max455", _
                            "Max4_X70", "Max465", "U55", "U65", "U75")
        lbTVInfo(0).Caption = arrProductModel(Val(recvData))

        If gintProductModel = Val(recvData) Then
            lbTVInfo(0).BackColor = &HFF00&
        Else
            lbTVInfo(0).BackColor = &HFF&
        End If
    End If

    If cmdIdx = 1 Then
        arrBacklightType = Array("PWM", "Local Dimming")
        lbTVInfo(1).Caption = arrBacklightType(Val(recvData) - 1)

        If gintBacklightType = Val(recvData) Then
            lbTVInfo(1).BackColor = &HFF00&
        Else
            lbTVInfo(1).BackColor = &HFF&
        End If
    End If

    If cmdIdx = 2 Then
        arrBoradModel = Array("EVT", "EVT2", "EVT3", _
                        "DVT", "DVT2", "DVT3", _
                        "PVT", "MP")
        lbTVInfo(2).Caption = arrBoradModel(Val(recvData) - 1)
                            
        If gintBoardModel = Val(recvData) Then
            lbTVInfo(2).BackColor = &HFF00&
        Else
            lbTVInfo(2).BackColor = &HFF&
        End If
    End If

    If cmdIdx = 3 Then
        arrHwVer = Array("H1000", "H2000", "H3000", "H5000", "H6000")
        lbTVInfo(3).Caption = arrHwVer(Val(recvData) - 1)
                            
        If gintHardwareVersion = Val(recvData) Then
            lbTVInfo(3).BackColor = &HFF00&
        Else
            lbTVInfo(3).BackColor = &HFF&
        End If
    End If

    If cmdIdx = 4 Then
        arrDimension = Array("2D", "3D")
        lbTVInfo(4).Caption = arrDimension(Val(recvData) - 1)
                            
        If gint2D3DModel = Val(recvData) Then
            lbTVInfo(4).BackColor = &HFF00&
        Else
            lbTVInfo(4).BackColor = &HFF&
        End If
    End If

    If cmdIdx = 5 Then
        arrPanelModel = Array("X4_70_2D", "X4_70_3D", "X3_55_120", _
                            "X3_55_60", "X4_65_Curve", "X4_55_Blade", _
                            "X4_70S", "X4_75S", "X4_55", "X4_65", _
                            "UNQ_55", "UNQ_65", "UNQ_75")
        lbTVInfo(5).Caption = arrPanelModel(Val(recvData) - 1)
                            
        If gintPanelModel = Val(recvData) Then
            lbTVInfo(5).BackColor = &HFF00&
        Else
            lbTVInfo(5).BackColor = &HFF&
        End If
    End If
    
    If cmdIdx = 6 Then
        lbTVInfo(6).Caption = recvData
                            
        If gstrSoftwareVersion = recvData Then
            lbTVInfo(6).BackColor = &HFF00&
        Else
            lbTVInfo(6).BackColor = &HFF&
        End If
    End If
End Sub
