VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Letv Max65 属性比对工具"
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
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   3960
      Width           =   1095
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
    Dim Counter As Integer
    Dim i, j, tmp, firstByteOfDataIdx As Integer
    
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
            isCmdDataRecv = True
            If Not isReadingSwVer Then
                For i = firstByteOfDataIdx To firstByteOfDataIdx + 2
                    If (ReceiveArr(i) < 16) Then
                        receiveData = receiveData + "0" + Hex(ReceiveArr(i)) + Space(1)
                    Else
                        receiveData = receiveData + Hex(ReceiveArr(i)) + Space(1)
                    End If
                Next i
            Else
                For i = firstByteOfDataIdx To firstByteOfDataIdx + 5
                    If (ReceiveArr(i) < 16) Then
                        receiveData = receiveData + "0" + Hex(ReceiveArr(i)) + Space(1)
                    Else
                        receiveData = receiveData + Hex(ReceiveArr(i)) + Space(1)
                    End If
                Next i
            End If
            
            Log_Info receiveData
        Else
            tbLogInfo.Text = tbLogInfo.Text + vbCrLf
            tbLogInfo.SelStart = Len(tbLogInfo.Text)
        End If
    End If
End Sub

