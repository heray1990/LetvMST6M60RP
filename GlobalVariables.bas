Attribute VB_Name = "GlobalVariables"
Option Explicit

Public gintTVComID As Integer
Public glngTVComBaud As Long

Public gintProductModel As Integer
Public gintBacklightType As Integer
Public gintBoardModel As Integer
Public gintHardwareVersion As Integer
Public gint2D3DModel As Integer
Public gintPanelModel As Integer
Public gstrSoftwareVersion As String

Public isCmdDataRecv As Boolean
Public countTime As Long
Public cmdIdentifyNum As Integer
Public Const strNoRecvData As String = "None"
Public Const itemNumOfTvInfo As Integer = 6
Public Const cmdReceiveWaitS As Integer = 5
Public Const cmdResendTimes As Integer = 2
