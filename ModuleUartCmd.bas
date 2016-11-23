Attribute VB_Name = "ModuleUartCmd"
'**********************************************
' Module for burning mode of Letv MST6M60.
'**********************************************
Option Explicit

Private mSendDataBuf(0 To 10) As Byte
Private mDDCDataWithoutChksum(0 To 5) As Byte
Private i As Integer

Private Sub SendCmd()
    Form1.MSComm1.Output = mSendDataBuf
    DelayMS 500
End Sub

Private Function CalDDCChkSum(ByRef data() As Byte) As Byte
    Dim tmp As Integer

    tmp = 0
    CalDDCChkSum = &H0

    For i = 0 To 5
        tmp = tmp + data(i)
    Next i
    
    CalDDCChkSum = tmp And &HF
End Function

Private Function CalChkSum(ByRef data() As Byte) As Byte
    Dim tmp As Integer

    tmp = 0
    CalChkSum = &H0

    For i = 0 To 9
        tmp = tmp + data(i)
    Next i
    
    CalChkSum = &HFF - tmp And &HFF
End Function

Private Sub DataToDDC()
    For i = 0 To 5
        mDDCDataWithoutChksum(i) = mSendDataBuf(i + 4)
    Next i
End Sub

Private Sub SaveCmdToLog(ByRef data() As Byte)
    Dim strSendData As String

    strSendData = ""
    For i = 0 To 10
        If (mSendDataBuf(i) < 16) Then
            strSendData = strSendData + "0" + Hex(data(i)) + " "
        Else
            strSendData = strSendData + Hex(data(i)) + " "
        End If
    Next i
    Log_Info strSendData
End Sub

Public Sub GetProperty(intProperty As Integer)
    cmdIdentifyNum = intProperty - 1
    'E0 0B 40 XD 03 XX 00 00 00 00 CHK
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H3
    mSendDataBuf(5) = intProperty
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &HD
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    
    SaveCmdToLog mSendDataBuf
    
    isCmdDataRecv = False
    SendCmd
End Sub

Public Sub GetSwVer()
    cmdIdentifyNum = 6
    'E0 0B 40 1D 01 00 00 00 00 00 B6
    mSendDataBuf(0) = &HE0
    mSendDataBuf(1) = &HB
    mSendDataBuf(2) = &H40
    mSendDataBuf(4) = &H1
    mSendDataBuf(5) = &H0
    mSendDataBuf(6) = &H0
    mSendDataBuf(7) = &H0
    mSendDataBuf(8) = &H0
    mSendDataBuf(9) = &H0

    DataToDDC
    mSendDataBuf(3) = CalDDCChkSum(mDDCDataWithoutChksum) * 16 + &HD
    mSendDataBuf(10) = CalChkSum(mSendDataBuf)
    
    SaveCmdToLog mSendDataBuf
    
    isCmdDataRecv = False
    SendCmd
End Sub
