VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "徕卡串口工具"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5385
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   5385
   StartUpPosition =   2  '屏幕中心
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame3 
      Caption         =   "独立指令："
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   5175
      Begin VB.CommandButton Command4 
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   600
         TabIndex        =   9
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "\r\n"
         Height          =   375
         Left            =   4680
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "%R1Q,"
         Height          =   200
         Left            =   200
         TabIndex        =   10
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3840
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "通讯操作："
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Frame Frame2 
      Caption         =   "操作日志："
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.Frame Frame4 
      Caption         =   "链接"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   5175
      Begin VB.Label Label1 
         Caption         =   "串口："
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FoundPort As Integer

Private LaserOn As Boolean


Private Sub Command1_Click()

    Dim ComPorts() As Integer
    Dim ComCount As Long
    Dim i As Long
    Dim PortNumber As Integer

    Dim Cmd As String
    Dim ResponseData As String
    Dim Response5004 As String
    Dim Response5003 As String

    Dim StartTime As Single
    Dim OpenError As String

    Dim InstrumentName As String
    Dim InstrumentNumber As String




    FoundPort = 0

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""


    AddLog "========================================"
    AddLog "开始扫描串口"
    AddLog "========================================"



    ComCount = GetExistingComPorts(ComPorts)


    If ComCount = 0 Then

        AddLog "电脑当前没有检测到串口"

        Exit Sub

    End If


    AddLog "电脑当前检测到 " & _
           CStr(ComCount) & _
           " 个串口"


    '----------------------------------------------------
    ' 显示 COM 列表
    '----------------------------------------------------

    For i = 0 To ComCount - 1

        AddLog "检测到：COM" & _
               CStr(ComPorts(i))

    Next i


    AddLog "----------------------------------------"
    AddLog "开始扫描实际存在的串口"
    AddLog "----------------------------------------"



    Cmd = "%R1Q,5004:" & vbCrLf


    AddLog "查询仪器型号：5004"

    AddLog "发送ASCII：" & _
           Replace(Cmd, vbCrLf, "\r\n")

    AddLog "发送HEX：" & _
           StringToHex(Cmd)




    For i = 0 To ComCount - 1

        PortNumber = ComPorts(i)


        On Error Resume Next

        If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
        End If

        Err.Clear

        On Error GoTo 0


        On Error Resume Next

        Err.Clear

        MSComm1.CommPort = PortNumber
        MSComm1.Settings = "9600,N,8,1"
        MSComm1.InputMode = comInputModeText

        MSComm1.InBufferCount = 0
        MSComm1.OutBufferCount = 0

        MSComm1.PortOpen = True


        If Err.Number <> 0 Then

            OpenError = Err.Description

            Err.Clear

            On Error GoTo 0


            AddLog "COM" & _
                   CStr(PortNumber) & _
                   " 无法打开"

            If Len(OpenError) > 0 Then
                AddLog "原因：" & OpenError
            End If

            GoTo NextPort

        End If

        On Error GoTo 0




        AddLog "----------------------------------------"

        AddLog "正在查询 COM" & _
               CStr(PortNumber)

        AddLog "COM" & _
               CStr(PortNumber) & _
               " 已打开"



        On Error Resume Next

        If MSComm1.InBufferCount > 0 Then
            ResponseData = MSComm1.Input
        End If

        MSComm1.InBufferCount = 0
        MSComm1.OutBufferCount = 0

        Err.Clear

        On Error GoTo 0



        AddLog "发送ASCII：" & _
               Replace(Cmd, vbCrLf, "\r\n")

        AddLog "发送HEX：" & _
               StringToHex(Cmd)


        On Error Resume Next

        Err.Clear

        MSComm1.Output = Cmd


        If Err.Number <> 0 Then

            OpenError = Err.Description

            Err.Clear

            On Error GoTo 0


            AddLog "COM" & _
                   CStr(PortNumber) & _
                   " 发送失败"

            If Len(OpenError) > 0 Then
                AddLog "原因：" & OpenError
            End If

            GoTo ClosePort

        End If

        On Error GoTo 0


        AddLog "Write：发送成功"


        Response5004 = ""

        StartTime = Timer


        Do

            DoEvents


            On Error Resume Next

            If MSComm1.InBufferCount > 0 Then

                Response5004 = _
                    Response5004 & _
                    MSComm1.Input

            End If

            Err.Clear

            On Error GoTo 0



            If InStr(1, _
                     Response5004, _
                     "%R1P,0,0:0", _
                     vbTextCompare) > 0 Then

                Exit Do

            End If



            If Timer >= StartTime Then

                If Timer - StartTime >= 3 Then
                    Exit Do
                End If

            Else

                If (86400 - StartTime) + Timer >= 3 Then
                    Exit Do
                End If

            End If


        Loop


 

        If InStr(1, _
                 Response5004, _
                 "%R1P,0,0:0", _
                 vbTextCompare) > 0 Then



            FoundPort = PortNumber

            Text1.Text = _
                "COM" & _
                CStr(PortNumber)


            AddLog "========================================"

            AddLog "找到 Leica GeoCom"

            AddLog "串口：COM" & _
                   CStr(PortNumber)



            AddLog "5004 收到ASCII：" & _
                   Replace( _
                       Replace( _
                           Response5004, _
                           vbCr, "\r"), _
                           vbLf, "\n")


            AddLog "5004 收到HEX：" & _
                   StringToHex(Response5004)



            InstrumentName = _
                GetInstrumentName(Response5004)


            If InstrumentName = "" Then

                InstrumentName = "未知"

                AddLog "5004 返回中没有找到仪器型号"

            Else

                AddLog "设备：" & _
                       InstrumentName

            End If


            If InstrumentName <> "未知" Then

                Form1.Caption = _
                    "徕卡串口工具：" & _
                    InstrumentName

            End If

            Cmd = "%R1Q,5003:" & vbCrLf


            AddLog "----------------------------------------"

            AddLog "查询仪器编号：5003"

            AddLog "发送ASCII：" & _
                   Replace(Cmd, vbCrLf, "\r\n")

            AddLog "发送HEX：" & _
                   StringToHex(Cmd)



            On Error Resume Next

            MSComm1.InBufferCount = 0
            MSComm1.OutBufferCount = 0

            Err.Clear

            On Error GoTo 0


            On Error Resume Next

            Err.Clear

            MSComm1.Output = Cmd


            If Err.Number <> 0 Then

                OpenError = Err.Description

                Err.Clear

                On Error GoTo 0


                AddLog "5003 发送失败：" & _
                       OpenError

                GoTo ClosePort

            End If

            On Error GoTo 0


            AddLog "Write：发送成功"



            Response5003 = ""

            StartTime = Timer


            Do

                DoEvents


                On Error Resume Next

                If MSComm1.InBufferCount > 0 Then

                    Response5003 = _
                        Response5003 & _
                        MSComm1.Input

                End If

                Err.Clear

                On Error GoTo 0




                If InStr(1, _
                         Response5003, _
                         "%R1P,0,0:0", _
                         vbTextCompare) > 0 Then

                    Exit Do

                End If



                If Timer >= StartTime Then

                    If Timer - StartTime >= 3 Then
                        Exit Do
                    End If

                Else

                    If (86400 - StartTime) + Timer >= 3 Then
                        Exit Do
                    End If

                End If


            Loop


            If InStr(1, _
                     Response5003, _
                     "%R1P,0,0:0", _
                     vbTextCompare) > 0 Then



                AddLog "5003 收到ASCII：" & _
                       Replace( _
                           Replace( _
                               Response5003, _
                               vbCr, "\r"), _
                               vbLf, "\n")


                AddLog "5003 收到HEX：" & _
                       StringToHex(Response5003)


                InstrumentNumber = _
                    GetInstrumentNumber(Response5003)


                If InstrumentNumber = "" Then

                    InstrumentNumber = "未知"

                    AddLog "5003 返回中没有找到仪器编号"

                Else

                    AddLog "编号：" & _
                           InstrumentNumber

                End If


                Text2.Text = _
                    "设备：""" & _
                    InstrumentName & _
                    """，编号：""" & _
                    InstrumentNumber & _
                    """"


                AddLog "========================================"

                AddLog "设备：""" & _
                       InstrumentName & _
                       """，编号：""" & _
                       InstrumentNumber & _
                       """"


                AddLog "========================================"


            Else


                InstrumentNumber = "未知"


                AddLog "5003：3秒内没有收到有效返回"


                Text2.Text = _
                    "设备：""" & _
                    InstrumentName & _
                    """，编号：""" & _
                    InstrumentNumber & _
                    """" & _
                    vbCrLf & _
                    vbCrLf & _
                    "【5004 仪器型号返回】" & _
                    vbCrLf & _
                    Response5004 & _
                    vbCrLf & _
                    "【5003 无有效返回】"

            End If


            Exit Sub


        Else



            If Len(Response5004) > 0 Then

                AddLog "COM" & _
                       CStr(PortNumber) & _
                       " 有返回，但不是有效 Leica GeoCom"

                AddLog "收到ASCII：" & _
                       Replace( _
                           Replace( _
                               Response5004, _
                               vbCr, "\r"), _
                               vbLf, "\n")


                AddLog "收到HEX：" & _
                       StringToHex(Response5004)

            Else

                AddLog "COM" & _
                       CStr(PortNumber) & _
                       " 无返回"

            End If

        End If


ClosePort:



        On Error Resume Next

        If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
        End If

        Err.Clear

        On Error GoTo 0


NextPort:

    Next i




    AddLog "========================================"

    AddLog "所有实际存在的串口扫描完成"

    AddLog "没有找到 Leica GeoCom"

    AddLog "========================================"

End Sub



Private Function GetExistingComPorts( _
                    ByRef Ports() As Integer) As Long

    Dim WMI As Object
    Dim Devices As Object
    Dim Device As Object

    Dim DeviceName As String
    Dim ComNumber As Integer

    Dim Count As Long
    Dim TempPorts() As Integer

    Dim i As Long


    On Error GoTo ErrorHandler


    Set WMI = GetObject( _
        "winmgmts:\\.\root\cimv2")


    Set Devices = WMI.ExecQuery( _
        "SELECT Name FROM Win32_PnPEntity " & _
        "WHERE Name LIKE '%(COM%)'")


    Count = 0


    For Each Device In Devices

        DeviceName = Device.Name


        ComNumber = _
            ExtractComNumber(DeviceName)


        If ComNumber > 0 Then

            If Not PortAlreadyExists( _
                    TempPorts, _
                    Count, _
                    ComNumber) Then


                Count = Count + 1


                ReDim Preserve _
                    TempPorts(0 To Count - 1)


                TempPorts(Count - 1) = _
                    ComNumber

            End If

        End If

    Next Device


    If Count = 0 Then

        GetExistingComPorts = 0

        Exit Function

    End If


    SortComPorts TempPorts, Count


    ReDim Ports(0 To Count - 1)


    For i = 0 To Count - 1

        Ports(i) = TempPorts(i)

    Next i


    GetExistingComPorts = Count

    Exit Function


ErrorHandler:

    GetExistingComPorts = 0

End Function



Private Function ExtractComNumber( _
                    ByVal DeviceName As String) As Integer

    Dim P1 As Long
    Dim P2 As Long

    Dim S As String


    ExtractComNumber = 0


    P1 = InStrRev( _
            UCase$(DeviceName), _
            "(COM")


    If P1 = 0 Then Exit Function


    P2 = InStr( _
            P1, _
            DeviceName, _
            ")")


    If P2 = 0 Then Exit Function


    S = Mid$( _
            DeviceName, _
            P1 + 4, _
            P2 - P1 - 4)


    If IsNumeric(S) Then

        If CInt(S) > 0 Then

            ExtractComNumber = _
                CInt(S)

        End If

    End If

End Function




Private Function GetInstrumentName( _
                    ByVal Response As String) As String

    Dim P1 As Long
    Dim P2 As Long


    GetInstrumentName = ""


    P1 = InStr( _
            1, _
            Response, _
            """", _
            vbBinaryCompare)


    If P1 = 0 Then Exit Function


    P2 = InStr( _
            P1 + 1, _
            Response, _
            """", _
            vbBinaryCompare)


    If P2 = 0 Then Exit Function


    GetInstrumentName = _
        Mid$( _
            Response, _
            P1 + 1, _
            P2 - P1 - 1)

End Function




Private Function GetInstrumentNumber( _
                    ByVal Response As String) As String

    Dim P1 As Long
    Dim P2 As Long
    Dim S As String


    GetInstrumentNumber = ""



    P1 = InStr( _
            1, _
            Response, _
            "%R1P,0,0:0,", _
            vbTextCompare)


    If P1 = 0 Then Exit Function



    P1 = _
        P1 + _
        Len("%R1P,0,0:0,")



    P2 = _
        InStr( _
            P1, _
            Response, _
            vbCr)


    If P2 = 0 Then

        P2 = _
            InStr( _
                P1, _
                Response, _
                vbLf)

    End If


    If P2 = 0 Then

        P2 = _
            Len(Response) + 1

    End If


    S = _
        Mid$( _
            Response, _
            P1, _
            P2 - P1)


    GetInstrumentNumber = _
        Trim$(S)

End Function


Private Function PortAlreadyExists( _
                    ByRef Ports() As Integer, _
                    ByVal Count As Long, _
                    ByVal PortNumber As Integer) As Boolean

    Dim i As Long


    PortAlreadyExists = False


    If Count <= 0 Then Exit Function


    For i = 0 To Count - 1

        If Ports(i) = PortNumber Then

            PortAlreadyExists = True

            Exit Function

        End If

    Next i

End Function



Private Sub SortComPorts( _
                    ByRef Ports() As Integer, _
                    ByVal Count As Long)

    Dim i As Long
    Dim j As Long

    Dim Temp As Integer


    For i = 0 To Count - 2

        For j = i + 1 To Count - 1

            If Ports(j) < Ports(i) Then

                Temp = Ports(i)

                Ports(i) = Ports(j)

                Ports(j) = Temp

            End If

        Next j

    Next i

End Sub


Private Sub Command2_Click()

    Dim Cmd As String
    Dim ResponseData As String
    Dim StartTime As Single


    If FoundPort = 0 Then

        MsgBox "请先扫描并连接 Leica 仪器！", _
               vbExclamation

        Exit Sub

    End If


    On Error Resume Next


    If MSComm1.PortOpen = False Then

        MSComm1.CommPort = FoundPort

        MSComm1.Settings = "9600,N,8,1"

        MSComm1.InputMode = comInputModeText

        MSComm1.InBufferCount = 0

        MSComm1.OutBufferCount = 0


        Err.Clear

        MSComm1.PortOpen = True


        If Err.Number <> 0 Then

            MsgBox "无法打开 COM" & _
                   CStr(FoundPort), _
                   vbExclamation

            Err.Clear

            On Error GoTo 0

            Exit Sub

        End If

    End If


    On Error GoTo 0

    Cmd = "%R1Q,17008:0" & vbCrLf

    Text2.Text = ""


    AddLog "========================================"

    AddLog "发送：%R1Q,17008:0"

    AddLog "发送ASCII：" & _
           Replace(Cmd, vbCrLf, "\r\n")

    AddLog "发送HEX：" & _
           StringToHex(Cmd)


    On Error Resume Next

    MSComm1.InBufferCount = 0

    Err.Clear

    MSComm1.Output = Cmd


    If Err.Number <> 0 Then

        AddLog "发送失败：" & _
               Err.Description

        Err.Clear

        On Error GoTo 0

        Exit Sub

    End If


    On Error GoTo 0


    AddLog "发送成功"


    ResponseData = ""

    StartTime = Timer


    Do

        DoEvents


        On Error Resume Next

        If MSComm1.InBufferCount > 0 Then

            ResponseData = _
                ResponseData & _
                MSComm1.Input

        End If

        Err.Clear

        On Error GoTo 0


        If InStr(1, _
                 ResponseData, _
                 "%R1P", _
                 vbTextCompare) > 0 Then

            Exit Do

        End If


        If Timer >= StartTime Then

            If Timer - StartTime >= 3 Then
                Exit Do
            End If

        Else

            If (86400 - StartTime) + Timer >= 3 Then
                Exit Do
            End If

        End If


    Loop


    If Len(ResponseData) > 0 Then

        Text2.Text = ResponseData

        AddLog "收到ASCII：" & _
               Replace( _
                   Replace( _
                       ResponseData, _
                       vbCr, "\r"), _
                       vbLf, "\n")

        AddLog "收到HEX：" & _
               StringToHex(ResponseData)

    Else

        Text2.Text = "无返回"

        AddLog "3秒内没有收到返回"

    End If



    Cmd = "%R1Q,2012:-0.034" & vbCrLf


    AddLog "----------------------------------------"

    AddLog "发送：%R1Q,2012:-0.034"

    AddLog "发送ASCII：" & _
           Replace(Cmd, vbCrLf, "\r\n")

    AddLog "发送HEX：" & _
           StringToHex(Cmd)


    On Error Resume Next

    MSComm1.InBufferCount = 0

    Err.Clear

    MSComm1.Output = Cmd


    If Err.Number <> 0 Then

        AddLog "发送失败：" & _
               Err.Description

        Err.Clear

        On Error GoTo 0

        Exit Sub

    End If


    On Error GoTo 0


    AddLog "发送成功"


    ResponseData = ""

    StartTime = Timer


    Do

        DoEvents


        On Error Resume Next

        If MSComm1.InBufferCount > 0 Then

            ResponseData = _
                ResponseData & _
                MSComm1.Input

        End If

        Err.Clear

        On Error GoTo 0


        If InStr(1, _
                 ResponseData, _
                 "%R1P", _
                 vbTextCompare) > 0 Then

            Exit Do

        End If


        If Timer >= StartTime Then

            If Timer - StartTime >= 3 Then
                Exit Do
            End If

        Else

            If (86400 - StartTime) + Timer >= 3 Then
                Exit Do
            End If

        End If


    Loop


    If Len(ResponseData) > 0 Then

        Text2.Text = _
            Text2.Text & _
            vbCrLf & _
            ResponseData


        AddLog "收到ASCII：" & _
               Replace( _
                   Replace( _
                       ResponseData, _
                       vbCr, "\r"), _
                       vbLf, "\n")


        AddLog "收到HEX：" & _
               StringToHex(ResponseData)

    Else

        AddLog "3秒内没有收到返回"

    End If

End Sub



Private Sub Command3_Click()

    Dim Cmd As String
    Dim ResponseData As String
    Dim StartTime As Single


    If FoundPort = 0 Then

        MsgBox "请先扫描并连接 Leica 仪器！", _
               vbExclamation

        Exit Sub

    End If


    On Error Resume Next


    If MSComm1.PortOpen = False Then

        MSComm1.CommPort = FoundPort

        MSComm1.Settings = "9600,N,8,1"

        MSComm1.InputMode = comInputModeText

        MSComm1.InBufferCount = 0

        MSComm1.OutBufferCount = 0

        Err.Clear

        MSComm1.PortOpen = True


        If Err.Number <> 0 Then

            MsgBox "无法打开 COM" & _
                   CStr(FoundPort), _
                   vbExclamation

            Err.Clear

            On Error GoTo 0

            Exit Sub

        End If

    End If


    On Error GoTo 0


    If LaserOn = False Then

        Cmd = "%R1Q,1004:1" & vbCrLf

        AddLog "========================================"

        AddLog "发送：打开激光"

    Else

        Cmd = "%R1Q,1004:0" & vbCrLf

        AddLog "========================================"

        AddLog "发送：关闭激光"

    End If


    AddLog "发送ASCII：" & _
           Replace(Cmd, vbCrLf, "\r\n")

    AddLog "发送HEX：" & _
           StringToHex(Cmd)


    ResponseData = ""


    On Error Resume Next

    MSComm1.InBufferCount = 0

    Err.Clear

    MSComm1.Output = Cmd


    If Err.Number <> 0 Then

        AddLog "发送失败：" & _
               Err.Description

        Err.Clear

        On Error GoTo 0

        Exit Sub

    End If


    On Error GoTo 0


    AddLog "发送成功"


    StartTime = Timer


    Do

        DoEvents


        On Error Resume Next

        If MSComm1.InBufferCount > 0 Then

            ResponseData = _
                ResponseData & _
                MSComm1.Input

        End If

        Err.Clear

        On Error GoTo 0


        If InStr(1, _
                 ResponseData, _
                 "%R1P", _
                 vbTextCompare) > 0 Then

            Exit Do

        End If


        If Timer >= StartTime Then

            If Timer - StartTime >= 3 Then
                Exit Do
            End If

        Else

            If (86400 - StartTime) + Timer >= 3 Then
                Exit Do
            End If

        End If


    Loop


    If Len(ResponseData) > 0 Then

        Text2.Text = ResponseData


        AddLog "收到ASCII：" & _
               Replace( _
                   Replace( _
                       ResponseData, _
                       vbCr, "\r"), _
                       vbLf, "\n")


        AddLog "收到HEX：" & _
               StringToHex(ResponseData)


        If InStr(1, _
                 ResponseData, _
                 "%R1P,0,0:0", _
                 vbTextCompare) > 0 Then


            If LaserOn = False Then

                LaserOn = True

                Command3.Caption = "关闭激光"

                AddLog "激光打开成功"

            Else

                LaserOn = False

                Command3.Caption = "打开激光"

                AddLog "激光关闭成功"

            End If

        Else

            AddLog "激光指令执行失败"

        End If


    Else

        Text2.Text = "无返回"

        AddLog "3秒内没有收到返回"

    End If

End Sub



Private Sub Command4_Click()

    Dim Cmd As String
    Dim ResponseData As String
    Dim StartTime As Single
    Dim InputCmd As String


    If FoundPort = 0 Then

        MsgBox "请先扫描并连接 Leica 仪器！", _
               vbExclamation

        Exit Sub

    End If


    InputCmd = Trim$(Text4.Text)


    If Len(InputCmd) = 0 Then

        MsgBox "请输入指令！", _
               vbExclamation

        Exit Sub

    End If


    If UCase$(Left$(InputCmd, 5)) = "%R1Q," Then

        Cmd = InputCmd & vbCrLf

    Else

        Cmd = "%R1Q," & _
              InputCmd & _
              vbCrLf

    End If


    On Error Resume Next


    If MSComm1.PortOpen = False Then

        MSComm1.CommPort = FoundPort

        MSComm1.Settings = "9600,N,8,1"

        MSComm1.InputMode = comInputModeText

        MSComm1.InBufferCount = 0

        MSComm1.OutBufferCount = 0

        Err.Clear

        MSComm1.PortOpen = True


        If Err.Number <> 0 Then

            MsgBox "无法打开 COM" & _
                   CStr(FoundPort), _
                   vbExclamation

            Err.Clear

            On Error GoTo 0

            Exit Sub

        End If

    End If


    On Error GoTo 0


    AddLog "========================================"

    AddLog "Text4 自定义指令"

    AddLog "发送ASCII：" & _
           Replace(Cmd, vbCrLf, "\r\n")

    AddLog "发送HEX：" & _
           StringToHex(Cmd)


    ResponseData = ""


    On Error Resume Next

    MSComm1.InBufferCount = 0

    Err.Clear

    MSComm1.Output = Cmd


    If Err.Number <> 0 Then

        AddLog "发送失败：" & _
               Err.Description

        Err.Clear

        On Error GoTo 0

        Exit Sub

    End If


    On Error GoTo 0


    AddLog "发送成功"


    StartTime = Timer


    Do

        DoEvents


        On Error Resume Next

        If MSComm1.InBufferCount > 0 Then

            ResponseData = _
                ResponseData & _
                MSComm1.Input

        End If

        Err.Clear

        On Error GoTo 0


        If InStr(1, _
                 ResponseData, _
                 "%R1P", _
                 vbTextCompare) > 0 Then

            Exit Do

        End If


        If Timer >= StartTime Then

            If Timer - StartTime >= 3 Then
                Exit Do
            End If

        Else

            If (86400 - StartTime) + Timer >= 3 Then
                Exit Do
            End If

        End If


    Loop


    If Len(ResponseData) > 0 Then

        Text2.Text = ResponseData


        AddLog "收到ASCII：" & _
               Replace( _
                   Replace( _
                       ResponseData, _
                       vbCr, "\r"), _
                       vbLf, "\n")


        AddLog "收到HEX：" & _
               StringToHex(ResponseData)

    Else

        Text2.Text = "无返回"

        AddLog "3秒内没有收到返回"

    End If

End Sub




Private Sub Form_Load()

    MSComm1.Settings = "9600,N,8,1"

    MSComm1.InputMode = comInputModeText

    MSComm1.InBufferSize = 4096

    MSComm1.OutBufferSize = 4096

    MSComm1.RThreshold = 0

    MSComm1.SThreshold = 0


    FoundPort = 0

    LaserOn = False


    Command1.Caption = "扫描 Leica"

    Command2.Caption = "设置棱镜参数"

    Command3.Caption = "打开激光"

    Command4.Caption = "发送指令"


    AddLog "程序启动"

End Sub




Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next


    If MSComm1.PortOpen = True Then

        MSComm1.PortOpen = False

    End If


    On Error GoTo 0

End Sub


Private Sub AddLog(ByVal Msg As String)

    Text3.Text = _
        Text3.Text & _
        Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
        "  " & _
        Msg & _
        vbCrLf


    Text3.SelStart = Len(Text3.Text)

End Sub





Private Function StringToHex(ByVal S As String) As String

    Dim i As Long

    Dim Result As String

    Dim Value As Long


    Result = ""


    For i = 1 To Len(S)

        Value = _
            Asc( _
                Mid$(S, i, 1))


        Result = _
            Result & _
            Right$( _
                "0" & _
                Hex$(Value), _
                2) & _
            " "

    Next i


    StringToHex = Trim$(Result)

End Function




Private Sub MSComm1_OnComm()

    Dim ReceiveData As String


    If MSComm1.CommEvent = comEvReceive Then

        On Error Resume Next


        ReceiveData = MSComm1.Input


        If Len(ReceiveData) > 0 Then

            Text2.Text = _
                Text2.Text & _
                ReceiveData

        End If


        Err.Clear


        On Error GoTo 0

    End If

End Sub

