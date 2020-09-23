VERSION 5.00
Begin VB.Form frmMac 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2460
      Top             =   90
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   90
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mac"
      Height          =   405
      Left            =   4140
      TabIndex        =   3
      Top             =   210
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1230
      TabIndex        =   1
      Top             =   330
      Width           =   45
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EDA84D&
      Height          =   2040
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   9000
      Left            =   240
      Picture         =   "frmMac.frx":0000
      Top             =   0
      Width           =   60
   End
   Begin VB.Image imgok 
      Height          =   750
      Left            =   420
      Picture         =   "frmMac.frx":042E
      Top             =   60
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgerr 
      Height          =   750
      Left            =   420
      Picture         =   "frmMac.frx":0A49
      Top             =   60
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgnormal 
      Height          =   750
      Left            =   420
      Picture         =   "frmMac.frx":1064
      Top             =   60
      Width           =   750
   End
End
Attribute VB_Name = "frmMac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Netbios Lib "netapi32.dll" _
   (pncb As NCB) As Byte
        
Private Declare Sub CopyMemory Lib "kernel32" Alias _
   "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource _
                    As Long, ByVal cbCopy As Long)
        
Private Declare Function GetProcessHeap Lib "kernel32" () _
                                                        As Long
        
Private Declare Function HeapAlloc Lib "kernel32" (ByVal _
                                                   hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes _
                                                   As Long) As Long
        
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap _
                                                  As Long, ByVal dwFlags As Long, lpMem As Any) As Long
        
Const NCBASTAT = &H33
Const NCBNAMSZ = 16
Const HEAP_ZERO_MEMORY = &H8
Const HEAP_GENERATE_EXCEPTIONS = &H4
Const NCBRESET = &H32

Private Type NCB
    ncb_command As Byte
    ncb_retcode As Byte
    ncb_lsn As Byte
    ncb_num As Byte
    ncb_buffer As Long
    ncb_length As Integer
    ncb_callname As String * NCBNAMSZ
    ncb_name As String * NCBNAMSZ
    ncb_rto As Byte
    ncb_sto As Byte
    ncb_post As Long
    ncb_lana_num As Byte
    ncb_cmd_cplt As Byte
    ncb_reserve(9) As Byte
    ncb_event As Long
End Type

Private Type ADAPTER_STATUS
    adapter_address(5) As Byte
    rev_major As Byte
    reserved0 As Byte
    adapter_type As Byte
    rev_minor As Byte
    duration As Integer
    frmr_recv As Integer
    frmr_xmit As Integer
    iframe_recv_err As Integer
    xmit_aborts As Integer
    xmit_success As Long
    recv_success As Long
    iframe_xmit_err As Integer
    recv_buff_unavail As Integer
    t1_timeouts As Integer
    ti_timeouts As Integer
    Reserved1 As Long
    free_ncbs As Integer
    max_cfg_ncbs As Integer
    max_ncbs As Integer
    xmit_buf_unavail As Integer
    max_dgram_size As Integer
    pending_sess As Integer
    max_cfg_sess As Integer
    max_sess As Integer
    max_sess_pkt_size As Integer
    name_count As Integer
End Type

Private Type NAME_BUFFER
    name As String * NCBNAMSZ
    name_num As Integer
    name_flags As Integer
End Type

Private Type ASTAT
    adapt As ADAPTER_STATUS
    NameBuff(30) As NAME_BUFFER
End Type
Dim WithEvents sink As SWbemSink
Attribute sink.VB_VarHelpID = -1
Dim TimerC As Integer
Dim NoAdaptercard$

Private Sub getmac()
    On Error Resume Next
    Dim myNcb As NCB
    Dim bRet As Byte
    Dim myASTAT As ASTAT, tempASTAT As ASTAT
    Dim pASTAT As Long
    
    myNcb.ncb_command = NCBRESET
    bRet = Netbios(myNcb)
    myNcb.ncb_command = NCBASTAT
    myNcb.ncb_lana_num = 0
    myNcb.ncb_callname = "* "
    
    myNcb.ncb_length = Len(myASTAT)
    
    pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS _
       Or HEAP_ZERO_MEMORY, myNcb.ncb_length)
             
    If pASTAT = 0 Then Exit Sub
    
    myNcb.ncb_buffer = pASTAT
    bRet = Netbios(myNcb)
    
    CopyMemory myASTAT, myNcb.ncb_buffer, Len(myASTAT)
    MACAddress = _
       HexEx(myASTAT.adapt.adapter_address(0)) & "-" & _
       HexEx(myASTAT.adapt.adapter_address(1)) & "-" & _
       HexEx(myASTAT.adapt.adapter_address(2)) & "-" & _
       HexEx(myASTAT.adapt.adapter_address(3)) & "-" & _
       HexEx(myASTAT.adapt.adapter_address(4)) & "-" & _
       HexEx(myASTAT.adapt.adapter_address(5))
    Call HeapFree(GetProcessHeap(), 0, pASTAT)

    If Replace(MACAddress, "-", "") = "000000000000" Then

        Timer2.Enabled = True
        Command1_Click

    Else
        
        imgnormal.Visible = False
        imgok.Visible = True
        frmMain.txtSerial.Text = MACAddress
        Timeout 1
        Unload Me

    End If
    
End Sub

Private Function HexEx(ByVal B&) As String
    On Error Resume Next
    Dim aa$
    
    aa = Hex(B)

    If Len(aa) < 2 Then aa = "0" & aa

    HexEx = aa

End Function

Private Sub Command1_Click()

    On Error Resume Next
    Text7.Text = 0
    Set sink = New SWbemSink
    Set adapter = GetObject("winmgmts:\\127.0.0.1")
    adapter.InstancesOfAsync sink, "Win32_NetworkAdapter"

End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    Set sink = New SWbemSink
    
    Center Me
    Timer1.Enabled = True
    
    NoAdaptercard$ = "You must install a Network Interface Card to continue with the installation."
    Label2.Caption = "Searching Network Adapters..."

End Sub

Private Sub Form_Unload(cancel As Integer)
    On Error Resume Next
    Timer2.Enabled = False
    Timer1.Enabled = False
    Set sink = Nothing
    Set adapter = Nothing

End Sub

Private Sub sink_OnObjectReady(ByVal objWbemObject As WbemScripting.ISWbemObject, ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)

    On Error Resume Next
    Dim i As Integer
    i = Text7.Text
    Set adapter = GetObject("winmgmts:Win32_NetworkAdapterConfiguration=" & i & "")
    Description = adapter.Description

    If IsNull(adapter.MACAddress) = False Then
        
        If InStr(1, adapter.MACAddress, ":") > 0 And Len(adapter.MACAddress) = 17 Then
            
            If InStr(1, Description, "WAN") = 0 Then

                MACAddress = adapter.MACAddress
                imgnormal.Visible = False
                imgok.Visible = True
                'imgerr.Visible = True
                frmMain.txtSerial.Text = Replace(MACAddress, ":", "-")
                Label2.Caption = Description
                Timeout 1
                Unload Me
                
            End If

        End If

    End If

    Text7.Text = i + 1
    
End Sub

Private Sub Timer1_Timer()

    On Error Resume Next
    getmac
    Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()

    On Error Resume Next
    
    TimerC = TimerC + 1

    If TimerC >= 15 Then

        If Len(MACAddress) = 0 Then
        
            MsgBox NoAdaptercard$, vbOKOnly, "ERROR"
            Timer2.Enabled = False
            imgnormal.Visible = False
            imgerr.Visible = True

        End If

        Set sink = Nothing
        Set adapter = Nothing
        Unload Me

    End If

End Sub

