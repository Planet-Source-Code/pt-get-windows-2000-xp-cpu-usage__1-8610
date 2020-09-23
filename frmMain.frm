VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Windows 2000 CPU Time"
   ClientHeight    =   1065
   ClientLeft      =   2445
   ClientTop       =   1515
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   5340
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   4200
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type CounterInfo
    hCounter As Long
    strName As String
End Type

Dim hQuery As Long
Dim Counters(0 To 99) As CounterInfo
Dim currentCounterIdx As Long
Dim iPerformanceDetail As PERF_DETAIL

Public Sub AddCounter(strCounterName As String, hQuery As Long)
    Dim pdhStatus As PDH_STATUS
    Dim hCounter As Long
    
    pdhStatus = PdhVbAddCounter(hQuery, strCounterName, hCounter)
    Counters(currentCounterIdx).hCounter = hCounter
    Counters(currentCounterIdx).strName = strCounterName
    currentCounterIdx = currentCounterIdx + 1
End Sub

Private Sub UpdateValues()
    Dim dblCounterValue As Double
    Dim pdhStatus As Long
    Dim strInfo As String
    Dim i As Long
        
    PdhCollectQueryData (hQuery)
    
    i = 0  'Only one counter but you can add more

    dblCounterValue = _
            PdhVbGetDoubleCounterValue(Counters(i).hCounter, pdhStatus)
        
        'Some error checking, make sure the query went through
        If (pdhStatus = PDH_CSTATUS_VALID_DATA) _
        Or (pdhStatus = PDH_CSTATUS_NEW_DATA) Then
        strInfo = "CPU Usage: " & Format$(dblCounterValue, "0.00")
        pb1.Value = dblCounterValue
        Me.Caption = Format$(dblCounterValue, "0") & "% - CPU Status"
        End If
        
    Label1 = strInfo
End Sub

Private Sub Form_Load()
    Dim pdhStatus As PDH_STATUS
    
    pdhStatus = PdhOpenQuery(0, 1, hQuery)
    If pdhStatus <> ERROR_SUCCESS Then
        MsgBox "OpenQuery failed"
        End
    End If
    
    ' Add the processor time query
    AddCounter "\Processor(0)\% Processor Time", hQuery
    UpdateValues    ' Force an immediate display of the counter values
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    PdhCloseQuery (hQuery)
End Sub

Private Sub Timer1_Timer()
    ' fires once per second, can be changed.
    UpdateValues
End Sub
