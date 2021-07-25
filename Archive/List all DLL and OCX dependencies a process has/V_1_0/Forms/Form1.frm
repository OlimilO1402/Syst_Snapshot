VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10095
   ClientLeft      =   405
   ClientTop       =   2880
   ClientWidth     =   17415
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   17415
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Timer:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Trav HeapList"
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton BtnTerminate 
      Caption         =   "Terminate Proc"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   1680
   End
   Begin VB.ListBox LstHeaps 
      Height          =   7665
      Left            =   14520
      TabIndex        =   6
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ListBox LstHeapLists 
      Height          =   7665
      Left            =   11640
      TabIndex        =   5
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ListBox LstThreads 
      Height          =   7665
      Left            =   8760
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ListBox LstModules 
      Height          =   7665
      Left            =   5880
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ListBox LstChildProcesses 
      Height          =   7665
      Left            =   3000
      TabIndex        =   1
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ListBox LstProcesses 
      Height          =   7860
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label LblHeap 
      Caption         =   "HeapEntry"
      Height          =   2055
      Left            =   14520
      TabIndex        =   11
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label LblHeapList 
      Caption         =   "HeapList"
      Height          =   2055
      Left            =   11640
      TabIndex        =   10
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label LblThread 
      Caption         =   "ThreadEntry"
      Height          =   2055
      Left            =   8760
      TabIndex        =   9
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label LblModule 
      Caption         =   "ModuleEntry"
      Height          =   2055
      Left            =   5880
      TabIndex        =   8
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label LblChildProcess 
      Caption         =   "ChildProcess"
      Height          =   2055
      Left            =   3000
      TabIndex        =   7
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label LblCurProcess 
      Caption         =   "ProcessEntry"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Snapshot As SnapShot
Private m_CurProcess  As ProcessEntry
Private m_CurHeapList As HeapList

Private Sub BtnTerminate_Click()
    m_CurProcess.KillProcess
End Sub

Private Sub Check1_Click()
    Timer1.Enabled = Check1.Value = vbChecked
End Sub

Private Sub Command1_Click()
    Check1.Value = vbUnchecked
    Dim mp As MousePointerConstants: mp = Screen.MousePointer
    Screen.MousePointer = MousePointerConstants.vbHourglass
    'Timer1.Enabled = False
    Me.LstProcesses.Visible = False
    Call MSnapShot.TraverseHeapList(Me.LstProcesses, Command1)
    Me.LstProcesses.Visible = True
    Me.LstProcesses.ZOrder 0
    Screen.MousePointer = mp
End Sub

Private Sub Form_Load()
    'm_Snapshot
    Text1.Text = CStr(Timer1.Interval)
    Check1.Value = IIf(Timer1.Enabled, vbChecked, vbUnchecked)
End Sub

Private Sub Form_Resize()
    'Dim w As Single: w = Me.ScaleWidth - (16 * Screen.TwipsPerPixelX)
    'Dim h As Single: h = Me.ScaleHeight - List1.Top - (8 * Screen.TwipsPerPixelX)
    
    'If w > 0 Then List1.Width = w
    'If h > 0 Then List1.Height = h
End Sub

Private Sub ClearLabels()
    'Me.LblCurProcess.Caption = "Process"
    Me.LblChildProcess.Caption = "ChildProcess"
    Me.LblModule.Caption = "ModuleEntry"
    Me.LblThread.Caption = "ThreadEntry"
    Me.LblHeapList.Caption = "HeapLists"
    Me.LblHeap.Caption = "HeapEntry"
End Sub

Private Sub LstProcesses_Click()
    Dim k As String: k = GetListBoxKey(Me.LstProcesses)
    If Len(k) > 0 Then
        Dim col As Collection: Set col = m_Snapshot.ProcessEntries
        If ContainsKey(col, k) Then
            Set m_CurProcess = col.Item(k)
            Me.LblCurProcess.Caption = m_CurProcess.ToString
            Call MISnapshotEntryList.ISnapshotEntryListToListBox( _
                m_CurProcess.ChildProcesses, Me.LstChildProcesses)
            Call MISnapshotEntryList.ISnapshotEntryListToListBox( _
                m_CurProcess.ModuleEntries, Me.LstModules)
            Call MISnapshotEntryList.ISnapshotEntryListToListBox( _
                m_CurProcess.ThreadEntries, Me.LstThreads)
            Call MISnapshotEntryList.ISnapshotEntryListToListBox( _
                m_CurProcess.HeapLists, Me.LstHeapLists)
            Call ClearLabels
        End If
    End If
End Sub
Private Sub LstChildProcesses_Click()
    Dim k As String: k = GetListBoxKey(Me.LstChildProcesses)
    Me.LstChildProcesses.ZOrder 0
    If Len(k) > 0 Then
        Dim col As Collection: Set col = m_CurProcess.ChildProcesses
        If ContainsKey(col, k) Then
            Dim p As ProcessEntry: Set p = col.Item(k)
            Me.LblChildProcess.Caption = p.ToString
        End If
    End If
End Sub
Private Sub LstModules_Click()
    Dim k As String: k = GetListBoxKey(Me.LstModules)
    Me.LstModules.ZOrder 0
    If Len(k) > 0 Then
        Dim col As Collection: Set col = m_CurProcess.ModuleEntries
        If ContainsKey(col, k) Then
            Dim m As ModuleEntry: Set m = col.Item(k)
            Me.LblModule.Caption = m.ToString
        End If
    End If
End Sub
Private Sub LstThreads_Click()
    Dim k As String: k = GetListBoxKey(Me.LstThreads)
    Me.LstThreads.ZOrder 0
    If Len(k) > 0 Then
        Dim col As Collection: Set col = m_CurProcess.ThreadEntries
        If ContainsKey(col, k) Then
            Dim t As ThreadEntry: Set t = col.Item(k)
            Me.LblThread.Caption = t.ToString
        End If
    End If
End Sub
Private Sub LstHeapLists_Click()
    Dim k As String: k = GetListBoxKey(Me.LstHeapLists)
    Me.LstHeapLists.ZOrder 0
    If Len(k) > 0 Then
        Dim col As Collection: Set col = m_CurProcess.HeapLists
        If ContainsKey(col, k) Then
            Set m_CurHeapList = col.Item(k)
            Me.LblHeapList.Caption = m_CurHeapList.ToString
            Set col = m_CurHeapList.HeapEntries
            Call MISnapshotEntryList.ISnapshotEntryListToListBox(col, Me.LstHeaps)
        End If
    End If
End Sub
Private Sub LstHeaps_Click()
    Dim k As String: k = GetListBoxKey(Me.LstHeaps)
    Me.LstHeaps.ZOrder 0
    If Len(k) > 0 Then
        Dim col As Collection: Set col = m_CurHeapList.HeapEntries
        If ContainsKey(col, k) Then
            Dim h As HeapEntry: Set h = col.Item(k)
            Me.LblHeap.Caption = h.ToString
        End If
    End If
End Sub


Private Function GetListBoxKey(aLB As ListBox) As Long 'String
    If aLB.ListCount > 0 Then
        If aLB.ListIndex >= 0 Then
            'GetListBoxKey = aLB.List(aLB.ListIndex)
            GetListBoxKey = aLB.ItemData(aLB.ListIndex)
        End If
    End If
End Function

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
        If IsNumeric(Text1.Text) Then
            Dim intv As Long: intv = CLng(Text1.Text)
            Timer1.Interval = intv
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Set m_Snapshot = New_Snapshot(SnapAll)
    Call MISnapshotEntryList.ISnapshotEntryListToListBox( _
             m_Snapshot.ProcessEntries, Me.LstProcesses)
    Call ClearLabels
End Sub

