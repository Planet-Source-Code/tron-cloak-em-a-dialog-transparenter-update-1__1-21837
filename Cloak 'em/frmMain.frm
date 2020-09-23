VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cloak 'em - Win2k & Higher Only                                [Update 1]"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton C_Refresh 
      Caption         =   "Refresh List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cloak Volume:"
      Height          =   3015
      Left            =   3840
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
      Begin MSComctlLib.Slider S_Volume 
         Height          =   2655
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   4683
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   0
         SmallChange     =   0
         Min             =   1
         Max             =   255
         SelStart        =   125
         TickStyle       =   2
         TickFrequency   =   0
         Value           =   125
      End
   End
   Begin VB.CommandButton C_Exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton C_Change 
      Caption         =   "Set Volume"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton C_UnSet 
      Caption         =   "Decloak"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton C_Set 
      Caption         =   "Cloak"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Icon Keys:"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   3615
      Begin VB.Label Label2 
         Caption         =   "that window/dialog is cloaked."
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "that window/dialog is not cloaked at all"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "frmMain.frx":0E42
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   120
         Picture         =   "frmMain.frx":1C84
         Top             =   480
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3918
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer_Load 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7320
      Top             =   120
   End
   Begin MSComctlLib.ListView LV_Windows 
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Current Windows that are open"
         Object.Width           =   5787
      EndProperty
   End
   Begin VB.ListBox List_Windows 
      Height          =   645
      Left            =   6720
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   7200
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List_IDs 
      Height          =   450
      Left            =   6720
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cloak'em v1.1 by TRON/Nathan Martin, site: http://t2n.dyndns.org"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' I only added and remove allot of junk from the code
' and created this project.  Most of the code you see is
' from some projects that were posted at PSC.  I do not
' intend to take credit for other authors work, but for
' my work and how I've made it useful for people to use
' unlike how most code authors who post at PSC that leave it
' all crapy and useless.  I did in fact did the whole GUI my self
' and the code with allot of comments like in Form_Load Sub that
' is code I've created for this project.
'
' Happy coding and Production! :)   -TRON(tron@ircd-net.org) Date of Release [03/21/2001], Date Updated [03/22/2001]
Dim TimeNumber As Integer
Dim Loading As Boolean

Public Function LoadWindows()
On Error Resume Next 'just in case we have problems
Dim X As Integer
Start = 1
Paren = 1
Retval = EnumWindows(AddressOf EnumChildProc, 0)
OnlyCaps = 0

For X = 0 To frmMain.List2.ListCount - 1
    Paren = 0
    phwnd = frmMain.List2.List(X)
    Retval = EnumChildWindows(phwnd, AddressOf EnumChildProc, 0)
Next X

End Function

Private Sub C_Change_Click()
    C_Set_Click ' Since this part will use the same code anyway used to
                ' set the cloak, mine as well just do a shortcut to that
                ' code that's already there, no need to have extra code laying around.
End Sub

Private Sub C_Exit_Click()
    Unload Me
End Sub

Private Sub C_Refresh_Click()
Dim Index As Integer
    Index = LV_Windows.SelectedItem.Index ' Get Selected Window
    SetupEverything 'let's redo all lists
    LV_Windows.ListItems(Index).Selected = True ' Select it back
End Sub

Private Sub C_Set_Click()
Dim Index As Integer
Dim ByteSet As Byte
    ByteSet = S_Volume.Value
    Index = LV_Windows.SelectedItem.Index
    If List_IDs.List(Index - 1) = Me.hwnd Then
        If ByteSet < 11 Then MsgBox "Yeah right, trying to cloak the cloaker eh?" & vbCrLf & "Well, it's not gunna happen ;p", vbInformation, "Nope, Not gunna happen ;p": Exit Sub
        ' ^- let's keep the user from cloaking the wrong window shall we... :)
    End If
    If Index = 0 Then MsgBox "Select a Window/Dialog first!", vbInformation, "Nothing to Cloak"
    SetLayered List_IDs.List(Index - 1), True, ByteSet ' Cloak window/dialog
    C_Refresh_Click 'let's refresh list
    LV_Windows.ListItems(Index).Selected = True ' Select item back
End Sub

Private Sub C_UnSet_Click()
Dim Index As Integer
    Index = LV_Windows.SelectedItem.Index
    If Index = 0 Then MsgBox "Select a Window/Dialog first!", vbInformation, "Nothing to decloak"
    SetLayered List_IDs.List(Index - 1), False, 0 ' Decloak window/dialog
    C_Refresh_Click 'let's refresh list
    LV_Windows.ListItems(Index).Selected = True ' Select item back
End Sub

Private Sub Form_Load()
On Error Resume Next
' Let's start off right :P
    Loading = True ' We're loading up, so let's tell the timer that
    Me.Hide ' let's just hide while loading for the cool effect :)
    SetupEverything ' let's get everything setup for user's use
    
    ' Now let's do our form cool effect ;)
    SetLayered Me.hwnd, True, 0 ' Cloak Form
    Me.Show ' need to bring us back, but cloaked ;)
    Timer_Load.Enabled = True ' Start timer to slowly uncloak form :P
    
    'Change buttons to C++ style buttons :D
    CButton C_Set
    CButton C_UnSet
    CButton C_Change
    CButton C_Refresh
    CButton C_Exit
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim TimeByte As Byte
Dim LoopEnds As Boolean
    'OK, you have two options, so let the timer do the closing effect or have the loop function do it which is kinda faster
    'and it looks better on most systems, but for those on you who have faster then 650MHz systems you may want to let the timer do it.
    LoopEnds = True ' Change this to False to allow the timer to do the unloading which is kinda slower on slower systems, but gives a better graphic effect on faster machines
    
    If LoopEnds = True Then
        Do Until TimeNumber = 0
            TimeNumber = TimeNumber - 25
            TimeByte = TimeNumber
            SetLayered Me.hwnd, True, TimeByte
        Loop
        
        Set frmMain = Nothing
        End
    Else
        Cancel = 1 ' Let's cancel Ending App here, so timer can do the close effect for us ;)
        Loading = False ' We're closing down App, so let's tell Timer that we are...
        Timer_Load.Enabled = True ' Let's have the timer to the rest for us =)
    End If
End Sub

Private Sub Timer_Load_Timer()
On Error Resume Next
Dim TimeByte As Byte
    If Loading = True Then
        TimeNumber = TimeNumber + 25
        TimeByte = TimeNumber
        SetLayered Me.hwnd, True, TimeByte
        If TimeNumber = 250 Then
            SetLayered Me.hwnd, False, 0 ' OK, We're done here, make us visible completly
            Timer_Load.Enabled = False ' We don't want this to continue, so let's disable it.
            Exit Sub ' And exit the sub, Note: this doesn't have to be here since it's at the
                     ' end of the sub, but some times VB can be wierd and loop the sub one more time anyany.
        End If
    Else
        TimeNumber = TimeNumber - 25
        TimeByte = TimeNumber
        SetLayered Me.hwnd, True, TimeByte
        If TimeNumber = 0 Then
            Set frmMain = Nothing
            End
        End If
    End If
End Sub

Sub SetupEverything()
On Error Resume Next
Dim X As Integer
    
    'Let's clear out all values since this sub is used to refresh list as well...
    List2.Clear
    List_IDs.Clear
    List_Windows.Clear
    LV_Windows.ListItems.Clear
    
    LoadWindows
ReScan:
    ' OK, now since we have the list of windows and their ID
    ' we'll just remove the ones that don't have caption at all.
    For X = 0 To List_Windows.ListCount - 1
        If List_Windows.List(X) = "" Or List_Windows.List(X) = "SysFader" Or List_Windows.List(X) = "MCI command handling window" Then ' Yes, we want SysFader and MCI command handling windows removed too.
            List_Windows.RemoveItem X ' Remove it from the list :)
            List_IDs.RemoveItem X     ' Got to remove ID for it too since it's useless now...
            GoTo ReScan
            ' Since we found one, we need to start over on the scan
            ' to be sure to remove all of them cause we don't want
            ' to miss any and still end-up show windows with empty
            ' captions which isn't useful since you don't know what
            ' that window is.
        End If
    Next X
    ' Now that's how you get rid of all those empty caption windows :)
    
    For X = 0 To List_Windows.ListCount - 1 ' Let's add all the windows to the visible list
        If CheckLayered(List_IDs.List(X)) = True Then
            LV_Windows.ListItems.Add , , List_Windows.List(X), , 2 ' Looks like the Window is Cloak, lets say it is
        Else
            LV_Windows.ListItems.Add , , List_Windows.List(X), , 1 ' Looks like the window isn't, so lets say it's normal.
        End If
    Next X
    ' Now we're done here :D
End Sub
