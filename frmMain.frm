VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "SuperClass Test..."
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   ScaleHeight     =   505
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Label lblHeading 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":0000
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      UseMnemonic     =   0   'False
      Width           =   29190
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Name.......... frmMain
'File.......... frmMain.frm
'Dependencies.. Requires cSuperClass & iSuperClass
'Description... Simple-ish demonstration of how to use the cSuperClass window subclasser.
'Author........ Paul_Caton@hotmail.com
'Date.......... June, 13th 2002
'Copyright..... None.

'
'---------
'Features
'---------
'
'No module!
'   AFAIK this is the only source-subclasser that doesn't require a module.
'
'Fast!
'   The subclassing WndProc is implemented in run-time dynamically generated machine-code!
'
'Flexible!
'   Filtered mode:    Will only callback on the messages that you specify, each of which
'                     can be individually set to call before or after default processing.
'   All message mode: Calls back for all messages after default processing.
'
'No Events!
'   Events are slow, the SuperClasser uses implemented interfaces.
'
'Safe in the IDE!
'   Seems to be immune to the End button (and End statement), see below.
'
'----------------------
'The VB IDE End button
'----------------------
'It's generally considered to be a BAD idea to hit the VB IDE end button with a source
'subclasser running as the IDE will usually blow-up (same goes for the End statement).
'
'However, on my machine at least, this one is totally resistant to the issue. This
'I think is because the WndProc remains executable even after the end button has stopped
'the application whereas a conventional subclasser isn't.
'
'Caveat Emptor: Your mileage may vary.
'
'As an ActiveX DLL
'-----------------
'You'll need to set the instancing properties as follows...
'cSuperClass - MultiUse
'iSuperClass - PublicNotCreatable
'
'You should also convert the Debug.Assert statements in cSuperClass to Err.Raise as
'assertions are useless for checking run-time parameters within a compiled component.
'
'Used privately within a UserControl
'-----------------------------------
'You'll need to set the instancing properties as follows...
'cSuperClass - Private
'iSuperClass - PublicNotCreatable
'
'Multiple Subclassers
'--------------------
'No issues, but realise that all cSuperClass instances will callback through the same interfaces.
'If common messages are used among them then use the hWnd parameter to distinguish the source.
'

Option Explicit

'Windows messages that we're going to filter for callback.
Private Const WM_MOVE             As Long = &H3
Private Const WM_SIZE             As Long = &H5
Private Const WM_NCMOUSEMOVE      As Long = &HA0
Private Const WM_NCLBUTTONDOWN    As Long = &HA1
Private Const WM_NCLBUTTONUP      As Long = &HA2
Private Const WM_NCLBUTTONDBLCLK  As Long = &HA3
Private Const WM_NCRBUTTONDOWN    As Long = &HA4
Private Const WM_NCMBUTTONDOWN    As Long = &HA7
Private Const WM_NCMBUTTONUP      As Long = &HA8
Private Const WM_NCMBUTTONDBLCLK  As Long = &HA9
Private Const WM_NCHITTEST        As Long = &H84
Private Const WM_SYSCOMMAND       As Long = &H112
Private Const WM_MOUSEMOVE        As Long = &H200
Private Const WM_LBUTTONDOWN      As Long = &H201
Private Const WM_LBUTTONUP        As Long = &H202
Private Const WM_LBUTTONDBLCLK    As Long = &H203
Private Const WM_RBUTTONDOWN      As Long = &H204
Private Const WM_RBUTTONUP        As Long = &H205
Private Const WM_RBUTTONDBLCLK    As Long = &H206
Private Const WM_MBUTTONDOWN      As Long = &H207
Private Const WM_MBUTTONUP        As Long = &H208
Private Const WM_MBUTTONDBLCLK    As Long = &H209
Private Const WM_MOUSEWHEEL       As Long = &H20A
Private Const WM_PAINT            As Long = &HF

'Message parameters
Private Const HTCLIENT            As Long = &H1
Private Const HTCAPTION           As Long = &H2
Private Const SC_MAXIMIZE         As Long = &HF030&

'API declarations that we'll use to scroll the displayed output
Private Const SW_INVALIDATE       As Long = &H2

Private Type RECT
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
End Type

Private Declare Function ScrollWindowEx Lib "user32" (ByVal hWnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As Any, ByVal fuScroll As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

'App variables
Private nTextHeight As Long         'Height of a text line
Private nLastLine   As Long         'Lowest vertical position where a line of text is completely visible
Private rc          As RECT         'Scrolling rectangle
Private sc          As cSuperClass  'Declare the subclasser

'We're implementing the interfaces declared in iSuperClass. Once the following declaration is in
'place you'll find an entry in the left hand combo-box at the top of the code window for iSuperClass.
Implements iSuperClass

Private Sub Form_Load()
  nTextHeight = TextHeight("My")  'Get the height of a line of text.
  rc.Top = nTextHeight + 1        'Set the top coordinate of the scrolling rectangle.
  CurrentY = ScaleHeight          'Set the vertical coordinate for the next print.
  
  'Performs an IntegralHeight ala a ListBox control
  Height = Height - ((ScaleHeight Mod nTextHeight) * Screen.TwipsPerPixelY)

  Show                            'So the form becomes visible before the first message shows up
  DoEvents
  
  Set sc = New cSuperClass        'Create a cSuperClass instance
  
  With sc
    'Tell the subclasser which messages to callback on (filtered mode).
    'Note: There's an optional second parameter to AddMsg which should be set to True if you
    '      wish to receive the message *before* default processing rather than the more usual
    '      *after*. For example, suppose you'd planned to do some custom painting upon receipt
    '      of a WM_PAINT message, if you'd added the message as *before* default processing then
    '      chances are that your painting would be over-painted by the windows painting that
    '      came after. One good use for *before* processing is to filter out messages, see the
    '      WM_SYSCOMMAND and WM_NCLBUTTONDBLCLK messages for an example.
    '
    'One truely insignificant optimization is to add the messages in frequency order, most
    'frequent first, least frequent last, for the subclasser tests the message number in the
    'order that they're added. Soonest matched, least tested. However, understand that we're
    'talking here about a few machine-code cycles per message number test.
    Call .AddMsg(WM_NCHITTEST)              'Test mouse pos to determine what part of the window it's over... frame, caption, client etc.
    Call .AddMsg(WM_MOUSEMOVE)              'Mouse movements in the client area
    Call .AddMsg(WM_NCMOUSEMOVE)            'Mouse movements in the non-client area (frame, caption etc)
    Call .AddMsg(WM_MOVE)                   'Window movements
    Call .AddMsg(WM_SIZE)                   'Window size changes
    Call .AddMsg(WM_MOUSEWHEEL)             'Mouse wheel
    Call .AddMsg(WM_LBUTTONDOWN)            'Left button down in client
    Call .AddMsg(WM_LBUTTONUP)              'Left button up in client
    Call .AddMsg(WM_LBUTTONDBLCLK)          'Left button double click in client
    Call .AddMsg(WM_MBUTTONDOWN)            'Middle button down in client
    Call .AddMsg(WM_MBUTTONUP)              'Middle button up in client
    Call .AddMsg(WM_MBUTTONDBLCLK)          'Middle button double click in client
    Call .AddMsg(WM_RBUTTONDOWN)            'Right button down in client
    Call .AddMsg(WM_RBUTTONUP)              'Right button up in client
    Call .AddMsg(WM_RBUTTONDBLCLK)          'Right button double click in client
    Call .AddMsg(WM_NCLBUTTONDOWN)          'Left button down in non-client
    Call .AddMsg(WM_NCLBUTTONUP)            'Left button up in non-client
    Call .AddMsg(WM_NCMBUTTONDOWN)          'Middle button down in non-client
    Call .AddMsg(WM_NCRBUTTONDOWN)          'Right button down in non-client

    'These messages will callback *before* default processing. We'll use them to filter out maximize commands.
    Call .AddMsg(WM_SYSCOMMAND, True)       'System menu / frame button command
    Call .AddMsg(WM_NCLBUTTONDBLCLK, True)  'Left  button double click in non-client

    'Start subclassing.
    'Note: hWnd could be the handle of another window, the desktop window even, a control, whatever...
    '      As long as there's an hWnd we can subclass it.
    'Note: There's an optional third parameter to SubClass which if set to true will callback on all
    '      messages. To try this out comment out the previous .AddMsg's and change the line below to...
    '    .Subclass(hWnd, Me, True) - Understand though that filtered mode will always be faster.
    Call .Subclass(hWnd, Me)
  End With
End Sub

Private Sub Form_Resize()
  If WindowState <> vbMinimized Then
  
    nLastLine = ScaleHeight - nTextHeight
    
    With rc
      rc.Right = ScaleWidth
      rc.Bottom = ScaleHeight
    End With
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set sc = Nothing  'Destroy the subclasser
End Sub

'This implemented interface is called AFTER default processing... that is, AFTER the previous WndProc
'See iSuperClass for parameter information
Private Sub iSuperClass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

  Select Case uMsg
  Case WM_PAINT
    'If we're in AllMsg mode then all messages come thru here, if we go on to display
    'this message it will induce another paint message callback producing a tight loop
    'with no real screen update.
    Exit Sub
    
  Case WM_NCHITTEST
  
    'Example of using the return value...
    If lReturn = HTCLIENT Then

      'Default processing has determined that the mouse is over the client area.
      '
      'Uncomment the line below to *lie* to windows that the mouse is over the caption
      'area of the form. This means that mouse messages over the client area will be
      'reported as non-client mouse messages. So, for instance, if you click and drag
      'on the client area you can move the form, if you double click the client area
      'the form will toggle between Normal and Maximized... Except that it wont toggle
      'Windowstates in this particular instance because we're filtering out maximise
      'messages in ISuperClass_Before. Nevertheless, uncommenting the line below will
      'allow client click/drag movement of the window.
      '
      'lReturn = HTCAPTION
    End If
    
  End Select

  Call Display("After ", lReturn, hWnd, uMsg, wParam, lParam)
End Sub

'This implemented interface is called BEFORE default processing... that is, BEFORE the previous WndProc
'See iSuperClass for parameter information
Private Sub iSuperClass_Before(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  '
  'We can filter certain messages out by setting lHandled to non-zero.
  '
  'Filter out the messages that'll maximize the form
  Select Case uMsg
  Case WM_SYSCOMMAND
    If wParam = SC_MAXIMIZE Then
    
      lHandled = True
      lReturn = 0                           'WinHelp says to return 0 if this message is handled
    End If

  Case WM_NCLBUTTONDBLCLK:
    lHandled = True
    lReturn = 0                             'WinHelp says to return 0 if this message is handled
  
  End Select
  
  Call Display("Before", lReturn, hWnd, uMsg, wParam, lParam)
End Sub

'Display the parameters... That's all folks! Appart from a little filtering and possible
'click drag on the client area to illustrate the lReturn and lHandled parameters, I'm not
'getting into what you can do with subclassing.
Private Sub Display(sWhen As String, lReturn As Long, hWnd As Long, uMsg As Long, wParam As Long, lParam As Long)
  Static nMsg As Long
  Dim sMsg    As String
  
  nMsg = nMsg + 1
  
  'Again, you'll get small performance improvements by ordering the Case tests by frequency.
  'Greater improvements here because this is VB not machine-code.
  Select Case uMsg
    Case WM_NCHITTEST:        sMsg = "WM_NCHITTEST"
    Case WM_MOUSEMOVE:        sMsg = "WM_MOUSEMOVE"
    Case WM_NCMOUSEMOVE:      sMsg = "WM_NCMOUSEMOVE"
    Case WM_MOVE:             sMsg = "WM_MOVE"
    Case WM_SIZE:             sMsg = "WM_SIZE"
    Case WM_MOUSEWHEEL:       sMsg = "WM_MOUSEWHEEL"
    Case WM_SYSCOMMAND:       sMsg = "WM_SYSCOMMAND"
    Case WM_MOUSEWHEEL:       sMsg = "WM_MOUSEWHEEL"
    Case WM_LBUTTONDOWN:      sMsg = "WM_LBUTTONDOWN"
    Case WM_LBUTTONUP:        sMsg = "WM_LBUTTONUP"
    Case WM_LBUTTONDBLCLK:    sMsg = "WM_LBUTTONDBLCLK"
    Case WM_MBUTTONDOWN:      sMsg = "WM_MBUTTONDOWN"
    Case WM_MBUTTONUP:        sMsg = "WM_MBUTTONUP"
    Case WM_MBUTTONDBLCLK:    sMsg = "WM_MBUTTONDBLCLK"
    Case WM_RBUTTONDOWN:      sMsg = "WM_RBUTTONDOWN"
    Case WM_RBUTTONUP:        sMsg = "WM_RBUTTONUP"
    Case WM_RBUTTONDBLCLK:    sMsg = "WM_RBUTTONDBLCLK"
    Case WM_NCLBUTTONDOWN:    sMsg = "WM_NCLBUTTONDOWN"
    Case WM_NCLBUTTONUP:      sMsg = "WM_NCLBUTTONUP"
    Case WM_NCLBUTTONDBLCLK:  sMsg = "WM_NCLBUTTONDBLCLK"
    Case WM_NCMBUTTONDOWN:    sMsg = "WM_NCMBUTTONDOWN"
    Case WM_NCMBUTTONUP:      sMsg = "WM_NCMBUTTONUP"
    Case WM_NCMBUTTONDBLCLK:  sMsg = "WM_NCMBUTTONDBLCLK"
    Case WM_NCRBUTTONDOWN:    sMsg = "WM_NCRBUTTONDOWN"
    Case Else:                sMsg = "WM_????"
  End Select

  If CurrentY > nLastLine Then

    'Scroll the displayed output up one line
    Call ScrollWindowEx(hWnd, 0, -nTextHeight, rc, rc, 0, ByVal 0&, SW_INVALIDATE)
    Call UpdateWindow(hWnd)
    
    CurrentY = nLastLine
  End If
  
  Print Format((nMsg Mod 1000000), "0##### ") & _
        sWhen & " " & fmt(lReturn) & " " & _
        fmt(hWnd) & " " & _
        fmt(uMsg) & " " & _
        fmt(wParam) & " " & _
        fmt(lParam) & " " & _
        sMsg
End Sub

'Return the Value parameter converted to a hex string padded to 8 characters with a leading &H
Private Function fmt(Value As Long) As String
  Dim s As String
  
  s = Hex$(Value)
  fmt = "&H" & String$(8 - Len(s), "0") & s
End Function
