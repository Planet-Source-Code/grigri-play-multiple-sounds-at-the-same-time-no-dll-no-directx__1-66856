VERSION 5.00
Begin VB.Form DemoWindow 
   Caption         =   "Grigri's Sound Manager - Test Application"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   343
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   577
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame I_HATE_FRAME_CONTROLS 
      Caption         =   "Instant Sound"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      TabIndex        =   70
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton btnInstantSound 
         Caption         =   "Instant"
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
         Left            =   1200
         TabIndex        =   71
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblInstant 
         Caption         =   "Click this one as fast as you can!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton btnPlayAll 
      Caption         =   "Play All (!)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   69
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton btnLoadAll 
      Caption         =   "Load All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   68
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton btnFree 
      Caption         =   "Free"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   7920
      TabIndex        =   67
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   7200
      TabIndex        =   66
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   6480
      TabIndex        =   65
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5760
      TabIndex        =   64
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox txtSoundFile 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   1680
      TabIndex        =   63
      Text            =   "Swedish.wav"
      Top             =   4440
      Width           =   3975
   End
   Begin VB.CommandButton btnFree 
      Caption         =   "Free"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   7920
      TabIndex        =   61
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   7200
      TabIndex        =   60
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   6480
      TabIndex        =   59
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   5760
      TabIndex        =   58
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtSoundFile 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   1680
      TabIndex        =   57
      Text            =   "Russian.wav"
      Top             =   4080
      Width           =   3975
   End
   Begin VB.CommandButton btnFree 
      Caption         =   "Free"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   7920
      TabIndex        =   55
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   7200
      TabIndex        =   54
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   6480
      TabIndex        =   53
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   5760
      TabIndex        =   52
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtSoundFile 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   1680
      TabIndex        =   51
      Text            =   "Polish.wav"
      Top             =   4800
      Width           =   3975
   End
   Begin VB.CommandButton btnFree 
      Caption         =   "Free"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   7920
      TabIndex        =   49
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   7200
      TabIndex        =   48
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6480
      TabIndex        =   47
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   5760
      TabIndex        =   46
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox txtSoundFile 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   45
      Text            =   "Norwegian.wav"
      Top             =   3720
      Width           =   3975
   End
   Begin VB.CommandButton btnFree 
      Caption         =   "Free"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   7920
      TabIndex        =   43
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   7200
      TabIndex        =   42
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6480
      TabIndex        =   41
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5760
      TabIndex        =   40
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtSoundFile 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   1680
      TabIndex        =   39
      Text            =   "Irish.wav"
      Top             =   3360
      Width           =   3975
   End
   Begin VB.CommandButton btnFree 
      Caption         =   "Free"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   7920
      TabIndex        =   37
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   7200
      TabIndex        =   36
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6480
      TabIndex        =   35
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   34
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtSoundFile 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1680
      TabIndex        =   33
      Text            =   "Icelandic.wav"
      Top             =   3000
      Width           =   3975
   End
   Begin VB.CommandButton btnFree 
      Caption         =   "Free"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   7920
      TabIndex        =   31
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   7200
      TabIndex        =   30
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6480
      TabIndex        =   29
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   28
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtSoundFile 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   27
      Text            =   "Hebrew.wav"
      Top             =   2640
      Width           =   3975
   End
   Begin VB.CommandButton btnFree 
      Caption         =   "Free"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   25
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7200
      TabIndex        =   24
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   23
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   22
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtSoundFile 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   21
      Text            =   "Danish.wav"
      Top             =   2280
      Width           =   3975
   End
   Begin VB.CommandButton btnFree 
      Caption         =   "Free"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   19
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7200
      TabIndex        =   18
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   17
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   16
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtSoundFile 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   15
      Text            =   "Czech.wav"
      Top             =   1920
      Width           =   3975
   End
   Begin VB.CommandButton btnFree 
      Caption         =   "Free"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   7920
      TabIndex        =   13
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   12
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   11
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtSoundFile 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   9
      Text            =   "Bulgarian.wav"
      Top             =   1560
      Width           =   3975
   End
   Begin VB.CommandButton btnFree 
      Caption         =   "Free"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtSoundFile 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Text            =   "Ukrainian.wav"
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton btnFreeAll 
      Caption         =   "Free All Sounds"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton btnStopAll 
      Caption         =   "Stop All Sounds"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   62
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   56
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   50
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   44
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   38
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   32
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   26
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "DemoWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements SoundManagerNotifier

Private Sub btnLoadAll_Click()
    Dim i As Integer
    On Error Resume Next
    For i = btnLoad.LBound To btnLoad.UBound
        If btnLoad(i).Enabled Then btnLoad_Click i
    Next
End Sub

Private Sub btnPlayAll_Click()
    Dim i As Integer
    On Error Resume Next
    For i = btnPlay.LBound To btnPlay.UBound
        If btnPlay(i).Enabled Then btnPlay_Click i
    Next
End Sub

Private Sub btnStopAll_Click()
    SoundManager.StopSound ALL_SOUND_BUFFERS
End Sub

Private Sub btnFreeAll_Click()
    SoundManager.FreeSound ALL_SOUND_BUFFERS
End Sub

Private Sub btnFree_Click(Index As Integer)
    SoundManager.FreeSound Index
End Sub

Private Sub btnLoad_Click(Index As Integer)
    SoundManager.LoadSoundFile Index, App.Path & "\Sounds\" & txtSoundFile(Index).Text
End Sub

Private Sub btnPlay_Click(Index As Integer)
    SoundManager.PlaySound Index
End Sub

Private Sub btnStop_Click(Index As Integer)
    SoundManager.StopSound Index
End Sub

Private Sub btnInstantSound_Click()
    ' Instantly load, play and free a sound using the first available buffer
    ' No notification is required
    SoundManager.LoadSoundFile SoundManager.FreeBuffer, App.Path & "\Sounds\Oshppis3.wav", BufferFlagInstant
End Sub

Private Sub Form_Load()
    Set SoundManager.Notifier = Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Very, Very Important. This MUST be called or a crash is inevitable
    SoundManager.DestroySoundManager
End Sub

Private Sub SoundManagerNotifier_SoundLoaded(ByVal BufferIndex As Long)
    With lblStatus(BufferIndex)
        .Caption = "Loaded"
        .ForeColor = vbBlue
    End With
    
    btnFree(BufferIndex).Enabled = True
    btnLoad(BufferIndex).Enabled = True
    btnPlay(BufferIndex).Enabled = True
    btnStop(BufferIndex).Enabled = False
End Sub

Private Sub SoundManagerNotifier_SoundPlayEnd(ByVal BufferIndex As Long)
    With lblStatus(BufferIndex)
        .Caption = "Stopped"
        .ForeColor = vbBlue
    End With
    
    btnFree(BufferIndex).Enabled = True
    btnLoad(BufferIndex).Enabled = True
    btnPlay(BufferIndex).Enabled = True
    btnStop(BufferIndex).Enabled = False
End Sub

Private Sub SoundManagerNotifier_SoundPlayStart(ByVal BufferIndex As Long)
    With lblStatus(BufferIndex)
        .Caption = "Playing"
        .ForeColor = vbGreen
    End With
    
    btnFree(BufferIndex).Enabled = True
    btnLoad(BufferIndex).Enabled = True
    btnPlay(BufferIndex).Enabled = False
    btnStop(BufferIndex).Enabled = True
End Sub

Private Sub SoundManagerNotifier_SoundUnloaded(ByVal BufferIndex As Long)
    With lblStatus(BufferIndex)
        .Caption = "Empty"
        .ForeColor = RGB(127, 127, 127)
    End With
    
    btnFree(BufferIndex).Enabled = False
    btnLoad(BufferIndex).Enabled = True
    btnPlay(BufferIndex).Enabled = False
    btnStop(BufferIndex).Enabled = False
End Sub
