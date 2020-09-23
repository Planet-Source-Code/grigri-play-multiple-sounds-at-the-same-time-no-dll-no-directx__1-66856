VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundManagerNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SoundLoaded(ByVal BufferIndex As Long)
    '
End Sub

Public Sub SoundPlayStart(ByVal BufferIndex As Long)
    '
End Sub

Public Sub SoundPlayEnd(ByVal BufferIndex As Long)
    '
End Sub

Public Sub SoundUnloaded(ByVal BufferIndex As Long)
    '
End Sub
