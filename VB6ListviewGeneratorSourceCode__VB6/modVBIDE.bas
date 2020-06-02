Attribute VB_Name = "modVBIDE"
Option Explicit

Public VBInstance As VBIDE.VBE

Public Type tListViewType
    GUID As String
    ProgID As String
    ListItemName As String
End Type

Public ListviewType As tListViewType


