VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Images"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hIcon As Long) As Long

Private Function ImageList_AddIcon(ByVal himl As Long, ByVal hIcon As Long, ByVal iIndex As Integer) As Long
  ImageList_AddIcon = ImageList_ReplaceIcon(himl, iIndex, hIcon)
End Function

Public Sub SetImageListIcons(ImlNormal As ImageList)
Call ImageList_AddIcon(ImlNormal.hImageList, LoadResPicture("HOUSE", vbResIcon), 0)
Call ImageList_AddIcon(ImlNormal.hImageList, LoadResPicture("BUTTERFLY", vbResIcon), 1)
Call ImageList_AddIcon(ImlNormal.hImageList, LoadResPicture("PEOPLE", vbResIcon), 2)
End Sub
