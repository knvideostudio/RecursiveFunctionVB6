VERSION 5.00
Begin VB.Form MyForm 
   Caption         =   "Build Categories - XML Structure"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Begin"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox MyText 
      Height          =   3495
      Left            =   0
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "My Text will be here"
      Top             =   600
      Width           =   7095
   End
End
Attribute VB_Name = "MyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *********************************************************************************************************************
' * Author:          Krassimir Nikov
' * Email:           nikov@rokamboll.com
' * Release date:    Apr 03, 2006
' * History:         Apr 03, 2006 - Initial Code
' * Web Site:        www.rokamboll.com/recursivefunc
' * Resume web site: www.rokamboll.com/my_profile.htm
' *********************************************************************************************************************

Option Explicit

Private Sub cmdTest_Click()
Dim ArrRoot() As String
    
    ' Create Root Categories
    MyText.Text = "Waiting ..."
    
    ' create an empty file with Root node
    On Error Resume Next
    Open App.Path & "\" & FILE_NAME For Output Access Write As #44
    Print #44, "<TreeView></TreeView>"
    Close #44
    
    ' get all Root categories
    ArrRoot = BuildCategoriesRoot()
    cmdTest.Enabled = False
 
    ' create the entire category structure
    BuildCatheoriesChildren ArrRoot

    MyText.Text = "Finishing ..."
    cmdTest.Enabled = True
    
End Sub
