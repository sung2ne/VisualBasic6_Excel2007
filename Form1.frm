VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_B2 
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txt_B1 
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txt_A2 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txt_A1 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Btn_Read 
      Caption         =   "불러오기"
      Height          =   615
      Index           =   1
      Left            =   3840
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Btn_Save 
      Caption         =   "저장하기"
      Height          =   615
      Index           =   0
      Left            =   5640
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lb_B1 
      Alignment       =   1  'Right Justify
      Caption         =   "B1 :"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lb_B2 
      Alignment       =   1  'Right Justify
      Caption         =   "B2 :"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lb_A2 
      Alignment       =   1  'Right Justify
      Caption         =   "A2 :"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lb_A1 
      Alignment       =   1  'Right Justify
      Caption         =   "A1 :"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object

Private Sub Btn_Read_Click(Index As Integer)

    oExcel.Workbooks.Open App.Path & "\" & "test.xls"
    Set oSheet = oExcel.Workbooks(1).Sheets(1)
    
    txt_A1.Text = oSheet.Range("A1").Value
    txt_A2.Text = oSheet.Range("A2").Value
    txt_B1.Text = oSheet.Range("B1").Value
    txt_B2.Text = oSheet.Range("B2").Value
    
    oExcel.Quit
    
End Sub

Private Sub Btn_Save_Click(Index As Integer)
  
    Set oBook = oExcel.Workbooks.Add
    
    'Add data to cells of the first worksheet in the new workbook
    Set oSheet = oBook.Worksheets(1)
    oSheet.Range("A1").Value = txt_A1.Text
    oSheet.Range("A2").Value = txt_A2.Text
    oSheet.Range("B1").Value = txt_B1.Text
    oSheet.Range("B2").Value = txt_B2.Text
    
    'Save the Workbook and Quit Excel
    oBook.SaveAs App.Path & "\" & "test.xls"
    oExcel.Quit
    
End Sub

Private Sub Form_Load()

    'Start a new workbook in Excel
    Set oExcel = CreateObject("Excel.Application")
            
End Sub

