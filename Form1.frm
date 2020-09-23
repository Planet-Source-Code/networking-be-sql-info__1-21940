VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SQL-Info"
   ClientHeight    =   5820
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Text            =   "sa"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "SQL Server login"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "NT Security"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Text            =   "employees"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "northwind"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "localhost"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Table"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Database"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Server"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "(very) short info"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    ' clear any previous results
    List1.Clear
    
    On Error GoTo ErrHandler
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim strCon As String
    
    Dim maxLen As Long
    
    cnn.Provider = "SQLOLEDB"
    strCon = "Server=" & Text1.Text & ";Database=" & Text2.Text & ";"
    
    If Option1.Value Then
        ' NT Security
        strCon = strCon & "Trusted_Connection=yes;"
    Else
        ' userid and password
        If Text4.Text <> "" Then
            strCon = strCon & "UID=" & Text4.Text & ";"
            If Text5.Text <> "" Then strCon = strCon & "PWD=" & Text5.Text & ";"
        Else
            MsgBox "Please supply a username"
        End If
    End If
    
    ' open the connection
    cnn.ConnectionString = strCon
    cnn.Open

    ' do a select on the table wich always yields 0 rows
    Set rst = cnn.Execute("SELECT * FROM " & Text3.Text & " WHERE 1=2")

    ' enumerate field to find out longest name (used when adding to list)
    Dim fld As Field
    For Each fld In rst.Fields
        If Len(fld.Name) > maxLen Then maxLen = Len(fld.Name)
    Next fld

    ' longest name +3
    maxLen = maxLen + 3
    
    ' add first row (titles)
    List1.AddItem "Name" & Space(maxLen - 4) & "Type" & Space(11) & "Size" & Space(5) & "Nullable"
    
    For Each fld In rst.Fields
        ' enumerate fields and fill list
        List1.AddItem fld.Name & Space(maxLen - Len(fld.Name)) & _
            GetTypeAndSize(fld.Type, fld.DefinedSize) & _
            IIf((fld.Attributes And adFldIsNullable) = adFldIsNullable, "X", " ")
    Next fld

    ' cleanup
    rst.Close
    cnn.Close

    Exit Sub
    
ErrHandler:

    Select Case Err.Number
    Case -2147467259 ' connection failed
        MsgBox "Failed to connect. Check server and database name"
    Case -2147217865 ' execute failed
        MsgBox "Failed to open the table. Make sure the name is spelled correctly"
    Case Else
        MsgBox "An unexpected error occured: " & vbCrLf & _
            Err.Number & " - " & Err.Description
    End Select

End Sub

Private Sub Command2_Click()

    Dim Fn, Fs
    
    ' first get data
    Command1_Click
    
    If List1.ListCount = 0 Then Exit Sub
    
    ' save printer settings
    Fn = Printer.FontName
    Fs = Printer.FontSize
    
    ' change settings
    Printer.FontName = List1.FontName
    Printer.FontSize = List1.FontSize
    
    ' send everything to printer
    Printer.Print "Table info for " & Text3.Text
    Printer.Print
    
    Dim T As Integer
    For T = 0 To List1.ListCount - 1
        Printer.Print "    " & List1.List(T)
    Next T

    Printer.EndDoc

    ' restore printer settings
    Printer.FontName = Fn
    Printer.FontSize = Fs

End Sub

Private Sub mnuFileExit_Click()

    End

End Sub

Private Sub mnuHelpAbout_Click()

    MsgBox "Created by Cakkie" & vbCrLf & "slisse@planetinternet.be", , "SQLInfo"

End Sub

Private Sub mnuHelpInfo_Click()

    MsgBox "Fill in the name of the server, the name of the database and the name of the table you want to examine. " & _
           "Select the way you want to log on to the server. Supply username and password (if not using NT Security). " & _
           "Press show to get the list on the screen, press print to send it to the printer.", , "SQLInfo"

End Sub

Private Sub Option1_Click()

    Text4.Enabled = False
    Text5.Enabled = False

End Sub

Private Sub Option2_Click()

    Text4.Enabled = True
    Text5.Enabled = True

End Sub

Private Function GetTypeAndSize(T As Long, S As Long) As String

    ' this gives a string vale for a fieldtype
    ' adds length
    
    Dim G

    Select Case T
    Case adArray
        G = "array"
    Case adBigInt
        G = "bigint"
    Case adBoolean
        G = "boolean"
    Case adBSTR
        G = "bstr"
    Case adChapter
        G = "chapter"
    Case adChar
        G = "char"
    Case adCurrency
        G = "currency"
    Case adDate
        G = "date"
    Case adDBDate
        G = "date"
    Case adDBTime
        G = "time"
    Case adDBTimeStamp
        G = "date/time"
    Case adDecimal
        G = "decimal"
    Case adDouble
        G = "double"
    Case adGUID
        G = "GUID"
    Case adInteger
        G = "int"
    Case adLongVarBinary
        G = "binary"
    Case adLongVarWChar
        G = "wtext"
    Case adLongVarChar
        G = "text"
    Case adNumeric
        G = "numeric"
    Case adSingle
        G = "single"
    Case adSmallInt
        G = "smallint"
    Case adTinyInt
        G = "tinyint"
    Case adUnsignedBigInt
        G = "uns. bigint"
    Case adUnsignedInt
        G = "uns. int"
    Case adUnsignedSmallInt
        G = "uns. smallint"
    Case adUnsignedTinyInt
        G = "uns. tinyint"
    Case adUserDefined
        G = "userdefined"
    Case adVarBinary
        G = "varbinary"
    Case adVarChar
        G = "varchar"
    Case adVariant
        G = "variant"
    Case adVarNumeric
        G = "varnumeric"
    Case adVarWChar
        G = "varwchar"
    Case adWChar
        G = "wchar"
    Case Else
        G = "unknown"
    End Select

    GetTypeAndSize = G & Space(15 - Len(G)) & S & Space(13 - Len(CStr(S)))

End Function
