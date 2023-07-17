VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "貨運貼紙生成"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "Report_UserForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandButton1_Click()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Dim sh3 As Worksheet
Set sh1 = ThisWorkbook.Worksheets("Mark_Final")
Set sh2 = ThisWorkbook.Worksheets("Mark")
'ComboBox1回傳指定工作表
Set sh3 = ThisWorkbook.Worksheets(ComboBox1.Value)
Dim po As String
Dim box_start, box_end, rowEnd As Integer

'取得報表結束的行數
rowEnd = sh3.Cells(Rows.Count, 1).End(xlUp).row

For i = 3 To rowEnd 'Packing_List列數

    box_start = sh3.Cells(i, "E").Value '計算箱號開始
    box_end = sh3.Cells(i, "F").Value   '計算箱號結束
    Dim acounts As Integer
    Dim remaining As Integer
    acounts = sh3.Cells(i, "L").Value   '一箱最多能放多少個數量
    remaining = sh3.Cells(i, "D").Value '需運送幾個數量
    
    For secondi = box_start To box_end
        po = sh3.Cells(i, 2).Value
        'counts = sh3.Cells(i, "D").Value / (box_end - box_start + 1) '出貨數量/總貨號


        
        If remaining >= acounts Then
            remaining = remaining - acounts

            If secondi Mod 2 = 1 Then '若箱號為單數放左邊貼紙
            
                sh2.Cells(4, "B") = po 'po資訊
                sh2.Cells(5, "B").Value = sh3.Cells(i, "C").Value 'part資訊
                sh2.Cells(5, "E").Value = "*" & sh3.Cells(i, "C").Value & "*" '
                sh2.Cells(8, "B").Value = acounts & "PCS/BOX" '出貨數量資訊
                sh2.Cells(8, "E").Value = acounts
                sh2.Cells(11, "B").Value = po & Format(Str(secondi), "0000") '出貨號碼
                sh2.Cells(11, "E").Value = "*" & po & Format(Str(secondi), "0000") & "*"
                If secondi = 1 Then
                    pasterow = 2    '首要貼上的地方為第2列
                Else
                    pasterow = pasterow + 12  '接續後12列貼上
                End If
                
                Sheets(2).Select         '點選第二個工作表"MARK"
                Range("A2:H12").Select
                Selection.Copy
                Sheets(1).Select          '點選第一個工作表"Final_Mark"
                Range("A" & pasterow).Select
                ActiveSheet.Paste          '貼上
                
            Else    '若箱號為雙數放右邊貼紙
               
                sh2.Cells(4, 11) = po
                sh2.Cells(5, 11).Value = sh3.Cells(i, "C").Value 'part資訊
                sh2.Cells(5, 14).Value = "*" & sh3.Cells(i, "C").Value & "*" '
                sh2.Cells(8, 11).Value = acounts & "PCS/BOX" '出貨數量資訊
                sh2.Cells(8, 14).Value = acounts
                sh2.Cells(11, 11).Value = po & Format(Str(secondi), "0000") '出貨號碼
                sh2.Cells(11, 14).Value = "*" & po & Format(Str(secondi), "0000") & "*"
                
                Sheets(2).Select
                Range("J2:Q12").Select
                Selection.Copy
                Sheets(1).Select
                Range("J" & pasterow).Select
                ActiveSheet.Paste
            End If
            
        Else
            remaining = remaining
            If secondi Mod 2 = 1 Then '若箱號為單數放左邊貼紙
            
                sh2.Cells(4, "B") = po  'po資訊
                sh2.Cells(5, "B").Value = sh3.Cells(i, "C").Value 'part資訊
                sh2.Cells(5, "E").Value = "*" & sh3.Cells(i, "C").Value & "*" '
                sh2.Cells(8, "B").Value = remaining & "PCS/BOX" '出貨數量資訊
                sh2.Cells(8, "E").Value = remaining
                sh2.Cells(11, "B").Value = po & Format(Str(secondi), "0000") '出貨號碼
                sh2.Cells(11, "E").Value = "*" & po & Format(Str(secondi), "0000") & "*"
                If secondi = 1 Then
                    pasterow = 2    '首要貼上的地方為第2列
                Else
                    pasterow = pasterow + 12  '接續後12列貼上
                End If
                
                Sheets(2).Select         '點選第二個工作表"MARK"
                Range("A2:H12").Select
                Selection.Copy
                Sheets(1).Select          '點選第一個工作表"Final_Mark"
                Range("A" & pasterow).Select
                ActiveSheet.Paste          '貼上
                
            Else    '若箱號為雙數放右邊貼紙
               
                sh2.Cells(4, 11) = po
                sh2.Cells(5, 11).Value = sh3.Cells(i, "C").Value 'part資訊
                sh2.Cells(5, 14).Value = "*" & sh3.Cells(i, "C").Value & "*" '
                sh2.Cells(8, 11).Value = remaining & "PCS/BOX" '出貨數量資訊
                sh2.Cells(8, 14).Value = remaining
                sh2.Cells(11, 11).Value = po & Format(Str(secondi), "0000") '出貨號碼
                sh2.Cells(11, 14).Value = "*" & po & Format(Str(secondi), "0000") & "*"
                
                Sheets(2).Select
                Range("J2:Q12").Select
                Selection.Copy
                Sheets(1).Select
                Range("J" & pasterow).Select
                ActiveSheet.Paste
            End If
        End If
    Next secondi
     

Next i


Sheets(1).Select
Cells.Select
With Selection.Font
    .Name = "Times New Roman"
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .TintAndShade = 0
    .ThemeFont = xlThemeFontNone
End With

MsgBox ("所有貼紙已生成完畢！")
End Sub


Private Sub CommandButton2_Click()

'
' 巨集2 巨集
'

'
    Sheets(1).Select
    Cells.Select
    Range("A29").Activate
    Selection.Delete Shift:=xlUp
End Sub

Private Sub Label1_Click()

End Sub

Private Sub PrintButton_Click()
'列印現有工作表（即生成之貼紙）
ActiveSheet.PrintOut
End Sub

Private Sub UserForm_Initialize()

Dim shtIdx As Integer
For shtIdx = 3 To Sheets.Count
ComboBox1.AddItem Sheets(shtIdx).Name
Next

End Sub
