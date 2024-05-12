VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CARIBARANG 
   Caption         =   "CARI BARANG"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6765
   OleObjectBlob   =   "frmCariBarang.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CARIBARANG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub masukkanKeTabelTransaksi()
On Error Resume Next
If Val(Me.tbBarang.Column(12)) < 1 Then
    MsgBox "Stok barang kosong!"
    Exit Sub
End If
DASHBOARD2.txtScan = Me.tbBarang.Column(1)
DASHBOARD2.txtScan.SetFocus
Application.SendKeys ("~")
Unload Me
End Sub
Private Sub tbBarang_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    masukkanKeTabelTransaksi
End Sub
Sub tampil()
On Error Resume Next
Set ws = Sheets("DATABARANG")
ws.Activate
akhir = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    Me.tbBarang.RowSource = "a2:m" & akhir
End Sub
Private Sub tbBarang_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub
Private Sub tbBarang_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    masukkanKeTabelTransaksi
End If
End Sub
Private Sub txtCari_AfterUpdate()
On Error Resume Next
    If Me.txtCari.Value = "" Then
    With Sheets("DATABARANG")
    .AutoFilterMode = False
    End With
    tampil
    Else
    Sheets("DATABARANG").Activate
    Application.ScreenUpdating = False
    Sheets("DATABARANG").Range("A1:G1").AutoFilter Field:=3, Criteria1:="*" & Me.txtCari.Text & "*", Operator:=xlAnd
    Sheets("HASILFILTER").Cells.clear
    Sheets("DATABARANG").Range("A1:m" & Cells(Rows.Count, 1).End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy _
    Destination:=Sheets("HASILFILTER").Range("A1")
    Sheets("HASILFILTER").Cells.EntireColumn.AutoFit
    Sheets("HASILFILTER").Select
    Me.tbBarang.RowSource = "HASILFILTER!a2:m" & Cells(Rows.Count, 1).End(xlUp).Row
    Application.ScreenUpdating = True
    End If
End Sub
Private Sub txtCari_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub
Private Sub UserForm_Initialize()
Call tampil
txtCari.SetFocus
Sheets("DATABARANG").AutoFilterMode = False
tbBarang.ColumnCount = 13
tbBarang.ColumnWidths = ("0pt;50pt;180pt;0pt;0pt;60pt;0pt;0pt;0pt;0pt;0pt;0pt;60pt")
End Sub
