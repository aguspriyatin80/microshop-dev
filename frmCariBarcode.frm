VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CARIBARCODE 
   Caption         =   "CARI BARCODE"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6735
   OleObjectBlob   =   "frmCariBarcode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CARIBARCODE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub tampil()
On Error Resume Next
Set ws = Sheets("DATABARANG")
ws.Activate
akhir = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    Me.tbBarang.ColumnCount = 3
    Me.tbBarang.ColumnWidths = "0pt;100pt;200pt"
    Me.tbBarang.RowSource = "a2:g" & akhir + 1
End Sub
Private Sub tbBarang_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
SETTINGBARCODE.cboKodeBarang.Value = Me.tbBarang.Column(1)
SETTINGBARCODE.txtNama.Value = Me.tbBarang.Column(2)
Unload Me
End Sub
Private Sub txtCari_Change()
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
Sheets("DATABARANG").Range("A1:g" & Cells(Rows.Count, 1).End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy _
Destination:=Sheets("HASILFILTER").Range("A1")
Sheets("HASILFILTER").Cells.EntireColumn.AutoFit
Sheets("HASILFILTER").Select
Me.tbBarang.RowSource = "HASILFILTER!a2:g" & Cells(Rows.Count, 1).End(xlUp).Row + 1
Application.ScreenUpdating = True
End If
End Sub
Private Sub UserForm_Initialize()
tampil
End Sub
