VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SETTINGBARCODE 
   Caption         =   "SETTING & CETAK BARCODE"
   ClientHeight    =   9075.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15210
   OleObjectBlob   =   "frmSettingBarcode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SETTINGBARCODE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws As Worksheet
Private Sub cboKodeBarang_Change()
On Error Resume Next
Dim rg As Range
Dim cIndex As Variant
Dim kode As Double
lastRow = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
Set rg = Sheets("DATABARANG").Range("B2:i" & lastRow)
kode = Val(Me.cboKodeBarang.Value)
cIndex = Application.Match(kode, rg.Columns(1), 0)
If Len(cIndex) > 0 Then
    Me.txtNama.Text = rg.Cells(cIndex, 2).Value
End If
End Sub
Private Sub cmdCariKode_Click()
CARIBARCODE.Show
End Sub
Private Sub cmdCetakBarcode_Click()
On Error Resume Next
If Me.cboKodeBarang.Value = "" Then
    MsgBox "Isi dulu kode barangnya!"
    Exit Sub
Else
Sheets("DATABARANG").Activate
Sheets("DATABARANG").AutoFilterMode = False
lastRow = Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row
Set findID = Sheets("DATABARANG").Range("b1:k" & lastRow).Find(what:=Me.cboKodeBarang.Value, LookIn:=xlValues)
If findID Is Nothing Then
MsgBox "ID tidak ditemukan"
Else
    Sheets("CETAKBARCODE2").Range("A1").Value = Me.cboKodeBarang.Value
    Sheets("CETAKBARCODE2").Range("B1").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("B4").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("B7").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("B10").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("B13").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("B16").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("B19").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("B22").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("B5").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("B2").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("B5").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("B8").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("B11").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("B14").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("B17").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("B20").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("B23").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("B26").Value = findID.Offset(0, 4).Value
    
    Sheets("CETAKBARCODE2").Range("C1").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("C4").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("C7").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("C10").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("C13").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("C16").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("C19").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("C22").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("C5").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("C2").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("C5").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("C8").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("C11").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("C14").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("C17").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("C20").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("C23").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("C26").Value = findID.Offset(0, 4).Value
                
    Sheets("CETAKBARCODE2").Range("E1").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("E4").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("E7").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("E10").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("E13").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("E16").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("E19").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("E22").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("E5").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("E2").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("E5").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("E8").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("E11").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("E14").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("E17").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("E20").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("E23").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("E26").Value = findID.Offset(0, 4).Value
                
    Sheets("CETAKBARCODE2").Range("G1").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("G4").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("G7").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("G10").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("G13").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("G16").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("G19").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("G22").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("G5").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("G2").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("G5").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("G8").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("G11").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("G14").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("G17").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("G20").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("G23").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("G26").Value = findID.Offset(0, 4).Value
                
    Sheets("CETAKBARCODE2").Range("I1").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("I4").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("I7").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("I10").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("I13").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("I16").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("I19").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("I22").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("I5").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("I2").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("I5").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("I8").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("I11").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("I14").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("I17").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("I20").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("I23").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("I26").Value = findID.Offset(0, 4).Value
                
    Sheets("CETAKBARCODE2").Range("K1").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("K4").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("K7").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("K10").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("K13").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("K16").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("K19").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("K22").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("K5").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("K2").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("K5").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("K8").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("K11").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("K14").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("K17").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("K20").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("K23").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("K26").Value = findID.Offset(0, 4).Value
                
    Sheets("CETAKBARCODE2").Range("M1").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("M4").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("M7").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("M10").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("M13").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("M16").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("M19").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("M22").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("M5").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("M2").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("M5").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("M8").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("M11").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("M14").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("M17").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("M20").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("M23").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("M26").Value = findID.Offset(0, 4).Value
                
    Sheets("CETAKBARCODE2").Range("O1").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("O4").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("O7").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("O10").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("O13").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("O16").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("O19").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("O22").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("O5").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("O2").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("O5").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("O8").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("O11").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("O14").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("O17").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("O20").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("O23").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("O26").Value = findID.Offset(0, 4).Value
                
    Sheets("CETAKBARCODE2").Range("Q1").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("Q4").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("Q7").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("Q10").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("Q13").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("Q16").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("Q19").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("Q22").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("Q5").Font.Name = "IDAHC39M Code 39 Barcode"
    Sheets("CETAKBARCODE2").Range("Q2").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("Q5").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("Q8").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("Q11").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("Q14").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("Q17").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("Q20").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("Q23").Value = findID.Offset(0, 4).Value
    Sheets("CETAKBARCODE2").Range("Q26").Value = findID.Offset(0, 4).Value
    
    Sheets("CETAKBARCODE2").Range("g35").Value = findID.Offset(0, 0).Value
    Sheets("CETAKBARCODE2").Range("g36").Value = findID.Offset(0, 1).Value
                
    Sheets("CETAKBARCODE2").Activate
    Application.Dialogs(xlDialogPrint).Show
End If
End If
End Sub
Private Sub UserForm_Initialize()
Me.cboKodeBarang.clear
kodeAwal = Sheets("DATABARANG").Range("b2").Value
lastRow = Val(Sheets("DATABARANG").Cells(Rows.Count, 1).End(xlUp).Row)
For i = 2 To lastRow
Me.cboKodeBarang.AddItem Sheets("DATABARANG").Range("b" & i).Text
Next i
Me.cboKodeBarang.Text = kodeAwal

Me.Label3.ZOrder 1
Me.Label4.ZOrder 1
Me.Label5.ZOrder 1

Set ws = Sheets("CETAKBARCODE2")
Me.B1.Caption = ""
ws.Range("B1").Font.Color = vbWhite
Me.C1.Caption = ""
ws.Range("C1").Font.Color = vbWhite
Me.E1.Caption = ""
ws.Range("E1").Font.Color = vbWhite
Me.G1.Caption = ""
ws.Range("G1").Font.Color = vbWhite
Me.I1.Caption = ""
ws.Range("I1").Font.Color = vbWhite
Me.K1.Caption = ""
ws.Range("K1").Font.Color = vbWhite
Me.M1.Caption = ""
ws.Range("M1").Font.Color = vbWhite
Me.O1.Caption = ""
ws.Range("O1").Font.Color = vbWhite
Me.Q1.Caption = ""
ws.Range("Q1").Font.Color = vbWhite
Me.B4.Caption = ""
ws.Range("B4").Font.Color = vbWhite
Me.C4.Caption = ""
ws.Range("C4").Font.Color = vbWhite
Me.E4.Caption = ""
ws.Range("E4").Font.Color = vbWhite
Me.G4.Caption = ""
ws.Range("G4").Font.Color = vbWhite
Me.I4.Caption = ""
ws.Range("I4").Font.Color = vbWhite
Me.K4.Caption = ""
ws.Range("K4").Font.Color = vbWhite
Me.M4.Caption = ""
ws.Range("M4").Font.Color = vbWhite
Me.O4.Caption = ""
ws.Range("O4").Font.Color = vbWhite
Me.Q4.Caption = ""
ws.Range("Q4").Font.Color = vbWhite
Me.B7.Caption = ""
ws.Range("B7").Font.Color = vbWhite
Me.C7.Caption = ""
ws.Range("C7").Font.Color = vbWhite
Me.E7.Caption = ""
ws.Range("E7").Font.Color = vbWhite
Me.G7.Caption = ""
ws.Range("G7").Font.Color = vbWhite
Me.I7.Caption = ""
ws.Range("I7").Font.Color = vbWhite
Me.K7.Caption = ""
ws.Range("K7").Font.Color = vbWhite
Me.M7.Caption = ""
ws.Range("M7").Font.Color = vbWhite
Me.O7.Caption = ""
ws.Range("O7").Font.Color = vbWhite
Me.Q7.Caption = ""
ws.Range("Q7").Font.Color = vbWhite
Me.B10.Caption = ""
ws.Range("B10").Font.Color = vbWhite
Me.C10.Caption = ""
ws.Range("C10").Font.Color = vbWhite
Me.E10.Caption = ""
ws.Range("E10").Font.Color = vbWhite
Me.G10.Caption = ""
ws.Range("G10").Font.Color = vbWhite
Me.I10.Caption = ""
ws.Range("I10").Font.Color = vbWhite
Me.K10.Caption = ""
ws.Range("K10").Font.Color = vbWhite
Me.M10.Caption = ""
ws.Range("M10").Font.Color = vbWhite
Me.O10.Caption = ""
ws.Range("O10").Font.Color = vbWhite
Me.Q10.Caption = ""
ws.Range("Q10").Font.Color = vbWhite
Me.B13.Caption = ""
ws.Range("B13").Font.Color = vbWhite
Me.C13.Caption = ""
ws.Range("C13").Font.Color = vbWhite
Me.E13.Caption = ""
ws.Range("E13").Font.Color = vbWhite
Me.G13.Caption = ""
ws.Range("G13").Font.Color = vbWhite
Me.I13.Caption = ""
ws.Range("I13").Font.Color = vbWhite
Me.K13.Caption = ""
ws.Range("K13").Font.Color = vbWhite
Me.M13.Caption = ""
ws.Range("M13").Font.Color = vbWhite
Me.O13.Caption = ""
ws.Range("O13").Font.Color = vbWhite
Me.Q13.Caption = ""
ws.Range("Q13").Font.Color = vbWhite
Me.B16.Caption = ""
ws.Range("B16").Font.Color = vbWhite
Me.C16.Caption = ""
ws.Range("C16").Font.Color = vbWhite
Me.E16.Caption = ""
ws.Range("E16").Font.Color = vbWhite
Me.G16.Caption = ""
ws.Range("G16").Font.Color = vbWhite
Me.I16.Caption = ""
ws.Range("I16").Font.Color = vbWhite
Me.K16.Caption = ""
ws.Range("K16").Font.Color = vbWhite
Me.M16.Caption = ""
ws.Range("M16").Font.Color = vbWhite
Me.O16.Caption = ""
ws.Range("O16").Font.Color = vbWhite
Me.Q16.Caption = ""
ws.Range("Q16").Font.Color = vbWhite
Me.B19.Caption = ""
ws.Range("B19").Font.Color = vbWhite
Me.C19.Caption = ""
ws.Range("C19").Font.Color = vbWhite
Me.E19.Caption = ""
ws.Range("E19").Font.Color = vbWhite
Me.G19.Caption = ""
ws.Range("G19").Font.Color = vbWhite
Me.I19.Caption = ""
ws.Range("I19").Font.Color = vbWhite
Me.K19.Caption = ""
ws.Range("K19").Font.Color = vbWhite
Me.M19.Caption = ""
ws.Range("M19").Font.Color = vbWhite
Me.O19.Caption = ""
ws.Range("O19").Font.Color = vbWhite
Me.Q19.Caption = ""
ws.Range("Q19").Font.Color = vbWhite
Me.B22.Caption = ""
ws.Range("B22").Font.Color = vbWhite
Me.C22.Caption = ""
ws.Range("C22").Font.Color = vbWhite
Me.E22.Caption = ""
ws.Range("E22").Font.Color = vbWhite
Me.G22.Caption = ""
ws.Range("G22").Font.Color = vbWhite
Me.I22.Caption = ""
ws.Range("I22").Font.Color = vbWhite
Me.K22.Caption = ""
ws.Range("K22").Font.Color = vbWhite
Me.M22.Caption = ""
ws.Range("M22").Font.Color = vbWhite
Me.O22.Caption = ""
ws.Range("O22").Font.Color = vbWhite
Me.Q22.Caption = ""
ws.Range("Q22").Font.Color = vbWhite
Me.B25.Caption = ""
ws.Range("B25").Font.Color = vbWhite
Me.C25.Caption = ""
ws.Range("C25").Font.Color = vbWhite
Me.E25.Caption = ""
ws.Range("E25").Font.Color = vbWhite
Me.G25.Caption = ""
ws.Range("G25").Font.Color = vbWhite
Me.I25.Caption = ""
ws.Range("I25").Font.Color = vbWhite
Me.K25.Caption = ""
ws.Range("K25").Font.Color = vbWhite
Me.M25.Caption = ""
ws.Range("M25").Font.Color = vbWhite
Me.O25.Caption = ""
ws.Range("O25").Font.Color = vbWhite
Me.Q25.Caption = ""
ws.Range("Q25").Font.Color = vbWhite
End Sub

'----------------------------------
Private Sub CheckBox1_Click()
If Me.chk_q7.Value = True Then
    Me.Q7.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("Q7").Font.Color = &H80000008
Else
    Me.Q7.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("Q7").Font.Color = vbWhite
End If
End Sub

Private Sub CheckBox2_Click()
If Me.chk_q10.Value = True Then
    Me.Q10.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("Q10").Font.Color = &H80000008
Else
    Me.Q10.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("Q10").Font.Color = vbWhite
End If
End Sub

Private Sub B1_Click()
If Me.chk_b1.Value = True Then
    Me.chk_b1.Value = False
Else
    Me.chk_b1.Value = True
End If
End Sub

Private Sub B10_Click()
If Me.chk_b10.Value = True Then
    Me.chk_b10.Value = False
Else
    Me.chk_b10.Value = True
End If
End Sub

Private Sub B13_Click()
If Me.chk_b13.Value = True Then
    Me.chk_b13.Value = False
Else
    Me.chk_b13.Value = True
End If
End Sub

Private Sub B16_Click()
If Me.chk_b16.Value = True Then
    Me.chk_b16.Value = False
Else
    Me.chk_b16.Value = True
End If
End Sub

Private Sub B19_Click()
If Me.chk_b19.Value = True Then
    Me.chk_b19.Value = False
Else
    Me.chk_b19.Value = True
End If
End Sub

Private Sub B22_Click()
If Me.chk_b22.Value = True Then
    Me.chk_b22.Value = False
Else
    Me.chk_b22.Value = True
End If
End Sub

Private Sub B25_Click()
If Me.chk_b25.Value = True Then
    Me.chk_b25.Value = False
Else
    Me.chk_b25.Value = True
End If
End Sub

Private Sub B4_Click()
If Me.chk_b4.Value = True Then
    Me.chk_b4.Value = False
Else
    Me.chk_b4.Value = True
End If
End Sub

Private Sub B7_Click()
If Me.chk_b7.Value = True Then
    Me.chk_b7.Value = False
Else
    Me.chk_b7.Value = True
End If
End Sub

Private Sub C1_Click()
If Me.chk_c1.Value = True Then
    Me.chk_c1.Value = False
Else
    Me.chk_c1.Value = True
End If
End Sub

Private Sub C10_Click()
If Me.chk_c10.Value = True Then
    Me.chk_c10.Value = False
Else
    Me.chk_c10.Value = True
End If
End Sub

Private Sub C13_Click()
If Me.chk_c13.Value = True Then
    Me.chk_c13.Value = False
Else
    Me.chk_c13.Value = True
End If
End Sub

Private Sub C16_Click()
If Me.chk_c16.Value = True Then
    Me.chk_c16.Value = False
Else
    Me.chk_c16.Value = True
End If
End Sub

Private Sub C19_Click()
If Me.chk_c19.Value = True Then
    Me.chk_c19.Value = False
Else
    Me.chk_c19.Value = True
End If
End Sub

Private Sub C22_Click()
If Me.chk_c22.Value = True Then
    Me.chk_c22.Value = False
Else
    Me.chk_c22.Value = True
End If
End Sub

Private Sub C25_Click()
If Me.chk_c25.Value = True Then
    Me.chk_c25.Value = False
Else
    Me.chk_c25.Value = True
End If
End Sub

Private Sub C4_Click()
If Me.chk_c4.Value = True Then
    Me.chk_c4.Value = False
Else
    Me.chk_c4.Value = True
End If
End Sub

Private Sub C7_Click()
If Me.chk_c7.Value = True Then
    Me.chk_c7.Value = False
Else
    Me.chk_c7.Value = True
End If
End Sub

Private Sub chk_b1_Click()
If Me.chk_b1.Value = True Then
    Me.B1.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("B1").Font.Color = &H80000008
Else
    Me.B1.BackColor = &H8000000F
    '&H8000000F&
    Sheets("CETAKBARCODE2").Range("B1").Font.Color = vbWhite
End If
End Sub

Private Sub chk_b4_Click()
If Me.chk_b4.Value = True Then
    Me.B4.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("B4").Font.Color = &H80000008
Else
    Me.B4.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("B4").Font.Color = vbWhite
End If

End Sub

Private Sub chk_b7_Click()
If Me.chk_b7.Value = True Then
    Me.B7.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("B7").Font.Color = &H80000008
Else
    Me.B7.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("B7").Font.Color = vbWhite
End If
End Sub

Private Sub chk_b10_Click()
If Me.chk_b10.Value = True Then
    Me.B10.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("B10").Font.Color = &H80000008
Else
    Me.B10.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("B10").Font.Color = vbWhite
End If
End Sub

Private Sub chk_b13_Click()
If Me.chk_b13.Value = True Then
    Me.B13.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("B13").Font.Color = &H80000008
Else
    Me.B13.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("B13").Font.Color = vbWhite
End If
End Sub

Private Sub chk_b16_Click()
If Me.chk_b16.Value = True Then
    Me.B16.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("B16").Font.Color = &H80000008
Else
    Me.B16.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("B16").Font.Color = vbWhite
End If
End Sub

Private Sub chk_b19_Click()
If Me.chk_b19.Value = True Then
    Me.B19.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("B19").Font.Color = &H80000008
Else
    Me.B19.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("B19").Font.Color = vbWhite
End If
End Sub

Private Sub chk_b22_Click()
If Me.chk_b22.Value = True Then
    Me.B22.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("B22").Font.Color = &H80000008
Else
    Me.B22.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("B22").Font.Color = vbWhite
End If
End Sub

Private Sub chk_b25_Click()
If Me.chk_b25.Value = True Then
    Me.B25.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("B25").Font.Color = &H80000008
Else
    Me.B25.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("B25").Font.Color = vbWhite
End If
End Sub

Private Sub chk_c1_Click()
If Me.chk_c1.Value = True Then
    Me.C1.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("C1").Font.Color = &H80000008
Else
    Me.C1.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("C1").Font.Color = vbWhite
End If
End Sub

Private Sub chk_c10_Click()
If Me.chk_c10.Value = True Then
    Me.C10.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("C10").Font.Color = &H80000008
Else
    Me.C10.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("C10").Font.Color = vbWhite
End If
End Sub

Private Sub chk_c13_Click()
If Me.chk_c13.Value = True Then
    Me.C13.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("C13").Font.Color = &H80000008
Else
    Me.C13.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("C13").Font.Color = vbWhite
End If
End Sub

Private Sub chk_c16_Click()
If Me.chk_c16.Value = True Then
    Me.C16.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("C16").Font.Color = &H80000008
Else
    Me.C16.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("C16").Font.Color = vbWhite
End If
End Sub

Private Sub chk_c19_Click()
If Me.chk_c19.Value = True Then
    Me.C19.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("C19").Font.Color = &H80000008
Else
    Me.C19.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("C19").Font.Color = vbWhite
End If
End Sub

Private Sub chk_c22_Click()
If Me.chk_c22.Value = True Then
    Me.C22.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("C22").Font.Color = &H80000008
Else
    Me.C22.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("C22").Font.Color = vbWhite
End If
End Sub

Private Sub chk_c25_Click()
If Me.chk_c25.Value = True Then
    Me.C25.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("C25").Font.Color = &H80000008
Else
    Me.C25.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("C25").Font.Color = vbWhite
End If
End Sub

Private Sub chk_c4_Click()
If Me.chk_c4.Value = True Then
    Me.C4.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("C4").Font.Color = &H80000008
Else
    Me.C4.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("C4").Font.Color = vbWhite
End If
End Sub

Private Sub chk_c7_Click()
If Me.chk_c7.Value = True Then
    Me.C7.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("C7").Font.Color = &H80000008
Else
    Me.C7.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("C7").Font.Color = vbWhite
End If
End Sub

Private Sub chk_e1_Click()
If Me.chk_e1.Value = True Then
    Me.E1.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("E1").Font.Color = &H80000008
Else
    Me.E1.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("E1").Font.Color = vbWhite
End If
End Sub

Private Sub chk_e10_Click()
If Me.chk_e10.Value = True Then
    Me.E10.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("E10").Font.Color = &H80000008
Else
    Me.E10.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("E10").Font.Color = vbWhite
End If
End Sub

Private Sub chk_e13_Click()
If Me.chk_e13.Value = True Then
    Me.E13.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("E13").Font.Color = &H80000008
Else
    Me.E13.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("E13").Font.Color = vbWhite
End If
End Sub

Private Sub chk_e16_Click()
If Me.chk_e16.Value = True Then
    Me.E16.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("E16").Font.Color = &H80000008
Else
    Me.E16.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("E16").Font.Color = vbWhite
End If
End Sub

Private Sub chk_e19_Click()
If Me.chk_e19.Value = True Then
    Me.E19.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("E19").Font.Color = &H80000008
Else
    Me.E19.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("E19").Font.Color = vbWhite
End If
End Sub

Private Sub chk_e22_Click()
If Me.chk_e22.Value = True Then
    Me.E22.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("E22").Font.Color = &H80000008
Else
    Me.E22.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("E22").Font.Color = vbWhite
End If
End Sub

Private Sub chk_e25_Click()
If Me.chk_e25.Value = True Then
    Me.E25.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("E25").Font.Color = &H80000008
Else
    Me.E25.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("E25").Font.Color = vbWhite
End If
End Sub

Private Sub chk_e4_Click()
If Me.chk_e4.Value = True Then
    Me.E4.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("E4").Font.Color = &H80000008
Else
    Me.E4.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("E4").Font.Color = vbWhite
End If
End Sub

Private Sub chk_e7_Click()
If Me.chk_e7.Value = True Then
    Me.E7.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("E7").Font.Color = &H80000008
Else
    Me.E7.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("E7").Font.Color = vbWhite
End If
End Sub

Private Sub chk_g1_Click()
If Me.chk_g1.Value = True Then
    Me.G1.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("G1").Font.Color = &H80000008
Else
    Me.G1.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("G1").Font.Color = vbWhite
End If
End Sub

Private Sub chk_g10_Click()
If Me.chk_g10.Value = True Then
    Me.G10.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("G10").Font.Color = &H80000008
Else
    Me.G10.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("G10").Font.Color = vbWhite
End If
End Sub

Private Sub chk_g13_Click()
If Me.chk_g13.Value = True Then
    Me.G13.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("G13").Font.Color = &H80000008
Else
    Me.G13.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("G13").Font.Color = vbWhite
End If
End Sub

Private Sub chk_g16_Click()
If Me.chk_g16.Value = True Then
    Me.G16.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("G16").Font.Color = &H80000008
Else
    Me.G16.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("G16").Font.Color = vbWhite
End If
End Sub

Private Sub chk_g19_Click()
If Me.chk_g19.Value = True Then
    Me.G19.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("G19").Font.Color = &H80000008
Else
    Me.G19.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("G19").Font.Color = vbWhite
End If
End Sub

Private Sub chk_g22_Click()
If Me.chk_g22.Value = True Then
    Me.G22.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("G22").Font.Color = &H80000008
Else
    Me.G22.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("G22").Font.Color = vbWhite
End If
End Sub

Private Sub chk_g25_Click()
If Me.chk_g25.Value = True Then
    Me.G25.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("G25").Font.Color = &H80000008
Else
    Me.G25.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("G25").Font.Color = vbWhite
End If
End Sub

Private Sub chk_g4_Click()
If Me.chk_g4.Value = True Then
    Me.G4.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("G4").Font.Color = &H80000008
Else
    Me.G4.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("G4").Font.Color = vbWhite
End If
End Sub

Private Sub chk_g7_Click()
If Me.chk_g7.Value = True Then
    Me.G7.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("G7").Font.Color = &H80000008
Else
    Me.G7.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("G7").Font.Color = vbWhite
End If
End Sub

Private Sub chk_i1_Click()
If Me.chk_i1.Value = True Then
    Me.I1.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("I1").Font.Color = &H80000008
Else
    Me.I1.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("I1").Font.Color = vbWhite
End If
End Sub

Private Sub chk_i10_Click()
If Me.chk_i10.Value = True Then
    Me.I10.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("I10").Font.Color = &H80000008
Else
    Me.I10.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("I10").Font.Color = vbWhite
End If
End Sub

Private Sub chk_i13_Click()
If Me.chk_i13.Value = True Then
    Me.I13.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("I13").Font.Color = &H80000008
Else
    Me.I13.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("I13").Font.Color = vbWhite
End If
End Sub

Private Sub chk_i16_Click()
If Me.chk_i16.Value = True Then
    Me.I16.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("I16").Font.Color = &H80000008
Else
    Me.I16.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("I16").Font.Color = vbWhite
End If
End Sub

Private Sub chk_i19_Click()
If Me.chk_i19.Value = True Then
    Me.I19.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("I19").Font.Color = &H80000008
Else
    Me.I19.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("I19").Font.Color = vbWhite
End If
End Sub

Private Sub chk_i22_Click()
If Me.chk_i22.Value = True Then
    Me.I22.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("I22").Font.Color = &H80000008
Else
    Me.I22.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("I22").Font.Color = vbWhite
End If
End Sub

Private Sub chk_i25_Click()
If Me.chk_i25.Value = True Then
    Me.I25.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("I25").Font.Color = &H80000008
Else
    Me.I25.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("I25").Font.Color = vbWhite
End If
End Sub

Private Sub chk_i4_Click()
If Me.chk_i4.Value = True Then
    Me.I4.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("I4").Font.Color = &H80000008
Else
    Me.I4.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("I4").Font.Color = vbWhite
End If
End Sub

Private Sub chk_i7_Click()
If Me.chk_i7.Value = True Then
    Me.I7.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("I7").Font.Color = &H80000008
Else
    Me.I7.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("I7").Font.Color = vbWhite
End If
End Sub

Private Sub chk_k1_Click()
If Me.chk_k1.Value = True Then
    Me.K1.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("K1").Font.Color = &H80000008
Else
    Me.K1.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("K1").Font.Color = vbWhite
End If
End Sub

Private Sub chk_k10_Click()
If Me.chk_k10.Value = True Then
    Me.K10.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("K10").Font.Color = &H80000008
Else
    Me.K10.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("K10").Font.Color = vbWhite
End If
End Sub

Private Sub chk_k13_Click()
If Me.chk_k13.Value = True Then
    Me.K13.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("K13").Font.Color = &H80000008
Else
    Me.K13.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("K13").Font.Color = vbWhite
End If
End Sub

Private Sub chk_k16_Click()
If Me.chk_k16.Value = True Then
    Me.K16.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("K16").Font.Color = &H80000008
Else
    Me.K16.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("K16").Font.Color = vbWhite
End If
End Sub

Private Sub chk_k19_Click()
If Me.chk_k19.Value = True Then
    Me.K19.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("K19").Font.Color = &H80000008
Else
    Me.K19.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("K19").Font.Color = vbWhite
End If
End Sub

Private Sub chk_k22_Click()
If Me.chk_k22.Value = True Then
    Me.K22.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("K22").Font.Color = &H80000008
Else
    Me.K22.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("K22").Font.Color = vbWhite
End If
End Sub

Private Sub chk_k25_Click()
If Me.chk_k25.Value = True Then
    Me.K25.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("K25").Font.Color = &H80000008
Else
    Me.K25.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("K25").Font.Color = vbWhite
End If
End Sub

Private Sub chk_k4_Click()
If Me.chk_k4.Value = True Then
    Me.K4.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("K4").Font.Color = &H80000008
Else
    Me.K4.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("K4").Font.Color = vbWhite
End If
End Sub

Private Sub chk_k7_Click()
If Me.chk_k7.Value = True Then
    Me.K7.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("K7").Font.Color = &H80000008
Else
    Me.K7.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("K7").Font.Color = vbWhite
End If
End Sub

Private Sub chk_m1_Click()
If Me.chk_m1.Value = True Then
    Me.M1.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("M1").Font.Color = &H80000008
Else
    Me.M1.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("M1").Font.Color = vbWhite
End If
End Sub

Private Sub chk_m10_Click()
If Me.chk_m10.Value = True Then
    Me.M10.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("M10").Font.Color = &H80000008
Else
    Me.M10.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("M10").Font.Color = vbWhite
End If
End Sub

Private Sub chk_m13_Click()
If Me.chk_m13.Value = True Then
    Me.M13.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("M13").Font.Color = &H80000008
Else
    Me.M13.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("M13").Font.Color = vbWhite
End If
End Sub

Private Sub chk_m16_Click()
If Me.chk_m16.Value = True Then
    Me.M16.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("M16").Font.Color = &H80000008
Else
    Me.M16.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("M16").Font.Color = vbWhite
End If
End Sub

Private Sub chk_m19_Click()
If Me.chk_m19.Value = True Then
    Me.M19.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("M19").Font.Color = &H80000008
Else
    Me.M19.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("M19").Font.Color = vbWhite
End If
End Sub

Private Sub chk_m22_Click()
If Me.chk_m22.Value = True Then
    Me.M22.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("M22").Font.Color = &H80000008
Else
    Me.M22.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("M22").Font.Color = vbWhite
End If
End Sub

Private Sub chk_m25_Click()
If Me.chk_m25.Value = True Then
    Me.M25.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("M25").Font.Color = &H80000008
Else
    Me.M25.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("M25").Font.Color = vbWhite
End If
End Sub

Private Sub chk_m4_Click()
If Me.chk_m4.Value = True Then
    Me.M4.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("M4").Font.Color = &H80000008
Else
    Me.M4.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("M4").Font.Color = vbWhite
End If
End Sub

Private Sub chk_m7_Click()
If Me.chk_m7.Value = True Then
    Me.M7.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("M7").Font.Color = &H80000008
Else
    Me.M7.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("M7").Font.Color = vbWhite
End If
End Sub

Private Sub chk_o1_Click()
If Me.chk_o1.Value = True Then
    Me.O1.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("O1").Font.Color = &H80000008
Else
    Me.O1.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("O1").Font.Color = vbWhite
End If
End Sub

Private Sub chk_o10_Click()
If Me.chk_o10.Value = True Then
    Me.O10.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("O10").Font.Color = &H80000008
Else
    Me.O10.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("O10").Font.Color = vbWhite
End If
End Sub

Private Sub chk_o13_Click()
If Me.chk_o13.Value = True Then
    Me.O13.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("O13").Font.Color = &H80000008
Else
    Me.O13.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("O13").Font.Color = vbWhite
End If
End Sub

Private Sub chk_o16_Click()
If Me.chk_o16.Value = True Then
    Me.O16.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("O16").Font.Color = &H80000008
Else
    Me.O16.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("O16").Font.Color = vbWhite
End If
End Sub

Private Sub chk_o19_Click()
If Me.chk_o19.Value = True Then
    Me.O19.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("O19").Font.Color = &H80000008
Else
    Me.O19.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("O19").Font.Color = vbWhite
End If
End Sub

Private Sub chk_o22_Click()
If Me.chk_o22.Value = True Then
    Me.O22.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("O22").Font.Color = &H80000008
Else
    Me.O22.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("O22").Font.Color = vbWhite
End If
End Sub

Private Sub chk_o25_Click()
If Me.chk_o25.Value = True Then
    Me.O25.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("O25").Font.Color = &H80000008
Else
    Me.O25.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("O25").Font.Color = vbWhite
End If
End Sub

Private Sub chk_o4_Click()
If Me.chk_o4.Value = True Then
    Me.O4.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("O4").Font.Color = &H80000008
Else
    Me.O4.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("O4").Font.Color = vbWhite
End If
End Sub

Private Sub chk_o7_Click()
If Me.chk_o7.Value = True Then
    Me.O7.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("O7").Font.Color = &H80000008
Else
    Me.O7.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("O7").Font.Color = vbWhite
End If
End Sub

Private Sub chk_q1_Click()
If Me.chk_q1.Value = True Then
    Me.Q1.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("Q1").Font.Color = &H80000008
Else
    Me.Q1.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("Q1").Font.Color = vbWhite
End If
End Sub

Private Sub chk_q10_Click()
If Me.chk_q10.Value = True Then
    Me.Q10.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("Q10").Font.Color = &H80000008
Else
    Me.Q10.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("Q10").Font.Color = vbWhite
End If
End Sub

Private Sub chk_q13_Click()
If Me.chk_q13.Value = True Then
    Me.Q13.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("Q13").Font.Color = &H80000008
Else
    Me.Q13.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("Q13").Font.Color = vbWhite
End If
End Sub

Private Sub chk_q16_Click()
If Me.chk_q16.Value = True Then
    Me.Q16.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("Q16").Font.Color = &H80000008
Else
    Me.Q16.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("Q16").Font.Color = vbWhite
End If
End Sub

Private Sub chk_q19_Click()
If Me.chk_q19.Value = True Then
    Me.Q19.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("Q19").Font.Color = &H80000008
Else
    Me.Q19.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("Q19").Font.Color = vbWhite
End If
End Sub

Private Sub chk_q22_Click()
If Me.chk_q22.Value = True Then
    Me.Q22.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("Q22").Font.Color = &H80000008
Else
    Me.Q22.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("Q22").Font.Color = vbWhite
End If
End Sub

Private Sub chk_q25_Click()
If Me.chk_q25.Value = True Then
    Me.Q25.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("Q25").Font.Color = &H80000008
Else
    Me.Q25.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("Q25").Font.Color = vbWhite
End If
End Sub

Private Sub chk_q4_Click()
If Me.chk_q4.Value = True Then
    Me.Q4.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("Q4").Font.Color = &H80000008
Else
    Me.Q4.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("Q4").Font.Color = vbWhite
End If
End Sub

Private Sub chk_q7_Click()
If Me.chk_q7.Value = True Then
    Me.Q7.BackColor = vbWhite
    Sheets("CETAKBARCODE2").Range("Q7").Font.Color = &H80000008
Else
    Me.Q7.BackColor = &H8000000F
    Sheets("CETAKBARCODE2").Range("Q7").Font.Color = vbWhite
End If
End Sub

Private Sub chkAll_Click()
If Me.chkAll.Value = True Then
    Me.chkKolom1.Value = True
    Me.chkKolom2.Value = True
    Me.chkKolom3.Value = True
    Me.chkKolom4.Value = True
    Me.chkKolom5.Value = True
    Me.chkKolom6.Value = True
    Me.chkKolom7.Value = True
    Me.chkKolom8.Value = True
    Me.chkKolom9.Value = True
    
    Me.chkBaris1.Value = True
    Me.chkBaris2.Value = True
    Me.chkBaris3.Value = True
    Me.chkBaris4.Value = True
    Me.chkBaris5.Value = True
    Me.chkBaris6.Value = True
    Me.chkBaris7.Value = True
    Me.chkBaris8.Value = True
    Me.chkBaris9.Value = True
Else
    Me.chkKolom1.Value = False
    Me.chkKolom2.Value = False
    Me.chkKolom3.Value = False
    Me.chkKolom4.Value = False
    Me.chkKolom5.Value = False
    Me.chkKolom6.Value = False
    Me.chkKolom7.Value = False
    Me.chkKolom8.Value = False
    Me.chkKolom9.Value = False
    
    Me.chkBaris1.Value = False
    Me.chkBaris2.Value = False
    Me.chkBaris3.Value = False
    Me.chkBaris4.Value = False
    Me.chkBaris5.Value = False
    Me.chkBaris6.Value = False
    Me.chkBaris7.Value = False
    Me.chkBaris8.Value = False
    Me.chkBaris9.Value = False
End If
End Sub

Private Sub chkBaris1_Click()
If Me.chkBaris1.Value = True Then
    Me.chk_b1.Value = True
    Me.chk_c1.Value = True
    Me.chk_e1.Value = True
    Me.chk_g1.Value = True
    Me.chk_i1.Value = True
    Me.chk_k1.Value = True
    Me.chk_m1.Value = True
    Me.chk_o1.Value = True
    Me.chk_q1.Value = True
    
    Me.B1.BackColor = vbWhite
    Me.C1.BackColor = vbWhite
    Me.E1.BackColor = vbWhite
    Me.G1.BackColor = vbWhite
    Me.I1.BackColor = vbWhite
    Me.K1.BackColor = vbWhite
    Me.M1.BackColor = vbWhite
    Me.O1.BackColor = vbWhite
    Me.Q1.BackColor = vbWhite
    
Else
    Me.chk_b1.Value = False
    Me.chk_c1.Value = False
    Me.chk_e1.Value = False
    Me.chk_g1.Value = False
    Me.chk_i1.Value = False
    Me.chk_k1.Value = False
    Me.chk_m1.Value = False
    Me.chk_o1.Value = False
    Me.chk_q1.Value = False
    
    Me.B1.BackColor = &H8000000F
    Me.C1.BackColor = &H8000000F
    Me.E1.BackColor = &H8000000F
    Me.G1.BackColor = &H8000000F
    Me.I1.BackColor = &H8000000F
    Me.K1.BackColor = &H8000000F
    Me.M1.BackColor = &H8000000F
    Me.O1.BackColor = &H8000000F
    Me.Q1.BackColor = &H8000000F
    
    
End If
End Sub

Private Sub chkBaris2_Click()
If Me.chkBaris2.Value = True Then
    Me.chk_b4.Value = True
    Me.chk_c4.Value = True
    Me.chk_e4.Value = True
    Me.chk_g4.Value = True
    Me.chk_i4.Value = True
    Me.chk_k4.Value = True
    Me.chk_m4.Value = True
    Me.chk_o4.Value = True
    Me.chk_q4.Value = True
    
    Me.B4.BackColor = vbWhite
    Me.C4.BackColor = vbWhite
    Me.E4.BackColor = vbWhite
    Me.G4.BackColor = vbWhite
    Me.I4.BackColor = vbWhite
    Me.K4.BackColor = vbWhite
    Me.M4.BackColor = vbWhite
    Me.O4.BackColor = vbWhite
    Me.Q4.BackColor = vbWhite
    
Else
    Me.chk_b4.Value = False
    Me.chk_c4.Value = False
    Me.chk_e4.Value = False
    Me.chk_g4.Value = False
    Me.chk_i4.Value = False
    Me.chk_k4.Value = False
    Me.chk_m4.Value = False
    Me.chk_o4.Value = False
    Me.chk_q4.Value = False
    
    Me.B4.BackColor = &H8000000F
    Me.C4.BackColor = &H8000000F
    Me.E4.BackColor = &H8000000F
    Me.G4.BackColor = &H8000000F
    Me.I4.BackColor = &H8000000F
    Me.K4.BackColor = &H8000000F
    Me.M4.BackColor = &H8000000F
    Me.O4.BackColor = &H8000000F
    Me.Q4.BackColor = &H8000000F
    
    
End If
End Sub

Private Sub chkBaris3_Click()
If Me.chkBaris3.Value = True Then
    Me.chk_b7.Value = True
    Me.chk_c7.Value = True
    Me.chk_e7.Value = True
    Me.chk_g7.Value = True
    Me.chk_i7.Value = True
    Me.chk_k7.Value = True
    Me.chk_m7.Value = True
    Me.chk_o7.Value = True
    Me.chk_q7.Value = True
    
    Me.B7.BackColor = vbWhite
    Me.C7.BackColor = vbWhite
    Me.E7.BackColor = vbWhite
    Me.G7.BackColor = vbWhite
    Me.I7.BackColor = vbWhite
    Me.K7.BackColor = vbWhite
    Me.M7.BackColor = vbWhite
    Me.O7.BackColor = vbWhite
    Me.Q7.BackColor = vbWhite
    
Else
    Me.chk_b7.Value = False
    Me.chk_c7.Value = False
    Me.chk_e7.Value = False
    Me.chk_g7.Value = False
    Me.chk_i7.Value = False
    Me.chk_k7.Value = False
    Me.chk_m7.Value = False
    Me.chk_o7.Value = False
    Me.chk_q7.Value = False
    
    Me.B7.BackColor = &H8000000F
    Me.C7.BackColor = &H8000000F
    Me.E7.BackColor = &H8000000F
    Me.G7.BackColor = &H8000000F
    Me.I7.BackColor = &H8000000F
    Me.K7.BackColor = &H8000000F
    Me.M7.BackColor = &H8000000F
    Me.O7.BackColor = &H8000000F
    Me.Q7.BackColor = &H8000000F
    
    
End If
End Sub

Private Sub chkBaris4_Click()
If Me.chkBaris4.Value = True Then
    Me.chk_b10.Value = True
    Me.chk_c10.Value = True
    Me.chk_e10.Value = True
    Me.chk_g10.Value = True
    Me.chk_i10.Value = True
    Me.chk_k10.Value = True
    Me.chk_m10.Value = True
    Me.chk_o10.Value = True
    Me.chk_q10.Value = True
    
    Me.B10.BackColor = vbWhite
    Me.C10.BackColor = vbWhite
    Me.E10.BackColor = vbWhite
    Me.G10.BackColor = vbWhite
    Me.I10.BackColor = vbWhite
    Me.K10.BackColor = vbWhite
    Me.M10.BackColor = vbWhite
    Me.O10.BackColor = vbWhite
    Me.Q10.BackColor = vbWhite
    
Else
    Me.chk_b10.Value = False
    Me.chk_c10.Value = False
    Me.chk_e10.Value = False
    Me.chk_g10.Value = False
    Me.chk_i10.Value = False
    Me.chk_k10.Value = False
    Me.chk_m10.Value = False
    Me.chk_o10.Value = False
    Me.chk_q10.Value = False
    
    Me.B10.BackColor = &H8000000F
    Me.C10.BackColor = &H8000000F
    Me.E10.BackColor = &H8000000F
    Me.G10.BackColor = &H8000000F
    Me.I10.BackColor = &H8000000F
    Me.K10.BackColor = &H8000000F
    Me.M10.BackColor = &H8000000F
    Me.O10.BackColor = &H8000000F
    Me.Q10.BackColor = &H8000000F
    
    
End If
End Sub

Private Sub chkBaris5_Click()
If Me.chkBaris5.Value = True Then
    Me.chk_b13.Value = True
    Me.chk_c13.Value = True
    Me.chk_e13.Value = True
    Me.chk_g13.Value = True
    Me.chk_i13.Value = True
    Me.chk_k13.Value = True
    Me.chk_m13.Value = True
    Me.chk_o13.Value = True
    Me.chk_q13.Value = True
    
    Me.B13.BackColor = vbWhite
    Me.C13.BackColor = vbWhite
    Me.E13.BackColor = vbWhite
    Me.G13.BackColor = vbWhite
    Me.I13.BackColor = vbWhite
    Me.K13.BackColor = vbWhite
    Me.M13.BackColor = vbWhite
    Me.O13.BackColor = vbWhite
    Me.Q13.BackColor = vbWhite
    
Else
    Me.chk_b13.Value = False
    Me.chk_c13.Value = False
    Me.chk_e13.Value = False
    Me.chk_g13.Value = False
    Me.chk_i13.Value = False
    Me.chk_k13.Value = False
    Me.chk_m13.Value = False
    Me.chk_o13.Value = False
    Me.chk_q13.Value = False
    
    Me.B13.BackColor = &H8000000F
    Me.C13.BackColor = &H8000000F
    Me.E13.BackColor = &H8000000F
    Me.G13.BackColor = &H8000000F
    Me.I13.BackColor = &H8000000F
    Me.K13.BackColor = &H8000000F
    Me.M13.BackColor = &H8000000F
    Me.O13.BackColor = &H8000000F
    Me.Q13.BackColor = &H8000000F
    
    
End If
End Sub

Private Sub chkBaris6_Click()
If Me.chkBaris6.Value = True Then
    Me.chk_b16.Value = True
    Me.chk_c16.Value = True
    Me.chk_e16.Value = True
    Me.chk_g16.Value = True
    Me.chk_i16.Value = True
    Me.chk_k16.Value = True
    Me.chk_m16.Value = True
    Me.chk_o16.Value = True
    Me.chk_q16.Value = True
    
    Me.B16.BackColor = vbWhite
    Me.C16.BackColor = vbWhite
    Me.E16.BackColor = vbWhite
    Me.G16.BackColor = vbWhite
    Me.I16.BackColor = vbWhite
    Me.K16.BackColor = vbWhite
    Me.M16.BackColor = vbWhite
    Me.O16.BackColor = vbWhite
    Me.Q16.BackColor = vbWhite
    
Else
    Me.chk_b16.Value = False
    Me.chk_c16.Value = False
    Me.chk_e16.Value = False
    Me.chk_g16.Value = False
    Me.chk_i16.Value = False
    Me.chk_k16.Value = False
    Me.chk_m16.Value = False
    Me.chk_o16.Value = False
    Me.chk_q16.Value = False
    
    Me.B16.BackColor = &H8000000F
    Me.C16.BackColor = &H8000000F
    Me.E16.BackColor = &H8000000F
    Me.G16.BackColor = &H8000000F
    Me.I16.BackColor = &H8000000F
    Me.K16.BackColor = &H8000000F
    Me.M16.BackColor = &H8000000F
    Me.O16.BackColor = &H8000000F
    Me.Q16.BackColor = &H8000000F
    
    
End If
End Sub

Private Sub chkBaris7_Click()
If Me.chkBaris7.Value = True Then
    Me.chk_b19.Value = True
    Me.chk_c19.Value = True
    Me.chk_e19.Value = True
    Me.chk_g19.Value = True
    Me.chk_i19.Value = True
    Me.chk_k19.Value = True
    Me.chk_m19.Value = True
    Me.chk_o19.Value = True
    Me.chk_q19.Value = True
    
    Me.B19.BackColor = vbWhite
    Me.C19.BackColor = vbWhite
    Me.E19.BackColor = vbWhite
    Me.G19.BackColor = vbWhite
    Me.I19.BackColor = vbWhite
    Me.K19.BackColor = vbWhite
    Me.M19.BackColor = vbWhite
    Me.O19.BackColor = vbWhite
    Me.Q19.BackColor = vbWhite
    
Else
    Me.chk_b19.Value = False
    Me.chk_c19.Value = False
    Me.chk_e19.Value = False
    Me.chk_g19.Value = False
    Me.chk_i19.Value = False
    Me.chk_k19.Value = False
    Me.chk_m19.Value = False
    Me.chk_o19.Value = False
    Me.chk_q19.Value = False
    
    Me.B19.BackColor = &H8000000F
    Me.C19.BackColor = &H8000000F
    Me.E19.BackColor = &H8000000F
    Me.G19.BackColor = &H8000000F
    Me.I19.BackColor = &H8000000F
    Me.K19.BackColor = &H8000000F
    Me.M19.BackColor = &H8000000F
    Me.O19.BackColor = &H8000000F
    Me.Q19.BackColor = &H8000000F
    
    
End If
End Sub

Private Sub chkBaris8_Click()
If Me.chkBaris8.Value = True Then
    Me.chk_b22.Value = True
    Me.chk_c22.Value = True
    Me.chk_e22.Value = True
    Me.chk_g22.Value = True
    Me.chk_i22.Value = True
    Me.chk_k22.Value = True
    Me.chk_m22.Value = True
    Me.chk_o22.Value = True
    Me.chk_q22.Value = True
    
    Me.B22.BackColor = vbWhite
    Me.C22.BackColor = vbWhite
    Me.E22.BackColor = vbWhite
    Me.G22.BackColor = vbWhite
    Me.I22.BackColor = vbWhite
    Me.K22.BackColor = vbWhite
    Me.M22.BackColor = vbWhite
    Me.O22.BackColor = vbWhite
    Me.Q22.BackColor = vbWhite
    
Else
    Me.chk_b22.Value = False
    Me.chk_c22.Value = False
    Me.chk_e22.Value = False
    Me.chk_g22.Value = False
    Me.chk_i22.Value = False
    Me.chk_k22.Value = False
    Me.chk_m22.Value = False
    Me.chk_o22.Value = False
    Me.chk_q22.Value = False
    
    Me.B22.BackColor = &H8000000F
    Me.C22.BackColor = &H8000000F
    Me.E22.BackColor = &H8000000F
    Me.G22.BackColor = &H8000000F
    Me.I22.BackColor = &H8000000F
    Me.K22.BackColor = &H8000000F
    Me.M22.BackColor = &H8000000F
    Me.O22.BackColor = &H8000000F
    Me.Q22.BackColor = &H8000000F
    
    
End If
End Sub

Private Sub chkBaris9_Click()
If Me.chkBaris9.Value = True Then
    Me.chk_b25.Value = True
    Me.chk_c25.Value = True
    Me.chk_e25.Value = True
    Me.chk_g25.Value = True
    Me.chk_i25.Value = True
    Me.chk_k25.Value = True
    Me.chk_m25.Value = True
    Me.chk_o25.Value = True
    Me.chk_q25.Value = True
    
    Me.B25.BackColor = vbWhite
    Me.C25.BackColor = vbWhite
    Me.E25.BackColor = vbWhite
    Me.G25.BackColor = vbWhite
    Me.I25.BackColor = vbWhite
    Me.K25.BackColor = vbWhite
    Me.M25.BackColor = vbWhite
    Me.O25.BackColor = vbWhite
    Me.Q25.BackColor = vbWhite
    
Else
    Me.chk_b25.Value = False
    Me.chk_c25.Value = False
    Me.chk_e25.Value = False
    Me.chk_g25.Value = False
    Me.chk_i25.Value = False
    Me.chk_k25.Value = False
    Me.chk_m25.Value = False
    Me.chk_o25.Value = False
    Me.chk_q25.Value = False
    
    Me.B25.BackColor = &H8000000F
    Me.C25.BackColor = &H8000000F
    Me.E25.BackColor = &H8000000F
    Me.G25.BackColor = &H8000000F
    Me.I25.BackColor = &H8000000F
    Me.K25.BackColor = &H8000000F
    Me.M25.BackColor = &H8000000F
    Me.O25.BackColor = &H8000000F
    Me.Q25.BackColor = &H8000000F
    
    
End If
End Sub

Private Sub chkKolom1_Click()
If Me.chkKolom1.Value = True Then
    Me.chk_b1.Value = True
    Me.chk_b4.Value = True
    Me.chk_b7.Value = True
    Me.chk_b10.Value = True
    Me.chk_b13.Value = True
    Me.chk_b16.Value = True
    Me.chk_b19.Value = True
    Me.chk_b22.Value = True
    Me.chk_b25.Value = True
    
    Me.B1.BackColor = vbWhite
    Me.B4.BackColor = vbWhite
    Me.B7.BackColor = vbWhite
    Me.B10.BackColor = vbWhite
    Me.B13.BackColor = vbWhite
    Me.B16.BackColor = vbWhite
    Me.B19.BackColor = vbWhite
    Me.B22.BackColor = vbWhite
    Me.B25.BackColor = vbWhite
    
Else
    Me.chk_b1.Value = False
    Me.chk_b4.Value = False
    Me.chk_b7.Value = False
    Me.chk_b10.Value = False
    Me.chk_b13.Value = False
    Me.chk_b16.Value = False
    Me.chk_b19.Value = False
    Me.chk_b22.Value = False
    Me.chk_b25.Value = False
    
    Me.B1.BackColor = &H8000000F
    Me.B4.BackColor = &H8000000F
    Me.B7.BackColor = &H8000000F
    Me.B10.BackColor = &H8000000F
    Me.B13.BackColor = &H8000000F
    Me.B16.BackColor = &H8000000F
    Me.B19.BackColor = &H8000000F
    Me.B22.BackColor = &H8000000F
    Me.B25.BackColor = &H8000000F
    
    
End If
End Sub

Private Sub Label8_Click()

End Sub

Private Sub chkKolom2_Click()
If Me.chkKolom2.Value = True Then
    Me.chk_c1.Value = True
    Me.chk_c4.Value = True
    Me.chk_c7.Value = True
    Me.chk_c10.Value = True
    Me.chk_c13.Value = True
    Me.chk_c16.Value = True
    Me.chk_c19.Value = True
    Me.chk_c22.Value = True
    Me.chk_c25.Value = True
    
    Me.C1.BackColor = vbWhite
    Me.C4.BackColor = vbWhite
    Me.C7.BackColor = vbWhite
    Me.C10.BackColor = vbWhite
    Me.C13.BackColor = vbWhite
    Me.C16.BackColor = vbWhite
    Me.C19.BackColor = vbWhite
    Me.C22.BackColor = vbWhite
    Me.C25.BackColor = vbWhite
    
Else
    Me.chk_c1.Value = False
    Me.chk_c4.Value = False
    Me.chk_c7.Value = False
    Me.chk_c10.Value = False
    Me.chk_c13.Value = False
    Me.chk_c16.Value = False
    Me.chk_c19.Value = False
    Me.chk_c22.Value = False
    Me.chk_c25.Value = False
    
    Me.C1.BackColor = &H8000000F
    Me.C4.BackColor = &H8000000F
    Me.C7.BackColor = &H8000000F
    Me.C10.BackColor = &H8000000F
    Me.C13.BackColor = &H8000000F
    Me.C16.BackColor = &H8000000F
    Me.C19.BackColor = &H8000000F
    Me.C22.BackColor = &H8000000F
    Me.C25.BackColor = &H8000000F
    
    
End If
End Sub

Private Sub chkKolom3_Click()
If Me.chkKolom3.Value = True Then
    Me.chk_e1.Value = True
    Me.chk_e4.Value = True
    Me.chk_e7.Value = True
    Me.chk_e10.Value = True
    Me.chk_e13.Value = True
    Me.chk_e16.Value = True
    Me.chk_e19.Value = True
    Me.chk_e22.Value = True
    Me.chk_e25.Value = True
    
    Me.E1.BackColor = vbWhite
    Me.E4.BackColor = vbWhite
    Me.E7.BackColor = vbWhite
    Me.E10.BackColor = vbWhite
    Me.E13.BackColor = vbWhite
    Me.E16.BackColor = vbWhite
    Me.E19.BackColor = vbWhite
    Me.E22.BackColor = vbWhite
    Me.E25.BackColor = vbWhite
    
Else
    Me.chk_e1.Value = False
    Me.chk_e4.Value = False
    Me.chk_e7.Value = False
    Me.chk_e10.Value = False
    Me.chk_e13.Value = False
    Me.chk_e16.Value = False
    Me.chk_e19.Value = False
    Me.chk_e22.Value = False
    Me.chk_e25.Value = False
    
    Me.E1.BackColor = &H8000000F
    Me.E4.BackColor = &H8000000F
    Me.E7.BackColor = &H8000000F
    Me.E10.BackColor = &H8000000F
    Me.E13.BackColor = &H8000000F
    Me.E16.BackColor = &H8000000F
    Me.E19.BackColor = &H8000000F
    Me.E22.BackColor = &H8000000F
    Me.E25.BackColor = &H8000000F
    
    
End If
End Sub

Private Sub chkKolom4_Click()
If Me.chkKolom4.Value = True Then
    Me.chk_g1.Value = True
    Me.chk_g4.Value = True
    Me.chk_g7.Value = True
    Me.chk_g10.Value = True
    Me.chk_g13.Value = True
    Me.chk_g16.Value = True
    Me.chk_g19.Value = True
    Me.chk_g22.Value = True
    Me.chk_g25.Value = True
    
    Me.G1.BackColor = vbWhite
    Me.G4.BackColor = vbWhite
    Me.G7.BackColor = vbWhite
    Me.G10.BackColor = vbWhite
    Me.G13.BackColor = vbWhite
    Me.G16.BackColor = vbWhite
    Me.G19.BackColor = vbWhite
    Me.G22.BackColor = vbWhite
    Me.G25.BackColor = vbWhite
    
Else
    Me.chk_g1.Value = False
    Me.chk_g4.Value = False
    Me.chk_g7.Value = False
    Me.chk_g10.Value = False
    Me.chk_g13.Value = False
    Me.chk_g16.Value = False
    Me.chk_g19.Value = False
    Me.chk_g22.Value = False
    Me.chk_g25.Value = False
    
    Me.G1.BackColor = &H8000000F
    Me.G4.BackColor = &H8000000F
    Me.G7.BackColor = &H8000000F
    Me.G10.BackColor = &H8000000F
    Me.G13.BackColor = &H8000000F
    Me.G16.BackColor = &H8000000F
    Me.G19.BackColor = &H8000000F
    Me.G22.BackColor = &H8000000F
    Me.G25.BackColor = &H8000000F
End If
End Sub

Private Sub chkKolom5_Click()
If Me.chkKolom5.Value = True Then
    Me.chk_i1.Value = True
    Me.chk_i4.Value = True
    Me.chk_i7.Value = True
    Me.chk_i10.Value = True
    Me.chk_i13.Value = True
    Me.chk_i16.Value = True
    Me.chk_i19.Value = True
    Me.chk_i22.Value = True
    Me.chk_i25.Value = True
    
    Me.I1.BackColor = vbWhite
    Me.I4.BackColor = vbWhite
    Me.I7.BackColor = vbWhite
    Me.I10.BackColor = vbWhite
    Me.I13.BackColor = vbWhite
    Me.I16.BackColor = vbWhite
    Me.I19.BackColor = vbWhite
    Me.I22.BackColor = vbWhite
    Me.I25.BackColor = vbWhite
    
Else
    Me.chk_i1.Value = False
    Me.chk_i4.Value = False
    Me.chk_i7.Value = False
    Me.chk_i10.Value = False
    Me.chk_i13.Value = False
    Me.chk_i16.Value = False
    Me.chk_i19.Value = False
    Me.chk_i22.Value = False
    Me.chk_i25.Value = False
    
    Me.I1.BackColor = &H8000000F
    Me.I4.BackColor = &H8000000F
    Me.I7.BackColor = &H8000000F
    Me.I10.BackColor = &H8000000F
    Me.I13.BackColor = &H8000000F
    Me.I16.BackColor = &H8000000F
    Me.I19.BackColor = &H8000000F
    Me.I22.BackColor = &H8000000F
    Me.I25.BackColor = &H8000000F
End If
End Sub

Private Sub chkKolom6_Click()
If Me.chkKolom6.Value = True Then
    Me.chk_k1.Value = True
    Me.chk_k4.Value = True
    Me.chk_k7.Value = True
    Me.chk_k10.Value = True
    Me.chk_k13.Value = True
    Me.chk_k16.Value = True
    Me.chk_k19.Value = True
    Me.chk_k22.Value = True
    Me.chk_k25.Value = True
    
    Me.K1.BackColor = vbWhite
    Me.K4.BackColor = vbWhite
    Me.K7.BackColor = vbWhite
    Me.K10.BackColor = vbWhite
    Me.K13.BackColor = vbWhite
    Me.K16.BackColor = vbWhite
    Me.K19.BackColor = vbWhite
    Me.K22.BackColor = vbWhite
    Me.K25.BackColor = vbWhite
    
Else
    Me.chk_k1.Value = False
    Me.chk_k4.Value = False
    Me.chk_k7.Value = False
    Me.chk_k10.Value = False
    Me.chk_k13.Value = False
    Me.chk_k16.Value = False
    Me.chk_k19.Value = False
    Me.chk_k22.Value = False
    Me.chk_k25.Value = False
    
    Me.K1.BackColor = &H8000000F
    Me.K4.BackColor = &H8000000F
    Me.K7.BackColor = &H8000000F
    Me.K10.BackColor = &H8000000F
    Me.K13.BackColor = &H8000000F
    Me.K16.BackColor = &H8000000F
    Me.K19.BackColor = &H8000000F
    Me.K22.BackColor = &H8000000F
    Me.K25.BackColor = &H8000000F
End If
End Sub

Private Sub chkKolom7_Click()
If Me.chkKolom7.Value = True Then
    Me.chk_m1.Value = True
    Me.chk_m4.Value = True
    Me.chk_m7.Value = True
    Me.chk_m10.Value = True
    Me.chk_m13.Value = True
    Me.chk_m16.Value = True
    Me.chk_m19.Value = True
    Me.chk_m22.Value = True
    Me.chk_m25.Value = True
    
    Me.M1.BackColor = vbWhite
    Me.M4.BackColor = vbWhite
    Me.M7.BackColor = vbWhite
    Me.M10.BackColor = vbWhite
    Me.M13.BackColor = vbWhite
    Me.M16.BackColor = vbWhite
    Me.M19.BackColor = vbWhite
    Me.M22.BackColor = vbWhite
    Me.M25.BackColor = vbWhite
    
Else
    Me.chk_m1.Value = False
    Me.chk_m4.Value = False
    Me.chk_m7.Value = False
    Me.chk_m10.Value = False
    Me.chk_m13.Value = False
    Me.chk_m16.Value = False
    Me.chk_m19.Value = False
    Me.chk_m22.Value = False
    Me.chk_m25.Value = False
    
    Me.M1.BackColor = &H8000000F
    Me.M4.BackColor = &H8000000F
    Me.M7.BackColor = &H8000000F
    Me.M10.BackColor = &H8000000F
    Me.M13.BackColor = &H8000000F
    Me.M16.BackColor = &H8000000F
    Me.M19.BackColor = &H8000000F
    Me.M22.BackColor = &H8000000F
    Me.M25.BackColor = &H8000000F
End If
End Sub

Private Sub chkKolom8_Click()
If Me.chkKolom8.Value = True Then
    Me.chk_o1.Value = True
    Me.chk_o4.Value = True
    Me.chk_o7.Value = True
    Me.chk_o10.Value = True
    Me.chk_o13.Value = True
    Me.chk_o16.Value = True
    Me.chk_o19.Value = True
    Me.chk_o22.Value = True
    Me.chk_o25.Value = True
    
    Me.O1.BackColor = vbWhite
    Me.O4.BackColor = vbWhite
    Me.O7.BackColor = vbWhite
    Me.O10.BackColor = vbWhite
    Me.O13.BackColor = vbWhite
    Me.O16.BackColor = vbWhite
    Me.O19.BackColor = vbWhite
    Me.O22.BackColor = vbWhite
    Me.O25.BackColor = vbWhite
    
Else
    Me.chk_o1.Value = False
    Me.chk_o4.Value = False
    Me.chk_o7.Value = False
    Me.chk_o10.Value = False
    Me.chk_o13.Value = False
    Me.chk_o16.Value = False
    Me.chk_o19.Value = False
    Me.chk_o22.Value = False
    Me.chk_o25.Value = False
    
    Me.O1.BackColor = &H8000000F
    Me.O4.BackColor = &H8000000F
    Me.O7.BackColor = &H8000000F
    Me.O10.BackColor = &H8000000F
    Me.O13.BackColor = &H8000000F
    Me.O16.BackColor = &H8000000F
    Me.O19.BackColor = &H8000000F
    Me.O22.BackColor = &H8000000F
    Me.O25.BackColor = &H8000000F
End If
End Sub

Private Sub chkKolom9_Click()
If Me.chkKolom9.Value = True Then
    Me.chk_q1.Value = True
    Me.chk_q4.Value = True
    Me.chk_q7.Value = True
    Me.chk_q10.Value = True
    Me.chk_q13.Value = True
    Me.chk_q16.Value = True
    Me.chk_q19.Value = True
    Me.chk_q22.Value = True
    Me.chk_q25.Value = True
    
    Me.Q1.BackColor = vbWhite
    Me.Q4.BackColor = vbWhite
    Me.Q7.BackColor = vbWhite
    Me.Q10.BackColor = vbWhite
    Me.Q13.BackColor = vbWhite
    Me.Q16.BackColor = vbWhite
    Me.Q19.BackColor = vbWhite
    Me.Q22.BackColor = vbWhite
    Me.Q25.BackColor = vbWhite
    
Else
    Me.chk_q1.Value = False
    Me.chk_q4.Value = False
    Me.chk_q7.Value = False
    Me.chk_q10.Value = False
    Me.chk_q13.Value = False
    Me.chk_q16.Value = False
    Me.chk_q19.Value = False
    Me.chk_q22.Value = False
    Me.chk_q25.Value = False
    
    Me.Q1.BackColor = &H8000000F
    Me.Q4.BackColor = &H8000000F
    Me.Q7.BackColor = &H8000000F
    Me.Q10.BackColor = &H8000000F
    Me.Q13.BackColor = &H8000000F
    Me.Q16.BackColor = &H8000000F
    Me.Q19.BackColor = &H8000000F
    Me.Q22.BackColor = &H8000000F
    Me.Q25.BackColor = &H8000000F
End If
End Sub

Private Sub E1_Click()
If Me.chk_e1.Value = True Then
    Me.chk_e1.Value = False
Else
    Me.chk_e1.Value = True
End If
End Sub

Private Sub E10_Click()
If Me.chk_e10.Value = True Then
    Me.chk_e10.Value = False
Else
    Me.chk_e10.Value = True
End If
End Sub

Private Sub E13_Click()
If Me.chk_e13.Value = True Then
    Me.chk_e13.Value = False
Else
    Me.chk_e13.Value = True
End If
End Sub

Private Sub E16_Click()
If Me.chk_e16.Value = True Then
    Me.chk_e16.Value = False
Else
    Me.chk_e16.Value = True
End If
End Sub

Private Sub E19_Click()
If Me.chk_e19.Value = True Then
    Me.chk_e19.Value = False
Else
    Me.chk_e19.Value = True
End If
End Sub

Private Sub E22_Click()
If Me.chk_e22.Value = True Then
    Me.chk_e22.Value = False
Else
    Me.chk_e22.Value = True
End If
End Sub

Private Sub E25_Click()
If Me.chk_e25.Value = True Then
    Me.chk_e25.Value = False
Else
    Me.chk_e25.Value = True
End If
End Sub

Private Sub E4_Click()
If Me.chk_e4.Value = True Then
    Me.chk_e4.Value = False
Else
    Me.chk_e4.Value = True
End If
End Sub

Private Sub E7_Click()
If Me.chk_e7.Value = True Then
    Me.chk_e7.Value = False
Else
    Me.chk_e7.Value = True
End If
End Sub

Private Sub G1_Click()
If Me.chk_g1.Value = True Then
    Me.chk_g1.Value = False
Else
    Me.chk_g1.Value = True
End If
End Sub

Private Sub G10_Click()
If Me.chk_g10.Value = True Then
    Me.chk_g10.Value = False
Else
    Me.chk_g10.Value = True
End If
End Sub

Private Sub G13_Click()
If Me.chk_g13.Value = True Then
    Me.chk_g13.Value = False
Else
    Me.chk_g13.Value = True
End If
End Sub

Private Sub G16_Click()
If Me.chk_g16.Value = True Then
    Me.chk_g16.Value = False
Else
    Me.chk_g16.Value = True
End If
End Sub

Private Sub G19_Click()
If Me.chk_g19.Value = True Then
    Me.chk_g19.Value = False
Else
    Me.chk_g19.Value = True
End If
End Sub

Private Sub G22_Click()
If Me.chk_g22.Value = True Then
    Me.chk_g22.Value = False
Else
    Me.chk_g22.Value = True
End If
End Sub

Private Sub G25_Click()
If Me.chk_g25.Value = True Then
    Me.chk_g25.Value = False
Else
    Me.chk_g25.Value = True
End If
End Sub

Private Sub G4_Click()
If Me.chk_g4.Value = True Then
    Me.chk_g4.Value = False
Else
    Me.chk_g4.Value = True
End If
End Sub

Private Sub G7_Click()
If Me.chk_g7.Value = True Then
    Me.chk_g7.Value = False
Else
    Me.chk_g7.Value = True
End If
End Sub

Private Sub I1_Click()
If Me.chk_i1.Value = True Then
    Me.chk_i1.Value = False
Else
    Me.chk_i1.Value = True
End If
End Sub

Private Sub I10_Click()
If Me.chk_i10.Value = True Then
    Me.chk_i10.Value = False
Else
    Me.chk_i10.Value = True
End If
End Sub

Private Sub I13_Click()
If Me.chk_i13.Value = True Then
    Me.chk_i13.Value = False
Else
    Me.chk_i13.Value = True
End If
End Sub

Private Sub I16_Click()
If Me.chk_i16.Value = True Then
    Me.chk_i16.Value = False
Else
    Me.chk_i16.Value = True
End If
End Sub

Private Sub I19_Click()
If Me.chk_i19.Value = True Then
    Me.chk_i19.Value = False
Else
    Me.chk_i19.Value = True
End If
End Sub

Private Sub I22_Click()
If Me.chk_i22.Value = True Then
    Me.chk_i22.Value = False
Else
    Me.chk_i22.Value = True
End If
End Sub

Private Sub I25_Click()
If Me.chk_i25.Value = True Then
    Me.chk_i25.Value = False
Else
    Me.chk_i25.Value = True
End If
End Sub

Private Sub I4_Click()
If Me.chk_i4.Value = True Then
    Me.chk_i4.Value = False
Else
    Me.chk_i4.Value = True
End If
End Sub

Private Sub I7_Click()
If Me.chk_i17.Value = True Then
    Me.chk_i17.Value = False
Else
    Me.chk_i17.Value = True
End If
End Sub

Private Sub K1_Click()
If Me.chk_k1.Value = True Then
    Me.chk_k1.Value = False
Else
    Me.chk_k1.Value = True
End If
End Sub

Private Sub K10_Click()
If Me.chk_k10.Value = True Then
    Me.chk_k10.Value = False
Else
    Me.chk_k10.Value = True
End If
End Sub

Private Sub K13_Click()
If Me.chk_k13.Value = True Then
    Me.chk_k13.Value = False
Else
    Me.chk_k13.Value = True
End If
End Sub

Private Sub K16_Click()
If Me.chk_k16.Value = True Then
    Me.chk_k16.Value = False
Else
    Me.chk_k16.Value = True
End If
End Sub

Private Sub K19_Click()
If Me.chk_k19.Value = True Then
    Me.chk_k19.Value = False
Else
    Me.chk_k19.Value = True
End If
End Sub

Private Sub K22_Click()
If Me.chk_k22.Value = True Then
    Me.chk_k22.Value = False
Else
    Me.chk_k22.Value = True
End If
End Sub

Private Sub K25_Click()
If Me.chk_k25.Value = True Then
    Me.chk_k25.Value = False
Else
    Me.chk_k25.Value = True
End If
End Sub

Private Sub K4_Click()
If Me.chk_k4.Value = True Then
    Me.chk_k4.Value = False
Else
    Me.chk_k4.Value = True
End If
End Sub

Private Sub K7_Click()
If Me.chk_k7.Value = True Then
    Me.chk_k7.Value = False
Else
    Me.chk_k7.Value = True
End If
End Sub

Private Sub M1_Click()
If Me.chk_m1.Value = True Then
    Me.chk_m1.Value = False
Else
    Me.chk_m1.Value = True
End If
End Sub

Private Sub M10_Click()
If Me.chk_m10.Value = True Then
    Me.chk_m10.Value = False
Else
    Me.chk_m10.Value = True
End If
End Sub

Private Sub M13_Click()
If Me.chk_m13.Value = True Then
    Me.chk_m13.Value = False
Else
    Me.chk_m13.Value = True
End If
End Sub

Private Sub M16_Click()
If Me.chk_m16.Value = True Then
    Me.chk_m16.Value = False
Else
    Me.chk_m16.Value = True
End If
End Sub

Private Sub M19_Click()
If Me.chk_m19.Value = True Then
    Me.chk_m19.Value = False
Else
    Me.chk_m19.Value = True
End If
End Sub

Private Sub M22_Click()
If Me.chk_m22.Value = True Then
    Me.chk_m22.Value = False
Else
    Me.chk_m22.Value = True
End If
End Sub

Private Sub M25_Click()
If Me.chk_m25.Value = True Then
    Me.chk_m25.Value = False
Else
    Me.chk_m25.Value = True
End If
End Sub

Private Sub M4_Click()
If Me.chk_m4.Value = True Then
    Me.chk_m4.Value = False
Else
    Me.chk_m4.Value = True
End If
End Sub

Private Sub M7_Click()
If Me.chk_m7.Value = True Then
    Me.chk_m7.Value = False
Else
    Me.chk_m7.Value = True
End If
End Sub

Private Sub O1_Click()
If Me.chk_o1.Value = True Then
    Me.chk_o1.Value = False
Else
    Me.chk_o1.Value = True
End If
End Sub

Private Sub O10_Click()
If Me.chk_o10.Value = True Then
    Me.chk_o10.Value = False
Else
    Me.chk_o10.Value = True
End If
End Sub

Private Sub O13_Click()
If Me.chk_o13.Value = True Then
    Me.chk_o13.Value = False
Else
    Me.chk_o13.Value = True
End If
End Sub

Private Sub O16_Click()
If Me.chk_o16.Value = True Then
    Me.chk_o16.Value = False
Else
    Me.chk_o16.Value = True
End If
End Sub

Private Sub O19_Click()
If Me.chk_o19.Value = True Then
    Me.chk_o19.Value = False
Else
    Me.chk_o19.Value = True
End If
End Sub

Private Sub O22_Click()
If Me.chk_o22.Value = True Then
    Me.chk_o22.Value = False
Else
    Me.chk_o22.Value = True
End If
End Sub

Private Sub O25_Click()
If Me.chk_o25.Value = True Then
    Me.chk_o25.Value = False
Else
    Me.chk_o25.Value = True
End If
End Sub

Private Sub O4_Click()
If Me.chk_o4.Value = True Then
    Me.chk_o4.Value = False
Else
    Me.chk_o4.Value = True
End If
End Sub

Private Sub O7_Click()
If Me.chk_o7.Value = True Then
    Me.chk_o7.Value = False
Else
    Me.chk_o7.Value = True
End If
End Sub

Private Sub Q1_Click()
If Me.chk_q1.Value = True Then
    Me.chk_q1.Value = False
Else
    Me.chk_q1.Value = True
End If
End Sub

Private Sub Q10_Click()
If Me.chk_q10.Value = True Then
    Me.chk_q10.Value = False
Else
    Me.chk_q10.Value = True
End If
End Sub

Private Sub Q13_Click()
If Me.chk_q13.Value = True Then
    Me.chk_q13.Value = False
Else
    Me.chk_q13.Value = True
End If
End Sub

Private Sub Q16_Click()
If Me.chk_q16.Value = True Then
    Me.chk_q16.Value = False
Else
    Me.chk_q16.Value = True
End If
End Sub

Private Sub Q19_Click()
If Me.chk_q19.Value = True Then
    Me.chk_q19.Value = False
Else
    Me.chk_q19.Value = True
End If
End Sub

Private Sub Q22_Click()
If Me.chk_q22.Value = True Then
    Me.chk_q22.Value = False
Else
    Me.chk_q22.Value = True
End If
End Sub

Private Sub Q25_Click()
If Me.chk_q25.Value = True Then
    Me.chk_q25.Value = False
Else
    Me.chk_q25.Value = True
End If
End Sub

Private Sub Q4_Click()
If Me.chk_q4.Value = True Then
    Me.chk_q4.Value = False
Else
    Me.chk_q4.Value = True
End If
End Sub

Private Sub Q7_Click()
If Me.chk_q7.Value = True Then
    Me.chk_q7.Value = False
Else
    Me.chk_q7.Value = True
End If
End Sub
