VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CETAKBARCODE 
   Caption         =   "CETAK BARCODE"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5550
   OleObjectBlob   =   "frmCetakBarcode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CETAKBARCODE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
End Sub
