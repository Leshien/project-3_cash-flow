
Attribute VB_Name = "CashFlowSimulator"
Sub ApplyRandomShock()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Simulator")

    ' Apply random ±5 day shift
    Randomize
    ws.Range("D6").Value = 30 + Int((60 - 30 + 1) * Rnd) + Int((Rnd() - 0.5) * 10)
    ws.Range("D7").Value = 20 + Int((45 - 20 + 1) * Rnd) + Int((Rnd() - 0.5) * 10)
    ws.Range("D8").Value = 30 + Int((70 - 30 + 1) * Rnd) + Int((Rnd() - 0.5) * 10)

    MsgBox "Random shock applied!", vbInformation
End Sub

Sub ResetToBaseline()
    Dim wsData As Worksheet, wsSim As Worksheet
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsSim = ThisWorkbook.Sheets("Simulator")

    ' Use first row as reference
    wsSim.Range("D6").Value = wsData.Range("D2").Value ' Receivables Days
    wsSim.Range("D7").Value = wsData.Range("E2").Value ' Inventory Days
    wsSim.Range("D8").Value = wsData.Range("F2").Value ' Payables Days

    MsgBox "Values reset to baseline!", vbInformation
End Sub

Sub Run100Simulations()
    Dim wsLog As Worksheet
    Dim i As Integer
    Dim dso As Integer, dio As Integer, dpo As Integer
    Dim fcf As Double

    Set wsLog = ThisWorkbook.Sheets("Simulation Runs")
    wsLog.Cells.ClearContents
    wsLog.Range("A1:E1").Value = Array("Run", "Receivables Days", "Inventory Days", "Payables Days", "Total FCF")

    For i = 1 To 100
        dso = 30 + Int((60 - 30 + 1) * Rnd) + Int((Rnd() - 0.5) * 10)
        dio = 20 + Int((45 - 20 + 1) * Rnd) + Int((Rnd() - 0.5) * 10)
        dpo = 30 + Int((70 - 30 + 1) * Rnd) + Int((Rnd() - 0.5) * 10)

        ' Fake FCF model — simplified estimation
        fcf = 50000 + (50 - dso) * 50 + (30 - dio) * 60 + (dpo - 50) * 40 + Rnd() * 10000

        wsLog.Cells(i + 1, 1).Value = i
        wsLog.Cells(i + 1, 2).Value = dso
        wsLog.Cells(i + 1, 3).Value = dio
        wsLog.Cells(i + 1, 4).Value = dpo
        wsLog.Cells(i + 1, 5).Value = Round(fcf, 2)
    Next i

    MsgBox "100 simulation runs completed. Check 'Simulation Runs' sheet.", vbInformation
End Sub
