# 🧾 Oracle Fusion Statement Reconciliation Assistant – Excel VBA Macro

 📌 Project Summary  
This Excel VBA macro automates the cleanup and **preparation of Supplier Statement of Account (SoA) exports** from Oracle Fusion (Feogen), enabling **faster and more accurate reconciliation**.  
The tool solved a critical operational bottleneck for our finance team, reducing hours of manual work each month and minimizing the risk of human error.

---

 ✅ Key Achievements  
- Reduced SoA processing time from hours to seconds  
- Improved data accuracy and reduced manual validation  
- Standardized input format for reconciliation-ready reports  
- Enhanced visibility into duplicate or suspicious invoice entries  
- Enabled faster vendor account **reconciliation**  

---

🧠 Problem It Solved  
Vendor **reconciliation** was one of our most time-consuming monthly tasks.  
The raw monthly Statement of Account (SoA) exports from Oracle Fusion came in inconsistent and unstructured formats:

- PO and Invoice numbers were merged (e.g. `912345/INV`)
- Values were partially numeric, inconsistent, or incomplete  
- Many rows required cleanup before reconciliation could even begin  
- Rows with zero value or quantity cluttered the file  
- Reversal matching (positive/negative pairs) was fully manual  
- Formatting inconsistencies slowed down review

---

 🚀 The VBA Macro Solution  
The macro, named `Split_PO_INV_WithSlash`, performs the following:

- Splits merged PO/INV fields into structured columns  
- Flags inconsistent or suspicious entries  
- Highlights duplicate invoice numbers via conditional formatting  
- Validates reversals based on value sign  
- Removes rows with zero quantity and amount  
- Auto-formats and filters the output for clean **reconciliation-ready** review  

All logic was tailored to our team's SoA structure and reconciliation workflow.

---
🔒 Note on Confidentiality  
Due to financial data sensitivity and internal business rules:

- The actual live Excel data is **not shared**  
- This repository is a **portfolio case study** to showcase real-world automation logic

---

🛠 Skills Used  
- ✅ Excel VBA (Advanced Macros)  
- ✅ String Parsing, Validation Logic, Conditional Formatting  
- ✅ Dictionary Objects & Duplicate Detection  
- ✅ ERP Awareness (Oracle Fusion, AP workflows)

---

👤 Author  
**Badriah Jaber**  
Finance Data Automation | Excel VBA Specialist | Process Optimizer  
🔗 [LinkedIn](https://www.linkedin.com/in/badriah-jaber)

💬 *Let’s connect if you’re tackling similar reconciliation challenges!*

🔧 Full VBA Macro Code – `Split_PO_INV_WithSlash`
#
```vba
Sub Split_PO_INV_WithSlash()
    Columns("J:J").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("K:K").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("L:L").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("J5").Value = "k"
    Range("K5").Value = "l"
    Range("L5").Value = "m"

    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim splitValues() As String
    Dim invoiceNumber As String

    lastRow = Cells(Rows.Count, "I").End(xlUp).Row

    For i = 6 To lastRow
        cellValue = Cells(i, "I").Value
        splitValues = Split(cellValue, "/", 2)

        If UBound(splitValues) >= 0 Then
            If Len(splitValues(0)) > 0 And Left(splitValues(0), 1) = "9" Then
                Cells(i, "J").Value = splitValues(0)
                If UBound(splitValues) > 0 Then Cells(i, "K").Value = splitValues(1)
            ElseIf InStr(cellValue, "INV") > 0 Or Not IsNumeric(Left(cellValue, 1)) Then
                Cells(i, "K").Value = cellValue
            ElseIf InStr(cellValue, "/") > 0 Then
                Cells(i, "J").Value = splitValues(0)
                If UBound(splitValues) > 0 Then Cells(i, "K").Value = splitValues(1)
            Else
                Cells(i, "J").Value = cellValue
            End If
        End If
    Next i

    For i = 6 To lastRow
        If IsEmpty(Cells(i, "K").Value) And Not IsEmpty(Cells(i, "J").Value) Then
            Cells(i, "K").Value = Cells(i, "J").Value
        End If
    Next i

    With Range("K6:K" & lastRow)
        .FormatConditions.AddUniqueValues
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        .FormatConditions(1).DupeUnique = xlDuplicate
        With .FormatConditions(1).Font
            .Color = -16383844
        End With
        With .FormatConditions(1).Interior
            .Color = 13551615
        End With
    End With

    For i = lastRow To 6 Step -1
        If Cells(i, "S").Value = 0 And Cells(i, "R").Value = 0 Then
            Rows(i).Delete
        End If
    Next i

    Dim invoiceDict As Object
    Set invoiceDict = CreateObject("Scripting.Dictionary")

    For i = 6 To lastRow
        invoiceNumber = Cells(i, "K").Value
        If invoiceDict.exists(invoiceNumber) Then
            invoiceDict(invoiceNumber) = invoiceDict(invoiceNumber) + 1
        Else
            invoiceDict.Add invoiceNumber, 1
        End If
    Next i

    For i = 6 To lastRow
        invoiceNumber = Cells(i, "K").Value
        If invoiceDict.exists(invoiceNumber) Then
            If invoiceDict(invoiceNumber) Mod 2 <> 0 Then
                Cells(i, "K").Interior.Color = RGB(204, 153, 255)
            End If
        End If
    Next i

    Dim positiveNegativeDict As Object
    Set positiveNegativeDict = CreateObject("Scripting.Dictionary")

    For i = 6 To lastRow
        invoiceNumber = Cells(i, "K").Value
        If invoiceDict.exists(invoiceNumber) And invoiceDict(invoiceNumber) Mod 2 = 0 Then
            If Not positiveNegativeDict.exists(invoiceNumber) Then
                positiveNegativeDict.Add invoiceNumber, Array(False, False)
            End If
            If Cells(i, "T").Value > 0 Then
                positiveNegativeDict(invoiceNumber)(0) = True
            ElseIf Cells(i, "T").Value < 0 Then
                positiveNegativeDict(invoiceNumber)(1) = True
            End If
        End If
    Next i

    For i = 6 To lastRow
        invoiceNumber = Cells(i, "K").Value
        If invoiceDict.exists(invoiceNumber) And invoiceDict(invoiceNumber) Mod 2 = 0 Then
            If positiveNegativeDict.exists(invoiceNumber) Then
                If positiveNegativeDict(invoiceNumber)(0) And positiveNegativeDict(invoiceNumber)(1) Then
                    Cells(i, "K").Interior.ColorIndex = xlNone
                End If
            End If
        End If
    Next i

    Range("A5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.AutoFilter
    Range("K6").Select
End Sub
