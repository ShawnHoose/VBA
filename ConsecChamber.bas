Attribute VB_Name = "ConsecChamber"
Sub successTestData() 'Shawn Hoose 8-2-19
    Dim SN As String
    Dim model As String

    SN = Sheet4.Cells(1, 2)
    model = Sheet4.Cells(3, 2)
    
    Sheet4.Range("F7:G10").ClearContents 'Clears percentages
    Sheet4.Range("P7:Y65336").ClearContents 'Clears pass rate calculations
    Sheet4.Range("A7:D65536").ClearContents 'Clears prior data
      
    If Sheet4.Cells(3, 2) <> "" Then
        Call gatherDataModel(model, 2)
    Else
        Call gatherData(SN, 1)
    End If
    
End Sub

Sub indvCalcData() 'Shawn Hoose 8-5-19
    Dim rowCount As Integer
    Dim i As Integer
    Dim firstPass As Integer
    Dim secondPass As Integer
    Dim thirdPass As Integer
    Dim fourthPass As Integer
    Dim firstFail As Integer
    Dim secondFail As Integer
    Dim thirdFail As Integer
    Dim fourthFail As Integer
    Dim firstSum As Integer
    Dim secondSum As Integer
    Dim thirdSum As Integer
    Dim fourthSum As Integer
    
    
    rowCount = Sheet4.Range("A65536").End(xlUp).Row 'last row of data
    
    For i = 7 To rowCount 'Check if pass or fail for each chamber test
        If Sheet4.Cells(i, 2) = "Pass" And Sheet4.Cells(i, 3) = 1 Then
            firstPass = firstPass + 1
        ElseIf Sheet4.Cells(i, 2) = "Fail" And Sheet4.Cells(i, 3) = 1 Then
            firstFail = firstFail + 1
            
        ElseIf Sheet4.Cells(i, 2) = "Pass" And Sheet4.Cells(i, 3) = 2 Then
            secondPass = secondPass + 1
        ElseIf Sheet4.Cells(i, 2) = "Fail" And Sheet4.Cells(i, 3) = 2 Then
            secondFail = secondFail + 1
          
        ElseIf Sheet4.Cells(i, 2) = "Pass" And Sheet4.Cells(i, 3) = 3 Then
            thirdPass = thirdPass + 1
        ElseIf Sheet4.Cells(i, 2) = "Fail" And Sheet4.Cells(i, 3) = 3 Then
            thirdFail = thirdFail + 1
            
        ElseIf Sheet4.Cells(i, 2) = "Pass" And Sheet4.Cells(i, 3) = 4 Then
            fourthPass = fourthPass + 1
        ElseIf Sheet4.Cells(i, 2) = "Fail" And Sheet4.Cells(i, 3) = 4 Then
            fourthFail = fourthFail + 1
        End If
    Next i
    
    firstSum = firstPass + firstFail
    secondSum = secondPass + secondFail
    thirdSum = thirdPass + thirdFail
    fourthSum = fourthPass + fourthFail
    
    On Error Resume Next
    
    Sheet4.Cells(7, 6) = (firstPass / (firstPass + firstFail)) 'First chamber pass rate
    If firstPass = 0 And firstPass + firstFail = 0 Then
        Sheet4.Cells(7, 6) = "-"
    End If
    
    Sheet4.Cells(8, 6) = (secondPass / (secondPass + secondFail)) 'Second pass rate
    If secondPass = 0 And secondPass + secondFail = 0 Then
        Sheet4.Cells(8, 6) = "-"
    End If
    
    Sheet4.Cells(9, 6) = (thirdPass / (thirdPass + thirdFail)) 'Third pass rate
    If thirdPass = 0 And thirdPass + thirdFail = 0 Then
        Sheet4.Cells(9, 6) = "-"
    End If
    
    Sheet4.Cells(10, 6) = (fourthPass / (fourthPass + fourthFail)) 'Fourth pass rate
    If fourthPass = 0 And fourthPass + fourthFail = 0 Then
        Sheet4.Cells(10, 6) = "-"
    End If
    
    Sheet4.Cells(7, 7) = firstSum
    Sheet4.Cells(8, 7) = secondSum
    Sheet4.Cells(9, 7) = thirdSum
    Sheet4.Cells(10, 7) = fourthSum
    
End Sub

Sub gatherData(laser As String, colNum As Integer) 'Shawn Hoose 8-5-19
    Dim rowCount As Integer
    Dim i As Integer
    Dim dataRow As Integer
    Dim y As Integer
    Dim startDate As String
    Dim endDate As String
    Dim cycleCount As String
    Dim prompt As String
    Dim prompt2 As String
    
    Application.Calculation = xlManual 'turn off automatic calculations
    
    prompt = "The last tested date is greater than 3 days ago. Was this system placed into chamber from the QC Shelf?"
    prompt2 = "Was this system placed in a weekend chamber cycle?"
    dataRow = 7
    cycleCount = 1
    startDate = Sheet4.Cells(1, 5)
    endDate = Sheet4.Cells(2, 5)
 
    rowCount = Sheet1.Range("A65536").End(xlUp).Row 'last row of data
    
    For i = 2 To rowCount
        If Sheet1.Cells(i, colNum) = laser Then
            If (DateValue(Sheet1.Cells(i, 3)) <= DateValue(endDate)) And (DateValue(Sheet1.Cells(i, 3)) >= DateValue(startDate)) Then 'Ensure within date range
                If Sheet1.Cells(i, 4) = 1 Or Sheet1.Cells(i, 4) = 0 And Sheet1.Cells(i, 4) <> "" Then 'If the SN in the data matches the requested SN
                    Sheet4.Cells(dataRow, 1) = Sheet1.Cells(i, 3) 'Copy date
                    
                    If Sheet1.Cells(i, 4) = 1 Then 'Determines if the test passed or failed
                        Sheet4.Cells(dataRow, 2) = "Pass"
                    Else
                        Sheet4.Cells(dataRow, 2) = "Fail"
                    End If
                    dataRow = dataRow + 1 'iter counter
                End If
            End If
        End If
    Next i
    
    For y = 7 To dataRow - 1
        If y > 7 Then
            If ((DateValue(Sheet4.Cells(y, 1)) - DateValue(Sheet4.Cells(y - 1, 1))) > 3) Then 'If prior date was a pass and is consecutive, then up the iterator
                If MsgBox(prompt, vbYesNo) = vbYes Then 'MsgBox asking if from QC Shelf
                    cycleCount = cycleCount
                Else
                    cycleCount = 1
                End If
                
                If Sheet4.Cells(y, 2) = "Pass" Then
                    If MsgBox(prompt2, vbYesNo) = vbYes Then 'MsgBox asking if weekend chamber
                        cycleCount = cycleCount + 1
                    Else
                        cycleCount = cycleCount
                    End If
                End If
                
            End If
        End If
        
        Sheet4.Cells(y, 3) = cycleCount
        
        If Sheet4.Cells(y, 2) = "Fail" Then 'Print cycle fail count
            cycleCount = 1 'reset cycleCount
        Else
            cycleCount = cycleCount + 1
        
        End If
               
    Next y
    
    Call indvCalcData
    Application.Calculation = xlAutomatic 'turn on automatic caluclations
    
End Sub


Sub gatherDataModel(laser As String, colNum As Integer) 'Shawn Hoose 8-6-19
    Dim rowCount As Integer
    Dim i As Integer
    Dim dataRow As Integer
    Dim y As Integer
    Dim startDate As String
    Dim endDate As String
    Dim cycleCount As String
    Dim dict As Scripting.Dictionary 'SN:cycleCount
    Dim prompt As String
    Dim prompt2 As String
    Dim dict2 As Scripting.Dictionary 'SN:Late date tested
        
    Set dict = New Scripting.Dictionary
    Set dict2 = New Scripting.Dictionary
    dataRow = 7
    cycleCount = 1
    startDate = Sheet4.Cells(1, 5)
    endDate = Sheet4.Cells(2, 5)
    
    Application.Calculation = xlManual 'Turn off automatic calculations
    
    rowCount = Sheet1.Range("A65536").End(xlUp).Row 'last row of data

    For i = 2 To rowCount
        If Sheet1.Cells(i, colNum) = laser Then
            If (DateValue(Sheet1.Cells(i, 3)) <= DateValue(endDate)) And (DateValue(Sheet1.Cells(i, 3)) >= DateValue(startDate)) Then 'Ensure within date range
                If Sheet1.Cells(i, 4) = 1 Or Sheet1.Cells(i, 4) = 0 And Sheet1.Cells(i, 4) <> "" Then 'If the SN in the data matches the requested SN
                    Sheet4.Cells(dataRow, 1) = Sheet1.Cells(i, 3) 'Copy date
                    Sheet4.Cells(dataRow, 4) = Sheet1.Cells(i, 1) 'Copy SN
                    
                    If Not (dict.Exists(Sheet1.Cells(i, 1).Value)) Then
                        dict(Sheet1.Cells(i, 1).Value) = 1 ' Set dict value to cycleCount default
                    End If
                    
                    If Not (dict2.Exists(Sheet1.Cells(i, 1).Value)) Then
                        dict2(Sheet1.Cells(i, 1).Value) = Sheet4.Cells(dataRow, 1) 'initialize second dictionary to hold dates
                    End If
                    
                    If Sheet1.Cells(i, 4) = 1 Then 'Determines if the test passed or failed
                        Sheet4.Cells(dataRow, 2) = "Pass"
                    Else
                        Sheet4.Cells(dataRow, 2) = "Fail"
                    End If
                    dataRow = dataRow + 1 'iter counter
                End If
            End If
        End If
    Next i
    
    Dim key As Variant
    i = 7
    For Each key In dict.Keys
        Sheet4.Cells(i, 16) = key 'print keys for data analysis
        i = i + 1
    Next key
    
    For y = 7 To dataRow - 1 'loop across all data rows
        For Each key In dict.Keys 'check row against all key:value pairs
        
            prompt = "The last test date for " & key & " is greater than 3 days ago. Was this system put into chamber from the QC Shelf?"
            prompt2 = "Was " & key & " placed in a weekend chamber?"
            
            If Sheet4.Cells(y, 2) = "Fail" And Sheet4.Cells(y, 4) = key Then 'Print cycle fail count, ensuring serial numbers match for each fail
                
                If (DateValue(Sheet4.Cells(y, 1)) - DateValue(dict2(key)) > 3) And (dict(key) <> 1) Then
                    If MsgBox(prompt, vbYesNo) = vbYes Then 'MsgBox asking if from QC Shelf
                        dict(key) = dict(key)
                    Else
                        dict(key) = 1
                    End If

                End If
                
                Sheet4.Cells(y, 3) = dict(key)
                dict2(key) = Sheet4.Cells(y, 1) 'update date of most recent test for SN
                dict(key) = 1 'reset cycleCount
                
            Else
                If Sheet4.Cells(y, 2) = "Pass" And Sheet4.Cells(y, 4) = key Then

                    If (DateValue(Sheet4.Cells(y, 1)) - DateValue(dict2(key)) > 3) Then
                        If MsgBox(prompt, vbYesNo) = vbYes Then 'MsgBox asking if from QC Shelf
                            dict(key) = dict(key)
                        Else
                            dict(key) = 1
                        End If
                        
                        If MsgBox(prompt2, vbYesNo) = vbYes Then 'MsgBox asking if weekend chamber
                            dict(key) = dict(key) + 1
                        Else
                            dict(key) = dict(key)
                        End If
                        
                    End If
                    
                    dict2(key) = Sheet4.Cells(y, 1) 'update date of most recent test for SN
                    Sheet4.Cells(y, 3) = dict(key) 'print chamber test number
                    
                End If
            End If
         Next key
    Next y
    
    Call modelCalcData(dict)
    Application.Calculation = xlAutomatic 'turn on automatic calculations
    
End Sub
Sub modelCalcData(dict As Dictionary) 'Shawn Hoose 8-5-19
    Dim rowCount As Integer
    Dim i As Integer
    Dim firstPass As Integer
    Dim secondPass As Integer
    Dim thirdPass As Integer
    Dim fourthPass As Integer
    Dim firstFail As Integer
    Dim secondFail As Integer
    Dim thirdFail As Integer
    Dim fourthFail As Integer
    Dim y As Integer
    Dim firstSum As Integer
    Dim secondSum As Integer
    Dim thirdSum As Integer
    Dim fourthSum As Integer
    Dim z As Integer
    Dim totalFirstPass As Integer
    Dim totalSecondPass As Integer
    Dim totalThirdPass As Integer
    Dim totalFourthPass As Integer
    
    y = 7
    rowCount = Sheet4.Range("A65536").End(xlUp).Row 'last row of data
    
    For Each key In dict.Keys
        'reset counters
        firstPass = 0
        firstFail = 0
        secondPass = 0
        secondFail = 0
        thirdPass = 0
        thirdFail = 0
        fourthPass = 0
        fourthFail = 0
        
        For i = 7 To rowCount 'Check if pass or fail for each chamber test and if it matches the serial number we're looking at
            If Sheet4.Cells(i, 2) = "Pass" And Sheet4.Cells(i, 3) = 1 And Sheet4.Cells(i, 4) = key Then
                firstPass = firstPass + 1
            ElseIf Sheet4.Cells(i, 2) = "Fail" And Sheet4.Cells(i, 3) = 1 And Sheet4.Cells(i, 4) = key Then
                firstFail = firstFail + 1
            
            
            ElseIf Sheet4.Cells(i, 2) = "Pass" And Sheet4.Cells(i, 3) = 2 And Sheet4.Cells(i, 4) = key Then
                secondPass = secondPass + 1
            ElseIf Sheet4.Cells(i, 2) = "Fail" And Sheet4.Cells(i, 3) = 2 And Sheet4.Cells(i, 4) = key Then
                secondFail = secondFail + 1
            
              
            ElseIf Sheet4.Cells(i, 2) = "Pass" And Sheet4.Cells(i, 3) = 3 And Sheet4.Cells(i, 4) = key Then
                thirdPass = thirdPass + 1
            ElseIf Sheet4.Cells(i, 2) = "Fail" And Sheet4.Cells(i, 3) = 3 And Sheet4.Cells(i, 4) = key Then
                thirdFail = thirdFail + 1
             
                
            ElseIf Sheet4.Cells(i, 2) = "Pass" And Sheet4.Cells(i, 3) = 4 And Sheet4.Cells(i, 4) = key Then
                fourthPass = fourthPass + 1
            ElseIf Sheet4.Cells(i, 2) = "Fail" And Sheet4.Cells(i, 3) = 4 And Sheet4.Cells(i, 4) = key Then
                fourthFail = fourthFail + 1
            
            
            End If
        Next i
        
        On Error Resume Next
        
        Sheet4.Cells(y, 17) = (firstPass / (firstPass + firstFail)) 'First chamber pass rate
        If firstPass = 0 And firstPass + firstFail = 0 Then
            Sheet4.Cells(y, 17) = "-"
        End If
        
        Sheet4.Cells(y, 18) = (secondPass / (secondPass + secondFail)) 'Second pass rate
        If secondPass = 0 And secondPass + secondFail = 0 Then
            Sheet4.Cells(y, 18) = "-"
        End If
        
        Sheet4.Cells(y, 19) = (thirdPass / (thirdPass + thirdFail)) 'Third pass rate
        If thirdPass = 0 And thirdPass + thirdFail = 0 Then
            Sheet4.Cells(y, 19) = "-"
        End If
        
        Sheet4.Cells(y, 20) = (fourthPass / (fourthPass + fourthFail)) 'Fourth pass rate
        If fourthPass = 0 And fourthPass + fourthFail = 0 Then
            Sheet4.Cells(y, 20) = "-"
        End If
        
        'Show total sums for all serial numbers for each numbered chamber test
        firstSum = firstSum + firstPass + firstFail
        secondSum = secondSum + secondPass + secondFail
        thirdSum = thirdSum + thirdPass + thirdFail
        fourthSum = fourthSum + fourthPass + fourthFail
        
        'Show how many times each serial number attempted each numbered chamber test
        Sheet4.Cells(y, 22) = firstPass + firstFail
        Sheet4.Cells(y, 23) = secondPass + secondFail
        Sheet4.Cells(y, 24) = thirdPass + thirdFail
        Sheet4.Cells(y, 25) = fourthPass + fourthFail
        
        totalFirstPass = totalFirstPass + firstPass
        totalSecondPass = totalSecondPass + secondPass
        totalThirdPass = totalThirdPass + thirdPass
        totalFourthPass = totalFourthPass + fourthPass
        
        y = y + 1
    Next key
    
    'display total number of tests per chamber number
    Sheet4.Cells(7, 7) = firstSum
    Sheet4.Cells(8, 7) = secondSum
    Sheet4.Cells(9, 7) = thirdSum
    Sheet4.Cells(10, 7) = fourthSum
    

    For z = 17 To 20 'loop over all 4 chamber test numbers
        If z = 17 Then
            If totalFirstPass = 0 And firstSum = 0 Then
                Sheet4.Cells(7, 6) = "-"
            Else
                Sheet4.Cells(7, 6) = totalFirstPass / firstSum
            End If
        ElseIf z = 18 Then
            If totalSecondPass = 0 And secondSum = 0 Then
                Sheet4.Cells(8, 6) = "-"
            Else
                Sheet4.Cells(8, 6) = totalSecondPass / secondSum
            End If
        ElseIf z = 19 Then
            If totalThirdPass = 0 And thirdSum = 0 Then
                Sheet4.Cells(9, 6) = "-"
            Else
                Sheet4.Cells(9, 6) = totalThirdPass / thirdSum
            End If
        Else
            If totalFourthPass = 0 And fourthSum = 0 Then
                Sheet4.Cells(10, 6) = "-"
            Else
                Sheet4.Cells(10, 6) = totalFourthPass / fourthSum
            End If
        End If
    Next z
    
End Sub
