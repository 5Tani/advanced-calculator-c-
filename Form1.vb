Public Class Frmbloodglucose
    Private Sub btngeneratereport_Click(sender As Object, e As EventArgs) Handles btngeneratereport.Click

        Dim BGlucose(10) As Double

        BGlucose(0) = CDbl(txtpatient1.Text)
        BGlucose(1) = CDbl(txtpatient2.Text)
        BGlucose(2) = CDbl(txtpatient3.Text)
        BGlucose(3) = CDbl(txtpatient4.Text)
        BGlucose(4) = CDbl(txtpatient5.Text)
        BGlucose(5) = CDbl(txtpatient6.Text)
        BGlucose(6) = CDbl(txtpatient7.Text)
        BGlucose(7) = CDbl(txtpatient8.Text)
        BGlucose(8) = CDbl(txtpatient9.Text)
        BGlucose(9) = CDbl(txtpatient10.Text)

        Dim Avg As Double = ComputeAvg(BGlucose)
        lstop.Items.Add("Average Blood Glucose is: " & Avg)
        lstop.Items.Add(vbNewLine)

        lstop.Items.Add("Highest Blood Glucose " & FindHighest(BGlucose))
        lstop.Items.Add("Lowest Blood Glucose" & FindLowest(BGlucose))
        lstop.Items.Add(vbNewLine)

        Dim Normal As Integar = 0
        Dim High As Integer = CountNormalHigh(BGlucose, Normal)
        lstop.Items.Add("Numbers of patient with Normal Blood Glucose: " & Normal)

        lstop.Items.Add("Numbers of Patient with High Blood Glucose " & High)
        lstop.Items.Add(vbNewLine)

        Call DisplayHighBlood(BGlucose)


    End Sub

    Module BloodGlucoseReport
        Function ComputAvg(ByVal BGlucose() As Double) As Double
            Dim Sum As Double = 0
            Dim Avg As Double
            For i As Integer = 0 To BGlucose.GetUpperBound(0) Step 1
                Sum += BGlucose(i)
            Next
            Avg = Sum / BGlucose.Length
            Return Avg
        End Function

        Function FindHighest(ByVal BGlucose() As Double) As Double
            Dim max As Double = BGlucose(0)
            For j As Integer = 1 To BGlucose.GetUpperBound(0) Step 1
                If max < BGlucose(j) Then
                    max = BGlucose(j)
                End If
            Next
            Return max
        End Function


        Function FindLowest(ByVal BGlucose() As Double) As Double
            Dim min As Double = BGlucose(0)
            For j As Integer = 1 To BGlucose.GetUpperBound(0) Step 1
                If min > BGlucose(j) Then
                    min = BGlucose(j)
                End If
            Next
            Return min
            Frmbloodglucose.lstResult.Items.Add(vbNewLine)
        End Function

        Function CountNormalHigh(ByVal BGlucose() As Double, ByRef Normal As Double) As Double
            Dim Count As Integer

            For k As Integer = 0 To BGlucose.GetUpperBound(0) Step 1
                If BGlucose(k) < 6.9 Then
                    Normal += 1
                ElseIf BGlucose(k) >= 6.9 Then
                    Count += 1
                End If
            Next
            Return Count
        End Function

        Sub DisplayHighBlood(ByVal BGlucose() As Double)

            Frmbloodglucose.lstResult.Items.Add("List of Patient with High Blood Glucose:")
            Frmbloodglucose.lstResult.Items.Add("Patient" & vbTab & "Blood Glucose")
            Frmbloodglucose.lstResult.Items.Add("--------------------------------------")
            For j As Integer = 0 To BGlucose.GetUpperBound(0) Step 1
                If BGlucose(j) >= 6.9 Then
                    Frmbloodglucose.lstResult.Items.Add(j + 1 & vbTab & BGlucose(j))
                End If
            Next
        End Sub
    End Module

End Class
