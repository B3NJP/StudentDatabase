Imports System.Text.RegularExpressions

Public Class Form1
    'set up a record or "class" for a student
    Class STUDENT
        Public firstname As String
        Public lastname As String
        Public DOB As Date
        Public gender As Char
        Public avMk As Single
        Public phoneNo As String
        Public paid As Boolean
    End Class
    Dim students(9) As STUDENT
    Dim studentCount As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'allocate memory
        For i = 0 To 9
            students(i) = New STUDENT
        Next
        'load 4 test records
        students(0).firstname = "Johnny"
        students(0).lastname = "Depp"
        students(0).DOB = "9/6/63"
        students(0).gender = "m"
        students(0).avMk = 78.2
        students(0).phoneNo = "0123456789"
        students(0).paid = False
        students(1).firstname = "Jennifer"
        students(1).lastname = "Lawrence"
        students(1).DOB = "15/8/90"
        students(1).gender = "f"
        students(1).avMk = 88.2
        students(1).phoneNo = "0987654321"
        students(1).paid = True
        students(2).firstname = "George"
        students(2).lastname = "Clooney"
        students(2).DOB = "6/5/61"
        students(2).gender = "m"
        students(2).avMk = 68.2
        students(2).phoneNo = "0213465789"
        students(2).paid = False
        students(3).firstname = "Scarlett"
        students(3).lastname = "Johansson"
        students(3).DOB = "22/11/84"
        students(3).gender = "f"
        students(3).avMk = 72.2
        students(3).phoneNo = "0918273645"
        students(3).paid = True

        'set the student count to the number of students which have been entered
        studentCount = 4
        displayList()
    End Sub
    Private Sub btnAddStud_Click(sender As Object, e As EventArgs) Handles btnAddStud.Click
        'place text from text boxes into the array - first students(0), then students(1), students(2) etc
        If Not ErrorChecking(txtFirstName.Text, txtLastName.Text, txtDOB.Text, txtGender.Text, txtAvMk.Text, txtPhone.Text) Then
            Exit Sub
        End If
        students(studentCount).firstname = txtFirstName.Text
        students(studentCount).lastname = txtLastName.Text
        students(studentCount).DOB = txtDOB.Text
        students(studentCount).gender = txtGender.Text
        students(studentCount).avMk = txtAvMk.Text
        students(studentCount).phoneNo = txtPhone.Text
        students(studentCount).paid = chkPaid.Checked
        studentCount += 1
        'return text boxes to blank ready for next entry
        txtFirstName.Text = ""
        txtLastName.Text = ""
        txtDOB.Text = ""
        txtGender.Text = ""
        txtAvMk.Text = ""
        txtPhone.Text = ""
        chkPaid.Checked = False
        displayList()
    End Sub
    Private Sub displayList()
        'clear the list box as it keeps the earlier loop
        lstStud.Items.Clear()
        'loop through the array to print all rows
        For i = 0 To studentCount - 1
            lstStud.Items.Add(students(i).firstname & " - " & students(i).lastname & " - " &
                              students(i).DOB & " - " & students(i).gender & " - " & students(i).avMk & " - " & students(i).phoneNo & " - " & students(i).paid & ".")
        Next
    End Sub

    Private Function ErrorChecking(fname, lname, dob, gender, avMk, PhoneNo)
        If fname = "" Then
            MsgBox("Please enter a first name")
            Return False
        End If
        If lname = "" Then
            MsgBox("Please enter a last name")
            Return False
        End If
        If gender <> "m" And gender <> "f" Then
            MsgBox("Gender should be 'm' or 'f' lowercase")
            Return False
        End If
        If Not IsNumeric(avMk) OrElse (Convert.ToDecimal(avMk) < 0 Or Convert.ToDecimal(avMk) > 100) Then
            MsgBox("Make sure average mark is a number between 0 and 100 inclusive")
            Return False
        End If
        If Not Regex.IsMatch(PhoneNo, "\([0-9]{3}\) [0-9]{3}-[0-9]{4}") Then
            MsgBox("Make sure the phone number is 10 digits")
            Return False
        End If
        Return True
    End Function

    Private Sub btnFindStud_Click(sender As Object, e As EventArgs) Handles btnFindStud.Click

    End Sub
End Class
