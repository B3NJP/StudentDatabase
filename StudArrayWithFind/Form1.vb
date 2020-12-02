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
    Dim students(20) As STUDENT
    Dim studentCount As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'allocate memory
        For i = 0 To 20
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
        students(4).firstname = "Johnny"
        students(4).lastname = "Deppy"
        students(4).DOB = "9/6/68"
        students(4).gender = "m"
        students(4).avMk = 78.2
        students(4).phoneNo = "0123456789"
        students(4).paid = False
        students(5).firstname = "Fred"
        students(5).lastname = "Bear"
        students(5).DOB = "9/4/63"
        students(5).gender = "m"
        students(5).avMk = 78.2
        students(5).phoneNo = "0123456789"
        students(5).paid = False
        students(6).firstname = "Mickey"
        students(6).lastname = "Mouse"
        students(6).DOB = "28/12/98"
        students(6).gender = "m"
        students(6).avMk = 78.2
        students(6).phoneNo = "0123456789"
        students(6).paid = False
        students(7).firstname = "Fred"
        students(7).lastname = "Flintstone"
        students(7).DOB = "4/11/03"
        students(7).gender = "m"
        students(7).avMk = 78.2
        students(7).phoneNo = "0123456789"
        students(7).paid = False
        students(8).firstname = "Minnie"
        students(8).lastname = "Mouse"
        students(8).DOB = "9/6/63"
        students(8).gender = "f"
        students(8).avMk = 78.2
        students(8).phoneNo = "0123455555"
        students(8).paid = True
        students(9).firstname = "Boo"
        students(9).lastname = "Depp"
        students(9).DOB = "9/6/01"
        students(9).gender = "m"
        students(9).avMk = 99
        students(9).phoneNo = "0123456789"
        students(9).paid = False
        students(10).firstname = "Lori Anne"
        students(10).lastname = "Allison"
        students(10).DOB = "9/6/63"
        students(10).gender = "f"
        students(10).avMk = 66
        students(10).phoneNo = "0123456789"
        students(10).paid = False
        students(11).firstname = "Amber"
        students(11).lastname = "Heard"
        students(11).DOB = "28/2/00"
        students(11).gender = "f"
        students(11).avMk = 77
        students(11).phoneNo = "0123456789"
        students(11).paid = False
        students(12).firstname = "Pistol"
        students(12).lastname = "Depp"
        students(12).DOB = "9/6/98"
        students(12).gender = "m"
        students(12).avMk = 3
        students(12).phoneNo = "0123456789"
        students(12).paid = False
        'set the student count to the number of students which have been entered
        studentCount = 13
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
        lstStud.SelectedIndex = studentCount - 1
    End Sub

    Private Function ErrorChecking(fname, lname, dob, gender, avMk, PhoneNo)
        If Not Regex.IsMatch(fname, "[A-Z][a-z]*") Then
            MsgBox("Please enter a first name")
            Return False
        End If
        If Not Regex.IsMatch(lname, "[A-Z][a-z]*") Then
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

    Private Function StringifyStudent(stud)
        Return stud.firstname & " - " & stud.lastname & " - " &
                              stud.DOB & " - " & stud.gender & " - " & stud.avMk & " - " & stud.phoneNo & " - " & stud.paid & "."
    End Function

    Private Sub btnFindStud_Click(sender As Object, e As EventArgs) Handles btnFindStud.Click
        If txtLastName.Text = "" Then
            MsgBox("Enter last name")
            Exit Sub
        End If
        Dim foundStud As Boolean = False
        For Each stud In students
            If stud.lastname = txtLastName.Text Then
                MsgBox(StringifyStudent(stud))
                foundStud = True
            End If
        Next
        If Not foundStud Then
            MsgBox("No students found")
        End If
    End Sub

    Private Sub lstStud_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstStud.SelectedIndexChanged
        selStud.Text = sender.SelectedItem
    End Sub

    Private Sub btnEditStud_Click(sender As Object, e As EventArgs) Handles btnEditStud.Click
        txtFirstName.Text = students(lstStud.SelectedIndex).firstname
        txtLastName.Text = students(lstStud.SelectedIndex).lastname
        txtDOB.Text = students(lstStud.SelectedIndex).DOB
        txtGender.Text = students(lstStud.SelectedIndex).gender
        txtAvMk.Text = students(lstStud.SelectedIndex).avMk
        txtPhone.Text = students(lstStud.SelectedIndex).phoneNo
        chkPaid.Checked = students(lstStud.SelectedIndex).paid
        btnAddStud.Enabled = False
        btnReplStud.Visible = True
        lstStud.Enabled = False
        btnEditStud.Enabled = False
    End Sub

    Private Sub btnReplStud_Click(sender As Object, e As EventArgs) Handles btnReplStud.Click
        If Not ErrorChecking(txtFirstName.Text, txtLastName.Text, txtDOB.Text, txtGender.Text, txtAvMk.Text, txtPhone.Text) Then
            Exit Sub
        End If
        students(lstStud.SelectedIndex).firstname = txtFirstName.Text
        students(lstStud.SelectedIndex).lastname = txtLastName.Text
        students(lstStud.SelectedIndex).DOB = txtDOB.Text
        students(lstStud.SelectedIndex).gender = txtGender.Text
        students(lstStud.SelectedIndex).avMk = txtAvMk.Text
        students(lstStud.SelectedIndex).phoneNo = txtPhone.Text
        students(lstStud.SelectedIndex).paid = chkPaid.Checked
        txtFirstName.Text = ""
        txtLastName.Text = ""
        txtDOB.Text = ""
        txtGender.Text = ""
        txtAvMk.Text = ""
        txtPhone.Text = ""
        chkPaid.Checked = False
        displayList()
        btnReplStud.Visible = False
        btnAddStud.Enabled = True
        lstStud.Enabled = True
        btnEditStud.Enabled = True
    End Sub
End Class
