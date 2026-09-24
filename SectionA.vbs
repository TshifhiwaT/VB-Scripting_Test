OPTION EXPLICIT

Class Account
    Public accNumber
    Public accHolder
    Private balance

    Private sub Class_Initialize()
        balance = 1000
    End Sub

    Public Sub SetDetails(accNo, accHold)
        accNumber = accNo
        accHolder = accHold
    End Sub
    
    Public Sub Withdrawal(amount)
        if amount <= 0 then
            MsgBox "Withdrawal amount must be greater than 0", 0
        ElseIf amount > balance then
            MsgBox "Insufficient funds for this withdrawal", 0
        Else
            balance = balance - amount
            MsgBox "Withdrawal of " & amount & " successful. New balance: " & balance, 0
        End If
    End Sub

    Public Sub Deposit(amount)
        if amount <= 0 then
            MsgBox "Deposit amount must be greater than 0", 0
        Else
            balance = balance + amount
            MsgBox "Deposit of " & amount & " successful. New balance: " & balance, 0
        End If
    End Sub

    Public Sub Transfer(trAccount, amount)
        if amount <= 0 then
            MsgBox "Transfer amount must be greater than 0", 0
        ElseIf amount > balance then
            MsgBox "Insufficient funds for this transfer", 0
        Else
            balance = balance - amount
            MsgBox "Transfer of " & amount & " to "& trAccount & " successful. New balance: " & balance, 0
        End If
    End Sub 

    Public Sub CheckBalance()
        MsgBox "Your Balance Is: "& balance, 0
    End Sub

    Public Function getAccountNumber()
        getAccountNumber = "Account Number: ACC-"&accNumber
    End Function

    Function getAccountDetails()
        getAccountDetails = "Account Holder: "& accHolder & VbNewline & _
                            "Account Number: "& accNumber & VbNewline & _
                            "Balance: " & balance
    End Function

End Class


'Main Script
Dim strTitle
Dim choice
Dim amount
Dim trAcc
Dim accNo
Dim accHold
Dim objAccount

Set objAccount = New Account
    
strTitle = "Balance Tracking System"

'Takein account number
Do 
    accNo = InputBox("Enter Account Number: (only enter the 7-10 digits)")

    if Len(accNo) < 7 or Len(accNo) > 10 then
        MsgBox "Account number must be between 7 and 10 digits", 0, strTitle
    ElseIf Not IsNumeric(accNo) then
        MsgBox "Account number must be numeric", 0, strTitle
    Else
        Exit Do
    End If
Loop 

'take in account holder name
accHold = Trim(InputBox("Enter Account Holder Fullname: "))

objAccount.SetDetails accNo, accHold 'store the account details in the object

'Take in option from the user and perform the corresponding action
Do

choice = InputBox("1-Withdrawal" & VbNewline &  _
                    "2-Deposit" & VbNewline & _
                    "3-Transfer" & VbNewline & _
                    "4-Check Balance" & VbNewline & _
                    "5-Exit")
Select Case (choice)
    Case "1"

        amount = InputBox("Enter Amount to Withdraw:")
        If IsNumeric(amount) Then
        objAccount.Withdrawal CDbl(amount)
        Else
            MsgBox "Invalid amount entered. Please enter a numeric value.", 0, strTitle
        End If

    Case "2"

        amount = InputBox("Enter Amount to Deposit")
        If IsNumeric(amount) Then
            objAccount.Deposit CDbl(amount)
        Else
            MsgBox "Invalid amount entered. Please enter a numeric value.", 0, strTitle
        End If

    Case "3"

        trAcc = InputBox("Enter Account number you want to transfer too")
        amount = InputBox("Enter Amount to Transfer")
        If IsNumeric(amount) Then
            objAccount.Transfer trAcc, CDbl(amount)
        Else
            MsgBox "Invalid amount entered. Please enter a numeric value.", 0, strTitle
        End If

    Case "4"

        objAccount.CheckBalance()

    Case "5"

        Exit Do

    Case Else
        
        MsgBox "Invalid option selected. Please Select a number between 1 and 5", 0, strTitle
End Select

Loop
