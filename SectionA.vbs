OPTION EXPLICIT

Class Account
    Public accNumber
    Public accHolder
    Private balance

    Private sub Class_Initialize()
        balance = 1000
    End Sub
    
    Public Sub Withdrawal(amount)
        if amount <= 0 then
            MsgBox "Withdrawal amount must be greater than 0", 0
        ElseIf amount > balance then
            MsgBox "Insufficient funds for this withdrawal", 0
        Else
            balance = balance - amount
            MsgBox "Withdrawal of " & amount & " successful." & VbNewLine &_
                    getAccountDetails(), 0
        End If
    End Sub

    Public Sub Deposit(amount)
        if amount <= 0 then
            MsgBox "Deposit amount must be greater than 0", 0
        Else
            balance = balance + amount
            MsgBox "Deposit of " & amount & " successful." & VbNewLine &_
                    getAccountDetails(), 0
        End If
    End Sub

    Public Sub Transfer(trAccount, amount)
        if amount <= 0 then
            MsgBox "Transfer amount must be greater than 0", 0
        ElseIf amount > balance then
            MsgBox "Insufficient funds for this transfer", 0
        Else
            balance = balance - amount
            MsgBox "Transfer of " & amount & " to "& trAccount & " successful." & VbNewLine &_
                    getAccountDetails(), 0
        End If
    End Sub 

    Public Sub CheckBalance()
        MsgBox getAccountDetails(), 0
    End Sub

    Public Function getAccountNumber()
        getAccountNumber = "ACC-"&accNumber
    End Function

    Function getAccountDetails()
        getAccountDetails = "Account Holder: "& accHolder & VbNewline & _
                            "Account Number: "& getAccountNumber() & VbNewline & _
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

'Take in account number
Do 
    accNo = InputBox("Enter Account Number: (only enter the 7-10 digits)")

    if Not IsNumeric(accNo)  then
        MsgBox "Account number must be numeric", 0, strTitle
    ElseIf Len(accNo) < 7 or Len(accNo) > 10 then
        MsgBox "Account number must be between 7 and 10 digits", 0, strTitle
    Else
        Exit Do
    End If
Loop 

'take in account holder name
Do 
    accHold = Trim(InputBox("Enter Account Holder Fullname: "))
    If accHold = "" Then
        MsgBox "Account holder name cannot be empty", 0, strTitle
    Else
        Exit Do
    End If
Loop

'Assign the account number and holder name to the account object
objAccount.accNumber = accNo
objAccount.accHolder = accHold

'Take in option from the user and perform the corresponding action
Do

choice = InputBox("1-Withdrawal" & VbNewline &  _
                    "2-Deposit" & VbNewline & _
                    "3-Transfer" & VbNewline & _
                    "4-Check Balance" & VbNewline & _
                    "5-Exit")
Select Case (choice)
    'Withdrawal
    Case "1"

        amount = InputBox("Enter Amount to Withdraw:")
        If IsNumeric(amount) Then
        objAccount.Withdrawal CDbl(amount)
        Else
            MsgBox "Invalid amount entered. Please enter a numeric value.", 0, strTitle
        End If

    'Deposit
    Case "2"

        amount = InputBox("Enter Amount to Deposit")
        If IsNumeric(amount) Then
            objAccount.Deposit CDbl(amount)
        Else
            MsgBox "Invalid amount entered. Please enter a numeric value.", 0, strTitle
        End If

    'Transfer
    Case "3"

        'capture account number to transfer to
        Do
            trAcc = InputBox("Enter Account number you want to transfer too")
             if Not IsNumeric(trAcc)  then
               MsgBox "Account number must be numeric", 0, strTitle
            ElseIf Len(trAcc) < 7 or Len(trAcc) > 10 then
               MsgBox "Account number must be between 7 and 10 digits", 0, strTitle
            Else
               Exit Do
            End If
        Loop
        
        'capture amount to transfer
        amount = InputBox("Enter Amount to Transfer")
        If IsNumeric(amount) Then
            objAccount.Transfer trAcc, CDbl(amount)
        Else
            MsgBox "Invalid amount entered. Please enter a numeric value.", 0, strTitle
        End If

    'Check Balance
    Case "4"

        objAccount.CheckBalance()

    'Exit
    Case "5"

        Exit Do

    Case Else
        
        MsgBox "Invalid option selected. Please Select a number between 1 and 5", 0, strTitle
End Select

Loop
