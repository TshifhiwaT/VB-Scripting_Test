OPTION EXPLICIT

Class Account
    Public accNumber
    Public accHolder
    Private balance
    
    Sub Withdrawal(amount)
        balance = balance - amount
    End Sub

    Sub Deposit(amount)
        balance = balance + amount
    End Sub

    Sub Transfer(trAccount, amount)
        balance = balance - amount
    End Sub 

    Sub CheckBalance()
        MsgBox "Your Balance Is: "& balance, 0
    End Sub

    Function getAccountNumber()
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
Dim option

Set objAccount = New Account
    
strTitle = "Balance Tracking"

objAccount.accHolder = InputBox("Enter Account Holder Fullname: ")
Do 
    objAccount.accNumber = IsNumeric(InputBox("Enter Account Number: (only enter the 7-10 digits)"))
Loop Until 7 <= objAccount.accNumber.Len <=10

option = IsNumeric(InputBox("1-Withdrawal" & VbNewline &  _
                            "2-Deposit" & VbNewline & _
                            "3-Transfer" & VbNewline & _
                            "4-Check Balance" & VbNewline & -
                            "5-Exit"))
Select Case (option)
    Dim temp
    Case 1
        temp = InputBox("Enter Amount to Withdraw:")
        objAccount.Withdrawal(temp)
    case 2
        temp = InputBox("Enter Amount to Deposit")
        objAccount.Deposit(temp)
    Case 3
        temp = InputBox("Enter Amount to Transfer")
        Dim trAcc = InputBox("Enter Account number you want to transfer too")
        objAccount.Transfer(trAcc, temp)
    Case 4
        objAccount.CheckBalance()
    Case 5

    Case Else
        MsgBox "Invalid option selected", 0, strTitle
End Select
