' A module for evaluating expressions, with support for
' parenthesis and many math functions
' Example:
'   Dim expr As String = "(SQR(9)^3)+COS(0)*3+ABS(-10)"
'   txtResult.Text = Evaluate(expr).ToString ' ==> 27+3+10 ==> 40    

Imports System.Text.RegularExpressions

Module EvaluateModule

    Function Evaluate(ByVal expr As String) As Double
        ' A number is a sequence of digits optionally followed by a dot and 
        ' another sequence of digits. The number in parenthesis in order to 
        ' define an unnamed group.
        Const Num As String = "(\-?\d+\.?\d*)"
        ' List of 1-operand functions.
        Const Func1 As String = "(exp|log|log10|abs|sqr|sqrt|sin|cos|tan|asin|ac" _
            & "os|atan)"
        ' List of 2-operand functions.
        Const Func2 As String = "(atan2)"
        ' List of N-operand functions.
        Const FuncN As String = "(min|max)"
        ' List of predefined constants.
        Const Constants As String = "(e|pi)"

        ' Define one Regex object for each supported operation.
        ' They are outside the loop, so that they are compiled only once.
        ' Binary operations are defined as two numbers with a symbol between 
        ' them
        ' optionally separated by spaces.
        Dim rePower As New Regex(Num & "\s*(\^)\s*" & Num)
        Dim reAddSub As New Regex(Num & "\s*([-+])\s*" & Num)
        Dim reMulDiv As New Regex(Num & "\s*([*/])\s*" & Num)
        ' These Regex objects resolve call to functions. (Case insensitivity.)
        Dim reFunc1 As New Regex(Func1 & "\(\s*" & Num & "\s*\)",
            RegexOptions.IgnoreCase)
        Dim reFunc2 As New Regex(Func2 & "\(\s*" & Num & "\s*,\s*" & Num &
            "\s*\)", RegexOptions.IgnoreCase)
        Dim reFuncN As New Regex(FuncN & "\((\s*" & Num & "\s*,)+\s*" & Num &
            "\s*\)", RegexOptions.IgnoreCase)
        ' This Regex object drop a + when it follows an operator.
        Dim reSign1 As New Regex("([-+/*^])\s*\+")
        ' This Regex object converts a double minus into a plus.
        Dim reSign2 As New Regex("\-\s*\-")
        ' This Regex object drops parenthesis around a number.
        ' (must not be preceded by an alphanum char (it might be a function 
        ' name)
        Dim rePar As New Regex("(?<![A-Za-z0-9])\(\s*([-+]?\d+.?\d*)\s*\)")
        ' A Regex object that tells that the entire expression is a number
        Dim reNum As New Regex("^\s*[-+]?\d+\.?\d*\s*$")

        ' The Regex object deals with constants. (Requires case insensitivity.)
        Dim reConst As New Regex("\s*" & Constants & "\s*",
            RegexOptions.IgnoreCase)
        ' This resolves predefined constants. (Can be kept out of the loop.)
        expr = reConst.Replace(expr, AddressOf DoConstants)

        ' Loop until the entire expression becomes just a number.
        Do Until reNum.IsMatch(expr)
            Dim saveExpr As String = expr

            ' Perform all the math operations in the source string.
            ' starting with operands with higher operands.
            ' Note that we continue to perform each operation until there are
            ' no matches, because we must account for expressions like (12*34*
            ' 56)

            ' Perform all power operations.
            Do While rePower.IsMatch(expr)
                expr = rePower.Replace(expr, AddressOf DoPower)
            Loop

            ' Perform all divisions and multiplications.
            Do While reMulDiv.IsMatch(expr)
                expr = reMulDiv.Replace(expr, AddressOf DoMulDiv)
            Loop

            ' Perform functions with variable numbers of arguments. 
            Do While reFuncN.IsMatch(expr)
                expr = reFuncN.Replace(expr, AddressOf DoFuncN)
            Loop

            ' Perform functions with 2 arguments. 
            Do While reFunc2.IsMatch(expr)
                expr = reFunc2.Replace(expr, AddressOf DoFunc2)
            Loop

            ' 1-operand functions must be processed last to deal correctly with 
            ' expressions such as SIN(ATAN(1)) before we drop parenthesis 
            ' pairs around numbers.
            Do While reFunc1.IsMatch(expr)
                expr = reFunc1.Replace(expr, AddressOf DoFunc1)
            Loop

            ' Discard + symbols (unary pluses)that follow another operator.
            expr = reSign1.Replace(expr, "$1")
            ' Simplify 2 consecutive minus signs into a plus sign.
            expr = reSign2.Replace(expr, "+")

            ' Perform all additions and subtractions.
            Do While reAddSub.IsMatch(expr)
                expr = reAddSub.Replace(expr, AddressOf DoAddSub)
            Loop

            ' attempt to discard parenthesis around numbers. We can do this
            expr = rePar.Replace(expr, "$1")

            ' if the expression didn't change, we have a syntax error.
            ' this serves to avoid endless loops
            If expr = saveExpr Then
                ' if it didn't work, exit with syntax error exception.
                Throw New SyntaxErrorException()
            End If
        Loop

        ' Return the expression, which is now a number.
        Return CDbl(expr)
    End Function

    ' These functions evaluate the actual math operations.
    ' In all cases the Match object on entry has groups that identify
    ' the two operands and the operator.

    Function DoConstants(ByVal m As Match) As String
        Select Case m.Groups(1).Value.ToUpper
            Case "PI"
                Return Math.PI.ToString
            Case "E"
                Return Math.E.ToString
        End Select
    End Function

    Function DoPower(ByVal m As Match) As String
        Dim n1 As Double = CDbl(m.Groups(1).Value)
        Dim n2 As Double = CDbl(m.Groups(3).Value)
        ' Group(2) is always the ^ character in this version.
        Return (n1 ^ n2).ToString
    End Function

    Function DoMulDiv(ByVal m As Match) As String
        Dim n1 As Double = CDbl(m.Groups(1).Value)
        Dim n2 As Double = CDbl(m.Groups(3).Value)
        Select Case m.Groups(2).Value
            Case "/"
                Return (n1 / n2).ToString
            Case "*"
                Return (n1 * n2).ToString
        End Select
    End Function

    Function DoAddSub(ByVal m As Match) As String
        Dim n1 As Double = CDbl(m.Groups(1).Value)
        Dim n2 As Double = CDbl(m.Groups(3).Value)
        Select Case m.Groups(2).Value
            Case "+"
                Return (n1 + n2).ToString
            Case "-"
                Return (n1 - n2).ToString
        End Select
    End Function

    ' These functions evaluate functions.

    Function DoFunc1(ByVal m As Match) As String
        ' function argument is 2nd group.
        Dim n1 As Double = CDbl(m.Groups(2).Value)
        ' function name is 1st group.
        Select Case m.Groups(1).Value.ToUpper
            Case "EXP"
                Return Math.Exp(n1).ToString
            Case "LOG"
                Return Math.Log(n1).ToString
            Case "LOG10"
                Return Math.Log10(n1).ToString
            Case "ABS"
                Return Math.Abs(n1).ToString
            Case "SQR", "SQRT"
                Return Math.Sqrt(n1).ToString
            Case "SIN"
                Return Math.Sin(n1).ToString
            Case "COS"
                Return Math.Cos(n1).ToString
            Case "TAN"
                Return Math.Tan(n1).ToString
            Case "ASIN"
                Return Math.Asin(n1).ToString
            Case "ACOS"
                Return Math.Acos(n1).ToString
            Case "ATAN"
                Return Math.Atan(n1).ToString
        End Select
    End Function

    Function DoFunc2(ByVal m As Match) As String
        ' function arguments are 2nd and 3rd group.
        Dim n1 As Double = CDbl(m.Groups(2).Value)
        Dim n2 As Double = CDbl(m.Groups(3).Value)
        ' function name is 1st group.
        Select Case m.Groups(1).Value.ToUpper
            Case "ATAN2"
                Return Math.Atan2(n1, n2).ToString
        End Select
    End Function

    Function DoFuncN(ByVal m As Match) As String
        ' function arguments are from group 2 onward.
        Dim args As New ArrayList()
        Dim i As Integer = 2
        ' Load all the arguments into the array.
        Do While m.Groups(i).Value <> ""
            ' Get the argument, replace any comma to space,
            '  and convert to double.
            args.Add(CDbl(m.Groups(i).Value.Replace(","c, " "c)))
            i += 1
        Loop

        ' function name is 1st group.
        Select Case m.Groups(1).Value.ToUpper
            Case "MIN"
                args.Sort()
                Return args(0).ToString
            Case "MAX"
                args.Sort()
                Return args(args.Count - 1).ToString
        End Select
    End Function

End Module


' Note: This code is taken from Francesco Balena's
' "Programming Microsoft Visual Basic .NET" - MS Press 2002, ISBN 0735613753
' You can read a free chapter of the book at 
' http://www.vb2themax.com/HtmlDoc.asp?Table=Books&ID=101000