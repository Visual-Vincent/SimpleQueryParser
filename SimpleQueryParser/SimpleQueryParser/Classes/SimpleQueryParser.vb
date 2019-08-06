'Copyright (c) 2019, Vincent Bengtsson
'All rights reserved.
'
'Redistribution and use in source and binary forms, with or without
'modification, are permitted provided that the following conditions are met:
'
'1. Redistributions of source code must retain the above copyright notice, this
'   list of conditions and the following disclaimer.
'
'2. Redistributions in binary form must reproduce the above copyright notice,
'   this list of conditions and the following disclaimer in the documentation
'   and/or other materials provided with the distribution.
'
'3. Neither the name of the copyright holder nor the names of its
'   contributors may be used to endorse or promote products derived from
'   this software without specific prior written permission.
'
'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
'AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
'IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
'DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
'FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
'DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
'SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
'CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
'OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

Imports System.Text

Public Delegate Sub QueryParsedEventHandler(ByVal sender As Object, e As QueryParseResult)

'TODO:
'  - Support functions.
'  - Support lists/arrays (both with and without parantheses? - i.e. "SELECT column1, column2, ...").
'  - Support for optional parameters?
'  - Support for a variable number of parameters? (low priority, may not be necessary even)
'  - Support for parameters with spaces without surrounding them with quotes? (i.e. SELECT <name> FROM <table> WHERE <expression string>)

''' <summary>
''' A class for parsing basic commands and queries.
''' </summary>
''' <remarks></remarks>
<DebuggerDisplay("\{SimpleQueryParser ({Commands.Count} command{If(Commands.Count = 1, """", ""s""),nq})\}")>
Public Class SimpleQueryParser
    Private ReadOnly Commands As New Dictionary(Of Integer, QueryCommand)
    Private ReadOnly FirstTokenLookup As New Dictionary(Of String, List(Of QueryCommand))(StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' Initializes a new instance of the SimpleQueryParser class.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the SimpleQueryParser class.
    ''' </summary>
    ''' <param name="Commands">A table of commands to be used by the query parser.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal Commands As QueryCommandTable)
        Me.AddCommands(Commands)
    End Sub

    ''' <summary>
    ''' Adds a new command to the query parser.
    ''' </summary>
    ''' <param name="ID">The ID to assign to the command.</param>
    ''' <param name="Tokens">An array of sequentially ordered tokens which define the structure of the command.</param>
    ''' <param name="Handler">Optional. A handler which will be called every time this command has been successfully parsed.</param>
    ''' <remarks></remarks>
    Protected Sub AddCommand(ByVal ID As Integer, ByVal Tokens As QueryToken(), Optional ByVal Handler As QueryParsedEventHandler = Nothing)
        Me.AddCommand(ID, New QueryCommand(ID, Tokens, Handler))
    End Sub

    ''' <summary>
    ''' Adds a new command to the query parser.
    ''' </summary>
    ''' <param name="ID">The ID to assign to the command.</param>
    ''' <param name="Command">The command to add.</param>
    ''' <remarks></remarks>
    Protected Sub AddCommand(ByVal ID As Integer, ByVal Command As QueryCommand)
        If Command Is Nothing Then Throw New ArgumentNullException("Command")
        If Command.Tokens.Length <= 0 Then Throw New ArgumentException("At least one token must be specified!", "Command.Tokens")
        If Command.Tokens(0).GetType() IsNot GetType(QueryKeyword) Then Throw New ArgumentException("First item must be of type QueryKeyword!", "Command.Tokens")
        If Array.IndexOf(Command.Tokens, Nothing) > -1 Then Throw New ArgumentException("'Command.Tokens' must not contain null!", "Command.Tokens")
        If Commands.ContainsKey(ID) Then Throw New ArgumentException("A command with ID " & ID & " already exists!", "ID")

        Dim Parameters As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each Parameter As QueryParameter In Command.Tokens.OfType(Of QueryParameter)()
            If Parameters.Contains(Parameter.Name) Then
                Throw New ArgumentException("Command contains multiple parameters by the same name! (" & Parameter.Name & ")", "Command")
            End If
            Parameters.Add(Parameter.Name)
        Next

        Me.Commands.Add(ID, Command)
        Me.FirstTokenLookup.Add(Command)
    End Sub

    ''' <summary>
    ''' Adds new commands to the query parser.
    ''' </summary>
    ''' <param name="CommandTable">The table of commands which to add.</param>
    ''' <remarks></remarks>
    Protected Sub AddCommands(ByVal CommandTable As QueryCommandTable)
        For Each kvp As KeyValuePair(Of Integer, QueryCommand) In CommandTable
            Me.AddCommand(kvp.Key, kvp.Value)
        Next
    End Sub

    ''' <summary>
    ''' Removes a command from the query parser.
    ''' </summary>
    ''' <param name="ID">The ID of the command which to remove.</param>
    ''' <remarks></remarks>
    Protected Function RemoveCommand(ByVal ID As Integer) As Boolean
        Dim Command As QueryCommand = Nothing
        If Not Me.Commands.TryGetValue(ID, Command) Then
            Return False
        End If

        Me.Commands.Remove(ID)
        Me.FirstTokenLookup.Remove(Command)
        Return True
    End Function

    ''' <summary>
    ''' Attempts to parse the specified input as one of the commands in this query parser.
    ''' </summary>
    ''' <param name="Input">The input string to parse.</param>
    ''' <remarks></remarks>
    Public Function Parse(ByVal Input As String) As QueryParseResult
        If Input Is Nothing Then Throw New ArgumentNullException("Input")
        Input = Input.Trim()

        Dim Tokens As New List(Of String)
        Dim TokenBuilder As New StringBuilder(32)

        Dim StartQuotationMark As Integer = -1
        Dim BackslashCount As Integer = 0
        Dim i As Integer = 1

        'Parse input.
        For Each c As Char In Input
            Select Case c

                Case " "c, CType(vbTab, Char)
                    'Skip consecutive whitespace.
                    If StartQuotationMark < 0 AndAlso TokenBuilder.Length = 0 Then Exit Select

                    If StartQuotationMark < 0 Then
                        'End of token, add to list.
                        Tokens.Add(TokenBuilder.ToString())
                        TokenBuilder.Clear()
                    Else
                        'If we're inside quotation marks append whitespace.
                        TokenBuilder.Append(c)
                    End If

                Case "\"c
                    BackslashCount += 1

                Case """"c
                    Dim AppendDoubleQuote As Boolean = False

                    'Are we inside quotation marks?
                    If StartQuotationMark >= 0 Then

                        If BackslashCount Mod 2 = 0 Then
                            'Non-escaped quotation mark.
                            StartQuotationMark = -1
                        Else
                            'Quotation mark escaped by backslash.
                            AppendDoubleQuote = True
                        End If

                    Else 'Not inside quotation marks.

                        If BackslashCount = 0 Then
                            'Unescaped quotation mark.
                            StartQuotationMark = (i - 1)

                        ElseIf BackslashCount > 1 AndAlso BackslashCount Mod 2 = 0 Then
                            'Do not allow backslashes before unescaped quotation mark.
                            Return New QueryParseResult("Unexpected '""' at position " & i)

                        ElseIf BackslashCount Mod 2 <> 0 Then
                            'Escaped quotation mark.
                            AppendDoubleQuote = True
                        End If

                    End If

                    'Append backslashes.
                    If BackslashCount > 0 Then TokenBuilder.Append(New String("\"c, Convert.ToInt32(Math.Floor(BackslashCount / 2))))
                    BackslashCount = 0

                    'Append the quote (if escaped).
                    If AppendDoubleQuote Then TokenBuilder.Append(c)

                Case Else
                    'Append backslashes.
                    If BackslashCount > 0 Then TokenBuilder.Append(New String("\"c, Convert.ToInt32(Math.Floor(BackslashCount / 2))))
                    BackslashCount = 0

                    'Append char.
                    TokenBuilder.Append(c)

            End Select

            i += 1
        Next

        If StartQuotationMark >= 0 Then
            Return New QueryParseResult("Unclosed quotation mark at position " & StartQuotationMark)
        End If

        'Add the last input token.
        If TokenBuilder.Length > 0 Then
            Tokens.Add(TokenBuilder.ToString())
            TokenBuilder.Clear()
        End If

        If Tokens.Count <= 0 OrElse String.IsNullOrWhiteSpace(Tokens(0)) Then Return New QueryParseResult("No input")

        'Get a list of possible commands based on the first input keyword.
        Dim Commands As List(Of QueryCommand) = Nothing
        If Not Me.FirstTokenLookup.TryGetValue(Tokens(0), Commands) Then
            Return New QueryParseResult("Unknown command")
        End If

        Dim Command As QueryCommand = Nothing
        Dim Parameters As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Dim StateTracker As New StateTracker

        'HighestToken   - Holds the highest index of all currently correct tokens. Used to determine partial commands.
        'InvalidToken   - Holds temporary information about an invalid token while we check for any other commands which might allow it.
        'ParameterError - Holds a temporary parameter error while we check for any other commands which might allow it.
        Dim HighestToken As TokenTracker = StateTracker.HighestToken
        Dim InvalidToken As InvalidTokenTracker = StateTracker.InvalidToken
        Dim ParameterError As ParameterErrorTracker = StateTracker.ParameterError

        HighestToken.Command = Commands(0)

        i = 0
        While i < Commands.Count
            Command = Commands(i)

            '+---------------------------------------------------------------------------+'
            '|                                   NOTE:                                   |'
            '|        Below statement was removed to favour more detailed errors.        |'
            '|   While the statement improves speed it makes the errors more ambiguous   |'
            '| as partial commands result in "Unknown command" instead of "Expected ..." |'
            '+---------------------------------------------------------------------------+'

            'If the number of tokens don't match then this isn't the right command.
            'If Tokens.Count <> Command.Tokens.Length Then
            '    i += 1
            '    Continue While
            'End If

            'If the current number of tokens are less then what we have successfully processed so far then this isn't the right command.
            If Command.Tokens.Length < HighestToken.Index + 1 Then
                i += 1
                Continue While
            End If

            Parameters.Clear()

            'Iterate input tokens.
            For x = 1 To Tokens.Count - 1
                Dim Token As String = Tokens(x)
                Dim CurrentToken As QueryToken = Command.Tokens(x)

                'Check input token against command token.
                Select Case CurrentToken.GetType()

                    Case GetType(QueryKeyword)
                        Dim Keyword As QueryKeyword = DirectCast(CurrentToken, QueryKeyword)

                        'Does the input keyword match the command keyword?
                        If Not Token.Equals(Keyword.Name, If(Keyword.CaseSensitive, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)) Then
                            i += 1

                            Dim CorrectPart As String = String.Join(" ", Tokens.Take(x))
                            If x > HighestToken.Index AndAlso (InvalidToken.CorrectPart Is Nothing OrElse CorrectPart = InvalidToken.CorrectPart) Then
                                InvalidToken.CorrectPart = CorrectPart
                                InvalidToken.Token = Token
                                InvalidToken.Expected.Add(Keyword.Name)

                                StateTracker.LastError = ErrorTypes.InvalidToken
                            End If

                            Continue While
                        End If

                    Case GetType(QueryParameter)
                        Dim Parameter As QueryParameter = DirectCast(CurrentToken, QueryParameter)

                        'Does the input parameter exceed the maximum allowed parameter length? (if any)
                        If Parameter.MaxLength > 0 AndAlso Token.Length > Parameter.MaxLength AndAlso x > ParameterError.Index AndAlso x > HighestToken.Index Then
                            ParameterError.Index = x
                            ParameterError.Error = "Parameter """ & Parameter.Name & """ must not be longer than " & Parameter.MaxLength & " characters"

                            StateTracker.LastError = ErrorTypes.ParameterError

                            i += 1
                            Continue While
                        End If

                        Dim Comparison As StringComparison = _
                            If(Parameter.CaseSensitive, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

                        'Does the input parameter match any of the predefined valid values? (if any)
                        If Parameter.ValidValues IsNot Nothing AndAlso Parameter.ValidValues.Length > 0 _
                            AndAlso Array.FindIndex(Parameter.ValidValues, Function(Item As String) Token.Equals(Item, Comparison)) < 0 Then

                            If x > ParameterError.Index AndAlso x > HighestToken.Index Then
                                ParameterError.Index = x
                                ParameterError.Error = "Invalid value """ & Token & """ for parameter """ & Parameter.Name & """" 'TODO: Include valid ones?

                                StateTracker.LastError = ErrorTypes.ParameterError
                            End If

                            i += 1
                            Continue While
                        End If

                        'Valid parameter.
                        Parameters.Add(Parameter.Name, Token)

                End Select

                If x > HighestToken.Index Then
                    HighestToken.Index = x
                    HighestToken.Command = Command

                    StateTracker.LastError = ErrorTypes.IncompleteCommand
                End If

                If ParameterError.Index >= 0 AndAlso x >= ParameterError.Index Then
                    'Matching parameter in other command found.
                    ParameterError.Reset()
                End If

                If x = Command.Tokens.Length - 1 Then
                    'Does the input contain only a partial command?
                    If Tokens.Count <> Command.Tokens.Length Then Exit For

                    'Success!
                    Dim Result As New QueryParseResult(Command.ID, Parameters)

                    If Command.Handler IsNot Nothing Then _
                        Command.Handler.Invoke(Me, Result)

                    Return Result
                End If
            Next

            i += 1
        End While

        'Invalid parameter; no other command allowed it.
        If ParameterError.Index >= 0 AndAlso ParameterError.Error IsNot Nothing Then
            Return New QueryParseResult(ParameterError.Error)
        End If

        'Invalid token specified.
        If InvalidToken.Token IsNot Nothing AndAlso InvalidToken.Expected.Count = 1 Then
            Return New QueryParseResult("Invalid token """ & InvalidToken.Token & """. Expected """ & InvalidToken.Expected(0) & """")
        ElseIf InvalidToken IsNot Nothing AndAlso InvalidToken.Expected.Count > 1 Then
            Dim Expected As HashSet(Of String) = InvalidToken.Expected
            Return New QueryParseResult("Invalid token """ & InvalidToken.Token & """. Expected either """ & String.Join(""", """, Expected.Take(Expected.Count - 1)) & """ or """ & Expected(Expected.Count - 1) & """")
        End If

        'Invalid number of parameters for command.
        If HighestToken.Command IsNot Nothing AndAlso HighestToken.Index >= 0 Then

            'Too few parameters.
            If HighestToken.Index < HighestToken.Command.Tokens.Length - 1 Then
                Dim ExpectedToken As QueryToken = HighestToken.Command.Tokens(HighestToken.Index + 1)

                Select Case ExpectedToken.GetType()
                    Case GetType(QueryKeyword)
                        Dim Expected As New HashSet(Of String)

                        'Contruct list of expected keywords.
                        For Each c As QueryCommand In Commands
                            If HighestToken.Index >= c.Tokens.Length - 1 Then Continue For

                            Dim Token As QueryKeyword = TryCast(c.Tokens(HighestToken.Index + 1), QueryKeyword)
                            If Token IsNot Nothing Then
                                Expected.Add(Token.Name)
                            End If
                        Next

                        If Expected.Count = 1 Then
                            Return New QueryParseResult("Expected """ & ExpectedToken.Name & """")
                        Else
                            Return New QueryParseResult("Expected either """ & String.Join(""", """, Expected.Take(Expected.Count - 1)) & """ or """ & Expected(Expected.Count - 1) & """")
                        End If

                    Case GetType(QueryParameter)
                        Return New QueryParseResult("Parameter """ & ExpectedToken.Name & """ must be specified")
                End Select

            Else 'Too many parameters.
                Return New QueryParseResult("Too many parameters. " & HighestToken.Command.Tokens.Length & " expected, got " & Tokens.Count)
            End If

        End If

        Return New QueryParseResult("Unknown command")
    End Function

    Private Enum ErrorTypes As Integer
        None = 0
        ParameterError
        InvalidToken
        IncompleteCommand
    End Enum

    Private Interface IStateTracker
        Sub Reset()
    End Interface

    Private Class StateTracker

        Private _lastError As ErrorTypes = ErrorTypes.None
        Private _highestToken As New TokenTracker
        Private _invalidToken As New InvalidTokenTracker
        Private _parameterError As New ParameterErrorTracker

        Public Property LastError As ErrorTypes
            Get
                Return _lastError
            End Get
            Set(value As ErrorTypes)
                _lastError = value

                Select Case _lastError
                    Case ErrorTypes.None
                        _invalidToken.Reset()
                        _highestToken.Reset()
                        _parameterError.Reset()

                    Case ErrorTypes.IncompleteCommand
                        _parameterError.Reset()
                        _invalidToken.Reset()

                    Case ErrorTypes.InvalidToken
                        _parameterError.Reset()

                    Case ErrorTypes.ParameterError
                        _invalidToken.Reset()

                End Select
            End Set
        End Property

        Public ReadOnly Property HighestToken As TokenTracker
            Get
                Return _highestToken
            End Get
        End Property

        Public ReadOnly Property InvalidToken As InvalidTokenTracker
            Get
                Return _invalidToken
            End Get
        End Property

        Public ReadOnly Property ParameterError As ParameterErrorTracker
            Get
                Return _parameterError
            End Get
        End Property
    End Class

    Private Class InvalidTokenTracker
        Implements IStateTracker

        Public Property Index As Integer = -1
        Public Property Token As String = Nothing
        Public Property CorrectPart As String = Nothing
        Public Property Expected As New HashSet(Of String)

        Public Sub Reset() Implements IStateTracker.Reset
            Me.Index = -1
            Me.Token = Nothing
            Me.CorrectPart = Nothing
            Me.Expected.Clear()
        End Sub
    End Class

    Private Class ParameterErrorTracker
        Implements IStateTracker

        Public Property Index As Integer = -1
        Public Property [Error] As String = Nothing

        Public Sub Reset() Implements IStateTracker.Reset
            Me.Index = -1
            Me.Error = Nothing
        End Sub
    End Class

    Private Class TokenTracker
        Implements IStateTracker

        Public Property Index As Integer = 0
        Public Property Command As QueryCommand = Nothing

        Public Sub Reset() Implements IStateTracker.Reset
            Me.Index = 0
            Me.Command = Nothing
        End Sub
    End Class
End Class
