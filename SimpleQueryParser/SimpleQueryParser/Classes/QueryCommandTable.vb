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

Imports System.Text.RegularExpressions

''' <summary>
''' A class holding information about multiple SimpleQueryParser commands.
''' </summary>
''' <remarks></remarks>
<DebuggerDisplay("\{QueryCommandTable ({Commands.Count} command{If(Commands.Count = 1, """", ""s""),nq})\}")>
Public NotInheritable Class QueryCommandTable
    Implements IEnumerable, IEnumerable(Of KeyValuePair(Of Integer, QueryCommand))

    Private Shared ReadOnly ParamRegex As New Regex("PARAM\((?<name>[^,\r\n]+?)(?:,[ \t]*(?<maxlength>[0-9]+)|,[ \t]*\{(?<validvalues>[ \t]*"".+?"")\})?\)", RegexOptions.Compiled Or RegexOptions.IgnoreCase)
    Private Shared ReadOnly ValuesRegex As New Regex("\A""(?<first>(?:\\""|[^""])+?)""(?:,[ \t]*""(?<consecutive>(?:\\""|[^""])+?)"")*\z", RegexOptions.Compiled Or RegexOptions.IgnoreCase)

    Private UnderlyingDictionary As New Dictionary(Of Integer, QueryCommand)

    ''' <summary>
    ''' Adds a command to the command table.
    ''' </summary>
    ''' <param name="ID">The unique ID to give the command.</param>
    ''' <param name="Tokens">An array of sequentially ordered tokens which to construct the command with.</param>
    ''' <param name="Handler">Optional. A method that will handle the result of the command when parsed.</param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal ID As Integer, ByVal Tokens As QueryToken(), Optional ByVal Handler As QueryParsedEventHandler = Nothing)
        If Tokens Is Nothing Then Throw New ArgumentNullException("Tokens")
        If Tokens.Length <= 0 Then Throw New ArgumentException("At least one token must be specified!", "Tokens")
        If Tokens(0).GetType() IsNot GetType(QueryKeyword) Then Throw New ArgumentException("First item must be of type QueryKeyword!", "Tokens")
        If Array.IndexOf(Tokens, Nothing) > -1 Then Throw New ArgumentException("'Tokens' must not contain null!", "Tokens")
        If UnderlyingDictionary.ContainsKey(ID) Then Throw New ArgumentException("A command with ID " & ID & " already exists!", "ID")

        UnderlyingDictionary.Add(ID, New QueryCommand(ID, Tokens, Handler))
    End Sub

    ''' <summary>
    ''' Parses and adds a command to the command table.
    ''' </summary>
    ''' <param name="ID">The unique ID to give the command.</param>
    ''' <param name="Command">The command string to parse. To include parameters use PARAM([name]) or PARAM([name], [maxlength]) or PARAM([name], {"list", "of", "valid", "values"}).</param>
    ''' <param name="CaseSensitive">Whether or not the command (including its parameters) is case-sensitive.</param>
    ''' <param name="Handler">Optional. A method that will handle the result of the command when parsed.</param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal ID As Integer, ByVal Command As String, Optional ByVal CaseSensitive As Boolean = False, Optional ByVal Handler As QueryParsedEventHandler = Nothing)
        If String.IsNullOrWhiteSpace(Command) Then Throw New ArgumentException("'Command' cannot be empty!", "Command")
        If UnderlyingDictionary.ContainsKey(ID) Then Throw New ArgumentException("A command with ID " & ID & " already exists!", "ID")

        Dim TokensList As New List(Of QueryToken)
        Dim ParamsList As New List(Of QueryParameter)

        Command = ParamRegex.Replace(Command, _
            Function(m As Match) As String
                Dim MaxLengthGroup As Group = m.Groups("maxlength")
                Dim ValidValuesGroup As Group = m.Groups("validvalues")

                If MaxLengthGroup.Success Then
                    Dim MaxLength As Integer = 0
                    If Not Integer.TryParse(MaxLengthGroup.Value, MaxLength) Then Throw New OverflowException("Invalid value for 'MaxLength'!")
                    ParamsList.Add(New QueryParameter(m.Groups("name").Value, MaxLength))

                ElseIf ValidValuesGroup.Success Then
                    Dim ValidValues As New List(Of String)
                    Dim Values As Match = ValuesRegex.Match(ValidValuesGroup.Value)

                    ValidValues.Add(Values.Groups("first").Value)

                    For Each Capture As Capture In Values.Groups("consecutive").Captures
                        ValidValues.Add(Capture.Value)
                    Next

                    ParamsList.Add(New QueryParameter(m.Groups("name").Value, CaseSensitive, ValidValues.ToArray()))

                Else
                    ParamsList.Add(New QueryParameter(m.Groups("name").Value))
                End If

                Return "{{PARAM_STUB}}"
            End Function
        )

        Dim Tokens As String() = Command.Split(New Char() {" "c, Convert.ToChar(vbTab)}, StringSplitOptions.RemoveEmptyEntries)
        Dim ParamIndex As Integer = 0

        For Each Token As String In Tokens
            If Token = "{{PARAM_STUB}}" Then
                TokensList.Add(ParamsList(ParamIndex))
                ParamIndex += 1
                Continue For
            End If

            TokensList.Add(New QueryKeyword(Token, CaseSensitive))
        Next

        UnderlyingDictionary.Add(ID, New QueryCommand(ID, TokensList.ToArray(), Handler))
    End Sub

    ''' <summary>
    ''' Removes the specified command from the table.
    ''' </summary>
    ''' <param name="ID">The unique ID of the command which to remove.</param>
    ''' <remarks></remarks>
    Public Function Remove(ByVal ID As Integer) As Boolean
        Return UnderlyingDictionary.Remove(ID)
    End Function

    Public Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of System.Collections.Generic.KeyValuePair(Of Integer, QueryCommand)) Implements System.Collections.Generic.IEnumerable(Of System.Collections.Generic.KeyValuePair(Of Integer, QueryCommand)).GetEnumerator
        Return UnderlyingDictionary.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return UnderlyingDictionary.GetEnumerator()
    End Function
End Class
