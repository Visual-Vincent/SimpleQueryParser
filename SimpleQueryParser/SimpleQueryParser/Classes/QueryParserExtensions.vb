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

Imports System.Runtime.CompilerServices

Public Module QueryParserExtensions

    ''' <summary>
    ''' Adds the specified command to the lookup table.
    ''' </summary>
    ''' <param name="Dictionary">The lookup table which to add the command to.</param>
    ''' <param name="Command">The command to add.</param>
    ''' <remarks></remarks>
    <Extension()>
    Public Sub Add(ByVal Dictionary As Dictionary(Of String, List(Of QueryCommand)), ByVal Command As QueryCommand)
        If Command Is Nothing Then Throw New ArgumentNullException("Command")
        If Command.Tokens.Length <= 0 Then Throw New ArgumentException("At least one token must be specified!", "Command.Tokens")
        If Command.Tokens(0).GetType() IsNot GetType(QueryKeyword) Then Throw New ArgumentException("First item must be of type QueryKeyword!", "Command.Tokens")
        If Array.IndexOf(Command.Tokens, Nothing) > -1 Then Throw New ArgumentException("'Command.Tokens' must not contain null!", "Command.Tokens")

        Dim Keyword As QueryKeyword = DirectCast(Command.Tokens(0), QueryKeyword)
        Dim CommandList As List(Of QueryCommand) = Nothing

        If Not Dictionary.TryGetValue(Keyword.Name, CommandList) Then
            CommandList = New List(Of QueryCommand)
            Dictionary.Add(Keyword.Name, CommandList)
        End If

        CommandList.Add(Command)
    End Sub

    ''' <summary>
    ''' Removes the specified command from the lookup table.
    ''' </summary>
    ''' <param name="Dictionary">The lookup table which to remove the command from.</param>
    ''' <param name="Command">The command to remove.</param>
    ''' <remarks></remarks>
    <Extension()>
    Public Function Remove(ByVal Dictionary As Dictionary(Of String, List(Of QueryCommand)), ByVal Command As QueryCommand) As Boolean
        If Command Is Nothing Then Throw New ArgumentNullException("Command")
        If Command.Tokens.Length <= 0 Then Throw New ArgumentException("At least one token must be specified!", "Command.Tokens")
        If Command.Tokens(0).GetType() IsNot GetType(QueryKeyword) Then Throw New ArgumentException("First item must be of type QueryKeyword!", "Command.Tokens")

        Dim Keyword As QueryKeyword = DirectCast(Command.Tokens(0), QueryKeyword)
        Dim CommandList As List(Of QueryCommand) = Nothing

        If Not Dictionary.TryGetValue(Keyword.Name, CommandList) Then Return False
        Return CommandList.Remove(Command)
    End Function
End Module
