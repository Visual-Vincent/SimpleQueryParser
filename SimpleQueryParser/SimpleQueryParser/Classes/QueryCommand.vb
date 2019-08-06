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

''' <summary>
''' A class holding information about a SimpleQueryParser command.
''' </summary>
''' <remarks></remarks>
<DebuggerDisplay("\{Command {ID} ({Tokens.Count} token{If(Tokens.Count = 1, """", ""s""),nq})\}")>
Public NotInheritable Class QueryCommand

    Private _id As Integer
    Private _handler As QueryParsedEventHandler
    Private _tokens As QueryToken()

    ''' <summary>
    ''' Gets a pointer to the handler (if any) called upon successful parsing of this command.
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property Handler As QueryParsedEventHandler
        Get
            Return _handler
        End Get
    End Property

    ''' <summary>
    ''' Gets the ID of the command.
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property ID As Integer
        Get
            Return _id
        End Get
    End Property

    ''' <summary>
    ''' Gets the array of sequentially ordered tokens which define the structure of the command.
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property Tokens As QueryToken()
        Get
            Return _tokens
        End Get
    End Property

    ''' <summary>
    ''' Initializes a new instance of the QueryCommand class.
    ''' </summary>
    ''' <param name="ID">The ID to assign to the command.</param>
    ''' <param name="Tokens">An array of sequentially ordered tokens which define the structure of the command.</param>
    ''' <param name="Handler">Optional. A handler which will be called every time this command has been successfully parsed.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal ID As Integer, ByVal Tokens As QueryToken(), Optional ByVal Handler As QueryParsedEventHandler = Nothing)
        If Tokens Is Nothing Then Throw New ArgumentNullException("FirstToken")
        If Array.IndexOf(Tokens, Nothing) > -1 Then Throw New ArgumentException("'Tokens' must not contain null!", "Tokens")

        _id = ID
        _tokens = Tokens
        _handler = _handler
    End Sub

    ''' <summary>
    ''' Returns the command's syntax in human-readable form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Function GetSyntax() As String
        Dim SyntaxBuilder As New StringBuilder(64)

        For i = 0 To Me.Tokens.Length - 1
            SyntaxBuilder.Append(Me.Tokens(i).GetSyntax())
            If i < Me.Tokens.Length - 1 Then SyntaxBuilder.Append(" ")
        Next

        Return SyntaxBuilder.ToString()
    End Function

    ''' <summary>
    ''' Returns the string representation of the command.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Function ToString() As String
        Return Me.GetSyntax()
    End Function
End Class