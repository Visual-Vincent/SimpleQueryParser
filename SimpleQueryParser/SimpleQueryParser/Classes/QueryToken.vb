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

''' <summary>
''' The base class for tokens used by the SimpleQueryParser commands.
''' </summary>
''' <remarks></remarks>
Public MustInherit Class QueryToken

    Private _caseSensitive As Boolean
    Private _name As String

    ''' <summary>
    ''' Gets whether the token is case-sensitive or not.
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property CaseSensitive As Boolean
        Get
            Return _caseSensitive
        End Get
    End Property

    ''' <summary>
    ''' Gets the name of the token.
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property Name As String
        Get
            Return _name
        End Get
    End Property

    ''' <summary>
    ''' Initializes a new instance of the QueryToken class.
    ''' </summary>
    ''' <param name="Name">The name of the token.</param>
    ''' <param name="CaseSensitive">Whether or not the token is case-sensitive.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal Name As String, ByVal CaseSensitive As Boolean)
        If String.IsNullOrWhiteSpace(Name) Then Throw New ArgumentException("A token name must be specified!", "Name")
        If Name.Contains(" "c) OrElse Name.Contains(vbTab) Then Throw New ArgumentException("Token name must not contain any whitespace!", "Name")

        _name = Name
        _caseSensitive = CaseSensitive
    End Sub

    ''' <summary>
    ''' Returns the token's syntax in human-readable form.
    ''' </summary>
    ''' <remarks></remarks>
    Public MustOverride Function GetSyntax() As String

    ''' <summary>
    ''' Returns the string representation of the token.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Function ToString() As String
        Return Me.GetSyntax()
    End Function
End Class
