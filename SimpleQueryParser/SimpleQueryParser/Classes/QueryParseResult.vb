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
''' A class holding information about the result of parsing a SimpleQueryParser query.
''' </summary>
''' <remarks></remarks>
<DebuggerDisplay("\{{If(Success, ""Result (command "" & CommandID & "" - "" & Parameters.Count & "" parameters)"", ""ERROR: "" & _error),nq}\}")>
Public NotInheritable Class QueryParseResult

    Private _commandID As Integer = -1
    Private _error As String = Nothing
    Private _parameters As Dictionary(Of String, String)

    ''' <summary>
    ''' Gets the error message describing why parsing failed.
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property [Error] As String
        Get
            Return _error
        End Get
    End Property

    ''' <summary>
    ''' Gets the ID of the parsed command.
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property CommandID As Integer
        Get
            Return _commandID
        End Get
    End Property

    ''' <summary>
    ''' Gets whether parsing was successful or not.
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property Success As Boolean
        Get
            Return _error Is Nothing AndAlso _commandID >= 0
        End Get
    End Property

    ''' <summary>
    ''' Gets a lookup table containing the parameters parsed from the query (if any).
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property Parameters As Dictionary(Of String, String)
        Get
            Return _parameters
        End Get
    End Property

    ''' <summary>
    ''' Initializes a new instance of the QueryParseResult class.
    ''' </summary>
    ''' <param name="CommandID">The ID of the parsed command.</param>
    ''' <param name="Parameters">Optional. A lookup table containing the parameters parsed from the query (if any).</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal CommandID As Integer, Optional ByVal Parameters As Dictionary(Of String, String) = Nothing)
        _commandID = CommandID
        _parameters = If(Parameters IsNot Nothing,
                         New Dictionary(Of String, String)(Parameters, StringComparer.OrdinalIgnoreCase),
                         New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase))
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the QueryParseResult class.
    ''' </summary>
    ''' <param name="Error">An error message describing why parsing failed.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal [Error] As String)
        _error = [Error]
    End Sub

    ''' <summary>
    ''' Gets the parsed value of the specified parameter. If the parameter does not exist this returns null.
    ''' </summary>
    ''' <param name="Name">The name of the parameter to get (case-insensitive).</param>
    ''' <remarks></remarks>
    Public Function GetParameter(ByVal Name As String) As String
        Dim Value As String = Nothing
        _parameters.TryGetValue(Name, Value)
        Return Value
    End Function
End Class
