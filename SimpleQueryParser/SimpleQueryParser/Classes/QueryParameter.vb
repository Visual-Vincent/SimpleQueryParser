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
''' A class holding information about a parameter in a SimpleQueryParser query.
''' </summary>
''' <remarks></remarks>
<DebuggerDisplay("{Name,nq}{If(MaxLength > 0, "" (max length: "" & MaxLength & "")"", If(ValidValues.Length > 0, "" ("" & ValidValues.Length & "" valid values)"", "" (parameter)"")),nq}")>
Public NotInheritable Class QueryParameter
    Inherits QueryToken

    Private _maxLength As Integer
    Private _validValues As String()

    ''' <summary>
    ''' Gets the maximum allowed length of the parameter's contents (unlimited = 0).
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property MaxLength As Integer
        Get
            Return _maxLength
        End Get
    End Property

    ''' <summary>
    ''' Gets an array of valid values for the parameter (if any).
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property ValidValues As String()
        Get
            Return _validValues
        End Get
    End Property

    ''' <summary>
    ''' Initializes a new instance of the QueryParameter class.
    ''' </summary>
    ''' <param name="Name">The name of the parameter.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal Name As String)
        Me.New(Name, False, {})
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the QueryParameter class.
    ''' </summary>
    ''' <param name="Name">The name of the parameter.</param>
    ''' <param name="MaxLength">The maximum allowed length of the parameter's contents (unlimited = 0).</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal Name As String, ByVal MaxLength As Integer)
        Me.New(Name, False, {})
        _maxLength = MaxLength
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the QueryParameter class.
    ''' </summary>
    ''' <param name="Name">The name of the parameter.</param>
    ''' <param name="CaseSensitive">Whether or not the parameter's contents are case-sensitive (only makes a difference if ValidValues is not empty)</param>
    ''' <param name="ValidValues">An array of valid values for the parameter.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal Name As String, ByVal CaseSensitive As Boolean, ByVal ParamArray ValidValues As String())
        MyBase.New(Name, CaseSensitive)
        If ValidValues Is Nothing Then Throw New ArgumentNullException("ValidValues")

        _validValues = New String(ValidValues.Length - 1) {}
        _maxLength = 0

        Array.Copy(ValidValues, 0, _validValues, 0, ValidValues.Length)
    End Sub

    ''' <summary>
    ''' Returns the parameter's syntax in human-readable form.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Function GetSyntax() As String
        Return "<" & Me.Name & ">"
    End Function
End Class
