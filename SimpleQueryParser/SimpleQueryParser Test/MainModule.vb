Imports TextParsing

Module MainModule

    'An enumeration of command IDs.
    Public Enum SqlCommands As Integer
        SelectFromTable = 0
        SelectFromTableOrderBy
        AlterTable
        AlterTableDrop
    End Enum

    Public Sub Main()
        'Declare commands.
        Dim Commands As New QueryCommandTable() From {
            {SqlCommands.SelectFromTable,        "SELECT PARAM(column) FROM PARAM(table)"},
            {SqlCommands.SelectFromTableOrderBy, "SELECT PARAM(column) FROM PARAM(table) ORDER BY PARAM(orderby)"},
            {SqlCommands.AlterTable,             "ALTER TABLE PARAM(table) PARAM(altertype, {""ADD"", ""MODIFY""}) PARAM(type, {""COLUMN"", ""INDEX""}) PARAM(name) PARAM(newtype)"},
            {SqlCommands.AlterTableDrop,         "ALTER TABLE PARAM(table) DROP PARAM(type, {""COLUMN"", ""INDEX""}) PARAM(name)"}
        }

        'Create parser.
        Dim SqlParser As New SimpleQueryParser(Commands)
        Dim sw As New Stopwatch

        'Begin program loop.
        While True
            Console.Clear()

            'Output possible commands.
            Console.ForegroundColor = ConsoleColor.Cyan
            Console.WriteLine("Enter any of the following commands:")

            Console.ForegroundColor = ConsoleColor.White
            For Each kvp As KeyValuePair(Of Integer, QueryCommand) In Commands
                Console.Write("    ")
                Console.WriteLine(kvp.Value.GetSyntax())
            Next

            'Retrieve input.
            Console.ResetColor()
            Console.WriteLine()
            Console.Write("> ")

            Dim Input As String = Console.ReadLine()
            Dim Result As QueryParseResult

            'Parse input.
            sw.Restart()
            Result = SqlParser.Parse(Input)
            sw.Stop()

            Console.Clear()

            If Result.Success Then
                'Parsed successfully.
                Console.ForegroundColor = ConsoleColor.Green
                Console.WriteLine("Parsing was successful!")

                Console.ForegroundColor = ConsoleColor.White
                Console.WriteLine("You entered command: " & CType(Result.CommandID, SqlCommands).ToString() & " (" & Result.CommandID & ")")

                Console.WriteLine()

                'Output parsed parameters.
                For Each Parameter As KeyValuePair(Of String, String) In Result.Parameters
                    Console.WriteLine(Parameter.Key & " = " & Parameter.Value)
                Next
            Else
                'An error occurred.
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("Parse failed:")
                Console.WriteLine(Result.Error)
            End If

            'Output execution time.
            Console.ResetColor()
            Console.WriteLine()
            Console.WriteLine("Execution time: " & sw.Elapsed.ToString())
            Console.WriteLine()

            Console.Write("Press ENTER to continue...")

            'Wait until the user presses ENTER.
            Dim KeyInfo As ConsoleKeyInfo = Console.ReadKey(True)
            While KeyInfo.Key <> ConsoleKey.Enter
                KeyInfo = Console.ReadKey(True)
            End While
        End While
    End Sub

End Module
