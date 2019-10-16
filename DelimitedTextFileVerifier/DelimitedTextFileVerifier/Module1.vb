Imports System.Text
Imports System.IO

Module Module1

    Sub Main()

        Dim DTF_InputFileName As String = " "
        Dim InputFileFormat As String = " "
        Dim NumberOfDelimiters As Integer = 999
        Dim InputAcquired As Boolean = False

        AcquireConsoleInput(DTF_InputFileName, InputFileFormat, NumberOfDelimiters, InputAcquired)

        If Not InputAcquired Then
            Environment.Exit(0)
        End If

        Dim InputRecords As String()

        Try
            InputRecords = System.IO.File.ReadAllLines(DTF_InputFileName)
        Catch
            Console.WriteLine("There has been an error in opening the input file path name that you supplied.")
            Console.WriteLine("Press <Enter> to exit, then run the program again with the correct file name.")
            Console.ReadLine()
            Environment.Exit(0)
        End Try


        Dim CorrectFileNameAndLoc As String = "c:\reports\DTFV_Correct.txt"
        Dim CorrectFS As New FileStream(CorrectFileNameAndLoc, FileMode.Create, FileAccess.Write)
        Dim CorrectSW As New StreamWriter(CorrectFS)

        Dim IncorrectFileNameAndLoc As String = "c:\reports\DTFV_Incorrect.txt"
        Dim IncorrectFS As New FileStream(IncorrectFileNameAndLoc, FileMode.Create, FileAccess.Write)
        Dim IncorrectSW As New StreamWriter(IncorrectFS)

        Dim CurrentDelimiters As Integer = 0
        Dim CurrentCount As Integer = 0

        For Each DTF_Record In InputRecords
            CurrentCount = CurrentCount + 1

            If CurrentCount > 1 Then
                If InputFileFormat.ToUpper() = "CSV" Then
                    CurrentDelimiters = DTF_Record.Length - DTF_Record.Replace(",", "").Length
                Else
                    CurrentDelimiters = DTF_Record.Length - DTF_Record.Replace(vbTab, "").Length
                End If

                If CurrentDelimiters = NumberOfDelimiters Then
                    CorrectSW.WriteLine(DTF_Record)
                    CorrectSW.Flush()
                Else
                    IncorrectSW.WriteLine(DTF_Record)
                    IncorrectSW.Flush()
                End If
            End If
        Next

        CorrectSW.Close()
        CorrectFS.Close()
        IncorrectSW.Close()
        IncorrectFS.Close()

    End Sub

    Private Sub AcquireConsoleInput(ByRef DTF_InputFileName As String, ByRef InputFileFormat As String, ByRef NumberOfDelimiters As Integer, ByRef InputAcquired As Boolean)

        InputAcquired = False

        Dim ConsoleInput As String = "ready"

        While ConsoleInput.ToUpper <> "EXIT"
            Console.WriteLine("Welcome to the Delimited Text File Verification program.  Result files will be written to the C:\Reports folder.")
            Console.WriteLine("The result file names will be DTFV_Correct.txt and DTFV_Incorrect.txt.")
            Console.WriteLine("----")
            Console.WriteLine("Please enter full path name of input file for delimited text file verification, OR enter 'Exit' to quit.")
            ConsoleInput = Console.ReadLine().Trim
            DTF_InputFileName = ConsoleInput
            If ConsoleInput.ToUpper = "EXIT" Then
                Exit Sub
            End If
            ConsoleInput = "EXIT"
        End While

        ConsoleInput = "ready"

        While ConsoleInput.ToUpper <> "EXIT"
            Console.WriteLine("----")
            Console.WriteLine("Please enter 'CSV' (comma separated values) or 'TSV' (tab separated values) for the file format. OR 'Exit' to quit.")
            ConsoleInput = Console.ReadLine().Trim

            If ConsoleInput.ToUpper = "EXIT" Then
                Exit Sub
            End If

            If ConsoleInput.ToUpper() = "CSV" Or ConsoleInput.ToUpper() = "TSV" Then
                InputFileFormat = ConsoleInput
                ConsoleInput = "EXIT"
            Else
                Console.WriteLine("You did not enter CSV or TSV.  Please try again.")
                ConsoleInput = "ready"
            End If

        End While


        ConsoleInput = "ready"

        While ConsoleInput.ToUpper <> "EXIT"
            Console.WriteLine("----")
            Console.WriteLine("How many fields should each record contain? OR enter 'Exit' to quit.")
            ConsoleInput = Console.ReadLine().Trim

            If ConsoleInput.ToUpper = "EXIT" Then
                Exit Sub
            End If

            If Integer.TryParse(ConsoleInput, NumberOfDelimiters) Then
                ConsoleInput = "EXIT"
                NumberOfDelimiters = NumberOfDelimiters - 1
                InputAcquired = True
                Exit Sub
            Else
                Console.WriteLine("You did not enter a valid whole number.  Please try again.")
                ConsoleInput = "ready"
            End If

        End While

    End Sub


End Module
