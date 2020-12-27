using System;
using System.Data.OleDb;
using System.Diagnostics;

namespace AccessViolationException_Jet_OleDb_x64
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            //
            // AccessViolationException_Jet_OleDb_x64 will result in an AccessViolationException in about 1 out of 5 runs:
            //

            Console.WriteLine($"Running as {(Environment.Is64BitProcess ? "x64" : "x86")} process.");
            Console.WriteLine();
            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.25));

            if (Environment.Is64BitProcess)
            {
                RunNorthwindTest(args.Length == 1 && string.Equals(args[0], "wait", StringComparison.OrdinalIgnoreCase));
            }
        }

        private static void RunNorthwindTest(bool waitForDebuggerAttachment)
        {
            try
            {
                //
                // Use either `Microsoft.ACE.OLEDB.16.0` or `Microsoft.ACE.OLEDB.12.0`:
                //   - Microsoft.ACE.OLEDB.16.0: https://www.microsoft.com/en-us/download/details.aspx?id=54920
                //   - Microsoft.ACE.OLEDB.12.0: https://www.microsoft.com/en-us/download/details.aspx?id=13255
                //

                using var connection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Northwind.accdb");
                connection.Open();

                for (var i = 0; i < 100; i++)
                {
                    Console.WriteLine($"{i:000}");

                    if (i == 10 &&
                        waitForDebuggerAttachment)
                    {
                        while (!Debugger.IsAttached &&
                               Console.ReadKey(true).Key != ConsoleKey.Enter)
                        {
                            System.Threading.Thread.Sleep(500);
                        }
                    }

                    //
                    // Select_Union:
                    //

                    using (var command1 = connection.CreateCommand())
                    {
                        command1.CommandText = @"SELECT [c].[Address]
FROM [Customers] AS [c]
WHERE [c].[City] = 'Berlin'
UNION
SELECT [c0].[Address]
FROM [Customers] AS [c0]
WHERE [c0].[City] = 'London'";

                        using var dataReader1 = command1.ExecuteReader();
                    }

                    // You can run any amount of queries in between.
                    /*
                    using (var command15 = connection.CreateCommand())
                    {
                        command15.CommandText = @"SELECT [c].[Address]
FROM [Customers] AS [c]
WHERE [c].[City] = 'Madrid'";

                        using var dataReader15 = command15.ExecuteReader();
                    }
                    */

                    //
                    // Select_bool_closure:
                    //

                    using (var command2 = connection.CreateCommand())
                    {
                        command2.CommandText = @"SELECT 1
FROM [Customers] AS [c]";

                        using var dataReader2 = command2.ExecuteReader();
                    }
                }
            }
            catch (AccessViolationException e)
            {
                Console.WriteLine(e);
            }
        }
    }
}
