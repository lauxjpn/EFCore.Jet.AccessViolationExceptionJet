using System;
using System.Data.Odbc;

namespace Working_Jet_Odbc_x64
{
    internal static class Program
    {
        private static void Main()
        {
            //
            // Working_Jet_Odbc_x64 will always work and never result in an AccessViolationException:
            //

            Console.WriteLine($"Running as {(Environment.Is64BitProcess ? "x64" : "x86")} process.");

            if (Environment.Is64BitProcess)
            {
                Console.WriteLine();
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));

                RunNorthwindTest();
            }
        }

        private static void RunNorthwindTest()
        {
            try
            {
                //
                // Use either `Microsoft.ACE.OLEDB.16.0` or `Microsoft.ACE.OLEDB.12.0`:
                //   - Microsoft.ACE.OLEDB.16.0: https://www.microsoft.com/en-us/download/details.aspx?id=54920
                //   - Microsoft.ACE.OLEDB.12.0: https://www.microsoft.com/en-us/download/details.aspx?id=13255
                //

                var filePath = System.IO.Path.GetFullPath("Northwind.accdb");
                using var connection = new OdbcConnection($"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={filePath}");
                connection.Open();

                for (var i = 0; i < 1000; i++)
                {
                    Console.WriteLine($"{i:000}");

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

                        using (var dataReader1 = command1.ExecuteReader())
                        {
                            while (dataReader1.Read())
                            {
                            }
                        }
                    }

                    // You can run any amount of queries in between.
                    /*
                    using (var command15 = connection.CreateCommand())
                    {
                        command15.CommandText = @"SELECT [c].[Address]
FROM [Customers] AS [c]
WHERE [c].[City] = 'Madrid'";

                        using (var dataReader15 = command15.ExecuteReader())
                        {
                            while (dataReader15.Read())
                            {
                            }
                        }
                    }
                    */

                    //
                    // Select_bool_closure:
                    //

                    using (var command2 = connection.CreateCommand())
                    {
                        command2.CommandText = @"SELECT 1
FROM [Customers] AS [c]";

                        using (var dataReader2 = command2.ExecuteReader())
                        {
                            while (dataReader2.Read())
                            {
                            }
                        }
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
