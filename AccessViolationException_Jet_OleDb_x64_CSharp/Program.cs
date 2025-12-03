using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

namespace AccessViolationException_Jet_OleDb_x64_CSharp
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
            //
            // Use either `Microsoft.ACE.OLEDB.16.0` or `Microsoft.ACE.OLEDB.12.0`:
            //   - Microsoft.ACE.OLEDB.16.0: https://www.microsoft.com/en-us/download/details.aspx?id=54920
            //   - Microsoft.ACE.OLEDB.12.0: https://www.microsoft.com/en-us/download/details.aspx?id=13255
            //

            var msdaInitializeType = Type.GetTypeFromCLSID(NativeCom.CLSID_MSDAINITIALIZE, true);
            var msdaInitialize = (NativeCom.IDataInitialize) Activator.CreateInstance(
                msdaInitializeType,
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                null,
                null,
                CultureInfo.InvariantCulture,
                null);
            if (msdaInitialize != null)
            {
                try
                {
                    const string connectionString = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=Northwind.accdb";

                    NativeCom.IDBInitialize dbInitialize = null;
                    dbInitialize = msdaInitialize.GetDataSource(
                        IntPtr.Zero,
                        (int) NativeCom.tagCLSCTX.CLSCTX_INPROC_SERVER,
                        connectionString,
                        typeof(NativeCom.IDBInitialize).GUID);

                    try
                    {
                        dbInitialize.Initialize();

                        try
                        {
                            var dbCreateSession = (NativeCom.IDBCreateSession) dbInitialize;

                            try
                            {
                                var session = dbCreateSession.CreateSession(IntPtr.Zero, NativeCom.IID_IUnknown);

                                try
                                {
                                    var dbCreateCommand = (NativeCom.IDBCreateCommand) session;

                                    try
                                    {
                                        for (var i = 0; i < 100; i++)
                                        {
                                            Console.WriteLine($"{i:000}");
                                            
                                            //
                                            // Select_Union:
                                            //

                                            var commandText = @"SELECT [c].[Address]
FROM [Customers] AS [c]
WHERE [c].[City] = 'Berlin'
UNION
SELECT [c0].[Address]
FROM [Customers] AS [c0]
WHERE [c0].[City] = 'London'";

                                            ExecuteCommand(dbCreateCommand, commandText);
                                            
                                            //
                                            // Select_bool_closure:
                                            //

                                            commandText = @"SELECT 1
FROM [Customers] AS [c]";
                                            
                                            ExecuteCommand(dbCreateCommand, commandText);
                                        }
                                    }
                                    finally
                                    {
                                        dbCreateCommand = null;
                                    }
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(session);
                                    session = null;
                                }
                            }
                            finally
                            {
                                dbCreateSession = null;
                            }
                        }
                        finally
                        {
                            dbInitialize.Uninitialize();
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(dbInitialize);
                        dbInitialize = null;
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(msdaInitialize);
                    msdaInitialize = null;
                }
            }
        }
        
        private static void ExecuteCommand(NativeCom.IDBCreateCommand dbCreateCommand, string sql)
        {
            var commandText = (NativeCom.ICommandText) dbCreateCommand.CreateCommand(
                IntPtr.Zero,
                typeof(NativeCom.ICommandText).GUID);

            try
            {
                commandText.SetCommandText(NativeCom.DBGUID_DEFAULT, sql);
            
                var rowset = commandText.Execute(
                    IntPtr.Zero,
                    typeof(NativeCom.IRowset).GUID,
                    null,
                    out var rowsAffected);

                try
                {

                }
                finally
                {
                    Marshal.ReleaseComObject(rowset);
                    rowset = null;
                }
            }
            finally
            {
                Marshal.ReleaseComObject(commandText);
                commandText = null;
            }
        }
    }
}
