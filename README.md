# Memory Access Violation Exception using the x64 OLE DB provider for Jet/ACE

This repository contains projects that reproduce a non-deterministic memory access violation when using the x64 OLE DB provider for Jet/ACE.

## How to reproduce

Use a vanilla Windows 11/10 x64 installation as the base image.

Install the [Microsoft Access Database Engine 2016 Redistributable](https://www.microsoft.com/en-us/download/details.aspx?id=54920) x64.

Install the [.NET 10 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/thank-you/sdk-10.0.100-windows-x64-installer).

Install [Git for Windows](https://github.com/git-for-windows/git/releases/download/v2.52.0.windows.1/Git-2.52.0-64-bit.exe).

Run the following PowerShell commands:
```powershell
mkdir 'C:\Repositories'
cd 'C:\Repositories'
git clone 'https://github.com/lauxjpn/EFCore.Jet.AccessViolationExceptionJet.git'
cd 'EFCore.Jet.AccessViolationExceptionJet'
dotnet --info
dotnet build '.\AccessViolationException_Jet_OleDb_x64\AccessViolationException_Jet_OleDb_x64.csproj' -c Debug
dotnet build '.\AccessViolationException_Jet_OleDb_x64_CSharp\AccessViolationException_Jet_OleDb_x64_CSharp.csproj' -c Debug
```

Now run one of the compiled projects a couple of times. It should result in a memory violation in at least 1 out of 5 runs. The following code runs the projects until they fail:
```powershell
# Run the projects that uses System.Data.OleDb:
while ($true) { .\AccessViolationException_Jet_OleDb_x64\bin\Debug\net10.0-windows\AccessViolationException_Jet_OleDb_x64.exe; if ($LASTEXITCODE -ne 0) { break; } }

# Run the projects that directly calls the OLE DB objects/interfaces:
while ($true) { .\AccessViolationException_Jet_OleDb_x64_CSharp\bin\Debug\net10.0-windows\AccessViolationException_Jet_OleDb_x64_CSharp.exe; if ($LASTEXITCODE -ne 0) { break; } }
```

## Example Runs

The following lists outputs of different configurations that all reproduced the issue (run on 2025-12-02).

<details>
<summary>Run: Windows 11 x64, .NET 10.0.0 SDK, AccessViolationException_Jet_OleDb_x64, System.Data.OleDb 5.0.0, Debug build, .NET 3.1.32 Runtime, Microsoft Access Database Engine 2016 Redistributable x64, Microsoft.ACE.OLEDB.12.0</summary>

```
PS C:\Repositories\EFCore.Jet.AccessViolationExceptionJet> .\AccessViolationException_Jet_OleDb_x64\bin\Debug\netcoreapp3.1\AccessViolationException_Jet_OleDb_x64.exe
Running as x64 process.

000
001
002
003
004
005
006
007
008
009
010
011
012
013
014
015
016
017
018
019
020
021
022
023
024
025
026
027
028
029
030
031
032
033
034
035
036
037
038
039
040
041
042
043
044
045
046
047
048
049
050
051
052
053
054
055
056
057
058
059
060
061
062
063
064
065
066
067
068
069
070
071
072
073
074
075
076
077
078
079
080
081
082
083
084
085
086
087
088
089
090
091
092
093
094
095
096
097
098
099
PS C:\Repositories\EFCore.Jet.AccessViolationExceptionJet> .\AccessViolationException_Jet_OleDb_x64\bin\Debug\netcoreapp3.1\AccessViolationException_Jet_OleDb_x64.exe
Running as x64 process.

000
001
002
Fatal error. System.AccessViolationException: Attempted to read or write protected memory. This is often an indication that other memory is corrupt.
   at System.Data.Common.UnsafeNativeMethods+ICommandText.Execute(IntPtr, System.Guid ByRef, System.Data.OleDb.tagDBPARAMS, IntPtr ByRef, System.Object ByRef)
   at System.Data.Common.UnsafeNativeMethods+ICommandText.Execute(IntPtr, System.Guid ByRef, System.Data.OleDb.tagDBPARAMS, IntPtr ByRef, System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteCommandTextForSingleResult(System.Data.OleDb.tagDBPARAMS, System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteCommandText(System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteCommand(System.Data.CommandBehavior, System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteReaderInternal(System.Data.CommandBehavior, System.String)
   at System.Data.OleDb.OleDbCommand.ExecuteReader(System.Data.CommandBehavior)
   at System.Data.OleDb.OleDbCommand.ExecuteReader()
   at AccessViolationException_Jet_OleDb_x64.Program.RunNorthwindTest(Boolean)
   at AccessViolationException_Jet_OleDb_x64.Program.Main(System.String[])
```

</details>
<details>
<summary>Run: Windows 11 x64, .NET 10.0.0 SDK, AccessViolationException_Jet_OleDb_x64, System.Data.OleDb 5.0.0, Debug build, .NET 3.1.32 Runtime, Microsoft Access Database Engine 2016 Redistributable x64, Microsoft.ACE.OLEDB.16.0</summary>  

```
PS C:\Repositories\EFCore.Jet.AccessViolationExceptionJet> .\AccessViolationException_Jet_OleDb_x64\bin\Debug\netcoreapp3.1\AccessViolationException_Jet_OleDb_x64.exe
Running as x64 process.

000
001
002
003
004
005
006
007
008
009
010
011
012
013
014
015
016
017
018
019
020
021
022
023
024
025
026
027
028
029
030
031
032
033
034
035
036
037
038
039
Fatal error. System.AccessViolationException: Attempted to read or write protected memory. This is often an indication that other memory is corrupt.
   at System.Data.Common.UnsafeNativeMethods+ICommandText.Execute(IntPtr, System.Guid ByRef, System.Data.OleDb.tagDBPARAMS, IntPtr ByRef, System.Object ByRef)
   at System.Data.Common.UnsafeNativeMethods+ICommandText.Execute(IntPtr, System.Guid ByRef, System.Data.OleDb.tagDBPARAMS, IntPtr ByRef, System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteCommandTextForSingleResult(System.Data.OleDb.tagDBPARAMS, System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteCommandText(System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteCommand(System.Data.CommandBehavior, System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteReaderInternal(System.Data.CommandBehavior, System.String)
   at System.Data.OleDb.OleDbCommand.ExecuteReader(System.Data.CommandBehavior)
   at System.Data.OleDb.OleDbCommand.ExecuteReader()
   at AccessViolationException_Jet_OleDb_x64.Program.RunNorthwindTest(Boolean)
   at AccessViolationException_Jet_OleDb_x64.Program.Main(System.String[])
```

</details>
<details>
<summary>Run: Windows 11 x64, .NET 10.0.0 SDK, AccessViolationException_Jet_OleDb_x64, System.Data.OleDb 10.0.0, Debug build, .NET 10.0.0 Runtime, Microsoft Access Database Engine 2016 Redistributable x64, Microsoft.ACE.OLEDB.16.0</summary>

```
PS C:\Repositories\EFCore.Jet.AccessViolationExceptionJet> .\AccessViolationException_Jet_OleDb_x64\bin\Debug\net10.0-windows\AccessViolationException_Jet_OleDb_x64.exe
Running as x64 process.

000
001
002
003
004
005
006
007
008
009
010
011
012
013
014
015
016
017
018
019
020
021
022
023
024
025
026
027
028
029
030
031
032
033
034
035
036
037
038
039
040
041
042
043
044
045
046
047
048
049
Fatal error.
0xC0000005
   at System.Data.Common.UnsafeNativeMethods+ICommandText.Execute(IntPtr, System.Guid ByRef, System.Data.OleDb.tagDBPARAMS, IntPtr ByRef, System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteCommandTextForSingleResult(System.Data.OleDb.tagDBPARAMS, System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteCommandText(System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteCommand(System.Data.CommandBehavior, System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteReaderInternal(System.Data.CommandBehavior, System.String)
   at System.Data.OleDb.OleDbCommand.ExecuteReader(System.Data.CommandBehavior)
   at System.Data.OleDb.OleDbCommand.ExecuteReader()
   at AccessViolationException_Jet_OleDb_x64.Program.RunNorthwindTest(Boolean)
   at AccessViolationException_Jet_OleDb_x64.Program.Main(System.String[])
```

</details>
<details>
<summary>Run: Windows 11 x64, .NET 10.0.0 SDK, AccessViolationException_Jet_OleDb_x64, System.Data.OleDb 10.0.0, Release build, .NET 10.0.0 Runtime, Microsoft Access Database Engine 2016 Redistributable x64, Microsoft.ACE.OLEDB.16.0</summary>

```
PS C:\Repositories\EFCore.Jet.AccessViolationExceptionJet> .\AccessViolationException_Jet_OleDb_x64\bin\Release\net10.0-windows\AccessViolationException_Jet_OleDb_x64.exe
Running as x64 process.

000
001
002
003
004
005
006
007
008
009
010
011
012
013
014
015
016
017
018
019
020
021
022
023
024
025
026
027
028
029
030
031
032
033
034
035
036
037
038
039
040
041
042
043
Fatal error.
0xC0000005
   at System.Data.Common.UnsafeNativeMethods+ICommandText.Execute(IntPtr, System.Guid ByRef, System.Data.OleDb.tagDBPARAMS, IntPtr ByRef, System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteCommandTextForSingleResult(System.Data.OleDb.tagDBPARAMS, System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteCommandText(System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteCommand(System.Data.CommandBehavior, System.Object ByRef)
   at System.Data.OleDb.OleDbCommand.ExecuteReaderInternal(System.Data.CommandBehavior, System.String)
   at System.Data.OleDb.OleDbCommand.ExecuteReader(System.Data.CommandBehavior)
   at System.Data.OleDb.OleDbCommand.ExecuteReader()
   at AccessViolationException_Jet_OleDb_x64.Program.RunNorthwindTest(Boolean)
   at AccessViolationException_Jet_OleDb_x64.Program.Main(System.String[])
```

</details>
<details>
<summary>Run: Windows 11 x64, .NET 10.0.0 SDK, AccessViolationException_Jet_OleDb_x64_CSharp, Debug build, .NET 10.0.0 Runtime, Microsoft Access Database Engine 2016 Redistributable x64, Microsoft.ACE.OLEDB.16.0</summary>

```
PS C:\Repositories\EFCore.Jet.AccessViolationExceptionJet> .\AccessViolationException_Jet_OleDb_x64_CSharp\bin\Debug\net10.0-windows\AccessViolationException_Jet_OleDb_x64_CSharp.exe
Running as x64 process.

000
001
002
003
004
005
006
007
008
009
010
011
012
013
014
015
016
017
018
019
020
021
022
023
024
025
026
027
028
029
030
031
032
033
034
035
036
037
038
039
040
041
042
043
044
045
046
047
048
049
050
051
052
053
054
055
056
057
058
059
060
061
062
063
064
065
066
067
068
069
070
071
072
073
074
075
076
077
078
079
080
081
082
083
084
085
086
087
088
089
090
091
092
093
094
095
096
097
098
099
PS C:\Repositories\EFCore.Jet.AccessViolationExceptionJet> .\AccessViolationException_Jet_OleDb_x64_CSharp\bin\Debug\net10.0-windows\AccessViolationException_Jet_OleDb_x64_CSharp.exe
Running as x64 process.

000
001
002
003
004
005
006
007
008
009
010
011
012
013
014
015
016
017
018
019
020
021
022
023
024
025
026
027
028
029
030
031
032
033
034
035
036
037
038
039
040
041
042
043
044
045
046
047
048
049
050
051
052
053
054
055
056
057
058
059
060
061
062
063
064
065
066
067
068
069
070
071
072
073
074
075
076
077
078
079
080
081
082
083
084
085
086
087
088
089
090
091
092
093
094
095
096
097
098
099
PS C:\Repositories\EFCore.Jet.AccessViolationExceptionJet> .\AccessViolationException_Jet_OleDb_x64_CSharp\bin\Debug\net10.0-windows\AccessViolationException_Jet_OleDb_x64_CSharp.exe
Running as x64 process.

000
001
002
003
004
005
006
007
008
009
010
011
012
013
014
015
016
017
018
019
020
021
022
023
024
025
026
027
028
029
030
031
032
033
034
035
036
037
038
039
040
041
042
043
044
045
046
047
048
Fatal error.
0xC0000005
   at AccessViolationException_Jet_OleDb_x64_CSharp.NativeCom+ICommandText.Execute(IntPtr, System.Guid, tagDBPARAMS, IntPtr ByRef)
   at AccessViolationException_Jet_OleDb_x64_CSharp.Program.ExecuteCommand(IDBCreateCommand, System.String)
   at AccessViolationException_Jet_OleDb_x64_CSharp.Program.RunNorthwindTest(Boolean)
   at AccessViolationException_Jet_OleDb_x64_CSharp.Program.Main(System.String[])
```

</details>
