using System;
using System.Runtime.InteropServices;
using System.Security;

namespace AccessViolationException_Jet_OleDb_x64_CSharp
{
    internal static class NativeCom
    {
        // MIDL_INTERFACE("2206CCB1-19C1-11D1-89E0-00C04FD7A829")
        // IDataInitialize : public IUnknown
        // {
        // public:
        //     virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE GetDataSource( 
        //         /* [in] */ __RPC__in_opt IUnknown *pUnkOuter,
        //         /* [in] */ DWORD dwClsCtx,
        //         /* [unique][in] */ __RPC__in_opt LPCOLESTR pwszInitializationString,
        //         /* [in] */ __RPC__in REFIID riid,
        //         /* [iid_is][out][in] */ __RPC__deref_inout_opt IUnknown **ppDataSource) = 0;
        //     
        //     virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE GetInitializationString( 
        //         /* [in] */ __RPC__in_opt IUnknown *pDataSource,
        //         /* [in] */ boolean fIncludePassword,
        //         /* [out] */ __RPC__deref_out_opt LPOLESTR *ppwszInitString) = 0;
        //     
        //     virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE CreateDBInstance( 
        //         /* [in] */ __RPC__in REFCLSID clsidProvider,
        //         /* [in] */ __RPC__in_opt IUnknown *pUnkOuter,
        //         /* [in] */ DWORD dwClsCtx,
        //         /* [unique][in] */ __RPC__in_opt LPOLESTR pwszReserved,
        //         /* [in] */ __RPC__in REFIID riid,
        //         /* [iid_is][out] */ __RPC__deref_out_opt IUnknown **ppDataSource) = 0;
        //     
        //     virtual /* [local][helpstring] */ HRESULT STDMETHODCALLTYPE CreateDBInstanceEx( 
        //         /* [in] */ REFCLSID clsidProvider,
        //         /* [annotation][in] */ 
        //         _In_opt_  IUnknown *pUnkOuter,
        //         /* [in] */ DWORD dwClsCtx,
        //         /* [annotation][unique][in] */ 
        //         _In_opt_z_  LPOLESTR pwszReserved,
        //         /* [annotation][unique][in] */ 
        //         _In_  COSERVERINFO *pServerInfo,
        //         /* [in] */ ULONG cmq,
        //         /* [annotation][size_is][out][in] */ 
        //         _Out_writes_(cmq)  MULTI_QI *rgmqResults) = 0;
        //     
        //     virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE LoadStringFromStorage( 
        //         /* [unique][in] */ __RPC__in_opt LPCOLESTR pwszFileName,
        //         /* [out] */ __RPC__deref_out_opt LPOLESTR *ppwszInitializationString) = 0;
        //     
        //     virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE WriteStringToStorage( 
        //         /* [unique][in] */ __RPC__in_opt LPCOLESTR pwszFileName,
        //         /* [unique][in] */ __RPC__in_opt LPCOLESTR pwszInitializationString,
        //         /* [in] */ DWORD dwCreationDisposition) = 0;
        //     
        // };
        [Guid("2206ccb1-19c1-11d1-89e0-00c04fd7a829"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
         ComImport,
         SuppressUnmanagedCodeSecurity]
        internal interface IDataInitialize
        {
            // virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE GetDataSource( 
            //         /* [in] */ __RPC__in_opt IUnknown *pUnkOuter,
            //         /* [in] */ DWORD dwClsCtx,
            //         /* [unique][in] */ __RPC__in_opt LPCOLESTR pwszInitializationString,
            //         /* [in] */ __RPC__in REFIID riid,
            //         /* [iid_is][out][in] */ __RPC__deref_inout_opt IUnknown **ppDataSource) = 0;
            [return: MarshalAs(UnmanagedType.Interface)]
            IDBInitialize GetDataSource(
                [In] IntPtr pUnkOuter,
                [In] int dwClsCtx,
                [In] string pwszInitializationString,
                [In] Guid riid/*,
                [Out, MarshalAs(UnmanagedType.Interface)] out IDBInitialize ppDataSource*/);

            // ...
        }

        // MIDL_INTERFACE("0c733a8b-2a1c-11ce-ade5-00aa0044773d")
        // IDBInitialize : public IUnknown
        // {
        // public:
        //     virtual /* [local] */ HRESULT STDMETHODCALLTYPE Initialize( void) = 0;
        //     
        //     virtual /* [local] */ HRESULT STDMETHODCALLTYPE Uninitialize( void) = 0;
        //
        // };
        [Guid("0c733a8b-2a1c-11ce-ade5-00aa0044773d"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
         ComImport,
         SuppressUnmanagedCodeSecurity]
        internal interface IDBInitialize
        {
            //virtual /* [local] */ HRESULT STDMETHODCALLTYPE Initialize( void) = 0;
            void Initialize();
        
            //virtual /* [local] */ HRESULT STDMETHODCALLTYPE Uninitialize( void) = 0;
            void Uninitialize();
        }
        
        // MIDL_INTERFACE("0c733a5d-2a1c-11ce-ade5-00aa0044773d")
        // IDBCreateSession : public IUnknown
        // {
        // public:
        //      virtual /* [local] */ HRESULT STDMETHODCALLTYPE CreateSession( 
        //                 /* [annotation][in] */ 
        //                 _In_opt_  IUnknown *pUnkOuter,
        //                 /* [annotation][in] */ 
        //                 _In_  REFIID riid,
        //                 /* [annotation][iid_is][out] */ 
        //                 _Outptr_  IUnknown **ppDBSession) = 0;
        //     
        // };
        [Guid("0c733a5d-2a1c-11ce-ade5-00aa0044773d"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
         ComImport,
         SuppressUnmanagedCodeSecurity]
        internal interface IDBCreateSession
        {
            // virtual /* [local] */ HRESULT STDMETHODCALLTYPE CreateSession( 
            //            /* [annotation][in] */ 
            //            _In_opt_  IUnknown *pUnkOuter,
            //            /* [annotation][in] */ 
            //            _In_  REFIID riid,
            //            /* [annotation][iid_is][out] */ 
            //            _Outptr_  IUnknown **ppDBSession) = 0;
            [return: MarshalAs(UnmanagedType.Interface)]
            object CreateSession(
                [In] IntPtr pUnkOuter,
                [In] Guid riid);
        }
        
        // MIDL_INTERFACE("0c733a1d-2a1c-11ce-ade5-00aa0044773d")
        // IDBCreateCommand : public IUnknown
        // {
        // public:
        //     virtual /* [local] */ HRESULT STDMETHODCALLTYPE CreateCommand( 
        //                /* [annotation][in] */ 
        //                _In_opt_  IUnknown *pUnkOuter,
        //                /* [in] */ REFIID riid,
        //                /* [annotation][iid_is][out] */ 
        //                _Outptr_  IUnknown **ppCommand) = 0;
        //
        // };
        [Guid("0c733a1d-2a1c-11ce-ade5-00aa0044773d"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
         ComImport,
         SuppressUnmanagedCodeSecurity]
        internal interface IDBCreateCommand
        {
            // virtual /* [local] */ HRESULT STDMETHODCALLTYPE CreateCommand( 
            //            /* [annotation][in] */ 
            //            _In_opt_  IUnknown *pUnkOuter,
            //           /* [in] */ REFIID riid,
            //           /* [annotation][iid_is][out] */ 
            //           _Outptr_  IUnknown **ppCommand) = 0;
            [return: MarshalAs(UnmanagedType.Interface)]
            object CreateCommand(
                [In] IntPtr pUnkOuter,
                [In] Guid riid);
        }
        
        // MIDL_INTERFACE("0c733a63-2a1c-11ce-ade5-00aa0044773d")
        // ICommand : public IUnknown
        // {
        // public:
        //     virtual /* [local] */ HRESULT STDMETHODCALLTYPE Cancel( void) = 0;
        //     
        //     virtual /* [local] */ HRESULT STDMETHODCALLTYPE Execute( 
        //         /* [annotation][in] */ 
        //         _In_opt_  IUnknown *pUnkOuter,
        //         /* [in] */ REFIID riid,
        //         /* [annotation][out][in] */ 
        //         _Inout_opt_  DBPARAMS *pParams,
        //         /* [annotation][out] */ 
        //         _Out_opt_  DBROWCOUNT *pcRowsAffected,
        //         /* [annotation][iid_is][out] */ 
        //         _Outptr_opt_  IUnknown **ppRowset) = 0;
        //     
        //     virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetDBSession( 
        //         /* [in] */ REFIID riid,
        //         /* [annotation][iid_is][out] */ 
        //         _Outptr_result_maybenull_  IUnknown **ppSession) = 0;
        //
        // };
        [Guid("0c733a63-2a1c-11ce-ade5-00aa0044773d"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
         ComImport,
         SuppressUnmanagedCodeSecurity]
        internal interface ICommand
        {
            // virtual /* [local] */ HRESULT STDMETHODCALLTYPE Cancel( void) = 0;
            void Cancel();
            
            // virtual /* [local] */ HRESULT STDMETHODCALLTYPE Execute( 
            //            /* [annotation][in] */ 
            //            _In_opt_  IUnknown *pUnkOuter,
            //            /* [in] */ REFIID riid,
            //            /* [annotation][out][in] */ 
            //            _Inout_opt_  DBPARAMS *pParams,
            //            /* [annotation][out] */ 
            //            _Out_opt_  DBROWCOUNT *pcRowsAffected,
            //            /* [annotation][iid_is][out] */ 
            //            _Outptr_opt_  IUnknown **ppRowset) = 0;
            [return: MarshalAs(UnmanagedType.Interface)]
            object Execute(
                [In] IntPtr pUnkOuter,
                [In] Guid riid,
                [In] tagDBPARAMS pParams,
                [Out] out IntPtr pcRowsAffected);
            
            // virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetDBSession( 
            //            /* [in] */
            //            REFIID riid,
            //            /* [annotation][iid_is][out] */ 
            //            _Outptr_result_maybenull_  IUnknown **ppSession) = 0;
            // [return: MarshalAs(UnmanagedType.Interface)]
            // object GetDBSession(
            //     [In] IntPtr pUnkOuter,
            //     [In] Guid riid);
            
            [Obsolete("not used", true)]
            void GetDBSession(/*deleted parameter signature*/);
        }
        
        // MIDL_INTERFACE("0c733a27-2a1c-11ce-ade5-00aa0044773d")
        // ICommandText : public ICommand
        // {
        // public:
        //     virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetCommandText( 
        //                /* [annotation][out][in] */ 
        //                _Inout_opt_  GUID *pguidDialect,
        //                /* [annotation][out] */ 
        //                _Outptr_  LPOLESTR *ppwszCommand) = 0;
        //     
        //     virtual /* [local] */ HRESULT STDMETHODCALLTYPE SetCommandText( 
        //                /* [in] */ REFGUID rguidDialect,
        //                /* [annotation][unique][in] */ 
        //                _In_opt_z_  LPCOLESTR pwszCommand) = 0;
        //
        // };
        [Guid("0c733a27-2a1c-11ce-ade5-00aa0044773d"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
         ComImport,
         SuppressUnmanagedCodeSecurity]
        internal interface ICommandText //: ICommand*
        {
            // virtual /* [local] */ HRESULT STDMETHODCALLTYPE Cancel( void) = 0;
            void Cancel();
            
            // virtual /* [local] */ HRESULT STDMETHODCALLTYPE Execute( 
            //            /* [annotation][in] */ 
            //            _In_opt_  IUnknown *pUnkOuter,
            //            /* [in] */ REFIID riid,
            //            /* [annotation][out][in] */ 
            //            _Inout_opt_  DBPARAMS *pParams,
            //            /* [annotation][out] */ 
            //            _Out_opt_  DBROWCOUNT *pcRowsAffected,
            //            /* [annotation][iid_is][out] */ 
            //            _Outptr_opt_  IUnknown **ppRowset) = 0;
            [return: MarshalAs(UnmanagedType.Interface)]
            object Execute(
                [In] IntPtr pUnkOuter,
                [In] Guid riid,
                [In] tagDBPARAMS pParams,
                [Out] out IntPtr pcRowsAffected);
            
            // virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetDBSession( 
            //            /* [in] */
            //            REFIID riid,
            //            /* [annotation][iid_is][out] */ 
            //            _Outptr_result_maybenull_  IUnknown **ppSession) = 0;
            // [return: MarshalAs(UnmanagedType.Interface)]
            // object GetDBSession(
            //     [In] IntPtr pUnkOuter,
            //     [In] Guid riid);
            
            [Obsolete("not used", true)]
            void GetDBSession(/*deleted parameter signature*/);

            // virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetCommandText( 
            //            /* [annotation][out][in] */ 
            //            _Inout_opt_  GUID *pguidDialect,
            //            /* [annotation][out] */ 
            //            _Outptr_  LPOLESTR *ppwszCommand) = 0;
            // [return: MarshalAs(UnmanagedType.LPWStr)]
            // string GetCommandText(
            //     [In] ref Guid pguidDialect);
            
            [Obsolete("not used", true)]
            void GetCommandText(/*deleted parameter signature*/);

            // virtual /* [local] */ HRESULT STDMETHODCALLTYPE SetCommandText( 
            //            /* [in] */
            //            REFGUID rguidDialect,
            //            /* [annotation][unique][in] */ 
            //            _In_opt_z_  LPCOLESTR pwszCommand) = 0;
            // void SetCommandText(
            //     [In] Guid rguidDialect,
            //     [In, MarshalAs(UnmanagedType.LPWStr)] string pwszCommand);
            
            void SetCommandText(
                [In] Guid rguidDialect,
                [In, MarshalAs(UnmanagedType.LPWStr)] string pwszCommand);
        }
        
        // MIDL_INTERFACE("0c733a7c-2a1c-11ce-ade5-00aa0044773d")
        // IRowset : public IUnknown
        // {
        // public:
        //     virtual HRESULT STDMETHODCALLTYPE AddRefRows( 
        //         /* [in] */ DBCOUNTITEM cRows,
        //     /* [size_is][in] */ const HROW rghRows[  ],
        //     /* [size_is][out] */ DBREFCOUNT rgRefCounts[  ],
        //         /* [size_is][out] */ DBROWSTATUS rgRowStatus[  ]) = 0;
        //     
        //     virtual HRESULT STDMETHODCALLTYPE GetData( 
        //         /* [in] */ HROW hRow,
        //         /* [in] */ HACCESSOR hAccessor,
        //         /* [out] */ void *pData) = 0;
        //     
        //     virtual HRESULT STDMETHODCALLTYPE GetNextRows( 
        //         /* [in] */ HCHAPTER hReserved,
        //         /* [in] */ DBROWOFFSET lRowsOffset,
        //         /* [in] */ DBROWCOUNT cRows,
        //         /* [out] */ DBCOUNTITEM *pcRowsObtained,
        //         /* [size_is][size_is][out] */ HROW **prghRows) = 0;
        //     
        //     virtual HRESULT STDMETHODCALLTYPE ReleaseRows( 
        //         /* [in] */ DBCOUNTITEM cRows,
        //     /* [size_is][in] */ const HROW rghRows[  ],
        //     /* [size_is][in] */ DBROWOPTIONS rgRowOptions[  ],
        //         /* [size_is][out] */ DBREFCOUNT rgRefCounts[  ],
        //     /* [size_is][out] */ DBROWSTATUS rgRowStatus[  ]) = 0;
        //     
        //     virtual HRESULT STDMETHODCALLTYPE RestartPosition( 
        //         /* [in] */ HCHAPTER hReserved) = 0;
        //     
        // };
        [Guid("0c733a7c-2a1c-11ce-ade5-00aa0044773d"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
         ComImport,
         SuppressUnmanagedCodeSecurity]
        internal interface IRowset
        {
            // ...
        }

        [StructLayout(LayoutKind.Sequential, Pack = 8)] // x64
        internal sealed class tagDBPARAMS
        {
            internal IntPtr pData;
            internal int cParamSets;
            internal IntPtr hAccessor;

            internal tagDBPARAMS()
            {
            }
        }
        
        internal enum tagCLSCTX
        {
            CLSCTX_INPROC_SERVER	= 0x1,
            
            // ...
        };

        internal static Guid CLSID_MSDAINITIALIZE = new Guid("2206CDB0-19C1-11D1-89E0-00C04FD7A829");
        internal static Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");
        internal static Guid DBGUID_DEFAULT = new Guid(0xc8b521fb, 0x5cf3, 0x11ce, 0xad, 0xe5, 0x00, 0xaa, 0x00, 0x44, 0x77, 0x3d);
    }
}