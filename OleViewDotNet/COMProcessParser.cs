﻿//    This file is part of OleViewDotNet.
//    Copyright (C) James Forshaw 2014, 2017
//
//    OleViewDotNet is free software: you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation, either version 3 of the License, or
//    (at your option) any later version.
//
//    OleViewDotNet is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with OleViewDotNet.  If not, see <http://www.gnu.org/licenses/>.

using NtApiDotNet;
using NtApiDotNet.Ndr;
using NtApiDotNet.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace OleViewDotNet
{
    internal class COMProcessParserConfig
    {
        public string DbgHelpPath { get; private set; }
        public string SymbolPath { get; private set; }
        public bool ParseStubMethods { get; private set; }
        public bool ResolveMethodNames { get; private set; }
        public bool ParseRegisteredClasses { get; private set; }

        public COMProcessParserConfig(string dbghelp_path, string symbol_path, bool parse_stubs_methods, bool resolve_method_names, bool parse_registered_classes)
        {
            DbgHelpPath = dbghelp_path;
            SymbolPath = symbol_path;
            ParseStubMethods = parse_stubs_methods;
            ResolveMethodNames = resolve_method_names;
            ParseRegisteredClasses = parse_registered_classes;
        }
    }

    internal static class COMProcessParser
    {
        private static T ReadStruct<T>(this NtProcess process, long address) where T : new()
        {
            try
            {
                return process.ReadMemory<T>(address);
            }
            catch (NtException)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("Error reading address {0:X}", address));
                return new T();
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        struct PageEntry
        {
            public IntPtr pNext;
            public int dwFlag;
        };

        interface IPageAllocator
        {
            int Pages { get; }
            int EntrySize { get; }
            int EntriesPerPage { get; }
            IntPtr[] ReadPages(NtProcess handle);

        }

        [StructLayout(LayoutKind.Sequential)]
        struct CInternalPageAllocator : IPageAllocator
        {
            public int _cPages;
            public IntPtr _pPageListStart;
            public IntPtr _pPageListEnd;
            public int _dwFlags;
            public PageEntry _ListHead;
            public IntPtr _cEntries;
            public IntPtr _cbPerEntry;
            public ushort _cEntriesPerPage;
            public IntPtr _pLock;

            int IPageAllocator.Pages
            {
                get
                {
                    return _cPages;
                }
            }

            int IPageAllocator.EntrySize
            {
                get
                {
                    return _cbPerEntry.ToInt32();
                }
            }

            int IPageAllocator.EntriesPerPage
            {
                get
                {
                    return _cEntriesPerPage;
                }
            }

            IntPtr[] IPageAllocator.ReadPages(NtProcess process)
            {
                return process.ReadMemoryArray<IntPtr>(_pPageListStart.ToInt64(), _cPages);
            }
        };

        [StructLayout(LayoutKind.Sequential)]
        struct CPageAllocator
        {
            public CInternalPageAllocator _pgalloc;
            public IntPtr _hHeap;
            public int _cbPerEntry;
            public int _lNumEntries;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct PageEntry32
        {
            public int pNext;
            public int dwFlag;
        };

        [StructLayout(LayoutKind.Sequential)]
        struct CInternalPageAllocator32 : IPageAllocator
        {
            public int _cPages;
            public int _pPageListStart;
            public int _pPageListEnd;
            public int _dwFlags;
            public PageEntry32 _ListHead;
            public int _cEntries;
            public int _cbPerEntry;
            public ushort _cEntriesPerPage;
            public int _pLock;

            int IPageAllocator.Pages
            {
                get
                {
                    return _cPages;
                }
            }

            int IPageAllocator.EntrySize
            {
                get
                {
                    return _cbPerEntry;
                }
            }

            int IPageAllocator.EntriesPerPage
            {
                get
                {
                    return _cEntriesPerPage;
                }
            }
            IntPtr[] IPageAllocator.ReadPages(NtProcess process)
            {
                return process.ReadMemoryArray<int>(_pPageListStart, _cPages).Select(i => new IntPtr(i)).ToArray();
            }
        };

        internal interface IPIDEntryNativeInterface
        {
            uint Flags { get; }
            IntPtr Interface { get; }
            IntPtr Stub { get; }
            Guid Ipid { get; }
            Guid Iid { get; }
            int StrongRefs { get; }
            int WeakRefs { get; }
            int PrivateRefs { get; }
            IOXIDEntry GetOxidEntry(NtProcess process);
        }

        [StructLayout(LayoutKind.Sequential)]
        struct IPIDEntryNative : IPIDEntryNativeInterface
        {
            public IntPtr pNextIPID;
            public uint dwFlags;
            public int cStrongRefs;
            public int cWeakRefs;
            public int cPrivateRefs;
            public IntPtr pv;
            public IntPtr pStub;
            public IntPtr pOXIDEntry;
            public Guid ipid;
            public Guid iid;
            public IntPtr pChnl;
            public IntPtr pIRCEntry;
            public IntPtr pOIDFLink;
            public IntPtr pOIDBLink;

            uint IPIDEntryNativeInterface.Flags
            {
                get
                {
                    return dwFlags;
                }
            }

            IntPtr IPIDEntryNativeInterface.Interface
            {
                get
                {
                    return pv;
                }
            }

            IntPtr IPIDEntryNativeInterface.Stub
            {
                get
                {
                    return pStub;
                }
            }

            Guid IPIDEntryNativeInterface.Ipid
            {
                get
                {
                    return ipid;
                }
            }

            Guid IPIDEntryNativeInterface.Iid
            {
                get
                {
                    return iid;
                }
            }

            int IPIDEntryNativeInterface.StrongRefs
            {
                get
                {
                    return cStrongRefs;
                }
            }

            int IPIDEntryNativeInterface.WeakRefs
            {
                get
                {
                    return cWeakRefs;
                }
            }

            int IPIDEntryNativeInterface.PrivateRefs
            {
                get
                {
                    return cPrivateRefs;
                }
            }

            IOXIDEntry IPIDEntryNativeInterface.GetOxidEntry(NtProcess process)
            {
                return process.ReadStruct<OXIDEntryNative>(pOXIDEntry.ToInt64());
            }
        };

        [StructLayout(LayoutKind.Sequential)]
        struct IPIDEntryNative32 : IPIDEntryNativeInterface
        {
            public int pNextIPID;
            public uint dwFlags;
            public int cStrongRefs;
            public int cWeakRefs;
            public int cPrivateRefs;
            public int pv;
            public int pStub;
            public int pOXIDEntry;
            public Guid ipid;
            public Guid iid;
            public int pChnl;
            public int pIRCEntry;
            public int pOIDFLink;
            public int pOIDBLink;

            uint IPIDEntryNativeInterface.Flags
            {
                get
                {
                    return dwFlags;
                }
            }

            IntPtr IPIDEntryNativeInterface.Interface
            {
                get
                {
                    return new IntPtr(pv);
                }
            }

            IntPtr IPIDEntryNativeInterface.Stub
            {
                get
                {
                    return new IntPtr(pStub);
                }
            }

            Guid IPIDEntryNativeInterface.Ipid
            {
                get
                {
                    return ipid;
                }
            }

            Guid IPIDEntryNativeInterface.Iid
            {
                get
                {
                    return iid;
                }
            }

            int IPIDEntryNativeInterface.StrongRefs
            {
                get
                {
                    return cStrongRefs;
                }
            }

            int IPIDEntryNativeInterface.WeakRefs
            {
                get
                {
                    return cWeakRefs;
                }
            }

            int IPIDEntryNativeInterface.PrivateRefs
            {
                get
                {
                    return cPrivateRefs;
                }
            }

            IOXIDEntry IPIDEntryNativeInterface.GetOxidEntry(NtProcess process)
            {
                return process.ReadStruct<OXIDEntryNative32>(pOXIDEntry);
            }
        };

        [StructLayout(LayoutKind.Sequential)]
        struct COMVERSION
        {
            public ushort MajorVersion;
            public ushort MinorVersion;
        }

        internal interface IOXIDEntry
        {
            int Pid { get; }
            int Tid { get; }
            Guid MOxid { get; }
            long Mid { get; }
            IntPtr ServerSTAHwnd { get; }
        }

        [StructLayout(LayoutKind.Sequential)]
        struct OXIDEntryNative : IOXIDEntry
        {
            public IntPtr _pNext;
            public IntPtr _pPrev;
            public int _dwPid;
            public int _dwTid;
            public Guid _moxid;
            public long _mid;
            public Guid _ipidRundown;
            public int _dwFlags;
            public IntPtr _hServerSTA;
            public IntPtr _pParentApt;
            public IntPtr _pSharedDefaultHandle;
            public IntPtr _pAuthId;
            public IntPtr _pBinding;
            public int _dwAuthnHint;
            public int _dwAuthnSvc;
            public IntPtr _pMIDEntry;
            public IntPtr _pRUSTA;
            public int _cRefs;
            public IntPtr _hComplete;
            public int _cCalls;
            public int _cResolverRef;
            public int _dwExpiredTime;
            COMVERSION _version;
            public IntPtr _pAppContainerServerSecurityDescriptor;
            public int _ulMarshaledTargetInfoLength;
            public IntPtr _pMarshaledTargetInfo;
            public IntPtr _pszServerPackageFullName;
            public Guid _guidProcessIdentifier;

            int IOXIDEntry.Pid
            {
                get
                {
                    return _dwPid;
                }
            }

            int IOXIDEntry.Tid
            {
                get
                {
                    return _dwTid;
                }
            }

            Guid IOXIDEntry.MOxid
            {
                get
                {
                    return _moxid;
                }
            }

            long IOXIDEntry.Mid
            {
                get
                {
                    return _mid;
                }
            }

            IntPtr IOXIDEntry.ServerSTAHwnd
            {
                get
                {
                    return _hServerSTA;
                }
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        struct OXIDEntryNative32 : IOXIDEntry
        {
            public int _pNext;
            public int _pPrev;
            public int _dwPid;
            public int _dwTid;
            public Guid _moxid;
            public long _mid;
            public Guid _ipidRundown;
            public int _dwFlags;
            public int _hServerSTA;
            public int _pParentApt;
            public int _pSharedDefaultHandle;
            public int _pAuthId;
            public int _pBinding;
            public int _dwAuthnHint;
            public int _dwAuthnSvc;
            public int _pMIDEntry;
            public int _pRUSTA;
            public int _cRefs;
            public int _hComplete;
            public int _cCalls;
            public int _cResolverRef;
            public int _dwExpiredTime;
            COMVERSION _version;
            public int _pAppContainerServerSecurityDescriptor;
            public int _ulMarshaledTargetInfoLength;
            public int _pMarshaledTargetInfo;
            public int _pszServerPackageFullName;
            public Guid _guidProcessIdentifier;

            int IOXIDEntry.Pid
            {
                get
                {
                    return _dwPid;
                }
            }

            int IOXIDEntry.Tid
            {
                get
                {
                    return _dwTid;
                }
            }

            Guid IOXIDEntry.MOxid
            {
                get
                {
                    return _moxid;
                }
            }

            long IOXIDEntry.Mid
            {
                get
                {
                    return _mid;
                }
            }

            IntPtr IOXIDEntry.ServerSTAHwnd
            {
                get
                {
                    return new IntPtr(_hServerSTA);
                }
            }
        }

        private class PageAllocator
        {
            public IntPtr[] Pages { get; private set; }
            public int EntrySize { get; private set; }
            public int EntriesPerPage { get; private set; }

            void Init<T>(NtProcess process, IntPtr ipid_table) where T : IPageAllocator, new()
            {
                IPageAllocator page_alloc = process.ReadStruct<T>(ipid_table.ToInt64());
                Pages = page_alloc.ReadPages(process);
                EntrySize = page_alloc.EntrySize;
                EntriesPerPage = page_alloc.EntriesPerPage;
            }

            public PageAllocator(NtProcess process, IntPtr ipid_table)
            {
                if (process.Is64Bit)
                {
                    Init<CInternalPageAllocator>(process, ipid_table);
                }
                else
                {
                    Init<CInternalPageAllocator32>(process, ipid_table);
                }
            }
        }

        interface ICClassEntry
        {
            Guid[] GetGuids();
        }

        [StructLayout(LayoutKind.Sequential)]
        struct CClassEntry : ICClassEntry
        {
            public IntPtr vfptr; // CClassCache::CCollectableVtbl* 
            public IntPtr _pNextCollectee; // CClassCache::CCollectable* 
            public ulong _qwTickLastTouched;

            // SMultiGUIDHashNode _hashNode;
            public IntPtr pNext; // SHashChain* 
            public IntPtr pPrev; // SHashChain* 
            public int cGUID;
            public IntPtr aGUID; // _GUID* 
            // END SMultiGUIDHashNode _hashNode;
            public Guid guids1;
            public Guid guids2;
            public uint _dwSig;
            public uint _dwFlags;
            public IntPtr _pTreatAsList; // CClassCache::CClassEntry* 
            public IntPtr _pBCEListFront; // CClassCache::CBaseClassEntry* 
            public IntPtr _pBCEListBack; // CClassCache::CBaseClassEntry* 
            public uint _cLocks;
            public uint _dwFailedContexts;
            public IntPtr _pCI; // IComClassInfo* 

            Guid[] ICClassEntry.GetGuids()
            {
                Guid[] ret = new Guid[2];
                ret[0] = guids1;
                ret[1] = guids2;
                return ret;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        struct CClassEntry32 : ICClassEntry
        {
            public int vfptr; // CClassCache::CCollectableVtbl* 
            public int _pNextCollectee; // CClassCache::CCollectable* 
            public ulong _qwTickLastTouched;

            // SMultiGUIDHashNode _hashNode;
            public int pNext; // SHashChain* 
            public int pPrev; // SHashChain* 
            public int cGUID;
            public int aGUID; // _GUID* 
            // END SMultiGUIDHashNode _hashNode;
            public Guid guids1;
            public Guid guids2;
            public uint _dwSig;
            public uint _dwFlags;
            public int _pTreatAsList; // CClassCache::CClassEntry* 
            public int _pBCEListFront; // CClassCache::CBaseClassEntry* 
            public int _pBCEListBack; // CClassCache::CBaseClassEntry* 
            public uint _cLocks;
            public uint _dwFailedContexts;
            public int _pCI; // IComClassInfo* 

            Guid[] ICClassEntry.GetGuids()
            {
                Guid[] ret = new Guid[2];
                ret[0] = guids1;
                ret[1] = guids2;
                return ret;
            }
        }

        interface ICLSvrClassEntry
        {
            IntPtr GetNext();
            IntPtr GetIUnknown();
            ICClassEntry GetClassEntry(NtProcess process);
            REGCLS GetRegFlags();
            uint GetCookie();
            CLSCTX GetContext();
        }

        [StructLayout(LayoutKind.Sequential)]
        struct CLSvrClassEntry : ICLSvrClassEntry
        {
            public IntPtr vfptr; // CClassCache::CBaseClassEntryVtbl*
            public IntPtr _pNext; // CClassCache::CBaseClassEntry*
            public IntPtr _pPrev; // CClassCache::CBaseClassEntry* 
            public IntPtr _pClassEntry; // CClassCache::CClassEntry* 
            public CLSCTX _dwContext;
            public int _dwSig;
            public IntPtr _pNextLSvr; // CClassCache::CLSvrClassEntry* 
            public IntPtr _pPrevLSvr; // CClassCache::CLSvrClassEntry*
            public IntPtr _pUnk; // IUnknown* 
            public REGCLS _dwRegFlags;
            public uint _dwFlags;
            public uint _dwScmReg;
            public uint _hApt;
            public IntPtr _hWndDdeServer;
            public IntPtr _pObjServer; // CObjServer*
            public uint _dwCookie;
            public uint _cUsing;
            public uint _ulServiceId;

            ICClassEntry ICLSvrClassEntry.GetClassEntry(NtProcess process)
            {
                if (_pClassEntry == IntPtr.Zero)
                {
                    return null;
                }
                return process.ReadStruct<CClassEntry>(_pClassEntry.ToInt64());
            }

            CLSCTX ICLSvrClassEntry.GetContext()
            {
                return _dwContext;
            }

            uint ICLSvrClassEntry.GetCookie()
            {
                return _dwCookie;
            }

            IntPtr ICLSvrClassEntry.GetIUnknown()
            {
                return _pUnk;
            }

            IntPtr ICLSvrClassEntry.GetNext()
            {
                return _pNextLSvr;
            }

            REGCLS ICLSvrClassEntry.GetRegFlags()
            {
                return _dwRegFlags;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        struct CLSvrClassEntry32 : ICLSvrClassEntry
        {
            public int vfptr; // CClassCache::CBaseClassEntryVtbl*
            public int _pNext; // CClassCache::CBaseClassEntry*
            public int _pPrev; // CClassCache::CBaseClassEntry* 
            public int _pClassEntry; // CClassCache::CClassEntry* 
            public CLSCTX _dwContext;
            public int _dwSig;
            public int _pNextLSvr; // CClassCache::CLSvrClassEntry* 
            public int _pPrevLSvr; // CClassCache::CLSvrClassEntry*
            public int _pUnk; // IUnknown* 
            public REGCLS _dwRegFlags;
            public uint _dwFlags;
            public uint _dwScmReg;
            public uint _hApt;
            public int _hWndDdeServer;
            public int _pObjServer; // CObjServer*
            public uint _dwCookie;
            public uint _cUsing;
            public uint _ulServiceId;

            ICClassEntry ICLSvrClassEntry.GetClassEntry(NtProcess process)
            {
                if (_pClassEntry == 0)
                {
                    return null;
                }
                return process.ReadStruct<CClassEntry32>(_pClassEntry);
            }

            IntPtr ICLSvrClassEntry.GetIUnknown()
            {
                return new IntPtr(_pUnk);
            }

            IntPtr ICLSvrClassEntry.GetNext()
            {
                return new IntPtr(_pNextLSvr);
            }
            uint ICLSvrClassEntry.GetCookie()
            {
                return _dwCookie;
            }

            REGCLS ICLSvrClassEntry.GetRegFlags()
            {
                return _dwRegFlags;
            }

            CLSCTX ICLSvrClassEntry.GetContext()
            {
                return _dwContext;
            }
        }

        static List<COMIPIDEntry> ParseIPIDEntries<T>(NtProcess process, IntPtr ipid_table, ISymbolResolver resolver, COMProcessParserConfig config, COMRegistry registry) 
            where T : struct, IPIDEntryNativeInterface
        {
            List<COMIPIDEntry> entries = new List<COMIPIDEntry>();
            PageAllocator palloc = new PageAllocator(process, ipid_table);
            if (palloc.Pages.Length == 0 || palloc.EntrySize < Marshal.SizeOf(typeof(T)))
            {
                return entries;
            }

            foreach (IntPtr page in palloc.Pages)
            {
                int total_size = palloc.EntriesPerPage * palloc.EntrySize;
                var data = process.ReadMemory(page.ToInt64(), palloc.EntriesPerPage * palloc.EntrySize);
                if (data.Length < total_size)
                {
                    continue;
                }

                using (var buf = new SafeHGlobalBuffer(data))
                {
                    for (int entry_index = 0; entry_index < palloc.EntriesPerPage; ++entry_index)
                    {
                        IPIDEntryNativeInterface ipid_entry = buf.Read<T>((ulong)(entry_index * palloc.EntrySize));
                        if ((ipid_entry.Flags != 0xF1EEF1EE) && (ipid_entry.Flags != 0))
                        {
                            entries.Add(new COMIPIDEntry(ipid_entry, process, resolver, config, registry));
                        }
                    }
                }
            }
            
            return entries;
        }

        static Dictionary<string, IntPtr> _resolved_32bit = new Dictionary<string, IntPtr>();
        static Dictionary<string, IntPtr> _resolved_64bit = new Dictionary<string, IntPtr>();

        static string _dllname = COMUtilities.GetCOMDllName();

        static string GetSymbolName(string name)
        {
            return String.Format("{0}!{1}", _dllname, name);
        }

        internal static IntPtr AddressFromSymbol(ISymbolResolver resolver, bool is64bit, string symbol)
        {
            Dictionary<string, IntPtr> resolved = is64bit ? _resolved_64bit : _resolved_32bit;
            if (resolved.ContainsKey(symbol))
            {
                return resolved[symbol];
            }

            IntPtr ret = resolver.GetAddressOfSymbol(symbol);
            if (ret != IntPtr.Zero)
            {
                resolved[symbol] = ret;
            }

            return ret;
        }

        internal static string SymbolFromAddress(ISymbolResolver resolver, bool is64bit, IntPtr address)
        {
            return String.Format("0x{0:X}", address.ToInt64());
        }

        static List<COMIPIDEntry> ParseIPIDEntries(NtProcess process, ISymbolResolver resolver, COMProcessParserConfig config, COMRegistry registry)
        {
            IntPtr ipid_table = AddressFromSymbol(resolver, process.Is64Bit, GetSymbolName("CIPIDTable::_palloc"));
            if (ipid_table == IntPtr.Zero)
            {
                return new List<COMIPIDEntry>();
            }

            if (process.Is64Bit)
            {
                return ParseIPIDEntries<IPIDEntryNative>(process, ipid_table, resolver, config, registry);
            }
            else
            {
                return ParseIPIDEntries<IPIDEntryNative32>(process, ipid_table, resolver, config, registry);
            }
        }

        private static Guid GetProcessAppId(NtProcess process, ISymbolResolver resolver)
        {
            IntPtr appid = AddressFromSymbol(resolver, process.Is64Bit, GetSymbolName("g_AppId"));
            if (appid == IntPtr.Zero)
            {
                return Guid.Empty;
            }
            return process.ReadStruct<Guid>(appid.ToInt64());
        }

        private static string ReadSecurityDescriptorFromAddress(NtProcess process, IntPtr address)
        {
            try
            {
                return new SecurityDescriptor(process, address).ToSddl();
            }
            catch (NtException)
            {
                return string.Empty;
            }
        }

        private static string ReadSecurityDescriptor(NtProcess process, ISymbolResolver resolver, string symbol)
        {
            IntPtr sd = AddressFromSymbol(resolver, process.Is64Bit, GetSymbolName(symbol));
            if (sd == IntPtr.Zero)
            {
                return String.Empty;
            }
            IntPtr sd_ptr;
            if (process.Is64Bit)
            {
                sd_ptr = process.ReadStruct<IntPtr>(sd.ToInt64());
            }
            else
            {
                sd_ptr = new IntPtr(process.ReadStruct<int>(sd.ToInt64()));
            }

            if (sd_ptr == IntPtr.Zero)
            {
                return "D:NO_ACCESS_CONTROL";
            }

            return ReadSecurityDescriptorFromAddress(process, sd_ptr);
        }

        private static string GetProcessAccessSecurityDescriptor(NtProcess process, ISymbolResolver resolver)
        {
            return ReadSecurityDescriptor(process, resolver, "gSecDesc");
        }

        private static string GetLrpcSecurityDescriptor(NtProcess process, ISymbolResolver resolver)
        {
            return ReadSecurityDescriptor(process, resolver, "gLrpcSecurityDescriptor");
        }

        private static ICLSvrClassEntry ReadCLSvrClassEntry(NtProcess process, IntPtr address)
        {
            return process.Is64Bit ? (ICLSvrClassEntry)process.ReadStruct<CLSvrClassEntry>(address.ToInt64())
                : process.ReadStruct<CLSvrClassEntry32>(address.ToInt64());
        }

        private static void ReadRegisteredClasses(NtProcess process, ISymbolResolver resolver, 
            IntPtr base_address, COMProcessClassApartment apartment, 
            int thread_id, List<COMProcessClassRegistration> classes)
        {
            if (base_address == IntPtr.Zero)
            {
                return;
            }

            IntPtr next = base_address;

            do
            {
                ICLSvrClassEntry entry = ReadCLSvrClassEntry(process, next);
                var class_entry = entry.GetClassEntry(process);
                if (class_entry != null)
                {
                    IntPtr vtable_ptr = ReadPointer(process, entry.GetIUnknown());
                    string vtable = resolver.GetModuleRelativeAddress(vtable_ptr);

                    classes.Add(new COMProcessClassRegistration(class_entry.GetGuids()[0], 
                        entry.GetIUnknown(), vtable,
                        entry.GetRegFlags(), entry.GetCookie(), thread_id, 
                        entry.GetContext(), apartment));
                }

                next = entry.GetNext();
            }
            while (next != base_address);
        }

        private static List<COMProcessClassRegistration> GetRegisteredClasses(NtProcess process, ISymbolResolver resolver, COMProcessParserConfig config)
        {
            var classes = new List<COMProcessClassRegistration>();
            if (!config.ParseRegisteredClasses)
            {
                return classes;
            }
            ReadRegisteredClasses(process, resolver, ReadPointer(process, resolver, "CClassCache::_MTALSvrsFront"), COMProcessClassApartment.MTA, -1, classes);
            ReadRegisteredClasses(process, resolver, ReadPointer(process, resolver, "CClassCache::_NTALSvrsFront"), COMProcessClassApartment.NTA, 0, classes);
            using (var list = process.GetThreads(ThreadAccessRights.QueryLimitedInformation).ToDisposableList())
            {
                foreach (var th in list)
                {
                    IntPtr sta = GetSTALSvrsFront(process, th);
                    if (sta == IntPtr.Zero)
                    {
                        continue;
                    }

                    ReadRegisteredClasses(process, resolver, sta, COMProcessClassApartment.STA, th.ThreadId, classes);
                }
            }
            return classes;
        }

        private static IntPtr GetSTALSvrsFront(NtProcess process, NtThread thread)
        {
            IntPtr p = GetReservedForOle(process, thread);
            if (p == IntPtr.Zero)
            {
                return IntPtr.Zero;
            }

            if (process.Is64Bit)
            {
                return ReadPointer(process, p + 0x118);
            }
            return ReadPointer(process, p + 0xa8);
        }

        private static IntPtr GetReservedForOle(NtProcess process, NtThread thread)
        {
            IntPtr teb = thread.TebBaseAddress;
            if (process.Is64Bit)
            {
                return ReadPointer(process, teb + 0x1758);
            }
            return ReadPointer(process, teb + 0xF80);
        }

        private static string ReadUnicodeString(NtProcess process, IntPtr ptr)
        {
            StringBuilder builder = new StringBuilder();
            int pos = 0;
            do
            {
                byte[] data = process.ReadMemory(ptr.ToInt64() + pos, 2);
                if (data.Length < 2)
                {
                    break;
                }
                char c = BitConverter.ToChar(data, 0);
                if (c == 0)
                {
                    break;
                }
                builder.Append(c);
                pos += 2;
            }
            while (true);
            return builder.ToString();
        }

        private static string ReadString(NtProcess process, ISymbolResolver resolver, string symbol)
        {
            IntPtr str = AddressFromSymbol(resolver, process.Is64Bit, GetSymbolName(symbol));
            if (str != IntPtr.Zero)
            {
                return ReadUnicodeString(process, str);
            }
            return String.Empty;
        }

        public static int ReadInt(NtProcess process, ISymbolResolver resolver, string symbol)
        {
            IntPtr p = AddressFromSymbol(resolver, process.Is64Bit, GetSymbolName(symbol));
            if (p != IntPtr.Zero)
            {
                return process.ReadStruct<int>(p.ToInt64());
            }
            return 0;
        }

        public static T ReadEnum<T>(NtProcess process, ISymbolResolver resolver, string symbol)
        {
            int value = ReadInt(process, resolver, symbol);
            return (T)Enum.ToObject(typeof(T), value);
        }

        public static IntPtr ReadPointer(NtProcess process, ISymbolResolver resolver, string symbol)
        {
            return ReadPointer(process, AddressFromSymbol(resolver, process.Is64Bit, GetSymbolName(symbol)));
        }

        public static IntPtr ReadPointer(NtProcess process, IntPtr p)
        {
            if (p != IntPtr.Zero)
            {
                if (process.Is64Bit)
                {
                    return process.ReadStruct<IntPtr>(p.ToInt64());
                }
                else
                {
                    return new IntPtr(process.ReadStruct<int>(p.ToInt64()));
                }
            }
            return IntPtr.Zero;
        }

        public static IntPtr[] ReadPointerArray(NtProcess process, IntPtr p, int count)
        {
            if (p == IntPtr.Zero)
            {
                return null;
            }
            if (process.Is64Bit)
            {
                return process.ReadMemoryArray<IntPtr>(p.ToInt64(), count);
            }
            else
            {
                var ptrs = process.ReadMemoryArray<int>(p.ToInt64(), count);
                return ptrs.Select(i => new IntPtr(i)).ToArray();
            }
        }

        private static string GetProcessFileName(NtProcess process)
        {
            return process.GetImageFilePath(false);
        }
        
        public static COMProcessEntry ParseProcess(int pid, COMProcessParserConfig config, COMRegistry registry)
        {
            using (var result = NtProcess.Open(pid, ProcessAccessRights.VmRead | ProcessAccessRights.QueryInformation, false))
            {
                if (!result.IsSuccess)
                {
                    return null;
                }

                NtProcess process = result.Result;

                if (process.Is64Bit && !Environment.Is64BitProcess)
                {
                    return null;
                }

                using (ISymbolResolver resolver = SymbolResolver.Create(process, config.DbgHelpPath, config.SymbolPath))
                {
                    Sid user = process.User;
                    return new COMProcessEntry(
                        pid,
                        GetProcessFileName(process),
                        ParseIPIDEntries(process, resolver, config, registry),
                        process.Is64Bit,
                        GetProcessAppId(process, resolver),
                        GetProcessAccessSecurityDescriptor(process, resolver),
                        GetLrpcSecurityDescriptor(process, resolver),
                        user.Name,
                        user.ToString(),
                        ReadString(process, resolver, "gwszLRPCEndPoint"),
                        ReadEnum<EOLE_AUTHENTICATION_CAPABILITIES>(process, resolver, "gCapabilities"),
                        ReadEnum<RPC_AUTHN_LEVEL>(process, resolver, "gAuthnLevel"),
                        ReadEnum<RPC_IMP_LEVEL>(process, resolver, "gImpLevel"),
                        ReadPointer(process, resolver, "gAccessControl"),
                        ReadPointer(process, resolver, "ghwndOleMainThread"),
                        GetRegisteredClasses(process, resolver, config));
                }
            }
        }

        public static IEnumerable<COMProcessEntry> GetProcesses(IEnumerable<Process> procs, COMProcessParserConfig config, IProgress<Tuple<string, int>> progress, COMRegistry registry)
        {
            List<COMProcessEntry> ret = new List<COMProcessEntry>();
            NtToken.EnableDebugPrivilege();
            int total_count = procs.Count();
            int current_count = 0;
            foreach (Process p in procs)
            {
                try
                {
                    if (progress != null)
                    {
                        progress.Report(new Tuple<string, int>(String.Format("Parsing process {0}", p.ProcessName),
                            100 * current_count++ / total_count));
                    }
                    COMProcessEntry proc = COMProcessParser.ParseProcess(p.Id,
                        config, registry);
                    if (proc != null)
                    {
                        ret.Add(proc);
                    }
                }
                catch (Win32Exception)
                {
                }
                finally
                {
                    p.Close();
                }
            }

            return ret;
        }
    }

    public enum COMProcessClassApartment
    {
        STA,
        MTA,
        NTA
    }

    public class COMProcessClassRegistration
    {
        public Guid Clsid { get; private set; }
        public IntPtr ClassFactory { get; private set; }
        public string VTable { get; private set; }
        public COMProcessClassApartment Apartment { get; private set; }
        public REGCLS RegFlags { get; private set; }
        public uint Cookie { get; private set; }
        public int ThreadId { get; private set; }
        public CLSCTX Context { get; private set; }
        internal COMProcessClassRegistration(
            Guid clsid, IntPtr class_factory, string vtable,
            REGCLS regflags, uint cookie, int thread_id, 
            CLSCTX context, COMProcessClassApartment apartment)
        {
            Clsid = clsid;
            ClassFactory = class_factory;
            VTable = vtable;
            Apartment = apartment;
            RegFlags = regflags;
            Cookie = cookie;
            ThreadId = thread_id;
            Context = context;
        }
    }

    public class COMProcessEntry
    {
        public int Pid { get; private set; }
        public string ExecutablePath { get; private set; }
        public string Name
        {
            get
            {
                return Path.GetFileNameWithoutExtension(ExecutablePath);
            }
        }
        public IEnumerable<COMIPIDEntry> Ipids { get; private set; }
        public bool Is64Bit { get; private set; }
        public Guid AppId { get; private set; }
        public string AccessPermissions { get; private set; }
        public string LRpcPermissions { get; private set; }
        public string User { get; private set; }
        public string UserSid { get; private set; }
        public string RpcEndpoint { get; private set; }
        public EOLE_AUTHENTICATION_CAPABILITIES Capabilities { get; private set; }
        public RPC_AUTHN_LEVEL AuthnLevel { get; private set; }
        public RPC_IMP_LEVEL ImpLevel { get; private set; }
        public IntPtr AccessControl { get; private set; }
        public IntPtr STAMainHWnd { get; private set; }
        public IEnumerable<COMProcessClassRegistration> Classes { get; private set; } 

        internal COMProcessEntry(int pid, string path, List<COMIPIDEntry> ipids, 
            bool is64bit, Guid appid, string access_perm, string lrpc_perm, string user,
            string user_sid, string rpc_endpoint, EOLE_AUTHENTICATION_CAPABILITIES capabilities,
            RPC_AUTHN_LEVEL authn_level, RPC_IMP_LEVEL imp_level,
            IntPtr access_control, IntPtr sta_main_hwnd, List<COMProcessClassRegistration> classes)
        {
            Pid = pid;
            ExecutablePath = path;
            Ipids = ipids.AsReadOnly();
            Is64Bit = is64bit;
            AppId = appid;
            AccessPermissions = access_perm;
            LRpcPermissions = lrpc_perm;
            User = user;
            UserSid = user_sid;
            if (!String.IsNullOrWhiteSpace(rpc_endpoint))
            {
                RpcEndpoint = "OLE" + rpc_endpoint;
            }
            else
            {
                RpcEndpoint = String.Empty;
            }
            Capabilities = capabilities;
            AuthnLevel = authn_level;
            ImpLevel = imp_level;
            AccessControl = access_control;
            STAMainHWnd = sta_main_hwnd;
            Classes = classes.AsReadOnly();
        }

        public override string ToString()
        {
            return Name;
        }
    }

    [Flags]
    public enum IPIDFlags : uint
    {
        IPIDF_CONNECTING = 0x1,
        IPIDF_DISCONNECTED = 0x2,
        IPIDF_SERVERENTRY = 0x4,
        IPIDF_NOPING = 0x8,
        IPIDF_COPY = 0x10,
        IPIDF_VACANT = 0x80,
        IPIDF_NONNDRSTUB = 0x100,
        IPIDF_NONNDRPROXY = 0x200,
        IPIDF_NOTIFYACT = 0x400,
        IPIDF_TRIED_ASYNC = 0x800,
        IPIDF_ASYNC_SERVER = 0x1000,
        IPIDF_DEACTIVATED = 0x2000,
        IPIDF_WEAKREFCACHE = 0x4000,
        IPIDF_STRONGREFCACHE = 0x8000,
        IPIDF_UNSECURECALLSALLOWED = 0x10000,
    }

    public class COMMethodEntry
    {
        public string Name { get; private set; }
        public string Address { get; private set; }
        public string Symbol { get; private set; }
        public NdrProcedureDefinition Procedure { get; internal set; }

        internal COMMethodEntry(string name, string address, string symbol)
        {
            Name = name;
            Address = address;
            Symbol = symbol;
        }
    }

    public class COMIPIDEntry
    {
        public Guid Ipid { get; private set; }
        public Guid Iid { get; private set; }
        public IPIDFlags Flags { get; private set; }
        public IntPtr Interface { get; private set; }
        public string InterfaceVTable { get; private set; }
        public IEnumerable<COMMethodEntry> Methods { get; private set; }
        public IEnumerable<NdrComplexTypeReference> ComplexTypes { get; private set; }
        public IntPtr Stub { get; private set; }
        public string StubVTable { get; private set; }
        public Guid Oxid { get; private set; }
        public int StrongRefs { get; private set; }
        public int WeakRefs { get; private set; }
        public int PrivateRefs { get; private set; }
        public IntPtr ServerSTAHwnd { get; private set; }
        public int ApartmentId
        {
            get
            {
                return COMUtilities.GetApartmentIdFromIPid(Ipid);
            }
        }
        
        public bool IsRunning
        {
            get
            {
                return (Flags & (IPIDFlags.IPIDF_DISCONNECTED | IPIDFlags.IPIDF_DEACTIVATED)) == 0;
            }
        }

        public byte[] ToObjref()
        {
            MemoryStream stm = new MemoryStream();
            BinaryWriter writer = new BinaryWriter(stm);
            writer.Write(Encoding.ASCII.GetBytes("MEOW"));
            writer.Write(1);
            writer.Write(Iid.ToByteArray());
            writer.Write(0);
            writer.Write(1);
            writer.Write(Oxid.ToByteArray(), 0, 8);
            RandomNumberGenerator rng = RandomNumberGenerator.Create();
            byte[] oid = new byte[8];
            rng.GetBytes(oid);
            writer.Write(oid);
            writer.Write(Ipid.ToByteArray());
            writer.Write(0);
            return stm.ToArray();
        }

        private Dictionary<IntPtr, COMMethodEntry> _method_cache = new Dictionary<IntPtr, COMMethodEntry>();

        private static int GetPointerSize(NtProcess process)
        {
            if (process.Is64Bit)
            {
                return 8;
            }
            else
            {
                return 4;
            }
        }

        private static string GetSymbolName(string symbol)
        {
            int last_index = symbol.LastIndexOf("::");
            if (last_index >= 0)
            {
                symbol = symbol.Substring(last_index + 2);
            }

            last_index = symbol.LastIndexOf("`");
            if (last_index >= 0)
            {
                symbol = symbol.Substring(last_index + 1);
            }
            return symbol;
        }

        private COMMethodEntry ResolveMethod(int index, IntPtr method_ptr, ISymbolResolver resolver, 
            COMProcessParserConfig config)
        {
            if (!_method_cache.ContainsKey(method_ptr))
            {
                string address = resolver.GetModuleRelativeAddress(method_ptr);
                string symbol = config.ResolveMethodNames ? resolver.GetSymbolForAddress(method_ptr) : string.Empty;
                string name = index > 2 ? GetSymbolName(symbol) : string.Empty;
                if (string.IsNullOrWhiteSpace(name))
                {
                    switch (index)
                    {
                        case 0:
                            name = "QueryInterface";
                            break;
                        case 1:
                            name = "AddRef";
                            break;
                        case 2:
                            name = "Release";
                            break;
                        default:
                            name = string.Format("Method{0}", index);
                            break;
                    }
                }
                _method_cache[method_ptr] = new COMMethodEntry(name, address, symbol);
            }
            return _method_cache[method_ptr];
        }

        internal COMIPIDEntry(COMProcessParser.IPIDEntryNativeInterface ipid, NtProcess process, 
            ISymbolResolver resolver, COMProcessParserConfig config, COMRegistry registry)
        {
            Ipid = ipid.Ipid;
            Iid = ipid.Iid;
            Flags = (IPIDFlags)ipid.Flags;
            Interface = ipid.Interface;
            Stub = ipid.Stub;
            var oxid = ipid.GetOxidEntry(process);
            Oxid = oxid.MOxid;
            ServerSTAHwnd = oxid.ServerSTAHwnd;
            StrongRefs = ipid.StrongRefs;
            WeakRefs = ipid.WeakRefs;
            PrivateRefs = ipid.PrivateRefs;
            List<COMMethodEntry> methods = new List<COMMethodEntry>();
            List<NdrComplexTypeReference> complex_types = new List<NdrComplexTypeReference>();
            IntPtr stub_vptr = IntPtr.Zero;
            if (Stub != IntPtr.Zero)
            {
                stub_vptr = COMProcessParser.ReadPointer(process, Stub);
                StubVTable = resolver.GetModuleRelativeAddress(stub_vptr);
            }
            if (Interface != IntPtr.Zero)
            {
                IntPtr vtable_ptr = COMProcessParser.ReadPointer(process, Interface);
                InterfaceVTable = resolver.GetModuleRelativeAddress(vtable_ptr);
                long count = 0;
                IntPtr server_info = IntPtr.Zero;

                // For standard stubs the following exists before the vtable pointer:
                // PMIDL_SERVER_INFO ServerInfo - If ForwardingDispatchTable is NULL
                // DWORD_PTR DispatchTableCount
                // PVOID ForwardingDispatchTable - Used presumably when there's code implementation.
                if (stub_vptr != IntPtr.Zero)
                {
                    IntPtr base_ptr = new IntPtr(stub_vptr.ToInt64() - (GetPointerSize(process) * 3));
                    IntPtr[] stub_info = COMProcessParser.ReadPointerArray(process, base_ptr, 3);
                    if (stub_info[2] == IntPtr.Zero)
                    {
                        server_info = stub_info[0];
                    }
                    count = stub_info[1].ToInt64();
                }
                else if (registry.Interfaces.ContainsKey(Iid))
                {
                    count = registry.Interfaces[Iid].NumMethods;
                }

                // Sanity check, 256 methods should be enough for anyone ;-)
                if (count < 3 || count > 256)
                {
                    count = 3;
                }

                IntPtr[] method_ptrs = COMProcessParser.ReadPointerArray(process, vtable_ptr, (int)count);
                if (method_ptrs != null)
                {
                    methods.AddRange(method_ptrs.Select((p, i) => ResolveMethod(i, p, resolver, config)));
                    if (config.ParseStubMethods && server_info != IntPtr.Zero && count > 3)
                    {
                        NdrParser parser = new NdrParser(process, resolver);
                        var procs = parser.ReadFromMidlServerInfo(server_info, 3, (int)count, methods.Skip(3).Select(m => m.Name).ToList()).ToArray();
                        for (int i = 0; i < procs.Length; ++i)
                        {
                            methods[i + 3].Procedure = procs[i];
                        }
                        complex_types.AddRange(parser.ComplexTypes);
                    }
                }
            }
            Methods = methods.AsReadOnly();
            ComplexTypes = complex_types.AsReadOnly();
        }

        internal COMProxyInstance ToProxyInstance(string name)
        {
            NdrComProxyDefinition entry = NdrComProxyDefinition.FromProcedures(name, Iid, COMInterfaceEntry.IID_IUnknown,
                Methods.Count(), Methods.SkipWhile(m => m.Procedure == null).Select(m => m.Procedure));
            return new COMProxyInstance(new NdrComProxyDefinition[] { entry }, ComplexTypes);
        }
    }
}
