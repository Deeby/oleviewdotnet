//    This file is part of OleViewDotNet.
//    Copyright (C) James Forshaw 2014
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

using BrightIdeasSoftware;
using IronPython.Hosting;
using Microsoft.Scripting;
using Microsoft.Scripting.Hosting;
using NtApiDotNet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OleViewDotNet
{
    /// <summary>
    /// Form to view the COM registration information
    /// </summary>
    public partial class COMRegistryViewer : UserControl
    {
        /// <summary>
        /// Current registry
        /// </summary>
        private readonly COMRegistry m_registry;
        private readonly HashSet<FilterType> m_filter_types;
        private readonly DisplayMode m_mode;
        private readonly IEnumerable<COMProcessEntry> m_processes;
        private RegistryViewerFilter m_filter;
        private Dictionary<COMCLSIDServerEntry, List<COMCLSIDEntry>> m_servers_to_clsids;
        private int m_total_count;

        /// <summary>
        /// Enumeration to indicate what to display
        /// </summary>
        public enum DisplayMode
        {
            CLSIDs,
            ProgIDs,
            CLSIDsByName,
            CLSIDsByServer,
            CLSIDsByLocalServer,
            CLSIDsWithSurrogate,
            Interfaces,
            InterfacesByName,
            ImplementedCategories,
            PreApproved,
            IELowRights,
            LocalServices,
            AppIDs,
            Typelibs,
            AppIDsWithIL,
            MimeTypes,
            AppIDsWithAC,
            ProxyCLSIDs,
            Processes,
            RuntimeClasses,
            RuntimeServers,
        }

        private const string FolderKey = "folder.ico";
        private const string InterfaceKey = "interface.ico";
        private const string ClassKey = "class.ico";
        private const string FolderOpenKey = "folderopen.ico";
        private const string ProcessKey = "process.ico";
        private const string ApplicationKey = "application.ico";

        class PlaceholderObject
        {
            public string Name { get; }
            public bool HasGuid { get; }
            public Guid Guid { get; }
            public IEnumerable<object> ChildObjects { get; }
            public object RealObject { get; }
            public string Tooltip { get; }
            public string IconKey { get; }
            public string ExpandedIconKey { get; }

            public PlaceholderObject(string name, bool has_guid, Guid guid, object real_object, IEnumerable<object> child_objects, string tooltip, string icon_key, string expanded_icon_key)
            {
                Name = name;
                HasGuid = has_guid;
                Guid = guid;
                ChildObjects = child_objects.ToList().AsReadOnly();
                RealObject = real_object;
                Tooltip = tooltip;
                IconKey = icon_key;
                ExpandedIconKey = expanded_icon_key;
            }
        }

        private static object GetRealObject(object obj)
        {
            if (obj is PlaceholderObject placeholder)
            {
                return placeholder.RealObject;
            }
            return obj;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="reg">The COM registry</param>
        /// <param name="mode">The display mode</param>
        public COMRegistryViewer(COMRegistry reg, DisplayMode mode, IEnumerable<COMProcessEntry> processes) 
            : this(reg, mode, processes, null, GetFilterTypes(mode), GetDisplayName(mode))
        {
        }

        private static string GetDisplayName(DisplayMode mode)
        {
            switch (mode)
            {
                case DisplayMode.CLSIDsByName:
                    return "CLSIDs by Name";
                case DisplayMode.CLSIDs:
                    return "CLSIDs";
                case DisplayMode.ProgIDs:
                    return "ProgIDs";
                case DisplayMode.CLSIDsByServer:
                    return "CLSIDs by Server";
                case DisplayMode.CLSIDsByLocalServer:
                    return "CLSIDs by Local Server";
                case DisplayMode.CLSIDsWithSurrogate:
                    return "CLSIDs with Surrogate";
                case DisplayMode.Interfaces:
                    return "Interfaces";
                case DisplayMode.InterfacesByName:
                    return "Interfaces by Name";
                case DisplayMode.ImplementedCategories:
                    return "Implemented Categories";
                case DisplayMode.PreApproved:
                    return "Explorer Pre-Approved";
                case DisplayMode.IELowRights:
                    return "IE Low Rights Policy";
                case DisplayMode.LocalServices:
                    return "Local Services";
                case DisplayMode.AppIDs:
                    return "AppIDs";
                case DisplayMode.AppIDsWithIL:
                    return "AppIDs with IL";
                case DisplayMode.AppIDsWithAC:
                    return "AppIDs with AC";
                case DisplayMode.Typelibs:
                    return "TypeLibs";
                case DisplayMode.MimeTypes:
                    return "MIME Types";
                case DisplayMode.ProxyCLSIDs:
                    return "Proxy CLSIDs";
                case DisplayMode.Processes:
                    return "COM Processes";
                case DisplayMode.RuntimeClasses:
                    return "Runtime Classes";
                case DisplayMode.RuntimeServers:
                    return "Runtime Servers";
                default:
                    throw new ArgumentException("Invalid mode value");
            }
        }

        private static IEnumerable<FilterType> GetFilterTypes(DisplayMode mode)
        {
            HashSet<FilterType> filter_types = new HashSet<FilterType>();
            switch (mode)
            {
                case DisplayMode.CLSIDsByName:
                case DisplayMode.CLSIDs:
                    filter_types.Add(FilterType.CLSID);
                    filter_types.Add(FilterType.Interface);
                    break;
                case DisplayMode.ProgIDs:
                    filter_types.Add(FilterType.ProgID);
                    filter_types.Add(FilterType.Interface);
                    break;
                case DisplayMode.CLSIDsByServer:
                case DisplayMode.CLSIDsByLocalServer:
                case DisplayMode.CLSIDsWithSurrogate:
                case DisplayMode.ProxyCLSIDs:
                    filter_types.Add(FilterType.CLSID);
                    filter_types.Add(FilterType.Server);
                    filter_types.Add(FilterType.Interface);
                    break;
                case DisplayMode.Interfaces:
                case DisplayMode.InterfacesByName:
                    filter_types.Add(FilterType.Interface);
                    break;
                case DisplayMode.ImplementedCategories:
                    filter_types.Add(FilterType.Category);
                    filter_types.Add(FilterType.CLSID);
                    filter_types.Add(FilterType.Interface);
                    break;
                case DisplayMode.PreApproved:
                    filter_types.Add(FilterType.CLSID);
                    filter_types.Add(FilterType.Interface);
                    break;
                case DisplayMode.IELowRights:
                    filter_types.Add(FilterType.LowRights);
                    filter_types.Add(FilterType.CLSID);
                    filter_types.Add(FilterType.Interface);
                    break;
                case DisplayMode.AppIDs:
                case DisplayMode.AppIDsWithIL:
                case DisplayMode.AppIDsWithAC:
                case DisplayMode.LocalServices:
                    filter_types.Add(FilterType.AppID);
                    filter_types.Add(FilterType.CLSID);
                    filter_types.Add(FilterType.Interface);
                    break;
                case DisplayMode.Typelibs:
                    filter_types.Add(FilterType.TypeLib);
                    break;
                case DisplayMode.MimeTypes:
                    filter_types.Add(FilterType.MimeType);
                    filter_types.Add(FilterType.CLSID);
                    filter_types.Add(FilterType.Interface);
                    break;
                case DisplayMode.Processes:
                    filter_types.Add(FilterType.Process);
                    filter_types.Add(FilterType.Ipid);
                    filter_types.Add(FilterType.AppID);
                    break;
                case DisplayMode.RuntimeClasses:
                    filter_types.Add(FilterType.RuntimeClass);
                    break;
                case DisplayMode.RuntimeServers:
                    filter_types.Add(FilterType.RuntimeServer);
                    break;
                default:
                    throw new ArgumentException("Invalid mode value");
            }
            return filter_types;
        }

        private void UpdateStatusLabel()
        {
            toolStripStatusLabelCount.Text = String.Format("Showing {0} of {1} entries", treeListView.GetItemCount(), m_total_count);
        }

        private bool CanExpand(object o)
        {
            if (o is PlaceholderObject placeholder)
            {
                return placeholder.ChildObjects.Any();
            }
            return o is COMCategory || o is ICOMClassEntry || o is COMCLSIDServerEntry || o is COMProcessEntry || o is COMAppIDEntry;
        }

        private IEnumerable<object> GetChildren(object o)
        {
            if (o is COMCategory category)
            {
                return category.Clsids.Select(c => m_registry.MapClsidToEntry(c)).Where(c => c != null).OrderBy(c => c.Name);
            }
            else if (o is COMCLSIDServerEntry server)
            {
                return m_servers_to_clsids[server].OrderBy(c => c.Name);
            }
            else if (o is COMProcessEntry process)
            {
                return process.Ipids;
            }
            else if (o is COMAppIDEntry appid)
            {

            }
            else if (o is PlaceholderObject placeholder)
            {
                return placeholder.ChildObjects;
            }

            return new object[0];
        }

        private string GetObjectName(object obj)
        {
            if (obj is COMInterfaceInstance intf)
            {
                return m_registry.MapIidToInterface(intf.Iid).Name;
            }
            else if (obj is PlaceholderObject placeholder)
            {
                return placeholder.Name;
            }
            return obj.ToString();
        }

        private void SetupColumns(DisplayMode mode)
        {
            olvColumnGuid.AspectGetter = o => CanGetGuid(o) ? GetGuidFromType(o).FormatGuid() : string.Empty;
            olvColumnName.AspectGetter = o => GetObjectName(o);
            olvColumnName.ImageGetter = o =>
            {
                if (o is ICOMClassEntry)
                {
                    return ClassKey;
                }
                else if (o is COMInterfaceEntry || o is COMIPIDEntry || o is COMInterfaceInstance)
                {
                    return InterfaceKey;
                }
                else if (o is COMCategory || o is COMCLSIDServerEntry || o is COMRuntimeServerEntry || o is COMAppIDEntry)
                {
                    if (treeListView.IsExpanded(o))
                    {
                        return FolderOpenKey;
                    }
                    return FolderKey;
                }
                else if (o is COMProcessEntry)
                {
                    return ProcessKey;
                }
                else if (o is PlaceholderObject placeholder)
                {
                    return treeListView.IsExpanded(o) ? placeholder.ExpandedIconKey : placeholder.IconKey;
                }
                return string.Empty;
            };
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="reg">The COM registry</param>
        /// <param name="mode">The display mode</param>
        public COMRegistryViewer(COMRegistry reg, DisplayMode mode, IEnumerable<COMProcessEntry> processes, 
            IEnumerable<object> nodes, IEnumerable<FilterType> filter_types, string text)
        {
            InitializeComponent();
            m_registry = reg;
            m_filter_types = new HashSet<FilterType>(filter_types);
            m_filter = new RegistryViewerFilter();
            m_mode = mode;
            m_processes = processes;
            treeImageList.Images.Add(ApplicationKey, SystemIcons.Application);
            SetupColumns(m_mode);

            foreach (FilterMode filter in Enum.GetValues(typeof(FilterMode)))
            {
                comboBoxMode.Items.Add(filter);
            }
            comboBoxMode.SelectedIndex = 0;

            Text = text;
            treeListView.CanExpandGetter = CanExpand;
            treeListView.ChildrenGetter = GetChildren;
            var root_objects = nodes ?? SetupTree(reg, mode, processes);
            m_total_count = root_objects.Count();
            treeListView.SetObjects(root_objects);
            treeListView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            treeListView.CellToolTipGetter = GetTooltip;
            UpdateStatusLabel();
        }

        private string GetTooltip(OLVColumn col, object obj)
        {
            if (obj is COMCLSIDEntry clsid)
            {
                return BuildCLSIDToolTip(m_registry, clsid);
            }
            else if (obj is COMProgIDEntry progid)
            {
                return BuildProgIDToolTip(m_registry, progid);
            }
            else if (obj is COMInterfaceEntry intf)
            {
                return BuildInterfaceToolTip(intf, null);
            }
            else if (obj is COMInterfaceInstance intf_instance)
            {
                COMInterfaceEntry intf_entry = m_registry.MapIidToInterface(intf_instance.Iid);
                return BuildInterfaceToolTip(intf_entry, intf_instance);
            }
            else if (obj is COMProcessEntry process)
            {
                return BuildCOMProcessTooltip(process);
            }
            else if (obj is PlaceholderObject placeholder)
            {
                if (placeholder.Tooltip != null)
                {
                    return placeholder.Tooltip;
                }
                return GetTooltip(col, placeholder.RealObject);
            }
            else if (obj is COMTypeLibVersionEntry typelib_version)
            {
                return BuildTypelibVersionTooltip(typelib_version);
            }

            return null;
        }

        private IEnumerable<object> SetupTree(COMRegistry registry, DisplayMode mode, IEnumerable<COMProcessEntry> processes)
        {   
            try
            {
                switch (mode)
                {
                    case DisplayMode.CLSIDsByName:
                        return registry.Clsids.Values.OrderBy(c => c.Name);
                    case DisplayMode.CLSIDs:
                        return registry.Clsids.Values.OrderBy(c => c.Clsid);
                    case DisplayMode.ProgIDs:
                        return registry.Progids.Values.OrderBy(p => p.Name);
                    case DisplayMode.CLSIDsByServer:
                        return LoadCLSIDByServer(registry, ServerType.None);
                    case DisplayMode.CLSIDsByLocalServer:
                        return LoadCLSIDByServer(registry, ServerType.Local);
                    case DisplayMode.CLSIDsWithSurrogate:
                        return LoadCLSIDByServer(registry, ServerType.Surrogate);
                    case DisplayMode.Interfaces:
                        return registry.Interfaces.Values.OrderBy(i => i.Iid);
                    case DisplayMode.InterfacesByName:
                        return registry.Interfaces.Values.OrderBy(i => i.Name);
                    case DisplayMode.ImplementedCategories:
                        return registry.ImplementedCategories.Values.OrderBy(c => c.Name);
                    case DisplayMode.PreApproved:
                        return registry.PreApproved.OrderBy(c => c.Name);
                    case DisplayMode.IELowRights:
                        return LoadIELowRights(registry);
                    case DisplayMode.LocalServices:
                        return LoadLocalServices(registry);
                    case DisplayMode.AppIDs:
                        return LoadAppIDs(registry, false, false);
                    case DisplayMode.AppIDsWithIL:
                        return LoadAppIDs(registry, true, false);
                    case DisplayMode.AppIDsWithAC:
                        return LoadAppIDs(registry, false, true);
                    case DisplayMode.Typelibs:
                        return LoadTypeLibs(registry);
                    case DisplayMode.MimeTypes:
                        return LoadMimeTypes(registry);
                    case DisplayMode.ProxyCLSIDs:
                        return LoadCLSIDByServer(registry, ServerType.Proxies);
                    case DisplayMode.Processes:
                        return LoadProcesses(registry, processes);
                    case DisplayMode.RuntimeClasses:
                        return registry.RuntimeClasses.Values.OrderBy(c => c.Name);
                    case DisplayMode.RuntimeServers:
                        return LoadRuntimeServers(registry);
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Program.ShowError(null, ex);
            }

            return new object[0];
        }

        /// <summary>
        /// Build a tooltip for a CLSID entry
        /// </summary>
        /// <param name="ent">The CLSID entry to build the tool tip from</param>
        /// <returns>A string tooltip</returns>
        private static string BuildCLSIDToolTip(COMRegistry registry, COMCLSIDEntry ent)
        {
            StringBuilder strRet = new StringBuilder();

            AppendFormatLine(strRet, "CLSID: {0}", ent.Clsid.FormatGuid());
            AppendFormatLine(strRet, "Name: {0}", ent.Name);
            AppendFormatLine(strRet, "{0}: {1}", ent.DefaultServerType.ToString(), ent.DefaultServer);
            IEnumerable<string> progids = registry.GetProgIdsForClsid(ent.Clsid).Select(p => p.ProgID);
            if (progids.Any())
            {
                strRet.AppendLine("ProgIDs:");
                foreach (string progid in progids)
                {
                    AppendFormatLine(strRet, "{0}", progid);
                }
            }
            if (ent.AppID != Guid.Empty)
            {
                AppendFormatLine(strRet, "AppID: {0}", ent.AppID.FormatGuid());
            }
            if (ent.TypeLib != Guid.Empty)
            {
                AppendFormatLine(strRet, "TypeLib: {0}", ent.TypeLib.FormatGuid());
            }

            COMInterfaceEntry[] proxies = registry.GetProxiesForClsid(ent);
            if (proxies.Length > 0)
            {
                AppendFormatLine(strRet, "Interface Proxies: {0}", proxies.Length);
            }

            if (ent.InterfacesLoaded)
            {
                AppendFormatLine(strRet, "Instance Interfaces: {0}", ent.Interfaces.Count());
                AppendFormatLine(strRet, "Factory Interfaces: {0}", ent.FactoryInterfaces.Count());
            }
            if (ent.DefaultServerType == COMServerType.InProcServer32)
            {
                COMCLSIDServerEntry server = ent.Servers[COMServerType.InProcServer32];
                if (server.HasDotNet)
                {
                    AppendFormatLine(strRet, "Assembly: {0}", server.DotNet.AssemblyName);
                    AppendFormatLine(strRet, "Class: {0}", server.DotNet.ClassName);
                    if (!String.IsNullOrWhiteSpace(server.DotNet.CodeBase))
                    {
                        AppendFormatLine(strRet, "Codebase: {0}", server.DotNet.CodeBase);
                    }
                    if (!String.IsNullOrWhiteSpace(server.DotNet.RuntimeVersion))
                    {
                        AppendFormatLine(strRet, "Runtime Version: {0}", server.DotNet.RuntimeVersion);
                    }
                }
            }

            return strRet.ToString();
        }

        /// <summary>
        /// Build a ProgID entry tooltip
        /// </summary>
        /// <param name="ent">The ProgID entry</param>
        /// <returns>The ProgID tooltip</returns>
        private static string BuildProgIDToolTip(COMRegistry registry, COMProgIDEntry ent)
        {
            string strRet;
            COMCLSIDEntry entry = registry.MapClsidToEntry(ent.Clsid);
            if (entry != null)
            {
                strRet = BuildCLSIDToolTip(registry, entry);
            }
            else
            {
                strRet = String.Format("CLSID: {0}\n", ent.Clsid.FormatGuid());
            }

            return strRet;
        }

        private static string BuildInterfaceToolTip(COMInterfaceEntry ent, COMInterfaceInstance instance)
        {            
            StringBuilder builder = new StringBuilder();

            AppendFormatLine(builder, "Name: {0}", ent.Name);
            AppendFormatLine(builder, "IID: {0}", ent.Iid.FormatGuid());
            if (ent.ProxyClsid != Guid.Empty)
            {
                AppendFormatLine(builder, "ProxyCLSID: {0}", ent.ProxyClsid.FormatGuid());
            }
            if (instance != null && instance.Module != null)
            {
                AppendFormatLine(builder, "VTable Address: {0}+0x{1:X}", instance.Module, instance.VTableOffset);
            }
            if (ent.HasTypeLib)
            {
                AppendFormatLine(builder, "TypeLib: {0}", ent.TypeLib.FormatGuid());
            }

            return builder.ToString();
        }

        private static IEnumerable<object> LoadRuntimeServers(COMRegistry registry)
        {
            List<PlaceholderObject> serverNodes = new List<PlaceholderObject>();
            foreach (var group in registry.RuntimeClasses.Values.GroupBy(p => p.Server.ToLower()))
            {
                COMRuntimeServerEntry server = registry.MapServerNameToEntry(group.Key);
                if (server == null)
                {
                    continue;
                }

                serverNodes.Add(new PlaceholderObject(server.Name, false, 
                    Guid.Empty, server, group.OrderBy(r => r.Name), null, FolderKey, FolderOpenKey));
            }
            return serverNodes.OrderBy(n => n.Name);
        }

        private static string BuildCOMProcessTooltip(COMProcessEntry proc)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendFormat("Path: {0}", proc.ExecutablePath).AppendLine();
            builder.AppendFormat("User: {0}", proc.User).AppendLine();
            if (proc.AppId != Guid.Empty)
            {
                builder.AppendFormat("AppID: {0}", proc.AppId).AppendLine();
            }
            builder.AppendFormat("Access Permissions: {0}", proc.AccessPermissions).AppendLine();
            builder.AppendFormat("LRPC Permissions: {0}", proc.LRpcPermissions).AppendLine();
            if (!String.IsNullOrEmpty(proc.RpcEndpoint))
            {
                builder.AppendFormat("LRPC Endpoint: {0}", proc.RpcEndpoint).AppendLine();
            }
            builder.AppendFormat("Capabilities: {0}", proc.Capabilities).AppendLine();
            builder.AppendFormat("Authn Level: {0}", proc.AuthnLevel).AppendLine();
            builder.AppendFormat("Imp Level: {0}", proc.ImpLevel).AppendLine();
            if (proc.AccessControl != IntPtr.Zero)
            {
                builder.AppendFormat("Access Control: 0x{0:X}", proc.AccessControl.ToInt64());
            }
            builder.Append(COMUtilities.FormatBitness(proc.Is64Bit));
            return builder.ToString();
        }

        private static string BuildCOMIpidTooltip(COMIPIDEntry ipid)
        {
            StringBuilder builder = new StringBuilder();
            AppendFormatLine(builder, "Interface: 0x{0:X}", ipid.Interface.ToInt64());
            if (!String.IsNullOrWhiteSpace(ipid.InterfaceVTable))
            {
                AppendFormatLine(builder, "Interface VTable: {0}", ipid.InterfaceVTable);
            }
            AppendFormatLine(builder, "Stub: 0x{0:X}", ipid.Stub.ToInt64());
            if (!String.IsNullOrWhiteSpace(ipid.StubVTable))
            {
                AppendFormatLine(builder, "Stub VTable: {0}", ipid.StubVTable);
            }
            AppendFormatLine(builder, "Flags: {0}", ipid.Flags);
            AppendFormatLine(builder, "Strong Refs: {0}", ipid.StrongRefs);
            AppendFormatLine(builder, "Weak Refs: {0}", ipid.WeakRefs);
            AppendFormatLine(builder, "Private Refs: {0}", ipid.PrivateRefs);
            
            return builder.ToString();
        }

        private static string BuildCOMProcessName(COMProcessEntry proc)
        {
            return String.Format("{0,-8} - {1} - {2}", proc.Pid, proc.Name, proc.User);
        }

        private static void PopulateIpids(COMRegistry registry, TreeNode node, COMProcessEntry proc)
        {
            //foreach (COMIPIDEntry ipid in proc.Ipids.Where(i => i.IsRunning))
            //{
            //    COMInterfaceEntry intf = registry.MapIidToInterface(ipid.Iid);
            //    TreeNode ipid_node = CreateNode(String.Format("IPID: {0} - IID: {1}", ipid.Ipid.FormatGuid(), intf.Name), InterfaceKey);
            //    ipid_node.ToolTipText = BuildCOMIpidTooltip(ipid);
            //    ipid_node.Tag = ipid;
            //    node.Nodes.Add(ipid_node);
            //}
        }
        
        private static TreeNode CreateCOMProcessNode(COMRegistry registry, COMProcessEntry proc, 
            IDictionary<int, IEnumerable<COMAppIDEntry>> appIdsByPid, IDictionary<Guid, List<COMCLSIDEntry>> clsidsByAppId)
        {
            return null;
            //TreeNode node = CreateNode(BuildCOMProcessName(proc), ApplicationKey);
            //node.ToolTipText = BuildCOMProcessTooltip(proc);
            //node.Tag = proc;

            //if (appIdsByPid.ContainsKey(proc.Pid) && appIdsByPid[proc.Pid].Count() > 0)
            //{
            //    TreeNode services_node = CreateNode("Services", FolderKey);
            //    foreach (COMAppIDEntry appid in appIdsByPid[proc.Pid])
            //    {
            //        if (clsidsByAppId.ContainsKey(appid.AppId))
            //        {
            //            services_node.Nodes.Add(CreateLocalServiceNode(registry, appid, clsidsByAppId[appid.AppId]));
            //        }
            //    }
            //    node.Nodes.Add(services_node);
            //}

            //var server_classes = proc.Classes.Where(c => (c.Context & CLSCTX.LOCAL_SERVER) != 0);

            //if (server_classes.Any())
            //{
            //    TreeNode classes_node = CreateNode("Classes", FolderKey);
            //    foreach (var c in server_classes)
            //    {
            //        classes_node.Nodes.Add(CreateCLSIDNode(registry, registry.MapClsidToEntry(c.Clsid)));
            //    }

            //    node.Nodes.Add(classes_node);
            //}

            //PopulatorIpids(registry, node, proc);
            //return node;
        }

        private static IEnumerable<TreeNode> LoadProcesses(COMRegistry registry, IEnumerable<COMProcessEntry> processes)
        {
            var servicesById = COMUtilities.GetServicePids();
            var appidsByService = registry.AppIDs.Values.Where(a => a.IsService).
                GroupBy(a => a.LocalService.Name, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g, StringComparer.OrdinalIgnoreCase);
            var clsidsByAppId = registry.ClsidsByAppId.ToDictionary(g => g.Key, g => g.ToList());
            var appsByPid = servicesById.ToDictionary(p => p.Key, p => p.Value.Where(v => appidsByService.ContainsKey(v)).SelectMany(v => appidsByService[v]));

            return processes.Where(p => p.Ipids.Any()).Select(p => CreateCOMProcessNode(registry, p, appsByPid, clsidsByAppId));
        }

        enum ServerType
        {
            None,
            Local,
            Surrogate,
            Proxies,
        }

        private static bool IsProxyClsid(COMRegistry registry, COMCLSIDEntry ent)
        {
            return ent.DefaultServerType == COMServerType.InProcServer32 && registry.GetProxiesForClsid(ent).Length > 0;
        }

        private static bool HasSurrogate(COMRegistry registry, COMCLSIDEntry ent)
        {
            return registry.AppIDs.ContainsKey(ent.AppID) && !String.IsNullOrWhiteSpace(registry.AppIDs[ent.AppID].DllSurrogate);
        }

        private class COMCLSIDServerEqualityComparer : IEqualityComparer<COMCLSIDServerEntry>
        {
            public bool Equals(COMCLSIDServerEntry x, COMCLSIDServerEntry y)
            {
                return x.Server.Equals(y.Server, StringComparison.OrdinalIgnoreCase);
            }

            public int GetHashCode(COMCLSIDServerEntry obj)
            {
                return obj.Server.ToLower().GetHashCode();
            }
        }

        private IEnumerable<COMCLSIDServerEntry> LoadCLSIDByServer(COMRegistry registry, ServerType serverType)
        {
            if (serverType == ServerType.Surrogate)
            {
                m_servers_to_clsids = registry.Clsids.Values.Where(c => HasSurrogate(registry, c))
                    .GroupBy(c => registry.AppIDs[c.AppID].DllSurrogate, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => new COMCLSIDServerEntry(COMServerType.LocalServer32, g.Key), g => g.AsEnumerable().ToList(),
                    new COMCLSIDServerEqualityComparer());
            }
            else
            {
                Dictionary<COMCLSIDServerEntry, List<COMCLSIDEntry>> dict =
                    new Dictionary<COMCLSIDServerEntry, List<COMCLSIDEntry>>(new COMCLSIDServerEqualityComparer());
                IEnumerable<COMCLSIDEntry> clsids = registry.Clsids.Values.Where(e => e.Servers.Count > 0);
                if (serverType == ServerType.Proxies)
                {
                    clsids = clsids.Where(c => IsProxyClsid(registry, c));
                }

                foreach (COMCLSIDEntry entry in clsids)
                {
                    foreach (COMCLSIDServerEntry server in entry.Servers.Values)
                    {
                        if (serverType == ServerType.Local && server.ServerType != COMServerType.LocalServer32)
                        {
                            continue;
                        }

                        if (!dict.ContainsKey(server))
                        {
                            dict[server] = new List<COMCLSIDEntry>();
                        }
                        dict[server].Add(entry);
                    }
                }
                m_servers_to_clsids = dict;
            }

            return m_servers_to_clsids.Keys.OrderBy(n => n.Name);
        }
        
        private static StringBuilder AppendFormatLine(StringBuilder builder, string format, params object[] ps)
        {
            return builder.AppendFormat(format, ps).AppendLine();
        }

        private static PlaceholderObject CreateLocalServiceNode(COMRegistry registry, COMAppIDEntry appidEnt, IEnumerable<COMCLSIDEntry> clsids)
        {
            string name = appidEnt.LocalService.DisplayName;
            if (String.IsNullOrWhiteSpace(name))
            {
                name = appidEnt.LocalService.Name;
            }

            return new PlaceholderObject(name, true, appidEnt.AppId, appidEnt, clsids.OrderBy(c => c.Name), BuildAppIdTooltip(appidEnt), FolderKey, FolderOpenKey);
        }

        private static IEnumerable<object> LoadLocalServices(COMRegistry registry)
        {
            List<IGrouping<Guid, COMCLSIDEntry>> clsidsByAppId = registry.ClsidsByAppId.ToList();
            IDictionary<Guid, COMAppIDEntry> appids = registry.AppIDs;

            List<PlaceholderObject> serverNodes = new List<PlaceholderObject>();
            foreach (IGrouping<Guid, COMCLSIDEntry> pair in clsidsByAppId)
            {   
                if(appids.ContainsKey(pair.Key) && appids[pair.Key].IsService)
                {
                    serverNodes.Add(CreateLocalServiceNode(registry, appids[pair.Key], pair));
                }
            }

            return serverNodes.OrderBy(n => n.Name);
        }

        static string LimitString(string s, int max)
        {
            if (s.Length > max)
            {
                return s.Substring(0, max) + "...";
            }
            return s;
        }

        private static string BuildAppIdTooltip(COMAppIDEntry appidEnt)
        {
            StringBuilder builder = new StringBuilder();

            AppendFormatLine(builder, "Name: {0}", appidEnt.Name);
            AppendFormatLine(builder, "AppID: {0}", appidEnt.AppId);
            if (!String.IsNullOrWhiteSpace(appidEnt.RunAs))
            {
                AppendFormatLine(builder, "RunAs: {0}", appidEnt.RunAs);
            }

            if (appidEnt.IsService)
            {
                COMAppIDServiceEntry service = appidEnt.LocalService;
                AppendFormatLine(builder, "Service Name: {0}", service.Name);
                if (!String.IsNullOrWhiteSpace(service.DisplayName))
                {
                    AppendFormatLine(builder, "Display Name: {0}", service.DisplayName);
                }
                if (!String.IsNullOrWhiteSpace(service.UserName))
                {
                    AppendFormatLine(builder, "Service User: {0}", service.UserName);
                }
                AppendFormatLine(builder, "Image Path: {0}", service.ImagePath);
                if (!String.IsNullOrWhiteSpace(service.ServiceDll))
                {
                    AppendFormatLine(builder, "Service DLL: {0}", service.ServiceDll);
                }
            }

            if (appidEnt.HasLaunchPermission)
            {
                AppendFormatLine(builder, "Launch: {0}", LimitString(appidEnt.LaunchPermission, 64));
            }

            if (appidEnt.HasAccessPermission)
            {
                AppendFormatLine(builder, "Access: {0}", LimitString(appidEnt.AccessPermission, 64));
            }

            if (appidEnt.RotFlags != COMAppIDRotFlags.None)
            {
                AppendFormatLine(builder, "RotFlags: {0}", appidEnt.RotFlags);
            }

            if (!String.IsNullOrWhiteSpace(appidEnt.DllSurrogate))
            {
                AppendFormatLine(builder, "DLL Surrogate: {0}", appidEnt.DllSurrogate);
            }

            if (appidEnt.Flags != COMAppIDFlags.None)
            {
                AppendFormatLine(builder, "Flags: {0}", appidEnt.Flags);
            }

            return builder.ToString();
        }

        private static IEnumerable<object> LoadAppIDs(COMRegistry registry, bool filterIL, bool filterAC)
        {
            IDictionary<Guid, List<COMCLSIDEntry>> clsidsByAppId = registry.ClsidsByAppId.ToDictionary(g => g.Key, g => g.ToList());
            IDictionary<Guid, COMAppIDEntry> appids = registry.AppIDs;

            List<PlaceholderObject> serverNodes = new List<PlaceholderObject>();
            foreach (var pair in appids)
            {
                COMAppIDEntry appidEnt = appids[pair.Key];

                if (filterIL && COMSecurity.GetILForSD(appidEnt.AccessPermission) == TokenIntegrityLevel.Medium &&
                    COMSecurity.GetILForSD(appidEnt.LaunchPermission) == TokenIntegrityLevel.Medium)
                {
                    continue;
                }

                if (filterAC && !COMSecurity.SDHasAC(appidEnt.AccessPermission) && !COMSecurity.SDHasAC(appidEnt.LaunchPermission))
                {
                    continue;
                }

                IEnumerable<COMCLSIDEntry> clsids = new COMCLSIDEntry[0];

                if (clsidsByAppId.ContainsKey(pair.Key))
                {
                    clsids = clsidsByAppId[pair.Key].OrderBy(c => c.Name);
                }

                serverNodes.Add(new PlaceholderObject(appidEnt.Name, true, appidEnt.AppId, appidEnt, clsids, BuildAppIdTooltip(appidEnt), FolderKey, FolderOpenKey));
            }

            return serverNodes.OrderBy(n => n.Name);
        }

        private static IEnumerable<object> LoadIELowRights(COMRegistry registry)
        {
            List<PlaceholderObject> entries = new List<PlaceholderObject>();
            foreach (COMIELowRightsElevationPolicy ent in registry.LowRights)
            {
                StringBuilder tooltip = new StringBuilder();
                List<COMCLSIDEntry> clsids = new List<COMCLSIDEntry>();
                if (ent.Clsid != Guid.Empty)
                {
                    clsids.Add(registry.MapClsidToEntry(ent.Clsid));
                }

                if (!String.IsNullOrWhiteSpace(ent.AppPath) && registry.ClsidsByServer.ContainsKey(ent.AppPath))
                {
                    clsids.AddRange(registry.ClsidsByServer[ent.AppPath]);
                    tooltip.AppendFormat("{0}", ent.AppPath);
                    tooltip.AppendLine();
                }

                if (clsids.Count == 0)
                {
                    continue;
                }

                tooltip.AppendFormat("Policy: {0}", ent.Policy);
                tooltip.AppendLine();
                entries.Add(new PlaceholderObject(ent.Name, true, ent.Uuid, ent, clsids, tooltip.ToString(), FolderKey, FolderOpenKey));
            }

            return entries.OrderBy(p => p.Name);
        }

        private static IEnumerable<object> LoadMimeTypes(COMRegistry registry)
        {
            List<PlaceholderObject> nodes = new List<PlaceholderObject>(registry.MimeTypes.Count());
            foreach (COMMimeType ent in registry.MimeTypes)
            {
                List<COMCLSIDEntry> clsids = new List<COMCLSIDEntry>();
                if (registry.Clsids.ContainsKey(ent.Clsid))
                {
                    clsids.Add(registry.Clsids[ent.Clsid]);
                }

                string tooltip = null;
                if (!String.IsNullOrWhiteSpace(ent.Extension))
                {
                    tooltip = $"Extension {ent.Extension}";
                }
                nodes.Add(new PlaceholderObject(ent.MimeType, false, Guid.Empty, ent, clsids, tooltip, FolderKey, FolderOpenKey));
            }

            return nodes;
        }

        private static string BuildTypelibVersionTooltip(COMTypeLibVersionEntry entry)
        {
            List<string> entries = new List<string>();
            entries.Add($"Version: {entry.Version}");
            if (!string.IsNullOrWhiteSpace(entry.Win32Path))
            {
                entries.Add($"Win32: {entry.Win32Path}");
            }
            if (!string.IsNullOrWhiteSpace(entry.Win64Path))
            {
                entries.Add($"Win64: {entry.Win64Path}");
            }
            return String.Join("\r\n", entries);
        }

        private static IEnumerable<object> LoadTypeLibs(COMRegistry registry)
        {
            List<PlaceholderObject> typelibNodes = new List<PlaceholderObject>();
            foreach (COMTypeLibEntry ent in registry.Typelibs.Values)
            {
                string name = ent.Versions.Select(v => v.Name).FirstOrDefault(n => !string.IsNullOrWhiteSpace(n));
                typelibNodes.Add(new PlaceholderObject(name ?? ent.TypelibId.FormatGuid(), 
                    true, ent.TypelibId, ent, ent.Versions, null, FolderKey, FolderOpenKey));
            }

            return typelibNodes;
        }

        private void AddInterfaceNodes(TreeNode node, IEnumerable<COMInterfaceInstance> intfs)
        {
            //node.Nodes.AddRange(intfs.Select(i => CreateInterfaceNameNode(m_registry, m_registry.MapIidToInterface(i.Iid), i)).OrderBy(n => n.Text).ToArray());
        }

        private Task SetupCLSIDNodeTree(TreeNode node, bool bRefresh)
        {
            return null;
            //ICOMClassEntry clsid = node.Tag as ICOMClassEntry;

            //if (clsid == null && node.Tag is COMProgIDEntry)
            //{
            //    clsid = m_registry.MapClsidToEntry(((COMProgIDEntry)node.Tag).Clsid);
            //}

            //if (clsid != null)
            //{
            //    node.Nodes.Clear();
            //    TreeNode wait_node = CreateNode("Please Wait, Populating Interfaces", InterfaceKey);
            //    node.Nodes.Add(wait_node);
            //    try
            //    {
            //        await clsid.LoadSupportedInterfacesAsync(bRefresh);
            //        int interface_count = clsid.Interfaces.Count();
            //        int factory_count = clsid.FactoryInterfaces.Count();
            //        if (interface_count == 0 && factory_count == 0)
            //        {
            //            wait_node.Text = "Error querying COM interfaces - Timeout";
            //        }
            //        else
            //        {
            //            if (interface_count > 0)
            //            {
            //                node.Nodes.Remove(wait_node);
            //                AddInterfaceNodes(node, clsid.Interfaces);
            //            }
            //            else
            //            {
            //                wait_node.Text = "Error querying COM interfaces - No Instance Interfaces";
            //            }

            //            if (factory_count > 0)
            //            {
            //                TreeNode factory = CreateNode("Factory Interfaces", FolderKey);
            //                AddInterfaceNodes(factory, clsid.FactoryInterfaces);
            //                node.Nodes.Add(factory);
            //            }
            //        }
            //    }
            //    catch (Win32Exception ex)
            //    {
            //        wait_node.Text = String.Format("Error querying COM interfaces - {0}", ex.Message);
            //    }
            //}
        }

        private async void treeComRegistry_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {            
            Cursor currCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            await SetupCLSIDNodeTree(e.Node, false);

            Cursor.Current = currCursor;
        }

        public enum CopyGuidType
        {
            CopyAsString,
            CopyAsStructure,
            CopyAsObject,
            CopyAsHexString,
        }

        public static void CopyTextToClipboard(string text)
        {
            int tries = 10;
            while (tries > 0)
            {
                try
                {
                    Clipboard.SetText(text);
                    break;
                }
                catch (ExternalException)
                {
                }
                System.Threading.Thread.Sleep(100);
                tries--;
            }
        }

        public static void CopyGuidToClipboard(Guid guid, CopyGuidType copyType)
        {
            string strCopy = null;

            switch (copyType)
            {
                case CopyGuidType.CopyAsObject:
                    strCopy = String.Format("<object id=\"obj\" classid=\"clsid:{0}\">NO OBJECT</object>",
                        guid.ToString());
                    break;
                case CopyGuidType.CopyAsString:
                    strCopy = guid.FormatGuid();
                    break;
                case CopyGuidType.CopyAsStructure:
                    {
                        strCopy = String.Format("GUID guidObject = {0:X};", guid);
                    }
                    break;
                case CopyGuidType.CopyAsHexString:
                    {
                        byte[] data = guid.ToByteArray();
                        strCopy = String.Join(" ", data.Select(b => String.Format("{0:X02}", b)));
                    }
                    break;
            }

            if (strCopy != null)
            {
                CopyTextToClipboard(strCopy);
            }
        }

        private static bool CanGetGuid(object obj)
        {
            if (obj is PlaceholderObject placeholder)
            {
                return placeholder.HasGuid;
            }

            return (obj is COMCLSIDEntry ||
                obj is COMInterfaceEntry ||
                obj is COMProgIDEntry ||
                obj is COMTypeLibVersionEntry ||
                obj is COMTypeLibEntry ||
                obj is Guid ||
                obj is COMAppIDEntry ||
                obj is COMIPIDEntry ||
                obj is COMCategory || 
                obj is COMRuntimeClassEntry ||
                obj is COMIELowRightsElevationPolicy);
        }

        private static Guid GetGuidFromType(object obj)
        {
            if (obj is COMCLSIDEntry clsid)
            {
                return clsid.Clsid;
            }
            else if (obj is COMInterfaceEntry intf)
            {
                return intf.Iid;
            }
            else if (obj is COMProgIDEntry progid)
            {
                return progid.Clsid;
            }
            else if (obj is COMTypeLibVersionEntry typelib_ver)
            {
                return typelib_ver.TypelibId;
            }
            else if (obj is COMTypeLibEntry typelib)
            {
                return typelib.TypelibId;
            }
            else if (obj is Guid guid)
            {
                return guid;
            }
            else if (obj is COMAppIDEntry appid)
            {
                return appid.AppId;
            }
            else if (obj is COMIPIDEntry ipid)
            {
                return ipid.Ipid;
            }
            else if (obj is COMCategory cat)
            {
                return cat.CategoryID;
            }
            else if (obj is COMRuntimeClassEntry runtime_class)
            {
                return runtime_class.Clsid;
            }
            else if (obj is PlaceholderObject placeholder && placeholder.HasGuid)
            {
                return placeholder.Guid;
            }
            else if (obj is COMIELowRightsElevationPolicy low_rights)
            {
                return low_rights.Uuid;
            }
            return Guid.Empty;
        }

        private void CopyGuid(CopyGuidType copy_type)
        {
            Guid guid = GetGuidFromType(treeListView.SelectedObject);
            if (guid != Guid.Empty)
            {
                CopyGuidToClipboard(guid, copy_type);
            }
        }

        private void copyGUIDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyGuid(CopyGuidType.CopyAsString);
        }

        private void copyGUIDCStructureToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyGuid(CopyGuidType.CopyAsStructure);
        }

        private void copyGUIDHexStringToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyGuid(CopyGuidType.CopyAsHexString);
        }

        private void copyObjectTagToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ICOMClassEntry obj = GetRealObject(treeListView.SelectedObject) as ICOMClassEntry;
            if (obj != null)
            {
                CopyGuidToClipboard(obj.Clsid, CopyGuidType.CopyAsObject);
            }
        }

        private async Task SetupObjectView(ICOMClassEntry ent, object obj, bool factory)
        {
            await Program.GetMainForm(m_registry).HostObject(ent, obj, factory);
        }

        private ICOMClassEntry GetSelectedClassEntry()
        {
            object obj = GetRealObject(treeListView.SelectedObject);
            if (obj is ICOMClassEntry class_entry)
            {
                return class_entry;
            }
            else if (obj is COMProgIDEntry progid)
            {
                return m_registry.MapClsidToEntry(progid.Clsid);
            }
            return null;
        }

        private async Task CreateInstance(CLSCTX clsctx, string server)
        {
            ICOMClassEntry ent = GetSelectedClassEntry();
            if (ent != null)
            {
                try
                {
                    object comObj = ent.CreateInstanceAsObject(clsctx, server);
                    if (comObj != null)
                    {
                        await SetupObjectView(ent, comObj, false);
                    }
                }
                catch (Exception ex)
                {
                    Program.ShowError(this, ex);
                }
            }
        }

        private async Task CreateClassFactory(string server)
        {
            ICOMClassEntry ent = GetSelectedClassEntry();
            if (ent != null)
            {
                try
                {
                    object comObj = ent.CreateClassFactory(server);
                    if (comObj != null)
                    {
                        await SetupObjectView(ent, comObj, true);
                    }
                }
                catch (Exception ex)
                {
                    Program.ShowError(this, ex);
                }
            }
        }

        private async void createInstanceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await CreateInstance(CLSCTX.ALL, null);
        }

        private void EnableViewPermissions(COMAppIDEntry appid)
        {
            if (appid.HasAccessPermission)
            {
                contextMenuStrip.Items.Add(viewAccessPermissionsToolStripMenuItem);
            }
            if (appid.HasLaunchPermission)
            {
                contextMenuStrip.Items.Add(viewLaunchPermissionsToolStripMenuItem);
            }
        }

        private void SetupCreateSpecialSessions()
        {
            createInSessionToolStripMenuItem.DropDownItems.Clear();
            createInSessionToolStripMenuItem.DropDownItems.Add(consoleToolStripMenuItem);
            foreach (int session_id in COMSecurity.GetSessionIds())
            {
                ToolStripMenuItem item = new ToolStripMenuItem(session_id.ToString());
                item.Tag = session_id.ToString();
                item.Click += consoleToolStripMenuItem_Click;
                createInSessionToolStripMenuItem.DropDownItems.Add(item);
            }
            createSpecialToolStripMenuItem.DropDownItems.Add(createInSessionToolStripMenuItem);
        }

        private static bool HasServerType(COMCLSIDEntry clsid, COMServerType type)
        {
            if (clsid == null)
            {
                return false;
            }

            if (clsid.DefaultServerType == COMServerType.UnknownServer)
            {
                // If we have no servers we assume anything is possible.
                return true;
            }

            return clsid.Servers.ContainsKey(type);
        }

        private void contextMenuStrip_Opening(object sender, CancelEventArgs e)
        {
            object obj = GetRealObject(treeListView.SelectedObject);
            if (obj is null)
            {
                e.Cancel = true;
                return;
            }

            contextMenuStrip.Items.Clear();
            contextMenuStrip.Items.Add(copyToolStripMenuItem);
            if (CanGetGuid(obj))
            {
                contextMenuStrip.Items.Add(copyGUIDToolStripMenuItem);
                contextMenuStrip.Items.Add(copyGUIDHexStringToolStripMenuItem);
                contextMenuStrip.Items.Add(copyGUIDCStructureToolStripMenuItem);
            }

            if ((obj is ICOMClassEntry) || (obj is COMProgIDEntry))
            {
                contextMenuStrip.Items.Add(copyObjectTagToolStripMenuItem);
                contextMenuStrip.Items.Add(createInstanceToolStripMenuItem);

                COMProgIDEntry progid = obj as COMProgIDEntry;
                COMCLSIDEntry clsid = obj as COMCLSIDEntry;
                COMRuntimeClassEntry runtime_class = obj as COMRuntimeClassEntry;
                ICOMClassEntry entry = obj as ICOMClassEntry;
                if (progid != null)
                {
                    clsid = m_registry.MapClsidToEntry(progid.Clsid);
                    entry = clsid;
                }

                createSpecialToolStripMenuItem.DropDownItems.Clear();

                if (HasServerType(clsid, COMServerType.InProcServer32))
                {
                    createSpecialToolStripMenuItem.DropDownItems.Add(createInProcServerToolStripMenuItem);
                }

                if (HasServerType(clsid, COMServerType.InProcHandler32))
                {
                    createSpecialToolStripMenuItem.DropDownItems.Add(createInProcHandlerToolStripMenuItem);
                }

                if (HasServerType(clsid, COMServerType.LocalServer32))
                {
                    createSpecialToolStripMenuItem.DropDownItems.Add(createLocalServerToolStripMenuItem);
                    SetupCreateSpecialSessions();
                    if (clsid.CanElevate)
                    {
                        createSpecialToolStripMenuItem.DropDownItems.Add(createElevatedToolStripMenuItem);
                    }
                    createSpecialToolStripMenuItem.DropDownItems.Add(createRemoteToolStripMenuItem);
                }

                createSpecialToolStripMenuItem.DropDownItems.Add(createClassFactoryToolStripMenuItem);
                if (entry != null && entry.SupportsRemoteActivation)
                {
                    createSpecialToolStripMenuItem.DropDownItems.Add(createClassFactoryRemoteToolStripMenuItem);
                }

                if (runtime_class != null && runtime_class.HasPermission)
                {
                    createSpecialToolStripMenuItem.DropDownItems.Add(createInRuntimeBrokerToolStripMenuItem);
                    createSpecialToolStripMenuItem.DropDownItems.Add(createInPerUserRuntimeBrokerToolStripMenuItem);
                    createSpecialToolStripMenuItem.DropDownItems.Add(createFactoryInRuntimeBrokerToolStripMenuItem);
                    createSpecialToolStripMenuItem.DropDownItems.Add(createFactoryInPerUserRuntimeBrokerToolStripMenuItem);
                }

                contextMenuStrip.Items.Add(createSpecialToolStripMenuItem);
                contextMenuStrip.Items.Add(refreshInterfacesToolStripMenuItem);

                if (clsid != null)
                {
                    if (m_registry.Typelibs.ContainsKey(clsid.TypeLib))
                    {
                        contextMenuStrip.Items.Add(viewTypeLibraryToolStripMenuItem);
                    }

                    if (m_registry.GetProxiesForClsid(clsid).Length > 0)
                    {
                        contextMenuStrip.Items.Add(viewProxyDefinitionToolStripMenuItem);
                    }

                    if (m_registry.AppIDs.ContainsKey(clsid.AppID))
                    {
                        EnableViewPermissions(m_registry.AppIDs[clsid.AppID]);
                    }
                }

                if (runtime_class != null)
                {
                    COMRuntimeServerEntry server =
                        runtime_class.HasServer
                            ? m_registry.MapServerNameToEntry(runtime_class.Server) : null;
                    if (runtime_class.HasPermission || (server != null && server.HasPermission))
                    {
                        contextMenuStrip.Items.Add(viewAccessPermissionsToolStripMenuItem);
                    }
                }
            }
            else if (obj is COMTypeLibVersionEntry)
            {
                contextMenuStrip.Items.Add(viewTypeLibraryToolStripMenuItem);
            }
            else if (obj is COMAppIDEntry)
            {
                EnableViewPermissions((COMAppIDEntry)obj);
            }
            else if (obj is COMInterfaceEntry)
            {
                COMInterfaceEntry intf = (COMInterfaceEntry)obj;
                if (intf.HasTypeLib)
                {
                    contextMenuStrip.Items.Add(viewTypeLibraryToolStripMenuItem);
                }

                if (intf.HasProxy && m_registry.Clsids.ContainsKey(intf.ProxyClsid))
                {
                    contextMenuStrip.Items.Add(viewProxyDefinitionToolStripMenuItem);
                }

                if (COMUtilities.RuntimeInterfaceMetadata.ContainsKey(intf.Iid))
                {
                    contextMenuStrip.Items.Add(viewRuntimeInterfaceToolStripMenuItem);
                }
            }
            else if (obj is COMProcessEntry)
            {
                contextMenuStrip.Items.Add(refreshProcessToolStripMenuItem);
                contextMenuStrip.Items.Add(viewAccessPermissionsToolStripMenuItem);
            }
            else if (obj is COMIPIDEntry ipid)
            {
                COMInterfaceEntry intf = m_registry.MapIidToInterface(ipid.Iid);

                if (intf.HasTypeLib)
                {
                    contextMenuStrip.Items.Add(viewTypeLibraryToolStripMenuItem);
                }

                if (intf.HasProxy && m_registry.Clsids.ContainsKey(intf.ProxyClsid))
                {
                    contextMenuStrip.Items.Add(viewProxyDefinitionToolStripMenuItem);
                }

                contextMenuStrip.Items.Add(unmarshalToolStripMenuItem);
            }
            else if (obj is COMRuntimeClassEntry runtime_class)
            {
                if (runtime_class.HasPermission)
                {
                    contextMenuStrip.Items.Add(viewAccessPermissionsToolStripMenuItem);
                }
            }
            else if (obj is COMRuntimeServerEntry runtime_server)
            {
                if (runtime_server.HasPermission)
                {
                    contextMenuStrip.Items.Add(viewAccessPermissionsToolStripMenuItem);
                }
            }

            if (m_filter_types.Contains(FilterType.CLSID))
            {
                contextMenuStrip.Items.Add(queryAllInterfacesToolStripMenuItem);
            }

            if (treeListView.GetItemCount() > 0)
            {
                contextMenuStrip.Items.Add(cloneTreeToolStripMenuItem);
                selectedToolStripMenuItem.Enabled = true;
            }

            if (PropertiesControl.SupportsProperties(obj))
            {
                contextMenuStrip.Items.Add(propertiesToolStripMenuItem);
            }
        }

        private void refreshInterfacesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //TreeNode node = treeComRegistry.SelectedNode;
            //if ((node != null) && (node.Tag != null))
            //{
            //    await SetupCLSIDNodeTree(node, true);
            //}
        }

        //private async void refreshInterfacesToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //TreeNode node = treeComRegistry.SelectedNode;
        //if ((node != null) && (node.Tag != null))
        //{
        //    await SetupCLSIDNodeTree(node, true);
        //}
        //}

        /// <summary>
        /// Convert a basic Glob to a regular expression
        /// </summary>
        /// <param name="glob">The glob string</param>
        /// <param name="ignoreCase">Indicates that match should ignore case</param>
        /// <returns>The regular expression</returns>
        private static Regex GlobToRegex(string glob, bool ignoreCase)
        {
            StringBuilder builder = new StringBuilder();

            builder.Append("^");

            foreach (char ch in glob)
            {
                if (ch == '*')
                {
                    builder.Append(".*");
                }
                else if (ch == '?')
                {
                    builder.Append(".");
                }
                else
                {
                    builder.Append(Regex.Escape(new String(ch, 1)));
                }
            }

            builder.Append("$");

            return new Regex(builder.ToString(), ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
        }

        private static Func<object, bool> CreatePythonFilter(string filter)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("from OleViewDotNet import *");
            builder.AppendLine("def run_filter(entry):");
            builder.AppendFormat("  return {0}", filter);
            builder.AppendLine();

            ScriptEngine engine = Python.CreateEngine();
            ScriptSource source = engine.CreateScriptSourceFromString(builder.ToString(), SourceCodeKind.File);
            ScriptScope scope = engine.CreateScope();
            scope.Engine.Runtime.LoadAssembly(typeof(COMCLSIDEntry).Assembly);
            source.Execute(scope);
            return scope.GetVariable<Func<object, bool>>("run_filter");
        }

        private static bool RunPythonFilter(object node, Func<object, bool> python_filter)
        {
            try
            {
                return python_filter(node);
            }
            catch 
            {
                return false;
            }
        }

        private FilterResult RunComplexFilter(object node, RegistryViewerFilter filter)
        {
            try
            {
                COMCLSIDEntry clsid = node as COMCLSIDEntry;
                FilterResult result = filter.Filter(node);
                if (result == FilterResult.None && clsid != null && clsid.InterfacesLoaded)
                {
                    foreach (COMInterfaceEntry intf in clsid.Interfaces.Concat(clsid.FactoryInterfaces).Select(i => m_registry.MapIidToInterface(i.Iid)))
                    {
                        result = filter.Filter(intf);
                        if (result != FilterResult.None)
                        {
                            break;
                        }
                    }
                }
                return result;
            }
            catch
            {
                return FilterResult.None;
            }
        }

        private FilterResult RunAccessibleFilter(object node, 
            Dictionary<string, bool> access_cache, 
            Dictionary<string, bool> launch_cache, 
            NtToken token, 
            string principal,
            COMAccessRights access_rights, 
            COMAccessRights launch_rights)
        {
            string launch_sddl = m_registry.DefaultLaunchPermission;
            string access_sddl = m_registry.DefaultAccessPermission;
            bool check_launch = true;

            if (node is COMProcessEntry process)
            {
                access_sddl = process.AccessPermissions;
                principal = process.UserSid;
                check_launch = false;
            }
            else if (node is COMAppIDEntry || node is COMCLSIDEntry)
            {
                COMAppIDEntry appid = node as COMAppIDEntry;
                if (appid == null && node is COMCLSIDEntry)
                {
                    COMCLSIDEntry clsid = (COMCLSIDEntry)node;
                    if (!m_registry.AppIDs.ContainsKey(clsid.AppID))
                    {
                        return FilterResult.Exclude;
                    }

                    appid = m_registry.AppIDs[clsid.AppID];
                }

                if (appid.HasLaunchPermission)
                {
                    launch_sddl = appid.LaunchPermission;
                }
                if (appid.HasAccessPermission)
                {
                    access_sddl = appid.AccessPermission;
                }
            }
            else if (node is COMRuntimeClassEntry runtime_class)
            {
                if (runtime_class.HasPermission)
                {
                    launch_sddl = runtime_class.Permissions;
                }
                else
                {
                    // Set to denied access.
                    launch_sddl = "O:SYG:SYD:";
                }
                access_sddl = launch_sddl;
            }
            else if (node is COMRuntimeServerEntry runtime_server)
            {
                if (runtime_server.HasPermission)
                {
                    launch_sddl = runtime_server.Permissions;
                }
                else
                {
                    launch_sddl = "O:SYG:SYD:";
                }
                access_sddl = launch_sddl;
            }
            else
            {
                return FilterResult.Exclude;
            }
            
            if (!access_cache.ContainsKey(access_sddl))
            {
                if (access_rights == 0)
                {
                    access_cache[access_sddl] = true;
                }
                else
                {
                    access_cache[access_sddl] = COMSecurity.IsAccessGranted(access_sddl, principal, token, false, false, access_rights);
                }
            }

            if (check_launch && !launch_cache.ContainsKey(launch_sddl))
            {
                if (launch_rights == 0)
                {
                    launch_cache[launch_sddl] = true;
                }
                else
                {
                    launch_cache[launch_sddl] = COMSecurity.IsAccessGranted(launch_sddl, principal, token,
                        true, true, access_rights);
                }
            }

            if (access_cache[access_sddl] && (!check_launch || launch_cache[launch_sddl]))
            {
                return FilterResult.Include;
            }
            return FilterResult.Exclude;
        }

        private enum FilterMode
        {
            Contains,
            BeginsWith,
            EndsWith,
            Equals,
            Glob,
            Regex,
            Python,
            Accessible,
            NotAccessible,
            Complex,
        }

        private Func<object, bool> CreateFilter(string filter, FilterMode mode, bool caseSensitive)
        {
            StringComparison comp;

            filter = filter.Trim();
            if (String.IsNullOrEmpty(filter))
            {
                return null;
            }

            if(caseSensitive)
            {
                comp = StringComparison.CurrentCulture;
            }
            else
            {
                comp = StringComparison.CurrentCultureIgnoreCase;
            }

            switch (mode)
            {
                case FilterMode.Contains:
                    if (caseSensitive)
                    {
                        return n => GetObjectName(n).Contains(filter);
                    }
                    else
                    {
                        filter = filter.ToLower();
                        return n => GetObjectName(n).ToLower().Contains(filter.ToLower());
                    }
                case FilterMode.BeginsWith:
                    return n => GetObjectName(n).StartsWith(filter, comp);
                case FilterMode.EndsWith:
                    return n => GetObjectName(n).EndsWith(filter, comp);
                case FilterMode.Equals:
                    return n => GetObjectName(n).Equals(filter, comp);
                case FilterMode.Glob:
                    {
                        Regex r = GlobToRegex(filter, caseSensitive);

                        return n => r.IsMatch(GetObjectName(n));
                    }
                case FilterMode.Regex:
                    {
                        Regex r = new Regex(filter, caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase);

                        return n => r.IsMatch(GetObjectName(n));
                    }
                case FilterMode.Python:
                    {
                        Func<object, bool> python_filter = CreatePythonFilter(filter);

                        return n => RunPythonFilter(n, python_filter);
                    }
                default:
                    throw new ArgumentException("Invalid mode value");
            }
        }

        // Check if top node or one of its subnodes matches the filter
        private FilterResult FilterNode(object obj, Func<object, FilterResult> filterFunc)
        {
            FilterResult result = filterFunc(obj);

            if (result == FilterResult.None)
            {
                foreach (object sub_obj in GetChildren(obj))
                {
                    result = FilterNode(sub_obj, filterFunc);
                    if (result == FilterResult.Include)
                    {
                        break;
                    }
                }
            }

            return result;
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            NtToken token = null;
            try
            {
                Func<object, FilterResult> filterFunc = null;
                FilterMode mode = (FilterMode)comboBoxMode.SelectedItem;

                if (mode == FilterMode.Complex)
                {
                    using (ViewFilterForm form = new ViewFilterForm(m_filter, m_filter_types))
                    {
                        if (form.ShowDialog(this) == DialogResult.OK)
                        {
                            m_filter = form.Filter;
                            if (m_filter.Filters.Count > 0)
                            {
                                filterFunc = n => RunComplexFilter(n, m_filter);
                            }
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                else if (mode == FilterMode.Accessible || mode == FilterMode.NotAccessible)
                {
                    using (SelectSecurityCheckForm form = new SelectSecurityCheckForm(m_mode == DisplayMode.Processes))
                    {
                        if (form.ShowDialog(this) == DialogResult.OK)
                        {
                            token = form.Token;
                            string principal = form.Principal;
                            COMAccessRights access_rights = form.AccessRights;
                            COMAccessRights launch_rights = form.LaunchRights;
                            Dictionary<string, bool> access_cache = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                            Dictionary<string, bool> launch_cache = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                            filterFunc = n => RunAccessibleFilter(n, access_cache, launch_cache, token, principal, access_rights, launch_rights);
                            if (mode == FilterMode.NotAccessible)
                            {
                                Func<object, FilterResult> last_filter = filterFunc;
                                filterFunc = n => last_filter(n) == FilterResult.Exclude ? FilterResult.Include : FilterResult.Exclude;
                            }
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                else
                {
                    Func<object, bool> filter = CreateFilter(textBoxFilter.Text, mode, false);
                    if (filter != null)
                    {
                        filterFunc = n => filter(n) ? FilterResult.Include : FilterResult.None;
                    }
                }

                if (filterFunc == null)
                {
                    treeListView.UseFiltering = false;
                    treeListView.ListFilter = null;
                }
                else
                {
                    treeListView.UseFiltering = true;
                    treeListView.ModelFilter = new ModelFilter(o => filterFunc(o) == FilterResult.Include);
                }

                UpdateStatusLabel();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (token != null)
                {
                    token.Close();
                }
            }
        }

        private void textBoxFilter_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return))
            {
                btnApply.PerformClick();
            }
        }

        private void viewTypeLibraryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            object obj = GetRealObject(treeListView.SelectedObject);
            COMTypeLibVersionEntry ent = obj as COMTypeLibVersionEntry;
            Guid selected_guid = Guid.Empty;

            if (ent == null)
            {
                COMCLSIDEntry clsid = obj as COMCLSIDEntry;
                COMProgIDEntry progid = obj as COMProgIDEntry;
                COMInterfaceEntry intf = obj as COMInterfaceEntry;
                if (progid != null)
                {
                    clsid = m_registry.MapClsidToEntry(progid.Clsid);
                }

                if (clsid != null && m_registry.Typelibs.ContainsKey(clsid.TypeLib))
                {
                    ent = m_registry.Typelibs[clsid.TypeLib].Versions.First();
                    selected_guid = clsid.Clsid;
                }

                if (intf != null && m_registry.Typelibs.ContainsKey(intf.TypeLib))
                {
                    ent = m_registry.GetTypeLibVersionEntry(intf.TypeLib, intf.TypeLibVersion);
                    selected_guid = intf.Iid;
                }
            }

            if (ent != null)
            {
                Assembly typelib = COMUtilities.LoadTypeLib(this, ent.NativePath);
                if (typelib != null)
                {
                    Program.GetMainForm(m_registry).HostControl(new TypeLibControl(ent.Name, typelib, selected_guid, false));
                }
            }
        }

        private void OpenProperties()
        {
            object obj = GetRealObject(treeListView.SelectedObject);
            if (PropertiesControl.SupportsProperties(obj))
            {
                Program.GetMainForm(m_registry).HostControl(new PropertiesControl(m_registry, GetObjectName(obj), obj));
            }
        }

        private void propertiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenProperties();
        }

        private void ViewPermissions(bool access)
        {
            object obj = GetRealObject(treeListView.SelectedObject);
            if (obj != null)
            {
                if (obj is COMProcessEntry proc)
                {
                    COMSecurity.ViewSecurity(m_registry, String.Format("{0} Access", proc.Name), proc.AccessPermissions, true);
                }
                else if (obj is COMRuntimeClassEntry || obj is COMRuntimeServerEntry)
                {
                    COMRuntimeServerEntry runtime_server = obj as COMRuntimeServerEntry;
                    COMRuntimeClassEntry runtime_class = obj as COMRuntimeClassEntry;
                    string name = runtime_class != null ? runtime_class.Name : runtime_server.Name;
                    if (runtime_class != null && runtime_class.HasServer)
                    {
                        runtime_server = m_registry.MapServerNameToEntry(runtime_class.Server);
                    }

                    string perms = runtime_server != null ? runtime_server.Permissions : runtime_class.Permissions;

                    COMSecurity.ViewSecurity(m_registry, string.Format("{0} Access", name), perms, false);
                }
                else
                {
                    COMAppIDEntry appid = obj as COMAppIDEntry;
                    if (appid == null)
                    {
                        COMCLSIDEntry clsid = obj as COMCLSIDEntry;
                        if (clsid == null && obj is COMProgIDEntry progid)
                        {
                            clsid = m_registry.MapClsidToEntry(progid.Clsid);
                        }

                        if (clsid != null && m_registry.AppIDs.ContainsKey(clsid.AppID))
                        {
                            appid = m_registry.AppIDs[clsid.AppID];
                        }
                    }

                    if (appid != null)
                    {
                        COMSecurity.ViewSecurity(m_registry, appid, access);
                    }
                }
            }
        }

        private void viewLaunchPermissionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ViewPermissions(false);
        }

        private void viewAccessPermissionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ViewPermissions(true);
        }

        private async void createLocalServerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await CreateInstance(CLSCTX.LOCAL_SERVER, null);
        }

        private async void createInProcServerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await CreateInstance(CLSCTX.INPROC_SERVER, null);
        }

        private async Task CreateFromMoniker(COMCLSIDEntry ent, string moniker)
        {
            try
            {
                object obj = COMUtilities.CreateFromMoniker(moniker, CLSCTX.LOCAL_SERVER);
                await SetupObjectView(ent, obj, obj is IClassFactory);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task CreateInSession(COMCLSIDEntry ent, string session_id)
        {
            await CreateFromMoniker(ent, String.Format("session:{0}!new:{1}", session_id, ent.Clsid));
        }

        private async Task CreateElevated(COMCLSIDEntry ent, bool factory)
        {
            await CreateFromMoniker(ent, String.Format("Elevation:Administrator!{0}:{1}", 
                factory ? "clsid" : "new", ent.Clsid));
        }

        private async void consoleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem item = sender as ToolStripMenuItem;
            COMCLSIDEntry ent = GetSelectedClassEntry() as COMCLSIDEntry;
            if (ent != null && item != null && item.Tag is string)
            {
                await CreateInSession(ent, (string)item.Tag);
            }
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            object obj = GetRealObject(treeListView.SelectedObject);
            CopyTextToClipboard(GetObjectName(obj));
        }

        private void viewProxyDefinitionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            object node = GetRealObject(treeListView.SelectedObject);
            if (node != null)
            {
                COMCLSIDEntry clsid = node as COMCLSIDEntry;
                Guid selected_iid = Guid.Empty;
                if (clsid == null && (node is COMInterfaceEntry || node is COMIPIDEntry))
                {
                    COMInterfaceEntry intf = node as COMInterfaceEntry;
                    if (intf == null)
                    {
                        intf = m_registry.MapIidToInterface(((COMIPIDEntry)node).Iid);
                    }

                    selected_iid = intf.Iid;
                    clsid = m_registry.Clsids[intf.ProxyClsid];
                }

                if (clsid != null)
                {
                    try
                    {
                        using (var resolver = Program.GetProxyParserSymbolResolver())
                        {
                            Program.GetMainForm(m_registry).HostControl(new TypeLibControl(m_registry,
                                Path.GetFileName(clsid.DefaultServer),
                                COMProxyInstance.GetFromCLSID(clsid, resolver), selected_iid));
                        }
                    }
                    catch (Exception ex)
                    {
                        Program.ShowError(this, ex);
                    }
                }
            }
        }

        private async void createClassFactoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await CreateClassFactory(null);
        }

        private void GetClsidsFromNodes(List<COMCLSIDEntry> clsids, TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Tag is COMCLSIDEntry)
                {
                    clsids.Add((COMCLSIDEntry)node.Tag);
                }

                if (node.Nodes.Count > 0)
                {
                    GetClsidsFromNodes(clsids, node.Nodes);
                }
            }
        }

        private void queryAllInterfacesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //using (QueryInterfacesOptionsForm options = new QueryInterfacesOptionsForm())
            //{
            //    if (options.ShowDialog(this) == DialogResult.OK)
            //    {
            //        List<COMCLSIDEntry> clsids = new List<COMCLSIDEntry>();
            //        GetClsidsFromNodes(clsids, treeComRegistry.Nodes);
            //        if (clsids.Count > 0)
            //        {
            //            COMUtilities.QueryAllInterfaces(this, clsids,
            //                options.ServerTypes, options.ConcurrentQueries,
            //                options.RefreshInterfaces);
            //        }
            //    }
            //}
        }

        private async void createInProcHandlerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await CreateInstance(CLSCTX.INPROC_HANDLER, null);
        }

        private async void instanceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            COMCLSIDEntry clsid = GetSelectedClassEntry() as COMCLSIDEntry;
            if (clsid != null)
            {
                await CreateElevated(clsid, false);
            }
        }

        private async void classFactoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            COMCLSIDEntry clsid = GetSelectedClassEntry() as COMCLSIDEntry;
            if (clsid != null)
            {
                await CreateElevated(clsid, true);
            }
        }

        private void comboBoxMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxMode.SelectedItem != null)
            {
                FilterMode mode = (FilterMode)comboBoxMode.SelectedItem;
                textBoxFilter.Enabled = mode != FilterMode.Complex && mode != FilterMode.Accessible && mode != FilterMode.NotAccessible;
            }
        }

        private async void createRemoteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (GetTextForm frm = new GetTextForm("localhost"))
            {
                frm.Text = "Enter Remote Host";
                if (frm.ShowDialog(this) == DialogResult.OK)
                {
                    await CreateInstance(CLSCTX.REMOTE_SERVER, frm.Data);
                }
            }
        }

        private async void createClassFactoryRemoteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (GetTextForm frm = new GetTextForm("localhost"))
            {
                frm.Text = "Enter Remote Host";
                if (frm.ShowDialog(this) == DialogResult.OK)
                {
                    await CreateClassFactory(frm.Data);
                }
            }
        }

        private void refreshProcessToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //TreeNode node = treeComRegistry.SelectedNode;
            //if (node != null && node.Tag is COMProcessEntry)
            //{
            //    COMProcessEntry process = (COMProcessEntry)node.Tag;
            //    process = COMProcessParser.ParseProcess(process.Pid, COMUtilities.GetProcessParserConfig(), m_registry);
            //    if (process == null)
            //    {
            //        treeComRegistry.Nodes.Remove(treeComRegistry.SelectedNode);
            //        m_originalNodes = m_originalNodes.Where(n => n != node).ToArray();
            //    }
            //    else
            //    {
            //        node.Tag = process;
            //        node.Nodes.Clear();
            //        PopulatorIpids(m_registry, node, process);
            //    }
            //}
        }

        private COMIPIDEntry GetSelectedIpid()
        {
            return GetRealObject(treeListView.SelectedObject) as COMIPIDEntry;
        }

        private void toHexEditorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            COMIPIDEntry ipid = GetSelectedIpid();
            if (ipid != null)
            {
                Program.GetMainForm(m_registry).HostControl(new ObjectHexEditor(m_registry, 
                    ipid.Ipid.ToString(),
                    ipid.ToObjref()));
            }
        }

        private void toFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            COMIPIDEntry ipid = GetSelectedIpid();
            if (ipid != null)
            {
                using (SaveFileDialog dlg = new SaveFileDialog())
                {
                    dlg.Filter = "All Files (*.*)|*.*";
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        try
                        {
                            File.WriteAllBytes(dlg.FileName, ipid.ToObjref());
                        }
                        catch (Exception ex)
                        {
                            Program.ShowError(this, ex);
                        }
                    }
                }
            }
        }

        private async void toObjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            COMIPIDEntry ipid = GetSelectedIpid();
            if (ipid != null)
            {
                try
                {
                    await Program.GetMainForm(m_registry).OpenObjectInformation(
                        COMUtilities.UnmarshalObject(ipid.ToObjref()),
                        String.Format("IPID {0}", ipid.Ipid));
                }
                catch (Exception ex)
                {
                    Program.ShowError(this, ex);
                }
            }
        }

        private void CreateClonedTree(IEnumerable nodes)
        {
            string text = Text;
            if (!text.StartsWith("Clone of "))
            {
                text = "Clone of " + text;
            }
            COMRegistryViewer viewer = new COMRegistryViewer(m_registry, m_mode, m_processes, nodes.Cast<object>(), m_filter_types, text);
            Program.GetMainForm(m_registry).HostControl(viewer);
        }

        private void allVisibleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateClonedTree(treeListView.FilteredObjects);
        }

        private void selectedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            object obj = treeListView.SelectedObject;
            if (obj != null)
            {
                CreateClonedTree(new object[] { obj });
            }
        }

        private async void filteredToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (ViewFilterForm form = new ViewFilterForm(new RegistryViewerFilter(), m_filter_types))
            {
                if (form.ShowDialog(this) == DialogResult.OK && form.Filter.Filters.Count > 0)
                {
                    IEnumerable<object> original_nodes = treeListView.Objects.Cast<object>();
                    IEnumerable<object> nodes =
                        await Task.Run(() => original_nodes.Where(n =>
                        FilterNode(n, x => RunComplexFilter(x, form.Filter)) == FilterResult.Include).ToArray());
                    CreateClonedTree(nodes);
                }
            }
        }

        private void allChildrenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var children = GetChildren(treeListView.SelectedObject);
            if (children.Any())
            {
                CreateClonedTree(children);
            }
        }

        private void viewRuntimeInterfaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (GetRealObject(treeListView.SelectedObject) is COMInterfaceEntry ent)
            {
                if (COMUtilities.RuntimeInterfaceMetadata.ContainsKey(ent.Iid))
                {
                    Assembly asm = COMUtilities.RuntimeInterfaceMetadata[ent.Iid].Assembly;
                    Program.GetMainForm(m_registry).HostControl(new TypeLibControl(asm.GetName().Name,
                        COMUtilities.RuntimeInterfaceMetadata[ent.Iid].Assembly, ent.Iid, false));
                }
            }
        }

        [Guid("D63B10C5-BB46-4990-A94F-E40B9D520160")]
        [ComImport]
        class RuntimeBroker
        {
        }

        [Guid("2593F8B9-4EAF-457C-B68A-50F6B8EA6B54")]
        [ComImport]
        class PerUserRuntimeBroker
        {
        }

        private IRuntimeBroker CreateBroker(bool per_user)
        {
            if (per_user)
            {
                return (IRuntimeBroker)new PerUserRuntimeBroker();
            }
            else
            {
                return (IRuntimeBroker)new RuntimeBroker();
            }
        }

        private async void CreateInRuntimeBroker(bool per_user, bool factory)
        {
            try
            {
                COMRuntimeClassEntry runtime_class = GetSelectedClassEntry() as COMRuntimeClassEntry;
                if (runtime_class != null)
                {
                    IRuntimeBroker broker = CreateBroker(per_user);
                    object comObj;
                    if (factory)
                    {
                        Guid iid = COMInterfaceEntry.IID_IUnknown;
                        comObj = broker.GetActivationFactory(runtime_class.Name, ref iid);
                    }
                    else
                    {
                        comObj = broker.ActivateInstance(runtime_class.Name);
                    }

                    await SetupObjectView(runtime_class, comObj, factory);
                }
            }
            catch (Exception ex)
            {
                Program.ShowError(this, ex);
            }
        }

        private void createInRuntimeBrokerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateInRuntimeBroker(false, false);
        }

        private void createInPerUserRuntimeBrokerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateInRuntimeBroker(true, false);
        }

        private void createFactoryInRuntimeBrokerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateInRuntimeBroker(false, true);
        }

        private void createFactoryInPerUserRuntimeBrokerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateInRuntimeBroker(false, true);
        }

        private void treeListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            Program.GetMainForm(m_registry).UpdatePropertyGrid(GetRealObject(treeListView.SelectedObject));
        }

        private void treeListView_DoubleClick(object sender, EventArgs e)
        {
            if (ModifierKeys == Keys.Control)
            {
                OpenProperties();
            }
            else
            {
                object obj = treeListView.SelectedObject;
                if (treeListView.IsExpanded(obj))
                {
                    treeListView.Collapse(obj);
                }
                else
                {
                    treeListView.Expand(obj);
                }
            }
        }
    }
}
