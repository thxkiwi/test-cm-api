using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using Windows.Win32;
using Windows.Win32.Devices.DeviceAndDriverInstallation;
using Windows.Win32.Devices.Properties;

namespace THX
{
    public class DeviceNode : IDisposable
    {
        private static char[] _StringListSeparators = new char[] { '\0' };

        static private Dictionary<uint, DeviceNode>   _deviceNodesByDevInst = new();
        static private Dictionary<string, DeviceNode> _deviceNodesByInstanceId = new();
        static private Dictionary<string, HashSet<DeviceNode>> _deviceNodesByHardwareId = new();

        /// \details Locks on 
        /// <ol>
        /// <li>_deviceNodesByInstanceId</li>
        /// <li>_deviceNodesByHardwareId</li>
        /// <li>_deviceNodesByDevInst</li>
        /// </ol>
        /// must be acquired in that order before calling this method.
        static private void Cache(DeviceNode node)
        {
            _deviceNodesByDevInst[node._devNode] = node;
            _deviceNodesByInstanceId[node.InstanceID] = node;

            foreach (string hardwareId in node.GetProperty<List<string>>(PInvoke.DEVPKEY_Device_HardwareIds)
                .Where(id => null != id))
            {
                var nodes = _deviceNodesByHardwareId.GetValueOrDefault(hardwareId, new());
                nodes.Add(node);
                _deviceNodesByHardwareId[hardwareId] = nodes;
            }
        }

        static public HashSet<DeviceNode> GetByHardwareId(string hardwareId)
        {
            lock (_deviceNodesByInstanceId)
            lock (_deviceNodesByHardwareId)
            lock (_deviceNodesByDevInst)
            {
                if (0 == _deviceNodesByHardwareId.Count)
                {
                    /// \todo Is there a better way to get all device IDs?
                    var dcs =  THX.DeviceClass.GetClasses(CM_ENUMERATE_FLAGS.CM_ENUMERATE_CLASSES_INSTALLER);
                    foreach (var dc in dcs)
                    {
                        var diids = THX.DeviceClass.GetDeviceIds(dc);
                        foreach (var diid in diids)
                        {
                            try
                            {
                                DeviceNode n = new (diid, PInvoke.DEVPKEY_Device_InstanceId);
                                Cache(n);
                            }
                            catch (KeyNotFoundException)
                            { }
                        }
                    }
                }

                if (0 == _deviceNodesByHardwareId.Count)
                {
                    throw new InvalidOperationException("_deviceNodesByHardwareId.Count == 0");
                }

                var deviceNodes = _deviceNodesByHardwareId[hardwareId];
                return deviceNodes;
            }
        }

        static public DeviceNode GetByInstanceId(string deviceId)
        {
            lock (_deviceNodesByInstanceId)
            lock (_deviceNodesByHardwareId)
            lock (_deviceNodesByDevInst)
            {
                if (_deviceNodesByInstanceId.TryGetValue(deviceId, out DeviceNode? node))
                {
                    return node;
                }
                unsafe
                {
                    fixed (char* usDeviceID = deviceId)
                    {
                        uint devNode = 0;
                        Windows.Win32.Foundation.PWSTR devNodeName = usDeviceID;
                        var ret = PInvoke.CM_Locate_DevNode(&devNode, devNodeName, CM_LOCATE_DEVNODE_FLAGS.CM_LOCATE_DEVNODE_PHANTOM);
                        if (CONFIGRET.CR_SUCCESS != ret)
                        {
                            throw new KeyNotFoundException($"CM_Locate_DevNode {deviceId} failed with {ret}");
                        }

                        DeviceNode deviceNode = new(devNode);
                        Cache(deviceNode);
                        return deviceNode;
                    }
                }
            }
        }

        static private uint GetDnInst(string deviceId, DEVPROPKEY propKey)
        {
            if (PInvoke.DEVPKEY_Device_InstanceId.Equals(propKey))
            {
                return GetByInstanceId(deviceId)._devNode;
            }

            if (PInvoke.DEVPKEY_Device_HardwareIds.Equals(propKey))
            {
                HashSet<DeviceNode> deviceNodes = GetByHardwareId(deviceId);
                if (deviceNodes.Count != 1)
                {
                    throw new InvalidOperationException($"GetByHardwareId {deviceId} returned {deviceNodes.Count} devices");
                }

                return deviceNodes.First()._devNode;
            }

            return 0;
        }

        private uint _devNode;

        private RegistryKey? _softwareKey;

        public RegistryKey SoftwareKey
        {
            get
            {
                if (_softwareKey != null)
                {
                    return _softwareKey;
                }

                unsafe
                {
                    SafeRegistryHandle hkey;
                    CONFIGRET ret;
                    ret = PInvoke.CM_Open_DevNode_Key(_devNode,
                        (uint)Windows.Win32.System.Registry.REG_SAM_FLAGS.KEY_READ,
                        0, // Use current hardware profile
                        PInvoke.RegDisposition_OpenExisting,
                        out hkey,
                        PInvoke.CM_REGISTRY_SOFTWARE);
                    switch (ret)
                    {
                        case CONFIGRET.CR_SUCCESS:
                            _softwareKey = RegistryKey.FromHandle(hkey);
                            return _softwareKey;
                        default:
                            throw new InvalidOperationException($"CM_Open_DevNode_Key failed with {ret}");
                    }
                }

            }
        }

        public string InstanceID { get; }

        public string FriendlyName { get; }


        private Dictionary<DEVPROPKEY, DEVPROPERTY> _availableProperties = new();

        public Guid DeviceClass => GetProperty<Guid>(PInvoke.DEVPKEY_Device_ClassGuid);

        public List<DeviceNode> Siblings
        {
            get
            {
                try
                {
                    return GetProperty<List<string>>(PInvoke.DEVPKEY_Device_Siblings)
                        .Select(sibling => new DeviceNode(sibling, PInvoke.DEVPKEY_Device_InstanceId)).ToList();
                }
                catch (KeyNotFoundException)
                {
                    return new List<DeviceNode>();
                }
            }
        }

        private List<DeviceNode> _children = new();
        public List<DeviceNode> Children
        {
            get
            {
                if (0 < _children.Count)
                {
                    return _children;
                }

                try
                {
                    _children = GetProperty<List<string>>(PInvoke.DEVPKEY_Device_Children)
                        .Select(child => new DeviceNode(child, PInvoke.DEVPKEY_Device_InstanceId)).ToList();
                    return _children;
                }
                catch (KeyNotFoundException)
                {
                    return _children = new List<DeviceNode>();
                }
            }
        }

        internal List<DEVPROPKEY> PropertyKeys
        {
            get
            {
                List<DEVPROPKEY> result = new List<DEVPROPKEY>();
                unsafe
                {
                    uint sz = 0;
                    PInvoke.CM_Get_DevNode_Property_Keys(
                                            _devNode,
                                            null,
                                            &sz,
                                            0);
                    if (0 == sz)
                    {
                        return new List<DEVPROPKEY>();
                    }

                    DEVPROPKEY* keys = stackalloc DEVPROPKEY[(int)sz];
                    PInvoke.CM_Get_DevNode_Property_Keys(_devNode,
                                                            keys,
                                                            &sz,
                                                            0);
                    for (int i = 0; i < sz; i++)
                    {
                        result.Add(keys[i]);
                    }
                }
                return result;
            }
        }

        public List<ExtensionInf> ExtensionInfs
        {
            get
            {
                try
                {
                    return GetProperty<List<string>>(PInvoke.DEVPKEY_Device_ExtendedConfigurationIds)
                        .Select(ExtensionInf.Parse)
                        .ToList();
                }
                catch (KeyNotFoundException)
                {
                    return new List<ExtensionInf>();
                }
            }
        }

        private List<AudioEndpoint> _audioEndpoints = new();

        public List<AudioEndpoint> AudioEndpoints
        {
            get
            {
                if (0 < _audioEndpoints.Count)
                {
                    return _audioEndpoints;
                }

                _audioEndpoints = Children
                    .Where(child => child.DeviceClass.Equals(THX.DeviceClass.AudioEndpoint))
                    .Select(child => new AudioEndpoint(child))
                    .ToList();
                return _audioEndpoints;
            }
        }

        private List<AudioProcessingObjectInf> _audioProcessingObjectInfs = new();

        public List<AudioProcessingObjectInf> AudioProcessingObjectInfs
        {
            get
            {
                if (0 < _audioProcessingObjectInfs.Count)
                {
                    return _audioProcessingObjectInfs;
                }

                _audioProcessingObjectInfs = Children
                    .Where(child => child.DeviceClass.Equals(THX.DeviceClass.AudioProcessingObject))
                    .Select(child => new AudioProcessingObjectInf(child))
                    .ToList();
                return _audioProcessingObjectInfs;
            }
        }

        ///
        internal DeviceNode(string deviceId,
                          DEVPROPKEY idPropertyKey)
            : this(GetDnInst(deviceId, idPropertyKey))
        { }

        public DeviceNode(uint devNode)
        {
            _devNode = devNode;

            foreach (var propKey in PropertyKeys)
            {
                uint sz = 0;
                unsafe
                {
                    DEVPROPTYPE propType = new DEVPROPTYPE();
                    PInvoke.CM_Get_DevNode_Property(
                        _devNode,
                        &propKey,
                        &propType,
                        null,
                        &sz,
                        0);

                    nint buffer = Marshal.AllocHGlobal((int)sz);
                    PInvoke.CM_Get_DevNode_Property(
                        _devNode,
                        &propKey,
                        &propType,
                        (byte*)buffer,
                        &sz,
                        0);

                    DEVPROPERTY prop = new DEVPROPERTY
                    {
                        Type = propType,
                        Buffer = (void*)buffer,
                        BufferSize = sz
                    };

                    _availableProperties.Add(propKey, prop);
                }
            }

            InstanceID = GetProperty<string>(PInvoke.DEVPKEY_Device_InstanceId);

            try
            {
                FriendlyName = GetProperty<string>(PInvoke.DEVPKEY_Device_FriendlyName);
            }
            catch (KeyNotFoundException)
            {
                FriendlyName = "<Not Found>";
            }

            Cache(this);
        }

        ~DeviceNode()
        {
            // Call Dispose() from the IDisposable interface.
            Dispose();
        }


        public void Dispose()
        {
            if (null != _availableProperties)
            {
                unsafe
                {
                    foreach (var prop in _availableProperties)
                    {
                        Marshal.FreeHGlobal((nint)prop.Value.Buffer);
                    }
                }
                _availableProperties.Clear();
            }
        }

        override public string ToString()
        {
            return $"{_devNode} {InstanceID} ({FriendlyName})";
        }

        public override bool Equals(object? obj)
        {
            return obj != null
                && obj is DeviceNode node
                && _devNode == node._devNode;
        }
        public override int GetHashCode()
        {
            return _devNode.GetHashCode();
        }

        private static readonly Dictionary<string, DEVPROPKEY> _headerProperties = new()
        {
            { "Friendly Name", PInvoke.DEVPKEY_Device_FriendlyName },
            { "Description", PInvoke.DEVPKEY_Device_DeviceDesc  },
            { "Device", PInvoke.DEVPKEY_Device_InstanceId},
            { "Class", PInvoke.DEVPKEY_Device_Class },
            { "Class GUID", PInvoke.DEVPKEY_Device_ClassGuid },
            { "Hardware IDs", PInvoke.DEVPKEY_Device_HardwareIds },
            { "Compatible IDs", PInvoke.DEVPKEY_Device_CompatibleIds },
            { "Location Info", PInvoke.DEVPKEY_Device_LocationInfo },
            { "Manufacturer", PInvoke.DEVPKEY_Device_Manufacturer },
            { "Provider", PInvoke.DEVPKEY_Device_DriverProvider },
            { "Driver", PInvoke.DEVPKEY_Device_Driver },
            { "Driver Description", PInvoke.DEVPKEY_Device_DriverDesc },
            { "Is Present", PInvoke.DEVPKEY_Device_IsPresent },
            { "Is Reboot Required", PInvoke.DEVPKEY_Device_IsRebootRequired },
            { "Install State", PInvoke.DEVPKEY_Device_InstallState },
            { "Has Problem?", PInvoke.DEVPKEY_Device_HasProblem },
            { "Problem Code", PInvoke.DEVPKEY_Device_ProblemCode },
            { "Problem Status", PInvoke.DEVPKEY_Device_ProblemStatus },
            { "Driver INF Path", PInvoke.DEVPKEY_Device_DriverInfPath },
            { "Children", PInvoke.DEVPKEY_Device_Children },
            { "Siblings", PInvoke.DEVPKEY_Device_Siblings }
        };

        private static readonly int _width = (int)(_headerProperties.Keys.ToList()
                .Concat(new[] { "Available Properties", "Extensions", "Children", "AudioProcessingObjects", "Unresolved AudioProcessingObjects" })
                .Max(s => (uint)s.Length) + 1);


        public void WriteDetailed(TextWriter writer, string indent = "")
        {
            writer.WriteLine($"{indent}{this}");
            foreach (var kv in _headerProperties)
            {
                try
                {
                    DEVPROPERTY prop = GetProperty(kv.Value);
                    writer.WriteLine($"{indent}{kv.Key.PadRight(_width)}: {Extension.GetPropertyString(prop)}");
                }
                catch (KeyNotFoundException)
                {
                    writer.WriteLine($"{indent}{kv.Key.PadRight(_width)}: <not found>");
                }
                catch (NotSupportedException nse)
                {
                    writer.WriteLine($"{indent}{kv.Key.PadRight(_width)}: {nse}");
                }
            }

            writer.WriteLine($"{indent}{"Available Properties".PadRight(_width)}: {_availableProperties.Count}");
            foreach (var kv in _availableProperties)
            {
                writer.WriteLine($"{indent}\t{ExtendDEVPROPKEY.ToString(kv.Key),-45} {kv.Value.Type,-35} : {Extension.GetPropertyString(kv.Value)}");
            }

            writer.WriteLine($"{indent}{"Extensions".PadRight(_width)} : {ExtensionInfs.Count}");
            foreach (var extension in ExtensionInfs)
            {
                writer.WriteLine($"{indent}\t{extension}");
            }

            writer.WriteLine($"{indent}{"Children".PadRight(_width)} : {Children.Count}");
            foreach (var child in Children)
            {
                if (child.DeviceClass.Equals(THX.DeviceClass.AudioEndpoint))
                {
                    writer.WriteLine($"{indent}\tAudioEndpoint");
                    new AudioEndpoint(child).WriteDetailed(writer, indent + "\t");
                }
                else if (child.DeviceClass.Equals(THX.DeviceClass.AudioProcessingObject))
                {
                    writer.WriteLine($"{indent}\tAudioProcessingObjectInf");
                    new AudioProcessingObjectInf(child).WriteDetailed(writer, indent + "\t");
                }
                child.WriteDetailed(writer, indent + "\t");

                writer.WriteLine();
            }
        }

        internal DEVPROPERTY GetProperty(DEVPROPKEY propKey)
        {
            if (!_availableProperties.TryGetValue(propKey, out DEVPROPERTY value))
            {
                throw new KeyNotFoundException($"Property {ExtendDEVPROPKEY.ToString(propKey)} not found on node {this}");
            }

            return value;
        }

        internal T GetProperty<T>(DEVPROPKEY propKey)
        {
            if (!_availableProperties.TryGetValue(propKey, out DEVPROPERTY prop))
            {
                throw new KeyNotFoundException($"Property {ExtendDEVPROPKEY.ToString(propKey)} not found on node {this}");
            }

            return Extension.GetPropertyValue<T>(prop);
        }

    }
}


