using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;

using Windows.Win32;
using Windows.Win32.Devices.DeviceAndDriverInstallation;
using Windows.Win32.Devices.Properties;

namespace TestCM
{
    class DeviceNode : IDisposable
    {
        static private Dictionary<string, DeviceNode>       _deviceNodesByInstanceId = new();
        static private Dictionary<string, HashSet<DeviceNode>> _deviceNodesByHardwareId = new();

        static private void Cache(DeviceNode node)
        {
            _deviceNodesByInstanceId[node.InstanceID] = node;

            foreach (string hardwareId in node.GetProperty<List<string>>(PInvoke.DEVPKEY_Device_HardwareIds)
                .Where(id => null != id))
            {
                var nodes = _deviceNodesByHardwareId.GetValueOrDefault(hardwareId, new ());
                nodes.Add(node);
                _deviceNodesByHardwareId[hardwareId] = nodes;
            }
        }

        static private DeviceNode GetByHardwareId(string hardwareId)
        {
            if (null == _deviceNodesByHardwareId)
            {
                foreach (var diid in DeviceClass.GetDeviceIds(Guid.Empty))
                {
                    try
                    {
                        DeviceNode n = new DeviceNode(diid, PInvoke.DEVPKEY_Device_InstanceId);
                        Cache(n);
                    }
                    catch (KeyNotFoundException)
                    { }
                }
            }

            if (null == _deviceNodesByHardwareId)
            {
                throw new NullReferenceException("_deviceNodesByHardwareId");
            }

            var deviceNodes = _deviceNodesByHardwareId[hardwareId];
            if (deviceNodes == null)
            {
                throw new KeyNotFoundException($"No devices found for hardware ID = {hardwareId}");
            }

            if (deviceNodes.Count != 1)
            {
                throw new InvalidOperationException($"{deviceNodes.Count} device nodes found for hardware ID = {hardwareId}");
            }

            return deviceNodes.First();
        }

        static private uint GetDnInst(string deviceId, DEVPROPKEY propKey)
        {
            uint devNode = 0;
            unsafe
            {
                fixed (char* usDeviceID = deviceId)
                {
                    if (PInvoke.DEVPKEY_Device_InstanceId.Equals(propKey))
                    {
                        Windows.Win32.Foundation.PWSTR devNodeName = usDeviceID;
                        var ret = PInvoke.CM_Locate_DevNode(&devNode, devNodeName, CM_LOCATE_DEVNODE_FLAGS.CM_LOCATE_DEVNODE_PHANTOM);
                        if (CONFIGRET.CR_SUCCESS != ret)
                        {
                            throw new KeyNotFoundException($"CM_Locate_DevNode {deviceId} failed with {ret}");
                        }

                        return devNode;
                    }

                    if (PInvoke.DEVPKEY_Device_HardwareIds.Equals(propKey))
                    {
                        return GetByHardwareId(deviceId)._devNode;
                    }
                }
            }

            return devNode;
        }

        ///
        public DeviceNode(string deviceId,
                          DEVPROPKEY idPropertyKey)
            : this(GetDnInst(deviceId, idPropertyKey))
        { }

        public DeviceNode(uint devNode)
        {
            _devNode = devNode;

            foreach (var propKey in this.PropertyKeys)
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

            _instanceID = GetProperty<string>(PInvoke.DEVPKEY_Device_InstanceId);

            try
            {
                _friendlyName = GetProperty<string>(PInvoke.DEVPKEY_Device_FriendlyName);
            }
            catch (KeyNotFoundException)
            {
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
            unsafe
            {
                foreach (var prop in _availableProperties)
                {
                    Marshal.FreeHGlobal((nint)prop.Value.Buffer);
                }
            }
        }

        override public string ToString()
        {
            return $"{_instanceID} ({_friendlyName})";
        }

        public override bool Equals(object? obj)
        {
            return (obj != null)
                && (obj is DeviceNode node)
                && (_devNode == node._devNode);
        }
        public override int GetHashCode()
        {
            return _devNode.GetHashCode();
        }

        public void WriteDetailed(TextWriter writer, string indent = "")
        {
            Dictionary<string, DEVPROPKEY> headerProperties = new()
        {
            { "Device", PInvoke.DEVPKEY_Device_InstanceId},
            { "Friendly Name", PInvoke.DEVPKEY_Device_FriendlyName },
            { "Description", PInvoke.DEVPKEY_Device_DeviceDesc  },
            { "Driver", PInvoke.DEVPKEY_Device_Driver },
            { "Driver Description", PInvoke.DEVPKEY_Device_DriverDesc },
            { "Install State", PInvoke.DEVPKEY_Device_InstallState },
            { "Problem Code", PInvoke.DEVPKEY_Device_ProblemCode },
            { "Problem Status", PInvoke.DEVPKEY_Device_ProblemStatus },
            { "Is Present", PInvoke.DEVPKEY_Device_IsPresent },
            { "Has Problem", PInvoke.DEVPKEY_Device_HasProblem },
            { "Is Reboot Required", PInvoke.DEVPKEY_Device_IsRebootRequired },
            { "Driver INF Path", PInvoke.DEVPKEY_Device_DriverInfPath }
        };

            int width = (int)(headerProperties.Keys
            .Max(s => (uint)s.Length)
            + 1);

            foreach (var kv in headerProperties)
            {
                try
                {
                    DEVPROPERTY prop = this.GetProperty(kv.Value);
                    writer.WriteLine($"{indent}{kv.Key.PadRight(width)}: {Extension.GetPropertyString(prop)}");
                }
                catch (KeyNotFoundException)
                {
                    writer.WriteLine($"{indent}{kv.Key.PadRight(width)}: <not found>");
                }
                catch (NotSupportedException nse)
                {
                    writer.WriteLine($"{indent}{kv.Key.PadRight(width)}: {nse}");
                }
            }

            writer.WriteLine($"{indent}Available Properties: {_availableProperties.Count}");
            foreach (var kv in _availableProperties)
            {
                writer.WriteLine($"{indent}\t{Extension.ToString(kv.Key)} ({kv.Value.Type}): {Extension.GetPropertyString(kv.Value)}");
            }

            writer.WriteLine($"{indent}Extensions: {this.ExtensionInfs.Count}");
            foreach (var extension in this.ExtensionInfs)
            {
                writer.WriteLine($"{indent}\t{extension}");
            }

            writer.WriteLine($"{indent}Children: {this.Children.Count}");
            foreach (var child in this.Children)
            {
                child.WriteDetailed(writer, indent + "\t");
                writer.WriteLine();
            }
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
                        (uint)1, //REG_CREATE_KEY_DISPOSITION.REG_OPENED_EXISTING_KEY is 2 but RegDisposition_OpenExisting == 0x1 in CfgMgr.h
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

        public string InstanceID { get { return _instanceID; } }

        private string _instanceID;
        public string FriendlyName { get { return _friendlyName; } }

        private string _friendlyName = "";

        private Dictionary<DEVPROPKEY, DEVPROPERTY> _availableProperties = new();

        public List<DeviceNode> Children
        {
            get
            {
                try
                {
                    return GetProperty<List<string>>(PInvoke.DEVPKEY_Device_Children)
                        .Select(child => new DeviceNode(child, PInvoke.DEVPKEY_Device_InstanceId)).ToList();
                }
                catch (KeyNotFoundException)
                {
                    return new List<DeviceNode>();
                }
            }
        }

        public List<DEVPROPKEY> PropertyKeys
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

        private static char[] _StringListSeparators = new char[] { '\0' };

        public DEVPROPERTY GetProperty(DEVPROPKEY propKey)
        {
            if (!_availableProperties.TryGetValue(propKey, out DEVPROPERTY value))
            {
                throw new KeyNotFoundException($"Propert {Extension.ToString(propKey)} not found on node {this}");
            }

            return value;
        }

        public T GetProperty<T>(DEVPROPKEY propKey)
        {
            if (!_availableProperties.TryGetValue(propKey, out DEVPROPERTY prop))
            {
                throw new KeyNotFoundException($"Property {Extension.ToString(propKey)} not found on node {this}");
            }

            return Extension.GetPropertyValue<T>(prop);
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
    }
}
