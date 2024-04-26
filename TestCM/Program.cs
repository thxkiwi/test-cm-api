using System.Runtime.InteropServices;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.Devices.DeviceAndDriverInstallation;
using Windows.Win32.Devices.Properties;
using System.Diagnostics.CodeAnalysis;
using Windows.Win32.Security;
using System.Numerics;

List<string> GetDeviceIDsForClass(Guid classId)
{
    string[] deviceIDs;
    unsafe
    {
        nint buffer = 0;
        uint sz = 0;
        fixed (char* filterRaw = classId.ToString("B"))
        {
            PCWSTR filter = filterRaw;
            var ret = PInvoke.CM_Get_Device_ID_List_Size(&sz, filter, PInvoke.CM_GETIDLIST_FILTER_CLASS);
            if (CONFIGRET.CR_SUCCESS != ret)
            {
                throw new InvalidOperationException($"CM_Get_Device_ID_List_Size for {classId} failed with {ret}");
            }

            if (0 == sz)
            {
                Console.WriteLine("No devices found");
                return [];
            }

            buffer = Marshal.AllocHGlobal((int)sz * 2);
            ret = PInvoke.CM_Get_Device_ID_List(filter, (char*)buffer, sz, PInvoke.CM_GETIDLIST_FILTER_CLASS);
            if (ret != CONFIGRET.CR_SUCCESS)
            {
                throw new InvalidOperationException($"CM_Get_Device_ID_List failed with {ret}");
            }
            string deviceIDList = Marshal.PtrToStringUni(buffer, (int)sz);
            deviceIDs = deviceIDList.Split('\0', StringSplitOptions.RemoveEmptyEntries);
            if (buffer != 0)
            {
                Marshal.FreeHGlobal(buffer);
            }
        }
    }

    return deviceIDs.ToList();
}

// Device classes
static string GetClassName(Guid cmClassId)
{
    string classname = "";
    unsafe
    {
        uint sz = 0;
        PInvoke.CM_Get_Class_Name(cmClassId, null, ref sz, 0);

        char* buffer = (char*)Marshal.AllocHGlobal((int)sz * 2);
        PInvoke.CM_Get_Class_Name(cmClassId, buffer, ref sz, 0);
        classname = Marshal.PtrToStringUni((nint)buffer, (int)sz);
        Marshal.FreeHGlobal((nint)buffer);
    }

    return classname.TrimEnd('\0');
}

static List<Guid> GetClasses(CM_ENUMERATE_FLAGS flags)
{
    HashSet<Guid> classes = new HashSet<Guid>();
    unsafe
    {
        CONFIGRET ret = CONFIGRET.CR_SUCCESS;
        uint idx = 0;
        while (CONFIGRET.CR_NO_SUCH_VALUE != ret)
        {
            Guid guid = Guid.Empty;
            ret = PInvoke.CM_Enumerate_Classes(idx, out guid, flags);
            if (CONFIGRET.CR_SUCCESS == ret)
            {
                classes.Add(guid);
            }
            ++idx;
        }
    }
    return classes.ToList();
}

ILookup<string, Guid> deviceSetupClasses =
    GetClasses(CM_ENUMERATE_FLAGS.CM_ENUMERATE_CLASSES_INSTALLER)
    .ToLookup(guid => GetClassName(guid), guid => guid);

Guid deviceClassGuidMedia = deviceSetupClasses
    .Where(kv => kv.Key.ToLower().Equals("media"))
    .SelectMany(kv => kv).Last();

ILookup<string, Guid> deviceInterfaceClasses =
    GetClasses(CM_ENUMERATE_FLAGS.CM_ENUMERATE_CLASSES_INTERFACE)
    .ToLookup(guid => GetClassName(guid), guid => guid);

foreach (var deviceID in GetDeviceIDsForClass(deviceClassGuidMedia))
{
    Console.WriteLine("===============================================");
    DeviceNode device = new(deviceID);
    device.WriteDetailed(Console.Out);
    Console.WriteLine("-----------------------------------------------\n");
}

internal static class Extension
{
    // Extension Method for DEVPROPKEY.ToString()
    internal static string ToString(ref this DEVPROPKEY key)
    {
        return $"{{{key.fmtid}}},{key.pid}";
    }

    internal static string ToString(this DEVPROPKEY key)
    {
        return $"{{{key.fmtid}}},{key.pid}";
    }

    internal static T GetPropertyValue<T>(DEVPROPERTY prop)
    {
        Type requestedType = typeof(T);

        unsafe
        {
            switch (prop.Type)
            {
                case DEVPROPTYPE.DEVPROP_TYPE_BOOLEAN:
                    if (requestedType != typeof(bool))
                    {
                        throw new InvalidCastException($"Unable to cast property {prop.CompKey.Key} of type DEVPROPTYPE == {prop.Type} to {requestedType}");
                    }
                    return (T)(object)(0 != Marshal.ReadByte((nint)prop.Buffer));
                case DEVPROPTYPE.DEVPROP_TYPE_BYTE:
                    if (requestedType != typeof(byte))
                    {
                        throw new InvalidCastException($"Unable to cast property {prop.CompKey.Key} of type DEVPROPTYPE == {prop.Type} to {requestedType}");
                    }
                    return (T)(object)Marshal.ReadByte((nint)prop.Buffer);
                case DEVPROPTYPE.DEVPROP_TYPE_STRING:
                    if (requestedType != typeof(string))
                    {
                        throw new InvalidCastException($"Unable to cast property {prop.CompKey.Key} of type DEVPROPTYPE == {prop.Type} to {requestedType}");
                    }
                    return (T)(object)Marshal.PtrToStringUni((nint)prop.Buffer, (int)prop.BufferSize / 2);

                case DEVPROPTYPE.DEVPROP_TYPE_STRING_LIST:
                    if (requestedType != typeof(List<string>))
                    {
                        throw new InvalidCastException($"Unable to cast property {prop.CompKey.Key} of type DEVPROPTYPE == {prop.Type} to {requestedType}");
                    }
                    string rawResult = Marshal.PtrToStringUni((nint)prop.Buffer, (int)prop.BufferSize / 2);
                    return (T)(object)rawResult.Split('\0', StringSplitOptions.RemoveEmptyEntries).ToList();

                case DEVPROPTYPE.DEVPROP_TYPE_BINARY:
                    byte[] result = new byte[prop.BufferSize];
                    Marshal.Copy((nint)prop.Buffer, result, 0, (int)prop.BufferSize);
                    return (T)(object)result;

                case DEVPROPTYPE.DEVPROP_TYPE_UINT32:
                    return (T)(object)Marshal.ReadInt32((nint)prop.Buffer);

                case DEVPROPTYPE.DEVPROP_TYPE_UINT64:
                    return (T)(object)Marshal.ReadInt64((nint)prop.Buffer);

                case DEVPROPTYPE.DEVPROP_TYPE_GUID:
                    return (T)(object)Marshal.PtrToStructure<Guid>((nint)prop.Buffer);

                case DEVPROPTYPE.DEVPROP_TYPE_NTSTATUS:
                    return (T)(object)Marshal.ReadInt32((nint)prop.Buffer);

                case DEVPROPTYPE.DEVPROP_TYPE_SECURITY_DESCRIPTOR:
                    if (requestedType != typeof(Windows.Win32.Security.SECURITY_DESCRIPTOR))
                    {
                        throw new InvalidCastException($"Unable to cast property {prop.CompKey.Key} of type DEVPROPTYPE == {prop.Type} to {requestedType}");
                    }

                    if (sizeof(Windows.Win32.Security.SECURITY_DESCRIPTOR) > prop.BufferSize)
                    {
                        throw new InvalidCastException(string.Format(
                            "Unable to cast property {0} of type DEVPROPTYPE == {1} to {2}: sizeof({5}}) [{3}] > prop.BufferSize [{4}]",
                            prop.CompKey,
                            prop.Type,
                            requestedType,
                            sizeof(SECURITY_DESCRIPTOR),
                            prop.BufferSize,
                            typeof(SECURITY_DESCRIPTOR)));
                    }
                    return (T)(object)Marshal.PtrToStructure<Windows.Win32.Security.SECURITY_DESCRIPTOR>((nint)prop.Buffer);

                case DEVPROPTYPE.DEVPROP_TYPE_FILETIME:
                    if (requestedType != typeof(FILETIME))
                    {
                        throw new InvalidCastException($"Unable to cast property {prop.CompKey.Key} of type DEVPROTYPE == {prop.Type} to {requestedType}");
                    }

                    if (prop.BufferSize < sizeof(FILETIME))
                    {
                        throw new InvalidCastException(string.Format(
                            "Unable to cast property {0} of type DEVPROPTYPE == {1} to {2}: sizeof({5}}) [{3}] > prop.BufferSize [{4}]",
                            prop.CompKey,
                            prop.Type,
                            requestedType,
                            sizeof(FILETIME),
                            prop.BufferSize,
                            typeof(FILETIME)));
                    }

                    return (T)(object)Marshal.PtrToStructure<FILETIME>((nint)prop.Buffer);

                default:
                    throw new InvalidCastException($"Unable to cast property of type DEVPROPTYPE == {prop.Type} to {typeof(T)}, buffer size = {prop.BufferSize}");
            }
        }
    }

    internal static string GetPropertyString(DEVPROPERTY prop)
    {
        unsafe
        {
            switch (prop.Type)
            {
                case DEVPROPTYPE.DEVPROP_TYPE_BOOLEAN:
                    return GetPropertyValue<bool>(prop).ToString();
                case DEVPROPTYPE.DEVPROP_TYPE_BYTE:
                    return GetPropertyValue<byte>(prop).ToString();
                case DEVPROPTYPE.DEVPROP_TYPE_STRING:
                    return GetPropertyValue<string>(prop);
                case DEVPROPTYPE.DEVPROP_TYPE_STRING_LIST:
                    return string.Join(',', GetPropertyValue<List<string>>(prop));
                case DEVPROPTYPE.DEVPROP_TYPE_BINARY:
                    return string.Join(",", GetPropertyValue<byte[]>(prop).Select(b => $"{b:X2}"));
                case DEVPROPTYPE.DEVPROP_TYPE_UINT32:
                    return GetPropertyValue<Int32>(prop).ToString();
                case DEVPROPTYPE.DEVPROP_TYPE_UINT64:
                    return GetPropertyValue<Int64>(prop).ToString();
                case DEVPROPTYPE.DEVPROP_TYPE_GUID:
                    return GetPropertyValue<Guid>(prop).ToString("B");
                case DEVPROPTYPE.DEVPROP_TYPE_NTSTATUS:
                    return GetPropertyValue<Int32>(prop).ToString();
                case DEVPROPTYPE.DEVPROP_TYPE_SECURITY_DESCRIPTOR:
                    SECURITY_DESCRIPTOR d = GetPropertyValue<SECURITY_DESCRIPTOR>(prop);
                    return $"Revision={d.Revision} Control=0x{d.Control:X} Owner={d.Owner} Group={d.Group} Sacl=<add support> Dacl=<add support>";
                case DEVPROPTYPE.DEVPROP_TYPE_FILETIME:
                    FILETIME ft = GetPropertyValue<FILETIME>(prop);
                    Int64 value = (ft.dwHighDateTime << 32) | ft.dwLowDateTime;
                    return value.ToString();
                default:
                    break;
            }
        }

        throw new NotSupportedException($"Property type {prop.Type} must be added for conversion to string. size = {prop.BufferSize}");
    }
}

#if true
class ExtensionInf
{
    internal static ExtensionInf Parse(string extendedConfigurationId)
    {
        // Parse the extended configuration ID based on the format:
        // oemXX.inf:DeviceInstanceID,DDInstall Section,Driver Date,Driver Version
        string[] parts = extendedConfigurationId.Split(':');
        if (parts.Length != 2)
        {
            throw new FormatException($"Invalid extended configuration ID: {extendedConfigurationId}");
        }

        string infPath = parts[0];

        parts = parts[1].Split(',');
        if (parts.Length != 4)
        {
            throw new FormatException($"Invalid extended configuration ID: {extendedConfigurationId}");
        }

        string deviceInstanceID = parts[0];
        string ddInstallSection = parts[1];
        string driverDate = parts[2];
        string driverVersion = parts[3];

        return new ExtensionInf()
        {
            InfPath = infPath,
            DeviceInstanceID = deviceInstanceID,
            DDInstallSection = ddInstallSection,
            DriverDate = driverDate,
            DriverVersion = driverVersion
        };
    }

    public DeviceNode Node
    {
        get
        {
            return new DeviceNode(DeviceInstanceID);
        }
    }

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.
    public string InfPath { get; private set; }

    public string DeviceInstanceID { get; private set; }

    public string DDInstallSection { get; private set; }

    public string DriverDate { get; private set; }

    public string DriverVersion { get; private set; }
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.

    public override string ToString()
    {
        return $"{InfPath}:{DeviceInstanceID},{DDInstallSection},{DriverDate},{DriverVersion}";
    }
}
#endif

class DeviceNode : IDisposable
{
    static private uint GetDnInst(string deviceInstanceId)
    {
        uint devNode = 0;
        unsafe
        {
            fixed (char* usDeviceID = deviceInstanceId)
            {
                PWSTR devNodeName = usDeviceID;
                var ret = PInvoke.CM_Locate_DevNode(&devNode, devNodeName, CM_LOCATE_DEVNODE_FLAGS.CM_LOCATE_DEVNODE_PHANTOM);
                if (CONFIGRET.CR_SUCCESS != ret)
                {
                    throw new InvalidOperationException($"CM_Locate_DevNode {deviceInstanceId} failed with {ret}");
                }
            }
        }

        return devNode;
    }

    public DeviceNode(string deviceInstanceId) : this(GetDnInst(deviceInstanceId))
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
                    .Select(child => new DeviceNode(child)).ToList();
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