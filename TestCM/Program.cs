using System.Runtime.InteropServices;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.Devices.DeviceAndDriverInstallation;
using Windows.Win32.Devices.Properties;

string filterClass = "{5989fce8-9cd0-467d-8a6a-5419e31529d4}";
string deviceIDList = string.Empty;
unsafe
{
    fixed (char* filterRaw = filterClass)
    {
        PCWSTR filter = filterRaw;
        uint sz = 0;
        PInvoke.CM_Get_Device_ID_List_Size(&sz, filter, PInvoke.CM_GETIDLIST_FILTER_CLASS);

        char* buffer = stackalloc char[(int)sz];
        PInvoke.CM_Get_Device_ID_List(filter, buffer, sz, PInvoke.CM_GETIDLIST_FILTER_CLASS);
        deviceIDList = new string(buffer, 0, (int)sz);
    }
}

// deviceIDList now contains a list of device IDs for devices of the specified class
// They are separated by a NULL character, and the list is terminated by two NULL characters.
// You can split the list like this:
string[] deviceIDs = deviceIDList.Split('\0', StringSplitOptions.RemoveEmptyEntries);

foreach (var deviceID in deviceIDs)
{
    Console.WriteLine(deviceID);
    unsafe
    {
        uint devNode = 0;
        fixed (char* usDeviceID = deviceID)
        {
            PWSTR devNodeName = usDeviceID;
            PInvoke.CM_Locate_DevNode(&devNode, devNodeName, CM_LOCATE_DEVNODE_FLAGS.CM_LOCATE_DEVNODE_NORMAL);
        }

        uint nPropertyKeys = 0;
        PInvoke.CM_Get_DevNode_Property_Keys(devNode, null, &nPropertyKeys, 0);

        IntPtr usPropertyKeys = Marshal.AllocHGlobal(Marshal.SizeOf<DEVPROPKEY>() * (int)nPropertyKeys);
        PInvoke.CM_Get_DevNode_Property_Keys(devNode, (DEVPROPKEY*)usPropertyKeys, &nPropertyKeys, 0);

        DEVPROPKEY[] propertyKeys = new DEVPROPKEY[nPropertyKeys];
        for (int i = 0; i < nPropertyKeys; ++i)
        {
            propertyKeys[i] = Marshal.PtrToStructure<DEVPROPKEY>(usPropertyKeys + i * Marshal.SizeOf<DEVPROPKEY>());
        }

        Console.WriteLine($"Found {nPropertyKeys} property keys for device {deviceID}:");
        foreach (var key in propertyKeys)
        {
            Console.WriteLine($"{{{key.fmtid}}},{key.pid}");
            //if (key.Equals(DEVPKEY_Device_Children) )
            //{

            //}
        }
    }
}
