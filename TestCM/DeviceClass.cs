using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

using Windows.Win32;
using Windows.Win32.Devices.Properties;
using Windows.Win32.Devices.DeviceAndDriverInstallation;
using Windows.Win32.Foundation;

namespace TestCM
{
    internal class DeviceClass
    {
        internal static List<string> GetDeviceIds(Guid classId)
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
        internal static string GetName(Guid cmClassId)
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

        internal static List<Guid> GetClasses(CM_ENUMERATE_FLAGS flags)
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

    }
}
