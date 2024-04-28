using System.Runtime.InteropServices;

using Windows.Win32;
using Windows.Win32.Devices.DeviceAndDriverInstallation;
using Windows.Win32.Foundation;

namespace THX
{
    public class DeviceClass
    {
        public static readonly Guid Media = new Guid("{4d36e96c-e325-11ce-bfc1-08002be10318}");
        public static readonly Guid AudioProcessingObject = new Guid("{5989fce8-9cd0-467d-8a6a-5419e31529d4}");
        public static readonly Guid SoftwareComponent = new Guid("{5c4c3332-344d-483c-8739-259e934c9cc8}");
        public static readonly Guid AudioEndpoint = new Guid("{c166523c-fe0c-4a94-a586-f1a80cfbbf3e}");

        public static List<string> GetDeviceIds(Guid classId)
        {
            string[] deviceIDs;
            unsafe
            {
                nint buffer = nint.Zero;

                /// The number of characters in the buffer. 
                /// For ANSI callers, this is also the number of bytes; 
                /// for Unicode callers, this is the number of characters.
                uint nCharacters = 0;
                fixed (char* filterRaw = classId.ToString("B"))
                {
                    PCWSTR filter = filterRaw;
                    var ret = PInvoke.CM_Get_Device_ID_List_Size(&nCharacters, filter, PInvoke.CM_GETIDLIST_FILTER_CLASS);
                    if (CONFIGRET.CR_SUCCESS != ret)
                    {
                        throw new InvalidOperationException($"CM_Get_Device_ID_List_Size for {classId} failed with {ret}");
                    }

                    if (0 == nCharacters)
                    {
                        Console.WriteLine("No devices found");
                        return new();
                    }

                    // CM API calls are using WCHARs. Buffer size is sizeof(WCHAR) * nCharacters
                    buffer = Marshal.AllocHGlobal((int)nCharacters * sizeof(char));
                    ret = PInvoke.CM_Get_Device_ID_List(filter, (char*)buffer, nCharacters, PInvoke.CM_GETIDLIST_FILTER_CLASS);
                    if (ret != CONFIGRET.CR_SUCCESS)
                    {
                        throw new InvalidOperationException($"CM_Get_Device_ID_List failed with {ret}");
                    }

                    // Device property strings have an extra NULL terminator
                    // at the end to indicate the end of the list.
                    // C# split does not handle the removal by marshaling 1-fewer.
                    string deviceIDList = Marshal.PtrToStringUni(buffer, (int)nCharacters - 1);
                    deviceIDs = deviceIDList.Split('\0', StringSplitOptions.RemoveEmptyEntries);
                    Marshal.FreeHGlobal(buffer);
                }
            }

            return deviceIDs.ToList();
        }

        public static string GetName(Guid cmClassId)
        {
            string classname = "";
            unsafe
            {
                uint nCharacters = 0;
                PInvoke.CM_Get_Class_Name(cmClassId, null, ref nCharacters, 0);

                char* buffer = (char*)Marshal.AllocHGlobal((int)nCharacters * sizeof(char));
                PInvoke.CM_Get_Class_Name(cmClassId, buffer, ref nCharacters, 0);

                // Device property strings have an extra NULL terminator
                // at the end to indicate the end of the list.
                // C# split does not handle the removal by marshaling 1-fewer.
                classname = Marshal.PtrToStringUni((nint)buffer, (int)nCharacters - 1);
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
