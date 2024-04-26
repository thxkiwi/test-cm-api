using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using Windows.Win32.Foundation;
using Windows.Win32.Security;
using Windows.Win32.Devices.Properties;

namespace TestCM
{
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
}
