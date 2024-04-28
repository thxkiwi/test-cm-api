using Microsoft.Win32;
using Windows.Win32.Devices.Properties;

namespace THX
{
    internal class Registry
    {
        internal static void Traverse(TextWriter writer, RegistryKey k, string indent = "")
        {
            writer.WriteLine($"{indent}{k.Name}");
            var valueNames = k.GetValueNames();
            if (valueNames.Length > 0)
            {
                uint maxLength = valueNames.Max(name => (uint)name.Length);
                foreach (string valueName in valueNames)
                {
                    var value = k.GetValue(valueName, "<no value>", RegistryValueOptions.None);
                    writer.WriteLine($"{indent}{valueName:maxLength + 1} = {value}");
                }
            }

            foreach (string subKeyName in k.GetSubKeyNames())
            {
                RegistryKey? subKey = k.OpenSubKey(subKeyName, false);
                if (null == subKey)
                {
                    writer.WriteLine($"{indent}Unable to open subkey {subKeyName}");
                    continue;
                }
                Traverse(writer, subKey, indent + "\t");
            }
        }

        internal static ValueT? GetValue<ValueT>(RegistryKey rk, DEVPROPKEY key, ValueT? defaultValue)
            where ValueT : class?
        {
            var value = rk.GetValue(key.ToString(), defaultValue);
            return null == value ? defaultValue : (ValueT)value;
        }

        internal static ValueT? GetValue<ValueT>(RegistryKey rk, Windows.Win32.UI.Shell.PropertiesSystem.PROPERTYKEY key, ValueT? defaultValue)
            where ValueT : class?
        {
            var value = rk.GetValue(key.ToString(), defaultValue);
            return (null == value) ? defaultValue : (ValueT)value;
        }

        internal static ValueT? GetValue<KeyT, ValueT>(RegistryKey rk, KeyT? key, ValueT? defaultValue)
            where KeyT : class?
            where ValueT : class?
        {
            var value = rk.GetValue(key?.ToString() ?? null, defaultValue);
            return (null == value) ? defaultValue : (ValueT)value;
        }

        internal static ValueT? GetValueTypeValue<KeyT, ValueT>(RegistryKey rk, KeyT? key, ValueT? defaultValue)
            where KeyT : class?
            where ValueT : struct
        {
            var value = rk.GetValue(key?.ToString() ?? null, defaultValue);
            return (null == value) ? defaultValue : (ValueT)value;
        }

    }
}
