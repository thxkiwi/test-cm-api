using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCM
{
    internal class Registry
    {
        internal static void Traverse(TextWriter writer, RegistryKey k, string indent = "")
        {
            writer.WriteLine($"{indent}{k.Name}");
            foreach (string valueName in k.GetValueNames())
            {
                var value = k.GetValue(valueName, "<no value>", RegistryValueOptions.None);
                writer.WriteLine($"{indent}{valueName} = {value}");
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
    }
}
