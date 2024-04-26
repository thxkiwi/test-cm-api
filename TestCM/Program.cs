using System.Runtime.InteropServices;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.Devices.DeviceAndDriverInstallation;
using Windows.Win32.Devices.Properties;
using System.Diagnostics.CodeAnalysis;
using Windows.Win32.Security;
using System.Numerics;
using Microsoft.Win32;
using Windows.Win32.System.Registry;
using Microsoft.Win32.SafeHandles;

using TestCM;

ILookup<string, Guid> deviceSetupClasses =
                TestCM.DeviceClass.GetClasses(CM_ENUMERATE_FLAGS.CM_ENUMERATE_CLASSES_INSTALLER)
                .ToLookup(guid => TestCM.DeviceClass.GetName(guid), guid => guid);

Guid DeviceClassMedia = deviceSetupClasses
                .Where(kv => kv.Key.ToLower().Equals("media"))
                .SelectMany(kv => kv).Last();

Guid DeviceClassAudioProcessingObject = deviceSetupClasses
                .Where(kv => kv.Key.ToLower().Equals("audioprocessingobject"))
                .SelectMany(kv => kv).Last();

Guid DeviceClassSoftwareComponent = deviceSetupClasses
                .Where(kv => kv.Key.ToLower().Equals("softwarecomponent"))
                .SelectMany(kv => kv).Last();

// Find MEDIA devices that have a THX APO associated with them
// Display
// - The device
// - The THX APO
// - The THX SoftwareComponents
foreach (var device in TestCM.DeviceClass.GetDeviceIds(DeviceClassMedia)
    .Select(deviceInstanceId => new DeviceNode(deviceInstanceId, PInvoke.DEVPKEY_Device_InstanceId))
    .Where(d => d.Children.Any(c => c.GetProperty<Guid>(PInvoke.DEVPKEY_Device_ClassGuid)
                                        .Equals(DeviceClassAudioProcessingObject)
                                 && c.GetProperty<string>(PInvoke.DEVPKEY_Device_DriverProvider)
                                        .Contains("THX")))
    .Where(d => d.InstanceID.Contains("USB\\VID_1532&PID_0529&MI_00"))
    )
{
    // Device details
    device.WriteDetailed(Console.Out);

    var thxChildren = device.Children.Where(c => c.GetProperty<string>(PInvoke.DEVPKEY_Device_DriverProvider)
                                                            .Contains("THX"));

    // Any Extension INFs that are provided by THX
    foreach (var extension in device.ExtensionInfs)
    {
        Console.WriteLine("THX Extension INF:");
        Console.WriteLine($"\t{extension.InfPath}");
    }

    // APO Software Key Details
    var thxapo = thxChildren.Where(c => c.GetProperty<Guid>(PInvoke.DEVPKEY_Device_ClassGuid)
                                                    .Equals(DeviceClassAudioProcessingObject))
                    .Last();
    Console.WriteLine("THX AudioProcessingObject:");
    Console.WriteLine($"OS Version: {Environment.OSVersion.Version}");
    TestCM.Registry.Traverse(Console.Out, thxapo.SoftwareKey, "\t");

    // On Windows 10, prior to 10.0.22000.0, the THX APOs are configured
    // via HKCR\CLSID and HKCR\AudioEngine\AudioProcessingObjects
    // Examine the Windows version and if it is less than 10.0.22000.0
    // then display the APOs from those locations.
    if (Environment.OSVersion.Version < new Version(10, 0, 22000, 0))
    {
        Console.WriteLine(Environment.OSVersion.Version);
        // Find all HKCR\CLSID\{CLSID} keys that have a FriendlyName value
        // that begins with "THX Spatial Audio"
        // HKCR\CLSID\{CLSID}\@ = FriendlyName ("THX Spatial Audio...")
        const string THX_APO_FRIENDLY_NAME_PREFIX = "THX Spatial Audio";
        var clsidKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(
                        "CLSID",
                        RegistryKeyPermissionCheck.ReadSubTree,
                        System.Security.AccessControl.RegistryRights.ReadKey);
        if (clsidKey != null)
        {
            var thxClsidKeys = clsidKey.GetSubKeyNames()
                        .Select(KeyName => (KeyName, clsidKey.OpenSubKey(KeyName)))
                        .Where(key => key.Item2.GetValue(null)?.ToString()?.StartsWith(THX_APO_FRIENDLY_NAME_PREFIX) ?? false);

            foreach (var thxClsidKey in thxClsidKeys)
            {
                Console.WriteLine($"\tCLSID\\{thxClsidKey.KeyName}");
                Console.WriteLine($"\t\tFriendlyName = {thxClsidKey.Item2.GetValue(null)}");
                RegistryKey inprocServer32 = thxClsidKey.Item2.OpenSubKey("InprocServer32");
                Console.WriteLine($"\t\tDLL Path = {inprocServer32?.GetValue(null) ?? "<InprocServer32 not found>"}");

                // HCKR\AudioEngine\AudioProcessingObjects\{CLSID}\@ = FriendlyName ("THX Spatial Audio...")
                RegistryKey? audioEngineAPO = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(
                                       $"AudioEngine\\AudioProcessingObjects\\{thxClsidKey.KeyName}",
                                        RegistryKeyPermissionCheck.ReadSubTree,
                                        System.Security.AccessControl.RegistryRights.ReadKey);

                Console.WriteLine($"\tAudioEngine\\AudioProcessingObjects\\{thxClsidKey.KeyName}");
                foreach (var propertyName in audioEngineAPO?.GetValueNames() ?? [])
                {
                    var kind = audioEngineAPO?.GetValueKind(propertyName);
                    var value = audioEngineAPO.GetValue(propertyName);
                    Console.WriteLine($"\t\t{propertyName} = {kind} {value}");
                }
            }
        }

    }

    // SoftwareComponents
    foreach (var softwareComponent in thxChildren.Where(c => c.GetProperty<Guid>(PInvoke.DEVPKEY_Device_ClassGuid)
                                                    .Equals(DeviceClassSoftwareComponent)))
    {
        Console.WriteLine("THX SoftwareComponent");
        softwareComponent.WriteDetailed(Console.Out, "\t");
    }

    Console.WriteLine();
}
