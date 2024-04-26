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
    TestCM.Registry.Traverse(Console.Out, thxapo.SoftwareKey, "\t");

    // SoftwareComponents
    foreach (var softwareComponent in thxChildren.Where(c => c.GetProperty<Guid>(PInvoke.DEVPKEY_Device_ClassGuid)
                                                    .Equals(DeviceClassSoftwareComponent)))
    {
        Console.WriteLine("THX SoftwareComponent");
        softwareComponent.WriteDetailed(Console.Out, "\t");
    }

    Console.WriteLine();
}
