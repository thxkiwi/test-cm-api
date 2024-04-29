using THX;
using Windows.Win32;

Console.WriteLine("Retireve all devices that match the Media device class...");

foreach (var device in THX.DeviceClass.GetDeviceIds(THX.DeviceClass.Media)
    .Select(deviceInstanceId => new THX.DeviceNode(deviceInstanceId, PInvoke.DEVPKEY_Device_InstanceId))

    )
{
    Console.WriteLine(new string('=', 80));
    device.WriteDetailed(Console.Out);
    Console.WriteLine(new string ('-', 80));
    Console.WriteLine();
}

/// DeviceNode instances are cached by
/// - Instance Id
/// - Hardware Id
/// 
/// The first query by hardward ID may take a long time if node devices have
/// been cached for that hardware ID because it has to load all devices
/// and then filter them. 
/// 
/// Subsequent calls to GetByHardwareId use cached data.
/// Instead of querying by outright hardware ID, first query by the device class
/// and then filter by the property. Now that device is cached subsequent calls
/// to GetByHardwareId for the same hardware ID will use cached data.
/// 
/// Because the media devices are cached above, calling GetByHardwareId for the
/// _after_ the media devices have been cached will be fast.
Console.WriteLine("Retrieve all devices that match a compatible hardtware ID...");

foreach (var device in DeviceNode.GetByHardwareId("USB\\VID_1532&PID_0529&MI_00"))
{
    Console.WriteLine(new string('=', 80));
    device.WriteDetailed(Console.Out);
    Console.WriteLine(new string('-', 80));
    Console.WriteLine();
}
