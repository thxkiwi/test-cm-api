using THX;
using Windows.Win32;

Console.WriteLine("Retireve all devices that match a compatible hardtware ID...");

foreach (var deviceNodes in DeviceNode.GetByHardwareId("USB\\VID_1532&PID_0529&MI_00"))
{
    Console.WriteLine(new string('=', 80));
    deviceNodes.WriteDetailed(Console.Out);
    Console.WriteLine(new string('-', 80));
    Console.WriteLine();
}

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
