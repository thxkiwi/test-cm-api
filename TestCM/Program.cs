using Windows.Win32;

foreach (var device in THX.DeviceClass.GetDeviceIds(THX.DeviceClass.Media)
    .Select(deviceInstanceId => new THX.DeviceNode(deviceInstanceId, PInvoke.DEVPKEY_Device_InstanceId))
    )
{
    Console.WriteLine(new string('=', 80));
    device.WriteDetailed(Console.Out);
    Console.WriteLine(new string ('-', 80));
    Console.WriteLine();
}
