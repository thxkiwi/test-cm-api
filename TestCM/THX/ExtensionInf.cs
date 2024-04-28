using Windows.Win32;

namespace THX
{
    public class ExtensionInf
    {
        internal static ExtensionInf Parse(string extendedConfigurationId)
        {
            // Parse the extended configuration ID based on the format:
            // oemXX.inf:DeviceInstanceID,DDInstall Section,Driver Date,Driver Version
            string[] parts = extendedConfigurationId.Split(':');
            if (parts.Length != 2)
            {
                throw new FormatException($"Invalid extended configuration ID: {extendedConfigurationId}");
            }

            string infPath = parts[0];

            parts = parts[1].Split(',');
            if (parts.Length != 4)
            {
                throw new FormatException($"Invalid extended configuration ID: {extendedConfigurationId}");
            }

            string deviceInstanceID = parts[0];
            string ddInstallSection = parts[1];
            string driverDate = parts[2];
            string driverVersion = parts[3];

            return new ExtensionInf()
            {
                InfPath = infPath,
                HardwareId = deviceInstanceID,
                DDInstallSection = ddInstallSection,
                DriverDate = driverDate,
                DriverVersion = driverVersion
            };
        }

        public DeviceNode Node
        {
            get
            {
                return new DeviceNode(HardwareId, PInvoke.DEVPKEY_Device_HardwareIds);
            }
        }

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.
        public string InfPath { get; private set; }

        public string HardwareId { get; private set; }

        public string DDInstallSection { get; private set; }

        public string DriverDate { get; private set; }

        public string DriverVersion { get; private set; }
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.

        public override string ToString()
        {
            return $"{InfPath}:{HardwareId},{DDInstallSection},{DriverDate},{DriverVersion}";
        }

    }
}
