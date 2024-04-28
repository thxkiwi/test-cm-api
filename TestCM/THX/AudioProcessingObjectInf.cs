using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace THX
{
    /// 
    /// \brief An AudioProcessingObject driver INF
    /// 
    /// An AudioProcessingObject driver INF defines one or more 
    /// AudioProcessingObject devices.
    /// 
    /// An AudioProcessingObject device is related starting from the 
    /// Manufacturer entries that relate an INF model to one ore TargetOSVersions. 
    /// 
    /// e.g.,
    /// [Manufacturer]
    /// manufacturer-name |
    /// %strkey%=models-section-name |
    /// %strkey%=models-section-name [,TargetOSVersion] [,TargetOSVersion] ...  (Windows XP and later versions of Windows)
    /// 
    /// A model section enumerates the model for each TargetOSVersion.
    /// 
    /// e.g.,
    /// [models-section-name] |
    /// [models-section-name.TargetOSVersion]  (Windows XP and later versions of Windows)
    ///
    /// device-description=install-section-name,[hw-id][, compatible-id...]
    /// [device-description=install-section-name,[hw-id][, compatible-id]...] ...
    /// 
    /// An install section provides the installation information for the 
    /// AudioProcessingObject.
    public class AudioProcessingObjectInf
    {
        /// \brief Access to the device node
        public DeviceNode DeviceNode { get; private set; }

        /// 
        /// \brief Dictionary of AudioProcessingObjects as installed for the 
        /// TargetOSVersion of the host OS.
        public List<AudioProcessingObject>
            AudioProcessingObjects
        { get; private set; }

        /// 
        /// \details Examines both the AudioProcessingObject device software 
        /// key, and the registry HKCR and HKLM.
        public AudioProcessingObjectInf(DeviceNode deviceNode)
        {
            DeviceNode = deviceNode;
            AudioProcessingObjects
                = new List<AudioProcessingObject>();

            // From the Software key of the device node, iterate over Classes\CLSID\*
            var clsidRegBase = deviceNode.SoftwareKey.OpenSubKey("Classes\\CLSID");
            var clsids = (clsidRegBase?.GetSubKeyNames() ?? []).Select(clsid => Guid.Parse(clsid));
            foreach (var clsid in clsids)
            {
                AudioProcessingObjects.Add(new AudioProcessingObject(clsidRegBase, clsid));
            }
        }

        public void WriteDetailed(TextWriter writer, string indent = "")
        {
            writer.WriteLine($"{indent}AudioProcessingObjects: {AudioProcessingObjects.Count}");
            foreach (var audioProcessingObject in AudioProcessingObjects)
            {
                audioProcessingObject.WriteDetailed(writer, indent + "\t");
            }
        }
    }
}
