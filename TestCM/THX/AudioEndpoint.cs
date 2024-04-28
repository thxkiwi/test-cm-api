﻿using Microsoft.Win32;
using Windows.Win32;
using Windows.Win32.UI.Shell.PropertiesSystem;

namespace THX
{
    /// 
    /// \brief Work-around for the lack of a managed API to access effects properties.
    /// 
    /// \todo Port to PInvoke using ActivateAudioInterfaceAsync to get IAudioEffectsPropertyStore.
    public class AudioEndpoint
    {
        private static string RegistryRoot = @"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\{0}\{1}";

        private static readonly Dictionary<PropertyKey, string> KeyToName = (new Dictionary<PROPERTYKEY, string>{
            { PInvoke.PKEY_CompositeFX_ModeEffectClsid, "PKEY_CompositeFX_ModeEffectClsid" },
            { PInvoke.PKEY_CompositeFX_EndpointEffectClsid, "PKEY_CompositeFX_ModeEffectClsid" },
            { PInvoke.PKEY_CompositeFX_Offload_StreamEffectClsid, "PKEY_CompositeFX_Offload_StreamEffectClsid" },
            { PInvoke.PKEY_CompositeFX_Offload_ModeEffectClsid, "PKEY_CompositeFX_Offload_ModeEffectClsid" },
            { PInvoke.PKEY_SFX_ProcessingModes_Supported_For_Streaming, "PKEY_SFX_ProcessingModes_Supported_For_Streaming" },
            { PInvoke.PKEY_SFX_Offload_ProcessingModes_Supported_For_Streaming, "PKEY_SFX_Offload_ProcessingModes_Supported_For_Streaming" },
            { PInvoke.PKEY_MFX_ProcessingModes_Supported_For_Streaming, "PKEY_MFX_ProcessingModes_Supported_For_Streaming" },
            { PInvoke.PKEY_MFX_Offload_ProcessingModes_Supported_For_Streaming, "PKEY_MFX_Offload_ProcessingModes_Supported_For_Streaming" },
            { PInvoke.PKEY_EFX_ProcessingModes_Supported_For_Streaming, "PKEY_EFX_ProcessingModes_Supported_For_Streaming" },
            { PInvoke.PKEY_FX_Association, "PKEY_FX_Association" }
        }).Select(pk => (PropertyKey.From(pk.Key), pk.Value)).ToDictionary();

        private static readonly Dictionary<Guid, string> ProcessingModesForStreaming = new Dictionary<Guid, string>
        {
            { Guid.Parse("C18E2F7E-933D-4965-B7D1-1EEF228D2AF3"), "Default" },
            { Guid.Parse("9E90EA20-B493-4FD1-A1A8-7E1361A956CF"), "Raw" },
            { Guid.Parse("B26FEB0D-EC94-477C-9494-D1AB8E753F6E"), "Movies" },
            { Guid.Parse("4780004E-7133-41D8-8C74-660DADD2C0EE"), "Media" },
            { Guid.Parse("FC1CFC9B-B9D6-4CFA-B5E0-4BB2166878B2"), "Speech" },
            { Guid.Parse("98951333-B9CD-48B1-A0A3-FF40682D73F7"), "Communications" },
            { Guid.Parse("9CF2A70B-F377-403B-BD6B-360863E0355C"), "Notification" }
        };

        private static readonly List<PropertyKey> ProcessingModesForStreamingProperties = (new List<PROPERTYKEY>
        {
            PInvoke.PKEY_SFX_ProcessingModes_Supported_For_Streaming,
            PInvoke.PKEY_SFX_Offload_ProcessingModes_Supported_For_Streaming,
            PInvoke.PKEY_MFX_ProcessingModes_Supported_For_Streaming,
            PInvoke.PKEY_MFX_Offload_ProcessingModes_Supported_For_Streaming,
            PInvoke.PKEY_EFX_ProcessingModes_Supported_For_Streaming,
        }).Select(k => PropertyKey.From(k)).ToList();

        public DeviceNode DeviceNode { get; private set; }

        public string DeviceInstanceId { get; private set; }

        public Guid EndpointId { get; private set; }

        public string MMDeviceID { get; private set; }

        public enum EDataFlow
        {
            Render = 0,
            Capture = 1,
            All = 2
        }

        public RegistryKey RegistryKey { get; private set; }

        public EDataFlow DataFlow { get; private set; }

        /// \brief AudioProcessingObjects indexed by APO slot.
        public Dictionary<PropertyKey, List<AudioProcessingObject>> AudioProcessingObjects
        {
            get;

            private set;
        }

        public List<Guid> UnresolvedAudioProcessingObjects { get; private set; }

        /// \param endpointDeviceInstanceId SWD\MMDEVAPI\{0.0.([01]).00000000}.<GUID> 
        /// 
        public AudioEndpoint(DeviceNode deviceNode,
            string regex = @"SWD\\MMDEVAPI\\({0\.0\.([01])\.00000000}\.(.*))")
        {
            DeviceNode = deviceNode;
            DeviceInstanceId = deviceNode.InstanceID;

            var match = System.Text.RegularExpressions.Regex.Match(DeviceInstanceId, regex);
            MMDeviceID = match.Groups[1].Value;
            string dataFlowStr = match.Groups[2].Value;
            EndpointId = Guid.Parse(match.Groups[3].Value);

            EDataFlow dataFlowEnum;
            if (!Enum.TryParse(dataFlowStr, out dataFlowEnum))
            {
                throw new ArgumentException($"Invalid device instance id: {DeviceInstanceId}; unknown data flow {dataFlowStr}.");
            }
            DataFlow = dataFlowEnum;

            string registryPath = string.Format(RegistryRoot, DataFlow, EndpointId.ToString("B"));
            var rk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(registryPath, false);
            if (null == rk)
            {
                throw new ArgumentException($"Failed to open registry key {registryPath}");
            }
            RegistryKey = rk;

            AudioProcessingObjects = new();
            UnresolvedAudioProcessingObjects = new();

            // Get the siblings that are AudioProcessingObjects
            var siblingAPOs = deviceNode.Siblings
                .Where(s => s.GetProperty<Guid>(PInvoke.DEVPKEY_Device_ClassGuid)
                                           .Equals(DeviceClass.AudioProcessingObject))
                .Select(s => new AudioProcessingObjectInf(s));

            // Form the set union of the AudioProcessingObjects
            var apos = siblingAPOs.SelectMany(s => s.AudioProcessingObjects)
                .ToDictionary(apo => apo.ClassId, apo => apo);

            // Projects built against WDK 10.0.22000.0 allow one to run on 
            // earlier versions of Windows 10. However, the Windows audio
            // before 10.0.22000.0 does not query the AudioProcessingObject 
            // software keys. 

            if (Environment.OSVersion.Version < new Version(10, 0, 22000, 0))
            {
                return;
            }

            // Build the Dictionary of AudioProcessingObjects
            foreach (var apoSlot in new PROPERTYKEY[]{
                            PInvoke.PKEY_CompositeFX_ModeEffectClsid,
                            PInvoke.PKEY_CompositeFX_EndpointEffectClsid,
                            PInvoke.PKEY_CompositeFX_Offload_StreamEffectClsid,
                            PInvoke.PKEY_CompositeFX_Offload_ModeEffectClsid
                        }
                    .Select(k => PropertyKey.From(k))
                )
            {
                var clsids = (GetProperty<PropertyKey, string[]>(
                    apoSlot,
                    PropertyLocation.Effects
                ) ?? [])
                .Select(s => Guid.Parse(s));

                foreach (var clsid in clsids)
                {
                    if (!apos.TryGetValue(clsid, out AudioProcessingObject? apo))
                    {
                        // An APO with the given CLSID was not found in an AudioProcessingObject
                        // software key. Attempt to instantiate by CLSID.
                        try
                        {
                            apo = new AudioProcessingObject(clsid);
                        }
                        catch (KeyNotFoundException)
                        {
                            // An APO with the given CLSID was not found in the registry.
                            // This is not an error condition.
                            UnresolvedAudioProcessingObjects.Add(clsid);
                            continue;
                        }
                    }

                    List<AudioProcessingObject>? audioProcessingObjects;
                    if (!AudioProcessingObjects.TryGetValue(apoSlot, out audioProcessingObjects))
                    {
                        audioProcessingObjects = new();
                        AudioProcessingObjects.Add(apoSlot, audioProcessingObjects);
                    }

                    audioProcessingObjects.Add(apo);
                }
            }
        }

        public enum PropertyLocation
        {
            Endpoint,
            Effects
        }

        public ValueT? GetProperty<KeyT, ValueT>(KeyT key, PropertyLocation location = PropertyLocation.Endpoint)
            where ValueT : class?
            where KeyT : class
        {
            RegistryKey? propertyStore = location switch
            {
                PropertyLocation.Endpoint => RegistryKey.OpenSubKey("Properties"),
                PropertyLocation.Effects => RegistryKey.OpenSubKey("FXProperties"),
                _ => throw new ArgumentException($"Invalid location: {location}")
            };

            if (null == propertyStore)
            {
                throw new ArgumentException($"Failed to open registry key {location}");
            }

            var value = Registry.GetValue(propertyStore, key, default(ValueT));

            return value;
        }

        public override string ToString()
        {
            return $"({DataFlow}) {MMDeviceID}";
        }

        public void WriteDetailed(TextWriter writer, string prefix = "")
        {
            writer.WriteLine($"{prefix}AudioProcessingObjects            : {AudioProcessingObjects.Count}");
            foreach (var apoSlot in AudioProcessingObjects)
            {
                writer.WriteLine($"{prefix}\t{apoSlot.Key} [{KeyToName[apoSlot.Key]}]:");
                foreach (var apo in apoSlot.Value)
                {
                    apo.WriteDetailed(writer, $"{prefix}\t\t");
                }
            }

            foreach (var field in ProcessingModesForStreamingProperties)
            {
                string[]? value = GetProperty<PropertyKey, string[]>(field, PropertyLocation.Effects);
                if (null == value)
                {
                    continue;
                }

                writer.Write($"{prefix}\t{field} [{KeyToName[field]}]: ");

                foreach (var mode in value)
                {
                    if (ProcessingModesForStreaming.TryGetValue(Guid.Parse(mode), out string? modeName))
                    {
                        writer.Write($"{modeName} ");
                    }
                    else
                    {
                        writer.Write($"{mode} ");
                    }
                }
                writer.WriteLine();
            }

            writer.WriteLine($"{prefix}Unresolved AudioProcessingObjects : {UnresolvedAudioProcessingObjects.Count}");
            foreach (var clsid in UnresolvedAudioProcessingObjects)
            {
                writer.WriteLine($"{prefix}\t{clsid}");
            }
        }
    }
}
