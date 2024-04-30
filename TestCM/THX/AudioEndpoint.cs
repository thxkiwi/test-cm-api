//#define ConsoleLoggingEnabled 

using Microsoft.Win32;
using System.Reflection.Metadata.Ecma335;
using System.Text;
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

        private static readonly Dictionary<PropertyKey, string> KeyToName = (new Dictionary<PROPERTYKEY, string>(){
            { PInvoke.PKEY_CompositeFX_StreamEffectClsid, "PKEY_CompositeFX_StreamEffectClsid" },
            { PInvoke.PKEY_CompositeFX_ModeEffectClsid, "PKEY_CompositeFX_ModeEffectClsid" },
            { PInvoke.PKEY_CompositeFX_EndpointEffectClsid, "PKEY_CompositeFX_EndpointEffectClsid" },
            { PInvoke.PKEY_CompositeFX_Offload_StreamEffectClsid, "PKEY_CompositeFX_Offload_StreamEffectClsid" },
            { PInvoke.PKEY_CompositeFX_Offload_ModeEffectClsid, "PKEY_CompositeFX_Offload_ModeEffectClsid" },
            { PInvoke.PKEY_SFX_ProcessingModes_Supported_For_Streaming, "PKEY_SFX_ProcessingModes_Supported_For_Streaming" },
            { PInvoke.PKEY_SFX_Offload_ProcessingModes_Supported_For_Streaming, "PKEY_SFX_Offload_ProcessingModes_Supported_For_Streaming" },
            { PInvoke.PKEY_MFX_ProcessingModes_Supported_For_Streaming, "PKEY_MFX_ProcessingModes_Supported_For_Streaming" },
            { PInvoke.PKEY_MFX_Offload_ProcessingModes_Supported_For_Streaming, "PKEY_MFX_Offload_ProcessingModes_Supported_For_Streaming" },
            { PInvoke.PKEY_EFX_ProcessingModes_Supported_For_Streaming, "PKEY_EFX_ProcessingModes_Supported_For_Streaming" },
            { PInvoke.PKEY_FX_Association, "PKEY_FX_Association" },
            { PInvoke.PKEY_AudioEndpoint_Association, "PKEY_AudioEndpoint_Association" }
        }).ToDictionary(kv => PropertyKey.From(kv.Key), kv => kv.Value);

        private static readonly Dictionary<Guid, string> GuidToName = new()
        {
            // KSNODETYPE_*
            { Guid.Empty, "KSNODETYPE_ANY" },
            { Guid.Parse("DFF21BE0-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_INPUT_UNDEFINED" },
            { Guid.Parse("DFF21BE1-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_MICROPHONE" },
            { Guid.Parse("DFF21BE2-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_DESKTOP_MICROPHONE" },
            { Guid.Parse("DFF21BE3-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_PERSONAL_MICROPHONE" },
            { Guid.Parse("DFF21BE4-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_OMNI_DIRECTIONAL_MICROPHONE" },
            { Guid.Parse("DFF21BE5-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_MICROPHONE_ARRAY" },
            { Guid.Parse("DFF21BE6-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_PROCESSING_MICROPHONE_ARRAY" },
            { Guid.Parse("DFF21CE0-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_OUTPUT_UNDEFINED" },
            { Guid.Parse("DFF21CE1-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_SPEAKER" },
            { Guid.Parse("DFF21CE2-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_HEADPHONES" },
            { Guid.Parse("DFF21CE3-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_HEAD_MOUNTED_DISPLAY_AUDIO" },
            { Guid.Parse("DFF21CE4-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_DESKTOP_SPEAKER" },
            { Guid.Parse("DFF21CE5-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_ROOM_SPEAKER" },
            { Guid.Parse("DFF21CE6-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_COMMUNICATION_SPEAKER" },
            { Guid.Parse("DFF21DE0-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_BIDIRECTIONAL_UNDEFINED" },
            { Guid.Parse("DFF21DE1-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_HANDSET" },
            { Guid.Parse("DFF21DE2-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_HEADSET" },
            { Guid.Parse("DFF21DE3-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_SPEAKERPHONE_NO_ECHO_REDUCTION" },
            { Guid.Parse("DFF21DE4-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_ECHO_SUPPRESSING_SPEAKERPHONE" },
            { Guid.Parse("DFF21DE5-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_ECHO_CANCELING_SPEAKERPHONE" },
            { Guid.Parse("DFF21EE0-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_TELEPHONY_UNDEFINED" },
            { Guid.Parse("DFF21EE1-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_PHONE_LINE" },
            { Guid.Parse("DFF21EE2-F70F-11D0-B917-00A0C9223196"), "KSNODETYPE_TELEPHONE" },

            // AUDIO_SIGNALPROCESSINGMODE_*
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

            // Get all of the AudioProcessingObjectInfs that are available on
            // the system.
            var systemAudioProcessingObjectInfs = DeviceClass.GetDeviceIds(DeviceClass.AudioProcessingObject)
                .Select(s => new AudioProcessingObjectInf(DeviceNode.GetByInstanceId(s)));

            // Form the set union of the AudioProcessingObjects obtained from
            // all INFs and map them from CLSID to AudioProcessingObject.
            var systemAudioProcessingObjectsByClsid 
                = systemAudioProcessingObjectInfs
                .SelectMany(s => s.AudioProcessingObjects)
                .ToDictionary(apo => apo.ClassId, apo => apo);

            // Build the Dictionary of AudioProcessingObjects
            foreach (var apoSlot in (new []{
                        PInvoke.PKEY_CompositeFX_StreamEffectClsid,
                        PInvoke.PKEY_CompositeFX_ModeEffectClsid,
                        PInvoke.PKEY_CompositeFX_EndpointEffectClsid,
                        PInvoke.PKEY_CompositeFX_Offload_StreamEffectClsid,
                        PInvoke.PKEY_CompositeFX_Offload_ModeEffectClsid
                    }).Select(x => PropertyKey.From(x))
                )
            {
#if ConsoleLoggingEnabled
                Console.Error.Write($"Processing {KeyToName[apoSlot]}: ");
#endif
                var clsidsForAPOSlot = (GetProperty<PropertyKey, string[]>(
                        apoSlot,
                        PropertyLocation.Effects
                    ) ?? []);

#if ConsoleLoggingEnabled
                Console.Error.WriteLine($"Found {clsidsForAPOSlot.Count()} [{
                    string.Join(", ", 
                        clsidsForAPOSlot.Select(g => g.ToString("B")))
                    }]");
#endif
                foreach (var clsidInSlotStr in clsidsForAPOSlot ?? [])
                {
                    if (null == clsidInSlotStr)
                    {
                        continue;
                    }

                    Guid clsidInSlot = Guid.Parse(clsidInSlotStr);

                    // Look for an APO for this CLSID in the sibling
                    // AudioProcessingObjectInfs.
                    if (!systemAudioProcessingObjectsByClsid.TryGetValue(
                            clsidInSlot, 
                            out AudioProcessingObject? apo))
                    {
                        // An APO with the given CLSID was not found in any of
                        // the sibling AudioProcessingObjectINF's software keys.
                        //
                        // This is not an error condition. The CLSID may be 
                        // registered in HKCR, or by an AudioProcessingObject INF
                        // that is not a sibling of this device.
                        //
                        // Attempt to instantiate directly by CLSID which reads from HKCR.
                        try
                        {
                            apo = new AudioProcessingObject(clsidInSlot);
                        }
                        catch (KeyNotFoundException)
                        {
                            // The CLSID was not found in HKCR.
                            UnresolvedAudioProcessingObjects.Add(clsidInSlot);
                            continue;
                        }
                        catch (Exception e)
                        {
                            Console.Error.WriteLine(
                                $"Unable to resolve AudioProcessingObject {clsidInSlot.ToString("B")}: {e}");

                            UnresolvedAudioProcessingObjects.Add(clsidInSlot);
                            continue;
                        }
                    }

                    // The CLSID was resolved.
                    //
                    // Add it to the AudioProcessingObjects for the given slot.
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
                return null;
            }

            return propertyStore.GetValue(key.ToString()) as ValueT;
        }

        public override string ToString()
        {
            return $"({DataFlow}) {MMDeviceID}";
        }

        public void WriteDetailed(TextWriter writer, string prefix = "")
        {
            PropertyKey AudioEndpointAssociation
                = PropertyKey.From(PInvoke.PKEY_AudioEndpoint_Association);

            string? endpointAssociationNodeType
                = GetProperty<PropertyKey, string>(
                    AudioEndpointAssociation,
                    PropertyLocation.Endpoint);

            string endpointAssociationNodeTypeName = endpointAssociationNodeType switch
            {
                null => "<unknown>",
                _ => GuidToName.TryGetValue(
                        Guid.Parse(endpointAssociationNodeType),
                        out string ? name)
                    ? name
                    : "<unknown>"
            };
           
            StringBuilder sb = new();
            sb.Append($"{prefix}{AudioEndpointAssociation} [{KeyToName[AudioEndpointAssociation]}] : ");
            sb.Append($"{endpointAssociationNodeType ?? "<unknown>"} [{endpointAssociationNodeTypeName}]");
            writer.WriteLine(sb.ToString());

            PropertyKey FxAssociation
                = PropertyKey.From(PInvoke.PKEY_FX_Association);

            string? fxAssociationNodeType
                = GetProperty<PropertyKey, string>(
                    FxAssociation,
                    PropertyLocation.Effects);

            string fxAssociationNodeTypeName = fxAssociationNodeType switch
            {
                null => "<unknown>",
                _ => GuidToName.TryGetValue(
                        Guid.Parse(fxAssociationNodeType),
                        out string ? name)
                    ? name
                    : "<unknown>"
            };

            sb = new();
            sb.Append($"{prefix}{FxAssociation} [{KeyToName[FxAssociation]}] : ");
            sb.Append($"{fxAssociationNodeType ?? "<unknown>"} [{fxAssociationNodeTypeName}]");
            writer.WriteLine(sb.ToString());

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
                    if (GuidToName.TryGetValue(Guid.Parse(mode), out string? modeName))
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
