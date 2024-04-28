using Microsoft.Win32;
using Windows.Wdk;

namespace THX
{
    public class AudioProcessingObject
    {
        private static readonly RegistryKey FallbackRegistryRoot
            = Microsoft.Win32.Registry.ClassesRoot;

        private static readonly string FallbackClsidRegBase
            = "CLSID";

        public string Source { get; private set; }

        public string Name { get; private set; }

        public Guid ClassId { get; private set; }

        public enum COMServerType
        {
            InProcServer32,
            LocalServer32
        }

        public COMServerType ServerType { get; private set; }

        public string ServerPath { get; private set; }

        public enum COMThreadingModel
        {
            Apartment,
            Free,
            Both,
            Neutral,
            Single
        }

        public COMThreadingModel ThreadingModel { get; private set; }

        public long MajorVersion { get; private set; }

        public long MinorVersion { get; private set; }

        public long Flags { get; private set; }

        public long MinInputConnections { get; private set; }

        public long MaxInputConnections { get; private set; }

        public long MinOutputConnections { get; private set; }

        public long MaxOutputConnections { get; private set; }

        public long MaxInstances { get; private set; }

        public List<Guid> APOInterfaces { get; private set; }

        private static readonly string MissingFriendlyName = "<FriendlyName Not Found>";

        private static readonly string MissingServerPath = "<ServerPath Not Found>";

        /// \details Searches the given root for the AudioProcessingObject with 
        /// the given clsid using 
        /// <ul>
        /// <li>&lt;root&gt;&lt;\&lt;clsidRegBase&gt;&lt;\clsid&gt;, and</li>
        /// <li>&lt;root&gt;&lt;\AudioEngine\AudioProcessingObjects\&lt;clsid&gt;</li>
        /// </ul>
        public AudioProcessingObject(RegistryKey root,
            Guid clsid,
            string clsidRegBase = "Classes\\CLSID")
        {
            Source = $"{root.Name}\\{clsidRegBase}\\{clsid.ToString("B")}";

            string clsidRegStr = clsid.ToString("B");
            string clsidRegPath = $"{clsidRegBase}\\{clsidRegStr}";
            var clsidKey = root.OpenSubKey(
                    clsidRegPath,
                    RegistryKeyPermissionCheck.ReadSubTree,
                    System.Security.AccessControl.RegistryRights.ReadKey);
            if (clsidKey == null)
            {
                throw new KeyNotFoundException($"Registry path {clsidRegPath} not found.");
            }

            Name = Registry.GetValue<string, string>(clsidKey, null, MissingFriendlyName)
                ?? MissingFriendlyName;

            ClassId = clsid;

            foreach (var serverType in Enum.GetNames(typeof(COMServerType)))
            {
                var serverKey = clsidKey.OpenSubKey(serverType);
                if (null == serverKey)
                {
                    continue;
                }

                ServerType = Enum.Parse<COMServerType>(serverType);
                ServerPath = Registry.GetValue<string, string>(serverKey, null, MissingServerPath)
                    ?? MissingServerPath;

                string? threadingModelStr = Registry.GetValue<string, string>(serverKey, "ThreadingModel", null);
                COMThreadingModel threadingModel;
                if (!Enum.TryParse(threadingModelStr, out threadingModel))
                {
                    throw new InvalidDataException($"AudioProcessingObject {clsidRegStr} has an unrecognized ThreadingModel ({threadingModelStr})");
                }
                ThreadingModel = threadingModel;
                break;
            }

            if (null == ServerPath)
            {
                throw new KeyNotFoundException($"ServerPath not found for AudioProcessingObject {clsidRegStr}");
            }

            string audioEngineRegPath = $"AudioEngine\\AudioProcessingObjects\\{clsidRegStr}";
            var audioEngine = root.OpenSubKey(audioEngineRegPath);
            if (null == audioEngine)
            {
                throw new KeyNotFoundException($"Registry path {audioEngineRegPath} not found.");
            }

            const long DefaultLongValue = -1;
            MajorVersion = Registry.GetValueTypeValue<string, int>(
                audioEngine, "MajorVersion", null)
                ?? DefaultLongValue;

            MinorVersion = Registry.GetValueTypeValue<string, int>(
                audioEngine, "MinorVersion", null)
                ?? DefaultLongValue;

            Flags = Registry.GetValueTypeValue<string, int>(
                audioEngine, "Flags", null)
                ?? DefaultLongValue;

            MinInputConnections = Registry.GetValueTypeValue<string, int>(
                audioEngine, "MinInputConnections", null)
                ?? DefaultLongValue;

            MaxInputConnections = Registry.GetValueTypeValue<string, int>(
                audioEngine, "MaxInputConnections", null)
                ?? DefaultLongValue;

            MinOutputConnections = Registry.GetValueTypeValue<string, int>(
                audioEngine, "MinOutputConnections", null)
                ?? DefaultLongValue;

            MaxOutputConnections = Registry.GetValueTypeValue<string, int>(
                audioEngine, "MaxOutputConnections", null)
                ?? DefaultLongValue;

            MaxInstances = Registry.GetValueTypeValue<string, int>(
                audioEngine, "MaxInstances", null)
                ?? DefaultLongValue;

            long numInterfaces = Registry.GetValueTypeValue<string, int>(
                audioEngine, "NumAPOInterfaces", null)
                ?? DefaultLongValue;

            APOInterfaces = new List<Guid>();
            for (long i = 0; i < numInterfaces; i++)
            {
                string interfaceStr = $"APOInterface{i}";

                if (Guid.TryParse(Registry.GetValue<string, string>(
                    audioEngine, interfaceStr, null), 
                        out Guid interfaceGuid))
                {
                    APOInterfaces.Add(interfaceGuid);
                }
            }
        }

        /// \brief Constructor to fallback to when an Endpoint cannot find an
        /// APO from the AudioProcessingObjectInf.SoftwareKey.
        public AudioProcessingObject(Guid clsid)
            : this(FallbackRegistryRoot, clsid, FallbackClsidRegBase)
        { }

        public override string ToString()
        {
            return $"{ClassId.ToString("B")} {Name}";
        }

        [Flags]
        private enum APO_FLAGS
        {
            /// <summary>Indicates that there are no flags enabled for this APO.</summary>
            APO_FLAG_NONE = 0,
			/// <summary>Indicates that this APO can perform in-place processing. This allows the processor to use a common buffer for input and output.</summary>
			APO_FLAG_INPLACE = 1,
			/// <summary>Indicates that the samples per frame for the input and output connections must match.</summary>
			APO_FLAG_SAMPLESPERFRAME_MUST_MATCH = 2,
			/// <summary>Indicates that the frames per second for the input and output connections must match.</summary>
			APO_FLAG_FRAMESPERSECOND_MUST_MATCH = 4,
			/// <summary>Indicates that bits per sample AND bytes per sample container for the  input and output connections must match.</summary>
			APO_FLAG_BITSPERSAMPLE_MUST_MATCH = 8,
			/// <summary></summary>
			APO_FLAG_MIXER = 16
        }

        public void WriteDetailed(TextWriter writer, string indent = "")
        {
            writer.WriteLine($"{indent}{ToString()}");
            writer.WriteLine($"{indent}  Source: {Source}");
            writer.WriteLine($"{indent}  ServerType: {ServerType}");
            writer.WriteLine($"{indent}  ServerPath: {ServerPath}");
            writer.WriteLine($"{indent}  ThreadingModel: {ThreadingModel}");
            writer.WriteLine($"{indent}  MajorVersion: {MajorVersion}");
            writer.WriteLine($"{indent}  MinorVersion: {MinorVersion}");
            writer.Write($"{indent}  Flags: 0x{Flags:x8} [");
            if (0 == Flags)
            {
                writer.Write("APO_FLAG_NONE");
            }

            var flagsSet = Enum.GetValues(typeof(APO_FLAGS))
                .Cast<APO_FLAGS>()
                .Where(flag => 0 != ((APO_FLAGS)(Flags) & flag))
                .Select(flag => flag.ToString())
                .ToList();
            writer.Write(string.Join(" | ", flagsSet));
            
            writer.WriteLine("]");
            writer.WriteLine($"{indent}  MinInputConnections: {MinInputConnections}");
            writer.WriteLine($"{indent}  MaxInputConnections: {MaxInputConnections}");
            writer.WriteLine($"{indent}  MinOutputConnections: {MinOutputConnections}");
            writer.WriteLine($"{indent}  MaxOutputConnections: {MaxOutputConnections}");
            writer.WriteLine($"{indent}  MaxInstances: {MaxInstances}");
            writer.WriteLine($"{indent}  APOInterfaces:");
            foreach (var interfaceGuid in APOInterfaces)
            {
                writer.WriteLine($"{indent}    {interfaceGuid.ToString("B")}");
            }


        }
    }
}
