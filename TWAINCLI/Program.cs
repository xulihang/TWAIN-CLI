using NTwain;
using NTwain.Data;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Controls;
using System.Windows.Forms;

namespace TWAINCLI
{
    internal class Program
    {
        private static TwainSession _session;
        private static string _outputPath = @"C:\Users\hp\";
        private static int _scanCount = 0;

        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            NTwain.PlatformInfo.Current.PreferNewDSM = false;

            // Parse command line arguments
            var config = ParseArguments(args);
            _outputPath = config.OutputPath;

            // If only listing scanners, execute and exit
            if (config.ListOnly)
            {
                ListScanners();
                return;
            }

            try
            {
                // Create application identity [[31]]
                var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, Assembly.GetExecutingAssembly());

                // Create TWAIN session
                _session = new TwainSession(appId);

                // Register event handlers [[31]]
                _session.TransferReady += TwainSession_TransferReady;
                _session.DataTransferred += TwainSession_DataTransferred;
                _session.SourceDisabled += TwainSession_SourceDisabled;

                // Open session
                var openResult = _session.Open();
                if (openResult != ReturnCode.Success)
                {
                    Console.WriteLine($"Failed to open TWAIN session: {openResult}");
                    return;
                }

                Console.WriteLine("TWAIN session opened");

                // Find and select scanner
                DataSource selectedSource = null;
                var sources = _session.GetSources().ToList();

                if (sources.Count == 0)
                {
                    Console.WriteLine("No scanners found");
                    _session.Close();
                    return;
                }

                // Select scanner by name (if -d parameter specified)
                if (!string.IsNullOrEmpty(config.ScannerName))
                {
                    selectedSource = sources.FirstOrDefault(s =>
                        s.Name.IndexOf(config.ScannerName, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (selectedSource == null)
                    {
                        Console.WriteLine($"No scanner found with name containing '{config.ScannerName}'");
                        _session.Close();
                        return;
                    }
                }
                else
                {
                    // Default to first available scanner
                    selectedSource = sources.FirstOrDefault();
                }

                if (selectedSource != null)
                {
                    Console.WriteLine($"\nSelected scanner: {selectedSource.Name}");

                    // Open data source [[21]]
                    var openDsResult = selectedSource.Open();
                    if (openDsResult == ReturnCode.Success)
                    {
                        if (config.ShowUI)
                        {
                            Console.WriteLine("Scanner opened, showing native UI...");
                            // Enable data source for scanning (with UI mode) [[21]]
                            var enableResult = selectedSource.Enable(SourceEnableMode.ShowUI, false, IntPtr.Zero);
                            if (enableResult != ReturnCode.Success)
                            {
                                Console.WriteLine($"Failed to enable scanner: {enableResult}");
                                selectedSource.Close();
                                return;
                            }
                        }
                        else
                        {
                            Console.WriteLine("Scanner opened, configuring parameters...");

                            // Apply configuration parameters
                            ConfigureSource(selectedSource, config);

                            // Enable data source for scanning (no UI mode) [[21]]
                            var enableResult = selectedSource.Enable(SourceEnableMode.NoUI, false, IntPtr.Zero);
                            if (enableResult != ReturnCode.Success)
                            {
                                Console.WriteLine($"Failed to enable scanner: {enableResult}");
                                selectedSource.Close();
                                return;
                            }
                        }

                        // Run message loop, wait for scan completion
                        Application.Run();
                    }
                    else
                    {
                        Console.WriteLine($"Failed to open scanner: {openDsResult}");
                        selectedSource.Close();
                    }
                }
                else
                {
                    Console.WriteLine("No suitable scanner found");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                if (_session != null)
                {
                    _session.Close();
                    Console.WriteLine("TWAIN session closed");
                }
            }
        }

        /// <summary>
        /// Configure scanner TWAIN Capabilities [[12]][[13]]
        /// </summary>
        private static void ConfigureSource(DataSource source, ScanConfig config)
        {
            // 1. Configure feeder/flatbed
            if (source.Capabilities.CapFeederEnabled.IsSupported)
            {
                bool useFeeder = config.SourceType?.ToLower() == "feeder";
                source.Capabilities.CapFeederEnabled.SetValue(useFeeder ? BoolType.True : BoolType.False);
            }

            // 2. Configure light path (positive/negative/flatbed) [[19]]
            if (source.Capabilities.ICapLightPath.IsSupported && !string.IsNullOrEmpty(config.SourceType))
            {
                switch (config.SourceType.ToLower())
                {
                    case "negative":
                        source.Capabilities.ICapLightPath.SetValue(LightPath.Transmissive);
                        if (source.Capabilities.ICapFilmType.IsSupported)
                            source.Capabilities.ICapFilmType.SetValue(FilmType.Negative);
                        break;
                    case "positive":
                        source.Capabilities.ICapLightPath.SetValue(LightPath.Transmissive);
                        if (source.Capabilities.ICapFilmType.IsSupported)
                            source.Capabilities.ICapFilmType.SetValue(FilmType.Positive);
                        break;
                    case "flatbed":
                    case "reflected":
                        source.Capabilities.ICapLightPath.SetValue(LightPath.Reflective);
                        break;
                }
            }

            // 3. Configure color mode (bw/gray/color) [[19]]
            if (source.Capabilities.ICapPixelType.IsSupported && !string.IsNullOrEmpty(config.ColorMode))
            {
                Console.WriteLine("Setting color mode");
                PixelType? pixelType = null;
                string mode = config.ColorMode.ToLower();
                if (mode == "bw" || mode == "blackandwhite" || mode == "1bit")
                {
                    pixelType = PixelType.BlackWhite;
                }
                else if (mode == "gray" || mode == "grayscale")
                {
                    pixelType = PixelType.Gray;
                }
                else if (mode == "color" || mode == "rgb")
                {
                    pixelType = PixelType.RGB;
                }
                // Keep null for other cases
                if (pixelType.HasValue && source.Capabilities.ICapPixelType.CanSet &&
                    source.Capabilities.ICapPixelType.GetValues().Contains(pixelType.Value))
                {
                    source.Capabilities.ICapPixelType.SetValue(pixelType.Value);
                }
            }

            // 4. Configure resolution [[13]][[14]]
            if (config.Resolution > 0)
            {
                Console.WriteLine("Setting resolution");
                if (source.Capabilities.ICapXResolution.IsSupported && source.Capabilities.ICapXResolution.CanSet)
                    source.Capabilities.ICapXResolution.SetValue(config.Resolution);
                if (source.Capabilities.ICapYResolution.IsSupported && source.Capabilities.ICapYResolution.CanSet)
                    source.Capabilities.ICapYResolution.SetValue(config.Resolution);
            }

            // 5. Configure duplex scanning [[12]]
            if (source.Capabilities.CapDuplexEnabled.IsSupported && source.Capabilities.CapDuplexEnabled.CanSet)
            {
                if (config.Duplex)
                {
                    Console.WriteLine("Enabling duplex scanning");
                    source.Capabilities.CapDuplexEnabled.SetValue(BoolType.True);
                }
                else
                {
                    Console.WriteLine("Disabling duplex scanning");
                    source.Capabilities.CapDuplexEnabled.SetValue(BoolType.False);
                }
            }

            // 6. Configure scan area (Frame) [[1]][[4]]
            // TWAIN Frame coordinates unit is determined by ICAP_UNITS, default is Inches
            // Note: Using ICapFrame (singular), not ICapFrames
            if (config.HasArea && source.Capabilities.ICapFrames.IsSupported && source.Capabilities.ICapFrames.CanSet)
            {
                Console.WriteLine("Setting scan area");
                var frame = new TWFrame
                {
                    Left = config.PageLeft,
                    Top = config.PageTop,
                    Right = config.PageLeft + config.PageWidth,
                    Bottom = config.PageTop + config.PageHeight
                };
                source.Capabilities.ICapFrames.SetValue(frame);
            }

            // 7. Enable lamp (required by some scanners)
            if (source.Capabilities.ICapLampState.IsSupported && source.Capabilities.ICapLampState.CanSet)
            {
                source.Capabilities.ICapLampState.SetValue(BoolType.True);
            }
        }

        /// <summary>
        /// List all available scanners
        /// </summary>
        private static void ListScanners()
        {
            try
            {
                var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, Assembly.GetExecutingAssembly());
                var session = new TwainSession(appId);

                var openResult = session.Open();
                if (openResult != ReturnCode.Success)
                {
                    Console.WriteLine($"Failed to open TWAIN session: {openResult}");
                    return;
                }

                var sources = session.GetSources().ToList();

                if (sources.Count == 0)
                {
                    Console.WriteLine("No scanners found");
                }
                else
                {
                    Console.WriteLine("Available scanners:");
                    foreach (var source in sources)
                    {
                        Console.WriteLine($"  - {source.Name} (ID: {source.Id})");
                    }
                }

                session.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing scanners: {ex.Message}");
            }
        }

        /// <summary>
        /// Parse command line arguments
        /// </summary>
        private static ScanConfig ParseArguments(string[] args)
        {
            var config = new ScanConfig
            {
                OutputPath = @"C:\Users\hp\",
                Resolution = 200,      // Default resolution 200 DPI
                ColorMode = "color",   // Default color
                SourceType = "flatbed", // Default flatbed
                ScannerName = ""
            };

            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];

                switch (arg)
                {
                    case "-l": // Page Left
                        if (i + 1 < args.Length && float.TryParse(args[++i], out float left))
                            config.PageLeft = left;
                        break;

                    case "-t": // Page Top (changed from -y to avoid conflict with height)
                        if (i + 1 < args.Length && float.TryParse(args[++i], out float top))
                            config.PageTop = top;
                        break;

                    case "-x": // Page Width
                        if (i + 1 < args.Length && float.TryParse(args[++i], out float width))
                            config.PageWidth = width;
                        break;

                    case "-y": // Page Height
                        if (i + 1 < args.Length && float.TryParse(args[++i], out float height))
                            config.PageHeight = height;
                        break;

                    case "-o": // Output Path
                        if (i + 1 < args.Length)
                        {
                            config.OutputPath = args[++i];
                            // Ensure path ends with separator
                            if (!config.OutputPath.EndsWith(@"\") && !config.OutputPath.EndsWith("/"))
                                config.OutputPath += Path.DirectorySeparatorChar;
                        }
                        break;

                    case "-m": // Color Mode
                        if (i + 1 < args.Length)
                            config.ColorMode = args[++i];
                        break;

                    case "-r": // Resolution
                        if (i + 1 < args.Length && int.TryParse(args[++i], out int res) && res > 0)
                            config.Resolution = res;
                        break;

                    case "-s": // Source Type
                        if (i + 1 < args.Length)
                            config.SourceType = args[++i];
                        break;

                    case "-d": // Scanner Name
                        if (i + 1 < args.Length)
                            config.ScannerName = args[++i];
                        break;

                    case "--duplex": // Duplex
                        config.Duplex = true;
                        break;

                    case "-L": // List scanners (uppercase)
                        config.ListOnly = true;
                        break;
                    case "--showUI": // Show native UI
                        config.ShowUI = true;
                        break;
                    case "-?":
                    case "--help":
                        PrintHelp();
                        Environment.Exit(0);
                        break;
                }
            }

            // Check if scan area is configured
            config.HasArea = config.PageWidth > 0 && config.PageHeight > 0;
            if (config.ScannerName == "" && config.ListOnly == false)
            {
                PrintHelp();
                Environment.Exit(0);
            }
            return config;
        }

        /// <summary>
        /// Print help information
        /// </summary>
        private static void PrintHelp()
        {
            Console.WriteLine("TWAIN CLI Scanner - Usage:");
            Console.WriteLine("  -L                List all available scanners");
            Console.WriteLine("  -s <name>         Specify source type: feeder, positive, negative, flatbed");
            Console.WriteLine("  -d <name>         Specify scanner by name (partial match supported)");
            Console.WriteLine("  -m <mode>         Specify color mode: bw, gray, color");
            Console.WriteLine("  -r <resolution>   Specify resolution in DPI (default: 200)");
            Console.WriteLine("  -o <path>         Output file/folder path");
            Console.WriteLine("  -l <left>         Scan area left margin (inches)");
            Console.WriteLine("  -t <top>          Scan area top margin (inches)");
            Console.WriteLine("  -x <width>        Scan area width (inches)");
            Console.WriteLine("  -y <height>       Scan area height (inches)");
            Console.WriteLine("  --duplex          Enable duplex scanning");
            Console.WriteLine("  --showUI          Show native scanner UI");
            Console.WriteLine("  -?, --help        Show this help message");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  TWAINCLI.exe -L");
            Console.WriteLine("  TWAINCLI.exe -d \"EPSON\" -s feeder -m color -r 300 -o C:\\Scans\\");
            Console.WriteLine("  TWAINCLI.exe -s negative -l 0 -t 0 -x 8.5 -y 11 --duplex");
        }

        // ===== Event Handlers =====

        private static void TwainSession_SourceDisabled(object sender, EventArgs e)
        {
            Console.WriteLine("Scan completed or cancelled by user");
            // Exit message loop
            Application.Exit();
        }

        private static void TwainSession_TransferReady(object sender, TransferReadyEventArgs e)
        {
            Console.WriteLine("Preparing to transfer image...");
            e.CancelAll = false;  // Ensure transfer is not cancelled
        }

        private static void TwainSession_DataTransferred(object sender, DataTransferredEventArgs e)
        {
            try
            {
                Console.WriteLine($"Data received, transfer type: {e.TransferType}");

                if (e.TransferType == XferMech.Native)
                {
                    // Handle Native transfer type (used by most scanners) [[31]]
                    using (var stream = e.GetNativeImageStream())
                    {
                        if (stream != null)
                        {
                            using (var image = System.Drawing.Image.FromStream(stream))
                            {
                                // Ensure output directory exists
                                Directory.CreateDirectory(_outputPath);

                                // Generate filename with timestamp
                                string fileName = Path.Combine(_outputPath,
                                    $"scan_{DateTime.Now:yyyyMMdd_HHmmss}_{_scanCount++}.jpg");
                                image.Save(fileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                                Console.WriteLine($"Image saved to: {fileName}");
                            }
                        }
                    }
                }
                else if (e.TransferType == XferMech.File)
                {
                    // Handle file transfer type
                    Console.WriteLine($"File transfer mode, file path: {e.FileDataPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing transferred data: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Scan configuration class
    /// </summary>
    internal class ScanConfig
    {
        public bool ShowUI { get; set; }              // --showUI: Use native scanner UI
        public bool ListOnly { get; set; }            // -L: List scanners only
        public string ScannerName { get; set; }       // -d: Scanner name
        public string SourceType { get; set; } = "flatbed";  // -s: Source type
        public string ColorMode { get; set; } = "color";     // -m: Color mode
        public int Resolution { get; set; } = 200;      // -r: Resolution
        public string OutputPath { get; set; }          // -o: Output path
        public float PageLeft { get; set; } = 0;        // -l: Left margin
        public float PageTop { get; set; } = 0;         // -t: Top margin
        public float PageWidth { get; set; } = 0;       // -x: Width
        public float PageHeight { get; set; } = 0;      // -y: Height
        public bool HasArea { get; set; }               // Whether scan area is configured
        public bool Duplex { get; set; }                // --duplex: Enable duplex scanning
    }
}