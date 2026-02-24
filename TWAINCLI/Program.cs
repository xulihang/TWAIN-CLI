using NTwain;
using NTwain.Data;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace TWAINCLI
{
    internal class Program
    {
        private static TwainSession _session;
        private static DataSource _selectedSource;          // [新增] 保存数据源引用用于取消
        private static string _outputPath = @"C:\Users\hp\";
        private static int _scanCount = 0;
        private static volatile bool _cancelRequested = false;  // [新增] 取消标志（volatile保证线程可见性）
        private static FileSystemWatcher _cancelWatcher = null; // [新增] 文件监控器

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
                // Create application identity
                var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, Assembly.GetExecutingAssembly());

                // Create TWAIN session
                _session = new TwainSession(appId);

                // Register event handlers
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
                    _selectedSource = selectedSource;  // [新增] 保存引用用于后续取消

                    // Open data source
                    var openDsResult = selectedSource.Open();
                    if (openDsResult == ReturnCode.Success)
                    {
                        if (config.ShowUI)
                        {
                            Console.WriteLine("Scanner opened, showing native UI...");
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

                            // Enable data source for scanning (no UI mode)
                            var enableResult = selectedSource.Enable(SourceEnableMode.NoUI, false, IntPtr.Zero);
                            if (enableResult != ReturnCode.Success)
                            {
                                Console.WriteLine($"Failed to enable scanner: {enableResult}");
                                selectedSource.Close();
                                return;
                            }
                        }

                        // [新增] 启动取消文件监控
                        StartCancelMonitoring(config.CancelFileName);

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
                // [新增] 停止取消文件监控
                StopCancelMonitoring();

                if (_session != null)
                {
                    _session.Close();
                    Console.WriteLine("TWAIN session closed");
                }
            }
        }

        /// <summary>
        /// [新增] 启动取消文件监控
        /// </summary>
        private static void StartCancelMonitoring(string cancelFileName)
        {
            _cancelRequested = false;

            if (_cancelWatcher != null)
            {
                _cancelWatcher.Dispose();
                _cancelWatcher = null;
            }

            string watchDir = AppDomain.CurrentDomain.BaseDirectory;
            string fileName = string.IsNullOrEmpty(cancelFileName) ? "cancel" : cancelFileName;

            _cancelWatcher = new FileSystemWatcher(watchDir, fileName)
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
                EnableRaisingEvents = true
            };

            _cancelWatcher.Created += (s, e) =>
            {
                Console.WriteLine("⚠ Cancel file detected, requesting scan cancellation...");
                _cancelRequested = true;

                // 尝试立即取消数据源（如果已启用）
                if (_selectedSource != null && _selectedSource.IsOpen)
                {
                    try
                    {
                        _selectedSource.Close();
                        Console.WriteLine("🔌 Scanner source closed");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"⚠ Error closing source: {ex.Message}");
                    }
                }

                // 退出消息循环
                Application.Exit();
            };

            // 兼容：如果cancel文件已存在，也触发取消
            if (File.Exists(Path.Combine(watchDir, fileName)))
            {
                Console.WriteLine("⚠ Cancel file already exists, requesting cancellation...");
                _cancelRequested = true;
            }
        }

        /// <summary>
        /// [新增] 停止取消文件监控
        /// </summary>
        private static void StopCancelMonitoring()
        {
            if (_cancelWatcher != null)
            {
                _cancelWatcher.EnableRaisingEvents = false;
                _cancelWatcher.Dispose();
                _cancelWatcher = null;
            }

            // 可选：自动删除cancel文件，避免下次误触发
            try
            {
                string cancelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cancel");
                if (File.Exists(cancelPath))
                    File.Delete(cancelPath);
            }
            catch { /* 忽略删除异常 */ }
        }

        /// <summary>
        /// Configure scanner TWAIN Capabilities
        /// </summary>
        private static void ConfigureSource(DataSource source, ScanConfig config)
        {
            // 1. Configure feeder/flatbed
            if (source.Capabilities.CapFeederEnabled.IsSupported)
            {
                bool useFeeder = config.SourceType?.ToLower() == "feeder";
                source.Capabilities.CapFeederEnabled.SetValue(useFeeder ? BoolType.True : BoolType.False);
            }

            // 2. Configure light path (positive/negative/flatbed)
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

            // 3. Configure color mode (bw/gray/color)
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
                if (pixelType.HasValue && source.Capabilities.ICapPixelType.CanSet &&
                    source.Capabilities.ICapPixelType.GetValues().Contains(pixelType.Value))
                {
                    source.Capabilities.ICapPixelType.SetValue(pixelType.Value);
                }
            }

            // 4. Configure resolution
            if (config.Resolution > 0)
            {
                Console.WriteLine("Setting resolution");
                if (source.Capabilities.ICapXResolution.IsSupported && source.Capabilities.ICapXResolution.CanSet)
                    source.Capabilities.ICapXResolution.SetValue(config.Resolution);
                if (source.Capabilities.ICapYResolution.IsSupported && source.Capabilities.ICapYResolution.CanSet)
                    source.Capabilities.ICapYResolution.SetValue(config.Resolution);
            }

            // 5. Configure duplex scanning
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

            // 6. Configure scan area (Frame)
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
                Resolution = 200,
                ColorMode = "color",
                SourceType = "flatbed",
                ScannerName = "",
                CancelFileName = "cancel"  // [新增] 默认取消文件名
            };

            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];

                switch (arg)
                {
                    case "-l":
                        if (i + 1 < args.Length && float.TryParse(args[++i], out float left))
                            config.PageLeft = left;
                        break;

                    case "-t":
                        if (i + 1 < args.Length && float.TryParse(args[++i], out float top))
                            config.PageTop = top;
                        break;

                    case "-x":
                        if (i + 1 < args.Length && float.TryParse(args[++i], out float width))
                            config.PageWidth = width;
                        break;

                    case "-y":
                        if (i + 1 < args.Length && float.TryParse(args[++i], out float height))
                            config.PageHeight = height;
                        break;

                    case "-o":
                        if (i + 1 < args.Length)
                        {
                            config.OutputPath = args[++i];
                            if (!config.OutputPath.EndsWith(@"\") && !config.OutputPath.EndsWith("/"))
                                config.OutputPath += Path.DirectorySeparatorChar;
                        }
                        break;

                    case "-m":
                        if (i + 1 < args.Length)
                            config.ColorMode = args[++i];
                        break;

                    case "-r":
                        if (i + 1 < args.Length && int.TryParse(args[++i], out int res) && res > 0)
                            config.Resolution = res;
                        break;

                    case "-s":
                        if (i + 1 < args.Length)
                            config.SourceType = args[++i];
                        break;

                    case "-d":
                        if (i + 1 < args.Length)
                            config.ScannerName = args[++i];
                        break;

                    case "--duplex":
                        config.Duplex = true;
                        break;

                    case "-L":
                        config.ListOnly = true;
                        break;

                    case "--showUI":
                        config.ShowUI = true;
                        break;

                    case "--cancelFile":  // [新增] 自定义取消文件名
                        if (i + 1 < args.Length)
                            config.CancelFileName = args[++i];
                        break;

                    case "-?":
                    case "--help":
                        PrintHelp();
                        Environment.Exit(0);
                        break;
                }
            }

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
            Console.WriteLine("  -L                    List all available scanners");
            Console.WriteLine("  -s <name>             Specify source type: feeder, positive, negative, flatbed");
            Console.WriteLine("  -d <name>             Specify scanner by name (partial match supported)");
            Console.WriteLine("  -m <mode>             Specify color mode: bw, gray, color");
            Console.WriteLine("  -r <resolution>       Specify resolution in DPI (default: 200)");
            Console.WriteLine("  -o <path>             Output file/folder path");
            Console.WriteLine("  -l <left>             Scan area left margin (inches)");
            Console.WriteLine("  -t <top>              Scan area top margin (inches)");
            Console.WriteLine("  -x <width>            Scan area width (inches)");
            Console.WriteLine("  -y <height>           Scan area height (inches)");
            Console.WriteLine("  --duplex              Enable duplex scanning");
            Console.WriteLine("  --showUI              Show native scanner UI");
            Console.WriteLine("  --cancelFile <name>   Custom cancel trigger filename (default: cancel)");
            Console.WriteLine("  -?, --help            Show this help message");
            Console.WriteLine();
            Console.WriteLine("Cancel scanning:");
            Console.WriteLine("  Create a file named 'cancel' in the program directory:");
            Console.WriteLine("    cmd:     echo. > cancel");
            Console.WriteLine("    powershell: New-Item -Path . -Name cancel -ItemType File");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  TWAINCLI.exe -L");
            Console.WriteLine("  TWAINCLI.exe -d \"EPSON\" -s feeder -m color -r 300 -o C:\\Scans\\");
            Console.WriteLine("  TWAINCLI.exe -s negative -l 0 -t 0 -x 8.5 -y 11 --duplex");
            Console.WriteLine("  TWAINCLI.exe -d \"EPSON\" --cancelFile stop.scan  (use 'stop.scan' to cancel)");
        }

        // ===== Event Handlers =====

        private static void TwainSession_SourceDisabled(object sender, EventArgs e)
        {
            if (_cancelRequested)
                Console.WriteLine("✅ Scan cancelled by user");
            else
                Console.WriteLine("✅ Scan completed");

            // Exit message loop
            Application.Exit();
        }

        private static void TwainSession_TransferReady(object sender, TransferReadyEventArgs e)
        {
            // [新增] 检查是否请求取消
            if (_cancelRequested)
            {
                Console.WriteLine("🚫 Transfer cancelled by user request");
                e.CancelAll = true;
                return;
            }

            Console.WriteLine("Preparing to transfer image...");
            e.CancelAll = false;
        }

        private static void TwainSession_DataTransferred(object sender, DataTransferredEventArgs e)
        {
            // [新增] 如果已取消，跳过处理
            if (_cancelRequested)
            {
                Console.WriteLine("⚠ Data transfer ignored (cancel requested)");
                return;
            }

            try
            {
                Console.WriteLine($"Data received, transfer type: {e.TransferType}");

                if (e.TransferType == XferMech.Native)
                {
                    using (var stream = e.GetNativeImageStream())
                    {
                        if (stream != null)
                        {
                            using (var image = System.Drawing.Image.FromStream(stream))
                            {
                                Directory.CreateDirectory(_outputPath);

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
        public bool ShowUI { get; set; }
        public bool ListOnly { get; set; }
        public string ScannerName { get; set; }
        public string SourceType { get; set; } = "flatbed";
        public string ColorMode { get; set; } = "color";
        public int Resolution { get; set; } = 200;
        public string OutputPath { get; set; }
        public float PageLeft { get; set; } = 0;
        public float PageTop { get; set; } = 0;
        public float PageWidth { get; set; } = 0;
        public float PageHeight { get; set; } = 0;
        public bool HasArea { get; set; }
        public bool Duplex { get; set; }
        public string CancelFileName { get; set; } = "cancel";  // [新增] 取消触发文件名
    }
}