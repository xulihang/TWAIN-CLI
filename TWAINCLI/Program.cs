using NTwain;
using NTwain.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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

        // [新增] 内存传输进度状态
        private static long _estimatedTotalPixels = 0;          // 驱动不提供尺寸时，用扫描区域估算
        private static int _pageNumber = 0;
        private static bool _pageInProgress = false;
        private static bool _pageCompressed = false;
        private static int _pageWidth = 0;
        private static int _pageHeight = 0;
        private static int _pageBitsPerPixel = 0;
        private static PixelType _pagePixelType = PixelType.RGB;
        private static long _totalPixels = 0;
        private static long _transferredPixels = 0;
        private static long _transferredBytes = 0;
        private static readonly List<MemoryChunk> _chunks = new List<MemoryChunk>();
        private static MemoryStream _compressedBytes = null;
        private static Stopwatch _pageStopwatch = null;
        private static long _lastProgressReportMs = 0;
        private static int _lastReportedPercent = -1;

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
            // 0. 选择传输方式（内存传输才能按已传像素给出百分比进度）
            ConfigureTransferMode(source, config);

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
        /// [新增] 选择传输方式：只有 Memory 传输能按已传像素计算进度
        /// TWAIN 规范只强制要求 Native 传输，Memory 是可选能力，所以必须检查并回退
        /// </summary>
        private static void ConfigureTransferMode(DataSource source, ScanConfig config)
        {
            // 驱动不返回图像尺寸时（ImageWidth/ImageLength 为 -1），用扫描区域估算总像素
            if (config.HasArea && config.Resolution > 0)
            {
                _estimatedTotalPixels = (long)Math.Ceiling(config.PageWidth * config.Resolution) *
                                        (long)Math.Ceiling(config.PageHeight * config.Resolution);
            }

            if (config.NativeTransfer)
            {
                Console.WriteLine("Transfer mode: Native (requested), no progress percentage");
                return;
            }

            // Kyocera Ecosys 的内存传输有已知问题（与 NAPS2 相同的处理）
            if (ContainsIgnoreCase(source.Manufacturer, "Kyocera") &&
                (ContainsIgnoreCase(source.ProductFamily, "Ecosys") || ContainsIgnoreCase(source.Name, "Ecosys")))
            {
                Console.WriteLine("Transfer mode: Native (Kyocera Ecosys, memory transfer known to be broken)");
                return;
            }

            try
            {
                var cap = source.Capabilities.ICapXferMech;
                if (cap.IsSupported && cap.CanSet &&
                    cap.GetValues().Contains(XferMech.Memory) &&
                    cap.SetValue(XferMech.Memory) == ReturnCode.Success)
                {
                    Console.WriteLine("Transfer mode: Memory, progress enabled");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠ Could not switch to memory transfer: {ex.Message}");
            }

            Console.WriteLine("Transfer mode: Native (memory transfer not supported), no progress percentage");
        }

        private static bool ContainsIgnoreCase(string value, string part)
        {
            return !string.IsNullOrEmpty(value) &&
                   value.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0;
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

                    case "--native":  // [新增] 强制 Native 传输（不做内存拼图，不输出百分比）
                        config.NativeTransfer = true;
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
            Console.WriteLine("  --native              Force native transfer (no progress percentage)");
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
            // [新增] 内存传输模式下，最后一页要等数据源禁用才算接收完成
            if (_cancelRequested)
                DiscardCurrentPage();
            else if (_pageInProgress)
                SaveCurrentPage();

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

            // [新增] ADF 模式下新一页就绪，说明上一页（内存传输）已接收完毕
            if (_pageInProgress)
                SaveCurrentPage();

            BeginPage(e);
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
                if (e.TransferType == XferMech.Memory)
                {
                    // [新增] 内存传输每块数据触发一次事件，在这里累加进度
                    HandleMemoryChunk(e);
                    return;
                }

                Console.WriteLine($"Data received, transfer type: {e.TransferType}");
                _pageInProgress = false;

                if (e.TransferType == XferMech.Native)
                {
                    using (var stream = e.GetNativeImageStream())
                    {
                        if (stream != null)
                        {
                            using (var image = System.Drawing.Image.FromStream(stream))
                            {
                                SaveImage(image);
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

        // ===== Progress / Memory Transfer =====

        /// <summary>
        /// [新增] 开始一页：记录整页图像信息，用于计算百分比
        /// </summary>
        private static void BeginPage(TransferReadyEventArgs e)
        {
            _pageNumber++;
            _pageInProgress = true;
            _pageCompressed = false;
            _chunks.Clear();
            _compressedBytes?.Dispose();
            _compressedBytes = null;

            _transferredPixels = 0;
            _transferredBytes = 0;
            _totalPixels = 0;
            _lastReportedPercent = -1;
            _lastProgressReportMs = 0;
            _pageWidth = 0;
            _pageHeight = 0;
            _pageBitsPerPixel = 0;
            _pagePixelType = PixelType.RGB;
            _pageStopwatch = Stopwatch.StartNew();

            try
            {
                var info = e.PendingImageInfo;
                if (info != null)
                {
                    _pageWidth = info.ImageWidth;
                    _pageHeight = info.ImageLength;
                    _pageBitsPerPixel = info.BitsPerPixel;
                    _pagePixelType = info.PixelType;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠ Could not read image info: {ex.Message}");
            }

            if (_pageWidth > 0 && _pageHeight > 0)
                _totalPixels = (long)_pageWidth * _pageHeight;
            else
                _totalPixels = _estimatedTotalPixels;
        }

        /// <summary>
        /// [新增] 处理内存传输的一块数据：累加像素并更新进度
        /// </summary>
        private static void HandleMemoryChunk(DataTransferredEventArgs e)
        {
            var mem = e.MemoryInfo;
            var data = e.MemoryData;
            if (mem == null || data == null || data.Length == 0)
                return;

            // 压缩传输（例如 JPEG）：无法按像素计算，按字节累计，整页接收完后整体解码
            if (mem.Compression != CompressionType.None)
            {
                _pageCompressed = true;
                if (_compressedBytes == null)
                    _compressedBytes = new MemoryStream();
                _compressedBytes.Write(data, 0, data.Length);
                _transferredBytes += data.Length;
                ReportByteProgress();
                return;
            }

            uint columns = mem.Columns;
            uint rows = mem.Rows;
            uint bytesPerRow = mem.BytesPerRow;

            // 部分驱动不填 Columns，按每行字节数反推
            if (columns == 0 && bytesPerRow > 0 && _pageBitsPerPixel > 0)
                columns = bytesPerRow * 8 / (uint)_pageBitsPerPixel;

            if (columns == 0 || rows == 0 || bytesPerRow == 0)
            {
                Console.WriteLine($"\r⚠ Invalid memory buffer: {columns}x{rows}, bytes/row={bytesPerRow}        ");
                return;
            }

            _chunks.Add(new MemoryChunk
            {
                Data = data,
                Columns = (int)columns,
                Rows = (int)rows,
                XOffset = (int)mem.XOffset,
                YOffset = (int)mem.YOffset,
                BytesPerRow = (int)bytesPerRow
            });

            _transferredPixels += (long)columns * rows;
            ReportPixelProgress();
        }

        private static void ReportPixelProgress()
        {
            if (_totalPixels <= 0)
            {
                ReportByteProgress();
                return;
            }

            int percent = (int)Math.Min(100L, _transferredPixels * 100L / _totalPixels);
            if (percent == _lastReportedPercent)
                return;

            long elapsed = _pageStopwatch == null ? 0 : _pageStopwatch.ElapsedMilliseconds;
            // 控制输出频率，避免每块都刷新
            if (percent < 100 && elapsed - _lastProgressReportMs < 100)
                return;

            _lastReportedPercent = percent;
            _lastProgressReportMs = elapsed;
            Console.Write($"\rScanning page {_pageNumber}: {percent,3}%");
        }

        private static void ReportByteProgress()
        {
            long elapsed = _pageStopwatch == null ? 0 : _pageStopwatch.ElapsedMilliseconds;
            if (elapsed - _lastProgressReportMs < 100)
                return;

            _lastProgressReportMs = elapsed;
            Console.Write($"\rScanning page {_pageNumber}: {_transferredBytes / 1048576.0:F1} MB received");
        }

        private static void ReportPageCompletion()
        {
            long elapsed = _pageStopwatch == null ? 0 : _pageStopwatch.ElapsedMilliseconds;
            Console.WriteLine($"\rScanning page {_pageNumber}: 100% ({elapsed / 1000.0:F1}s)          ");
        }

        /// <summary>
        /// [新增] 保存当前页（内存传输模式下，数据在多个分块里）
        /// </summary>
        private static void SaveCurrentPage()
        {
            if (!_pageInProgress)
                return;
            _pageInProgress = false;

            try
            {
                if (_pageCompressed)
                {
                    byte[] compressed = _compressedBytes?.ToArray();
                    if (compressed == null || compressed.Length == 0)
                    {
                        Console.WriteLine($"\r⚠ Page {_pageNumber}: no data received                  ");
                        return;
                    }

                    using (var stream = new MemoryStream(compressed))
                    using (var image = System.Drawing.Image.FromStream(stream))
                    {
                        ReportPageCompletion();
                        SaveImage(image);
                    }
                    return;
                }

                if (_chunks.Count == 0)
                {
                    Console.WriteLine($"\r⚠ Page {_pageNumber}: no image data received              ");
                    return;
                }

                int width = Math.Max(_pageWidth, 0);
                int height = Math.Max(_pageHeight, 0);
                foreach (var chunk in _chunks)
                {
                    width = Math.Max(width, chunk.XOffset + chunk.Columns);
                    height = Math.Max(height, chunk.YOffset + chunk.Rows);
                }

                if (width <= 0 || height <= 0)
                {
                    Console.WriteLine($"\r⚠ Page {_pageNumber}: invalid image size {width}x{height}          ");
                    return;
                }

                using (var bitmap = CreatePageBitmap(width, height))
                {
                    CopyChunksToBitmap(bitmap, width, height);
                    ReportPageCompletion();
                    SaveBitmapAsJpeg(bitmap);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\r⚠ Error saving page {_pageNumber}: {ex.Message}");
            }
            finally
            {
                _pageStopwatch = null;
                _chunks.Clear();
                _compressedBytes?.Dispose();
                _compressedBytes = null;
            }
        }

        private static void DiscardCurrentPage()
        {
            _pageInProgress = false;
            _pageStopwatch = null;
            _chunks.Clear();
            _compressedBytes?.Dispose();
            _compressedBytes = null;
        }

        /// <summary>
        /// [新增] 按位深创建空白位图（TWAIN 内存传输支持的三种格式）
        /// </summary>
        private static Bitmap CreatePageBitmap(int width, int height)
        {
            if (_pageBitsPerPixel == 1)
            {
                var bitmap = new Bitmap(width, height, PixelFormat.Format1bppIndexed);
                var palette = bitmap.Palette;
                palette.Entries[0] = Color.Black;
                palette.Entries[1] = Color.White;
                bitmap.Palette = palette;
                return bitmap;
            }

            if (_pageBitsPerPixel == 8)
            {
                var bitmap = new Bitmap(width, height, PixelFormat.Format8bppIndexed);
                var palette = bitmap.Palette;
                for (int i = 0; i < 256; i++)
                    palette.Entries[i] = Color.FromArgb(i, i, i);
                bitmap.Palette = palette;
                return bitmap;
            }

            if (_pageBitsPerPixel == 24)
                return new Bitmap(width, height, PixelFormat.Format24bppRgb);

            if (_pageBitsPerPixel == 0)
            {
                // 拿不到位深时按最常见的 24 位彩色处理
                Console.WriteLine("\r⚠ Image info unavailable, assuming 24-bit RGB          ");
                return new Bitmap(width, height, PixelFormat.Format24bppRgb);
            }

            throw new NotSupportedException(
                $"Unsupported pixel format for memory transfer: {_pagePixelType}, {_pageBitsPerPixel} bpp");
        }

        /// <summary>
        /// [新增] 把各分块按 XOffset/YOffset 拷贝到整页位图
        /// </summary>
        private static void CopyChunksToBitmap(Bitmap bitmap, int width, int height)
        {
            bool isRgb = bitmap.PixelFormat == PixelFormat.Format24bppRgb;
            bool isBw = bitmap.PixelFormat == PixelFormat.Format1bppIndexed;
            int bytesPerPixel = isRgb ? 3 : 1;

            var locked = bitmap.LockBits(new Rectangle(0, 0, width, height),
                ImageLockMode.WriteOnly, bitmap.PixelFormat);
            try
            {
                foreach (var chunk in _chunks)
                {
                    int destX = isBw ? chunk.XOffset / 8 : chunk.XOffset * bytesPerPixel;
                    if (isBw && chunk.XOffset % 8 != 0)
                    {
                        Console.WriteLine($"\r⚠ 1-bit chunk with unaligned XOffset={chunk.XOffset}, image may be misaligned");
                    }

                    int rowBytes = isBw
                        ? Math.Min((chunk.Columns + 7) / 8, chunk.BytesPerRow)
                        : Math.Min(chunk.Columns * bytesPerPixel, chunk.BytesPerRow);
                    rowBytes = Math.Min(rowBytes, locked.Stride - destX);
                    if (rowBytes <= 0)
                        continue;

                    byte[] swap = isRgb ? new byte[rowBytes] : null;

                    for (int row = 0; row < chunk.Rows; row++)
                    {
                        int destY = chunk.YOffset + row;
                        if (destY >= height)
                            break;

                        int srcOffset = row * chunk.BytesPerRow;
                        if (srcOffset >= chunk.Data.Length)
                            break;

                        int copy = Math.Min(rowBytes, chunk.Data.Length - srcOffset);
                        IntPtr destPtr = IntPtr.Add(locked.Scan0, destY * locked.Stride + destX);

                        if (isRgb)
                        {
                            // TWAIN 内存传输的 RGB 是 R,G,B 顺序，GDI+ 的 Format24bppRgb 内存是 B,G,R
                            copy = copy / 3 * 3;
                            for (int i = 0; i + 2 < copy; i += 3)
                            {
                                swap[i] = chunk.Data[srcOffset + i + 2];
                                swap[i + 1] = chunk.Data[srcOffset + i + 1];
                                swap[i + 2] = chunk.Data[srcOffset + i];
                            }
                            Marshal.Copy(swap, 0, destPtr, copy);
                        }
                        else
                        {
                            Marshal.Copy(chunk.Data, srcOffset, destPtr, copy);
                        }
                    }
                }
            }
            finally
            {
                bitmap.UnlockBits(locked);
            }
        }

        private static void SaveImage(System.Drawing.Image image)
        {
            Directory.CreateDirectory(_outputPath);
            string fileName = Path.Combine(_outputPath,
                $"scan_{DateTime.Now:yyyyMMdd_HHmmss}_{_scanCount++}.jpg");
            image.Save(fileName, ImageFormat.Jpeg);
            Console.WriteLine($"Image saved to: {fileName}");
        }

        private static void SaveBitmapAsJpeg(Bitmap bitmap)
        {
            Directory.CreateDirectory(_outputPath);
            string fileName = Path.Combine(_outputPath,
                $"scan_{DateTime.Now:yyyyMMdd_HHmmss}_{_scanCount++}.jpg");

            if (bitmap.PixelFormat == PixelFormat.Format24bppRgb)
            {
                bitmap.Save(fileName, ImageFormat.Jpeg);
            }
            else
            {
                // 索引色位图先转 24 位再存 JPEG
                using (var rgb = new Bitmap(bitmap.Width, bitmap.Height, PixelFormat.Format24bppRgb))
                {
                    if (bitmap.HorizontalResolution > 0)
                        rgb.SetResolution(bitmap.HorizontalResolution, bitmap.VerticalResolution);
                    using (var g = Graphics.FromImage(rgb))
                    {
                        g.DrawImage(bitmap, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
                    }
                    rgb.Save(fileName, ImageFormat.Jpeg);
                }
            }

            Console.WriteLine($"Image saved to: {fileName}");
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
        public bool NativeTransfer { get; set; }                // [新增] 强制 Native 传输（关闭进度百分比）
    }

    /// <summary>
    /// [新增] 内存传输的一块图像数据
    /// </summary>
    internal class MemoryChunk
    {
        public byte[] Data { get; set; }
        public int Columns { get; set; }
        public int Rows { get; set; }
        public int XOffset { get; set; }
        public int YOffset { get; set; }
        public int BytesPerRow { get; set; }
    }
}