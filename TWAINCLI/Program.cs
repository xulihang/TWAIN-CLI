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

            // 解析命令行参数
            var config = ParseArguments(args);

            // 如果只列出扫描仪，执行后退出
            if (config.ListOnly)
            {
                ListScanners();
                return;
            }

            try
            {
                // 创建应用程序标识 [[31]]
                var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, Assembly.GetExecutingAssembly());

                // 创建 TWAIN 会话
                _session = new TwainSession(appId);

                // 注册事件处理 [[31]]
                _session.TransferReady += TwainSession_TransferReady;
                _session.DataTransferred += TwainSession_DataTransferred;
                _session.SourceDisabled += TwainSession_SourceDisabled;

                // 打开会话
                var openResult = _session.Open();
                if (openResult != ReturnCode.Success)
                {
                    Console.WriteLine($"打开 TWAIN 会话失败: {openResult}");
                    return;
                }

                Console.WriteLine("TWAIN 会话已打开");

                // 查找并选择扫描仪
                DataSource selectedSource = null;
                var sources = _session.GetSources().ToList();

                if (sources.Count == 0)
                {
                    Console.WriteLine("未找到任何扫描仪");
                    _session.Close();
                    return;
                }

                // 按名称选择扫描仪（如果指定了 -d 参数）
                if (!string.IsNullOrEmpty(config.ScannerName))
                {
                    selectedSource = sources.FirstOrDefault(s =>
                        s.Name.IndexOf(config.ScannerName, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (selectedSource == null)
                    {
                        Console.WriteLine($"未找到名称包含 '{config.ScannerName}' 的扫描仪");
                        _session.Close();
                        return;
                    }
                }
                else
                {
                    // 默认选择第一个可用扫描仪
                    selectedSource = sources.FirstOrDefault();
                }

                if (selectedSource != null)
                {
                    Console.WriteLine($"\n选择扫描仪: {selectedSource.Name}");

                    // 打开数据源 [[21]]
                    var openDsResult = selectedSource.Open();
                    if (openDsResult == ReturnCode.Success)
                    {
                        Console.WriteLine("扫描仪已打开，正在配置参数...");

                        // 应用配置参数
                        ConfigureSource(selectedSource, config);

                        // 启用数据源进行扫描（无 UI 模式）[[21]]
                        var enableResult = selectedSource.Enable(SourceEnableMode.NoUI, false, IntPtr.Zero);
                        if (enableResult != ReturnCode.Success)
                        {
                            Console.WriteLine($"启用扫描仪失败: {enableResult}");
                            selectedSource.Close();
                            return;
                        }

                        // 运行消息循环，等待扫描完成
                        Application.Run();
                    }
                    else
                    {
                        Console.WriteLine($"打开扫描仪失败: {openDsResult}");
                        selectedSource.Close();
                    }
                }
                else
                {
                    Console.WriteLine("未找到合适的扫描仪");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                if (_session != null)
                {
                    _session.Close();
                    Console.WriteLine("TWAIN 会话已关闭");
                }
            }
        }

        /// <summary>
        /// 配置扫描仪的 TWAIN Capabilities [[12]][[13]]
        /// </summary>
        private static void ConfigureSource(DataSource source, ScanConfig config)
        {
            // 1. 配置进纸器/平板 (feeder/flatbed)
            if (source.Capabilities.CapFeederEnabled.IsSupported)
            {
                bool useFeeder = config.SourceType?.ToLower() == "feeder";
                source.Capabilities.CapFeederEnabled.SetValue(useFeeder ? BoolType.True : BoolType.False);
            }

            // 2. 配置光路类型 (positive/negative/flatbed) [[19]]
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

            // 3. 配置颜色模式 (bw/gray/color) [[19]]
            if (source.Capabilities.ICapPixelType.IsSupported && !string.IsNullOrEmpty(config.ColorMode))
            {
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
                // 其他情况保持 null
                if (pixelType.HasValue && source.Capabilities.ICapPixelType.CanSet &&
                    source.Capabilities.ICapPixelType.GetValues().Contains(pixelType.Value))
                {
                    source.Capabilities.ICapPixelType.SetValue(pixelType.Value);
                }
            }

            // 4. 配置分辨率 [[13]][[14]]
            if (config.Resolution > 0)
            {
                if (source.Capabilities.ICapXResolution.IsSupported && source.Capabilities.ICapXResolution.CanSet)
                    source.Capabilities.ICapXResolution.SetValue(config.Resolution);
                if (source.Capabilities.ICapYResolution.IsSupported && source.Capabilities.ICapYResolution.CanSet)
                    source.Capabilities.ICapYResolution.SetValue(config.Resolution);
            }

            // 5. 配置双面扫描 [[12]]

            if (config.Duplex && source.Capabilities.CapDuplexEnabled.IsSupported && source.Capabilities.CapDuplexEnabled.CanSet)
            {
                source.Capabilities.CapDuplexEnabled.SetValue(BoolType.True);
            }

            // 6. 配置扫描区域 (Frame) [[1]][[4]]
            // TWAIN 的 Frame 坐标单位由 ICAP_UNITS 决定，默认为 Inches
            if (config.HasArea && source.Capabilities.ICapFrames.IsSupported && source.Capabilities.ICapFrames.CanSet)
            {
                if (source.Capabilities.ICapUnits.IsSupported && source.Capabilities.ICapUnits.CanSet) {
                    source.Capabilities.ICapUnits.SetValue(Unit.Millimeters);
                }
                var frame = new TWFrame
                {
                    Left = config.PageLeft,
                    Top = config.PageTop,
                    Right = config.PageLeft + config.PageWidth,
                    Bottom = config.PageTop + config.PageHeight
                };
                source.Capabilities.ICapFrames.SetValue(frame);
            }

            // 7. 开启灯源（部分扫描仪需要）
            if (source.Capabilities.ICapLampState.IsSupported && source.Capabilities.ICapLampState.CanSet)
            {
                source.Capabilities.ICapLampState.SetValue(BoolType.True);
            }
        }

        /// <summary>
        /// 列出所有可用扫描仪
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
                    Console.WriteLine($"打开 TWAIN 会话失败: {openResult}");
                    return;
                }

                var sources = session.GetSources().ToList();

                if (sources.Count == 0)
                {
                    Console.WriteLine("未找到任何扫描仪");
                }
                else
                {
                    Console.WriteLine("可用扫描仪列表:");
                    foreach (var source in sources)
                    {
                        Console.WriteLine($"  - {source.Name} (ID: {source.Id})");
                    }
                }

                session.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"列出扫描仪时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 解析命令行参数
        /// </summary>
        private static ScanConfig ParseArguments(string[] args)
        {
            var config = new ScanConfig
            {
                OutputPath = @"C:\Users\hp\",
                Resolution = 200,      // 默认分辨率 200 DPI
                ColorMode = "color",   // 默认彩色
                SourceType = "flatbed", // 默认平板
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

                    case "-t": // Page Top (原请求用-y，但与height冲突，改用-t)
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
                            // 确保路径以分隔符结尾
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

                    case "-L": // List scanners (大写)
                        config.ListOnly = true;
                        break;

                    case "-?":
                    case "--help":
                        PrintHelp();
                        Environment.Exit(0);
                        break;
                }
            }

            // 判断是否设置了扫描区域
            config.HasArea = config.PageWidth > 0 && config.PageHeight > 0;
            if (config.ScannerName == "") {
                PrintHelp();
                Environment.Exit(0);
            }
            return config;
        }

        /// <summary>
        /// 打印帮助信息
        /// </summary>
        private static void PrintHelp()
        {
            Console.WriteLine("TWAIN CLI Scanner - 使用说明:");
            Console.WriteLine("  -L                列出所有可用扫描仪");
            Console.WriteLine("  -s <name>         指定源类型: feeder(进纸器), positive(正片), negative(负片), flatbed(平板)");
            Console.WriteLine("  -d <name>         按名称指定扫描仪（支持部分匹配）");
            Console.WriteLine("  -m <mode>         指定颜色模式: bw(黑白), gray(灰度), color(彩色)");
            Console.WriteLine("  -r <resolution>   指定分辨率 DPI (默认: 200)");
            Console.WriteLine("  -o <path>         输出文件/文件夹路径");
            Console.WriteLine("  -l <left>         扫描区域左边距 (英寸)");
            Console.WriteLine("  -t <top>          扫描区域上边距 (英寸)");
            Console.WriteLine("  -x <width>        扫描区域宽度 (英寸)");
            Console.WriteLine("  -y <height>       扫描区域高度 (英寸)");
            Console.WriteLine("  --duplex          启用双面扫描");
            Console.WriteLine("  -?, --help        显示此帮助信息");
            Console.WriteLine();
            Console.WriteLine("示例:");
            Console.WriteLine("  TWAINCLI.exe -L");
            Console.WriteLine("  TWAINCLI.exe -d \"EPSON\" -s feeder -m color -r 300 -o C:\\Scans\\");
            Console.WriteLine("  TWAINCLI.exe -s negative -l 0 -t 0 -x 8.5 -y 11 --duplex");
        }

        // ===== 事件处理程序 =====

        private static void TwainSession_SourceDisabled(object sender, EventArgs e)
        {
            Console.WriteLine("扫描完成或用户取消了扫描");
            // 退出消息循环
            Application.Exit();
        }

        private static void TwainSession_TransferReady(object sender, TransferReadyEventArgs e)
        {
            Console.WriteLine("准备传输图像...");
            e.CancelAll = false;  // 确保不取消传输
        }

        private static void TwainSession_DataTransferred(object sender, DataTransferredEventArgs e)
        {
            try
            {
                Console.WriteLine($"接收到数据，传输类型: {e.TransferType}");

                if (e.TransferType == XferMech.Native)
                {
                    // 处理 Native 传输类型（大多数扫描仪使用）[[31]]
                    using (var stream = e.GetNativeImageStream())
                    {
                        if (stream != null)
                        {
                            using (var image = System.Drawing.Image.FromStream(stream))
                            {
                                // 确保输出目录存在
                                Directory.CreateDirectory(_outputPath);

                                // 生成带时间戳的文件名
                                string fileName = Path.Combine(_outputPath,
                                    $"scan_{DateTime.Now:yyyyMMdd_HHmmss}_{_scanCount++}.jpg");
                                image.Save(fileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                                Console.WriteLine($"图像已保存到: {fileName}");
                            }
                        }
                    }
                }
                else if (e.TransferType == XferMech.File)
                {
                    // 处理文件传输类型
                    Console.WriteLine($"文件传输模式，文件路径: {e.FileDataPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理传输数据时出错: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 扫描配置类
    /// </summary>
    internal class ScanConfig
    {
        public bool ListOnly { get; set; }              // -L: 仅列出扫描仪
        public string ScannerName { get; set; }         // -d: 扫描仪名称
        public string SourceType { get; set; } = "flatbed";  // -s: 源类型
        public string ColorMode { get; set; } = "color";     // -m: 颜色模式
        public int Resolution { get; set; } = 200;      // -r: 分辨率
        public string OutputPath { get; set; }          // -o: 输出路径
        public float PageLeft { get; set; } = 0;        // -l: 左边距
        public float PageTop { get; set; } = 0;         // -t: 上边距
        public float PageWidth { get; set; } = 0;       // -x: 宽度
        public float PageHeight { get; set; } = 0;      // -y: 高度
        public bool HasArea { get; set; }               // 是否设置了扫描区域
        public bool Duplex { get; set; }                // --duplex: 双面扫描
    }
}