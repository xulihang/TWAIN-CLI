using NTwain;
using NTwain.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
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

        // [新增] ShowUI 模式：隐藏父窗口 + 强制激活扫描仪 UI
        private static Form _parentWindow = null;               // 传给驱动的宿主窗口句柄
        private static Timer _uiActivationTimer = null;         // 轮询并激活扫描仪 UI 的定时器
        private static int _uiActivationTicks = 0;              // 已轮询次数
        private static readonly HashSet<IntPtr> _knownWindows = new HashSet<IntPtr>(); // Enable 前已存在的窗口

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

        // [新增] 实时预览（周期快照）
        private static string _previewPath = null;
        private static int _previewIntervalMs = 500;
        private static int _previewMaxWidth = 800;
        private static long _lastPreviewMs = 0;
        private static bool _previewErrorShown = false;
        private static Bitmap _previewBitmap = null;              // 增量绘制的预览图
        private static double _previewScale = 1.0;
        private static int _previewPageWidth = 0;
        private static int _previewPageHeight = 0;
        private static bool _unalignedWarningShown = false;

        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            NTwain.PlatformInfo.Current.PreferNewDSM = false;

            // Parse command line arguments
            var config = ParseArguments(args);
            _outputPath = config.OutputPath;
            _previewPath = config.PreviewPath;
            _previewIntervalMs = config.PreviewIntervalMs;
            _previewMaxWidth = config.PreviewMaxWidth;

            // If only listing scanners, execute and exit
            if (config.ListOnly)
            {
                ListScanners();
                return;
            }

            // [新增] NTwain 的 WindowsFormsMessageLoopHook 要求消息循环已经在 UI 线程运行。
            // 先启动 WinForms 消息循环，再在循环内打开 TWAIN 会话、显示驱动 UI，
            // 这样驱动 UI 由正常的 WinForms 消息泵处理（而不是 NTwain 内部的 WPF 消息线程）。
            EnsureParentWindow();
            _parentWindow.BeginInvoke(new Action(() => BeginScan(config)));
            try
            {
                Application.Run();
            }
            finally
            {
                CleanupSession();
            }
        }

        /// <summary>
        /// [新增] 在消息循环内打开 TWAIN 会话并启动扫描
        /// </summary>
        private static void BeginScan(ScanConfig config)
        {
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
                // [新增] 使用 WinForms 消息钩子，并把宿主窗口句柄同时作为 DSM 窗口，
                // 与 NAPS2 的做法一致：所有 TWAIN 调用都在 WinForms 消息泵线程上执行。
                var openResult = _session.Open(new WindowsFormsMessageLoopHook(_parentWindow.Handle));
                if (openResult != ReturnCode.Success)
                {
                    Console.WriteLine($"Failed to open TWAIN session: {openResult}");
                    Application.Exit();
                    return;
                }

                Console.WriteLine("TWAIN session opened");

                // Find and select scanner
                DataSource selectedSource = null;
                var sources = _session.GetSources().ToList();

                if (sources.Count == 0)
                {
                    Console.WriteLine("No scanners found");
                    Application.Exit();
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
                        Application.Exit();
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

                            // [新增] 记录 Enable 前已存在的可见窗口，用于找出驱动刚创建的 UI
                            SnapshotTopLevelWindows();

                            var enableResult = selectedSource.Enable(SourceEnableMode.ShowUI, false, _parentWindow.Handle);
                            if (enableResult != ReturnCode.Success)
                            {
                                Console.WriteLine($"Failed to enable scanner: {enableResult}");
                                selectedSource.Close();
                                Application.Exit();
                                return;
                            }

                            // [新增] 部分驱动的 UI 显示后不会自行激活，先把焦点切入一次
                            StartScannerUIActivation(selectedSource.Name);
                        }
                        else
                        {
                            Console.WriteLine("Scanner opened, configuring parameters...");

                            // Apply configuration parameters
                            ConfigureSource(selectedSource, config);

                            // Enable data source for scanning (no UI mode)
                            var enableResult = selectedSource.Enable(SourceEnableMode.NoUI, false, _parentWindow.Handle);
                            if (enableResult != ReturnCode.Success)
                            {
                                Console.WriteLine($"Failed to enable scanner: {enableResult}");
                                selectedSource.Close();
                                Application.Exit();
                                return;
                            }
                        }

                        // [新增] 启动取消文件监控
                        StartCancelMonitoring(config.CancelFileName);

                        // [新增] 已经在消息循环内，不能再调用 Application.Run()；
                        // 扫描完成/取消时由事件处理里的 Application.Exit() 结束循环。
                    }
                    else
                    {
                        Console.WriteLine($"Failed to open scanner: {openDsResult}");
                        selectedSource.Close();
                        Application.Exit();
                    }
                }
                else
                {
                    Console.WriteLine("No suitable scanner found");
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                Application.Exit();
            }
        }

        /// <summary>
        /// [新增] 消息循环结束后清理 TWAIN 会话和宿主窗口
        /// </summary>
        private static void CleanupSession()
        {
            // [新增] 停止取消文件监控
            StopCancelMonitoring();

            // [新增] 停止扫描仪 UI 激活轮询
            StopScannerUIActivation();

            if (_session != null)
            {
                try
                {
                    if (_session.IsDsmOpen)
                        _session.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error closing TWAIN session: {ex.Message}");
                }
                Console.WriteLine("TWAIN session closed");
            }

            // [新增] 释放 ShowUI 宿主窗口
            if (_parentWindow != null)
            {
                try { _parentWindow.Dispose(); } catch { /* 忽略释放异常 */ }
                _parentWindow = null;
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
        /// [新增] 创建一个不可见的父窗口，作为驱动 UI 的宿主窗口。
        /// 部分驱动（如 Canon）在 hParent 为 IntPtr.Zero 时不会把 UI 激活到前台。
        /// </summary>
        private static void EnsureParentWindow()
        {
            if (_parentWindow != null && !_parentWindow.IsDisposed)
                return;

            var area = Screen.PrimaryScreen != null
                ? Screen.PrimaryScreen.WorkingArea
                : new Rectangle(0, 0, 1, 1);
            _parentWindow = new Form
            {
                ShowInTaskbar = false,
                FormBorderStyle = FormBorderStyle.None,
                StartPosition = FormStartPosition.Manual,
                Location = new Point(area.Left + area.Width / 2, area.Top + area.Height / 2),
                Size = new Size(1, 1),
                Text = "TWAINCLI"
            };

            // 访问 Handle 强制创建窗口句柄；不调用 Show()，窗口保持不可见
            IntPtr handle = _parentWindow.Handle;
            GC.KeepAlive(handle);
        }

        /// <summary>
        /// [新增] 记录当前所有可见的顶层窗口（用于识别驱动随后创建的扫描仪 UI）
        /// </summary>
        private static void SnapshotTopLevelWindows()
        {
            _knownWindows.Clear();
            EnumWindows((hWnd, lParam) =>
            {
                if (IsWindowVisible(hWnd))
                    _knownWindows.Add(hWnd);
                return true;
            }, IntPtr.Zero);
        }

        /// <summary>
        /// [新增] 轮询等待扫描仪 UI 出现，并把它切换为前台窗口。
        /// 相当于手动“切换一次焦点”，让 Canon 等驱动初始化界面。
        /// </summary>
        private static void StartScannerUIActivation(string sourceName)
        {
            StopScannerUIActivation();

            _uiActivationTicks = 0;
            _uiActivationTimer = new Timer { Interval = 300 };
            _uiActivationTimer.Tick += (sender, e) =>
            {
                _uiActivationTicks++;

                IntPtr scannerWindow = FindScannerWindow(sourceName);
                if (scannerWindow != IntPtr.Zero && TryActivateWindow(scannerWindow))
                {
                    Console.WriteLine($"Scanner UI activated (hwnd=0x{scannerWindow.ToInt64():X})");
                    StopScannerUIActivation();
                    return;
                }

                // 最多等待 15 秒，避免一直抢占焦点或空转
                if (_uiActivationTicks >= 50)
                    StopScannerUIActivation();
            };
            _uiActivationTimer.Start();
        }

        private static void StopScannerUIActivation()
        {
            if (_uiActivationTimer != null)
            {
                _uiActivationTimer.Stop();
                _uiActivationTimer.Dispose();
                _uiActivationTimer = null;
            }
        }

        /// <summary>
        /// [新增] 在 Enable 之后新出现（或刚变为可见）的顶层窗口中查找扫描仪 UI
        /// </summary>
        private static IntPtr FindScannerWindow(string sourceName)
        {
            IntPtr result = IntPtr.Zero;
            uint currentPid = GetCurrentProcessId();

            EnumWindows((hWnd, lParam) =>
            {
                // 只关注 Enable 之后新出现的可见窗口
                if (_knownWindows.Contains(hWnd) || !IsWindowVisible(hWnd))
                    return true;

                uint windowPid;
                GetWindowThreadProcessId(hWnd, out windowPid);

                var titleBuilder = new StringBuilder(512);
                GetWindowText(hWnd, titleBuilder, titleBuilder.Capacity);
                string title = titleBuilder.ToString();

                var classBuilder = new StringBuilder(256);
                GetClassName(hWnd, classBuilder, classBuilder.Capacity);
                string className = classBuilder.ToString();

                // 经典 TWAIN DS 跑在应用进程内，因此同进程的新窗口优先；
                // 也兼容独立进程承载 UI 的驱动（按标题/类名识别）
                bool sameProcess = windowPid == currentPid;
                bool looksLikeTwain = className.IndexOf("TWNDS", StringComparison.OrdinalIgnoreCase) >= 0;
                bool titleMatches = title.Length > 0 && (
                    (!string.IsNullOrEmpty(sourceName) && title.IndexOf(sourceName, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    title.IndexOf("ScanGear", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    title.IndexOf("Canon", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    title.IndexOf("scan", StringComparison.OrdinalIgnoreCase) >= 0);

                if (!sameProcess && !looksLikeTwain && !titleMatches)
                    return true;

                // 跳过提示框、隐藏的辅助窗口等小窗口
                RECT rect;
                if (!GetWindowRect(hWnd, out rect))
                    return true;
                if (rect.Right - rect.Left < 150 || rect.Bottom - rect.Top < 100)
                    return true;

                result = hWnd;
                return false; // 找到第一个候选窗口即停止枚举
            }, IntPtr.Zero);

            return result;
        }

        /// <summary>
        /// [新增] 把窗口切换为前台窗口（等价于用户手动切换一次焦点）
        /// </summary>
        private static bool TryActivateWindow(IntPtr hWnd)
        {
            if (GetForegroundWindow() == hWnd)
                return true;

            // 先直接尝试
            ShowWindow(hWnd, SW_RESTORE);
            BringWindowToTop(hWnd);
            SetForegroundWindow(hWnd);

            if (GetForegroundWindow() == hWnd)
                return true;

            // 系统通常只允许前台进程调用 SetForegroundWindow。
            // 临时把输入队列附加到当前前台线程，可绕过该限制（经典做法）。
            uint currentThreadId = GetCurrentThreadId();
            IntPtr foregroundWindow = GetForegroundWindow();
            uint foregroundPid;
            uint foregroundThreadId = GetWindowThreadProcessId(foregroundWindow, out foregroundPid);
            uint targetPid;
            uint targetThreadId = GetWindowThreadProcessId(hWnd, out targetPid);

            bool attachedToForeground = false;
            bool attachedToTarget = false;
            try
            {
                // 前台窗口无响应时不要 attach，否则可能把本进程一起卡死
                if (foregroundThreadId != 0 && foregroundThreadId != currentThreadId &&
                    !IsHungAppWindow(foregroundWindow))
                {
                    attachedToForeground = AttachThreadInput(foregroundThreadId, currentThreadId, true);
                }

                if (targetThreadId != 0 && targetThreadId != currentThreadId &&
                    targetThreadId != foregroundThreadId)
                {
                    attachedToTarget = AttachThreadInput(targetThreadId, currentThreadId, true);
                }

                BringWindowToTop(hWnd);
                ShowWindow(hWnd, SW_SHOW);
                SetForegroundWindow(hWnd);
                SetActiveWindow(hWnd);
                SetFocus(hWnd);
            }
            finally
            {
                if (attachedToTarget)
                    AttachThreadInput(targetThreadId, currentThreadId, false);
                if (attachedToForeground)
                    AttachThreadInput(foregroundThreadId, currentThreadId, false);
            }

            return GetForegroundWindow() == hWnd;
        }

        // ===== Win32 interop：扫描仪 UI 激活 =====
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsHungAppWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern IntPtr SetActiveWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentProcessId();

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private const int SW_RESTORE = 9;
        private const int SW_SHOW = 5;

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
                WarnPreviewRequiresMemory(config);
                return;
            }

            // Kyocera Ecosys 的内存传输有已知问题（与 NAPS2 相同的处理）
            if (ContainsIgnoreCase(source.Manufacturer, "Kyocera") &&
                (ContainsIgnoreCase(source.ProductFamily, "Ecosys") || ContainsIgnoreCase(source.Name, "Ecosys")))
            {
                Console.WriteLine("Transfer mode: Native (Kyocera Ecosys, memory transfer known to be broken)");
                WarnPreviewRequiresMemory(config);
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
            WarnPreviewRequiresMemory(config);
        }

        private static void WarnPreviewRequiresMemory(ScanConfig config)
        {
            if (!string.IsNullOrEmpty(config.PreviewPath))
                Console.WriteLine("⚠ Live preview requires memory transfer; preview will be written when the scan completes");
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

                    case "--preview":  // [新增] 实时预览输出路径
                        if (i + 1 < args.Length)
                            config.PreviewPath = args[++i];
                        break;

                    case "--previewInterval":  // [新增] 预览刷新间隔（毫秒）
                        if (i + 1 < args.Length && int.TryParse(args[++i], out int previewInterval) && previewInterval >= 100)
                            config.PreviewIntervalMs = previewInterval;
                        break;

                    case "--previewWidth":  // [新增] 预览最大宽度（像素，0 = 原始尺寸）
                        if (i + 1 < args.Length && int.TryParse(args[++i], out int previewWidth) && previewWidth >= 0)
                            config.PreviewMaxWidth = previewWidth;
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
            Console.WriteLine("  --preview <path>      Live preview JPEG updated while scanning (memory transfer only)");
            Console.WriteLine("  --previewInterval <ms> Preview refresh interval (default: 500)");
            Console.WriteLine("  --previewWidth <px>   Max preview width (default: 800, 0 = full size)");
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
            _lastPreviewMs = 0;
            ResetPreviewBitmap();
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

            if (_pageBitsPerPixel <= 0)
                Console.WriteLine("⚠ Image info unavailable, assuming 24-bit RGB");
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

            var chunk = new MemoryChunk
            {
                Data = data,
                Columns = (int)columns,
                Rows = (int)rows,
                XOffset = (int)mem.XOffset,
                YOffset = (int)mem.YOffset,
                BytesPerRow = (int)bytesPerRow
            };
            _chunks.Add(chunk);

            _transferredPixels += (long)columns * rows;
            ReportPixelProgress();
            RenderChunkToPreview(chunk);
            MaybeWritePreview();
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

                using (var bitmap = BuildPageBitmap(out int contentWidth, out int contentHeight))
                {
                    if (bitmap == null)
                        return;

                    ReportPageCompletion();
                    SaveBitmapAsJpeg(bitmap);
                    UpdatePreview(bitmap, contentWidth, contentHeight);
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
                ResetPreviewBitmap();
            }
        }

        private static void DiscardCurrentPage()
        {
            _pageInProgress = false;
            _pageStopwatch = null;
            _chunks.Clear();
            _compressedBytes?.Dispose();
            _compressedBytes = null;
            ResetPreviewBitmap();
        }

        /// <summary>
        /// [新增] 用当前已接收的分块构建整页位图
        /// </summary>
        private static Bitmap BuildPageBitmap(out int contentWidth, out int contentHeight)
        {
            contentWidth = 0;
            contentHeight = 0;

            int width = Math.Max(_pageWidth, 0);
            int height = Math.Max(_pageHeight, 0);
            foreach (var chunk in _chunks)
            {
                contentWidth = Math.Max(contentWidth, chunk.XOffset + chunk.Columns);
                contentHeight = Math.Max(contentHeight, chunk.YOffset + chunk.Rows);
            }
            width = Math.Max(width, contentWidth);
            height = Math.Max(height, contentHeight);

            if (width <= 0 || height <= 0)
            {
                Console.WriteLine($"\r⚠ Page {_pageNumber}: invalid image size {width}x{height}          ");
                return null;
            }

            var bitmap = CreatePageBitmap(width, height);
            try
            {
                CopyChunksToBitmap(bitmap, width, height);
                return bitmap;
            }
            catch
            {
                bitmap.Dispose();
                throw;
            }
        }

        /// <summary>
        /// [新增] 按时间间隔把当前预览图写成一张完整的 JPEG，供第三方程序轮询显示
        /// </summary>
        private static void MaybeWritePreview()
        {
            if (string.IsNullOrEmpty(_previewPath))
                return;

            long now = _pageStopwatch == null ? 0 : _pageStopwatch.ElapsedMilliseconds;
            if (now - _lastPreviewMs < _previewIntervalMs)
                return;
            _lastPreviewMs = now;

            try
            {
                if (_previewBitmap != null)
                {
                    SavePreviewBitmap();
                }
                else if (_chunks.Count > 0)
                {
                    // 拿不到整页尺寸时退化为按分块重建
                    using (var bitmap = BuildPageBitmap(out int contentWidth, out int contentHeight))
                    {
                        if (bitmap != null)
                            UpdatePreview(bitmap, contentWidth, contentHeight);
                    }
                }
            }
            catch (Exception ex)
            {
                ShowPreviewError(ex);
            }
        }

        private static void SavePreviewBitmap()
        {
            if (_previewBitmap == null)
                return;

            string fullPath = Path.GetFullPath(_previewPath);
            string dir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);
            string tempPath = fullPath + ".tmp";

            _previewBitmap.Save(tempPath, ImageFormat.Jpeg);
            ReplaceFile(tempPath, fullPath);
        }

        /// <summary>
        /// [新增] 把新到达的一块数据直接画进预览图（增量绘制，避免每次重建整页）
        /// </summary>
        private static void RenderChunkToPreview(MemoryChunk chunk)
        {
            if (string.IsNullOrEmpty(_previewPath))
                return;

            try
            {
                if (!EnsurePreviewBitmap())
                    return;

                int contentWidth = chunk.XOffset + chunk.Columns;
                int contentHeight = chunk.YOffset + chunk.Rows;
                if (contentWidth > _previewPageWidth || contentHeight > _previewPageHeight)
                {
                    // 实际内容超出驱动声明的尺寸：按新尺寸重建预览并重绘已有分块
                    RebuildPreviewBitmap(
                        Math.Max(_previewPageWidth, contentWidth),
                        Math.Max(_previewPageHeight, contentHeight));
                    return;
                }

                DrawChunkScaled(chunk);
            }
            catch (Exception ex)
            {
                ShowPreviewError(ex);
            }
        }

        private static bool EnsurePreviewBitmap()
        {
            if (_previewBitmap != null)
                return true;

            int pageWidth = Math.Max(_pageWidth, 0);
            int pageHeight = Math.Max(_pageHeight, 0);
            if (pageWidth <= 0 || pageHeight <= 0)
                return false;

            AllocatePreviewBitmap(pageWidth, pageHeight);
            return true;
        }

        private static void AllocatePreviewBitmap(int pageWidth, int pageHeight)
        {
            _previewBitmap?.Dispose();
            _previewPageWidth = pageWidth;
            _previewPageHeight = pageHeight;

            int width = pageWidth;
            int height = pageHeight;
            if (_previewMaxWidth > 0 && width > _previewMaxWidth)
            {
                _previewScale = (double)_previewMaxWidth / width;
                width = _previewMaxWidth;
                height = Math.Max(1, (int)Math.Round(pageHeight * _previewScale));
            }
            else
            {
                _previewScale = 1.0;
            }

            _previewBitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(_previewBitmap))
            {
                g.Clear(Color.White);
            }
        }

        private static void RebuildPreviewBitmap(int pageWidth, int pageHeight)
        {
            AllocatePreviewBitmap(pageWidth, pageHeight);
            foreach (var chunk in _chunks)
                DrawChunkScaled(chunk);
        }

        private static void DrawChunkScaled(MemoryChunk chunk)
        {
            if (_previewBitmap == null)
                return;

            int destX = (int)Math.Floor(chunk.XOffset * _previewScale);
            int destY = (int)Math.Floor(chunk.YOffset * _previewScale);
            int destRight = (int)Math.Ceiling((chunk.XOffset + chunk.Columns) * _previewScale);
            int destBottom = (int)Math.Ceiling((chunk.YOffset + chunk.Rows) * _previewScale);

            if (destX < 0 || destY < 0 || destX >= _previewBitmap.Width || destY >= _previewBitmap.Height)
                return;

            int destWidth = Math.Min(_previewBitmap.Width - destX, Math.Max(1, destRight - destX));
            int destHeight = Math.Min(_previewBitmap.Height - destY, Math.Max(1, destBottom - destY));
            if (destWidth <= 0 || destHeight <= 0)
                return;

            using (var chunkBitmap = CreateChunkBitmap(chunk))
            using (var g = Graphics.FromImage(_previewBitmap))
            {
                g.InterpolationMode = InterpolationMode.Bilinear;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(chunkBitmap,
                    new Rectangle(destX, destY, destWidth, destHeight),
                    new Rectangle(0, 0, chunk.Columns, chunk.Rows),
                    GraphicsUnit.Pixel);
            }
        }

        private static Bitmap CreateChunkBitmap(MemoryChunk chunk)
        {
            var bitmap = CreatePageBitmap(chunk.Columns, chunk.Rows);
            bool isRgb = bitmap.PixelFormat == PixelFormat.Format24bppRgb;
            bool isBw = bitmap.PixelFormat == PixelFormat.Format1bppIndexed;

            var locked = bitmap.LockBits(new Rectangle(0, 0, chunk.Columns, chunk.Rows),
                ImageLockMode.WriteOnly, bitmap.PixelFormat);
            try
            {
                CopyChunkRows(locked, isRgb, isBw, chunk, 0, 0);
            }
            finally
            {
                bitmap.UnlockBits(locked);
            }
            return bitmap;
        }

        private static void ResetPreviewBitmap()
        {
            _previewBitmap?.Dispose();
            _previewBitmap = null;
            _previewScale = 1.0;
            _previewPageWidth = 0;
            _previewPageHeight = 0;
        }

        /// <summary>
        /// [新增] 将已接收部分缩放后原子写入预览文件（始终是一张有效的完整 JPEG）
        /// </summary>
        private static void UpdatePreview(System.Drawing.Image source, int sourceWidth, int sourceHeight)
        {
            if (string.IsNullOrEmpty(_previewPath) || source == null)
                return;

            try
            {
                if (sourceWidth <= 0 || sourceHeight <= 0)
                {
                    sourceWidth = source.Width;
                    sourceHeight = source.Height;
                }
                sourceWidth = Math.Min(sourceWidth, source.Width);
                sourceHeight = Math.Min(sourceHeight, source.Height);

                int width = sourceWidth;
                int height = sourceHeight;
                if (_previewMaxWidth > 0 && width > _previewMaxWidth)
                {
                    double scale = (double)_previewMaxWidth / width;
                    width = Math.Max(1, (int)Math.Round(width * scale));
                    height = Math.Max(1, (int)Math.Round(height * scale));
                }

                string fullPath = Path.GetFullPath(_previewPath);
                string dir = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(dir))
                    Directory.CreateDirectory(dir);
                string tempPath = fullPath + ".tmp";

                using (var preview = new Bitmap(width, height, PixelFormat.Format24bppRgb))
                {
                    using (var g = Graphics.FromImage(preview))
                    {
                        g.Clear(Color.White);
                        g.InterpolationMode = InterpolationMode.Bilinear;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        g.DrawImage(source,
                            new Rectangle(0, 0, width, height),
                            new Rectangle(0, 0, sourceWidth, sourceHeight),
                            GraphicsUnit.Pixel);
                    }
                    preview.Save(tempPath, ImageFormat.Jpeg);
                }

                ReplaceFile(tempPath, fullPath);
            }
            catch (Exception ex)
            {
                ShowPreviewError(ex);
            }
        }

        private static void ReplaceFile(string tempPath, string targetPath)
        {
            try
            {
                if (File.Exists(targetPath))
                    File.Replace(tempPath, targetPath, null);
                else
                    File.Move(tempPath, targetPath);
            }
            catch (IOException)
            {
                // 目标文件可能正被第三方程序占用，保留临时文件，下次预览时覆盖
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        private static void ShowPreviewError(Exception ex)
        {
            if (_previewErrorShown)
                return;
            _previewErrorShown = true;
            Console.WriteLine($"\r⚠ Preview update failed: {ex.Message}          ");
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
                    CopyChunkRows(locked, isRgb, isBw, chunk, destX, chunk.YOffset);
                }
            }
            finally
            {
                bitmap.UnlockBits(locked);
            }
        }

        /// <summary>
        /// [新增] 将一块原始像素数据拷贝到位图（RGB 会做 R,B 交换）
        /// </summary>
        private static void CopyChunkRows(BitmapData locked, bool isRgb, bool isBw,
            MemoryChunk chunk, int destX, int destY)
        {
            int bytesPerPixel = isRgb ? 3 : 1;

            if (isBw && chunk.XOffset % 8 != 0 && !_unalignedWarningShown)
            {
                _unalignedWarningShown = true;
                Console.WriteLine($"\r⚠ 1-bit chunk with unaligned XOffset={chunk.XOffset}, image may be misaligned");
            }

            int rowBytes = isBw
                ? Math.Min((chunk.Columns + 7) / 8, chunk.BytesPerRow)
                : Math.Min(chunk.Columns * bytesPerPixel, chunk.BytesPerRow);
            rowBytes = Math.Min(rowBytes, locked.Stride - destX);
            if (rowBytes <= 0)
                return;

            byte[] swap = isRgb ? new byte[rowBytes] : null;

            for (int row = 0; row < chunk.Rows; row++)
            {
                int y = destY + row;
                if (y < 0 || y >= locked.Height)
                    break;

                int srcOffset = row * chunk.BytesPerRow;
                if (srcOffset >= chunk.Data.Length)
                    break;

                int copy = Math.Min(rowBytes, chunk.Data.Length - srcOffset);
                IntPtr destPtr = IntPtr.Add(locked.Scan0, y * locked.Stride + destX);

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

        private static void SaveImage(System.Drawing.Image image)
        {
            Directory.CreateDirectory(_outputPath);
            string fileName = Path.Combine(_outputPath,
                $"scan_{DateTime.Now:yyyyMMdd_HHmmss}_{_scanCount++}.jpg");
            image.Save(fileName, ImageFormat.Jpeg);
            Console.WriteLine($"Image saved to: {fileName}");
            UpdatePreview(image, image.Width, image.Height);
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
        public string PreviewPath { get; set; }                 // [新增] 实时预览 JPEG 路径
        public int PreviewIntervalMs { get; set; } = 500;       // [新增] 预览刷新间隔（毫秒）
        public int PreviewMaxWidth { get; set; } = 800;         // [新增] 预览最大宽度（0 = 原始尺寸）
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