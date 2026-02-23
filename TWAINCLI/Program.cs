using NTwain;
using NTwain.Data;
using System;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace TWAINCLI
{
    internal class Program
    {
        private static TwainSession _session;
        private static bool _scanCompleted = false;

        [STAThread]  // 需要STAThread属性
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            NTwain.PlatformInfo.Current.PreferNewDSM = false;
            try
            {
                // 创建应用程序标识
                var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, Assembly.GetExecutingAssembly());

                // 创建TWAIN会话
                _session = new TwainSession(appId);

                // 注册事件处理
                _session.TransferReady += TwainSession_TransferReady;
                _session.DataTransferred += TwainSession_DataTransferred;
                _session.SourceDisabled += TwainSession_SourceDisabled;  // 添加SourceDisabled事件

                // 打开会话
                var openResult = _session.Open();
                if (openResult != ReturnCode.Success)
                {
                    Console.WriteLine($"打开TWAIN会话失败: {openResult}");
                    return;
                }

                Console.WriteLine("TWAIN会话已打开");

                // 查找并选择扫描仪
                DataSource selectedSource = null;
                var sources = _session.GetSources().ToList();

                if (sources.Count == 0)
                {
                    Console.WriteLine("未找到任何扫描仪");
                    _session.Close();
                    return;
                }

                Console.WriteLine("找到以下扫描仪:");
                foreach (var source in sources)
                {
                    Console.WriteLine($"  - {source.Name} (ID: {source.Id})");

                    // 选择第一个可用的扫描仪，或者特定品牌的扫描仪
                    if (selectedSource == null || source.Name.StartsWith("EPSON"))
                    {
                        selectedSource = source;
                    }
                }

                if (selectedSource != null)
                {
                    Console.WriteLine($"\n选择扫描仪: {selectedSource.Name}");

                    // 打开数据源
                    var openDsResult = selectedSource.Open();
                    if (openDsResult == ReturnCode.Success)
                    {
                        Console.WriteLine("扫描仪已打开，准备扫描...");

                        // 启用数据源进行扫描
                        // 使用模态窗口模式，这样用户可以设置扫描参数

                        if (selectedSource.Capabilities.ICapLampState.IsSupported)
                        {
                            selectedSource.Capabilities.ICapLampState.SetValue(BoolType.True);
                        }
                        if (selectedSource.Capabilities.CapFeederEnabled.IsSupported)
                        {
                            selectedSource.Capabilities.CapFeederEnabled.SetValue(BoolType.False);
                        }

                        if (selectedSource.Capabilities.ICapLightPath.IsSupported)
                        {
                            selectedSource.Capabilities.ICapLightPath.SetValue(LightPath.Transmissive);
                        }
                        if (selectedSource.Capabilities.ICapFilmType.IsSupported)
                        {
                            selectedSource.Capabilities.ICapFilmType.SetValue(FilmType.Negative);
                        }
                        selectedSource.Enable(SourceEnableMode.NoUI, false, IntPtr.Zero);
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
                if (_session != null && _session.IsSourceOpen)
                {
                    _session.Close();
                    Console.WriteLine("TWAIN会话已关闭");
                }
            }
        }

        private static void TwainSession_SourceDisabled(object sender, EventArgs e)
        {
            Console.WriteLine("扫描完成或用户取消了扫描");

            // 退出消息循环
            Application.Exit();
        }

        private static void TwainSession_TransferReady(object sender, TransferReadyEventArgs e)
        {
            Console.WriteLine("准备传输图像...");
            // 可以在这里设置是否接受传输
            e.CancelAll = false;  // 确保不取消传输
        }

        private static void TwainSession_DataTransferred(object sender, DataTransferredEventArgs e)
        {
            try
            {
                Console.WriteLine($"接收到数据，传输类型: {e.TransferType}");

                if (e.TransferType == XferMech.Native)
                {
                    // 处理Native传输类型（大多数扫描仪使用这种）
                    using (var stream = e.GetNativeImageStream())
                    {
                        if (stream != null)
                        {
                            using (var image = Image.FromStream(stream))
                            {
                                // 生成带时间戳的文件名
                                string fileName = $"C:\\Users\\hp\\scan_{DateTime.Now:yyyyMMdd_HHmmss}.jpg";
                                image.Save(fileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                                Console.WriteLine($"图像已保存到: {fileName}");
                            }
                        }
                    }
                }
                else if (e.TransferType == XferMech.File)
                {
                    // 处理文件传输类型
                    Console.WriteLine("文件传输模式");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理传输数据时出错: {ex.Message}");
            }
        }
    }
}