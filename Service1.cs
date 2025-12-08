using System;
using System.IO;
using System.ServiceProcess;
using System.Timers;
using System.Diagnostics;

namespace B1DITest_WindowsService_CSharp
{
    public partial class Service1 : ServiceBase
    {
        private Timer timer;
        private ConfigModel config;  // 确保这是类级别的字段
        private SAPB1Connection sapConnection;
        private bool isStarted = false;
        private bool isStarting = false; // 新增：防止重复启动标志

        //private string logPath = @"C:\Logs\B1DITest_WindowsService_CSharp\B1DITest.log";
        // 日志目录
        private readonly string logDirectory = @"C:\Logs\B1DITest_WindowsService_CSharp";

        // 动态获取当前小时的日志文件路径
        private string GetCurrentLogPath()
        {
            string fileName = $"B1DITest_{DateTime.Now:yyyyMMddHH}.log";
            return Path.Combine(logDirectory, fileName);
        }
        public Service1()
        {
            InitializeComponent();
            
            // 设置服务名称
            this.ServiceName = "B1DITest_WindowsService_CSharp";

            this.AutoLog = true;
            this.CanStop = true;
            this.CanPauseAndContinue = false;
        }

        protected override void OnStart(string[] args)
        {
            /*
            try { 
                // 创建日志目录
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
                WriteLog("========================================");
                WriteLog("=== 服务启动 ===");
                WriteLog("========================================");

                // 读取配置文件
                LoadConfiguration();

                // 连接 SAP B1
                ConnectToSAPB1();

                // 创建定时器，每30秒执行一次
                timer = new Timer(30000);
                timer.Elapsed += OnTimerElapsed;
                timer.Start();

                WriteLog("服务启动成功");
                WriteLog("========================================");
            }
            catch (Exception ex)
            {
                WriteLog($"服务启动失败: {ex.Message}");
                WriteLog($"堆栈跟踪: {ex.StackTrace}");
                throw;
            }*/

            // 防止重复启动
            if (isStarting || isStarted)
            {
                WriteLog("警告: 服务已经在启动或运行中，忽略重复启动请求");
                return;
            }

            isStarting = true;

            try
            {
                WriteLog("========================================");
                WriteLog("=== 服务启动开始 ===");
                WriteLog($"时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                WriteLog($"程序目录: {AppDomain.CurrentDomain.BaseDirectory}");
                WriteLog($"配置文件路径: {ConfigManager.GetConfigFilePath()}");
                WriteLog("========================================");

                EnsureLogDirectory();

                System.Threading.ThreadPool.QueueUserWorkItem(new System.Threading.WaitCallback(StartService), args);

                WriteLog("服务启动请求已提交");
            }
            catch (Exception ex)
            {
                WriteLog($"OnStart 发生严重错误: {ex.Message}");
                WriteLog($"堆栈跟踪: {ex.StackTrace}");
                WriteEventLog($"服务启动失败: {ex.Message}", EventLogEntryType.Error);
                throw;
            }
        }

        private void StartService(object state)
        {
            try
            {
                string[] args = state as string[];

                WriteLog("开始初始化服务组件...");

                // 1. 读取配置文件
                WriteLog("步骤 1/3: 读取配置文件");
                if (!TryLoadConfiguration())
                {
                    WriteLog("配置文件加载失败，服务将以有限功能运行");
                    isStarting = false;
                    return;
                }

                // 2. 连接 SAP B1
                WriteLog("步骤 2/3: 连接 SAP Business One");
                if (!TryConnectToSAPB1())
                {
                    WriteLog("SAP B1 连接失败，将在定时器中重试");
                }

                // 3. 启动定时器
                WriteLog("步骤 3/3: 启动定时器");
                StartTimer();

                isStarted = true;
                isStarting = false;
                WriteLog("========================================");
                WriteLog("=== 服务启动成功 ===");
                WriteLog("========================================");
            }
            catch (Exception ex)
            {
                isStarting = false;
                WriteLog($"StartService 发生错误: {ex.Message}");
                WriteLog($"堆栈跟踪: {ex.StackTrace}");
                WriteEventLog($"服务初始化失败: {ex.Message}", EventLogEntryType.Error);
            }
        }
        protected override void OnStop()
        {
            /*
            try
            {
                WriteLog("========================================");
                WriteLog("=== 服务停止 ===");
                WriteLog("========================================");

                timer?.Stop();
                timer?.Dispose();

                // 断开 SAP B1 连接
                DisconnectFromSAPB1();

                WriteLog("服务停止成功");
                WriteLog("========================================");
            }
            catch (Exception ex)
            {
                WriteLog($"服务停止时发生错误: {ex.Message}");
            }
            */
            try
            {
                WriteLog("========================================");
                WriteLog("=== 服务停止开始 ===");
                WriteLog($"时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                WriteLog("========================================");

                isStarted = false;

                if (timer != null)
                {
                    WriteLog("正在停止定时器...");
                    timer.Stop();
                    timer.Dispose();
                    timer = null;
                    WriteLog("定时器已停止");
                }

                DisconnectFromSAPB1();

                WriteLog("========================================");
                WriteLog("=== 服务停止成功 ===");
                WriteLog("========================================");
            }
            catch (Exception ex)
            {
                WriteLog($"OnStop 发生错误: {ex.Message}");
                WriteLog($"堆栈跟踪: {ex.StackTrace}");
                WriteEventLog($"服务停止时发生错误: {ex.Message}", EventLogEntryType.Warning);
            }
        }

        private void EnsureLogDirectory()
        {
            /*
            try
            {
                string directory = Path.GetDirectoryName(logPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                    WriteLog($"已创建日志目录: {directory}");
                }
            }
            catch (Exception ex)
            {
                logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "service.log");
                WriteLog($"无法创建日志目录，使用程序目录: {ex.Message}");
            }
            */
            try
            {
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                    WriteLog($"已创建日志目录: {logDirectory}");
                }
            }
            catch (Exception ex)
            {
                WriteLog($"无法创建日志目录: {ex.Message}");
            }
        }


        private bool TryLoadConfiguration()
        {
            try
            {
                WriteLog("正在读取配置文件...");
                WriteLog($"配置文件位置: {ConfigManager.GetConfigFilePath()}");

                // 检查配置文件是否存在
                if (!ConfigManager.ConfigFileExists())
                {
                    WriteLog("配置文件不存在，正在创建默认配置文件...");
                    ConfigManager.CreateDefaultConfig();
                    WriteLog($"默认配置文件已创建: {ConfigManager.GetConfigFilePath()}");
                    WriteLog("请修改配置文件并重启服务");
                    return false;
                }

                config = ConfigManager.LoadConfig();

                if (config == null)
                {
                    WriteLog("错误: 配置对象为空");
                    return false;
                }

                WriteLog("配置文件读取成功:");
                WriteLog($"  ServerName: {config.ServerName}");
                WriteLog($"  DatabaseType: {config.DatabaseType}");
                WriteLog($"  Company: {config.Company}");
                WriteLog($"  B1Username: {config.B1Username}");
                WriteLog($"  SLDServer: {config.SLDServer}");
                WriteLog($"  LicenseServer: {config.LicenseServer}");

                return true;
            }
            catch (FileNotFoundException ex)
            {
                WriteLog($"配置文件不存在: {ex.Message}");
                WriteLog("正在创建默认配置文件...");

                try
                {
                    ConfigManager.CreateDefaultConfig();
                    WriteLog($"默认配置文件已创建: {ConfigManager.GetConfigFilePath()}");
                    WriteLog("请修改配置文件并重启服务");
                }
                catch (Exception createEx)
                {
                    WriteLog($"创建配置文件失败: {createEx.Message}");
                }

                return false;
            }
            catch (Exception ex)
            {
                WriteLog($"读取配置文件失败: {ex.Message}");
                WriteLog($"堆栈跟踪: {ex.StackTrace}");
                return false;
            }
        }

        private bool TryConnectToSAPB1()
        {
            try
            {
                if (config == null)
                {
                    WriteLog("错误: 配置对象为空，无法连接 SAP B1");
                    return false;
                }

                WriteLog("正在连接到 SAP Business One...");
                WriteLog($"  服务器: {config.ServerName}");
                WriteLog($"  数据库类型: {config.DatabaseType}");
                WriteLog($"  公司数据库: {config.Company}");
                WriteLog($"  B1用户: {config.B1Username}");
                //WriteLog($"  数据库用户: {config.DatabaseUserName}");
                WriteLog($"  SLD服务器: {config.SLDServer}");
                WriteLog($"  License服务器: {config.LicenseServer}");

                if (sapConnection == null)
                {
                    sapConnection = new SAPB1Connection();
                }

                // 传递日志函数
                bool connected = sapConnection.Connect(config);

                if (connected)
                {
                    WriteLog("成功连接到 SAP Business One！");
                    WriteLog(sapConnection.GetCompanyInfo());
                    return true;
                }
                else
                {
                    WriteLog("连接失败: 未知错误");
                    return false;
                }
            }
            catch (Exception ex)
            {
                WriteLog($"连接 SAP B1 失败 in private bool TryConnectToSAPB1: {ex.Message}");
                WriteLog($"堆栈跟踪: {ex.StackTrace}");

                if (sapConnection != null)
                {
                    WriteLog($"详细错误(sapConnection.GetLastError): {sapConnection.GetLastError()}");
                }

                return false;
            }
        }
        private void DisconnectFromSAPB1()
        {
            try
            {
                if (sapConnection != null)
                {
                    WriteLog("正在断开 SAP Business One 连接...");
                    sapConnection.Disconnect();
                    sapConnection = null;
                    WriteLog("已断开 SAP B1 连接");
                }
            }
            catch (Exception ex)
            {
                WriteLog($"断开连接时发生错误: {ex.Message}");
            }
        }

        private void StartTimer()
        {
            try
            {
                if (timer != null)
                {
                    timer.Stop();
                    timer.Dispose();
                }

                timer = new Timer(60000);
                timer.Elapsed += OnTimerElapsed;
                timer.AutoReset = true;
                timer.Start();

                WriteLog("定时器已启动 (间隔: 30秒)");
            }
            catch (Exception ex)
            {
                WriteLog($"启动定时器失败: {ex.Message}");
                throw;
            }
        }

        private void OnTimerElapsed(object sender, ElapsedEventArgs e)
        {
            /*
            // 执行定时任务
            //WriteLog($"服务正在运行... {DateTime.Now}");
            try
            {
                WriteLog($"--- 定时任务执行 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ---");

                // 检查连接状态
                if (!sapConnection.TestConnection())
                {
                    WriteLog("警告: SAP B1 连接已断开，尝试重新连接...");
                    ConnectToSAPB1();
                    return;
                }

                // 执行业务逻辑
                ProcessBusinessLogic();
            }
            catch (Exception ex)
            {
                WriteLog($"执行定时任务时发生错误: {ex.Message}");
                WriteLog($"堆栈跟踪: {ex.StackTrace}");
            }*/

            if (!isStarted)
                return;

            try
            {
                WriteLog($"--- 定时任务执行 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ---");

                if (config == null)
                {
                    WriteLog("配置未加载，跳过任务");
                    return;
                }

                if (sapConnection == null || !sapConnection.TestConnection())
                {
                    WriteLog("警告: SAP B1 连接已断开");
                    WriteLog("尝试重新连接...");

                    if (TryConnectToSAPB1())
                    {
                        WriteLog("重新连接成功");
                    }
                    else
                    {
                        WriteLog("重新连接失败，将在下次定时任务时重试");
                        return;
                    }
                }

                ProcessBusinessLogic();

                WriteLog("定时任务执行完成");
            }
            catch (Exception ex)
            {
                WriteLog($"执行定时任务时发生错误: {ex.Message}");
                WriteLog($"堆栈跟踪: {ex.StackTrace}");
            }
        }
        private void ProcessBusinessLogic()
        {
            try
            {
                WriteLog("开始处理业务逻辑...");

                if (sapConnection == null || !sapConnection.IsConnected)
                {
                    WriteLog("错误: 未连接到 SAP B1");
                    return;
                }

                if (sapConnection.Company == null)
                {
                    WriteLog("错误: Company 对象为空");
                    return;
                }

                WriteLog($"当前连接公司: {config.Company}");
                WriteLog($"当前用户: {config.B1Username}");

                try
                {
                    string companyName = sapConnection.Company.CompanyName;
                    WriteLog($"公司名称: {companyName}");
                }
                catch (Exception ex)
                {
                    WriteLog($"获取公司信息失败: {ex.Message}");
                }

                // 示例：读取业务伙伴信息
                // SAPbobsCOM.BusinessPartners oBusinessPartner = 
                //     (SAPbobsCOM.BusinessPartners)sapConnection.Company.GetBusinessObject(
                //         SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                // if (oBusinessPartner.GetByKey("C00001"))
                // {
                //     WriteLog($"业务伙伴: {oBusinessPartner.CardCode} - {oBusinessPartner.CardName}");
                // }

                // System.Runtime.InteropServices.Marshal.ReleaseComObject(oBusinessPartner);

                //示例：读取 Item A001
                SAPbobsCOM.Items oItems =
                     (SAPbobsCOM.Items)sapConnection.Company.GetBusinessObject(
                         SAPbobsCOM.BoObjectTypes.oItems);

                if (oItems.GetByKey("A001"))
                {
                    WriteLog($"Item: {oItems.ItemCode} - {oItems.ItemName}");
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItems);

                WriteLog("业务逻辑处理完成");
            }
            catch (Exception ex)
            {
                WriteLog($"处理业务逻辑失败: {ex.Message}");
                WriteLog($"堆栈跟踪: {ex.StackTrace}");
                if (sapConnection != null)
                {
                    WriteLog($"SAP 详细错误: {sapConnection.GetLastError()}");
                }
            }
        }

        /// <summary>
        /// 写入日志（按小时分文件）
        /// </summary>
        private void WriteLog(string message)
        {
            /*
            try
            {
                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
                File.AppendAllText(logPath, logMessage + Environment.NewLine);

                if (Environment.UserInteractive)
                {
                    Console.WriteLine(logMessage);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    string backupLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "service_backup.log");
                    File.AppendAllText(backupLog,
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [日志写入失败: {ex.Message}] {message}{Environment.NewLine}");
                }
                catch
                {
                    WriteEventLog($"日志系统失败: {message}", EventLogEntryType.Warning);
                }
            }
            */
            try
            {
                string logPath = GetCurrentLogPath();
                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";

                // 确保目录存在
                string directory = Path.GetDirectoryName(logPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // 写入日志文件
                File.AppendAllText(logPath, logMessage + Environment.NewLine);

                // 如果是控制台模式，同时输出到控制台
                if (Environment.UserInteractive)
                {
                    Console.WriteLine(logMessage);
                }
            }
            catch (Exception ex)
            {
                // 尝试写入备用日志
                try
                {
                    string backupLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "service_error.log");
                    File.AppendAllText(backupLog,
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [日志写入失败: {ex.Message}] {message}{Environment.NewLine}");
                }
                catch
                {
                    WriteEventLog($"日志系统失败: {message}", EventLogEntryType.Warning);
                }
            }
        }


        private void WriteEventLog(string message, EventLogEntryType type)
        {
            try
            {
                if (!EventLog.SourceExists(this.ServiceName))
                {
                    EventLog.CreateEventSource(this.ServiceName, "Application");
                }

                EventLog.WriteEntry(this.ServiceName, message, type);
            }
            catch
            {
                // 忽略事件日志错误
            }
        }

        // 用于调试的入口点
        public void StartDebug(string[] args)
        {
            OnStart(args);
            System.Threading.Thread.Sleep(5000);
        }

        public void StopDebug()
        {
            OnStop();
        }
    }
}
