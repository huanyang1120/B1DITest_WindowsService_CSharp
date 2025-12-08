using System;
using System.ServiceProcess;

namespace B1DITest_WindowsService_CSharp
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        //static void Main()
        //{
        //    ServiceBase[] ServicesToRun;
        //    ServicesToRun = new ServiceBase[]
        //    {
        //        new Service1()
        //    };
        //    ServiceBase.Run(ServicesToRun);
        //}
        static void Main(string[] args)
        {

            // 检查是否带有测试参数
            if (args.Length > 0 && args[0].ToLower() == "/test")
            {
                // 运行连接测试
                ConnectionTester.RunTest();
                Console.WriteLine("\n按任意键退出...");
                Console.ReadKey();
                return;
            }

            if (Environment.UserInteractive)
            {
                // 控制台模式（用于调试）
                Console.WriteLine("========================================");
                Console.WriteLine("  SAP B1 DI API 服务 - 调试模式");
                Console.WriteLine("========================================");
                Console.WriteLine();
                Console.WriteLine("提示:");
                Console.WriteLine("  使用 /test 参数运行连接测试");
                Console.WriteLine("  示例: B1DITest_WindowsService_CSharp.exe /test");
                Console.WriteLine();

                Service1 service = new Service1();

                try
                {
                    service.StartDebug(args);

                    Console.WriteLine();
                    Console.WriteLine("服务正在运行... 按 Ctrl+C 或任意键停止");
                    Console.WriteLine();

                    // 等待用户按键
                    Console.ReadKey();

                    service.StopDebug();
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine();
                    Console.WriteLine($"错误: {ex.Message}");
                    Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
                    Console.ResetColor();
                    Console.WriteLine();
                    Console.WriteLine("按任意键退出...");
                    Console.ReadKey();
                }
            }
            else
            {
                // Windows 服务模式
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                    new Service1()
                };
                ServiceBase.Run(ServicesToRun);
            }
        }
    }
}
