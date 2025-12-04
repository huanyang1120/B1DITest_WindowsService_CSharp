using System;
using SAPbobsCOM;

namespace WindowsServiceTestB1DI
{
    /// <summary>
    /// 连接测试工具类（不包含 Main 方法）
    /// </summary>
    public class ConnectionTester
    {
        /// <summary>
        /// 执行连接测试
        /// </summary>
        public static void RunTest()
        {
            Console.WriteLine("========================================");
            Console.WriteLine("  SAP B1 连接测试工具");
            Console.WriteLine("========================================");
            Console.WriteLine();

            try
            {
                // 读取配置
                Console.WriteLine("正在读取配置文件...");
                var config = ConfigManager.LoadConfig();
                Console.WriteLine($"配置文件路径: {ConfigManager.GetConfigFilePath()}");
                Console.WriteLine();

                // 显示配置
                Console.WriteLine("配置信息:");
                Console.WriteLine(config.ToString());
                Console.WriteLine();

                // 测试连接
                Console.WriteLine("开始测试连接...");
                Console.WriteLine("----------------------------------------");

                var connection = new SAPB1Connection();

                // 使用 lambda 表达式传递日志回调
                bool success = connection.Connect(config, (msg) => Console.WriteLine(msg));

                Console.WriteLine("----------------------------------------");
                Console.WriteLine();

                if (success)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("✓ 连接成功！");
                    Console.ResetColor();
                    Console.WriteLine();

                    Console.WriteLine("公司信息:");
                    Console.WriteLine(connection.GetCompanyInfo());
                    Console.WriteLine();

                    // 测试数据访问
                    Console.WriteLine("测试数据访问...");
                    TestDataAccess(connection.Company);

                    connection.Disconnect();
                    Console.WriteLine("已断开连接");
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("✗ 连接失败");
                    Console.ResetColor();
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n错误: {ex.Message}");
                Console.WriteLine($"\n堆栈跟踪:\n{ex.StackTrace}");
                Console.ResetColor();
            }
        }

        /// <summary>
        /// 测试数据访问
        /// </summary>
        private static void TestDataAccess(Company oCompany)
        {
            try
            {
                Console.WriteLine($"  公司名称: {oCompany.CompanyName}");
                //Console.WriteLine($"  服务器时间: {oCompany.ServerTime}");
                //Console.WriteLine($"  系统货币: {oCompany.SystemCurrency}");
               // Console.WriteLine($"  本地货币: {oCompany.LocalCurrency}");

                // 测试获取用户信息
                SAPbobsCOM.Users oUser = (SAPbobsCOM.Users)oCompany.GetBusinessObject(BoObjectTypes.oUsers);
                if (oUser.GetByKey(oCompany.UserSignature))
                {
                    Console.WriteLine($"  当前用户: {oUser.UserName} ({oUser.UserCode})");
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUser);

                Console.WriteLine("  数据访问测试成功 ✓");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  数据访问测试失败: {ex.Message}");
            }
        }
    }
}