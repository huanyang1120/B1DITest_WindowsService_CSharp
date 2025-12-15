using System;
using System.Reflection.Emit;
using System.Text;
using SAPbobsCOM;

namespace B1DITest_WindowsService_CSharp
{
    public class SAPB1Connection
    {
        private Company oCompany;
        private ConfigModel config;
        private bool isConnected = false;

        public Company Company => oCompany;
        public bool IsConnected => isConnected && oCompany != null && oCompany.Connected;

        public bool GetSAPB1CompanyList(ConfigModel configuration, Action<string> logAction = null)
        {

            if (configuration == null)
            {
                throw new ArgumentNullException(nameof(configuration), "配置对象不能为空");
            }

            try
            {
                config = configuration;

                /*
                // 释放旧的连接
                if (oCompany != null && oCompany.Connected)
                {
                    //Disconnect();
                    Log(logAction, "Already connected to a B1 Company - " + oCompany.CompanyDB);
                    return true;
                }
                */

                // 创建 Company 对象
                oCompany = new Company();

                // 设置数据库服务器类型
                oCompany.DbServerType = config.GetDbServerType();
                oCompany.Server = config.ServerName;
                oCompany.UseTrusted = false;
                oCompany.DbUserName = config.DatabaseUserName;
                oCompany.DbPassword = config.DatabasePassword;

                // once the Server property of the Company is set
                // 获取公司列表
                var companyList = oCompany.GetCompanyList();
                Log(logAction, $"Get the company list Successfully. The RecordCount is: {companyList.RecordCount}");

                Log(logAction, $"The company object is: {oCompany.ToString()} {((object)oCompany).GetHashCode()}");
                if (oCompany.GetCompanyList().RecordCount>0)
                {
                    //Log(logAction, "Get the company list Successfully");
                    return true;
                }
                else
                {
                    Log(logAction, "B1 Company list is empty");
                    return false;
                }
            }
            catch (Exception ex)
            {
                // Use GetLastError method directly after a function
                // which doesn't have a return code
                // you may also use the On Error GoTo.
                // functions with no return codes throws exceptions
                try
                {
                    int errCode;
                    string errMsg;
                    oCompany.GetLastError(out errCode, out errMsg);

                    //throw new Exception($"连接失败 [错误代码: {errCode}]: {errMsg}");
                    throw new Exception($"Get Company List 失败 in public bool GetSAPB1CompanyList: \r\n oCompany.GetLastError is {errCode} {errMsg}\r\nex.Message is {ex.Message} \r\nex.StackTrace is {ex.StackTrace}", ex);
                }
                catch (Exception exc)
                {
                    throw new Exception($"Get Company List 失败 in public bool GetSAPB1CompanyList Catch: \r\nex.Message is {exc.Message} \r\nex.StackTrace is {exc.StackTrace}", exc);
                }
            }
        }

        /// <summary>
        /// 连接到 SAP Business One
        /// </summary>
        /// <param name="configuration">配置信息</param>
        /// <param name="logAction">日志回调函数（可选）</param>
        public bool Connect(ConfigModel configuration, Action<string> logAction = null)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException(nameof(configuration), "配置对象不能为空");
            }

            try
            {
                config = configuration;

                // 释放旧的连接
                if (oCompany != null && oCompany.Connected)
                {
                    //Disconnect();
                    Log(logAction, "Already connected to a B1 Company - "+oCompany.CompanyDB);
                    return true;
                }

                // 创建 Company 对象
                oCompany = new Company();


                // 记录连接参数（用于诊断）
                Log(logAction, "=== SAP B1 连接参数 ===");
                Log(logAction, $"Server: {config.ServerName}");
                Log(logAction, $"DbServerType: {config.GetDbServerType()} ({(int)config.GetDbServerType()})");
                Log(logAction, $"DbUserName: {config.DatabaseUserName}");
                Log(logAction, $"CompanyDB: {config.Company}");
                Log(logAction, $"UserName: {config.B1Username}");
                Log(logAction, $"LicenseServer: {config.LicenseServer}");
                Log(logAction, $"SLDServer: {config.SLDServer}");
                Log(logAction, "========================");

                // 设置数据库服务器类型
                oCompany.DbServerType = config.GetDbServerType();


                // 设置公司数据库
                oCompany.CompanyDB = config.Company;

                // 设置 B1 用户凭据
                oCompany.UserName = config.B1Username;
                oCompany.Password = config.B1Password;

                // 设置语言（可选）
                oCompany.language = BoSuppLangs.ln_English;

                oCompany.Server = config.ServerName;
                // 设置数据库连接信息
                //oCompany.DbUserName = config.DatabaseUserName ?? "";
                //oCompany.DbPassword = config.DatabasePassword ?? "";

                oCompany.LicenseServer = config.LicenseServer;
                oCompany.SLDServer = config.SLDServer;


                oCompany.UseTrusted = false;


                Log(logAction, $"The company object is: {oCompany.ToString()} {((object)oCompany).GetHashCode()}");
                // 尝试连接
                Log(logAction, "开始连接...");
                int lRetCode = oCompany.Connect();
                // 尝试连接
                Log(logAction, "连接...After Connect");

                if (lRetCode != 0)
                {
                    int errCode;
                    string errMsg;
                    oCompany.GetLastError(out errCode, out errMsg);

                    //throw new Exception($"连接失败 [错误代码: {errCode}]: {errMsg}");
                    // 提供更详细的错误信息
                    string detailedError = GetDetailedErrorMessage(errCode, errMsg);
                    throw new Exception(detailedError);
                }

                isConnected = true;
                Log(logAction, "连接成功！");
                return true;
            }
            catch (Exception ex)
            {
                isConnected = false;
                throw new Exception($"连接 SAP Business One 失败 in public bool Connect: \r\nex.Message is {ex.Message} \r\nex.StackTrace is {ex.StackTrace}", ex);
            }
    }


    /// <summary>
    /// 辅助日志方法
    /// </summary>
    private void Log(Action<string> logAction, string message)
    {
        logAction?.Invoke(message);
    }

    private string GetDetailedErrorMessage(int errCode, string errMsg)
    {
        string suggestion = "";

        switch (errCode)
        {
        case -132:
            suggestion = "\n\n可能的原因:\n" +
                       "1. 用户名或密码错误\n" +
                       "2. 用户账户被锁定或禁用\n" +
                       "3. 没有该公司的访问权限\n" +
                       "4. 许可证已过期或数量不足\n" +
                       "\n解决方案:\n" +
                       "- 验证用户名和密码是否正确\n" +
                       "- 在 SAP B1 客户端尝试用相同凭据登录\n" +
                       "- 检查用户是否被授权访问该公司数据库\n" +
                       "- 检查许可证管理器中的许可证状态\n" +
                       "- 运行 SQL: SELECT USER_CODE, U_NAME, LOCKED FROM OUSR WHERE USER_CODE = 'your_username'";
            break;

        case -104:
            suggestion = "\n\n无法连接到公司数据库\n" +
                       "- 检查公司数据库名称是否正确\n" +
                       "- 验证数据库是否在线\n" +
                       "- 运行 SQL: SELECT name FROM sys.databases WHERE name = 'your_company_db'";
            break;

        case -4004:
            suggestion = "\n\nSSL 证书问题\n" +
                       "- 在 config.json 中设置 \"TrustServerCertificate\": true\n" +
                       "- 或安装受信任的 SSL 证书";
            break;

        case -5002:
            suggestion = "\n\n无法连接到服务器\n" +
                        "- 检查服务器名称/IP 地址\n" +
                        "- 检查防火墙设置\n" +
                        "- 验证 SQL Server 是否运行\n" +
                        "- 测试命令: sqlcmd -S your_server -U sa -P password";
            break;


        case -1102:
            suggestion = "\n\n数据库用户认证失败\n" +
                       "- 检查数据库用户名和密码\n" +
                       "- 验证 SQL Server 认证模式（混合模式）";
            break;

        case -111:
            suggestion = "\n\n无效的许可证\n" +
                       "- 检查许可证是否已安装\n" +
                       "- 验证许可证是否过期\n" +
                       "- 检查许可证服务器连接";
            break;
        }

        return $"连接失败 [错误代码: {errCode}]: {errMsg}{suggestion}";
    }

    /// <summary>
    /// 断开连接
    /// </summary>
    public void Disconnect()
    {
        try
        {
            if (oCompany != null && oCompany.Connected)
            {
                oCompany.Disconnect();
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"断开连接失败: {ex.Message}", ex);
        }
        finally
        {
            if (oCompany != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompany);
                oCompany = null;
            }
            isConnected = false;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    /// <summary>
    /// 获取公司信息
    /// </summary>
    public string GetCompanyInfo()
    {
        if (!IsConnected)
        {
        return "未连接到 SAP B1";
        }

        try
        {
            var sb = new StringBuilder();
            sb.AppendLine($"公司名称: {oCompany.CompanyName}");
            sb.AppendLine($"公司数据库: {oCompany.CompanyDB}");
            sb.AppendLine($"SAP B1 版本: {oCompany.Version}");
            sb.AppendLine($"服务器: {oCompany.Server}");
            sb.AppendLine($"用户: {oCompany.UserName}");
            sb.AppendLine($"License服务器: {oCompany.LicenseServer}");
            sb.Append($"SLD服务器: {oCompany.SLDServer}"); // 最后一行不要换行

            return sb.ToString();
        }
        catch (Exception ex)
        {
        return $"获取公司信息失败: {ex.Message}";
        }
    }

    /// <summary>
    /// 获取最后的错误信息
    /// </summary>
    public string GetLastError()
    {
        if (oCompany == null)
        return "Company 对象未初始化";

        try
        {
        int errCode;
        string errMsg;
        oCompany.GetLastError(out errCode, out errMsg);
        return $"[错误代码: {errCode}] {errMsg}";
        }
        catch
        {
        return "无法获取错误信息";
        }
    }

    /// <summary>
    /// 测试连接
    /// </summary>
    public bool TestConnection()
    {
        try
        {
        return IsConnected && oCompany != null && oCompany.Connected;
        }
        catch
        {
        return false;
        }
        }
    }
}