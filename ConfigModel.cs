using SAPbobsCOM;

namespace B1DITest_WindowsService_CSharp
{
    public class ConfigModel
    {
        public string ServerName { get; set; }
        public string DatabaseType { get; set; }
        public string DatabaseUserName { get; set; }
        public string DatabasePassword { get; set; }
        public string LicenseServer { get; set; }
        public string SLDServer { get; set; }
        public string Company { get; set; }
        public string B1Username { get; set; }
        public string B1Password { get; set; }


        /// <summary>
        /// 获取数据库类型枚举值
        /// </summary>
        public SAPbobsCOM.BoDataServerTypes GetDbServerType()
        {
            if (string.IsNullOrEmpty(DatabaseType))
            {
                return BoDataServerTypes.dst_MSSQL2022;
            }

            switch (DatabaseType?.ToUpper())
            {
                case "DST_MSSQL2012":
                case "MSSQL2012":
                case "7":
                    return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

                case "DST_MSSQL2014":
                case "MSSQL2014":
                case "8":
                    return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;

                case "DST_HANADB":
                case "HANA":
                case "HANADB":
                case "9":
                    return SAPbobsCOM.BoDataServerTypes.dst_HANADB;

                case "DST_MSSQL2016":
                case "MSSQL2016":
                case "10":
                    return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;

                case "DST_MSSQL2017":
                case "MSSQL2017":
                case "11":
                    return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;

                case "DST_MSSQL2019":
                case "MSSQL2019":
                case "15":
                    return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;


                case "DST_MSSQL2022":
                case "MSSQL2022":
                case "17":
                    return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2022;

                default:
                    // 默认使用 MSSQL2022 (17)
                    return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2022;
            }
        }

        /// <summary>
        /// 验证配置
        /// </summary>
        public bool IsValid()
        {
            return !string.IsNullOrEmpty(ServerName) &&
                   !string.IsNullOrEmpty(Company) &&
                   !string.IsNullOrEmpty(B1Username) &&
                   !string.IsNullOrEmpty(B1Password);
        }

        /// <summary>
        /// 显示配置信息（隐藏密码）
        /// </summary>
        public override string ToString()
        {
            return $"ServerName: {ServerName}\n" +
                   $"DatabaseType: {DatabaseType}\n" +
                   $"DatabaseUserName: {DatabaseUserName}\n" +
                   $"DatabasePassword: ****\n" +
                   $"LicenseServer: {LicenseServer}\n" +
                   $"SLDServer: {SLDServer}\n" +
                   $"Company: {Company}\n" +
                   $"B1Username: {B1Username}\n" +
                   $"B1Password: ****" ;
        }
    }
}