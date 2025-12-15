using System;
using System.IO;
using System.Text.Json;

namespace B1DITest_WindowsService_CSharp
{
    public class ConfigManager
    {
        //private static readonly string ConfigFilePath = @"C:\Temp\COM DI\CSharp\B1DITest_WindowsService_CSharp\config.json";
        // 使用相对路径，配置文件与 exe 在同一目录
        private static string ConfigFilePath
        {
            get
            {
                // 获取当前执行程序的目录
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                return Path.Combine(baseDirectory, "config.json");
            }
        }

        /// <summary>
        /// 读取配置文件
        /// </summary>
        public static ConfigModel LoadConfig()
        {
            try
            {
                string configPath = ConfigFilePath;

                if (!File.Exists(configPath))
                {
                    throw new FileNotFoundException($"配置文件不存在: {configPath}");
                }

                string jsonContent = File.ReadAllText(configPath);

                if (string.IsNullOrWhiteSpace(jsonContent))
                {
                    throw new Exception("配置文件为空");
                }

                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    ReadCommentHandling = JsonCommentHandling.Skip,
                    AllowTrailingCommas = true
                };

                var config = JsonSerializer.Deserialize<ConfigModel>(jsonContent, options);

                if (config == null)
                {
                    throw new Exception("配置文件解析后为空对象");
                }

                if (!config.IsValid())
                {
                    throw new Exception("配置文件缺少必要参数（ServerName, Company, B1Username, B1Password）");
                }

                return config;
            }
            catch (Exception ex)
            {
                throw new Exception($"读取配置文件失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 保存配置文件
        /// </summary>
        public static void SaveConfig(ConfigModel config)
        {
            try
            {
                string configPath = ConfigFilePath;
                string directory = Path.GetDirectoryName(configPath);

                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                string jsonContent = JsonSerializer.Serialize(config, options);
                File.WriteAllText(configPath, jsonContent);
            }
            catch (Exception ex)
            {
                throw new Exception($"保存配置文件失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 创建默认配置文件
        /// </summary>
        public static void CreateDefaultConfig()
        {
            var defaultConfig = new ConfigModel
            {
                ServerName = "iZ7y4gbss1x67fZ",
                DatabaseType = "dst_MSSQL2022",
                DatabaseUserName = "sa",
                DatabasePassword = "1234",
                LicenseServer = "iZ7y4gbss1x67fZ:30000",
                SLDServer = "iZ7y4gbss1x67fZ:40000",
                Company = "TestCN",
                B1Username = "manager",
                B1Password = "1234"
            };

            SaveConfig(defaultConfig);
        }

        /// <summary>
        /// 获取配置文件的完整路径
        /// </summary>
        public static string GetConfigFilePath()
        {
            return ConfigFilePath;
        }

        /// <summary>
        /// 检查配置文件是否存在
        /// </summary>
        public static bool ConfigFileExists()
        {
            return File.Exists(ConfigFilePath);
        }
    }
}