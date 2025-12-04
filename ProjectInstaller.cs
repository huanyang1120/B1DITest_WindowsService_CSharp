using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace WindowsServiceTestB1DI
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        private ServiceProcessInstaller serviceProcessInstaller;
        private ServiceInstaller serviceInstaller;
        public ProjectInstaller()
        {
            InitializeComponent();

            // 配置服务进程安装器
            serviceProcessInstaller = new ServiceProcessInstaller();
            serviceProcessInstaller.Account = ServiceAccount.LocalSystem;

            // 配置服务安装器
            serviceInstaller = new ServiceInstaller();
            serviceInstaller.ServiceName = "TestB1DIService";
            serviceInstaller.DisplayName = "我的 B1DI Windows 服务";
            serviceInstaller.Description = "这是一个Test B1DI 示例 Windows 服务";
            serviceInstaller.StartType = ServiceStartMode.Automatic;

            Installers.Add(serviceProcessInstaller);
            Installers.Add(serviceInstaller);
        }
    }
}
