using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;

namespace GSRScheduleMailerService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }

        private void serviceProcessInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {

        }
    }
}