using System.ComponentModel;

namespace Tecfy.OCR
{
    [RunInstaller(true)]
    public partial class Installer : System.Configuration.Install.Installer
    {
        public Installer()
        {
            InitializeComponent();
        }

        private void serviceInstaller_AfterInstall(object sender, System.Configuration.Install.InstallEventArgs e)
        {

        }
    }
}
