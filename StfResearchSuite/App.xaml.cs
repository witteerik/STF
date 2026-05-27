using STFN.Core;
using STFN;
using System.Diagnostics;

namespace OstfTabletSuite
{
    public partial class App : Application
    {
        public App()
        {
            InitializeComponent();
        }


        protected override Window CreateWindow(IActivationState? activationState)
        {
            var mainPage = new MainPage();
            mainPage.Unloaded += MainPage_Unloaded;
            return new Window(mainPage);
        }

        private void MainPage_Unloaded(object? sender, EventArgs e)
        {
           Globals.StfBase.TerminateSTF();
        }
    }
}