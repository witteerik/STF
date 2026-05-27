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
            var window = base.CreateWindow(activationState);

            var mainPage = new MainPage();

            // Adding an event handler that disposes the sound player 
            mainPage.Unloaded += MainPage_Unloaded;

            window.Page = mainPage;

            return window;
        }

        private void MainPage_Unloaded(object sender, EventArgs e)
        {
           Globals.StfBase.TerminateSTF();
        }
    }
}