using STFN.Core;
using STFM.Views;

using System.ComponentModel;

namespace STFM.Extension.Views
{

    [ToolboxItem(true)] // Marks this class as available for the Toolbox

    public partial class TestResultView_MultiFixedSiP : TestResultsView
    {
        public TestResultView_MultiFixedSiP()
        {
            InitializeComponent();
            switch (SharedSpeechTestObjects.GuiLanguage)
            {
                case STFN.Core.Utils.EnumCollection.Languages.Swedish:
                    StartButton.Text = "Start";
                    StopButton.Text = "Stop";
                    PauseButton.Text = "Pause";
                    ScreenShotButton.Text = "Spara";

                    break;
                default:
                    StartButton.Text = "Start";
                    StopButton.Text = "Stop";
                    PauseButton.Text = "Pause";
                    ScreenShotButton.Text = "Save";

                    break;
            }

        }

        public override void UpdateStartButtonText(string text)
        {
            StartButton.Text = text;
        }

        public override void ShowTestResults(string results)
        {
            // Not used
        }

        public override void ShowTestResults(SpeechTest speechTest)
        {

            // Making sure the current Twg has a panel
            List<string> panelTwgs = new List<string>();
            if (TwgPanel1.Twg != "") { panelTwgs.Add(TwgPanel1.Twg); }
            if (TwgPanel2.Twg != "") { panelTwgs.Add(TwgPanel2.Twg); }
            if (TwgPanel3.Twg != "") { panelTwgs.Add(TwgPanel3.Twg); }
            if (TwgPanel4.Twg != "") { panelTwgs.Add(TwgPanel4.Twg); }
            if (TwgPanel5.Twg != "") { panelTwgs.Add(TwgPanel5.Twg); }

            // Getting the observed test trials that belong to the last presented test word group
            var ObservedTrials = speechTest.GetObservedTestTrials();
            if (ObservedTrials.Count() == 0)
            {
                return;
            }

            string Twg = speechTest.GetObservedTestTrials().Last().SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation;
            List<TestTrial> ObservedTwgTrials = new List<TestTrial>();
            foreach (var testTrial in ObservedTrials)
            {
                if (testTrial.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation == Twg)
                {
                    ObservedTwgTrials.Add(testTrial);
                }
            }

            if (panelTwgs.Contains(Twg) == false)
            {
                // Adding the TWG string to the first available panel
                if (TwgPanel1.Twg == "") { TwgPanel1.Twg = Twg; }
                else if (TwgPanel2.Twg == "") { TwgPanel2.Twg = Twg; }
                else if (TwgPanel3.Twg == "") { TwgPanel3.Twg = Twg; }
                else if (TwgPanel4.Twg == "") { TwgPanel4.Twg = Twg; }
                else if (TwgPanel5.Twg == "") { TwgPanel5.Twg = Twg; }
            }

            // Determining which group to update
            TwgPanelView_ConstantStimuli? currentTwgPanel = null;

            if (Twg == TwgPanel1.Twg)
                currentTwgPanel = TwgPanel1;
            else if (Twg == TwgPanel2.Twg)
                currentTwgPanel = TwgPanel2;
            else if (Twg == TwgPanel3.Twg)
                currentTwgPanel = TwgPanel3;
            else if (Twg == TwgPanel4.Twg)
                currentTwgPanel = TwgPanel4;
            else if (Twg == TwgPanel5.Twg)
                currentTwgPanel = TwgPanel5;

            if (currentTwgPanel is null)
                return;

            // Referencing the SnrDiagram locally
            TestResultsDiagram MySnrDiagram = (TestResultsDiagram)currentTwgPanel.GetSnrView().Drawable;
            
            currentTwgPanel.Twg = Twg;

            // Reference level
            currentTwgPanel.ReferenceLevel = speechTest.ReferenceLevel.ToString() + " dB SPL";

            // Trial count / progress
            currentTwgPanel.TrialNumber = (1 + ObservedTwgTrials.Count()).ToString();


            // SNR diagram (Not updating the SNR diagram in practise tests)
            if (speechTest.IsPractiseTest == false)
            {

                // Setting up default scales in the diagram
                List<float> X_PnrArray = new List<float>();
                SortedSet<double> TestPNRs = new SortedSet<double>();
                
                foreach (STFN.Core.SipTest.SipTrial trial in ObservedTwgTrials)
                {
                        TestPNRs.Add(trial.PNR);
                }

                foreach (double pnr in TestPNRs)
                {
                    X_PnrArray.Add((float)pnr);
                }

                MySnrDiagram.SetXlim(X_PnrArray.Min() - 2.5f, X_PnrArray.Max() + 2.5f);
                MySnrDiagram.SetYlim(0, 105);

                List<float> YaxisLinePositions = new List<float>();
                List<float> YaxisTextPositions = new List<float>();
                List<string> YaxisTextValues = new List<string>();
                for (int i = 0; i < 100 + 10; i += 20)
                {
                    YaxisTextPositions.Add((float)i);
                    YaxisTextValues.Add(i.ToString());
                    if (i != 0) { YaxisLinePositions.Add((float)i); }
                }

                List<float> XaxisTextPositions = new List<float>();
                List<string> XaxisTextValues = new List<string>();
                for (int i = 0; i < X_PnrArray.Count; i++)
                {
                    XaxisTextPositions.Add(X_PnrArray[i]);
                    XaxisTextValues.Add(X_PnrArray[i].ToString());
                }

                MySnrDiagram.SetTickTextsY(YaxisTextPositions, YaxisTextValues.ToArray());
                MySnrDiagram.SetTickTextsX(XaxisTextPositions, XaxisTextValues.ToArray());

                MySnrDiagram.SetYaxisDashedGridLinePositions(YaxisLinePositions);

                MySnrDiagram.SetSizeModificationStrategy(PlotBase.SizeModificationStrategies.Horizontal);
                MySnrDiagram.SetTextSizeAxisX(0.8f);
                MySnrDiagram.SetTextSizeAxisY(0.8f);

                // Calculating psychometric function
                SortedList<double, List<double>> psychometricFunctionData = new SortedList<double, List<double>>();
                foreach (var pnr in TestPNRs)
                {
                    psychometricFunctionData.Add(pnr, new List<double>());
                }

                // Sorting results into PNRs
                foreach (STFN.Core.SipTest.SipTrial trial in ObservedTwgTrials)
                {
                    if (trial.IsTSFCTrial)
                    {
                        psychometricFunctionData[trial.PNR].Add(trial.GradedResponse);
                    }
                    else
                    {
                        if (trial.IsCorrect == true)
                        {
                            psychometricFunctionData[trial.PNR].Add(1);
                        }
                        else 
                        {
                            psychometricFunctionData[trial.PNR].Add(0);
                        }
                    }
                }

                List<float> PresentedPnrs = new List<float>();
                foreach (var pnr in psychometricFunctionData.Keys)
                {
                    PresentedPnrs.Add((float)pnr);
                }

                List<float> AverageScoresPerPnr = new List<float>();
                foreach (var pnr in psychometricFunctionData.Keys)
                {
                    // Calculating average and convert to percentage
                    AverageScoresPerPnr.Add((float)psychometricFunctionData[pnr].Average()*100);
                }

                MySnrDiagram.PointSeries.Clear();
                MySnrDiagram.Lines.Clear();

                MySnrDiagram.PointSeries.Add(new PointSerie() { Color = Colors.Red, PointSize = 1, Type = PointSerie.PointTypes.Cross, XValues = PresentedPnrs.ToArray(), YValues = AverageScoresPerPnr.ToArray() });
                MySnrDiagram.Lines.Add(new Line() { Color = Colors.Blue, Dashed = false, LineWidth = 2, XValues = PresentedPnrs.ToArray(), YValues = AverageScoresPerPnr.ToArray() });

                MySnrDiagram.UpdateLayout();

            }

        }

        private void StartButton_Clicked(object sender, EventArgs e)
        {
            OnStartedFromTestResultView(new EventArgs());
        }

        private void PauseButton_Clicked(object sender, EventArgs e)
        {
            OnPausedFromTestResultView(new EventArgs());
        }
        private void StopButton_Clicked(object sender, EventArgs e)
        {
            OnStoppedFromTestResultView(new EventArgs());
        }

        private void ScreenShotButton_Clicked(object sender, EventArgs e)
        {
            TakeScreenShot();
        }
        public override void SetGuiLayoutState(SpeechTestView.GuiLayoutStates currentTestPlayState)
        {


            switch (currentTestPlayState)
            {
                case SpeechTestView.GuiLayoutStates.InitialState:
                    StartButton.IsEnabled = false;
                    PauseButton.IsEnabled = false;
                    StopButton.IsEnabled = false;
                    ScreenShotButton.IsEnabled = false;

                    break;
                case SpeechTestView.GuiLayoutStates.TestSelection:
                    StartButton.IsEnabled = false;
                    PauseButton.IsEnabled = false;
                    StopButton.IsEnabled = false;
                    ScreenShotButton.IsEnabled = false;

                    break;
                case SpeechTestView.GuiLayoutStates.SpeechMaterialSelection:
                    StartButton.IsEnabled = false;
                    PauseButton.IsEnabled = false;
                    StopButton.IsEnabled = false;
                    ScreenShotButton.IsEnabled = false;

                    break;
                case SpeechTestView.GuiLayoutStates.TestOptions_StartButton_TestResultsOnForm:
                    StartButton.IsEnabled = true;
                    PauseButton.IsEnabled = false;
                    StopButton.IsEnabled = false;
                    ScreenShotButton.IsEnabled = false;

                    break;
                case SpeechTestView.GuiLayoutStates.TestOptions_StartButton_TestResultsOffForm:
                    StartButton.IsEnabled = true;
                    PauseButton.IsEnabled = false;
                    StopButton.IsEnabled = false;
                    ScreenShotButton.IsEnabled = false;

                    break;
                case SpeechTestView.GuiLayoutStates.TestIsRunning:
                    StartButton.IsEnabled = false;
                    PauseButton.IsEnabled = true;
                    StopButton.IsEnabled = true;
                    ScreenShotButton.IsEnabled = false;

                    break;
                case SpeechTestView.GuiLayoutStates.TestIsPaused:
                    StartButton.IsEnabled = true;
                    PauseButton.IsEnabled = false;
                    StopButton.IsEnabled = true;
                    ScreenShotButton.IsEnabled = true;

                    break;
                case SpeechTestView.GuiLayoutStates.TestIsStopped:
                    StartButton.IsEnabled = false;
                    PauseButton.IsEnabled = false;
                    StopButton.IsEnabled = false;
                    ScreenShotButton.IsEnabled = true;

                    break;
                default:
                    break;
            }


        }

    }
}