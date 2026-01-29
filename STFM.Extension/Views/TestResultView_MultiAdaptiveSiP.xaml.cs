// License
// This project is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
// Commercial use requires a separate commercial license.See the `LICENSE` file for details.
// 
// Copyright (c) 2025 Erik Witte

using STFN.Core;
using STFM.Views;

using System.ComponentModel;

namespace STFM.Extension.Views
{

    [ToolboxItem(true)] // Marks this class as available for the Toolbox
    public partial class TestResultView_MultiAdaptiveSiP : TestResultsView
    {
        public TestResultView_MultiAdaptiveSiP()
        {
            InitializeComponent();
            //this.LoadFromXaml(typeof(TestResultView_Adaptive));


            //// Assign the custom drawable to the GraphicsView
            //SnrView.Drawable = new TestResultsDiagram(SnrView);
            //TestResultsDiagram MySnrDiagram = (TestResultsDiagram)SnrView.Drawable;
            //MySnrDiagram.SetSizeModificationStrategy(PlotBase.SizeModificationStrategies.Horizontal);
            //MySnrDiagram.SetXlim(0.5f, 1.5f);
            //MySnrDiagram.SetYlim(-10, 10);
            ////MySnrDiagram.TransitionHeightRatio = 0.86f;
            ////MySnrDiagram.Background = Colors.DarkSlateGray;
            //MySnrDiagram.UpdateLayout();

            //// Force redraw on size change
            //SnrView.SizeChanged += (s, e) => SnrView.Invalidate();


            switch (SharedSpeechTestObjects.GuiLanguage)
            {
                case STFN.Core.Utils.EnumCollection.Languages.Swedish:
                    StartButton.Text = "Start";
                    StopButton.Text = "Stop";
                    PauseButton.Text = "Pause";
                    ScreenShotButton.Text = "Spara";

                    //TwgNameLabel.Text = "Grupp:";
                    //TargetScoreNameLabel.Text = "Mål, andel korrekt:";
                    //ReferenceLevelNameLabel.Text = "Referensnivå:";
                    //AdaptiveLevelNameLabel.Text = "Adaptiv nivå:";
                    //TrialNumberNameLabel.Text = "Försök nummer:";
                    //FinalResultNameLabel.Text = "Hörtröskel:";

                    //SnrGridLabelY.Text = "PNR (dB)";
                    //SnrGridLabelX.Text = "Försök nummer";

                    //GroupScoreTitleLabel.Text = "Andel korrekt sista 10 försöken";
                    //TwgNameLabel1.Text = "Grupp 1";
                    //TwgNameLabel2.Text = "Grupp 2";
                    //TwgNameLabel3.Text = "Grupp 3";
                    //TwgNameLabel4.Text = "Grupp 4";
                    //TwgNameLabel5.Text = "Grupp 5";
                    //TenLastScoreNameLabel.Text = "Alla grupper";

                    break;
                default:
                    StartButton.Text = "Start";
                    StopButton.Text = "Stop";
                    PauseButton.Text = "Pause";
                    ScreenShotButton.Text = "Save";

                    //SnrGridLabelY.Text = "PNR (dB)";
                    //SnrGridLabelX.Text = "Test trial";

                    //TargetScoreNameLabel.Text = "Target score:";
                    //ReferenceLevelNameLabel.Text = "Reference level:";
                    //AdaptiveLevelNameLabel.Text = "Adaptive level:";
                    //TrialNumberNameLabel.Text = "Trial number:";
                    //FinalResultNameLabel.Text = "SRT";

                    //GroupScoreTitleLabel.Text = "Average scores in last 10 trials";
                    //TwgNameLabel1.Text = "Group 1";
                    //TwgNameLabel2.Text = "Group 2";
                    //TwgNameLabel3.Text = "Group 3";
                    //TwgNameLabel4.Text = "Group 4";
                    //TwgNameLabel5.Text = "Group 5";
                    //TenLastScoreNameLabel.Text = "Overall score";

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

            TestProtocol testProtocol = speechTest.TestProtocol;

            if (speechTest.TestProtocol == null)
            {
                // This works bad without a test protocol. Returning to prevent exceptions
                return;
            }

            // Making sure the current Twg has a panel
            List<string> panelTwgs = new List<string>();
            if (TwgPanel1.Twg != "") { panelTwgs.Add(TwgPanel1.Twg); }
            if (TwgPanel2.Twg != "") { panelTwgs.Add(TwgPanel2.Twg); }
            if (TwgPanel3.Twg != "") { panelTwgs.Add(TwgPanel3.Twg); }
            if (TwgPanel4.Twg != "") { panelTwgs.Add(TwgPanel4.Twg); }
            if (TwgPanel5.Twg != "") { panelTwgs.Add(TwgPanel5.Twg); }

            if (panelTwgs.Contains(speechTest.TestProtocol.TargetStimulusSet) == false)
            {
                // Adding the TWG string to the first available panel
                if (TwgPanel1.Twg == "") { TwgPanel1.Twg = speechTest.TestProtocol.TargetStimulusSet; }
                else if (TwgPanel2.Twg == "") { TwgPanel2.Twg = speechTest.TestProtocol.TargetStimulusSet; }
                else if (TwgPanel3.Twg == "") { TwgPanel3.Twg = speechTest.TestProtocol.TargetStimulusSet; }
                else if (TwgPanel4.Twg == "") { TwgPanel4.Twg = speechTest.TestProtocol.TargetStimulusSet; }
                else if (TwgPanel5.Twg == "") { TwgPanel5.Twg = speechTest.TestProtocol.TargetStimulusSet; }
            }

            // Determining which group to update
            TwgPanelView? currentTwgPanel = null;

            if (speechTest.TestProtocol.TargetStimulusSet == TwgPanel1.Twg)
                currentTwgPanel = TwgPanel1;
            else if (speechTest.TestProtocol.TargetStimulusSet == TwgPanel2.Twg)
                currentTwgPanel = TwgPanel2;
            else if (speechTest.TestProtocol.TargetStimulusSet == TwgPanel3.Twg)
                currentTwgPanel = TwgPanel3;
            else if (speechTest.TestProtocol.TargetStimulusSet == TwgPanel4.Twg)
                currentTwgPanel = TwgPanel4;
            else if (speechTest.TestProtocol.TargetStimulusSet == TwgPanel5.Twg)
                currentTwgPanel = TwgPanel5;

            if (currentTwgPanel is null)
                return;



            // Referencing the SnrDiagram locally
            TestResultsDiagram MySnrDiagram = (TestResultsDiagram)currentTwgPanel.GetSnrView().Drawable;
            //TestResultsDiagram MySnrDiagram = (TestResultsDiagram)SnrView.Drawable;

            // Getting all trials so far
            var AllObservedTestTrials = speechTest.GetObservedTestTrials();

            // Adding the trial only if it's in the same TWG as the currently tested
            List<TestTrial> ObservedTestTrials = new List<TestTrial>();
            foreach (var trial in AllObservedTestTrials)
            {
                if (trial.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation == speechTest.TestProtocol.TargetStimulusSet)
                {
                    ObservedTestTrials.Add(trial);
                }
            }

            // Test word group
            if (ObservedTestTrials.Count() > 0)
            {
                var currentTWG = ObservedTestTrials.Last();
                currentTwgPanel.Twg = currentTWG.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation;
            }


            // Target score
            if (testProtocol.TargetScore.HasValue)
            {
                if (speechTest.IsPractiseTest == false)
                {
                    currentTwgPanel.TargetScore = System.Math.Round(100 * testProtocol.TargetScore.Value, 0).ToString() + "%";
                }
                else
                {
                    currentTwgPanel.TargetScore = "—";
                }
            }

            // Reference level
            currentTwgPanel.ReferenceLevel = speechTest.ReferenceLevel.ToString() + " dB SPL";

            // PNR
            double? CurrentAdaptiveValue = testProtocol.GetCurrentAdaptiveValue();
            if (CurrentAdaptiveValue.HasValue)
            {
                currentTwgPanel.AdaptiveLevel = System.Math.Round(CurrentAdaptiveValue.Value, 1).ToString() + " dB PNR";
            }
            else
            {
                currentTwgPanel.AdaptiveLevel = "—";
            }

            // Trial count / progress
            if (ObservedTestTrials.Count() > 0)
            {
                currentTwgPanel.TrialNumber = (1 + ObservedTestTrials.Count()).ToString() + " / " + speechTest.GetTotalTrialCount().ToString();
            }
            else
            {
                currentTwgPanel.TrialNumber = "1 / " + speechTest.GetTotalTrialCount().ToString();
            }

            // SRT:
            double? FinalResult = testProtocol.GetFinalResultValue();
            //switch (SharedSpeechTestObjects.GuiLanguage)
            //{
            //    case STFN.Core.Utils.EnumCollection.Languages.English:
            //        FinalResultNameLabel.Text = "SRT:";
            //        break;
            //    case STFN.Core.Utils.EnumCollection.Languages.Swedish:
            //        FinalResultNameLabel.Text = "Hörtröskel:";
            //        break;
            //    default:
            //        FinalResultNameLabel.Text = "SRT:";
            //        break;
            //}
            if (FinalResult.HasValue)
            {
                currentTwgPanel.FinalResult = System.Math.Round(FinalResult.Value, 1).ToString() + " dB PNR";
            }
            else
            {
                currentTwgPanel.FinalResult = "—";
            }


            // SNR diagram (Not updating the SNR diagram in practise tests)
            if (speechTest.IsPractiseTest == false)
            {

                List<float> PresentedPnrs = new List<float>();
                List<float> PresentedTrials = new List<float>();

                float presentedTrialIndex = 1;
                foreach (TestTrial trial in ObservedTestTrials)
                {
                    PresentedPnrs.Add((float)trial.AdaptiveProtocolValue);
                    PresentedTrials.Add(presentedTrialIndex);
                    presentedTrialIndex += 1;
                }

                // Adding also the current adaptive value (which has not yet been stored)
                if (CurrentAdaptiveValue != null)
                {
                    PresentedPnrs.Add((float)CurrentAdaptiveValue.Value);
                    PresentedTrials.Add(presentedTrialIndex);
                }


                MySnrDiagram.SetSizeModificationStrategy(PlotBase.SizeModificationStrategies.Horizontal);
                MySnrDiagram.SetTextSizeAxisX(0.8f);
                MySnrDiagram.SetTextSizeAxisY(0.8f);

                MySnrDiagram.SetXlim(0.5f, PresentedPnrs.Count + 0.5f);

                float Ymin = PresentedPnrs.Min() - 5f;
                float Ymax = PresentedPnrs.Max() + 5f;
                MySnrDiagram.SetYlim(Ymin, Ymax);


                List<float> YaxisTextPositions = new List<float>();
                List<string> YaxisTextValues = new List<string>();

                int Steps = 6;
                int StepSize = (int)((Ymax - Ymin) / Steps);
                for (int i = 0; i < Steps + 10; i++)
                {
                    YaxisTextPositions.Add((float)System.Math.Round(Ymin, 0) + (i * StepSize));
                    YaxisTextValues.Add(((float)System.Math.Round(Ymin, 0) + (i * StepSize)).ToString());
                }

                List<float> XaxisTextPositions = new List<float>();
                List<string> XaxisTextValues = new List<string>();

                int TrialSteps = 1 + (int)(PresentedTrials.Count / 15);
                for (int i = 0; i < PresentedTrials.Count; i += TrialSteps)
                {
                    XaxisTextPositions.Add(PresentedTrials[i]);
                    XaxisTextValues.Add(PresentedTrials[i].ToString());
                }

                MySnrDiagram.SetTickTextsY(YaxisTextPositions, YaxisTextValues.ToArray());
                MySnrDiagram.SetTickTextsX(XaxisTextPositions, XaxisTextValues.ToArray());

                MySnrDiagram.PointSeries.Clear();
                MySnrDiagram.Lines.Clear();

                MySnrDiagram.PointSeries.Add(new PointSerie() { Color = Colors.Red, PointSize = 1, Type = PointSerie.PointTypes.Cross, XValues = PresentedTrials.ToArray(), YValues = PresentedPnrs.ToArray() });
                MySnrDiagram.Lines.Add(new Line() { Color = Colors.Blue, Dashed = false, LineWidth = 2, XValues = PresentedTrials.ToArray(), YValues = PresentedPnrs.ToArray() });

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
