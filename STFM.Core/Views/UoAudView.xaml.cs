// License
// This project is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
// Commercial use requires a separate commercial license.See the `LICENSE` file for details.
//
// Copyright (c) 2025 Erik Witte

using STFN.Core;

namespace STFM.Views;

public partial class UoAudView : ContentView
{

    public event EventHandler<EventArgs> EnterFullScreenMode;
    public event EventHandler<EventArgs> ExitFullScreenMode;

    /// <summary>
    /// The event fires upon test completion.
    /// </summary>
    public event EventHandler<EventArgs> Finished;

    /// <summary>
    /// This variable can be set to true to show the results directly on in the view.
    /// </summary>
    public bool ShowResultsInView = false;

    private Brush NonPressedBrush = Colors.LightGray;
    private Brush PressedBrush = Colors.Yellow;

    private UoPta CurrentTest;

    UoPta.PtaTrial CurrentTrial;

    private STFN.Core.Audio.Formats.WaveFormat WaveFormat;

    private SortedList<int, double> RetSplList;
    private SortedList<int, double> PureToneCalibrationList = new SortedList<int, double>();

    private STFN.Core.Audio.Sound silentSound = null;

    IDispatcherTimer TrialEndTimer;

    private bool TestIsStarted = false;
    private bool permanentlyStopped;
    public event EventHandler<UoAudErrorEventArgs>? Failed;
    public bool StopWhenUnloaded { get; set; }

    /// <summary>Stops this run and invalidates any pending responses. Use a new view to restart.</summary>
    public void StopTest()
    {
        if (permanentlyStopped) return;
        permanentlyStopped = true;
        TestIsStarted = false;
        if (TrialEndTimer is not null)
        {
            TrialEndTimer.Stop();
            TrialEndTimer.Tick -= TrialEndTimer_Tick;
        }
        CurrentTrial = null;
        MidResponseButton.IsEnabled = false;
        MessageButton.IsEnabled = false;
        MidResponseButton.Background = NonPressedBrush;
        if (Globals.StfBase.SoundPlayer is not null)
        {
            try { Globals.StfBase.SoundPlayer.FadeOutPlayback(); }
            catch { Globals.StfBase.SoundPlayer.CloseStream(); }
        }
    }

    private void FailTest(Exception exception)
    {
        try { StopTest(); }
        catch (Exception stopException) { System.Diagnostics.Debug.WriteLine(stopException); }
        Failed?.Invoke(this, new UoAudErrorEventArgs(exception.Message, exception));
    }

    public UoAudView(UoPta CurrentTest)
    {
        InitializeComponent();
        Unloaded += (_, _) =>
        {
            if (!StopWhenUnloaded) return;
            try { StopTest(); }
            catch (Exception exception) { FailTest(exception); }
        };

        // Setting default texts
        switch (SharedSpeechTestObjects.GuiLanguage)
        {
            case STFN.Core.Utils.EnumCollection.Languages.English:

                MessageButton.Text = "Click here to start!";
                MessageLabel.Text = "Press when you hear a tone.\nRelease when the tone disappears.";

                LevelLabel.Text = "";
                RatingLabel.Text = "";
                ResultView.Text = "";

                break;

            case STFN.Core.Utils.EnumCollection.Languages.Swedish:

                MessageButton.Text = "Tryck här för att starta!";
                MessageLabel.Text = "Tryck på knappen nedanför när du hör en ton.\nSläpp knappen när tonen tystnar.";

                MidResponseButton.Text = "Tryck här!";

                LevelLabel.Text = "";
                RatingLabel.Text = "";
                ResultView.Text = "";

                break;
            default:
                // Using English as default

                MessageButton.Text = "Click here to start!";
                MessageLabel.Text = "Press when you hear a tone.\nRelease when the tone disappears.";

                LevelLabel.Text = "";
                RatingLabel.Text = "";
                ResultView.Text = "";

                break;
        }

        // Setting default colors
        MidResponseButton.Background = NonPressedBrush;
        MidResponseButton.TextColor = Colors.DarkGrey;

        // Setting default visibility
        SetVisibility(VisibilityTypes.MessageMode);

        // Initializing the CurrentTest
        this.CurrentTest = CurrentTest;

        // Creating a default wave format and silent sound
        WaveFormat = new STFN.Core.Audio.Formats.WaveFormat(48000, 32, 2, "", STFN.Core.Audio.Formats.WaveFormat.WaveFormatEncodings.IeeeFloatingPoints);

        silentSound = STFN.Core.DSP.CreateSilence(ref this.WaveFormat, null, 3);

        // RETSPL values for DD65v2
        RetSplList = new SortedList<int, double>();
        RetSplList.Add(125, 30.5);
        RetSplList.Add(160, 25.5);
        RetSplList.Add(200, 21.5);
        RetSplList.Add(250, 17);
        RetSplList.Add(315, 14);
        RetSplList.Add(400, 10.5);
        RetSplList.Add(500, 8);
        RetSplList.Add(630, 6.5);
        RetSplList.Add(750, 5.5);
        RetSplList.Add(800, 5);
        RetSplList.Add(1000, 4.5);
        RetSplList.Add(1250, 3.5);
        RetSplList.Add(1500, 2.5);
        RetSplList.Add(1600, 2.5);
        RetSplList.Add(2000, 2.5);
        RetSplList.Add(2500, 2);
        RetSplList.Add(3000, 2);
        RetSplList.Add(3150, 3);
        RetSplList.Add(4000, 9.5);
        RetSplList.Add(5000, 15.5);
        RetSplList.Add(6000, 21);
        RetSplList.Add(6300, 21);
        RetSplList.Add(8000, 21);

        // Adding PTA calibration gain values
        var CurrentMixer = Globals.StfBase.SoundPlayer.GetMixer();
        //RightInfoLabel2.Text = CurrentMixer.GetParentTransducerSpecification.Name;
        //RightInfoLabel3.Text = CurrentMixer.GetParentTransducerSpecification.ParentAudioApiSettings.GetSelectedOutputDeviceName();

        if (CurrentMixer.GetParentTransducerSpecification.PtaCalibrationGain != null)
        {
            PureToneCalibrationList = CurrentMixer.GetParentTransducerSpecification.PtaCalibrationGain;
        }

    }

    private enum VisibilityTypes
    {
        MessageMode,
        TestMode,
        ResultMode
    }

    private void SetVisibility(VisibilityTypes visibilityType)
    {
        switch (visibilityType)
        {
            case VisibilityTypes.MessageMode:
                MessageLabel.IsVisible = false;
                MidResponseButton.IsVisible = false;
                MessageButton.IsVisible = true;
                ResultView.IsVisible = false;
                break;
            case VisibilityTypes.TestMode:
                MessageLabel.IsVisible = true;
                MidResponseButton.IsVisible = true;
                MessageButton.IsVisible = false;
                ResultView.IsVisible = false;
                break;

            case VisibilityTypes.ResultMode:
                MessageLabel.IsVisible = false;
                MidResponseButton.IsVisible = false;
                MessageButton.IsVisible = true;
                ResultView.IsVisible = true;
                break;

            default:
                throw new NotImplementedException("Unknown VisibilityType");
        }

    }

    public void StartTest()
    {

        // Returns emmediately if test is already started
        if (TestIsStarted || permanentlyStopped)
        {
            return;
        }

        TestIsStarted = true;

        // Goes into test mode and starts the next trial
        SetVisibility(VisibilityTypes.TestMode);

        // And into fullscreen mode
        SetFullScreenMode(true);

        // Create the timer before the first trial; completion/errors may stop it.
        try
        {
            StartTimer();
            StartNextTrial();
        }
        catch (Exception exception)
        {
            FailTest(exception);
        }
    }

    private void MessageButton_Clicked(object sender, EventArgs e)
    {

        if (CurrentTest.IsCompleted() == false)
        {
            StartTest();
        }
        else
        {
            // Stopping the testing loop
            TrialEndTimer.Stop();

            // Leaves the fullscreen mode
            SetFullScreenMode(false);

            // Shows test results
            SetVisibility(VisibilityTypes.ResultMode);
            ResultView.Text = CurrentTest.ResultSummary;

        }
    }

    public string GetResults()
    {
        return CurrentTest.GetResults(false);
    }

    private void StartNextTrial()
    {
        if (permanentlyStopped || !TestIsStarted) return;
        CurrentTrial = CurrentTest.GetNextTrial();

        // Checking if the test is completed, and messaging the user
        if (CurrentTrial == null)
        {
            if (CurrentTest.IsCompleted() == true)
            {
                // The test is completed
                SetVisibility(VisibilityTypes.MessageMode);

                // Stopping the trial end timer
                TrialEndTimer.Stop();
                TestIsStarted = false;
                Globals.StfBase.SoundPlayer.FadeOutPlayback();

                // Saving test result
                if (UoPta.SaveAudiogramDataToFile == true)
                {
                    CurrentTest.ExportAudiogramData();
                }

                if (ShowResultsInView == false)
                {
                    EventHandler<EventArgs> handler = Finished;
                    EventArgs e2 = new EventArgs();
                    if (handler != null)
                    {
                        handler(this, e2);
                    }
                    return;
                }

                switch (SharedSpeechTestObjects.GuiLanguage)
                {
                    case STFN.Core.Utils.EnumCollection.Languages.English:
                        MessageButton.Text = "This hearing test is now completed.\n\nClick here to continue!";
                        LevelLabel.Text = "";
                        RatingLabel.Text = "";
                        break;
                    case STFN.Core.Utils.EnumCollection.Languages.Swedish:
                        MessageButton.Text = "Detta test är nu klart.\n\nTryck här för att gå vidare!";
                        LevelLabel.Text = "";
                        RatingLabel.Text = "";
                        break;
                    default:
                        // Using English as default
                        MessageButton.Text = "This hearing test is now completed.\n\nClick here to continue!";
                        LevelLabel.Text = "";
                        RatingLabel.Text = "";
                        break;
                }
            }
            else
            {
                // Something else has caused the CurrentTrial to be null. Why? Message the user...
            }
            // Returning from the method early
            return;
        }

        // Showing the presented level
        //LevelLabel.Text = CurrentTrial.ToneLevel.ToString() + " dB HL";
        //RatingLabel.Text = "";

        int frequency = CurrentTrial.ParentSubTest.Frequency;
        if (!RetSplList.TryGetValue(frequency, out var retspl))
            throw new InvalidOperationException($"Missing RETSPL value for {frequency} Hz.");
        if (!PureToneCalibrationList.TryGetValue(frequency, out var calibration) || !double.IsFinite(calibration))
            throw new InvalidOperationException($"Missing or invalid pure-tone calibration for {frequency} Hz.");
        double RetSplCorrectedLevel_FS = STFN.Core.DSP.Standard_dBSPL_To_dBFS(
            CurrentTrial.ToneLevel + retspl + calibration);
        if (!double.IsFinite(RetSplCorrectedLevel_FS) || RetSplCorrectedLevel_FS > -3)
            throw new InvalidOperationException($"The configured output cannot present {CurrentTrial.ToneLevel} dB HL at {frequency} Hz without clipping. Check calibration and output settings.");

        // Creating the sine
        STFN.Core.Audio.Sound SineSound = STFN.Core.DSP.CreateSineWave(ref this.WaveFormat, 1, CurrentTrial.ParentSubTest.Frequency, (decimal)RetSplCorrectedLevel_FS, STFN.Core.DSP.SoundDataUnit.dB,
            CurrentTrial.ToneDuration.TotalSeconds, TimeUnits.seconds, 0, true);
        STFN.Core.DSP.Fade(ref SineSound, null, 0, 1, 0, (int)(WaveFormat.SampleRate * 0.1), STFN.Core.DSP.FadeSlopeType.Linear);
        STFN.Core.DSP.Fade(ref SineSound, 0, null, 1, (int)(-WaveFormat.SampleRate * 0.1), null, STFN.Core.DSP.FadeSlopeType.Linear);

        // Padding the sine with the initial silence
        SineSound.ZeroPad(CurrentTrial.ToneOnsetTime.TotalSeconds, null, false);

        // Creating a stereo sound
        STFN.Core.Audio.Sound TrialSound = new STFN.Core.Audio.Sound(WaveFormat);

        switch (CurrentTrial.ParentSubTest.Side)
        {
            case STFN.Core.Utils.EnumCollection.Sides.Left:

                TrialSound.WaveData.set_SampleData(1, SineSound.WaveData.get_SampleData(1));
                // Leaving the other channel empty

                break;
            case STFN.Core.Utils.EnumCollection.Sides.Right:

                TrialSound.WaveData.set_SampleData(2, SineSound.WaveData.get_SampleData(1));
                // Leaving the other channel empty

                break;
            default:
                throw new Exception("Invalid value for sides");
                //break;
        }

        // Storing the sound in the trial (probabilty unecessary)
        CurrentTrial.MixedSound = TrialSound;

        // Storing the trial start time
        CurrentTrial.TrialStartTime = DateTime.Now;

        // Starting the playback
        if (!Globals.StfBase.SoundPlayer.SwapOutputSounds(ref TrialSound))
            throw new InvalidOperationException("The sound player could not play the audiometry stimulus.");

    }


    private void OnMidResponseButtonPressed(object sender, EventArgs e)
    {
        if (!TestIsStarted || permanentlyStopped) return;
        MidResponseButton.Background = PressedBrush;

        if (CurrentTrial != null)
        {
            CurrentTrial.ResponseOnsetTime = DateTime.Now - CurrentTrial.TrialStartTime;
        }

    }


    private void OnMidResponseButtonReleased(object sender, EventArgs e)
    {
        if (!TestIsStarted || permanentlyStopped) return;
        MidResponseButton.Background = NonPressedBrush;

        if (CurrentTrial != null)
        {
            CurrentTrial.ResponseOffsetTime = DateTime.Now - CurrentTrial.TrialStartTime;

            //RatingLabel.Text = CurrentTrial.Result.ToString();

        }

    }


    private void StartTimer()
    {

        // Create and setup timer
        TrialEndTimer = Application.Current.Dispatcher.CreateTimer();
        TrialEndTimer.Interval = TimeSpan.FromMilliseconds(50); // Checking if trial has ended every 50th ms
        TrialEndTimer.Tick += TrialEndTimer_Tick;
        TrialEndTimer.IsRepeating = true;
        TrialEndTimer.Start();

    }

    void TrialEndTimer_Tick(object? sender, EventArgs e)
    {
        if (!TestIsStarted || permanentlyStopped) return;
        try
        {
            if (CurrentTrial != null)
            {
                if (CurrentTrial.Result == UoPta.PtaResults.TruePositive)
                {
                    // We should wait PostTruePositiveResponseDuration seconds after a true positive response before next trial is started
                    if (DateTime.Now > CurrentTrial.TrialStartTime + CurrentTrial.ResponseOffsetTime + TimeSpan.FromSeconds(CurrentTest.PostTruePositiveResponseDuration))
                    {
                        //Saving trial data to file
                        if (UoPta.LogTrialData == true) { CurrentTrial.ExportPtaTrialData(); }

                        // Setting CurrentTrial to null and starting next trial
                        CurrentTrial = null;
                        StartNextTrial();
                    }
                }
                else
                {
                    // If we don't have a true positive response, we should wait PostToneDuration seconds after the end of the signal before next trial is started
                    if (DateTime.Now > CurrentTrial.TrialStartTime + CurrentTrial.ToneOffsetTime + TimeSpan.FromSeconds(CurrentTest.PostToneDuration))
                    {
                        //Saving trial data to file
                        if (UoPta.LogTrialData == true) { CurrentTrial.ExportPtaTrialData(); }

                        // Setting CurrentTrial to null and starting next trial
                        CurrentTrial = null;
                        StartNextTrial();
                    }
                }
            }

        }
        catch (Exception exception)
        {
            FailTest(exception);
        }
    }

    private void SetFullScreenMode(bool Fullscreen)
    {

        if (Fullscreen == true)
        {
            EventHandler<EventArgs> handler = EnterFullScreenMode;
            EventArgs e = new EventArgs();
            if (handler != null)
            {
                handler(this, e);
            }
        }
        else
        {
            EventHandler<EventArgs> handler = ExitFullScreenMode;
            EventArgs e = new EventArgs();
            if (handler != null)
            {
                handler(this, e);
            }
        }
    }


}

public sealed class UoAudErrorEventArgs(string message, Exception exception) : EventArgs
{
    public string Message { get; } = message;
    public Exception Exception { get; } = exception;
}
