// License
// This project is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
// Commercial use requires a separate commercial license.See the `LICENSE` file for details.
// 
// Copyright (c) 2025 Erik Witte


using STFN.Core;
using STFM.Views;

namespace STFM.Extension.Views;

public partial class ResponseView_TSFC : ResponseView
{

    TSFC_Triangle CurrentTSFC_Triangle = null;

    private IDispatcherTimer HideAllTimer;


    bool isPressed = false;
    public ResponseView_TSFC()
	{
		InitializeComponent();

        // Assign the custom drawable to the GraphicsView
        TSFC_TriangleView.Drawable = new TSFC_Triangle(TSFC_TriangleView);
        
        // Referencing the TSFC_Triangle so that it gets directly accessible in code
        CurrentTSFC_Triangle = (TSFC_Triangle)TSFC_TriangleView.Drawable;
        //CurrentTSFC_Triangle.Background = Color.FromArgb("#2F2F2F");

        TSFC_TriangleView.StartInteraction += OnStartInteraction;
        TSFC_TriangleView.DragInteraction += OnDragInteraction;
        TSFC_TriangleView.EndInteraction += OnEndInteraction;

        // Cascades loaded and any resizing to the CurrentTSFC_Triangle
        TSFC_TriangleView.Loaded += CascadeResizeToTriangle;
        TSFC_TriangleView.SizeChanged += CascadeResizeToTriangle;

        // Adds a handler to rezise the buttons
        this.SizeChanged += ResizeButtons;

        // Setting unpressed button colors
        SetPressedState(false);

        // Creating a hide-all timer
        HideAllTimer = Microsoft.Maui.Controls.Application.Current.Dispatcher.CreateTimer();
        HideAllTimer.Interval = TimeSpan.FromMilliseconds(300);
        HideAllTimer.Tick += ClearLayout;
        HideAllTimer.IsRepeating = false;

    }

    private void ResizeButtons(object sender, EventArgs e)
    {

        double ButtonDiameter = this.Height * 0.3;

        LeftButton.HeightRequest = ButtonDiameter;
        LeftButton.WidthRequest = ButtonDiameter;
        LeftButton.CornerRadius = (int)ButtonDiameter / 2;

        RightButton.HeightRequest = ButtonDiameter;
        RightButton.WidthRequest = ButtonDiameter;
        RightButton.CornerRadius = (int)ButtonDiameter / 2;

    }

    private void CascadeResizeToTriangle(object sender, EventArgs e)
    {
        // Updating the cached size of the CurrentTSFC_Triangle
        var width = TSFC_TriangleView.Width;
        var height = TSFC_TriangleView.Height;
        CurrentTSFC_Triangle.ViewportSize = new SizeF((float)width, (float)height);

        // Forcíng redraw on size change
        TSFC_TriangleView.Invalidate();
    }


    private void OnStartInteraction(object sender, TouchEventArgs e)
    {
        // User has started touching
        PointF touchPoint = new PointF(e.Touches.First().X, e.Touches.First().Y);

        if (CurrentTSFC_Triangle != null)
        {
            CurrentTSFC_Triangle.StartInteraction(touchPoint);
        }

    }

    private void OnDragInteraction(object sender, TouchEventArgs e)
    {
        // Tracking the movement
        PointF touchPoint= new PointF( e.Touches.First().X, e.Touches.First().Y);

        if (CurrentTSFC_Triangle != null)
        {
            CurrentTSFC_Triangle.DragInteraction(touchPoint);
        }

    }

    private void OnEndInteraction(object sender, TouchEventArgs e)
    {
        // Touch finished

        PointF touchPoint = new PointF(e.Touches.First().X, e.Touches.First().Y);

        if (CurrentTSFC_Triangle != null)
        {
            CurrentTSFC_Triangle.EndInteraction(touchPoint);
        }

    }

    private void OnButtonPressed(object sender, EventArgs e)
    {
        SetPressedState(true);
    }

    private void OnButtonReleased(object sender, EventArgs e)
    {
        SetPressedState(false);
    }

    void SetPressedState(bool isPressed)
    {
        var background = isPressed ? Colors.LightBlue : Colors.LightGray;
        var textColor = isPressed ? Colors.White : Color.FromArgb("#2F2F2F");

        LeftButton.BackgroundColor = background;
        LeftButton.TextColor = textColor;

        RightButton.BackgroundColor = background;
        RightButton.TextColor = textColor;
    }

    private void OnButtonClicked(object sender, EventArgs e)
    {

        // Initiating test response 

        // Getting the coordinates
        ReportResult();

    }

    private readonly SemaphoreSlim _reportResultLock = new SemaphoreSlim(1, 1);
    private bool responseIsGiven = false;

    private async void ReportResult()
    {
        // Prevent double-entry
        if (!await _reportResultLock.WaitAsync(0))
            return;

        try
        {

            // Returns is a response has already been given 
            if (responseIsGiven) {return; }
            responseIsGiven = true;

            // Disable response buttons
            LeftButton.IsEnabled = false;
            RightButton.IsEnabled = false;

            // Getting the barycentric coordinates of the point in relation to all response alternatives
            SortedList<string, double> responseBox = new SortedList<string, double>();

            if (CurrentTSFC_Triangle != null)
            {
                responseBox.Add(TestWordLabel_Left.Text, CurrentTSFC_Triangle.PointLocations[0]);
                responseBox.Add(TestWordLabel_Right.Text, CurrentTSFC_Triangle.PointLocations[1]);
                responseBox.Add(TestWordLabel_Bottom.Text, CurrentTSFC_Triangle.PointLocations[2]);
            }

            // Storing the raw response
            SpeechTestInputEventArgs args = new SpeechTestInputEventArgs
            {
                LinguisticResponseTime = DateTime.Now,
                Box = responseBox
            };

            // Raise the response event
            OnResponseGiven(args);

            // Clear UI
            ClearLayout();
        }
        finally
        {
            // Always release lock
            _reportResultLock.Release();
        }
    }


    public override void AddSourceAlternatives(VisualizedSoundSource[] soundSources)
    {
        //throw new NotImplementedException();
    }


    public override void HideVisualCue()
    {
        // throw new NotImplementedException();
    }

    public override void InitializeNewTrial()
    {
        StopAllTimers();
        ClearLayout();
    }



    public override void ResponseTimesOut()
    {
        // throw new NotImplementedException();
    }

    public override void ShowMessage(string Message)
    {
        // throw new NotImplementedException();
    }

    public override void ShowResponseAlternativePositions(List<List<SpeechTestResponseAlternative>> ResponseAlternatives)
    {
        //  throw new NotImplementedException();
    }

    public override void ShowResponseAlternatives(List<List<SpeechTestResponseAlternative>> ResponseAlternatives)
    {

        if (ResponseAlternatives.Count > 1)
        {
            throw new ArgumentException("ShowResponseAlternatives is not yet implemented for multidimensional sets of response alternatives");
        }

        List<SpeechTestResponseAlternative> localResponseAlternatives = ResponseAlternatives[0];

        TestWordLabel_Left.Text = localResponseAlternatives[0].Spelling;
        TestWordLabel_Right.Text = localResponseAlternatives[1].Spelling;
        TestWordLabel_Bottom.Text = localResponseAlternatives[2].Spelling;

        // Reset point to center
        CurrentTSFC_Triangle.ResetPointToCenter();

        // Resets responseIsGiven 
        responseIsGiven = false;

        // Enables the response buttons
        LeftButton.IsEnabled = true;
        RightButton.IsEnabled = true;

    }

    public override void ShowVisualCue()
    {
        //  throw new NotImplementedException();
    }

    public override void StopAllTimers()
    {
        HideAllTimer.Stop();
    }

    public async override void UpdateTestFormProgressbar(int Value, int Maximum, int Minimum)
    {
        if (PtcProgressBar != null)
        {
            ProgressFrame.IsVisible = true;
            double range = Maximum - Minimum;
            double progressProp = Value / range;
            await PtcProgressBar.ProgressTo(progressProp, 50, Easing.Linear);
        }
    }

    private void ClearLayout(object sender, EventArgs e)
    {
        ClearLayout();
    }

    private void ClearLayout()
    {
        //Clear things
        TestWordLabel_Left.Text = "";
        TestWordLabel_Right.Text = "";
        TestWordLabel_Bottom.Text = "";

        // Hiding the circle
        CurrentTSFC_Triangle.CircleIsVisible = false;

    }

    public override void HideAllItems()
    {
        ClearLayout();

        LeftButton.IsVisible = false;
        RightButton.IsVisible = false;

        ProgressFrame.IsVisible = false;
    }


}