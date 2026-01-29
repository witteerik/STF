namespace STFM.Extension.Views;
using STFN.Core;
using STFM.Views;

public partial class TwgPanelView : ContentView
{
    public TwgPanelView()
    {
        InitializeComponent();

        // Assign the custom drawable to the GraphicsView
        SnrView.Drawable = new TestResultsDiagram(SnrView);
        TestResultsDiagram MySnrDiagram = (TestResultsDiagram)SnrView.Drawable;
        MySnrDiagram.SetSizeModificationStrategy(PlotBase.SizeModificationStrategies.Horizontal);
        MySnrDiagram.SetXlim(0.5f, 1.5f);
        MySnrDiagram.SetYlim(-10, 10);
        //MySnrDiagram.TransitionHeightRatio = 0.86f;
        //MySnrDiagram.Background = Colors.DarkSlateGray;
        MySnrDiagram.UpdateLayout();

        // Force redraw on size change
        SnrView.SizeChanged += (s, e) => SnrView.Invalidate();

        switch (SharedSpeechTestObjects.GuiLanguage)
        {
            case STFN.Core.Utils.EnumCollection.Languages.Swedish:

                TwgNameLabel.Text = "Grupp:";
                TargetScoreNameLabel.Text = "Mål, andel korrekt:";
                ReferenceLevelNameLabel.Text = "Referensnivå:";
                AdaptiveLevelNameLabel.Text = "Adaptiv nivå:";
                TrialNumberNameLabel.Text = "Försök nummer:";
                FinalResultNameLabel.Text = "Hörtröskel:";

                SnrGridLabelY.Text = "PNR (dB)";
                SnrGridLabelX.Text = "Försök nummer";

                break;
            default:

                SnrGridLabelY.Text = "PNR (dB)";
                SnrGridLabelX.Text = "Test trial";

                TargetScoreNameLabel.Text = "Target score:";
                ReferenceLevelNameLabel.Text = "Reference level:";
                AdaptiveLevelNameLabel.Text = "Adaptive level:";
                TrialNumberNameLabel.Text = "Trial number:";
                FinalResultNameLabel.Text = "SRT";

                break;
        }


    }


    // If SnrView uses an IDrawable, expose a helper for that too:
    public void SetSnrDrawable(IDrawable drawable)
    {
        SnrView.Drawable = drawable;
        SnrView.Invalidate();
    }

    public GraphicsView GetSnrView()
    {
            return SnrView;
    }

    public string Twg
    {
        get => TwgValueLabel.Text;
        set => TwgValueLabel.Text = value;
    }

    public string TargetScore
    {
        get => TargetScoreValueLabel.Text;
        set => TargetScoreValueLabel.Text = value;
    }

    public string ReferenceLevel
    {
        get => ReferenceLevelValueLabel.Text;
        set => ReferenceLevelValueLabel.Text = value;
    }

    public string AdaptiveLevel
    {
        get => AdaptiveLevelValueLabel.Text;
        set => AdaptiveLevelValueLabel.Text = value;
    }

    public string TrialNumber
    {
        get => TrialNumberValueLabel.Text;
        set => TrialNumberValueLabel.Text = value;
    }

    public string FinalResult
    {
        get => FinalResultValueLabel.Text;
        set => FinalResultValueLabel.Text = value;
    }


}