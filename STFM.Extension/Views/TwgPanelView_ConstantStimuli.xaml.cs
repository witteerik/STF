namespace STFM.Extension.Views;
using STFN.Core;
using STFM.Views;

public partial class TwgPanelView_ConstantStimuli : ContentView
{
    public TwgPanelView_ConstantStimuli()
    {
        InitializeComponent();

        // Assign the custom drawable to the GraphicsView
        SnrView.Drawable = new TestResultsDiagram(SnrView);
        TestResultsDiagram MySnrDiagram = (TestResultsDiagram)SnrView.Drawable;
        MySnrDiagram.SetSizeModificationStrategy(PlotBase.SizeModificationStrategies.Horizontal);
        //MySnrDiagram.TransitionHeightRatio = 0.86f;
        //MySnrDiagram.Background = Colors.DarkSlateGray;

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
        List<float> X_PnrArray = [0];

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

        MySnrDiagram.SetXlim(X_PnrArray.Min() - 2.5f, X_PnrArray.Max() + 2.5f);
        MySnrDiagram.SetYlim(0, 105);

        MySnrDiagram.UpdateLayout();

        // Force redraw on size change
        SnrView.SizeChanged += (s, e) => SnrView.Invalidate();

        switch (SharedSpeechTestObjects.GuiLanguage)
        {
            case STFN.Core.Utils.EnumCollection.Languages.Swedish:

                TwgNameLabel.Text = "Grupp:";
                ReferenceLevelNameLabel.Text = "Referensnivå:";
                TrialNumberNameLabel.Text = "Försök nummer:";

                SnrGridLabelY.Text = "% Korrekt";
                SnrGridLabelX.Text = "PNR (dB)";

                break;
            default:

                SnrGridLabelY.Text = "% Score";
                SnrGridLabelX.Text = "PNR (dB)";

                ReferenceLevelNameLabel.Text = "Reference level:";
                TrialNumberNameLabel.Text = "Trial number:";

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

    public string ReferenceLevel
    {
        get => ReferenceLevelValueLabel.Text;
        set => ReferenceLevelValueLabel.Text = value;
    }

    public string TrialNumber
    {
        get => TrialNumberValueLabel.Text;
        set => TrialNumberValueLabel.Text = value;
    }


}