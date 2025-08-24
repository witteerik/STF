// License
// This project is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
// Commercial use requires a separate commercial license.See the `LICENSE` file for details.
// 
// Copyright (c) 2025 Erik Witte

using STFN.Core;
using STFM.Views;

namespace STFM.Extension.Views;

public class TestResultsView_Text : TestResultsView
{

    ScrollView MainScrollView;
    Label ResultsLabel;

    public TestResultsView_Text()
	{

        MainScrollView = new Microsoft.Maui.Controls.ScrollView();
        ResultsLabel = new Label();
        MainScrollView.Content = ResultsLabel;

        //MainTestResultsGrid = new Grid
        //{
        //    RowDefinitions = { new RowDefinition { Height = new GridLength(1, GridUnitType.Star) } },
        //    ColumnDefinitions = { new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) } }
        //};

        Content = MainScrollView;
    }

    public override void SetGuiLayoutState(SpeechTestView.GuiLayoutStates currentTestPlayState)
    {
        // Ignores any call
    }

    public override void ShowTestResults(string results)
    {
        ResultsLabel.Text = results;
     }

    public override void ShowTestResults(SpeechTest speechTest)
    {

    }

    public override void UpdateStartButtonText(string text) 
    {
        // Ignores this
    }

}