// License
// This project is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
// Commercial use requires a separate commercial license.See the `LICENSE` file for details.
// 
// Copyright (c) 2025 Erik Witte


namespace STFM.Pages;

public class TestResultPage : ContentPage
{

	STFM.Views.TestResultsView CurrentTestResultsView;

    public TestResultPage(ref STFM.Views.TestResultsView currentTestResultsView)

    {

        CurrentTestResultsView = currentTestResultsView;

        Content = CurrentTestResultsView;

        // Prevent closing: https://learn.microsoft.com/en-us/answers/questions/1336207/how-to-remove-close-and-maximize-button-for-a-maui

    }
}