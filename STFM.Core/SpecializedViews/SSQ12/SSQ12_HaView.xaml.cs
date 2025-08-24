// License
// This project is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
// Commercial use requires a separate commercial license.See the `LICENSE` file for details.
// 
// Copyright (c) 2025 Erik Witte

using STFN.Core;
namespace STFM.SpecializedViews.SSQ12;

public partial class SSQ12_HaView : ContentView
{

    public SSQ12_HaView()
	{
		InitializeComponent();

        if (SSQ12_MainView.MinimalVersion == false)
        {
            // Hiding the HA-time GUI objects
            HA_UseTimeQuestion_Label.IsVisible = false;
            HA_UseTime_EditorBorder.IsVisible = false;
        }

        switch (SharedSpeechTestObjects.GuiLanguage)
        {

            case STFN.Core.Utils.EnumCollection.Languages.Swedish:

                HA_UseQuestion_Label.Text = "Anv�nder du h�rapparat?";
                HA_DoUse_RadioButton.Content = "Jag anv�nder h�rapparat";
                HA_NotUse_RadioButton.Content = "Jag anv�nder inte h�rapparat";

                HA_Use_Label.Text = "Ange ett av alternativen";
                HA_DoUseLeft_RadioButton.Content = "Jag anv�nder h�rapparat till v�nster �ra";
                HA_DoUseRight_RadioButton.Content = "Jag anv�nder h�rapparat till h�ger �ra";
                HA_DoUseBoth_RadioButton.Content = "Jag anv�nder h�rapparat till b�da �ronen";

                HA_UseTimeQuestion_Label.Text = "Hur l�nge har du anv�nt din h�rapparat?";
                HA_UseTime_Editor.Text = "";

                break;
            default:
                // Using English as default

                HA_UseQuestion_Label.Text = "";
                HA_DoUse_RadioButton.Content = "";
                HA_NotUse_RadioButton.Content = "";

                HA_Use_Label.Text = "";
                HA_DoUseLeft_RadioButton.Content = "";
                HA_DoUseRight_RadioButton.Content = "";
                HA_DoUseBoth_RadioButton.Content = "";

                HA_UseTimeQuestion_Label.Text = "";
                HA_UseTime_Editor.Text = "";

                break;
        }

    }

    private void HA_DoUse_RadioButton_CheckedChanged(object sender, CheckedChangedEventArgs e)
    {
        if (e.Value)
        {
            HA_Details_StackLayout.IsVisible = true;

            if (SSQ12_MainView.MinimalVersion)
            {
                HA_UseTimeQuestion_Label.IsVisible = false;
                HA_UseTime_EditorBorder.IsVisible = false;
            }

        }
    }

    private void HA_NotUse_RadioButton_CheckedChanged(object sender, CheckedChangedEventArgs e)
    {
        if (e.Value)
        {
            HA_Details_StackLayout.IsVisible = false;
        }
    }

    /// <summary>
    /// Checks if side of hearing aid was selected if hearing aid use is specified. However, everything else is passed through, since this question is not mandatory.
    /// </summary>
    /// <returns></returns>
    public async Task<bool> HasResponse()
    {

        if (HA_DoUse_RadioButton.IsChecked)
        {

            if (HA_DoUseLeft_RadioButton.IsChecked | HA_DoUseRight_RadioButton.IsChecked | HA_DoUseBoth_RadioButton.IsChecked)
            {
                return true;
            }
            else
            {
                // Asks the user to specify left / righ / or both sides
                switch (SharedSpeechTestObjects.GuiLanguage)
                {
                    case STFN.Core.Utils.EnumCollection.Languages.Swedish:
                        await Messager.MsgBoxAsync("V�nligen ange om du anv�nder h�rapparat p� v�nster, h�ger eller b�da �ronen!", Messager.MsgBoxStyle.Information, "Ange h�rapparatsida!");

                        break;
                    default:
                        // Using English as default
                        await Messager.MsgBoxAsync("Please specify if you use a hearing aid on left, right or both ears.", Messager.MsgBoxStyle.Information, "Specify hearing aid side!");

                        break;
                }
                return false;
            }
        }
        else
        {
            return true;
        }

    }

    /// <summary>
    /// Gets a string formatted result of the hearing aid use questions. Note that HasResponse should have been evaluated before calling this function.
    /// </summary>
    /// <returns></returns>
    public string GetResultString()
    {

        List<string> ReturnValueList = new List<string>();

        switch (SharedSpeechTestObjects.GuiLanguage)
        {
            case STFN.Core.Utils.EnumCollection.Languages.Swedish:

                if (SSQ12_MainView.MinimalVersion == false)
                {
                    ReturnValueList.Add("H�RAPPARATANV�NDNING");
                }

                // Adding hearing aid use
                if (HA_DoUse_RadioButton.IsChecked)
                {
                    if (HA_DoUseLeft_RadioButton.IsChecked)
                    {
                        ReturnValueList.Add("Anv�nder h�rapparat till v�nster �ra.");
                    }
                    else if (HA_DoUseRight_RadioButton.IsChecked)
                    {
                        ReturnValueList.Add("Anv�nder h�rapparat till h�ger �ra.");
                    }
                    else if (HA_DoUseRight_RadioButton.IsChecked)
                    {
                        ReturnValueList.Add("Anv�nder h�rapparat till b�da �ronen.");
                    }

                    // Adding time comment
                    if (HA_UseTime_Editor.Text.Trim() != "")
                    {
                        ReturnValueList.Add("TID: " + HA_UseTime_Editor.Text);
                    }

                }
                else if (HA_NotUse_RadioButton.IsChecked)
                {
                    ReturnValueList.Add("Anv�nder inte h�rapparat.");
                }
                else
                {
                    ReturnValueList.Add("H�rapparatanv�ndning ej angiven.");
                }

                break;
            default:

                if (SSQ12_MainView.MinimalVersion == false)
                {
                    ReturnValueList.Add("HEARING AID USE");
                }

                // Adding hearing aid use
                if (HA_DoUse_RadioButton.IsChecked)
                {
                    if (HA_DoUseLeft_RadioButton.IsChecked)
                    {
                        ReturnValueList.Add("Hearing aid on left ear.");
                    }
                    else if (HA_DoUseRight_RadioButton.IsChecked)
                    {
                        ReturnValueList.Add("Hearing aid on right ear.");
                    }
                    else if (HA_DoUseRight_RadioButton.IsChecked)
                    {
                        ReturnValueList.Add("Hearing aids on both ears.");
                    }

                    // Adding time comment
                    if (HA_UseTime_Editor.Text.Trim() != "")
                    {
                        ReturnValueList.Add("TIME: " + HA_UseTime_Editor.Text);
                    }

                }
                else if (HA_NotUse_RadioButton.IsChecked)
                {
                    ReturnValueList.Add("Not using hearing aids.");
                }
                else
                {
                    ReturnValueList.Add("H�rapparatanv�ndning ej angiven.");
                }
                break;
        }

        return string.Join("\n", ReturnValueList);

    }

    public enum BoolWithNotSet
    {
        Yes,
        No,
        NotSet
    }

    /// <summary>
    /// Gets a BoolWithNotSet specifying the response concerning hearing aid use. Note that HasResponse should have been evaluated before calling this function.
    /// </summary>
    /// <returns></returns>
    public BoolWithNotSet UsingHearingAids()
    {

        // Adding hearing aid use
        if (HA_DoUse_RadioButton.IsChecked)
        {
            return BoolWithNotSet.Yes;
        } else if (HA_NotUse_RadioButton.IsChecked)
        {
            return BoolWithNotSet.No;
        } else
        {
            return BoolWithNotSet.NotSet;
        }

    }

}