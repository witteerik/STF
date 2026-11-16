// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using Microsoft.Maui.Graphics;
using STFM.Views;

namespace STFM.Extension.Screening.Reporting;

/// <summary>Uses STF's existing log-frequency plotting/coordinate system for screening audiometry.</summary>
public sealed class ScreeningAudiogram : PlotBase, IDrawable
{
    private readonly AudiometryPoint[] points;

    public ScreeningAudiogram(AudiometryResult result)
    {
        points = result.Points.GroupBy(p => (p.Ear, p.FrequencyHz)).Select(group => group.Last()).ToArray();
        XlimMin = 125; XlimMax = 8000; Xlog = true; XlogBase = 2;
        YlimMin = -10; YlimMax = Math.Max(110, result.MaximumLevelDbHl + 10); Yreversed = true;
        PlotAreaRelativeMarginLeft = .10f; PlotAreaRelativeMarginRight = .06f;
        PlotAreaRelativeMarginTop = .09f; PlotAreaRelativeMarginBottom = .05f;
        PlotAreaBorderColor = Color.FromArgb("#78868F");
        GridLineColor = Color.FromArgb("#D0D8DD"); DashedGridLineColor = Color.FromArgb("#E4E9ED");
        XaxisGridLinePositions = [125, 250, 500, 1000, 2000, 4000, 8000];
        XaxisDashedGridLinePositions = [750, 1500, 3000, 6000];
        XaxisTextPositions = XaxisGridLinePositions;
        XaxisTextValues = ["125", "250", "500", "1k", "2k", "4k", "8k"];
        XaxisDrawTop = true; XaxisDrawBottom = false; XaxisTextSize = .65f;
        YaxisGridLinePositions = Enumerable.Range(0, (int)YlimMax / 10 + 1).Select(v => (float)(v * 10)).ToList();
        YaxisTextPositions = Enumerable.Range(0, (int)YlimMax / 20 + 1).Select(v => (float)(v * 20)).ToList();
        YaxisTextValues = YaxisTextPositions.Select(v => v.ToString("0", System.Globalization.CultureInfo.InvariantCulture)).ToArray();
        YaxisTextSize = .65f;
    }

    void IDrawable.Draw(ICanvas canvas, RectF dirtyRect) => DrawScreening(canvas, dirtyRect);

    public void DrawScreening(ICanvas canvas, RectF bounds)
    {
        canvas.SaveState();
        canvas.Translate(bounds.X, bounds.Y);
        var local = new RectF(0, 0, bounds.Width, bounds.Height);
        base.Draw(canvas, local);
        float left = PlotAreaRelativeMarginLeft * bounds.Width;
        float top = PlotAreaRelativeMarginTop * bounds.Height;
        float width = bounds.Width * (1 - PlotAreaRelativeMarginLeft - PlotAreaRelativeMarginRight);
        float height = bounds.Height * (1 - PlotAreaRelativeMarginTop - PlotAreaRelativeMarginBottom);
        float bottom = top + height;
        canvas.StrokeDashPattern = null;
        canvas.StrokeSize = 1.5f;
        // Do not connect across missing, uncertain or no-response points. Screening
        // bounds are marked explicitly instead of implying exact threshold curves.
        foreach (var point in points)
        {
            if (point.LevelDbHl is not { } level || point.FrequencyHz < XlimMin || point.FrequencyHz > XlimMax ||
                point.Status is AudiometryPointStatus.NotMeasured or AudiometryPointStatus.Uncertain)
                continue;
            bool right = point.Ear.Equals("Right", StringComparison.OrdinalIgnoreCase);
            var color = right ? Color.FromArgb("#B53838") : Color.FromArgb("#2466A5");
            canvas.StrokeColor = color;
            canvas.FontColor = color;
            float x = XValueToCoordinate(point.FrequencyHz, left, width, XlimMax - XlimMin);
            float y = YValueToCoordinate(level, top, bottom, height, YlimMax - YlimMin);
            if (right)
                canvas.DrawEllipse(x - 3.6f, y - 3.6f, 7.2f, 7.2f);
            else
            {
                canvas.DrawLine(x - 3.6f, y - 3.6f, x + 3.6f, y + 3.6f);
                canvas.DrawLine(x - 3.6f, y + 3.6f, x + 3.6f, y - 3.6f);
            }
            if (point.Status == AudiometryPointStatus.NoResponse)
            {
                float direction = right ? 1 : -1;
                float endX = x + direction * 9;
                float endY = y + 11;
                canvas.DrawLine(x + direction * 2, y + 2, endX, endY);
                canvas.DrawLine(endX, endY, endX - direction * 5, endY - 1);
                canvas.DrawLine(endX, endY, endX, endY - 5);
            }
            else if (point.AtScreeningLevel)
            {
                canvas.FontSize = 8;
                canvas.DrawString("≤", x + 5, y - 4, 12, 12, HorizontalAlignment.Left, VerticalAlignment.Top);
            }
        }
        canvas.RestoreState();
    }
}
