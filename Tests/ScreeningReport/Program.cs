global using Microsoft.Maui.Graphics;
using STFM.Extension.Screening;
using STFM.Extension.Screening.Reporting;
using SkiaSharp;
using Microsoft.Maui.Graphics.Skia;
var output = args.Length == 0
    ? Path.Combine(Path.GetTempPath(), "ResonicScreeningReportChecks", Guid.NewGuid().ToString("N"))
    : Path.GetFullPath(args[0]);
Directory.CreateDirectory(output);
var photo = Path.Combine(output,"demo.jpg");
using(var bitmap = new SKBitmap(720,720))
{
 using var c=new SKCanvas(bitmap); c.Clear(SKColors.LightGray);
 using var p=new SKPaint{Color=SKColors.SlateGray}; c.DrawCircle(360,360,250,p);
 using var f=new SKFont(SKTypeface.Default,30); p.Color=SKColors.White;
 c.DrawText("DEMO IMAGE",360,360,SKTextAlign.Center,f,p);
 using var im=SKImage.FromBitmap(bitmap); using var d=im.Encode(SKEncodedImageFormat.Jpeg,95); using var s=File.Create(photo); d.SaveTo(s);
}
var now = new DateTimeOffset(2026,11,14,14,15,0,TimeSpan.FromHours(1));
var exam = new ScreeningExamination {
 Reference="DEMO – INGA PATIENTDATA",StartedAt=now,CompletedAt=now.AddMinutes(10),Stage=ScreeningStage.Report,
 Otoscopy=new(){LeftImagePath=photo,RightImagePath=photo,LeftCapturedAt=now,RightCapturedAt=now.AddMinutes(1),Notes="Demonstration av rapportens layout. Bilderna är syntetiska testbilder och innehåller inga patientdata."},
 Ssq12=new(now.AddMinutes(4),Enumerable.Range(1,12).Select(i=>new Ssq12Answer(i,"Fråga "+i,i==8?null:i%11,i==8)).ToArray(),false,false,false),
 QuickSip=new(now.AddMinutes(6),23,30,30,2,23d/30,"DD65 v2 – kalibrerade hörtelefoner","Lista 1"),
 Audiometry=new(now.AddMinutes(10),new[]{
 new AudiometryPoint("Right",500,25,AudiometryPointStatus.Reached,false,true),
 new AudiometryPoint("Right",1000,25,AudiometryPointStatus.Reached,false,true),
 new AudiometryPoint("Right",2000,35,AudiometryPointStatus.Reached,false,false),
 new AudiometryPoint("Right",4000,50,AudiometryPointStatus.Reached,false,false),
 new AudiometryPoint("Right",6000,80,AudiometryPointStatus.NoResponse,false,false),
 new AudiometryPoint("Right",8000,60,AudiometryPointStatus.Uncertain,false,false),
 new AudiometryPoint("Left",500,25,AudiometryPointStatus.Reached,false,true),
 new AudiometryPoint("Left",1000,25,AudiometryPointStatus.Reached,false,true),
 new AudiometryPoint("Left",2000,30,AudiometryPointStatus.Reached,false,false),
 new AudiometryPoint("Left",4000,55,AudiometryPointStatus.Reached,false,false),
 new AudiometryPoint("Left",6000,70,AudiometryPointStatus.Reached,false,false),
 new AudiometryPoint("Left",8000,null,AudiometryPointStatus.NotMeasured,false,false),
 new AudiometryPoint("Left",1000,30,AudiometryPointStatus.Reached,true,false)
 },"DD65 v2","SAME96_Screening",25,80)
};
var result=await ScreeningReportRenderer.CreateAsync(exam,output);
Console.WriteLine(result.PdfPath);
foreach(var path in result.PreviewImagePaths) Console.WriteLine(path);
if(result.PageCount!=2)throw new Exception("Expected two normal report pages.");
// A reprint snapshot must remain stable after the caller changes mutable examination fields.
exam.Reference="CHANGED"; exam.Otoscopy.Notes="CHANGED";
using(var print=File.Create(Path.Combine(output,"selected-page.pdf"))) result.WritePages!(print,new[]{1},595.276f,841.890f,18,default);
exam.Otoscopy.Notes=string.Join(" ",Enumerable.Repeat("En lång anteckning som inte får tappas bort.",150));
var longReport=await ScreeningReportRenderer.CreateAsync(exam,Path.Combine(output,"long"));
if(longReport.PageCount<=2)throw new Exception("Expected continuation pages.");
exam.Otoscopy = new() {LeftNotCapturedReason="Kameran var inte tillgänglig.", RightNotCapturedReason="Kameran var inte tillgänglig.", Notes="Inga bilder registrerades."};
exam.Ssq12=exam.Ssq12 with {Answers=Enumerable.Range(1,12).Select(i=>new Ssq12Answer(i,"Fråga "+i,null,true)).ToArray()};
exam.QuickSip=exam.QuickSip with {ScoreFraction=null};
var missingReport=await ScreeningReportRenderer.CreateAsync(exam,Path.Combine(output,"missing"));
Console.WriteLine(missingReport.PdfPath);
using (var canceled = new CancellationTokenSource())
{
 canceled.Cancel();
 try { await ScreeningReportRenderer.CreateAsync(exam, Path.Combine(output,"cancelled"), canceled.Token); throw new Exception("Expected cancellation."); }
 catch(OperationCanceledException) { }
 if(Directory.Exists(Path.Combine(output,"cancelled"))) throw new Exception("Cancelled report must not write files.");
}
try { await ScreeningReportRenderer.CreateAsync(new ScreeningExamination(), Path.Combine(output,"incomplete")); throw new Exception("Expected incomplete examination rejection."); }
catch(InvalidOperationException) { }
using(var invalid=new MemoryStream())
{
 try { result.WritePages!(invalid,new[]{99},595.276f,841.890f,0,default); throw new Exception("Expected invalid range rejection."); }
 catch(ArgumentException) { }
}
// The print snapshot must own its image data, rather than reading a new/replaced patient's paths.
File.Delete(photo);
using(var print=File.Create(Path.Combine(output,"snapshot-after-photo-deletion.pdf"))) result.WritePages!(print,new[]{0},595.276f,841.890f,0,default);
foreach(var path in result.PreviewImagePaths)
{
 using var image=SKBitmap.Decode(path);
 if(image.Width != (int)Math.Ceiling(ScreeningReportRenderer.A4Width*2) || image.Height != (int)Math.Ceiling(ScreeningReportRenderer.A4Height*2))
  throw new Exception("Preview is not A4 at 144 dpi.");
}
if(!System.Text.Encoding.ASCII.GetString(File.ReadAllBytes(result.PdfPath),0,5).Equals("%PDF-"))throw new Exception("Report is not a PDF.");
// The MAUI plot bridge must not alter the surrounding native PDF/preview canvas.
using(var bitmap=new SKBitmap(1200,1700))
using(var canvas=new SKCanvas(bitmap))
{
 canvas.Scale(1.75f); canvas.Translate(9,13); canvas.ClipRect(new SKRect(2,4,600,850));
 var matrix=canvas.TotalMatrix; var clip=canvas.DeviceClipBounds; int saves=canvas.SaveCount;
 using(var bridge=new SkiaCanvas {Canvas=canvas})
  new ScreeningAudiogram(exam.Audiometry!).DrawScreening(bridge,new RectF(36,359,523.276f,278));
 if(!canvas.TotalMatrix.Equals(matrix) || !canvas.DeviceClipBounds.Equals(clip) || canvas.SaveCount!=saves)
  throw new Exception("The audiogram bridge changed the report canvas transform, clip or save stack.");
 using var paint=new SKPaint {Color=SKColors.Black}; canvas.DrawRect(10,10,20,20,paint);
}
Console.WriteLine($"Report checks passed: normal {result.PageCount} pages, long notes {longReport.PageCount} pages.");
public sealed class GraphicsView {public void Invalidate() {}}
