using ImageMagick;
using IronOcr;
using IronPdf;
using IronSoftware.Drawing;
using System.Text.RegularExpressions;
using Tesseract;
using Xceed.Document.NET;
using Xceed.Words.NET;
using static System.Net.Mime.MediaTypeNames;



Console.WriteLine("Extracting started!");
IronOcr.License.LicenseKey =  "your-key";// to be put in external file
bool licenseStatus = IronOcr.License.IsValidLicense(IronOcr.License.LicenseKey);

string dataPath = @"./tessdata";
string imagesPath = @"./images";
string wordPath = Path.Combine(dataPath,"OCR_Output.docx");
List<string> extractedText = new List<string>();
string[] imageFiles = Directory.GetFiles(imagesPath,"*.jpg");// to be configurable
List<OcrInput> inputs = new List<OcrInput>();


var myPath = Directory.GetFiles(dataPath,"*.pdf");
if(!myPath.Any()) {
    Console.WriteLine("No pdf was found, please add a pdf file and restart the service");
}
var pdf = PdfDocument.FromFile(myPath[0]);
pdf.RasterizeToImageFiles( imagesPath + @"\page-*.png");




var Ocr = new IronTesseract();
Ocr.Language = OcrLanguage.Arabic;



using OcrInput input = new OcrInput();
foreach(string imageFilePath in imageFiles) {
    
    input.LoadImage(imageFilePath);
    var result = Ocr.Read(input);
    extractedText.Add(result.Text);
    input.RemovePage(0);
}

using(var doc = DocX.Create(wordPath)) {
    foreach(string text in extractedText) {
        var p = doc.InsertParagraph(text);
        p.Font("Arial");
        p.FontSize(14);
        p.Alignment = Alignment.right;
        p.Direction = Direction.RightToLeft;

        doc.InsertSectionPageBreak();
    }
    doc.Save();
}
Console.WriteLine($"Task completed! Word document saved at: {Path.GetFullPath(wordPath)}");