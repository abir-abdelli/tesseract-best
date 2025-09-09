using Tesseract;
using Xceed.Document.NET;
using Xceed.Words.NET;


Console.WriteLine("Extracting started!");
string tessDataPath = @"./tessdata";
string imagesPath = @"./images";
string wordPath = Path.Combine(tessDataPath,"OCR_Output.docx");
List<string> extractedText = new List<string>();


using var engine = new TesseractEngine(tessDataPath,"ara",EngineMode.Default);
engine.SetVariable("tessedit_char_whitelist","ءآأؤإئابةتثجحخدذرزسشصضطظعغفقكلمنهوىيًٌٍَُِّْ٠۰١۱٢۲٣۳٤۴٥۵٦۶٧۷٨۸٩۹ ");



string[] imageFiles = Directory.GetFiles(imagesPath,"*.*");

foreach(string imageFile in imageFiles) {
    try {
        if(imageFile.EndsWith(".png",StringComparison.OrdinalIgnoreCase)) {
            using var pix = Pix.LoadFromFile(imageFile);

            using var page = engine.Process(pix);
            string text = page.GetText();
            extractedText.Add(text);
        }
    } catch(Exception ex) {
        Console.WriteLine($"Error processing {imageFile}: {ex.Message}");
    }
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
