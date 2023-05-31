// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlLab;

static IEnumerable<MeasuredParagraph> ExtractMeasuredParagraphs(string filepath)
{
    using (var doc = WordprocessingDocument.Open(filepath, false))
    {
        var output = new List<MeasuredParagraph>();
        Body? body = doc.MainDocumentPart?.Document.Body;

        if (body != null)
        {
            foreach (var child in body.ChildElements)
            {
                var paragraph = child as Paragraph;
                if (paragraph == null)
                {
                    continue;
                }

                var p = new MeasuredParagraph(paragraph);
                output.Add(p);
            }
        }

        return output;
    }
}

static void ReadWordDoc(string filepath)
{
    var fileName = Path.GetFileName(filepath);
    var list = ExtractMeasuredParagraphs(filepath);
    foreach (var item in list)
    {
        Console.WriteLine($"\"{fileName}\",{item.ParaId},\"{item.SampleText}\",{item.InnerTextLength}");
    }
}

if (args.Length < 1)
{
    Console.WriteLine("please include filepath");
    return;
}

ReadWordDoc(args[0]);

