using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

if (args.Length < 1)
{
    Console.WriteLine("Usage: <inputFilePath> [outputDirectoryPath]");
    Console.ReadLine();
    return;
}

var inputFilePath = args[0];
var outputDirectoryPath = args.Length > 1 ? args[1] : Path.GetDirectoryName(inputFilePath);

if (string.IsNullOrEmpty(inputFilePath) || !File.Exists(inputFilePath))
{
    throw new ArgumentException("Invalid input file path.", nameof(inputFilePath));
}

if (string.IsNullOrEmpty(outputDirectoryPath) || !Directory.Exists(outputDirectoryPath))
{
    throw new ArgumentException("Invalid output directory path.", nameof(outputDirectoryPath));
}

var protoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "proto.docx");
if (!File.Exists(protoPath))
{
    throw new FileNotFoundException($"Proto file not found: {protoPath}");
}

var allText = ReadAllTextFromDocx(inputFilePath);

var outputFilePath = Path.Combine(outputDirectoryPath, "Output.docx");
File.Copy(protoPath, outputFilePath, true);

InsertTextToDocx(outputFilePath, allText);

Console.WriteLine($"File created at: {outputFilePath}");

static string ReadAllTextFromDocx(string filePath)
{
    using WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false);

    var body = wordDoc.MainDocumentPart?.Document.Body;
    if (body == null)
        throw new ArgumentException(nameof(body));

    return string.Join("", body.Elements<Paragraph>().Select(p => p.InnerText));
}

static void InsertTextToDocx(string filePath, string text)
{
    using WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true);
    var body = wordDoc.MainDocumentPart?.Document.Body;
    if (body == null)
        throw new ArgumentException(nameof(body));

    var paragraph = body.AppendChild(new Paragraph());
    var run = paragraph.AppendChild(new Run());
    run.AppendChild(new Text(text));

    wordDoc.MainDocumentPart!.Document.Save();
}
