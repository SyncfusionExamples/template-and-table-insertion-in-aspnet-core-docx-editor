using System.Text;
using BitMiracle.LibTiff.Classic;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using SkiaSharp;
using Syncfusion.DocIORenderer;
using Syncfusion.EJ2.DocumentEditor;
using Syncfusion.EJ2.SpellChecker;
using Syncfusion.Office;
using Syncfusion.Pdf;
using FormatType = Syncfusion.EJ2.DocumentEditor.FormatType;
using WDocument = Syncfusion.DocIO.DLS.WordDocument;
using WFormatType = Syncfusion.DocIO.FormatType;

namespace AspNetCoreSample.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentEditorController : Controller
    {


        [Route("Import")]
        public string Import(IFormCollection data)
        {
            if (data.Files.Count == 0)
                return null;
            Stream stream = new MemoryStream();
            IFormFile file = data.Files[0];
            int index = file.FileName.LastIndexOf('.');
            string type = index > -1 && index < file.FileName.Length - 1 ?
                file.FileName.Substring(index) : ".docx";
            file.CopyTo(stream);
            stream.Position = 0;
            // Create instance first
            WDocument wDocument = new WDocument();
            wDocument.Settings.SkipIncrementalSaveValidation = true;

            wDocument.Open(stream, WFormatType.Automatic);

            //Hooks MetafileImageParsed event.
            WordDocument.MetafileImageParsed += OnMetafileImageParsed;
            //Converts Stream DOM to SFDT DOM.
            WordDocument document = WordDocument.Load(wDocument);
            //Unhooks MetafileImageParsed event.
            WordDocument.MetafileImageParsed -= OnMetafileImageParsed;

            string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
            document.Dispose();
            return json;
        }

        private static void OnMetafileImageParsed(object sender, MetafileImageParsedEventArgs args)
        {
            if (args.IsMetafile)
            {
                //MetaFile image conversion(EMF and WMF)
                //You can write your own method definition for converting metafile to raster image using any third-party image converter.
                //args.ImageStream = ConvertMetafileToRasterImage(args.MetafileStream);
            }
            else
            {
                //TIFF image conversion
                args.ImageStream = ConvertTiffToRasterImage(args.MetafileStream);
            }
        }

        private static MemoryStream ConvertTiffToRasterImage(Stream tiffStream)
        {
            MemoryStream imageStream = new MemoryStream();
            using (Tiff tif = Tiff.ClientOpen("in-memory", "r", tiffStream, new TiffStream()))
            {
                // Find the width and height of the image
                FieldValue[] value = tif.GetField(BitMiracle.LibTiff.Classic.TiffTag.IMAGEWIDTH);
                int width = value[0].ToInt();

                value = tif.GetField(BitMiracle.LibTiff.Classic.TiffTag.IMAGELENGTH);
                int height = value[0].ToInt();

                // Read the image into the memory buffer
                int[] raster = new int[height * width];
                if (!tif.ReadRGBAImage(width, height, raster))
                {
                    throw new Exception("Could not read image");
                }

                // Create a bitmap image using SkiaSharp.
                using (SKBitmap sKBitmap = new SKBitmap(width, height, SKImageInfo.PlatformColorType, SKAlphaType.Premul))
                {
                    // Convert a RGBA value to byte array.
                    byte[] bitmapData = new byte[sKBitmap.RowBytes * sKBitmap.Height];
                    for (int y = 0; y < sKBitmap.Height; y++)
                    {
                        int rasterOffset = y * sKBitmap.Width;
                        int bitsOffset = (sKBitmap.Height - y - 1) * sKBitmap.RowBytes;

                        for (int x = 0; x < sKBitmap.Width; x++)
                        {
                            int rgba = raster[rasterOffset++];
                            bitmapData[bitsOffset++] = (byte)((rgba >> 16) & 0xff);
                            bitmapData[bitsOffset++] = (byte)((rgba >> 8) & 0xff);
                            bitmapData[bitsOffset++] = (byte)(rgba & 0xff);
                            bitmapData[bitsOffset++] = (byte)((rgba >> 24) & 0xff);
                        }
                    }

                    // Convert a byte array to SKColor array.
                    SKColor[] sKColor = new SKColor[bitmapData.Length / 4];
                    int index = 0;
                    for (int i = 0; i < bitmapData.Length; i++)
                    {
                        sKColor[index] = new SKColor(bitmapData[i + 2], bitmapData[i + 1], bitmapData[i], bitmapData[i + 3]);
                        i += 3;
                        index += 1;
                    }

                    // Set the SKColor array to SKBitmap.
                    sKBitmap.Pixels = sKColor;

                    // Save the SKBitmap to PNG image stream.
                    sKBitmap.Encode(SKEncodedImageFormat.Png, 100).SaveTo(imageStream);
                    imageStream.Flush();
                }
            }
            return imageStream;
        }
        public class CustomParameter
        {
            public string content { get; set; }
            public string type { get; set; }
        }


        [Route("SystemClipboard")]
        public string SystemClipboard([FromBody] CustomParameter param)
        {
            if (param.content != null && param.content != "")
            {
                try
                {
                    //Hooks MetafileImageParsed event.
                    WordDocument.MetafileImageParsed += OnMetafileImageParsed;
                    WordDocument document = WordDocument.LoadString(param.content, GetFormatType(param.type.ToLower()));
                    //Unhooks MetafileImageParsed event.
                    WordDocument.MetafileImageParsed -= OnMetafileImageParsed;
                    string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                    document.Dispose();
                    return json;
                }
                catch (Exception)
                {
                    return "";
                }
            }
            return "";
        }


        [Route("protectDocument")]
        public string protectDocument([FromBody] SaveParameter data)
        {
            string sfdtText = "";
            // Converts the sfdt to stream
            Stream doc = WordDocument.Save(data.Content, FormatType.Docx);
            Syncfusion.DocIO.DLS.WordDocument document = new Syncfusion.DocIO.DLS.WordDocument(doc, Syncfusion.DocIO.FormatType.Docx);
            // Find start and end markers
            List<Syncfusion.DocIO.DLS.Entity> startEntity = document.FindAllItemsByProperty(Syncfusion.DocIO.DLS.EntityType.MergeField, "FieldName", "UserEditStart");

            List<Syncfusion.DocIO.DLS.Entity> endEntity = document.FindAllItemsByProperty(Syncfusion.DocIO.DLS.EntityType.MergeField, "FieldName", "UserEditEnd");

            for (int i = 0; i < startEntity.Count && i < endEntity.Count; i++)
            {
                Syncfusion.DocIO.DLS.WMergeField startField = startEntity[i] as Syncfusion.DocIO.DLS.WMergeField;
                Syncfusion.DocIO.DLS.WMergeField endField = endEntity[i] as Syncfusion.DocIO.DLS.WMergeField;

                if (startField != null && endField != null)
                {
                    Syncfusion.DocIO.DLS.WParagraph startPara = startField.OwnerParagraph;
                    Syncfusion.DocIO.DLS.WParagraph endPara = endField.OwnerParagraph;

                    // Get index of merge fields
                    int startIndex = startPara.ChildEntities.IndexOf(startField);


                    // Create editable range
                    Syncfusion.DocIO.DLS.EditableRangeStart rangeStart = new Syncfusion.DocIO.DLS.EditableRangeStart(document);
                    Syncfusion.DocIO.DLS.EditableRangeEnd rangeEnd = new Syncfusion.DocIO.DLS.EditableRangeEnd(document, rangeStart);

                    // Insert start at exact position of merge field
                    startPara.ChildEntities.Insert(startIndex, rangeStart);

                    int endIndex = endPara.ChildEntities.IndexOf(endField);
                    // Insert end at exact position of end marker
                    endPara.ChildEntities.Insert(endIndex, rangeEnd);

                    // Remove the merge fields after creating editable range
                    startPara.ChildEntities.Remove(startField);
                    endPara.ChildEntities.Remove(endField);
                }
            }

            // Protect document (only editable regions allowed)
            document.Protect(Syncfusion.DocIO.ProtectionType.AllowOnlyReading, "password");
            MemoryStream stream = new MemoryStream();
            document.Save(stream, WFormatType.Docx);
            stream.Position = 0;
            WordDocument mergeDocument = WordDocument.Load(stream, FormatType.Docx);
            sfdtText = Newtonsoft.Json.JsonConvert.SerializeObject(mergeDocument);
            doc.Dispose();
            stream.Dispose();
            return sfdtText;
        }

        public class SpellCheckJsonData
        {
            public int LanguageID { get; set; }
            public string TexttoCheck { get; set; }
            public bool CheckSpelling { get; set; }
            public bool CheckSuggestion { get; set; }
            public bool AddWord { get; set; }

        }

        [Route("SpellCheck")]
        public string SpellCheck([FromBody] SpellCheckJsonData spellChecker)
        {
            try
            {
                SpellChecker spellCheck = new SpellChecker();
                spellCheck.GetSuggestions(spellChecker.LanguageID, spellChecker.TexttoCheck, spellChecker.CheckSpelling, spellChecker.CheckSuggestion, spellChecker.AddWord);
                return Newtonsoft.Json.JsonConvert.SerializeObject(spellCheck);
            }
            catch
            {
                return "{\"SpellCollection\":[],\"HasSpellingError\":false,\"Suggestions\":null}";
            }
        }



        [Route("SpellCheckByPage")]
        public string SpellCheckByPage([FromBody] SpellCheckJsonData spellChecker)
        {
            try
            {
                SpellChecker spellCheck = new SpellChecker();
                spellCheck.CheckSpelling(spellChecker.LanguageID, spellChecker.TexttoCheck);
                return Newtonsoft.Json.JsonConvert.SerializeObject(spellCheck);
            }
            catch
            {
                return "{\"SpellCollection\":[],\"HasSpellingError\":false,\"Suggestions\":null}";
            }
        }
        public class SaveParameter
        {
            public string Content { get; set; }
            public string FileName { get; set; }
            public int RowIndex { get; set; }
        }
        public class InputParameter
        {
            public string HtmlContent { get; set; }
        }

        [Route("ExportPdf")]

        public void ExportPdf([FromBody] SaveParameter data)
        {
            // Converts the sfdt to stream
            Stream document = WordDocument.Save(data.Content, FormatType.Docx);

            Syncfusion.DocIO.DLS.WordDocument doc = new Syncfusion.DocIO.DLS.WordDocument(document, Syncfusion.DocIO.FormatType.Docx);
            //Instantiation of DocIORenderer for Word to PDF conversion
            DocIORenderer render = new DocIORenderer();
            //Converts Word document into PDF document
            PdfDocument pdfDocument = render.ConvertToPDF(doc);
            // Saves the document to server machine file system, you can customize here to save into databases or file servers based on requirement.
            FileStream fileStream = new FileStream("Sample.pdf", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            //Saves the PDF file
            pdfDocument.Save(fileStream);
            pdfDocument.Close();
            fileStream.Close();
            document.Close();
        }

        [Route("Save")]

        public void Save([FromBody] SaveParameter data)
        {
            string name = data.FileName;
            if (string.IsNullOrEmpty(name))
            {
                name = "Document1";
            }
            WDocument document = WordDocument.Save(data.Content);
            FileStream fileStream = new FileStream(name + ".docx", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            document.Save(fileStream, WFormatType.Docx);
            document.Close();
            fileStream.Close();
        }

        [Route("LoadStringHtml")]
        public string LoadStringHtml([FromBody] InputParameter data)
        {
            // You can also load HTML file/string from server side.
            Syncfusion.EJ2.DocumentEditor.WordDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.LoadString(data.HtmlContent, FormatType.Html); // Convert the HTML to SFDT format.
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
            document.Dispose();
            return json;
        }

        public class TemplateRequest
        {
            public string TemplateId { get; set; }
        }

        [Route("ImportTemplate")]
        public string ImportTemplate([FromBody] TemplateRequest request)
        {
            try
            {
                // Map template IDs to file names
                string fileName = request.TemplateId switch
                {
                    "AdventureWorks_CoverPage" => "AdventureWorks_CoverPage.docx",
                    "Sustainability_Template" => "Sustainability_Template.docx",
                    "Market_Analysis_Template" => "Market_Analysis_Template.docx",
                    "empty_template" => "empty_template.docx",
                    _ => "empty_template.docx"
                };

                string filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", fileName);

                // Check if file exists
                if (!System.IO.File.Exists(filePath))
                {
                    return "{}"; // Return empty SFDT if file not found
                }

                using FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);

                WordDocument.MetafileImageParsed += OnMetafileImageParsed;

                WordDocument document = WordDocument.Load(stream, FormatType.Docx);

                WordDocument.MetafileImageParsed -= OnMetafileImageParsed;

                string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);

                document.Dispose();

                return json;
            }
            catch (Exception ex)
            {
                return "{}"; // Return empty SFDT on error
            }
        }

        internal static FormatType GetFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            switch (format.ToLower())
            {
                case ".dotx":
                case ".docx":
                case ".docm":
                case ".dotm":
                    return FormatType.Docx;
                case ".dot":
                case ".doc":
                    return FormatType.Doc;
                case ".rtf":
                    return FormatType.Rtf;
                case ".txt":
                    return FormatType.Txt;
                case ".xml":
                    return FormatType.WordML;
                case ".html":
                    return FormatType.Html;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }

        private string RetrieveFileType(string name)
        {
            int index = name.LastIndexOf('.');
            string format = index > -1 && index < name.Length - 1 ?
                name.Substring(index) : ".doc";
            return format;
        }
        internal static WFormatType GetWFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            switch (format.ToLower())
            {
                case ".dotx":
                    return WFormatType.Dotx;
                case ".docx":
                    return WFormatType.Docx;
                case ".docm":
                    return WFormatType.Docm;
                case ".dotm":
                    return WFormatType.Dotm;
                case ".dot":
                    return WFormatType.Dot;
                case ".doc":
                    return WFormatType.Doc;
                case ".rtf":
                    return WFormatType.Rtf;
                case ".html":
                    return WFormatType.Html;
                case ".txt":
                    return WFormatType.Txt;
                case ".xml":
                    return WFormatType.WordML;
                case ".odt":
                    return WFormatType.Odt;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }
    }


}
