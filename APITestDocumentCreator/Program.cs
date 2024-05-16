using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System.Text;
using System.Text.RegularExpressions;

namespace APITestDocumentCreator
{
    internal class Program
    {
        static void Main()
        {
            // Base, pictures and result folder paths for the application, so it can read the input file and export the final .docx
            string baseFolder = $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\\API_Test_Document_Creator";
            string resultFolder = $"{baseFolder}\\Result";
            string picturesFolder = $"{baseFolder}\\Pictures";

            // Create folders and input files necessary for the application.
            CreateApplicationBasicStructure(baseFolder, resultFolder, picturesFolder);

            // Retrieving basic document information.
            string titleText, documentAuthor;
            DocumentBasicInformation(out titleText, out documentAuthor);

            // Ask the user if the data inside the 'Input_Txt.txt' follows a specific pattern, that it's showed in the console.
            DataPatternApresentationAndVerification();

            // Data validation of every input file
            string[] picturesList; // Array that contains the path of each file found in 'Pictures' folder.
            List<InputData> dataList; // This list contains all lines in the 'Input_Data.txt' file.
            HighlightParameters? highlightParameters; // Holds all JSON parameters that need highlight in the final document.
            List<SectionProperties> sectionList; // List that will inform all properties of each section of the document.

            InputFilesValidation(baseFolder, picturesFolder, out picturesList, out dataList, out highlightParameters, out sectionList);

            // Creating the .docx document
            using (XWPFDocument document = new())
            {
                // DOCUMENT PROPERTIES
                POIXMLProperties properties = document.GetProperties();

                NPOI.OpenXml4Net.OPC.Internal.PackagePropertiesPart underlyingProp = properties.CoreProperties.GetUnderlyingProperties();
                underlyingProp.SetCreatorProperty(documentAuthor);

                NPOI.OpenXmlFormats.CT_ExtendedProperties extendedProp = properties.ExtendedProperties.GetUnderlyingProperties();
                extendedProp.Application = "Microsoft Office Word";

                // DOCUMENT TITLE
                XWPFParagraph titleParagraph = document.CreateParagraph();

                ParagraphStylizer(titleParagraph, ParagraphAlignment.CENTER, TextAlignment.CENTER, Borders.Single);
                RunStylizer(titleParagraph, 18, titleText.ToUpper(), true);

                int tempSectionNumber = 0; // This counter will be used to track lines in 'Input_Data' that have the same section number.

                // Writing in the document every line that the application read in the 'Input_Data'.
                foreach (InputData data in dataList)
                {
                    if (data.SectionNumber > tempSectionNumber)
                    {
                        // Veryifing section properties based on the tempSectionNumber so we can use later.
                        SectionProperties sectionNow = sectionList.SingleOrDefault(section => section.SectionNumber.Equals(tempSectionNumber + 1));

                        if (tempSectionNumber > 0)
                        {
                            // Adding a page break in every new section after the first one.
                            XWPFParagraph addBreak = document.CreateParagraph();
                            XWPFRun addBreakRun = addBreak.CreateRun();
                            addBreakRun.AddBreak(BreakType.PAGE);
                        }

                        // SECTION TITLE
                        XWPFParagraph documentSection = document.CreateParagraph();
                        ParagraphStylizer(documentSection, ParagraphAlignment.LEFT);

                        string sectionText = $"{sectionNow.SectionNumber} - {sectionNow.SectionTitle.ToUpper()}";
                        RunStylizer(documentSection, 14, sectionText, true, UnderlinePatterns.Single, "44AE2F");

                        // SECTION PICTURES
                        List<string> sectionPictures = new();

                        // In picturesList is save the full patch to every picture, so the application will retrieve only the name of the file.
                        foreach (string picture in picturesList)
                        {
                            string pictureName = Path.GetFileNameWithoutExtension(picture);
                            string[] pictureNameParts = pictureName.Split('_');

                            if (pictureNameParts[0].Equals(data.SectionNumber.ToString()) == true)
                            {
                                sectionPictures.Add(picture);
                            }
                        }

                        if (sectionPictures.Count > 0)
                        {
                            XWPFParagraph documentSectionPictures = document.CreateParagraph();
                            ParagraphStylizer(documentSectionPictures, ParagraphAlignment.CENTER);

                            // Here we can define the image dimensions with ease, so the application will convert after (in EMUs (https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/)).
                            int widthCentimeters = 15;
                            int heightCentimeters = 10;

                            int widthEmus = widthCentimeters * 360000;
                            int heightEmus = heightCentimeters * 360000;

                            foreach (string picPath in sectionPictures)
                            {
                                XWPFRun pictureRun = documentSectionPictures.CreateRun();

                                using (FileStream picData = new FileStream(picPath, FileMode.Open, FileAccess.Read))
                                {
                                    pictureRun.AddPicture(picData, (int)PictureType.PNG, "image1", widthEmus, heightEmus);
                                }

                                pictureRun.AddCarriageReturn();
                            }
                        }

                        // SECTION DESCRIPTION

                        if (sectionNow != null)
                        {
                            XWPFParagraph sectionDescription = document.CreateParagraph();
                            ParagraphStylizer(sectionDescription, ParagraphAlignment.BOTH);

                            string sectionDescriptionText = $"Descrição: {sectionNow.Description}";
                            RunStylizer(sectionDescription, 10, sectionDescriptionText, false, UnderlinePatterns.None, "000000", "Calibri", true);
                        }

                        tempSectionNumber = data.SectionNumber;
                    }

                    // ENDPOINT REQUEST TITLE
                    XWPFParagraph endpointRequest = document.CreateParagraph();
                    ParagraphStylizer(endpointRequest, ParagraphAlignment.CENTER, TextAlignment.CENTER, Borders.Single);

                    string endpointRequestText = $"REQUISIÇÃO - {data.MethodName.ToUpper()}";
                    RunStylizer(endpointRequest, 12, endpointRequestText, true, UnderlinePatterns.None, "297FC2");

                    // ENDPOINT REQUEST TITLE - URL USED
                    XWPFParagraph endpointRequestURL = document.CreateParagraph();
                    ParagraphStylizer(endpointRequestURL);

                    string URLText = $"URL: {data.URL}";
                    RunStylizer(endpointRequestURL, 10,  URLText);

                    // ENDPOINT REQUEST - JSON TEXT
                    XWPFParagraph endpointRequestJSON = document.CreateParagraph();
                    ParagraphStylizer(endpointRequestJSON);

                    XWPFRun endpointRequestJSONRun = endpointRequestJSON.CreateRun();
                    endpointRequestJSONRun.FontFamily = "Calibri"; // Set font to maintain preformatted style
                    endpointRequestJSONRun.FontSize = 10;

                    string jsonRequestText = data.Request;

                    if (jsonRequestText == "NULL")
                    {
                        endpointRequestJSONRun.SetText($"BODY: null");
                    }
                    else
                    {
                        endpointRequestJSONRun.SetText($"BODY:");
                        JSONFormatter(highlightParameters, document, jsonRequestText);
                    }

                    // ENDPOINT RESPONSE TITLE
                    XWPFParagraph endpointResponse = document.CreateParagraph();
                    ParagraphStylizer(endpointResponse, ParagraphAlignment.CENTER, TextAlignment.CENTER, Borders.Single);

                    string responseTitleText = $"RESPOSTA - {data.MethodName.ToUpper()}";
                    RunStylizer(endpointResponse, 12, responseTitleText, true, UnderlinePatterns.None, "FF0000");

                    // ENDPOINT RESPONSE - JSON TEXT
                    XWPFRun endpointResponseJSONRun = endpointResponse.CreateRun();
                    endpointResponseJSONRun.FontFamily = "Calibri"; // Set font to maintain preformatted style
                    endpointResponseJSONRun.FontSize = 10;

                    string jsonResponseText = data.Response;

                    JSONFormatter(highlightParameters, document, jsonResponseText);
                }

                // Create an docx. file and writes the document content into it
                using (FileStream fs = new($"{resultFolder}\\{titleText}.docx", FileMode.Create, FileAccess.Write))
                {
                    document.Write(fs);
                }
            }
        }

        private static void JSONFormatter(HighlightParameters? highlightParameters, XWPFDocument document, string jsonText)
        {
            // This variables will help in the parts where we describe the request and response text
            string jsonWithIdentation;
            string[] separator = new[] { "\r\n", "\r", "\n" };
            string[] lines;

            jsonWithIdentation = PrettyJson(jsonText); // Formatting the JSON
            lines = jsonWithIdentation.Split(separator, StringSplitOptions.None);

            // Create a paragraph within the cell for each line of the JSON content
            foreach (string line in lines)
            {
                XWPFParagraph endpointJSON = document.CreateParagraph();

                XWPFRun run = endpointJSON.CreateRun();
                run.SetText(line);
                run.FontFamily = "Calibri"; // Set font to maintain preformatted style
                run.FontSize = 10;

                // Every line of the JSON is composed of a parameter name and its value and this function will extract the name and compare
                // to a JSON list created and populated by the user.
                HighlightRun(highlightParameters, line, run);

                // Set indentation to mimic JSON structure
                int indentationLevel = GetIndentationLevel(line);
                for (int i = 0; i < indentationLevel; i++)
                {
                    endpointJSON.IndentationFirstLine = i * 720; // 720 twips = 1/2 inch
                }
            }
        }

        private static void ParagraphStylizer(XWPFParagraph paragraph, ParagraphAlignment paragraphAlignment = ParagraphAlignment.LEFT, TextAlignment textAlignment = TextAlignment.CENTER, Borders borderStyle = Borders.None)
        {
            paragraph.Alignment = paragraphAlignment;
            paragraph.VerticalAlignment = textAlignment;
            paragraph.BorderTop = borderStyle;
            paragraph.BorderLeft = borderStyle;
            paragraph.BorderRight = borderStyle;
            paragraph.BorderBottom = borderStyle;
        }

        private static void RunStylizer(XWPFParagraph paragraph, int fontSize, string printText, bool bold = false, UnderlinePatterns underline = UnderlinePatterns.None, string color = "000000", string fontFamily = "Calibri", bool italic = false)
        {
            XWPFRun run = paragraph.CreateRun();
            run.FontFamily = fontFamily;
            run.FontSize = fontSize;
            run.IsBold = bold;
            run.Underline = underline;
            run.IsItalic = italic;
            run.SetColor(color);
            run.SetText(printText);
        }

        private static void InputFilesValidation(string baseFolder, string picturesFolder, out string[] picturesList, out List<InputData> dataList, out HighlightParameters? highlightParameters, out List<SectionProperties> sectionList)
        {
            Console.WriteLine("\n\n-- INPUT FILES VALIDATION --\n");

            // Initializing key variables.
            dataList = Enumerable.Empty<InputData>().ToList();
            sectionList = Enumerable.Empty<SectionProperties>().ToList();
            picturesList = [];
            highlightParameters = new();

            FileStreamOptions options = new() { Access = FileAccess.Read, Mode = FileMode.Open, Options = FileOptions.None };

            try
            {
                // Every line of the file will be transformed in a instance of InputData so the values can be accessed as a parameter.
                using (StreamReader streamInputData = new($"{baseFolder}\\Input_Data.txt", Encoding.UTF8, false, options))
                {
                    string fileLine;

                    // Each line is stored inside the 'fileLine' variable so it can be analyzed.
                    while ((fileLine = streamInputData.ReadLine()) != null)
                    {
                        string[] dataFields = fileLine.Split(';');

                        InputData data = new()
                        {
                            SectionNumber = int.Parse(dataFields[0]),
                            MethodName = Regex.Replace(dataFields[1], "([A-Z])(?![A-Z])", " $1").ToUpper(), // Separates each word in the field.
                            URL = dataFields[2],
                            Request = dataFields[3],
                            Response = dataFields[4]
                        };

                        dataList.Add(data);
                    }
                }
            }
            catch (IOException ioException)
            {
                Console.WriteLine($"\n\n[ERROR] Another program is using the file, you need to close it before run this application!\n> DETAILS: {ioException.Message}");
            }
            catch (Exception exception)
            {
                PrintGenericErrorException(exception);
            }

            // Check if the 'Input_Data.txt' file have any information, because the application won't work properly without it.
            if (dataList.Count > 0)
            {
                Console.WriteLine($"[INFO] All {dataList.Count} lines in 'Input_Data.txt' was read by the application!");

                try
                {
                    // Every line of the file will be transformed in a instance of SectionProperties so the values can be accessed as a parameter.
                    using (StreamReader streamSectionFile = new($"{baseFolder}\\Section_Information.txt", Encoding.UTF8, false, options))
                    {
                        string sectionFileLine;

                        // Each line is stored inside the 'fileLine' variable so it can be analyzed.
                        while ((sectionFileLine = streamSectionFile.ReadLine()) != null)
                        {
                            string[] sectionProperties = sectionFileLine.Split(';');

                            SectionProperties section = new()
                            {
                                SectionNumber = int.Parse(sectionProperties[0]),
                                SectionTitle = sectionProperties[1],
                                Description = sectionProperties[2]
                            };

                            sectionList.Add(section);
                        }
                    }
                }
                catch (IOException ioException)
                {
                    Console.WriteLine($"\n\n[ERROR] Another program is using the file, you need to close it before run this application!\n> DETAILS: {ioException.Message}");
                }
                catch (Exception exception)
                {
                    PrintGenericErrorException(exception);
                }

                // Retrieving all prints stored in the 'Pictures' folder.
                picturesList = Directory.GetFiles(picturesFolder);

                if(picturesList.Length > 0)
                {
                    Console.WriteLine($"[INFO] {dataList.Count} images were detected in the 'Pictures' folder!");
                }
                else
                {
                    Console.WriteLine($"[INFO] Nothing was found in the 'Pictures' folder!");
                }

                // Retrieving the parameters name list in the JSON file that the user wants to highlight in the document.
                string highlightParametersJson = File.ReadAllText($"{baseFolder}\\HighlightParameters.json");
                highlightParameters = JsonConvert.DeserializeObject<HighlightParameters>(highlightParametersJson);

                if(highlightParameters.ParametersList.Length > 0)
                {
                    Console.WriteLine($"[INFO] {highlightParameters.ParametersList.Length} different parameters will be highlighted in the document!");
                }
                else
                {
                    Console.WriteLine($"[INFO] No highlight will be needed in the document");
                }
            }
            else
            {
                Console.WriteLine("[INFO] No information was provided in the 'Input_Data.txt' file, the application can't progress without it!");
                Environment.Exit(0);
            }
        }

        private static void DocumentBasicInformation(out string? titleText, out string? documentAuthor)
        {
            Console.WriteLine("\n-- DOCUMENT PROPERTIES --\n");

            // Registering the user input for document title.
            Console.Write("> What is the title of the document? (The title will be the result file name)\nTitle: ");
            titleText = Console.ReadLine();
            while (titleText == "")
            {
                Console.Write("[INFO] The title can't be empty!\nTitle: ");
                titleText = Console.ReadLine();
            }

            Console.Write("\n> What's the document's author name?\nAuthor: ");
            documentAuthor = Console.ReadLine();
            while (documentAuthor == "")
            {
                Console.Write("[INFO] The author's name can't be empty!\nAuthor: ");
                documentAuthor = Console.ReadLine();
            }
        }

        private static void HighlightRun(HighlightParameters highlightParameters, string line, XWPFRun run)
        {
            string[] parameterKeyValue = line.Split(':');
            string adjustedParameterName = parameterKeyValue[0].Replace("\"", "").Trim();
            bool highlightCondition = Array.Exists(highlightParameters.ParametersList, name => name.Equals(adjustedParameterName));

            if (highlightCondition == true)
            {
                run.GetCTR().AddNewRPr().highlight = new CT_Highlight
                {
                    val = ST_HighlightColor.yellow
                };
            }
        }

        private static void PrintGenericErrorException(Exception exception)
        {
            Console.WriteLine($"\n[ERROR]: An error has occurred! See details below: \n{exception.Message}");
        }

        private static void DataPatternApresentationAndVerification()
        {
            Console.WriteLine("\n-- INPUT DATA EXPLANATION --\n");

            Console.WriteLine("[INFO] Before the application read the 'Input_Data.txt' file is important that the data inside it follows the pattern below, separated with semi-colon (;):\n");
            Console.WriteLine("1) SECTION NUMBER: Section number in which all the input file will be showed in the document, if the section has more than one method put the same number but write in the order of appeareance.");
            Console.WriteLine("2) SECTION NAME: The name used to define the section, if more than one line has the same section number, the name must be the same.");
            Console.WriteLine("3) METHOD NAME: Name of the API method that will appear in the document.");
            Console.WriteLine("4) URL: What URL was used in the request?");
            Console.WriteLine("5) REQUEST: JSON used in the request (don't need pre-formatting).");
            Console.WriteLine("6) RESPONSE: JSON received in the response (don't need pre-formatting).");

            Console.WriteLine("\n> Before you continue, check if there is data inside the input file and follows the pattern above.\nProceed?");

            bool patternDecisionLoop = true;
            while (patternDecisionLoop)
            {
                Console.Write("\n- Type '1' if it's OK\n- Type '2' to exit application\nResponse: ");
                ConsoleKeyInfo patternDecision = Console.ReadKey();

                switch (patternDecision.Key)
                {
                    case ConsoleKey.NumPad1:
                        patternDecisionLoop = false;
                        break;
                    case ConsoleKey.D1:
                        patternDecisionLoop = false;
                        break;
                    case ConsoleKey.NumPad2:
                        Console.WriteLine("\n\n[INFO] Exiting application...");
                        Environment.Exit(0);
                        break;
                    case ConsoleKey.D2:
                        Console.WriteLine("\n\n[INFO] Exiting application...");
                        Environment.Exit(0);
                        break;
                    default:
                        break;
                }
            }
        }

        private static void CreateApplicationBasicStructure(string baseFolder, string resultFolder, string picturesFolder)
        {
            Console.WriteLine("-- BASIC APPLICATION STRUCTURE -- \n");
            try
            {
                // Checking if the folder structure already exists.
                if(Directory.Exists(resultFolder))
                {
                    Console.WriteLine($"[INFO] Basic folder already exists! No need for creation.");
                }
                else
                {
                    // Creating all basic folders.
                    Console.WriteLine($"[INFO] Creating basic folder strucutre in the following path: {baseFolder}");
                    Directory.CreateDirectory(resultFolder); // Here the application will create both base folder ande result.
                    Directory.CreateDirectory(picturesFolder);
                }

                // Checking if the input file don't exists so we don't accidentally recreate the file and delete the data inside it.
                if (!File.Exists($"{baseFolder}\\Input_Data.txt"))
                {
                    File.Create($"{baseFolder}\\Input_Data.txt").Close();
                    Console.WriteLine($"[INFO] Input file 'Input_Data.txt' has been created inside the folder");
                }
                else
                {
                    Console.WriteLine($"[INFO] Input file already exist, please put the data inside it!");
                }

                // Here the application will verify if the highlight file don't exists and if is true another one will be created.
                if (!File.Exists($"{baseFolder}\\HighlightParameters.json"))
                {
                    File.Create($"{baseFolder}\\HighlightParameters.json").Close();
                    Console.WriteLine($"[INFO] Highlight file 'HighlightParameters.json' has been created inside the folder");
                }
                else
                {
                    Console.WriteLine($"[INFO] Highlight file already exist, please put the parameter names in it!");
                }

                // This file will contain information about the section in the document (number, title and description)
                if (!File.Exists($"{baseFolder}\\Section_Information.txt"))
                {
                    File.Create($"{baseFolder}\\Section_Information.txt").Close();
                    Console.WriteLine($"[INFO] Section data file 'Section_Information.txt' has been created inside the folder");
                }
                else
                {
                    Console.WriteLine($"[INFO] Section data file already exist, please put the number, title and description of each section in the document!");
                }
            }
            catch (Exception exception)
            {
                PrintGenericErrorException(exception);
            }
        }

        // Apply indentation to the raw JSON string and return to be used in the document
        public static string PrettyJson(string unPrettyJson)
        {
            string jsonWithoutDoubleQuotation = unPrettyJson.Replace("\"\"", "\""); // Adjusting double quotation marks that appears when the application reads the JSON.
            string jsonBracketsRemovedStartEnd = jsonWithoutDoubleQuotation.Substring(1, jsonWithoutDoubleQuotation.Length - 2).TrimStart('[').TrimEnd(']'); // When a JSON starts and ends with a bracket, this will remove it because Newtonsoft interprets as an array instead of JSON and will cause an error.

            JObject parsedJson = JObject.Parse(jsonBracketsRemovedStartEnd);
            return JsonConvert.SerializeObject(parsedJson, Formatting.Indented);
        }

        // Helper method to get the indentation level of a JSON line
        static int GetIndentationLevel(string line)
        {
            int level = 0;
            foreach (char c in line)
            {
                if (c == '{' || c == '[')
                    level++;
                else if (c == '}' || c == ']')
                    level--;
            }
            return level;
        }
    }

    // Auxiliary classes to be able to serialize the JSON
    public class InputData()
    {
        public int SectionNumber { get; set; }
        public string MethodName { get; set; }
        public string URL { get; set; }
        public string Request { get; set; }
        public string Response { get; set; }
    }

    public class HighlightParameters()
    {
        public string[] ParametersList { get; set; }
    }

    public class SectionProperties()
    {
        public int SectionNumber {get; set; }
        public string SectionTitle { get; set; }
        public string Description { get; set; }
    }
}
