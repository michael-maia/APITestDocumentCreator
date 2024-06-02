using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System.Drawing;
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
            List<HighlightParameters> highlightParametersList; // Holds all JSON parameters that need highlight in the final document.
            List<SectionProperties> sectionList; // List that will inform all properties of each section of the document.

            InputFilesValidation(baseFolder, picturesFolder, out picturesList, out dataList, out highlightParametersList, out sectionList);

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

                            foreach (string picPath in sectionPictures)
                            {
                                // Here we can define the image dimensions with ease, so the application will convert after (in EMUs (https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/)).
                                int widthCentimeters = 15;
                                int heightCentimeters = 10;

                                int widthEmus = 0;
                                int heightEmus = 0;

                                // Creating a instance of Bitmap so we can access the image dimensions.
                                Bitmap image = new(picPath);

                                // If the dimensions of the image are small, considering its size in comparison to others, the application will use another value for width and height.
                                if (image.Width * image.Height < 100000)
                                {
                                    widthCentimeters = 10;
                                    heightCentimeters = 4;
                                }

                                widthEmus = widthCentimeters * 360000;
                                heightEmus = heightCentimeters * 360000;

                                XWPFRun pictureRun = documentSectionPictures.CreateRun();

                                using (FileStream picData = new FileStream(picPath, FileMode.Open, FileAccess.Read))
                                {
                                    pictureRun.AddPicture(picData, (int)PictureType.PNG, "image1", widthEmus, heightEmus);
                                }

                                // This double carriage returns is to create a space between images.
                                pictureRun.AddCarriageReturn();
                                pictureRun.AddCarriageReturn();
                            }
                        }

                        // SECTION DESCRIPTION
                        if (sectionNow.Description.Trim() != "")
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
                    RunStylizer(endpointRequestURL, 10, URLText);

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
                        JSONFormatter(highlightParametersList, document, jsonRequestText, data.SectionNumber);
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

                    JSONFormatter(highlightParametersList, document, jsonResponseText, data.SectionNumber);
                }

                // Create an docx. file and writes the document content into it
                using (FileStream fs = new($"{resultFolder}\\{titleText}.docx", FileMode.Create, FileAccess.Write))
                {
                    document.Write(fs);
                }
            }
        }

        private static void CreateApplicationBasicStructure(string baseFolder, string resultFolder, string picturesFolder)
        {
            Console.WriteLine("-- BASIC APPLICATION STRUCTURE -- \n");
            try
            {
                // Checking if the folder structure already exists.
                if (Directory.Exists(resultFolder))
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
                    File.WriteAllText($"{baseFolder}\\Input_Data.txt", "1;Method Example;https://url.com;\"{\"parameter1\":\"value1\"}\";\"{\"parameter2\":\"value2\"}\"\n2;Method Example2;https://url2.com;\"{\"parameter1\":\"value1\"}\";\"{\"parameter2\":\"value2\"}\"", Encoding.UTF8);
                    Console.WriteLine($"[INFO] Input file 'Input_Data.txt' has been created inside the folder, with a example inside it.");
                }
                else
                {
                    Console.WriteLine($"[INFO] Input file already exist, please put the data inside it!");
                }

                // Here the application will verify if the highlight file don't exists and if is true another one will be created.
                if (!File.Exists($"{baseFolder}\\Highlight_Parameters.txt"))
                {
                    File.WriteAllText($"{baseFolder}\\Highlight_Parameters.txt", "parameter_example1;;;\nexample2;;;\nparameter3;;;", Encoding.UTF8);
                    Console.WriteLine($"[INFO] Highlight file 'Highlight_Parameters.txt' has been created inside the folder, with a example inside it.");
                }
                else
                {
                    Console.WriteLine($"[INFO] Highlight file already exist, please put the parameter names in it!");
                }

                // This file will contain information about the section in the document (number, title and description)
                if (!File.Exists($"{baseFolder}\\Section_Information.txt"))
                {
                    File.WriteAllText($"{baseFolder}\\Section_Information.txt", "1;section title1;this is the section title1\n2;section json request;this is the section request", Encoding.UTF8);
                    Console.WriteLine($"[INFO] Section data file 'Section_Information.txt' has been created inside the folder, with a example inside it.");
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

        private static void InputFilesValidation(string baseFolder, string picturesFolder, out string[] picturesList, out List<InputData> dataList, out List<HighlightParameters> highlightParametersList, out List<SectionProperties> sectionList)
        {
            Console.WriteLine("\n\n-- INPUT FILES VALIDATION --\n");

            // Initializing key variables.
            dataList = Enumerable.Empty<InputData>().ToList();
            sectionList = Enumerable.Empty<SectionProperties>().ToList();
            highlightParametersList = Enumerable.Empty<HighlightParameters>().ToList();
            picturesList = [];

            FileStreamOptions options = new() { Access = FileAccess.Read, Mode = FileMode.Open, Options = FileOptions.None };

            try
            {
                // Every line of the file will be transformed in a instance of InputData so the values can be accessed as a parameter.
                using (StreamReader streamInputData = new($"{baseFolder}\\Input_Data.txt", Encoding.UTF8, false, options))
                {
                    string fileLine;
                    int inputFileLineCounter = 1;
                    int inputFileSectionNumber = 0;

                    // Each line is stored inside the 'fileLine' variable so it can be analyzed.
                    while ((fileLine = streamInputData.ReadLine()) != null)
                    {
                        string[] dataFields = fileLine.Split(';');

                        // Some field validations before proceed with the creation of the document.
                        if (dataFields.Any(field => field.Trim().Equals("")) == true)
                        {
                            Console.WriteLine($"[ERROR | LINE {inputFileLineCounter}] One of the fields in this line are empty in the file 'Input_Data.txt'! All field are required for the document, check all lines on the file and run the application again.");
                            Environment.Exit(0);
                        }
                        else if (int.TryParse(dataFields[0], out inputFileSectionNumber) == false)
                        {
                            Console.WriteLine($"[ERROR | LINE {inputFileLineCounter}] The first field of the line must be a INTEGER NUMBER, check all lines on the file 'Input_Data.txt' and run the application again.");
                            Environment.Exit(0);
                        }

                        string methodNameWithoutSpaces = dataFields[1].Replace(" ", ""); // Remove all spaces so the Regex logic always works.

                        InputData data = new()
                        {
                            SectionNumber = inputFileSectionNumber,
                            MethodName = Regex.Replace(methodNameWithoutSpaces, "([A-Z])(?![A-Z])", " $1").ToUpper().TrimStart(), // Separates each word in the field.
                            URL = dataFields[2],
                            Request = dataFields[3],
                            Response = dataFields[4]
                        };

                        dataList.Add(data);
                        inputFileLineCounter++;
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

            Console.WriteLine($"[INFO] All {dataList.Count} lines in 'Input_Data.txt' was read by the application!");

            try
            {
                // Every line of the file will be transformed in a instance of SectionProperties so the values can be accessed as a parameter.
                using (StreamReader streamSectionFile = new($"{baseFolder}\\Section_Information.txt", Encoding.UTF8, false, options))
                {
                    string sectionFileLine;
                    int sectionFileLineCounter = 1;
                    int sectionNumber = 0;

                    // Each line is stored inside the 'fileLine' variable so it can be analyzed.
                    while ((sectionFileLine = streamSectionFile.ReadLine()) != null)
                    {
                        string[] sectionProperties = sectionFileLine.Split(';');

                        // Some field validations before proceed with the creation of the document.
                        if (sectionProperties[0].Trim() == "" || sectionProperties[1].Trim() == "")
                        {
                            Console.WriteLine($"\n[ERROR | LINE {sectionFileLineCounter}] Section Number and / or Section Title on file 'Section_Information.txt' are empty and both are need for the application! Check all lines on the file and run the application again.");
                            Environment.Exit(0);
                        }
                        else if (int.TryParse(sectionProperties[0], out sectionNumber) == false)
                        {
                            Console.WriteLine($"\n[ERROR | LINE {sectionNumber}] The first field of the line must be a INTEGER NUMBER, check all lines on the file 'Section_Information.txt' and run the application again.");
                            Environment.Exit(0);
                        }

                        SectionProperties section = new()
                        {
                            SectionNumber = sectionNumber,
                            SectionTitle = sectionProperties[1],
                            Description = sectionProperties[2]
                        };

                        sectionList.Add(section);
                        sectionFileLineCounter++;
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

            if (picturesList.Length > 0)
            {
                Console.WriteLine($"[INFO] {dataList.Count} images were detected in the 'Pictures' folder!");
            }
            else
            {
                Console.WriteLine($"[INFO] Nothing was found in the 'Pictures' folder!");
            }

            // Retrieving the parameters name list in the .txt file that the user wants to highlight in the document.
            try
            {
                // Every line of the file will be transformed in a instance of SectionProperties so the values can be accessed as a parameter.
                using (StreamReader streamHighlightFile = new($"{baseFolder}\\Highlight_Parameters.txt", Encoding.UTF8, false, options))
                {
                    string highlightParametersFileLine;

                    // Each line is stored inside the 'fileLine' variable so it can be analyzed.
                    while ((highlightParametersFileLine = streamHighlightFile.ReadLine()) != null)
                    {
                        string[] highlightParameterProperties = highlightParametersFileLine.Split(';');

                        int highlightCode = 1;

                        if (highlightParameterProperties[1].Trim() != "" || highlightParameterProperties[2].Trim() != "" || highlightParameterProperties[3].Trim() != "")
                        {
                            highlightCode = 2;
                        }

                        //// Some field validations before proceed with the creation of the document.
                        //if (sectionProperties[0].Trim() == "" || sectionProperties[1].Trim() == "")
                        //{
                        //    Console.WriteLine($"\n[ERROR | LINE {sectionFileLineCounter}] Section Number and / or Section Title on file 'Section_Information.txt' are empty and both are need for the application! Check all lines on the file and run the application again.");
                        //    Environment.Exit(0);
                        //}
                        //else if (int.TryParse(sectionProperties[0], out sectionNumber) == false)
                        //{
                        //    Console.WriteLine($"\n[ERROR | LINE {sectionNumber}] The first field of the line must be a INTEGER NUMBER, check all lines on the file 'Section_Information.txt' and run the application again.");
                        //    Environment.Exit(0);
                        //}

                        HighlightParameters highlightParameter = new()
                        {
                            ParameterName = highlightParameterProperties[0],
                            SectionReferenceNumber = highlightParameterProperties[1].Trim() != "" ? int.Parse(highlightParameterProperties[1]) : null,
                            ParameterReferenceName = highlightParameterProperties[2].Trim() != "" ? highlightParameterProperties[2] : null,
                            ParameterReferenceValue = highlightParameterProperties[3].Trim() != "" ? highlightParameterProperties[3] : null,
                            HighlighCode = highlightCode
                        };

                        highlightParametersList.Add(highlightParameter);
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
            //string highlightFile = File.ReadAllText($"{baseFolder}\\HighlightParameters.txt");
            //string[] highlightSplit = highlightFile.Split("\r\n");
            //highlightParameters.ParametersList = highlightSplit;

            //if (highlightParameters.ParametersList.Length > 0)
            //{
            //    Console.WriteLine($"[INFO] {highlightParameters.ParametersList.Length} different parameters will be highlighted in the document!");
            //}
            //else
            //{
            //    Console.WriteLine($"[INFO] No highlight will be needed in the document");
            //}
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

        private static void HighlightRun(List<HighlightParameters> highlightParameters, string line, XWPFRun run, int sectionNumber)
        {
            string[] parameterKeyValue = line.Split(':');

            if(parameterKeyValue.Length >= 2)
            {
                string adjustedParameterName = parameterKeyValue[0].Replace("\"", "").Trim();
                string adjustedParameterValue = parameterKeyValue[1].Replace("\"", "").Replace(",","").Trim();
                //string adjustedParameterName = parameterKeyValue[0];
                //string adjustedParameterValue = parameterKeyValue[1];

                var teste2 = highlightParameters.Where(hp => hp.SectionReferenceNumber == null);
                var teste = highlightParameters.Where(hp => hp.SectionReferenceNumber == sectionNumber);

                if (teste2.Count() > 0)
                {
                    foreach (var hp2 in teste2)
                    {
                        if (hp2.ParameterName == adjustedParameterName)
                        {
                            run.GetCTR().AddNewRPr().highlight = new CT_Highlight
                            {
                                val = ST_HighlightColor.yellow
                            };
                        }
                    }
                }


                if (teste.Count() > 0)
                {
                    foreach (var hp in teste)
                    {
                        if (hp.HighlighCode == 2)
                        {
                            if (hp.ParameterReferenceName == adjustedParameterName && hp.ParameterReferenceValue == adjustedParameterValue)
                            {
                                hp.HighlighCode = 3;
                            }
                        }
                        else if (hp.HighlighCode == 3)
                        {
                            if (hp.ParameterName == adjustedParameterName)
                            {
                                run.GetCTR().AddNewRPr().highlight = new CT_Highlight
                                {
                                    val = ST_HighlightColor.yellow
                                };

                                hp.HighlighCode = 4;
                            }
                        }
                    }
                }

            }
            //bool highlightCondition = Array.Exists(highlightParameters.ParametersList, name => name.Equals(adjustedParameterName));

            //if (highlightCondition == true)
            //{
            //    run.GetCTR().AddNewRPr().highlight = new CT_Highlight
            //    {
            //        val = ST_HighlightColor.yellow
            //    };
            //}
        }

        // Apply indentation to the raw JSON string and return to be used in the document
        public static string PrettyJson(string unPrettyJson)
        {
            string jsonWithoutDoubleQuotation = unPrettyJson.Replace("\"\"", "\""); // Adjusting double quotation marks that appears when the application reads the JSON.
            jsonWithoutDoubleQuotation = jsonWithoutDoubleQuotation.Substring(1, jsonWithoutDoubleQuotation.Length - 2); // Ignoring the first and last quotation marks that are left overs of the Replace().

            string parsedJsonString = jsonWithoutDoubleQuotation;

            // If a JSON string does not start with curly braces, we can assume that the JSON either starts with a bracket or is broken and have none.
            if (jsonWithoutDoubleQuotation.StartsWith('{') == false)
            {
                // If the JSON does not have either brackets or curly braces, the only solution is by adding brackets at the start and end.
                if(jsonWithoutDoubleQuotation.StartsWith('[') == false)
                {
                    parsedJsonString = $"[{{{jsonWithoutDoubleQuotation}}}]";
                }

                // In this case we need to use JArray because the JSON will start with brackets.
                JArray parsedJson = new();

                try
                {
                    parsedJson = JArray.Parse(parsedJsonString);
                }
                catch (JsonException jsonException)
                {
                    Console.WriteLine($"\n[ERROR] Check all JSON strings inside the 'Input_Data.txt' because one of then is not an object and cannot be parsed to a JSON identation.\n> Error Details: {jsonException.Message}");
                    Environment.Exit(0);
                }

                return JsonConvert.SerializeObject(parsedJson, Formatting.Indented);
            }
            // If a JSON string start with curly braces
            else
            {
                JObject parsedJson = new();

                try
                {
                    parsedJson = JObject.Parse(parsedJsonString);
                }
                catch (JsonException jsonException)
                {
                    Console.WriteLine($"\n[ERROR] Check all JSON strings inside the 'Input_Data.txt' because one of then is not an object and cannot be parsed to a JSON identation.\n> Error Details: {jsonException.Message}");
                    Environment.Exit(0);
                }

                return JsonConvert.SerializeObject(parsedJson, Formatting.Indented);
            }
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

        private static void JSONFormatter(List<HighlightParameters> highlightParameters, XWPFDocument document, string jsonText, int sectionNumber)
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
                HighlightRun(highlightParameters, line, run, sectionNumber);

                // Set indentation to mimic JSON structure
                int indentationLevel = GetIndentationLevel(line);
                for (int i = 0; i < indentationLevel; i++)
                {
                    endpointJSON.IndentationFirstLine = i * 720; // 720 twips = 1/2 inch
                }
            }
        }

        private static void PrintGenericErrorException(Exception exception)
        {
            Console.WriteLine($"\n[ERROR]: An error has occurred! See details below: \n{exception.Message}");
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
        //public string[] ParametersList { get; set; }
        public string ParameterName {  get; set; }
        public int? SectionReferenceNumber { get; set; }
        public string? ParameterReferenceName { get; set; }
        public string? ParameterReferenceValue { get; set; }

        /** Highlight Codes
         *  1 = Don't have reference. => This will be used if all reference properties are null.
         *  2 = The reference was not found yet.
         *  3 = The reference was found and the parameter is ready to be highlighted.
         *  4 = Parameter was highlighted.
         */
        public int HighlighCode { get; set; }
    }

    public class SectionProperties()
    {
        public int SectionNumber { get; set; }
        public string SectionTitle { get; set; }
        public string Description { get; set; }
    }
}
