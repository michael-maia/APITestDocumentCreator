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
            Console.WriteLine("======================================\n" +
                              "       API TEST DOCUMENT CREATOR\n" +
                              "======================================" +
                              "\n[INFO] If is your first time using this application, it's possible to create an example document to see how it works, do you wanna proceed with this example?");

            bool wasRequestedExample = false;
            bool correctButtonPressed = false;

            while (correctButtonPressed != true)
            {
                Console.Write("\n- Type '1' to create example\n- Type '2' to customize input data\n\nAwaiting user decision... ");
                ConsoleKeyInfo exampleDecision = Console.ReadKey();

                switch (exampleDecision.Key)
                {
                    case ConsoleKey.NumPad1:
                        wasRequestedExample = true;
                        correctButtonPressed = true;
                        Console.WriteLine("\n\n[INFO] User chose to proceed with example!\n");
                        Thread.Sleep(3000);
                        Console.Clear();
                        break;
                    case ConsoleKey.D1:
                        wasRequestedExample = true;
                        correctButtonPressed = true;
                        Console.WriteLine("\n\n[INFO] User chose to proceed with example!\n");
                        Thread.Sleep(3000);
                        Console.Clear();
                        break;
                    case ConsoleKey.NumPad2:
                        correctButtonPressed = true;
                        Console.WriteLine("\n\n[INFO] User will input it's own data!\n");
                        Thread.Sleep(3000);
                        break;
                    case ConsoleKey.D2:
                        correctButtonPressed = true;
                        Console.WriteLine("\n\n[INFO] User will input it's own data!\n");
                        Thread.Sleep(3000);
                        break;
                    default:
                        break;
                }
            }

            // Base, pictures and result folder paths for the application, so it can read the input file and export the final .docx
            string baseFolder = $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\\API_Test_Document_Creator";
            string resultFolder = $"{baseFolder}\\Result";
            string picturesFolder = $"{baseFolder}\\Pictures";

            string inputFileName = wasRequestedExample == true ? "Input_Data_Example" : "Input_Data";
            string sectionFileName = wasRequestedExample == true ? "Section_Information_Example" : "Section_Information";
            string highlightFileName = wasRequestedExample == true ? "Highlight_Parameters_Example" : "Highlight_Parameters";
            string[] fileNamesList = { inputFileName, sectionFileName, highlightFileName };

            // Create folders and input files necessary for the application.
            CreateApplicationBasicStructure(baseFolder, resultFolder, picturesFolder, wasRequestedExample, fileNamesList);

            // Retrieving basic document information.
            string titleText, documentAuthor;

            if (wasRequestedExample == true)
            {
                titleText = "Test Document";
                documentAuthor = "Test User";
            }
            else
            {
                DocumentBasicInformation(out titleText, out documentAuthor);

                // Ask the user if the data inside the 'Input_Txt.txt' follows a specific pattern, that it's showed in the console.
                DataPatternApresentationAndVerification();
            }

            // Data validation of every input file
            string[] picturesList; // Array that contains the path of each file found in 'Pictures' folder.
            List<InputData> dataList; // This list contains all lines in the 'Input_Data.txt' file.
            List<HighlightParameters> highlightParametersList; // Holds all JSON parameters that need highlight in the final document.
            List<SectionProperties> sectionList; // List that will inform all properties of each section of the document.

            InputFilesValidation(baseFolder, picturesFolder, wasRequestedExample, fileNamesList, out picturesList, out dataList, out highlightParametersList, out sectionList);

            // Creating the .docx document
            using (XWPFDocument document = new())
            {
                Console.WriteLine("\n======================================\n" +
                              "          DOCUMENT CREATION\n" +
                              "======================================");
                Console.WriteLine("[INFO] Starting creation process.");

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
                        List<string> sectionPictures = [];

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

                                using (FileStream picData = new(picPath, FileMode.Open, FileAccess.Read))
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
                    Console.WriteLine($"[INFO] Document created in the following path: {resultFolder}\\{titleText}.docx");
                }
            }
        }

        private static void CreateApplicationBasicStructure(string baseFolder, string resultFolder, string picturesFolder, bool exampleRequested, string[] fileNamesList)
        {
            Console.WriteLine("======================================\n" +
                              "     BASIC APPLICATION STRUCTURE\n" +
                              "======================================");
            Console.WriteLine("[INFO] In this step the application will check if the base folders and files are already created or need creation.\n");

            try
            {
                // Checking if the folder structure already exists.
                if (Directory.Exists(resultFolder))
                {
                    Console.WriteLine($"[INFO] Basic folder structure already exists! No need for creation. The path is {baseFolder}");
                }
                else
                {
                    // Creating all basic folders.
                    Console.WriteLine($"[INFO] Creating basic folder strucutre in the following path: {baseFolder}");

                    // Here the application will create both base folder and result.
                    Directory.CreateDirectory(resultFolder);
                    Directory.CreateDirectory(picturesFolder);
                }
            }
            catch (Exception exception)
            {
                PrintGenericErrorException(exception);
            }

            try
            {
                if (exampleRequested == true)
                {
                    // Creating a file with an example inside it so the user can run the program to check this result.
                    File.WriteAllText($"{baseFolder}\\{fileNamesList[0]}.txt", "1;Users Connected;https://system.com/api/users;\"{\"only_connected_users\":\"true\"}\";\"[{\"id\":1,\"first_name\":\"Jeanette\",\"last_name\":\"Penddreth\",\"email\":\"jpenddreth0@census.gov\",\"gender\":\"Female\",\"ip_address\":\"26.58.193.2\"},{\"id\":2,\"first_name\":\"Giavani\",\"last_name\":\"Frediani\",\"email\":\"gfrediani1@senate.gov\",\"gender\":\"Male\",\"ip_address\":\"229.179.4.212\"}]\"\n2;Users By Company;https://system.com/api/company;\"{\"company_name\":\"ARTIQ\"}\";\"[{\"_id\":\"5973782bdb9a930533b05cb2\",\"isActive\":true,\"balance\":\"$1,446.35\",\"age\":32,\"eyeColor\":\"green\",\"name\":\"LoganKeller\",\"gender\":\"male\",\"company\":\"ARTIQ\",\"email\":\"logankeller@artiq.com\",\"phone\":\"+1(952)533-2258\",\"friends\":[{\"id\":0,\"name\":\"ColonSalazar\"},{\"id\":1,\"name\":\"FrenchMcneil\"},{\"id\":2,\"name\":\"JackPaul\"}],\"favoriteFruit\":\"banana\"},{\"_id\":\"4987255bdb9a930533j50bv2\",\"isActive\":false,\"balance\":\"$10,644.27\",\"age\":40,\"eyeColor\":\"blue\",\"name\":\"JackPaul\",\"gender\":\"male\",\"company\":\"ARTIQ\",\"email\":\"jackpaul@artiq.com\",\"phone\":\"+1(952)355-3348\",\"friends\":[{\"id\":0,\"name\":\"LoganKeller\"},{\"id\":1,\"name\":\"FrenchMcneil\"},{\"id\":2,\"name\":\"CarolMartin\"}],\"favoriteFruit\":\"banana\"}]\"", Encoding.UTF8);
                    Console.WriteLine($"[INFO] Input file '{fileNamesList[0]}.txt' has been created inside the folder, with an example inside it.");

                    File.WriteAllText($"{baseFolder}\\{fileNamesList[1]}.txt", "1;Connected Users;This section will show all users connected to our system.\n2;Users With Same Company;In this part the API returned a list of users that are from a specific company.", Encoding.UTF8);
                    Console.WriteLine($"[INFO] Section data file '{fileNamesList[1]}.txt' has been created inside the folder, with a example inside it.");

                    File.WriteAllText($"{baseFolder}\\{fileNamesList[2]}.txt", "email;;;\nid;1;;\nfirst_name;1;id;1\nlast_name;1;id;1\nemail;1;id;1\n_id;2;;\nisActive;2;_id;5973782bdb9a930533b05cb2\nbalance;2;_id;5973782bdb9a930533b05cb2\nage;2;_id;5973782bdb9a930533b05cb2\nname;2;_id;5973782bdb9a930533b05cb2", Encoding.UTF8);
                    Console.WriteLine($"[INFO] Highlight file '{fileNamesList[2]}.txt' has been created inside the folder, with a example inside it.");
                }
                else
                {

                    // Checking if the input file don't exists so we don't accidentally recreate the file and delete the data inside it.
                    // TODO: REFACTOR THE IF STRUCTURE BELOW
                    if (!File.Exists($"{baseFolder}\\{fileNamesList[0]}.txt"))
                    {
                        File.WriteAllText($"{baseFolder}\\{fileNamesList[0]}.txt", "", Encoding.UTF8);
                        Console.WriteLine($"[INFO] Input file '{fileNamesList[0]}.txt' has been created inside the folder.");
                    }
                    else
                    {
                        Console.WriteLine($"[INFO] Input file already exist, please check your data inside it!");
                    }

                    // This file will contain information about the section in the document (number, title and description)
                    if (!File.Exists($"{baseFolder}\\{fileNamesList[1]}.txt"))
                    {
                        File.WriteAllText($"{baseFolder}\\{fileNamesList[1]}.txt", "", Encoding.UTF8);
                        Console.WriteLine($"[INFO] Section data file '{fileNamesList[1]}.txt' has been created inside the folder.");
                    }
                    else
                    {
                        Console.WriteLine($"[INFO] Section data file already exist, please review if the file is in the correct pattern!");
                    }

                    // Here the application will verify if the highlight file don't exists and if is true another one will be created.
                    if (!File.Exists($"{baseFolder}\\{fileNamesList[2]}.txt"))
                    {
                        File.WriteAllText($"{baseFolder}\\{fileNamesList[2]}.txt", "", Encoding.UTF8);
                        Console.WriteLine($"[INFO] Hightlight file '{fileNamesList[2]}.txt' has been created inside the folder.");
                    }
                    else
                    {
                        Console.WriteLine($"[INFO] Highlight file already exist, please put the parameter names on it!");
                    }

                    // Adding a stand-by in the program so the user can read this section informations and then the screen will be cleared.
                    Console.Write("\nPress any key to proceed...");
                    Console.ReadKey();
                    Console.Clear();
                }
            }
            catch (Exception exception)
            {
                PrintGenericErrorException(exception);
            }
        }

        private static void DocumentBasicInformation(out string? titleText, out string? documentAuthor)
        {
            Console.WriteLine("======================================\n" +
                              "          DOCUMENT PROPERTIES\n" +
                              "======================================");
            //Console.WriteLine("\n-- DOCUMENT PROPERTIES --");
            Console.WriteLine("[INFO] This step is to input some informations about the properties of the document.\n");

            // Registering the user input for document title.
            Console.Write("> What is the title of the document? (The title will be the result file name too)\nTitle: ");

            // This regex will allow the application to only use characters and numbers so it wont't generate an error when creating the .docx file.
            Regex regexTitleText = new("[A-Za-z0-9_-]+");

            bool titleTextValidation = false; // This bool variable will indicate if the title is valid or not.

            // Here the application will make some validations on the title text and will not proceed until a valid name is input.
            do
            {
                titleText = Console.ReadLine().Trim();

                if (titleText == "")
                {
                    titleTextValidation = false;
                    Console.Write("\n[INFO] The title can't be empty!\nPlease insert a new title: ");
                }
                else if (regexTitleText.IsMatch(titleText) == false)
                {
                    titleTextValidation = false;
                    Console.Write("\n[INFO] The title text can only contain the following:\n- NUMBERS (0-9)\n- ALPHABETICAL CHARACTERS (A-Z)\n- UNDERSCORE (_)\n- HYPHEN (-)\n\nPlease insert a new title: ");
                }
                else
                {
                    titleTextValidation = true;
                }
            }
            while (titleTextValidation == false);

            documentAuthor = Environment.UserName;

            Console.Clear();
        }

        private static void DataPatternApresentationAndVerification()
        {
            Console.WriteLine("======================================\n" +
                              "   DATA FILES - PATTERN EXPLANATION\n" +
                              "======================================");
            Console.WriteLine("[INFO] This is an important step of the application, where is explained how the 4 types of input files patterns are, pay attention to the details bellow and don't forget that each field in the files are separated with a semi-colon (;)\n");
            Console.WriteLine("[INFO] The files created in the base folder structure have an example showing how to application works, don't change the data inside if you wanna check how it works.\n");

            Console.WriteLine("\nPress any key to proceed to the first part...\n");
            Console.ReadKey();

            Console.WriteLine("------------------------------\n" +
                              "      INPUT_DATA LAYOUT\n" +
                              "------------------------------");
            Console.WriteLine("[INFO] The main file of the application, this is where all API request / response logs should be stored.\n");

            Console.WriteLine("1) SECTION NUMBER*: Section number in which all the input will be showed in the document, if the section has more than one input line, apply the same number but write in the order of appeareance.");
            Console.WriteLine("2) METHOD NAME*: Name of the API endpoint that will appear in the document close to the JSON request and response.");
            Console.WriteLine("3) URL*: URL used in the request to the API.");
            Console.WriteLine("4) REQUEST*: JSON used in the request (don't need pre-formatting).");
            Console.WriteLine("5) RESPONSE*: JSON received in the response (don't need pre-formatting).");

            Console.WriteLine("\n* - Required field.");

            Console.WriteLine("\nPress any key to proceed to the next part...\n");
            Console.ReadKey();

            Console.WriteLine("--------------------------------\n" +
                              "   SECTION_INFORMATION LAYOUT\n" +
                              "--------------------------------");
            Console.WriteLine("[INFO] In this file you will fill all informations about the sectors of the document.\n");

            Console.WriteLine("1) SECTION NUMBER*: Number of the section");
            Console.WriteLine("2) SECTION NAME*: The name used to define the section.");
            Console.WriteLine("3) DESCRIPTION: A text that should briefly explain what happened in the sector.");

            Console.WriteLine("\n* - Required field.");

            Console.WriteLine("\nPress any key to proceed to the next part...\n");
            Console.ReadKey();

            Console.WriteLine("--------------------------------\n" +
                              "  HIGHLIGHT_PARAMETERS LAYOUT\n" +
                              "--------------------------------");
            Console.WriteLine("[INFO] Here the objective is to read the parameters that you want to be highlighted in yellow, this part is completely optional, but if you wanna highlight anything in the JSON, some fields are required.\n");

            Console.WriteLine("1) PARAMETER NAME*: Name of the parameter inside the JSON (both request and response)");
            Console.WriteLine("2) SECTION NUMBER REFERENCE: If you want to highlight a parameter only in a specific sector, just put the number in this field.");
            Console.WriteLine("3) PARAMETER NAME REFERENCE: If the parameter that will be highlighted is inside an array in the JSON with multiple objects, an example, you can define the name of the parameter that will be used as reference.");
            Console.WriteLine("4) PARAMETER VALUE REFERENCE: If the parameter that will be highlighted is inside an array in the JSON with multiple objects, an example, you can define the value of the parameter that will be used as reference.");

            Console.WriteLine("\n* - Required field.");
            Console.WriteLine("PS: If you wanna highlight a parameter based on another parameter / value reference, all fields need to be filled!");

            Console.WriteLine("\nPress any key to proceed to the next part...\n");
            Console.ReadKey();

            Console.WriteLine("--------------------------------\n" +
                              "     PICTURES NAME PATTERN\n" +
                              "--------------------------------");
            Console.WriteLine("[INFO] This step is completely optional, but if you wanna put some images on each section of the document, a pattern must be followed.\n");
            Console.WriteLine("IMAGE NAME: Each image must be declare by indicating its section and order, an example is '2_1.png' where this image is in sector 2 and is the first image to be showed");

            Console.WriteLine("\n> Before continue to the validation of all 4 types of files, check if there is data inside the text files and that each one follows the respective patterns above.\n\nProceed?");

            bool patternDecisionLoop = true;
            while (patternDecisionLoop)
            {
                Console.Write("- Type '1' if it's OK\n- Type '2' to exit application\n\nAwaiting user decision... ");
                ConsoleKeyInfo patternDecision = Console.ReadKey();

                switch (patternDecision.Key)
                {
                    case ConsoleKey.NumPad1:
                        patternDecisionLoop = false;
                        Console.Clear();
                        break;
                    case ConsoleKey.D1:
                        patternDecisionLoop = false;
                        Console.Clear();
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

        private static void InputFilesValidation(string baseFolder, string picturesFolder, bool exampleRequested, string[] fileNamesList, out string[] picturesList, out List<InputData> dataList, out List<HighlightParameters> highlightParametersList, out List<SectionProperties> sectionList)
        {
            if (exampleRequested == false)
            {
                Console.WriteLine("======================================\n" +
                                  "        INPUT FILES VALIDATION\n" +
                                  "======================================" +
                                  "\n[INFO] Here the application will make some validations on each file to verify if it follows what is expected.\n");
            }

            // Initializing key variables.
            dataList = Enumerable.Empty<InputData>().ToList();
            sectionList = Enumerable.Empty<SectionProperties>().ToList();
            highlightParametersList = Enumerable.Empty<HighlightParameters>().ToList();
            picturesList = [];

            FileStreamOptions options = new() { Access = FileAccess.Read, Mode = FileMode.Open, Options = FileOptions.None };

            bool validationStatus = true;

            try
            {
                // Every line of the file will be transformed in a instance of InputData so the values can be accessed as a parameter.
                using (StreamReader streamInputData = new($"{baseFolder}\\{fileNamesList[0]}.txt", Encoding.UTF8, false, options))
                {
                    List<string> inputErrorList = [];
                    string fileLine;
                    int inputFileLineCounter = 1;
                    int inputFileSectionNumber = 0;

                    // Each line is stored inside the 'fileLine' variable so it can be analyzed.
                    while ((fileLine = streamInputData.ReadLine()) != null)
                    {
                        string[] dataFields = fileLine.Split(';');
                        bool inputSectionNumberIsValid = int.TryParse(dataFields[0], out inputFileSectionNumber); // It will check if the string is a valid number and change the inputSectionNumber if the TryParse is true.

                        if (exampleRequested == false)
                        {
                            // Some field validations before proceed with the creation of the document.
                            if (dataFields.Any(field => field.Trim().Equals("")) == true)
                            {
                                inputErrorList.Add($"[ERROR | LINE {inputFileLineCounter}] One of the fields in this line are empty!");
                                validationStatus = false;
                            }
                            else if (inputSectionNumberIsValid == false)
                            {
                                inputErrorList.Add($"[ERROR | LINE {inputFileLineCounter}] The first field of the line must be a INTEGER NUMBER!");
                                validationStatus = false;
                            }
                        }

                        string methodNameWithoutSpaces = dataFields[1].Replace(" ", ""); // Remove all spaces so the Regex logic always works.

                        InputData data = new()
                        {
                            SectionNumber = inputFileSectionNumber, // This number will change based on the TryParse function.
                            MethodName = Regex.Replace(methodNameWithoutSpaces, "([A-Z])(?![A-Z])", " $1").ToUpper().TrimStart(), // Separate each word that starts with a capital letter.
                            URL = dataFields[2],
                            Request = dataFields[3],
                            Response = dataFields[4]
                        };

                        dataList.Add(data);
                        inputFileLineCounter++;
                    }

                    if (exampleRequested == false)
                    {
                        Console.Write("1) INPUT_DATA STATUS: ");
                        if (inputErrorList.Any() == false)
                        {
                            Console.WriteLine("[OK]");
                        }
                        else
                        {
                            Console.WriteLine("[NOT OK]");
                            foreach (string errorDetail in inputErrorList)
                            {
                                Console.WriteLine(errorDetail);
                            }
                        }
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

            try
            {
                // Every line of the file will be transformed in a instance of SectionProperties so the values can be accessed as a parameter.
                using (StreamReader streamSectionFile = new($"{baseFolder}\\{fileNamesList[1]}.txt", Encoding.UTF8, false, options))
                {
                    List<string> sectionErrorList = [];
                    string sectionFileLine;
                    int sectionFileLineCounter = 1;
                    int sectionNumber = 0;

                    // Each line is stored inside the 'fileLine' variable so it can be analyzed.
                    while ((sectionFileLine = streamSectionFile.ReadLine()) != null)
                    {
                        string[] sectionProperties = sectionFileLine.Split(';');
                        bool sectionNumberIsValid = int.TryParse(sectionProperties[0], out sectionNumber); // It will check if the string is a valid number and change the sectionNumber if the TryParse is true.

                        if (exampleRequested == false)
                        {
                            // Some field validations before proceed with the creation of the document.
                            if (sectionProperties[0].Trim() == "" || sectionProperties[1].Trim() == "")
                            {
                                sectionErrorList.Add($"\n[ERROR | LINE {sectionFileLineCounter}] Section Number and / or Section Title are empty and both are need for the application!");
                                validationStatus = false;
                            }
                            else if (sectionNumberIsValid == false)
                            {
                                sectionErrorList.Add($"\n[ERROR | LINE {sectionNumber}] The first field of the line must be a INTEGER NUMBER!");
                                validationStatus = false;
                            }
                        }

                        SectionProperties section = new()
                        {
                            SectionNumber = sectionNumber, // This number will change based on the TryParse function.
                            SectionTitle = sectionProperties[1],
                            Description = sectionProperties[2]
                        };

                        sectionList.Add(section);
                        sectionFileLineCounter++;
                    }

                    if (exampleRequested == false)
                    {
                        Console.Write("2) SECTION_INFORMATION STATUS: ");
                        if (sectionErrorList.Any() == false)
                        {
                            Console.WriteLine("[OK]");
                        }
                        else
                        {
                            Console.WriteLine("[NOT OK]");
                            foreach (string errorDetail in sectionErrorList)
                            {
                                Console.WriteLine(errorDetail);
                            }
                        }
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

            // Retrieving the parameters name list in the .txt file that the user wants to highlight in the document.
            try
            {
                // Every line of the file will be transformed in a instance of SectionProperties so the values can be accessed as a parameter.
                using (StreamReader streamHighlightFile = new($"{baseFolder}\\{fileNamesList[2]}.txt", Encoding.UTF8, false, options))
                {
                    List<string> highlightErrorList = [];
                    string highlightParametersFileLine;
                    int highlightFileLineCounter = 1;

                    // Each line is stored inside the 'fileLine' variable so it can be analyzed.
                    while ((highlightParametersFileLine = streamHighlightFile.ReadLine()) != null)
                    {
                        string[] highlightParameterProperties = highlightParametersFileLine.Split(';');

                        // In this part the application will verify the type of reference was defined in the file and by default it will assume that is Global and don't have any references (basic type)
                        HighlightType highlightType = HighlightType.Global;
                        HighlightCode highlightCode = HighlightCode.NoReference;

                        // If SectionReferenceNumber has data or not in the file
                        if (highlightParameterProperties[1].Trim() != "")
                        {
                            // If ParameterReference is filled in the highlight file we are already define this highlight as a section and parameter dependency
                            if (highlightParameterProperties[2].Trim() != "")
                            {
                                highlightType = HighlightType.SectionAndParameterReference;

                                if (exampleRequested == false)
                                {
                                    if (highlightParameterProperties[3].Trim() == "")
                                    {
                                        highlightErrorList.Add($"[ERROR | LINE {highlightFileLineCounter}] The field 'Parameter Value' cannot be empty if the field 'Parameter Name' is filled!");
                                        validationStatus = false;
                                    }
                                }
                            }
                            // If does not have data the application will consider that is a highlight by section.
                            else
                            {
                                highlightType = HighlightType.SectionOnly;
                            }

                            highlightCode = HighlightCode.ReferenceNotFound;
                        }

                        // Creating the object that holds all features of this highlight.
                        HighlightParameters highlightParameter = new()
                        {
                            ParameterName = highlightParameterProperties[0],
                            HighlightType = highlightType,
                            HighlighCode = highlightCode,
                            SectionReferenceNumber = highlightParameterProperties[1].Trim() != "" ? int.Parse(highlightParameterProperties[1]) : null,
                            ParameterReferenceName = highlightParameterProperties[2].Trim() != "" ? highlightParameterProperties[2] : null,
                            ParameterReferenceValue = highlightParameterProperties[3].Trim() != "" ? highlightParameterProperties[3] : null
                        };

                        highlightParametersList.Add(highlightParameter);
                        highlightFileLineCounter++;
                    }

                    if (exampleRequested == false)
                    {
                        Console.Write("3) HIGHLIGHT_PARAMETERS STATUS: ");
                        if (highlightErrorList.Any() == false)
                        {
                            Console.WriteLine("[OK]");
                        }
                        else
                        {
                            Console.WriteLine("[NOT OK]");
                            foreach (string errorDetail in highlightErrorList)
                            {
                                Console.WriteLine(errorDetail);
                            }
                        }
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

            if (exampleRequested == false)
            {
                Console.WriteLine("4) PICTURES STATUS: [OK]");
                if (picturesList.Length > 0)
                {
                    Console.WriteLine($"[INFO] {dataList.Count} images were detected in the 'Pictures' folder!");
                }
                else
                {
                    Console.WriteLine($"[INFO] Nothing was found in the 'Pictures' folder!");
                }

                if (validationStatus == true)
                {
                    Console.WriteLine($"\n[INFO] All validations are OK and you can proceed with the creation of the document!\n");
                    Console.WriteLine("Press any key to start creation...\n");
                    Console.ReadKey();
                    Console.Clear();
                }
                else
                {
                    Console.WriteLine($"\n[ERROR] Some validations have failed, fix every error showed before and restart the application!");
                    Console.WriteLine("Closing program...\n");
                    Environment.Exit(0);
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

        private static void HighlightRun(List<HighlightParameters> highlightParameters, string line, XWPFRun run, int sectionNumber)
        {
            string[] parameterKeyValue = line.Split(':');

            if (parameterKeyValue.Length >= 2)
            {
                // First we adjust the parameter and value for better comparison.
                string adjustedParameterName = parameterKeyValue[0].Replace("\"", "").Trim();
                string adjustedParameterValue = parameterKeyValue[1].Replace("\"", "").Replace(",", "").Trim();

                // Check if there is any global highlight with the same name as the adjusted parameter.
                bool highlightGlobal = highlightParameters.Any(hp => hp.ParameterName == adjustedParameterName && hp.SectionReferenceNumber == null);

                if (highlightGlobal != false)
                {
                    run.GetCTR().AddNewRPr().highlight = new CT_Highlight
                    {
                        val = ST_HighlightColor.yellow
                    };
                }

                // Create a list of all highlight parameters with same section as reference.
                IEnumerable<HighlightParameters> highlightReferenceList = highlightParameters.Where(hp => hp.SectionReferenceNumber == sectionNumber);

                if (highlightReferenceList.Any())
                {
                    // In this loop the application will check some conditions to see if the parameter in the JSON is OK to be highlighted.
                    foreach (HighlightParameters hp in highlightReferenceList)
                    {
                        if (hp.ParameterName == adjustedParameterName && (hp.HighlightType == HighlightType.SectionOnly || hp.HighlighCode == HighlightCode.ReferenceFound))
                        {
                            run.GetCTR().AddNewRPr().highlight = new CT_Highlight
                            {
                                val = ST_HighlightColor.yellow
                            };

                            hp.HighlighCode = HighlightCode.ParameterHighlighted;
                        }
                        else if (hp.ParameterReferenceName == adjustedParameterName && hp.ParameterReferenceValue == adjustedParameterValue && hp.HighlighCode == HighlightCode.ReferenceNotFound)
                        {
                            hp.HighlighCode = HighlightCode.ReferenceFound;
                        }
                    }
                }
            }
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
                if (jsonWithoutDoubleQuotation.StartsWith('[') == false)
                {
                    parsedJsonString = $"[{{{jsonWithoutDoubleQuotation}}}]";
                }

                // In this case we need to use JArray because the JSON will start with brackets.
                JArray parsedJson = [];

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
                JObject parsedJson = [];

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
            string[] separator = ["\r\n", "\r", "\n"];
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

    public enum HighlightType
    {
        Global = 1,
        SectionOnly = 2,
        SectionAndParameterReference = 3
    }

    public enum HighlightCode
    {
        /** Highlight Codes
         *  1 = Don't have reference. => This will be used if all reference properties are null.
         *  2 = The reference was not found yet.
         *  3 = The reference was found and the parameter is ready to be highlighted.
         *  4 = Parameter was highlighted.
         */

        NoReference = 1,
        ReferenceNotFound = 2,
        ReferenceFound = 3,
        ParameterHighlighted = 4
    }

    public class HighlightParameters()
    {
        public string ParameterName { get; set; }
        public HighlightType HighlightType { get; set; }
        public HighlightCode HighlighCode { get; set; }
        public int? SectionReferenceNumber { get; set; }
        public string? ParameterReferenceName { get; set; }
        public string? ParameterReferenceValue { get; set; }
    }

    public class SectionProperties()
    {
        public int SectionNumber { get; set; }
        public string SectionTitle { get; set; }
        public string Description { get; set; }
    }
}
