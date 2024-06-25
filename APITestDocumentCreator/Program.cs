using APITestDocumentCreator.Classes;
using NPOI;
using NPOI.XWPF.UserModel;
using System.Drawing;

namespace APITestDocumentCreator
{
    public class Program
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
            string[] fileNamesList = [ inputFileName, sectionFileName, highlightFileName ];

            // Create folders and input files necessary for the application.
            ProgramExtensions.CreateApplicationBasicStructure(baseFolder, resultFolder, picturesFolder, wasRequestedExample, fileNamesList);

            // Retrieving basic document information.
            string titleText, documentAuthor;

            if (wasRequestedExample == true)
            {
                titleText = "Test Document";
                documentAuthor = "Test User";
            }
            else
            {
                ProgramExtensions.DocumentBasicInformation(out titleText, out documentAuthor);

                // Ask the user if the data inside the 'Input_Txt.txt' follows a specific pattern, that it's showed in the console.
                ProgramExtensions.DataPatternApresentationAndVerification();
            }

            // Data validation of every input file
            string[] picturesList; // Array that contains the path of each file found in 'Pictures' folder.
            List<InputData> dataList; // This list contains all lines in the 'Input_Data.txt' file.
            List<HighlightParameters> highlightParametersList; // Holds all JSON parameters that need highlight in the final document.
            List<SectionProperties> sectionList; // List that will inform all properties of each section of the document.

            ProgramExtensions.InputFilesValidation(baseFolder, picturesFolder, wasRequestedExample, fileNamesList, out picturesList, out dataList, out highlightParametersList, out sectionList);

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

                ProgramExtensions.ParagraphStylizer(titleParagraph, ParagraphAlignment.CENTER, TextAlignment.CENTER, Borders.Single);
                ProgramExtensions.RunStylizer(titleParagraph, 18, titleText.ToUpper(), true);

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
                        ProgramExtensions.ParagraphStylizer(documentSection, ParagraphAlignment.LEFT);

                        string sectionText = $"{sectionNow.SectionNumber} - {sectionNow.SectionTitle.ToUpper()}";
                        ProgramExtensions.RunStylizer(documentSection, 14, sectionText, true, UnderlinePatterns.Single, "44AE2F");

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
                            ProgramExtensions.ParagraphStylizer(documentSectionPictures, ParagraphAlignment.CENTER);

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
                            ProgramExtensions.ParagraphStylizer(sectionDescription, ParagraphAlignment.BOTH);

                            string sectionDescriptionText = $"Descrição: {sectionNow.Description}";
                            ProgramExtensions.RunStylizer(sectionDescription, 10, sectionDescriptionText, false, UnderlinePatterns.None, "000000", "Calibri", true);
                        }

                        tempSectionNumber = data.SectionNumber;
                    }

                    // ENDPOINT REQUEST TITLE
                    XWPFParagraph endpointRequest = document.CreateParagraph();
                    ProgramExtensions.ParagraphStylizer(endpointRequest, ParagraphAlignment.CENTER, TextAlignment.CENTER, Borders.Single);

                    string endpointRequestText = $"REQUISIÇÃO - {data.MethodName.ToUpper()}";
                    ProgramExtensions.RunStylizer(endpointRequest, 12, endpointRequestText, true, UnderlinePatterns.None, "297FC2");

                    // ENDPOINT REQUEST TITLE - URL USED
                    XWPFParagraph endpointRequestURL = document.CreateParagraph();
                    ProgramExtensions.ParagraphStylizer(endpointRequestURL);

                    string URLText = $"URL: {data.URL}";
                    ProgramExtensions.RunStylizer(endpointRequestURL, 10, URLText);

                    // ENDPOINT REQUEST - JSON TEXT
                    XWPFParagraph endpointRequestJSON = document.CreateParagraph();
                    ProgramExtensions.ParagraphStylizer(endpointRequestJSON);

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
                        ProgramExtensions.JSONFormatter(highlightParametersList, document, jsonRequestText, data.SectionNumber);
                    }

                    // ENDPOINT RESPONSE TITLE
                    XWPFParagraph endpointResponse = document.CreateParagraph();
                    ProgramExtensions.ParagraphStylizer(endpointResponse, ParagraphAlignment.CENTER, TextAlignment.CENTER, Borders.Single);

                    string responseTitleText = $"RESPOSTA - {data.MethodName.ToUpper()}";
                    ProgramExtensions.RunStylizer(endpointResponse, 12, responseTitleText, true, UnderlinePatterns.None, "FF0000");

                    // ENDPOINT RESPONSE - JSON TEXT
                    XWPFRun endpointResponseJSONRun = endpointResponse.CreateRun();
                    endpointResponseJSONRun.FontFamily = "Calibri"; // Set font to maintain preformatted style
                    endpointResponseJSONRun.FontSize = 10;

                    string jsonResponseText = data.Response;

                    ProgramExtensions.JSONFormatter(highlightParametersList, document, jsonResponseText, data.SectionNumber);
                }

                // Create an docx. file and writes the document content into it
                using (FileStream fs = new($"{resultFolder}\\{titleText}.docx", FileMode.Create, FileAccess.Write))
                {
                    document.Write(fs);
                    Console.WriteLine($"[INFO] Document created in the following path: {resultFolder}\\{titleText}.docx");
                }

                Console.WriteLine("\nPress any key to proceed with the close of this application...");
                Console.ReadKey();
            }
        }
    }
}
