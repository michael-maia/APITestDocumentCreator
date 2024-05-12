using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.XWPF.UserModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace APITestDocumentCreator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Base, pictures and result folder paths for the application, so it can read the input file and export the final .docx
            string baseFolder = $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\\API_Test_Document_Creator";
            string resultFolder = $"{baseFolder}\\Result";
            string picturesFolder = $"{baseFolder}\\Pictures";

            // Create folders and input file necessary for the application.
            CreateApplicationBasicFolderStructure(baseFolder, resultFolder, picturesFolder);

            // Registering the user input for document title.
            Console.Write("\n> What is the title of the document?\nTitle: ");
            string? titleText = Console.ReadLine();

            // Checking if the user will input a title text to avoid be empty.
            while (titleText == "")
            {
                Console.Write("[INFO] The title can't be empty!\nTitle: ");
                titleText = Console.ReadLine();
            }

            // Ask the User if the data inside the 'Input_Txt.txt' follows a specific pattern, that it's showed in the console.
            DataPatternApresentationAndVerification();

            // Retrieving all prints stored in the 'Pictures' folder.
            string[] picturesList = Directory.GetFiles(picturesFolder);

            FileStreamOptions options = new()
            {
                Access = FileAccess.Read,
                Mode = FileMode.Open,
                Options = FileOptions.None
            };

            // List that will hold every line of the file.
            List<InputData> dataList = Enumerable.Empty<InputData>().ToList();

            // Every line of the file will be transformed in a instance of InputData so the values can be accessed as a parameter.
            try
            {
                using (StreamReader streamInputData = new($"{baseFolder}\\Input_Data.txt", Encoding.Default, false, options))
                {
                    string fileLine;

                    // Each line is stored inside the 'fileLine' variable so it can be analyzed.
                    while ((fileLine = streamInputData.ReadLine()) != null)
                    {
                        string[] dataFields = fileLine.Split(';');

                        InputData data = new ()
                        {
                            SectionNumber = int.Parse(dataFields[0]),
                            SectionName = dataFields[1],
                            MethodName = Regex.Replace(dataFields[2], "([A-Z])(?![A-Z])", " $1").ToUpper(), // Separates each word in the field.
                            URL = dataFields[3],
                            Request = dataFields[4],
                            Response = dataFields[5]
                        };

                        dataList.Add(data);
                    }
                }
            }
            catch(IOException ioException)
            {
                Console.WriteLine($"\n\n[ERROR] Another program is using the file, you need to close it before run this application!\n> DETAILS: {ioException.Message}");
            }
            catch (Exception exception)
            {
                PrintGenericErrorException(exception);
            }

            // Creating the .docx document
            using (XWPFDocument document = new())
            {
                // Document title
                XWPFParagraph titleParagraph = document.CreateParagraph();
                titleParagraph.Alignment = ParagraphAlignment.CENTER;
                titleParagraph.VerticalAlignment = TextAlignment.CENTER;
                titleParagraph.BorderTop = Borders.Single;
                titleParagraph.BorderLeft = Borders.Single;
                titleParagraph.BorderRight = Borders.Single;
                titleParagraph.BorderBottom = Borders.Single;

                XWPFRun titleRun = titleParagraph.CreateRun();
                titleRun.FontFamily = "Calibri";
                titleRun.FontSize = 18;
                titleRun.IsBold = true;
                titleRun.SetText(titleText.ToUpper());

                int tempSectionNumber = 0;

                foreach(InputData data in dataList)
                {
                    if(data.SectionNumber > tempSectionNumber)
                    {
                        if(tempSectionNumber > 0)
                        {
                            // Adding a page break in every new section after the first one.
                            XWPFParagraph addBreak = document.CreateParagraph();
                            XWPFRun addBreakRun = addBreak.CreateRun();
                            addBreakRun.AddBreak(BreakType.PAGE);
                        }

                        // Document section
                        XWPFParagraph documentSection = document.CreateParagraph();
                        documentSection.Alignment = ParagraphAlignment.LEFT;
                        documentSection.VerticalAlignment = TextAlignment.CENTER;

                        XWPFRun documentSectionRun = documentSection.CreateRun();
                        documentSectionRun.FontFamily = "Calibri";
                        documentSectionRun.FontSize = 14;
                        documentSectionRun.IsBold = true;
                        documentSectionRun.Underline = UnderlinePatterns.Single;
                        documentSectionRun.SetColor("44AE2F");
                        documentSectionRun.SetText($"{data.SectionNumber} - {data.SectionName.ToUpper()}");

                        List<string> sectionPictures = new();

                        foreach(string picture in picturesList)
                        {
                            string pictureName = Path.GetFileNameWithoutExtension(picture);

                            if(pictureName.StartsWith(data.SectionNumber.ToString()) == true)
                            {
                                sectionPictures.Add(picture);
                            }
                        }

                        if(sectionPictures.Count > 0)
                        {
                            XWPFParagraph documentSectionPictures = document.CreateParagraph();
                            documentSectionPictures.Alignment = ParagraphAlignment.CENTER;
                            documentSectionPictures.VerticalAlignment = TextAlignment.CENTER;

                            int widthCentimeters = 15;
                            int heightCentimeters = 10;

                            int widthEmus = widthCentimeters * 360000;
                            int heightEmus = heightCentimeters * 360000;

                            foreach(string picPath in sectionPictures)
                            {
                                XWPFRun pictureRun = documentSectionPictures.CreateRun();

                                using (FileStream picData = new FileStream(picPath, FileMode.Open, FileAccess.Read))
                                {
                                    pictureRun.AddPicture(picData, (int)PictureType.PNG, "image1", widthEmus, heightEmus);
                                }

                                pictureRun.AddCarriageReturn();
                            }
                        }

                        tempSectionNumber = data.SectionNumber;
                    }

                    // This variables will help in the parts where we describe the request and response text
                    string jsonWithIdentation;
                    string[] separator = new[] { "\r\n", "\r", "\n" };
                    string[] lines;

                    // Endpoint request title
                    XWPFParagraph endpointRequest = document.CreateParagraph();
                    endpointRequest.Alignment = ParagraphAlignment.CENTER;
                    endpointRequest.VerticalAlignment = TextAlignment.CENTER;
                    endpointRequest.BorderTop = Borders.Single;
                    endpointRequest.BorderLeft = Borders.Single;
                    endpointRequest.BorderRight = Borders.Single;
                    endpointRequest.BorderBottom = Borders.Single;

                    XWPFRun endpointRequestRun = endpointRequest.CreateRun();
                    endpointRequestRun.FontFamily = "Calibri";
                    endpointRequestRun.FontSize = 12;
                    endpointRequestRun.IsBold = true;
                    endpointRequestRun.SetColor("297FC2");
                    endpointRequestRun.SetText($"REQUISIÇÃO - {data.MethodName}");

                    // Endpoint URL used in the request.
                    XWPFParagraph endpointRequestURL = document.CreateParagraph();
                    endpointRequestURL.Alignment = ParagraphAlignment.LEFT;
                    endpointRequestURL.VerticalAlignment = TextAlignment.CENTER;

                    XWPFRun endpointRequestURLRun = endpointRequestURL.CreateRun();
                    endpointRequestURLRun.FontFamily = "Calibri"; // Set font to maintain preformatted style
                    endpointRequestURLRun.FontSize = 10;
                    endpointRequestURLRun.SetText($"URL: {data.URL}");

                    // Endpoint JSON request
                    XWPFParagraph endpointRequestJSON = document.CreateParagraph();
                    endpointRequestJSON.Alignment = ParagraphAlignment.LEFT;

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

                        jsonWithIdentation = PrettyJson(jsonRequestText); // Formatting the JSON
                        lines = jsonWithIdentation.Split(separator, StringSplitOptions.None);

                        // Create a paragraph within the cell for each line of the JSON content
                        foreach (string line in lines)
                        {
                            XWPFParagraph endpointJSON = document.CreateParagraph();

                            XWPFRun run = endpointJSON.CreateRun();
                            run.SetText(line);
                            run.FontFamily = "Calibri"; // Set font to maintain preformatted style
                            run.FontSize = 10;

                            // Set indentation to mimic JSON structure
                            int indentationLevel = GetIndentationLevel(line);
                            for (int i = 0; i < indentationLevel; i++)
                            {
                                endpointJSON.IndentationFirstLine = i * 720; // 720 twips = 1/2 inch
                            }
                        }
                    }

                    // Endpoint response title
                    XWPFParagraph endpointResponse = document.CreateParagraph();
                    endpointResponse.Alignment = ParagraphAlignment.CENTER;
                    endpointResponse.VerticalAlignment = TextAlignment.CENTER;
                    endpointResponse.BorderTop = Borders.Single;
                    endpointResponse.BorderLeft = Borders.Single;
                    endpointResponse.BorderRight = Borders.Single;
                    endpointResponse.BorderBottom = Borders.Single;

                    XWPFRun endpointResponseRun = endpointResponse.CreateRun();
                    endpointResponseRun.FontFamily = "Calibri";
                    endpointResponseRun.FontSize = 12;
                    endpointResponseRun.IsBold = true;
                    endpointResponseRun.SetColor("ff0000");
                    endpointResponseRun.SetText($"RESPOSTA - {data.MethodName}");

                    // Endpoint JSON response
                    XWPFRun endpointResponseJSONRun = endpointResponse.CreateRun();
                    endpointResponseJSONRun.FontFamily = "Calibri"; // Set font to maintain preformatted style
                    endpointResponseJSONRun.FontSize = 10;

                    string jsonResponseText = data.Response;
                    jsonWithIdentation = PrettyJson(jsonResponseText); // Formatting the JSON
                    lines = jsonWithIdentation.Split(separator, StringSplitOptions.None);

                    // Create a paragraph within the cell for each line of the JSON content
                    foreach (string line in lines)
                    {
                        XWPFParagraph endpointJSON = document.CreateParagraph();

                        XWPFRun run = endpointJSON.CreateRun();
                        run.SetText(line);
                        run.FontFamily = "Calibri"; // Set font to maintain preformatted style
                        run.FontSize = 10;

                        // Set indentation to mimic JSON structure
                        int indentationLevel = GetIndentationLevel(line);
                        for (int i = 0; i < indentationLevel; i++)
                        {
                            endpointJSON.IndentationFirstLine = i * 720; // 720 twips = 1/2 inch
                        }
                    }
                }

                // Create an docx. file and writes the document content into it
                using (FileStream fs = new($"{resultFolder}\\API_Test_Document.docx", FileMode.Create, FileAccess.Write))
                {
                    document.Write(fs);
                }
            }
        }

        private static void PrintGenericErrorException(Exception exception)
        {
            Console.WriteLine($"[ERROR]: An error has occurred! See details below: \n{exception.Message}");
        }

        private static void DataPatternApresentationAndVerification()
        {
            Console.WriteLine("\n[INFO] Before reading the 'Input_Data.txt' file is important that the data inside it follows the pattern below, separated with semi-colon (;):");
            Console.WriteLine("1) SECTION NUMBER: Section number in which all the input file will be showed in the document, if the section has more than one method put the same number but write in the order of appeareance.");
            Console.WriteLine("2) METHOD NAME: Name of the API method that will appear in the document.");
            Console.WriteLine("3) URL: What URL was used in the request?");
            Console.WriteLine("4) REQUEST: JSON used in the request (don't need pre-formatting).");
            Console.WriteLine("5) RESPONSE: JSON received in the response (don't need pre-formatting).");

            Console.WriteLine("\n> Before you continue, check if there is data inside the input file.\nCan we proceed?");

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

        private static void CreateApplicationBasicFolderStructure(string baseFolder, string resultFolder, string picturesFolder)
        {
            try
            {
                // Checking if the folder structure already exists.
                if(Directory.Exists(resultFolder))
                {
                    Console.WriteLine($"[INFO] Basic folder already exists! No need for creation.");
                }
                else
                {
                    // Creating bot base and result folder in the same function.
                    Console.WriteLine($"[INFO] Creating basic folder strucutre in the following path: {baseFolder}");
                    Directory.CreateDirectory($"{resultFolder}");
                    Directory.CreateDirectory($"{picturesFolder}");
                }

                // Checking if the input file exists so we don't accidentally recreate the file and delete the data inside it.
                if (!File.Exists($"{baseFolder}\\Input_Data.txt"))
                {
                    // Creating the file for the user to input all data that should be read and exported to the document.
                    File.Create($"{baseFolder}\\Input_Data.txt").Close();
                    Console.WriteLine($"[INFO] Input file 'Input_Data.txt' has been created inside the folder");
                }
                else
                {
                    Console.WriteLine($"[INFO] Input file already exist, please put the data inside it!");
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

    // Auxiliary class to be able to serialize the JSON
    public class InputData()
    {
        public int SectionNumber { get; set; }
        public string SectionName { get; set; }
        public string MethodName { get; set; }
        public string URL { get; set; }
        public string Request { get; set; }
        public string Response { get; set; }
    }
}
