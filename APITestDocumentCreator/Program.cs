using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.XWPF.UserModel;
using System.Text;

namespace APITestDocumentCreator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Base and result folder paths for the application, so it can read the input file and export the final .docx
            string baseFolder = $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\\API_Test_Document_Creator";
            string resultFolder = $"{baseFolder}\\Result";

            // Create folders and input file necessary for the application.
            CreateApplicationBasicFolderStructure(baseFolder, resultFolder);

            // Checking user input for document title.
            Console.Write("\n> What is the title of the document?\nTitle: ");
            string titleText = Console.ReadLine();

            while (titleText == "") // Checking if the user will input a title text to avoid be empty.
            {
                Console.Write("[INFO] The title can't be empty!\nTitle: ");
                titleText = Console.ReadLine();
            }

            // Ask the User if the data inside the 'Input_Txt.txt' follows a specific pattern, that it's showed in the console.
            DataPatternApresentationAndVerification();

            FileStreamOptions options = new()
            {
                Access = FileAccess.Read,
                Mode = FileMode.Open,
                Options = FileOptions.None
            };

            List<InputData> dataList = Enumerable.Empty<InputData>().ToList();

            try
            {
                using (StreamReader streamInputData = new($"{baseFolder}\\Input_Data.txt", Encoding.UTF8, false, options))
                {
                    string fileLine;

                    while ((fileLine = streamInputData.ReadLine()) != null)
                    {
                        string[] dataFields = fileLine.Split(';');

                        InputData data = new ()
                        {
                            SectionNumber = int.Parse(dataFields[0]),
                            MethodName = dataFields[1],
                            URL = dataFields[2],
                            Request = dataFields[3],
                            Response = dataFields[4]
                        };

                        dataList.Add(data);
                        //Console.WriteLine(data);
                    }
                }
            }
            catch(IOException ioException)
            {
                Console.WriteLine($"\n\n[ERROR] Another program is using the file, you need to close it before run this application!");
            }
            catch (Exception exception)
            {
                PrintGenericErrorException(exception);
            }

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
                titleRun.FontSize = 20;
                titleRun.IsBold = true;
                titleRun.SetText(titleText);

                // Document section
                XWPFParagraph documentSection = document.CreateParagraph();
                documentSection.Alignment = ParagraphAlignment.LEFT;
                documentSection.VerticalAlignment = TextAlignment.CENTER;

                XWPFRun documentSectionRun = documentSection.CreateRun();
                documentSectionRun.FontFamily = "Calibri";
                documentSectionRun.FontSize = 14;
                documentSectionRun.IsBold = true;
                documentSectionRun.SetColor("44AE2F");
                documentSectionRun.SetText("SECTION TEST");

                // Endpoint request / response title
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
                endpointRequestRun.SetText("REQUEST TEST");

                // Endpoint JSON request / response
                string jsonText = "{\"name\":\"Adeel Solangi\",\"language\":\"Sindhi\",\"id\":\"V59OF92YF627HFY0\",\"bio\":\"Donec lobortis eleifend condimentum. Cras dictum dolor lacinia lectus vehicula rutrum. Maecenas quis nisi nunc. Nam tristique feugiat est vitae mollis. Maecenas quis nisi nunc.\",\"version\":6.1}";
                string jsonWithIdentation = PrettyJson(jsonText); // Formatting the JSON
                string[] separator = new[] { "\r\n", "\r", "\n" };
                string[] lines = jsonWithIdentation.Split(separator, StringSplitOptions.None);

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

            Console.WriteLine("\n> The data is following the pattern?");

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

        private static void CreateApplicationBasicFolderStructure(string baseFolder, string resultFolder)
        {
            try
            {
                // Checking if the all the folder structure already exist.
                if(Directory.Exists(resultFolder))
                {
                    Console.WriteLine($"[INFO] Basic folder already exists! No need for creation.");
                }
                else
                {
                    // Creating base and result folder, so both will be created in the same function
                    Console.WriteLine($"[INFO] Creating basic folder strucutre in the following path: {baseFolder}");
                    Directory.CreateDirectory($"{resultFolder}");
                }

                // Checking if the input file exists, so we don't delete accidentally recreate the file with the data inside it.
                if (!File.Exists($"{baseFolder}\\Input_Data.txt"))
                {
                    // Creating the file for the user to input all data that should read and exported to the document.
                    File.Create($"{baseFolder}\\Input_Data.txt").Close();
                    Console.WriteLine($"[INFO] Input file 'Input_Data.txt' has been created inside the folder");
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
            JObject parsedJson = JObject.Parse(unPrettyJson);
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
    public class InputData()
    {
        public int SectionNumber { get; set; }
        public string MethodName { get; set; }
        public string URL { get; set; }
        public string Request { get; set; }
        public string Response { get; set; }
    }
}
