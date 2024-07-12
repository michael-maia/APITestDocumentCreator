using APITestDocumentCreator.Classes;
using APITestDocumentCreator.Enums;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System.Text.RegularExpressions;
using System.Text;
using System.Resources;

namespace APITestDocumentCreator
{
    public static class ProgramExtensions
    {
        public static void CreateApplicationBasicStructure(ResourceManager resourceManager, string baseFolder, string resultFolder, string picturesFolder, bool exampleRequested, string[] fileNamesList)
        {
            Console.WriteLine("======================================\n" +
                              resourceManager.GetString("BasicApplicationStructureTitle") +
                              "======================================");
            Console.WriteLine(resourceManager.GetString("BasicApplicationStructureDescription"));

            try
            {
                // Checking if the folder structure already exists.
                if (Directory.Exists(resultFolder))
                {
                    //Console.WriteLine($"[INFO] Basic folder structure already exists! No need for creation. The path is {baseFolder}");
                    Console.WriteLine(resourceManager.GetString("BasicApplicationStructureExists"), baseFolder);
                }
                else
                {
                    // Creating all basic folders.
                    //Console.WriteLine($"[INFO] Creating basic folder strucutre in the following path: {baseFolder}");
                    Console.WriteLine(resourceManager.GetString("BasicApplicationStructureCreated"), baseFolder);

                    // Here the application will create both base folder and result.
                    Directory.CreateDirectory(resultFolder);
                    Directory.CreateDirectory(picturesFolder);
                }
            }
            catch (Exception exception)
            {
                ProgramExtensions.PrintGenericErrorException(resourceManager, exception);
            }

            try
            {
                if (exampleRequested == true)
                {
                    // Creating a file with an example inside it so the user can run the program to check this result.
                    File.WriteAllText($"{baseFolder}\\{fileNamesList[0]}.txt", "1;Users Connected;https://system.com/api/users;\"{\"only_connected_users\":\"true\"}\";\"[{\"id\":1,\"first_name\":\"Jeanette\",\"last_name\":\"Penddreth\",\"email\":\"jpenddreth0@census.gov\",\"gender\":\"Female\",\"ip_address\":\"26.58.193.2\"},{\"id\":2,\"first_name\":\"Giavani\",\"last_name\":\"Frediani\",\"email\":\"gfrediani1@senate.gov\",\"gender\":\"Male\",\"ip_address\":\"229.179.4.212\"}]\"\n2;Users By Company;https://system.com/api/company;\"{\"company_name\":\"ARTIQ\"}\";\"[{\"_id\":\"5973782bdb9a930533b05cb2\",\"isActive\":true,\"balance\":\"$1,446.35\",\"age\":32,\"eyeColor\":\"green\",\"name\":\"LoganKeller\",\"gender\":\"male\",\"company\":\"ARTIQ\",\"email\":\"logankeller@artiq.com\",\"phone\":\"+1(952)533-2258\",\"friends\":[{\"id\":0,\"name\":\"ColonSalazar\"},{\"id\":1,\"name\":\"FrenchMcneil\"},{\"id\":2,\"name\":\"JackPaul\"}],\"favoriteFruit\":\"banana\"},{\"_id\":\"4987255bdb9a930533j50bv2\",\"isActive\":false,\"balance\":\"$10,644.27\",\"age\":40,\"eyeColor\":\"blue\",\"name\":\"JackPaul\",\"gender\":\"male\",\"company\":\"ARTIQ\",\"email\":\"jackpaul@artiq.com\",\"phone\":\"+1(952)355-3348\",\"friends\":[{\"id\":0,\"name\":\"LoganKeller\"},{\"id\":1,\"name\":\"FrenchMcneil\"},{\"id\":2,\"name\":\"CarolMartin\"}],\"favoriteFruit\":\"banana\"}]\"", Encoding.UTF8);
                    //Console.WriteLine($"[INFO] Input file '{fileNamesList[0]}.txt' has been created inside the folder, with an example inside it.");
                    Console.WriteLine(resourceManager.GetString("BasicApplicationStructureInputExampleCreated"), fileNamesList[0]);

                    File.WriteAllText($"{baseFolder}\\{fileNamesList[1]}.txt", "1;Connected Users;This section will show all users connected to our system.\n2;Users With Same Company;In this part the API returned a list of users that are from a specific company.", Encoding.UTF8);
                    //Console.WriteLine($"[INFO] Section data file '{fileNamesList[1]}.txt' has been created inside the folder, with a example inside it.");
                    Console.WriteLine(resourceManager.GetString("BasicApplicationStructureSectionExampleCreated"), fileNamesList[1]);

                    File.WriteAllText($"{baseFolder}\\{fileNamesList[2]}.txt", "email;;;\nid;1;;\nfirst_name;1;id;1\nlast_name;1;id;1\nemail;1;id;1\n_id;2;;\nisActive;2;_id;5973782bdb9a930533b05cb2\nbalance;2;_id;5973782bdb9a930533b05cb2\nage;2;_id;5973782bdb9a930533b05cb2\nname;2;_id;5973782bdb9a930533b05cb2", Encoding.UTF8);
                    //Console.WriteLine($"[INFO] Highlight file '{fileNamesList[2]}.txt' has been created inside the folder, with a example inside it.");
                    Console.WriteLine(resourceManager.GetString("BasicApplicationStructureHighlightExampleCreated"), fileNamesList[2]);
                }
                else
                {

                    // Checking if the input file don't exists so we don't accidentally recreate the file and delete the data inside it.
                    if (!File.Exists($"{baseFolder}\\{fileNamesList[0]}.txt"))
                    {
                        File.WriteAllText($"{baseFolder}\\{fileNamesList[0]}.txt", "", Encoding.UTF8);
                        //Console.WriteLine($"[INFO] Input file '{fileNamesList[0]}.txt' has been created inside the folder.");
                        Console.WriteLine(resourceManager.GetString("BasicApplicationStructureInputCreated"), fileNamesList[0]);

                    }
                    else
                    {
                        //Console.WriteLine($"[INFO] Input file already exist, please check your data inside it!");
                        Console.WriteLine(resourceManager.GetString("BasicApplicationStructureInputExists"));
                    }

                    // This file will contain information about the section in the document (number, title and description)
                    if (!File.Exists($"{baseFolder}\\{fileNamesList[1]}.txt"))
                    {
                        File.WriteAllText($"{baseFolder}\\{fileNamesList[1]}.txt", "", Encoding.UTF8);
                        //Console.WriteLine($"[INFO] Section data file '{fileNamesList[1]}.txt' has been created inside the folder.");
                        Console.WriteLine(resourceManager.GetString("BasicApplicationStructureSectionCreated"), fileNamesList[1]);
                    }
                    else
                    {
                        //Console.WriteLine($"[INFO] Section data file already exist, please review if the file is in the correct pattern!");
                        Console.WriteLine(resourceManager.GetString("BasicApplicationStructureSectionExists"));
                    }

                    // Here the application will verify if the highlight file don't exists and if is true another one will be created.
                    if (!File.Exists($"{baseFolder}\\{fileNamesList[2]}.txt"))
                    {
                        File.WriteAllText($"{baseFolder}\\{fileNamesList[2]}.txt", "", Encoding.UTF8);
                        //Console.WriteLine($"[INFO] Hightlight file '{fileNamesList[2]}.txt' has been created inside the folder.");
                        Console.WriteLine(resourceManager.GetString("BasicApplicationStructureHighlightCreated"), fileNamesList[2]);
                    }
                    else
                    {
                        //Console.WriteLine($"[INFO] Highlight file already exist, please put the parameter names on it!");
                        Console.WriteLine(resourceManager.GetString("BasicApplicationStructureHighlightExists"));
                    }

                    // Adding a stand-by in the program so the user can read this section informations and then the screen will be cleared.
                    // Console.Write("\nPress any key to proceed...");
                    Console.WriteLine(resourceManager.GetString("PressAnyKeyToProceed"));
                    Console.ReadKey();
                    Console.Clear();
                }
            }
            catch (Exception exception)
            {
                ProgramExtensions.PrintGenericErrorException(resourceManager, exception);
            }
        }

        public static void DocumentBasicInformation(ResourceManager resourceManager, out string? titleText, out string? documentAuthor)
        {
            Console.WriteLine("======================================\n" +
                              resourceManager.GetString("DocumentBasicInformationTitle") +
                              "======================================");
            Console.WriteLine(resourceManager.GetString("DocumentBasicInformationDescription"));

            // Registering the user input for document title.
            Console.Write(resourceManager.GetString("DocumentBasicInformationTitleInput"));

            // This regex will allow the application to only use characters and numbers so it wont't generate an error when creating the .docx file.
            Regex regexTitleText = new("[A-Za-z0-9_-]+");

            bool titleTextValidation; // This bool variable will indicate if the title is valid or not.

            // Here the application will make some validations on the title text and will not proceed until a valid name is input.
            do
            {
                titleText = Console.ReadLine().Trim();

                if (titleText == "")
                {
                    titleTextValidation = false;
                    Console.Write(resourceManager.GetString("DocumentBasicInformationTitleInputEmptyValidation"));
                }
                else if (regexTitleText.IsMatch(titleText) == false)
                {
                    titleTextValidation = false;
                    Console.Write(resourceManager.GetString("DocumentBasicInformationTitleInputRegexValidation"));
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

        public static void DataPatternApresentationAndVerification(ResourceManager resourceManager)
        {
            Console.WriteLine("======================================\n" +
                              resourceManager.GetString("DataPatternExplanationTitle") +
                              "======================================");
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationDescription"));
            Console.WriteLine(resourceManager.GetString("PressAnyKeyToProceed"));
            Console.ReadKey();

            Console.WriteLine("------------------------------\n" +
                              "      INPUT_DATA LAYOUT\n" +
                              "------------------------------");
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationInputDataDescription"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationInputDataDescriptionSectionNumber"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationInputDataDescriptionMethodName"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationInputDataDescriptionURL"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationInputDataDescriptionRequest"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationInputDataDescriptionResponse"));
            Console.WriteLine(resourceManager.GetString("RequiredField"));
            Console.WriteLine(resourceManager.GetString("PressAnyKeyToProceed"));
            Console.ReadKey();

            Console.WriteLine("--------------------------------\n" +
                              "   SECTION_INFORMATION LAYOUT\n" +
                              "--------------------------------");
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationSectionInformationDescription"));

            Console.WriteLine(resourceManager.GetString("DataPatternExplanationSectionInformationDescriptionSectionNumber"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationSectionInformationDescriptionSectionName"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationSectionInformationDescriptionSection"));
            Console.WriteLine(resourceManager.GetString("RequiredField"));
            Console.WriteLine(resourceManager.GetString("PressAnyKeyToProceed"));
            Console.ReadKey();

            Console.WriteLine("--------------------------------\n" +
                              "  HIGHLIGHT_PARAMETERS LAYOUT\n" +
                              "--------------------------------");
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationHighlightParametersDescription"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationHighlightParametersDescriptionParameterName"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationHighlightParametersDescriptionSectionNumberReference"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationHighlightParametersDescriptionParameterNameReference"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationHighlightParametersDescriptionParameterValueReference"));
            Console.WriteLine(resourceManager.GetString("RequiredField"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationHighlightParametersDescriptionReminder"));
            Console.WriteLine(resourceManager.GetString("PressAnyKeyToProceed"));
            Console.ReadKey();

            Console.WriteLine("--------------------------------\n" +
                              resourceManager.GetString("DataPatternExplanationPicturesTitle") +
                              "--------------------------------");
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationPicturesDescription"));
            Console.WriteLine(resourceManager.GetString("DataPatternExplanationPicturesImageName"));

            Console.WriteLine(resourceManager.GetString("DataPatternValidationQuestion"));

            bool patternDecisionLoop = true;
            while (patternDecisionLoop)
            {
                Console.Write(resourceManager.GetString("DataPatternValidationQuestionOptions"));
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
                        Console.WriteLine(resourceManager.GetString("ExitingApplication"));
                        Thread.Sleep(3000);
                        Environment.Exit(0);
                        break;
                    case ConsoleKey.D2:
                        Console.WriteLine(resourceManager.GetString("ExitingApplication"));
                        Thread.Sleep(3000);
                        Environment.Exit(0);
                        break;
                    default:
                        break;
                }
            }
        }

        public static void InputFilesValidation(ResourceManager resourceManager, string baseFolder, string picturesFolder, bool exampleRequested, string[] fileNamesList, out string[] picturesList, out List<InputData> dataList, out List<HighlightParameters> highlightParametersList, out List<SectionProperties> sectionList)
        {
            if (exampleRequested == false)
            {
                Console.WriteLine("======================================\n" +
                                  resourceManager.GetString("InputFilesValidationTitle") +
                                  "======================================");
                Console.WriteLine(resourceManager.GetString("InputFilesValidationDescription"));
            }

            // Initializing key variables.
            dataList = Enumerable.Empty<InputData>().ToList();
            sectionList = Enumerable.Empty<SectionProperties>().ToList();
            highlightParametersList = Enumerable.Empty<HighlightParameters>().ToList();
            picturesList = [];
            int inputDataLastSectionNumber = 0;

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
                                inputErrorList.Add(string.Format(resourceManager.GetString("InputFilesValidationInputDataTest1"), inputFileLineCounter));
                                validationStatus = false;
                            }
                            else if (inputSectionNumberIsValid == false)
                            {
                                inputErrorList.Add(string.Format(resourceManager.GetString("InputFilesValidationInputDataTest2"), inputFileLineCounter));
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

                    // Store the last section number of the input data file so it can be later used to compare with the last number of the section information file.
                    inputDataLastSectionNumber = (dataList.Last()).SectionNumber;

                    if (exampleRequested == false)
                    {
                        Console.Write("1) INPUT_DATA STATUS: ");
                        if ((inputErrorList.Count != 0) == false)
                        {
                            Console.WriteLine("[OK]");
                        }
                        else
                        {
                            Console.WriteLine("[N-OK]");
                            foreach (string errorDetail in inputErrorList)
                            {
                                Console.WriteLine($"\t> {errorDetail}");
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                PrintGenericErrorException(resourceManager, exception);
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
                                sectionErrorList.Add(string.Format(resourceManager.GetString("InputFilesValidationSectionInformationTest1"), sectionFileLineCounter));
                                validationStatus = false;
                            }
                            else if (sectionNumberIsValid == false)
                            {
                                sectionErrorList.Add(string.Format(resourceManager.GetString("InputFilesValidationSectionInformationTest2"), sectionFileLineCounter));
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

                    if(inputDataLastSectionNumber != (sectionList.Last().SectionNumber) == true)
                    {
                        sectionErrorList.Add(string.Format(resourceManager.GetString("InputFilesValidationSectionInformationTestLastLine"), sectionList.Last().SectionNumber));
                        validationStatus = false;
                    }

                    if (exampleRequested == false)
                    {
                        Console.Write("2) SECTION_INFORMATION STATUS: ");
                        if ((sectionErrorList.Count != 0) == false)
                        {
                            Console.WriteLine("[OK]");
                        }
                        else
                        {
                            Console.WriteLine("[N-OK]");
                            foreach (string errorDetail in sectionErrorList)
                            {
                                Console.WriteLine($"\t> {errorDetail}");
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                PrintGenericErrorException(resourceManager, exception);
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
                                        highlightErrorList.Add(string.Format(resourceManager.GetString("InputFilesValidationSectionInformationTest2"), highlightFileLineCounter));
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
                        if ((highlightErrorList.Count != 0) == false)
                        {
                            Console.WriteLine("[OK]");
                        }
                        else
                        {
                            Console.WriteLine("[N-OK]");
                            foreach (string errorDetail in highlightErrorList)
                            {
                                Console.WriteLine($"\t> {errorDetail}");
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                PrintGenericErrorException(resourceManager, exception);
            }

            // Retrieving all prints stored in the 'Pictures' folder.
            picturesList = Directory.GetFiles(picturesFolder);

            if (exampleRequested == false)
            {
                Console.WriteLine("4) PICTURES STATUS: [OK]");
                if (picturesList.Length > 0)
                {
                    Console.WriteLine(resourceManager.GetString("InputFilesValidationPicturesFolderCounterMoreThanZero"), dataList.Count);
                }
                else
                {
                    Console.WriteLine(resourceManager.GetString("InputFilesValidationPicturesFolderCounterEqualZero"));
                }

                if (validationStatus == true)
                {
                    Console.WriteLine(resourceManager.GetString("InputFilesValidationPassed"));
                    Console.WriteLine(resourceManager.GetString("PressAnyKeyStartCreation"));
                    Console.ReadKey();
                    Console.Clear();
                }
                else
                {
                    Console.WriteLine(resourceManager.GetString("InputFilesValidationFailed"));
                    Console.WriteLine(resourceManager.GetString("PressAnyKeyToExitApplication"));
                    Console.ReadKey();
                    Environment.Exit(0);
                }
            }
        }

        public static void ParagraphStylizer(XWPFParagraph paragraph, ParagraphAlignment paragraphAlignment = ParagraphAlignment.LEFT, TextAlignment textAlignment = TextAlignment.CENTER, Borders borderStyle = Borders.None)
        {
            paragraph.Alignment = paragraphAlignment;
            paragraph.VerticalAlignment = textAlignment;
            paragraph.BorderTop = borderStyle;
            paragraph.BorderLeft = borderStyle;
            paragraph.BorderRight = borderStyle;
            paragraph.BorderBottom = borderStyle;
        }

        public static void RunStylizer(XWPFParagraph paragraph, int fontSize, string printText, bool bold = false, UnderlinePatterns underline = UnderlinePatterns.None, string color = "000000", string fontFamily = "Calibri", bool italic = false)
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

        public static void HighlightRun(List<HighlightParameters> highlightParameters, string line, XWPFRun run, int sectionNumber)
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
        public static int GetIndentationLevel(string line)
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

        public static void JSONFormatter(List<HighlightParameters> highlightParameters, XWPFDocument document, string jsonText, int sectionNumber)
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

        public static void PrintGenericErrorException(ResourceManager resourceManager, Exception exception)
        {
            //Console.WriteLine($"\n[ERROR]: An error has occurred! See details below: \n{exception.Message}");
            Console.WriteLine(resourceManager.GetString("PrintGenericErrorException"));
        }
    }
}
