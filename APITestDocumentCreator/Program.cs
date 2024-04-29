using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.XWPF.UserModel;

namespace APITestDocumentCreator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (XWPFDocument document = new())
            {
                // Document title
                XWPFParagraph title = document.CreateParagraph();
                title.Alignment = ParagraphAlignment.CENTER;
                title.VerticalAlignment = TextAlignment.CENTER;
                title.BorderTop = Borders.Single;
                title.BorderLeft = Borders.Single;
                title.BorderRight = Borders.Single;
                title.BorderBottom = Borders.Single;

                XWPFRun titleRun = title.CreateRun();
                titleRun.FontFamily = "Calibri";
                titleRun.FontSize = 20;
                titleRun.IsBold = true;
                titleRun.SetText("TITLE TEST");

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
                string[] lines = jsonWithIdentation.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

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

                // Saves the file in the user's desktop folder for easy access
                string userFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                // Create an docx. file and writes the document content into it
                using (FileStream fs = new($"{userFolder}\\API_Test_Document.docx", FileMode.Create, FileAccess.Write))
                {
                    document.Write(fs);
                }
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
}
