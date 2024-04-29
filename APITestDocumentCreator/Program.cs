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
                XWPFParagraph endpointJSONRequest = document.CreateParagraph();
                endpointJSONRequest.Alignment = ParagraphAlignment.LEFT;

                XWPFRun endpointJSONRequestRun = endpointJSONRequest.CreateRun();
                endpointRequestRun.FontFamily = "Calibri";
                endpointRequestRun.FontSize = 10;
                endpointJSONRequestRun.SetText("JSON REQUEST TEST");

                // Saves the file in the user's desktop folder for easy access
                string userFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                using (FileStream fs = new ($"{userFolder}\\API_Test_Document.docx", FileMode.Create))
                {
                    document.Write(fs);
                }
            }
        }
    }
}
