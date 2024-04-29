using NPOI.SS.Formula;
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

                XWPFRun titleRun = title.CreateRun();
                titleRun.SetText("TITLE TEST");

                // API endpoint name
                XWPFParagraph endpointName = document.CreateParagraph();
                endpointName.Alignment = ParagraphAlignment.LEFT;
                endpointName.VerticalAlignment = TextAlignment.CENTER;

                XWPFRun endpointNameRun = endpointName.CreateRun();
                endpointNameRun.SetText("ENDPOINT TEST");

                // Endpoint request title
                XWPFParagraph endpointRequest = document.CreateParagraph();
                endpointRequest.Alignment = ParagraphAlignment.CENTER;
                endpointRequest.VerticalAlignment = TextAlignment.CENTER;

                XWPFRun endpointRequestRun = endpointRequest.CreateRun();
                endpointRequestRun.SetText("REQUEST TEST");

                // Endpoint JSON request
                XWPFParagraph endpointJSONRequest = document.CreateParagraph();
                endpointJSONRequest.Alignment = ParagraphAlignment.LEFT;

                XWPFRun endpointJSONRequestRun = endpointJSONRequest.CreateRun();
                endpointJSONRequestRun.SetText("JSON REQUEST TEST");

                string userFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                using (FileStream fs = new ($"{userFolder}\\API_Test_Document.docx", FileMode.Create))
                {
                    document.Write(fs);
                }
            }
        }
    }
}
