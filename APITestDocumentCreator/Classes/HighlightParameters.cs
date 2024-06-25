using APITestDocumentCreator.Enums;

namespace APITestDocumentCreator.Classes
{
    public class HighlightParameters
    {
        public string ParameterName { get; set; }
        public HighlightType HighlightType { get; set; }
        public HighlightCode HighlighCode { get; set; }
        public int? SectionReferenceNumber { get; set; }
        public string? ParameterReferenceName { get; set; }
        public string? ParameterReferenceValue { get; set; }
    }
}
