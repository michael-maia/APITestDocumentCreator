namespace APITestDocumentCreator.Enums
{
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
}
