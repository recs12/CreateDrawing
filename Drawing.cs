using DocumentTypeConstants = SolidEdgeFramework.DocumentTypeConstants;
using static System.Console;

namespace CreateDrawing
{
    class Drawing
    {

        public static void DispatchDraft(SolidEdgeFramework.Application application)
        {
            if (application.ActiveDocumentType != DocumentTypeConstants.igDraftDocument)
            {

                switch (application.ActiveDocumentType)
                {
                    case DocumentTypeConstants.igDraftDocument:
                        // Nothing needs to be done if draft.
                        WriteLine("This document is already a draft.");
                        break;

                    case DocumentTypeConstants.igAssemblyDocument:

                        // call on the class.
                        Drafting.Assembly.Draw(application);
                        break;

                    //partDocument
                    case DocumentTypeConstants.igPartDocument:

                        // call on the class.
                        Drafting.Part.Draw(application);
                        break;

                    // SheetMetal
                    case DocumentTypeConstants.igSheetMetalDocument:

                        // call on the class.
                        Drafting.SheetMetals.Draw(application);
                        break;

                    case DocumentTypeConstants.igUnknownDocument:

                        WriteLine("Document type is Unknown.");
                        break;
                }
            }
        }
    }
}
