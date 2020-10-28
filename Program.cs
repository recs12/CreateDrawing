using SolidEdgeFramework;
using System;
using static System.Console;

namespace CreateDrawing
{
    internal class Program
    {
        [STAThread]
        private static void Main()
        {

            string __version__ = "0.0.0";
            string __author__ = "recs";
            string __update__ = "2020-10-27";
            string __project__ = "CreateDrawing";

            try
            {
                // See "Handling 'Application is Busy' and 'Call was Rejected By Callee' errors" topic.
                SolidEdgeCommunity.OleMessageFilter.Register();

                SolidEdgeAssembly.AssemblyDocument assemblyDocument = null;
                SolidEdgePart.PartDocument partDocument = null;
                SolidEdgePart.SheetMetalDocument sheetMetalDocument = null;
                SolidEdgeDraft.DraftDocument seDraftDocument = null;
                SolidEdgeDraft.Sheets sheets = null;
                SolidEdgeDraft.Sheet activeSheet = null;
                SolidEdgeDraft.Sheet sheetWithBackground = null;
                SolidEdgeDraft.ModelLinks modelLinks = null;
                SolidEdgeDraft.ModelLink modelLink = null;
                SolidEdgeDraft.DrawingViews drawingViews = null;
                SolidEdgeDraft.DrawingView principalView = null;
                SolidEdgeDraft.DrawingView frontView = null;
                SolidEdgeDraft.DrawingView rightView = null;
                SolidEdgeDraft.DrawingView isoView = null;
                SolidEdgeDraft.DrawingView flattenView = null;
                SolidEdgeDraft.PartsLists partsLists = null;
                SolidEdgeDraft.PartsList partsList = null;
                string fullName = null;

                WriteLine(
                    $"{__project__}  --author: {__author__} --version: {__version__} --last-update :{__update__} ");

                WriteLine("Create draft, press y/[Y] to proceed: ");
                string resp = ReadLine().ToLower();
                string anwser = "y";

                if (resp != anwser)
                {
                    WriteLine("You have exit the application.");
                }
                else
                {
                    var application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);

                    WriteLine($"[+] Document type:   {application.ActiveDocumentType}");

                    switch (application.ActiveDocumentType)
                    {
                        case DocumentTypeConstants.igDraftDocument:
                            WriteLine("This document is already a draft.");
                            break;

                        case DocumentTypeConstants.igAssemblyDocument:
                            // call on the class.

                            break;

                        //partDocument
                        case DocumentTypeConstants.igPartDocument:

                            // call on the class.

                            break;

                        // SheetMetal
                        case DocumentTypeConstants.igSheetMetalDocument:

                            // call on the class.

                            break;

                        case DocumentTypeConstants.igUnknownDocument:
                            WriteLine("Document type is Unknown.");
                            break;
                    }
                    WriteLine("Drawing created successfully.");
                }
            }
            catch (Exception e)
            {
                WriteLine(e.Message);
                throw;
            }
            finally
            {
                SolidEdgeCommunity.OleMessageFilter.Unregister();
                WriteLine("\nPress any key to exit.");
                ReadKey();
            }
        }
    }
}