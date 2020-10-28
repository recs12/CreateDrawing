using System;
using SolidEdgeDraft;
using static System.Console;


namespace CreateDrawing.Drafting
{

    public class Part
    {
        public static void Draw(SolidEdgeFramework.Application application)
        {
            SolidEdgePart.PartDocument partDocument = null;
            SolidEdgeDraft.DraftDocument seDraftDocument = null;
            SolidEdgeDraft.Sheets sheets = null;
            SolidEdgeDraft.Sheet activeSheet = null;
            SolidEdgeDraft.Sheet sheetWithBackground = null;
            SolidEdgeDraft.ModelLinks modelLinks = null;
            SolidEdgeDraft.ModelLink modelLink = null;
            SolidEdgeDraft.DrawingViews drawingViews = null;
            SolidEdgeDraft.DrawingView principalView = null;
            string fullName = null;

            partDocument = (SolidEdgePart.PartDocument)application.ActiveDocument;

            fullName = partDocument.FullName;

            // Open draft document.(make a function of it)
            seDraftDocument = (DraftDocument)application.Documents.Add("SolidEdge.DraftDocument");
            Console.WriteLine("[+] Draft created");

            // little pause for solid edge.
            application.DoIdle();

            // Add the view of the active part.
            modelLinks = seDraftDocument.ModelLinks;
            modelLink = modelLinks.Add(fullName);
            sheets = seDraftDocument.Sheets;

            // Change the background to part.
            sheetWithBackground = sheets.Item(2);
            sheetWithBackground.ReplaceBackground(
                "J:\\PTCR\\_Solidedge\\Template\\Draft (part) TC.dft",
                "Background FORMAT B"
            );
            Console.WriteLine("[+] Background replaced.");

            // Add the views in the drawing.
            activeSheet = seDraftDocument.ActiveSheet;
            drawingViews = activeSheet.DrawingViews;

            principalView = drawingViews.AddPartView(
                From: modelLink,
                Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopView,
                Scale: 0.5,
                x: 0.150,
                y: 0.125,
                ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView
            );
            Console.WriteLine("\t+ TopView");

            var frontView = drawingViews.AddByFold(From: principalView, foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldUp, x: 0.150, y: 0.200);
            Console.WriteLine("\t+ FoldUp");

            var rightView = drawingViews.AddByFold(From: principalView, foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldRight, x: 0.260, y: 0.125);
            Console.WriteLine("\t+ FoldRight");

            var isoView = drawingViews.AddPartView(From: modelLink, Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopFrontRightView, Scale: 0.5, x: 0.300, y: 0.200, ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView);
            Console.WriteLine("\t+ TopFrontRight");

        }
    }

}