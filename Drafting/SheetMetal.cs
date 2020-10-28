using System;
using SolidEdgeDraft;
using static System.Console;

namespace CreateDrawing.Drafting
{
    public class SheetMetals
    {
        public static void Draw(SolidEdgeFramework.Application application)
        {
            SolidEdgePart.SheetMetalDocument sheetMetalDocument = null;
            SolidEdgeDraft.DraftDocument seDraftDocument = null;
            SolidEdgeDraft.Sheets sheets = null;
            SolidEdgeDraft.Sheet activeSheet = null;
            SolidEdgeDraft.Sheet sheetWithBackground = null;
            SolidEdgeDraft.ModelLinks modelLinks = null;
            SolidEdgeDraft.ModelLink modelLink = null;
            SolidEdgeDraft.DrawingViews drawingViews = null;
            SolidEdgeDraft.DrawingView principalView = null;
            string fullName = null;
            sheetMetalDocument = (SolidEdgePart.SheetMetalDocument)application.ActiveDocument;

            fullName = sheetMetalDocument.FullName;

            seDraftDocument = (DraftDocument)application.Documents.Add("SolidEdge.DraftDocument");
            Console.WriteLine("[+] Draft created");

            // little pause for solidedge.
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

            principalView = drawingViews.AddSheetMetalView(
                From: modelLink,
                Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopView,
                Scale: 0.5,
                x: 0.150,
                y: 0.125,
                ViewType: SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView
            );
            WriteLine("\t+ TopView");

            var frontView = drawingViews.AddByFold(From: principalView, foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldUp, x: 0.150, y: 0.200);
            WriteLine("\t+ FoldUp");

            var rightView = drawingViews.AddByFold(From: principalView, foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldRight, x: 0.260, y: 0.125);
            WriteLine("\t+ FoldRight");

            var isoView = drawingViews.AddSheetMetalView(From: modelLink, Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopFrontRightView, Scale: 0.5, x: 0.300, y: 0.200, ViewType: SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView);
            WriteLine("\t+ TopFrontRight");

            // if flatten exists import flatten.
            if (sheetMetalDocument.FlatPatternModels.Count == 1)
            {
                var flattenView = drawingViews.AddSheetMetalView(From: modelLink, Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopView, Scale: 0.5, x: 0.300, y: 0.200, ViewType: SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalFlatView);
                WriteLine("\t+ TopView - SheetMetalFlatView");
            }
        }
    }

}