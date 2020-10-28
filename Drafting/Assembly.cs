using System;
using SolidEdgeDraft;
using static System.Console;

namespace CreateDrawing.Drafting
{
    public class Assembly
    {
        public static void Draw(SolidEdgeFramework.Application application)
        {
            SolidEdgeAssembly.AssemblyDocument assemblyDocument = null;
            SolidEdgeDraft.DraftDocument seDraftDocument = null;
            SolidEdgeDraft.Sheets sheets = null;
            SolidEdgeDraft.Sheet activeSheet = null;
            SolidEdgeDraft.Sheet sheetWithBackground = null;
            SolidEdgeDraft.ModelLinks modelLinks = null;
            SolidEdgeDraft.ModelLink modelLink = null;
            SolidEdgeDraft.DrawingViews drawingViews = null;
            SolidEdgeDraft.DrawingView principalView = null;
            SolidEdgeDraft.DrawingView isoView = null;
            SolidEdgeDraft.PartsLists partsLists = null;
            string fullName = null;

            assemblyDocument = (SolidEdgeAssembly.AssemblyDocument)application.ActiveDocument;

            fullName = assemblyDocument.FullName;

            // Open draft document.(make a fonction of it)
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
            "J:\\PTCR\\_Solidedge\\Template\\Draft (assembly).dft",
                "Background FORMAT B"
            );
            Console.WriteLine("[+] Background replaced.");

            // Add the views in the drawing.
            activeSheet = seDraftDocument.ActiveSheet;
            drawingViews = activeSheet.DrawingViews;

            Console.WriteLine("[+] Views:");
            principalView = drawingViews.AddAssemblyView(
                From: modelLink,
                Orientation: SolidEdgeDraft.ViewOrientationConstants.igFrontView,
                Scale: 0.10,
                x: 0.150,
                y: 0.125,
                ViewType: SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView
            );
            WriteLine("\t+ TopView");

            drawingViews.AddByFold(
                From: principalView,
                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldUp,
                x: 0.150,
                y: 0.200
            );
            WriteLine("\t+ FoldUp");

            drawingViews.AddByFold(
                From: principalView,
                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                x: 0.260,
                y: 0.125
            );
            WriteLine("\t+ FoldRight");

            isoView = drawingViews.AddAssemblyView(
                From: modelLink,
                Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopFrontRightView,
                Scale: 0.10,
                x: 0.300,
                y: 0.200,
                ViewType: SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView
            );
            WriteLine("\t+ TopFrontRight");
            partsLists = seDraftDocument.PartsLists;
            partsLists.Add(
                isoView,
                "BILL_ANGLAIS",
                1,
                1
            );
        }
    }

}