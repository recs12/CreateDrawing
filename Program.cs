using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SolidEdgeFramework;
using SolidEdgeCommunity;
using SolidEdgePart;
using SolidEdgeDraft;
using SolidEdgeCommunity.Extensions;

namespace NewDraft
{
    class Program
    {

        [STAThread]
        static void Main()
        {

            try
            {
                // See "Handling 'Application is Busy' and 'Call was Rejected By Callee' errors" topic.
                SolidEdgeCommunity.OleMessageFilter.Register();

                // Send Application foreground 
                // application.foregournd()

                SolidEdgeFramework.Application application = null;
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
                string fullName = null;


                Console.WriteLine("Create Drawing");
                Console.WriteLine("==============================================================");
                Console.WriteLine("- Version:                                             '0.0.1'");
                Console.WriteLine("- Author:                                                 RECS");
                Console.WriteLine("- Maintainer:                                   Slimane RECHDI");
                Console.WriteLine("- Last update:                                      2020-05-16");
                Console.WriteLine("==============================================================");
                Console.WriteLine("Create draft, press y/[Y] to proceed:");
                string resp = Console.ReadLine().ToLower();
                string anws = "y";

                if (!resp.Equals(value: anws))
                {
                    Console.WriteLine("You have exit the application.");
                }
                else
                {
                    string userName = System.Environment.UserName.ToLower();
                    Console.WriteLine($"User: {userName}");
                    // check authorizations of user.
                    string[] MyTeam = { "recs", "Slimane" };

                    Console.WriteLine("[+] user with permissions");
                    application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);
                    Console.WriteLine("[+] connected");

                    Console.WriteLine($"Part type:   {application.ActiveDocumentType}");
                    Console.WriteLine($"Version SE: {application.Name}");

                    switch (application.ActiveDocumentType)
                    {
                        case DocumentTypeConstants.igDraftDocument:
                            Console.WriteLine("Type: Draft");
                            Console.WriteLine("This document is already a draft.");
                            break;

                        // Reports
                        //J:\PTCR\_Solidedge\SE_config_files
                        case DocumentTypeConstants.igAssemblyDocument:
                            Console.WriteLine("Type: Assembly");
                            assemblyDocument = (SolidEdgeAssembly.AssemblyDocument)application.ActiveDocument;

                            fullName = assemblyDocument.FullName;
                            Console.WriteLine($"Path: {fullName}");

                            // Open draft document.(make a fonction of it)
                            seDraftDocument = (DraftDocument)application.Documents.Add("SolidEdge.DraftDocument");
                            Console.WriteLine("Draft created");

                            // little pause for solidedge.
                            application.DoIdle();

                            // Add the view of the active part.
                            modelLinks = seDraftDocument.ModelLinks;
                            modelLink = modelLinks.Add(fullName);
                            sheets = seDraftDocument.Sheets;

                            // Change the background to part.
                            sheetWithBackground = sheets.Item(2);
                            Console.WriteLine($"sheet name: {sheetWithBackground.Name}");
                            sheetWithBackground.ReplaceBackground(
                            "J:\\PTCR\\_Solidedge\\Template\\Draft (assembly).dft",
                                "Background FORMAT B"
                            );
                            //sheetWithBackground.ReplaceBackground(
                            //    "C:\\Users\\Slimane\\Desktop\\draft_macro\\Draft (assembly).dft",
                            //    "Background FORMAT B"
                            //);
                            Console.WriteLine("Background replaced.");


                            // Add the views in the drawing.
                            activeSheet = seDraftDocument.ActiveSheet;
                            drawingViews = activeSheet.DrawingViews;

                            principalView = drawingViews.AddPartView(
                                From: modelLink,
                                Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopView,
                                Scale: 0.25,
                                x: 0.150,
                                y: 0.125,
                                ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView
                            );
                            Console.WriteLine("igTopView");

                            frontView = drawingViews.AddByFold(
                                From: principalView,
                                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldUp,
                                x: 0.150,
                                y: 0.200
                            );
                            Console.WriteLine("igFoldUp");

                            rightView = drawingViews.AddByFold(
                                From: principalView,
                                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                                x: 0.260,
                                y: 0.125
                            );
                            Console.WriteLine("igFoldRight");

                            isoView = drawingViews.AddPartView(
                                From: modelLink,
                                Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopFrontRightView,
                                Scale: 0.25,
                                x: 0.300,
                                y: 0.200,
                                ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView
                            );
                            Console.WriteLine("igTopFrontRight");

                            break;


                        //partDocument 
                        case DocumentTypeConstants.igPartDocument:

                            Console.WriteLine("Type: PartDocument");
                            partDocument = (SolidEdgePart.PartDocument)application.ActiveDocument;

                            fullName = partDocument.FullName;
                            Console.WriteLine($"Path: {fullName}");

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
                            Console.WriteLine($"sheet name: {sheetWithBackground.Name}");
                            sheetWithBackground.ReplaceBackground(
                                "J:\\PTCR\\_Solidedge\\Template\\Draft (part) TC.dft",
                                "Background FORMAT B"
                            );
                            //sheetWithBackground.ReplaceBackground(
                            //    "C:\\Users\\Slimane\\Desktop\\draft_macro\\Draft (part) TC.dft",
                            //    "Background FORMAT B"
                            //);
                            Console.WriteLine("Background replaced.");


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
                            Console.WriteLine("igTopView");

                            frontView = drawingViews.AddByFold(
                                From: principalView,
                                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldUp,
                                x: 0.150,
                                y: 0.200
                            );
                            Console.WriteLine("igFoldUp");

                            rightView = drawingViews.AddByFold(
                                From: principalView,
                                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                                x: 0.260,
                                y: 0.125
                            );
                            Console.WriteLine("igFoldRight");

                            isoView = drawingViews.AddPartView(
                                From: modelLink,
                                Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopFrontRightView,
                                Scale: 0.5,
                                x: 0.300,
                                y: 0.200,
                                ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView
                            );
                            Console.WriteLine("igTopFrontRight");

                            break;

                        // SheetMetal
                        case DocumentTypeConstants.igSheetMetalDocument:

                            Console.WriteLine("Type: SheetMetal");
                            sheetMetalDocument = (SolidEdgePart.SheetMetalDocument)application.ActiveDocument;

                            fullName = sheetMetalDocument.FullName;
                            Console.WriteLine($"Path: {fullName}");

                            // Open draft document.(make a fonction of it)
                            seDraftDocument = (DraftDocument)application.Documents.Add("SolidEdge.DraftDocument");
                            Console.WriteLine("Draft created");

                            // little pause for solidedge.
                            application.DoIdle();

                            // Add the view of the active part.
                            modelLinks = seDraftDocument.ModelLinks;
                            modelLink = modelLinks.Add(fullName);
                            sheets = seDraftDocument.Sheets;

                            // Change the background to part.
                            sheetWithBackground = sheets.Item(2);
                            Console.WriteLine($"sheet name: {sheetWithBackground.Name}");
                            //sheetWithBackground.ReplaceBackground(
                            //    "C:\\Users\\Slimane\\Desktop\\draft_macro\\Draft (part) TC.dft",
                            //    "Background FORMAT B"
                            //);
                            sheetWithBackground.ReplaceBackground(
                                "J:\\PTCR\\_Solidedge\\Template\\Draft (part) TC.dft",
                                "Background FORMAT B"
                            );
                            Console.WriteLine("Background replaced.");


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
                            Console.WriteLine("igTopView");

                            frontView = drawingViews.AddByFold(
                                From: principalView,
                                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldUp,
                                x: 0.150,
                                y: 0.200
                            );
                            Console.WriteLine("igFoldUp");

                            rightView = drawingViews.AddByFold(
                                From: principalView,
                                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                                x: 0.260,
                                y: 0.125
                            );
                            Console.WriteLine("igFoldRight");

                            isoView = drawingViews.AddPartView(
                                From: modelLink,
                                Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopFrontRightView,
                                Scale: 0.5,
                                x: 0.300,
                                y: 0.200,
                                ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView
                            );
                            Console.WriteLine("igTopFrontRight");

                            break;

                        case DocumentTypeConstants.igUnknownDocument:
                            Console.WriteLine("Type: Unknown");
                            break;

                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
            finally
            {
                SolidEdgeCommunity.OleMessageFilter.Unregister();
                Console.WriteLine("Process completed successfully.");
                Console.ReadKey();
            }
        }
    }
}
