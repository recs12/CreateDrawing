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
                SolidEdgeDraft.DrawingView flattenView = null;
                SolidEdgeDraft.PartsLists partsLists = null;
                SolidEdgeDraft.PartsList partsList = null;
                string fullName = null;


                Console.WriteLine("Create Drawing");
                Console.WriteLine("==============================================================");
                Console.WriteLine("- Version:                                             '0.0.1'");
                Console.WriteLine("- Author:                                                 RECS");
                Console.WriteLine("- Maintainer:                                   Slimane RECHDI");
                Console.WriteLine("- Last update:                                      2020-06-04");
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
                    string[] MyTeam = { "recs",
                                        "Slimane",
                                        "nunk", 
                    };
                    Console.WriteLine("[+] user with permissions");

                    application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);
                    Console.WriteLine("[+] connected");

                    Console.WriteLine($"[+] Part Type:   {application.ActiveDocumentType}");
                    Console.WriteLine($"[+] Version SolidEdge: {application.Name}");

                    switch (application.ActiveDocumentType)
                    {
                        case DocumentTypeConstants.igDraftDocument:
                            Console.WriteLine("This document is already a draft.");
                            break;

                        case DocumentTypeConstants.igAssemblyDocument:
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
                            Console.WriteLine("\t+ TopView");

                            frontView = drawingViews.AddByFold(
                                From: principalView,
                                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldUp,
                                x: 0.150,
                                y: 0.200
                            );
                            Console.WriteLine("\t+ FoldUp");

                            rightView = drawingViews.AddByFold(
                                From: principalView,
                                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                                x: 0.260,
                                y: 0.125
                            );
                            Console.WriteLine("\t+ FoldRight");

                            isoView = drawingViews.AddAssemblyView(
                                From: modelLink,
                                Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopFrontRightView,
                                Scale: 0.10,
                                x: 0.300,
                                y: 0.200,
                                ViewType: SolidEdgeDraft.AssemblyDrawingViewTypeConstants.seAssemblyDesignedView
                            );
                            Console.WriteLine("\t+ TopFrontRight");
                            partsLists = seDraftDocument.PartsLists;
                            partsList = partsLists.Add(
                                isoView,
                                "BILL_ANGLAIS",
                                1,
                                1
                            );

                            break;


                        //partDocument 
                        case DocumentTypeConstants.igPartDocument:

                            partDocument = (SolidEdgePart.PartDocument)application.ActiveDocument;

                            fullName = partDocument.FullName;

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

                            frontView = drawingViews.AddByFold(
                                From: principalView,
                                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldUp,
                                x: 0.150,
                                y: 0.200
                            );
                            Console.WriteLine("\t+ FoldUp");

                            rightView = drawingViews.AddByFold(
                                From: principalView,
                                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                                x: 0.260,
                                y: 0.125
                            );
                            Console.WriteLine("\t+ FoldRight");

                            isoView = drawingViews.AddPartView(
                                From: modelLink,
                                Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopFrontRightView,
                                Scale: 0.5,
                                x: 0.300,
                                y: 0.200,
                                ViewType: SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView
                            );
                            Console.WriteLine("\t+ TopFrontRight");

                            break;

                        // SheetMetal
                        case DocumentTypeConstants.igSheetMetalDocument:

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
                            Console.WriteLine("\t+ TopView");

                            frontView = drawingViews.AddByFold(
                                From: principalView,
                                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldUp,
                                x: 0.150,
                                y: 0.200
                            );
                            Console.WriteLine("\t+ FoldUp");

                            rightView = drawingViews.AddByFold(
                                From: principalView,
                                foldDir: SolidEdgeDraft.FoldTypeConstants.igFoldRight,
                                x: 0.260,
                                y: 0.125
                            );
                            Console.WriteLine("\t+ FoldRight");

                            isoView = drawingViews.AddSheetMetalView(
                                From: modelLink,
                                Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopFrontRightView,
                                Scale: 0.5,
                                x: 0.300,
                                y: 0.200,
                                ViewType: SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView
                            );
                            Console.WriteLine("\t+ TopFrontRight");

                            // if flatten exists import flatten. 
                            if (sheetMetalDocument.FlatPatternModels.Count==1) {
                                flattenView = drawingViews.AddSheetMetalView(
                                    From: modelLink,
                                    Orientation: SolidEdgeDraft.ViewOrientationConstants.igTopView,
                                    Scale: 0.5,
                                    x: 0.300,
                                    y: 0.200,
                                    ViewType: SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalFlatView
                                );
                                Console.WriteLine("\t+ TopView - SheetMetalFlatView");
                            }

                            break;

                        case DocumentTypeConstants.igUnknownDocument:
                            Console.WriteLine("Type: Unknown");
                            break;

                    }
                Console.WriteLine("[+]New drawing created.");
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
                Console.WriteLine("\nPress any key to exit.");
                Console.ReadKey();
            }
        }
    }
}
