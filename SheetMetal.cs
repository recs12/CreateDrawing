using System;

public class Class1
{
	public Class1()
	{
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
		if (sheetMetalDocument.FlatPatternModels.Count == 1)
		{
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
	}
}
