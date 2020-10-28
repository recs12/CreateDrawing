using System;

public class Class1
{
	public Class1()
	{

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
	}
}
