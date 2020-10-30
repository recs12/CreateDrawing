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


                WriteLine(
                    $"{__project__}\t--author:{__author__}\t --version:{__version__}\t --last-update :{__update__}\t\n");

                WriteLine(" Would you like make a draft of the opened document? Press y/[Y] to proceed or any key to exit... ");
                WriteLine(" (Option: Press '*' for processing documents in batch)");
                string resp = ReadLine()?.ToLower();
                const string anwserYes = "y";
                const string answerAll = "*";

                if (resp == answerAll)
                {
                    var application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);
                    System.Diagnostics.Debug.WriteLine($"SE : {application.Name}");
                    foreach (SolidEdgeDocument window in application.Documents)
                    {
                        WriteLine($"\n-Item: {window.Name}");

                        if (window.Type != DocumentTypeConstants.igDraftDocument)
                        {
                            window.Activate();
                            Drawing.DispatchDraft(application);
                        }
                    }
                }
                if (resp == anwserYes)
                {
                    var application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);
                    System.Diagnostics.Debug.WriteLine($"SE : {application.Name}");
                    if (application.ActiveDocumentType != DocumentTypeConstants.igDraftDocument)
                    {
                        Drawing.DispatchDraft(application);
                    }
                }
                else
                {
                    WriteLine("You have exit the application.");
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