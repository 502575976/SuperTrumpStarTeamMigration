/*
 * logstserver.aspx.cs revised 2004/10/19
 * Logging wrapper for SuperTRUMP Server.
 *
 * Copyright (c) 2002-2004 Ivory Consulting Corporation.  All rights reserved.
 *
 *
 */

using System.IO;
using System.Text; 
using System.Web;

namespace loggehc {

/// <summary>
/// Summary description for WebForm.
/// </summary>

    public class WebForm : System.Web.UI.Page {

        private void DumpTextToFile( string Filename, string Content ) {
            StreamWriter sw = new StreamWriter( Filename );
            sw.Write( Content );
            sw.Flush();
            sw.Close();
        }

        private void CreateLogFilenames(out string QueryFilename, out string ResultFilename) {

            // The base filename may resemble
            // c:\inetpub\wwwroot\x-trump.net\GEHC\2004-10-19T13_15_00_
            // To this we will add a sequence number and "query.xml" or "result.xml".
            string BaseFilename = 
                HttpContext.Current.ApplicationInstance.Server.MapPath(".") // Folder.
                + '\\'
                + System.DateTime.Now.ToString("s").Replace(':', '_') // Current date & time.
                + '_';
            string TestQueryFilename;
            string TestResultFilename;
            int fileCounter = 0;

            // Check and revise the filenames until we find a pair of unused names.
            do {
                string strCounter = fileCounter.ToString("D3");
                TestQueryFilename  = BaseFilename + strCounter + "query.xml";
                TestResultFilename = BaseFilename + strCounter + "result.xml";
                ++fileCounter;
            } while (File.Exists(TestQueryFilename) || File.Exists(TestResultFilename));

            QueryFilename  = TestQueryFilename;
            ResultFilename = TestResultFilename;
        }

        // "StServer.STApplication" is the X-TRUMP ProgId for GEHC.
        private static STSERVER.STApplication xTrump = new STSERVER.STApplicationClass();

    private void Page_Load( object sender, System.EventArgs e ) {
        // Put user code to initialize the page here

        string xmlIn = "";
        bool inputIsOkay;
        string xmlOut = "";

        if ( 0 == Request.Form.Count ) {
            
            // First case: no parameters; this must be an HTTP GET.
            // Just report the server version and build number.

            xmlIn = "<SuperTrump><AppData><Version query=\"true\" /><BuildNumber query=\"true\" /></AppData></SuperTrump>";
            inputIsOkay = true;
        } else {
            string [] values = Request.Form.GetValues( 0 );
            if ( 1 == values.GetLength( 0 ) ) {

                // Second (main) case: single parameter;
                // this is a normal XML query in the form of HTTP POST.

                xmlIn = values[0];
                inputIsOkay = true;
            } else {

                // Third case: more than one parameter.
                // Why are there multiple name/value pairs?
                // This can happen if an ampersand (&) is not encoded as #26
                // in the HTTP POST.

                xmlOut = "<STWrapperError>Multiple parameters were found but only one is expected. (Perhaps were ampersands (&amp;) not encoded as \"%26\"?)</STWrapperError>";
                inputIsOkay = false;
            }
        }

        // Ensure there is only one request being handled at a time.
        System.Threading.Monitor.Enter( xTrump );

        string QueryFilename;
        string ResultFilename;
        CreateLogFilenames(out QueryFilename, out ResultFilename);
        DumpTextToFile(QueryFilename, xmlIn);

        if (inputIsOkay) {
            xmlOut = xTrump.XmlInOut( xmlIn );
        }
        
        DumpTextToFile(ResultFilename, xmlOut);

        System.Threading.Monitor.Exit( xTrump );

        Response.Write( xmlOut );
    }

    private void Page_Error( object sender, System.EventArgs e ) {

        System.Exception ex = Server.GetLastError();

        string errorMessage = ex.ToString() + "\n" + ex.StackTrace;

        WebForm.log( errorMessage, System.Diagnostics.EventLogEntryType.Error );

        Server.ClearError();
    }

    private static void log( string message, System.Diagnostics.EventLogEntryType messageType ) {

        string eventSourceName = "gehc.WebForm";
        string logName = "GEHC X-TRUMP";

        if ( !System.Diagnostics.EventLog.SourceExists( eventSourceName ) ) {
            System.Diagnostics.EventLog.CreateEventSource( eventSourceName, logName );
        }

        System.Diagnostics.EventLog eventLog = new System.Diagnostics.EventLog( logName );
        eventLog.Source = eventSourceName;

        eventLog.WriteEntry( message, messageType );
    }

#region Web Form Designer generated code

    override protected void OnInit( System.EventArgs e ) {

        // CODEGEN: This call is required by the ASP.NET Web Form Designer.
        this.InitializeComponent();
        base.OnInit( e );
    }

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent() {
        this.Load += new System.EventHandler( this.Page_Load );
    }

#endregion
}
}
