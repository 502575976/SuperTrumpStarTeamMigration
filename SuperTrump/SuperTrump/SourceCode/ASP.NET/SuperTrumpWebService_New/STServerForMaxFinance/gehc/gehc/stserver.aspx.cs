/*
 * stserver.aspx.cs revised 2004/04/01
 *
 * Copyright (c) 2002-2004 Ivory Consulting Corporation.  All rights reserved.
 *
 *
 */

namespace gehc {

/// <summary>
/// Summary description for WebForm.
/// </summary>

public class WebForm : System.Web.UI.Page {

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

        if (inputIsOkay) {

            // Ensure there is only one request being handled at a time.
System.Threading.Monitor.Enter( xTrump );

            xmlOut = xTrump.XmlInOut( xmlIn );

System.Threading.Monitor.Exit( xTrump );
        }

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
