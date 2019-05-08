using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
public enum mwsEnumLogLevels
{
    TYPE_FATAL = 0,
    TYPE_ERROR,
    TYPE_WARNING,
    TYPE_INFORMATION,
    TYPE_NONE,
    TYPE_DEBUG,
    TYPE_TEST
}


public class mwsClassLogMessage
{

    public mwsClassLogMessage(mwsEnumLogLevels logLevel, String productName)
    {
        pLogLevel = logLevel;
        mwsProductName = productName;
        pLogFile = Path.GetDirectoryName(Path.GetTempFileName()) + "\\MWSLogMessage.txt";
    }
    public mwsClassLogMessage(mwsEnumLogLevels logLevel, String productName, String fileName)
    {
        pLogLevel = logLevel;
        mwsProductName = productName;
        pLogFile = fileName;
    }

    /// <summary>
    /// current log level
    /// </summary>
    public mwsEnumLogLevels pLogLevel { get; set; }
    public String mwsProductName { get; set; }

    public mwsEnumLogLevels mwsStringTologlevel(String value)
    {
        mwsEnumLogLevels logLevel = mwsEnumLogLevels.TYPE_INFORMATION;
        switch (value)
        {
            case "FATAL":
                logLevel = mwsEnumLogLevels.TYPE_FATAL;
                break;

            case "ERROR":
                logLevel = mwsEnumLogLevels.TYPE_ERROR;
                break;

            case "WARNING":
                logLevel = mwsEnumLogLevels.TYPE_WARNING;
                break;

            case "INFORMATION":
                logLevel = mwsEnumLogLevels.TYPE_INFORMATION;
                break;

            case "DEBUG":
                logLevel = mwsEnumLogLevels.TYPE_DEBUG;
                break;

            case "TEST":
                logLevel = mwsEnumLogLevels.TYPE_TEST;
                break;
        }

        return logLevel;
    }
    /// <summary>
    /// Product Name
    /// </summary>
    private String iProductName
    {
        get { return this.mwsProductName; }

    }

    private String iGetUserName
    {
        //get { return sThisAddIn.Application.UserName; }
        get { return "Unknown"; }
    }

    private String iGetVersion
    {
        get
        {
            String emwVersion = String.Empty;

            emwVersion = GetType().Assembly.GetName().Version.ToString();
            return emwVersion;
        }
    }


    /// <summary>
    /// Logging file handle
    /// </summary>
    private StreamWriter _iLogStream = null;
    private StreamWriter iLogStream
    {
        get { return this._iLogStream; }
        set { this._iLogStream = value; }
    }


    /// <summary>
    /// Default log file name
    /// </summary>
    private String pLogFile
    {
        set;
        get;
    }


    /// <summary>
    /// Log file full path and name
    /// </summary>
    private String _iLogFile;
    public String mwsLogFile
    {
        get { return this._iLogFile; }
        set { this._iLogFile = value; }
    }


    public String mwsVersion()
    {
        String emwVersion = String.Empty;

        emwVersion = GetType().Assembly.GetName().Version.ToString();
        return emwVersion;
    }

    public void mwsLogMessage(mwsEnumLogLevels level, String text)
    {
        mwsLogMessage(null, level, text);
    }

    public String mwsGetLogLevel(mwsEnumLogLevels loglevel)
    {
        String returnValue = String.Empty;

        switch (this.pLogLevel)
        {
            case mwsEnumLogLevels.TYPE_ERROR:
                returnValue = "Error";
                break;
            case mwsEnumLogLevels.TYPE_FATAL:
                returnValue = "Fatal";
                break;
            case mwsEnumLogLevels.TYPE_INFORMATION:
                returnValue = "Information";
                break;
            case mwsEnumLogLevels.TYPE_NONE:
                returnValue = "None";
                break;
            case mwsEnumLogLevels.TYPE_WARNING:
                returnValue = "Warning";
                break;
            case mwsEnumLogLevels.TYPE_DEBUG:
                returnValue = "Debug";
                break;
        }
        return returnValue;
    }

    public String mwsGetLogLevel()
    {
        return mwsGetLogLevel(this.pLogLevel);
    }

    public void mwsLogMessage(StackTrace est, mwsEnumLogLevels level, String text)
    {
        StackTrace st = new StackTrace(true);
        String stringLevel = String.Empty;
        String logDateTime = DateTime.Now.ToString();

        // Check current log level
        if (pLogLevel < level) return;

        switch (level)
        {
            case mwsEnumLogLevels.TYPE_FATAL:
                stringLevel = "Fatal";
                wmsLogMessage(iProductName + " Version " + mwsVersion() + " " + logDateTime + " " + stringLevel + ": (" + st.GetFrame(1).GetMethod().Name + ") - " + text);
                break;

            case mwsEnumLogLevels.TYPE_ERROR:
                stringLevel = "Error";
                wmsLogMessage(iProductName + " Version " + mwsVersion() + " " + logDateTime + " " + stringLevel + ": (" + st.GetFrame(1).GetMethod().Name + ") - " + text);
                wmsLogMessage("Strack Trace from error:");
                for (int i = 1; i < st.FrameCount && st.GetFrame(i).GetFileLineNumber() > 0; i++)
                {
                    if (Debugger.IsAttached == false && st.GetFrame(i).GetMethod().ToString().Substring(0, 2) != "emw") break;
                    wmsLogMessage(st.GetFrame(i).GetMethod() + " - " + st.GetFrame(i).GetFileLineNumber());
                }

                if (Debugger.IsAttached && est != null)
                {
                    wmsLogMessage("Strack Trace from Exception:");
                    for (int i = est.FrameCount - 1; i >= 0 && est.GetFrame(i).GetFileLineNumber() > 0; i--)
                        wmsLogMessage(est.GetFrame(i).GetMethod() + " - " + est.GetFrame(i).GetFileLineNumber());
                }
                break;

            case mwsEnumLogLevels.TYPE_WARNING:
                stringLevel = "Warning";
                wmsLogMessage(iProductName + " " + logDateTime + " " + stringLevel + ": (" + st.GetFrame(2).GetMethod().Name + ") - " + text);
                break;

            case mwsEnumLogLevels.TYPE_INFORMATION:
                stringLevel = "Information";
                wmsLogMessage(iProductName + " " + logDateTime + " " + stringLevel + ": " + text);
                break;

            case mwsEnumLogLevels.TYPE_DEBUG:
                stringLevel = "Debuging";

                String path = st.GetFrame(2).GetMethod().Name + " at " + st.GetFrame(2).GetFileLineNumber();
                if (path.Equals("pchDebugMessage")) path = st.GetFrame(3).GetMethod().Name;

                wmsLogMessage(iProductName + " " + logDateTime + " " + stringLevel + ": (" + path + ") - " + text);
                break;
        }
    }


    private void wmsLogMessage(String text)
    {
        mwsSetLogStream();
        if (iLogStream == null) return;
        iLogStream.WriteLine(text);
        iLogStream.Flush();
    }

    public void mwsSetLogStream()
    {
        if (iLogStream == null)
        {
            iLogStream = new StreamWriter(pLogFile, false);
            mwsLogFile = pLogFile;
        }
        if (iLogStream == null)
        {
            mwsLogFile = null;
        }
    }

    public void mwsStopLogStream()
    {
        if (iLogStream != null) iLogStream.Close();
        File.Delete(mwsLogFile);
        iLogStream = null;
    }
       
    }

