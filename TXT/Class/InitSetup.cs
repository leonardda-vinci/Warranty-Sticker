using WarrantySticker;
//using Renci.SshNet;
using System;
using System.IO;
using System.Windows.Forms;

public static class SystemSetup
{
    public static string ErrorFile, SuccessFile;
    public static bool gb_error = false;
    public static string ErrMsg, strErrMsg;
    public static long retval, ErrNumber, intRetCode;
    public static int intErrCode;
    public static string LogFile;
    public static string LogName;

    public static bool DefaultFolders()
    {
        try
        {
            if (!Directory.Exists(Application.StartupPath + @"\Success Log"))
            {
                Directory.CreateDirectory(Application.StartupPath + @"\Success Log");
            }

            if (!Directory.Exists(Application.StartupPath + @"\SuccessFile"))
            {
                Directory.CreateDirectory(Application.StartupPath + @"\SuccessFile");
            }

            SuccessFile = Application.StartupPath + @"\Success Log\" + DateTime.Now.ToString("yyyy-MM-dd") + "- SuccessFile.txt";
            if (!File.Exists(SuccessFile))
            {
                var SuccessLog = new FileInfo(SuccessFile);
                var StreamWriter = SuccessLog.CreateText();
                StreamWriter.WriteLine(DateTime.Today.ToString() + "          Batch Program Success Log");
                StreamWriter.Close();
            }

            if (!Directory.Exists(Application.StartupPath + @"\Generated Image"))
            {
                Directory.CreateDirectory(Application.StartupPath + @"\Generated Image");
            }

            if (!Directory.Exists(Application.StartupPath + @"\Error Log"))
            {
                Directory.CreateDirectory(Application.StartupPath + @"\Error Log");
            }

            ErrorFile = Application.StartupPath + @"\Error Log\" + DateTime.Now.ToString("yyyy-MM-dd") + "- ErrorFile.txt";
            if (!File.Exists(ErrorFile))
            {
                var ErrorLog = new FileInfo(ErrorFile);
                var StreamWriter = ErrorLog.CreateText();
                StreamWriter.WriteLine(DateTime.Today.ToString() + "          Batch Program Error Log");
                StreamWriter.Close();
            }

            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
            return false;
        }
    }

    public static void createFile(string text)
    {
        LogName = "UserData." + text + ".dat";
        LogFile = Application.StartupPath + @"\UploadFile\" + LogName;
        if (!File.Exists(LogFile))
        {
            var TimeLog = new FileInfo(LogFile);
            var StreamWriter = TimeLog.CreateText();
            StreamWriter.Close();
        }
    }


    public static void ErrorAppend(string text)
    {
        var ErrorLog = new FileInfo(ErrorFile);
        var StreamWriter = ErrorLog.AppendText();
        StreamWriter.WriteLine(DateTime.Now.TimeOfDay.ToString() + "        " + text);
        StreamWriter.Close();
        gb_error = true;
    }

    public static void SuccessAppend(string text)
    {
        try
        {
            var SuccessLog = new FileInfo(SuccessFile);
            var StreamWriter = SuccessLog.AppendText();
            StreamWriter.WriteLine(DateTime.Now.TimeOfDay.ToString() + "        " + text);
            StreamWriter.Close();
            gb_error = true;
        }
        catch (Exception er)
        {
            MessageBox.Show(er.Message.ToString());
        }

    }
}