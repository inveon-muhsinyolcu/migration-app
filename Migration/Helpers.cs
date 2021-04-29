using System;
using System.IO;

namespace Migration
{
    public class Helpers
    {
        public static string SetMaxLength(object value, int maxLength)
        {
            string stringValue = value.ToString();

            if (string.IsNullOrEmpty(stringValue))
                return null;

            if (stringValue.Length > maxLength)
                return stringValue.Substring(0, maxLength);

            return stringValue;
        }
        public static void LogMessage(string msg)
        {
            try
            {
                StreamWriter log;
                FileStream fileStream = null;
                DirectoryInfo logDirInfo = null;
                FileInfo logFileInfo;

                string logFilePath = @"C:\Users\Muhsin\Desktop\TODO\AyakkabiDunyasi\0503_2021_New_Migration\Log\CustomerScripts.txt";
                logFilePath = logFilePath + "Log-" + DateTime.Today.ToString("MM-dd-yyyy") + "." + "txt";
                logFileInfo = new FileInfo(logFilePath);
                logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
                if (!logDirInfo.Exists) logDirInfo.Create();
                if (!logFileInfo.Exists)
                {
                    fileStream = logFileInfo.Create();
                }
                else
                {
                    fileStream = new FileStream(logFilePath, FileMode.Append);
                }
                log = new StreamWriter(fileStream);
                log.WriteLine(msg);
                log.Close();
            }
            catch (Exception)
            {
                //Logger patlarsa artık yapacak bişey yok =)
            }
        }
    }
}
