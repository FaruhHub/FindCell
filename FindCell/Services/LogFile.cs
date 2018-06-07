using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindCell.Services
{
    public class LogFile
    {
        public static bool WriteToLog(string message, string execptionMsg = null, string path = null)
        {
            try
            {
                if (path == null)
                    path = "logs\\logFile.txt";

                string fullPath = System.IO.Path.GetFullPath(path);

                String text = Environment.NewLine + DateTime.Now.ToString() + ": " + message;
                text = execptionMsg != null ? text + " ,Error Details: " + execptionMsg : text;
                File.AppendAllText(fullPath, text, Encoding.Default);

                return true;

            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
