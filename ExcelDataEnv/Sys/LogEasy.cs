using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;


namespace ExcelData.Sys
{
    public static  class LogEasy
    {
        
        public static void  WriteLog(string eventName, string path)
        {
            if ((eventName != string.Empty) && (path != string.Empty))
            {
                    using (StreamWriter logger = new StreamWriter(path, true))
                    {
                        logger.WriteLine(DateTime.Now.ToLongTimeString() + " - " + eventName);
                    }
            }
        }

        //private static void CreateFileLog(string pathLog)
        //{
            
        //}

        public static void DeleteFileLog(string pathLog)
        {
            if (File.Exists(pathLog))
            {
                var file = new FileInfo(pathLog);
                file.Delete();
            }
        }

        /// <summary>
        /// Размер файла в Мб.
        /// </summary>
        /// <param name="filePath">Путь к файлу.</param>
        /// <returns></returns>
        public static float GetFileSizeMb(string filePath)
        {
            // Console.WriteLine(File.Exists(curFile) ? "File exists." : "File does not exist.");
            if (File.Exists(filePath))
            {
                var info = new FileInfo(filePath);
                return (float)info.Length / 1024 / 1024;
            }
            else
            {
                return 0.0f;
            }
        }
    }

    /*
    public static class LogTesting
    {
        [CommandMethod("LogTestErr")]
        public static void LogTestErr()
        {
            // internal static readonly string PathLog = GetPathApp() + "\\<название приложения>.log";

            try
            {   
                // вызовем исключение из AutoCAD
                throw new AcEx(ErrorStatus.HandleExists,"Generated Exception from AutoCAD. Texting");
            }
            catch (AcEx e)
            {
                string eventThis =
                    "AcadExeption " + "LogTestErr " + e.Message.ToString();
                LogEasy.WriteLog(eventThis, Pathes.PathLog);
                throw;
            }
            catch (System.Exception e )
            {

                throw;
            }

        }
    }*/

}
