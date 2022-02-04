using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Reflection;
using Autodesk.Windows;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using ExcelData;
using ExcelData.Class;

[assembly: CommandClass(typeof(AcadInc.DialCreateSet))]

namespace AcadInc
{
    /// <summary>
    /// Запускаемый класс - точка входа.
    /// При загрузке данной dll в AutoCAD выполняется код в методе IExtensionApplication.Initialize()
    /// </summary>
    internal class InitSelf : IExtensionApplication
    {

        /// <summary>
        /// Инициализация.
        /// для запуска своих методов при загрузке dll в acad
        /// через команду _netload дописать здесь свой код 
        /// </summary>
        void IExtensionApplication.Initialize()
        {
            // Вывод данных о приложении в ком строку AutoCAD
            //InitThis.InitOne();
            // Подключение обработчиков основных событий.
            //InitThis.BasicEventHadlerlersConnect();
            // Загрузка интерфейса
            //InitThis.LoadUserInterface();

           

        }

        /// <summary>
        /// Метод, выполняемый при выгрузке плагина
        /// в нашем случае, при выгрузке экземляра acad.exe
        /// </summary>
        void IExtensionApplication.Terminate()
        {

            
        }

    }
    public static class DialCreateSet
    {


        [CommandMethod("eOpenDial")]
        public static void OpenDial()
        {
            DataExcel ED = new DataExcel();

            string[,] strTable = ED.GetDataExel();

            int rows = strTable.GetUpperBound(0) + 1;    // количество строк
            int columns = strTable.GetUpperBound(1) + 1;// strTable.Length / rows;        // количество столбцов

            Console.WriteLine($"Строк={rows}  Столбцов={columns}");
        }

    }


}
