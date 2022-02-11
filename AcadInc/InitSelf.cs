/* Кирилл Уваров 2022г. 10 февраля u.k.send@gmail.com. +79062644029
 * добавить для автозагрузки в файл:
 * c:\program files\autodesk\AutoCAD 2022\Support\en-us\acad2022doc.lsp
 * строку:
 * (command "_netload" "u:/dev/ExcelConnect/ExcelData/AcadInc/bin/Release/AcadInc.dll")
 * 
 * для нумерации блоков - https://forum.dwg.ru/showthread.php?t=144669
*/

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


}
