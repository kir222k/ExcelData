/* Кирилл Уваров 2022г. 10 февраля. u.k.send@gmail.com. +79062644029
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Reflection;
using Autodesk.Windows;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using ExcelData;
using ExcelData.Class;
using ExcelData.Model;


namespace AcadInc
{
    public static class ExtData
    {
        //public static void WriteToExtDataExcelFileInfo ((string,string) dataToWrite)
        //{

        // https://adn-cis.org/forum/index.php?topic=8184.15
        //[CommandMethod("XRecord_test")]
        public static void WriteToExtDataExcelFileInfo((string file, string sheet) dataToWrite)
        {
           // Database db = Application.DocumentManager.MdiActiveDocument.Database;
            AcadSendMess acadSend = new AcadSendMess();

            try
            {
                #region УДАЛИТЬ ?
                //               using (Transaction trans = db.TransactionManager.StartTransaction())
                //               {
                //                   BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead); // ForRead - IN STOCK
                //                   BlockTableRecord ms = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                //                   // создадим словарь в модели, как в любом примитиве, но проверим, вдруг уже есть?
                //                   // https://www.programmerall.com/article/8393129803/
                //                   if (ms.ExtensionDictionary.IsNull)
                //                   {
                //                       ms.UpgradeOpen();
                //                       ms.CreateExtensionDictionary();
                //                       ms.DowngradeOpen();
                //                   }
                //                   #region СПРАВКА
                //                   /*
                //           if (obj.ExtensionDictionary.IsNull)//If the object has no extended dictionary, create it
                //14             {
                //15                 obj.UpgradeOpen();//Switch object to write state
                //16                 obj.CreateExtensionDictionary();//Create an extended dictionary for an object, an object can only have one extended dictionary
                //17                 obj.DowngradeOpen();//Switch the object to read state
                //18             }
                //19             //Open the extended dictionary of the object
                //20         DBDictionary dict = obj.ExtensionDictionary.GetObject(OpenMode.ForRead) as DBDictionary;
                //21             //If the specified extended record object is already contained in the extended dictionary  
                //                   */
                //                   #endregion

                //                   // откроем словарь на запись
                //                   DBDictionary dict = (DBDictionary)trans.GetObject(ms.ExtensionDictionary, OpenMode.ForWrite);
                //                   // что-то пока непонятное, ну да ладно
                //                   ResultBuffer rb = new ResultBuffer();
                //                   // какая-то переменная, кот. передаем в буфер
                //                   TypedValue tv = new TypedValue((short)DxfCode.ExtendedDataAsciiString, dataToWrite.file); // ПУТЬ к ФАЙЛУ EXCEL
                //                   // передали в буфер
                //                   rb.Add(tv);


                //                   if (!dict.Contains("samexceldatapath"))
                //                   {
                //                       Xrecord xrec = new Xrecord();
                //                       xrec.Data = rb;
                //                       dict.SetAt("samexceldatapath", xrec);

                //                       trans.AddNewlyCreatedDBObject(xrec, true);
                //                       trans.Commit();
                //                   }
                //                   else 
                //                   {
                //                       // если хзапись с ключом samexceldatapath уже есть, зайдем в нее и заменим данные
                //                       // т.е. перезапишем путь к файлу
                //                       ObjectId xrecordId = dict.GetAt("samexceldatapath");
                //                       Xrecord xrec = xrecordId.GetObject(OpenMode.ForWrite) as Xrecord;
                //                       xrec.Data = rb;

                //                       trans.Commit();
                //                   }
                //                   #region СПРАВКА
                //                                       /*
                //                                   if (!dict.Contains(xRecordSearchKey))
                //                    82             {
                //                    83                 return null;//If there is no extended record containing the specified keyword in the extended dictionary, null will be returned;
                //                    84             }

                //                    85             //First get the extended dictionary of the object or the well-known object dictionary in the graph, and then get the extended record to be queried in the dictionary
                //                    86             ObjectId xrecordId = dict.GetAt(xRecordSearchKey);//Get the id of the extended record object
                //                    87             Xrecord xrecord = xrecordId.GetObject(OpenMode.ForRead) as Xrecord;//Get extended record object according to id
                //                    88             TypedValueList values = xrecord.Data;
                //                    89             return values;//The values ​​array should be sequential
                //                                       */
                //                   #endregion




                //                   trans.Dispose();
                //               }
                #endregion

                // путь
                WriteToExtDataModel("samexceldatapath", dataToWrite.file);

                // лист
                WriteToExtDataModel("samexceldatasheet", dataToWrite.sheet);
            }
            catch (System.Exception e)
            {
                acadSend.SendStringDebugStars(e.Message);
               // throw;
            }
        }

        // универс. метод
        // запись в расш.данные модели
        private static void WriteToExtDataModel (string xKey, string xData)
        {
            Database db = Application.DocumentManager.MdiActiveDocument.Database;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead); // ForRead - IN STOCK
                BlockTableRecord ms = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                // создадим словарь в модели, как в любом примитиве, но проверим, вдруг уже есть?
                // https://www.programmerall.com/article/8393129803/
                if (ms.ExtensionDictionary.IsNull)
                {
                    ms.UpgradeOpen();
                    ms.CreateExtensionDictionary();
                    ms.DowngradeOpen();
                }
                #region СПРАВКА
                /*
        if (obj.ExtensionDictionary.IsNull)//If the object has no extended dictionary, create it
14             {
15                 obj.UpgradeOpen();//Switch object to write state
16                 obj.CreateExtensionDictionary();//Create an extended dictionary for an object, an object can only have one extended dictionary
17                 obj.DowngradeOpen();//Switch the object to read state
18             }
19             //Open the extended dictionary of the object
20         DBDictionary dict = obj.ExtensionDictionary.GetObject(OpenMode.ForRead) as DBDictionary;
21             //If the specified extended record object is already contained in the extended dictionary  
                */
                #endregion

                // откроем словарь на запись
                DBDictionary dict = (DBDictionary)trans.GetObject(ms.ExtensionDictionary, OpenMode.ForWrite);
                // что-то пока непонятное, ну да ладно
                ResultBuffer rb = new ResultBuffer();
                // какая-то переменная, кот. передаем в буфер
                TypedValue tv = new TypedValue((short)DxfCode.ExtendedDataAsciiString, xData);
                // передали в буфер
                rb.Add(tv);

                if (!dict.Contains(xKey))
                {
                    Xrecord xrec = new Xrecord();
                    xrec.Data = rb;
                    dict.SetAt(xKey, xrec);

                    trans.AddNewlyCreatedDBObject(xrec, true);
                    trans.Commit();
                }
                else
                {
                    // если хзапись с ключом samexceldatapath уже есть, зайдем в нее и заменим данные
                    // т.е. перезапишем путь к файлу
                    ObjectId xrecordId = dict.GetAt(xKey);
                    Xrecord xrec = xrecordId.GetObject(OpenMode.ForWrite) as Xrecord;
                    xrec.Data = rb;

                    trans.Commit();
                }
                #region СПРАВКА
                /*
            if (!dict.Contains(xRecordSearchKey))
82             {
83                 return null;//If there is no extended record containing the specified keyword in the extended dictionary, null will be returned;
84             }

85             //First get the extended dictionary of the object or the well-known object dictionary in the graph, and then get the extended record to be queried in the dictionary
86             ObjectId xrecordId = dict.GetAt(xRecordSearchKey);//Get the id of the extended record object
87             Xrecord xrecord = xrecordId.GetObject(OpenMode.ForRead) as Xrecord;//Get extended record object according to id
88             TypedValueList values = xrecord.Data;
89             return values;//The values ​​array should be sequential
                */
                #endregion
                trans.Dispose();
            }
        }

        private static string ReadAndGetExtDataModel(string xKey)
        {
            string xData = string.Empty;



            return xData;
        }



        #region СПРАВКА
        //DateTime t1 = DateTime.Now;

        //Document doc = Application.DocumentManager.MdiActiveDocument;
        //Database db = doc.Database;
        //Editor ed = doc.Editor;

        //using (Transaction trans = db.TransactionManager.StartTransaction())
        //{
        //    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
        //    BlockTableRecord ms = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

        //    Line l = new Line(new Point3d(0, 0, 0), new Point3d(10, 10, 10));
        //    ms.AppendEntity(l);
        //    trans.AddNewlyCreatedDBObject(l, true);

        //    l.CreateExtensionDictionary();
        //    DBDictionary dict = (DBDictionary)trans.GetObject(l.ExtensionDictionary, OpenMode.ForWrite);

        //    ResultBuffer rb = new ResultBuffer();

        //    for (int i = 1; i < 10000; i++)
        //    {
        //        string s = "";
        //        for (int j = 1; j < 10000; j++)
        //        {
        //            s += "0";
        //        }

        //        TypedValue tv = new TypedValue((short)DxfCode.ExtendedDataAsciiString, s);
        //        rb.Add(tv);
        //    }

        //    Xrecord xrec = new Xrecord();
        //    xrec.Data = rb;
        //    dict.SetAt("myxrecordkey", xrec);
        //    trans.AddNewlyCreatedDBObject(xrec, true);

        //    trans.Commit();
        //}

        //DateTime t2 = DateTime.Now;
        //ed.WriteMessage("\n" + t2.Subtract(t1).TotalSeconds.ToString());


        //[CommandMethod("XRecord_test2")]
        //public static void XRecord_test2()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor ed = doc.Editor;

        //    using (Transaction trans = db.TransactionManager.StartTransaction())
        //    {
        //        BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
        //        BlockTableRecord ms = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

        //        Line l = (Line)trans.GetObject(ed.GetEntity("\nline:").ObjectId, OpenMode.ForRead);

        //        DBDictionary dict = (DBDictionary)trans.GetObject(l.ExtensionDictionary, OpenMode.ForWrite);

        //        Xrecord xrec = (Xrecord)trans.GetObject(dict.GetAt("myxrecordkey"), OpenMode.ForRead);
        //        ResultBuffer rb = xrec.Data;

        //        int i = 0;
        //        int d = 0;
        //        foreach (TypedValue tv in rb)
        //        {
        //            i++;
        //            d += tv.Value.ToString().Length;
        //        }

        //        ed.WriteMessage("\nКол-во записей: " + i.ToString() + "\nКол-во символов: " + d.ToString());

        //        trans.Commit();
        //    }
        //}




        //}
        #endregion
    }
}
