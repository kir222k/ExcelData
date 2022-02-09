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
        //public static void WriteToExtDataFile ((string,string) dataToWrite)
        //{

        // https://adn-cis.org/forum/index.php?topic=8184.15
        //[CommandMethod("XRecord_test")]
        public static void WriteToExtDataFile((string, string) dataToWrite)
            {



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
            }

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


        }

    //}
}
