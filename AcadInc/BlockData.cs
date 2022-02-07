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

[assembly: CommandClass(typeof(AcadInc.BlockData))]


namespace AcadInc
{
    public static class BlockData
    {

        // https://www.theswamp.org/index.php?topic=55238.0
        //[CommandMethod("selb")]
        public static void BlockRefModifity(List<ExcelData.Model.BlockData> blockDatas)
        {
            AcadSendMess AcMess = new AcadSendMess();

            foreach (ObjectId blockRefId in selectDynamicBlockReferences())
            {
                    AcMess.SendStringDebug(BlockRefAttributeRefWrite(blockRefId, blockDatas) );
            }
        }


        public static string BlockRefAttributeRefWrite(ObjectId bed, List<ExcelData.Model.BlockData> blockDatas)
        {
            Database db = Application.DocumentManager.MdiActiveDocument.Database;

            using (Transaction rbTrans = db.TransactionManager.StartTransaction())
            {
                string result = "";


                BlockReference bref = (BlockReference)rbTrans.GetObject(bed, OpenMode.ForWrite);
                BlockTableRecord bdef = (BlockTableRecord)rbTrans.GetObject(bref.DynamicBlockTableRecord, OpenMode.ForWrite);

                foreach (ExcelData.Model.BlockData blockData in blockDatas)
                {
                    if (bdef.Name == blockData.BlockName)
                    {

                        if (bdef.HasAttributeDefinitions == true)
                        {

                            foreach (ObjectId id in bref.AttributeCollection)
                            {
                                AttributeReference attref = (AttributeReference)rbTrans.GetObject(id, OpenMode.ForWrite);

                                //if (attref.Tag ==  Const.BlockAttrApparatTag)
                                //{
                                //    //result = attref.Tag;
                                //    //attref.TextString = "Я тута!";
                                //}

                                // проверим, что совпадают номер участка (секция) и QF

                                bool isChek = false;

                                foreach (AttrData attrData in blockData.ListAttributes)
                                {
                                    if (attref.Tag == attrData.AttributeTag)
                                    {
                                        attref.TextString = attrData.AttributeValue;
                                    }
                                }







                                foreach (AttrData attrData in blockData.ListAttributes)
                                {
                                    if (attref.Tag == attrData.AttributeTag)
                                    {
                                        attref.TextString = attrData.AttributeValue;
                                    }
                                }
                            }
                        }
                    }

                }



                rbTrans.Commit();
                rbTrans.Dispose();

                return "Запись атр. выполнена.";
            }

        }


        // Взято из 15/07/2013:
        // https://adn-cis.org/kak-najti-vse-vstavki-dinamicheskogo-bloka.html
        //[CommandMethod("selb")]
        public static ObjectIdCollection selectDynamicBlockReferences()
        {
            ObjectIdCollection resultCollection = null;

            //Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = Application.DocumentManager.MdiActiveDocument.Database;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // получаем таблицу блоков и проходим по всем записям таблицы блоков
                BlockTable bt =
                  (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in bt)
                {
                    // получаем запись таблицы блоков и смотри анонимная ли она
                    BlockTableRecord btr =
                      (BlockTableRecord)trans.GetObject(btrId, OpenMode.ForRead);
                    if (btr.IsDynamicBlock)
                    {
                        // получаем все анонимные блоки динамического блока
                        ObjectIdCollection anonymousIds = btr.GetAnonymousBlockIds();
                        // получаем все прямые вставки динамического блока
                        ObjectIdCollection dynBlockRefs = btr.GetBlockReferenceIds(true, true);
                        foreach (ObjectId anonymousBtrId in anonymousIds)
                        {
                            // получаем анонимный блок
                            BlockTableRecord anonymousBtr =
                                 (BlockTableRecord)trans.GetObject(anonymousBtrId, OpenMode.ForRead);
                            // получаем все вставки этого блока
                            ObjectIdCollection blockRefIds =
                                 anonymousBtr.GetBlockReferenceIds(true, true);
                            foreach (ObjectId id in blockRefIds)
                            {
                                dynBlockRefs.Add(id);
                                // зайдем в блок и пройдемся по атрибутам
                                


                            }
                        }
                        // Что-нибудь делаем с созданным нами набором
                        //ed.WriteMessage("\nДинамическому блоку \"{0}\" соответствуют {1} анонимных блоков и {2} вставок блока\n",
                        //    btr.Name, anonymousIds.Count, dynBlockRefs.Count);
                        resultCollection = dynBlockRefs;





                    }
                }
            }

            return resultCollection;
        }
    }
}
