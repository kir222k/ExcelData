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

[assembly: CommandClass(typeof(AcadInc.BlockData))]

namespace AcadInc
{
    public static class BlockData
    {
        // https://www.theswamp.org/index.php?topic=55238.0
        //[CommandMethod("selb")]
        public static void BlockRefModifity(List<ExcelData.Model.BlockData> blockDatas)
        {
            //AcadSendMess AcMess = new AcadSendMess();

            // пройдемся по всем вхождениям всех блоков и будем подсовывать им наш blockDatas
            foreach (ObjectId blockRefId in BlockUni.selectDynamicBlockReferences())
            {
                // AcMess.SendStringDebug(c);
                string str = BlockRefAttributeRefWrite(blockRefId, blockDatas);
            }



        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="blockDatas"></param>
        public static void BlockRefNotDynamicModifity(List<ExcelData.Model.BlockData> blockDatas)
        {
            //AcadSendMess AcMess = new AcadSendMess();

            // пройдемся по всем вхождениям всех блоков и будем подсовывать им наш blockDatas
            foreach ( ObjectIdCollection objectIdCollection in BlockUni.selectNotDynamicBlockReferences())
            {
                foreach (ObjectId blockRefId in objectIdCollection)
                {
                    // AcMess.SendStringDebug(c);
                    string str = BlockRefAttributeRefWrite2(blockRefId, blockDatas);
                }

            }
        }


        public static string BlockRefAttributeRefWrite(ObjectId bed, List<ExcelData.Model.BlockData> blockDatas)
        {
            Database db = Application.DocumentManager.MdiActiveDocument.Database;
            using (Transaction rbTrans = db.TransactionManager.StartTransaction())
            {
                BlockReference blRef = (BlockReference)rbTrans.GetObject(bed, OpenMode.ForWrite);
                BlockTableRecord blRefTabRec = (BlockTableRecord)rbTrans.GetObject(blRef.DynamicBlockTableRecord, OpenMode.ForWrite);
                //BlockTableRecord blRefTabRec = (BlockTableRecord)rbTrans.GetObject(blRef.BlockTableRecord, OpenMode.ForWrite);
                // пройдемся по нашему списку блоков с атрибутами из Excel
                // и сравним/поработаем с атрибутами данного вхождения блока:
                foreach (ExcelData.Model.BlockData blockData in blockDatas)
                {
                    if (blRefTabRec.Name == blockData.BlockName)
                    {
                        if (blRefTabRec.HasAttributeDefinitions == true)
                        {

                            // по-умолчанию номер участка (секция) не совпадает
                            bool isChekSection = false;
                            // по-умолчанию  QF не совпадает
                            bool isChekQF = false;

                            // проверим, что совпадают УЧАСТОК и N.АПП1
                            // для этого пройдем по коллекции атрибутов тек вх. блока
                            foreach (ObjectId id in blRef.AttributeCollection)
                            {
                                // получим атрибут
                                AttributeReference attref = (AttributeReference)rbTrans.GetObject(id, OpenMode.ForRead);

                                // проход по атрибутам, получ. из Excel
                                foreach (AttrData attrData in blockData.ListAttributes)
                                {
                                    //string tagAttrRef = attref.Tag;

                                    if (
                                        (attref.Tag.Equals(attrData.AttributeTag)) &&        // если тэг текущего атрибута вх. блока совпал с КАКИМ-ТО атр.  из Excel
                                        (attref.TextString.Equals(attrData.AttributeValue))     // и если совпадает значение атрибута вх. блока и атрибута из Excel
                                       ) 
                                    {
                                        // проверим, что тэг = "УЧАСТОК"
                                        if (attref.Tag.Equals(Const.BlockAttrApparatSect))       
                                        {
                                            isChekSection = true;
                                        }
                                        // проверим, что тэг = "N.АПП1"
                                        if (attref.Tag.Equals(Const.BlockAttrApparatQF))       
                                        {
                                            isChekQF = true;
                                        }
                                    }
                                }
                            }

                            // Если значения атрибутов  N.АПП1 и УЧАСТОК
                            // совпадают в данном вхождении блока с атрибутами, получ. из Excel:
                            if (isChekSection && isChekQF)
                            {
                                // тогда пройдем по коллекции атрибутов тек вх. блока
                                foreach (ObjectId id in blRef.AttributeCollection)
                                {
                                    // получим атрибут
                                    AttributeReference attref = (AttributeReference)rbTrans.GetObject(id, OpenMode.ForWrite);

                                    // проход по атрибутам, получ. из Excel
                                    foreach (AttrData attrData in blockData.ListAttributes)
                                    {
                                        // если тэг текущего атрибута вх. блока совпал с атр. из Excel
                                        if (
                                            (attref.Tag.Equals(attrData.AttributeTag)) && // И если 
                                            (!attref.Tag.Equals(Const.BlockAttrApparatQF)) &&
                                            (!attref.Tag.Equals(Const.BlockAttrApparatSect))
                                           )
                                        {
                                            // тогда запишем в него свое значение
                                            attref.TextString = attrData.AttributeValue;
                                        }
                                    }
                                }

                            } // если элемент типа BlockData из списка не подходит для текущего блока, ничего не пишем в его атрибуты 
                        }
                    }

                } // и берем для манипуляций след. элемент типа BlockData в списке 

                rbTrans.Commit();
                rbTrans.Dispose();

                return "BlockRefAttributeRefWrite is completed.";
            }

        }

        public static string BlockRefAttributeRefWrite2(ObjectId bed, List<ExcelData.Model.BlockData> blockDatas)
        {
            Database db = Application.DocumentManager.MdiActiveDocument.Database;
            using (Transaction rbTrans = db.TransactionManager.StartTransaction())
            {
                BlockReference blRef = (BlockReference)rbTrans.GetObject(bed, OpenMode.ForWrite);
                //BlockTableRecord blRefTabRec = (BlockTableRecord)rbTrans.GetObject(blRef.DynamicBlockTableRecord, OpenMode.ForWrite);
                BlockTableRecord blRefTabRec = (BlockTableRecord)rbTrans.GetObject(blRef.BlockTableRecord, OpenMode.ForWrite);
                // пройдемся по нашему списку блоков с атрибутами из Excel
                // и сравним/поработаем с атрибутами данного вхождения блока:
                foreach (ExcelData.Model.BlockData blockData in blockDatas)
                {
                    if (blRefTabRec.Name == blockData.BlockName)
                    {
                        if (blRefTabRec.HasAttributeDefinitions == true)
                        {

                            // по-умолчанию номер участка (секция) не совпадает
                            bool isChekSection = false;
                            // по-умолчанию  QF не совпадает
                            bool isChekQF = false;

                            // проверим, что совпадают УЧАСТОК и N.АПП1
                            // для этого пройдем по коллекции атрибутов тек вх. блока
                            foreach (ObjectId id in blRef.AttributeCollection)
                            {
                                // получим атрибут
                                AttributeReference attref = (AttributeReference)rbTrans.GetObject(id, OpenMode.ForRead);

                                // проход по атрибутам, получ. из Excel
                                foreach (AttrData attrData in blockData.ListAttributes)
                                {
                                    //string tagAttrRef = attref.Tag;

                                    if (
                                        (attref.Tag.Equals(attrData.AttributeTag)) &&        // если тэг текущего атрибута вх. блока совпал с КАКИМ-ТО атр.  из Excel
                                        (attref.TextString.Equals(attrData.AttributeValue))     // и если совпадает значение атрибута вх. блока и атрибута из Excel
                                       )
                                    {
                                        // проверим, что тэг = "УЧАСТОК"
                                        if (attref.Tag.Equals(Const.BlockAttrApparatSect))
                                        {
                                            isChekSection = true;
                                        }
                                        // проверим, что тэг = "N.АПП1"
                                        if (attref.Tag.Equals(Const.BlockAttrApparatQF))
                                        {
                                            isChekQF = true;
                                        }
                                    }
                                }
                            }

                            // Если значения атрибутов  N.АПП1 и УЧАСТОК
                            // совпадают в данном вхождении блока с атрибутами, получ. из Excel:
                            if (isChekSection && isChekQF)
                            {
                                // тогда пройдем по коллекции атрибутов тек вх. блока
                                foreach (ObjectId id in blRef.AttributeCollection)
                                {
                                    // получим атрибут
                                    AttributeReference attref = (AttributeReference)rbTrans.GetObject(id, OpenMode.ForWrite);

                                    // проход по атрибутам, получ. из Excel
                                    foreach (AttrData attrData in blockData.ListAttributes)
                                    {
                                        // если тэг текущего атрибута вх. блока совпал с атр. из Excel
                                        if (
                                            (attref.Tag.Equals(attrData.AttributeTag)) && // И если 
                                            (!attref.Tag.Equals(Const.BlockAttrApparatQF)) &&
                                            (!attref.Tag.Equals(Const.BlockAttrApparatSect))
                                           )
                                        {
                                            // тогда запишем в него свое значение
                                            attref.TextString = attrData.AttributeValue;
                                        }
                                    }
                                }

                            } // если элемент типа BlockData из списка не подходит для текущего блока, ничего не пишем в его атрибуты 
                        }
                    }

                } // и берем для манипуляций след. элемент типа BlockData в списке 

                rbTrans.Commit();
                rbTrans.Dispose();

                return "BlockRefAttributeRefWrite is completed.";
            }

        }



    }
}
