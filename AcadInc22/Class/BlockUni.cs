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
    public static class BlockUni
    {


        /// <summary>
        /// Возращает список Id вхождений нединамического блока по его имени (имени определения блока). 
        /// </summary>
        /// <returns>список objectId</returns>
        public static  List<ObjectId> GetListIdsNotDynamicBlockRefsFromNameBlock ()
        {
            List<ObjectId> objectIdsList= null;
            // проверим все вхождения всех недин. блоков

            // и добавим в список

            // Заглушка
            ObjectId objId = new ObjectId();
            objectIdsList.Add(objId);


            return objectIdsList;
        }


        /// <summary>
        /// Записывает значения в атрибуты вхождения блока. Возращает строку- отчет.
        /// </summary>
        /// <param name="blockId"> Id вхождения блока</param>
        /// <param name="attrDatas">Список пар "Тэг атт - значение" </param>
        /// <returns></returns>
        public static string WriteDataToAttrsOfBlock (ObjectId blockId, List<AttrData> attrDatas)
        {
            string str = string.Empty;

            // откроем вхождение блока на запись и запишем в его атрибуты свои значения

            // Проверим, есть ли вообще атрибуты.

            // цикл по списку атрибутов - при совпадении тэга тек атр. с тэгом из списка пар "тэг-атр"

            return str;
        }


        /// <summary>
        /// Возращает 
        /// </summary>
        /// <returns></returns>
        public static List<ObjectIdCollection> selectNotDynamicBlockReferences()
        {
            List<ObjectIdCollection> resultCollection = new List<ObjectIdCollection>();

            //Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = Application.DocumentManager.MdiActiveDocument.Database;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // получаем таблицу блоков и проходим по всем записям таблицы блоков
                BlockTable bt =
                  (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in bt)
                {
                    // получаем запись таблицы блоков 
                    BlockTableRecord btr =
                      (BlockTableRecord)trans.GetObject(btrId, OpenMode.ForRead);
                    if (!btr.IsDynamicBlock)
                    {
                      
                        // получаем все прямые вставки динамического блока
                        ObjectIdCollection notDynBlockRefs = btr.GetBlockReferenceIds(true, true);

                        resultCollection.Add(notDynBlockRefs);
                    }

                }

            }

            return resultCollection;
        }


        /// <summary>
        /// <para>
        /// Пример из 2013г: </para>
        /// Огромный респект Ривилису:
        /// <br/>
        /// <a href="https://adn-cis.org/kak-najti-vse-vstavki-dinamicheskogo-bloka.html"></a>
        /// <br/><br/>
        /// и Баладжи Рамамурти: 
        /// <br/>
        /// <a href="https://adndevblog.typepad.com/autocad/2012/06/finding-all-block-references-of-a-dynamic-block.html"></a>
        /// </summary>
        /// <returns></returns>
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
