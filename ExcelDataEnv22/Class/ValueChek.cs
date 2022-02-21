using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData.Class
{
    public static class ValueChek
    {
        public static string GetAddZeroStr(string strWithoutZero, int round)
        {
            string str = strWithoutZero;

            if (str.Contains(","))
            {
                int lenZeroRound = str.Substring(str.IndexOf(",") + 1).Length;
                if (lenZeroRound < round)
                {
                    int deltaZero = round - lenZeroRound;

                    for (int k = 0; k < deltaZero; k++) // 0.02 а  разрядов - 3, тогда deltaZero=1 и добавим 1 ноль справа
                    {
                        str += "0";
                    }
                }
            }
            else
            {
                str += ",";
                for (int k = 0; k < round; k++) // 0.02 а  разрядов - 3, тогда deltaZero=1 и добавим 1 ноль справа
                {
                    str += "0";
                }
            }
            return str;
        }


        /// <summary>
        /// Проверяет, может ли строка быть преобр. в число. Разделитель - "," или "."
        /// </summary>
        /// <param name="str">строка-предположительно-число</param>
        /// <returns></returns>
        public static bool IsDigitStr (string str)
        {
            // "23,42".Replace(",", "").ToCharArray().All(char.IsDigit)

            // если все символы - цифры.
            if (str.ToCharArray().All(char.IsDigit))
                // то пройдено!
                return true;
            // если не все символы - цифры.
            else
            {
                // если нашли ","
                if (str.Contains(","))
                {
                    // проверим, единственный ли это разделитель
                    if (IsPointRoundExclusive(str))
                    {
                        // и если его убрать,
                        str = str.Replace(",", "");
                        // будут ли тогда все символы строки цифрами?
                        if (str.ToCharArray().All(char.IsDigit))
                        {
                            // если да, то пройдено!
                            return true;
                        }
                        // если нет, т.е. есть еще нецифровые символы, кроме ","
                        else
                            // соотв. - не пройдено!
                            return false;

                    }
                    // если  это - не единственный разделитель,
                    else
                        // тогда - не пройдено!
                        return false;
                }
                // если не все символы - цифры И если не нашли ","
                else
                {
                    // но нашли "."
                    if (str.Contains("."))
                    {
                        // тогда проверим, единственный ли это разделитель
                        if (IsPointRoundExclusive(str))
                        {
                            // и если его убрать,
                            str = str.Replace(".", "");
                            // будут ли тогда все символы строки цифрами?
                            if (str.ToCharArray().All(char.IsDigit))
                            {
                                // если да, то пройдено!
                                return true;
                            }
                            // если нет, т.е. есть еще нецифровые символы, кроме "."
                            else
                                // соотв. - не пройдено!
                                return false;
                        }
                        // если  это - не единственный разделитель,
                        else
                            // тогда - не пройдено!
                            return false;

                    }
                    // если не все символы - цифры
                    // И
                    // если не нашли "," 
                    // И
                    // не нашли "."
                    else
                        // тогда строка не число - не пройдено!
                        return false;
                }
            }

        }


        /// <summary>
        /// Возращает строку, где "." заменены на ","
        /// </summary>
        /// <param name="str">строка-возможно-с-точкой</param>
        /// <returns></returns>
        public static string GetStringWithPointCorrect (string str)
        {
            if (str.Contains("."))
            {
                return str.Replace(".", ",");

            }
            else
                return str;
        }


        /// <summary>
        /// Проверяет, является ли разделитель ("," или ".") единственным в строке-числе
        /// </summary>
        /// <param name="str">строка, д. содержать хотя бы 1 символ "," или "."</param>
        /// <returns>false - если нет разделителей, или их несколько в строке, true - если есть единств. разделитель</returns>
        public static bool IsPointRoundExclusive (string str)
        {
            // по классике запросов SQL.
            /*
            var strOnlyPointsClearLinq = from charX in str.ToCharArray()
                                where  charX.ToString() == "," || 
                                       charX.ToString() == "."
                                select charX;
            */

            // работа со строкой, как с коллекцией.
            // "1,2,2.32.3".Where(c => c.ToString() == "," || c.ToString() == ".").ToList()

            var strOnlyPointsLinqOfCollection =
                str.Where(c => c.ToString() == "," || c.ToString() == ".").ToList();

            if (strOnlyPointsLinqOfCollection.Count==1)
                return true;
            else
                return false;

        }


        /// <summary>
        /// Возращает кол-во знаков после запятой в строке-числе
        /// </summary>
        /// <param name="str">строка</param>
        /// <returns>кол-во знаков после запятой</returns>
        public static int GetQuantOfPoint(string  str)
        {
            int numOfPt = 0;

            if (IsPointRoundExclusive(str)) // если разделитель есть  и он один
            {
                if (str.Contains(","))
                {
                    // "1,343".Substring("1,343".IndexOf(",")+1).Length
                    numOfPt = str.Substring(str.IndexOf(",") + 1).Length;
                }

                if (str.Contains("."))
                {
                    // "1,343".Substring("1,343".IndexOf(",")+1).Length
                    numOfPt = str.Substring(str.IndexOf(".") + 1).Length;
                }

            }

            return numOfPt;
        }





    }



}
