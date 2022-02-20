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



        public static bool IsDigitStr (string str)
        {
            //bool isChek = false;
            // "1,00434500034563".Replace(",", "").ToCharArray().All(char.IsDigit)
            // "23,42".Replace(",", "").ToCharArray().All(char.IsDigit)


            if (str.ToCharArray().All(char.IsDigit))
                return true;
            else
            {


                if (str.Contains(","))
                {
                    str = str.Replace(",", "");
                    if (str.ToCharArray().All(char.IsDigit))
                    {
                        return true;
                    }
                    else
                        return false;
                }
                else
                {
                    if (str.Contains("."))
                    {
                        str = str.Replace(".", "");
                        if (str.ToCharArray().All(char.IsDigit))
                        {
                            return true;
                        }
                        else
                            return false;
                    }
                    else
                        return false;
                }

            }

            //if (
            //    (!str.Contains(",")) &&
            //    (!str.Contains("."))
            //   )
            //{
            //    if (str.ToCharArray().All(char.IsDigit))
            //        return true;
            //    else
            //        return false;
            //}
            //else
            //{
            //    if (!str.ToCharArray().All(char.IsDigit))
            //        return false;
            //}




            // return isChek;
        }



        public static string GetPointValid (string str)
        {
            if (str.Contains("."))
            {
                return str.Replace(".", ",");

            }
            else
                return str;
        }


    }



}
