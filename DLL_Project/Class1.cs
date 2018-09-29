using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;

using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace DLL_Project
{
    public class Class1
    {
        //Variablen zur parallelen Summenbildung
        public long sum = 0;
        public long ret = 0;

        static readonly object _locker = new object();
        public object[,] ar;
        
        public long get_Sum2(Range pRng, int WorkerThreads) {
            sum = 0;
            ar = pRng.Value2;


            //System.Diagnostics.Debugger.Launch();
            //System.Diagnostics.Debug.WriteLine("Debug");
 
            Parallel.For(1, ar.GetUpperBound(1) + 1, i => Addition2(i));
            
            return sum;
        }

        public void Addition2(int Col) {
            int tmpSum = 0;
            for (int i = ar.GetLowerBound(0); i <= ar.GetUpperBound(0); i++)
                tmpSum += Convert.ToInt32(ar[i, Col]);

            lock (_locker) {
                sum += tmpSum;
                //if (counter <= ar.GetUpperBound(1))
                //    Parallel.Invoke(() => Addition2(counter));
                
            }
        }

        //weitere Berechnungsverfahren
        //Multiplikation
        public long get_Produkt(Range pRng) {
            ret = 0;
            ar = pRng.Value2;

            Parallel.For(1, ar.GetUpperBound(1) + 1, i => Multiply(i));

            return ret;
        }
        public void Multiply(int Col) {
            int tmpSum = 0;
            for (int i = ar.GetLowerBound(0); i <= ar.GetUpperBound(0); i++)
                tmpSum *= Convert.ToInt32(ar[i, Col]);
           
            lock (_locker) {
                ret += tmpSum;
            }
        }

        //Textsuche
        public int get_String_Matches(Range pRng, String pSearch_String) {
            ret = 0;
            ar = pRng.Value2;

            Parallel.For(1, ar.GetUpperBound(1) + 1, i => search_String(i, pSearch_String));

            return Convert.ToInt16(ret);
        }
        public void search_String(int Col, String pSearch_String) {
            int tmpSum = 0;
            for (int i = ar.GetLowerBound(0); i <= ar.GetUpperBound(0); i++)
                if (ar[i, Col].ToString() == pSearch_String)
                    tmpSum++;

            lock (_locker)
            {
                ret += tmpSum;
            }
        }

        //Subtextsuche
        public int get_SubString_Matches(Range pRng, String pSearch_String)
        {
            ret = 0;
            ar = pRng.Value2;

            Parallel.For(1, ar.GetUpperBound(1) + 1, i => search_SubString(i, pSearch_String));

            return Convert.ToInt16(ret);
        }
        public void search_SubString(int Col, String pSearch_String)
        {
            int tmpSum = 0;
            for (int i = ar.GetLowerBound(0); i <= ar.GetUpperBound(0); i++)
                for (int j = 0; j < ar[i, Col].ToString().Length - (pSearch_String.Length-1); j++)
                    if (ar[i, Col].ToString().Substring(j, pSearch_String.Length) == pSearch_String)
                        tmpSum++;

            lock (_locker)
            {
                ret += tmpSum;
            }
        }

        //Dividieren
        public void Divide(Range pRng, int Divisor)
        {
            ar = pRng.Value2;

            Parallel.For(1, ar.GetUpperBound(1) + 1, i => Divide_Parallel(i, Divisor));
        }
        public void Divide_Parallel(int Col, int Divisor)
        {
            Double tmpSum = 0;
            for (int i = ar.GetLowerBound(0); i <= ar.GetUpperBound(0); i++)
                tmpSum = Convert.ToInt32(ar[i, Col]) / Divisor;

        }
    }
}
