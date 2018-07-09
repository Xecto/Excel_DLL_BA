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
        //Variablen
        private int Nr_Threads = 1;
        private int max_Threads = 1;

        //Variablen zur parallelen Summenbildung
        public long sum = 0;
        public int counter = 1;
        static readonly object _locker = new object();
        public object[,] ar;

        //Variablen für den Last Test
        uint[,,] Zahlen_Array = new uint[Environment.ProcessorCount * 2, 9999998, 2]; //Primzahlen von 2 bis 999999 (10.000.000)

        Boolean Test_Done = false;
        int CPU_Cores = Environment.ProcessorCount;
        int Iterationen = 8;


        public long Last_Test(int Thread_Anzahl)
        {
            Stopwatch stopwatch = new Stopwatch();
            int avg_Time = 0;
            //TO-DO Berechnung durchführen mit unterschiedlicher Anzal an Threads
            //Primzahlentest
            //Init des Arrays//Zahlen 2 bis 999//zweite Dimension 0=nicht gestrichen; 1=gestrichen
            if (Thread_Anzahl >= CPU_Cores * 2)
                Thread_Anzahl = CPU_Cores * 2;

            for (int j=0;j < CPU_Cores * 2; j++)
                for (uint i = 2; i < 10000000; i++)
                {
                    Zahlen_Array[j, i - 2, 0] = i;
                    Zahlen_Array[j, i - 2, 1] = 0;
                }

            //maximale Anzahl an Threads einstellen
            max_Threads = Thread_Anzahl;

            //mehrere Iterationen durchführen und avg-time bestimmen
            for (int a = 0; a < Iterationen; a++)
            {
                stopwatch.Start();
                counter = max_Threads - 1;
                Parallel.For(0, max_Threads, i => get_Prime_Numbers(i));
                stopwatch.Stop();
                avg_Time += (int)stopwatch.ElapsedMilliseconds;
            }
            return avg_Time/Iterationen;
        }

        /// <summary>
        /// Ermittelt die Primzahlen in 4 Arrays, mittels des Sieb des Eratosthenes, bis 10.000.000 (zehn Millionen)
        /// </summary>
        /// <param name="index">Bestimmt welches der 4 Arrays berechnet werden soll</param>
        public void get_Prime_Numbers(int index) {
            Increase();

            if (index == (CPU_Cores * 2)-1)
                Test_Done = true;
           
            for (uint i = 0; i < 10000000 - 2; i++)
            {
                if (Zahlen_Array[index, i, 1] != 1)
                {
                    for (uint j = i; j < 10000000 - 2; j+= Zahlen_Array[index, i, 0])
                    {
                        Zahlen_Array[index, j, 1] = 1;
                    }
                }
            }

            //neuen Thread erstellen
            lock (_locker)
            {
                counter++;
                sum++;
                if (Test_Done == false)
                    Parallel.Invoke(() => get_Prime_Numbers(counter));
            }
            Decrease();
        }

        public long get_Sum(Range pRng, int pAnzahl)
        {
            //Objekte definieren und max. Anzahl der Spalten/Zeilen ermitteln
            ar = pRng.Value2;

            max_Threads = pAnzahl;
            counter = ar.GetLowerBound(1);
            sum = 0;

            if (max_Threads > ar.GetUpperBound(1))
                max_Threads = ar.GetUpperBound(1);

            counter = max_Threads;
            Parallel.For(1, max_Threads + 1, i => Addition(i, i));

            return sum;
        }

        /// <summary>
        /// Bildet die Summe über die angegebene Spalte
        /// </summary>
        /// <param name="Col">Gibt an, über welche Spalte die Summe erstellt werden soll.</param>
        public void Addition(int Col, int Thread_Number)
        {
            Increase();
            int tmpSum = 0;
            for (int i = ar.GetLowerBound(0); i <= ar.GetUpperBound(0); i++)
                tmpSum += Convert.ToInt32(ar[i, Col]);


            lock (_locker)
            {
                counter++;
                sum += tmpSum;

                if (counter <= ar.GetUpperBound(1))
                    Parallel.Invoke(() => Addition(counter, Thread_Number));
            }
            Decrease();
        }

        /// <summary>
        /// Erhöht die Anzahl (Zählervariable) der Threads, wenn diese die maximale Anzahl an Threads nicht überschreitet
        /// </summary>
        public void Increase()
        {
            lock (_locker)
            {
                if (Nr_Threads < max_Threads)
                    Nr_Threads++;
            }
        }

        /// <summary>
        /// Reduziert die Anzahl der Threads (Zählervariable), solange diese nicht unter 1 fällt
        /// </summary>
        public void Decrease()
        {
            lock (_locker)
            {
                if (Nr_Threads >= 1)
                    Nr_Threads--;
            }
        }
    }
}
