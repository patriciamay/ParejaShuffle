using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel.Design;
using static System.Net.Mime.MediaTypeNames;

namespace ParejaShuffle
{
        internal class Program
        {
            public static int time = 1;
            public static void shuffleList(List<string> list)
            {

            }
            public static void GetuserCards(List<string> ucards)
            {

            }
            public static void deckofCards(Stack<string> deck)
            {

            }
            private static void TimerCallback(object o)
            {

            }
            public static void setTimer()
            {

            }
            static void Main(string[] args)
            {
                List<string> list = new List<string>();
                List<string> usercards = new List<string>();

                Application excelApp = new Application();

                if (excelApp == null)
                {
                    Console.WriteLine("Excel is not installed!");
                }

                Workbook.excelBook = excelApp.Workbooks.Open(@"C:\Users\22-0281C\Downloads\deckofcards (1).xlsx");
                _Worksheet excelSheet = excelBook.Sheets[1];
                Range excelRange = excelSheet.UsedRange;

                int rows = excelRange.Rows.Count;
                int cols = excelRange.Columns.Count;

                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= cols; j++)
                    {
                        //write the console
                        if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value != null)
                            list.Add(excelRange.Cells[i, j].Value2.ToString());
                    }
                }

                Console.WriteLine("Shuffling Card in ");
                //shuffles cards in 5 secs
                setTimer();
                //display shuffled cards
                shuffleList(list);

                Stack<string> deckcards = new Stack<string>(list);

                deckofCards(deckcards);

                Console.WriteLine("\nGenerate User Cards? Pat: ");
                string ans = Console.ReadLine().ToLower();
                if (ans == "y")
                {
                    Console.WriteLine("Generating User Cards in");
                    setTimer();
                    GetuserCards(usercards, deckcards);
                    deckofCards(deckcards);
                }

                Console.Write(ans"\nDraw First Card? Pat: ");
                string ans2 = Console.ReadLine().ToLower();
                if (ans2 == "y")
                {
                    usercards.RemoveAt(0);
                    displayUserCards(usercards);
                    usercards.Add(deckcards.Pop());
                    displayUserCards(usercards);
                    deckofCards(deckcards);
                }

                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                Console.ReadLine();

            }
        }
    }