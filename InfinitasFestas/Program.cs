using InfinitasFestas.Classes;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace InfinitasFestas
{
    public class Program
    {
        static List<Produto> produtos;
        static List<Item> categorias;
        static List<Item> cores;
        static Excel.Application xlApp;
        static Excel.Workbook xlWorkBook;
        static Excel.Worksheet xlWorkSheet, xlCores, xlCategorias;
        static Excel.Range range;

        static void Main(string[] args)
        {
            int rCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"E:\Downloads\Planilha modelo.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlCores = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            range = xlCores.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            cores = new List<Item>();
            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                Item newItem = new Item();
                newItem.ID = Convert.ToInt32((range.Cells[rCnt, 1] as Excel.Range).Value);
                newItem.descricao = (range.Cells[rCnt, 2] as Excel.Range).Value;
                cores.Add(newItem);
            }

            xlCategorias = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
            range = xlCategorias.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            categorias = new List<Item>();
            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                Item newItem = new Item();
                newItem.ID = Convert.ToInt32((range.Cells[rCnt, 1] as Excel.Range).Value);
                newItem.descricao = (range.Cells[rCnt, 2] as Excel.Range).Value;
                newItem.nroSeq = 1;
                categorias.Add(newItem);
            }

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            for (rCnt = 2; rCnt <= 101; rCnt++)
            {
                string catDesc = (range.Cells[rCnt, 5] as Excel.Range).Value;
                string corDesc = (range.Cells[rCnt, 6] as Excel.Range).Value;

                int IDCat = 0;
                int IDCor = 0;
                int nroSeq = 0;

                foreach (Item cat in categorias)
                {
                    if (catDesc.Equals(cat.descricao))
                    {
                        IDCat = cat.ID;
                        nroSeq = cat.nroSeq++;
                        break;
                    }
                }

                foreach (Item cor in cores)
                {
                    if (corDesc.Equals(cor.descricao))
                    {
                        IDCor = cor.ID;
                        break;
                    }
                }

                string codigo = IDCat.ToString().PadLeft(2, '0') + IDCor.ToString().PadLeft(2, '0') + nroSeq.ToString().PadLeft(3, '0');
                range.Cells[rCnt, 2] = codigo;
            }

            xlWorkSheet.SaveAs(@"E:\Downloads\Planilha modelo_TOBIAS.xlsx");

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
        }
    }
}