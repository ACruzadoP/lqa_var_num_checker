using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Equity_Checking
{
    class TheClass
    {
        private static Excel.Application oXL;

        private static ArrayList WorkBooks = new ArrayList();
        private static Excel._Workbook oWB;
        private static Excel._Worksheet oWS;

        private static Excel._Workbook oWBReport;
        private static Excel._Worksheet WorkSheetReport;

        private static string cellValueS;
        private static string cellID;
        private static string openingVar;
        private static string closingVar;
        private static List<string> variablesInString = new List<string>();
        private static string numerosenSource = "";
        private static bool issuefounded = false;

        public static void createapliXL()
        {
            if (oXL == null)
            {
                oXL = new Excel.Application();
                oXL.DisplayAlerts = false;
            }
        }
        public static string openingWB(string ruta)
        {
            oWB = oXL.Workbooks.Open(ruta,
                0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                true, false, 0, true, false, false);
            return oWB.Name;
        }
        public static void createnewWB(string nombreSheet)
        {
            if (oWBReport == null)
            {
                oWBReport = (Excel._Workbook)oXL.Workbooks.Add();
            }
            if (WorkBooks.Count >= 3)
            {
                oWBReport.Worksheets.Add(Type.Missing, (Excel.Worksheet)oWBReport.Worksheets.get_Item(WorkBooks.Count), Type.Missing, Type.Missing);
            }
            WorkSheetReport = (Excel.Worksheet)oWBReport.Worksheets.get_Item(WorkBooks.Count + 1);
            WorkSheetReport.Name = nombreSheet;
            WorkSheetReport.Cells[1, 1] = "String IDs";
            WorkSheetReport.Cells[1, 2] = "SOURCE";
            WorkSheetReport.Cells[1, 3] = "ISSUE str.";
        }
        public static void addnewWBtoArrayLst()
        {
            WorkBooks.Add(oWB);
        }

        public static void cerrartodo()
        {
            cerraroWB();
            if (oWBReport != null)
            {
                oWBReport.Close(false);
                if (WorkSheetReport != null)
                {
                    releaseObject(WorkSheetReport);
                    WorkSheetReport = null;
                }
                releaseObject(oWBReport);
                oWBReport = null;
            }
            if (oXL != null)
            {
                oXL.Quit();
                releaseObject(oXL);
                oXL = null;
            }
        }
        public static void cerraroWB()
        {
            if (oWB != null)
            {
                oWB.Close(false);
                if (oWS != null)
                {
                    releaseObject(oWS);
                    oWS = null;
                }
                releaseObject(oWB);
                oWB = null;
            }
        }
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                System.Windows.Forms.MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        public static void cleanWorkbooksArray()
        {
            WorkBooks.Clear();
        }

        public static bool checklimitcolums(string colSour, string colID, string[] cellsCol, System.Drawing.Color color)
        {
            bool lessthanZ = true;
            int i = 0;
            while ((i < cellsCol.Length) && (lessthanZ == true))
            {
                if (!cellsCol[i].Contains("^") && !cellsCol[i].Contains("`") && !cellsCol[i].Contains("´") && !cellsCol[i].Contains("¨"))
                {
                    if (color == System.Drawing.Color.DarkRed)
                    {
                        if (cellsCol[i] != "")
                        {
                            if ((int.Parse(cellsCol[i]) > 36) || (int.Parse(cellsCol[i]) == 0))
                            {
                                lessthanZ = false;
                            }
                        }
                    }
                }
                else
                {
                    lessthanZ = false;
                }
                i++;
            }
            if (!colSour.Contains("^") && !colSour.Contains("`") && !colSour.Contains("´") && !colSour.Contains("¨") && !colID.Contains("^") && !colID.Contains("`") && !colID.Contains("´") && !colID.Contains("¨"))
            {
                if (color == System.Drawing.Color.DarkRed)
                {
                    if ((int.Parse(colSour) > 36) || (int.Parse(colID) > 36) || (int.Parse(colSour) == 0) || (int.Parse(colID) == 0))
                    {
                        lessthanZ = false;
                    }
                }
            }
            else
            {
                lessthanZ = false;
            }
            return lessthanZ;
        }
        public static int GetIndexInAlphabet(char value)
        {
            // Uses the uppercase character unicode code point. 'A' = U+0042 = 65, 'Z' = U+005A = 90
            char upper = char.ToUpper(value);
            return (int)upper - 64;
        }
        public static bool thereareissues()
        {
            for (int i = 1; i <= WorkBooks.Count; i++)
            {
                WorkSheetReport = (Excel.Worksheet)oWBReport.Worksheets.get_Item(i);
                if (WorkSheetReport.UsedRange.Rows.Count > 1)
                {
                    return true;
                }
            }
            return false;
        }
        public static Excel._Workbook returnWBReport()
        {
            return oWBReport;
        }
        private static void actualizarSID(string cellValueS, string cellID, string openingVar, string closingVar)
        {
            TheClass.cellValueS = cellValueS;
            TheClass.cellID = cellID;
            TheClass.openingVar = openingVar;
            TheClass.closingVar = closingVar;
        }

        public static void superswitch(int colSour, int colID, string[] cellsCol, string openingVar, string closingVar, bool Numericalcheck, bool VariableCheck)
        {
            for (int owbb = 0; owbb < WorkBooks.Count; owbb++)
            {
                oWB = (Excel._Workbook)WorkBooks[owbb];
                for (int owss = 1; owss <= oWB.Worksheets.Count; owss++)
                {
                    oWS = (Excel._Worksheet)oWB.Worksheets.get_Item(owss);
                    foreach (Excel.Range fila in oWS.UsedRange.Rows)
                    {
                        actualizarSID((string)(oWS.Cells[fila.Row, colSour] as Excel.Range).Value, (string)(oWS.Cells[fila.Row, colID] as Excel.Range).Value, openingVar, closingVar);
                        if (cellValueS != null)
                        {
                            variablesInString = getVariablesinString(cellValueS);
                            if (Numericalcheck == true)
                            {
                                numerosenSource = getNumbers(cellValueS, variablesInString);
                            }
                            switch (cellsCol.Length)
                            {
                                case 1:
                                    if (cellsCol[0] != "")
                                    {
                                        if (VariableCheck)
                                        {
                                            issuefounded = VarCompare((string)(oWS.Cells[fila.Row, int.Parse(cellsCol[0])] as Excel.Range).Value, owbb);
                                        }
                                        if ((issuefounded == false)&&(Numericalcheck == true))
                                        {
                                            NumCompare((string)(oWS.Cells[fila.Row, int.Parse(cellsCol[0])] as Excel.Range).Value, owbb);
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.Forms.MessageBox.Show("Please make sure that you fill \nin all the mandatory fields properly.");
                                    }
                                    break;
                                case 2:
                                    int[] columnass = { -1, -1 };
                                    if ((cellsCol[0] != null) && (cellsCol[0] != ""))
                                    {
                                        columnass[0] = int.Parse(cellsCol[0]);
                                    }
                                    if ((cellsCol[1] != null) && (cellsCol[1] != ""))
                                    {
                                        columnass[1] = int.Parse(cellsCol[1]);
                                    }

                                    if ((columnass[0] != -1) && (columnass[1] != -1))
                                    {
                                        if (columnass[0] > columnass[1])
                                        {
                                            int exc = columnass[0];
                                            columnass[0] = columnass[1];
                                            columnass[1] = exc;
                                        }
                                        for (int b = columnass[0]; b <= columnass[1]; b++)
                                        {
                                            if (VariableCheck)
                                            {
                                                issuefounded = VarCompare((string)(oWS.Cells[fila.Row, int.Parse(cellsCol[0])] as Excel.Range).Value, owbb);
                                            }
                                            if ((issuefounded == false) && (Numericalcheck == true))
                                            {
                                                NumCompare((string)(oWS.Cells[fila.Row, b] as Excel.Range).Value, owbb);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if ((columnass[0] == -1) && (columnass[1] == -1))
                                        {
                                            System.Windows.Forms.MessageBox.Show("Please make sure that you fill \nin all the mandatory fields properly.");
                                        }
                                        else
                                        {
                                            if (columnass[0] == -1)
                                            {
                                                columnass[0] = columnass[1];
                                            }
                                            if (VariableCheck)
                                            {
                                                issuefounded = VarCompare((string)(oWS.Cells[fila.Row, int.Parse(cellsCol[0])] as Excel.Range).Value, owbb);
                                            }
                                            if ((issuefounded == false) && (Numericalcheck == true))
                                            {
                                                NumCompare((string)(oWS.Cells[fila.Row, columnass[0]] as Excel.Range).Value, owbb);
                                            }
                                        }
                                    }
                                    break;
                                default:
                                    foreach (string cols in cellsCol)
                                    {
                                        if ((cols != null) && (cols != ""))
                                        {
                                            if (VariableCheck)
                                            {
                                                issuefounded = VarCompare((string)(oWS.Cells[fila.Row, int.Parse(cellsCol[0])] as Excel.Range).Value, owbb);
                                            }
                                            if ((issuefounded == false) && (Numericalcheck == true))
                                            {
                                                NumCompare((string)(oWS.Cells[fila.Row, int.Parse(cols)] as Excel.Range).Value, owbb);
                                            }
                                        }
                                    }
                                    break;
                            }
                        }
                    }
                }
            }
        }

        private static List<string> getVariablesinString(string madre)
        {
            List<string> variables = new List<string>();
            int index = 0;
            int numeroderepeticiones;

            if (CountStringOccurrences(madre, openingVar) >= CountStringOccurrences(madre, closingVar))
            {
                numeroderepeticiones = CountStringOccurrences(madre, closingVar);
            }
            else
            {
                numeroderepeticiones = CountStringOccurrences(madre, openingVar);
            }

            for (int i = 0; i < numeroderepeticiones; i++)
            {
                int indiceopening = madre.IndexOf(openingVar, index);
                int indiceclosing = madre.IndexOf(closingVar, indiceopening) + closingVar.Length;

                while ((madre.IndexOf(openingVar, indiceopening + openingVar.Length) < indiceclosing) && (madre.IndexOf(openingVar, indiceopening + openingVar.Length) != -1))
                {
                    indiceopening = madre.IndexOf(openingVar, indiceopening + openingVar.Length);
                }

                variables.Add(madre.Substring(indiceopening, indiceclosing - indiceopening));
                index = indiceclosing;

            }

            return variables;
        }
        private static int CountStringOccurrences(string text, string pattern)
        {
            int count = 0;
            int i = 0;
            while ((i = text.IndexOf(pattern, i)) != -1)
            {
                i += pattern.Length;
                count++;
            }
            return count;
        }
        public static string getNumbers(string cadena, List<string> variables)
        {
            List<string> list = new List<string>();
            string cadenaresultante = deleteVariablesfromString(cadena, variables);
            for (int i = 0; i < cadenaresultante.Length; i++)
            {
                if (Char.IsNumber(cadenaresultante[i]) == true)
                {
                    if ((int.Parse(cadenaresultante[i].ToString()) >= 0) && (int.Parse(cadenaresultante[i].ToString()) <= 9))
                    {
                        list.Add(cadenaresultante[i].ToString());
                    }
                }
            }
            string numbers = string.Join(",", list.ToArray());
            return numbers;
        }
        private static string deleteVariablesfromString(string cadena, List<string> variables)
        {
            string cadenaresultante = cadena;
            for (int i = 0; i < variables.Count; i++)
            {
                if (cadenaresultante.Contains(variables[i]))
                {
                    cadenaresultante = cadenaresultante.Replace(variables[i], "");
                }
            }
            return cadenaresultante;
        }
        private static bool VarCompare(string cellValueL, int numerodeOWB)
        {
            if (cellValueL != null)
            {
                if (!getVariablesinString(cellValueL).SequenceEqual(variablesInString))
                {
                    WorkSheetReport = (Excel.Worksheet)oWBReport.Worksheets.get_Item(numerodeOWB + 1);
                    WorkSheetReport.Cells[WorkSheetReport.UsedRange.Rows.Count + 1, 1] = cellID;
                    WorkSheetReport.Cells[WorkSheetReport.UsedRange.Rows.Count, 2] = cellValueS;
                    WorkSheetReport.Cells[WorkSheetReport.UsedRange.Rows.Count, 3] = cellValueL;
                    return true;
                }
            }
            return false;
        }
        private static void NumCompare(string cellValueL, int numerodeOWB)
        {
            if (cellValueL != null)
            {
                if (getNumbers(cellValueL, getVariablesinString(cellValueL)) != numerosenSource)
                {
                    WorkSheetReport = (Excel.Worksheet)oWBReport.Worksheets.get_Item(numerodeOWB + 1);
                    WorkSheetReport.Cells[WorkSheetReport.UsedRange.Rows.Count + 1, 1] = cellID;
                    WorkSheetReport.Cells[WorkSheetReport.UsedRange.Rows.Count, 2] = cellValueS;
                    WorkSheetReport.Cells[WorkSheetReport.UsedRange.Rows.Count, 3] = cellValueL;
                }
            }
        }
    }
}
