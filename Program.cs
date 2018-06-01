using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Text;
using System.Threading.Tasks;

namespace Combine2Code
{
    public class Program
    {
        public static List<Record> subList = new List<Record>();
        public static List<Record> masterList = new List<Record>();
        public static List<Record> combinedList = new List<Record>();
        public static List<string> orphans = new List<string>();

        static void Main(string[] args)
        {
            if (!ReadFilesToList())
                Console.WriteLine("Something went wrong");
            if(!CombineLists())
                Console.WriteLine("Something went wrong");
            if(!FindOrphans())
                Console.WriteLine("Something went wrong");
            if (!ExportToExcel())
                Console.WriteLine("Something went wrong");
            Console.WriteLine("Process has completed - please check out put folder");
        }

        public static bool ReadFilesToList()
        {
            try
            {
                string[] files = Directory.GetFiles(ConfigurationManager.AppSettings["FolderHome"]);
                foreach (string file in files)
                {
                    if (file.Contains("Sub_csv") )
                    {
                        subList = ReadBPTFile(file);
                    }
                    if (file.Contains("Master_csv"))
                    {
                        masterList = ReadMasterFile(file);
                    }
                }
                if(subList.Count > 0 && masterList.Count > 0)
                    return true;
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        public static List<Record> ReadMasterFile(string file)
        {
            List<Record> list = new List<Record>();
            try
            {
                string[] lines = File.ReadAllLines(file);
                for(int i = 0; lines.Length > i; i++)
                {
                    if (i != 0)
                        list.Add(Record.MasterFromCsv(lines[i]));
                }
                return list;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        public static List<Record> ReadBPTFile(string file)
        {
            List<Record> list = new List<Record>();
            try
            {
                string[] lines = File.ReadAllLines(file);
                for (int i = 0; lines.Length > i; i++)
                {
                    if (i != 0)
                        list.Add(Record.BPTFromCsv(lines[i]));
                }
                return list;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        public static bool CombineLists()
        {
            try
            {
                combinedList = new List<Record>();
                foreach (Record sub in subList)
                {
                    foreach (Record master in masterList)
                    {
                        if (master.ID == sub.ID)
                        {
                            combinedList.Add(new Record()
                            {
                                ID = master.ID,
                            });
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        public static bool FindOrphans()
        {
            try
            {
                orphans = new List<string>();
                List<string> subIdOnly = new List<string>();
                List<string> masterIdOnly = new List<string>();
                foreach (Record sr in subList)
                {
                    subIdOnly.Add(sr.ID.ToString());
                }
                foreach (Record sr in masterList)
                {
                    masterIdOnly.Add(sr.ID.ToString());
                }
                var exceptBPT = subIdOnly.Except(masterIdOnly);
                var exceptMaster = masterIdOnly.Except(subIdOnly);
                foreach (string id in exceptBPT)
                {
                    orphans.Add(id + " In sub list ONLY");
                }
                foreach (string id in exceptMaster)
                {
                    orphans.Add(id + " In master list ONLY");
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        public static bool ExportToExcel()
        {
            try
            {
                string fileName = String.Format(@"\ExportCombinedMasterAndSub_{0}.xlsx", DateTime.Now.ToString("yyyyMMddhhmmss"));
                string workbookPath = ConfigurationManager.AppSettings["FolderHome"] + fileName;
                var excel = new Application();
                Workbook workbook = excel.Workbooks.Add(true);
                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.Workbooks.Add();

                System.Data.DataTable masterTable = RecordListToDataTable(combinedList);
                AddTableToWorkSheet(workbook, masterTable, "Master_Results");

                System.Data.DataTable orphanTable = StringListToDataTable(orphans);
                Worksheet mismatchSheet = AddTableToWorkSheet(workbook, orphanTable, "Orphans");

                workbook.SaveAs(workbookPath);
                workbook.Close();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        public static System.Data.DataTable StringListToDataTable(List<string> list)
        {
            try
            {
                System.Data.DataTable MethodResult = null;
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Columns.Add("Orphans");

                foreach (string s in list)
                {
                    DataRow dr = dt.NewRow();
                    dr[0] = s;
                    dt.Rows.Add(dr);

                }
                dt.AcceptChanges();
                MethodResult = dt;
                return MethodResult;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        public static System.Data.DataTable RecordListToDataTable(List<Record> list)
        {
            try
            {
                System.Data.DataTable MethodResult = null;
                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Columns.Add("ID");
                //Add your additional column fields here
                foreach (Record sr in list)
                {
                    DataRow dr = dt.NewRow();
                    dr[0] = sr.ID;
                    //Add the extra rows to the table here
                    dt.Rows.Add(dr);
                }
                dt.AcceptChanges();
                MethodResult = dt;
                return MethodResult;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }

        public static Worksheet ExportTableToExcel(System.Data.DataTable tbl,Worksheet workSheet)
        {
            try
            {
                if (tbl == null || tbl.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");
                // load excel, and create a new workbook
                var excelApp = new Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;
                excelApp.Workbooks.Add();
                workSheet = excelApp.ActiveSheet;
                // column headings
                for (var i = 0; i < tbl.Columns.Count; i++)
                {
                    workSheet.Cells[1, i + 1] = tbl.Columns[i].ColumnName;
                }
                // rows
                for (var i = 0; i < tbl.Rows.Count; i++)
                {
                    // to do: format datetime values before printing
                    for (var j = 0; j < tbl.Columns.Count; j++)
                    {
                        workSheet.Cells[i + 2, j + 1] = tbl.Rows[i][j];
                    }
                }
                return workSheet;
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }

        public static Worksheet AddTableToWorkSheet(Workbook wb, System.Data.DataTable t1, string name)
        {
            try
            {
                Sheets sheets = wb.Sheets;
                Worksheet newSheet = sheets.Add();
                newSheet.Name = name;
                int iCol = 0;
                foreach (DataColumn c in t1.Columns)
                {
                    iCol++;
                    newSheet.Cells[1, iCol] = c.ColumnName;
                }
                int iRow = 0;
                foreach (DataRow r in t1.Rows)
                {
                    iRow++;
                    // add each row's cell data...
                    iCol = 0;
                    foreach (DataColumn c in t1.Columns)
                    {
                        iCol++;
                        newSheet.Cells[iRow + 1, iCol] = r[c.ColumnName];
                    }
                }

                return newSheet;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}
