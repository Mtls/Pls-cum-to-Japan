using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace OpenFileTest
{
    public partial class Form1 : Form
    {
        string file = "";
        bool fileSelected = false;

        List<String> UniqueStudents = new List<String>();
        List<String> UniqueUnits = new List<String>();
        List<String> UniqueExamLocations = new List<String>();
        List<Sitting> sittings = new List<Sitting>();

        Excel.Application excel = null;
        Excel.Workbook wkb = null;
        Excel.Worksheet sheet1 = null;
        Excel.Range range = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Open File was clicked....");
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) //Test result.
            {
                fileSelected = true;
                file = openFileDialog1.FileName;
                Console.WriteLine(file);
                MessageBox.Show(file, "File");
                //file = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (fileSelected == true)
            {
                try
                {
                    button1.Enabled = false;
                    button2.Enabled = false;
                    Console.WriteLine("Import was clicked.... trying...");
                    //Initialise excel stuff
                    excel = new Excel.Application();
                    wkb = OpenBook(excel, file, true, false, false);
                    sheet1 = wkb.Sheets["Sheet1"] as Excel.Worksheet;
                    //Excel.Worksheet sheet2 = wkb.Sheets["Sheet2"] as Excel.Worksheet;


                    range = sheet1.UsedRange;

                    int rowCount = range.Rows.Count;
                    int SittingCount = 0;

                    Sitting tempSitting = new Sitting();
                    string tempStu;
                    string tempUnit;
                    string tempLoc;

                    //read in the sittings
                    Console.WriteLine("Importing data, please wait...");
                    Console.WriteLine("Read:");
                    for (int i = 2; i <= rowCount; i++)
                    {
                        //create a new sitting
                        ///tempSitting = new Sitting();
                        //Set data
                        tempStu = range.Cells[i, 1].Value2.ToString();
                        tempUnit = range.Cells[i, 2].Value2.ToString();
                        tempLoc = range.Cells[i, 3].Value2.ToString();
                        ///tempSitting.StudentO.StudentCode = range.Cells[i, 1].Value2.ToString();
                        ///tempSitting.UnitO.UnitCode = range.Cells[i, 2].Value2.ToString();
                        ///tempSitting.LocationO.LocationCode = range.Cells[i, 3].Value2.ToString();

                        ///sittings.Add(tempSitting);
                        SittingCount++;

                        if (UniqueStudents.Contains(tempStu) == false)
                            UniqueStudents.Add(tempStu);
                        
                        ///if (UniqueStudents.Contains(tempSitting.StudentO.StudentCode) == false)
                        ///    UniqueStudents.Add(tempSitting.StudentO.StudentCode);

                        if (UniqueUnits.Contains(tempUnit) == false)
                            UniqueUnits.Add(tempUnit);

                        ///if (UniqueUnits.Contains(tempSitting.UnitO.UnitCode) == false)
                        ///    UniqueUnits.Add(tempSitting.UnitO.UnitCode);

                        if (UniqueExamLocations.Contains(tempLoc) == false)
                            UniqueExamLocations.Add(tempLoc);

                        ///if (UniqueExamLocations.Contains(tempSitting.LocationO.LocationCode) == false)
                        ///    UniqueExamLocations.Add(tempSitting.LocationO.LocationCode);

                        if (SittingCount % 1000 == 0)
                        {
                            Console.WriteLine("{0}", SittingCount);
                        }

                    }

                    //DisplaySittings(sittings);
                    Console.WriteLine("Imported {0} Sittings", SittingCount);
                    Console.WriteLine("Unique Students: {0}", UniqueStudents.Count);
                    Console.WriteLine("Unique Units: {0}", UniqueUnits.Count);
                    Console.WriteLine("Unique Exam Locations: {0}", UniqueExamLocations.Count);

                    foreach (var loc in UniqueExamLocations)
                    {
                        Console.WriteLine(loc);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    if (range != null)
                    {
                        Console.WriteLine("Releasing range");
                        ReleaseRCM(range);
                    }
                    if (sheet1 != null)
                    {
                        Console.WriteLine("Releasing sheet1");
                        ReleaseRCM(sheet1);
                    }
                    if (wkb != null)
                    {
                        Console.WriteLine("Releasing wkb");
                        ReleaseRCM(wkb);
                    }
                    if (excel != null)
                    {
                        Console.WriteLine("Releasing excel");
                        ReleaseRCM(excel);
                    }
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        public static Excel.Workbook OpenBook(Excel.Application excelInstance, string fileName, bool readOnly, bool editable, bool updateLinks)
        {
            Excel.Workbook book = excelInstance.Workbooks.Open(
                fileName, updateLinks, readOnly,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, editable, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
            return book;
        }

        public static void ReleaseRCM(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch
            {
            }
            finally
            {
                o = null;
            }
        }

    }

    public class Sitting
    {
        public Student StudentO = new Student();
        public Unit UnitO = new Unit();
        public Location LocationO = new Location();
        public bool pairChecked = false;
    }

    public class Unit
    {
        public string UnitCode { get; set; }
        public List<Student> students = new List<Student>();

    }

    public class Student
    {
        public string StudentCode { get; set; }
        public List<Unit> units = new List<Unit>();
    }

    public class Location
    {
        public string LocationCode { get; set; }
    }
}