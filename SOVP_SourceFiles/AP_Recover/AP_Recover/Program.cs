using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;

namespace AccelerometerProcessing_FiveSecond
{



    class Program
    {
        static List<TeacherFolderDateTime> TeacherFolderDateTimeList = new List<TeacherFolderDateTime>();

        /// <summary>
        /// questions, concerns, comments? why the f^&* are we using a bgworker? plz ask Laur to ask Will; thx
        /// </summary>
        ///        
        [STAThread]
        static void Main(string[] args)
        {

            

            UserMessages();

            string parentSchoolFolder = FiveSecCheck();

            List<Student> studentData = new List<Student>();            
            studentData = getWorkSheetsForCurrentTeacherFolder();

            fire_worker_no_multithreading(parentSchoolFolder, studentData);

            Console.WriteLine("Complete!");
            Console.ReadLine();
        }

        static void fire_worker_no_multithreading(string parentSchoolFolder, List<Student> studentData)
        {
          
            createCleanCSV(parentSchoolFolder, studentData);

            // creates master merged file per folder
            createMasterMergePerFolder(parentSchoolFolder);

            // creates highest Level folder merged
            createHighestMerge(parentSchoolFolder);
        }

        static void createCleanCSV(string schoolFoo, List<Student> studentData)
        {

            HashSet<string> worksheetTabsThatwereNotFound = new HashSet<string>();

            var teacherfolders = Directory.GetDirectories(schoolFoo);

            foreach (var folderPath in teacherfolders)
            {

                string filepath = folderPath;
                DirectoryInfo d = new DirectoryInfo(filepath);

                foreach (var file in d.GetFiles())
                {

                    // load up excel file with student schedule for merging

                    string activityNumber = "0"; // reset to zero upon new file
                    var rowCounter = 0; // ignore first 11 rows of csv and reset to 0 upon next file

                    string xAxisReg1 = "0";
                    string xAxisReg2 = "0";
                    string xAxisReg3 = "0";
                    string xAxisReg4 = "0";

                    string dateInLine;

                    string timeInLine;

                    string xAxisData;

                    string yAxisData;

                    string zAxisData;

                    string VecMagData;

                    List<DateAndActivityStack> daasList = new List<DateAndActivityStack>();

                    var csv = new StringBuilder(); // new csv SB per file

                    using (var reader = new StreamReader(file.FullName))
                    {


                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();

                            // for each line read increase counter
                            rowCounter++;

                            if (rowCounter == 12 || rowCounter > 12)
                            {



                                // file.Name.TrimEnd('5', 's', 'e', 'c', '.', 'c', 's', 'v') <---- use this for the student number
                                string studentIdtolookfor = file.Name.ToString().Substring(0, file.Name.ToString().Length - 8);

                                Student result = studentData.Find(x => x.StudentID == studentIdtolookfor);

                                if (result is null)
                                {
                                    worksheetTabsThatwereNotFound.Add(studentIdtolookfor + " was not found!");
                                }

                                string lineString = line;

                                string[] lineSplit = lineString.Split(',');

                                dateInLine = lineSplit[0];

                                timeInLine = lineSplit[1];

                                xAxisData = lineSplit[2];

                                yAxisData = lineSplit[3];

                                zAxisData = lineSplit[4];

                                // string stepsData = lineSplit[5];

                                // string incOffData = lineSplit[6];

                                //string incStandData = lineSplit[7];

                                // string incSitData = lineSplit[8];

                                //  string incLyingData = lineSplit[9];

                                VecMagData = lineSplit[11];

                                string xAxisChoice = "0";

                                ////////////////////////////// x axis column choices

                                if (Int32.Parse(xAxisData) >= 0 && Int32.Parse(xAxisData) <= 8)
                                {
                                    xAxisChoice = "1";
                                }
                                if (Int32.Parse(xAxisData) >= 9 && Int32.Parse(xAxisData) <= 190)
                                {
                                    xAxisChoice = "2";
                                }
                                if (Int32.Parse(xAxisData) >= 191 && Int32.Parse(xAxisData) <= 334)
                                {
                                    xAxisChoice = "3";
                                }
                                if (Int32.Parse(xAxisData) >= 335)
                                {
                                    xAxisChoice = "4";
                                }

                                ///////////////////////////////////// stakeholder mind change switch old 1-4 choice into registers and write columns


                                if (xAxisChoice == "1")
                                {
                                    xAxisReg1 = "1";
                                    xAxisReg2 = "0";
                                    xAxisReg3 = "0";
                                    xAxisReg4 = "0";
                                }
                                if (xAxisChoice == "2")
                                {
                                    xAxisReg1 = "0";
                                    xAxisReg2 = "1";
                                    xAxisReg3 = "0";
                                    xAxisReg4 = "0";
                                }
                                if (xAxisChoice == "3")
                                {
                                    xAxisReg1 = "0";
                                    xAxisReg2 = "0";
                                    xAxisReg3 = "1";
                                    xAxisReg4 = "0";
                                }
                                if (xAxisChoice == "4")
                                {
                                    xAxisReg1 = "0";
                                    xAxisReg2 = "0";
                                    xAxisReg3 = "0";
                                    xAxisReg4 = "1";
                                }


                                if (result is null)
                                {
                                    activityNumber = "null";
                                }

                                if(result != null) // student was found in schedule set now attach the activity number for the correct date and time
                                {
                                    foreach (var row  in result.StudentRows)
                                    {
                                        if (row.col1 == dateInLine) // confirm date
                                        {

                                            if (timeInLine[0] == '0') // check for leading zero causing mismatch in time stamps (can remove later if data gets matched up)
                                            {
                                               timeInLine = timeInLine.Substring(1);
                                            }

                                            if (row.col2 == timeInLine) // confirm time, update activity number
                                            {
                                                activityNumber = row.col3;
                                            }
                                        }
                                    }
                                }

                                 




                                // string stepsData = lineSplit[5];

                                // string incOffData = lineSplit[6];

                                //string incStandData = lineSplit[7];

                                // string incSitData = lineSplit[8];

                                //  string incLyingData = lineSplit[9];


                                var getACount = (folderPath.ToString().Substring(folderPath.LastIndexOf('\\') + 1));



                                // try   wth
                                // {


                                // wth original
                                // DateAndActivityStack daas = new DateAndActivityStack(dateInLine.ToString(), timeInLine.ToString(), file.Name.ToString().Substring(0, file.Name.ToString().Length - 8), folderPath.ToString().Substring(folderPath.LastIndexOf('\\') + 1).ToString(), activityNumber.ToString(), Int32.Parse(xAxisData), Int32.Parse(yAxisData), Int32.Parse(zAxisData), Single.Parse(VecMagData));
                                DateAndActivityStack daas = new DateAndActivityStack(dateInLine.ToString(), timeInLine.ToString(), file.Name.ToString().Substring(0, file.Name.ToString().Length - 8), folderPath.ToString().Substring(folderPath.LastIndexOf('\\') + 1).ToString(), activityNumber.ToString(), Int32.Parse(xAxisReg1), Int32.Parse(xAxisReg2), Int32.Parse(xAxisReg3), Int32.Parse(xAxisReg4));
                                daasList.Add(daas);


                                // debugging
                                // Console.WriteLine( "Teacher:" + daas.TeacherFolderName.ToString() + "\n" +"Student:" + daas.StudentNumber.ToString() + "\n"+"Time:" + daas.Time.ToString());
                                // }
                                // catch(Exception e)
                                // {

                                //   Console.WriteLine(e.ToString() + "\n");

                                //   Console.WriteLine("\n\n\n ************* \n failed due to something in this:" );

                                // Console.ReadLine();
                                // }  





                                // this will write final csv
                                // ','+ timeInLine + ','+ xAxisData + 
                                //csv.AppendLine(dateInLine + ',' + file.Name.ToString().Substring(0, file.Name.ToString().Length - 8) + ',' + folderPath.ToString().Substring(folderPath.LastIndexOf('\\')+1) + ',' + activityNumber + ',' + xAxisReg1 + ',' + xAxisReg2 + ',' + xAxisReg3 + ',' + xAxisReg4);


                                // Console.ReadLine();
                            }


                            // may do final csv here after
                            // csv.AppendLine(dateInLine + ',' + file.Name.ToString().Substring(0, file.Name.ToString().Length - 8) + ',' + folderPath.ToString().Substring(folderPath.LastIndexOf('\\') + 1) + ',' + activityNumber + ',' + xAxisReg1 + ',' + xAxisReg2 + ',' + xAxisReg3 + ',' + xAxisReg4);
                        }

                        // per each csv merge here

                        List<DateAndActivityStack> rowsToBeMerged = new List<DateAndActivityStack>();

                        foreach (DateAndActivityStack row in daasList)
                        {
                            if (rowsToBeMerged.Count == 0)
                            {
                                // first row to be added and compared too

                                rowsToBeMerged.Add(row);

                                continue; // break to next iteration
                            }


                            // matching date and activity number so add to merge list 
                            if (row.Date == rowsToBeMerged.First().Date && row.ActivityNumber == rowsToBeMerged.First().ActivityNumber)
                            {
                                // add rows until break in activityNumber or Date
                                rowsToBeMerged.Add(row);
                            }
                            else // break occured operate on rows
                            {
                                // new date or activity number detected, merge the existing list, sum the regs, empty the list, add this one as the new start
                                float regESUM = 0;
                                float regFSUM = 0;
                                float regGSUM = 0;
                                float regHSUM = 0;

                                foreach (DateAndActivityStack item in rowsToBeMerged)
                                {                                  
                                    regESUM += item.AxisRegSumE;
                                    regFSUM += item.AxisRegSumF;
                                    regGSUM += item.AxisRegSumG;
                                    regHSUM += item.AxisRegSumH;
                                }

                                //div by 12
                                regESUM = regESUM / 12.0f;
                                regFSUM = regFSUM / 12.0f;
                                regGSUM = regGSUM / 12.0f;
                                regHSUM = regHSUM / 12.0f;

                                // write it
                                csv.AppendLine(rowsToBeMerged.First().Date + ',' + rowsToBeMerged.First().StudentNumber + ',' + rowsToBeMerged.First().TeacherFolderName + ',' + rowsToBeMerged.First().ActivityNumber + ',' + regESUM + ',' + regFSUM + ',' + regGSUM + ',' + regHSUM);
                                // line got written clear for restart
                                rowsToBeMerged.Clear();
                            }
                            // write csv line // wth heel

                            //rowsToBeMerged.Add(row); // continue on
                            //string dateInLine = lineSplit[0];                           

                        }
                    }

                    // if breakdown of ind E's later

                    //string[] lines = csv.ToString().Split(',');

                    //foreach (string someLine in lines)
                    //{

                    //}

                    File.WriteAllText(filepath + "\\" + "merged_" + file.Name, csv.ToString());

                    Console.WriteLine(file + " processed");
                }

            }

            foreach (var item in worksheetTabsThatwereNotFound)
            {
                Console.WriteLine(item);
            }


        }

        static void createMasterMergePerFolder(string schoolFoo)
        {
            var teacherfolders = Directory.GetDirectories(schoolFoo);


            foreach (var folderPath in teacherfolders)
            {

                var csv = new StringBuilder(); // new csv SB per folder
                string filepath = folderPath;
                DirectoryInfo d = new DirectoryInfo(filepath);

                foreach (var file in d.GetFiles())
                {
                    // if it's a merge file bring in the csv data
                    if (file.Name.Contains("merged"))
                    {
                        using (var reader = new StreamReader(file.FullName))
                        {

                            while (!reader.EndOfStream)
                            {
                                var line = reader.ReadLine();

                                csv.AppendLine(line);

                            }
                        }


                    }
                    else
                    {
                        // skipping
                    }


                }

                File.WriteAllText(filepath + "\\" + "master_merger.csv", csv.ToString());

                Console.WriteLine(filepath + " processed");
            }


        }

        static void createHighestMerge(string schoolFoo)
        {
            var teacherfolders = Directory.GetDirectories(schoolFoo);
            var csv = new StringBuilder(); // new csv SB per folder

            foreach (var folderPath in teacherfolders)
            {

                string filepath = folderPath;
                DirectoryInfo d = new DirectoryInfo(filepath);

                foreach (var file in d.GetFiles())
                {
                    // if it's a merge file bring in the csv data
                    if (file.Name.Contains("master_merger"))
                    {
                        using (var reader = new StreamReader(file.FullName))
                        {

                            while (!reader.EndOfStream)
                            {
                                var line = reader.ReadLine();

                                csv.AppendLine(line);

                            }
                        }


                    }
                    else
                    {
                        // skipping
                    }


                }


            }
            File.WriteAllText(schoolFoo + "\\" + "highest_master_merger.csv", csv.ToString());

            Console.WriteLine("highest merge processed");

            Console.WriteLine("Completed press enter!");     

        }
        
        static List<Student> getWorkSheetsForCurrentTeacherFolder()
        {
            //////////////////////////////////////////////////////////
            List<Student> StudentListPerStudentScheduleExcelSheet = new List<Student>();

            // Show the dialog and get result.        // getting xls file path from user    
            OpenFileDialog openFileDialog = new OpenFileDialog();

            DialogResult result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
            }
            // Console.WriteLine(result); // <-- For debugging use.

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();  // Creates a new Excel Application
            excelApp.Visible = false;  // Makes Excel visible to the user.           
                                       // The following code opens an existing workbook
            string workbookPath = openFileDialog.FileName;
            Workbook excelWorkbook = null;
            try
            {
                excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0,
                false, 5, "", "", false, XlPlatform.xlWindows, "", true,
                false, 0, true, false, false);
            }
            catch
            {
                //Create a new workbook if the existing workbook failed to open.
                excelWorkbook = excelApp.Workbooks.Add();
            }
            // The following gets the Worksheets collection
            Microsoft.Office.Interop.Excel.Sheets excelSheets = excelWorkbook.Worksheets;
            // Console.WriteLine(excelSheets.Count.ToString()); //dat count


            foreach (Worksheet worksheet in excelSheets)
            {
                Console.WriteLine(worksheet.Name.ToString());

                Student stackStudent = new Student();

                stackStudent.StudentID = worksheet.Name.ToString();



                //Get the used Range
                Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;

                // Console.WriteLine(worksheet.Rows);
                foreach (Microsoft.Office.Interop.Excel.Range row in usedRange.Rows)
                {
                    //Do something with the row.

                    usedRange.Columns.AutoFit();

                    //Ex. Iterate through the row's data and put in a string array
                    String[] rowData = new String[row.Columns.Count];

                    Row stackRow = new Row();

                    for (int i = 0; i < row.Columns.Count; i++)
                    {
                        try
                        {
                            // write individual cell data
                            //  Console.Write(rowData[i] = row.Cells[1, i + 1].Text.ToString());

                            if (i == 0)
                            {
                                stackRow.col0 = row.Cells[1, i + 1].Text.ToString();
                            }
                            if (i == 1)
                            {
                                stackRow.col1 = row.Cells[1, i + 1].Text.ToString();
                            }
                            if (i == 2)
                            {
                                stackRow.col2 = row.Cells[1, i + 1].Text.ToString();
                            }
                            if (i == 3)
                            {
                                stackRow.col3 = row.Cells[1, i + 1].Text.ToString();
                            }
                        }
                        catch
                        {
                            continue;
                        }


                    }

                    // add row before 

                    stackStudent.StudentRows.Add(stackRow);

                    // Console.WriteLine("\n");
                }

                StudentListPerStudentScheduleExcelSheet.Add(stackStudent);

            }

            // try closing excel
            excelWorkbook.Close(0);
            excelApp.Quit();

            return StudentListPerStudentScheduleExcelSheet;

        }      

        static void UserMessages()
        {
            Console.WriteLine("First, you will select the parent-most folder 'i.e. school' of the .csv files you would like to contexualize. Hit Enter");
            Console.ReadLine();
        }

        static string GetParticipantScheduleDir()
        {
            string someDir = GetDir();

            return someDir;

        }

        static string FiveSecCheck() // this checks name of csv's
        {
            string schoolFoo = GetDir(); // get parent folder

            var teacherfolders = Directory.GetDirectories(schoolFoo);

            // get start/end date and time for later from each teacher folder also write this to diagnostic report 

            foreach (var folder in teacherfolders)
            {

                CheckAgainForThisFolder:

                Console.WriteLine("Would you like to add sample from: " + folder.ToString().Split('\\').Last() + "? (y/n)");
                string choice = Console.ReadLine();

                if (choice == "y")
                {
                    Console.WriteLine("Enter start date as it appears in dataset for: " + folder.ToString().Split('\\').Last());
                    string sDate = Console.ReadLine();
                    Console.WriteLine("Enter start time as it appears in dataset for: " + folder.ToString().Split('\\').Last());
                    string sTime = Console.ReadLine();
                    Console.WriteLine("Enter end date as it appears in dataset for: " + folder.ToString().Split('\\').Last());
                    string eDate = Console.ReadLine();
                    Console.WriteLine("Enter end time as it appears in dataset for: " + folder.ToString().Split('\\').Last());
                    string eTime = Console.ReadLine();

                    TeacherFolderDateTimeList.Add(new TeacherFolderDateTime(folder.ToString().Split('\\').Last(), sDate, sTime, eDate, eTime));

                    goto CheckAgainForThisFolder;
                }
                else
                {
                    // do nothing iterate to next teacher folder
                }

                


                
            }

            

            foreach (var folderPath in teacherfolders)
            {

                string filepath = folderPath;
                DirectoryInfo d = new DirectoryInfo(filepath);

                foreach (var file in d.GetFiles())
                {
                    if (!file.Name.Contains("5sec.csv"))
                    {
                        Console.WriteLine(file.Name + ":  ths file is not in the 5 second format. Please fix, then re-run the program.");
                        Console.ReadLine();

                        //System.Windows.Forms.Application.Exit();
                        Environment.Exit(0);

                    }

                    //  Console.WriteLine(file); debug to see what files it's enumerating
                }

            }
            return schoolFoo;
        }

        static string GetDir()
        {

            string schoolfolder = "";

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    // string[] files = Directory.GetFiles(fbd.SelectedPath);
                    schoolfolder = fbd.SelectedPath;
                    //System.Windows.Forms.MessageBox.Show("Files found: " + files.Length.ToString(), "Message");
                }


            }

            return schoolfolder;

        }

    }
}
