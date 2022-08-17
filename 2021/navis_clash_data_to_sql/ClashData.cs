//Navisworks API references
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Automation;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Internal.ApiImplementation;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Text.RegularExpressions;



namespace ClashData
{
    //plugin attributes require Name, DeveloperID and optional parameters
    [PluginAttribute("clash_sql_export", "Derek B", DisplayName = "Export Clashes to SQL", ToolTip = "Exports clashs data to SQL Database.", ExtendedToolTip = "Version 2022.1.0.1")]
    [AddInPluginAttribute(AddInLocation.AddIn, Icon = "C:\\Program Files\\Autodesk\\Navisworks Manage 2021\\Plugins\\clash_sql_export\\resources\\16x16_sql_export_img.bmp",
        LargeIcon = "C:\\Program Files\\Autodesk\\Navisworks Manage 2021\\Plugins\\clash_sql_export\\resources\\32x32_sql_export_img.bmp")]

    public class ClashData : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;
            DocumentClash documentClash = document.GetClash();
            var allTests_copy = documentClash.TestsData.CreateCopy();
            //DocumentModels docModel = document.Models;

            string filePath = Autodesk.Navisworks.Api.Application.MainDocument.FileName.ToString();
            var regMatchNum = @"([0-9]{5})";
            Match regMatch = Regex.Match(filePath, regMatchNum);

            if (regMatch.Success)
            {
                string projectNum = regMatch.Value;
                MessageBox.Show(projectNum, "Project Number from File Name.");

                //-----------------------------------------------------------------------------------------
                //Check if Clash Test have even been created
                //If no clash tests created, exit program
                int check = allTests_copy.Tests.Count;
                if (check == 0)
                {
                    MessageBox.Show("No clash tests currently exist!");
                }
                else
                {
                    ////Get save folder for XML files.Used with old Python Scipt-----------------------------------------------------
                    //FolderBrowserDialog fbd = new FolderBrowserDialog();
                    //fbd.Description = "Select folder where the clash data (XML) files will be saved";
                    //string sSelectedPath = "";
                    //if (fbd.ShowDialog() == DialogResult.OK)
                    //{
                    //    sSelectedPath = fbd.SelectedPath;
                    //}
                    //else
                    //{

                    //}
                    ////---------------------------------------------------------------------------------------------------------

                    //Clash Report date------------------------------------------------------------
                    //string export_date = DateTime.Now.ToString("yyyy.MM.dd");
                    var dqr = new date_question.DateQuestion();
                    dqr.ShowDialog();

                    string export_date = "";

                    if (dqr.dateQuestResponse == "Yes")
                    {
                        export_date = DateTime.Now.ToString("yyyy.MM.dd");
                    }
                    else
                    {
                        var dp = new date_picker_1.DatePicker();
                        dp.ShowDialog();

                        export_date = dp.userDate;
                    }
                    ////Use this is the date question is not used-----------------------------
                    //var dp = new date_picker_1.DatePicker();
                    //dp.ShowDialog();

                    //string export_date = dp.userDate;
                    ////--------------------------------------------------------------------

                    ////list of clash test the were and were not successfully written to database
                    List<string> passList = new List<string>();
                    List<string> failList = new List<string>();
                    List<string> emptyTestList = new List<string>();

                    ////good to proceed indicator collecting and writing clashes
                    string proceed = "yes";
                    ////test database connection
                    string dbTestResult = fill_in_db.FillDB.SQL_test();

                    if (proceed == "yes")
                    {
                        //Store clash data in data table--------------------------------------
                        foreach (ClashTest test in allTests_copy.Tests)
                        {
                            //This location will produce one xml per test.----------------------------------------
                            //Data set to store data tables
                            //DataSet ds = new DataSet();

                            //test id an name seperated-----------------------------------------------------------
                            string testFullName = test.DisplayName.ToString();
                            string testId = testFullName.Substring(0, 5);
                            //MessageBox.Show(testId);
                            string testName = testFullName.Substring(6);
                            //MessageBox.Show(testName);

                            //----------------------------------------------------------------------------------

                            if (test.LastRun != null && proceed == "yes")
                            {
                                //Create data table to stor clash data--------------------------------------------
                                //DataTable dt_test_data = new DataTable("clash_results"); //old entry: test.DisplayName.ToString()
                                DataTable dt_test_data = new DataTable(test.DisplayName.ToString()); //old entry: test.DisplayName.ToString()
                                dt_test_data.Columns.Add(new DataColumn("key_id", typeof(System.String)));
                                dt_test_data.Columns.Add(new DataColumn("test_id", typeof(System.String)));
                                //dt_test_data.Columns.Add(new DataColumn("test_name", typeof(System.String)));
                                dt_test_data.Columns.Add(new DataColumn("clash_guid", typeof(System.String)));
                                dt_test_data.Columns.Add(new DataColumn("clash_id", typeof(System.String)));
                                dt_test_data.Columns.Add(new DataColumn("date_created", typeof(System.String)));
                                dt_test_data.Columns.Add(new DataColumn("group_name", typeof(System.String)));
                                dt_test_data.Columns.Add(new DataColumn("status", typeof(System.String)));
                                dt_test_data.Columns.Add(new DataColumn("element_1_guid", typeof(System.String)));
                                dt_test_data.Columns.Add(new DataColumn("element_2_guid", typeof(System.String)));
                                dt_test_data.Columns.Add(new DataColumn("export_date", typeof(System.String)));
                                dt_test_data.Columns.Add(new DataColumn("assigned_to", typeof(System.String)));
                                dt_test_data.Columns.Add(new DataColumn("project_num", typeof(System.String)));


                                foreach (SavedItem issue in test.Children)
                                {
                                    //Sort individual clashes------------------------------------------------------------
                                    if (issue.IsGroup == false)
                                    {
                                        ClashResult clashResult = issue as ClashResult;

                                        DateTime cTime = (DateTime)clashResult.CreatedTime;
                                        String newTime = cTime.ToString("yyyy.MM.dd");

                                        if (clashResult.Item1 != null && clashResult.Item2 != null)
                                        {
                                            dt_test_data.Rows.Add(clashResult.Guid + export_date, testId, clashResult.Guid, clashResult.DisplayName, newTime, "not_grouped", clashResult.Status, clashResult.Item1.InstanceGuid, clashResult.Item2.InstanceGuid, export_date, clashResult.AssignedTo, projectNum);
                                        }
                                        if (clashResult.Item1 == null && clashResult.Item2 == null)
                                        {
                                            dt_test_data.Rows.Add(clashResult.Guid + export_date, testId, clashResult.Guid, clashResult.DisplayName, newTime, "not_grouped", clashResult.Status, "no_guid", "no_guid", export_date, clashResult.AssignedTo, clashResult.Guid + export_date, projectNum);
                                        }
                                        if (clashResult.Item1 != null && clashResult.Item2 == null)
                                        {
                                            dt_test_data.Rows.Add(clashResult.Guid + export_date, testId, clashResult.Guid, clashResult.DisplayName, newTime, "not_grouped", clashResult.Status, clashResult.Item1.InstanceGuid, "no_guid", export_date, clashResult.AssignedTo, projectNum);
                                        }
                                        if (clashResult.Item1 == null && clashResult.Item2 != null)
                                        {
                                            dt_test_data.Rows.Add(clashResult.Guid + export_date, testId, clashResult.Guid, clashResult.DisplayName, newTime, "not_grouped", clashResult.Status, "no_guid", clashResult.Item2.InstanceGuid, export_date, clashResult.AssignedTo, projectNum);
                                        }

                                        //dt_test_data.Rows.Add(testId, testName, clashResult.Guid, clashResult.DisplayName, newTime, "not_grouped", clashResult.Status, clashResult.Item1.InstanceGuid, clashResult.Item2.InstanceGuid, export_date);
                                    }

                                    //find a property
                                    //clashResult.Items.PropertyCategories.FindPropertyByDisplayName("Item", "GUID") 


                                    //sort clashes in groups-----------------------------------------------------------
                                    if (issue.IsGroup)
                                    {
                                        var group_name = issue.DisplayName;

                                        //skip clashes places in a group called "Approved"
                                        if (group_name.ToLower() != "approved")
                                        {
                                            foreach (SavedItem groupedClashes in ((GroupItem)issue).Children)
                                            {
                                                ClashResult gclashResult = groupedClashes as ClashResult;
                                                //var testing_param = groupedClashes.DisplayName;
                                                //MessageBox.Show(testing_param);

                                                DateTime cTime = (DateTime)gclashResult.CreatedTime;
                                                String newTime = cTime.ToString("yyyy.MM.dd");

                                                if (gclashResult.Item1 != null && gclashResult.Item2 != null)
                                                {
                                                    dt_test_data.Rows.Add(gclashResult.Guid + export_date, testId, gclashResult.Guid, gclashResult.DisplayName, newTime, group_name, gclashResult.Status, gclashResult.Item1.InstanceGuid, gclashResult.Item2.InstanceGuid, export_date, gclashResult.AssignedTo, projectNum);
                                                }
                                                if (gclashResult.Item1 == null && gclashResult.Item2 == null)
                                                {
                                                    dt_test_data.Rows.Add(gclashResult.Guid + export_date, testId, gclashResult.Guid, gclashResult.DisplayName, newTime, group_name, gclashResult.Status, "no_guid", "no_guid", export_date, gclashResult.AssignedTo, projectNum);
                                                }
                                                if (gclashResult.Item1 != null && gclashResult.Item2 == null)
                                                {
                                                    dt_test_data.Rows.Add(gclashResult.Guid + export_date, testId, gclashResult.Guid, gclashResult.DisplayName, newTime, group_name, gclashResult.Status, gclashResult.Item1.InstanceGuid, "no_guid", export_date, gclashResult.AssignedTo, projectNum);
                                                }
                                                if (gclashResult.Item1 == null && gclashResult.Item2 != null)
                                                {
                                                    dt_test_data.Rows.Add(gclashResult.Guid + export_date, testId, gclashResult.Guid, gclashResult.DisplayName, newTime, group_name, gclashResult.Status, "no_guid", gclashResult.Item2.InstanceGuid, export_date, gclashResult.AssignedTo, projectNum);
                                                }

                                                //dt_test_data.Rows.Add(testId, testName, gclashResult.Guid, gclashResult.DisplayName, newTime, group_name, gclashResult.Status, gclashResult.Item1.InstanceGuid, gclashResult.Item2.InstanceGuid, export_date);
                                            }
                                        }
                                    }
                                }
                                ////Write Clash date to XML file. Used with old Python Scipt-------------------------------------------------------
                                //dt_test_data.WriteXml(sSelectedPath + @"\" + testFullName + "_" + export_date + ".xml", XmlWriteMode.WriteSchema);
                                ////--------------------------------------------------------------------------------------------------------------

                                ////used to display test data tables in a date set
                                //ds.Tables.Add(dt_test_data);

                                //using (var f = new display_xml.dataForm())
                                //{
                                //    f.dataGridView1.DataSource = dt_test_data;
                                //    f.ShowDialog();
                                //}

                                //Fillin database one test at a time----------------------------------------------------------------------------

                                if (dbTestResult == "Success")
                                {
                                    string dbResult = fill_in_db.FillDB.Main_C(dt_test_data);
                                    if (dbResult.Substring(0, 1) == "P")
                                    {
                                        //MessageBox.Show(dbResult[0].ToString()+ dbResult[1].ToString()+dbResult[2].ToString());
                                        passList.Add(dbResult.Substring(2, 5));
                                    }
                                    if (dbResult.Substring(0, 1) == "F")
                                    {
                                        //MessageBox.Show(dbResult[0].ToString() + dbResult[1].ToString() + dbResult[2].ToString());
                                        failList.Add(dbResult.Substring(2, 5));
                                    }
                                    if (dbResult.Substring(0, 1) == "E")
                                    {
                                        //MessageBox.Show(dbResult[0].ToString() + dbResult[1].ToString() + dbResult[2].ToString());
                                        emptyTestList.Add(dbResult.Substring(2, 5));
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Unable to write data to database. Check network/VPN", "Database Error");
                                    proceed = "no";
                                }

                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Process was terminated.", "Error");
                    }

                    ////Show list of tests passed and failed------------------------------------------------------------------------------------------
                    string delim = ", ";
                    //string messageP = String.Join(delim, passList);
                    //string messageF = String.Join(delim, failList);
                    if (passList.Count > 0)
                    {
                        MessageBox.Show(string.Join(delim, passList), "Tests Succesfully added to Database");
                    }
                    if (failList.Count > 0)
                    {
                        MessageBox.Show(string.Join(delim, failList), "Tests Failed to add to Database");
                    }
                    if (emptyTestList.Count > 0)
                    {
                        MessageBox.Show(string.Join(delim, emptyTestList), "Tests that had no data");
                    }
                }
            }
            else
            {
                MessageBox.Show("A valid project number was not found in the file name");
            }

            ////Display window asking if user wants to write data to database----------------------------
            //var d = new database_question.DBQuestion();
            //d.ShowDialog();

            return 0;
        }
    }
}
