using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;
using Aras.IOM;
using JPC_Reports.Models.BPM;
using System.IO;
using System.Diagnostics;

namespace EmployeeBPMToAras
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        private void btnTransfer_Click(object sender, EventArgs e)
        {
            BPMRepository BPMRepo = new BPMRepository(txtConnStr.Text);

            DataTable dtS = BPMRepo.GetEmployeeLevel();
            DataTable dt = new DataTable();
            dt.Columns.Add("Employee");
            dt.Columns.Add("cn_level0_manager");
            dt.Columns.Add("cn_level1_manager");
            dt.Columns.Add("cn_level2_manager");
            dt.Columns.Add("cn_level3_manager");
            dt.Columns.Add("cn_level4_manager");
            dt.Columns.Add("cn_level5_manager");
            dt.Columns.Add("cn_level6_manager");
            dt.Columns.Add("cn_level7_manager");
            dt.Columns.Add("cn_level8_manager");
            dt.Columns.Add("cn_level9_manager");
            dt.TableName = "Employee";
            int i = 0;
            foreach (DataRow row in dtS.Rows)
            {
                
                DataRow dtRow = dt.NewRow();
                string occu = row["Occu"].ToString();
                string level = row["functionLevelName"].ToString();
                string dept_name = row["DeptName"].ToString();
                string super_dept = row["SuperDeptName"].ToString();
                string manager = row["Manager"].ToString();
                string isMain = row["isMain"].ToString();
                string org_name = row["OrgName"].ToString();
                if (isMain == "0") continue;
                
                occu = occu.Replace("（", "(");
                int splitIndex = occu.IndexOf('(');
                if (splitIndex > 0)
                {
                    occu = occu.Substring(0, splitIndex);
                }
                dtRow["Employee"] = occu;

                dtRow = ReturnManager(dtS, dtRow, occu, manager, org_name, dept_name);

                dt.Rows.Add(dtRow);
            }
            XLWorkbook nWb = new XLWorkbook();
            nWb.AddWorksheet(dt);
            nWb.SaveAs("Result.xlsx");
        }
        private DataRow ReturnManager(DataTable wsheet, DataRow dtRow, string occu, string manager, string org_name, string dept_name)
        {
            if (manager == "NULL")
            {
                return dtRow;
            }
            else
            {
                DataRow[] isResultManager = wsheet.Select("Occu ='" + manager + "' and OrgName='" + org_name + "' and DeptName='" + dept_name + "'");
                if (isResultManager.Count() > 0)
                {
                    //Log("Occu ='" + manager + "' and OrgName='"+org_name+"' and DeptName='"+dept_name+"'");
                    string occu2 = isResultManager[0]["Occu"].ToString();
                    string dept_name2 = isResultManager[0]["DeptName"].ToString();
                    string super_dept2 = isResultManager[0]["SuperDeptName"].ToString();
                    string level_name = isResultManager[0]["functionLevelName"].ToString();
                    string manager_name = isResultManager[0]["Manager"].ToString();
                    manager = manager.Replace("（", "(");
                    int splitIndex = manager.IndexOf('(');
                    if (splitIndex > 0)
                    {
                        manager = manager.Substring(0, splitIndex);
                    }
                    switch (level_name)
                    {
                        case "董事長級":
                            dtRow["cn_level0_manager"] = manager;
                            break;
                        case "總經理級":
                            dtRow["cn_level1_manager"] = manager;
                            break;
                        case "BU副總經理":
                            dtRow["cn_level2_manager"] = manager;
                            break;
                        case "集團主管":
                            dtRow["cn_level3_manager"] = manager;
                            break;
                        case "處長/協理/副總級":
                            dtRow["cn_level4_manager"] = manager;
                            break;
                        case "資深經理級":
                            dtRow["cn_level5_manager"] = manager;
                            break;
                        case "經理級":
                            dtRow["cn_level6_manager"] = manager;
                            break;
                        case "副理級":
                            dtRow["cn_level7_manager"] = manager;
                            break;
                        case "課長級":
                            dtRow["cn_level8_manager"] = manager;
                            break;
                        case "組長級":
                            dtRow["cn_level9_manager"] = manager;
                            break;
                    }
                    dtRow = ReturnManager(wsheet, dtRow, occu2, manager_name, org_name, super_dept2);
                }
                else
                {
                    //找不到部門中的長官，有可能是跨部門，找回主要部門
                    DataRow[] isResultManager2 = wsheet.Select("Occu ='" + manager + "' and isMain='1'");
                    if (isResultManager2.Count() > 0)
                    {
                        string occu2 = isResultManager2[0]["Occu"].ToString();
                        string dept_name2 = isResultManager2[0]["DeptName"].ToString();
                        string super_dept2 = isResultManager2[0]["SuperDeptName"].ToString();
                        string level_name = isResultManager2[0]["functionLevelName"].ToString();
                        string manager_name = isResultManager2[0]["Manager"].ToString();
                        manager = manager.Replace("（", "(");
                        int splitIndex = manager.IndexOf('(');
                        if (splitIndex > 0)
                        {
                            manager = manager.Substring(0, splitIndex);
                        }
                        switch (level_name)
                        {
                            case "董事長級":
                                dtRow["cn_level0_manager"] = manager;
                                break;
                            case "總經理級":
                                dtRow["cn_level1_manager"] = manager;
                                break;
                            case "BU副總經理":
                                dtRow["cn_level2_manager"] = manager;
                                break;
                            case "集團主管":
                                dtRow["cn_level3_manager"] = manager;
                                break;
                            case "處長/協理/副總級":
                                dtRow["cn_level4_manager"] = manager;
                                break;
                            case "資深經理級":
                                dtRow["cn_level5_manager"] = manager;
                                break;
                            case "經理級":
                                dtRow["cn_level6_manager"] = manager;
                                break;
                            case "副理級":
                                dtRow["cn_level7_manager"] = manager;
                                break;
                            case "課長級":
                                dtRow["cn_level8_manager"] = manager;
                                break;
                            case "組長級":
                                dtRow["cn_level9_manager"] = manager;
                                break;
                        }
                        dtRow = ReturnManager(wsheet, dtRow, occu2, manager_name, org_name, super_dept2);
                    }
                }
            }
            return dtRow;
        }

        private void btnOpenResult_Click(object sender, EventArgs e)
        {
            Process.Start(@"Result.xlsx");
        }
    }
}
