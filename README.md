 static public void ExportTimeSheetToExcel_EPPLUS(string tieude, DataTable timesheet, DataTable listEmp, DataTable listDivision, string departmentID, string divisionID, string yearmonth)
        {
            try
            {
                FileInfo Template_File = new FileInfo(System.Windows.Forms.Application.StartupPath + "\\Template\\Timesheet.xlsx");
                using (ExcelPackage Package = new ExcelPackage(Template_File))
                {
                    int SheetIndex = 1;
                    string strFilter = "1=1";
                    if (departmentID != "")
                        strFilter = strFilter + " AND DEPARTMENT_ID=" + departmentID;
                    if (divisionID != "")
                        strFilter = strFilter + " AND DIVISION_ID=" + divisionID;
                    DataRow[] division = listDivision.Select(strFilter, "DIVISION_ID asc");

                    foreach (DataRow row_div in division)
                    {
                        string strFilt = "DIVISION_ID=" + Convert.ToInt32(row_div["DIVISION_ID"]);
                        string SName = row_div["DIVISION_NAME_VI"].ToString();
                        DataRow[] rowDep = listEmp.Select(strFilt);
                        int maxIndex = 0;
                        //Check sheetname exist
                        var sheetNames = Package.Workbook.Worksheets.Select(w => w.Name).ToList();
                        string SName_temp = SName;
                        foreach (var sheetName in sheetNames)
                        {
                            if (sheetName.StartsWith(SName_temp, StringComparison.OrdinalIgnoreCase))
                            {
                                int startIndex = SName_temp.Length + 2; // Vị trí của số sau "Tên ("
                                int endIndex = sheetName.LastIndexOf(")");
                                if (endIndex <= 0)
                                    maxIndex = 1;
                                else
                                {
                                    maxIndex = Convert.ToInt32(sheetName.Substring(startIndex, endIndex - startIndex)) + 1;
                                }
                                SName = SName_temp + " (" + maxIndex.ToString() + ")";
                            }
                        }

                        ExcelWorksheet worksheet = Package.Workbook.Worksheets.Copy("Vp", SName);
                        double totalNC = 0;
                        double total_150 = 0;
                        double total_195 = 0;
                        double total_CN = 0;
                        double total_N = 0;
                        int i = 12;
                        string strFilt_Emp = "";
                        int stt = 1;
                        worksheet.Cells[1, 1].Value = "WORKINGTIME SHEET - BẢNG CHẤM CÔNG THÁNG: " + clsGlobal.Right(yearmonth, 2) + "/" + clsGlobal.Left(yearmonth, 4);
                        worksheet.Cells[4, 3].Value = SName_temp;
                        int totalEmp = rowDep.Length * 5;
                        worksheet.InsertRow(i, totalEmp);
                        foreach (DataRow dataRow in rowDep)
                        {
                            int intdate = 0;
                            worksheet.SelectedRange["A7:AR11"].Copy(worksheet.Cells["A" + (i).ToString() + ":AR" + (i + 4).ToString()]);
                            worksheet.Cells[i, 1].Value = stt;
                            worksheet.Cells[i, 1, i + 4, 1].Merge = true;
                            worksheet.Cells[i, 2].Value = dataRow["EMP_CODE"];
                            worksheet.Cells[i, 2, i + 4, 2].Merge = true;
                            worksheet.Cells[i, 3].Value = dataRow["FULLNAME"];
                            worksheet.Cells[i, 3, i + 4, 3].Merge = true;
                            worksheet.Cells[i, 4].Value = dataRow["HIRE_DAY"];
                            worksheet.Cells[i, 4, i + 4, 4].Merge = true;
                            //an dong gio hanh chanh 
                            if (frmDailyAttendent.chkTC == 1)
                                worksheet.Row(i).Hidden = true;
                            strFilt_Emp = "EMPLOYEE_ID = " + Convert.ToInt32(dataRow["EMPLOYEE_ID"]);
                            DataRow[] sheetrow = timesheet.Select(strFilt_Emp, "DIVISION_ID asc");
                            foreach (DataRow eachrow in sheetrow)
                            {
                                if (eachrow["DATEOFMONTH"] != null)
                                {
                                    intdate = Convert.ToDateTime(eachrow["DATEOFMONTH"]).Day;
                                }
                                worksheet.Cells[6, 5 + intdate].Value = eachrow["DATENAME"];
                                if (eachrow["DATENAME"].ToString() == "SUN")
                                {
                                    worksheet.Cells[i, 5 + intdate, i + 4, 5 + intdate].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    worksheet.Cells[i, 5 + intdate, i + 4, 5 + intdate].Style.Fill.BackgroundColor.SetColor(Color.Gray);

                                }
                                if (eachrow["STATUS"].ToString() == "4")
                                {
                                    worksheet.Cells[i, 5 + intdate].Value = "K";
                                }
                                else if (eachrow["STATUS"].ToString() == "5")
                                {
                                    worksheet.Cells[i, 5 + intdate].Value = "L";
                                }
                                else if (eachrow["STATUS"].ToString() == "6")
                                {
                                    double anualleave;
                                    if (eachrow["ANUAL_LEAVE"].ToString() == "1" || eachrow["ANUAL_LEAVE2"].ToString() == "1")
                                    {
                                        worksheet.Cells[i, 5 + intdate].Value = "N";
                                        total_N += 1;
                                    }
                                    else
                                    {
                                        anualleave = (Convert.ToDouble(eachrow["ANUAL_LEAVE"]) + Convert.ToDouble(eachrow["ANUAL_LEAVE2"])) * 8;
                                        worksheet.Cells[i, 5 + intdate].Value = "N/" + anualleave.ToString();
                                        total_N += anualleave / 8;
                                    }
                                }
                                else if (eachrow["STATUS"].ToString() == "7")
                                {
                                    double otherleave;
                                    if (eachrow["OTHER_LEAVE"].ToString() == "1" || eachrow["OTHER_LEAVE2"].ToString() == "1")
                                        worksheet.Cells[i, 5 + intdate].Value = "P";
                                    else
                                    {
                                        otherleave = (Convert.ToDouble(eachrow["OTHER_LEAVE"]) + Convert.ToDouble(eachrow["OTHER_LEAVE2"])) * 8;
                                        worksheet.Cells[i, 5 + intdate].Value = "P/" + otherleave.ToString();
                                    }
                                }
                                else if (eachrow["STATUS"].ToString() == "8")
                                {
                                    worksheet.Cells[i, 5 + intdate].Value = "C";
                                }
                                else if (eachrow["STATUS"].ToString() == "9")
                                {
                                    worksheet.Cells[i, 5 + intdate].Value = "KT";
                                }
                                else if (eachrow["STATUS"].ToString() == "10")
                                {
                                    worksheet.Cells[i, 5 + intdate].Value = "HL";
                                }
                                else
                                {
                                    if (Convert.ToDouble(eachrow["HOUR_WORK"]) != 0)
                                    {
                                        worksheet.Cells[i, 5 + intdate].Value = eachrow["HOUR_WORK"];
                                        totalNC += Convert.ToDouble(eachrow["HOUR_WORK"]);
                                    }
                                }
                                if (Convert.ToDouble(eachrow["OT_WORK"]) != 0)
                                {
                                    worksheet.Cells[i + 1, 5 + intdate].Value = Convert.ToDouble(eachrow["OT_WORK"]) + Convert.ToDouble(eachrow["OT_WORK_2"]);
                                    total_150 += Convert.ToDouble(eachrow["OT_WORK"]) + Convert.ToDouble(eachrow["OT_WORK_2"]);
                                }
                                if (Convert.ToDouble(eachrow["OT195_WORK"]) != 0)
                                {
                                    worksheet.Cells[i + 2, 5 + intdate].Value = eachrow["OT195_WORK"];
                                    total_195 += Convert.ToDouble(eachrow["OT195_WORK"]);
                                }
                                if (Convert.ToDouble(eachrow["OT200_WORK"]) != 0)
                                {
                                    worksheet.Cells[i + 3, 5 + intdate].Value = eachrow["OT200_WORK"];
                                    total_CN += Convert.ToDouble(eachrow["OT200_WORK"]);
                                }
                                if (Convert.ToDouble(eachrow["NIGHT_TIME"]) != 0)
                                {
                                    worksheet.Cells[i + 4, 5 + intdate].Value = eachrow["NIGHT_TIME"];
                                    //total_CN += Convert.ToDouble(eachrow["NIGHT_TIME"]);
                                }
                            }
                            i = i + 5;
                            stt++;
                        }
                        worksheet.Cells[1, 39].Value = totalNC / 8;
                        worksheet.Cells[2, 39].Value = total_150;
                        worksheet.Cells[3, 39].Value = total_195;
                        worksheet.Cells[4, 39].Value = total_CN;
                        worksheet.Cells[4, 42].Value = total_N;
                    }

                    Package.Workbook.Worksheets.Delete("Vp");
                    SaveFileDialog dlg = new SaveFileDialog();
                    dlg.Filter = "Excel files  .xlsx|*.xlsx";
                    dlg.ShowDialog();
                    string path = dlg.FileName;
                    if (path != "")
                        Package.SaveAs(new FileInfo(path));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }