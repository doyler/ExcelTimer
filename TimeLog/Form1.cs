using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace TimeLog
{
    public partial class Form1 : Form
    {
        private int timeCounter = 0;
        private DateTime startTime = new DateTime();
        private DateTime endTime = new DateTime();
        int startRow = 1;

        string path = Directory.GetCurrentDirectory();

        public Form1()
        {
            InitializeComponent();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            btnEdit.Enabled = false;
            btnLock.Enabled = false;
            btnStart.Enabled = false;
            btnStop.Enabled = true;
            btnSave.Enabled = false;

            myTimer.Start();
            startTime = DateTime.Now;
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            btnEdit.Enabled = true;
            btnLock.Enabled = false;
            btnStart.Enabled = false; //TODO: Add stopping/startting/resetting
            btnStop.Enabled = false;
            btnSave.Enabled = true;

            myTimer.Stop();
            endTime = DateTime.Now;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            btnEdit.Enabled = false;
            btnLock.Enabled = false;
            btnStart.Enabled = true; //TODO: Reset after a save?
            btnStop.Enabled = true;
            btnSave.Enabled = false;

            try
            {
                var existingFile = new FileInfo("TimeLog.xlsx");

                using (var package = new ExcelPackage(existingFile))
                {
                    ExcelWorkbook workBook = package.Workbook;
                    if (workBook != null)
                    {
                        if (workBook.Worksheets.Count > 0)
                        {
                            ExcelWorksheet currentWorksheet = workBook.Worksheets.First();

                            object col1Header = currentWorksheet.Cells[startRow, 1].Value;
                            object col2Header = currentWorksheet.Cells[startRow, 2].Value;

                            if (((col1Header != null) && (col1Header.ToString() == "Start Time")) && ((col2Header != null) && (col2Header.ToString() == "End Time")))
                            {
                                int endRow = currentWorksheet.Dimension.End.Row + 1;

                                for (int rowNumber = startRow + 1; rowNumber <= endRow; rowNumber++)
                                {
                                    object col1Value = currentWorksheet.Cells[rowNumber, 1].Value;
                                    object col2Value = currentWorksheet.Cells[rowNumber, 2].Value;

                                    if ((col1Value == null) && (col2Value == null))
                                    {
                                        currentWorksheet.Cells[rowNumber, 1].Value = startTime.ToString();
                                        currentWorksheet.Cells[rowNumber, 2].Value = endTime.ToString();
                                        currentWorksheet.Cells[rowNumber, 3].Value = 0.1 * Math.Ceiling(10 * (timeCounter / (60.0 * 60.0)));
                                        try
                                        {
                                            package.Save();
                                        }
                                        catch (Exception ex)
                                        {
                                            showError(ex.ToString());
                                            showError(ex.Message);
                                        }
                                        finally
                                        {
                                            showError("Save complete!");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("Example data incorrectly formatted.");
                            }
                        }
                    }
                    workBook = null;
                }
            }
            catch (Exception ex)
            {
                showError(ex.ToString());
                showError(ex.Message);
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            showError("Note that if you edit the time, then it will not match up with the Start Time and End Time");

            btnEdit.Enabled = false;
            btnLock.Enabled = true;
            btnStart.Enabled = false;
            btnStop.Enabled = false;
            btnSave.Enabled = false;

            theTime.Enabled = true;
        }

        private void btnLock_Click(object sender, EventArgs e)
        {
            btnEdit.Enabled = true;
            btnLock.Enabled = false;
            btnStart.Enabled = false;
            btnStop.Enabled = false;
            btnSave.Enabled = true;

            theTime.Enabled = false;
            timeCounter = (int)TimeSpan.Parse(theTime.Text).TotalSeconds;
        }

        private void showError(string theError)
        {
            MessageBox.Show(theError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
        }

        private void myTimer_Tick(object sender, EventArgs e)
        {
            timeCounter++;
            theTime.Text = (timeCounter / (60 * 60)).ToString().PadLeft(2, '0') + ":" + ((timeCounter / 60) % 60).ToString().PadLeft(2, '0') + ":" + (timeCounter % 60).ToString().PadLeft(2, '0');
        }
    }
}
