using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Spliter {
    public partial class Form1 : Form {
        private string filepath;
        private string[] lines = new string[500];
        public Form1() {
            InitializeComponent();
            progressBar1.Minimum = 1;
            progressBar1.Value = 1;
            progressBar1.Step = 1;
        }


        private void openToolStripMenuItem_Click(object sender, EventArgs e) {

            if (openFileDialog1.ShowDialog() == DialogResult.OK) {
                filepath = (openFileDialog1.FileName);
                richTextBox1.Text = "File loaded";
                button1.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e) {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK) {
                string name = saveFileDialog1.FileName;
                var excelApp = new Excel.Application();
                
                

                excelApp.Workbooks.Open(filepath);
                

                Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
                

                int row = 1;
                int copyrow = 0;
                bool isDone = false;
                int filecount = 0;

                progressBar1.Maximum = workSheet.UsedRange.Rows.Count;

                while (isDone == false) {
                    var excelAppCopy = new Excel.Application();
                    excelAppCopy.Workbooks.Add();
                    Excel._Worksheet workSheetcopy = (Excel.Worksheet)excelAppCopy.ActiveSheet;
                    for (int i = 0; i < 500;) {
                        if (workSheet.Cells[row, "D"] == null) workSheet.Cells[row, "4"] = "no";
                        if (workSheet.Cells[row, "D"].value == "yes") {
                            copyrow++;
                            i++;
                            if (workSheet.Cells[row + 1, "A"].value == null) {
                                i = 500;
                                isDone = true;
                            }
                            workSheetcopy.Cells[copyrow, "A"] = workSheet.Cells[row, "A"];
                            workSheetcopy.Cells[copyrow, "B"] = workSheet.Cells[row, "B"];
                            workSheetcopy.Cells[copyrow, "C"] = workSheet.Cells[row, "C"];
                        }
                        row++;
                        textBox1.Text = string.Format("{0},{1}",row,workSheet.UsedRange.Rows.Count);
                        progressBar1.PerformStep();

                    }
                    copyrow = 0;
                    filecount++;
                    if (filecount >= 11) {
                       name = name.Remove(name.Length - 2,2);
                    }else if (filecount > 1){
                        name = name.Remove(name.Length - 1,1);
                    }
                    name = string.Format(name + filecount);
                    workSheetcopy.SaveAs(name);
                    excelAppCopy.Quit();
                }
                excelApp.Quit();
            }
            





        }

       
    }
}
