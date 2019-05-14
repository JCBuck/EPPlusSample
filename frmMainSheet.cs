using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;

namespace EPPlusSample
{
    
    public partial class frmMainSheet : Form
    {
        


        public frmMainSheet()
        {
            //https://stackoverflow.com/a/6362414/1764583
            //https://stackoverflow.com/a/27280598/1764583
            // To embed a dll in a compiled exe:
            // 1 - Change the properties of the dll in References so that Copy Local=false
            // 2 - Add the dll file to the project as an additional file not just a reference
            // 3 - Change the properties of the file so that Build Action=Embedded Resource
            // 4 - Paste this code before Application.Run in the main exe
            AppDomain.CurrentDomain.AssemblyResolve += (Object sender, ResolveEventArgs args) =>
            {
                String thisExe = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
                System.Reflection.AssemblyName embeddedAssembly = new System.Reflection.AssemblyName(args.Name);
                String resourceName = thisExe + "." + embeddedAssembly.Name + ".dll";

                using (var stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    Byte[] assemblyData = new Byte[stream.Length];
                    stream.Read(assemblyData, 0, assemblyData.Length);
                    return System.Reflection.Assembly.Load(assemblyData);
                }
            };

            InitializeComponent();

        }
        ExcelPackage excelPackage;

        private void BtnNew_Click(object sender, EventArgs e)
        {
            String[,] arr = new String[5, 5];

            var columnCount = arr.GetUpperBound(1) + 1;
            var rowCount = arr.GetUpperBound(0) + 1;


            if (dgvSheet.DataSource == null)
            {
                dgvSheet.Rows.Clear();
                dgvSheet.Columns.Clear();
            }
            else
            {
                //((DataTable)dgvSheet.DataSource).Clear();
                dgvSheet.DataSource = null;
            }


            for (int i = 0; i < columnCount; i++)
            {
                dgvSheet.Columns.Add(i.ToString(), " ");
            }
            for (int i = 0; i < rowCount; i++)
            {
                dgvSheet.Rows.AddCopy(0);
                for (int k = 0; k < columnCount; k++)
                {
                    dgvSheet.Rows[i].Cells[k].Value = arr[i, k];
                }
            }
            
            if(excelPackage != null)
                excelPackage.Dispose();
            excelPackage = new ExcelPackage();
            excelPackage.Workbook.Worksheets.Add("MySheet");
            
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (excelPackage == null)
                return;
            String path = Path.Combine(Application.StartupPath, "test.xlsx");
            var ws = excelPackage.Workbook.Worksheets[1];
            for(int row = 0; row < dgvSheet.RowCount; row++)
            {
                for(int col = 0; col < dgvSheet.ColumnCount; col++)
                {
                    ws.Cells[row + 1, col + 1].Value = dgvSheet.Rows[row].Cells[col].Value;
                }
            }

            if (excelPackage.File != null)
            {
                excelPackage.Save();
            }
            else
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Open Excel Files|*.xlsx";

                DialogResult dr = saveFileDialog.ShowDialog();

                if (dr == DialogResult.OK)
                {
                    excelPackage.SaveAs(new FileInfo(saveFileDialog.FileName));

                    frmMainSheet.ActiveForm.Text = saveFileDialog.FileName;
                }
                
            }


        }

        private void Open_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Open Excel Files|*.xlsx";
            
            DialogResult dr = openFileDialog.ShowDialog();

            //openFileDialog.FileOk
            if (dr==DialogResult.OK)
            {
                excelPackage = new ExcelPackage(new FileInfo(openFileDialog.FileName));


                frmMainSheet.ActiveForm.Text = openFileDialog.FileName;
                var ws = excelPackage.Workbook.Worksheets[1];

                int iColCnt = ws.Dimension.End.Column;
                int iRowCnt = ws.Dimension.End.Row;


                if (dgvSheet.DataSource == null)
                {
                    dgvSheet.Rows.Clear();
                    dgvSheet.Columns.Clear();
                }
                else
                {
                    //((DataTable)dgvSheet.DataSource).Clear();
                    dgvSheet.DataSource = null;
                }



                //===========Method 1 Fastest=============

                DataTable dt = new DataTable();
                for (int i = 0; i < iColCnt; i++)
                    dt.Columns.Add(i.ToString(), typeof(string));

                String[] cols = new String[iColCnt];

                for (int i = 0; i < iRowCnt; i++)
                {
                    
                    
                    for (int k = 0; k < iColCnt; k++)
                        cols[k] = ws.Cells[i + 1, k + 1].GetValue<String>();
                    dt.Rows.Add(cols);

                }
                dgvSheet.DataSource = dt;
                //===========Method 2 Faster =================

                //for (int i = 0; i < iColCnt; i++)
                //    dgvSheet.Columns.Add(i.ToString(), " ");

                //https://10tec.com/articles/why-datagridview-slow.aspx
                //List<DataGridViewRow> rows = new List<DataGridViewRow>();
                //for (int i = 0; i <  iRowCnt; i++)
                //{
                //    DataGridViewRow row = new DataGridViewRow();
                //    row.CreateCells(dgvSheet);
                //    for (int k = 0; k < iColCnt; k++)
                //        row.Cells[k].Value = ws.Cells[i + 1, k + 1].Value;

                //    //row.Cells[0].Value = String.Format("Text {0}", i);
                //    //row.Cells[1].Value = i;
                //    //row.Cells[2].Value = DateTime.Now;
                //    rows.Add(row);
                //}

                //dgvSheet.Rows.AddRange(rows.ToArray());
                //========Method 3 Slow ==============
                //for (int i = 0; i < iColCnt; i++)
                //    dgvSheet.Columns.Add(i.ToString(), " ");

                //This causes it to load slower
                //for (int i = 0; i < iRowCnt; i++)
                //{
                //    dgvSheet.Rows.AddCopy(0);
                //    

                //    for (int k = 0; k < iColCnt; k++)
                //    {
                //        dgvSheet.Rows[i].Cells[k].Value = ws.Cells[i + 1, k + 1].Value;
                //    }
                //}

            }


        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            if (excelPackage != null)
            {
                excelPackage.Dispose();
                if (dgvSheet.DataSource == null)
                {
                    dgvSheet.Rows.Clear();
                    dgvSheet.Columns.Clear();
                } else
                {
                    //((DataTable)dgvSheet.DataSource).Clear();
                    dgvSheet.DataSource = null;
                }

                frmMainSheet.ActiveForm.Text = "";
            }
            

        }

        private void FrmMainSheet_Load(object sender, EventArgs e)
        {
            //https://10tec.com/articles/why-datagridview-slow.aspx
            //enables double buffering so it's much faster
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dgvSheet.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dgvSheet, true, null);
            }
        }
    }
}
