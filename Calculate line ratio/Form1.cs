using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;

namespace Calculate_line_ratio
{
    public partial class Form1 : Form
    {
        DataTable dtA = new DataTable();

        public Form1()
        {
            InitializeComponent();
        }

        private void fBtnFileLoad_Click(object sender, EventArgs e)
        {
            Console.WriteLine("btn_aFileLoad_Clicked");
            OpenFileDialog dialog = new OpenFileDialog();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = string.Empty;


                filePath = dialog.FileName;
                Console.WriteLine(String.Format("{0} was imported by btn_aFileLoad_Click", filePath));

                if (filePath.EndsWith(".csv"))
                    dtA = CSVconvertToDataTable(filePath, "dtA");   // 이 함수 쓰려면 csv 파일 맨 윗줄의 cell이 하나라도 비어있으면 안됨.
                else if (filePath.EndsWith(".xlsx") || filePath.EndsWith(".xls"))
                { 
                    //dtA = Xlsx_xlsConvertToDataTable(filePath, "dtA");
                    dtA = exceldata(filePath);
                    dgv1.DataSource = dtA;

                }
            }
        }

        private DataTable CSVconvertToDataTable(string filePath, string dtType)
        {
            DataTable dt = new DataTable();
            string[] lines = System.IO.File.ReadAllLines(filePath);

            if (lines.Length > 0)
            {
                string firstLine = lines[0]; // 2

                string[] headerLabels = firstLine.Split(',');

                short columnIndex = 0;
                foreach (string headerWord in headerLabels)   // create header of Datatale
                {
                    dt.Columns.Add(new DataColumn(headerWord));
                    columnIndex++;
                }


                for (int lineNum = 1; lineNum < lines.Length; lineNum++) // fill data to Datatale // 3
                {
                    string[] dataWords = lines[lineNum].Split(',');
                    DataRow dr = dt.NewRow();
                    columnIndex = 0;
                    foreach (string headerWord in headerLabels)
                    {
                        dr[headerWord] = dataWords[columnIndex++];
                    }
                    dt.Rows.Add(dr);
                }

                if (dt.Rows.Count > 0)
                {
                    if (dtType == "dtA")
                        dgv1.DataSource = dt;
                    //else if (dtType == "dtB")
                    //    dgv_2.DataSource = dt;
                }
            }
            return dt;
        }

        private DataTable Xlsx_xlsConvertToDataTable(string filePath, string dtType)
        {

            string dirName = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileName(filePath);
            string fileExtension = Path.GetExtension(filePath);
            string pathConn = string.Empty;
            string excelsql = string.Empty;
            try
            {
                switch (fileExtension)
                {
                    case ".xls":
                        pathConn = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={filePath};" + "Extended Properties=\"Excel 8.0; HDR=Yes; IMEX=1\"";
                        excelsql = "SELECT * FROM [WorkSheet$]";
                        break;

                    case ".xlsx":
                        pathConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=YES'";
                        excelsql = @"select * from[WorkSheet$]";
                        break;

                    case ".csv":
                        pathConn = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dirName};" + "Extended Properties=\"text; HDR=Yes; IMEX=1; FMT=Delimited\"";
                        excelsql = $"SELECT * FROM [{fileName}]";
                        break;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("오류의 건:" + e.Message);
            }
            OleDbConnection conn = new OleDbConnection(pathConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(excelsql, conn);
            DataSet excelDs = new DataSet();
            _ = myDataAdapter.Fill(excelDs);
            DataTable excelTable = excelDs.Tables[0];
            if (dtType == "dtA")
                dgv1.DataSource = excelTable;
            //else if (dtType == "dtB")
            //    dgv_2.DataSource = excelTable;
            return excelTable;

        }

        public static DataTable exceldata(string filePath)
        {
            DataTable dtexcel = new DataTable();
            bool hasHeaders = false;
            string HDR = hasHeaders ? "Yes" : "No";
            string strConn;
            if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xlsx")
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
            else
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            //Looping Total Sheet of Xl File
            /*foreach (DataRow schemaRow in schemaTable.Rows)
            {
            }*/
            //Looping a first Sheet of Xl File
            DataRow schemaRow = schemaTable.Rows[0];
            string sheet = schemaRow["TABLE_NAME"].ToString();
            if (!sheet.EndsWith("_"))
            {
                string query = "SELECT  * FROM [" + "sheet1" + "]";
                OleDbDataAdapter daexcel = new OleDbDataAdapter(query, conn);
                dtexcel.Locale = CultureInfo.CurrentCulture;
                daexcel.Fill(dtexcel);
            }

            conn.Close();
            return dtexcel;

        }

        //private void btnCalLineRatio_Click(object sender, EventArgs e)
        //{
        //    MessageBox.Show(dtA.Columns.Count.ToString());

        //    MessageBox.Show(dtA.Columns[1].ToString());

        //    int dtA_ColumnCount = dtA.Columns.Count - 1;

        //    for (int x1=1 ; x1<=dtA_ColumnCount ; x1++) 
        //    {
        //        for (int x2 = 1 ; x2 <= dtA_ColumnCount ; x2++)
        //        {
        //            if (x2 != x1)  // column이 동일하면 그냥 skip
        //            {
        //                string newColumnName = dtA.Columns[x1] +"nm"+"/" + dtA.Columns[x2]+"nm";
        //                if (!dtA.Columns.Contains(newColumnName))
        //                {
        //                    dtA.Columns.Add(newColumnName);
        //                }

                       

        //                for (int rowNum = 0; rowNum <= dtA.Rows.Count - 1; rowNum++)
        //                {
        //                    double dx1Value = double.Parse(dtA.Rows[rowNum][x1].ToString());
        //                    double dx2Value = double.Parse(dtA.Rows[rowNum][x2].ToString());

        //                    dtA.Rows[rowNum][newColumnName] = dx1Value / dx2Value;
        //                }
                        

        //            }
        //        }
        //    }

        //    dgv1.DataSource = dtA;

        //}

        private void btnCalLineRatio_Click(object sender, EventArgs e)
        {
            MessageBox.Show(dtA.Columns.Count.ToString());

            MessageBox.Show(dtA.Columns[1].ToString());

            int dtA_ColumnCount = dtA.Columns.Count - 1;

            for (int x1 = 1; x1 <= dtA_ColumnCount; x1++)
            {
                for (int x2 = 1; x2 <= dtA_ColumnCount; x2++)
                {
                    if (x2 != x1)  // column이 동일하면 그냥 skip
                    {
                        string newColumnName = dtA.Columns[x1] + "nm" + "/" + dtA.Columns[x2] + "nm";
                        if (!dtA.Columns.Contains(newColumnName))
                        {
                            dtA.Columns.Add(newColumnName);
                        }



                        for (int rowNum = 1; rowNum <= dtA.Rows.Count - 1; rowNum++)
                        {
                            double dx1Value = double.Parse(dtA.Rows[rowNum][x1].ToString());
                            double dx2Value = double.Parse(dtA.Rows[rowNum][x2].ToString());

                            dtA.Rows[rowNum][newColumnName] = dx1Value / dx2Value;
                        }


                    }
                }
            }

            dgv1.DataSource = dtA;

        }
    }
}
