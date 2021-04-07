using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;


namespace Calculate_line_ratio
{
    public partial class Form1 : Form
    {
        DataTable _dtA = new DataTable();
        int _preFilteredDtAColumnNum = 0;
        int _MedianSamplingNum = 5;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<int> vs = new List<int>();
            vs.Add(100);
            vs.Add(100);
            vs.Add(100);
            Console.WriteLine(vs);

            vs.RemoveAt(0);

            vs.Add(300);

            Console.WriteLine(vs);

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
                { 
                    _dtA = CSVconvertToDataTable(filePath, "dtA");   // 이 함수 쓰려면 csv 파일 맨 윗줄의 cell이 하나라도 비어있으면 안됨.
                    lblFileName.Text = filePath;
                    Console.WriteLine("file loaded");
                }
                else if (filePath.EndsWith(".xlsx") || filePath.EndsWith(".xls"))
                { 
                    //dtA = Xlsx_xlsConvertToDataTable(filePath, "dtA");
                    _dtA = exceldata(filePath);
                    lblFileName.Text = filePath;
                    dgv1.DataSource = _dtA;

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
                    { 
                        //dgv1.DataSource = dt;
                    }
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

        //    for (int x1 = 1; x1 <= dtA_ColumnCount; x1++)
        //    {
        //        for (int x2 = 1; x2 <= dtA_ColumnCount; x2++)
        //        {
        //            if (x2 != x1)  // column이 동일하면 그냥 skip
        //            {
        //                string newColumnName = dtA.Columns[x1] + "nm" + "/" + dtA.Columns[x2] + "nm";
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
            MessageBox.Show("칼럼 갯수 : " + _dtA.Columns.Count.ToString());

            MessageBox.Show("첫번째 칼럼명 :" + _dtA.Columns[1].ToString());

            int dtA_ColumnCount = _dtA.Columns.Count - 1;

            for (int x1 = 1; x1 <= dtA_ColumnCount; x1++)
            {
                for (int x2 = 1; x2 <= dtA_ColumnCount; x2++)
                {
                    if (x2 != x1)  // column이 동일하면 그냥 skip
                    {
                        string newColumnName = _dtA.Columns[x1] + "nm" + "/" + _dtA.Columns[x2] + "nm";
                        if (!_dtA.Columns.Contains(newColumnName))
                        {
                            _dtA.Columns.Add(newColumnName);
                        }



                        //for (int rowNum = 0; rowNum <= dtA.Rows.Count - 1; rowNum++)
                        //{
                            int rowNum = 0;
                            
                            double dx1Value = double.Parse(_dtA.Rows[rowNum][x1].ToString());
                            double dx2Value = double.Parse(_dtA.Rows[rowNum][x2].ToString());

                            _dtA.Rows[rowNum][newColumnName] = dx1Value / dx2Value;
                            
                        //}


                    }
                }
                Console.WriteLine("진행률 :" + x1.ToString() + "/" + (dtA_ColumnCount).ToString());
            }
            Console.WriteLine("done");
            MessageBox.Show("done");
            ExportToCSV(_dtA, "lineratio222.csv");
            //dgv1.DataSource = dtA;
            MessageBox.Show("!!!!!!!!!!");
            //dgv1.DataSource = dtA;

        }

        public static void ExportToCSV(DataTable dtDataTable, string strFilePath)
        {

            StreamWriter sw = new StreamWriter(strFilePath, false, System.Text.Encoding.Default);
            //headers   
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i].ToString().Trim());
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString().Trim();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString().Trim());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        //private void btnCalElectronTemp_Click(object sender, EventArgs e)
        //{
        //    int dtA_ColumnCount = dtA.Columns.Count - 1;

        //    for (int iX1 = 1; iX1 <= dtA_ColumnCount; iX1++)
        //    {
        //        for (int iX2 = 1; iX2 <= dtA_ColumnCount; iX2++)
        //        {
        //            double dX1Wavelgth = double.Parse(dtA.Columns[iX1].ToString()) - 2;
        //            double dX1Energy = fFindNearestWavelgthEnergy(dX1Wavelgth);

        //            if (iX2 != iX1)  // column index가 동일하면 그냥 skip
        //            {
        //                double dX2Wavelgth = double.Parse(dtA.Columns[iX2].ToString()) - 2;
        //                double dX2Energy = fFindNearestWavelgthEnergy(dX2Wavelgth);

        //                string newColumnName = "Te("+ (dX1Wavelgth+2).ToString() + "nm" + "/" + (dX2Wavelgth+2).ToString() + "nm)";
        //                if (!dtA.Columns.Contains(newColumnName))
        //                {
        //                    dtA.Columns.Add(newColumnName);
        //                }

        //                for (int rowNum = 0; rowNum <= dtA.Rows.Count - 1; rowNum++)
        //                {
        //                    double dX1Intensity = double.Parse(dtA.Rows[rowNum][iX1].ToString());
        //                    double dX2Intensity = double.Parse(dtA.Rows[rowNum][iX2].ToString());

        //                    double dTe = -(dX1Energy-dX2Energy)/Math.Log(dX1Intensity / dX2Intensity);

        //                    if (double.IsNaN(dTe) == true) dtA.Rows[rowNum][newColumnName] = 0;
        //                    else dtA.Rows[rowNum][newColumnName] = dTe;
        //                }


        //            }
        //        }
        //        Console.WriteLine("진행률 :" + iX1.ToString() + "/" + (dtA_ColumnCount).ToString());
        //    }

        //    for (int iX1 = dtA_ColumnCount+1 ; iX1 <= dtA.Columns.Count-1; iX1++)
        //    {
        //        /// column의 초기 row[0], row[1], row[2] 값이 0이면 그 칼럼은 건너뛰게 설정
        //        if (dtA.Rows[0][iX1].ToString() != "0" && dtA.Rows[1][iX1].ToString() != "0" && dtA.Rows[2][iX1].ToString() != "0")
        //        {
        //            for (int rowNum = 0; rowNum <= dtA.Rows.Count - 1; rowNum++)
        //            {
        //                List<double> tempList = new List<double>();
        //                if (rowNum <= dtA.Rows.Count - _MedianSamplingNum)
        //                {
        //                    int rowNumforMedian = rowNum;
        //                    for (; rowNumforMedian <= rowNum + _MedianSamplingNum-1; rowNumforMedian++)
        //                    {
        //                        double dX1Intensity = double.Parse(dtA.Rows[rowNumforMedian][iX1].ToString());
        //                        tempList.Add(dX1Intensity);
        //                    }
        //                }
        //                else
        //                {
        //                    int rowNumforMedian = rowNum;
        //                    for (; rowNumforMedian >= rowNum -_MedianSamplingNum+1; rowNumforMedian--)
        //                    {
        //                        double dX1Intensity = double.Parse(dtA.Rows[rowNumforMedian][iX1].ToString());
        //                        tempList.Add(dX1Intensity);
        //                    }
        //                }
        //                tempList.Sort();
        //                dtA.Rows[rowNum][iX1] = tempList[_MedianSamplingNum/2];
        //                tempList.Clear();
        //            }
        //            Console.WriteLine("Median filtering 진행률 :" + iX1.ToString() + "/" + (dtA.Columns.Count - 1).ToString());
        //        }
        //    }

        //    Console.WriteLine("done");
        //    MessageBox.Show("done");
        //    ExportToCSV(dtA, "Te2.csv");
        //    //dgv1.DataSource = dtA;
        //    MessageBox.Show("!!!!!!!!!!");
        //    //dgv1.DataSource = dtA;
        //}

        private void btnCalElectronTemp_Click(object sender, EventArgs e)
        {
            _dtA = fCalElectronTemp(_dtA);

            _dtA = fElectronTempMedianFiltering(_dtA);

            Console.WriteLine("done");
            MessageBox.Show("done");
            ExportToCSV(_dtA, "PP_Te.csv");
            //dgv1.DataSource = dtA;
            MessageBox.Show("!!!!!!!!!!");
            //dgv1.DataSource = dtA;
        }

        /// <summary>
        /// partial pressure 고려 안하고 짠거!
        /// </summary>
        /// <param name="dtA"></param>
        /// <returns></returns>
        //DataTable fCalElectronTemp(DataTable dtA)
        //{
        //    _preFilteredDtAColumnNum = dtA.Columns.Count - 1;

        //    for (int iX1 = 1; iX1 <= _preFilteredDtAColumnNum; iX1++)
        //    {
        //        for (int iX2 = 1; iX2 <= _preFilteredDtAColumnNum; iX2++)
        //        {
        //            double dX1Wavelgth = double.Parse(dtA.Columns[iX1].ToString()) - 2;
        //            double dX1Energy = fFindNearestWavelgthEnergy(dX1Wavelgth);

        //            if (iX2 != iX1)  // column index가 동일하면 그냥 skip
        //            {
        //                double dX2Wavelgth = double.Parse(dtA.Columns[iX2].ToString()) - 2;
        //                double dX2Energy = fFindNearestWavelgthEnergy(dX2Wavelgth);

        //                string newColumnName = "Te(" + (dX1Wavelgth + 2).ToString() + "nm" + "/" + (dX2Wavelgth + 2).ToString() + "nm)";
        //                if (!dtA.Columns.Contains(newColumnName))
        //                {
        //                    dtA.Columns.Add(newColumnName);
        //                }

        //                for (int rowNum = 0; rowNum <= dtA.Rows.Count - 1; rowNum++)
        //                {
        //                    double dX1Intensity = double.Parse(dtA.Rows[rowNum][iX1].ToString());
        //                    double dX2Intensity = double.Parse(dtA.Rows[rowNum][iX2].ToString());

        //                    double dTe = -(dX1Energy - dX2Energy) / Math.Log(dX1Intensity / dX2Intensity);

        //                    if (double.IsNaN(dTe) == true) dtA.Rows[rowNum][newColumnName] = 0;
        //                    else if (100000 < dTe) dtA.Rows[rowNum][newColumnName] = 65000;
        //                    else if (dTe < -100000) dtA.Rows[rowNum][newColumnName] = -65000;
        //                    else dtA.Rows[rowNum][newColumnName] = dTe;
        //                }


        //            }
        //        }
        //        Console.WriteLine("진행률 :" + iX1.ToString() + "/" + (_preFilteredDtAColumnNum).ToString());
        //    }
        //    return dtA;
        //}    

        DataTable fElectronTempMedianFiltering(DataTable dtA)
        {
            for (int iX1 = _preFilteredDtAColumnNum+1; iX1 <= dtA.Columns.Count - 1; iX1++)
            {
                /// column의 초기 row[0], row[1], row[2] 값이 0이면 그 칼럼은 건너뛰게 설정
                if (dtA.Rows[0][iX1].ToString() != "0" && dtA.Rows[1][iX1].ToString() != "0" && dtA.Rows[2][iX1].ToString() != "0")
                {
                    for (int rowNum = 0; rowNum <= dtA.Rows.Count - 1; rowNum++)
                    {
                        List<double> tempList = new List<double>();
                        if (rowNum <= dtA.Rows.Count - _MedianSamplingNum)
                        {
                            int rowNumforMedian = rowNum;
                            for (; rowNumforMedian <= rowNum + _MedianSamplingNum - 1; rowNumforMedian++)
                            {
                                double dX1Intensity = double.Parse(dtA.Rows[rowNumforMedian][iX1].ToString());
                                tempList.Add(dX1Intensity);
                            }
                        }
                        else
                        {
                            int rowNumforMedian = rowNum;
                            for (; rowNumforMedian >= rowNum - _MedianSamplingNum + 1; rowNumforMedian--)
                            {
                                double dX1Intensity = double.Parse(dtA.Rows[rowNumforMedian][iX1].ToString());
                                tempList.Add(dX1Intensity);
                            }
                        }
                        tempList.Sort();
                        dtA.Rows[rowNum][iX1] = tempList[_MedianSamplingNum / 2];
                        tempList.Clear();
                    }
                    Console.WriteLine("Median filtering 진행률 :" + iX1.ToString() + "/" + (dtA.Columns.Count - 1).ToString());
                }
            }
            return dtA;
        }
        
        double fFindNearestWavelgthEnergy(double tempWaveLgth)
        {
            double nearestWaveLgth = CLineRatio_Te.dWaveLgthIntensity.Keys.Max();

            foreach (double Key in CLineRatio_Te.dWaveLgthIntensity.Keys)
            {
                double diffOfMaxAndTestValue = nearestWaveLgth - tempWaveLgth; // dWaveLgthIntensity key값의 최대값과 testValue의 차이
                double diffOfNowKeyAndTestValue = Key - tempWaveLgth; // dWaveLgthIntensity 현재 key 값과 testValue의 차이

                if (Math.Abs(diffOfNowKeyAndTestValue) < Math.Abs(diffOfMaxAndTestValue))
                {
                    nearestWaveLgth = Key;
                }
            }
            //Console.WriteLine(nearestWaveLgth.ToString()+"    "+CLineRatio_Te.dWaveLgthIntensity[nearestWaveLgth]);
            return CLineRatio_Te.dWaveLgthIntensity[nearestWaveLgth]; 
        }


        /// <summary>
        /// O2 779.~(777.~) Partial pressure 고려해서 짠거
        /// </summary>
        /// <param name="dtA"></param>
        /// <returns></returns>
        DataTable fCalElectronTemp(DataTable dtA)
        {
            _preFilteredDtAColumnNum = dtA.Columns.Count - 1;

            for (int iX1 = 1; iX1 <= _preFilteredDtAColumnNum; iX1++)
            {
                for (int iX2 = 1; iX2 <= _preFilteredDtAColumnNum; iX2++)
                {
                    double dX1Wavelgth = double.Parse(dtA.Columns[iX1].ToString()) - 2;
                    double dX1Energy = fFindNearestWavelgthEnergy(dX1Wavelgth);

                    if (iX2 != iX1)  // column index가 동일하면 그냥 skip
                    {
                        double dX2Wavelgth = double.Parse(dtA.Columns[iX2].ToString()) - 2;
                        double dX2Energy = fFindNearestWavelgthEnergy(dX2Wavelgth);

                        string newColumnName = "Te(" + (dX1Wavelgth + 2).ToString() + "nm" + "/" + (dX2Wavelgth + 2).ToString() + "nm)";
                        if (!dtA.Columns.Contains(newColumnName))
                        {
                            dtA.Columns.Add(newColumnName);
                        }

                        for (int rowNum = 0; rowNum <= dtA.Rows.Count - 1; rowNum++)
                        {
                            double dX1Intensity = double.Parse(dtA.Rows[rowNum][iX1].ToString());
                            double dX2Intensity = double.Parse(dtA.Rows[rowNum][iX2].ToString());
                            double dTe = 0;

                            // Partial pressure 부분 추가를 위한 부분
                            if ( CLineRatio_Te.o2List.Contains(dX1Wavelgth+2) && !CLineRatio_Te.o2List.Contains(dX2Wavelgth+2) )            dTe = -(dX1Energy - dX2Energy) / Math.Log(dX1Intensity / dX2Intensity*0.072);    // O/Ar    P(Ar/O)
                            else if ( !CLineRatio_Te.o2List.Contains(dX1Wavelgth+2) && CLineRatio_Te.o2List.Contains(dX2Wavelgth+2) )       dTe = -(dX1Energy - dX2Energy) / Math.Log(dX1Intensity / dX2Intensity*13.75);    //  Ar/O   P(O/Ar)
                            else                                                                                                                              dTe = -(dX1Energy - dX2Energy) / Math.Log(dX1Intensity / dX2Intensity); 


                            // 무한대 값이나 NaN 값을 제거하기 위한 부분
                            if (double.IsNaN(dTe) == true)                                                     dtA.Rows[rowNum][newColumnName] = 0;
                            else if (100000 < dTe)                                                                dtA.Rows[rowNum][newColumnName] = 65000;
                            else if (dTe < -100000)                                                               dtA.Rows[rowNum][newColumnName] = -65000;
                            else                                                                                      dtA.Rows[rowNum][newColumnName] = dTe;
                        }


                    }
                }
                Console.WriteLine("진행률 :" + iX1.ToString() + "/" + (_preFilteredDtAColumnNum).ToString());
            }
            return dtA;
        }

    }
}
