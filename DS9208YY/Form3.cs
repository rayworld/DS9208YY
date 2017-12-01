using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Controls;
using DS9208YY.Models;
using Ray.Framework.Config;
using Ray.Framework.DBUtility;
using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;



namespace DS9208YY
{
    public partial class Form3 : Office2007Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        string fName = "";
        DataTable dt = new DataTable();
        string sql = "";

        #region 事件

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form2_Load(object sender, EventArgs e)
        {
            this.styleManager1.ManagerStyle = (eStyle)Enum.Parse(typeof(eStyle), ConfigHelper.ReadValueByKey(ConfigHelper.ConfigurationFile.AppConfig, "FormStyle"));
        }

        /// <summary>
        /// 打开
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonX2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = "c:\\";//注意这里写路径时要用c:\\而不是c:\
            dialog.Filter = "Excel97-2003文本文件|*.xls|Excel 2007文件|*.xlsx|所有文件|*.*";
            dialog.RestoreDirectory = true;
            dialog.FilterIndex = 1;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                fName = dialog.FileName;
            }

            if (!string.IsNullOrEmpty(fName))
            {
                dt = ReadExcelFile(fName, "Maotai");
                //dt.Rows.RemoveAt(dt.Rows.Count - 1);
                dataGridViewX1.DataSource = dt;
                dataGridViewX1.Columns["fdate"].HeaderText = "日期";
                dataGridViewX1.Columns["fbillNo"].HeaderText = "单据编号";
                dataGridViewX1.Columns["fEntryID"].HeaderText = "分录号";
                dataGridViewX1.Columns["FSupplyIDName"].HeaderText = "购货单位";
                dataGridViewX1.Columns["AAAAA"].HeaderText = "发货仓库";
                dataGridViewX1.Columns["FItemName"].HeaderText = "产品名称";
                dataGridViewX1.Columns["FCUUnitQty"].HeaderText = "实发数量";
                dataGridViewX1.Columns["fBatchNo"].HeaderText = "批号";
                dataGridViewX1.Columns["BBBBB"].HeaderText = "摘要";
                dataGridViewX1.Columns["FSupplyIDName"].Width = 240;
                dataGridViewX1.Columns["FItemName"].Width = 300;
            }
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {
                int recCount = 0;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ImportBill bill = new ImportBill();
                    ///对应关系修改
                    bill.日期 = dt.Rows[i]["fdate"].ToString();
                    bill.单据编号 = dt.Rows[i]["fbillNo"].ToString();
                    bill.EntryID = dt.Rows[i]["fEntryID"].ToString();
                    bill.购货单位 = dt.Rows[i]["FSupplyIDName"].ToString();
                    bill.发货仓库 = "";
                    bill.产品名称 = dt.Rows[i]["FItemName"].ToString();
                    bill.规格型号 = "";
                    bill.实发数量 = GetQty(dt.Rows[i]["FCUUnitQty"].ToString());
                    bill.批号 = dt.Rows[i]["fBatchNo"].ToString();
                    bill.摘要 = "";
                    bill.fAuxQty = 0;

                    //去重复
                    sql = string.Format("SELECT COUNT(*) FROM [icstock] WHERE [单据编号] = '{0}' AND fEntryID = {1}", bill.单据编号, bill.EntryID);
                    object obj = SqlHelper.ExecuteScalar(sql);
                    if (obj != null && int.Parse(obj.ToString()) < 1)
                    {
                        sql = string.Format("INSERT INTO [icstock] ([日期],[单据编号],[FEntryID],[购货单位],[发货仓库] ,[产品名称] ,[规格型号] ,[实发数量] ,[批号] ,[摘要] ,[FActQty]) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}',{7},'{8}','{9}', {10})", bill.日期, bill.单据编号, bill.EntryID, bill.购货单位, bill.发货仓库, bill.产品名称, bill.规格型号, bill.实发数量, bill.批号, bill.摘要, bill.fAuxQty);
                        if (SqlHelper.ExecuteNonQuery(sql) > 0)
                        {
                            recCount++;
                        }
                    }
                    else
                    {
                        DesktopAlert.Show("数据无效或重复！");
                    }
                }
                DesktopAlert.Show("<h2>" + "共成功导入 " + recCount.ToString() + " 条记录！" + "</h2>");
            }

        }
        #endregion

        #region 私有过程

        /// <summary>
        /// 将Excel文件转成DataTable
        /// </summary>
        /// <param name="strFileName">文件名</param>
        /// <param name="strSheetName">工作簿名</param>
        /// <returns></returns>
        private DataTable ReadExcelFile(string strFileName, string strSheetName)
        {
            if (strFileName != "")
            {
                ////office 2003 
                ////string conn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFileName + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1'";
                ////office 2007
                ////"Provider=Microsoft.ACE.OLEDB.12.0; Persist Security Info=False;Data Source=" + 文件选择的路径 + "; Extended Properties=Excel 8.0";
                //string conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1'";  

                string strConn = strFileName.EndsWith(".xlsx") ? @"Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + strFileName + ";Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'" : @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFileName + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1'";

                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                OleDbDataAdapter myCommand = null;
                DataTable dt = null;
                sql = "SELECT * FROM [Maotai$] ORDER BY fentryID";
                myCommand = new OleDbDataAdapter(sql, strConn);
                dt = new DataTable();
                try
                {
                    myCommand.Fill(dt);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                return dt;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 对含有小数的数量进行处理，大于0 就+1；小于0 就-1
        /// </summary>
        /// <param name="sFQty"></param>
        /// <returns></returns>
        private float GetQty(string sFQty)
        {
            float retVal = 0;

            retVal = float.Parse(sFQty);
            if (retVal - (int)retVal != 0)//是小数
            {
                if (retVal > 0)//是正数
                {
                    retVal = (int)retVal + 1;
                }
                else
                {
                    retVal = (int)retVal - 1;
                }
            }

            return retVal;
        }
        #endregion
    }
}
