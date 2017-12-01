using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Controls;
using Ray.Framework.Converter;
using Ray.Framework.DBUtility;
using System;
using System.Data;
using System.Text;

namespace DS9208YY
{
    public partial class Form1 : Office2007Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataTable dt = new DataTable();
        //string mingQRCode = "";
        //string info = "";
        string sql = "";
        string sBillNo = "";

        /// <summary>
        /// 查询
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonX1_Click(object sender, EventArgs e)
        {
            //[日期],[FInterID],[单据编号],[FEntryID],[购货单位],[发货仓库],
            //[产品名称],[规格型号],[实发数量],[批号],[摘要] ,
            //[FActQty] ,[FQRCode]
            string QRCode = tbQRCode.Text;
            string startDate = dtiStartDate.Value.Year + "-" + dtiStartDate.Value.Month + "-" + dtiStartDate.Value.Day;
            string endDate = dtiEndDate.Value.Year + "-" + dtiEndDate.Value.Month + "-" + dtiEndDate.Value.Day;
            string customName = tbCustomName.Text;
            string billNo = tbBillNo.Text;
            string productName = tbProductName.Text;

            if (startDate.StartsWith("1-") && endDate.StartsWith("1-") && QRCode == "" && customName == "" && billNo == "" && productName == "")
            {
                DesktopAlert.Show(Utils.H2("请先设置条件!"));
                return;

            }
            sql = "SELECT [日期],[FInterID] as 内部编号 ,[单据编号],[FEntryID]as 分录号,[购货单位],[发货仓库], [产品名称],[规格型号],[实发数量],[批号],[摘要] ,[FActQty] as 扫描数量 ,[FQRCode] as 条码 FROM [View_QRCode] WHERE ";

            StringBuilder sb = new StringBuilder();
            if (!startDate.StartsWith("1-"))
            {
                sb.Append(" And [日期] >= '" + startDate + "'");
            }
            if (!endDate.StartsWith("1-"))
            {
                sb.Append(" And [日期] <= '" + endDate + "'");
            }
            if (!String.IsNullOrEmpty(QRCode))
            {
                sb.Append(" And [条码] = '" + QRCode + "' ");
            }
            if (!string.IsNullOrEmpty(billNo))
            {
                sb.Append(" And [单据编号] = '" + billNo + "' ");
            }
            if (!string.IsNullOrEmpty(productName))
            {
                sb.Append(" And [产品名称] like '%" + productName + "%'");
            }
            if (!string.IsNullOrEmpty(customName))
            {
                sb.Append(" And [购货单位] LIKE '%" + customName + "%'");
            }
            sb.Append(" Order By [单据编号] ASC");

            dt = SqlHelper.ExecuteDataTable(sql + sb.ToString().Substring(4));

            //去掉重复的行
            RemoveRepeat();
            dataGridViewX1.DataSource = dt;
            dataGridViewX1.Columns["条码"].Width = 200;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            //受EXCEL限制,在每个条码前加上'
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i]["条码"] = "'" + dt.Rows[i]["条码"];
                }
            }
            //导出EXCEL
            DataTable2Excel.DataToExcel(dt);
        }


        /// <summary>
        /// 去掉重复的行
        /// </summary>
        private void RemoveRepeat()
        {
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (sBillNo == dt.Rows[i]["单据编号"].ToString())
                    {
                        //[日期],[FInterID],[单据编号],[FEntryID],[购货单位],[发货仓库],
                        //[产品名称],[规格型号],[实发数量],[批号],[摘要] ,
                        //[FActQty] ,[FQRCode]
                        dt.Rows[i]["日期"] = DBNull.Value;
                        dt.Rows[i]["内部编号"] = DBNull.Value;
                        dt.Rows[i]["单据编号"] = "";
                        dt.Rows[i]["分录号"] = DBNull.Value;
                        dt.Rows[i]["购货单位"] = "";
                        dt.Rows[i]["发货仓库"] = "";
                        dt.Rows[i]["产品名称"] = "";
                        dt.Rows[i]["规格型号"] = "";
                        dt.Rows[i]["实发数量"] = DBNull.Value;
                        dt.Rows[i]["批号"] = "";
                        dt.Rows[i]["摘要"] = "";
                        dt.Rows[i]["扫描数量"] = DBNull.Value;
                    }
                    else
                    {
                        sBillNo = dt.Rows[i]["单据编号"].ToString();
                    }
                }
            }
            else
            {
                DesktopAlert.Show(Utils.H2("没有数据！"));
            }
        }
    }
}
