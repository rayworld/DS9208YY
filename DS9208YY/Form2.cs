using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Controls;
using Ray.Framework.DBUtility;
using System;
using System.Data;
using System.IO;


namespace DS9208YY
{
    public partial class Form2 : Office2007Form
    {

        public Form2()
        {
            InitializeComponent();
        }

        DataTable dt = (DataTable)null;
        string sql = "";
        string info = "";

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonX1_Click(object sender, EventArgs e)
        {
            string startDate = this.dateTimeInput1.Value.Date.ToShortDateString();
            string endDate = this.dateTimeInput2.Value.Date.ToShortDateString();
            string billType = switchButton1.Value == false ? "XSD" : "RKD";
            sql = string.Format("SELECT * FROM [View_QRCode] WHERE ([产品名称] LIKE '%商超%' OR [产品名称] LIKE '%餐饮%') AND [单据编号] LIKE '{0}%' AND [日期] >='{1}' AND [日期] <='{2}'", billType, startDate, endDate);
            dt = SqlHelper.ExecuteDataTable(sql);
            dataGridViewX1.DataSource = dt;
            dataGridViewX1.Columns["FQRCode"].Width = 200;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonX2_Click(object sender, EventArgs e)
        {
            if (dataGridViewX1.Rows.Count > 0)
            {
                StreamWriter sw = new StreamWriter("D:\\1.txt");
                string w = "";
                for (int i = 0; i < dataGridViewX1.Rows.Count; i++)
                {
                    w += dataGridViewX1.Rows[i].Cells["FQRCode"].Value.ToString() + "\r\n";
                }
                w = w.Substring(0, w.Length - 2);
                sw.Write(w);
                sw.Close();
                info = string.Format(" D:\\1.txt 导出成功，共导出 {0} 条记录！", dataGridViewX1.Rows.Count.ToString());
                DesktopAlert.Show(Utils.H2(info));
            }
            else
            {
                DesktopAlert.Show(Utils.H2("请先查询记录！"));
            }
        }
    }
}
