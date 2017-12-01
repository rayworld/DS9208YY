using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Controls;
using Ray.Framework.DBUtility;
using System;
using System.Data;
using System.Windows.Forms;


namespace DS9208YY
{
    public partial class Form4 : Office2007Form
    {
        public Form4()
        {
            InitializeComponent();
        }

        string mingQRCodes = "";
        DataTable dt = (DataTable)null;
        string sql = "";
        string info = "";
        int lvRowNo = 0;
        string billNo = "";
        string billType = "";
        string entryID = "";

        int currVal = 0;

        #region 事件

        /// <summary>
        /// 用户重新选择了分录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridViewX1_SelectionChanged(object sender, EventArgs e)
        {
            //重置而维码列表
            mingQRCodes = "";
            //重置listview
            listViewEx1.Items.Clear();
            lvRowNo = 0;
            //准备输入
            textBoxX6.Focus();

        }

        private void expandableSplitter1_ExpandedChanged(object sender, ExpandedChangeEventArgs e)
        {
            panelEx2.Width = expandableSplitter1.Expanded == true ? 360 : 0;
            dataGridViewX1.Width = this.Width - panelEx2.Width;
        }

        private void textBoxX6_KeyDown(object sender, KeyEventArgs e)
        {

            //用户按下回车键
            if (e.KeyCode == Keys.Enter)
            {
                if (dataGridViewX1.Rows.Count == 0)
                {
                    DesktopAlert.Show(Utils.H2("请先输入单号并选择分录！"));
                    return;
                }

                //显示状态信息
                billType = comboBoxEx2.SelectedIndex == 0 ? "XSD" : "RKD";
                billNo = billType + textBoxX1.Text;
                entryID = dataGridViewX1.SelectedRows[0].Cells[1].Value.ToString();

                //单据编号和分录编号不为空
                if (billNo == "" || entryID == "")
                {
                    DesktopAlert.Show(Utils.H2("请先输入出库单编号，选择明细分录！"));
                    return;
                }

                //如果已经扫描二维码个数小于该分录总数，则继续扫描，
                int maxVal = int.Parse(dataGridViewX1.SelectedRows[0].Cells[4].Value.ToString());
                currVal = int.Parse(dataGridViewX1.SelectedRows[0].Cells[5].Value.ToString());

                if (currVal < maxVal)
                {
                    //去掉回车换行符
                    string QRCode = textBoxX6.Text.Trim().Replace(" ", "").Replace("\n", "").Replace("\r\n", "");
                    //揭秘成明码
                    string mingQRCode = QRCode;
                    //显示明码
                    textBoxX6.Text = mingQRCode;

                    //扫描窗口重新获得焦点
                    textBoxX6.Text = "";
                    //labelItem2.Text = "";
                    textBoxX6.Focus();

                    //限定二维码信息
                    if (string.IsNullOrEmpty(mingQRCode))
                    {
                        DesktopAlert.Show(Utils.H2("二维码为空！"));
                        return;
                    }

                    if (mingQRCode.Length != 10 && mingQRCode.Length != 24 && mingQRCode.Length != 15)
                    {
                        DesktopAlert.Show(Utils.H2("二维码长度不正确！"));
                        return;
                    }

                    if (IsNumber(mingQRCode) == false)
                    {
                        DesktopAlert.Show(Utils.H2("二维码未能正确识别！"));
                        return;
                    }


                    //查重
                    int index = mingQRCodes.IndexOf(mingQRCode);
                    if (index > -1)
                    {
                        DesktopAlert.Show(Utils.H2("此二维码录入重复！"));
                        return;
                    }
                    mingQRCodes += mingQRCode + ";";

                    //写入T_QRCode
                    //billNo = billNo.Substring(0, 1) + billNo.Substring(4);
                    insertQRCode2T_QRCode(mingQRCode, billNo, entryID);
                    //更新icstock
                    updateICStockByActQty(billNo, entryID);
                    
                    //更新Listview
                    ListViewItem lvi = new ListViewItem();
                    lvi.Text = (lvRowNo + 1).ToString();
                    lvi.SubItems.Add(mingQRCode);
                    this.listViewEx1.Items.Add(lvi);
                    lvRowNo++;



                    //更新状态栏
                    currVal++;
                    dataGridViewX1.SelectedRows[0].Cells[5].Value = currVal;

                    if (currVal == maxVal)//此分录已经完成
                    {
                        dataGridViewX1.Rows.Remove(dataGridViewX1.SelectedRows[0]);
                        //此出库单已经全部录入完成
                        if (dataGridViewX1.Rows.Count == 0)
                        {
                            DesktopAlert.Show(Utils.H2("此出库单已经全部录入完成！"));
                        }
                        else//此分录已经全部录入完成
                        {
                            dataGridViewX1.Rows[0].Selected = true;
                            DesktopAlert.Show(Utils.H2("此分录已经全部录入完成！"));
                        }
                        //清空二维码录入记录
                        mingQRCodes = "";
                    }
                }
                else
                {
                    DesktopAlert.Show(Utils.H2("二维码数量超过范围！"));
                    return;
                }
            }
        }

        /// <summary>                                                                           
        /// 用户输入新的出库单号并确认
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBoxX1_KeyDown(object sender, KeyEventArgs e)
        {
            ////用户按下回车键
            if (e.KeyCode == Keys.Enter)
            {
                //清空选项，
                dataGridViewX1.DataSource = (DataTable)null;
                dataGridViewX1.Rows.Clear();
                dataGridViewX1.Columns.Clear();
                listViewEx1.Items.Clear();
                lvRowNo = 0;
                //单据编号为数字
                if (!string.IsNullOrEmpty(textBoxX1.Text) && IsNumber(textBoxX1.Text))
                {
                    //清空二维码列表，
                    mingQRCodes = "";
                    //得到单据编号
                    billType = comboBoxEx2.SelectedIndex == 0 ? "XSD" : "RKD";
                    billNo = billType + textBoxX1.Text;
                    //收到单据分录信息
                    sql = string.Format("SELECT COUNT(*) FROM icstock WHERE [单据编号] ='{0}' AND [FActQty] < [实发数量]", billNo);
                    object obj = SqlHelper.ExecuteScalar(sql);
                    int recCount = obj != null ? int.Parse(obj.ToString()) : 0;
                    if (recCount > 0)
                    {
                        sql = string.Format("SELECT TOP 1 [日期],[购货单位],[发货仓库],[摘要] FROM icstock WHERE [单据编号] ='{0}' AND [FActQty] < [实发数量]", billNo);
                        DataTable dtmaster = SqlHelper.ExecuteDataTable(sql);
                        textBoxX2.Text = dtmaster.Rows[0][0].ToString();
                        textBoxX3.Text = dtmaster.Rows[0][1].ToString();
                        textBoxX4.Text = dtmaster.Rows[0][2].ToString();

                        sql = string.Format("SELECT [fEntryID] as 分录号,[产品名称],[批号],[实发数量] as 应发,[FActQty] as 实发  FROM icstock WHERE [单据编号] ='{0}' AND [FActQty] < [实发数量] ORDER BY fEntryID", billNo);
                        dt = SqlHelper.ExecuteDataTable(sql);
                        dataGridViewX1.DataSource = dt;
                        DataGridViewCheckBoxColumn newColumn = new DataGridViewCheckBoxColumn();
                        newColumn.HeaderText = "选择";
                        dataGridViewX1.Columns.Insert(0, newColumn);
                        dataGridViewX1.Columns["产品名称"].Width = 400;
                        dataGridViewX1.Rows[0].Selected = true;
                        //
                        textBoxX6.Focus();
                    }
                    else
                    {
                        DesktopAlert.Show(Utils.H2("无数据，请检查单据编号的输入!"));
                    }
                }
                else
                {
                    DesktopAlert.Show(Utils.H2("请检查单据编号的输入!"));
                }
            }
        }


        /// <summary>
        /// 程序启动时运行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form4_Load(object sender, EventArgs e)
        {
            comboBoxEx2.SelectedIndex = 0;
            //textBoxItem1.TextBoxWidth = 200;
            expandableSplitter1.Left = dataGridViewX1.Width;
            expandableSplitter1.Expanded = true;
        }
        #endregion

        #region 私有过程

        /// <summary>  
        /// 判读字符串是否为数值型
        /// </summary>  
        /// <param name="strNumber">字符串</param>  
        /// <returns>是否</returns>  
        public static bool IsNumber(string strNumber)
        {
            System.Text.RegularExpressions.Regex r = new System.Text.RegularExpressions.Regex(@"^-?\d+\.?\d*$");
            return r.IsMatch(strNumber);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="mingQRCode"></param>
        /// <param name="billNo"></param>
        /// <param name="EntryID"></param>
        /// <returns></returns>
        public int insertQRCode2T_QRCode(string mingQRCode, string billNo, string EntryID)
        {
            string tableName = "t_QRCode";
            string EntryNo = billNo + EntryID.PadLeft(4, '0');
            sql = string.Format("INSERT INTO [{0}] ([FQRCode],[FEntryID]) VALUES('{1}','{2}')", tableName, mingQRCode, EntryNo);
            return SqlHelper.ExecuteNonQuery(sql);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="mingQRCode"></param>
        /// <param name="billNo"></param>
        /// <param name="EntryID"></param>
        /// <returns></returns>
        public int deleteQRCode2T_QRCode(string mingQRCode, string billNo, string EntryID)
        {
            string tableName = "t_QRCode";
            string EntryNo = billNo + EntryID.PadLeft(4, '0');
            sql = string.Format("DELETE [{0}] WHERE [FQRCode] = '{1}' AND [FEntryID] = '{2}' ", tableName, mingQRCode, EntryNo);
            return SqlHelper.ExecuteNonQuery(sql);
        }

        /// <summary>
        /// 加总表
        /// </summary>
        /// <param name="billNo"></param>
        /// <param name="EntryID"></param>
        /// <returns></returnsT
        public int updateICStockByActQty(string billNo, string EntryID)
        {
            sql = string.Format("UPDATE [icstock] SET [FActQty] = [FActQty] + 1 WHERE [单据编号] = '{0}' AND [FEntryID] ={1}", billNo, EntryID);
            return SqlHelper.ExecuteNonQuery(sql);
        }

        /// <summary>
        /// 减总表
        /// </summary>
        /// <param name="billNo"></param>
        /// <param name="EntryID"></param>
        /// <returns></returnsT
        public int deleteICStockByActQty(string billNo, string EntryID)
        {
            sql = string.Format("UPDATE [icstock] SET [FActQty] = [FActQty] - 1 WHERE [单据编号] = '{0}' AND [FEntryID] ={1}", billNo, EntryID);
            return SqlHelper.ExecuteNonQuery(sql);
        }

        #endregion

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(this, "真的要删除这些记录吗？", "系统警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes) 
            {
                for (int i = 0; i < listViewEx1.Items.Count; i++)
                {
                    if (listViewEx1.Items[i].Checked == true)
                    {
                        //DesktopAlert.Show(listViewEx1.Items[i].SubItems[1].Text);
                        //删QRCode
                        string delQRCode = listViewEx1.Items[i].SubItems[1].Text;
                        int ret = deleteQRCode2T_QRCode(delQRCode, billNo, entryID);
                        //删icstock
                        ret += deleteICStockByActQty(billNo,entryID);
                        //删listview
                        if (ret >= 2) //都成功
                        {
                            listViewEx1.Items[i].Remove();
                            currVal--;
                            dataGridViewX1.SelectedRows[0].Cells[5].Value = currVal;
                        }
                        else 
                        {
                            info= string.Format("{0} 条码删除失败！",delQRCode);
                            DesktopAlert.Show(info);
                        }
                    }
                }
            }
        }


        

    }
}
