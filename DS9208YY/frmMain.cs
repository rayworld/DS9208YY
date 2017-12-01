using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Controls;
using Ray.Framework.Config;
using Ray.Framework.Utilities;

namespace DS9208YY
{
    public partial class frmMain : Office2007Form
    {
        public frmMain()
        {
            InitializeComponent();
        }
        /// <summary>
        /// ��������ʱִ��
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmMain_Load(object sender, EventArgs e)
        {
            this.ribbonControl1.TitleText = ConfigHelper.ReadValueByKey(ConfigHelper.ConfigurationFile.AppConfig, "AppName");
            this.styleManager1.ManagerStyle = (eStyle)Enum.Parse(typeof(eStyle), ConfigHelper.ReadValueByKey(ConfigHelper.ConfigurationFile.AppConfig, "FormStyle"));
        }

        /// <summary>
        /// �ı���ʽ����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AppCommandTheme_Executed(object sender, EventArgs e)
        {
            ICommandSource source = sender as ICommandSource;
            if (source.CommandParameter is string)
            {
                eStyle style = (eStyle)Enum.Parse(typeof(eStyle), source.CommandParameter.ToString());
                // Using StyleManager change the style and color tinting
                if (StyleManager.IsMetro(style))
                {
                    // More customization is needed for Metro
                    // Capitalize App Button and tab
                    //buttonFile.Text = buttonFile.Text.ToUpper();
                    //foreach (BaseItem item in RibbonControl.Items)
                    //{
                    //    // Ribbon Control may contain items other than tabs so that needs to be taken in account
                    //    RibbonTabItem tab = item as RibbonTabItem;
                    //    if (tab != null)
                    //        tab.Text = tab.Text.ToUpper();
                    //}

                    //buttonFile.BackstageTabEnabled = true; // Use Backstage for Metro

                    ribbonControl1.RibbonStripFont = new System.Drawing.Font("Segoe UI", 9.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                    if (style == eStyle.Metro)
                        StyleManager.MetroColorGeneratorParameters = DevComponents.DotNetBar.Metro.ColorTables.MetroColorGeneratorParameters.DarkBlue;

                    // Adjust size of switch button to match Metro styling
                    //switchButtonItem1.SwitchWidth = 16;
                    //switchButtonItem1.ButtonWidth = 48;
                    //switchButtonItem1.ButtonHeight = 19;

                    // Adjust tab strip style
                    //tabStrip1.Style = eTabStripStyle.Metro;

                    StyleManager.Style = style; // BOOM
                }
                else
                {
                    // If previous style was Metro we need to update other properties as well
                    //if (StyleManager.IsMetro(StyleManager.Style))
                    //{
                    //    ribbonControl1.RibbonStripFont = null;
                    //    // Fix capitalization App Button and tab
                    //    //buttonFile.Text = ToTitleCase(buttonFile.Text);
                    //foreach (BaseItem item in RibbonControl.Items)
                    //{
                    //    // Ribbon Control may contain items other than tabs so that needs to be taken in account
                    //    RibbonTabItem tab = item as RibbonTabItem;
                    //    if (tab != null)
                    //        tab.Text = ToTitleCase(tab.Text);
                    //}
                    //    // Adjust size of switch button to match Office styling
                    //    switchButtonItem1.SwitchWidth = 28;
                    //    switchButtonItem1.ButtonWidth = 62;
                    //    switchButtonItem1.ButtonHeight = 20;
                    //}
                    // Adjust tab strip style
                    //tabStrip1.Style = eTabStripStyle.Office2007Document;
                    StyleManager.ChangeStyle(style, Color.Empty);
                    //if (style == eStyle.Office2007Black || style == eStyle.Office2007Blue || style == eStyle.Office2007Silver || style == eStyle.Office2007VistaGlass)
                    //    buttonFile.BackstageTabEnabled = false;
                    //else
                    //    buttonFile.BackstageTabEnabled = true;
                }
            }
            else if (source.CommandParameter is Color)
            {
                if (StyleManager.IsMetro(StyleManager.Style))
                    StyleManager.MetroColorGeneratorParameters = new DevComponents.DotNetBar.Metro.ColorTables.MetroColorGeneratorParameters(Color.White, (Color)source.CommandParameter);
                else
                    StyleManager.ColorTint = (Color)source.CommandParameter;
            }
            //�����û�����
            ConfigHelper.UpdateOrCreateAppSetting(ConfigHelper.ConfigurationFile.AppConfig, "FormStyle", source.CommandParameter.ToString());
        }


        #region ˽�й���


        /// <summary>
        /// ����������ʾһ�����ĵ�����ҳ��
        /// </summary>
        /// <param name="caption">�������</param>
        /// <param name="formType">��������</param>
        public void SetMdiForm(string caption, Type formType)
        {
            bool IsOpened = false;

            //�������е�Tabҳ�棬������ڣ���ô����Ϊѡ�м���
            foreach (SuperTabItem tabitem in NavTabControl.Tabs)
            {
                if (tabitem.Name == caption)
                {
                    NavTabControl.SelectedTab = tabitem;
                    IsOpened = true;
                    break;
                }
            }

            //���������Tabҳ����û���ҵ�����ô��Ҫ��ʼ����Tabҳ����
            if (!IsOpened)
            {
                //Ϊ�˷����������LoadMdiForm����������һ���µĴ��壬����ΪMDI���Ӵ���
                //Ȼ������SuperTab�ؼ�������һ��SuperTabItem����ʾ
                DevComponents.DotNetBar.Office2007Form form = ChildWinManagement.LoadMdiForm(this, formType)
                    as DevComponents.DotNetBar.Office2007Form;

                SuperTabItem tabItem = NavTabControl.CreateTab(caption);
                tabItem.Name = caption;
                tabItem.Text = caption;

                form.FormBorderStyle = FormBorderStyle.None;
                form.TopLevel = false;
                form.Visible = true;
                form.Dock = DockStyle.Fill;
                //tabItem.Icon = form.Icon;
                tabItem.AttachedControl.Controls.Add(form);

                NavTabControl.SelectedTab = tabItem;
            }
        }

        #endregion

        /// <summary>
        /// ��Ϣ��ѯ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonItem20_Click(object sender, EventArgs e)
        {
            SetMdiForm("��Ϣ��ѯ", typeof(Form1));
        }

        /// <summary>
        /// ������Ϣ����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonItem18_Click(object sender, EventArgs e)
        {
            SetMdiForm("������Ϣ����", typeof(Form3));
        }

        /// <summary>
        /// ���ϴ�����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonItem14_Click(object sender, EventArgs e)
        {
            SetMdiForm("���ϴ�����", typeof(Form2));
        }

        /// <summary>
        /// �ֹ���¼
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonItem15_Click(object sender, EventArgs e)
        {
            SetMdiForm("�ֹ���¼", typeof(Form4));
        }


    }
}