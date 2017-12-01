using System;
using System.Collections.Generic;
using System.Text;

namespace DS9208YY.Models
{
    public partial class ImportBill
    {
        public string 日期 { get; set; }
        public string 单据编号 { get; set; }
        public string EntryID { get; set; }
        public string 购货单位 { get; set; }
        public string 发货仓库 { get; set; }
        public string 产品名称 { get; set; }
        public string 规格型号 { get; set; }
        public float 实发数量 { get; set; }
        public string 批号 { get; set; }
        public string 摘要 { get; set; }
        public float fAuxQty { get; set; }
    }
}
