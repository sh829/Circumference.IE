using Circumference.ImportAndExport.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace Web.Models
{
    public class ImportBillsDto
    {
        /// <summary>
        /// 清单编号
        /// </summary>
        [ImporterHeader(Name = "清单编号", Description = "必填")]
        [Required(ErrorMessage = "清单编号是必填的")]
        public string BillCode { get; set; }
        /// <summary>
        /// 清单名称
        /// </summary>
        [ImporterHeader(Name = "清单名称", Description = "必填")]
        [Required(ErrorMessage = "清单名称是必填的")]
        public string BillName { get; set; }
        [ImporterHeader(Name = "清单单位")]
        public decimal Unit { get; set; }
       
        [ImporterHeader(Name = "单价")]
        public decimal Price { get; set; }
        [ImporterHeader(Name = "工程量")]
        public decimal Quantity { get; set; }
        [ImporterHeader(Name = "金额")]
        public decimal Amount { get; set; }
       
        [ImporterHeader(Name = "清单类型")]
        public string BillType { get; set; }
        [ImporterHeader(Name = "项目特征")]
        public string Feature { get; set; }
    }
}