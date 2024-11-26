using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ThankBoss.Models
{
    public class OrderModel
    {
        [Required(ErrorMessage = "Please enter input content")]
        [Display(Name = "Input Content")]
        public string InputContent { get; set; }
        public string title { get; set; }
        public string subtitle { get; set; }
        public int quantity { get; set; }
        public decimal price { get; set; }
    }
}