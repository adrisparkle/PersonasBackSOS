﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Dist_LogErrores")]
    public class Dist_LogErrores
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }

        public int UserId { get; set; }
        public int DistProcessId { get; set; }
        public string CUNI { get; set; }
        public Error Error { get; set; }
        public int ErrorId { get; set; }
        public string Archivos { get; set; }
        public string State { get; set; }
    }
}