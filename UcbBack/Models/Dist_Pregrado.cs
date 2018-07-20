﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace UcbBack.Models
{
    [Table("ADMNALRRHH.Dist_Pregrado")]
    public class Dist_Pregrado
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { set; get; }

        public string Document { get; set; }
        public string FirstName { get; set; }
        public string FirstSurName { get; set; }
        public string SecondSurName { get; set; }
        public string MariedSurName { get; set; }
        public Decimal TotalNeto { get; set; }
        public string Carrera { get; set; }
        public string CUNI { get; set; }
        public string Dependency { get; set; }

        public decimal Porcentaje { get; set; }
        public string segmentoOrigen { get; set; }
        [StringLength(2)]
        public string mes { get; set; }
        [StringLength(4)]
        public string gestion { get; set; }

        public Dist_File DistFile { get; set; }
        public int DistFileId { get; set; }
    }
}