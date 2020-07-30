using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;

namespace UcbBack.Models
{
    [CustomSchema("TipoTarea")]
    public class TipoTarea
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { get; set; }
        public string Tarea { get; set; }
        public string Abr { get; set; }
    }
}