using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using UcbBack.Models;

namespace UcbBack.Controllers
{
    public class TipoTareaController : ApiController
    {
        private ApplicationDbContext _context;

        public TipoTareaController()
        {
            _context = new ApplicationDbContext();
        }

        public IHttpActionResult Get()
        {
            var tipoTarea = _context.TipoTarea.ToList();
            return Ok(tipoTarea);
        }
    }
}