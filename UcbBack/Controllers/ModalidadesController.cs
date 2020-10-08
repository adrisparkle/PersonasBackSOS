using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using UcbBack.Models;

namespace UcbBack.Controllers
{
    public class ModalidadesController : ApiController
    {
        private ApplicationDbContext _context;

        public ModalidadesController()
        {
            _context = new ApplicationDbContext();
        }

        public IHttpActionResult Get()
        {
            //var modalidades = _context.Modalidades.OrderBy(Name).ToList();
            var modalidades = _context.Modalidades.Select(x =>
                new
                {
                    x.Id,
                    x.Abr,
                    x.Modalidad
                }).OrderBy(x => x.Abr);
            return Ok(modalidades);
        }
    }
}