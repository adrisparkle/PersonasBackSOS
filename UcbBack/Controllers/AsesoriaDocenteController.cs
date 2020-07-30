using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Data;
using System.Data.Entity;
using UcbBack.Logic;
using UcbBack.Models;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;
using UcbBack.Models.Not_Mapped.ViewMoldes;
using ClosedXML.Excel;
using System.Net.Http.Headers;
using UcbBack.Models.Not_Mapped;
using System.IO;
using UcbBack.Logic.B1;
using Newtonsoft.Json.Linq;
using System.Configuration;
using UcbBack.Logic.ExcelFiles;
using UcbBack.Models.Dist;
using System.Diagnostics;
using UcbBack.Logic.ExcelFiles.Serv;

namespace UcbBack.Controllers
{
    public class AsesoriaDocenteController : ApiController
    {
        private ApplicationDbContext _context;
        private ValidateAuth auth;
        private B1Connection B1;

        public AsesoriaDocenteController()
        {
            _context = new ApplicationDbContext();
            B1 = B1Connection.Instance();
            auth = new ValidateAuth();
        }
        //convertir a mes literal
        public List<AsesoriaDocenteViewModel> mesLiteral(string query)
        {
            string[] _months = {
                        "ENE",
                        "FEB",
                        "MAR",
                        "ABR",
                        "MAY",
                        "JUN",
                        "JUL",
                        "AGO",
                        "SEP",
                        "OCT",
                        "NOV",
                        "DIC"
                    };
            //Mes a literal
            var rawresult = _context.Database.SqlQuery<AsesoriaDocenteViewModel>(query).ToList();
            List<AsesoriaDocenteViewModel> list = new List<AsesoriaDocenteViewModel>();
            foreach (var element in rawresult)
            {
                element.MesLiteral = _months[element.Mes - 1];
                list.Add(element);
            };
            return list;
        }

        //registro por Id
        [HttpGet]
        [Route("api/AsesoriaDocente/{id}")]
        public IHttpActionResult getIndividualRecord(int id)
        {
            //datos para la tabla histórica
            var uniqueRecord = _context.AsesoriaDocente.FirstOrDefault(x => x.Id == id);
            if (uniqueRecord == null)
            {
                return BadRequest("Ese registro no existe");
            }
            else
            {
                return Ok(uniqueRecord);
            }
        }
        
        //obtener registros de tutorias segun su estado
        [HttpGet]
        [Route("api/AsesoriaDocente")]
        public IHttpActionResult getAsesoria([FromUri] string by)
        {
            //datos para la tabla histórica
            string query = "select a.\"Id\",a.\"TeacherFullName\", a.\"TeacherCUNI\", a.\"TeacherBP\", a.\"Categoría\", " +
                                "case when (a.\"Acta\") is null or (a.\"Acta\")='' then 'S/N' when (a.\"Acta\") is not null then a.\"Acta\" end as \"Acta\", a.\"ActaFecha\", a.\"BranchesId\", br.\"Abr\" as \"Regional\", a.\"Carrera\", a.\"DependencyCod\", a.\"Horas\", " +
                                "a.\"MontoHora\", a.\"TotalNeto\", a.\"TotalBruto\", a.\"StudentFullName\", a.\"Mes\", a.\"Gestion\", " +
                                "a.\"Observaciones\", a.\"Deduccion\", t.\"Abr\" as \"TipoTarea\", tm.\"Abr\" as \"Modalidad\", null as \"MesLiteral\", a.\"Origen\" " +
                                "from " + CustomSchema.Schema + ".\"AsesoriaDocente\" a " +
                                "inner join " + CustomSchema.Schema + ".\"TipoTarea\" t " +
                                "on a.\"TipoTareaId\"=t.\"Id\" " +
                                "inner join " + CustomSchema.Schema + ".\"Modalidades\" tm " +
                                "on a.\"ModalidadId\"=tm.\"Id\" " +
                                "inner join " + CustomSchema.Schema + ".\"Branches\" br " +
                                "on a.\"BranchesId\"=br.\"Id\" ";
            string orderBy = "order by a.\"Gestion\" desc, a.\"Mes\" desc, a.\"Carrera\" asc, a.\"TeacherCUNI\" asc ";

            var rawresult = new List<AsesoriaDocenteViewModel>();
            var user = auth.getUser(Request);

            if (by.Equals("APROBADO")) {
                string customQuery = query + "where a.\"Estado\"='APROBADO' " + orderBy;
                //Mes a literal
                rawresult = mesLiteral(customQuery);
                var filteredList = auth.filerByRegional(rawresult.AsQueryable(), user).ToList()
                    .Select(x => new { x.Id, x.Acta, x.Carrera, Profesor = x.TeacherFullName, Estudiante=x.StudentFullName, Tarea=x.TipoTarea, x.MesLiteral, x.Origen, x.Gestion });
                return Ok(filteredList);

            } else if (by.Equals("PRE-APROBADO")){
                string customQuery = query + "where a.\"Estado\"='PRE-APROBADO' " + orderBy;
                rawresult = _context.Database.SqlQuery<AsesoriaDocenteViewModel>(customQuery).ToList();
                var filteredList = auth.filerByRegional(rawresult.AsQueryable(), user).ToList()
                    .Select(x => new {x.Id, x.TeacherFullName, x.Acta, x.Carrera, x.StudentFullName, x.TipoTarea, x.Modalidad, x.TotalNeto, x.TotalBruto });;
                return Ok(filteredList);

            } else if (by.Equals("REGISTRADO-DEP")){
                //para la pantalla de aprobación nos interesan los registrados nada más
                string customQuery = query + "where a.\"Estado\"='REGISTRADO' " + "and a.\"Origen\"='DEPEN' " +orderBy;
                rawresult = _context.Database.SqlQuery<AsesoriaDocenteViewModel>(customQuery).ToList();
                var filteredList = auth.filerByRegional(rawresult.AsQueryable(), user).ToList()
                    .Select(x => new { x.Id, x.TeacherFullName, x.Acta, x.Carrera, x.StudentFullName, x.TipoTarea, x.Modalidad, x.TotalNeto, x.TotalBruto }); ; ;
                return Ok(filteredList);

            } else if (by.Equals("REGISTRADO-INDEP")) {
            //para la pantalla de aprobación nos interesan los registrados nada más
                string customQuery = query + "where a.\"Estado\"='REGISTRADO' " +"and a.\"Origen\"='INDEP' " +orderBy;
                rawresult = _context.Database.SqlQuery<AsesoriaDocenteViewModel>(customQuery).ToList();
                var filteredList = auth.filerByRegional(rawresult.AsQueryable(), user).ToList()
                    .Select(x => new { x.Id, x.TeacherFullName, x.Acta, x.Carrera, x.StudentFullName, x.TipoTarea, x.Modalidad, x.TotalNeto, x.TotalBruto }); ; ;
                return Ok(filteredList);
            }
            else if (by.Equals("REGISTRADO-OR"))
            {
                //para la pantalla de aprobación nos interesan los registrados nada más
                string customQuery = query + "where a.\"Estado\"='REGISTRADO' " + "and a.\"Origen\"='OR' " + orderBy;
                rawresult = _context.Database.SqlQuery<AsesoriaDocenteViewModel>(customQuery).ToList();
                var filteredList = auth.filerByRegional(rawresult.AsQueryable(), user).ToList()
                    .Select(x => new { x.Id, x.TeacherFullName, x.Acta, x.Carrera, x.StudentFullName, x.TipoTarea, x.Modalidad, x.TotalNeto, x.TotalBruto });
                return Ok(filteredList);
            }
            else {
                return BadRequest();            
            }

        }

        //conseguir los registros del docente por nombre completo, esto se debe a que no todos los registros tienen cuni o socio de negocio
        [HttpGet]
        [Route("api/TeacherStudent/{id}")]
        public IHttpActionResult TeachingRecords(int id)
        {
            var record = _context.AsesoriaDocente.FirstOrDefault(x => x.Id == id).TeacherFullName;
            //muestra los registros aprobados del docente X
            var query = "select a.*, t.\"Abr\" as \"TipoTarea\", tm.\"Abr\" as \"Modalidad\" from " + CustomSchema.Schema + ".\"AsesoriaDocente\" a " +
                        "inner join " + CustomSchema.Schema + ".\"TipoTarea\" t " +
                             "on a.\"TipoTareaId\"=t.\"Id\" " +
                        "inner join " + CustomSchema.Schema + ".\"Modalidades\" tm " +
                             "on a.\"ModalidadId\"=tm.\"Id\" " +
                        "where " +
                        "  \"TeacherFullName\"= '" + record + "' " +
                        "   and \"Estado\"='APROBADO' " +
                        "order by a.\"Gestion\" desc, a.\"Mes\" desc, a.\"Carrera\" asc, a.\"TeacherCUNI\" asc ";
            var allTeachingRecords = mesLiteral(query).Select(x => new { x.Id, x.Modalidad, x.TipoTarea, x.Carrera, x.Horas, x.MontoHora, x.TotalBruto, x.Deduccion, x.TotalNeto, Estudiante=x.StudentFullName, x.MesLiteral, x.Gestion});

            return Ok(allTeachingRecords);
        }

        // lista de docentes para el registro
        [HttpGet]
        [Route("api/DocentesList")]
        public IHttpActionResult DocentesList()
        {
            //Hacer un union con los docentes que no sean indepedientes, es decir que sean de civil nomas, por su jobTitle
            var activeDocentes = _context.Database.SqlQuery<AsesoriaTeachers>("(select lc.\"CUNI\", fn.\"FullName\",lc.\"StartDate\", lc.\"EndDate\", lc.\"BranchesId\", true as \"TipoPago\", lc.\"Categoria\" " +
            "from " + CustomSchema.Schema + ".\"ContractDetail\" lc " +
            "inner join " + CustomSchema.Schema + ".\"FullName\" fn " +
            "on fn.\"CUNI\"=lc.\"CUNI\" " +
            "where lc.\"Categoria\" is not null " +
            "and ( year(lc.\"EndDate\")*100+month(lc.\"EndDate\")>= year(current_date)*100+month(current_date) " +
            "or lc.\"EndDate\" is null)) " +
                //aquí juntamos a las personas de ADMNALRHH con los profesores independientes, es decir que estan como socios de negocio
            "UNION ALL " +
            "(select cv.\"SAPId\" as \"CUNI\",cv.\"FullName\", null as \"StartDate\", null as \"EndDate\", br.\"Id\" as \"BranchesId\", false as \"TipoPago\", cv.\"Categoria\" " +
            "from " + CustomSchema.Schema + ".\"Civil\" cv " +
            "inner join " + ConfigurationManager.AppSettings["B1CompanyDB"] + ".CRD8 " +
            "on cv.\"SAPId\" = crd8.\"CardCode\" " +
            "inner join " + CustomSchema.Schema + ".\"Branches\" br " +
            "on crd8.\"BPLId\"=br.\"CodigoSAP\") " +
            "order by \"FullName\" "
                //"where oh.\"jobTitle\" like '%DOCENTE%' "
            ).ToList();


            var user = auth.getUser(Request);

            var filteredList = auth.filerByRegional(activeDocentes.AsQueryable(), user);

            return Ok(filteredList);
        }

        //para obtener el cuerpo del reporte PDF
        [HttpGet]
        [Route("api/PDFReportBody")]
        public IHttpActionResult PDFReport([FromUri] string part)
        {
            string query = "";
            var report = new List<AsesoriaDocenteViewModel>();
            string[] data = part.Split(';');
            string section = data[0];
            string state = data[1];
            string origin = data[2];
            //query para generar todos los datos de cada docente, ordenado por carrera y docente
            switch (section)
            {
                case "Body":
                    //obtiene el cuerpo de la tabla para el PDF
                    //join para el nombre de la carrera
                    query = "select " +
                            "\"TeacherFullName\", \"Categoría\", " +
                            "m.\"Abr\" as \"Modalidad\", " +
                            "t.\"Abr\" as \"TipoTarea\", " +
                            "a.\"Carrera\" ||"+ " ' ' " +"|| op.\"PrcName\" as \"Carrera\" " + ", \"StudentFullName\" , " +
                            "\"Acta\", \"ActaFecha\" , " +
                            "\"Horas\", \"MontoHora\", " +
                            "\"TotalBruto\" , " +
                            "\"TotalNeto\" , " +
                            "\"Observaciones\", \"BranchesId\" " +
                        "from " +
                            CustomSchema.Schema + ".\"AsesoriaDocente\" a " +
                        "inner join " +
                            CustomSchema.Schema + ".\"TipoTarea\" t " +
                            "on a.\"TipoTareaId\"=t.\"Id\" " +
                        "inner join " +
                            CustomSchema.Schema + ".\"Modalidades\" m " +
                            "on a.\"ModalidadId\"=m.\"Id\" " +
                        "inner join " +
                            ConfigurationManager.AppSettings["B1CompanyDB"] + ".\"OPRC\" op " +
                            "on a.\"Carrera\"= op.\"PrcCode\" "+
                        "where " +
                            "a.\"Estado\"='" + state + "' "+
                            "and a.\"Origen\" like '%"+ origin + "%'" +
                            "and op.\"DimCode\" = 3 "+
                        "order by \"Carrera\", \"TeacherFullName\" ";
                    report = _context.Database.SqlQuery<AsesoriaDocenteViewModel>(query).ToList();
                    break;

                case "Results":
                    //obtiene los resultados al pie de cada tabla, por carrera
                    query = "select " +
                            "(a.\"Carrera\" ||"+ " ' ' " +"|| op.\"PrcName\") as \"Carrera\", "+
                            "sum(\"TotalBruto\") as \"TotalBruto\", " +
                            "sum(\"TotalNeto\") as \"TotalNeto\", \"BranchesId\" " +
                        "from " +
                            CustomSchema.Schema + ".\"AsesoriaDocente\" a " +
                        "inner join " +
                            ConfigurationManager.AppSettings["B1CompanyDB"] + ".\"OPRC\" op " +
                            "on a.\"Carrera\"= op.\"PrcCode\" " +
                        "where " +
                            "\"Estado\"='"+ state + "' " +
                            "and a.\"Origen\" like '%" + origin + "%'" +
                        "group by \"Carrera\", \"PrcName\", \"BranchesId\" " +
                        "order by \"Carrera\" ";
                    report = _context.Database.SqlQuery<AsesoriaDocenteViewModel>(query).ToList();
                    break;

                case "FinalResult":
                    //obtiene los resultados al pie de cada tabla, por carrera
                    query = "select " +
                            "sum(\"TotalBruto\") as \"TotalBruto\", " +
                            "sum(\"TotalNeto\") as \"TotalNeto\", \"BranchesId\" " +
                        "from " +
                            CustomSchema.Schema + ".\"AsesoriaDocente\" a " +
                        "where " +
                            "\"Estado\"='" + state + "' " +
                            "and a.\"Origen\" like '%" + origin + "%'" +
                        "group by \"BranchesId\" ";
                    report = _context.Database.SqlQuery<AsesoriaDocenteViewModel>(query).ToList();
                    break;

                default:
                    return BadRequest();
            }
            //Filtro de datos por regional
            var user = auth.getUser(Request);
            if (section.Equals("Body"))
            {
                var filteredListBody = auth.filerByRegional(report.AsQueryable(), user).ToList().Select(x => new
                {
                    Carrera = x.Carrera,
                    Docente = x.TeacherFullName,
                    Categ = x.Categoría,
                    Modal = x.Modalidad,
                    Tarea = x.TipoTarea,
                    Alumno = x.StudentFullName,
                    Acta = x.Acta,
                    Fecha = x.ActaFecha != null? x.ActaFecha.ToString("dd-MM-yyyy"): null,
                    Horas = x.Horas,
                    Costo_Hora = x.MontoHora,
                    Total_Bruto = x.TotalBruto,
                    Total_Neto = x.TotalNeto,
                    Observaciones = x.Observaciones
                });

                return Ok(filteredListBody);
            }
            else if (section.Equals("Results"))
            {
                var filteredListResult = auth.filerByRegional(report.AsQueryable(), user).ToList().Select(x => new
                {
                    Carrera = x.Carrera,
                    Total_Bruto = x.TotalBruto,
                    Total_Neto = x.TotalNeto,
                });
                return Ok(filteredListResult);
            }
            else
            {
                var filteredListResult = auth.filerByRegional(report.AsQueryable(), user).ToList().Select(x => new
                {
                    Total_Bruto = x.TotalBruto,
                    Total_Neto = x.TotalNeto,
                });
                return Ok(filteredListResult);
            }
        }

        //para generar el archivo PREGRADO de SALOMON
        [HttpGet]
        [Route("api/ToPregradoFile")]
        public HttpResponseMessage ToPregradoFile([FromUri] string data)
        {
            string[] info = data.Split(';');
            int segmentoId = Convert.ToInt16(info[0]);
            string segmento = _context.Branch.FirstOrDefault(x=>x.Id == segmentoId).Abr;
            string mes = (info[1]);
            string gestion = info[2];

            var process = _context.DistProcesses.FirstOrDefault(x => x.mes.Equals(mes) && x.gestion.Equals(gestion) && x.Branches.Abr.Equals(segmento) && x.State.Equals("INSAP"));
            //validar que ese proceso en SALOMON sea válido para la generación de datos
            if (process!=null)
            {
                HttpResponseMessage response =
                            new HttpResponseMessage(HttpStatusCode.InternalServerError);
                            response.Content = new StringContent("El periodo seleccionado no es válido para la generación del archivo PREGRADO en la regional "+ segmento);
                            response.RequestMessage = Request;
                            return response;
            }
            else {
                var user = auth.getUser(Request);
                //El query genera el archivo PREGRADO de SALOMON en base a los datos de las tutorías PRE-APROBADAS
                string query = "select  " +
                                    "\"Document\" ,\"FirstSurName\", \"SecondSurName\",  \"Names\", \"MariedSurName\", sum(\"TotalNeto\") as \"TotalNeto\", \"Carrera\" , \"CUNI\", \"Dependency\" " +
                                "from( " +
                                    "select " +
                                        "p.\"Document\" ,p.\"FirstSurName\", p.\"SecondSurName\", " +
                                        "p.\"Names\", p.\"MariedSurName\", " +
                                        "a.\"TotalNeto\", a.\"Carrera\" , " +
                                        "a.\"TeacherCUNI\" as \"CUNI\", a.\"DependencyCod\" as \"Dependency\", a.\"BranchesId\" " +
                                    "from " +
                                        CustomSchema.Schema + ".\"AsesoriaDocente\" a " +
                                        "inner join " + CustomSchema.Schema + ".\"People\" p " +
                                        "on a.\"TeacherCUNI\"=p.\"CUNI\" " +
                                        "inner join " + CustomSchema.Schema + ".\"LASTCONTRACTS\" lc " +
                                        "on a.\"TeacherCUNI\"=lc.\"CUNI\" " +
                                        "inner join " + CustomSchema.Schema + ".\"Branches\" br " +
                                        "on a.\"BranchesId\"=br.\"Id\" " +
                                    "where " +
                                        "a.\"Estado\"='PRE-APROBADO' " +
                                        "and br.\"Abr\" ='" + segmento + "' "+
                                        "and a.\"Origen\"='DEPEN' " +
                                    "order by a.\"Id\" desc) " +
                                 "group by \"Document\" ,\"FirstSurName\", \"SecondSurName\",  \"Names\", \"MariedSurName\", \"Carrera\" , \"CUNI\", \"Dependency\", \"BranchesId\" " +
                                 "order by \"Carrera\" asc, \"FirstSurName\" ";
                
                var excelContent = _context.Database.SqlQuery<DistPregradoViewModel>(query).ToList();

                var filteredWithoutCol = excelContent.Select(x => new { x.Document, x.FirstSurName, x.SecondSurName, x.Names, x.MariedSurName, x.TotalNeto, x.Carrera, x.CUNI, x.Dependency }).ToList();

                //--------------------------------------------------------Generación del excel------------------------------------------------------------------------
                //Para las columnas del excel
                string[] header = new string[]{"Carnet Identidad", "Primer Apellido", "Segundo Apellido", 
                                            "Nombres", "Apellido Casada", "Total Neto Ganado", "Código de Carrera", "CUNI", 
                                            "Identificador de dependencia"};
                var workbook = new XLWorkbook();

                //Se agrega la hoja de excel
                var ws = workbook.Worksheets.Add("PREGRADO");

                // Título
                ws.Cell("A1").Value = "PREGRADO";

                //Formato Cabecera
                ws.Cell(1, 1).Style.Font.Bold = true;
                ws.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1);
                ws.Cell(1, 1).Style.Font.FontName = "Bahnschrift SemiLight";
                ws.Cell(1, 1).Style.Font.FontSize = 20;
                ws.Cell(1, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                // Rango hoja excel
                //1,1: es la posicion inicial; 2,header.Length: es el alto y el ancho
                var rngTable = ws.Range(1, 1, 2, header.Length);

                //Bordes para las columnas
                var columns = ws.Range(3, 1, 2 + excelContent.Count, header.Length);
                columns.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                columns.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;


                //Para juntar celdas de la cabecera
                rngTable.Row(1).Merge();

                //auxiliar: desde qué línea ponemos los nombres de columna
                var headerPos = 2;

                //Ciclo para asignar los nombres a las columnas y darles formato
                for (int i = 0; i < header.Length; i++)
                {
                    ws.Column(i + 1).Width = 13;
                    ws.Cell(headerPos, i + 1).Value = header[i];
                    ws.Cell(headerPos, i + 1).Style.Alignment.WrapText = true;
                    ws.Cell(headerPos, i + 1).Style.Font.Bold = true;
                    ws.Cell(headerPos, i + 1).Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1);
                }

                //Aquí hago el attachment del query a mi hoja de de excel
                ws.Cell(3, 1).Value = filteredWithoutCol.AsEnumerable();

                //Ajustar contenidos
                ws.Columns().AdjustToContents();

                //Carga el objeto de la respuesta
                HttpResponseMessage response = new HttpResponseMessage();

                //Array de bytes
                var ms = new MemoryStream();
                workbook.SaveAs(ms);
                response.StatusCode = HttpStatusCode.OK;
                response.Content = new StreamContent(ms);
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                response.Content.Headers.ContentDisposition.FileName = segmento + mes + gestion + "PREG.xlsx";
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                response.Content.Headers.ContentLength = ms.Length;
                //La posicion para el comienzo del stream
                ms.Seek(0, SeekOrigin.Begin);

                //-----------------------------------------------------Cambios en PRE-APROBADOS ---------------------------------------------------------------------
                //Actualizar con la fecha a los registros pre-aprobados
                var docentesPorAprobar = _context.AsesoriaDocente.Where(x => x.Origen.Equals("DEPEN") && x.Estado.Equals("PRE-APROBADO") && x.BranchesId == segmentoId).ToList();
                //Se sobrescriben los registros con la fecha actual y el nuevo estado
                foreach (var docente in docentesPorAprobar)
                {
                    docente.Mes = Convert.ToInt16(mes);
                    docente.Gestion = Convert.ToInt16(gestion);
                    docente.Estado = "APROBADO";
                }

                _context.SaveChanges();

                return response;
            }
        }

        //para generar el archivo PREGRADO de SARAI
        [HttpGet]
        [Route("api/ToCarreraFile")]
        public HttpResponseMessage ToCarreraFile([FromUri] string data)
        {
            string[] info = data.Split(';');
            int segmentoId = Convert.ToInt16(info[0]);
            string segmento = _context.Branch.FirstOrDefault(x => x.Id == segmentoId).Abr;
            // el mes y la gestion son necesarios para guardar el registro histórico ISAAC
            string mes = (info[1]);
            string gestion = info[2];
            //El query genera el archivo PREGRADO de SALOMON en base a los datos de las tutorías PRE-APROBADAS
            string query =      
                "select " +
                    "a.\"TeacherBP\" as \"Codigo_Socio\", a.\"TeacherFullName\" as \"Nombre_Socio\", "+
                    "a.\"DependencyCod\" as \"Cod_Dependencia\", 'PO' as \"PEI_PO\", " +
                    "t.\"Tarea\" as \"Nombre_del_Servicio\", a.\"Carrera\" as \"Codigo_Carrera\" ,a.\"Acta\" as \"DocumentNumber\", " +
                    "a.\"StudentFullName\" as \"Postulante\", t.\"Type\" as \"Tipo_Tarea_Asignada\", 'CC_TEMPORAL' as \"Cuenta_Asignada\", " +
                    "a.\"TotalBruto\" as \"Monto_Contrato\", a.\"IUE\" as \"Monto_IUE\", a.\"IT\" as \"Monto_IT\", a.\"TotalNeto\" as \"Monto_a_Pagar\",  " +
                    "a.\"Observaciones\" " +
                "from " +
                    CustomSchema.Schema + ".\"AsesoriaDocente\" a " +
                    "inner join " + CustomSchema.Schema + ".\"Civil\" c " +
                    "on a.\"TeacherBP\"=c.\"SAPId\" " +
                    "inner join " + CustomSchema.Schema + ".\"TipoTarea\" t " +
                    "on a.\"TipoTareaId\"=t.\"Id\" " +
                    "inner join " + CustomSchema.Schema + ".\"Branches\" br " +
                    "on a.\"BranchesId\"=br.\"Id\" " +
                "where " +
                   "a.\"Estado\"='PRE-APROBADO' " +
                   "and br.\"Abr\" ='" + segmento + "' " +
                   "and a.\"Origen\"='INDEP' " +
                "order by a.\"Id\" desc";

            var excelContent = _context.Database.SqlQuery<Serv_PregradoViewModel>(query).ToList();

            //Para las columnas del excel
            string[] header = new string[]{"Codigo_Socio", "Nombre_Socio", "Cod", 
                                            "PEI_PO", "Nombre_del_Servicio", "Codigo_Carrera", "Documento_Base", "Postulante", 
                                            "Tipo_Tarea_Asignada", "Cuenta_Asignada", 
                                            "Monto_Contrato","Monto_IUE","Monto_IT","Monto_a_Pagar", "Observaciones"};
            var workbook = new XLWorkbook();

            //Se agrega la hoja de excel
            var ws = workbook.Worksheets.Add("Plantilla_CARRERA");

            // Rango hoja excel
            //1,1: es la posicion inicial; 2,header.Length: es el alto y el ancho
            var rngTable = ws.Range(1, 1, 2, header.Length);

            //Bordes para las columnas
            var columns = ws.Range(2, 1, excelContent.Count +1, header.Length);
            columns.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            columns.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            //auxiliar: desde qué línea ponemos los nombres de columna
            var headerPos = 1;

            //Ciclo para asignar los nombres a las columnas y darles formato
            for (int i = 0; i < header.Length; i++)
            {
                ws.Cell(headerPos, i + 1).Value = header[i];
                ws.Cell(headerPos, i + 1).Style.Font.Bold = true;
                ws.Cell(headerPos, i + 1).Style.Font.FontColor = XLColor.White;
                ws.Cell(headerPos, i + 1).Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1);
            }

            //Aquí hago el attachment del query a mi hoja de de excel
            ws.Cell(2, 1).Value = excelContent.AsEnumerable();

            //Ajustar contenidos
            ws.Columns().AdjustToContents();

            //Carga el objeto de la respuesta
            HttpResponseMessage response = new HttpResponseMessage();

            //Array de bytes
            var ms = new MemoryStream();
            workbook.SaveAs(ms);
            response.StatusCode = HttpStatusCode.OK;
            response.Content = new StreamContent(ms);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentDisposition.FileName = segmento+"-CC_CARRERA.xlsx";
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.Content.Headers.ContentLength = ms.Length;
            //La posicion para el comienzo del stream
            ms.Seek(0, SeekOrigin.Begin);

            //-----------------------------------------------------Cambios en PRE-APROBADOS INDEP ---------------------------------------------------------------------
            //Actualizar con la fecha a los registros pre-aprobados
            var branchesId = _context.Branch.FirstOrDefault(x=>x.Abr ==segmento);
            var docentesPorAprobar = _context.AsesoriaDocente.Where(x => x.Origen.Equals("INDEP") && x.Estado.Equals("PRE-APROBADO") && x.BranchesId == segmentoId).ToList();
            //Se sobrescriben los registros con la fecha actual y el nuevo estado
            foreach (var docente in docentesPorAprobar)
            {
                docente.Mes = Convert.ToInt16(mes);
                docente.Gestion = Convert.ToInt16(gestion);
                docente.Estado = "APROBADO";
            }

            _context.SaveChanges();

            return response;
        }

        //registro de la tutoria
        [HttpPost]
        [Route("api/AsesoriaDocente")]
        public IHttpActionResult Post([FromBody] AsesoriaDocente asesoria)
        {
            var B1conn = B1Connection.Instance();
            var user = auth.getUser(Request);
            // Ver que la persona esté disponible en nuestra base de personas y de civiles    
            if ((asesoria.Origen.Equals("DEPEN") || asesoria.Origen.Equals("OR")) && !_context.Person.ToList().Any(x => x.CUNI == asesoria.TeacherCUNI))
            {
                return BadRequest("La persona no existe en BD");
            }

            if (asesoria.Origen.Equals("INDEP") && !_context.Civils.ToList().Any(x => x.SAPId == asesoria.TeacherBP))
            {
                return BadRequest("La persona no existe en BD Civil");
            }

            if (!_context.Modalidades.ToList().Any(x => x.Id == asesoria.ModalidadId))
            {
                return BadRequest("La modalidad no existe en BD");
            }

            if (!_context.TipoTarea.ToList().Any(x => x.Id == asesoria.TipoTareaId))
            {
                return BadRequest("El tipo de tarea no existe en BD");
            }

            //Validación de la carrera
            List<dynamic> careerList = B1conn.getCareers();
            //Validar el noombre del código de la carrera con el ingresado
            if (!careerList.Exists(x => x.cod == asesoria.Carrera))
            {
                return BadRequest("La carrera no existe en SAP, al menos para esa regional");
            }
            if (asesoria.TotalBruto <= 0 || asesoria.TotalNeto <= 0)
            {
                return BadRequest("no se pueden ingresar datos con valores negativos o iguales a 0");
            }
            if (asesoria.Ignore == false && (_context.AsesoriaDocente.Where(x => x.StudentFullName == asesoria.StudentFullName && x.TeacherCUNI == asesoria.TeacherCUNI && x.Acta == asesoria.Acta).FirstOrDefault() != null))
            {
                return BadRequest("La combinación de docente, estudiante y acta ya existe en la BD");
            }
            else
            {
                //El branchesId es del último puesto de quién registra
                var userCUNI = user.People.CUNI;
                var regional = asesoria.Carrera;
                string[] Abr = regional.Split('-');
                var regionalId = Abr[0].ToString();

                asesoria.BranchesId = _context.Branch.FirstOrDefault(x => x.Abr.Equals(regionalId)).Id;
                //el Id del siguiente registro
                asesoria.Id = AsesoriaDocente.GetNextId(_context);
                //asegura que no se junte el nuevo registro con los históricos
                asesoria.Estado = "REGISTRADO";
                //identifica la dependencia del registro en base al nombre de la carrera y la regional
                var dep = _context.Database.SqlQuery<int>("select de.\"Cod\" " +
                                        "from " +
                                        "   " + ConfigurationManager.AppSettings["B1CompanyDB"] + ".oprc op " +
                                        "inner join " + ConfigurationManager.AppSettings["B1CompanyDB"] + ".\"@T_GEN_CARRERAS\" tg " +
                                        "    on op.\"PrcCode\" = tg.\"U_CODIGO_CARRERA\" " +
                                        "inner join " + CustomSchema.Schema + ".\"OrganizationalUnit\" ou " +
                                        "    on tg.\"U_CODIGO_DEPARTAMENTO\"=ou.\"Cod\" " +
                                        "inner join " + CustomSchema.Schema + ".\"Dependency\" de " +
                                        "    on ou.\"Id\"=de.\"OrganizationalUnitId\" " +
                                        "where " +
                                            "op.\"DimCode\"=3 " +
                                            "and op.\"PrcCode\" ='" + asesoria.Carrera + "' " +
                                            "and de.\"BranchesId\"=" + asesoria.BranchesId).FirstOrDefault().ToString();
                asesoria.DependencyCod = dep;
                //agregar el nuevo registro en el contexto
                _context.AsesoriaDocente.Add(asesoria);
                _context.SaveChanges();
                return Ok("Información registrada");
            }
        }

        [HttpPut]
        [Route("api/AsesoriaDocente/{id}")]
        public IHttpActionResult Put(int id, [FromBody] AsesoriaDocente asesoria)
        {
            if (!_context.AsesoriaDocente.ToList().Any(x => x.Id == id))
            {
                return BadRequest("No existe el registro correspondiente");
            }
            else
            {
                if (!ModelState.IsValid)
                {
                    return BadRequest("Datos inválidos para el registro");
                }
                var thisAsesoria = _context.AsesoriaDocente.FirstOrDefault(x => x.Id == id);
                //Temporalidad
                thisAsesoria.Mes = asesoria.Mes;
                thisAsesoria.Gestion = asesoria.Gestion;
                //Carrera y Dep
                thisAsesoria.DependencyCod = asesoria.DependencyCod;
                thisAsesoria.Carrera = asesoria.Carrera;
                //Docente
                thisAsesoria.TeacherCUNI = asesoria.TeacherCUNI;
                thisAsesoria.TeacherFullName = asesoria.TeacherFullName;
                thisAsesoria.TeacherBP = asesoria.TeacherBP;
                thisAsesoria.Categoría = asesoria.Categoría;
                thisAsesoria.Origen = asesoria.Origen;
                //Estudiante
                thisAsesoria.StudentFullName = asesoria.StudentFullName;
                //Sobre la tutoria
                thisAsesoria.TipoTareaId = asesoria.TipoTareaId;
                thisAsesoria.ModalidadId = asesoria.ModalidadId;
                thisAsesoria.Ignore = asesoria.Ignore;
                //Sobre costos
                thisAsesoria.Horas = asesoria.Horas;
                thisAsesoria.MontoHora = asesoria.MontoHora;
                thisAsesoria.TotalBruto = asesoria.TotalBruto;
                thisAsesoria.TotalNeto = asesoria.TotalNeto;
                thisAsesoria.Deduccion = asesoria.Deduccion;
                thisAsesoria.Observaciones = asesoria.Observaciones;
                thisAsesoria.IUE = asesoria.IUE;
                thisAsesoria.IT = asesoria.IT;
                //Del Acta
                thisAsesoria.Acta = asesoria.Acta;
                thisAsesoria.ActaFecha = asesoria.ActaFecha;
                thisAsesoria.BranchesId = asesoria.BranchesId;
                //Modifica su estado
                thisAsesoria.Estado = asesoria.Estado;
                _context.SaveChanges();
                return Ok("Se actualizaron los datos correctamente");
            }
        }

        //para la instancia de el modulo de aprobacion Isaac, pasar a pre-aprobacion
        [HttpPut]
        [Route("api/ToPreAprobacion")]
        public IHttpActionResult ToPreAprobacion([FromUri] string myArray)
        {
            if (myArray == null)
            {
                return BadRequest("No se ha seleccionado ningún registro para aprobación");
            }
            else
            {
                var countRegister = 0;
                int[] array = Array.ConvertAll(myArray.Split(','), int.Parse);
                int[] failedUpdates = new int[array.Length];
                for (int i = 0; i < array.Length; i++)
                {
                    int currentElement = array[i];
                    var thisAsesoria = _context.AsesoriaDocente.FirstOrDefault(x => x.Id == currentElement);
                    if (thisAsesoria != null)
                    {
                        if (thisAsesoria.Origen.Equals("OR"))
                        {
                            thisAsesoria.Estado = "APROBADO";
                        }
                        else {
                            thisAsesoria.Estado = "PRE-APROBADO";
                        }
                        _context.SaveChanges();
                    }
                    else
                    {
                        //Hubieron elementos del array que no se pudieron actualizar
                        failedUpdates[countRegister] = array[i];
                        countRegister += 1;
                    }
                }
                //Si tenemos todos los Ids
                if (countRegister == 0)
                {
                    return Ok("Se actualizaron los registros exitosamente");
                }
                //Si fallan todos los Ids
                else if (countRegister == array.Length)
                {
                    return BadRequest("No se pudo actualizar ningún registro");
                }
                //Si solo fallan algunos
                else
                {
                    return Ok("No se pudieron actualizar los siguientes registros:" + failedUpdates);//aquí meterle el concat por comas
                }
            }
        }

        //para la instancia de el modulo de aprobacion Isaac
        [HttpDelete]
        [Route("api/DeleteRecord/{id}")]
        public IHttpActionResult DeleteRecord(int id)
        {
            //solo borrarlo en la primera instancia, no se eliminan los aprobados
            var recordForDeletion = _context.AsesoriaDocente.FirstOrDefault(x => x.Id == id && x.Estado == "REGISTRADO");
            if (recordForDeletion == null)
            {
                return BadRequest("El registro no existe en BD");
            }
            else
            {
                _context.AsesoriaDocente.Remove(recordForDeletion);
                _context.SaveChanges();
                return Ok("Se eliminó el registro exitosamente");
            }
        }

    }
}

