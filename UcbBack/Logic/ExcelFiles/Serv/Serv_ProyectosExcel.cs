using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using UcbBack.Logic.B1;
using UcbBack.Models;
using UcbBack.Models.Auth;
using UcbBack.Models.Serv;
using System.Configuration;
using UcbBack.Models.Not_Mapped.CustomDataAnnotations;
using System.Globalization;
using System.Data.Entity;
using Newtonsoft.Json.Linq;

namespace UcbBack.Logic.ExcelFiles.Serv
{
    public class Serv_ProyectosExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Codigo Socio", typeof(string)), 
            new Excelcol("Nombre Socio", typeof(string)),
            new Excelcol("Cod Dependencia", typeof(string)),
            new Excelcol("PEI PO", typeof(string)),
            new Excelcol("Nombre del Servicio", typeof(string)),
            new Excelcol("Codigo Proyecto SAP", typeof(string)),
            new Excelcol("Nombre del Proyecto", typeof(string)),
            new Excelcol("Version", typeof(string)),
            new Excelcol("Periodo Academico", typeof(string)),
            new Excelcol("Tipo Tarea Asignada", typeof(string)),
            new Excelcol("Cuenta Asignada", typeof(string)),
            new Excelcol("Monto Contrato", typeof(double)),
            new Excelcol("Monto IUE", typeof(double)),
            new Excelcol("Monto IT", typeof(double)),
            new Excelcol("Monto a Pagar", typeof(double)),
            new Excelcol("Observaciones", typeof(string)),
        };

        private ApplicationDbContext _context;
        private ServProcess process;
        private CustomUser user;

        public Serv_ProyectosExcel(string fileName, int headerin = 1)
            : base(cols, fileName, headerin)
        { }

        public Serv_ProyectosExcel(Stream data, ApplicationDbContext context, string fileName, ServProcess process, CustomUser user,int headerin = 1, int sheets = 1, string resultfileName = "Result")
            : base(cols, data, fileName, headerin, sheets, resultfileName, context)
        {
            this.process = process;
            this.user = user;
            _context = context;
            isFormatValid();
        }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                _context.ServProyectoses.Add(ToServVarios(i));
            }

            _context.SaveChanges();
        }

        public Serv_Proyectos ToServVarios(int row, int sheet = 1)
        {
            Serv_Proyectos data = new Serv_Proyectos();
            data.Id = Serv_Proyectos.GetNextId(_context);

            data.CardCode = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            data.CardName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            var cod = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            var depId = _context.Dependencies
                .FirstOrDefault(x => x.Cod == cod);
            data.DependencyId = depId.Id;
            data.PEI = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
            data.ServiceName = wb.Worksheet(sheet).Cell(row, 5).Value.ToString();
            data.ProjectSAPCode = wb.Worksheet(sheet).Cell(row, 6).Value.ToString();
            data.ProjectSAPName = wb.Worksheet(sheet).Cell(row, 7).Value.ToString();
            data.Version = wb.Worksheet(sheet).Cell(row, 8).Value.ToString();
            data.Periodo = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();
            data.AssignedJob = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();

            data.AssignedAccount = wb.Worksheet(sheet).Cell(row, 11).Value.ToString();
            data.ContractAmount = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 12).Value.ToString());
            data.IUE = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 13).Value.ToString());
            data.IT = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 14).Value.ToString());
            data.TotalAmount = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 15).Value.ToString());
            data.Comments = wb.Worksheet(sheet).Cell(row, 16).Value.ToString();
            data.Serv_ProcessId = process.Id;
            return data;
        }

        public override bool ValidateFile()
        {
            if (isValid())
            {
                var connB1 = B1Connection.Instance();

                if (!connB1.connectedtoHana)
                {
                    addError("Error en SAP", "No se puedo conectar con SAP B1, es posible que algunas validaciones cruzadas con SAP no sean ejecutadas");
                }

                bool v1 = VerifyBP(1, 2,process.BranchesId,user);
                bool v2 = VerifyColumnValueIn(3, _context.Dependencies.Where(x => x.BranchesId == this.process.BranchesId).Select(x => x.Cod).ToList(), comment: "Esta Dependencia no es Válida");
                var pei = connB1.getCostCenter(B1Connection.Dimension.PEI).Cast<String>().ToList();
                bool v3 = VerifyColumnValueIn(4, pei, comment: "Este PEI no existe en SAP.");
                bool v4 = VerifyLength(5, 50);
                bool v5 = verifyproject(dependency:3);
                var periodo = connB1.getCostCenter(B1Connection.Dimension.Periodo).Cast<string>().ToList();
                bool v6 = VerifyColumnValueIn(9, periodo, comment: "Este Periodo no existe en SAP.");
                bool v7 = VerifyColumnValueIn(10, new List<string> { "PROF", "TG", "REL", "LEC", "REV", "PAN", "OTR" }, comment: "No existe este tipo de Tarea Asignada.");
                bool v8 = VerifyColumnValueIn(11, new List<string> { "CC_POST", "CC_EC", "CC_FC", "CC_INV", "CC_SA" }, comment: "No existe este tipo de Cuenta Asignada.");
                //Nueva validación para comprobar que la cuenta asignada corresponde al proyecto
                bool v9 = true;
                foreach (var i in new List<int>(){1,2,3,4,5  ,7  ,9,10,11,12,13,14,15})
                {
                    v9 = VerifyNotEmpty(i) && v9;
                }
                bool v10 = verifyAccounts(dependency:3);
                bool v11 = verifyDates(dependency: 3);

                return v1 && v2 && v3 && v4 && v5 && v6 && v7 && v8 && v9;
            }

            return false;

        }

        private bool verifyproject(int dependency, int sheet = 1)
        {
            string commnet = "Este proyecto no existe en SAP.";
            var connB1 = B1Connection.Instance();
            var br = _context.Branch.FirstOrDefault(x => x.Id == process.BranchesId);
            var list = connB1.getProjects("*").Where(x => x.U_Sucursal == br.Abr).Select(x => new { x.PrjCode, x.U_UORGANIZA }).ToList();
            int index = 6;
            int tipoproy = 11;
            bool res = true;
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            
            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                var strproject = index != -1 ? wb.Worksheet(sheet).Cell(i, index).Value.ToString() : null;
                var strdependency = dependency != -1 ? wb.Worksheet(sheet).Cell(i, dependency).Value.ToString() : null;
                var dep = _context.Dependencies.Where(x => x.BranchesId == br.Id).Include(x => x.OrganizationalUnit).FirstOrDefault(x => x.Cod == strdependency);
                //------------------------------------Valida existencia del proyecto--------------------------------
                //Si no existe en esta rama un proyecto que haga match con el proyecto del excel
                if (!list.Exists(x => string.Equals(x.PrjCode.ToString(), strproject, StringComparison.OrdinalIgnoreCase)))
                {
                    //Si el tipo de proyecto, no es de los siguientes tipos y el codigo del proyecto no viene vacío
                    if (!(
                        (
                            wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString() == "CC_EC"
                            || wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString() == "CC_FC"
                            || wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString() == "CC_SA"
                        )
                        &&
                        wb.Worksheet(sheet).Cell(i, index).Value.ToString() == ""
                    ))
                    {
                        res = false;
                        paintXY(index, i, XLColor.Red, commnet);
                    }
                }
                else
                {
                    //como ya sabemos que existe el proyecto, ahora preguntamos de la UO
                    //dep es de la celda correcta
                    var row = list.FirstOrDefault(x => x.PrjCode == strproject);
                    string UO = row.U_UORGANIZA.ToString();
                    string UOName = _context.OrganizationalUnits.FirstOrDefault(x => x.Cod == UO).Name;
                    if (!string.Equals(dep.OrganizationalUnit.Cod.ToString(), row.U_UORGANIZA.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        //Si la UO para esta fila es diferente de la UO registrada en SAP, marcamos error
                        res = false;
                        paintXY(dependency, i, XLColor.Red, "Este proyecto debe tener una dependencia asociada a la Unidad Org: " + row.U_UORGANIZA + " " + UOName);
                    }

                }
            }
            valid = valid && res;
            if (!res)
            {
                addError("Valor no valido", "Proyecto o proyectos no validos en la columna: " + index, false);
            }
            
            return res;
        }

       private bool verifyAccounts(int dependency, int sheet = 1)
       {
           string commnet;//especifica el error
           var connB1 = B1Connection.Instance();
           var br = _context.Branch.FirstOrDefault(x => x.Id == process.BranchesId);
           //todos los proyectos de esa rama
           var list = connB1.getProjects("*").Where(x => x.U_Sucursal == br.Abr).Select(x => new { x.PrjCode, x.U_Tipo, x.ValidFrom, x.ValidTo, x.U_UORGANIZA }).ToList();
           //columnas del excel
           int index = 6;
           int tipoProyecto = 11;
           bool res = true;
           int badAccount = 0;
           int badType = 0;
           IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
           for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
           {
               if (list.Exists(x => string.Equals(x.PrjCode.ToString(), wb.Worksheet(sheet).Cell(i, index).Value.ToString(), StringComparison.OrdinalIgnoreCase)))
               {

                   if (wb.Worksheet(sheet).Cell(i, tipoProyecto).Value.ToString() != "CAP")
                   {
                        //-----------------------------Validaciones de la cuenta--------------------------------
                       var tiposBD = _context.TableOfTableses.Where(x => x.Type.Equals("TIPOS_P&C_SARAI")).Select(x => x.Value).ToList();
                       var projectType = list.Where(x => x.PrjCode == wb.Worksheet(sheet).Cell(i, index).Value.ToString()).FirstOrDefault().U_Tipo.ToString();//tipo de proyecto del proyecto en la celda
                       string tipo = projectType;//variable auxiliar, no puede usarse la de arriba en EF por ser dinámica
                       var typeExists = tiposBD.Exists(x => string.Equals(x.Split(':')[0], tipo, StringComparison.OrdinalIgnoreCase));
                       //el tipo de proyecto existe en nuestra tabla de tablas?
                       if (!typeExists)
                       {
                           commnet = "El tipo de proyecto: " + tipo + " no es válido.";
                           paintXY(index, i, XLColor.Red, commnet);
                           res = false;
                           badType++;
                       }
                       else
                       {
                           var projectAccount = wb.Worksheet(sheet).Cell(i, tipoProyecto).Value.ToString();
                           var assignedAccount = tiposBD.Where(x => x.Split(':')[0].Equals(tipo)).FirstOrDefault().ToString().Split(':')[1];
                           if (projectAccount != assignedAccount)
                           {
                               commnet = "La cuenta asignada es incorrecta, debería ser: " + assignedAccount;
                               paintXY(tipoProyecto, i, XLColor.Red, commnet);
                               res = false;
                               badAccount++;
                           }
                       }
                   }
                   
               }
           }
           valid = valid && res;
           if (!res && badAccount > 0 && badType > 0) { addError("Valor no valido", "Tipos de proyectos no válidos en la columna: " + index + " y cuentas asignadas no válidas en la columna: " + tipoProyecto, false); }
           else if (!res && badAccount > 0 && badType == 0) { addError("Valor no valido", "Cuentas asignadas no válidas en la columna: " + tipoProyecto, false); }
           else if (!res && badAccount == 0 && badType > 0) { addError("Valor no valido", "Tipos de proyectos no válidos en la columna: " + index, false); }
           
           return res;
       }

       private bool verifyDates(int dependency, int sheet = 1)
       {
           string commnet;//especifica el error
           var connB1 = B1Connection.Instance();
           var br = _context.Branch.FirstOrDefault(x => x.Id == process.BranchesId);
           //todos los proyectos de esa rama
           var list = connB1.getProjects("*").Where(x => x.U_Sucursal == br.Abr).Select(x => new { x.PrjCode, x.U_Tipo, x.ValidFrom, x.ValidTo, x.U_UORGANIZA }).ToList();
           //columnas del excel
           int index = 6;
           bool res = true;
           IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
           var l = UsedRange.LastRow().RowNumber();

           for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
           {
               if (list.Exists(x => string.Equals(x.PrjCode.ToString(), wb.Worksheet(sheet).Cell(i, index).Value.ToString(), StringComparison.OrdinalIgnoreCase)))
               {
                   var strproject = index != -1 ? wb.Worksheet(sheet).Cell(i, index).Value.ToString() : null;
                   var row = list.FirstOrDefault(x => x.PrjCode == strproject);
                   string UO = row.U_UORGANIZA.ToString();
                   var strdependency = dependency != -1 ? wb.Worksheet(sheet).Cell(i, dependency).Value.ToString() : null;
                   var dep = _context.Dependencies.Where(x => x.BranchesId == br.Id).Include(x => x.OrganizationalUnit).FirstOrDefault(x => x.Cod == strdependency);
                   //Si la UO hace match también
                   if(row.U_UORGANIZA==dep.OrganizationalUnit.Cod){
                       //-----------------------------Validaciones de la fecha del proyecto--------------------------------
                       var projectInitialDate = list.Where(x => x.PrjCode == wb.Worksheet(sheet).Cell(i, index).Value.ToString()).FirstOrDefault().ValidFrom.ToString();
                       DateTime parsedIni = Convert.ToDateTime(projectInitialDate);
                       var projectFinalDate = list.Where(x => x.PrjCode == wb.Worksheet(sheet).Cell(i, index).Value.ToString()).FirstOrDefault().ValidTo.ToString();
                       DateTime parsedFin = Convert.ToDateTime(projectFinalDate);

                       //si el tiempo actual es menor al inicio del proyecto en SAP ó si el tiempo actual es mayor a la fecha límite del proyectoSAP
                       if (System.DateTime.Now < parsedIni || System.DateTime.Now > parsedFin)
                       {
                           res = false;
                           commnet = "La fecha de este proyecto ya está cerrada, estuvo disponible del " + parsedIni + " al " + parsedFin;
                           paintXY(index, i, XLColor.Red, commnet);
                       }
                   }
               }
           }
           valid = valid && res;
           if (!res) { addError("Valor no valido", "Proyecto/s con fechas no válidas en la columna:" + index, false); }

           return res;
       }


      
    }
}