﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using UcbBack.Logic.B1;
using UcbBack.Models;
using UcbBack.Models.Dist;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2013.Word;
using ExcelDataReader;
using Newtonsoft.Json.Linq;
using System.Data.Entity;
using System.Diagnostics;
using UcbBack.Controllers;
using UcbBack.Models.Auth;


namespace UcbBack.Logic.ExcelFiles
{
    public class PostgradoExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Carnet Identidad", typeof(string)), 
            new Excelcol("Primer Apellido", typeof(string)),
            new Excelcol("Segundo Apellido", typeof(string)),
            new Excelcol("Nombres", typeof(string)),
            new Excelcol("Apellido Casada", typeof(string)),
            new Excelcol("Nombre del Proyecto", typeof(string)),
            new Excelcol("Versión", typeof(string)),
            new Excelcol("Total Neto Ganado", typeof(double)),
            new Excelcol("Identificador de Dependencia", typeof(string)),
            new Excelcol("CUNI", typeof(string)),
            new Excelcol("Tipo Proyecto", typeof(string)),
            new Excelcol("Tipo de tarea asignada", typeof(string)),
            new Excelcol("PEI", typeof(string)),
            new Excelcol("Periodo académico", typeof(string)),
            new Excelcol("Código Proyecto SAP", typeof(string))
        };
        private ApplicationDbContext _context;
        private string mes, gestion, segmentoOrigen;
        private Dist_File file;
        public PostgradoExcel(Stream data, ApplicationDbContext context, string fileName, string mes, string gestion, string segmentoOrigen, Dist_File file, int headerin = 3, int sheets = 1, string resultfileName = "Result")
            : base(cols, data, fileName, headerin, sheets, resultfileName)
        {
            this.segmentoOrigen = segmentoOrigen;
            this.gestion = gestion;
            switch (mes)
            {
                case "13":
                    this.mes = "01";
                    break;
                case "14":
                    this.mes = "02";
                    break;
                case "15":
                    this.mes = "03";
                    break;
                case "16":
                    this.mes = "04";
                    break;
                default:
                    this.mes = mes;
                    break;
            }
            this.file = file;
            _context = context;
            isFormatValid();
        }
        public PostgradoExcel(string fileName, int headerin = 1)
            : base(cols, fileName, headerin)
        { }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                _context.DistPosgrados.Add(ToDistDiscounts(i));
            }

            _context.SaveChanges();
        }

        public override bool ValidateFile()
        {
            var connB1 = B1Connection.Instance();

            //bool v1 = VerifyColumnValueIn(6, connB1.getProjects(col: "PrjName").ToList(), comment: "Este Proyecto no existe en SAP.");
            int brId = Int32.Parse(this.segmentoOrigen);
            bool v2 = VerifyColumnValueIn(9, _context.Dependencies.Where(x => x.BranchesId == brId).Select(m => m.Cod).Distinct().ToList(), comment: "No existe esta dependencia.");
            //qué tipo de proy debe ser para asignar a CAP?
            bool v3 = VerifyColumnValueIn(11, _context.TipoEmpleadoDists.Select(x => x.Name).ToList().Where(x => new List<string> { "POST", "EC", "INV", "FC", "SA", "CAP" }.Contains(x)).ToList(), comment: "No existe este tipo de proyecto.");
            bool v4 = VerifyColumnValueIn(12, new List<string> { "PROF", "TG", "REL", "LEC", "REV", "OTR", "PAN" }, comment: "No existe este tipo de tarea asignada.");
            bool v5 = VerifyColumnValueIn(13, connB1.getCostCenter(B1Connection.Dimension.PEI, mes: this.mes, gestion: this.gestion).Cast<string>().ToList(), comment: "Este PEI no existe en SAP.");
            bool v6 = VerifyColumnValueIn(14, connB1.getCostCenter(B1Connection.Dimension.Periodo, mes: this.mes, gestion: this.gestion).Cast<string>().ToList(), comment: "Este periodo no existe en SAP.");
            //bool v7 = VerifyColumnValueIn(15, connB1.getProjects(), comment: "Este proyecto no existe en SAP.");
            bool v7 = verifyproject(dependency: 9);
            bool v8 = VerifyPerson(ci: 1, fullname: 2, CUNI: 10, date: gestion + "-" + mes + "-01", personActive: false);
            bool v9 = verifyAcount(dependency: 9);
            bool v10 = verifyDates(dependency:9);
            return isValid() && v2 && v3 && v4 && v8 && v5 && v6 && v7 && v9 && v10;
        }

        private bool verifyproject(int dependency, int sheet = 1)
        {
            string commnet = "Este proyecto no existe en SAP.";
            var connB1 = B1Connection.Instance();
            int branchId = Convert.ToInt16(segmentoOrigen);
            var branch = _context.Branch.FirstOrDefault(x => x.Id==branchId).Abr;
            var list = connB1.getProjects("*").Where(x => x.U_Sucursal == branch).Select(x => new { x.PrjCode, x.U_UORGANIZA }).ToList();
            int index = 15;
            int tipoproy = 11;
            bool res = true;            
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            var l = UsedRange.LastRow().RowNumber();

            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                var strdependency = dependency != -1 ? wb.Worksheet(sheet).Cell(i, dependency).Value.ToString() : null;
                var strproject = index != -1 ? wb.Worksheet(sheet).Cell(i, index).Value.ToString() : null;
                var dep = _context.Dependencies.Where(x => x.BranchesId == branchId).Include(x => x.OrganizationalUnit).FirstOrDefault(x => x.Cod == strdependency);
                if (!list.Exists(x => string.Equals(x.PrjCode.ToString(), strproject, StringComparison.OrdinalIgnoreCase)))
                //Si el proyecto no existe en la lista de proyectos SAP. Solo entrará si está vacío o es un proyecto no registrado segun Codigo y su UO correspondiente
                {
                    var a1 = wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString();
                    var a2 = wb.Worksheet(sheet).Cell(i, index).Value.ToString();
                    if (!(
                            (
                        //Estos tipos de proyecto pueden no tener codigo de proyecto
                        //POST e INV siempre deben estar con un codigo de proyecto
                                wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString() == "EC"
                                || wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString() == "FC"
                                || wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString() == "SA"
                                || wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString() == "CAP"
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
                    var row=list.FirstOrDefault(x=>x.PrjCode==strproject);
                    string UO=row.U_UORGANIZA.ToString();
                    string UOName=_context.OrganizationalUnits.FirstOrDefault(x=>x.Cod==UO).Name;
                    if (!string.Equals(dep.OrganizationalUnit.Cod.ToString(), row.U_UORGANIZA.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        //Si la UO para esta fila es diferente de la UO registrada en SAP, marcamos error
                        res = false;
                        paintXY(dependency, i, XLColor.Red, "Este proyecto debe tener una dependencia asociada a la Unidad Org: "+ UOName );
                    }
                    
                }

            }
            valid = valid && res;
            if (!res) { 
                addError("Valor no valido", "Proyecto o proyectos no válidos en la columna: " + index, false); 
            }
                
            return res;
        }

        private bool verifyAcount(int dependency, int sheet=1) {
            bool res = true;
            string commnet = "";
            int index = 15;
            int tipoproy = 11;
            int badType = 0;
            int badAccount = 0;
            var connB1 = B1Connection.Instance();
            int branchId = Convert.ToInt16(segmentoOrigen);
            var branch = _context.Branch.FirstOrDefault(x => x.Id == branchId).Abr;
            var listParams = connB1.getProjects("*").Where(x => x.U_Sucursal == branch).Select(x => new { x.PrjCode, x.U_Tipo, x.U_UORGANIZA }).ToList();
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            var l = UsedRange.LastRow().RowNumber();


            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                //Si el proyecto existe en SAP ahí validamos los tipos de proyecto y cuentas asignadas
                
                if (listParams.Exists(x => string.Equals(x.PrjCode.ToString(), wb.Worksheet(sheet).Cell(i, index).Value.ToString(), StringComparison.OrdinalIgnoreCase)))
                {
                    var strproject = index != -1 ? wb.Worksheet(sheet).Cell(i, index).Value.ToString() : null;
                    var row = listParams.FirstOrDefault(x => x.PrjCode == strproject);
                    string UO = row.U_UORGANIZA.ToString();
                    var strdependency = dependency != -1 ? wb.Worksheet(sheet).Cell(i, dependency).Value.ToString() : null;
                    var dep = _context.Dependencies.Where(x => x.BranchesId == branchId).Include(x => x.OrganizationalUnit).FirstOrDefault(x => x.Cod == strdependency);
                    //Validamos aunque la UO no iguale, esto con el objetivo de que marque los errores
                    //CAP se deja pasar
                    if (wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString() != "CAP" && row.U_UORGANIZA!=dep.OrganizationalUnit.Cod)
                    {
                        //-----------------------------Validacion del tipo--------------------------------
                        var projectAccount = wb.Worksheet(sheet).Cell(i, tipoproy).Value.ToString();
                        var projectType = listParams.Where(x => x.PrjCode == wb.Worksheet(sheet).Cell(i, index).Value.ToString()).FirstOrDefault().U_Tipo.ToString();
                        string tipo = projectType;//variable auxiliar, no puede usarse la de arriba en EF por ser dinámica
                        var typeExists = _context.TableOfTableses.ToList().Exists(x => string.Equals(x.Type, tipo, StringComparison.OrdinalIgnoreCase));
                        if (!typeExists)
                        {
                            //si el proyecto tiene un U_Tipo en SAP y no lo tenemos en personas, entonces no es válido. Pasa para los proyectos con U_Tipo 'I', ESOS NUNCA ENTRAN EN PLANILLAS, este If es una formalidad
                            commnet = "El tipo de proyecto no es válido";
                            paintXY(index, i, XLColor.Red, commnet);
                            badType++;
                            res = false;
                        }
                        else
                        {
                            //-----------------------------Validacion de la cuenta asignada--------------------------------
                            var assignedAccount = _context.TableOfTableses.Where(x => x.Type == tipo).Where(x => x.Id >= 29 && x.Id <= 33).FirstOrDefault().Value.ToString();//la cuenta asignada a ese tipo de proyecto
                            //si el tipo existe en nuestra BD entonces verificar que la cuenta corresponda a ese tipo
                            if (projectAccount != assignedAccount)
                            {
                                commnet = "La cuenta asignada es incorrecta, debería ser: " + assignedAccount;
                                paintXY(tipoproy, i, XLColor.Red, commnet);
                                badAccount++;
                                res = false;
                            }
                        }
                    }
                }
            }
            valid = valid && res;
            if (!res && badAccount>0 && badType>0) { addError("Valor no valido", "Tipos de proyectos no válidos en la columna: "+index+ " y cuentas asignadas no válidas en la columna: " + tipoproy , false); }
            else if (!res && badAccount > 0 && badType == 0) { addError("Valor no valido", "Cuentas asignadas no válidas en la columna: " + tipoproy, false); }
            else if (!res && badAccount == 0 && badType > 0) { addError("Valor no valido", "Tipos de proyectos no válidos en la columna: " + index, false); }
            return res;

        }

        private bool verifyDates(int dependency, int sheet = 1)
        {
            string commnet;//especifica el error
            var connB1 = B1Connection.Instance();
            int branchId = Convert.ToInt16(segmentoOrigen);
            var branch = _context.Branch.FirstOrDefault(x => x.Id == branchId).Abr;
            var list = connB1.getProjects("*").Where(x => x.U_Sucursal == branch).Select(x => new { x.PrjCode, x.U_Tipo, x.ValidFrom, x.ValidTo, x.U_UORGANIZA }).ToList();
            int index = 15;
            bool res = true;
            IXLRange UsedRange = wb.Worksheet(sheet).RangeUsed();
            var l = UsedRange.LastRow().RowNumber();


            for (int i = headerin + 1; i <= UsedRange.LastRow().RowNumber(); i++)
            {
               
                //Si el proyecto existe en SAP ahí validamos fechas
                if (list.Exists(x => string.Equals(x.PrjCode.ToString(), wb.Worksheet(sheet).Cell(i, index).Value.ToString(), StringComparison.OrdinalIgnoreCase)))
                {
                    var strproject = index != -1 ? wb.Worksheet(sheet).Cell(i, index).Value.ToString() : null;
                    //-----------------------------Validaciones de la fecha del proyecto--------------------------------
                    var projectInitialDate = list.Where(x => x.PrjCode == strproject).FirstOrDefault().ValidFrom.ToString();
                    DateTime parsedIni = Convert.ToDateTime(projectInitialDate);
                    var projectFinalDate = list.Where(x => x.PrjCode == strproject).FirstOrDefault().ValidTo.ToString();
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
            valid = valid && res;
            if (!res) { addError("Valor no valido", "Proyecto/s con fechas no válidas en la columna:" +index, false); }

            return res;
        }

        public Dist_Posgrado ToDistDiscounts(int row, int sheet = 1)
        {
            Dist_Posgrado dis = new Dist_Posgrado();
            dis.Id = Dist_Posgrado.GetNextId(_context);
            dis.Document = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            dis.Names = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            dis.FirstSurName = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            dis.SecondSurName = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
            dis.MariedSurName = wb.Worksheet(sheet).Cell(row, 5).Value.ToString();
            dis.ProjectName = wb.Worksheet(sheet).Cell(row, 6).Value.ToString();
            dis.Vesion = (int) strToDecimal(row, 7);
            dis.TotalPagado = strToDecimal(row,8);
            dis.Dependency = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();
            dis.CUNI = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();
            dis.ProjectType = wb.Worksheet(sheet).Cell(row, 11).Value.ToString();
            dis.TipoTarea = wb.Worksheet(sheet).Cell(row, 12).Value.ToString();
            dis.PEI = wb.Worksheet(sheet).Cell(row, 13).Value.ToString();
            dis.Periodo = wb.Worksheet(sheet).Cell(row, 14).Value.ToString();
            dis.ProjectCode = wb.Worksheet(sheet).Cell(row, 15).Value.ToString();

            dis.Porcentaje = 0m;

            dis.mes = this.mes;
            dis.gestion = this.gestion;
            dis.segmentoOrigen = this.segmentoOrigen;

            dis.DistFileId = file.Id;
            return dis;
        }
    }
}