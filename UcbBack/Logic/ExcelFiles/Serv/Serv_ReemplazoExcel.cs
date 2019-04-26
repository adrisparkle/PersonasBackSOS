﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using ClosedXML.Excel;
using UcbBack.Models;
using UcbBack.Models.Serv;

namespace UcbBack.Logic.ExcelFiles.Serv
{
    public class Serv_ReemplazoExcel : ValidateExcelFile
    {
        private static Excelcol[] cols = new[]
        {
            new Excelcol("Codigo Socio", typeof(string)),
            new Excelcol("Nombre Socio", typeof(string)),
            new Excelcol("Cod Dependencia", typeof(string)),
            new Excelcol("PEI PO", typeof(string)),
            new Excelcol("Nombre del Servicio", typeof(string)),
            new Excelcol("Periodo Academico", typeof(string)),
            new Excelcol("Sigla Asignatura", typeof(string)),
            new Excelcol("Paralelo", typeof(string)),
            new Excelcol("Codigo Paralelo SAP", typeof(string)),
            new Excelcol("Cuenta Asignada", typeof(string)),
            new Excelcol("Monto Contrato", typeof(double)),
            new Excelcol("Monto IUE", typeof(double)),
            new Excelcol("Monto IT", typeof(double)),
            new Excelcol("Monto a Pagar", typeof(double)),
            new Excelcol("Observaciones", typeof(string)),
        };

        private ApplicationDbContext _context;
        private ServProcess process;

        public Serv_ReemplazoExcel(string fileName, int headerin = 1)
            : base(cols, fileName, headerin)
        {
        }

        public Serv_ReemplazoExcel(Stream data, ApplicationDbContext context, string fileName, ServProcess process,
            int headerin = 1, int sheets = 1, string resultfileName = "Result")
            : base(cols, data, fileName, headerin, sheets, resultfileName, context)
        {
            this.process = process;
            _context = context;
            isFormatValid();
        }

        public override void toDataBase()
        {
            IXLRange UsedRange = wb.Worksheet(1).RangeUsed();

            for (int i = 1 + headerin; i <= UsedRange.LastRow().RowNumber(); i++)
            {
                _context.ServReemplazos.Add(ToServVarios(i));
            }

            _context.SaveChanges();
        }

        public Serv_Paralelo ToServVarios(int row, int sheet = 1)
        {
            Serv_Paralelo data = new Serv_Paralelo();
            data.Id = Serv_Paralelo.GetNextId(_context);

            data.CardCode = wb.Worksheet(sheet).Cell(row, 1).Value.ToString();
            data.CardName = wb.Worksheet(sheet).Cell(row, 2).Value.ToString();
            var cod = wb.Worksheet(sheet).Cell(row, 3).Value.ToString();
            var depId = _context.Dependencies
                .FirstOrDefault(x => x.Cod == cod);
            data.DependencyId = depId.Id;
            data.PEI = wb.Worksheet(sheet).Cell(row, 4).Value.ToString();
            data.ServiceName = wb.Worksheet(sheet).Cell(row, 5).Value.ToString();
            data.Periodo = wb.Worksheet(sheet).Cell(row, 6).Value.ToString();
            data.Sigla = wb.Worksheet(sheet).Cell(row, 7).Value.ToString();
            data.ParalelNumber = wb.Worksheet(sheet).Cell(row, 8).Value.ToString();
            data.ParalelSAP = wb.Worksheet(sheet).Cell(row, 9).Value.ToString();

            data.AssignedAccount = wb.Worksheet(sheet).Cell(row, 10).Value.ToString();
            data.ContractAmount = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 11).Value.ToString());
            data.IUE = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 12).Value.ToString());
            data.IT = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 13).Value.ToString());
            data.TotalAmount = Decimal.Parse(wb.Worksheet(sheet).Cell(row, 14).Value.ToString());
            data.Comments = wb.Worksheet(sheet).Cell(row, 15).Value.ToString();
            data.Serv_ProcessId = process.Id;
            return data;
        }

        public override bool ValidateFile()
        {
            return true;
        }
    }
}