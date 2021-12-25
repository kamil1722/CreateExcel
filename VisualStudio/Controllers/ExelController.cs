using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using MyExel.Models;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace MyExel.Controllers
{
    public class ExelController : Controller
    {
        private DBCtx Context { get; }
        public ExelController(DBCtx _context)
        {
            this.Context = _context;
        }
        public static DataTable ColumnsExel()
        {
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[4] {
                                        new DataColumn("Code"),
                                        new DataColumn("NameSMO"),
                                        new DataColumn("NameMO"),
                                        new DataColumn("NameMKB")
                                        });
            return dt;
        }
        public IList<ExelViewModel> GetHospitalList()
        {
            var valuesExelList = (from smo in this.Context.smo
                                  join med in this.Context.med on smo.code equals med.cont
                                  join mo in this.Context.mo on med.mo equals mo.code
                                  join mkb in this.Context.mkb on med.ds1 equals mkb.code
                                  orderby med.cont, mo.name
                                  select new ExelViewModel
                                  {
                                      Code = smo.code + "-" + mo.code + "-" + mkb.code,
                                      NameSMO = smo.name,
                                      NameMO = mo.name,
                                      NameMKB = mkb.name
                                  }).ToList();
            return valuesExelList;
        }
        public IXLWorksheet GroupWs(IXLWorksheet firstWs)
        {
            var newValues = GetHospitalList();
            DataTable newDt = ColumnsExel();

            //Сортировка по SMO
            for (var i = 2; i < newValues.Where(x => x.Code.StartsWith("5-")).Count() + 1; i++)
                firstWs.Row(i).OutlineLevel = 1;
            for (var i = 90; i < newValues.Where(x => x.Code.StartsWith("8-")).Count() + 89; i++)
                firstWs.Row(i).OutlineLevel = 1;
            for (var i = 119; i < newValues.Where(x => x.Code.Contains("10-")).Count() + 118; i++)
                firstWs.Row(i).OutlineLevel = 1;
            for (var i = 558; i < newValues.Where(x => x.Code.Contains("15-")).Count() + 557; i++)
                firstWs.Row(i).OutlineLevel = 1;
         
            return firstWs;
        }
        public IXLWorksheet InsertData(IXLWorksheet secondWs)
        {
            var newValues = GetHospitalList();
            //Запросы для получения сводной информации по подразделениям и организациям
            var MOCount = Context.med.GroupBy(x => x.cont).Select(x => new { Count = x.Select(z => z.mo).Distinct().Count() });
            var MKBCount = Context.med.GroupBy(x => x.cont).Select(x => new { Count = x.Select(z => z.ds1).Count() }); //для организаций      
            var MKBCount2 = Context.med.Join(Context.mo, x => x.mo, y => y.code, (x, y) => new { NameMO = y.name, CodeDS = x.ds1 })
                .GroupBy(x => x.NameMO).Select(x => new { Count = x.Select(z => z.CodeDS).Count() });//для подраздлений
            var MONames = Context.med.Join(Context.mo, x => x.mo, y => y.code, (x, y) => new { NameMO = y.name, CodeDS = x.ds1 });
            //внесения результатов запросов в ячейки
            secondWs.Cell(1, 1).SetValue("Организация:");
            secondWs.Cell(2, 1).SetValue(newValues.Select(x => x.NameSMO).Distinct());
            secondWs.Cell(1, 2).SetValue("Количество подразделений:");
            secondWs.Cell(2, 2).SetValue(MOCount.Select(x => x.Count));
            secondWs.Cell(1, 3).SetValue("Общее количество болезней:");
            secondWs.Cell(2, 3).SetValue(MKBCount.Select(x => x.Count));

            secondWs.Cell(7, 1).SetValue("Подразделение:");
            secondWs.Cell(8, 1).SetValue(MONames.Select(x => x.NameMO).Distinct());
            secondWs.Cell(7, 3).SetValue("Количество болезней:");
            secondWs.Cell(8, 3).SetValue(MKBCount2.Select(x => x.Count));

            return secondWs;
        }


        public ActionResult ExportToExcel()
        {
            DataTable newDt = ColumnsExel();
            var newValues = GetHospitalList();
            foreach (var insertExelList in newValues)
            {
                newDt.Rows.Add(insertExelList.Code, insertExelList.NameSMO, insertExelList.NameMO, insertExelList.NameMKB);
            }
            //test
            using (XLWorkbook workbook = new XLWorkbook())
            {
                var firstWs = workbook.AddWorksheet(newDt);
                var secontWs = workbook.AddWorksheet("Report");
                firstWs = GroupWs(firstWs);
                secontWs = InsertData(secontWs);
                firstWs.Columns("A:H").AdjustToContents();
                secontWs.Columns("A:H").AdjustToContents();
                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ImportData1.xlsx");
                }
            }
        }
    }
        
}
