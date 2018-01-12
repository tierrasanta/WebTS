using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using TierraSanta.Models;
using TierraSanta.Helper;
using OfficeOpenXml;
using iTextSharp.text;
using System.IO;
using iTextSharp.text.pdf;

namespace TierraSanta.Controllers
{

    public class FINCASIndexViewModel
    {
        public List<FINCAS> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class FincasController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();


        // GET: Fincas
        public ActionResult Index(int? page, String Search)
        {
            //
            var viewModel = new FINCASIndexViewModel();

            //if (Search == null || Search.Equals(""))
            //{
            var pager = new Pager(db.FINCAS.Count(), page);
            viewModel.Items = db.FINCAS
                    .OrderBy(c => c.idfinca)
                    .Skip((pager.CurrentPage - 1) * pager.PageSize)
                    .Take(pager.PageSize).ToList();
            viewModel.Pager = pager;
            //        }
            //        else
            //        {
            //var pager = new Pager(db.FINCAS.Where(c => c.AgregarVariableAbuscar.Contains(Search)).Count(), page);
            //            viewModel.Items = db.FINCAS.Where(c => c.AgregarVariableAbuscar.Contains(Search))
            //                    .OrderBy(c => c.FINCASID)
            //                    .Skip((pager.CurrentPage - 1) * pager.PageSize)
            //                    .Take(pager.PageSize).ToList();
            //viewModel.Pager = pager;
            //@ViewBag.Search = Search;
            //        }
            return View(viewModel);
        }



        // GET: Fincas/Details/5
        public ActionResult Details(string id, string id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            FINCAS fINCAS = db.FINCAS.Find(id, id2);
            if (fINCAS == null)
            {
                return HttpNotFound();
            }
            return View(fINCAS);
        }

        // GET: Fincas/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Fincas/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idfinca,idusuario,descripcion,fechacreacion,fechacambio")] FINCAS fINCAS)
        {
            fINCAS.idempresa = "01";
            List<FINCAS> f = db.FINCAS.ToList();
            if (f.Count == 0) { fINCAS.idfinca = "01"; }
            else { fINCAS.idfinca = getidfinca(Convert.ToInt32(f.Last().idfinca)); }
            fINCAS.idusuario = "0001";
            if (ModelState.IsValid)
            {
                //fINCAS.Creado = DateTime.Now;
                //fINCAS.Modificado = DateTime.Now;
                db.FINCAS.Add(fINCAS);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(fINCAS);
        }

        private string getidfinca(int idfinca)
        {
            idfinca = idfinca + 1;
            int digitos = Convert.ToString(idfinca).Length;
            if (digitos == 1) { return "0" + idfinca; }
            else { return Convert.ToString(idfinca); }
        }

        // GET: Fincas/Edit/5
        public ActionResult Edit(string id, string id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            FINCAS fINCAS = db.FINCAS.Find(id, id2);
            if (fINCAS == null)
            {
                return HttpNotFound();
            }
            return View(fINCAS);
        }

        // POST: Fincas/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idfinca,idusuario,descripcion,fechacreacion,fechacambio")] FINCAS fINCAS)
        {
            fINCAS.fechacambio = DateTime.Now;
            if (ModelState.IsValid)
            {
                db.Entry(fINCAS).State = EntityState.Modified;
                //fINCAS.Modificado = DateTime.Now;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(fINCAS);
        }

        // GET: Fincas/Delete/5
        public ActionResult Delete(string id, string id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            FINCAS fINCAS = db.FINCAS.Find(id, id2);
            if (fINCAS == null)
            {
                return HttpNotFound();
            }
            return View(fINCAS);
        }

        // POST: Fincas/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id, string id2)
        {
            FINCAS fINCAS = db.FINCAS.Find(id, id2);
            db.FINCAS.Remove(fINCAS);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        public ActionResult ReportExcel()
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Report");
                ExcelWorksheet ws = package.Workbook.Worksheets[1];
                ws.Name = "Report";
                ws.Cells.Style.Font.Size = 11;
                ws.Cells.Style.Font.Name = "Calibri";
                //
                List<FINCAS> list = db.FINCAS.ToList();
                int pos = 4;
                ws.Cells[pos, 4].Value = "idusuario";
                ws.Cells[pos, 5].Value = "descripcion";
                ws.Cells[pos, 6].Value = "fechacreacion";
                ws.Cells[pos, 7].Value = "fechacambio";
                ws.Cells[pos, 8].Value = "LOTES";

                foreach (var item in list)
                {
                    pos++;
                    ws.Cells[pos, 4].Value = item.idusuario == null ? "" : item.idusuario.ToString();
                    ws.Cells[pos, 5].Value = item.descripcion == null ? "" : item.descripcion.ToString();
                    ws.Cells[pos, 6].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();
                    ws.Cells[pos, 7].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();
                    ws.Cells[pos, 8].Value = item.LOTES == null ? "" : item.LOTES.ToString();
                }
                ws.Cells["B3:F" + pos].AutoFitColumns();


                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=Report.xlsx");
                Response.BinaryWrite(package.GetAsByteArray());
                Response.Flush();
                Response.Close();
            }

            return null;
        }

        public ActionResult ReportPDF()
        {
            var document = new Document(PageSize.A4, 50, 50, 25, 25);
            var output = new MemoryStream();
            var writer = PdfWriter.GetInstance(document, output);
            document.Open();


            var table = new PdfPTable(5);

            var boldTableFont = FontFactory.GetFont("Arial", 10, Font.BOLD);
            var bodyFont = FontFactory.GetFont("Arial", 10, Font.NORMAL);

            table.AddCell(new Phrase("idusuario", boldTableFont));
            table.AddCell(new Phrase("descripcion", boldTableFont));
            table.AddCell(new Phrase("fechacreacion", boldTableFont));
            table.AddCell(new Phrase("fechacambio", boldTableFont));
            table.AddCell(new Phrase("LOTES", boldTableFont));

            //
            List<FINCAS> list = db.FINCAS.ToList();

            foreach (var item in list)
            {

                table.AddCell(new Phrase(item.idusuario == null ? "" : item.idusuario.ToString(), bodyFont));
                table.AddCell(new Phrase(item.descripcion == null ? "" : item.descripcion.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechacreacion == null ? "" : item.fechacreacion.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));
                table.AddCell(new Phrase(item.LOTES == null ? "" : item.LOTES.ToString(), bodyFont));
            }



            document.Add(table);
            document.Close();

            Response.Clear();
            Response.ContentType = "application/pdf";
            Response.AddHeader("Content-Disposition", "attachment; filename=Report.pdf");
            Response.BinaryWrite(output.ToArray());
            Response.Flush();
            Response.Close();

            return null;
        }
    }
}
