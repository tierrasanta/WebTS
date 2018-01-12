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

    public class LOTESIndexViewModel
    {
        public List<LOTES> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class LotesController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();


        // GET: Lotes
        public ActionResult Index(int? page, String Search)
        {
            //
            var viewModel = new LOTESIndexViewModel();

            //if (Search == null || Search.Equals(""))
            //{
            var pager = new Pager(db.LOTES.Count(), page);
            viewModel.Items = db.LOTES.Include(l => l.FINCAS)
                    .OrderBy(c => c.idlote)
                    .Skip((pager.CurrentPage - 1) * pager.PageSize)
                    .Take(pager.PageSize).ToList();
            viewModel.Pager = pager;
            //        }
            //        else
            //        {
            //var pager = new Pager(db.LOTES.Where(c => c.AgregarVariableAbuscar.Contains(Search)).Count(), page);
            //            viewModel.Items = db.LOTES.Include(l => l.FINCAS).Where(c => c.AgregarVariableAbuscar.Contains(Search))
            //                    .OrderBy(c => c.LOTESID)
            //                    .Skip((pager.CurrentPage - 1) * pager.PageSize)
            //                    .Take(pager.PageSize).ToList();
            //viewModel.Pager = pager;
            //@ViewBag.Search = Search;
            //        }
            return View(viewModel);
        }



        // GET: Lotes/Details/5
        public ActionResult Details(string id, string id2, string id3)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            LOTES lOTES = db.LOTES.Find(id, id2, id3);
            if (lOTES == null)
            {
                return HttpNotFound();
            }
            return View(lOTES);
        }

        // GET: Lotes/Create
        public ActionResult Create()
        {
            ViewBag.idfinca = new SelectList(db.FINCAS, "idfinca", "descripcion");
            return View();
        }

        // POST: Lotes/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idfinca,idlote,idusuario,descripcion,area,fechacreacion,fechacambio")] LOTES lOTES)
        {
            lOTES.idempresa = "01";
            List<LOTES> l = db.LOTES.ToList();
            if (l.Count == 0) { lOTES.idlote = "0001"; }
            else { lOTES.idlote = getidlote(Convert.ToInt32(l.Last().idlote)); }
            lOTES.idusuario = "0001";
            if (ModelState.IsValid)
            {
                //lOTES.Creado = DateTime.Now;
                //lOTES.Modificado = DateTime.Now;
                db.LOTES.Add(lOTES);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.idempresa = new SelectList(db.FINCAS, "idempresa", "idusuario", lOTES.idempresa);
            return View(lOTES);
        }

        private string getidlote(int idlote)
        {
            idlote = idlote + 1;
            int digitos = Convert.ToString(idlote).Length;
            if (digitos == 1) { return "000" + idlote; }
            else if (digitos == 2) { return "00" + idlote; }
            else if (digitos == 3) { return "0" + idlote; }
            else { return Convert.ToString(idlote); }
        }

        // GET: Lotes/Edit/5
        public ActionResult Edit(string id, string id2, string id3)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            LOTES lOTES = db.LOTES.Find(id, id2, id3);
            if (lOTES == null)
            {
                return HttpNotFound();
            }
            ViewBag.idempresa = new SelectList(db.FINCAS, "idempresa", "idusuario", lOTES.idempresa);
            return View(lOTES);
        }

        // POST: Lotes/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idfinca,idlote,idusuario,descripcion,area,fechacreacion,fechacambio")] LOTES lOTES)
        {
            lOTES.fechacambio = DateTime.Now;
            if (ModelState.IsValid)
            {
                db.Entry(lOTES).State = EntityState.Modified;
                //lOTES.Modificado = DateTime.Now;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.idempresa = new SelectList(db.FINCAS, "idempresa", "idusuario", lOTES.idempresa);
            return View(lOTES);
        }

        // GET: Lotes/Delete/5
        public ActionResult Delete(string id, string id2, string id3)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            LOTES lOTES = db.LOTES.Find(id, id2, id3);
            if (lOTES == null)
            {
                return HttpNotFound();
            }
            return View(lOTES);
        }

        // POST: Lotes/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id, string id2, string id3)
        {
            LOTES lOTES = db.LOTES.Find(id, id2, id3);
            db.LOTES.Remove(lOTES);
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
                //.Include(l => l.FINCAS)
                List<LOTES> list = db.LOTES.Include(l => l.FINCAS).ToList();
                int pos = 4;
                ws.Cells[pos, 5].Value = "idusuario";
                ws.Cells[pos, 6].Value = "descripcion";
                ws.Cells[pos, 7].Value = "area";
                ws.Cells[pos, 8].Value = "fechacreacion";
                ws.Cells[pos, 9].Value = "fechacambio";
                ws.Cells[pos, 10].Value = "FINCAS";

                foreach (var item in list)
                {
                    pos++;
                    ws.Cells[pos, 5].Value = item.idusuario == null ? "" : item.idusuario.ToString();
                    ws.Cells[pos, 6].Value = item.descripcion == null ? "" : item.descripcion.ToString();
                    ws.Cells[pos, 7].Value = item.area == null ? "" : item.area.ToString();
                    ws.Cells[pos, 8].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();
                    ws.Cells[pos, 9].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();
                    ws.Cells[pos, 10].Value = item.FINCAS == null ? "" : item.FINCAS.ToString();
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


            var table = new PdfPTable(6);

            var boldTableFont = FontFactory.GetFont("Arial", 10, Font.BOLD);
            var bodyFont = FontFactory.GetFont("Arial", 10, Font.NORMAL);

            table.AddCell(new Phrase("idusuario", boldTableFont));
            table.AddCell(new Phrase("descripcion", boldTableFont));
            table.AddCell(new Phrase("area", boldTableFont));
            table.AddCell(new Phrase("fechacreacion", boldTableFont));
            table.AddCell(new Phrase("fechacambio", boldTableFont));
            table.AddCell(new Phrase("FINCAS", boldTableFont));

            //
            List<LOTES> list = db.LOTES.Include(l => l.FINCAS).ToList();

            foreach (var item in list)
            {

                table.AddCell(new Phrase(item.idusuario == null ? "" : item.idusuario.ToString(), bodyFont));
                table.AddCell(new Phrase(item.descripcion == null ? "" : item.descripcion.ToString(), bodyFont));
                table.AddCell(new Phrase(item.area == null ? "" : item.area.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechacreacion == null ? "" : item.fechacreacion.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));
                table.AddCell(new Phrase(item.FINCAS == null ? "" : item.FINCAS.ToString(), bodyFont));
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
