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

    public class PlantillaCultivoDetalleIndexViewModel
    {
        public List<PlantillaCultivoDetalle> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class PlantillaCultivoDetallesController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();


        // GET: PlantillaCultivoDetalles
        public ActionResult Index(int? page, String Search)
        {
            //
            var viewModel = new PlantillaCultivoDetalleIndexViewModel();

            var pager = new Pager(db.PlantillaCultivoDetalle.Count(), page);
            viewModel.Items = db.PlantillaCultivoDetalle.Include(p => p.PlantillaCultivoCabecera).Include(p => p.TablaActividades)
                    .OrderBy(c => c.idplantilla)
                    .Skip((pager.CurrentPage - 1) * pager.PageSize)
                    .Take(pager.PageSize).ToList();
            viewModel.Pager = pager;

            return View(viewModel);
        }



        // GET: PlantillaCultivoDetalles/Details/5
        public ActionResult Details(string id, int id2, string id3)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PlantillaCultivoDetalle plantillaCultivoDetalle = db.PlantillaCultivoDetalle.Find(id, id2, id3);
            if (plantillaCultivoDetalle == null)
            {
                return HttpNotFound();
            }
            return View(plantillaCultivoDetalle);
        }

        // GET: PlantillaCultivoDetalles/Create
        public ActionResult Create()
        {
            ViewBag.idplantilla = new SelectList(db.PlantillaCultivoCabecera, "idplantilla", "descripcion");
            ViewBag.idactividad = new SelectList(db.TablaActividades, "idactividad", "descripcion");
            return View();
        }

        // POST: PlantillaCultivoDetalles/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idplantilla,idusuario,idactividad,cantidad,fechacreacion,fechacambio")] PlantillaCultivoDetalle plantillaCultivoDetalle)
        {
            plantillaCultivoDetalle.idempresa = "01";
            plantillaCultivoDetalle.idusuario = "0001";
            if (ModelState.IsValid)
            {
                db.PlantillaCultivoDetalle.Add(plantillaCultivoDetalle);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            //ViewBag.idempresa = new SelectList(db.PlantillaCultivoCabecera, "idempresa", "descripcion", plantillaCultivoDetalle.idempresa);
            //ViewBag.idempresa = new SelectList(db.TablaActividades, "idempresa", "idusuario", plantillaCultivoDetalle.idempresa);
            return View(plantillaCultivoDetalle);
        }

        // GET: PlantillaCultivoDetalles/Edit/5
        public ActionResult Edit(string id, int id2, string id3)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PlantillaCultivoDetalle plantillaCultivoDetalle = db.PlantillaCultivoDetalle.Find(id, id2, id3);
            if (plantillaCultivoDetalle == null)
            {
                return HttpNotFound();
            }
            ViewBag.idactividad = new SelectList(db.TablaActividades, "idactividad", "descripcion", plantillaCultivoDetalle.idempresa);
            return View(plantillaCultivoDetalle);
        }

        // POST: PlantillaCultivoDetalles/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idplantilla,idusuario,idactividad,cantidad,fechacreacion,fechacambio")] PlantillaCultivoDetalle plantillaCultivoDetalle)
        {
            plantillaCultivoDetalle.fechacambio = DateTime.Now;
            if (ModelState.IsValid)
            {
                db.Entry(plantillaCultivoDetalle).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.idempresa = new SelectList(db.PlantillaCultivoCabecera, "idempresa", "descripcion", plantillaCultivoDetalle.idempresa);
            ViewBag.idempresa = new SelectList(db.TablaActividades, "idempresa", "idusuario", plantillaCultivoDetalle.idempresa);
            return View(plantillaCultivoDetalle);
        }

        // GET: PlantillaCultivoDetalles/Delete/5
        public ActionResult Delete(string id, int id2, string id3)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PlantillaCultivoDetalle plantillaCultivoDetalle = db.PlantillaCultivoDetalle.Find(id, id2, id3);
            if (plantillaCultivoDetalle == null)
            {
                return HttpNotFound();
            }
            return View(plantillaCultivoDetalle);
        }

        // POST: PlantillaCultivoDetalles/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id, int id2, string id3)
        {
            PlantillaCultivoDetalle plantillaCultivoDetalle = db.PlantillaCultivoDetalle.Find(id, id2, id3);
            db.PlantillaCultivoDetalle.Remove(plantillaCultivoDetalle);
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
                //.Include(p => p.PlantillaCultivoCabecera).Include(p => p.TablaActividades)
                List<PlantillaCultivoDetalle> list = db.PlantillaCultivoDetalle.Include(p => p.PlantillaCultivoCabecera).Include(p => p.TablaActividades).ToList();
                int pos = 4;
                ws.Cells[pos, 4].Value = "Plantilla";
                ws.Cells[pos, 6].Value = "Actividad";
                ws.Cells[pos, 7].Value = "Cantidad";
                ws.Cells[pos, 8].Value = "Fecha de creacion";
                ws.Cells[pos, 9].Value = "Fecha de cambio";

                foreach (var item in list)
                {
                    pos++;
                    ws.Cells[pos, 4].Value = item.PlantillaCultivoCabecera.descripcion == null ? "" : item.PlantillaCultivoCabecera.descripcion.ToString();
                    ws.Cells[pos, 6].Value = item.TablaActividades.descripcion == null ? "" : item.TablaActividades.descripcion.ToString();
                    ws.Cells[pos, 7].Value = item.cantidad == null ? "" : item.cantidad.ToString();
                    ws.Cells[pos, 8].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();
                    ws.Cells[pos, 9].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();
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

            table.AddCell(new Phrase("Plantilla", boldTableFont));
            table.AddCell(new Phrase("Actividad", boldTableFont));
            table.AddCell(new Phrase("cantidad", boldTableFont));
            table.AddCell(new Phrase("Fecha de creacion", boldTableFont));
            table.AddCell(new Phrase("Fecha de cambio", boldTableFont));
            //
            List<PlantillaCultivoDetalle> list = db.PlantillaCultivoDetalle.Include(p => p.PlantillaCultivoCabecera).Include(p => p.TablaActividades).ToList();

            foreach (var item in list)
            {

                table.AddCell(new Phrase(item.PlantillaCultivoCabecera.descripcion == null ? "" : item.PlantillaCultivoCabecera.descripcion.ToString(), bodyFont));
                table.AddCell(new Phrase(item.TablaActividades.descripcion == null ? "" : item.TablaActividades.descripcion.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechacreacion == null ? "" : item.fechacreacion.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));                
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
