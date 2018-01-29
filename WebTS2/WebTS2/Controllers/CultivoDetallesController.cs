using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using WebTS2.Models;
using WebTS2.Helper;
using OfficeOpenXml;
using iTextSharp.text;
using System.IO;
using iTextSharp.text.pdf;

namespace WebTS2.Controllers
{

    //public class CultivoDetalleIndexViewModel
    //{
    //    public List<CultivoDetalle> Items { get; set; }
    //    public Pager Pager { get; set; }
    //}

    public class CultivoDetallesController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();


        // GET: CultivoDetalles
        //public ActionResult Index(int? page, String Search)
        //{
        //    //
        //    var viewModel = new CultivoDetalleIndexViewModel();
        //    var pager = new Pager(db.CultivoDetalle.Count(), page);
        //    viewModel.Items = db.CultivoDetalle.Include(c => c.Cultivo).Include(c => c.TablaActividades)
        //            .OrderBy(c => c.idcultivodetalle)
        //            .Skip((pager.CurrentPage - 1) * pager.PageSize)
        //            .Take(pager.PageSize).ToList();
        //    viewModel.Pager = pager;

        //    return View(viewModel);
        //}



        // GET: CultivoDetalles/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            CultivoDetalle cultivoDetalle = db.CultivoDetalle.Find(id);
            if (cultivoDetalle == null)
            {
                return HttpNotFound();
            }
            return View(cultivoDetalle);
        }

        // GET: CultivoDetalles/Create
        public ActionResult Create(int idcultivo, int idactividades, int idplantilla)
        {
            CultivoDetalle cultivodetalle = new CultivoDetalle();
            cultivodetalle.idcultivo = idcultivo;
            ViewBag.idreturn = idcultivo;
            ViewBag.idcultivo = new SelectList(db.Cultivo, "idcultivo", db.Cultivo);
            List<SelectListItem> list = new List<SelectListItem>();
            var culdetalle = db.PlantillaCultivoDetalle.Include(c => c.TablaActividades).Where(t => t.TablaActividades.idparent == idactividades && t.TablaActividades.abreviatura != "" && t.idplantilla==idplantilla).OrderBy(t => t.TablaActividades.descripcion);
            foreach (var det in culdetalle)
            {
                list.Add(new SelectListItem() { Text = det.TablaActividades.descripcion, Value = det.idactividad.ToString() });
            }
            ViewBag.idactividad = new SelectList(list, "Value", "Text");
            //ViewBag.nombreActividad = 
            //ViewBag.idactividad = new SelectList(db.TablaActividades.Where(t => t.idparent == idactividades && t.abreviatura != ""), "idactividades", "descripcion");

            return View(cultivodetalle);
        }

        // POST: CultivoDetalles/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idusuario,idcultivo,idcultivodetalle,idactividad,cantidad,fechallenado,fechacreacion,fechacambio")] CultivoDetalle cultivoDetalle)
        {
            //Cultivo cul = new Cultivo();
            List<Cultivo> cult = db.Cultivo.Where(c => c.idcultivo == cultivoDetalle.idcultivo).ToList();
            var culti = db.TablaActividades.Where(a => a.idactividades == cultivoDetalle.idactividad).ToList();
            //int parent = Convert.ToInt32(cultivoDetalle.TablaActividades.idparent);
            cultivoDetalle.idempresa = "01";
            cultivoDetalle.idusuario = "0001";
            if (ModelState.IsValid)
            {

                db.CultivoDetalle.Add(cultivoDetalle);
                db.SaveChanges();
                return RedirectToAction("details", "Cultivos", new { @id = cultivoDetalle.idcultivo });
            }
            //cultivoDetalle = Create(cultivoDetalle.idcultivo,Convert.ToInt32( culti[0].idparent),cult[0].idplantilla);
            
            return Create(cultivoDetalle.idcultivo, Convert.ToInt32(culti[0].idparent), cult[0].idplantilla);
        }

        // GET: CultivoDetalles/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            CultivoDetalle cultivoDetalle = db.CultivoDetalle.Find(id);
            if (cultivoDetalle == null)
            {
                return HttpNotFound();
            }
            ViewBag.idreturn = cultivoDetalle.idcultivo;
            ViewBag.idcultivo = new SelectList(db.Cultivo, "idcultivo", "idempresa", cultivoDetalle.idcultivo);
            ViewBag.idactividad = new SelectList(db.TablaActividades, "idactividades", "idempresa", cultivoDetalle.idactividad);
            return View(cultivoDetalle);
        }

        // POST: CultivoDetalles/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idusuario,idcultivo,idcultivodetalle,idactividad,cantidad,fechallenado,fechacreacion,fechacambio")] CultivoDetalle cultivoDetalle)
        {
            cultivoDetalle.fechacambio = DateTime.Now;
            if (ModelState.IsValid)
            {
                db.Entry(cultivoDetalle).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("details", "Cultivos", new { @id = cultivoDetalle.idcultivo });
            }
            ViewBag.idcultivo = new SelectList(db.Cultivo, "idcultivo", "idempresa", cultivoDetalle.idcultivo);
            ViewBag.idactividad = new SelectList(db.TablaActividades, "idactividades", "idempresa", cultivoDetalle.idactividad);
            return Edit(cultivoDetalle.idcultivodetalle);
        }

        // GET: CultivoDetalles/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            CultivoDetalle cultivoDetalle = db.CultivoDetalle.Find(id);
            if (cultivoDetalle == null)
            {
                return HttpNotFound();
            }
            return View(cultivoDetalle);
        }

        // POST: CultivoDetalles/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            CultivoDetalle cultivoDetalle = db.CultivoDetalle.Find(id);
            int idcultivo = cultivoDetalle.idcultivo;
            db.CultivoDetalle.Remove(cultivoDetalle);
            db.SaveChanges();
            return RedirectToAction("details", "Cultivos", new { @id = idcultivo });
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
                //.Include(c => c.Cultivo).Include(c => c.TablaActividades)
                List<CultivoDetalle> list = db.CultivoDetalle.Include(c => c.Cultivo).Include(c => c.TablaActividades).ToList();
                int pos = 4;
                ws.Cells[pos, 2].Value = "idempresa";
                ws.Cells[pos, 3].Value = "idusuario";
                ws.Cells[pos, 7].Value = "cantidad";
                ws.Cells[pos, 8].Value = "fechallenado";
                ws.Cells[pos, 9].Value = "fechacreacion";
                ws.Cells[pos, 10].Value = "fechacambio";
                ws.Cells[pos, 11].Value = "Cultivo";
                ws.Cells[pos, 12].Value = "TablaActividades";

                foreach (var item in list)
                {
                    pos++;
                    ws.Cells[pos, 2].Value = item.idempresa == null ? "" : item.idempresa.ToString();
                    ws.Cells[pos, 3].Value = item.idusuario == null ? "" : item.idusuario.ToString();
                    ws.Cells[pos, 7].Value = item.cantidad == null ? "" : item.cantidad.ToString();
                    ws.Cells[pos, 8].Value = item.fechallenado == null ? "" : item.fechallenado.ToString();
                    ws.Cells[pos, 9].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();
                    ws.Cells[pos, 10].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();
                    ws.Cells[pos, 11].Value = item.Cultivo == null ? "" : item.Cultivo.ToString();
                    ws.Cells[pos, 12].Value = item.TablaActividades == null ? "" : item.TablaActividades.ToString();
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


            var table = new PdfPTable(8);

            var boldTableFont = FontFactory.GetFont("Arial", 10, Font.BOLD);
            var bodyFont = FontFactory.GetFont("Arial", 10, Font.NORMAL);

            table.AddCell(new Phrase("idempresa", boldTableFont));
            table.AddCell(new Phrase("idusuario", boldTableFont));
            table.AddCell(new Phrase("cantidad", boldTableFont));
            table.AddCell(new Phrase("fechallenado", boldTableFont));
            table.AddCell(new Phrase("fechacreacion", boldTableFont));
            table.AddCell(new Phrase("fechacambio", boldTableFont));
            table.AddCell(new Phrase("Cultivo", boldTableFont));
            table.AddCell(new Phrase("TablaActividades", boldTableFont));

            //
            List<CultivoDetalle> list = db.CultivoDetalle.Include(c => c.Cultivo).Include(c => c.TablaActividades).ToList();

            foreach (var item in list)
            {

                table.AddCell(new Phrase(item.idempresa == null ? "" : item.idempresa.ToString(), bodyFont));
                table.AddCell(new Phrase(item.idusuario == null ? "" : item.idusuario.ToString(), bodyFont));
                table.AddCell(new Phrase(item.cantidad == null ? "" : item.cantidad.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechallenado == null ? "" : item.fechallenado.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechacreacion == null ? "" : item.fechacreacion.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));
                table.AddCell(new Phrase(item.Cultivo == null ? "" : item.Cultivo.ToString(), bodyFont));
                table.AddCell(new Phrase(item.TablaActividades == null ? "" : item.TablaActividades.ToString(), bodyFont));
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
