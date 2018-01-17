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

	public class ActividadesFertilizanteIndexViewModel
    {
		public List<TablaActividades> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class ActividadesFertilizanteController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();

		
        // GET: ActividadesFertilizante
        public ActionResult Index(int? page, String Search)
        {
		//
			var viewModel = new ActividadesFertilizanteIndexViewModel();            

            var pager = new Pager(db.TablaActividades.Count(), page);
                viewModel.Items = db.TablaActividades.Include(t => t.TablaCultivos)
                        .OrderBy(c => c.idactividad)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;
            
            return View(viewModel);
        }



        // GET: ActividadesFertilizante/Details/5
        public ActionResult Details(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TablaActividades tablaActividades = db.TablaActividades.Find(id);
            if (tablaActividades == null)
            {
                return HttpNotFound();
            }
            return View(tablaActividades);
        }

        // GET: ActividadesFertilizante/Create
        public ActionResult Create()
        {
            ViewBag.idempresa = new SelectList(db.TablaCultivos, "idempresa", "idusuario");
            return View();
        }

        // POST: ActividadesFertilizante/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idactividad,idusuario,descripcion,abreviatura,unimedida,costo1,fechacreacion,fechacambio")] TablaActividades tablaActividades)
        {
            if (ModelState.IsValid)
            {               
                db.TablaActividades.Add(tablaActividades);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.idempresa = new SelectList(db.TablaCultivos, "idempresa", "idusuario", tablaActividades.idempresa);
            return View(tablaActividades);
        }

        // GET: ActividadesFertilizante/Edit/5
        public ActionResult Edit(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TablaActividades tablaActividades = db.TablaActividades.Find(id);
            if (tablaActividades == null)
            {
                return HttpNotFound();
            }
            ViewBag.idempresa = new SelectList(db.TablaCultivos, "idempresa", "idusuario", tablaActividades.idempresa);
            return View(tablaActividades);
        }

        // POST: ActividadesFertilizante/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idactividad,idusuario,descripcion,abreviatura,unimedida,costo1,fechacreacion,fechacambio")] TablaActividades tablaActividades)
        {
            if (ModelState.IsValid)
            {
                db.Entry(tablaActividades).State = EntityState.Modified;				
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.idempresa = new SelectList(db.TablaCultivos, "idempresa", "idusuario", tablaActividades.idempresa);
            return View(tablaActividades);
        }

        // GET: ActividadesFertilizante/Delete/5
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TablaActividades tablaActividades = db.TablaActividades.Find(id);
            if (tablaActividades == null)
            {
                return HttpNotFound();
            }
            return View(tablaActividades);
        }

        // POST: ActividadesFertilizante/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            TablaActividades tablaActividades = db.TablaActividades.Find(id);
            db.TablaActividades.Remove(tablaActividades);
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
				//.Include(t => t.TablaCultivos)
                List<TablaActividades> list = db.TablaActividades.Include(t => t.TablaCultivos).ToList();
                int pos = 4;
								ws.Cells[pos, 4].Value = "idusuario";
									ws.Cells[pos, 5].Value = "descripcion";
									ws.Cells[pos, 6].Value = "abreviatura";
									ws.Cells[pos, 8].Value = "costo1";
									ws.Cells[pos, 9].Value = "fechacreacion";
									ws.Cells[pos, 10].Value = "fechacambio";
									ws.Cells[pos, 11].Value = "TablaCultivos";
					
                foreach (var item in list)
                {
                    pos++;
								ws.Cells[pos, 4].Value = item.idusuario == null ? "" : item.idusuario.ToString();				
									ws.Cells[pos, 5].Value = item.descripcion == null ? "" : item.descripcion.ToString();				
									ws.Cells[pos, 6].Value = item.abreviatura == null ? "" : item.abreviatura.ToString();				
									ws.Cells[pos, 8].Value = item.costo1 == null ? "" : item.costo1.ToString();				
									ws.Cells[pos, 9].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();				
									ws.Cells[pos, 10].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();				
									ws.Cells[pos, 11].Value = item.TablaCultivos == null ? "" : item.TablaCultivos.ToString();				
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

			
            var table = new PdfPTable(7);

            var boldTableFont = FontFactory.GetFont("Arial", 10, Font.BOLD);
            var bodyFont = FontFactory.GetFont("Arial", 10, Font.NORMAL);

							table.AddCell(new Phrase("idusuario", boldTableFont));
									table.AddCell(new Phrase("descripcion", boldTableFont));
									table.AddCell(new Phrase("abreviatura", boldTableFont));
									table.AddCell(new Phrase("costo1", boldTableFont));
									table.AddCell(new Phrase("fechacreacion", boldTableFont));
									table.AddCell(new Phrase("fechacambio", boldTableFont));
									table.AddCell(new Phrase("TablaCultivos", boldTableFont));
									              
//
            List<TablaActividades> list = db.TablaActividades.Include(t => t.TablaCultivos).ToList();

			foreach (var item in list)
                {
                    
								table.AddCell(new Phrase(item.idusuario == null ? "" : item.idusuario.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.descripcion == null ? "" : item.descripcion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.abreviatura == null ? "" : item.abreviatura.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.costo1 == null ? "" : item.costo1.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacreacion == null ? "" : item.fechacreacion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.TablaCultivos == null ? "" : item.TablaCultivos.ToString(), bodyFont));			
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
