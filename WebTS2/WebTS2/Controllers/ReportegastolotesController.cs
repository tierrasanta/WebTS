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

	public class ReportegastolotesIndexViewModel
    {
		public List<Lote> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class ReportegastolotesController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();

		
        // GET: Reportegastolotes
        public ActionResult Index(int? page, String Search)
        {
			var viewModel = new LoteIndexViewModel();
            if (Search == null || Search.Equals(""))
            {
				var pager = new Pager(db.Lote.Count(), page);
                viewModel.Items = db.Lote.Include(l => l.Fundo)
                        .OrderBy(c => c.idlote)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;
            }
            else
            {
				var pager = new Pager(db.Lote.Where(c => c.descripcion.Contains(Search)).Count(), page);
                viewModel.Items = db.Lote.Include(l => l.Fundo).Where(c => c.descripcion.Contains(Search))
                        .OrderBy(c => c.idlote)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
				viewModel.Pager = pager;
				@ViewBag.Search = Search;
            }
            return View(viewModel);
        }



        // GET: Reportegastolotes/Details/5
        public ActionResult Details(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Lote lote = db.Lote.Find(id);
            if (lote == null)
            {
                return HttpNotFound();
            }
            return View(lote);
        }

        // GET: Reportegastolotes/Create
        public ActionResult Create()
        {
            ViewBag.idempresa = new SelectList(db.Fundo, "idempresa", "idusuario");
            return View();
        }

        // POST: Reportegastolotes/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idfundo,idlote,idusuario,descripcion,area,fechacreacion,fechacambio")] Lote lote)
        {
            if (ModelState.IsValid)
            {
                //lote.Creado = DateTime.Now;
                //lote.Modificado = DateTime.Now;
                db.Lote.Add(lote);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.idempresa = new SelectList(db.Fundo, "idempresa", "idusuario", lote.idempresa);
            return View(lote);
        }

        // GET: Reportegastolotes/Edit/5
        public ActionResult Edit(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Lote lote = db.Lote.Find(id);
            if (lote == null)
            {
                return HttpNotFound();
            }
            ViewBag.idempresa = new SelectList(db.Fundo, "idempresa", "idusuario", lote.idempresa);
            return View(lote);
        }

        // POST: Reportegastolotes/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idfundo,idlote,idusuario,descripcion,area,fechacreacion,fechacambio")] Lote lote)
        {
            if (ModelState.IsValid)
            {
                db.Entry(lote).State = EntityState.Modified;
				//lote.Modificado = DateTime.Now;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.idempresa = new SelectList(db.Fundo, "idempresa", "idusuario", lote.idempresa);
            return View(lote);
        }

        // GET: Reportegastolotes/Delete/5
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Lote lote = db.Lote.Find(id);
            if (lote == null)
            {
                return HttpNotFound();
            }
            return View(lote);
        }

        // POST: Reportegastolotes/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            Lote lote = db.Lote.Find(id);
            db.Lote.Remove(lote);
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
				//.Include(l => l.Fundo)
                List<Lote> list = db.Lote.Include(l => l.Fundo).ToList();
                int pos = 4;
								ws.Cells[pos, 5].Value = "idusuario";
									ws.Cells[pos, 6].Value = "descripcion";
									ws.Cells[pos, 7].Value = "area";
									ws.Cells[pos, 8].Value = "fechacreacion";
									ws.Cells[pos, 9].Value = "fechacambio";
									ws.Cells[pos, 10].Value = "Fundo";
									ws.Cells[pos, 11].Value = "Cultivo";
					
                foreach (var item in list)
                {
                    pos++;
								ws.Cells[pos, 5].Value = item.idusuario == null ? "" : item.idusuario.ToString();				
									ws.Cells[pos, 6].Value = item.descripcion == null ? "" : item.descripcion.ToString();				
									ws.Cells[pos, 7].Value = item.area == null ? "" : item.area.ToString();				
									ws.Cells[pos, 8].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();				
									ws.Cells[pos, 9].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();				
									ws.Cells[pos, 10].Value = item.Fundo == null ? "" : item.Fundo.ToString();				
									ws.Cells[pos, 11].Value = item.Cultivo == null ? "" : item.Cultivo.ToString();				
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
									table.AddCell(new Phrase("area", boldTableFont));
									table.AddCell(new Phrase("fechacreacion", boldTableFont));
									table.AddCell(new Phrase("fechacambio", boldTableFont));
									table.AddCell(new Phrase("Fundo", boldTableFont));
									table.AddCell(new Phrase("Cultivo", boldTableFont));
									              
//
            List<Lote> list = db.Lote.Include(l => l.Fundo).ToList();

			foreach (var item in list)
                {
                    
								table.AddCell(new Phrase(item.idusuario == null ? "" : item.idusuario.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.descripcion == null ? "" : item.descripcion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.area == null ? "" : item.area.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacreacion == null ? "" : item.fechacreacion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.Fundo == null ? "" : item.Fundo.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.Cultivo == null ? "" : item.Cultivo.ToString(), bodyFont));			
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
