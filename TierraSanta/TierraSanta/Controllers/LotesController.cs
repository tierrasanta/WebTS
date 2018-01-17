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

	public class LoteIndexViewModel
    {
		public List<Lote> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class LotesController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();

		
        // GET: Lotes
        public ActionResult Index(int? page, String Search)
        {
		//
			var viewModel = new LoteIndexViewModel();            

           	var pager = new Pager(db.Lote.Count(), page);
                viewModel.Items = db.Lote.Include(l => l.Fundo)
                        .OrderBy(c => c.idlote)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;
            
            return View(viewModel);
        }



        // GET: Lotes/Details/5
        public ActionResult Details(string id, string id2, string id3)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Lote lote = db.Lote.Find(id,id2,id3);
            if (lote == null)
            {
                return HttpNotFound();
            }
            return View(lote);
        }

        // GET: Lotes/Create
        public ActionResult Create()
        {
            ViewBag.idfundo = new SelectList(db.Fundo, "idfundo", "descripcion");
            return View();
        }

        // POST: Lotes/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idlote,idusuario,descripcion,area,fechacreacion,fechacambio,idfundo")] Lote lote)
        {
            lote.idempresa = "01";
            List<Lote> l = db.Lote.ToList();
            if (l.Count == 0) { lote.idlote = "0001"; }
            else { lote.idlote = getidlote(Convert.ToInt32(l.Last().idlote)); }
            lote.idusuario = "0001";
            if (ModelState.IsValid)
            {
                
                db.Lote.Add(lote);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.idfundo = new SelectList(db.Fundo, "idfundo", "descripcion", lote.idempresa);
            return View(lote);
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
            Lote lote = db.Lote.Find(id,id2,id3);
            if (lote == null)
            {
                return HttpNotFound();
            }
            ViewBag.idfundo = new SelectList(db.Fundo, "idfundo", "descripcion", lote.idempresa);
            return View(lote);
        }

        // POST: Lotes/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idlote,idusuario,descripcion,area,fechacreacion,fechacambio,idfundo")] Lote lote)
        {
            lote.fechacambio = DateTime.Now;
            if (ModelState.IsValid)
            {
                db.Entry(lote).State = EntityState.Modified;				
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.idfundo = new SelectList(db.Fundo, "idfundo", "descripcion", lote.idempresa);
            return View(lote);
        }

        // GET: Lotes/Delete/5
        public ActionResult Delete(string id, string id2, string id3)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Lote lote = db.Lote.Find(id,id2,id3);
            if (lote == null)
            {
                return HttpNotFound();
            }
            return View(lote);
        }

        // POST: Lotes/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id, string id2, string id3)
        {
            Lote lote = db.Lote.Find(id,id2,id3);
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
								ws.Cells[pos, 4].Value = "Fundo";
									ws.Cells[pos, 5].Value = "descripcion";
									ws.Cells[pos, 6].Value = "area";
									ws.Cells[pos, 7].Value = "fechacreacion";
									ws.Cells[pos, 8].Value = "fechacambio";
														
                foreach (var item in list)
                {
                    pos++;
								ws.Cells[pos, 4].Value = item.Fundo.descripcion == null ? "" : item.Fundo.descripcion.ToString();				
									ws.Cells[pos, 5].Value = item.descripcion == null ? "" : item.descripcion.ToString();				
									ws.Cells[pos, 6].Value = item.area == null ? "" : item.area.ToString();				
									ws.Cells[pos, 7].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();				
									ws.Cells[pos, 8].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();																	
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

							table.AddCell(new Phrase("Fundo", boldTableFont));
									table.AddCell(new Phrase("Lote", boldTableFont));
									table.AddCell(new Phrase("area", boldTableFont));
									table.AddCell(new Phrase("fecha de creacion", boldTableFont));
									table.AddCell(new Phrase("fecha de cambio", boldTableFont));									
									              
//
            List<Lote> list = db.Lote.Include(l => l.Fundo).ToList();

			foreach (var item in list)
                {
                    
								table.AddCell(new Phrase(item.Fundo.descripcion == null ? "" : item.Fundo.descripcion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.descripcion == null ? "" : item.descripcion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.area == null ? "" : item.area.ToString(), bodyFont));			
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
