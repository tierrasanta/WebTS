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

	public class FundoIndexViewModel
    {
		public List<Fundo> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class FundosController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();

		
        // GET: Fundos
        public ActionResult Index(int? page, String Search)
        {
		//
			var viewModel = new FundoIndexViewModel();            

            var pager = new Pager(db.Fundo.Count(), page);
                viewModel.Items = db.Fundo
                        .OrderBy(c => c.idfundo)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;           
            return View(viewModel);
        }



        // GET: Fundos/Details/5
        public ActionResult Details(string id, string id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Fundo fundo = db.Fundo.Find(id, id2);
            if (fundo == null)
            {
                return HttpNotFound();
            }
            return View(fundo);
        }

        // GET: Fundos/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Fundos/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idusuario,descripcion,fechacreacion,fechacambio,idfundo")] Fundo fundo)
        {
            fundo.idempresa = "01";
            List<Fundo> f = db.Fundo.ToList();
            if (f.Count == 0) { fundo.idfundo = "01"; }
            else { fundo.idfundo = getidfundo(Convert.ToInt32(f.Last().idfundo)); }
            fundo.idusuario = "0001";
            if (ModelState.IsValid)
            {                
                db.Fundo.Add(fundo);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(fundo);
        }

        private string getidfundo(int v)
        {
            v = v + 1;
            int digitos = Convert.ToString(v).Length;
            if (digitos == 1) { return "0" + v; }
            else { return Convert.ToString(v); }
        }

        // GET: Fundos/Edit/5
        public ActionResult Edit(string id, string id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Fundo fundo = db.Fundo.Find(id,id2);
            if (fundo == null)
            {
                return HttpNotFound();
            }
            return View(fundo);
        }

        // POST: Fundos/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idfundo,idusuario,descripcion,fechacreacion,fechacambio")] Fundo fundo)
        {
            fundo.fechacambio = DateTime.Now;
            if (ModelState.IsValid)
            {
                db.Entry(fundo).State = EntityState.Modified;				
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(fundo);
        }

        // GET: Fundos/Delete/5
        public ActionResult Delete(string id, string id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Fundo fundo = db.Fundo.Find(id,id2);
            if (fundo == null)
            {
                return HttpNotFound();
            }
            return View(fundo);
        }

        // POST: Fundos/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id, string id2)
        {
            Fundo fundo = db.Fundo.Find(id, id2);
            db.Fundo.Remove(fundo);
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
                List<Fundo> list = db.Fundo.ToList();
                int pos = 4;
								
									ws.Cells[pos, 4].Value = "descripcion";
									ws.Cells[pos, 5].Value = "fechacreacion";
									ws.Cells[pos, 6].Value = "fechacambio";
									
					
                foreach (var item in list)
                {
                    pos++;
								
									ws.Cells[pos, 4].Value = item.descripcion == null ? "" : item.descripcion.ToString();				
									ws.Cells[pos, 5].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();				
									ws.Cells[pos, 6].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();				
									
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
							
									table.AddCell(new Phrase("descripcion", boldTableFont));
									table.AddCell(new Phrase("fechacreacion", boldTableFont));
									table.AddCell(new Phrase("fechacambio", boldTableFont));									
									              
//
            List<Fundo> list = db.Fundo.ToList();

			foreach (var item in list)
                {                    
								
									table.AddCell(new Phrase(item.descripcion == null ? "" : item.descripcion.ToString(), bodyFont));			
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
