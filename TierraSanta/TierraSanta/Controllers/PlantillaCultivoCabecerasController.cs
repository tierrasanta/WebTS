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

	public class PlantillaCultivoCabeceraIndexViewModel
    {
		public List<PlantillaCultivoCabecera> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class PlantillaCultivoCabecerasController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();

		
        // GET: PlantillaCultivoCabeceras
        public ActionResult Index(int? page, String Search)
        {
		//
			var viewModel = new PlantillaCultivoCabeceraIndexViewModel();            

           	var pager = new Pager(db.PlantillaCultivoCabecera.Count(), page);
                viewModel.Items = db.PlantillaCultivoCabecera
                        .OrderBy(c => c.idplantilla)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;
           
            return View(viewModel);
        }



        // GET: PlantillaCultivoCabeceras/Details/5
        public ActionResult Details(string id, int id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PlantillaCultivoCabecera plantillaCultivoCabecera = db.PlantillaCultivoCabecera.Find(id, id2);
            if (plantillaCultivoCabecera == null)
            {
                return HttpNotFound();
            }
            return View(plantillaCultivoCabecera);
        }

        // GET: PlantillaCultivoCabeceras/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: PlantillaCultivoCabeceras/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idplantilla,descripcion,idusuario,fechacreacion,fechacambio")] PlantillaCultivoCabecera plantillaCultivoCabecera)
        {
            plantillaCultivoCabecera.idempresa = "01";
            plantillaCultivoCabecera.idusuario = "0001";
            if (ModelState.IsValid)
            {               
                db.PlantillaCultivoCabecera.Add(plantillaCultivoCabecera);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(plantillaCultivoCabecera);
        }

        // GET: PlantillaCultivoCabeceras/Edit/5
        public ActionResult Edit(string id, int id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PlantillaCultivoCabecera plantillaCultivoCabecera = db.PlantillaCultivoCabecera.Find(id, id2);
            if (plantillaCultivoCabecera == null)
            {
                return HttpNotFound();
            }
            return View(plantillaCultivoCabecera);
        }

        // POST: PlantillaCultivoCabeceras/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idplantilla,descripcion,idusuario,fechacreacion,fechacambio")] PlantillaCultivoCabecera plantillaCultivoCabecera)
        {
            plantillaCultivoCabecera.fechacambio = DateTime.Now;
            if (ModelState.IsValid)
            {
                db.Entry(plantillaCultivoCabecera).State = EntityState.Modified;				
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(plantillaCultivoCabecera);
        }

        // GET: PlantillaCultivoCabeceras/Delete/5
        public ActionResult Delete(string id, int id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PlantillaCultivoCabecera plantillaCultivoCabecera = db.PlantillaCultivoCabecera.Find(id,id2);
            if (plantillaCultivoCabecera == null)
            {
                return HttpNotFound();
            }
            return View(plantillaCultivoCabecera);
        }

        // POST: PlantillaCultivoCabeceras/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id, int id2)
        {
            PlantillaCultivoCabecera plantillaCultivoCabecera = db.PlantillaCultivoCabecera.Find(id, id2);
            db.PlantillaCultivoCabecera.Remove(plantillaCultivoCabecera);
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
                List<PlantillaCultivoCabecera> list = db.PlantillaCultivoCabecera.ToList();
                int pos = 4;
								ws.Cells[pos, 4].Value = "descripcion";									
									ws.Cells[pos, 6].Value = "fechacreacion";
									ws.Cells[pos, 7].Value = "fechacambio";							
					
                foreach (var item in list)
                {
                    pos++;
								ws.Cells[pos, 4].Value = item.descripcion == null ? "" : item.descripcion.ToString();					
									ws.Cells[pos, 6].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();				
									ws.Cells[pos, 7].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();					
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
            List<PlantillaCultivoCabecera> list = db.PlantillaCultivoCabecera.ToList();

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
