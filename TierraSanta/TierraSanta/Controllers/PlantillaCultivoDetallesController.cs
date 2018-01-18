using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using TierraSanta.Models;
//using AgregarNamespace.Helper;
using OfficeOpenXml;
using iTextSharp.text;
using System.IO;
using iTextSharp.text.pdf;

namespace TierraSanta.Controllers
{

	//public class PlantillaCultivoDetalleIndexViewModel
 //   {
	//	public List<PlantillaCultivoDetalle> Items { get; set; }
 //       public Pager Pager { get; set; }
 //   }

    public class PlantillaCultivoDetallesController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();

		
        // GET: PlantillaCultivoDetalles
  //      public ActionResult Index(int? page, String Search)
  //      {
		////
		//	var viewModel = new PlantillaCultivoDetalleIndexViewModel();            

  //          if (Search == null || Search.Equals(""))
  //          {
		//		var pager = new Pager(db.PlantillaCultivoDetalle.Count(), page);
  //              viewModel.Items = db.PlantillaCultivoDetalle.Include(p => p.PlantillaCultivoCabecera).Include(p => p.TablaActividades)
  //                      .OrderBy(c => c.PlantillaCultivoDetalleID)
  //                      .Skip((pager.CurrentPage - 1) * pager.PageSize)
  //                      .Take(pager.PageSize).ToList();
  //              viewModel.Pager = pager;
  //          }
  //          else
  //          {
		//		var pager = new Pager(db.PlantillaCultivoDetalle.Where(c => c.AgregarVariableAbuscar.Contains(Search)).Count(), page);
  //              viewModel.Items = db.PlantillaCultivoDetalle.Include(p => p.PlantillaCultivoCabecera).Include(p => p.TablaActividades).Where(c => c.AgregarVariableAbuscar.Contains(Search))
  //                      .OrderBy(c => c.PlantillaCultivoDetalleID)
  //                      .Skip((pager.CurrentPage - 1) * pager.PageSize)
  //                      .Take(pager.PageSize).ToList();
		//		viewModel.Pager = pager;
		//		@ViewBag.Search = Search;
  //          }
  //          return View(viewModel);
  //      }



        // GET: PlantillaCultivoDetalles/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PlantillaCultivoDetalle plantillaCultivoDetalle = db.PlantillaCultivoDetalle.Find(id);
            if (plantillaCultivoDetalle == null)
            {
                return HttpNotFound();
            }
            return View(plantillaCultivoDetalle);
        }

        // GET: PlantillaCultivoDetalles/Create
        public ActionResult Create(int idplantilla, int idactividades)
        {
            PlantillaCultivoDetalle plantillaCultivoDetalle = new PlantillaCultivoDetalle();
            plantillaCultivoDetalle.idplantilla = idplantilla;
            ViewBag.idplantilla = new SelectList(db.PlantillaCultivoCabecera, "idplantilla", "idempresa");
            ViewBag.idactividad = new SelectList(db.TablaActividades.Where(t => t.idparent == idactividades), "idactividades", "descripcion");
            return View(plantillaCultivoDetalle);
        }

        // POST: PlantillaCultivoDetalles/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idplantilla,idplantilladetalle,idactividad,idusuario,cantidad,fechacreacion,fechacambio")] PlantillaCultivoDetalle plantillaCultivoDetalle)
        {
            
            if (ModelState.IsValid)
            {
                plantillaCultivoDetalle.idempresa = "01";
                plantillaCultivoDetalle.idusuario = "01";
                
                //
                //plantillaCultivoDetalle.Modificado = DateTime.Now;
                db.PlantillaCultivoDetalle.Add(plantillaCultivoDetalle);
                db.SaveChanges();
                return RedirectToAction("Details", "PlantillaCultivoCabeceras", new { @id = plantillaCultivoDetalle.idplantilla });
            }

            ViewBag.idplantilla = new SelectList(db.PlantillaCultivoCabecera, "idplantilla", "idempresa", plantillaCultivoDetalle.idplantilla);
            ViewBag.idactividad = new SelectList(db.TablaActividades, "idactividades", "idempresa", plantillaCultivoDetalle.idactividad);
            return View(plantillaCultivoDetalle);
        }

        // GET: PlantillaCultivoDetalles/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PlantillaCultivoDetalle plantillaCultivoDetalle = db.PlantillaCultivoDetalle.Find(id);
            if (plantillaCultivoDetalle == null)
            {
                return HttpNotFound();
            }
            ViewBag.idplantilla = new SelectList(db.PlantillaCultivoCabecera, "idplantilla", "idempresa", plantillaCultivoDetalle.idplantilla);
            ViewBag.idactividad = new SelectList(db.TablaActividades, "idactividades", "idempresa", plantillaCultivoDetalle.idactividad);
            return View(plantillaCultivoDetalle);
        }

        // POST: PlantillaCultivoDetalles/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idplantilla,idplantilladetalle,idactividad,idusuario,cantidad,fechacreacion,fechacambio")] PlantillaCultivoDetalle plantillaCultivoDetalle)
        {
            if (ModelState.IsValid)
            {
                db.Entry(plantillaCultivoDetalle).State = EntityState.Modified;
				//plantillaCultivoDetalle.Modificado = DateTime.Now;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.idplantilla = new SelectList(db.PlantillaCultivoCabecera, "idplantilla", "idempresa", plantillaCultivoDetalle.idplantilla);
            ViewBag.idactividad = new SelectList(db.TablaActividades, "idactividades", "idempresa", plantillaCultivoDetalle.idactividad);
            return View(plantillaCultivoDetalle);
        }

        // GET: PlantillaCultivoDetalles/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PlantillaCultivoDetalle plantillaCultivoDetalle = db.PlantillaCultivoDetalle.Find(id);
            if (plantillaCultivoDetalle == null)
            {
                return HttpNotFound();
            }
            return View(plantillaCultivoDetalle);
        }

        // POST: PlantillaCultivoDetalles/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            PlantillaCultivoDetalle plantillaCultivoDetalle = db.PlantillaCultivoDetalle.Find(id);
            int idplantilla = plantillaCultivoDetalle.idplantilla;

            db.PlantillaCultivoDetalle.Remove(plantillaCultivoDetalle);
            db.SaveChanges();
            return RedirectToAction("Details", "PlantillaCultivoCabeceras", new { @id = idplantilla });
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
								ws.Cells[pos, 2].Value = "idempresa";
									ws.Cells[pos, 6].Value = "idusuario";
									ws.Cells[pos, 7].Value = "cantidad";
									ws.Cells[pos, 8].Value = "fechacreacion";
									ws.Cells[pos, 9].Value = "fechacambio";
									ws.Cells[pos, 10].Value = "PlantillaCultivoCabecera";
									ws.Cells[pos, 11].Value = "TablaActividades";
					
                foreach (var item in list)
                {
                    pos++;
								ws.Cells[pos, 2].Value = item.idempresa == null ? "" : item.idempresa.ToString();				
									ws.Cells[pos, 6].Value = item.idusuario == null ? "" : item.idusuario.ToString();				
									ws.Cells[pos, 7].Value = item.cantidad == null ? "" : item.cantidad.ToString();				
									ws.Cells[pos, 8].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();				
									ws.Cells[pos, 9].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();				
									ws.Cells[pos, 10].Value = item.PlantillaCultivoCabecera == null ? "" : item.PlantillaCultivoCabecera.ToString();				
									ws.Cells[pos, 11].Value = item.TablaActividades == null ? "" : item.TablaActividades.ToString();				
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

							table.AddCell(new Phrase("idempresa", boldTableFont));
									table.AddCell(new Phrase("idusuario", boldTableFont));
									table.AddCell(new Phrase("cantidad", boldTableFont));
									table.AddCell(new Phrase("fechacreacion", boldTableFont));
									table.AddCell(new Phrase("fechacambio", boldTableFont));
									table.AddCell(new Phrase("PlantillaCultivoCabecera", boldTableFont));
									table.AddCell(new Phrase("TablaActividades", boldTableFont));
									              
//
            List<PlantillaCultivoDetalle> list = db.PlantillaCultivoDetalle.Include(p => p.PlantillaCultivoCabecera).Include(p => p.TablaActividades).ToList();

			foreach (var item in list)
                {
                    
								table.AddCell(new Phrase(item.idempresa == null ? "" : item.idempresa.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.idusuario == null ? "" : item.idusuario.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.cantidad == null ? "" : item.cantidad.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacreacion == null ? "" : item.fechacreacion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.PlantillaCultivoCabecera == null ? "" : item.PlantillaCultivoCabecera.ToString(), bodyFont));			
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
