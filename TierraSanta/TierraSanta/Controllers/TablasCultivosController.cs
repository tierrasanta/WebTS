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

	public class TABLAS_CULTIVOSIndexViewModel
    {
		public List<TABLAS_CULTIVOS> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class TablasCultivosController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();

		
        // GET: TablasCultivos
        public ActionResult Index(int? page, String Search)
        {
		//
			var viewModel = new TABLAS_CULTIVOSIndexViewModel();            

            //if (Search == null || Search.Equals(""))
            //{
				var pager = new Pager(db.TABLAS_CULTIVOS.Count(), page);
                viewModel.Items = db.TABLAS_CULTIVOS
                        .OrderBy(c => c.idcodigo)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;
    //        }
    //        else
    //        {
				//var pager = new Pager(db.TABLAS_CULTIVOS.Where(c => c.AgregarVariableAbuscar.Contains(Search)).Count(), page);
    //            viewModel.Items = db.TABLAS_CULTIVOS.Where(c => c.AgregarVariableAbuscar.Contains(Search))
    //                    .OrderBy(c => c.TABLAS_CULTIVOSID)
    //                    .Skip((pager.CurrentPage - 1) * pager.PageSize)
    //                    .Take(pager.PageSize).ToList();
				//viewModel.Pager = pager;
				//@ViewBag.Search = Search;
    //        }
            return View(viewModel);
        }



        // GET: TablasCultivos/Details/5
        public ActionResult Details(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TABLAS_CULTIVOS tABLAS_CULTIVOS = db.TABLAS_CULTIVOS.Find(id);
            if (tABLAS_CULTIVOS == null)
            {
                return HttpNotFound();
            }
            return View(tABLAS_CULTIVOS);
        }

        // GET: TablasCultivos/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: TablasCultivos/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idcodigo,idusuario,descripcion,abreviatura,valornum,valorchar,fechacreacion,fechacambio")] TABLAS_CULTIVOS tABLAS_CULTIVOS)
        {
            if (ModelState.IsValid)
            {
                //tABLAS_CULTIVOS.Creado = DateTime.Now;
                //tABLAS_CULTIVOS.Modificado = DateTime.Now;
                db.TABLAS_CULTIVOS.Add(tABLAS_CULTIVOS);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(tABLAS_CULTIVOS);
        }

        // GET: TablasCultivos/Edit/5
        public ActionResult Edit(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TABLAS_CULTIVOS tABLAS_CULTIVOS = db.TABLAS_CULTIVOS.Find(id);
            if (tABLAS_CULTIVOS == null)
            {
                return HttpNotFound();
            }
            return View(tABLAS_CULTIVOS);
        }

        // POST: TablasCultivos/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idcodigo,idusuario,descripcion,abreviatura,valornum,valorchar,fechacreacion,fechacambio")] TABLAS_CULTIVOS tABLAS_CULTIVOS)
        {
            if (ModelState.IsValid)
            {
                db.Entry(tABLAS_CULTIVOS).State = EntityState.Modified;
				//tABLAS_CULTIVOS.Modificado = DateTime.Now;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(tABLAS_CULTIVOS);
        }

        // GET: TablasCultivos/Delete/5
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TABLAS_CULTIVOS tABLAS_CULTIVOS = db.TABLAS_CULTIVOS.Find(id);
            if (tABLAS_CULTIVOS == null)
            {
                return HttpNotFound();
            }
            return View(tABLAS_CULTIVOS);
        }

        // POST: TablasCultivos/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            TABLAS_CULTIVOS tABLAS_CULTIVOS = db.TABLAS_CULTIVOS.Find(id);
            db.TABLAS_CULTIVOS.Remove(tABLAS_CULTIVOS);
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
                List<TABLAS_CULTIVOS> list = db.TABLAS_CULTIVOS.ToList();
                int pos = 4;
								ws.Cells[pos, 4].Value = "idusuario";
									ws.Cells[pos, 5].Value = "descripcion";
									ws.Cells[pos, 6].Value = "abreviatura";
									ws.Cells[pos, 7].Value = "valornum";
									ws.Cells[pos, 8].Value = "valorchar";
									ws.Cells[pos, 9].Value = "fechacreacion";
									ws.Cells[pos, 10].Value = "fechacambio";
									ws.Cells[pos, 11].Value = "CULTIVOS";
									ws.Cells[pos, 12].Value = "PLANTILLAS_CULTIVOS";
									ws.Cells[pos, 13].Value = "TABLAS_ACTIVIDADES";
					
                foreach (var item in list)
                {
                    pos++;
								ws.Cells[pos, 4].Value = item.idusuario == null ? "" : item.idusuario.ToString();				
									ws.Cells[pos, 5].Value = item.descripcion == null ? "" : item.descripcion.ToString();				
									ws.Cells[pos, 6].Value = item.abreviatura == null ? "" : item.abreviatura.ToString();				
									ws.Cells[pos, 7].Value = item.valornum == null ? "" : item.valornum.ToString();				
									ws.Cells[pos, 8].Value = item.valorchar == null ? "" : item.valorchar.ToString();				
									ws.Cells[pos, 9].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();				
									ws.Cells[pos, 10].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();				
									ws.Cells[pos, 11].Value = item.CULTIVOS == null ? "" : item.CULTIVOS.ToString();				
									ws.Cells[pos, 12].Value = item.PLANTILLAS_CULTIVOS == null ? "" : item.PLANTILLAS_CULTIVOS.ToString();				
									ws.Cells[pos, 13].Value = item.TABLAS_ACTIVIDADES == null ? "" : item.TABLAS_ACTIVIDADES.ToString();				
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

			
            var table = new PdfPTable(10);

            var boldTableFont = FontFactory.GetFont("Arial", 10, Font.BOLD);
            var bodyFont = FontFactory.GetFont("Arial", 10, Font.NORMAL);

							table.AddCell(new Phrase("idusuario", boldTableFont));
									table.AddCell(new Phrase("descripcion", boldTableFont));
									table.AddCell(new Phrase("abreviatura", boldTableFont));
									table.AddCell(new Phrase("valornum", boldTableFont));
									table.AddCell(new Phrase("valorchar", boldTableFont));
									table.AddCell(new Phrase("fechacreacion", boldTableFont));
									table.AddCell(new Phrase("fechacambio", boldTableFont));
									table.AddCell(new Phrase("CULTIVOS", boldTableFont));
									table.AddCell(new Phrase("PLANTILLAS_CULTIVOS", boldTableFont));
									table.AddCell(new Phrase("TABLAS_ACTIVIDADES", boldTableFont));
									              
//
            List<TABLAS_CULTIVOS> list = db.TABLAS_CULTIVOS.ToList();

			foreach (var item in list)
                {
                    
								table.AddCell(new Phrase(item.idusuario == null ? "" : item.idusuario.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.descripcion == null ? "" : item.descripcion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.abreviatura == null ? "" : item.abreviatura.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.valornum == null ? "" : item.valornum.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.valorchar == null ? "" : item.valorchar.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacreacion == null ? "" : item.fechacreacion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.CULTIVOS == null ? "" : item.CULTIVOS.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.PLANTILLAS_CULTIVOS == null ? "" : item.PLANTILLAS_CULTIVOS.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.TABLAS_ACTIVIDADES == null ? "" : item.TABLAS_ACTIVIDADES.ToString(), bodyFont));			
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
