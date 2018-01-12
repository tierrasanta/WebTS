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

	public class FertilizanteIndexViewModel
    {
		public List<TABLAS_ACTIVIDADES> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class ActividadesFertilizanteController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();

		
        // GET: ActividadesFertilizante
        public ActionResult Index(int? page, String Search)
        {
		//
			var viewModel = new FertilizanteIndexViewModel();

            //if (Search == null || Search.Equals(""))
            //{
            var pager = new Pager(db.TABLAS_ACTIVIDADES.Where(x => x.idactividad.StartsWith("03")).Count(), page);
            viewModel.Items = db.TABLAS_ACTIVIDADES.Where(x => x.idactividad.StartsWith("03")).Include(t => t.TABLAS_CULTIVOS)
               .OrderBy(c => c.idactividad)
                    .Skip((pager.CurrentPage - 1) * pager.PageSize)
                    .Take(pager.PageSize).ToList();
            viewModel.Pager = pager;
    //    }
    //        else
    //        {
				//var pager = new Pager(db.TABLAS_ACTIVIDADES.Where(c => c.AgregarVariableAbuscar.Contains(Search)).Count(), page);
    //            viewModel.Items = db.TABLAS_ACTIVIDADES.Include(t => t.TABLAS_CULTIVOS).Where(c => c.AgregarVariableAbuscar.Contains(Search))
    //                    .OrderBy(c => c.TABLAS_ACTIVIDADESID)
    //                    .Skip((pager.CurrentPage - 1) * pager.PageSize)
    //                    .Take(pager.PageSize).ToList();
				//viewModel.Pager = pager;
				//@ViewBag.Search = Search;
    //        }
            return View(viewModel);
        }



        // GET: ActividadesFertilizante/Details/5
        public ActionResult Details(string id, string id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TABLAS_ACTIVIDADES tABLAS_ACTIVIDADES = db.TABLAS_ACTIVIDADES.Find(id, id2);
            if (tABLAS_ACTIVIDADES == null)
            {
                return HttpNotFound();
            }
            return View(tABLAS_ACTIVIDADES);
        }

        // GET: ActividadesFertilizante/Create
        public ActionResult Create()
        {
            ViewBag.unimedida = new SelectList(db.TABLAS_CULTIVOS.Where(x => x.idcodigo.StartsWith("01") && x.idcodigo != "010000"), "idcodigo", "descripcion");
            ViewBag.idactividad = new SelectList(db.TABLAS_ACTIVIDADES.Where(x => x.idactividad.StartsWith("030") && x.idactividad.EndsWith("0000") && x.idactividad != "03000000"), "idactividad", "descripcion");
            return View();
        }

        // POST: ActividadesFertilizante/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idactividad,idusuario,descripcion,abreviatura,unimedida,costo1,fechacreacion,fechacambio")] TABLAS_ACTIVIDADES tABLAS_ACTIVIDADES)
        {
            List<TABLAS_ACTIVIDADES> ta = db.TABLAS_ACTIVIDADES.Where(x => x.idactividad.StartsWith(tABLAS_ACTIVIDADES.idactividad.Substring(0, 4)) && x.idactividad.Substring(4, 7) != "0000" && x.idactividad != "03000000").ToList();
            tABLAS_ACTIVIDADES.idempresa = "01";
            tABLAS_ACTIVIDADES.idusuario = "0001";
            if (ta.Count == 0) { tABLAS_ACTIVIDADES.idactividad = tABLAS_ACTIVIDADES.idactividad.Substring(0, 4) + "0001"; }
            else { tABLAS_ACTIVIDADES.idactividad = getactividades(tABLAS_ACTIVIDADES.idactividad, Convert.ToInt32(ta.Last().idactividad.Substring(4))); }
            if (ModelState.IsValid)
            {
                //tABLAS_ACTIVIDADES.Creado = DateTime.Now;
                //tABLAS_ACTIVIDADES.Modificado = DateTime.Now;
                db.TABLAS_ACTIVIDADES.Add(tABLAS_ACTIVIDADES);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.idempresa = new SelectList(db.TABLAS_CULTIVOS, "idempresa", "idusuario", tABLAS_ACTIVIDADES.idempresa);
            return View(tABLAS_ACTIVIDADES);
        }

        private string getactividades(string idactividad, int v)
        {
            String idact = idactividad.Substring(0, 4);
            v = v + 1;
            int digitos = Convert.ToString(v).Length;
            if (digitos == 1) { return idact + "000" + v; }
            else if (digitos == 2) { return idact + "00" + v; }
            else if (digitos == 3) { return idact + "0" + v; }
            else if (digitos == 4) { return idact + v; }
            else { return Convert.ToString(v); }
        }

        // GET: ActividadesFertilizante/Edit/5
        public ActionResult Edit(string id, string id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TABLAS_ACTIVIDADES tABLAS_ACTIVIDADES = db.TABLAS_ACTIVIDADES.Find(id, id2);
            if (tABLAS_ACTIVIDADES == null)
            {
                return HttpNotFound();
            }
            ViewBag.unimedida = new SelectList(db.TABLAS_CULTIVOS.Where(x => x.idcodigo.StartsWith("01") && x.idcodigo != "010000"), "idcodigo", "descripcion");
            return View(tABLAS_ACTIVIDADES);
        }

        // POST: ActividadesFertilizante/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idactividad,idusuario,descripcion,abreviatura,unimedida,costo1,fechacreacion,fechacambio")] TABLAS_ACTIVIDADES tABLAS_ACTIVIDADES)
        {
            tABLAS_ACTIVIDADES.fechacambio = DateTime.Now;
            if (ModelState.IsValid)
            {
                db.Entry(tABLAS_ACTIVIDADES).State = EntityState.Modified;
				//tABLAS_ACTIVIDADES.Modificado = DateTime.Now;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.idempresa = new SelectList(db.TABLAS_CULTIVOS, "idempresa", "idusuario", tABLAS_ACTIVIDADES.idempresa);
            return View(tABLAS_ACTIVIDADES);
        }

        // GET: ActividadesFertilizante/Delete/5
        public ActionResult Delete(string id, string id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            TABLAS_ACTIVIDADES tABLAS_ACTIVIDADES = db.TABLAS_ACTIVIDADES.Find(id, id2);
            if (tABLAS_ACTIVIDADES == null)
            {
                return HttpNotFound();
            }
            return View(tABLAS_ACTIVIDADES);
        }

        // POST: ActividadesFertilizante/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id, string id2)
        {
            TABLAS_ACTIVIDADES tABLAS_ACTIVIDADES = db.TABLAS_ACTIVIDADES.Find(id, id2);
            db.TABLAS_ACTIVIDADES.Remove(tABLAS_ACTIVIDADES);
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
				//.Include(t => t.TABLAS_CULTIVOS)
                List<TABLAS_ACTIVIDADES> list = db.TABLAS_ACTIVIDADES.Include(t => t.TABLAS_CULTIVOS).ToList();
                int pos = 4;
								ws.Cells[pos, 4].Value = "idusuario";
									ws.Cells[pos, 5].Value = "descripcion";
									ws.Cells[pos, 6].Value = "abreviatura";
									ws.Cells[pos, 8].Value = "costo1";
									ws.Cells[pos, 9].Value = "fechacreacion";
									ws.Cells[pos, 10].Value = "fechacambio";
									ws.Cells[pos, 11].Value = "PLANTILLAS_CULTIVOS";
									ws.Cells[pos, 12].Value = "TABLAS_CULTIVOS";
					
                foreach (var item in list)
                {
                    pos++;
								ws.Cells[pos, 4].Value = item.idusuario == null ? "" : item.idusuario.ToString();				
									ws.Cells[pos, 5].Value = item.descripcion == null ? "" : item.descripcion.ToString();				
									ws.Cells[pos, 6].Value = item.abreviatura == null ? "" : item.abreviatura.ToString();				
									ws.Cells[pos, 8].Value = item.costo1 == null ? "" : item.costo1.ToString();				
									ws.Cells[pos, 9].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();				
									ws.Cells[pos, 10].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();				
									ws.Cells[pos, 11].Value = item.PLANTILLAS_CULTIVOS == null ? "" : item.PLANTILLAS_CULTIVOS.ToString();				
									ws.Cells[pos, 12].Value = item.TABLAS_CULTIVOS == null ? "" : item.TABLAS_CULTIVOS.ToString();				
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

							table.AddCell(new Phrase("idusuario", boldTableFont));
									table.AddCell(new Phrase("descripcion", boldTableFont));
									table.AddCell(new Phrase("abreviatura", boldTableFont));
									table.AddCell(new Phrase("costo1", boldTableFont));
									table.AddCell(new Phrase("fechacreacion", boldTableFont));
									table.AddCell(new Phrase("fechacambio", boldTableFont));
									table.AddCell(new Phrase("PLANTILLAS_CULTIVOS", boldTableFont));
									table.AddCell(new Phrase("TABLAS_CULTIVOS", boldTableFont));
									              
//
            List<TABLAS_ACTIVIDADES> list = db.TABLAS_ACTIVIDADES.Include(t => t.TABLAS_CULTIVOS).ToList();

			foreach (var item in list)
                {
                    
								table.AddCell(new Phrase(item.idusuario == null ? "" : item.idusuario.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.descripcion == null ? "" : item.descripcion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.abreviatura == null ? "" : item.abreviatura.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.costo1 == null ? "" : item.costo1.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacreacion == null ? "" : item.fechacreacion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.PLANTILLAS_CULTIVOS == null ? "" : item.PLANTILLAS_CULTIVOS.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.TABLAS_CULTIVOS == null ? "" : item.TABLAS_CULTIVOS.ToString(), bodyFont));			
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
