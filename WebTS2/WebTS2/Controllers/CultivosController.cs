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

	public class CultivoIndexViewModel
    {
		public List<Cultivo> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class CultivosController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();

		
        // GET: Cultivos
        public ActionResult Index(int? page, String Search)
        {
		//
			var viewModel = new CultivoIndexViewModel();            

            
				//var pager = new Pager(db.Cultivo.Count(), page);
    //            viewModel.Items = db.Cultivo.Include(c => c.Fundo).Include(c => c.Lote).Include(c => c.PlantillaCultivoCabecera)
    //                    .OrderBy(c => c.idcultivo)
    //                    .Skip((pager.CurrentPage - 1) * pager.PageSize)
    //                    .Take(pager.PageSize).ToList();
    //            viewModel.Pager = pager;          
           

            if (Search == null || Search.Equals(""))
            {
                var pager = new Pager(db.Cultivo.Count(), page);
                viewModel.Items = db.Cultivo.Include(c => c.Fundo).Include(c => c.Lote).Include(c => c.PlantillaCultivoCabecera).Include(c=>c.TablaCultivos)
                        .OrderBy(c => c.idcultivo)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;
            }
            else
            {
                
                var pager = new Pager(db.Cultivo.Include(c => c.Fundo).Include(c => c.Lote).Include(c => c.PlantillaCultivoCabecera).Include(c => c.TablaCultivos).Where(c => c.PlantillaCultivoCabecera.descripcion.Contains(Search) || c.Fundo.descripcion.Contains(Search) || c.Lote.descripcion.Contains(Search) || c.TablaCultivos.descripcion.Contains(Search)).Count() , page);
                viewModel.Items = db.Cultivo.Include(c => c.Fundo).Include(c => c.Lote).Include(c => c.PlantillaCultivoCabecera).Include(c => c.TablaCultivos).Where(c => c.PlantillaCultivoCabecera.descripcion.Contains(Search) || c.Fundo.descripcion.Contains(Search) || c.Lote.descripcion.Contains(Search) || c.TablaCultivos.descripcion.Contains(Search))
                        .OrderBy(c => c.PlantillaCultivoCabecera.descripcion)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;
                @ViewBag.Search = Search;
            }
            return View(viewModel);
        }



        // GET: Cultivos/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Cultivo cultivo = db.Cultivo.Find(id);
            if (cultivo == null)
            {
                return HttpNotFound();
            }
            List<DetalleCultivoVM> detalles = new List<DetalleCultivoVM>();
            var actividades = db.TablaActividades.Where(a => a.idparent == null && a.abreviatura == "").ToList();
            foreach(var actividad in actividades)
            {
                DetalleCultivoVM detallec = new DetalleCultivoVM();
                detallec.NombreActividad = actividad.descripcion.ToUpper();
                detallec.IdActividad = actividad.idactividades;
                detallec.IdCabecera = Convert.ToInt32(id);
                detallec.detalle = db.CultivoDetalle.Where(d => d.idcultivo == id && d.TablaActividades.idparent== actividad.idactividades).ToList();
                detalles.Add(detallec);
            }
            ViewBag.Detalles = detalles;
            //return this.RedirectToAction("create","CultivoDetallesController");
            return View(cultivo);
        }

        // GET: Cultivos/Create
        public ActionResult Create()
        {
            ViewBag.idfundo = new SelectList(db.Fundo, "idfundo", "descripcion");
            ViewBag.idlote = new SelectList(db.Lote, "idlote", "descripcion");
            ViewBag.idplantilla = new SelectList(db.PlantillaCultivoCabecera, "idplantilla", "descripcion");
            ViewBag.idtablacultivos = new SelectList(db.TablaCultivos.Where(t => t.idcodigo.StartsWith("02") && t.abreviatura != ""),"pktablacultivos","descripcion");
            return View();
        }

        public ActionResult Createprorrateo()
        {
            ViewBag.idactividadprorrateo = new SelectList(db.TablaActividades.Where(t=>t.prorrateo == true), "idactividades", "descripcion");            
            return View();
        }

        // POST: Cultivos/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idfundo,idlote,idcultivo,idtablacultivos,idusuario,idplantilla,area,fechainicio,fechafin,fechacreacion,fechacambio")] Cultivo cultivo)
        {
            cultivo.idempresa = "01";
            cultivo.idusuario = "0001";
            int id =  cultivo.idtablacultivos;
            if (ModelState.IsValid)
            {                
                db.Cultivo.Add(cultivo);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.idfundo = new SelectList(db.Fundo, "idfundo", "descripcion");
            ViewBag.idlote = new SelectList(db.Lote, "idlote", "descripcion");
            ViewBag.idplantilla = new SelectList(db.PlantillaCultivoCabecera, "idplantilla", "descripcion");
            ViewBag.idtablacultivos = new SelectList(db.TablaCultivos.Where(t => t.idcodigo.StartsWith("02") && t.abreviatura != ""), "pktablacultivos", "descripcion");
            return View(cultivo);
        }

        // GET: Cultivos/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Cultivo cultivo = db.Cultivo.Find(id);
            if (cultivo == null)
            {
                return HttpNotFound();
            }
            ViewBag.idfundo = new SelectList(db.Fundo, "idfundo", "descripcion");
            ViewBag.idlote = new SelectList(db.Lote, "idlote", "descripcion");
            ViewBag.idplantilla = new SelectList(db.PlantillaCultivoCabecera, "idplantilla", "descripcion");
            ViewBag.idtablacultivos = new SelectList(db.TablaCultivos.Where(t => t.idcodigo.StartsWith("02") && t.abreviatura != ""), "pktablacultivos", "descripcion");
            return View(cultivo);
        }

        // POST: Cultivos/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idfundo,idlote,idcultivo,idtablacultivos,idusuario,idplantilla,area,fechainicio,fechafin,fechacreacion,fechacambio")] Cultivo cultivo)
        {
            cultivo.fechacambio = DateTime.Now;
            if (ModelState.IsValid)
            {
                db.Entry(cultivo).State = EntityState.Modified;				
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.idfundo = new SelectList(db.Fundo, "idfundo", "descripcion");
            ViewBag.idlote = new SelectList(db.Lote, "idlote", "descripcion");
            ViewBag.idplantilla = new SelectList(db.PlantillaCultivoCabecera, "idplantilla", "descripcion");
            ViewBag.idtablacultivos = new SelectList(db.TablaCultivos.Where(t => t.idcodigo.StartsWith("02") && t.abreviatura != ""), "pktablacultivos", "descripcion");
            return View(cultivo);
        }

        // GET: Cultivos/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Cultivo cultivo = db.Cultivo.Find(id);
            if (cultivo == null)
            {
                return HttpNotFound();
            }
            return View(cultivo);
        }

        // POST: Cultivos/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Cultivo cultivo = db.Cultivo.Find(id);
            db.Cultivo.Remove(cultivo);
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
				//.Include(c => c.Fundo).Include(c => c.Lote).Include(c => c.PlantillaCultivoCabecera)
                List<Cultivo> list = db.Cultivo.Include(c => c.Fundo).Include(c => c.Lote).Include(c => c.PlantillaCultivoCabecera).ToList();
                int pos = 4;
								ws.Cells[pos, 6].Value = "idusuario";
									ws.Cells[pos, 8].Value = "area";
									ws.Cells[pos, 9].Value = "fechainicio";
									ws.Cells[pos, 10].Value = "fechafin";
									ws.Cells[pos, 11].Value = "fechacreacion";
									ws.Cells[pos, 12].Value = "fechacambio";
									ws.Cells[pos, 13].Value = "Fundo";
									ws.Cells[pos, 14].Value = "Lote";
									ws.Cells[pos, 15].Value = "PlantillaCultivoCabecera";
									ws.Cells[pos, 16].Value = "CultivoDetalle";
					
                foreach (var item in list)
                {
                    pos++;
								ws.Cells[pos, 6].Value = item.idusuario == null ? "" : item.idusuario.ToString();				
									ws.Cells[pos, 8].Value = item.area == null ? "" : item.area.ToString();				
									ws.Cells[pos, 9].Value = item.fechainicio == null ? "" : item.fechainicio.ToString();				
									ws.Cells[pos, 10].Value = item.fechafin == null ? "" : item.fechafin.ToString();				
									ws.Cells[pos, 11].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();				
									ws.Cells[pos, 12].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();				
									ws.Cells[pos, 13].Value = item.Fundo == null ? "" : item.Fundo.ToString();				
									ws.Cells[pos, 14].Value = item.Lote == null ? "" : item.Lote.ToString();				
									ws.Cells[pos, 15].Value = item.PlantillaCultivoCabecera == null ? "" : item.PlantillaCultivoCabecera.ToString();				
									ws.Cells[pos, 16].Value = item.CultivoDetalle == null ? "" : item.CultivoDetalle.ToString();				
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
									table.AddCell(new Phrase("area", boldTableFont));
									table.AddCell(new Phrase("fechainicio", boldTableFont));
									table.AddCell(new Phrase("fechafin", boldTableFont));
									table.AddCell(new Phrase("fechacreacion", boldTableFont));
									table.AddCell(new Phrase("fechacambio", boldTableFont));
									table.AddCell(new Phrase("Fundo", boldTableFont));
									table.AddCell(new Phrase("Lote", boldTableFont));
									table.AddCell(new Phrase("PlantillaCultivoCabecera", boldTableFont));
									table.AddCell(new Phrase("CultivoDetalle", boldTableFont));
									              
//
            List<Cultivo> list = db.Cultivo.Include(c => c.Fundo).Include(c => c.Lote).Include(c => c.PlantillaCultivoCabecera).ToList();

			foreach (var item in list)
                {
                    
								table.AddCell(new Phrase(item.idusuario == null ? "" : item.idusuario.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.area == null ? "" : item.area.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechainicio == null ? "" : item.fechainicio.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechafin == null ? "" : item.fechafin.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacreacion == null ? "" : item.fechacreacion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.Fundo == null ? "" : item.Fundo.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.Lote == null ? "" : item.Lote.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.PlantillaCultivoCabecera == null ? "" : item.PlantillaCultivoCabecera.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.CultivoDetalle == null ? "" : item.CultivoDetalle.ToString(), bodyFont));			
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
