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



            if (Search == null || Search.Equals(""))
            {
                var pager = new Pager(db.PlantillaCultivoCabecera.Count(), page);
                viewModel.Items = db.PlantillaCultivoCabecera
                        .OrderBy(c => c.idplantilla)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;
            }
            else
            {

                var pager = new Pager(db.PlantillaCultivoCabecera.Where(c => c.descripcion.Contains(Search)).Count(), page);
                viewModel.Items = db.PlantillaCultivoCabecera.Where(c => c.descripcion.Contains(Search))
                    .OrderBy(c => c.descripcion)
                    .Skip((pager.CurrentPage - 1) * pager.PageSize)
                    .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;
                @ViewBag.Search = Search;
            }
            return View(viewModel);
        }



        // GET: PlantillaCultivoCabeceras/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PlantillaCultivoCabecera plantillaCultivoCabecera = db.PlantillaCultivoCabecera.Find(id);
            if (plantillaCultivoCabecera == null)
            {
                return HttpNotFound();
            }
            List<DetalleVM> detalles = new List<DetalleVM>();
            var actividades = db.TablaActividades.Where(t => t.idparent == null).ToList();
            foreach (var actividad in actividades)
            {
                DetalleVM detalle = new DetalleVM();
                detalle.NombreActividad = actividad.descripcion.ToUpper();
                detalle.IdActividad = actividad.idactividades;
                detalle.IdCabecera = Convert.ToInt32(id);

                //int? idactividades = actividad.idactividades;Convert.ToInt32(id)

                detalle.detalle = db.PlantillaCultivoDetalle
                    .Where(p => p.idplantilla == id && p.TablaActividades.idparent == actividad.idactividades).ToList();

                //detalle.detalle.FirstOrDefault().TablaActividades.descripcion
                //    db.TablaActividades.Where(t => t.idactividad.EndsWith("000000")).ToList();


                detalles.Add(detalle);

            }

            ViewBag.Detalles = detalles;
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
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            PlantillaCultivoCabecera plantillaCultivoCabecera = db.PlantillaCultivoCabecera.Find(id);
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
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PlantillaCultivoCabecera plantillaCultivoCabecera = db.PlantillaCultivoCabecera.Find(id);
            if (plantillaCultivoCabecera == null)
            {
                return HttpNotFound();
            }
            return View(plantillaCultivoCabecera);
        }

        // POST: PlantillaCultivoCabeceras/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            PlantillaCultivoCabecera plantillaCultivoCabecera = db.PlantillaCultivoCabecera.Find(id);
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

        public ActionResult ReportExcel(int id)
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Report");
                ExcelWorksheet ws = package.Workbook.Worksheets[1];
                ws.Name = "Report";
                ws.Cells.Style.Font.Size = 11;
                ws.Cells.Style.Font.Name = "Calibri";
                //
                List<TablaActividades> list = db.TablaActividades.Where(t => t.idparent == null && t.abreviatura == "").ToList();
                int pos = 4;
                int calculofinal = 0;
                //falta agregar nombre de la plantilla
                //ws.Cells[pos, 2].Value = "idempresa";
                //ws.Cells[pos, 4].Value = "descripcion";
                //ws.Cells[pos, 5].Value = "idusuario";
                //ws.Cells[pos, 6].Value = "fechacreacion";
                //ws.Cells[pos, 7].Value = "fechacambio";
                //ws.Cells[pos, 8].Value = "PlantillaCultivoDetalle";
                List<string> posiciones = new List<string>();
                foreach (var item in list)
                {
                    string linicio = "";
                    string lfin = "";
                    int inicio = 0;
                    pos++;
                    ws.Cells[pos, 2].Value = item.descripcion == null ? "" : item.descripcion.ToString();
                    pos++;
                    ws.Cells[pos, 2].Value = "Descripcion";
                    ws.Cells[pos, 4].Value = "Cantidad";
                    ws.Cells[pos, 5].Value = "Costo";
                    ws.Cells[pos, 6].Value = "Subtotal";
                    List<PlantillaCultivoDetalle> listdetalle = db.PlantillaCultivoDetalle.Where(d => d.idplantilla == id && d.TablaActividades.idparent == item.idactividades).ToList();//Where(t => t.idplantilla == id && t.TablaActividades.idparent == item.idparent).ToList();
                    foreach (var itemdetalle in listdetalle)
                    {
                        pos++;
                        if (inicio == 0)
                        {
                            linicio = ws.Cells[pos, 6].FullAddress;
                        }
                        if (inicio == listdetalle.Count - 1)
                        {
                            lfin = ws.Cells[pos, 6].FullAddress;
                        }

                        ws.Cells[pos, 2].Value = itemdetalle.TablaActividades.descripcion == null ? "" : itemdetalle.TablaActividades.descripcion.ToString();
                        ws.Cells[pos, 4].Value = itemdetalle.cantidad;
                        ws.Cells[pos, 5].Value = itemdetalle.TablaActividades.costo1;
                        ws.Cells[pos, 6].Formula = "PRODUCT(" + ws.Cells[pos, 4].FullAddress + ":" + ws.Cells[pos, 5].FullAddress + ")";
                        inicio++;
                    }
                    pos++;
                    ws.Cells[pos, 2].Value = "Total por actividad";
                    ws.Cells[pos, 6].Formula = "SUM(" + linicio + ":" + lfin + ")";
                    posiciones.Add(ws.Cells[pos, 6].FullAddress);
                    pos++;
                    ws.Cells[pos, 2].Value = "";
                    //ws.Cells[pos, 4].Value = item.descripcion == null ? "" : item.descripcion.ToString();
                    //ws.Cells[pos, 5].Value = item.idusuario == null ? "" : item.idusuario.ToString();
                    //ws.Cells[pos, 6].Value = item.fechacreacion == null ? "" : item.fechacreacion.ToString();
                    //ws.Cells[pos, 7].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();
                    //ws.Cells[pos, 8].Value = item.PlantillaCultivoDetalle == null ? "" : item.PlantillaCultivoDetalle.ToString();
                }
                pos++;
                string formula = "";
                ws.Cells[pos, 2].Value = "Total de la plantilla";
                foreach (var item in posiciones)
                {
                    if (calculofinal == 0)
                    {
                        formula = "SUM(" + item + ";";
                    }
                    else if (calculofinal == posiciones.Count - 1)
                    {
                        formula = formula + item + ")";
                    }
                    else
                    {
                        formula = formula + item + ";";
                    }
                    calculofinal++;
                }
                ws.Cells[pos, 2].Formula = formula;
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

            table.AddCell(new Phrase("idempresa", boldTableFont));
            table.AddCell(new Phrase("descripcion", boldTableFont));
            table.AddCell(new Phrase("idusuario", boldTableFont));
            table.AddCell(new Phrase("fechacreacion", boldTableFont));
            table.AddCell(new Phrase("fechacambio", boldTableFont));
            table.AddCell(new Phrase("PlantillaCultivoDetalle", boldTableFont));

            //
            List<PlantillaCultivoCabecera> list = db.PlantillaCultivoCabecera.ToList();

            foreach (var item in list)
            {

                table.AddCell(new Phrase(item.idempresa == null ? "" : item.idempresa.ToString(), bodyFont));
                table.AddCell(new Phrase(item.descripcion == null ? "" : item.descripcion.ToString(), bodyFont));
                table.AddCell(new Phrase(item.idusuario == null ? "" : item.idusuario.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechacreacion == null ? "" : item.fechacreacion.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));
                table.AddCell(new Phrase(item.PlantillaCultivoDetalle == null ? "" : item.PlantillaCultivoDetalle.ToString(), bodyFont));
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
