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

    public class ActividadesProrrateoIndexViewModel
    {
        public List<indexprorrateoVM> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class ActividadesProrrateoController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();


        // GET: ActividadesProrrateo
        public ActionResult Index(int? page, String Search)
        {
            //
            var viewModel = new ActividadesProrrateoIndexViewModel();

            if (Search == null || Search.Equals(""))
            {
                var pager = new Pager(db.PlantillaCultivoDetalle.Include(p => p.PlantillaCultivoCabecera).Include(p => p.TablaActividades).Where(p => p.TablaActividades.prorrateo == true).GroupBy(p => p.TablaActividades.idactividades).Select(p => p.FirstOrDefault()).Count(), page);
                List<PlantillaCultivoDetalle> list = new List<PlantillaCultivoDetalle>();
                List<indexprorrateoVM> index = new List<indexprorrateoVM>();
                list = db.PlantillaCultivoDetalle.Include(p => p.PlantillaCultivoCabecera).Include(p => p.TablaActividades).Where(p => p.TablaActividades.prorrateo == true).GroupBy(p => p.TablaActividades.idactividades).Select(p => p.FirstOrDefault())
                        .OrderBy(c => c.idplantilladetalle)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                foreach (var item in list)
                {
                    indexprorrateoVM itemprorrateo = new indexprorrateoVM();
                    itemprorrateo.cantidad = item.cantidad;
                    itemprorrateo.fechacambio = item.fechacambio;
                    itemprorrateo.fechacreacion = item.fechacreacion;
                    itemprorrateo.idactividad = item.idactividad;
                    itemprorrateo.idempresa = item.idempresa;
                    itemprorrateo.idplantilla = item.idplantilla;
                    itemprorrateo.idplantilladetalle = item.idplantilladetalle;
                    itemprorrateo.idusuario = item.idusuario;
                    itemprorrateo.PlantillaCultivoCabecera = item.PlantillaCultivoCabecera;
                    itemprorrateo.TablaActividades = item.TablaActividades;

                    decimal costo = 0;
                    var detallecultivo = db.CultivoDetalle.Where(p => p.idactividad == item.idactividad);
                    foreach (var itemdetalle in detallecultivo)
                    {
                        if (itemdetalle.costo != null)
                        {
                            costo = costo + Convert.ToDecimal(itemdetalle.costo);
                        }

                    }

                    itemprorrateo.costo = costo;
                    index.Add(itemprorrateo);
                }
                viewModel.Items = index;
                viewModel.Pager = pager;
            }
            else
            {
                var pager = new Pager(db.PlantillaCultivoDetalle.Where(c => c.TablaActividades.descripcion.Contains(Search)).Count(), page);
                List<indexprorrateoVM> index = new List<indexprorrateoVM>();
                List<PlantillaCultivoDetalle> list = new List<PlantillaCultivoDetalle>();
                list = db.PlantillaCultivoDetalle.Include(p => p.PlantillaCultivoCabecera).Include(p => p.TablaActividades).Where(c => c.TablaActividades.descripcion.Contains(Search))
                        .OrderBy(c => c.idplantilladetalle)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;
                foreach (var item in list)
                {
                    indexprorrateoVM itemprorrateo = new indexprorrateoVM();
                    itemprorrateo.cantidad = item.cantidad;
                    itemprorrateo.fechacambio = item.fechacambio;
                    itemprorrateo.fechacreacion = item.fechacreacion;
                    itemprorrateo.idactividad = item.idactividad;
                    itemprorrateo.idempresa = item.idempresa;
                    itemprorrateo.idplantilla = item.idplantilla;
                    itemprorrateo.idplantilladetalle = item.idplantilladetalle;
                    itemprorrateo.idusuario = item.idusuario;
                    itemprorrateo.PlantillaCultivoCabecera = item.PlantillaCultivoCabecera;
                    itemprorrateo.TablaActividades = item.TablaActividades;

                    decimal costo = 0;
                    var detallecultivo = db.CultivoDetalle.Where(p => p.idactividad == item.idactividad);
                    foreach (var itemdetalle in detallecultivo)
                    {
                        if (itemdetalle.costo != null)
                        {
                            costo = costo + Convert.ToDecimal(itemdetalle.costo);
                        }

                    }
                    itemprorrateo.costo = costo;
                    index.Add(itemprorrateo);
                }
                viewModel.Items = index;
                @ViewBag.Search = Search;
            }
            return View(viewModel);
        }



        // GET: ActividadesProrrateo/Details/5
        public ActionResult Details(int? idplantilladetalle, int idactividad, DateTime fechaingreso)
        {

            if (idplantilladetalle == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            PlantillaCultivoDetalle plantillaCultivoDetalle = db.PlantillaCultivoDetalle.Find(idplantilladetalle);
            if (plantillaCultivoDetalle == null)
            {
                return HttpNotFound();
            }
            List<ActividadesProrrateoVM> list = new List<ActividadesProrrateoVM>();

            var detalle = db.PlantillaCultivoDetalle.Include(p => p.PlantillaCultivoCabecera).Where(p => p.idactividad == idactividad);
            foreach (var item in detalle)
            {
                ActividadesProrrateoVM prorrateo = new ActividadesProrrateoVM();
                var cultivo = db.Cultivo.Where(c => c.idplantilla == item.idplantilla);
                foreach (var itemcultivo in cultivo)
                {
                    var cdetalle = db.CultivoDetalle.Where(c => c.idcultivo == itemcultivo.idcultivo && c.idactividad == idactividad && c.fechallenado == fechaingreso);
                    foreach (var itemdetalle in cdetalle)
                    {
                        prorrateo.idlote = itemcultivo.idlote;
                        prorrateo.descripcion = itemcultivo.Lote.descripcion;
                        prorrateo.costo = itemdetalle.costo;
                        list.Add(prorrateo);
                    }
                }
            }
            ViewBag.idactividad = idactividad;
            ViewBag.list = list;
            return View(plantillaCultivoDetalle);
        }
        public ActionResult Detailsactividad(int? idplantilladetalle, int? idactividad)
        {

            var cultivodetalle = db.CultivoDetalle.Where(p => p.idactividad == idactividad).GroupBy(p => p.fechallenado);
            List<DetailsProrrateoVM> detalles = new List<DetailsProrrateoVM>();
            foreach (var itemcultivodetalle in cultivodetalle)
            {
                DetailsProrrateoVM detalle = new DetailsProrrateoVM();
                detalle.fechaingreso = itemcultivodetalle.Key;
                double monto = 0;
                foreach (var itemgrupo in itemcultivodetalle)
                {
                    detalle.DescripcionActividad = itemgrupo.TablaActividades.descripcion;
                    monto = monto + Convert.ToDouble(itemgrupo.costo);
                }
                detalle.monto = monto;
                detalles.Add(detalle);
            }
            CultivoDetalle cultivoDetalle = db.CultivoDetalle.Find(idplantilladetalle);
            ViewBag.idplantilladetalle = idplantilladetalle;
            ViewBag.descripcion = detalles[0].DescripcionActividad;
            ViewBag.idactividad = idactividad;
            ViewBag.list = detalles;
            return View(cultivoDetalle);
        }
        // GET: ActividadesProrrateo/Create
        public ActionResult Create(int? idactividad)
        {
            List<SelectListItem> list = new List<SelectListItem>();
            var plantilladetalle = db.PlantillaCultivoDetalle.Where(p => p.idactividad == idactividad);
            foreach (var itemplantilladetalle in plantilladetalle)
            {
                var plantillacabecera = db.PlantillaCultivoCabecera.Where(p => p.idplantilla == itemplantilladetalle.idplantilla);
                foreach (var itemplantillacabecera in plantillacabecera)
                {
                    var cultivo = db.Cultivo.Where(p => p.idplantilla == itemplantillacabecera.idplantilla).GroupBy(p => p.idlote).Select(p => p.FirstOrDefault());
                    foreach (var itemcultivo in cultivo)
                    {
                        var lote = db.Lote.Where(p => p.idlote == itemcultivo.idlote).OrderBy(p => p.descripcion);
                        foreach (var itemlote in lote)
                        {
                            list.Add(new SelectListItem() { Text = itemlote.descripcion, Value = itemlote.idlote.ToString() });
                        }
                    }
                }
            }
            ViewBag.idlote = new SelectList(list, "Value", "Text");
            ViewBag.idactividad = idactividad;
            //ViewBag.idactividad = new SelectList(db.TablaActividades, "idactividades", "idempresa");
            return View();
        }

        public ActionResult Createprorrateo(int? idactividad)
        {
            List<SelectListItem> list = new List<SelectListItem>();
            var plantilladetalle = db.PlantillaCultivoDetalle.Where(p => p.idactividad == idactividad);
            foreach (var itemplantilladetalle in plantilladetalle)
            {
                var plantillacabecera = db.PlantillaCultivoCabecera.Where(p => p.idplantilla == itemplantilladetalle.idplantilla);
                foreach (var itemplantillacabecera in plantillacabecera)
                {
                    var cultivo = db.Cultivo.Where(p => p.idplantilla == itemplantillacabecera.idplantilla).GroupBy(p => p.idlote).Select(p => p.FirstOrDefault());
                    foreach (var itemcultivo in cultivo)
                    {
                        var lote = db.Lote.Where(p => p.idlote == itemcultivo.idlote).OrderBy(p => p.descripcion);
                        foreach (var itemlote in lote)
                        {
                            list.Add(new SelectListItem() { Text = itemlote.descripcion, Value = itemlote.idlote.ToString() });
                        }
                    }
                }
            }
            ViewBag.idlote = new SelectList(list, "Value", "Text");
            CultivoDetalle cultivodetalle = new CultivoDetalle();

            //string idlote = Request.Form["idlote"].ToString();
            return View();
        }

        // POST: ActividadesProrrateo/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idplantilla,idplantilladetalle,idactividad,idusuario,cantidad,fechacreacion,fechacambio,idlote")] ActividadProrrateoVM plantillaCultivoDetalle)
        {

            return View(plantillaCultivoDetalle);
        }

        // GET: ActividadesProrrateo/Edit/5
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

        // POST: ActividadesProrrateo/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idplantilla,idplantilladetalle,idactividad,idusuario,cantidad,fechacreacion,fechacambio")] PlantillaCultivoDetalle plantillaCultivoDetalle)
        {
            var actividad = db.TablaActividades.Where(p => p.idactividades == plantillaCultivoDetalle.idactividad);
            TablaActividades actividadedit = null;
            foreach (var itemactividad in actividad)
            {
                actividadedit = itemactividad;
            }
            actividadedit.costo1 = plantillaCultivoDetalle.cantidad;
            if (ModelState.IsValid)
            {
                db.Entry(plantillaCultivoDetalle).State = EntityState.Modified;

                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.idplantilla = new SelectList(db.PlantillaCultivoCabecera, "idplantilla", "idempresa", plantillaCultivoDetalle.idplantilla);
            ViewBag.idactividad = new SelectList(db.TablaActividades, "idactividades", "idempresa", plantillaCultivoDetalle.idactividad);
            return View(plantillaCultivoDetalle);
        }

        // GET: ActividadesProrrateo/Delete/5
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

        // POST: ActividadesProrrateo/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            PlantillaCultivoDetalle plantillaCultivoDetalle = db.PlantillaCultivoDetalle.Find(id);
            db.PlantillaCultivoDetalle.Remove(plantillaCultivoDetalle);
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
