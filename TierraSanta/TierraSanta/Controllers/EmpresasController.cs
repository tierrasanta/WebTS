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

    public class EmpresaIndexViewModel
    {
        public List<Empresa> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class EmpresasController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();


        // GET: Empresas
        public ActionResult Index(int? page, String Search)
        {
            //
            var viewModel = new EmpresaIndexViewModel();


            var pager = new Pager(db.Empresa.Count(), page);
            viewModel.Items = db.Empresa
                    .OrderBy(c => c.idempresa)
                    .Skip((pager.CurrentPage - 1) * pager.PageSize)
                    .Take(pager.PageSize).ToList();
            viewModel.Pager = pager;

            return View(viewModel);
        }



        // GET: Empresas/Details/5
        public ActionResult Details(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Empresa empresa = db.Empresa.Find(id);
            if (empresa == null)
            {
                return HttpNotFound();
            }
            return View(empresa);
        }

        // GET: Empresas/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Empresas/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idubicacion,idpersoneria,idregimen,idtipo,idmoneda,idcomision,idusuario,control,parametros,ruc,razonsocial,direccion,Abreviatura,estado,logotipo,fondo")] Empresa empresa)
        {
            List<Empresa> f = db.Empresa.ToList();
            if (f.Count == 0) { empresa.idempresa = "01"; }
            else { empresa.idempresa = getidempresa(Convert.ToInt32(f.Last().idempresa)); }
            empresa.idubicacion = "01040101";
            empresa.idpersoneria = "200002";
            empresa.idregimen = "100002";
            empresa.idtipo = "040001";
            empresa.idmoneda = "020001";
            empresa.idcomision = "120001";
            empresa.idusuario = "0001";
            empresa.control = "11111111111110000000";
            empresa.parametros = "0002210N0101006101000012010000";
            empresa.estado = "1";
            empresa.logotipo = "";
            empresa.fondo = "";            

            if (ModelState.IsValid)
            {
                db.Empresa.Add(empresa);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(empresa);
        }
        private string getidempresa(int v)
        {
            v = v + 1;
            int digitos = Convert.ToString(v).Length;
            if (digitos == 1) { return "0" + v; }
            else { return Convert.ToString(v); }
        }

        // GET: Empresas/Edit/5
        public ActionResult Edit(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Empresa empresa = db.Empresa.Find(id);
            if (empresa == null)
            {
                return HttpNotFound();
            }
            return View(empresa);
        }

        // POST: Empresas/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idubicacion,idpersoneria,idregimen,idtipo,idmoneda,idcomision,idusuario,control,parametros,ruc,razonsocial,direccion,Abreviatura,fechainicio,estado,fechacambio,logotipo,fondo")] Empresa empresa)
        {
            empresa.fechacambio = DateTime.Now;
            empresa.logotipo = "";
            empresa.fondo = "";
            if (ModelState.IsValid)
            {
                db.Entry(empresa).State = EntityState.Modified;                
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(empresa);
        }

        // GET: Empresas/Delete/5
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Empresa empresa = db.Empresa.Find(id);
            if (empresa == null)
            {
                return HttpNotFound();
            }
            return View(empresa);
        }

        // POST: Empresas/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            Empresa empresa = db.Empresa.Find(id);
            db.Empresa.Remove(empresa);
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
                List<Empresa> list = db.Empresa.ToList();
                int pos = 4;
                ws.Cells[pos, 3].Value = "idubicacion";
                ws.Cells[pos, 4].Value = "idpersoneria";
                ws.Cells[pos, 5].Value = "idregimen";
                ws.Cells[pos, 6].Value = "idtipo";
                ws.Cells[pos, 7].Value = "idmoneda";
                ws.Cells[pos, 8].Value = "idcomision";
                ws.Cells[pos, 9].Value = "idusuario";
                ws.Cells[pos, 10].Value = "control";
                ws.Cells[pos, 11].Value = "parametros";
                ws.Cells[pos, 12].Value = "ruc";
                ws.Cells[pos, 13].Value = "razonsocial";
                ws.Cells[pos, 14].Value = "direccion";
                ws.Cells[pos, 15].Value = "Abreviatura";
                ws.Cells[pos, 16].Value = "fechainicio";
                ws.Cells[pos, 17].Value = "estado";
                ws.Cells[pos, 18].Value = "fechacambio";
                ws.Cells[pos, 19].Value = "logotipo";
                ws.Cells[pos, 20].Value = "fondo";

                foreach (var item in list)
                {
                    pos++;
                    ws.Cells[pos, 3].Value = item.idubicacion == null ? "" : item.idubicacion.ToString();
                    ws.Cells[pos, 4].Value = item.idpersoneria == null ? "" : item.idpersoneria.ToString();
                    ws.Cells[pos, 5].Value = item.idregimen == null ? "" : item.idregimen.ToString();
                    ws.Cells[pos, 6].Value = item.idtipo == null ? "" : item.idtipo.ToString();
                    ws.Cells[pos, 7].Value = item.idmoneda == null ? "" : item.idmoneda.ToString();
                    ws.Cells[pos, 8].Value = item.idcomision == null ? "" : item.idcomision.ToString();
                    ws.Cells[pos, 9].Value = item.idusuario == null ? "" : item.idusuario.ToString();
                    ws.Cells[pos, 10].Value = item.control == null ? "" : item.control.ToString();
                    ws.Cells[pos, 11].Value = item.parametros == null ? "" : item.parametros.ToString();
                    ws.Cells[pos, 12].Value = item.ruc == null ? "" : item.ruc.ToString();
                    ws.Cells[pos, 13].Value = item.razonsocial == null ? "" : item.razonsocial.ToString();
                    ws.Cells[pos, 14].Value = item.direccion == null ? "" : item.direccion.ToString();
                    ws.Cells[pos, 15].Value = item.Abreviatura == null ? "" : item.Abreviatura.ToString();
                    ws.Cells[pos, 16].Value = item.fechainicio == null ? "" : item.fechainicio.ToString();
                    ws.Cells[pos, 17].Value = item.estado == null ? "" : item.estado.ToString();
                    ws.Cells[pos, 18].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();
                    ws.Cells[pos, 19].Value = item.logotipo == null ? "" : item.logotipo.ToString();
                    ws.Cells[pos, 20].Value = item.fondo == null ? "" : item.fondo.ToString();
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


            var table = new PdfPTable(18);

            var boldTableFont = FontFactory.GetFont("Arial", 10, Font.BOLD);
            var bodyFont = FontFactory.GetFont("Arial", 10, Font.NORMAL);

            table.AddCell(new Phrase("idubicacion", boldTableFont));
            table.AddCell(new Phrase("idpersoneria", boldTableFont));
            table.AddCell(new Phrase("idregimen", boldTableFont));
            table.AddCell(new Phrase("idtipo", boldTableFont));
            table.AddCell(new Phrase("idmoneda", boldTableFont));
            table.AddCell(new Phrase("idcomision", boldTableFont));
            table.AddCell(new Phrase("idusuario", boldTableFont));
            table.AddCell(new Phrase("control", boldTableFont));
            table.AddCell(new Phrase("parametros", boldTableFont));
            table.AddCell(new Phrase("ruc", boldTableFont));
            table.AddCell(new Phrase("razonsocial", boldTableFont));
            table.AddCell(new Phrase("direccion", boldTableFont));
            table.AddCell(new Phrase("Abreviatura", boldTableFont));
            table.AddCell(new Phrase("fechainicio", boldTableFont));
            table.AddCell(new Phrase("estado", boldTableFont));
            table.AddCell(new Phrase("fechacambio", boldTableFont));
            table.AddCell(new Phrase("logotipo", boldTableFont));
            table.AddCell(new Phrase("fondo", boldTableFont));

            //
            List<Empresa> list = db.Empresa.ToList();

            foreach (var item in list)
            {

                table.AddCell(new Phrase(item.idubicacion == null ? "" : item.idubicacion.ToString(), bodyFont));
                table.AddCell(new Phrase(item.idpersoneria == null ? "" : item.idpersoneria.ToString(), bodyFont));
                table.AddCell(new Phrase(item.idregimen == null ? "" : item.idregimen.ToString(), bodyFont));
                table.AddCell(new Phrase(item.idtipo == null ? "" : item.idtipo.ToString(), bodyFont));
                table.AddCell(new Phrase(item.idmoneda == null ? "" : item.idmoneda.ToString(), bodyFont));
                table.AddCell(new Phrase(item.idcomision == null ? "" : item.idcomision.ToString(), bodyFont));
                table.AddCell(new Phrase(item.idusuario == null ? "" : item.idusuario.ToString(), bodyFont));
                table.AddCell(new Phrase(item.control == null ? "" : item.control.ToString(), bodyFont));
                table.AddCell(new Phrase(item.parametros == null ? "" : item.parametros.ToString(), bodyFont));
                table.AddCell(new Phrase(item.ruc == null ? "" : item.ruc.ToString(), bodyFont));
                table.AddCell(new Phrase(item.razonsocial == null ? "" : item.razonsocial.ToString(), bodyFont));
                table.AddCell(new Phrase(item.direccion == null ? "" : item.direccion.ToString(), bodyFont));
                table.AddCell(new Phrase(item.Abreviatura == null ? "" : item.Abreviatura.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechainicio == null ? "" : item.fechainicio.ToString(), bodyFont));
                table.AddCell(new Phrase(item.estado == null ? "" : item.estado.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));
                table.AddCell(new Phrase(item.logotipo == null ? "" : item.logotipo.ToString(), bodyFont));
                table.AddCell(new Phrase(item.fondo == null ? "" : item.fondo.ToString(), bodyFont));
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
