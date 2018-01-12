﻿using System;
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

	public class CUENTAS_CONTABLESIndexViewModel
    {
		public List<CUENTAS_CONTABLES> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class CuentasContablesController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();

		
        // GET: CuentasContables
        public ActionResult Index(int? page, String Search)
        {
		//
			var viewModel = new CUENTAS_CONTABLESIndexViewModel();            

            //if (Search == null || Search.Equals(""))
            //{
				var pager = new Pager(db.CUENTAS_CONTABLES.Count(), page);
                viewModel.Items = db.CUENTAS_CONTABLES
                        .OrderBy(c => c.idcuenta)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;
    //        }
    //        else
    //        {
				//var pager = new Pager(db.CUENTAS_CONTABLES.Where(c => c.AgregarVariableAbuscar.Contains(Search)).Count(), page);
    //            viewModel.Items = db.CUENTAS_CONTABLES.Where(c => c.AgregarVariableAbuscar.Contains(Search))
    //                    .OrderBy(c => c.CUENTAS_CONTABLESID)
    //                    .Skip((pager.CurrentPage - 1) * pager.PageSize)
    //                    .Take(pager.PageSize).ToList();
				//viewModel.Pager = pager;
				//@ViewBag.Search = Search;
    //        }
            return View(viewModel);
        }



        // GET: CuentasContables/Details/5
        public ActionResult Details(string id, string id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            CUENTAS_CONTABLES cUENTAS_CONTABLES = db.CUENTAS_CONTABLES.Find(id, id2);
            if (cUENTAS_CONTABLES == null)
            {
                return HttpNotFound();
            }
            return View(cUENTAS_CONTABLES);
        }

        // GET: CuentasContables/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: CuentasContables/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "idempresa,idcuenta,idusuario,cuenta,cuentaanterior,descripcion,estado,fechacambio,resultado,ctacargo,ctaabono,porcentajecargo,porcentajeabono,idespecial,idrendicion,iddetalle,idunegocio,idgrupo")] CUENTAS_CONTABLES cUENTAS_CONTABLES)
        {
            cUENTAS_CONTABLES.idempresa = "01";
            List<CUENTAS_CONTABLES> cc = db.CUENTAS_CONTABLES.ToList();
            if (cc.Count == 0) { cUENTAS_CONTABLES.idcuenta = "11000001"; }
            else { cUENTAS_CONTABLES.idcuenta = "11"+getidcuenta(Convert.ToInt32(cc.Last().idcuenta.Substring(2))); }
            cUENTAS_CONTABLES.idusuario = "0001";
            cUENTAS_CONTABLES.cuentaanterior = "";
            cUENTAS_CONTABLES.estado = "1";
            cUENTAS_CONTABLES.resultado = "0";
            cUENTAS_CONTABLES.ctacargo = "";
            cUENTAS_CONTABLES.ctaabono = "";
            cUENTAS_CONTABLES.porcentajecargo = 0;
            cUENTAS_CONTABLES.porcentajeabono = 0;
            cUENTAS_CONTABLES.idespecial = "280000";
            cUENTAS_CONTABLES.idrendicion = "0";
            cUENTAS_CONTABLES.iddetalle = "0";
            cUENTAS_CONTABLES.idunegocio = "0";
            cUENTAS_CONTABLES.idgrupo = "630000";
            if (ModelState.IsValid)
            {
                //cUENTAS_CONTABLES.Creado = DateTime.Now;
                //cUENTAS_CONTABLES.Modificado = DateTime.Now;
                db.CUENTAS_CONTABLES.Add(cUENTAS_CONTABLES);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(cUENTAS_CONTABLES);
        }

        // GET: CuentasContables/Edit/5
        public ActionResult Edit(string id, string id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            CUENTAS_CONTABLES cUENTAS_CONTABLES = db.CUENTAS_CONTABLES.Find(id, id2);
            if (cUENTAS_CONTABLES == null)
            {
                return HttpNotFound();
            }
            return View(cUENTAS_CONTABLES);
        }

        // POST: CuentasContables/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "idempresa,idcuenta,idusuario,cuenta,cuentaanterior,descripcion,estado,fechacambio,resultado,ctacargo,ctaabono,porcentajecargo,porcentajeabono,idespecial,idrendicion,iddetalle,idunegocio,idgrupo")] CUENTAS_CONTABLES cUENTAS_CONTABLES)
        {
            cUENTAS_CONTABLES.cuentaanterior = "";
            cUENTAS_CONTABLES.ctacargo = "";
            cUENTAS_CONTABLES.ctaabono = "";
            if (ModelState.IsValid)
            {
                db.Entry(cUENTAS_CONTABLES).State = EntityState.Modified;
				//cUENTAS_CONTABLES.Modificado = DateTime.Now;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(cUENTAS_CONTABLES);
        }

        private string getidcuenta(int v)
        {
            v = v + 1;
            int digitos = Convert.ToString(v).Length;
            if (digitos == 1) { return "00000" + v; }
            else if (digitos == 2) { return "0000" + v; }
            else if (digitos == 3) { return "000" + v; }
            else if (digitos == 4) { return "00" + v; }
            else if (digitos == 5) { return "0" + v; }
            else if (digitos == 6) { return v.ToString(); }
            else { return Convert.ToString(v); }
        }

        // GET: CuentasContables/Delete/5
        public ActionResult Delete(string id, string id2)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            CUENTAS_CONTABLES cUENTAS_CONTABLES = db.CUENTAS_CONTABLES.Find(id, id2);
            if (cUENTAS_CONTABLES == null)
            {
                return HttpNotFound();
            }
            return View(cUENTAS_CONTABLES);
        }

        // POST: CuentasContables/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id, string id2)
        {
            CUENTAS_CONTABLES cUENTAS_CONTABLES = db.CUENTAS_CONTABLES.Find(id, id2);
            db.CUENTAS_CONTABLES.Remove(cUENTAS_CONTABLES);
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
                List<CUENTAS_CONTABLES> list = db.CUENTAS_CONTABLES.ToList();
                int pos = 4;
								ws.Cells[pos, 4].Value = "idusuario";
									ws.Cells[pos, 5].Value = "cuenta";
									ws.Cells[pos, 6].Value = "cuentaanterior";
									ws.Cells[pos, 7].Value = "descripcion";
									ws.Cells[pos, 8].Value = "estado";
									ws.Cells[pos, 9].Value = "fechacambio";
									ws.Cells[pos, 10].Value = "resultado";
									ws.Cells[pos, 11].Value = "ctacargo";
									ws.Cells[pos, 12].Value = "ctaabono";
									ws.Cells[pos, 13].Value = "porcentajecargo";
									ws.Cells[pos, 14].Value = "porcentajeabono";
									ws.Cells[pos, 15].Value = "idespecial";
									ws.Cells[pos, 16].Value = "idrendicion";
									ws.Cells[pos, 17].Value = "iddetalle";
									ws.Cells[pos, 18].Value = "idunegocio";
									ws.Cells[pos, 19].Value = "idgrupo";
					
                foreach (var item in list)
                {
                    pos++;
								ws.Cells[pos, 4].Value = item.idusuario == null ? "" : item.idusuario.ToString();				
									ws.Cells[pos, 5].Value = item.cuenta == null ? "" : item.cuenta.ToString();				
									ws.Cells[pos, 6].Value = item.cuentaanterior == null ? "" : item.cuentaanterior.ToString();				
									ws.Cells[pos, 7].Value = item.descripcion == null ? "" : item.descripcion.ToString();				
									ws.Cells[pos, 8].Value = item.estado == null ? "" : item.estado.ToString();				
									ws.Cells[pos, 9].Value = item.fechacambio == null ? "" : item.fechacambio.ToString();				
									ws.Cells[pos, 10].Value = item.resultado == null ? "" : item.resultado.ToString();				
									ws.Cells[pos, 11].Value = item.ctacargo == null ? "" : item.ctacargo.ToString();				
									ws.Cells[pos, 12].Value = item.ctaabono == null ? "" : item.ctaabono.ToString();				
									ws.Cells[pos, 13].Value = item.porcentajecargo == null ? "" : item.porcentajecargo.ToString();				
									ws.Cells[pos, 14].Value = item.porcentajeabono == null ? "" : item.porcentajeabono.ToString();				
									ws.Cells[pos, 15].Value = item.idespecial == null ? "" : item.idespecial.ToString();				
									ws.Cells[pos, 16].Value = item.idrendicion == null ? "" : item.idrendicion.ToString();				
									ws.Cells[pos, 17].Value = item.iddetalle == null ? "" : item.iddetalle.ToString();				
									ws.Cells[pos, 18].Value = item.idunegocio == null ? "" : item.idunegocio.ToString();				
									ws.Cells[pos, 19].Value = item.idgrupo == null ? "" : item.idgrupo.ToString();				
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

			
            var table = new PdfPTable(16);

            var boldTableFont = FontFactory.GetFont("Arial", 10, Font.BOLD);
            var bodyFont = FontFactory.GetFont("Arial", 10, Font.NORMAL);

							table.AddCell(new Phrase("idusuario", boldTableFont));
									table.AddCell(new Phrase("cuenta", boldTableFont));
									table.AddCell(new Phrase("cuentaanterior", boldTableFont));
									table.AddCell(new Phrase("descripcion", boldTableFont));
									table.AddCell(new Phrase("estado", boldTableFont));
									table.AddCell(new Phrase("fechacambio", boldTableFont));
									table.AddCell(new Phrase("resultado", boldTableFont));
									table.AddCell(new Phrase("ctacargo", boldTableFont));
									table.AddCell(new Phrase("ctaabono", boldTableFont));
									table.AddCell(new Phrase("porcentajecargo", boldTableFont));
									table.AddCell(new Phrase("porcentajeabono", boldTableFont));
									table.AddCell(new Phrase("idespecial", boldTableFont));
									table.AddCell(new Phrase("idrendicion", boldTableFont));
									table.AddCell(new Phrase("iddetalle", boldTableFont));
									table.AddCell(new Phrase("idunegocio", boldTableFont));
									table.AddCell(new Phrase("idgrupo", boldTableFont));
									              
//
            List<CUENTAS_CONTABLES> list = db.CUENTAS_CONTABLES.ToList();

			foreach (var item in list)
                {
                    
								table.AddCell(new Phrase(item.idusuario == null ? "" : item.idusuario.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.cuenta == null ? "" : item.cuenta.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.cuentaanterior == null ? "" : item.cuentaanterior.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.descripcion == null ? "" : item.descripcion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.estado == null ? "" : item.estado.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.fechacambio == null ? "" : item.fechacambio.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.resultado == null ? "" : item.resultado.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.ctacargo == null ? "" : item.ctacargo.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.ctaabono == null ? "" : item.ctaabono.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.porcentajecargo == null ? "" : item.porcentajecargo.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.porcentajeabono == null ? "" : item.porcentajeabono.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.idespecial == null ? "" : item.idespecial.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.idrendicion == null ? "" : item.idrendicion.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.iddetalle == null ? "" : item.iddetalle.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.idunegocio == null ? "" : item.idunegocio.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.idgrupo == null ? "" : item.idgrupo.ToString(), bodyFont));			
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
