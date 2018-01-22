using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebTS2.Models;

namespace WebTS2.Controllers
{
    public class HomeController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Login()
        {            
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Login(String login, String clave)
        {
            Usuario usuario = db.Usuario.Where(u => u.Login.Equals(login) && u.Clave.Equals(clave) && u.Estado == true).FirstOrDefault();
            if (usuario != null)
            {
                Session["Usuario"] = usuario;
                return RedirectToAction("Index", "Home");
            }
            ViewBag.Message = "Usuario o clave no válidas";
            return View();
        }

        public ActionResult Logout()
        {
            Session["Usuario"] = null;
            Session.Clear();
            return RedirectToAction("Login", "Home", new { Area = "" });

        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}