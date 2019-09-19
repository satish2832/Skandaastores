using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Caching;
using System.Web.Mvc;
using StoreSite.Helpers;
using StoreSite.Models;


namespace StoreSite.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            var lsProducts = GetProducts();
            ViewBag.Products = lsProducts;
            return View();
        }
        private List<Product> GetProducts()
        {
            var lsProducts = HttpContext.Cache.Get("ProductDetails") as List<Product>;
            if (lsProducts == null)
            {
                var helper = new DataHelper();
                lsProducts = helper.GetProducts();
                HttpContext.Cache.Insert("ProductDetails", lsProducts, null, DateTime.Now.AddMinutes(60), Cache.NoSlidingExpiration);
            }
            return lsProducts;
        }
        private Product GetProductsByCode(List<Product> products,string code)
        {           
            return products.Where(x=>x.Code==code).FirstOrDefault();
        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";
            return View();
        }
        public ActionResult ProductDetails(string code,string description)
        {
            try
            {
                var lsProducts = GetProducts();
                ViewBag.Products = lsProducts;
                var product = GetProductsByCode(lsProducts,code);
                ViewBag.product = product;
                ViewBag.code = code;
                ViewBag.description = description;
                if (product!=null)
                {
                    ViewBag.oldPrice = product.OldValue;
                    ViewBag.newPrice =product.NewValue;
                }
                
                if (product.Colors.Split(',').Length > 1)
                {
                    ViewBag.colors = product.Colors.ToLower().Split(',');
                }
                return View();
            }
            catch
            {
                return View("Index");
            }
        }
        [HttpPost]
        public ActionResult SaveEnquiry(CustomerEnquiry enquiry)
        {
            var helper = new DataHelper();
            helper.AddCustomerEnquiry(enquiry);
            return Json(new { result = true }, JsonRequestBehavior.AllowGet);
        }
        public ActionResult DownloadEnquiry()
        {
            string filePath = System.Web.HttpContext.Current.Server.MapPath("~/DataSources/Customer_Enquiry.xlsx");
            return File(filePath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }
        public ActionResult UploadProductDetails()
        {
            var fileProductDetails = Request.Files.Get(0) as HttpPostedFileBase;           
            string targetFolder = System.Web.HttpContext.Current.Server.MapPath("~/DataSources");
            string targetPath = System.IO.Path.Combine(targetFolder, fileProductDetails.FileName);
            fileProductDetails.SaveAs(targetPath);
            return Json(new { success=true},JsonRequestBehavior.AllowGet);
        }
        public ActionResult UploadProductImages()
        {
            var fileProductImages = Request.Files;
            for(var index=0;index<fileProductImages.AllKeys.Count();index++)
            {
                var fileImage = Request.Files.Get(index) as HttpPostedFileBase;
                string targetFolder = System.Web.HttpContext.Current.Server.MapPath("~/Content/Site/images/product");
                string targetPath = System.IO.Path.Combine(targetFolder, fileImage.FileName);
                fileImage.SaveAs(targetPath);
            }
            return Json(new { success = true }, JsonRequestBehavior.AllowGet);
        }
        public ActionResult Contact()
        {
            return View();
        }
        public ActionResult Admin()
        {
            return View();
        }
    }
}