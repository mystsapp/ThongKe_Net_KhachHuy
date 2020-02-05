using System.Web.Optimization;

namespace ThongKe.Web
{
    public class BundleConfig
    {
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
                        "~/Scripts/jquery - 1.10.2.min.js",
                        "~/Scripts/jquery.unobtrusive-ajax.min.js",
                        "~/Scripts/jquery-{version}.js"));

            bundles.Add(new ScriptBundle("~/bundles/jqueryval").Include(
                        //"~/Scripts/jquery.validate.min.js",
                        "~/Scripts/jquery.validate.unobtrusive.js",
                        "~/Scripts/jquery.validate*"));

            // Use the development version of Modernizr to develop with and learn from. Then, when you're
            // ready for production, use the build tool at http://modernizr.com to pick only the tests you need.
            bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
                        "~/Scripts/modernizr-*"));

            bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
                      //"~/Content/boostrap/vendor/jquery/jquery.min.js",
                      "~/Scripts/bootstrap.js",
                      "~/Scripts/respond.js"));
            bundles.Add(new ScriptBundle("~/bundles/vendor").Include(
                     "~/Content/boostrap/vendor/metisMenu/metisMenu.min.js",
                     "~/Content/boostrap/vendor/raphael/raphael.min.js",
                     //"~/Assets/Admin/Js/Plugins/ckfinder/ckfinder.js",
                     //"~/Assets/Admin/Js/Plugins/ckeditor/ckeditor.js",
                     "~/Content/boostrap/dist/js/sb-admin-2.js",
                     "~/Content/boxover/js/boxover.js"));

            bundles.Add(new StyleBundle("~/Content/css").Include(
                      //"~/Content/bootstrap.css",
                      "~/Assets/Admin/vendor/bootstrap/css/bootstrap.css",
                      "~/Assets/Admin/js/User/jquery-ui/jquery-ui.css",
                      "~/Content/boostrap/vendor/metisMenu/metisMenu.min.css",
                      "~/Content/boostrap/dist/css/sb-admin-2.css",
                      "~/Content/boxover/css/boxover.css",
                      "~/Content/boostrap/vendor/morrisjs/morris.css",
                      "~/Assets/Admin/fontawesome-free-5.1.0-web/css/all.css",
                      "~/Content/Site.css"));
        }

        //public static void RegisterBundles(BundleCollection bundles)
        //{
        //    bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
        //                "~/Scripts/jquery - 1.10.2.min.js",
        //                "~/Scripts/jquery.unobtrusive-ajax.min.js",
        //                "~/Scripts/jquery-{version}.js"));

        //    bundles.Add(new ScriptBundle("~/bundles/jqueryval").Include(
        //                //"~/Scripts/jquery.validate.min.js",
        //                "~/Scripts/jquery.validate.unobtrusive.js",
        //                "~/Scripts/jquery.validate*"));

        //    // Use the development version of Modernizr to develop with and learn from. Then, when you're
        //    // ready for production, use the build tool at http://modernizr.com to pick only the tests you need.
        //    bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
        //                "~/Scripts/modernizr-*"));

        //    bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
        //              //"~/Content/boostrap/vendor/jquery/jquery.min.js",
        //              "~/Scripts/bootstrap.js",
        //              "~/Scripts/respond.js"));
        //    bundles.Add(new ScriptBundle("~/bundles/vendor").Include(
        //             "~/Content/boostrap/vendor/metisMenu/metisMenu.min.js",
        //             "~/Content/boostrap/vendor/raphael/raphael.min.js",
        //             "~/Assets/Admin/Js/Plugins/ckfinder/ckfinder.js",
        //             "~/Assets/Admin/Js/Plugins/ckeditor/ckeditor.js",
        //             "~/Content/boostrap/dist/js/sb-admin-2.js",
        //             "~/Content/boxover/js/boxover.js"));

        //    bundles.Add(new StyleBundle("~/Content/css").Include(
        //              "~/Content/bootstrap.css",
        //              "~/Assets/Admin/js/User/jquery-ui/jquery-ui.css",
        //              "~/Content/boostrap/vendor/metisMenu/metisMenu.min.css",
        //              "~/Content/boostrap/dist/css/sb-admin-2.css",
        //              "~/Content/boxover/css/boxover.css",
        //              "~/Content/boostrap/vendor/morrisjs/morris.css",
        //              "~/Content/boostrap/vendor/font-awesome/css/font-awesome.min.css",
        //              "~/Assets/Admin/fontawesome-free-5.1.0-web/css/all.css",
        //              "~/Content/Site.css"));
        //}
    }
}