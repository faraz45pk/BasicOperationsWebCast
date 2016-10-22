using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

using System.Web.Mvc;

namespace BasicOperationsWebCastWeb.Controllers
{
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            User spUser = null;

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    spUser = clientContext.Web.CurrentUser;

                    clientContext.Load(spUser, user => user.Title);

                    clientContext.ExecuteQuery();

                    ViewBag.UserName = spUser.Title;
                }
            }

            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }


        public ActionResult DoStuff()
        {
            var authManager = new AuthenticationManager();

            var context = 
                authManager.GetAppOnlyAuthenticatedContext
                (   "https://mydev2016.sharepoint.com/",
                    ConfigurationManager.AppSettings["ClientId"],
                    ConfigurationManager.AppSettings["ClientSecret"]);

            using (context)
            {


                var list2 = 
                    context.Web.CreateList(
                        ListTemplateType.GenericList,
                        "WithPnP", false);

                var view2 =
                    list2.CreateView("New View 2",
                                Microsoft.SharePoint.Client.ViewType.Html,
                                new string[] { "Title", "Modified" }, 10, true);


                context.Web.SetPropertyBagValue("ourwebkey2", Guid.NewGuid().ToString());


                context.Web.AddIndexedPropertyBagKey("ourwebkey2");







                #region CommentCode
                /*
                // Create a list
                var list = context.Web.Lists.Add(
                    new ListCreationInformation {
                        Title = "NoPnP2",
                        Description = "A list created for testing.",
                        Url = "Lists/NoPnP2",
                        TemplateType = (int)ListTemplateType.GenericList,
                        TemplateFeatureId = 
                        new Guid("00BFEA71-DE22-43B2-A848-C05709900100")
            });

            context.Load(list);
            context.ExecuteQueryRetry();

            //create a view on list
            ViewCreationInformation
                viewCreationInformation = new ViewCreationInformation();

            viewCreationInformation.Title = "New View";
            viewCreationInformation.ViewTypeKind =
                Microsoft.SharePoint.Client.ViewType.Html;
            viewCreationInformation.RowLimit = 10;
            viewCreationInformation.ViewFields = new string[] { "Title","Modified"};
            viewCreationInformation.PersonalView = false;
            viewCreationInformation.SetAsDefaultView = true;
            viewCreationInformation.Paged = false;

            var view = 
                list.Views.Add(viewCreationInformation);


            list.Context.Load(view);
            list.Context.ExecuteQueryRetry();



            // set web property

            var props = context.Web.AllProperties;
            context.Web.Context.Load(props);
            context.Web.Context.ExecuteQueryRetry();

            props["ourwebkey"] = Guid.NewGuid().ToString();
            context.Web.Update();

            context.Web.Context.ExecuteQueryRetry();
                */
                #endregion
            }

            return new HttpStatusCodeResult(System.Net.HttpStatusCode.OK);
        }

    }
}
