using Microsoft.AspNetCore.Mvc.Razor;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Text.Encodings.Web;
using System.Text.RegularExpressions;
using Rotativa.AspNetCore;
using DocumentFormat.OpenXml.InkML;

namespace SGID.Models.Services
{
    public class RazorPageAsPdf : AsPdfResultBase
    {
        private readonly IRazorViewEngine _razorViewEngine;
        private readonly ITempDataProvider _tempDataProvider;
        private readonly IRazorPageActivator _activator;
        private string _razorPageName { get; set; }
        private string Caminho { get; set; }
        public PageModel PageModel { get; set; }
        public RazorPageAsPdf(PageModel pageModel,string Path)
        {
            Caminho = Path;
            PageModel = pageModel;
            var httpContext = pageModel.HttpContext;
            _razorPageName = httpContext.Request.RouteValues["page"].ToString().Trim('/');
            if (string.IsNullOrEmpty(_razorPageName))
            {
                throw new ArgumentException("there's no such a 'page' in this context");
            }
            _razorViewEngine = httpContext.RequestServices.GetRequiredService<IRazorViewEngine>();
            _tempDataProvider = httpContext.RequestServices.GetRequiredService<ITempDataProvider>();
            _activator = httpContext.RequestServices.GetRequiredService<IRazorPageActivator>();
        }

        protected override string GetUrl(ActionContext context) => string.Empty;
        private ViewContext GetViewContext(ActionContext actionContext, IRazorPage page, StringWriter sw)
        {
            var view = new RazorView(_razorViewEngine, _activator, new List<IRazorPage>(), page, HtmlEncoder.Default, new DiagnosticListener(nameof(RazorPageAsPdf)));
            return new ViewContext(actionContext, view, PageModel.ViewData, PageModel.TempData, sw, new HtmlHelperOptions());
        }

        private async Task<string> RenderPageAsString(ActionContext actionContext)
        {
            using (var sw = new StringWriter())
            {
                var pageResult = _razorViewEngine.FindPage(actionContext, _razorPageName);
                if (pageResult.Page == null)
                {
                    throw new ArgumentNullException($"The page {_razorPageName} cannot be found.");
                }
                var viewContext = GetViewContext(actionContext, pageResult.Page, sw);
                var page = (Page)pageResult.Page;
                page.PageContext = PageModel.PageContext;
                page.ViewContext = viewContext;
                _activator.Activate(page, viewContext);
                await page.ExecuteAsync();
                return sw.ToString();
            }
        }

        protected override async Task<byte[]> CallTheDriver(ActionContext actionContext)
        {
            var html = await RenderPageAsString(actionContext);
            // copied from https://github.com/webgio/Rotativa.AspNetCore/blob/c907afa8c7dd6a565d307901741c336c429fc698/Rotativa.AspNetCore/ViewAsPdf.cs#L147-L151
            string baseUrl = string.Format("{0}://{1}", actionContext.HttpContext.Request.Scheme, actionContext.HttpContext.Request.Host);
            var htmlForWkhtml = Regex.Replace(html.ToString(), "<head>", string.Format("<head><base href=\"{0}\" />", baseUrl), RegexOptions.IgnoreCase);
            var teste = GetConvertOptions();
            byte[] fileContent = WkhtmltopdfDriver.ConvertHtml(Caminho,teste , htmlForWkhtml);
            return fileContent;
        }
    }
}
