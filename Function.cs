using Google.Cloud.Functions.Framework;
using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;
using GrapeCity.Documents.Excel;
using System.IO;

namespace diodocs_gcfunctions
{
    public class Function : IHttpFunction
    {
        /// <summary>
        /// Logic for your function goes here.
        /// </summary>
        /// <param name="context">The HTTP context, containing the request and the response.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        public async Task HandleAsync(HttpContext context)
        {
            // await context.Response.WriteAsync("Hello, Functions Framework.");

            HttpRequest request = context.Request;
            string name = ((string)request.Query["name"]) ?? "world";

            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Range["A1"].Value = $"Hello, {name}!!";

            byte[] output;

            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFileFormat.Xlsx);
                output = ms.ToArray();
            }

            context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            context.Response.Headers.Add("Content-Disposition", "attachment;filename=Result.xlsx");
            await context.Response.Body.WriteAsync(output);
        }
    }
}
