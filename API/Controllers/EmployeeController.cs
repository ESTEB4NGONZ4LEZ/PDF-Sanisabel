
using Infrastructure;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class EmployeeController : ControllerBase
    {
        private readonly MainContext _context;
        public EmployeeController(MainContext context)
        {
            _context = context;
        }

        [HttpGet]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<ActionResult> GetEmployees()
        {            
            MemoryStream memoryStream = new();
            PdfWriter writer = new(memoryStream);
            PdfDocument pdf = new(writer);
            Document document = new (pdf, PageSize.A4.Rotate());
            document.SetMargins(0, 0, 0, 0);


            ImageData imageData = ImageDataFactory.Create("C:/Users/Esteban/Documents/PDFReport/img/img1.png");
            Image pagImg1 = new(imageData);
            pagImg1.SetHeight(10);
            pagImg1.SetWidth(10);

            // PAGINA 1
            
            // PAGINA 1 SECCION 1
            Div container1 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 30f / 100f - 4)
                                      .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4);

            // Parrafo 1
            Paragraph parrafo = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                               .SetMarginTop(20)
                                               .SetMarginLeft(20)
                                               .SetMarginBottom(5)
                                               .SetMarginRight(20)
                                               .SetFontSize(8)
                                               .Add("Esta es la primera línea del párrafo justificado.")
                                               .Add("Esta es la segunda línea del párrafo justificado.")
                                               .Add("Esta es la tercera línea del párrafo");

            container1.Add(parrafo);

            // Subtitulo 1 
            Paragraph subtitulo1 = new Paragraph().Add("SUBTITULO 1")
                                                  .SetFontColor(ColorConstants.WHITE)
                                                  .SetFontSize(7);

            Div subtituloContainer1 = new Div().SetWidth(90f)
                                               .Add(subtitulo1);

            Table tableSubtitulo1 = new Table(2).SetBackgroundColor(new DeviceRgb(53, 142, 54))
                                                .SetWidth(container1.GetWidth())
                                                .SetMarginTop(5)
                                                .SetMarginLeft(10)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(0);

            tableSubtitulo1.AddCell(new Cell().Add(pagImg1)
                                              .SetBorder(Border.NO_BORDER)
                                              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                                              .SetPaddingLeft(10));

            tableSubtitulo1.AddCell(new Cell().Add(subtituloContainer1)
                                              .SetBorder(Border.NO_BORDER))
                                              .SetTextAlignment(TextAlignment.LEFT);
        
            container1.Add(tableSubtitulo1);

            // Parrafo 2
            Paragraph parrafo2 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(8)
                                                .Add("Segundo parrafo separado por un enter justificado. Segundo parrafo separado por un enter justificado.")
                                                .Add("\n \n")
                                                .Add("Segundo parrafo separado por un enter justificado. Segundo parrafo separado por un enter justificado.");

            container1.Add(parrafo2);

            // Subtitulo 2
            Paragraph subtitulo2 = new Paragraph().Add("SUBTITULO 2")
                                                  .SetFontColor(ColorConstants.WHITE)
                                                  .SetFontSize(7);

            Div subtituloContainer2 = new Div().SetWidth(90f)
                                               .Add(subtitulo2);

            Table tableSubtitulo2 = new Table(2).SetBackgroundColor(new DeviceRgb(53, 142, 54))
                                                .SetWidth(container1.GetWidth())
                                                .SetMarginTop(5)
                                                .SetMarginLeft(10)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(0);

            tableSubtitulo2.AddCell(new Cell().Add(pagImg1)
                                              .SetBorder(Border.NO_BORDER)
                                              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                                              .SetPaddingLeft(10));

            tableSubtitulo2.AddCell(new Cell().Add(subtituloContainer2)
                                              .SetBorder(Border.NO_BORDER)
                                              .SetTextAlignment(TextAlignment.LEFT));
        
            container1.Add(tableSubtitulo2);

            // Parrafo 3
            Paragraph parrafo3 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(8)
                                                .Add("Tercer parrafo color de letra negro justificado. Tercer parrafo color de letra negro justificado.");

            container1.Add(parrafo3);

            // Subtitulo 3
            Paragraph subtitulo3 = new Paragraph().Add("SUBTITULO 3")
                                                  .SetFontColor(ColorConstants.WHITE)
                                                  .SetFontSize(7);

            Div subtituloContainer3 = new Div().SetWidth(90f)
                                               .Add(subtitulo3);

            Table tableSubtitulo3 = new Table(2).SetBackgroundColor(new DeviceRgb(53, 142, 54)).SetWidth(container1.GetWidth())
                                                .SetMarginTop(5)
                                                .SetMarginLeft(10)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(0);

            tableSubtitulo3.AddCell(new Cell().Add(pagImg1)
                                              .SetBorder(Border.NO_BORDER)
                                              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                                              .SetPaddingLeft(10));

            tableSubtitulo3.AddCell(new Cell().Add(subtituloContainer3)
                                              .SetBorder(Border.NO_BORDER)
                                              .SetTextAlignment(TextAlignment.LEFT));
        
            container1.Add(tableSubtitulo3);

            // Parrafo 4
            Paragraph parrafo4 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(8)
                                                .Add("Cuarto parrafo separado por un enter justificado. Cuarto parrafo separado por un enter justificado.")
                                                .Add("\n \n")
                                                .Add("Cuarto parrafo separado por un enter justificado. Cuarto parrafo separado por un enter justificado.");

            container1.Add(parrafo4);

            // Subtitulo 4
            Paragraph subtitulo4 = new Paragraph().Add("SUBTITULO 4")
                                                  .SetFontColor(ColorConstants.WHITE)
                                                  .SetFontSize(7);

            Div subtituloContainer4 = new Div().SetWidth(90f)
                                               .Add(subtitulo4);

            Table tableSubtitulo4 = new Table(2).SetBackgroundColor(new DeviceRgb(53, 142, 54))
                                                .SetWidth(container1.GetWidth())
                                                .SetMarginTop(5)
                                                .SetMarginLeft(10)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(0);

            tableSubtitulo4.AddCell(new Cell().Add(pagImg1)
                                              .SetBorder(Border.NO_BORDER)
                                              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                                              .SetPaddingLeft(10));

            tableSubtitulo4.AddCell(new Cell().Add(subtituloContainer4)
                                              .SetBorder(Border.NO_BORDER)
                                              .SetTextAlignment(TextAlignment.LEFT));
        
            container1.Add(tableSubtitulo4);

            // Parrafo 5
            Paragraph parrafo5 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(8)
                                                .Add("Quinto parrafo separado por un enter justificado. Quinto parrafo separado por un enter justificado.")
                                                .Add("\n \n")
                                                .Add("Quinto parrafo separado por un enter justificado. Quinto parrafo separado por un enter justificado.");

            container1.Add(parrafo5);

            // Parrafo 6
            Paragraph parrafo6 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(8)
                                                .Add("Sexto parrafo color de letra negro justificado. Sexto parrafo color de letra negro justificado.");

            container1.Add(parrafo6);

            // Data Cliente
            
            Paragraph dataCliente1 = new Paragraph("De:").SetFontSize(8)
                                                         .SetBorderBottom(new SolidBorder(new DeviceRgb(53, 142, 54), 1))
                                                         .SetMarginTop(5)
                                                         .SetMarginLeft(30)
                                                         .SetMarginBottom(5)
                                                         .SetMarginRight(30);

            Paragraph dataCliente2 = new Paragraph("Texto 1:").SetFontSize(8)
                                                             .SetBorderBottom(new SolidBorder(new DeviceRgb(53, 142, 54), 1))
                                                             .SetMarginTop(5)
                                                             .SetMarginLeft(30)
                                                             .SetMarginBottom(5)
                                                             .SetMarginRight(30);

            Paragraph dataCliente3 = new Paragraph("Texto 2:").SetFontSize(8)
                                                             .SetBorderBottom(new SolidBorder(new DeviceRgb(53, 142, 54), 1))
                                                             .SetMarginTop(5)
                                                             .SetMarginLeft(30)
                                                             .SetMarginBottom(5)
                                                             .SetMarginRight(30);

            container1.Add(dataCliente1);
            container1.Add(dataCliente2);
            container1.Add(dataCliente3);

            // Parrafo 7
            Paragraph parrafo7 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(8)
                                                .Add("Septimo parrafo color de letra negro justificado.");

            container1.Add(parrafo7);

            // PAGINA 1 SECCION 2
            Div container2 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 70f / 100f - 4)
                                      .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4)
                                      .SetBackgroundColor(ColorConstants.GRAY);

            Table tablePage1 = new (2);
            tablePage1.AddCell(new Cell().Add(container1).SetBorder(Border.NO_BORDER));
            tablePage1.AddCell(new Cell().Add(container2).SetBorder(Border.NO_BORDER));
            document.Add(tablePage1);

            // -------------------------------------------------

            Table tablePage2 = new (2);

            Div container3 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 50f / 100f - 4)
                                      .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4)
                                      .SetBackgroundColor(ColorConstants.YELLOW);
            
            Div container4 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 50f / 100f - 4)
                                      .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4)
                                      .SetBackgroundColor(ColorConstants.GRAY);

            tablePage2.AddCell(new Cell().Add(container3).SetBorder(Border.NO_BORDER));
            tablePage2.AddCell(new Cell().Add(container4).SetBorder(Border.NO_BORDER));

            document.Add(tablePage2);

            document.Close();   


            Response.Headers.Add("Content-Disposition", "attachment; filename=reporte.pdf");
            Response.Headers.Add("Content-Type", "application/pdf");
            return File(memoryStream.ToArray(), "application/pdf");
        }
    }
}