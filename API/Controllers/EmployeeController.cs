
using Infrastructure;
using iText.IO.Font.Constants;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
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

            PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

            DeviceRgb verde = new(53, 142, 54);
            DeviceRgb plomo = new(127, 125, 126);
            var negro = ColorConstants.BLACK;
            var blanco = ColorConstants.WHITE;

            int sizeTitulo = 8;
            int sizeSubtitulo = 7;
            int sizeTexto = 7;
            int sizeTabla = 6;

            // PRIMERA PAGINA
            // Primera seccion
            Div container1 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 30f / 100f - 4)
                                      .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4)
                                      .SetVerticalAlignment(VerticalAlignment.MIDDLE);

            ImageData imageData = ImageDataFactory.Create("C:/Users/Esteban/Documents/PDFReport/img/img1.png"); // Cambiar ruta
            Image img1 = new(imageData);
            img1.SetHeight(10); // Mantener alto
            img1.SetWidth(10); // Mantener ancho

            // Primera Seccion - Primer Parrafo
            Paragraph parrafo = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                               .SetMarginLeft(20)
                                               .SetMarginBottom(5)
                                               .SetMarginRight(20)
                                               .SetFontSize(sizeTexto)
                                               .Add("Esta es la primera línea del párrafo justificado.")
                                               .Add("Esta es la segunda línea del párrafo justificado.")
                                               .Add("Esta es la tercera línea del párrafo justificado.");
            
            container1.Add(parrafo);

            // Primera Seccion - Subtitulo 1 
            Paragraph subtitulo1 = new Paragraph().Add("SUBTITULO 1") // Subtitulo
                                                  .SetFontColor(blanco)
                                                  .SetFontSize(sizeSubtitulo)
                                                  .SetFont(boldFont);

            Div subtituloContainer1 = new Div().SetWidth(90f)
                                               .Add(subtitulo1);

            Table tableSubtitulo1 = new Table(2).SetBackgroundColor(verde)
                                                .SetWidth(container1.GetWidth())
                                                .SetMarginTop(5)
                                                .SetMarginLeft(10)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(0);

            tableSubtitulo1.AddCell(new Cell().Add(img1) // Imagen 
                                              .SetBorder(Border.NO_BORDER)
                                              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                                              .SetPaddingLeft(10));

            tableSubtitulo1.AddCell(new Cell().Add(subtituloContainer1)
                                              .SetBorder(Border.NO_BORDER))
                                              .SetTextAlignment(TextAlignment.LEFT);
        
            container1.Add(tableSubtitulo1);

            // Seccion 1 - Parrafo 2
            Paragraph parrafo2 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(sizeTexto)
                                                .Add("Segundo parrafo separado por un enter justificado. Segundo parrafo separado por un enter justificado. Segundo parrafo separado por un enter justificado." + "\n")
                                                .Add("\n" + "Segundo parrafo separado por un enter justificado. Segundo parrafo separado por un enter justificado. Segundo parrafo separado por un enter justificado.");

            container1.Add(parrafo2);

            // Seccion 1 - Subtitulo 2
            Paragraph subtitulo2 = new Paragraph().Add("SUBTITULO 2") // Subtitulo
                                                  .SetFontColor(blanco)
                                                  .SetFontSize(sizeSubtitulo)
                                                  .SetFont(boldFont);

            Div subtituloContainer2 = new Div().SetWidth(90f)
                                               .Add(subtitulo2);

            Table tableSubtitulo2 = new Table(2).SetBackgroundColor(verde)
                                                .SetWidth(container1.GetWidth())
                                                .SetMarginTop(5)
                                                .SetMarginLeft(10)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(0);

            tableSubtitulo2.AddCell(new Cell().Add(img1) // Imagen
                                              .SetBorder(Border.NO_BORDER)
                                              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                                              .SetPaddingLeft(10));

            tableSubtitulo2.AddCell(new Cell().Add(subtituloContainer2)
                                              .SetBorder(Border.NO_BORDER)
                                              .SetTextAlignment(TextAlignment.LEFT));
        
            container1.Add(tableSubtitulo2);

            // Seccion 1 - Parrafo 3
            Paragraph parrafo3 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(sizeTexto)
                                                .Add("Tercer parrafo color de letra negro justificado. Tercer parrafo color de letra negro justificado. Tercer parrafo color de letra negro justificado.");

            container1.Add(parrafo3);

            // Seccion 1 - Subtitulo 3
            Paragraph subtitulo3 = new Paragraph().Add("SUBTITULO 3") // Subtitulo
                                                  .SetFontColor(blanco)
                                                  .SetFontSize(sizeSubtitulo)
                                                  .SetFont(boldFont);

            Div subtituloContainer3 = new Div().SetWidth(90f)
                                               .Add(subtitulo3);

            Table tableSubtitulo3 = new Table(2).SetBackgroundColor(verde).SetWidth(container1.GetWidth())
                                                .SetMarginTop(5)
                                                .SetMarginLeft(10)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(0);

            tableSubtitulo3.AddCell(new Cell().Add(img1) // Imagen
                                              .SetBorder(Border.NO_BORDER)
                                              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                                              .SetPaddingLeft(10));

            tableSubtitulo3.AddCell(new Cell().Add(subtituloContainer3)
                                              .SetBorder(Border.NO_BORDER)
                                              .SetTextAlignment(TextAlignment.LEFT));
        
            container1.Add(tableSubtitulo3);

            // Seccion 1 - Parrafo 4
            Paragraph parrafo4 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(sizeTexto)
                                                .Add("Cuarto parrafo separado por un enter justificado. Cuarto parrafo separado por un enter justificado. Cuarto parrafo separado por un enter justificado." + "\n")
                                                .Add("\n" + "Cuarto parrafo separado por un enter justificado. Cuarto parrafo separado por un enter justificado. Cuarto parrafo separado por un enter justificado.");

            container1.Add(parrafo4);

            // Seccion 1 - Subtitulo 4
            Paragraph subtitulo4 = new Paragraph().Add("SUBTITULO 4") // Subtitulo
                                                  .SetFontColor(blanco)
                                                  .SetFontSize(sizeSubtitulo)
                                                  .SetFont(boldFont);

            Div subtituloContainer4 = new Div().SetWidth(90f)
                                               .Add(subtitulo4);

            Table tableSubtitulo4 = new Table(2).SetBackgroundColor(verde)
                                                .SetWidth(container1.GetWidth())
                                                .SetMarginTop(5)
                                                .SetMarginLeft(10)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(0);

            tableSubtitulo4.AddCell(new Cell().Add(img1) // Imagen
                                              .SetBorder(Border.NO_BORDER)
                                              .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                                              .SetPaddingLeft(10));

            tableSubtitulo4.AddCell(new Cell().Add(subtituloContainer4)
                                              .SetBorder(Border.NO_BORDER)
                                              .SetTextAlignment(TextAlignment.LEFT));
        
            container1.Add(tableSubtitulo4);

            // Seccion 1 - Parrafo 5
            Paragraph parrafo5 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(sizeTexto)
                                                .Add("Quinto parrafo separado por un enter justificado. Quinto parrafo separado por un enter justificado. Quinto parrafo separado por un enter justificado." + "\n")
                                                .Add("\n" + "Quinto parrafo separado por un enter justificado. Quinto parrafo separado por un enter justificado. Quinto parrafo separado por un enter justificado.");

            container1.Add(parrafo5);

            // Seccion 1 - Parrafo 6
            Paragraph parrafo6 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(sizeTexto)
                                                .Add("Sexto parrafo color de letra negro justificado. Sexto parrafo color de letra negro justificado.");

            container1.Add(parrafo6);

            // Seccion 1 - Data Cliente
            string cliente = "Nombre cliente";
            container1.Add(new Paragraph($"De: {cliente}").SetFontSize(sizeTexto)
                                                          .SetBorderBottom(new SolidBorder(verde, 1))
                                                          .SetMarginTop(5)
                                                          .SetMarginLeft(30)
                                                          .SetMarginBottom(5)
                                                          .SetMarginRight(30));

            string dataClienteTexto1 = "texto de interes 1";
            container1.Add(new Paragraph($"Texto 1: {dataClienteTexto1}").SetFontSize(sizeTexto)
                                                                         .SetBorderBottom(new SolidBorder(verde, 1))
                                                                         .SetMarginTop(5)
                                                                         .SetMarginLeft(30)
                                                                         .SetMarginBottom(5)
                                                                         .SetMarginRight(30));

            string dataClienteTexto2 = "texto de interes 2";
            container1.Add(new Paragraph($"Texto 2: {dataClienteTexto2}").SetFontSize(sizeTexto)
                                                                         .SetBorderBottom(new SolidBorder(verde, 1))
                                                                         .SetMarginTop(5)
                                                                         .SetMarginLeft(30)
                                                                         .SetMarginBottom(5)
                                                                         .SetMarginRight(30));

            // Seccion 1 - Parrafo 7
            Paragraph parrafo7 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(sizeTexto)
                                                .Add("Septimo parrafo color de letra negro justificado.");

            container1.Add(parrafo7);

            // Seccion 2
            Div container2 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 70f / 100f - 4)
                                      .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4)
                                      .SetBackgroundColor(blanco);
            
            // Seccion 2 - Header
            ImageData imageData2 = ImageDataFactory.Create("C:/Users/Esteban/Documents/PDFReport/img/img2.png"); // Cambiar ruta logo
            Image img2 = new(imageData2);
            img2.SetWidth(100f); // Mantener ancho
            // img2.SetHeight(100) Si necesitamos modificar el alto

            Table headerSeccion2 = new Table(2).SetBackgroundColor(verde)
                                               .SetWidth(container2.GetWidth())
                                               .SetHeight(90);

            Div containerheader1 = new Div().SetWidth(150)
                                            .Add(new Paragraph("PALABRAS EN TAMANIO GRANDE EN BLANCO") // Reemplazar 
                                            .SetFontSize(13).SetMarginBottom(0)) 
                                            .Add(new Paragraph("Texto Descriptivo")  // Reemplazar
                                            .SetFontSize(8).SetMarginTop(0))
                                            .SetTextAlignment(TextAlignment.CENTER);

            headerSeccion2.AddCell(new Cell().Add(containerheader1)
                                             .SetBorder(Border.NO_BORDER)
                                             .SetFontColor(blanco)
                                             .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                                             .SetPaddingLeft(50));

            Div containerheader2 = new Div().SetWidth(100f)
                                            .Add(img2) // Logo
                                            .Add(new Paragraph("Logo con texto debajo") // Texto abajo logo
                                            .SetFontSize(sizeTexto))
                                            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                                            .SetTextAlignment(TextAlignment.CENTER);

            headerSeccion2.AddCell(new Cell().Add(containerheader2)
                                             .SetBorder(Border.NO_BORDER)
                                             .SetFontColor(blanco)
                                             .SetVerticalAlignment(VerticalAlignment.MIDDLE));

            container2.Add(headerSeccion2);

            // Seccion 2 - Titulo 1
            Div titulo1 = new Div().SetBackgroundColor(verde)
                                   .SetWidth(container2.GetWidth())
                                   .SetMarginTop(2)
                                   .SetTextAlignment(TextAlignment.CENTER);

            titulo1.Add(new Paragraph("TITULO 1").SetFontColor(blanco)
                                                 .SetFontSize(sizeTitulo));

            container2.Add(titulo1);

            // Seccion 2 - Subtitulo 1         
            Div subtitulo1Sec2 = new Div().SetBackgroundColor(plomo)
                                          .SetWidth(container2.GetWidth())
                                          .SetMarginTop(2)
                                          .SetTextAlignment(TextAlignment.CENTER);

            subtitulo1Sec2.Add(new Paragraph("SUBTITULO 1").SetFontColor(blanco)
                                                           .SetFontSize(sizeSubtitulo));

            container2.Add(subtitulo1Sec2);

            // Seccion 2 - Table 1
            Div containerTableSec2 = new Div().SetMarginTop(3)
                                              .SetBorderLeft(new SolidBorder(verde, 1))
                                              .SetMarginLeft(20);

            Table tableSec2 = new Table(UnitValue.CreatePercentArray(12)).UseAllAvailableWidth()
                                                                         .SetTextAlignment(TextAlignment.CENTER)
                                                                         .SetFontSize(sizeTabla);

            tableSec2.AddCell(new Cell(2, 2).Add(new Paragraph(""))
                                            .SetBorder(Border.NO_BORDER));
            
            tableSec2.AddCell(new Cell(1, 2).Add(new Paragraph("CATG1")
                                            .SetFontColor(verde)
                                            .SetBorderBottom(new SolidBorder(verde, 1)))
                                            .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell(1, 2).Add(new Paragraph("CATG2")
                                            .SetFontColor(plomo)
                                            .SetBorderBottom(new SolidBorder(plomo, 1)))
                                            .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell(1, 2).Add(new Paragraph("CATG3")
                                            .SetBorderBottom(new SolidBorder(negro, 1)))
                                            .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell(1, 2).Add(new Paragraph("CATG4")
                                            .SetFontColor(verde)
                                            .SetBorderBottom(new SolidBorder(verde, 1)))
                                            .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell(1, 2).Add(new Paragraph("CATG5")
                                            .SetBorderBottom(new SolidBorder(negro, 1)))
                                            .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell().Add(new Paragraph("SUB1")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1))
                                        .SetBorderLeft(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell().Add(new Paragraph("SUB2")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1))
                                        .SetBorderRight(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell().Add(new Paragraph("SUB3")
                                        .SetFontColor(plomo)
                                        .SetBorderBottom(new SolidBorder(plomo, 1))
                                        .SetBorderLeft(new SolidBorder(plomo, 1)))
                                        .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell().Add(new Paragraph("SUB4")
                                        .SetFontColor(plomo)
                                        .SetBorderBottom(new SolidBorder(plomo, 1))
                                        .SetBorderRight(new SolidBorder(plomo, 1)))
                                        .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell().Add(new Paragraph("SUB5")
                                        .SetFontColor(negro)
                                        .SetBorderBottom(new SolidBorder(negro, 1))
                                        .SetBorderLeft(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell().Add(new Paragraph("SUB6")
                                        .SetFontColor(negro)
                                        .SetBorderBottom(new SolidBorder(negro, 1))
                                        .SetBorderRight(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell().Add(new Paragraph("SUB7")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1))
                                        .SetBorderLeft(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell().Add(new Paragraph("SUB8")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1))
                                        .SetBorderRight(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell().Add(new Paragraph("SUB9")
                                        .SetFontColor(negro)
                                        .SetBorderBottom(new SolidBorder(negro, 1))
                                        .SetBorderLeft(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER));

            tableSec2.AddCell(new Cell().Add(new Paragraph("SUB10")
                                        .SetFontColor(negro)
                                        .SetBorderBottom(new SolidBorder(negro, 1))
                                        .SetBorderRight(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER));

            for(int i = 0; i < 4; i++)
            {
                tableSec2.AddCell(new Cell(1, 2).Add(new Paragraph($"Texto {i + 1}")
                                                .SetTextAlignment(TextAlignment.CENTER)
                                                .SetBorderBottom(new SolidBorder(negro, 1)))
                                                .SetBorder(Border.NO_BORDER));

                tableSec2.AddCell(new Cell().Add(new Paragraph("TEXT")
                                            .SetFontColor(verde)
                                            .SetBorderBottom(new SolidBorder(verde, 1)))
                                            .SetBorder(Border.NO_BORDER));

                tableSec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                            .SetFontColor(verde)
                                            .SetBorderBottom(new SolidBorder(verde, 1)))
                                            .SetBorder(Border.NO_BORDER));

                tableSec2.AddCell(new Cell().Add(new Paragraph("TEXT")
                                            .SetFontColor(plomo)
                                            .SetBorderBottom(new SolidBorder(plomo, 1)))
                                            .SetBorder(Border.NO_BORDER));

                tableSec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                            .SetFontColor(plomo)
                                            .SetBorderBottom(new SolidBorder(plomo, 1)))
                                            .SetBorder(Border.NO_BORDER));

                tableSec2.AddCell(new Cell().Add(new Paragraph("TEXT")
                                            .SetFontColor(negro)
                                            .SetBorderBottom(new SolidBorder(negro, 1)))
                                            .SetBorder(Border.NO_BORDER));

                tableSec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                            .SetFontColor(negro)
                                            .SetBorderBottom(new SolidBorder(negro, 1)))
                                            .SetBorder(Border.NO_BORDER));

                tableSec2.AddCell(new Cell().Add(new Paragraph("TEXT")
                                            .SetFontColor(verde)
                                            .SetBorderBottom(new SolidBorder(verde, 1)))
                                            .SetBorder(Border.NO_BORDER));

                tableSec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                            .SetFontColor(verde)
                                            .SetBorderBottom(new SolidBorder(verde, 1)))
                                            .SetBorder(Border.NO_BORDER));

                tableSec2.AddCell(new Cell().Add(new Paragraph("TEXT")
                                            .SetFontColor(negro)
                                            .SetBorderBottom(new SolidBorder(negro, 1)))
                                            .SetBorder(Border.NO_BORDER));

                tableSec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                            .SetFontColor(negro)
                                            .SetBorderBottom(new SolidBorder(negro, 1)))
                                            .SetBorder(Border.NO_BORDER));
            }
            
            containerTableSec2.Add(tableSec2);
            container2.Add(containerTableSec2);

            // Seccion 2 - Subtitulo 2
            Div containerSubtitulo2Sec2 = new Div().SetMarginTop(3)
                                                   .SetFontSize(sizeSubtitulo)
                                                   .SetBackgroundColor(plomo)
                                                   .SetHeight(subtitulo1Sec2.GetHeight());

            Table table2Subtitulo2 = new Table(UnitValue.CreatePercentArray(4))
                                                        .UseAllAvailableWidth()
                                                        .SetTextAlignment(TextAlignment.CENTER);

            table2Subtitulo2.AddCell(new Cell(1, 2).Add(new Paragraph("SUBTITULO2")
                                                .SetFontColor(blanco))
                                                .SetBorder(Border.NO_BORDER)
                                                .SetMarginBottom(5));

            table2Subtitulo2.AddCell(new Cell().Add(new Paragraph("SUBTITULO 2.1")
                                                .SetFontColor(blanco))
                                                .SetBorder(Border.NO_BORDER)
                                                .SetMarginBottom(5));

            table2Subtitulo2.AddCell(new Cell().Add(new Paragraph("SUBTITULO 2.2")
                                                .SetFontColor(blanco))
                                                .SetBorder(Border.NO_BORDER)
                                                .SetMarginBottom(5));
            
            containerSubtitulo2Sec2.Add(table2Subtitulo2);
            container2.Add(containerSubtitulo2Sec2);

            // Seccion 2 - Table 2
            Div containerTable2Sec2 = new Div().SetMarginTop(3)
                                               .SetBorderLeft(new SolidBorder(verde, 1))
                                               .SetMarginLeft(20);

            Table table2Sec2 = new Table(UnitValue.CreatePercentArray(4)).UseAllAvailableWidth()
                                                                         .SetTextAlignment(TextAlignment.CENTER)
                                                                         .SetFontSize(sizeTabla);

            for(int i = 0; i < 2; i++)
            {
                table2Sec2.AddCell(new Cell(1, 2).Add(new Paragraph($"Texto {i + 1}")
                                                 .SetBorderBottom(new SolidBorder(negro, 1)))
                                                 .SetBorder(Border.NO_BORDER));

                table2Sec2.AddCell(new Cell().Add(new Paragraph($"VAL {i + 1}")
                                             .SetBorderBottom(new SolidBorder(negro, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table2Sec2.AddCell(new Cell().Add(new Paragraph($"VAL {i + 1}")
                                             .SetBorderBottom(new SolidBorder(negro, 1)))
                                             .SetBorder(Border.NO_BORDER));
            }

            containerTable2Sec2.Add(table2Sec2);
            container2.Add(containerTable2Sec2);

            // Seccion 2 - Subititulo 3
            Div subtitulo3Sec2 = new Div().SetBackgroundColor(plomo)
                                          .SetWidth(container2.GetWidth())
                                          .SetMarginTop(2)
                                          .SetTextAlignment(TextAlignment.CENTER);

            subtitulo3Sec2.Add(new Paragraph("SUBTITULO 3").SetFontColor(blanco)
                                                           .SetFontSize(sizeSubtitulo));

            container2.Add(subtitulo3Sec2);

            // Seccion 2 - Tabla 3
            Div containerTable3Sec2 = new Div().SetMarginTop(3)
                                               .SetBorderLeft(new SolidBorder(verde, 1))
                                               .SetMarginLeft(20);

            Table table3Sec2 = new Table(UnitValue.CreatePercentArray(7)).UseAllAvailableWidth()
                                                                         .SetTextAlignment(TextAlignment.CENTER)
                                                                         .SetFontSize(sizeTabla);

            table3Sec2.AddCell(new Cell(1, 2).Add(new Paragraph(""))
                                             .SetBorder(Border.NO_BORDER));

            table3Sec2.AddCell(new Cell().Add(new Paragraph("CATG 1")
                                         .SetFontColor(verde)
                                         .SetBorderBottom(new SolidBorder(verde, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table3Sec2.AddCell(new Cell().Add(new Paragraph("CATG 2") 
                                         .SetFontColor(plomo)
                                         .SetBorderBottom(new SolidBorder(plomo, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table3Sec2.AddCell(new Cell().Add(new Paragraph("CATG 3")
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table3Sec2.AddCell(new Cell().Add(new Paragraph("CATG 4")
                                         .SetFontColor(verde)
                                         .SetBorderBottom(new SolidBorder(verde, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table3Sec2.AddCell(new Cell().Add(new Paragraph("CATG 5")
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));
                                                    
            table3Sec2.AddCell(new Cell(1, 2).Add(new Paragraph("Texto")
                                             .SetBorderBottom(new SolidBorder(negro, 1)))
                                             .SetBorder(Border.NO_BORDER));

            table3Sec2.AddCell(new Cell().Add(new Paragraph("SUBCATG 1")
                                         .SetFontColor(verde)
                                         .SetBorderBottom(new SolidBorder(verde, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table3Sec2.AddCell(new Cell().Add(new Paragraph("SUBCATG 2") 
                                         .SetFontColor(plomo)
                                         .SetBorderBottom(new SolidBorder(plomo, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table3Sec2.AddCell(new Cell().Add(new Paragraph("SUBCATG 3")
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table3Sec2.AddCell(new Cell().Add(new Paragraph("SUBCATG 4")
                                         .SetFontColor(verde)
                                         .SetBorderBottom(new SolidBorder(verde, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table3Sec2.AddCell(new Cell().Add(new Paragraph("SUBCATG 5")
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));
            
            for(int i = 0; i < 3; i++)
            {
                table3Sec2.AddCell(new Cell(1, 2).Add(new Paragraph($"Texto {i + 1}")
                                                 .SetBorderBottom(new SolidBorder(negro, 1)))
                                                 .SetBorder(Border.NO_BORDER));

                table3Sec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetFontColor(verde)
                                             .SetBorderBottom(new SolidBorder(verde, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table3Sec2.AddCell(new Cell().Add(new Paragraph("VAL") 
                                             .SetFontColor(plomo)
                                             .SetBorderBottom(new SolidBorder(plomo, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table3Sec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetBorderBottom(new SolidBorder(negro, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table3Sec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetFontColor(verde)
                                             .SetBorderBottom(new SolidBorder(verde, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table3Sec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetBorderBottom(new SolidBorder(negro, 1)))
                                             .SetBorder(Border.NO_BORDER));
            }

            containerTable3Sec2.Add(table3Sec2);
            container2.Add(containerTable3Sec2);

            // Seccion 2 - Titulo 2
            Div titulo2 = new Div().SetBackgroundColor(verde)
                                   .SetWidth(container2.GetWidth())
                                   .SetMarginTop(2)
                                   .SetTextAlignment(TextAlignment.CENTER);

            titulo2.Add(new Paragraph("TITULO 2").SetFontColor(blanco)
                                                 .SetFontSize(sizeTitulo));

            container2.Add(titulo2);

            // Seccion 2 - Subtitulo 4
            Div subtitulo4Sec2 = new Div().SetBackgroundColor(plomo)
                                          .SetWidth(container2
                                          .GetWidth())
                                          .SetMarginTop(2)
                                          .SetTextAlignment(TextAlignment.CENTER);

            subtitulo4Sec2.Add(new Paragraph("SUBTITULO 4").SetFontColor(blanco)
                                                           .SetFontSize(sizeSubtitulo));

            container2.Add(subtitulo4Sec2);

            // Seccion 2 - Table 4 
            Div containerTable4Sec2 = new Div().SetMarginTop(3)
                                               .SetBorderLeft(new SolidBorder(verde, 1))
                                               .SetMarginLeft(20);

            Table table4Sec2 = new Table(UnitValue.CreatePercentArray(7))
                                                  .UseAllAvailableWidth()
                                                  .SetTextAlignment(TextAlignment.CENTER)
                                                  .SetFontSize(sizeTabla);

            table4Sec2.AddCell(new Cell(1, 2).Add(new Paragraph(""))
                                             .SetBorder(Border.NO_BORDER));

            table4Sec2.AddCell(new Cell().Add(new Paragraph("CATG 1")
                                         .SetFontColor(verde)
                                         .SetBorderBottom(new SolidBorder(verde, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table4Sec2.AddCell(new Cell().Add(new Paragraph("CATG 2") 
                                         .SetFontColor(plomo)
                                         .SetBorderBottom(new SolidBorder(plomo, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table4Sec2.AddCell(new Cell().Add(new Paragraph("CATG 3")
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table4Sec2.AddCell(new Cell().Add(new Paragraph("CATG 4")
                                         .SetFontColor(verde)
                                         .SetBorderBottom(new SolidBorder(verde, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table4Sec2.AddCell(new Cell().Add(new Paragraph("CATG 5")
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));
            
            for(int i = 0; i < 5; i++)
            {
                table4Sec2.AddCell(new Cell(1, 2).Add(new Paragraph($"Texto {i + 1}")
                                                 .SetBorderBottom(new SolidBorder(negro, 1)))
                                                 .SetBorder(Border.NO_BORDER));

                table4Sec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetFontColor(verde)
                                             .SetBorderBottom(new SolidBorder(verde, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table4Sec2.AddCell(new Cell().Add(new Paragraph("VAL") 
                                             .SetFontColor(plomo)
                                             .SetBorderBottom(new SolidBorder(plomo, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table4Sec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetBorderBottom(new SolidBorder(negro, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table4Sec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetFontColor(verde)
                                             .SetBorderBottom(new SolidBorder(verde, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table4Sec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetBorderBottom(new SolidBorder(negro, 1)))
                                             .SetBorder(Border.NO_BORDER));
            }

            containerTable4Sec2.Add(table4Sec2);
            container2.Add(containerTable4Sec2);

            // Seccion 2 - Subtitulo 5
            Div subtitulo5Sec2 = new Div().SetBackgroundColor(plomo)
                                          .SetWidth(container2.GetWidth())
                                          .SetMarginTop(2)
                                          .SetTextAlignment(TextAlignment.CENTER);

            subtitulo5Sec2.Add(new Paragraph("SUBTITULO 5").SetFontColor(blanco)
                                                           .SetFontSize(sizeSubtitulo));

            container2.Add(subtitulo5Sec2);

            // Seccion 2 - Table 5
            Div containerTable5Sec2 = new Div().SetMarginTop(3)
                                               .SetBorderLeft(new SolidBorder(verde, 1))
                                               .SetMarginLeft(20);

            Table table5Sec2 = new Table(UnitValue.CreatePercentArray(7))
                                                  .UseAllAvailableWidth()
                                                  .SetTextAlignment(TextAlignment.CENTER)
                                                  .SetFontSize(sizeTabla);

            table5Sec2.AddCell(new Cell(1, 2).Add(new Paragraph(""))
                                             .SetBorder(Border.NO_BORDER));

            table5Sec2.AddCell(new Cell().Add(new Paragraph("CATG 1")
                                         .SetFontColor(verde)
                                         .SetBorderBottom(new SolidBorder(verde, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table5Sec2.AddCell(new Cell().Add(new Paragraph("CATG 2") 
                                         .SetFontColor(plomo)
                                         .SetBorderBottom(new SolidBorder(plomo, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table5Sec2.AddCell(new Cell().Add(new Paragraph("CATG 3")
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table5Sec2.AddCell(new Cell().Add(new Paragraph("CATG 4")
                                         .SetFontColor(verde)
                                         .SetBorderBottom(new SolidBorder(verde, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table5Sec2.AddCell(new Cell().Add(new Paragraph("CATG 5")
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));
            
            for(int i = 0; i < 3; i++)
            {
                table5Sec2.AddCell(new Cell(1, 2).Add(new Paragraph($"Texto {i + 1}")
                                                 .SetBorderBottom(new SolidBorder(negro, 1)))
                                                 .SetBorder(Border.NO_BORDER));

                table5Sec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetFontColor(verde)
                                             .SetBorderBottom(new SolidBorder(verde, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table5Sec2.AddCell(new Cell().Add(new Paragraph("VAL") 
                                             .SetFontColor(plomo)
                                             .SetBorderBottom(new SolidBorder(plomo, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table5Sec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetBorderBottom(new SolidBorder(negro, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table5Sec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetFontColor(verde)
                                             .SetBorderBottom(new SolidBorder(verde, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table5Sec2.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetBorderBottom(new SolidBorder(negro, 1)))
                                             .SetBorder(Border.NO_BORDER));
            }

            containerTable5Sec2.Add(table5Sec2);
            container2.Add(containerTable5Sec2);

            Table tablePage1 = new (2);
            tablePage1.AddCell(new Cell().Add(container1).SetBorder(Border.NO_BORDER));
            tablePage1.AddCell(new Cell().Add(container2).SetBorder(Border.NO_BORDER));
            document.Add(tablePage1);

            // SEGUNDA PAGINA
            // Tercera Seccion
            Div container3 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 50f / 100f - 4)
                                      .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4)
                                      .SetVerticalAlignment(VerticalAlignment.MIDDLE);

            // Seccion 3 - Titulo 3
            Div titulo3 = new Div().SetBackgroundColor(verde)
                                   .SetWidth(container2.GetWidth())
                                   .SetMarginTop(2)
                                   .SetTextAlignment(TextAlignment.CENTER);

            titulo3.Add(new Paragraph("TITULO 3").SetFontColor(blanco)
                                                 .SetFontSize(sizeTitulo));

            container3.Add(titulo3);

            // Seccion 3 - Subtitulo 6
            Div subtitulo6Sec3 = new Div().SetBackgroundColor(plomo)
                                          .SetWidth(container2.GetWidth())
                                          .SetMarginTop(2)
                                          .SetTextAlignment(TextAlignment.CENTER);

            subtitulo6Sec3.Add(new Paragraph("SUBTITULO 6").SetFontColor(blanco)
                                                           .SetFontSize(sizeSubtitulo));

            container3.Add(subtitulo6Sec3);

            // Seccion 3 - Tabla 6
            Div containerTable6 = new Div().SetMarginTop(3)
                                           .SetBorderLeft(new SolidBorder(verde, 1))
                                           .SetMarginLeft(20)
                                           .SetMarginRight(20);

            Table table6 = new Table(UnitValue.CreatePercentArray(10)).UseAllAvailableWidth()
                                                                      .SetTextAlignment(TextAlignment.CENTER)
                                                                      .SetFontSize(sizeTabla);

            table6.AddCell(new Cell(1, 2).Add(new Paragraph(""))
                                         .SetBorder(Border.NO_BORDER));
            
            table6.AddCell(new Cell(1, 2).Add(new Paragraph("CATG1")
                                         .SetFontColor(verde)
                                         .SetBorderBottom(new SolidBorder(verde, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table6.AddCell(new Cell(1, 2).Add(new Paragraph("CATG2")
                                         .SetFontColor(plomo)
                                         .SetBorderBottom(new SolidBorder(plomo, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table6.AddCell(new Cell(1, 2).Add(new Paragraph("CATG3")
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));

            table6.AddCell(new Cell(1, 2).Add(new Paragraph("CATG4")
                                         .SetFontColor(negro)
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));

            for(int i = 0; i < 10; i++)
            {
                table6.AddCell(new Cell(1, 2).Add(new Paragraph($"Texto {i + 1}")
                                             .SetTextAlignment(TextAlignment.CENTER)
                                             .SetBorderBottom(new SolidBorder(negro, 1)))
                                             .SetBorder(Border.NO_BORDER));

                table6.AddCell(new Cell().Add(new Paragraph("TEXT")
                                         .SetFontColor(verde)
                                         .SetBorderBottom(new SolidBorder(verde, 1)))
                                         .SetBorder(Border.NO_BORDER));

                table6.AddCell(new Cell().Add(new Paragraph("TEXT")
                                         .SetFontColor(verde)
                                         .SetBorderBottom(new SolidBorder(verde, 1)))
                                         .SetBorder(Border.NO_BORDER));

                table6.AddCell(new Cell().Add(new Paragraph("TEXT")
                                         .SetFontColor(plomo)
                                         .SetBorderBottom(new SolidBorder(plomo, 1)))
                                         .SetBorder(Border.NO_BORDER));

                table6.AddCell(new Cell().Add(new Paragraph("TEXT")
                                         .SetFontColor(plomo)
                                         .SetBorderBottom(new SolidBorder(plomo, 1)))
                                         .SetBorder(Border.NO_BORDER));

                table6.AddCell(new Cell().Add(new Paragraph("TEXT")
                                         .SetFontColor(negro)
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));

                table6.AddCell(new Cell().Add(new Paragraph("TEXT")
                                         .SetFontColor(negro)
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));

                table6.AddCell(new Cell().Add(new Paragraph("TEXT")
                                         .SetFontColor(negro)
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));

                table6.AddCell(new Cell().Add(new Paragraph("TEXT")
                                         .SetFontColor(negro)
                                         .SetBorderBottom(new SolidBorder(negro, 1)))
                                         .SetBorder(Border.NO_BORDER));
            }
            
            containerTable6.Add(table6);
            container3.Add(containerTable6);

            // Seccion 3 - Subtitulo 7
            Div subtitulo6 = new Div().SetBackgroundColor(plomo)
                                      .SetWidth(container2.GetWidth())
                                      .SetMarginTop(2)
                                      .SetTextAlignment(TextAlignment.CENTER);

            subtitulo6.Add(new Paragraph("SUBTITULO 7").SetFontColor(blanco)
                                                       .SetFontSize(sizeSubtitulo));

            container3.Add(subtitulo6);

            // Seccion 3 - Tabla 7
            Div containerTable7 = new Div().SetMarginTop(3)
                                           .SetMarginLeft(20)
                                           .SetMarginRight(150)
                                           .SetMarginTop(10);

            Table table7 = new Table(UnitValue.CreatePercentArray(5)).UseAllAvailableWidth()
                                                                     .SetTextAlignment(TextAlignment.CENTER)
                                                                     .SetFontSize(sizeTabla)
                                                                     .SetBorder(new SolidBorder(verde, 1));

            for(int i = 0; i < 2; i++)
            {
                table7.AddCell(new Cell(4, 2).Add(new Paragraph($"Texto {i + 1}")
                                             .SetTextAlignment(TextAlignment.CENTER)
                                             .SetMaxWidth(100))
                                             .SetBorder(new SolidBorder(verde, 1)));

                table7.AddCell(new Cell(1, 2).Add(new Paragraph("CATG 1")
                                             .SetFontColor(plomo))
                                             .SetBorder(new SolidBorder(verde, 1)));

                table7.AddCell(new Cell().Add(new Paragraph("CATG 2")
                                         .SetFontColor(plomo))
                                         .SetBorder(new SolidBorder(verde, 1)));

                for(int j = 0; j < 3; j++)
                {
                    table7.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetFontColor(negro))
                                             .SetBorder(new SolidBorder(verde, 1)));

                    table7.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetFontColor(negro))
                                             .SetBorder(new SolidBorder(verde, 1)));

                    table7.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetFontColor(negro))
                                             .SetBorder(new SolidBorder(verde, 1)));
                }
            }

            for(int i = 0; i < 2; i++)
            {
                table7.AddCell(new Cell(5, 2).Add(new Paragraph($"Texto {i + 1}")
                                             .SetTextAlignment(TextAlignment.CENTER))
                                             .SetBorder(new SolidBorder(verde, 1)));

                table7.AddCell(new Cell(1, 2).Add(new Paragraph("CATG 1")
                                             .SetFontColor(plomo))
                                             .SetBorder(new SolidBorder(verde, 1)));

                table7.AddCell(new Cell().Add(new Paragraph("CATG 2")
                                         .SetFontColor(plomo))
                                         .SetBorder(new SolidBorder(verde, 1)));

                for(int j = 0; j < 4; j++)
                {
                    table7.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetFontColor(negro))
                                             .SetBorder(new SolidBorder(verde, 1)));

                    table7.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetFontColor(negro))
                                             .SetBorder(new SolidBorder(verde, 1)));

                    table7.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetFontColor(negro))
                                             .SetBorder(new SolidBorder(verde, 1)));
                }
            }
            
            containerTable7.Add(table7);
            container3.Add(containerTable7);

            // Seccion 3 - Texto 1
            Div texto1Sec3 = new Div().SetWidth(container2.GetWidth())
                                      .SetMarginTop(10)
                                      .SetTextAlignment(TextAlignment.CENTER)
                                      .SetMarginLeft(20)
                                      .SetMarginRight(150);

            texto1Sec3.Add(new Paragraph("TEXTO 1").SetFontColor(negro)
                                                   .SetFontSize(sizeSubtitulo));

            container3.Add(texto1Sec3);

            // Seccion 3 - Texto 2
            Div texto2Sec3 = new Div().SetWidth(container2.GetWidth())
                                      .SetMarginTop(2)
                                      .SetTextAlignment(TextAlignment.CENTER)
                                      .SetMarginLeft(20)
                                      .SetMarginRight(150);

            texto2Sec3.Add(new Paragraph("TEXTO 2").SetFontColor(negro)
                                                   .SetFontSize(sizeSubtitulo));

            container3.Add(texto2Sec3);

            // Cuarta Seccion
            Div container4 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 50f / 100f - 4)
                                      .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4)
                                      .SetVerticalAlignment(VerticalAlignment.MIDDLE);

            // Seccion 4 - Subtitulo 8
            Div subtitulo8 = new Div().SetBackgroundColor(plomo)
                                      .SetWidth(container2.GetWidth())
                                      .SetTextAlignment(TextAlignment.CENTER);

            subtitulo8.Add(new Paragraph("SUBTITULO 8").SetFontColor(blanco)
                                                       .SetFontSize(sizeSubtitulo));

            container4.Add(subtitulo8);

            // Seccion 4 - Tabla 8 
            Div containerTable8 = new Div().SetMarginTop(3)
                                           .SetMarginLeft(20)
                                           .SetBorderLeft(new SolidBorder(verde, 1));

            Table table8 = new Table(UnitValue.CreatePercentArray(7)).UseAllAvailableWidth()
                                                                     .SetTextAlignment(TextAlignment.CENTER)
                                                                     .SetFontSize(sizeTabla);

            for(int i = 0; i < 1; i++)
            {
                table8.AddCell(new Cell(2, 2).Add(new Paragraph($"Texto {i + 1}")
                                             .SetTextAlignment(TextAlignment.CENTER)
                                             .SetMaxWidth(105))
                                             .SetBorder(Border.NO_BORDER));

                table8.AddCell(new Cell().Add(new Paragraph("CATG 1")
                                         .SetBorderBottom(new SolidBorder(verde, 1))
                                         .SetFontColor(verde))
                                         .SetBorder(Border.NO_BORDER));

                table8.AddCell(new Cell().Add(new Paragraph("CATG 2")
                                         .SetBorderBottom(new SolidBorder(plomo, 1))
                                         .SetFontColor(plomo))
                                         .SetBorder(Border.NO_BORDER));

                table8.AddCell(new Cell().Add(new Paragraph("CATG 3")
                                         .SetBorderBottom(new SolidBorder(negro, 1))
                                         .SetFontColor(negro))
                                         .SetBorder(Border.NO_BORDER));

                table8.AddCell(new Cell().Add(new Paragraph("CATG 4")
                                         .SetBorderBottom(new SolidBorder(verde, 1))
                                         .SetFontColor(verde))
                                         .SetBorder(Border.NO_BORDER));

                table8.AddCell(new Cell().Add(new Paragraph("CATG 5")
                                         .SetBorderBottom(new SolidBorder(negro, 1))
                                         .SetFontColor(negro))
                                         .SetBorder(Border.NO_BORDER));

                for(int j = 0; j < 1; j++)
                {
                    table8.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetBorderBottom(new SolidBorder(verde, 1))
                                             .SetFontColor(verde))
                                             .SetBorder(Border.NO_BORDER));

                    table8.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetBorderBottom(new SolidBorder(plomo, 1))
                                             .SetFontColor(plomo))
                                             .SetBorder(Border.NO_BORDER));

                    table8.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetBorderBottom(new SolidBorder(negro, 1))
                                             .SetFontColor(negro))
                                             .SetBorder(Border.NO_BORDER));

                    table8.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetBorderBottom(new SolidBorder(verde, 1))
                                             .SetFontColor(verde))
                                             .SetBorder(Border.NO_BORDER));

                    table8.AddCell(new Cell().Add(new Paragraph("VAL")
                                             .SetBorderBottom(new SolidBorder(negro, 1))
                                             .SetFontColor(plomo))
                                             .SetBorder(Border.NO_BORDER));
                }
            }

            containerTable8.Add(table8);
            container4.Add(containerTable8);

            // Seccion 4 - Subtitulo 9
            Div subtitulo9 = new Div().SetBackgroundColor(plomo)
                                      .SetWidth(container4.GetWidth())
                                      .SetMarginTop(2)
                                      .SetTextAlignment(TextAlignment.CENTER);

            subtitulo9.Add(new Paragraph("SUBTITULO 9").SetFontColor(blanco)
                                                       .SetFontSize(sizeSubtitulo));

            container4.Add(subtitulo9);

            // Seccion 4 = Texto 4
            Div texto3 = new Div().SetWidth(container4.GetWidth())
                                  .SetTextAlignment(TextAlignment.CENTER);

            texto3.Add(new Paragraph("TEXTO").SetFontColor(negro)
                                             .SetFontSize(sizeSubtitulo)
                                             .SetMarginLeft(20));

            container4.Add(texto3);

            // Seccion 4 - Container tabla y texto
            Div container6 = new Div().SetWidth(container4.GetWidth().GetValue() * 90f / 100f);
            
            Div container5 = new Div().SetWidth(8)
                                      .SetPadding(3)
                                      .SetMarginLeft(17)
                                      .SetBorderLeft(new SolidBorder(verde, 1))
                                      .SetHeight(100f)
                                      .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                                      .SetTextAlignment(TextAlignment.CENTER);
            // Seccion 4 = Texto 4
            Div texto4 = new Div().SetWidth(container5.GetWidth());

            texto4.Add(new Paragraph("TEXTO").SetFontColor(negro)
                                             .SetFontSize(sizeSubtitulo));

            container5.Add(texto4);

            // Seccion 4 - Tabla 9
            Div containerTable9 = new();

            Table table9 = new Table(UnitValue.CreatePercentArray(5)).UseAllAvailableWidth()
                                                                     .SetTextAlignment(TextAlignment.CENTER)
                                                                     .SetFontSize(sizeTabla);

            table9.AddCell(new Cell().Add(new Paragraph($"")
                                     .SetTextAlignment(TextAlignment.CENTER))
                                     .SetBorder(Border.NO_BORDER));

            table9.AddCell(new Cell().Add(new Paragraph($"Texto 1")
                                     .SetTextAlignment(TextAlignment.CENTER)
                                     .SetBorderBottom(new SolidBorder(verde, 1))
                                     .SetFontColor(verde))
                                     .SetBorder(Border.NO_BORDER));

            table9.AddCell(new Cell().Add(new Paragraph($"Texto 2")
                                     .SetTextAlignment(TextAlignment.CENTER)
                                     .SetBorderBottom(new SolidBorder(plomo, 1))
                                     .SetFontColor(plomo))
                                     .SetBorder(Border.NO_BORDER));

            table9.AddCell(new Cell().Add(new Paragraph($"Texto 3")
                                     .SetTextAlignment(TextAlignment.CENTER)
                                     .SetBorderBottom(new SolidBorder(verde, 1))
                                     .SetFontColor(verde))
                                     .SetBorder(Border.NO_BORDER));

            table9.AddCell(new Cell().Add(new Paragraph($"Texto 4")
                                     .SetTextAlignment(TextAlignment.CENTER)
                                     .SetBorderBottom(new SolidBorder(plomo, 1))
                                     .SetFontColor(plomo))
                                     .SetBorder(Border.NO_BORDER));

            for(int i = 0; i < 6; i++)
            {
                table9.AddCell(new Cell().Add(new Paragraph($"Texto {i + 1}")
                                         .SetBorderBottom(new SolidBorder(negro, 1))
                                         .SetTextAlignment(TextAlignment.CENTER))
                                         .SetBorder(Border.NO_BORDER));

                table9.AddCell(new Cell().Add(new Paragraph("CATG 1")
                                         .SetBorderBottom(new SolidBorder(verde, 1))
                                         .SetFontColor(verde))
                                         .SetBorder(Border.NO_BORDER));

                table9.AddCell(new Cell().Add(new Paragraph("CATG 2")
                                         .SetBorderBottom(new SolidBorder(plomo, 1))
                                         .SetFontColor(plomo))
                                         .SetBorder(Border.NO_BORDER));

                table9.AddCell(new Cell().Add(new Paragraph("CATG 3")
                                         .SetBorderBottom(new SolidBorder(verde, 1))
                                         .SetFontColor(verde))
                                         .SetBorder(Border.NO_BORDER));

                table9.AddCell(new Cell().Add(new Paragraph("CATG 4")
                                         .SetBorderBottom(new SolidBorder(plomo, 1))
                                         .SetFontColor(plomo))
                                         .SetBorder(Border.NO_BORDER));
            }

            containerTable9.Add(table9);
            container6.Add(containerTable9);

            Table table9Seccion4 = new (2);
            table9Seccion4.AddCell(new Cell().Add(container5).SetBorder(Border.NO_BORDER));
            table9Seccion4.AddCell(new Cell().Add(container6).SetBorder(Border.NO_BORDER));
            container4.Add(table9Seccion4);

            // Seccion 4 - Texto 5
            Div texto5 = new Div().SetWidth(container4.GetWidth())
                                  .SetMarginLeft(20);

            texto5.Add(new Paragraph("TEXTO").SetFontColor(negro)
                                             .SetFontSize(sizeSubtitulo));

            container4.Add(texto5);

            // Seccion 4 - Subtitulo 10
            Div subtitulo10 = new Div().SetBackgroundColor(plomo)
                                       .SetWidth(container4.GetWidth())
                                       .SetMarginTop(2)
                                       .SetMarginBottom(3)
                                       .SetTextAlignment(TextAlignment.CENTER);

            subtitulo10.Add(new Paragraph("SUBTITULO 10").SetFontColor(blanco)
                                                         .SetFontSize(sizeSubtitulo));

            container4.Add(subtitulo10);

            // Seccion 4 - Container tabla 10
            Div containerTabla10 = new Div().SetWidth(container4.GetWidth())
                                            .SetBorderLeft(new SolidBorder(verde, 1))
                                            .SetMarginLeft(20)
                                            .SetMarginRight(150)
                                            .SetBorderLeft(new SolidBorder(verde, 1));

            Table tabla10 = new Table(UnitValue.CreatePercentArray(4)).UseAllAvailableWidth()
                                                                      .SetTextAlignment(TextAlignment.CENTER)
                                                                      .SetFontSize(sizeTabla)
                                                                      .SetMarginTop(10)
                                                                      .SetMarginLeft(5);

            tabla10.AddCell(new Cell(2, 2).Add(new Paragraph($"TEXTO")
                                          .SetBorderBottom(new SolidBorder(negro, 1))
                                          .SetTextAlignment(TextAlignment.CENTER)
                                          .SetFontSize(8)
                                          .SetMaxWidth(118)
                                          .SetHeight(20))
                                         .SetBorder(Border.NO_BORDER));

            tabla10.AddCell(new Cell(2, 2).Add(new Paragraph($"TEXTO")
                                          .SetBorderBottom(new SolidBorder(negro, 1))
                                          .SetTextAlignment(TextAlignment.CENTER)
                                          .SetFontSize(8)
                                          .SetMaxWidth(118)
                                          .SetHeight(20))
                                          .SetBorder(Border.NO_BORDER));

            for(int i = 0; i < 5; i++)
            {
                tabla10.AddCell(new Cell(1, 2).Add(new Paragraph($"Texto {i + 1}")
                                              .SetBorderBottom(new SolidBorder(negro, 1))
                                              .SetTextAlignment(TextAlignment.CENTER))
                                              .SetBorder(Border.NO_BORDER));

                tabla10.AddCell(new Cell(1, 2).Add(new Paragraph($"Texto {i + 1}")
                                              .SetBorderBottom(new SolidBorder(negro, 1))
                                              .SetTextAlignment(TextAlignment.CENTER))
                                              .SetBorder(Border.NO_BORDER));
            }

            containerTabla10.Add(tabla10);
            container4.Add(containerTabla10);

            // Seccion 4 - Parrafo tabla 10
            containerTabla10.Add(new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetPadding(5)
                                                .SetFontSize(sizeTexto)
                                                .Add("Esta es la primera línea del párrafo justificado.")
                                                .Add("Esta es la segunda línea del párrafo justificado.")
                                                .Add("Esta es la tercera línea del párrafo justificado."));

            // Seccion 4 - Subtitulo 11
            Div subtitulo11 = new Div().SetBackgroundColor(plomo)
                                       .SetWidth(container4.GetWidth())
                                       .SetMarginTop(2)
                                       .SetMarginBottom(3)
                                       .SetTextAlignment(TextAlignment.CENTER);

            subtitulo11.Add(new Paragraph("SUBTITULO 11").SetFontColor(blanco)
                                                         .SetFontSize(sizeSubtitulo));

            container4.Add(subtitulo11);

            // Seccion 4 - Parrafo dos lineas justificado
            Div containerParrafo = new Div().SetWidth(container4.GetWidth())
                                            .SetBorderLeft(new SolidBorder(verde, 1))
                                            .SetMarginLeft(20)
                                            .SetMarginRight(150)
                                            .SetBorderLeft(new SolidBorder(verde, 1));

            containerParrafo.Add(new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetPadding(5)
                                                .SetFontSize(sizeTexto)
                                                .Add("Esta es la primera línea del párrafo justificado.")
                                                .Add("Esta es la segunda línea del párrafo justificado."));

            container4.Add(containerParrafo);

            // Seccion 4 - Subtitulo 12
            Div subtitulo12 = new Div().SetBackgroundColor(plomo)
                                       .SetWidth(container4.GetWidth())
                                       .SetMarginTop(2)
                                       .SetTextAlignment(TextAlignment.CENTER);

            subtitulo12.Add(new Paragraph("SUBTITULO 12").SetFontColor(blanco)
                                                         .SetFontSize(sizeSubtitulo));

            container4.Add(subtitulo12);

            // Seccion 4 - Tabla 11
            // Seccion 2 - Table 5
            Div containerTable11 = new Div().SetMarginTop(3)
                                            .SetBorderLeft(new SolidBorder(verde, 1))
                                            .SetMarginLeft(20);

            Table table11 = new Table(UnitValue.CreatePercentArray(7))
                                               .UseAllAvailableWidth()
                                               .SetTextAlignment(TextAlignment.CENTER)
                                               .SetFontSize(sizeTabla);

            table11.AddCell(new Cell(1, 2).Add(new Paragraph(""))
                                          .SetBorder(Border.NO_BORDER));

            table11.AddCell(new Cell().Add(new Paragraph("CATG 1")
                                      .SetFontColor(verde)
                                      .SetBorderBottom(new SolidBorder(verde, 1)))
                                      .SetBorder(Border.NO_BORDER));

            table11.AddCell(new Cell().Add(new Paragraph("CATG 2") 
                                      .SetFontColor(plomo)
                                      .SetBorderBottom(new SolidBorder(plomo, 1)))
                                      .SetBorder(Border.NO_BORDER));

            table11.AddCell(new Cell().Add(new Paragraph("CATG 3")
                                      .SetBorderBottom(new SolidBorder(negro, 1)))
                                      .SetBorder(Border.NO_BORDER));

            table11.AddCell(new Cell().Add(new Paragraph("CATG 4")
                                      .SetFontColor(verde)
                                      .SetBorderBottom(new SolidBorder(verde, 1)))
                                      .SetBorder(Border.NO_BORDER));

            table11.AddCell(new Cell().Add(new Paragraph("CATG 5")
                                      .SetBorderBottom(new SolidBorder(negro, 1)))
                                      .SetBorder(Border.NO_BORDER));
            
            for(int i = 0; i < 4; i++)
            {
                table11.AddCell(new Cell(1, 2).Add(new Paragraph($"Texto {i + 1}")
                                              .SetBorderBottom(new SolidBorder(negro, 1)))
                                              .SetBorder(Border.NO_BORDER));

                table11.AddCell(new Cell().Add(new Paragraph("VAL")
                                          .SetFontColor(verde)
                                          .SetBorderBottom(new SolidBorder(verde, 1)))
                                          .SetBorder(Border.NO_BORDER));

                table11.AddCell(new Cell().Add(new Paragraph("VAL") 
                                          .SetFontColor(plomo)
                                          .SetBorderBottom(new SolidBorder(plomo, 1)))
                                          .SetBorder(Border.NO_BORDER));

                table11.AddCell(new Cell().Add(new Paragraph("VAL")
                                          .SetBorderBottom(new SolidBorder(negro, 1)))
                                          .SetBorder(Border.NO_BORDER));

                table11.AddCell(new Cell().Add(new Paragraph("VAL")
                                          .SetFontColor(verde)
                                          .SetBorderBottom(new SolidBorder(verde, 1)))
                                          .SetBorder(Border.NO_BORDER));

                table11.AddCell(new Cell().Add(new Paragraph("VAL")
                                          .SetBorderBottom(new SolidBorder(negro, 1)))
                                          .SetBorder(Border.NO_BORDER));
            }

            containerTable11.Add(table11);
            container4.Add(containerTable11);

            Table tablePage2 = new (2);
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