
using System.Drawing;
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

            ImageData imageData = ImageDataFactory.Create("C:/Users/Esteban/Documents/PDFReport/img/img1.png");
            Image pagImg1 = new(imageData);
            pagImg1.SetHeight(10);
            pagImg1.SetWidth(10);

            DeviceRgb verde = new(53, 142, 54);
            DeviceRgb plomo = new(127, 125, 126);
            var negro = ColorConstants.BLACK;
            var blanco = ColorConstants.WHITE;

            int sizeTitulo = 8;
            int sizeSubtitulo = 7;
            int sizeTexto = 7;
            int sizeTabla = 6;

            // PAGINA 1
            
            // PAGINA 1 SECCION 1
            Div container1 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 30f / 100f - 4)
                                      .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4);

            // Parrafo 1
            Paragraph parrafo = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                               .SetMarginTop(5)
                                               .SetMarginLeft(20)
                                               .SetMarginBottom(5)
                                               .SetMarginRight(20)
                                               .SetFontSize(sizeTexto)
                                               .Add("Esta es la primera línea del párrafo justificado.")
                                               .Add("Esta es la segunda línea del párrafo justificado.")
                                               .Add("Esta es la tercera línea del párrafo justificado.");

            container1.Add(parrafo);

            // Subtitulo 1 
            Paragraph subtitulo1 = new Paragraph().Add("SUBTITULO 1")
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
                                                .SetFontSize(sizeTexto)
                                                .Add("Segundo parrafo separado por un enter justificado. Segundo parrafo separado por un enter justificado. Segundo parrafo separado por un enter justificado.")
                                                .Add("\n \n")
                                                .Add("Segundo parrafo separado por un enter justificado. Segundo parrafo separado por un enter justificado. Segundo parrafo separado por un enter justificado.");

            container1.Add(parrafo2);

            // Subtitulo 2
            Paragraph subtitulo2 = new Paragraph().Add("SUBTITULO 2")
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
                                                .SetFontSize(sizeTexto)
                                                .Add("Tercer parrafo color de letra negro justificado. Tercer parrafo color de letra negro justificado. Tercer parrafo color de letra negro justificado.");

            container1.Add(parrafo3);

            // Subtitulo 3
            Paragraph subtitulo3 = new Paragraph().Add("SUBTITULO 3")
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
                                                .SetFontSize(sizeTexto)
                                                .Add("Cuarto parrafo separado por un enter justificado. Cuarto parrafo separado por un enter justificado. Cuarto parrafo separado por un enter justificado.")
                                                .Add("\n \n")
                                                .Add("Cuarto parrafo separado por un enter justificado. Cuarto parrafo separado por un enter justificado. Cuarto parrafo separado por un enter justificado.");

            container1.Add(parrafo4);

            // Subtitulo 4
            Paragraph subtitulo4 = new Paragraph().Add("SUBTITULO 4")
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
                                                .SetFontSize(sizeTexto)
                                                .Add("Quinto parrafo separado por un enter justificado. Quinto parrafo separado por un enter justificado. Quinto parrafo separado por un enter justificado.")
                                                .Add("\n \n")
                                                .Add("Quinto parrafo separado por un enter justificado. Quinto parrafo separado por un enter justificado. Quinto parrafo separado por un enter justificado.");

            container1.Add(parrafo5);

            // Parrafo 6
            Paragraph parrafo6 = new Paragraph().SetTextAlignment(TextAlignment.JUSTIFIED)
                                                .SetMarginTop(5)
                                                .SetMarginLeft(20)
                                                .SetMarginBottom(5)
                                                .SetMarginRight(20)
                                                .SetFontSize(sizeTexto)
                                                .Add("Sexto parrafo color de letra negro justificado. Sexto parrafo color de letra negro justificado.");

            container1.Add(parrafo6);

            // Data Cliente
            
            Paragraph dataCliente1 = new Paragraph("De:").SetFontSize(sizeTexto)
                                                         .SetBorderBottom(new SolidBorder(verde, 1))
                                                         .SetMarginTop(5)
                                                         .SetMarginLeft(30)
                                                         .SetMarginBottom(5)
                                                         .SetMarginRight(30);

            Paragraph dataCliente2 = new Paragraph("Texto 1:").SetFontSize(sizeTexto)
                                                             .SetBorderBottom(new SolidBorder(verde, 1))
                                                             .SetMarginTop(5)
                                                             .SetMarginLeft(30)
                                                             .SetMarginBottom(5)
                                                             .SetMarginRight(30);

            Paragraph dataCliente3 = new Paragraph("Texto 2:").SetFontSize(sizeTexto)
                                                             .SetBorderBottom(new SolidBorder(verde, 1))
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
                                                .SetFontSize(sizeTexto)
                                                .Add("Septimo parrafo color de letra negro justificado.");

            container1.Add(parrafo7);

            // PAGINA 1 SECCION 2
            Div container2 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 70f / 100f - 4)
                                      .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4)
                                      .SetBackgroundColor(blanco);
            
            // Header seccion 2
            ImageData imageData2 = ImageDataFactory.Create("C:/Users/Esteban/Documents/PDFReport/img/img2.png");
            Image pagImg2 = new(imageData2);
            pagImg2.SetWidth(100f);

            Table headerSeccion2 = new Table(2).SetBackgroundColor(verde)
                                               .SetWidth(container2.GetWidth())
                                               .SetHeight(90);

            Div containerheader1 = new Div().SetWidth(150)
                                            .Add(new Paragraph("PALABRAS EN TAMANIO GRANDE EN BLANCO").SetFontSize(sizeTitulo))
                                            .Add(new Paragraph("Texto Descriptivo").SetFontSize(sizeTexto))
                                            .SetTextAlignment(TextAlignment.CENTER);

            headerSeccion2.AddCell(new Cell().Add(containerheader1)
                                             .SetBorder(Border.NO_BORDER)
                                             .SetFontColor(blanco)
                                             .SetVerticalAlignment(VerticalAlignment.MIDDLE)
                                             .SetPaddingLeft(50));

            Div containerheader2 = new Div().SetWidth(100f)
                                            .Add(pagImg2)
                                            .Add(new Paragraph("Logo con texto debajo").SetFontSize(sizeTexto))
                                            .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                                            .SetTextAlignment(TextAlignment.CENTER);

            headerSeccion2.AddCell(new Cell().Add(containerheader2)
                                             .SetBorder(Border.NO_BORDER)
                                             .SetFontColor(blanco)
                                             .SetVerticalAlignment(VerticalAlignment.MIDDLE));

            container2.Add(headerSeccion2);

            // Titulo 1
            Div titulo1 = new Div().SetBackgroundColor(verde)
                                   .SetWidth(container2.GetWidth())
                                   .SetMarginTop(2)
                                   .SetTextAlignment(TextAlignment.CENTER);

            titulo1.Add(new Paragraph("TITULO 1").SetFontColor(blanco)
                                                 .SetFontSize(sizeTitulo));

            container2.Add(titulo1);

            // Subtitulo 1         
            Div subtitulo1Sec2 = new Div().SetBackgroundColor(plomo)
                                   .SetWidth(container2.GetWidth())
                                   .SetMarginTop(2)
                                   .SetTextAlignment(TextAlignment.CENTER);

            subtitulo1Sec2.Add(new Paragraph("SUBTITULO 1").SetFontColor(blanco)
                                                 .SetFontSize(sizeSubtitulo));

            container2.Add(subtitulo1Sec2);

            // Table 1 seccion 2
            Div containerTableSec2 = new Div().SetMarginTop(3);

            Table tableSec2 = new Table(UnitValue.CreatePercentArray(12)).UseAllAvailableWidth()
                                                                         .SetTextAlignment(TextAlignment.CENTER)
                                                                         .SetFontSize(sizeTabla);

            Cell spaceless = new Cell(2, 2).Add(new Paragraph("")).SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(spaceless);

            Cell categ1 = new Cell(1, 2).Add(new Paragraph("CATG1")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(categ1);

            Cell categ2 = new Cell(1, 2).Add(new Paragraph("CATG2")
                                        .SetFontColor(plomo)
                                        .SetBorderBottom(new SolidBorder(plomo, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(categ2);

            Cell categ3 = new Cell(1, 2).Add(new Paragraph("CATG3")
                                        .SetBorderBottom(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(categ3);

            Cell categ4 = new Cell(1, 2).Add(new Paragraph("CATG4")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(categ4);

            Cell categ5 = new Cell(1, 2).Add(new Paragraph("CATG5")
                                        .SetBorderBottom(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(categ5);

            Cell sub1 = new Cell().Add(new Paragraph("SUB1")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1))
                                        .SetBorderLeft(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(sub1);

            Cell sub2 = new Cell().Add(new Paragraph("SUB2")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1))
                                        .SetBorderRight(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(sub2);

            Cell sub3 = new Cell().Add(new Paragraph("SUB3")
                                        .SetFontColor(plomo)
                                        .SetBorderBottom(new SolidBorder(plomo, 1))
                                        .SetBorderLeft(new SolidBorder(plomo, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(sub3);

            Cell sub4 = new Cell().Add(new Paragraph("SUB4")
                                        .SetFontColor(plomo)
                                        .SetBorderBottom(new SolidBorder(plomo, 1))
                                        .SetBorderRight(new SolidBorder(plomo, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(sub4);

            Cell sub5 = new Cell().Add(new Paragraph("SUB5")
                                        .SetFontColor(negro)
                                        .SetBorderBottom(new SolidBorder(negro, 1))
                                        .SetBorderLeft(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(sub5);

            Cell sub6 = new Cell().Add(new Paragraph("SUB6")
                                        .SetFontColor(negro)
                                        .SetBorderBottom(new SolidBorder(negro, 1))
                                        .SetBorderRight(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(sub6);

            Cell sub7 = new Cell().Add(new Paragraph("SUB7")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1))
                                        .SetBorderLeft(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(sub7);

            Cell sub8 = new Cell().Add(new Paragraph("SUB8")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1))
                                        .SetBorderRight(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(sub8);

            Cell sub9 = new Cell().Add(new Paragraph("SUB9")
                                        .SetFontColor(negro)
                                        .SetBorderBottom(new SolidBorder(negro, 1))
                                        .SetBorderLeft(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(sub9);

            Cell sub10 = new Cell().Add(new Paragraph("SUB10")
                                        .SetFontColor(negro)
                                        .SetBorderBottom(new SolidBorder(negro, 1))
                                        .SetBorderRight(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER);
            tableSec2.AddCell(sub10);

            for(int i = 0; i < 4; i++)
            {
                Cell texto1 = new Cell(1, 2).Add(new Paragraph($"Texto {i + 1}")
                                            .SetTextAlignment(TextAlignment.CENTER)
                                            .SetBorderBottom(new SolidBorder(negro, 1)))
                                            .SetBorder(Border.NO_BORDER);
                tableSec2.AddCell(texto1);

                Cell textVal1 = new Cell().Add(new Paragraph("TEXT")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER);
                tableSec2.AddCell(textVal1);

                Cell textVal2 = new Cell().Add(new Paragraph("VAL")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER);
                tableSec2.AddCell(textVal2);

                Cell textVal3 = new Cell().Add(new Paragraph("TEXT")
                                        .SetFontColor(plomo)
                                        .SetBorderBottom(new SolidBorder(plomo, 1)))
                                        .SetBorder(Border.NO_BORDER);
                tableSec2.AddCell(textVal3);

                Cell textVal4 = new Cell().Add(new Paragraph("VAL")
                                        .SetFontColor(plomo)
                                        .SetBorderBottom(new SolidBorder(plomo, 1)))
                                        .SetBorder(Border.NO_BORDER);
                tableSec2.AddCell(textVal4);

                Cell textVal5 = new Cell().Add(new Paragraph("TEXT")
                                        .SetFontColor(negro)
                                        .SetBorderBottom(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER);
                tableSec2.AddCell(textVal5);

                Cell textVal6 = new Cell().Add(new Paragraph("VAL")
                                        .SetFontColor(negro)
                                        .SetBorderBottom(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER);
                tableSec2.AddCell(textVal6);

                Cell textVal7 = new Cell().Add(new Paragraph("TEXT")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER);
                tableSec2.AddCell(textVal7);

                Cell textVal8 = new Cell().Add(new Paragraph("VAL")
                                        .SetFontColor(verde)
                                        .SetBorderBottom(new SolidBorder(verde, 1)))
                                        .SetBorder(Border.NO_BORDER);
                tableSec2.AddCell(textVal8);

                Cell textVal9 = new Cell().Add(new Paragraph("TEXT")
                                        .SetFontColor(negro)
                                        .SetBorderBottom(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER);
                tableSec2.AddCell(textVal9);

                Cell textVal10 = new Cell().Add(new Paragraph("VAL")
                                        .SetFontColor(negro)
                                        .SetBorderBottom(new SolidBorder(negro, 1)))
                                        .SetBorder(Border.NO_BORDER);
                tableSec2.AddCell(textVal10);
            }
            
            containerTableSec2.Add(tableSec2);

            container2.Add(containerTableSec2);

            // Table 2 Seccion 2c 
            Div containerTable2Sec2 = new Div().SetMarginTop(3);

            Table table2Sec2 = new Table(UnitValue.CreatePercentArray(4)).UseAllAvailableWidth()
                                                                         .SetTextAlignment(TextAlignment.CENTER)
                                                                         .SetFontSize(sizeTabla);

            Cell sec2subtitulo2 = new Cell(1, 2).Add(new Paragraph("SUBTITULO2")
                                                .SetFontColor(blanco))
                                                .SetBorder(Border.NO_BORDER)
                                                .SetBackgroundColor(plomo)
                                                .SetMarginBottom(5);
            table2Sec2.AddCell(sec2subtitulo2);

            Cell sec2subtitulo21 = new Cell().Add(new Paragraph("SUBTITULO 2.1")
                                                .SetFontColor(blanco))
                                                .SetBorder(Border.NO_BORDER)
                                                .SetBackgroundColor(plomo)
                                                .SetMarginBottom(5);
            table2Sec2.AddCell(sec2subtitulo21);

            Cell sec2subtitulo22 = new Cell().Add(new Paragraph("SUBTITULO 2.2")
                                                .SetFontColor(blanco))
                                                .SetBorder(Border.NO_BORDER)
                                                .SetBackgroundColor(plomo)
                                                .SetMarginBottom(5);
            table2Sec2.AddCell(sec2subtitulo22);

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

            // Subititulo 3
            Div subtitulo3Sec2 = new Div().SetBackgroundColor(plomo)
                                   .SetWidth(container2.GetWidth())
                                   .SetMarginTop(2)
                                   .SetTextAlignment(TextAlignment.CENTER);

            subtitulo3Sec2.Add(new Paragraph("SUBTITULO 3").SetFontColor(blanco)
                                                 .SetFontSize(sizeSubtitulo));

            container2.Add(subtitulo3Sec2);

            // Table 2 Seccion 2
            Div containerTable3Sec2 = new Div().SetMarginTop(3);

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

            // Titulo 2
            Div titulo2 = new Div().SetBackgroundColor(verde)
                                   .SetWidth(container2.GetWidth())
                                   .SetMarginTop(2)
                                   .SetTextAlignment(TextAlignment.CENTER);

            titulo2.Add(new Paragraph("TITULO 2").SetFontColor(blanco)
                                                 .SetFontSize(sizeTitulo));

            container2.Add(titulo2);

            // Subtitulo 4
            Div subtitulo4Sec2 = new Div().SetBackgroundColor(plomo)
                                   .SetWidth(container2.GetWidth())
                                   .SetMarginTop(2)
                                   .SetTextAlignment(TextAlignment.CENTER);

            subtitulo4Sec2.Add(new Paragraph("SUBTITULO 4").SetFontColor(blanco)
                                                 .SetFontSize(sizeSubtitulo));

            container2.Add(subtitulo4Sec2);

            // Table 2 Seccion 2
            Div containerTable4Sec2 = new Div().SetMarginTop(3);

            Table table4Sec2 = new Table(UnitValue.CreatePercentArray(7)).UseAllAvailableWidth()
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

            // Subtitulo 5
            Div subtitulo5Sec2 = new Div().SetBackgroundColor(plomo)
                                   .SetWidth(container2.GetWidth())
                                   .SetMarginTop(2)
                                   .SetTextAlignment(TextAlignment.CENTER);

            subtitulo5Sec2.Add(new Paragraph("SUBTITULO 5").SetFontColor(blanco)
                                                 .SetFontSize(sizeSubtitulo));

            container2.Add(subtitulo5Sec2);

            // Table 5 Seccion 2
            Div containerTable5Sec2 = new Div().SetMarginTop(3);

            Table table5Sec2 = new Table(UnitValue.CreatePercentArray(7)).UseAllAvailableWidth()
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

            // ------------------------------------------------------------------------------------------------------------

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