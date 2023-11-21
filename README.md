
Empezar a construir a partir de (Primera Pagina): 
```           
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
                          .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4);

// Primera Seccion - Primer Parrafo

// Primera Seccion - Subtitulo 1

// Seccion 1 - Parrafo 2

// Seccion 1 - Subtitulo 2

// Seccion 1 - Parrafo 3

// Seccion 1 - Subtitulo 3

// Seccion 1 - Parrafo 4

// Seccion 1 - Subtitulo 4

// Seccion 1 - Parrafo 5

// Seccion 1 - Parrafo 6

// Seccion 1 - Data Cliente

// Seccion 1 - Parrafo 7

// Seccion 2
Div container2 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 70f / 100f - 4)
                                       .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4)
                                       .SetBackgroundColor(blanco);

// Seccion 2 - Header

// Seccion 2 - Titulo 1

// Seccion 2 - Subtitulo 1         

// Seccion 2 - Table 1

// Seccion 2 - Tabla 2

// Seccion 2 - Subititulo 3

// Seccion 2 - Tabla 3

// Seccion 2 - Titulo 2

// Seccion 2 - Subtitulo 4

// Seccion 2 - Table 4 

// Seccion 2 - Subtitulo 5

// Seccion 2 - Table 5


Table tablePage1 = new (2);
tablePage1.AddCell(new Cell().Add(container1).SetBorder(Border.NO_BORDER));
tablePage1.AddCell(new Cell().Add(container2).SetBorder(Border.NO_BORDER));
document.Add(tablePage1);

// SEGUNDA PAGINA
// Tercera Seccion
Div container3 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 50f / 100f - 4)
                            .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4)
                            .SetBackgroundColor(ColorConstants.YELLOW);

Div container4 = new Div().SetWidth(pdf.GetDefaultPageSize().GetWidth() * 50f / 100f - 4)
                            .SetHeight(pdf.GetDefaultPageSize().GetHeight() - 4)
                            .SetBackgroundColor(ColorConstants.GRAY);

Table tablePage2 = new (2);
tablePage2.AddCell(new Cell().Add(container3).SetBorder(Border.NO_BORDER));
tablePage2.AddCell(new Cell().Add(container4).SetBorder(Border.NO_BORDER));
document.Add(tablePage2);

document.Close();   

Response.Headers.Add("Content-Disposition", "attachment; filename=reporte.pdf");
Response.Headers.Add("Content-Type", "application/pdf");
return File(memoryStream.ToArray(), "application/pdf");
```

![Estructura](/estructura.jpeg)