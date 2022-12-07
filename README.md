# excel2pdf
---

Excel(.xls or .xlsx) file convert to PDF file
- JDK1.8
- Apache POI
- Itext7

## Use
```java
public class Excel2PDFTest {
    public static void main(String[] args){
        try(InputStream is = new FileInputStream("you_excel_file_path.xlsx");
            OutputStream os = new FileOutputStream("generated_pdf_file_path.pdf")
        ) {
            Excel2PDF.process(is, os, document -> {
                // set A4 Page size, rotated
                document.getPdfDocument().setDefaultPageSize(PageSize.A4.rotate());
                // set margin, default value is 36.0F
                document.setTopMargin(12.0F);
                document.setRightMargin(6.0F);
                document.setBottomMargin(12.0F);
                document.setLeftMargin(6.0F);
            });
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
```