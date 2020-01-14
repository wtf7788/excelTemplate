# excelTemplate
使用excel制作模板生成对应的pdf 或者图片

使用方法：
String excelPath ="e:/asd.xls"; //指定模板路径
//默认只读取第一个sheet 中的模板内容
//ExcelExReader templateReader = ExcelTemplateUtil.getReader(excelPath2,1); 可以指定读取第几个sheet
ExcelExReader templateReader = ExcelTemplateUtil.getReader(excelPath);   
ExcelObject excelObject = new ExcelObject(templateReader,new FileOutputStream(new File("e:/templateConvertPdf.pdf"))); //指定生成pdf路径
excelObject.convert();
