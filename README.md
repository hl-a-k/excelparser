Excat informations from excel
```$xslt
boolean isValid = parser.verify(excelFile, errorExcel, cfg);
```
verify the excel file. If find any error, will generate a new excel file. All error will be marked.

```$xslt
 List<HashMap> list = parser.extract(HashMap.class);
```
excat informations from excel
