# Export-As utilities

A convenience library for exporting data in different formats.

## Example

```java
List<List<Object>> data = getTestData();

byte[] xls  = ExportAs.xls(data,  "test");

byte[] xlsx = ExportAs.xlsx(data, "test");

byte[] csv  = ExportAs.csv(data,  "test");

byte[] html = ExportAs.html(data, "test");

byte[] pdf  = ExportAs.pdf(data,  "test");

byte[] json = ExportAs.json(data, "test");
```

## Build

- `git clone https://github.com/giosil/export-as.git`
- `mvn clean install`

## Contributors

* [Giorgio Silvestris](https://github.com/giosil)
