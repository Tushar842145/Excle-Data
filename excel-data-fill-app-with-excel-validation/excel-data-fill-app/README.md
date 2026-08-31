# Excel Data Fill Application

## Run
```bash
mvn spring-boot:run
```

Swagger:
```text
http://localhost:8080/swagger-ui.html
```

API:
```text
POST /api/excel/fill
```

Multipart fields:
```text
templateFile = Excel template .xlsx
jsonFile     = JSON file .json
```

## Important
- Row 1 = header.
- Row 2 = mapping, for example `employees.Id`, `employees.salary`, `companies.loc.location`.
- Object name comes from Row 2 first part.
- Sheet name can be anything.
- Existing Excel validation is read from template before data insertion.
- If an existing Excel validation fails for any mapped column, the full row is skipped.
- Data values are center aligned.
- Case-insensitive JSON path lookup is supported.
- Empty formatted rows are detected correctly, so data starts from Row 3 if Row 3 is empty.
