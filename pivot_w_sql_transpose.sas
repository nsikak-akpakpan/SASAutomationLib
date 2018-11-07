/*create pivot tables with Proc SQL and Transpose */

PROC SQL;
CREATE TABLE CategorySales AS
SELECT Category, Product, Sum(Sales) as Sales
FROM SampleData
GROUP BY Category, Product
ORDER BY Category, Sales Desc;
RUN;

PROC SQL;
ALTER TABLE CategorySales
ADD SalesPerson Character(25) label = “Sales Person”;
UPDATE CategorySales
SET SalesPerson = “JOHN DOE”;
TITLE “Adding SalesPerson Column”;
SELECT Category, Product, Sales, SalesPerson
FROM CategorySales;
RUN;

* Pivot dataset by Quarter variable;
PROC TRANSPOSE DATA=ProductSales OUT=ProductSalesByQuarters NAME=Sales;
BY Category Product;
VAR Sales;
ID Quarter;
RUN;
