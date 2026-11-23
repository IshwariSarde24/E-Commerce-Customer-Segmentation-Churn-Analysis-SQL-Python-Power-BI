/*=============================================================================
  PROJECT : IKEA Gifting - Customer Analytics & RFM Segmentation
  DATABASE: EcommerceCustomerAnalytics
  PURPOSE : Combine the two Online Retail II source tables, validate the data,
            remove exact duplicate transaction lines, and create a clean table.
  TOOL    : Microsoft SQL Server

  EXPECTED SOURCE TABLES:
      dbo.[Year 2009-2010]
      dbo.[Year 2010-2011]

  OUTPUT TABLES:
      dbo.RetailCombined
      dbo.RetailCleaned

  NOTE:
      Import the two source-year tables into this database before running
      this script. Re-running this script replaces only the two output tables.
=============================================================================*/

/*=============================================================================
  1. CREATE AND SELECT DATABASE
=============================================================================*/
IF DB_ID(N'EcommerceCustomerAnalytics') IS NULL
BEGIN
    CREATE DATABASE EcommerceCustomerAnalytics;
END;
GO

USE EcommerceCustomerAnalytics;
GO

/*=============================================================================
  2. CHECK SOURCE DATA ROW COUNTS
=============================================================================*/
SELECT COUNT(*) AS Rows_2009_2010
FROM dbo.[Year 2009-2010];

SELECT COUNT(*) AS Rows_2010_2011
FROM dbo.[Year 2010-2011];

/*=============================================================================
  3. PREPARE OUTPUT TABLES FOR A RE-RUN

  The source tables are not dropped or modified. Only the output tables
  created by this script are replaced.
=============================================================================*/
DROP TABLE IF EXISTS dbo.RetailCleaned;
DROP TABLE IF EXISTS dbo.RetailCombined;
GO

/*=============================================================================
  4. COMBINE BOTH YEARS AND CONVERT DATA TYPES

  - UNION ALL appends records from both source tables.
  - TRY_CONVERT safely converts date, quantity, and price values.
  - COALESCE tries the supported date formats in sequence.
  - SourceSheet records which source table each row came from.
=============================================================================*/
SELECT
    Invoice,
    StockCode,
    Description,
    TRY_CONVERT(INT, Quantity) AS Quantity,
    COALESCE(
        TRY_CONVERT(DATETIME2, InvoiceDate, 120),
        TRY_CONVERT(DATETIME2, InvoiceDate, 105)
    ) AS InvoiceDate,
    TRY_CONVERT(DECIMAL(10, 2), Price) AS Price,
    Customer_ID,
    Country,
    CAST('2009-2010' AS VARCHAR(9)) AS SourceSheet
INTO dbo.RetailCombined
FROM dbo.[Year 2009-2010]

UNION ALL

SELECT
    Invoice,
    StockCode,
    Description,
    TRY_CONVERT(INT, Quantity) AS Quantity,
    COALESCE(
        TRY_CONVERT(DATETIME2, InvoiceDate, 120),
        TRY_CONVERT(DATETIME2, InvoiceDate, 105)
    ) AS InvoiceDate,
    TRY_CONVERT(DECIMAL(10, 2), Price) AS Price,
    Customer_ID,
    Country,
    CAST('2010-2011' AS VARCHAR(9)) AS SourceSheet
FROM dbo.[Year 2010-2011];
GO

/*=============================================================================
  5. VALIDATE DATA TYPE CONVERSIONS
=============================================================================*/
SELECT
    COUNT(*) AS Total_Rows,
    SUM(CASE WHEN InvoiceDate IS NULL THEN 1 ELSE 0 END) AS Failed_Dates,
    SUM(CASE WHEN Quantity IS NULL THEN 1 ELSE 0 END) AS Failed_Quantities,
    SUM(CASE WHEN Price IS NULL THEN 1 ELSE 0 END) AS Failed_Prices
FROM dbo.RetailCombined;

/*=============================================================================
  6. CHECK DATA QUALITY
=============================================================================*/

-- 6.1 Missing customer IDs
SELECT COUNT(*) AS Missing_Customer_ID
FROM dbo.RetailCombined
WHERE Customer_ID IS NULL;

-- 6.2 Zero or negative quantities
-- These records are flagged for review, not automatically deleted.
SELECT COUNT(*) AS Zero_or_Negative_Quantity
FROM dbo.RetailCombined
WHERE Quantity <= 0;

-- 6.3 Missing product descriptions
SELECT COUNT(*) AS Missing_Description
FROM dbo.RetailCombined
WHERE Description IS NULL;

-- 6.4 Missing invoice dates
SELECT COUNT(*) AS Missing_InvoiceDate
FROM dbo.RetailCombined
WHERE InvoiceDate IS NULL;

-- 6.5 Missing or invalid prices
SELECT COUNT(*) AS Missing_or_Invalid_Price
FROM dbo.RetailCombined
WHERE Price IS NULL;

/*=============================================================================
  7. IDENTIFY EXACT DUPLICATE TRANSACTION ROWS

  SourceSheet is intentionally excluded from the duplicate definition so
  identical transaction records can be identified across both source tables.
=============================================================================*/
SELECT
    Invoice,
    StockCode,
    Description,
    Quantity,
    InvoiceDate,
    Price,
    Customer_ID,
    Country,
    COUNT(*) AS ExactDuplicateCount
FROM dbo.RetailCombined
GROUP BY
    Invoice,
    StockCode,
    Description,
    Quantity,
    InvoiceDate,
    Price,
    Customer_ID,
    Country
HAVING COUNT(*) > 1
ORDER BY ExactDuplicateCount DESC;

/*=============================================================================
  8. REMOVE EXACT DUPLICATES AND CREATE THE CLEANED TABLE

  ROW_NUMBER assigns a sequence to each group of identical transaction rows.
  One row is retained from each group. SourceSheet is not part of PARTITION BY,
  so duplicates across the two source sheets are also considered.

  If identical records occur in both source sheets, the 2009-2010 source label
  is retained first. The transaction values themselves are identical.
=============================================================================*/
WITH Deduped AS
(
    SELECT
        Invoice,
        StockCode,
        Description,
        Quantity,
        InvoiceDate,
        Price,
        Customer_ID,
        Country,
        SourceSheet,
        ROW_NUMBER() OVER
        (
            PARTITION BY
                Invoice,
                StockCode,
                Description,
                Quantity,
                InvoiceDate,
                Price,
                Customer_ID,
                Country
            ORDER BY
                CASE
                    WHEN SourceSheet = '2009-2010' THEN 1
                    ELSE 2
                END
        ) AS RowNum
    FROM dbo.RetailCombined
)
SELECT
    Invoice,
    StockCode,
    Description,
    Quantity,
    InvoiceDate,
    Price,
    Customer_ID,
    Country,
    SourceSheet
INTO dbo.RetailCleaned
FROM Deduped
WHERE RowNum = 1;
GO

/*=============================================================================
  9. VALIDATE DEDUPLICATION RESULTS
=============================================================================*/
SELECT
    (SELECT COUNT(*) FROM dbo.RetailCombined) AS OriginalRows,
    (SELECT COUNT(*) FROM dbo.RetailCleaned) AS CleanedRows,
    (SELECT COUNT(*) FROM dbo.RetailCombined)
      - (SELECT COUNT(*) FROM dbo.RetailCleaned) AS DuplicateRowsRemoved;

/*=============================================================================
  10. REVIEW THE FINAL CLEANED DATASET
=============================================================================*/

-- Total number of cleaned records
SELECT COUNT(*) AS Rows_Count_Cleaned
FROM dbo.RetailCleaned;

-- Preview the first 100 cleaned records
SELECT TOP (100)
    Invoice,
    StockCode,
    Description,
    Quantity,
    InvoiceDate,
    Price,
    Customer_ID,
    Country,
    SourceSheet
FROM dbo.RetailCleaned
ORDER BY InvoiceDate, Invoice;

-- Final data quality summary
SELECT
    COUNT(*) AS Total_Cleaned_Rows,
    SUM(CASE WHEN Customer_ID IS NULL THEN 1 ELSE 0 END) AS Missing_Customer_ID,
    SUM(CASE WHEN Description IS NULL THEN 1 ELSE 0 END) AS Missing_Description,
    SUM(CASE WHEN InvoiceDate IS NULL THEN 1 ELSE 0 END) AS Missing_InvoiceDate,
    SUM(CASE WHEN Quantity IS NULL THEN 1 ELSE 0 END) AS Invalid_Quantity,
    SUM(CASE WHEN Price IS NULL THEN 1 ELSE 0 END) AS Invalid_Price
FROM dbo.RetailCleaned;

/*=============================================================================
  END OF SCRIPT

  Final analysis table: dbo.RetailCleaned
=============================================================================*/
