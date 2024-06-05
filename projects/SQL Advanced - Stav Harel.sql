USE AdventureWorks2019;

--- Mission 1
SELECT pro.ProductID,
	pro.Name,
	pro.Color,
	pro.ListPrice,
	pro.Size
FROM Production.Product pro
WHERE pro.ProductID NOT IN (
	SELECT sod.ProductID
	FROM Sales.SalesOrderDetail sod
)
ORDER BY pro.ProductID;

--- Mission 2 ---
SELECT c.CustomerID,
	pp.BusinessEntityID AS PersonID,
	ISNULL(pp.LastName, 'Unknown') AS LastName,
	ISNULL(pp.FirstName, 'Unknown') AS FirstName
FROM Sales.Customer c
	LEFT JOIN Person.Person pp
		ON pp.BusinessEntityID = c.CustomerID
WHERE c.CustomerID NOT IN (
	SELECT soh.CustomerID
	FROM Sales.SalesOrderHeader soh
) AND CustomerID > 273
	AND CustomerID < 692
ORDER BY c.CustomerID

--- Mission 3
SELECT TOP 10
	c.CustomerID,
	pp.FirstName,
	pp.LastName,
	COUNT(soh.SalesOrderID) AS CountOfOrders
FROM Sales.Customer c
	JOIN Person.Person pp 
		ON c.PersonID = pp.BusinessEntityID
	JOIN Sales.SalesOrderHeader soh 
		ON c.CustomerID = soh.CustomerID
GROUP BY c.CustomerID,
	pp.FirstName,
	pp.LastName
ORDER BY CountOfOrders DESC,
	c.CustomerID ASC;

--- Mission 4
WITH
	EmployeeCounts
	AS
	(
		SELECT e.JobTitle,
			COUNT(*) AS CountOfTitle
		FROM HumanResources.Employee e
		GROUP BY e.JobTitle
	)
SELECT
	pp.FirstName,
	pp.LastName,
	e.JobTitle,
	e.HireDate,
	ec.CountOfTitle
FROM HumanResources.Employee e
	JOIN EmployeeCounts ec 
		ON e.JobTitle = ec.JobTitle
	JOIN Person.Person pp 
		ON e.BusinessEntityID = pp.BusinessEntityID;

--- Mission 5

WITH
	MostRecentOrder
	AS
	(
		SELECT
			SalesOrderID,
			CustomerID,
			OrderDate,
			ROW_NUMBER() OVER (PARTITION BY CustomerID ORDER BY OrderDate DESC) AS OrderRank
		FROM Sales.SalesOrderHeader
	),
	PreviousOrder
	AS
	(
		SELECT *
		FROM MostRecentOrder
		WHERE OrderRank = 2
	)
SELECT
	mro.SalesOrderID,
	mro.CustomerID,
	pp.LastName,
	pp.FirstName,
	mro.OrderDate AS CurrentOrderDate,
	po.OrderDate AS PreviousOrderDate
FROM MostRecentOrder mro
	LEFT JOIN PreviousOrder po 
		ON mro.CustomerID = po.CustomerID
	JOIN Sales.Customer c 
		ON mro.CustomerID = c.CustomerID
	JOIN Person.Person pp 
		ON c.PersonID = pp.BusinessEntityID
WHERE mro.OrderRank = 1;

--- Mission 6

WITH
	MostExpensiveOrderPerYear
	AS
	(
		SELECT
			YEAR(soh.OrderDate) AS Year,
			soh.SalesOrderID,
			soh.CustomerID,
			SUM(sod.UnitPrice * (1 - sod.UnitPriceDiscount) * sod.OrderQty) AS Total,
			ROW_NUMBER() OVER (PARTITION BY YEAR(soh.OrderDate) 
                           ORDER BY SUM(sod.UnitPrice * (1 - sod.UnitPriceDiscount) * sod.OrderQty)
						   	DESC) AS RowNum
		FROM Sales.SalesOrderDetail sod
			JOIN Sales.SalesOrderHeader soh
				ON sod.SalesOrderID = soh.SalesOrderID
		GROUP BY YEAR(soh.OrderDate),
			soh.SalesOrderID,
				soh.CustomerID
	)

SELECT
	meopy.Year,
	meopy.SalesOrderID,
	Customer.PersonID AS Customer,
	Person.LastName,
	Person.FirstName,
	FORMAT(meopy.Total, 'N1') AS Total
FROM MostExpensiveOrderPerYear meopy
	JOIN Sales.Customer 
		ON meopy.CustomerID = Customer.CustomerID
	JOIN Person.Person 
		ON Customer.PersonID = Person.BusinessEntityID
WHERE meopy.RowNum = 1
ORDER BY meopy.Year;

--- Mission 7

WITH
	MonthlyOrderCounts
	AS
	(
		SELECT
			MONTH(OrderDate) AS Month,
			YEAR(OrderDate) AS Year,
			COUNT(*) AS OrderCount
		FROM
			Sales.SalesOrderHeader
		GROUP BY
        	MONTH(OrderDate),
        	YEAR(OrderDate)
	)
SELECT
	Month,
	COALESCE([2011], 0) AS [2011],
	COALESCE([2012], 0) AS [2012],
	COALESCE([2013], 0) AS [2013],
	COALESCE([2014], 0) AS [2014]
FROM
	MonthlyOrderCounts
PIVOT (
    SUM(OrderCount) FOR Year IN ([2011], [2012], [2013], [2014])
) AS PivotTable
ORDER BY Month;

--- Mission 8

DROP TABLE IF EXISTS #MonthlyTotals;
DROP TABLE IF EXISTS #FinalResult;

CREATE TABLE #MonthlyTotals
(
	[Year] INT,
	[Month] INT,
	Sum_Price DECIMAL(18,2)
);

INSERT INTO #MonthlyTotals
SELECT
	YEAR(soh.OrderDate) AS [Year],
	MONTH(soh.OrderDate) AS [Month],
	SUM((sod.UnitPrice * (1 - sod.UnitPriceDiscount)) * sod.OrderQty) AS Sum_Price
FROM Sales.SalesOrderHeader soh
	JOIN Sales.SalesOrderDetail sod
		ON soh.SalesOrderID = sod.SalesOrderID
GROUP BY YEAR(soh.OrderDate),
	MONTH(soh.OrderDate);

CREATE TABLE #FinalResult
(
	[Year] INT,
	[Month] VARCHAR(15),
	Sum_Price VARCHAR(20),
	CumSum VARCHAR(20),
	SortOrder INT
);

DECLARE @Year INT,
	@Month INT,
	@Sum_Price DECIMAL(18,2),
	@CumSum DECIMAL(18,2) = 0;

DECLARE TotalsCursor CURSOR FOR
SELECT [Year], [Month], Sum_Price
FROM #MonthlyTotals
ORDER BY [Year], [Month];

OPEN TotalsCursor;

FETCH NEXT FROM TotalsCursor 
	INTO @Year,
		@Month,
		@Sum_Price;

WHILE @@FETCH_STATUS = 0
BEGIN
	SET @CumSum = @CumSum + @Sum_Price;

	INSERT INTO #FinalResult
		(
			[Year],
			[Month],
			Sum_Price,
			CumSum,
			SortOrder
		)
	VALUES
		(
			@Year,
			CAST(@Month AS VARCHAR(10)),
			FORMAT(@Sum_Price, 'N'),
			FORMAT(@CumSum, 'N'), 0
		);

	IF (
		@Month = 12 OR (
		SELECT COUNT(*)
		FROM #MonthlyTotals
		WHERE [Year] = @Year AND [Month] > @Month) = 0
	)
    SET @CumSum = 0;

	FETCH NEXT FROM TotalsCursor 
		INTO @Year,
			@Month,
			@Sum_Price;
END

CLOSE TotalsCursor;
DEALLOCATE TotalsCursor;

INSERT INTO #FinalResult
(
	[Year],
	[Month],
	Sum_Price,
	CumSum,
	SortOrder
)

SELECT [Year],
	'Grand Total:',
	NULL,
	FORMAT(SUM(Sum_Price), 'N'), 1
FROM #MonthlyTotals
GROUP BY [Year];

SELECT [Year],
	[Month],
	Sum_Price,
	CumSum
FROM #FinalResult
ORDER BY 
    [Year],
    SortOrder,
    CASE WHEN [Month] = 'Grand Total:' THEN 'Grand Total:' 
		ELSE RIGHT('0' + [Month], 2) 
			END;

--- Mission 9

WITH
	EmployeeHierarchy
	AS
	(
		SELECT
			e.BusinessEntityID AS EmployeeId,
			pp.FirstName + ' ' + pp.LastName AS FullName,
			e.HireDate,
			edh.DepartmentID,
			(DATEDIFF(MONTH, e.HireDate, GETDATE())- 11) AS Seniority,
			LAG(pp.FirstName + ' ' + pp.LastName) 
				OVER (PARTITION BY edh.DepartmentID ORDER BY e.HireDate) AS PreviousEmpName,
			LAG(e.HireDate) 
				OVER (PARTITION BY edh.DepartmentID ORDER BY e.HireDate) AS PreviousEmpHDate
		FROM
			HumanResources.Employee e
				JOIN HumanResources.EmployeeDepartmentHistory edh 
					ON e.BusinessEntityID = edh.BusinessEntityID
				JOIN Person.Person pp 
					ON e.BusinessEntityID = pp.BusinessEntityID
),
	DepartmentEmployees
	AS
	(
		SELECT
			d.Name AS DepartmentName,
			eh.EmployeeId,
			eh.FullName,
			eh.HireDate,
			eh.Seniority,
			NULLIF(eh.PreviousEmpName, '') AS PreviousEmpName,
			eh.PreviousEmpHDate,
			ABS(COALESCE(DATEDIFF(DAY, eh.PreviousEmpHDate, eh.HireDate), 0)) AS DiffDays
		FROM
			EmployeeHierarchy eh
				JOIN HumanResources.Department d 
					ON eh.DepartmentID = d.DepartmentID
	)
SELECT
	de.DepartmentName,
	de.EmployeeId AS [Employee's Id],
	de.FullName AS [Employee's FullName],
	de.HireDate,
	de.Seniority,
	de.PreviousEmpName,
	de.PreviousEmpHDate,
	de.DiffDays
FROM
	DepartmentEmployees de
ORDER BY
    de.DepartmentName, 
	de.HireDate DESC; 

--- Mission 10

WITH
	EmployeeHierarchy
	AS
	(
		SELECT
			e.BusinessEntityID,
			e.OrganizationNode,
			e.HireDate,
			(
				SELECT TOP 1
					edh.DepartmentID
				FROM HumanResources.EmployeeDepartmentHistory edh
				WHERE edh.BusinessEntityID = e.BusinessEntityID
				ORDER BY edh.StartDate DESC
			) AS DepartmentID,
			CONCAT(e.BusinessEntityID, ' ', pp.LastName, ' ', pp.FirstName) AS FullName
		FROM HumanResources.Employee e
			JOIN Person.Person pp 
				ON e.BusinessEntityID = pp.BusinessEntityID
	)
SELECT
	eh.HireDate,
	eh.DepartmentID,
	STRING_AGG(FullName, ', ') WITHIN GROUP (ORDER BY OrganizationNode) AS TeamEmployees
FROM EmployeeHierarchy eh
GROUP BY HireDate,
	DepartmentID
ORDER BY HireDate DESC;