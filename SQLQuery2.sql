WITH noi_orig
AS (
	SELECT r.ItemCode
		,d.NumAtCard
		,r.OpenCreQty nmatr
		,'noi' AS lato
	--,r.DocEntry
	--,r.LineNum
	--,d.DocNum
	--,d.TaxDate
	FROM PDN1 r
	JOIN OPDN d ON d.DocEntry = r.DocEntry
	JOIN OITM a ON r.ItemCode = a.ItemCode
	WHERE d.CardCode = 'F0029'
		AND r.LineStatus = 'O'
		AND a.ManSerNum = 'Y'
	)
	,noi
AS (
	SELECT r.ItemCode
		,d.NumAtCard
		,sum(r.OpenCreQty) AS nmatr
		,'noi' AS lato
	--,r.DocEntry
	--,r.LineNum
	--,d.DocNum
	--,d.TaxDate
	FROM PDN1 r
	JOIN OPDN d ON d.DocEntry = r.DocEntry
	JOIN OITM a ON r.ItemCode = a.ItemCode
	WHERE d.CardCode = 'F0029'
		AND r.LineStatus = 'O'
		AND a.ManSerNum = 'Y'
	GROUP BY r.ItemCode
		,d.NumAtCard
	)
	,loro
AS (
	SELECT itemcode collate SQL_Latin1_General_CP850_CI_AS AS itemcode
		,numatcard collate SQL_Latin1_General_CP850_CI_AS AS numatcard
		,count(distnumber) AS 'nmatr'
		,'loro' AS 'lato'
	FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0;HDR=YES;Database=C:\Users\ingrassiaadm\Desktop\ddt_gn.xlsx', 'select * from [loro$]')
	GROUP BY NumAtCard
		,itemcode
	)
SELECT l.*
	,n.*
	,isnull(l.nmatr, 0) - isnull(n.nmatr, 0) AS 'to do on SAP',LEFT(n.numatcard, CHARINDEX(' ',n.numatcard)),right(n.numatcard, CHARINDEX(' ',n.numatcard))
FROM loro l
FULL OUTER JOIN noi n ON l.itemcode = n.ItemCode
	AND l.numatcard = n.NumAtCard
--AND l.nmatr = n.nmatr
WHERE (
		isnull(l.nmatr, 0) - isnull(n.nmatr, 0) <> 0
		--		l.itemcode IS NULL
		--		OR n.ItemCode IS NULL
		)
	AND (
		l.itemcode = '18182200'
		OR n.ItemCode = '18182200'
		)
ORDER BY l.itemcode
	,l.numatcard
