'use strict';

const database = require('../config/database');

// Kept equivalent to the legacy query. The date value is always parameter-bound.
const SHIPMENT_SQL = `
SELECT DISTINCT
    D.iIndex AS IINDEX, D.LineNo AS LINENO,
    S.DDay AS DDAY, S.Code AS SALESCODE,
    I.Category AS CATEGORY, I.Code AS ITEMCODE,
    I.Name AS ITEMNAME, I.Description AS DESCRIPTION,
    D.SubClass AS SUBCLASS, D.WaitingQty AS WAITINGQTY,
    D.Qty AS QTY, D.CancelQty AS RETURNQTY,
    D.Qty - D.CancelQty AS RSQTY,
    D.Unit AS UNIT, D.Price AS UNITPRICE,
    CASE WHEN (D.Qty - D.CancelQty) = 0 THEN 0
         ELSE D.Tax / (D.Qty - D.CancelQty) END AS UNITTAX,
    ((D.Qty - D.CancelQty) * D.Price) AS SALESPRICE,
    D.Tax AS STAX,
    (((D.Qty - D.CancelQty) * D.Price) + D.Tax) AS SUBTOTAL,
    (D.Cost * (D.Qty - D.CancelQty)) AS COST,
    (((D.Qty - D.CancelQty) * D.Price) - ((D.Qty - D.CancelQty) * D.Cost)) AS PROFIT,
    COALESCE(A.Currency, '') AS CURRENCY,
    D.Volume AS VOLUME, (D.Qty - D.CancelQty) * D.Volume AS VSUBTOTAL,
    D.VolumeUnit AS VOLUMEUNIT,
    D.Weight AS WEIGHT, (D.Qty - D.CancelQty) * D.Weight AS WSUBTOTAL,
    D.WeightUnit AS WEIGHTUNIT,
    S.StaffName AS STAFFNAME,
    S.CustomerCode AS CUSTOMERCODE, C.Name AS CUSTOMERNAME,
    C.Phone AS PHONE, C.Address AS ADDRESS,
    S.OutLocation AS OUTLOCATION, L.Name AS LOCATIONNAME,
    D.Remark AS REMARK, D.BaseType AS BASETYPE
FROM SALES AS S
JOIN SALESDETAIL AS D ON S.Code = D.SalesCode
JOIN ITEM AS I ON D.ItemCode = I.Code
LEFT JOIN CUSTOMER AS C ON C.Code = S.CustomerCode
LEFT JOIN LOCATION AS L ON L.Code = S.LocationCode
LEFT JOIN (
    SELECT SourceCode, Currency FROM ACCOUNT WHERE PSType = 1
) AS A ON S.Code = A.SourceCode
WHERE S.TradeType = 0
  AND S.DDay <= CAST(? AS DATE)
ORDER BY S.Code, D.LineNo
`;

function findThroughDate(throughDate) {
  return database.query(SHIPMENT_SQL, [throughDate]);
}

module.exports = { findThroughDate, SHIPMENT_SQL };
