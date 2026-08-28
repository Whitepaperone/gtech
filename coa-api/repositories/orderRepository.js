'use strict';

const database = require('../config/database');

const ORDERS_SQL = `
SELECT
    T.SO_CODE, T.ORDER_DATE, T.DUE_DATE,
    T.CUSTOMER_CODE, T.CUSTOMER_NAME,
    T.ITEM_CODE, T.ITEM_NAME, T.ITEM_DESCRIPTION,
    T.ORDER_QTY, T.CANCEL_QTY, T.VALID_QTY, T.PROCESSED_QTY,
    (T.VALID_QTY - T.PROCESSED_QTY) AS REMAIN_QTY,
    T.UNIT, T.PRICE, T.REMARK,
    T.PROJECT_CODE, T.PROJECT_NAME, T.STATUS, T.DETAIL_INDEX
FROM (
    SELECT
        S.Code AS SO_CODE,
        S.DDay AS ORDER_DATE,
        D.DueDate AS DUE_DATE,
        S.CustomerCode AS CUSTOMER_CODE,
        C.Name AS CUSTOMER_NAME,
        D.ItemCode AS ITEM_CODE,
        I.Name AS ITEM_NAME,
        I.Description AS ITEM_DESCRIPTION,
        D.Qty AS ORDER_QTY,
        D.CancelQty AS CANCEL_QTY,
        (D.Qty - D.CancelQty) AS VALID_QTY,
        (
            SELECT COALESCE(SUM(PD.Qty - PD.CancelQty), 0)
            FROM SALESDETAIL PD
            JOIN BASECODE B ON PD.iIndex = B.ObjIndex
            WHERE B.SrcTable = 6
              AND B.SrcIndex = D.iIndex
              AND B.ObjTable = 7
        ) AS PROCESSED_QTY,
        D.Unit AS UNIT,
        D.Price AS PRICE,
        D.Remark AS REMARK,
        S.ProjectCode AS PROJECT_CODE,
        (SELECT FIRST 1 P.Name FROM PROJECT P WHERE P.Code = S.ProjectCode) AS PROJECT_NAME,
        S.Status AS STATUS,
        D.iIndex AS DETAIL_INDEX
    FROM SALESORDER S
    JOIN SODETAIL D ON D.SOCode = S.Code
    LEFT JOIN CUSTOMER C ON C.Code = S.CustomerCode
    LEFT JOIN ITEM I ON I.Code = D.ItemCode
    WHERE S.Status IN (1, 2)
) T
WHERE (T.VALID_QTY - T.PROCESSED_QTY) > 0
ORDER BY T.DUE_DATE, T.SO_CODE, T.DETAIL_INDEX
`;

function findOpenOrders() {
  return database.query(ORDERS_SQL);
}

module.exports = { findOpenOrders, ORDERS_SQL };
