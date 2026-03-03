clos all
clea all
clea

SQLServer = "TRAVERSE"
SQLUsr = "uid=sa;pwd=janiTR@verse"
SQLRegion = "INT"
*SQLRegion = "ADV"

rptDate 	= {09/27/2024}
fromDate	= dtoc(rptDate)
toDate 		= dtoc(rptDate+1)

xSql = "select space(6) as company_no, ourAcctNum, space(1) as RecType, "
xSql = xSql + "CheckNum, netpaidcalc as check_amt, "
xSql = xSql + "checkdate, "
xSql = xSql + "PayToName, "
xSql = xSql + "VoidDate, Voidyn "
xSql = xSql + "from tblAPCheckHist ck "
xSql = xSql + "inner join tblSmBankAcct ba on ck.bankid = ba.bankid "
xSql = xSql + "where left(CheckNum,3) <> 'ACH' and left(CheckNum,3) <> 'CLR' and left(CheckNum,3) <> 'WIR' " 
xSql = xSql + "and (checkdate = '" + fromDate + "' or checkdate = '" + toDate + "') "
xSql = xSql + "and left(ba.Name,12) = 'North Dallas' "

*--> Create Connection to SQL Database
oConn=sqlStringConnect("Driver={SQL Server};Server="+SQLServer+";"+SQLUsr+";Database="+SQLRegion)

*--> Execute Query 
x=SQLEXEC(oConn, xsql, 'myCursor') 

sele myCursor
brow last
