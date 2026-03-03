clos all
clea all
clea
set cent on

rptDate = " ckdate = date()-1 "
rptDate2 = date()-1
rptName = "PosPay_" + left(dtoc(rptDate2),2)+substr(dtoc(rptDate2),4,2)+right(dtoc(rptDate2),4)

*rptDate = " between(ckdate,date()-3,date()+1) "
*rptName = "PossPay_05_01_2015"

*use i:\frrpts\data_loc in 0 order region
use \\jkfradb\fr_acct\FrAcctRpts\Users\jillbean\data_loc in 0 order region

sele data_loc
set filt to !empty(lockbox) and !empty(traverse_d)

create cursor tmpData ;
	(company_no C(6), ;
	AcctNum C(12), ;
	RecType C(1), ;
	CheckNum C(10), ;
	Amount C(12), ;
	xAmount N(12,2), ;
	CkDate C(8), ;
	Fill C(1))
	
sele data_loc
go top
do while !eof()
	curRegion  = alltrim(Data_Loc.region)
	xPath = alltrim(data_loc.bank_data) + "\"
	do GetData
	sele data_loc
	skip
enddo
wait clear
sele tmpData
go top

*xyPath=sys(5)+sys(2003)+"\"
xyPath="i:\tmp\"
xFileOut=xyPath+rptName

SET CONSOLE OFF
SET TEXTMERGE ON
SET TEXTMERGE TO (xFileOut)
* \\ Stay on same line, \ new line

cTxt=tmpData.AcctNum
\\<<cTxt>>
cTxt=tmpData.RecType
\\<<cTxt>>
cTxt=tmpData.CheckNum
\\<<cTxt>>
cTxt=tmpData.Amount
\\<<cTxt>>
cTxt=tmpData.CkDate
\\<<cTxt>>
cTxt=tmpData.Fill
\\<<cTxt>>
skip

do while !eof()
	*--> Write File
	cTxt=tmpData.AcctNum
	\<<cTxt>>
	cTxt=tmpData.RecType
	\\<<cTxt>>
	cTxt=tmpData.CheckNum
	\\<<cTxt>>
	cTxt=tmpData.Amount
	\\<<cTxt>>
	cTxt=tmpData.CkDate
	\\<<cTxt>>
	cTxt=tmpData.Fill
	\\<<cTxt>>
	skip
enddo

SET TEXTMERGE TO
SET CONSOLE ON

sele tmpData
sum all xAmount to xAmt
append blank
repla amount with "Total"
repla xAmount with xAmt

copy to "i:\tmp\" + rptName type xls
clos all

proc GetData
	wait window "Now Processing " + curRegion + "..." nowait
	use xPath + "jkcmpfil" in 0
	sele jkcmpfil
	loca for account = "1005" 
	curAcct = alltrim(jkcmpfil.acct_num)
	curAcct = strtran(curAcct,"\","")
	curAcct = strtran(curAcct,"/","")
	curAcct = strtran(curAcct,"-","")
	curAcct = strtran(curAcct," ","")
	curAcct = padl(curAcct,12,"0")
	m.AcctNum = curAcct
	use
	sele space(6) as company_no, left(dtoc(ckdate),2)+substr(dtoc(ckdate),4,2)+right(dtoc(ckdate),4) as ckdate, ;
		padl(alltrim(cknumber),10,"0") as CheckNum, ;
		padl(strtran(alltrim(str(amount,12,2)),".",""),12,"0") as amount, ;
		amount as xAmount ;
		from xPath + "ckbook" ;
	where &rptDate ;
	and cknumber <> "DEP" and ! "WIRE" $(cknumber) ;
	and ! "ADJ" $(cknumber) and satype <> "V" ;
	into cursor tmpCks
	
	sele tmpCks
	scan	
		scatter memvar
		insert into tmpData from memvar
	endscan
	sele tmpData
	repla all for empty(company_no) company_no with data_loc.company_no
	sele tmpCks
	use
	sele ckBook
	use
return	
	
	