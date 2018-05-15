HA$PBExportHeader$a_excel.sra
$PBExportComments$Generated Application Object
forward
global type a_excel from application
end type
global transaction sqlca
global dynamicdescriptionarea sqlda
global dynamicstagingarea sqlsa
global error error
global message message
end forward

global type a_excel from application
string appname = "a_excel"
end type
global a_excel a_excel

on a_excel.create
appname="a_excel"
message=create message
sqlca=create transaction
sqlda=create dynamicdescriptionarea
sqlsa=create dynamicstagingarea
error=create error
end on

on a_excel.destroy
destroy(sqlca)
destroy(sqlda)
destroy(sqlsa)
destroy(error)
destroy(message)
end on

event open;// Profile alphaproductos
SQLCA.DBMS = "OLE DB"
SQLCA.LogPass = "admin"
SQLCA.LogId = "sa"
SQLCA.AutoCommit = False
//SQLCA.DBParm = "PROVIDER='SQLNCLI11',DATASOURCE='10.7.8.111\SQLSERVER2016',PROVIDERSTRING='database=alpha_productos'"
//SQLCA.DBParm = "PROVIDER='SQLNCLI11',DATASOURCE='PC-141',PROVIDERSTRING='database=alpha_productos'"
//SQLCA.DBParm = "PROVIDER='SQLNCLI11',DATASOURCE='10.7.10.122',PROVIDERSTRING='database=alpha_restored'"
SQLCA.DBParm = "PROVIDER='SQLOLEDB',DATASOURCE='10.7.10.122',PROVIDERSTRING='database=alpha_restored'"


connect;
IF SQLCA.SQLCode = -1 THEN 

        MessageBox("SQL error", SQLCA.SQLErrText)

END IF
open(w_excel)

end event

