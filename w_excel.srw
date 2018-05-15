HA$PBExportHeader$w_excel.srw
forward
global type w_excel from window
end type
type ddlb_mes from dropdownlistbox within w_excel
end type
type st_10 from statictext within w_excel
end type
type em_dia from editmask within w_excel
end type
type cb_1 from commandbutton within w_excel
end type
type st_9 from statictext within w_excel
end type
type st_8 from statictext within w_excel
end type
type st_7 from statictext within w_excel
end type
type st_6 from statictext within w_excel
end type
type st_5 from statictext within w_excel
end type
type mle_archivo from multilineedit within w_excel
end type
type st_4 from statictext within w_excel
end type
type st_3 from statictext within w_excel
end type
type st_2 from statictext within w_excel
end type
type st_1 from statictext within w_excel
end type
type sle_12 from singlelineedit within w_excel
end type
type sle_11 from singlelineedit within w_excel
end type
type sle_10 from singlelineedit within w_excel
end type
type sle_9 from singlelineedit within w_excel
end type
type sle_8 from singlelineedit within w_excel
end type
type sle_7 from singlelineedit within w_excel
end type
type sle_6 from singlelineedit within w_excel
end type
type sle_5 from singlelineedit within w_excel
end type
type sle_4 from singlelineedit within w_excel
end type
type sle_3 from singlelineedit within w_excel
end type
type dw_1 from datawindow within w_excel
end type
type sle_2 from singlelineedit within w_excel
end type
type sle_1 from singlelineedit within w_excel
end type
type cb_procesar from commandbutton within w_excel
end type
end forward

global type w_excel from window
integer x = 6501
integer y = 200
integer width = 1787
integer height = 1772
boolean titlebar = true
string title = "Reporte Semanal Consolidado Ventas"
boolean controlmenu = true
boolean minbox = true
long backcolor = 67108864
string icon = "AppIcon!"
boolean center = true
ddlb_mes ddlb_mes
st_10 st_10
em_dia em_dia
cb_1 cb_1
st_9 st_9
st_8 st_8
st_7 st_7
st_6 st_6
st_5 st_5
mle_archivo mle_archivo
st_4 st_4
st_3 st_3
st_2 st_2
st_1 st_1
sle_12 sle_12
sle_11 sle_11
sle_10 sle_10
sle_9 sle_9
sle_8 sle_8
sle_7 sle_7
sle_6 sle_6
sle_5 sle_5
sle_4 sle_4
sle_3 sle_3
dw_1 dw_1
sle_2 sle_2
sle_1 sle_1
cb_procesar cb_procesar
end type
global w_excel w_excel

type prototypes
Function long SHGetFolderPath ( long hwndOwner, long nFolder, long hToken, long dwFlags, Ref string pszPath ) Library "shell32.dll" alias For "SHGetFolderPathW"  
end prototypes

forward prototypes
public function boolean wf_campos_completos ()
public function string wf_specialfolders (long folder)
end prototypes

public function boolean wf_campos_completos ();boolean ret = true

if LEN(sle_1.text)=0 then
	ret = false
end if

if LEN(sle_2.text)=0 then
	ret = false
end if

if LEN(sle_3.text)=0 then
	ret = false
end if

if LEN(sle_4.text)=0 then
	ret = false
end if

if LEN(sle_5.text)=0 then
	ret = false
end if

if LEN(sle_6.text)=0 then
	ret = false
end if

if LEN(sle_7.text)=0 then
	ret = false
end if

if LEN(sle_8.text)=0 then
	ret = false
end if

if LEN(sle_9.text)=0 then
	ret = false
end if

if LEN(sle_10.text)=0 then
	ret = false
end if

if LEN(sle_11.text)=0 then
	ret = false
end if

if LEN(sle_12.text)=0 then
	ret = false
end if

return ret
end function

public function string wf_specialfolders (long folder);Constant Long CSIDL_PERSONAL = 5 // current user My Documents  
Constant Long CSIDL_APPDATA = 26 // current user Application Data  
Constant Long CSIDL_LOCAL_APPDATA = 28 // local settings Application Data  
Constant Long CSIDL_COMMON_DOCUMENTS = 46 // all users My Documents  
Constant Long CSIDL_COMMON_APPDATA = 35 // all users Application Data  

string ls_path  
ulong lul_handle, lul_rc, lul_hToken  

ls_path = Space(256)  
lul_handle = Handle(This)  
SetNull(lul_hToken)  
lul_rc = SHGetFolderPath(lul_handle, CSIDL_PERSONAL, lul_hToken, 0, ls_path)  

RETURN ls_path // path  
end function

on w_excel.create
this.ddlb_mes=create ddlb_mes
this.st_10=create st_10
this.em_dia=create em_dia
this.cb_1=create cb_1
this.st_9=create st_9
this.st_8=create st_8
this.st_7=create st_7
this.st_6=create st_6
this.st_5=create st_5
this.mle_archivo=create mle_archivo
this.st_4=create st_4
this.st_3=create st_3
this.st_2=create st_2
this.st_1=create st_1
this.sle_12=create sle_12
this.sle_11=create sle_11
this.sle_10=create sle_10
this.sle_9=create sle_9
this.sle_8=create sle_8
this.sle_7=create sle_7
this.sle_6=create sle_6
this.sle_5=create sle_5
this.sle_4=create sle_4
this.sle_3=create sle_3
this.dw_1=create dw_1
this.sle_2=create sle_2
this.sle_1=create sle_1
this.cb_procesar=create cb_procesar
this.Control[]={this.ddlb_mes,&
this.st_10,&
this.em_dia,&
this.cb_1,&
this.st_9,&
this.st_8,&
this.st_7,&
this.st_6,&
this.st_5,&
this.mle_archivo,&
this.st_4,&
this.st_3,&
this.st_2,&
this.st_1,&
this.sle_12,&
this.sle_11,&
this.sle_10,&
this.sle_9,&
this.sle_8,&
this.sle_7,&
this.sle_6,&
this.sle_5,&
this.sle_4,&
this.sle_3,&
this.dw_1,&
this.sle_2,&
this.sle_1,&
this.cb_procesar}
end on

on w_excel.destroy
destroy(this.ddlb_mes)
destroy(this.st_10)
destroy(this.em_dia)
destroy(this.cb_1)
destroy(this.st_9)
destroy(this.st_8)
destroy(this.st_7)
destroy(this.st_6)
destroy(this.st_5)
destroy(this.mle_archivo)
destroy(this.st_4)
destroy(this.st_3)
destroy(this.st_2)
destroy(this.st_1)
destroy(this.sle_12)
destroy(this.sle_11)
destroy(this.sle_10)
destroy(this.sle_9)
destroy(this.sle_8)
destroy(this.sle_7)
destroy(this.sle_6)
destroy(this.sle_5)
destroy(this.sle_4)
destroy(this.sle_3)
destroy(this.dw_1)
destroy(this.sle_2)
destroy(this.sle_1)
destroy(this.cb_procesar)
end on

event open;string ls_dia
string ls_mes
integer li_dia

ls_dia =  ProfileString("a_excel.ini", "fecha", "dia", "01")
li_dia = integer(ls_dia) + 7
if li_dia < 10 then
	ls_dia = '0' + string(li_dia)
else
	ls_dia = string(li_dia)
end if

em_dia.text = ls_dia
ddlb_mes.selectitem(month(today()) )
dw_1.settransobject( sqlca)


end event

type ddlb_mes from dropdownlistbox within w_excel
integer x = 617
integer y = 656
integer width = 709
integer height = 476
integer taborder = 150
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
boolean sorted = false
boolean vscrollbar = true
string item[] = {"Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"}
borderstyle borderstyle = stylelowered!
end type

type st_10 from statictext within w_excel
integer x = 430
integer y = 664
integer width = 178
integer height = 92
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Mes:"
boolean focusrectangle = false
end type

type em_dia from editmask within w_excel
integer x = 192
integer y = 660
integer width = 169
integer height = 112
integer taborder = 150
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
string mask = "##"
end type

type cb_1 from commandbutton within w_excel
integer x = 1353
integer y = 644
integer width = 320
integer height = 132
integer taborder = 140
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Cierres"
end type

event clicked;datetime ldt_fecharep
Date ld_Date 
Time lt_Time = time("00:00:00")
string ls_cierre
string ls_mes, ls_dia 
Long ll_index

ls_dia = em_dia.Text

if integer(ls_dia) < 10 then
	ls_dia = '0' + ls_dia
end if

ll_index = ddlb_mes.FindItem(ddlb_mes.Text, 0)

if ll_index < 10 then
	ls_mes = '0' + string(ll_index)
else
	ls_mes = string(ll_index)
end if

////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//A$$HEX1$$f100$$ENDHEX$$o 1
ld_Date = Date('01/01/' + st_1.text ) 
ldt_fecharep = DateTime( ld_Date, lt_Time) 

select id_cierre_diario into :ls_cierre from cierre_diario where fecha_operacion = :ldt_fecharep;

sle_1.text = ls_cierre

ld_Date = Date('01/' + ls_mes + '/' + st_1.text ) 
ldt_fecharep = DateTime( ld_Date, lt_Time) 

select id_cierre_diario into :ls_cierre from cierre_diario where fecha_operacion = :ldt_fecharep;

sle_2.text = ls_cierre

ld_Date = Date(em_dia.Text + '/' + ls_mes + '/' + st_1.text ) 
ldt_fecharep = DateTime( ld_Date, lt_Time) 

select id_cierre_diario into :ls_cierre from cierre_diario where fecha_operacion = :ldt_fecharep;

sle_3.text = ls_cierre
////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//A$$HEX1$$f100$$ENDHEX$$o 2
ld_Date = Date('01/01/' + st_2.text ) 
ldt_fecharep = DateTime( ld_Date, lt_Time) 

select id_cierre_diario into :ls_cierre from cierre_diario where fecha_operacion = :ldt_fecharep;

sle_4.text = ls_cierre

ld_Date = Date('01/' + ls_mes + '/' + st_2.text ) 
ldt_fecharep = DateTime( ld_Date, lt_Time) 

select id_cierre_diario into :ls_cierre from cierre_diario where fecha_operacion = :ldt_fecharep;

sle_5.text = ls_cierre

ld_Date = Date(em_dia.Text + '/' + ls_mes + '/' + st_2.text ) 
ldt_fecharep = DateTime( ld_Date, lt_Time) 

select id_cierre_diario into :ls_cierre from cierre_diario where fecha_operacion = :ldt_fecharep;

sle_6.text = ls_cierre
////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//A$$HEX1$$f100$$ENDHEX$$o 3
ld_Date = Date('01/01/' + st_3.text ) 
ldt_fecharep = DateTime( ld_Date, lt_Time) 

select id_cierre_diario into :ls_cierre from cierre_diario where fecha_operacion = :ldt_fecharep;

sle_7.text = ls_cierre

ld_Date = Date('01/' + ls_mes + '/' + st_3.text ) 
ldt_fecharep = DateTime( ld_Date, lt_Time) 

select id_cierre_diario into :ls_cierre from cierre_diario where fecha_operacion = :ldt_fecharep;

sle_8.text = ls_cierre

ld_Date = Date(em_dia.Text + '/' + ls_mes + '/' + st_3.text ) 
ldt_fecharep = DateTime( ld_Date, lt_Time) 

select id_cierre_diario into :ls_cierre from cierre_diario where fecha_operacion = :ldt_fecharep;

sle_9.text = ls_cierre
////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//A$$HEX1$$f100$$ENDHEX$$o 4
ld_Date = Date('01/01/' + st_4.text ) 
ldt_fecharep = DateTime( ld_Date, lt_Time) 

select id_cierre_diario into :ls_cierre from cierre_diario where fecha_operacion = :ldt_fecharep;

sle_10.text = ls_cierre

ld_Date = Date('01/' + ls_mes + '/' + st_4.text ) 
ldt_fecharep = DateTime( ld_Date, lt_Time) 

select id_cierre_diario into :ls_cierre from cierre_diario where fecha_operacion = :ldt_fecharep;

sle_11.text = ls_cierre

ld_Date = Date(em_dia.Text + '/' + ls_mes + '/' + st_4.text ) 
ldt_fecharep = DateTime( ld_Date, lt_Time) 

select id_cierre_diario into :ls_cierre from cierre_diario where fecha_operacion = :ldt_fecharep;

sle_12.text = ls_cierre

setProfileString("a_excel.ini", "fecha", "dia", ls_dia)
setProfileString("a_excel.ini", "fecha", "mes", ls_mes)
end event

type st_9 from statictext within w_excel
integer x = 46
integer y = 680
integer width = 155
integer height = 92
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "D$$HEX1$$ed00$$ENDHEX$$a:"
boolean focusrectangle = false
end type

type st_8 from statictext within w_excel
integer x = 1289
integer y = 64
integer width = 411
integer height = 92
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "HASTA"
alignment alignment = center!
boolean focusrectangle = false
end type

type st_7 from statictext within w_excel
integer x = 169
integer y = 68
integer width = 201
integer height = 76
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Cierres"
alignment alignment = center!
boolean focusrectangle = false
end type

type st_6 from statictext within w_excel
integer x = 859
integer y = 64
integer width = 411
integer height = 92
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "INICIO"
alignment alignment = center!
boolean focusrectangle = false
end type

type st_5 from statictext within w_excel
integer x = 453
integer y = 64
integer width = 389
integer height = 92
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "ENERO 1$$HEX1$$b000$$ENDHEX$$"
alignment alignment = center!
boolean focusrectangle = false
end type

type mle_archivo from multilineedit within w_excel
integer x = 32
integer y = 1456
integer width = 1719
integer height = 188
integer taborder = 150
integer textsize = -10
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
boolean vscrollbar = true
boolean displayonly = true
borderstyle borderstyle = stylelowered!
end type

type st_4 from statictext within w_excel
integer x = 87
integer y = 508
integer width = 343
integer height = 92
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "2018"
alignment alignment = center!
boolean focusrectangle = false
end type

type st_3 from statictext within w_excel
integer x = 82
integer y = 392
integer width = 343
integer height = 92
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "2017"
alignment alignment = center!
boolean focusrectangle = false
end type

type st_2 from statictext within w_excel
integer x = 87
integer y = 280
integer width = 343
integer height = 92
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "2016"
alignment alignment = center!
boolean focusrectangle = false
end type

type st_1 from statictext within w_excel
integer x = 87
integer y = 164
integer width = 343
integer height = 92
integer textsize = -12
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "2015"
alignment alignment = center!
boolean focusrectangle = false
end type

type sle_12 from singlelineedit within w_excel
integer x = 1289
integer y = 508
integer width = 402
integer height = 92
integer taborder = 120
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_11 from singlelineedit within w_excel
integer x = 864
integer y = 508
integer width = 402
integer height = 92
integer taborder = 110
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_10 from singlelineedit within w_excel
integer x = 439
integer y = 508
integer width = 402
integer height = 92
integer taborder = 100
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_9 from singlelineedit within w_excel
integer x = 1289
integer y = 396
integer width = 402
integer height = 92
integer taborder = 90
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_8 from singlelineedit within w_excel
integer x = 864
integer y = 396
integer width = 402
integer height = 92
integer taborder = 80
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_7 from singlelineedit within w_excel
integer x = 439
integer y = 396
integer width = 402
integer height = 92
integer taborder = 70
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_6 from singlelineedit within w_excel
integer x = 1289
integer y = 284
integer width = 402
integer height = 92
integer taborder = 60
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_5 from singlelineedit within w_excel
integer x = 864
integer y = 280
integer width = 402
integer height = 92
integer taborder = 50
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_4 from singlelineedit within w_excel
integer x = 439
integer y = 280
integer width = 402
integer height = 92
integer taborder = 40
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_3 from singlelineedit within w_excel
integer x = 1289
integer y = 164
integer width = 402
integer height = 92
integer taborder = 30
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type dw_1 from datawindow within w_excel
integer x = 46
integer y = 992
integer width = 1687
integer height = 416
integer taborder = 170
string title = "none"
string dataobject = "dw_valores_consolidado"
boolean vscrollbar = true
boolean livescroll = true
borderstyle borderstyle = stylelowered!
end type

type sle_2 from singlelineedit within w_excel
integer x = 864
integer y = 164
integer width = 402
integer height = 92
integer taborder = 20
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_1 from singlelineedit within w_excel
integer x = 439
integer y = 164
integer width = 402
integer height = 92
integer taborder = 10
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type cb_procesar from commandbutton within w_excel
integer x = 46
integer y = 820
integer width = 1659
integer height = 112
integer taborder = 130
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Procesar"
end type

event clicked;int li_rtn, li_row, li_col = 1
string ls_range, ls_path, ls_title, ls_mes
oleobject lole_excel, lole_workbook, lole_worksheet, lole_range
string docpath, docname[]
integer i, li_cnt, li_filenum
datetime ldt_fechafin
decimal vtactdoasoc,vtactdonoasoc,vtacredasoc, tonctdoasoc,tonctdonoasoc,toncredasoc

vtactdoasoc = 0
vtactdonoasoc = 0
vtacredasoc = 0
tonctdoasoc = 0
tonctdonoasoc = 0
toncredasoc = 0

cb_1.triggerevent( Clicked!)

if not wf_campos_completos() then
	MessageBox("Par$$HEX1$$e100$$ENDHEX$$metros", "Faltan codigos de fecha... verifique.", StopSign!, OK!)
	return
end if

ls_path = wf_specialfolders(1)

li_rtn = GetFileOpenName("Select File", &
   docpath, docname[], "Excel", &
   + "Excel 1997-2003 (*.xls),*.xls," &
   + "Excel 2010-2013 (*.xlsx),*.xlsx," &
   + "All Files (*.*), *.*", &
   ls_path, 18)

mle_archivo.text = ""
IF li_rtn < 1 THEN return
li_cnt = Upperbound(docname)

// if only one file is picked, docpath contains the 
// path and file name
if li_cnt = 1 then
	mle_archivo.text = string(docpath)
else
	// if multiple files are picked, docpath contains the 
	// path only - concatenate docpath and docname
   	for i=1 to li_cnt
      	mle_archivo.text += string(docpath) + "\" +(string(docname[i]))+"~r~n"
   	next
end if

//return
lole_excel = create oleobject
li_rtn = lole_excel.ConnectToNewObject("excel.application")
if li_rtn <> 0 then
	choose case li_rtn
		case -1  
			 MessageBox( "Error", 'Invalid Call: the argument is the Object property of a control')
		case -2  
			 MessageBox( "Error", 'Class name not found')
		case -3  
			 MessageBox( "Error", 'Object could not be created')
		case -4  
			 MessageBox( "Error", 'Could not connect to object')
		case -9   
			MessageBox( "Error", 'Other error')
		case -15  
			 MessageBox( "Error", 'COM+ is not loaded on this computer')
		case -16  
			 MessageBox( "Error", 'Invalid Call: this function not applicable')
		case else
	      	MessageBox( "Error", 'Error running MS Excel api.')
	end choose
    destroy lole_Excel
else


  lole_excel.WorkBooks.Open(docpath) 

  lole_workbook = lole_excel.application.workbooks(1)
  lole_worksheet = lole_workbook.worksheets(1)
  lole_excel.visible=true

//VALORES MONETARIOS
//valores acumulados anual
ls_title = sle_12.text
SELECT FECHA_OPERACION INTO :ldt_fechafin FROM CIERRE_DIARIO WHERE ID_CIERRE_DIARIO = :ls_title;
ls_range = string(month(date(ldt_fechafin)))
ls_range = fill('0', 2 - len(ls_range)) + ls_range
SELECT NOMBRE_MES INTO :ls_mes FROM MES WHERE MES = :ls_range;
ls_title = 'Ventas comparadas a$$HEX1$$f100$$ENDHEX$$o '+  st_1.text + ', '+  st_2.text + ', '+  st_3.text + ' y '+  st_4.text + ' del 1 de Enero  al '+  string(day(date(ldt_fechafin))) + ' de ' + mid(ls_mes,1,1) + lower(mid(ls_mes,2))  + ' (valores sin iva)'
lole_worksheet.cells(1,2).value = ls_title //it is cells(line, column)

for li_col = 1 to 4
	if li_col = 1 then
		li_row = 10
		dw_1.retrieve( sle_1.text, sle_3.text )
	end if
	if li_col = 2 then
		li_row = 7
		dw_1.retrieve( sle_4.text, sle_6.text )
	end if
	if li_col = 3 then
		li_row = 4
		dw_1.retrieve( sle_7.text, sle_9.text )
	end if
	if li_col = 4 then
		li_row = 3
		dw_1.retrieve( sle_10.text, sle_12.text )
		vtactdoasoc = 0 //153.94
		vtactdonoasoc = 0 //238.64
		vtacredasoc = 0 //75.00
	end if
  // Set the cell value
  //contado ss
  lole_worksheet.cells(8,li_row).value = dw_1.getitemdecimal( 1, 2) + vtactdoasoc //it is cells(line, column)
  lole_worksheet.cells(9,li_row).value = dw_1.getitemdecimal( 2, 2) + vtactdonoasoc //it is cells(line, column)
  lole_worksheet.cells(10,li_row).value = dw_1.getitemdecimal( 3, 2) //it is cells(line, column)
  lole_worksheet.cells(11,li_row).value = dw_1.getitemdecimal( 4, 2) //it is cells(line, column)
  
//  credito ss
  lole_worksheet.cells(26,li_row).value = dw_1.getitemdecimal( 5, 2) + vtacredasoc//it is cells(line, column)
  lole_worksheet.cells(27,li_row).value = dw_1.getitemdecimal( 6, 2) //it is cells(line, column)
  lole_worksheet.cells(28,li_row).value = dw_1.getitemdecimal( 7, 2) //it is cells(line, column)
  lole_worksheet.cells(29,li_row).value = dw_1.getitemdecimal( 8, 2) //it is cells(line, column)
  
  //contado sm
  lole_worksheet.cells(13,li_row).value = dw_1.getitemdecimal( 9, 2) //it is cells(line, column)
  lole_worksheet.cells(14,li_row).value = dw_1.getitemdecimal( 10, 2) //it is cells(line, column)
  lole_worksheet.cells(15,li_row).value = dw_1.getitemdecimal( 11, 2) //it is cells(line, column)
  lole_worksheet.cells(16,li_row).value = dw_1.getitemdecimal( 12, 2) //it is cells(line, column)
  
//  credito sm
  lole_worksheet.cells(31,li_row).value = dw_1.getitemdecimal( 13, 2) //it is cells(line, column)
  lole_worksheet.cells(32,li_row).value = dw_1.getitemdecimal( 14, 2) //it is cells(line, column)
  lole_worksheet.cells(33,li_row).value = dw_1.getitemdecimal( 15, 2) //it is cells(line, column)
  lole_worksheet.cells(34,li_row).value = dw_1.getitemdecimal( 16, 2) //it is cells(line, column)

  //preprensa contado
  lole_worksheet.cells(45,li_row).value = dw_1.getitemdecimal( 17, 2) //it is cells(line, column)
  lole_worksheet.cells(46,li_row).value = dw_1.getitemdecimal( 18, 2) //it is cells(line, column)
  lole_worksheet.cells(47,li_row).value = dw_1.getitemdecimal( 19, 2) //it is cells(line, column)
  lole_worksheet.cells(48,li_row).value = dw_1.getitemdecimal( 20, 2) //it is cells(line, column)
  
//  preprensa credito
  lole_worksheet.cells(57,li_row).value = dw_1.getitemdecimal( 21, 2) //it is cells(line, column)
  lole_worksheet.cells(58,li_row).value = dw_1.getitemdecimal( 22, 2) //it is cells(line, column)
  lole_worksheet.cells(59,li_row).value = dw_1.getitemdecimal( 23, 2) //it is cells(line, column)
  lole_worksheet.cells(60,li_row).value = dw_1.getitemdecimal( 24, 2) //it is cells(line, column)

next


//valores acumulados mensuales
  lole_worksheet = lole_workbook.worksheets(2)
  ls_title = 'Ventas comparadas a$$HEX1$$f100$$ENDHEX$$o '+  st_1.text + ', '+  st_2.text + ', '+  st_3.text + ' y '+  st_4.text + ' del 1 al '+  string(day(date(ldt_fechafin))) + ' de ' + mid(ls_mes,1,1) + lower(mid(ls_mes,2))  + ' (valores sin iva)'
  lole_worksheet.cells(1,2).value = ls_title //it is cells(line, column)
  ls_title = 'Del 1 al '+  string(day(date(ldt_fechafin))) + ' de ' + mid(ls_mes,1,1) + lower(mid(ls_mes,2)) + ' San Salvador'
  lole_worksheet.cells(5,2).value = ls_title //it is cells(line, column)
  ls_title = 'Del 1 al '+  string(day(date(ldt_fechafin))) + ' de ' + mid(ls_mes,1,1) + lower(mid(ls_mes,2)) + ' San Miguel'
  lole_worksheet.cells(6,2).value = ls_title //it is cells(line, column)
  ls_title = 'Del 1 al '+  string(day(date(ldt_fechafin))) + ' de ' + mid(ls_mes,1,1) + lower(mid(ls_mes,2)) 
  lole_worksheet.cells(7,2).value = ls_title //it is cells(line, column)
  ls_title = 'Del 1 de Enero al '+  string(day(date(ldt_fechafin))) + ' de ' + mid(ls_mes,1,1) + lower(mid(ls_mes,2)) 
  lole_worksheet.cells(8,2).value = ls_title //it is cells(line, column)
  ls_title = 'Valores al '+  string(day(date(ldt_fechafin))) + ' de ' + mid(ls_mes,1,1) + lower(mid(ls_mes,2))
  lole_worksheet.cells(11,3).value = ls_title //it is cells(line, column)

  for li_col = 1 to 4
	if li_col = 1 then
		li_row = 10
		dw_1.retrieve( sle_2.text, sle_3.text )
	end if
	if li_col = 2 then
		li_row = 7
		dw_1.retrieve( sle_5.text, sle_6.text )
	end if
	if li_col = 3 then
		li_row = 4
		dw_1.retrieve( sle_8.text, sle_9.text )
	end if
	if li_col = 4 then
		li_row = 3
		dw_1.retrieve( sle_11.text, sle_12.text )
	end if
  // Set the cell value
  //contado ss
  lole_worksheet.cells(14,li_row).value = dw_1.getitemdecimal( 1, 2) //it is cells(line, column)
  lole_worksheet.cells(15,li_row).value = dw_1.getitemdecimal( 2, 2) //it is cells(line, column)
  lole_worksheet.cells(16,li_row).value = dw_1.getitemdecimal( 3, 2) //it is cells(line, column)
  lole_worksheet.cells(17,li_row).value = dw_1.getitemdecimal( 4, 2) //it is cells(line, column)
  
  //credito ss
  lole_worksheet.cells(32,li_row).value = dw_1.getitemdecimal( 5, 2) //it is cells(line, column)
  lole_worksheet.cells(33,li_row).value = dw_1.getitemdecimal( 6, 2) //it is cells(line, column)
  lole_worksheet.cells(34,li_row).value = dw_1.getitemdecimal( 7, 2) //it is cells(line, column)
  lole_worksheet.cells(35,li_row).value = dw_1.getitemdecimal( 8, 2) //it is cells(line, column)
  
  //contado sm
  lole_worksheet.cells(19,li_row).value = dw_1.getitemdecimal( 9, 2) //it is cells(line, column)
  lole_worksheet.cells(20,li_row).value = dw_1.getitemdecimal( 10, 2) //it is cells(line, column)
  lole_worksheet.cells(21,li_row).value = dw_1.getitemdecimal( 11, 2) //it is cells(line, column)
  lole_worksheet.cells(22,li_row).value = dw_1.getitemdecimal( 12, 2) //it is cells(line, column)
  
  //credito sm
  lole_worksheet.cells(37,li_row).value = dw_1.getitemdecimal( 13, 2) //it is cells(line, column)
  lole_worksheet.cells(38,li_row).value = dw_1.getitemdecimal( 14, 2) //it is cells(line, column)
  lole_worksheet.cells(39,li_row).value = dw_1.getitemdecimal( 15, 2) //it is cells(line, column)
  lole_worksheet.cells(40,li_row).value = dw_1.getitemdecimal( 16, 2) //it is cells(line, column)

  //preprensa contado
  lole_worksheet.cells(54,li_row).value = dw_1.getitemdecimal( 17, 2) //it is cells(line, column)
  lole_worksheet.cells(55,li_row).value = dw_1.getitemdecimal( 18, 2) //it is cells(line, column)
  lole_worksheet.cells(56,li_row).value = dw_1.getitemdecimal( 19, 2) //it is cells(line, column)
  lole_worksheet.cells(57,li_row).value = dw_1.getitemdecimal( 20, 2) //it is cells(line, column)
  
  //preprensa credito
  lole_worksheet.cells(66,li_row).value = dw_1.getitemdecimal( 21, 2) //it is cells(line, column)
  lole_worksheet.cells(67,li_row).value = dw_1.getitemdecimal( 22, 2) //it is cells(line, column)
  lole_worksheet.cells(68,li_row).value = dw_1.getitemdecimal( 23, 2) //it is cells(line, column)
  lole_worksheet.cells(69,li_row).value = dw_1.getitemdecimal( 24, 2) //it is cells(line, column)

next

//TONELADAS VENDIDAS
dw_1.reset()
dw_1.dataobject='dw_toneladas_consolidado'
dw_1.settransobject( sqlca)

 lole_worksheet = lole_workbook.worksheets(3)
 ls_title = 'Ventas toneladas comparadas a$$HEX1$$f100$$ENDHEX$$o '+  st_1.text + ', '+  st_2.text + ', '+  st_3.text + ' y '+  st_4.text + ' del 1 de Enero al '+  string(day(date(ldt_fechafin))) + ' de ' + mid(ls_mes,1,1) + lower(mid(ls_mes,2))
 lole_worksheet.cells(1,2).value = ls_title //it is cells(line, column)
 ls_title = 'Toneladas al '+  string(day(date(ldt_fechafin))) + ' de ' + mid(ls_mes,1,1) + lower(mid(ls_mes,2))
 lole_worksheet.cells(5,3).value = ls_title //it is cells(line, column)

 //valores acumulados anual
for li_col = 1 to 4
	if li_col = 1 then
		li_row = 10
		dw_1.retrieve( sle_1.text, sle_3.text )
	end if
	if li_col = 2 then
		li_row = 7
		dw_1.retrieve( sle_4.text, sle_6.text )
	end if
	if li_col = 3 then
		li_row = 4
		dw_1.retrieve( sle_7.text, sle_9.text )
	end if
	if li_col = 4 then
		li_row = 3
		dw_1.retrieve( sle_10.text, sle_12.text )
		tonctdoasoc = 0 //0.1176
		tonctdonoasoc = 0 //0.1790
		toncredasoc = 0 //0.0339
	end if
  // Set the cell value
  //contado ss
  lole_worksheet.cells(8,li_row).value = dw_1.getitemdecimal( 1, 2) + tonctdoasoc //it is cells(line, column)
  lole_worksheet.cells(9,li_row).value = dw_1.getitemdecimal( 2, 2)  + tonctdonoasoc//it is cells(line, column)
  lole_worksheet.cells(10,li_row).value = dw_1.getitemdecimal( 3, 2) //it is cells(line, column)
  lole_worksheet.cells(11,li_row).value = dw_1.getitemdecimal( 4, 2) //it is cells(line, column)
  
//  credito ss
  lole_worksheet.cells(22,li_row).value = dw_1.getitemdecimal( 5, 2) + toncredasoc //it is cells(line, column)
  lole_worksheet.cells(23,li_row).value = dw_1.getitemdecimal( 6, 2) //it is cells(line, column)
  lole_worksheet.cells(24,li_row).value = dw_1.getitemdecimal( 7, 2) //it is cells(line, column)
  lole_worksheet.cells(25,li_row).value = dw_1.getitemdecimal( 8, 2) //it is cells(line, column)
  
  //contado sm
  lole_worksheet.cells(13,li_row).value = dw_1.getitemdecimal( 9, 2) //it is cells(line, column)
  lole_worksheet.cells(14,li_row).value = dw_1.getitemdecimal( 10, 2) //it is cells(line, column)
  lole_worksheet.cells(15,li_row).value = dw_1.getitemdecimal( 11, 2) //it is cells(line, column)
  lole_worksheet.cells(16,li_row).value = dw_1.getitemdecimal( 12, 2) //it is cells(line, column)
  
//  credito sm
  lole_worksheet.cells(27,li_row).value = dw_1.getitemdecimal( 13, 2) //it is cells(line, column)
  lole_worksheet.cells(28,li_row).value = dw_1.getitemdecimal( 14, 2) //it is cells(line, column)
  lole_worksheet.cells(29,li_row).value = dw_1.getitemdecimal( 15, 2) //it is cells(line, column)
  lole_worksheet.cells(30,li_row).value = dw_1.getitemdecimal( 16, 2) //it is cells(line, column)

  //preprensa contado
  lole_worksheet.cells(37,li_row).value = dw_1.getitemdecimal( 17, 2) //it is cells(line, column)
  lole_worksheet.cells(38,li_row).value = dw_1.getitemdecimal( 18, 2) //it is cells(line, column)
  lole_worksheet.cells(39,li_row).value = dw_1.getitemdecimal( 19, 2) //it is cells(line, column)
  lole_worksheet.cells(40,li_row).value = dw_1.getitemdecimal( 20, 2) //it is cells(line, column)
  
//  preprensa credito
  lole_worksheet.cells(45,li_row).value = dw_1.getitemdecimal( 21, 2) //it is cells(line, column)
  lole_worksheet.cells(46,li_row).value = dw_1.getitemdecimal( 22, 2) //it is cells(line, column)
  lole_worksheet.cells(47,li_row).value = dw_1.getitemdecimal( 23, 2) //it is cells(line, column)
  lole_worksheet.cells(48,li_row).value = dw_1.getitemdecimal( 24, 2) //it is cells(line, column)

next

//valores acumulados mensuales
  lole_worksheet = lole_workbook.worksheets(4)
 ls_title = 'Ventas toneladas comparadas a$$HEX1$$f100$$ENDHEX$$o '+  st_1.text + ', '+  st_2.text + ', '+  st_3.text + ' y '+  st_4.text + ' del 1 al '+  string(day(date(ldt_fechafin))) + ' de ' + mid(ls_mes,1,1) + lower(mid(ls_mes,2))
 lole_worksheet.cells(1,2).value = ls_title //it is cells(line, column)

  for li_col = 1 to 4
	if li_col = 1 then
		li_row = 10
		dw_1.retrieve( sle_2.text, sle_3.text )
	end if
	if li_col = 2 then
		li_row = 7
		dw_1.retrieve( sle_5.text, sle_6.text )
	end if
	if li_col = 3 then
		li_row = 4
		dw_1.retrieve( sle_8.text, sle_9.text )
	end if
	if li_col = 4 then
		li_row = 3
		dw_1.retrieve( sle_11.text, sle_12.text )
	end if
  // Set the cell value
  //contado ss
  lole_worksheet.cells(14,li_row).value = dw_1.getitemdecimal( 1, 2) //it is cells(line, column)
  lole_worksheet.cells(15,li_row).value = dw_1.getitemdecimal( 2, 2) //it is cells(line, column)
  lole_worksheet.cells(16,li_row).value = dw_1.getitemdecimal( 3, 2) //it is cells(line, column)
  lole_worksheet.cells(17,li_row).value = dw_1.getitemdecimal( 4, 2) //it is cells(line, column)
  
  //credito ss
  lole_worksheet.cells(28,li_row).value = dw_1.getitemdecimal( 5, 2) //it is cells(line, column)
  lole_worksheet.cells(29,li_row).value = dw_1.getitemdecimal( 6, 2) //it is cells(line, column)
  lole_worksheet.cells(30,li_row).value = dw_1.getitemdecimal( 7, 2) //it is cells(line, column)
  lole_worksheet.cells(31,li_row).value = dw_1.getitemdecimal( 8, 2) //it is cells(line, column)
  
  //contado sm
  lole_worksheet.cells(19,li_row).value = dw_1.getitemdecimal( 9, 2) //it is cells(line, column)
  lole_worksheet.cells(20,li_row).value = dw_1.getitemdecimal( 10, 2) //it is cells(line, column)
  lole_worksheet.cells(21,li_row).value = dw_1.getitemdecimal( 11, 2) //it is cells(line, column)
  lole_worksheet.cells(22,li_row).value = dw_1.getitemdecimal( 12, 2) //it is cells(line, column)
  
  //credito sm
  lole_worksheet.cells(33,li_row).value = dw_1.getitemdecimal( 13, 2) //it is cells(line, column)
  lole_worksheet.cells(34,li_row).value = dw_1.getitemdecimal( 14, 2) //it is cells(line, column)
  lole_worksheet.cells(35,li_row).value = dw_1.getitemdecimal( 15, 2) //it is cells(line, column)
  lole_worksheet.cells(36,li_row).value = dw_1.getitemdecimal( 16, 2) //it is cells(line, column)

  //preprensa contado
  lole_worksheet.cells(46,li_row).value = dw_1.getitemdecimal( 17, 2) //it is cells(line, column)
  lole_worksheet.cells(47,li_row).value = dw_1.getitemdecimal( 18, 2) //it is cells(line, column)
  lole_worksheet.cells(48,li_row).value = dw_1.getitemdecimal( 19, 2) //it is cells(line, column)
  lole_worksheet.cells(49,li_row).value = dw_1.getitemdecimal( 20, 2) //it is cells(line, column)
  
  //preprensa credito
  lole_worksheet.cells(54,li_row).value = dw_1.getitemdecimal( 21, 2) //it is cells(line, column)
  lole_worksheet.cells(55,li_row).value = dw_1.getitemdecimal( 22, 2) //it is cells(line, column)
  lole_worksheet.cells(56,li_row).value = dw_1.getitemdecimal( 23, 2) //it is cells(line, column)
  lole_worksheet.cells(57,li_row).value = dw_1.getitemdecimal( 24, 2) //it is cells(line, column)

next

//COSTO DE VENTAS DE BODEGAS
dw_1.reset()
dw_1.dataobject='dw_costo_vta_bodega_consolidado'
dw_1.settransobject( sqlca)
//valores acumulados mensuales
lole_worksheet = lole_workbook.worksheets(2)

for li_col = 1 to 4
	if li_col = 1 then
		li_row = 21
		dw_1.retrieve( sle_2.text, sle_3.text )
	end if
	if li_col = 2 then
		li_row = 17
		dw_1.retrieve( sle_5.text, sle_6.text )
	end if
	if li_col = 3 then
		li_row = 13
		dw_1.retrieve( sle_8.text, sle_9.text )
	end if
	if li_col = 4 then
		li_row = 9
		dw_1.retrieve( sle_11.text, sle_12.text )
	end if

  lole_worksheet.cells(li_row,14).value = dw_1.getitemdecimal( 1, 2) //it is cells(line, column)
  lole_worksheet.cells(li_row,15).value = dw_1.getitemdecimal( 2, 2) //it is cells(line, column)
  lole_worksheet.cells(li_row,16).value = dw_1.getitemdecimal( 3, 2) //it is cells(line, column)
  lole_worksheet.cells(li_row,17).value = dw_1.getitemdecimal( 4, 2) //it is cells(line, column)
next

// Save
//  lole_workbook.save()
  // Quit
//  lole_excel.application.quit()
  lole_excel.DisconnectObject()
  destroy lole_Excel
end if
end event

