Attribute Module_Name = "mfrmCalendar880"
//jsa880配套日历面板 使用时配合jsa880框架一起使用
//导出frmCalenDar880和mfrmCalenDar880在其他项目导入 即可使用
//顶部<< 内层<为月往前翻 外层<为年往前翻 右侧>>为往后翻
//今 点击后切换到当天对应的月份 选 点击后 切换到目标单元格或者控件里对应的日期
//Calendar880 作者 郑广学 2024.6.6 V1.0.1 vbayyds.com
//表格事件一行代码指定区域弹出日历面板
//function Workbook_SheetSelectionChange(Sh, Target)
//{
//	Calendar880.showInRange(Target,Sheets("日历").Range("C2:C24,G2:G24"),"yyyy-MM-dd")
//}
//窗体在文本框或者组合框的鼠标弹起事件里 显示日历面板 在窗体的click事件里Calendar880.close()关闭日历面板
//function UserForm2_TextEdit1_MouseUp(btn, shift, x, y)
//{
//	Calendar880.showInFrm(UserForm2.TextEdit1);
//}

var Calendar880={
	getMonthData(){},
	showInRange(Target,targetRange,dateFormat='',showmodle=0){},
	showInFrm(Target,dateFormat='yyyy-MM-dd',showmodle=0){},
	close(){}
}
eval(function(p,a,c,k,e,r){e=function(c){return(c<a?'':e(parseInt(c/a)))+((c=c%a)>35?String.fromCharCode(c+29):c.toString(36))};if(!''.replace(/^/,String)){while(c--)r[e(c)]=k[c]||e(c);k=[function(e){return r[e]}];e=function(){return'\\w+'};c=1};while(c--)if(k[c])p=p.replace(new RegExp('\\b'+e(c)+'\\b','g'),k[c]);return p}('j A=(()=>{B t{C(t){2.4=f(t),2.g=3.g(2.4),2.h=3.h(2.4),2.b=5 D(2.g,2.h-1,1),2.k=3.E(2.b),2.l=$.m(2.b,1-2.k),2.F=3.G(2.b)}H(t=!1){j a=5 n(6).o().p(()=>5 n(7).o(0)),e=0;q(u t=0;t<6;t++)q(u r=0;r<7;r++)a[t][r]=$.m(2.l,e++);8 t?a.p(t=>3.I(t)):a}J(){8 5 t(3.c(2.4,1))}K(){8 5 t(3.c(2.4,-1))}L(){8 5 t(3.c(2.4,v))}M(){8 5 t(3.c(2.4,-v))}d i(t,a=0,e="w-x-y",r){9.z(0),N(r)&&!O(r)&&加载日历(r),9.P=e,9.Q=t,1==a&&9.z(1)}d R(a,e,r="",s=0){8 $.S(a(1),e)?(t.i(a(1),s,T(r)?U(a(1)).V:r,f(a(1).W)),!0):(9.X(),!1)}d Y(a,e="w-x-y",r=0){t.i(a,r,e,f(a.Z))}d 10(){9.11()}}8 t})();',62,64,'||this|DateUtils|date|new|||return|frmCalendar880||day1|addMonths|static||asDate|year|month|ShowFrmCalendar|var|day1Weekday|firstDay|addDays|Array|fill|map|for||||let|12|yyyy|MM|dd|Show|Calendar880|class|constructor|Date|weekday|days|daysOfMonth|getMonthData|day|getNextMonth|getPreMonth|getNextYear|getPreYear|isDate|isNaN|dateformat|target|showInRange|hitRange|isEmpty|asRange|NumberFormatLocal|Value2|Hide|showInFrm|Text|close|Close'.split('|'),0,{}));

var calendarfrm=new Calendar880("2024-6-5");
var frm_Calendar880=frmCalendar880;
function frmCalendar880_Initialize()
{	
	frm_Calendar880=frmCalendar880;
	frmCalendar880.Frame1.Caption="";
	frmCalendar880.z今天=frmCalendar880.Label13;
	frmCalendar880.z选中=frmCalendar880.Label89;
	var tarr=[1,3,4].map(x=>frmCalendar880["TextEdit"+x]);
	//debugger
	frmCalendar880.dateformat="yyyy-MM-dd";//日期写入格式
	frmCalendar880.日期控件数组=$$.getMatrix(42,7).map2d(x=>frmCalendar880["Label"+(x+15)]);
	加载日历(new Date());
}

const 加载日历=(d)=>{
	d=defautvalue(d,new Date());
	calendarfrm=new Calendar880(d);
	calendarfrm.current=calendarfrm;
	显示日历();
}
const 显示日历=()=>{var r=frmCalendar880.日期控件数组,a=calendarfrm.current.getMonthData();frmCalendar880.Frame1.Visible=!1,frmCalendar880.TextEdit4.Text=calendarfrm.current.year,frmCalendar880.TextEdit3.Text=calendarfrm.current.month,a.forEach((a,e)=>{a.forEach((a,t)=>{r[e][t].ForeColor=JSA_Colors.Black,r[e][t].BackColor=parseInt("#f8f8f8".substring(1),16),r[e][t].date=a,r[e][t].Caption=a.getDate(),calendarfrm.current.month!=$.month(a)?r[e][t].ForeColor=parseInt("#bababa".substring(1),16):(cdate(a)==$.justDate(calendarfrm.date)&&(r[e][t].BackColor=parseInt("#ff557f".substring(1),16),r[e][t].ForeColor=JSA_Colors.White),cdate(a)==$.justDate(new Date)&&(r[e][t].ForeColor=parseInt("#0000ff".substring(1),16)))})}),frmCalendar880.Frame1.Visible=!0,frmCalendar880.Height=200};

const 前后翻月=(m=1)=>{
	if(isDate(m)){
		calendarfrm.current= new Calendar880(m);
	}else{
		calendarfrm.current= new Calendar880(DateUtils.addMonths(calendarfrm.current.date,m));
	}
	显示日历();
}
const frmCalendar880_输入=(c)=>{
	if(isString(c)){c=frmCalendar880[c]}
	var rs=JSA.text(cdate(c.date),frmCalendar880.dateformat);
	if(isRange(frmCalendar880.target)){
		frmCalendar880.target.Value2 = rs;
	}else{
		frmCalendar880.target.Text = rs;
	}
	frmCalendar880.Close();
}

/**
 * 上个月
 */
function frmCalendar880_Label57_Click()
{
	前后翻月(-1);
}
/**
 * 下个月
 */
function frmCalendar880_Label83_Click()
{
	前后翻月(1);
}
//上一年
function frmCalendar880_Label79_Click()
{
	前后翻月(-12);
}
//下一年
function frmCalendar880_Label84_Click()
{
	前后翻月(12);
}

//今天
function frmCalendar880_Label13_Click()
{
	前后翻月(new Date());
}


function frmCalendar880_Label89_Click()
{
	前后翻月(calendarfrm.date);//回到初始选择界面
}
function frmCalendar880_TextEdit4_KeyUp(keycode, shift)
{
	if(keycode==13){
		var d=calendarfrm.current.date;
		前后翻月(new Date(`${frmCalendar880.TextEdit4.Text}-${frmCalendar880.TextEdit3.Text}-${1}`));
	}
}
function frmCalendar880_TextEdit3_KeyUp(keycode, shift)
{
	if(keycode==13){
		var d=calendarfrm.current.date;
		前后翻月(new Date(`${frmCalendar880.TextEdit4.Text}-${frmCalendar880.TextEdit3.Text}-${1}`));
	}
}
//function frmCalendar880_Label31_Click(){frmCalendar880_输入('Label31')} //点击输入
function frmCalendar880_Label15_Click(){frmCalendar880_输入('Label15')}
function frmCalendar880_Label16_Click(){frmCalendar880_输入('Label16')}
function frmCalendar880_Label17_Click(){frmCalendar880_输入('Label17')}
function frmCalendar880_Label18_Click(){frmCalendar880_输入('Label18')}
function frmCalendar880_Label19_Click(){frmCalendar880_输入('Label19')}
function frmCalendar880_Label20_Click(){frmCalendar880_输入('Label20')}
function frmCalendar880_Label21_Click(){frmCalendar880_输入('Label21')}
function frmCalendar880_Label22_Click(){frmCalendar880_输入('Label22')}
function frmCalendar880_Label23_Click(){frmCalendar880_输入('Label23')}
function frmCalendar880_Label24_Click(){frmCalendar880_输入('Label24')}
function frmCalendar880_Label25_Click(){frmCalendar880_输入('Label25')}
function frmCalendar880_Label26_Click(){frmCalendar880_输入('Label26')}
function frmCalendar880_Label27_Click(){frmCalendar880_输入('Label27')}
function frmCalendar880_Label28_Click(){frmCalendar880_输入('Label28')}
function frmCalendar880_Label29_Click(){frmCalendar880_输入('Label29')}
function frmCalendar880_Label30_Click(){frmCalendar880_输入('Label30')}
function frmCalendar880_Label31_Click(){frmCalendar880_输入('Label31')}
function frmCalendar880_Label32_Click(){frmCalendar880_输入('Label32')}
function frmCalendar880_Label33_Click(){frmCalendar880_输入('Label33')}
function frmCalendar880_Label34_Click(){frmCalendar880_输入('Label34')}
function frmCalendar880_Label35_Click(){frmCalendar880_输入('Label35')}
function frmCalendar880_Label36_Click(){frmCalendar880_输入('Label36')}
function frmCalendar880_Label37_Click(){frmCalendar880_输入('Label37')}
function frmCalendar880_Label38_Click(){frmCalendar880_输入('Label38')}
function frmCalendar880_Label39_Click(){frmCalendar880_输入('Label39')}
function frmCalendar880_Label40_Click(){frmCalendar880_输入('Label40')}
function frmCalendar880_Label41_Click(){frmCalendar880_输入('Label41')}
function frmCalendar880_Label42_Click(){frmCalendar880_输入('Label42')}
function frmCalendar880_Label43_Click(){frmCalendar880_输入('Label43')}
function frmCalendar880_Label44_Click(){frmCalendar880_输入('Label44')}
function frmCalendar880_Label45_Click(){frmCalendar880_输入('Label45')}
function frmCalendar880_Label46_Click(){frmCalendar880_输入('Label46')}
function frmCalendar880_Label47_Click(){frmCalendar880_输入('Label47')}
function frmCalendar880_Label48_Click(){frmCalendar880_输入('Label48')}
function frmCalendar880_Label49_Click(){frmCalendar880_输入('Label49')}
function frmCalendar880_Label50_Click(){frmCalendar880_输入('Label50')}
function frmCalendar880_Label51_Click(){frmCalendar880_输入('Label51')}
function frmCalendar880_Label52_Click(){frmCalendar880_输入('Label52')}
function frmCalendar880_Label53_Click(){frmCalendar880_输入('Label53')}
function frmCalendar880_Label54_Click(){frmCalendar880_输入('Label54')}
function frmCalendar880_Label55_Click(){frmCalendar880_输入('Label55')}
function frmCalendar880_Label56_Click(){frmCalendar880_输入('Label56')}
//function 批量生成输入事件()
//{
//	//frmCalendar880.日期控件数组=$$.getMatrix(42,7).map2d(x=>frmCalendar880["Label"+(x+15)]);
//	$$.getMatrix(42,7).map2d(x=>{
//		console.log(`function frmCalendar880_Label${x+15}_Click(){frmCalendar880_输入('Label${x+15}')}`)
//	})
//}
//本模块里不要添加代码
/**
 * frmCalendar880_Frame1_Click Macro
 */

//年

//Calendar880.showInFrm(UserForm1.TextEdit1);
//Calendar880.showInRange(Target,Sheets("日历").Range("C17:I24"),"yyyy-MM-dd")


