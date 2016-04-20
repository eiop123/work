<% @Page Language="JScript" aspcompat="true" %>
<!--#include file="../../common/sjs/xmlUtil.js"-->
<!--#include file="../../common/sjs/common.js"-->
<%
var reqDoc=new XMLDoc(true);
var Result=new XMLDoc(true);
var ConnectStr=Application("ConnectStr");
Result.loadXML("<RESPONSE></RESPONSE>");
var reader=new System.IO.StreamReader(Request.InputStream);
var xmlData = reader.ReadToEnd();
if (xmlData!="")
	reqDoc.loadXML(xmlData);
else
	reqDoc.loadXML("<REQUEST></REQUEST>");
var Conn=Server.CreateObject("ADODB.Connection");
var ISEND_LIST_ID=0;
var Print_Flag = 0;

try
{
//Result.appendChild("DEBUG",ConnectStr);
	Conn.Open(ConnectStr);
	Conn.CursorLocation=3;
	var CMD = Server.CreateObject("Adodb.Command");
	CMD.ActiveConnection = Conn;
	CMD.CommandText="sp_edesk_home_save";
	CMD.CommandType=4;
	var Count=0;
//	var Print_Flag=reqDoc.getAttributeInt("PRINT_LIST");
	var Flow_Type1=reqDoc.getAttributeInt("FLOW_TYPE1");
	var Flow_Type2=reqDoc.getAttributeInt("FLOW_TYPE2");
	Print_Flag=reqDoc.getAttributeInt("PRINT_FLAG");
//Result.appendChild("DEBUG",Print_Flag);

	if (!((reqDoc.getAttributeInt("USER_ID")) && (reqDoc.getAttributeInt("ROLE_ID")) && (reqDoc.getAttributeInt("COME_NO_FLAG"))))
	{
		Result.appendChild("RETURN",1);		//1:無系統資訊(USER_ID,ROLE_ID)
	}
	else
	{
		var USER_ID=reqDoc.getAttributeInt("USER_ID");
		var ROLE_ID=reqDoc.getAttributeInt("ROLE_ID");
		var Come_No_Flag=reqDoc.getAttributeInt("COME_NO_FLAG");
		var IFlag=0;
		if (!(reqDoc.toFirstChild()))
		{
			Result.appendChild("RETURN",100);		//100:無收文資訊
		}
		else
		{
			do
			{
				try
				{
					try{
						if (Count==0)
							AppendPara();
						else
							SetParaValue();
					}catch(e){
						var errStr = "[資料值有誤]"
						if(Count==0)
							errStr+="[append]"
						else
							errStr+="[set]"
						errStr += "["+reqDoc.m_curNode.nodeName+"]["+reqDoc.m_curNode.text+"]"+e.description
						throw errStr
					}

					CMD.Execute();

					Result.appendChild("RV",CMD.Parameters("RETURN").Value);

					if ((CMD.Parameters("RETURN").Value=="0") && (IFlag==0))
					{
						ISEND_LIST_ID=CMD.Parameters("RETURN_SEND_LIST_ID").Value;
						IFlag=1;
					}
				}
				catch(e)
				{
					Result.appendChild("RV",2);
					Result.appendChild("ERROR",e.description);
					break;
				}
				reqDoc.toParent();
				Count++;
			}while (reqDoc.toNext("ITEM"));
		}
	}
}
catch(e)
{
	Result.appendChild("ERROR",e.description);
}
CMD=null;

Result.toChild("/RESPONSE");

Result.appendChild("SEND_LIST");
Result.toChild("SEND_LIST");
Result.setAttribute("SEND_LIST_ID",ISEND_LIST_ID);

if (Print_Flag==10)
{
	
	var Server_Date=new Date();
	var Server_Year=Server_Date.getFullYear()-1911;
	var Server_Month=Server_Date.getMonth()+1;
	var Server_Day=Server_Date.getDate();
	var Server_Hour=Server_Date.getHours();
	var Server_Min=Server_Date.getMinutes();
	Result.setAttribute("Y",Server_Year);
	Result.setAttribute("M",Server_Month);
	Result.setAttribute("D",Server_Day);
	Result.setAttribute("H",Server_Hour);
	Result.setAttribute("MIN",Server_Min);
	if (ISEND_LIST_ID>0)
	{
		try
		{
			var CMD = Server.CreateObject("Adodb.Command");
			CMD.ActiveConnection = Conn;
			CMD.CommandText="sp_get_send_list";
			CMD.Parameters.Append(CMD.CreateParameter("Send_List_ID",3,1,4,ISEND_LIST_ID));
			CMD.Parameters.Append(CMD.CreateParameter("FLOW_TYPE1",3,1,4,Flow_Type1));
			CMD.Parameters.Append(CMD.CreateParameter("FLOW_TYPE2",3,1,4,Flow_Type2));
			var Rec=CMD.EXECUTE();
			var SEND_TO_ID=0;
			if (Rec.State==1)
			{
				while (!(Rec.EOF))
				{
					if (SEND_TO_ID!=Rec.Fields("SEND_TO_DEPT_ID").Value)
					{
						Result.toChild("/RESPONSE/SEND_LIST");
						SEND_TO_ID=Rec.Fields("SEND_TO_DEPT_ID").Value;
						Result.appendChild("ITEM");
						Result.toLastChild();
						Result.setAttribute("LIST_NO",ISEND_LIST_ID);
						Result.setAttribute("SEND_TO_DEPT_ID",Check_Cols_Data(Rec.Fields("SEND_TO_DEPT_ID").Value));
						Result.setAttribute("SEND_TO_DEPT_IDNO",Check_Cols_Data(Rec.Fields("SEND_TO_DEPT_IDNO").Value));
						Result.setAttribute("SEND_TO_DEPT_NAME",Check_Cols_Data(Rec.Fields("SEND_TO_DEPT_NAME").Value));
						Result.setAttribute("REG_USER",Check_Cols_Data(Rec.Fields("REG_USER").Value));
						Result.setAttribute("REG_UNIT",Check_Cols_Data(Rec.Fields("REG_UNIT").Value));
						Result.appendChild("DOC_NO",Check_Cols_Data(Rec.Fields("DOC_NO").Value));
						Result.toLastChild("DOC_NO");
						Result.setAttribute("SEND_TO",Check_Cols_Data(Rec.Fields("SEND_TO").Value));
						Result.setAttribute("SEND_TO_NAME",Check_Cols_Data(Rec.Fields("SEND_TO_NAME").Value));
						Result.toParent();
						
					}
					else
					{
						Result.appendChild("DOC_NO",Check_Cols_Data(Rec.Fields("DOC_NO").Value));
						Result.toLastChild("DOC_NO");
						Result.setAttribute("SEND_TO",Check_Cols_Data(Rec.Fields("SEND_TO").Value));
						Result.setAttribute("SEND_TO_NAME",Check_Cols_Data(Rec.Fields("SEND_TO_NAME").Value));
						Result.toParent();
					}
					Rec.MoveNext();
				}
				Rec.Close();
			}
		}
		catch (e)
		{
			Result.appendChild("ERROR",e.description);
		}
	}
	else
	{
		Result.appendChild("ERROR");
		Result.toLastChild();
		Result.setAttribute("NO",1001);
	}
}
if (Conn.State==1)
        Conn.Close();
CMD=null;
Conn=null;
Response.ContentType = "text/xml";
Response.Write(Result.xml());

function AppendPara()
{
	CMD.Parameters.Append(CMD.CreateParameter("RETURN",3,4));
	reqDoc.toChild("RD_ID");
	Check_Before_AppendPara("RD_ID",3,1,4);

	reqDoc.toNext("COME_FROM_NAME");
	Check_Before_AppendPara("COME_FROM_NAME",202,1,255);
	reqDoc.toNext("DOC_TYPE");
	Check_Before_AppendPara("DOC_TYPE",202,1,2);
	reqDoc.toNext("COPY_TYPE");
	Check_Before_AppendPara("COPY_TYPE",202,1,2);
	reqDoc.toNext("SPEED");
	Check_Before_AppendPara("SPEED",202,1,2);
	reqDoc.toNext("SECRET");
	Check_Before_AppendPara("SECRET",202,1,2);
	reqDoc.toNext("DESECRET");
	Check_Before_AppendPara("DESECRET",202,1,30);
	reqDoc.toNext("COME_DATE");
	Check_Before_AppendPara("COME_DATE",135,1,8);
	reqDoc.toNext("COME_WORD");
	Check_Before_AppendPara("COME_WORD",202,1,30);
	//reqDoc.toNext("COME_YEAR");
	//Check_Before_AppendPara("COME_YEAR",129,1,3);
	reqDoc.toNext("COME_NO");
	Check_Before_AppendPara("COME_NO",202,1,30);
	reqDoc.toNext("MDOC_FLAG");
	Check_Before_AppendPara("MDOC_FLAG",3,1,4);
	reqDoc.toNext("MTYPE_NO");
	Check_Before_AppendPara("MTYPE_NO",202,1,1);
	reqDoc.toNext("DOC_CAT");
	Check_Before_AppendPara("DOC_CAT",3,1,4);
	reqDoc.toNext("EMAIL_DOC");
	Check_Before_AppendPara("EMAIL_DOC",202,1,1);
	reqDoc.toNext("LIMIT_DOC");
	Check_Before_AppendPara("LIMIT_DOC",202,1,1);
	reqDoc.toNext("ATT_TYPE");
	Check_Before_AppendPara("ATT_TYPE",202,1,2);
	reqDoc.toNext("ATT_NO");
	Check_Before_AppendPara("ATT_NO",3,1,4);
	reqDoc.toNext("ATT_UNIT");
	Check_Before_AppendPara("ATT_UNIT",202,1,10);
	reqDoc.toNext("ATT_DESC");
	Check_Before_AppendPara("ATT_DESC",202,1,100);
	reqDoc.toNext("ATT_NAME");
	Check_Before_AppendPara("ATT_NAME",202,1,50);
	reqDoc.toNext("SUBJECT");
	Check_Before_AppendPara("SUBJECT",202,1,255);
	reqDoc.toNext("REG_NO");
	Check_Before_AppendPara("REG_NO",202,1,20);
	reqDoc.toNext("SEND_TO");
	Check_Before_AppendPara("SEND_TO",3,1,4);
	reqDoc.toNext("SEND_TO_DEPT_ID");
	Check_Before_AppendPara("SEND_TO_DEPT_ID",3,1,4);
	reqDoc.toNext("PRE_RECV_NO");
	Check_Before_AppendPara("PRE_RECV_NO",129,1,3);
	reqDoc.toNext("MRECV_NO_ID");
	Check_Before_AppendPara("MRECV_NO_ID",3,1,4);
	reqDoc.toNext("MRECV_NO");
	Check_Before_AppendPara("MRECV_NO",202,1,15);
	reqDoc.toNext("POST_RECV_NO");
	Check_Before_AppendPara("POST_RECV_NO",202,1,12);
	reqDoc.toNext("RECV_NO");
	Check_Before_AppendPara("RECV_NO",202,1,30);
	reqDoc.toNext("RECV_TIME");
	Check_Before_AppendPara("RECV_TIME",135,1,8);
	reqDoc.toNext("SEND_TIME");
	Check_Before_AppendPara("SEND_TIME",135,1,8);
	reqDoc.toNext("FILE_ID");
	Check_Before_AppendPara("FILE_ID",3,1,4);
	reqDoc.toNext("RECV_DOC_STYLE");
	Check_Before_AppendPara("RECV_DOC_STYLE",3,1,4);
	reqDoc.toNext("REMARK");
	Check_Before_AppendPara("REMARK",202,1,255);
	reqDoc.toNext("ACC_ID");
	Check_Before_AppendPara("ACC_ID",3,1,4);
	reqDoc.toNext("EXG_FLAG");
	Check_Before_AppendPara("EXG_FLAG",3,1,4);
	reqDoc.toNext("DOC_NAME");
	Check_Before_AppendPara("DOC_NAME",202,1,50);
	reqDoc.toNext("FILE_GROUP_ID");
	Check_Before_AppendPara("FILE_GROUP_ID",3,1,4);
	reqDoc.toNext("ATT_FILE_FLAG");
	Check_Before_AppendPara("ATT_FILE_FLAG",3,1,4);
	CMD.Parameters.Append(CMD.CreateParameter("EFILE_FLAG",202,1,2,"0"));
	CMD.Parameters.Append(CMD.CreateParameter("COME_DOC_TYPE",202,1,2,"1"));
	CMD.Parameters.Append(CMD.CreateParameter("SEND_DOC_TYPE",202,1,2,"1"));
	CMD.Parameters.Append(CMD.CreateParameter("STATUS",202,1,1,"1"));
	reqDoc.toNext("ITEM_ID");
	CMD.Parameters.Append(CMD.CreateParameter("ITEM_ID",3,1,4,reqDoc.getTextInt()));
	reqDoc.toNext("ITEM_STATUS");
	CMD.Parameters.Append(CMD.CreateParameter("ITEM_STATUS",3,1,4,reqDoc.getTextInt()));
	reqDoc.toNext("NO");
	CMD.Parameters.Append(CMD.CreateParameter("NO",202,1,12,reqDoc.getText()));
	reqDoc.toNext("NO_STATUS");
	CMD.Parameters.Append(CMD.CreateParameter("NO_STATUS",3,1,4,reqDoc.getTextInt()));
	reqDoc.toNext("DEL_FLAG");
	CMD.Parameters.Append(CMD.CreateParameter("DEL_FLAG",3,1,4,reqDoc.getTextInt()));
	reqDoc.toNext("ATT_COMBINE");
	CMD.Parameters.Append(CMD.CreateParameter("ATT_COMBINE",202,1,255,reqDoc.getText()));
	//reqDoc.toNext("DOC_DUE_DATE");
	//CMD.Parameters.Append(CMD.CreateParameter("DOC_DUE_DATE",135,1,8,reqDoc.getText()));
	reqDoc.toNext("USER_TYPE");
	CMD.Parameters.Append(CMD.CreateParameter("USER_TYPE",3,1,4,reqDoc.getTextInt()));
	reqDoc.toNext("DOC_STEP_DUE_DATE");
	CMD.Parameters.Append(CMD.CreateParameter("DOC_STEP_DUE_DATE",135,1,8,reqDoc.getText()));
	reqDoc.toNext("SEND_TO_ACTION");
	CMD.Parameters.Append(CMD.CreateParameter("SEND_TO_ACTION",3,1,4,reqDoc.getTextInt()));
	reqDoc.toNext("MY_ACTION_NAME");
	CMD.Parameters.Append(CMD.CreateParameter("MY_ACTION_NAME",202,1,30,reqDoc.getText()));
	
	reqDoc.toNext("DUE_DAYS");
	Check_Before_AppendPara("DUE_DAYS",3,1,4);
	reqDoc.toNext("DUE_DAYS_TYPE");
	Check_Before_AppendPara("DUE_DAYS_TYPE",3,1,4);
	reqDoc.toNext("DOC_DUE_DATE");
	Check_Before_AppendPara("DOC_DUE_DATE",135,1,8);
	
	reqDoc.toNext("EXT_DATE_DATE1");
	Check_Before_AppendPara("EXT_DATE_DATE1",135,1,8);
	
	reqDoc.toNext("ATT_FILE_GRP_ID");
	Check_Before_AppendPara("ATT_FILE_GRP_ID",3,1,4);
	
	reqDoc.toNext("FLOW_SIGN_FLAG");
	Check_Before_AppendPara("FLOW_SIGN_FLAG",3,1,4);
	
	reqDoc.toNext("COME_DOC_METHOD");
	Check_Before_AppendPara("COME_DOC_METHOD",3,1,4);
	reqDoc.toNext("PAPER_AFTER_FLAG");
	Check_Before_AppendPara("PAPER_AFTER_FLAG",3,1,4);

	CMD.Parameters.Append(CMD.CreateParameter("SEND_LIST_FLAG",3,1,4,1));
	CMD.Parameters.Append(CMD.CreateParameter("SEND_LIST_ID",3,1,4,0));
	CMD.Parameters.Append(CMD.CreateParameter("RETURN_SEND_LIST_ID",3,2,4,0));

	CMD.Parameters.Append(CMD.CreateParameter("PRINT_FLAG",3,1,4,Print_Flag));
	CMD.Parameters.Append(CMD.CreateParameter("ROLE_ID",3,1,4,ROLE_ID));
	CMD.Parameters.Append(CMD.CreateParameter("USER_ID",3,1,4,USER_ID));
	CMD.Parameters.Append(CMD.CreateParameter("COME_NO_FLAG",3,1,4,Come_No_Flag));

	reqDoc.toNext("MASTER_COME_DOC_FLAG");
	Check_Before_AppendPara("MASTER_COME_DOC_FLAG",3,1,4);

	reqDoc.toNext("MEETING_DATES");
	Check_Before_AppendPara("MEETING_DATES",135,1,8);
	reqDoc.toNext("SUB_DOC_CAT");
	Check_Before_AppendPara("SUB_DOC_CAT",3,1,4);
	
	reqDoc.toNext("DOC_CASE");
	Check_Before_AppendPara("DOC_CASE",3,1,4);
}

function SetParaValue()
{
	reqDoc.toChild("RD_ID");
	Check_Before_SetParaValue(1,3);

	reqDoc.toNext("COME_FROM_NAME");
	Check_Before_SetParaValue(2,202);
	reqDoc.toNext("DOC_TYPE");
	Check_Before_SetParaValue(3,202);
	reqDoc.toNext("COPY_TYPE");
	Check_Before_SetParaValue(4,202);
	reqDoc.toNext("SPEED");
	Check_Before_SetParaValue(5,202);
	reqDoc.toNext("SECRET");
	Check_Before_SetParaValue(6,202);
	reqDoc.toNext("DESECRET");
	Check_Before_SetParaValue(7,202);
	reqDoc.toNext("COME_DATE");
	Check_Before_SetParaValue(8,135);
	reqDoc.toNext("COME_WORD");
	Check_Before_SetParaValue(9,202);
	//reqDoc.toNext("COME_YEAR");
	//Check_Before_SetParaValue(9,3);
	reqDoc.toNext("COME_NO");
	Check_Before_SetParaValue(10,202);
	reqDoc.toNext("MDOC_FLAG");
	Check_Before_SetParaValue(11,3);
	reqDoc.toNext("MTYPE_NO");
	Check_Before_SetParaValue(12,202);
	reqDoc.toNext("DOC_CAT");
	Check_Before_SetParaValue(13,3);
	reqDoc.toNext("EMAIL_DOC");
	Check_Before_SetParaValue(14,202);
	reqDoc.toNext("LIMIT_DOC");
	Check_Before_SetParaValue(15,202);
	reqDoc.toNext("ATT_TYPE");
	Check_Before_SetParaValue(16,202);
	reqDoc.toNext("ATT_NO");
	Check_Before_SetParaValue(17,3);
	reqDoc.toNext("ATT_UNIT");
	Check_Before_SetParaValue(18,202);
	reqDoc.toNext("ATT_DESC");
	Check_Before_SetParaValue(19,202);
	reqDoc.toNext("ATT_NAME");
	Check_Before_SetParaValue(20,202);
	reqDoc.toNext("SUBJECT");
	Check_Before_SetParaValue(21,202);
	reqDoc.toNext("REG_NO");
	Check_Before_SetParaValue(22,202);
	reqDoc.toNext("SEND_TO");
	Check_Before_SetParaValue(23,3);
	reqDoc.toNext("SEND_TO_DEPT_ID");
	Check_Before_SetParaValue(24,3);
	reqDoc.toNext("PRE_RECV_NO");
	Check_Before_SetParaValue(25,202);
	reqDoc.toNext("MRECV_NO_ID");
	Check_Before_SetParaValue(26,3);
	reqDoc.toNext("MRECV_NO");
	Check_Before_SetParaValue(27,202);
	reqDoc.toNext("POST_RECV_NO");
	Check_Before_SetParaValue(28,202);
	reqDoc.toNext("RECV_NO");
	Check_Before_SetParaValue(29,202);
	reqDoc.toNext("RECV_TIME");
	Check_Before_SetParaValue(30,135);
	reqDoc.toNext("SEND_TIME");
	Check_Before_SetParaValue(31,135);
	reqDoc.toNext("FILE_ID");
	Check_Before_SetParaValue(32,4);
	reqDoc.toNext("RECV_DOC_STYLE");
	Check_Before_SetParaValue(33,4);
	reqDoc.toNext("REMARK");
	Check_Before_SetParaValue(34,202);
	reqDoc.toNext("ACC_ID");
	Check_Before_SetParaValue(35,4);
	reqDoc.toNext("EXG_FLAG");
	Check_Before_SetParaValue(36,4);
	reqDoc.toNext("DOC_NAME");
	Check_Before_SetParaValue(37,202);
	reqDoc.toNext("FILE_GROUP_ID");
	Check_Before_SetParaValue(38,3);
	reqDoc.toNext("ATT_FILE_FLAG");
	Check_Before_SetParaValue(39,3);
	//CMD.Parameters.Append(CMD.CreateParameter("EFILE_FLAG",202,1,2,"0"));
	//CMD.Parameters.Append(CMD.CreateParameter("COME_DOC_TYPE",202,1,2,"1"));
	//CMD.Parameters.Append(CMD.CreateParameter("SEND_DOC_TYPE",202,1,2,"1"));
	//CMD.Parameters.Append(CMD.CreateParameter("STATUS",202,1,1,"1"));
	reqDoc.toNext("ITEM_ID");
	Check_Before_SetParaValue(44,3);
	reqDoc.toNext("ITEM_STATUS");
	Check_Before_SetParaValue(45,3);
	reqDoc.toNext("NO");
	Check_Before_SetParaValue(46,202);
	reqDoc.toNext("NO_STATUS");
	Check_Before_SetParaValue(47,3);
	reqDoc.toNext("DEL_FLAG");
	Check_Before_SetParaValue(48,3);
	reqDoc.toNext("ATT_COMBINE");
	Check_Before_SetParaValue(49,202);
	//reqDoc.toNext("DOC_DUE_DATE");
	//Check_Before_SetParaValue(47,135);
	reqDoc.toNext("USER_TYPE");
	Check_Before_SetParaValue(50,3);
	reqDoc.toNext("DOC_STEP_DUE_DATE");
	Check_Before_SetParaValue(51,135);
	reqDoc.toNext("SEND_TO_ACTION");
	Check_Before_SetParaValue(52,3);
	reqDoc.toNext("MY_ACTION_NAME");
	Check_Before_SetParaValue(53,202);
	
	reqDoc.toNext("DUE_DAYS");
	Check_Before_SetParaValue(54,3);
	reqDoc.toNext("DUE_DAYS_TYPE");
	Check_Before_SetParaValue(55,3);
	reqDoc.toNext("DOC_DUE_DATE");
	Check_Before_SetParaValue(56,135);
	
	reqDoc.toNext("EXT_DATE_DATE1");
	Check_Before_SetParaValue(57,135);
	
	reqDoc.toNext("ATT_FILE_GRP_ID");
	Check_Before_SetParaValue(58,3);
	
	reqDoc.toNext("FLOW_SIGN_FLAG");
	Check_Before_SetParaValue(59,3);
	
	reqDoc.toNext("COME_DOC_METHOD");
	Check_Before_SetParaValue(60,3);
	reqDoc.toNext("PAPER_AFTER_FLAG");
	Check_Before_SetParaValue(61,3);
	
	CMD.Parameters.Item(62).Value=0;
	CMD.Parameters.Item(63).Value=ISEND_LIST_ID;

	reqDoc.toNext("MASTER_COME_DOC_FLAG");
	Check_Before_SetParaValue(69,3);

	reqDoc.toNext("MEETING_DATES");
	Check_Before_SetParaValue(70,135);
	reqDoc.toNext("SUB_DOC_CAT");
	Check_Before_SetParaValue(71,3);
	
	reqDoc.toNext("DOC_CASE");
	Check_Before_SetParaValue(72,3);	
}
%>