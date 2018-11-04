<SCRIPT LANGUAGE=VBScript RUNAT=Server>
function CreateVBArray(elem1,elem2,elem3,elem4)

	elem1 = "" + elem1
	elem2 = "" + elem2
	elem3 = "" + elem3
	elem4 = "" + elem4

	if (Len(elem1) = 0) then
		elem1 = Empty
	end if

	if (Len(elem2) = 0) then
		elem2 = Empty
	end if

	if (Len(elem3) = 0) then
		elem3 = Empty
	end if

	if (Len(elem4) = 0) then
		elem4 = Empty
	end if

	CreateVBArray = Array(elem1,elem2,elem3,elem4)

end function
</SCRIPT>


<SCRIPT LANGUAGE=JavaScript RUNAT=Server>

function CreateMMConnection(ConnectionString,UserName,Password,Timeout)
{
	var Object;
	Object = new MMConnection(ConnectionString,UserName,Password,Timeout);
	return Object;
}

function MMConnection(ConnectionString,UserName,Password,Timeout)
{
	MMConnReconnect(this);
	this.isOpen = false;
	this.ConnectionString = ConnectionString;
	this.UserName		  = String(UserName);
	this.Password		  = String(Password);
	this.Connection		  = Server.CreateObject("ADODB.Connection");
	this.Connection.ConnectionTimeout = Timeout;
}


function MMConnReconnect(Object)
{
	Object.GetODBCDSNs				= ConnGetODBCDSNs;
	Object.Open						= ConnOpen;
	Object.GetTables				= ConnGetTables;
	Object.GetViews					= ConnGetViews;
	Object.GetProcedures			= ConnGetProcedures;
	Object.GetColumnsOfTable		= ConnGetColumns;
	Object.GetParametersOfProcedure = ConnGetParametersOfProcedure;
	Object.ExecuteSQL				= ConnExecuteSQL;
	Object.ExecuteSP				= ConnExecuteSP;
	Object.ReturnsResultSet			= ConnReturnsResultSet;
	Object.SupportsProcedure		= ConnSupportsProcedure;
	Object.GetProviderTypes			= ConnGetProviderTypes;
	Object.HandleExceptions			= ConnHandleExceptions;
	Object.TestOpen					= ConnIsOpen;
	Object.Close					= ConnClose;
}


function ConnOpen()
{
	theConnectionString = this.ConnectionString;
	if (this.UserName && this.UserName.length)
	{
		theConnectionString = theConnectionString + ";uid=" + this.UserName;
	}
	if (this.Password && this.Password.length)
	{
		theConnectionString = theConnectionString + ";pwd=" + this.Password;
	}

	var aConn = ConnEval(theConnectionString);
	this.Connection.Open(aConn);
	this.isOpen = (this.Connection.State == adStateOpen);
}

function ConnIsOpen()
{
	var xmlOutput = "";

	if (this.isOpen)
	{
		xmlOutput = xmlOutput + "<TEST status=";
		xmlOutput = xmlOutput + this.isOpen;
		xmlOutput = xmlOutput + "></TEST>";
	}
	else
	{
		xmlOutput = this.HandleExceptions();
	}

	return xmlOutput;
}

function ConnClose()
{
	if (this.Connection && this.isOpen)
	{
		this.Connection.Close();
	}
}

function ConnGetTables(SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		var VBVariant = new VBArray(CreateVBArray(CatalogName,SchemaName,"","TABLE"));
		return MarshallRecordsetIntoHTML(this.Connection.OpenSchema(adSchemaTables,VBVariant));
	}

	return null;
}

function ConnGetViews(SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		var VBVariant = new VBArray(CreateVBArray(CatalogName,SchemaName,"","VIEW"));
		return MarshallRecordsetIntoHTML(this.Connection.OpenSchema(adSchemaTables,VBVariant));
	}

	return null;
}

function ConnGetProcedures(SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		var VBVariant = new VBArray(CreateVBArray(CatalogName,SchemaName,"",""));
		return MarshallRecordsetIntoHTML(this.Connection.OpenSchema(adSchemaProcedures,VBVariant));
	}

	return null;
}

function ConnGetColumns(TableName,SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		var VBVariant = new VBArray(CreateVBArray(CatalogName,SchemaName,TableName,""));
		return MarshallRecordsetIntoHTML(this.Connection.OpenSchema(adSchemaColumns,VBVariant));
	}

	return null;
}

function ConnGetParametersOfProcedure(ProcedureName,SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		var VBVariant = new VBArray(CreateVBArray(CatalogName,SchemaName,ProcedureName,""));
		return this.Connection.OpenSchema(adSchemaProcedureParameters,VBVariant);
	}

	return null;
}

function ConnExecuteSQL(aStatement,MaxRows)
{
	if (this.Connection && this.isOpen)
	{
		var oRecordset = Server.CreateObject("ADODB.Recordset");
		if (oRecordset)
		{
			aStatement = "" + aStatement;
			oRecordset.MaxRecords = MaxRows;
			oRecordset.Open(aStatement,this.Connection);
			return MarshallRecordsetIntoHTML(oRecordset);
		}
	}

	return null;
}

function ConnGetProviderTypes()
{
	if (this.Connection && this.isOpen)
	{
		return MarshallRecordsetIntoHTML(this.Connection.OpenSchema(adSchemaProviderTypes));
	}

	return null;
}

function ConnExecuteSP(aProcStatement,TimeOut,Parameters)
{
	if (this.Connection && this.isOpen)
	{
		var oCommand = Server.CreateObject("ADODB.Command");

		aProcStatement = "" + aProcStatement;
		oCommand.CommandTimeout = TimeOut;
		oCommand.CommandText = aProcStatement;
		oCommand.CommandType = adCmdStoredProc;
		oCommand.ActiveConnection = this.Connection;

		Parameters = "" + Parameters;

		if (!Parameters.length)
		{
			if (oCommand)
			{
				return MarshallRecordsetIntoHTML(oCommand.Execute());
			}
		}
		else
		{
			//Substitute Parameters.
			var Params = Parameters;
			var ParamArray = new Array();

			if (Params && Params != "undefined")
			{
				var cSize = 0;
				for (;;)
				{
					var index = Params.indexOf(",");
					if (index == -1)
					{
						index = Params.length;
					}

					var name = Params.substring(0,index);

					Params = Params.substring(index+1,Params.length);
					index = Params.indexOf(",");
					if (index == -1)
					{
						index = Params.length;
					}

					var value = Params.substring(0,index);

					var Pair = new Object();

					Pair.name = name;
					Pair.value = value;

					ParamArray[cSize] = Pair;
					cSize++;

					if (index >= Params.length)
					{
						break;
					}

					Params = Params.substring(index+1,Params.length);
				}


				if (oCommand.Parameters.Count == -1)
				{
					//Create Parameters
					var oRecordset = ConnGetParametersOfProcedure(aProcStatement);
					if (oRecordset)
					{
						var pCount=0;
						while (!oRecordset.EOF)
						{
							var pName    = oRecordset.Fields.Item("PARAMETER_NAME").Value;
							var pOrdinal = oRecordset.Fields.Item("ORDINAL_POSITION").Value;
							var pType	 = oRecordset.Fields.Item("PARAMETER_TYPE").Value;
							var pDataType = oRecordset.Fields.Item("DATA_TYPE").Value;
							switch (pDataType)
							{
								case adBinary:
								case adBSTR:
								case adChar:
								case adLongVarBinary:
								case adLongVarChar:
								case adLongVarWChar:
								case adLongVarChar:
								case adVarBinary:
								case adVarChar:
								case adVarWChar:
								{
									var pSize = oRecordset.Fields.Item("CHARACTER_MAXIMUM_LENGTH").Value;
								}
								default:
								{
									var pSize = null;
								}
							}

							if ((pType == adParamInput) || (pType == adParamInputOutput))
							{
								var pValue = ParamArray[pName];
								//if we could not find parameter by name ..try to find 
								//parameter by index.
								if (!pValue)
								{
									//try the case when the parameter is set by index.
									pStrCount = "" + pCount;
									pValue = ParamArray[pStrCount];
								}
								oCommand.CreateParameter(pName,pDataType,pType,pSize,pValue);
							}
							else
							{
								var pValue = null;
								oCommand.CreateParameter(pName,pDataType,pType,pSize,pValue);
							}
							oRecordset.MoveNext();
							pCount++;
						}
					}	
				 }
				 else
				 {
					for (var i =0 ; i < ParamArray.length ; i++)
					{
						Pair = ParamArray[i];

						if (Pair.value)
						{
							var pIndex = "" + parseInt(Pair.name);

							if (pIndex == Pair.name)
							{
								var aParameter = oCommand.Parameters(parseInt(Pair.name));
							}
							else
							{
								var aParameter = oCommand.Parameters(Pair.name);
							}

							if (aParameter)
							{
								if ((aParameter.Direction == adParamInput) || (aParameter.Direction == adParamInputOutput))
								{
									aParameter.Value = Pair.value;
								}
							}
						}
					}
					return MarshallRecordsetIntoHTML(oCommand.Execute());
				 }
			}
		}
	}

	return null;
}

function ConnReturnsResultSet(ProcedureName,SchemaName,CatalogName)
{
	if (this.Connection && this.isOpen)
	{
		var VBVariant =  new VBArray(CreateVBArray(CatalogName,SchemaName,ProcedureName,""));
		var oRecordset = this.Connection.OpenSchema(adSchemaProcedureColumns,VBVariant);

		var status = "true";
		if (oRecordset.EOF) 
		{
			status = "false";
		}

		var xmlOutput = "";
		xmlOutput = xmlOutput + "<RETURNSRESULTSET status=";
		xmlOutput = xmlOutput + status;
		xmlOutput = xmlOutput + "></RETURNSRESULTSET>";
		return xmlOutput;
	}
}

function ConnSupportsProcedure()
{	
	if (this.Connection && this.isOpen)
	{
		var aProvider = "" + this.Connection.Provider;

		var status = "true";

		if (aProvider.indexOf("Microsoft.Jet") != -1)
		{
			status = "false";
		}

		if (aProvider.indexOf("MSDASQL")!=-1)
		{
			var ProviderTypes = this.Connection.OpenSchema(adSchemaProviderTypes);

			if (ProviderTypes.Fields.Count > 0)
			{
				//Access
				aProviderType = ProviderTypes.Fields(0).Value;
				aProviderType = aProviderType.toLowerCase();

				if (aProviderType == "guid")
				{
					status = "false";
				}//Paradox/DBaseIII.
				else if (aProviderType == "short")
				{
					status = "false";
				}
				else if (aProviderType == "image")
				{
					status = "false";
				}
				else if (aProviderType == "logical")
				{
					status = "false";
				} //For FoxPro
				else if (aProviderType == "l")
				{
					status = "false";
				} //For MySQL....
				else if (aProviderType == "tinyint")
				{
					status = "false";
				}
			}
		}

		var xmlOutput = "";
		xmlOutput = xmlOutput + "<SUPPORTSPROCEDURE status=";
		xmlOutput = xmlOutput + status;
		xmlOutput = xmlOutput + "></SUPPORTSPROCEDURE>";
		return xmlOutput;
	}
}

function ConnHandleExceptions()
{
	var xmlOutput = "";

	xmlOutput = xmlOutput + "<ERRORS>";
	if (this.Connection)
	{
		var Errors = this.Connection.Errors;

		for (var i =0 ; i < Errors.Count ; i++)
		{ 
			xmlOutput = xmlOutput + "<ERROR";

			xmlOutput = xmlOutput + " Identification=\""
			xmlOutput = xmlOutput + Errors(i).Number;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " Source=\""
			xmlOutput = xmlOutput + Errors(i).Source;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " Description=\""
			xmlOutput = xmlOutput + Errors(i).Description;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " HelpFile=\""
			xmlOutput = xmlOutput + Errors(i).HelpFile;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " HelpContext=\""
			xmlOutput = xmlOutput + Errors(i).HelpContext;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + "></ERROR>";
		}
	}
	xmlOutput = xmlOutput + "</ERRORS>";

	return xmlOutput;
}

function MarshallRecordsetIntoHTML(aResultSet)
{
	var xmlOutput = "";
	if (aResultSet)
	{
		xmlOutput = xmlOutput + "<RESULTSET>";
		xmlOutput = xmlOutput + "<FIELDS>";

		for(var i=0 ;i < aResultSet.Fields.Count ; i++)
		{
			xmlOutput = xmlOutput + "<FIELD";

			xmlOutput = xmlOutput + " name=\""
			xmlOutput = xmlOutput + aResultSet.Fields(i).Name;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " type=\""
			xmlOutput = xmlOutput + aResultSet.Fields(i).Type;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " definedSize=\""
			xmlOutput = xmlOutput + aResultSet.Fields(i).DefinedSize;
			xmlOutput = xmlOutput + "\"";


			xmlOutput = xmlOutput + " actualsize=\""

			if (!aResultSet.EOF)
			{
				xmlOutput = xmlOutput + aResultSet.Fields(i).ActualSize;
			}
			else
			{
				xmlOutput = xmlOutput + "-1";
			}

			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " precision=\""
			xmlOutput = xmlOutput + aResultSet.Fields(i).Precision;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + " scale=\""
			xmlOutput = xmlOutput + aResultSet.Fields(i).NumericScale;
			xmlOutput = xmlOutput + "\"";

			xmlOutput = xmlOutput + "></FIELD>";

		}

		xmlOutput = xmlOutput + "</FIELDS>";
		xmlOutput = xmlOutput + "<ROWS>";

		while (!aResultSet.EOF)
		{
			xmlOutput = xmlOutput + "<ROW";
			for(var i=0 ;i < aResultSet.Fields.Count ; i++)
			{
				xmlOutput = xmlOutput + " " + aResultSet.Fields(i).Name + "=\""
				var aValue = aResultSet.Fields(i).Value;
				if (aValue && aValue.length)
				{
					xmlOutput = xmlOutput + Server.HTMLEncode(aValue);
				}
				else
				{
					xmlOutput = xmlOutput + aResultSet.Fields(i).Value;
				}
				xmlOutput = xmlOutput + "\"";
			}
			xmlOutput = xmlOutput + "></ROW>";
			aResultSet.MoveNext()
		}

		xmlOutput = xmlOutput + "</ROWS>";
		xmlOutput = xmlOutput + "</RESULTSET>";
		aResultSet.Close();
	}
	return xmlOutput;
}

function ConnGetODBCDSNs()
{
   var fso = new ActiveXObject("Scripting.FileSystemObject");
   var dsnList=new Array();
   var OdbcIniFile = null;
   var e = new Enumerator(fso.Drives);
   var xmlOutput="";
   for (; !e.atEnd(); e.moveNext())
   {
	  var x = e.item();

	  //Skip Drive that not ready...
	  if (!fso.DriveExists(x) || !x.IsReady || (x.DriveType==1))
	  {
		continue;
	  }

	  var driverLetter = x.DriveLetter;
	  var WinFolderName1 = driverLetter + ":\\" + "Winnt";
	  var WinFolderName2 = driverLetter + ":\\" + "Windows";
	  if (fso.FolderExists(WinFolderName1))
	  {
		var odbcFileName = WinFolderName1 + "\\" + "ODBC.INI";
		if (fso.FileExists(odbcFileName))
		{
			//Get the ODBC FileName.
			OdbcIniFile = fso.OpenTextFile(odbcFileName,1,false);
		}
		break;
	  }
	  else if (fso.FolderExists(WinFolderName2))
	  {
		var odbcFileName = WinFolderName2 + "\\" + "ODBC.INI";
		if (fso.FileExists(odbcFileName))
		{
 			//Get the ODBC FileName.
			OdbcIniFile = fso.OpenTextFile(odbcFileName,1,false);
		}
		break;
	  }
   }
   if (OdbcIniFile)
   {
	 var i =0;
	 var odbcSection = -1;
	 while (!OdbcIniFile.AtEndOfStream)
	 {
		 var aLine = OdbcIniFile.ReadLine();
		 var odbcSection = aLine.indexOf("[ODBC");
		 if (odbcSection != -1)
		 {
			break;
		 }
	 }
	 if (odbcSection != -1)
	 {
		 while (!OdbcIniFile.AtEndOfStream)
		 {
			 var aLine = OdbcIniFile.ReadLine();
			 if (aLine.charAt(0) != "[")
			 {
				 var anIndex = aLine.indexOf("=");
				 if (anIndex != -1)
				 {
					var dsnName = aLine.substring(0,anIndex);
					dsnList[dsnList.length]= dsnName;
				 }
			}
			else
			{
				break;
			}
		 }
	  }
	 OdbcIniFile.Close();
   }

   xmlOutput = "<RESULTSET>";
   xmlOutput = xmlOutput + "<FIELDS>";
   xmlOutput = xmlOutput + "<FIELD";
   xmlOutput = xmlOutput + " name=\""
   xmlOutput = xmlOutput + "NAME";
   xmlOutput = xmlOutput + "\"";
   xmlOutput = xmlOutput + "></FIELD>";
   xmlOutput = xmlOutput + "</FIELDS>";
   xmlOutput = xmlOutput + "<ROWS>";

   if (dsnList.length)
   {
		for (var i =0 ; i < dsnList.length; i++)
		{
			xmlOutput = xmlOutput + "<ROW ";
			xmlOutput = xmlOutput + " NAME=\""
			xmlOutput = xmlOutput + dsnList[i];
			xmlOutput = xmlOutput + "\"";
			xmlOutput = xmlOutput + "></ROW>";
		}
   }

   xmlOutput = xmlOutput + "</ROWS>";
   xmlOutput = xmlOutput + "</RESULTSET>";

   return xmlOutput;
}

function ConnEval(ConnString)
{
	ConnString = "" + ConnString;
	if (ConnString.length)
	{
		var delimiter = (ConnString.indexOf("+") != -1) ? "+" : "&";
		var aConnString = "";

		for (;;)
		{
			var index = ConnString.indexOf(delimiter);

			if (index == -1)
			{
				index = ConnString.length;
			}

			var aStringlet	= ConnString.substring(0,index);
			aConnString = aConnString + eval(aStringlet);

			if (index >= ConnString.length)
			{
				break;
			}

			ConnString = ConnString.substring(index+1,ConnString.length);
		}

		return aConnString;
	}

	return ConnString;
}

</SCRIPT>

<SCRIPT LANGUAGE=JavaScript RUNAT=Server>

//---- ObjectStateEnum Values ----
var adStateClosed = 0x00000000;
var adStateOpen = 0x00000001;
var adStateConnecting = 0x00000002;
var adStateExecuting = 0x00000004;
var adStateFetching = 0x00000008;

//---- DataTypeEnum Values ----
var adEmpty = 0;
var adTinyInt = 16;
var adSmallInt = 2;
var adInteger = 3;
var adBigInt = 20;
var adUnsignedTinyInt = 17;
var adUnsignedSmallInt = 18;
var adUnsignedInt = 19;
var adUnsignedBigInt = 21;
var adSingle = 4;
var adDouble = 5;
var adCurrency = 6;
var adDecimal = 14;
var adNumeric = 131;
var adBoolean = 11;
var adError = 10;
var adUserDefined = 132;
var adVariant = 12;
var adIDispatch = 9;
var adIUnknown = 13;
var adGUID = 72;
var adDate = 7;
var adDBDate = 133;
var adDBTime = 134;
var adDBTimeStamp = 135;
var adBSTR = 8;
var adChar = 129;
var adVarChar = 200;
var adLongVarChar = 201;
var adWChar = 130;
var adVarWChar = 202;
var adLongVarWChar = 203;
var adBinary = 128;
var adVarBinary = 204;
var adLongVarBinary = 205;
var adChapter = 136;
var adFileTime = 64;
var adDBFileTime = 137;
var adPropVariant = 138;
var adVarNumeric = 139;

//---- PositionEnum Values ----
var adPosUnknown = -1;
var adPosBOF = -2;
var adPosEOF = -3;

//---- ParameterDirectionEnum Values ----
var adParamUnknown = 0x0000;
var adParamInput = 0x0001;
var adParamOutput = 0x0002;
var adParamInputOutput = 0x0003;
var adParamReturnValue = 0x0004;

//---- CommandTypeEnum Values ----
var adCmdUnknown = 0x0008;
var adCmdText = 0x0001;
var adCmdTable = 0x0002;
var adCmdStoredProc = 0x0004;
var adCmdFile = 0x0100;
var adCmdTableDirect = 0x0200;


//---- SchemaEnum Values ----
var adSchemaProviderSpecific = -1;
var adSchemaAsserts = 0;
var adSchemaCatalogs = 1;
var adSchemaCharacterSets = 2;
var adSchemaCollations = 3;
var adSchemaColumns = 4;
var adSchemaCheckConstraints = 5;
var adSchemaConstraintColumnUsage = 6;
var adSchemaConstraintTableUsage = 7;
var adSchemaKeyColumnUsage = 8;
var adSchemaReferentialConstraints = 9;
var adSchemaTableConstraints = 10;
var adSchemaColumnsDomainUsage = 11;
var adSchemaIndexes = 12;
var adSchemaColumnPrivileges = 13;
var adSchemaTablePrivileges = 14;
var adSchemaUsagePrivileges = 15;
var adSchemaProcedures = 16;
var adSchemaSchemata = 17;
var adSchemaSQLLanguages = 18;
var adSchemaStatistics = 19;
var adSchemaTables = 20;
var adSchemaTranslations = 21;
var adSchemaProviderTypes = 22;
var adSchemaViews = 23;
var adSchemaViewColumnUsage = 24;
var adSchemaViewTableUsage = 25;
var adSchemaProcedureParameters = 26;
var adSchemaForeignKeys = 27;
var adSchemaPrimaryKeys = 28;
var adSchemaProcedureColumns = 29;
var adSchemaDBInfoKeywords = 30;
var adSchemaDBInfoLiterals = 31;
var adSchemaCubes = 32;
var adSchemaDimensions = 33;
var adSchemaHierarchies = 34;
var adSchemaLevels = 35;
var adSchemaMeasures = 36;
var adSchemaProperties = 37;
var adSchemaMembers = 38;
</SCRIPT>


