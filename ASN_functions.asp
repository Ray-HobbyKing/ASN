<!-- #include file = "../db.asp" -->
<!-- #include file = "../BackendSecurity.asp" -->
<!-- #include file = "../../emailhead.asp" -->
 <!--#include virtual="/UNITEDHOBBIES/includes/sendmail.asp"--> 
<%
	access = 0
	poweruser = "Jasmine,Raymond.Lam,Ester.Chan,Stanley.Kan,Frederick.Lau,Natalie Lee,Jesi,"

	if instr(poweruser, session("bkusername")) > 0 then
		access = 1
		warehouse = "HK"
		iddept = 6
	else
		set rs1 = Server.CreateObject("ADODB.RecordSet")
			strsql =	"SELECT bd.name, bu.iddept " &_
						"FROM backenduser AS bu WITH(NOLOCK) " &_
						"LEFT JOIN backenduser_dept bd WITH(NOLOCK) ON bd.id = bu.iddept " &_
						"WHERE bu.username = '" & session("bkusername") & "'"
			rs1.Open strsql, pDatabaseConnectionString

			if not rs1.eof then
				select case rs1("name")
					case "Warehouse - AU"
						warehouse = "AU"
						fwhere = "AND left(fp.countrycode, 2) = 'AU' "
					case "Warehouse - BR"
						warehouse = "BR"
						fwhere = "AND left(fp.countrycode, 2) = 'BR' "
					case "Warehouse - UK"
						warehouse = "GB"
						fwhere = "AND left(fp.countrycode, 2) = 'GB' "
					case "Warehouse - US"
						warehouse = "US"
						fwhere = "AND left(fp.countrycode, 2) = 'US' "
					case "Warehouse - NL"
						warehouse = "NL"
						fwhere = "AND left(fp.countrycode, 2) = 'NL' "
					case else
						warehouse = "HK"
						fwhere = "AND fp.countrycode IS NULL "
				end select
				access = 1
				iddept = rs1("iddept")
			else
				response.write "You have no access right to this page, please contact IT."
				response.end
			end if
		rs1.close
	end if

	if request.form("mode") <> "" then
		select case lcase(request.form("mode"))
			case "showdata"
				call displayASNData( replace(replace(trim(request.form("idASN")), "''", ""), "--", "") )
			case "deletedata"
				call deletedata( replace(replace(trim(request.form("idASN")), "''", ""), "--", "") )
			case "showref"
				call showRef( replace(replace(trim(request.form("ASNType")), "''", ""), "--", "") )
			case "recordscan"
				recordScan()
			case "showlinedetails"
				showLineDetails()
			case "autosave"
				autoSave()
			case "updateqty"
				updateQty()
			case "goodsarrival"
				goodsArrival()
			case "goodsarrivalreverse"
				goodsArrivalReverse()
		end select
	end if

	function getUserInfo(bkusername)
		bkusername = replace(replace(trim(bkusername), "'", "''"), "--", "")
		set rsUser = Server.CreateObject("ADODB.RecordSet")
			strsql =	"SELECT wms.* " &_
						"FROM WMS_users AS wms WITH(NOLOCK) " &_
						"LEFT JOIN backenduser AS bu WITH(NOLOCK) ON bu.idUser = wms.idUser " &_
						"WHERE bu = '" & bkusername & "'"
			rsUser.Open strsql, pDatabaseConnectionString

			if not rsUser.eof then
				getUserInfo = rsUser("idUser") & "," & rsUser("permission") & "," & rsUser("active")
			else
				getUserInfo = false
			end if
		set rsUser = nothing
	end function

	function setUserPermission(idUser, permission)
		idUser     = clng(replace(replace(trim(idUser), "'", "''"), "--", ""))
		permission = clng(replace(replace(trim(permission), "'", "''"), "--", ""))

		set rsUser = Server.CreateObject("ADODB.RecordSet")
			strsql =	"UPDATE WMS_users SET permission = " & permission & " " &_
						"WHERE idUser = " & isUser & " " &_
						"IF @@ROWCOUNT = 0 " &_
							"SELECT 0 AS result " &_
						"ELSE " &_
							"SELECT 1 AS result"
			rsUser.execute strsql, pDatabaseConnectionString

			if rsUser("result") = 1 then
				setUserPermission = true
			else
				setUserPermission = false
			end if
		set rsUser = nothing
	end function

	sub createASNHeader(ASNtype, referenceNum)
		ASNtype = replace(replace(trim(ASNtype), "'", "''"), "--", "")
		referenceNum = replace(replace(trim(referenceNum), "'", "''"), "--", "")

		set connASN = Server.CreateObject("ADODB.Connection")
			connASN.Open pDatabaseConnectionString

			strsql =	"INSERT INTO ASN (createdBy, ASNtype, referenceNum) VALUES ((SELECT idUser FROM backenduser WITH(NOLOCK) WHERE username = '" & session("bkusername") & "'), " & ASNtype & ", '" & referenceNum & "');"
			connASN.Execute strsql

			set rs1 = Server.CreateObject("ADODB.RecordSet")
				strsql =	"SELECT idASN FROM ASN WITH(NOLOCK) " &_
							"WHERE active = 1 AND referenceNum = '" & referenceNum & "'"
				rs1.Open strsql, connASN

				if not rs1.eof then
					idASN = rs1("idASN")
				else
					response.write "Can't find record"
					response.end
				end if
			set rs1 = nothing

			' Type 2 is Intransit
			if ASNtype = 2 then
				strsql =	"INSERT INTO ASN_products (idASN, idProduct, SKU, NBReason, ExpQuantity, to_bin) " &_
							"SELECT " & idASN & ", bm.from_idproduct, p.sku, isnull(nb.NBReason, 0) AS NBReason, quantity, 'stock' " &_
							"FROM BINMovements AS bm WITH(NOLOCK) " &_
							"INNER JOIN products AS p WITH(NOLOCK) ON p.idproduct = bm.from_idproduct " &_
							"LEFT JOIN products_NBReason AS nb WITH(NOLOCK) ON nb.idproduct = p.idproduct " &_
							"WHERE bm.BINType = 6 AND obs = '" & referenceNum & "'"
				connASN.Execute strsql
			elseif ASNtype = 1 then
				strsql =	"INSERT INTO ASN_products (idASN, idProduct, SKU, NBStatus, ExpQuantity) " &_
							"SELECT " & idASN & ", p.idproduct, p.sku, isnull(nb.NBReason, 0) AS NBReason, a1.qty " &_
							"FROM products AS p WITH(NOLOCK) " &_
							"LEFT JOIN fproducts AS fp WITH(NOLOCK) ON fp.idproduct = p.idproduct " &_
							"LEFT JOIN products_NBReason AS nb WITH(NOLOCK) ON nb.idproduct = p.idproduct " &_
							"LEFT JOIN ( " &_
							"	SELECT DISTINCT p.sku, sum(pt.qty) AS qty " &_
							"	FROM po_table AS pt WITH(NOLOCK) " &_
							"	LEFT JOIN products AS p WITH(NOLOCK) ON p.idproduct = pt.idproduct " &_
							"	WHERE p.sku IN ( " &_
							"		SELECT DISTINCT sku " &_
							"		FROM products WITH(NOLOCK) " &_
							"		LEFT JOIN po_table WITH(NOLOCK) ON po_table.idproduct = products.idproduct " &_
							"		WHERE idpo = " & referenceNum & " " &_
							"	) " &_
							"	GROUP BY p.sku " &_
							") AS a1 ON a1.sku = p.sku " &_
							"WHERE fp.countrycode IS NULL " &_
							"AND p.sku IN ( " &_
							"	SELECT DISTINCT sku " &_
							"	FROM products WITH(NOLOCK) " &_
							"	LEFT JOIN po_table WITH(NOLOCK) ON po_table.idproduct = products.idproduct " &_
							"	WHERE idpo = " & referenceNum & " " &_
							") " &_
							"ORDER BY p.sku"
				connASN.Execute strsql
			end if

			response.write "<p>New ASN has been created</p>"
		set rsASN = nothing
	end sub

	sub createASNForm()
		%>
		<form id="ASN" name="ASN" action="ASN_Exec.asp" method="POST">
			<input type="hidden" name="createASNHeader" value="1" />
			<label for="ASNType">ASN Type</label>
			<select id="ASNType" name="ASNType" onchange="showRef(this.value);">
				<!-- <option value="1">Purchase Order</option> -->
				<option value="2">InTransit</option>
			</select>
			<br />
			<label for="refno">Reference Number</label>
			<span id="refSpan" name="refSpan">
				<input type="text" id="refno" name="refno" />
			</span>
			<br />
			<input type="submit" value="Create" />
		</form>
		<%
	end sub

	sub showRef(ASNtype)
		set rs = Server.CreateObject("ADODB.RecordSet")
			if ASNtype = 2 then
				strsql = 	"SELECT DISTINCT obs FROM BINMovements WITH(NOLOCK) " &_
							"LEFT JOIN ASN ON ASN.referenceNum = BINMovements.obs " &_
							"WHERE BINType = 6 AND transferCompleted = 0 " &_
							"AND (from_bin <> 'HK14F' AND to_bin <> 'HK15F') AND (from_bin <> 'HK15F' AND to_bin <> 'HK14F') " &_
							"ORDER BY obs"
							'"AND ASN.referenceNum <> BINMovements.obs " &_
				rs.Open strsql, pDatabaseConnectionString
				if not rs.eof then
		%>
		<select id="refno" name="refno">
		<%
					while not rs.eof
		%>
			<option value="<%=rs("obs")%>"><%=rs("obs")%></option>
		<%
						rs.movenext
					wend
		%>
		</select>
		<%
				end if
			elseif ASNtype = 1 then
				strsql =	"SELECT idpo FROM po_header WITH(NOLOCK) " &_
							"WHERE status = 2 AND del = 0 " &_
							"ORDER BY idpo"
				rs.Open strsql, pDatabaseConnectionString
				if not rs.eof then
		%>
		<select id="refno" name="refno">
		<%
					while not rs.eof
		%>
			<option value="<%=rs("idpo")%>">PO <%=rs("idpo")%></option>
		<%
						rs.movenext
					wend
		%>
		</select>
		<%
				end if
			end if
		set rs = nothing
	end sub

	sub importASNForm()
		%>
		<form id="frmASNData" name="frmASNData" action="ASN_upload.asp" method="POST" ENCTYPE="multipart/form-data">
			<input type="file" name="filename" />
			<br />
			<input type="hidden" id="uploadData" name="uploadData" value="1" />
			<input type="submit" value="upload" />
		</form>
		<%
	end sub

	sub goodsArrival()
		idASN = replace(replace(trim(request.form("idasn")), "--", ""), "'", "''")

		set conn = Server.CreateObject("ADODB.Connection")
			conn.Open pDatabaseConnectionString
			strsql =	"UPDATE ASN SET goodsArrivalDate = getDate(), goodsArrivalBy = (SELECT idUser FROM backenduser WHERE username = '" & session("bkusername") & "') " &_
						"WHERE idASN = " & idASN
			'response.write strsql & "<br />"
			conn.Execute strsql

			set rs = Server.CreateObject("ADODB.RecordSet")
				strsql =	"SELECT ASN.goodsArrivalDate, bu.username " &_
							"FROM ASN with(nolock) " &_
							"INNER JOIN backenduser AS bu with(nolock) ON bu.idUser = ASN.goodsArrivalBy " &_
							"WHERE idASN = " & idASN
				rs.open strsql, conn

				if not rs.eof then
					response.write rs("goodsArrivalDate") & "|||" & rs("username")
				end if
				rs.close
			set rs = nothing
			conn.close
		set conn = nothing
	end sub

	sub goodsarrivalreverse()
		idASN = replace(replace(trim(request.form("idasn")), "--", ""), "'", "''")

		set conn = Server.CreateObject("ADODB.Connection")
			conn.Open pDatabaseConnectionString

			strsql =	"UPDATE ASN SET goodsArrivalDate = NULL, goodsArrivalBy = NULL " &_
						"WHERE idASN = " & idASN
			conn.Execute strsql
			conn.close
		set conn = nothing
		response.write "<span id=""arrival_" & idASN & """><button onclick=""goodsArrival(" & idASN & ");"">Arrived</button></span>"
	end sub

	sub displayASNList()
		%>
		<h2><% if (iddept = 9 or iddept = 24 or iddept = 25 or iddept = 6) then %>Step 2 - <% end if %>Available ASN</h2>
		<div style="height: 400px; overflow: auto; padding: 1px;">
			<table cellpadding="5" cellspacing="0" style="width:100%;">
				<thead>
					<tr>
						<th>ASN ID</th>
						<th>Created Date</th>
						<th>ASN Type</th>
						<th>Receiving Warehouse</th>
						<th>Goods Arrived</th>
						<th>Reference Number</th>
						<th>Supplier</th>
						<th>Action</th>
					</tr>
				</thead>
				<tbody style="overflow:auto; height:300px;">
				<%
					set rsASN = Server.CreateObject("ADODB.RecordSet")
						strsql =	"SELECT a.*, bk.username, sup.supplierName, isnull(a.goodsArrivalDate, 0) AS goodsArrivalDate, bk2.username AS goodsArrivalBy, " &_
									"CAST((cast(floor(cast(getdate() as float)) AS datetime) - cast(floor(cast(a.createdDate as float)) AS datetime)) AS int) AS diff " &_
									"FROM ASN AS a WITH(NOLOCK) " &_
									"LEFT JOIN backenduser AS bk WITH(NOLOCK) ON bk.idUser = a.createdby " &_
									"LEFT JOIN backenduser AS bk2 WITH(NOLOCK) ON bk2.idUser = a.goodsArrivalBy " &_
									"LEFT JOIN suppliers AS sup WITH(NOLOCK) ON sup.idSupplier = a.idsupplier " &_
									"WHERE a.active = 1 " &_
									"ORDER BY a.idASN DESC"
						rsASN.Open strsql, pDatabaseConnectionString
						if not rsASN.eof then
							while not rsASN.eof
				%>
					<tr id="<%=rsASN("idASN")%>">
						<td style="text-align:center;">
							<a href="ASN_receiving.asp?type=po&amp;idASN=<%=rsASN("idASN")%>" target="_blank"><%=rsASN("idASN")%></a>
						</td>
						<td style="text-align:center;"><%=rsASN("createdDate")%><br /><span style="font-weight: bold;">[By: <%=rsASN("username")%>]</span></td>
						<td style="text-align:center;"><% if rsASN("ASNtype") = 1 then response.write "PO" else response.write "InTransit" %></td>
						<td style="text-align:center;"><%=ucase(rsASN("receivingWarehouse"))%></td>
						<td style="text-align:center;">
							<% if rsASN("goodsArrivalDate") = "1/01/1900" then %>
							<span id="arrival_<%=rsASN("idASN")%>">
								<button onclick="goodsArrival(<%=rsASN("idASN")%>);">Arrived</button>
							</span>
							<% else %>
							<span id="arrival_<%=rsASN("idASN")%>">
								<%=rsASN("goodsArrivalDate")%><br />
								<strong>[By: <%=rsASN("goodsArrivalBy")%>]</strong><br />
								<%
									if iddept = 6 then
								%>
								<button onclick="goodsArrivalReverse(<%=rsASN("idASN")%>);">Delete Date</button>
								<%
									end if
								%>
							</span>
							<% end if %>
						</td>
						<td style="text-align:center;">
							<a href="ASN_receiving.asp?type=po&amp;idASN=<%=rsASN("idASN")%>" target="_blank"><%=rsASN("referenceNum")%></a>
						</td>
						<td style="text-align:center;">
							<%
								response.write rsASN("idsupplier")
								if (iddept = 9 or iddept = 24 or iddept = 25 or iddept = 6) then
									response.write "<br />" & rsASN("supplierName")
								end if
							%>
						</td>
						<td style="text-align:center;">
							<% if (iddept = 9 or iddept = 24 or iddept = 25 or iddept = 6) then %>
							<a href="ASN_edit.asp?idASN=<%=rsASN("idASN")%>" target="_blank">Edit</a>&nbsp;
							<button onclick="deletedata(<%=rsASN("idASN")%>);">Delete</button><br />
							<% end if %>
							<a href="ASN_print.asp?idASN=<%=rsASN("idASN")%>" target="_blank">Print</a>
						</td>
					</tr>
				<%
								rsASN.movenext
							wend
						end if
					set rsASN = nothing
				%>
				</tbody>
			</table>
		</div>
		<%
	end sub

	sub displayASNData(idASN)
		%>
		<h1>ASN ID: <%=idASN%></h1>
		<table style="width: 100%;" border="1" cellpadding="5" cellspacing="0">
			<thead>
				<tr>
					<th>SKU</th>
					<th>Bar Code</th>
					<th>Quantity</th>
				</tr>
			</thead>
			<tbody>
		<%
		set rsData = Server.CreateObject("ADODB.RecordSet")
			strsql =	"SELECT sku, idproduct, quantity " &_
						"FROM ASN_products WITH(NOLOCK) " &_
						"WHERE idASN = " & idASN & " " &_
						"ORDER BY sku"
			rsData.Open strsql, pDatabaseConnectionString

			if not rsData.eof then
				while not rsData.eof %>
				<tr>
					<td style="text-align:center;"><%=rsData("sku")%></td>
					<td style="text-align:center;"><%=rsData("idproduct")%></td>
					<td style="text-align:center;"><%=rsData("quantity")%></td>
				</tr><%
					rsData.movenext
				wend
			end if
		set rsData = nothing
		%>
			</tbody>
		</table>
		<%
	end sub

	sub deletedata(idASN)
		set delASN = Server.CreateObject("ADODB.Connection")
			delASN.Open pDatabaseConnectionString

			strsql =	"DELETE FROM ASN WHERE idASN = " & idASN
			delASN.Execute strsql

			strsql =	"DELETE FROM ASN_products " &_
						"WHERE idASN = " & idASN
			delASN.Execute strsql

			response.write "<p>ASN has been deleted</p>"

		set delASN = nothing
	end sub

	sub ASNcomplete(idASN)
		set connComplete = Server.CreateObject("ADODB.Connection")
			connComplete.Open pDatabaseConnectionString

			strsql =	"INSERT INTO ASN_products (sku, idASN, idproduct, quantity, NBStatus, poNumber, receivingWarehouse, remark, to_bin, receivingDate) " &_
						"SELECT DISTINCT tmp.sku, idASN, idproduct, a.qty, NBStatus, poNumber, receivingWarehouse, remark, to_bin, getdate() " &_
						"FROM ASN_ReceivingLine as tmp WITH(NOLOCK) " &_
						"LEFT JOIN ( " &_
						"	SELECT sku, sum(quantity) AS qty " &_
						"	FROM ASN_ReceivingLine WITH(NOLOCK) " &_
						"	GROUP BY sku " &_
						") AS a ON a.sku = tmp.sku"
			connComplete.Execute strsql
		set connComplete = nothing
	end sub

	'------------------------------------------------------------------------------------------
	' Can be used by PO and InTransit
	'------------------------------------------------------------------------------------------
	sub sumUpStock(idASN)
		' Copy all the sum receive to ASN_products
		set rs3 = Server.CreateObject("ADODB.RecordSet")
			strsql =	"SELECT ap.idMap, sum(isnull(a.qty, 0)) AS qty " &_
						"FROM ASN_products AS ap WITH(NOLOCK) " &_
						"LEFT JOIN ( " &_
						"	SELECT ap3.idMap, sum(ReceivedQtyLine) AS qty " &_
						"	FROM ASN_ReceivingLine AS ar WITH(NOLOCK) " &_
						"	INNER JOIN ASN_products AS ap3 WITH(NOLOCK) ON ap3.idMap = ar.idMap " &_
						"	WHERE ap3.idProduct IN (SELECT ap2.idProduct FROM ASN_products AS ap2 WITH(NOLOCK) WHERE ap2.idASN = " & idASN & ") " &_
						"	GROUP BY ap3.idMap " &_
						") AS a ON a.idMap = ap.idMap " &_
						"WHERE ap.idASN = " & idASN & " AND ap.lineStatus = 0 " &_
						"GROUP BY ap.idMap"
			rs3.Open strsql, pDatabaseConnectionString

			if not rs3.eof then
				while not rs3.eof
					strsql =	"UPDATE ASN_products SET receivingDate = getDate(), receivedQty = " & rs3("qty") & ", lineStatus = 1 " &_
								"WHERE idMap = " & rs3("idMap") & " AND idASN = " & idASN
					conn1.Execute strsql
					rs3.movenext
				wend
			end if
		set rs3 = nothing
	end sub

	'------------------------------------------------------------------------------------------
	' PO Receiving Begin
	'------------------------------------------------------------------------------------------
		sub ASNStockAllocation(idASN)
			'NBConstraints = "2,3,5,6,7,8,9,10,12"
			' Modified by RL on 2015-01-19, requested by Nino
			' 2 = Minor update, 6 = Alternative suppliers, 7 = Quanlity Issues, 12 = Others, 17 = Wholesales Customization
			NBConstraints = "2,6,7,12,17"
			NBArray = split(NBConstraints, ",")

			set connAll = Server.CreateObject("ADODB.Connection")
				connAll.Open pDatabaseConnectionString
				set rsAll = Server.CreateObject("ADODB.RecordSet")
					strsql =	"SELECT ap.idMap, ap.idproduct, isnull(ap.receivedQty, 0) AS receivedQty, ap.remark, " &_
								"isnull(ap.NBReason, 0) AS NBReason, ap.to_BIN, ap.poNumber, " &_
								"ASN.receivingWarehouse, ASN.referenceNum " &_
								"FROM ASN_products AS ap WITH(NOLOCK) " &_
								"LEFT JOIN ASN WITH(NOLOCK) ON ASN.idASN = ap.idASN " &_
								"WHERE ap.idASN = " & idASN & " AND ap.lineStatus = 1"
					rsAll.Open strsql, connAll

					if not rsAll.eof then
						while not rsAll.eof
							idMap = rsAll("idMap")
							idproduct = rsAll("idproduct")
							receivedQty = rsAll("receivedQty")
							remark = rsAll("remark")
							NBReason = rsAll("NBReason")
							poNumber = rsAll("poNumber")

							' if the product has a matching NBReason

							'response.write "rsAll(""NBReason"")? " & rsAll("NBReason") & "<br />"
							'response.write "NBReason? " & inArray(NBArray, NBReason) & "<br />"
							'response.write "rsAll(""to_BIN"")? " & rsAll("to_BIN") & "<br />"
							'response.end

							if inArray(NBArray, NBReason) = true AND rsAll("to_BIN") = "" then
								if lcase(rsAll("receivingWarehouse")) = lcase("hk") then
									' Investigating Stocks HK
									set temp = Server.CreateObject("ADODB.Connection")
										temp.Open pDatabaseConnectionString
										set cmd = Server.CreateObject("ADODB.Command")
											set cmd.ActiveConnection = temp

											cmd.CommandText = "ASN_addToBin"
											' 4 = Stored Proc
											cmd.CommandType = 4

											cmd.Parameters(1) = session("bkusername")
											cmd.Parameters(2) = idProduct
											cmd.Parameters(3) = receivedQty
											cmd.Parameters(4) = 4
											cmd.Parameters(5) = poNumber
											cmd.Parameters(6) = remark
											cmd.Parameters(7) = idASN
											cmd.Parameters(8) = idMap

											cmd.Execute

											if cmd.Parameters(0) = 1 then
												call displayCheckPoint(idASN, 1)
											end if
										set cmd = nothing
									set temp = nothing
								elseif lcase(rsAll("receivingWarehouse")) = lcase("ca") then
									' Investigating Stocks CN
									set temp = Server.CreateObject("ADODB.Connection")
										temp.Open pDatabaseConnectionString
										set cmd = Server.CreateObject("ADODB.Command")
											set cmd.ActiveConnection = temp

											cmd.CommandText = "ASN_addToBin"
											' 4 = Stored Proc
											cmd.CommandType = 4

											cmd.Parameters(1) = session("bkusername")
											cmd.Parameters(2) = idProduct
											cmd.Parameters(3) = receivedQty
											cmd.Parameters(4) = 5
											cmd.Parameters(5) = poNumber
											cmd.Parameters(6) = remark
											cmd.Parameters(7) = idASN
											cmd.Parameters(8) = idMap

											cmd.Execute

											if cmd.Parameters(0) = 1 then
												call displayCheckPoint(idASN, 1)
											end if
										set cmd = nothing
									set temp = nothing
								end if
							else
								' If the product doesn't fall into specified NBReason, allocation to CSV allocated BIN
								to_bin = lcase(trim(rsAll("to_BIN")))

								if to_bin = "" then
									to_bin = "stock"
								end if

								tobinList = "Stock, Goods hold at forwarders,Suspended BIN,Stock Storage for OS (HK),Stock Storage for OS (CN),Stock Storage for HK (CN),Investigating Stocks (HK),Investigating Stocks (CN),Wholesale BIN"
								tobinArr = split(tobinList, ",")

								select case to_bin
									case "stock"
										set temp = Server.CreateObject("ADODB.Connection")
											temp.Open pDatabaseConnectionString
											set cmd = Server.CreateObject("ADODB.Command")
												set cmd.ActiveConnection = temp

												cmd.CommandText = "addStockRoutine"
												' 4 = Stored Proc
												cmd.CommandType = 4

												cmd.Parameters(1) = session("bkusername")
												cmd.Parameters(2) = idproduct
												cmd.Parameters(3) = receivedQty
												cmd.Parameters(4) = poNumber
												cmd.Parameters(5) = remark
												cmd.Parameters(6) = idASN
												cmd.Parameters(7) = idMap

												cmd.Execute

												if cmd.Parameters(0) = 1 then
													call displayCheckPoint(idASN, 1)
												end if
											set cmd = nothing
										set temp = nothing

									case "goods hold at forwarders"
										set temp = Server.CreateObject("ADODB.Connection")
											temp.Open pDatabaseConnectionString
											set cmd = Server.CreateObject("ADODB.Command")
												set cmd.ActiveConnection = temp

												cmd.CommandText = "ASN_addToszstock"
												' 4 = Stored Proc
												cmd.CommandType = 4

												cmd.Parameters(1) = session("bkusername")
												cmd.Parameters(2) = idProduct
												cmd.Parameters(3) = receivedQty
												cmd.Parameters(4) = poNumber
												cmd.Parameters(5) = remark
												cmd.Parameters(6) = idASN
												cmd.Parameters(7) = idMap

												cmd.Execute

												if cmd.Parameters(0) = 1 then
													call displayCheckPoint(idASN, 1)
												end if
											set cmd = nothing
										set temp = nothing

										'strsql =	"EXEC ASN_addToszstock '" & session("bkusername") & "', " & idproduct & ", " & receivedQty & ", " & poNumber & ", '" & remark & "', " & idASN
										'connAll.Execute strsql

									case "suspended bin"
										set temp = Server.CreateObject("ADODB.Connection")
											temp.Open pDatabaseConnectionString
											set cmd = Server.CreateObject("ADODB.Command")
												set cmd.ActiveConnection = temp

												cmd.CommandText = "ASN_addToBin"
												' 4 = Stored Proc
												cmd.CommandType = 4

												cmd.Parameters(1) = session("bkusername")
												cmd.Parameters(2) = idProduct
												cmd.Parameters(3) = receivedQty
												cmd.Parameters(4) = 1
												cmd.Parameters(5) = poNumber
												cmd.Parameters(6) = remark
												cmd.Parameters(7) = idASN
												cmd.Parameters(8) = idMap

												cmd.Execute

												if cmd.Parameters(0) = 1 then
													call displayCheckPoint(idASN, 1)
												end if
											set cmd = nothing
										set temp = nothing

										'strsql =	"EXEC ASN_addToBin '" & session("bkusername") & "', " & idproduct & ", " & receivedQty & ", 1, " & poNumber & ", '" & remark & "', " & idASN
										'connAll.Execute strsql

									case "stock storage for os (hk)"
										set temp = Server.CreateObject("ADODB.Connection")
											temp.Open pDatabaseConnectionString
											set cmd = Server.CreateObject("ADODB.Command")
												set cmd.ActiveConnection = temp

												cmd.CommandText = "ASN_addToszstock2"
												' 4 = Stored Proc
												cmd.CommandType = 4

												cmd.Parameters(1) = session("bkusername")
												cmd.Parameters(2) = idProduct
												cmd.Parameters(3) = receivedQty
												cmd.Parameters(4) = poNumber
												cmd.Parameters(5) = remark
												cmd.Parameters(6) = idASN
												cmd.Parameters(7) = idMap

												cmd.Execute

												if cmd.Parameters(0) = 1 then
													call displayCheckPoint(idASN, 1)
												end if
											set cmd = nothing
										set temp = nothing

										'strsql =	"EXEC ASN_addToszstock2 '" & session("bkusername") & "', " & idproduct & ", " & receivedQty & ", " & poNumber & ", '" & remark & "', " & idASN
										'connAll.Execute strsql

									case "stock storage for os (cn)"
										set temp = Server.CreateObject("ADODB.Connection")
											temp.Open pDatabaseConnectionString
											set cmd = Server.CreateObject("ADODB.Command")
												set cmd.ActiveConnection = temp

												cmd.CommandText = "ASN_addToBin"
												' 4 = Stored Proc
												cmd.CommandType = 4

												cmd.Parameters(1) = session("bkusername")
												cmd.Parameters(2) = idProduct
												cmd.Parameters(3) = receivedQty
												cmd.Parameters(4) = 2
												cmd.Parameters(5) = poNumber
												cmd.Parameters(6) = remark
												cmd.Parameters(7) = idASN
												cmd.Parameters(8) = idMap

												cmd.Execute

												if cmd.Parameters(0) = 1 then
													call displayCheckPoint(idASN, 1)
												end if
											set cmd = nothing
										set temp = nothing

										'strsql =	"EXEC ASN_addToBin '" & session("bkusername") & "', " & idproduct & ", " & receivedQty & ", 2, " & poNumber & ", '" & remark & "', " & idASN
										'connAll.Execute strsql

									case "stock storage for hk (cn)"
										set temp = Server.CreateObject("ADODB.Connection")
											temp.Open pDatabaseConnectionString
											set cmd = Server.CreateObject("ADODB.Command")
												set cmd.ActiveConnection = temp

												cmd.CommandText = "usp_ASN_addToszstock3"
												' 4 = Stored Proc
												cmd.CommandType = 4

												cmd.Parameters(1) = session("bkusername")
												cmd.Parameters(2) = idProduct
												cmd.Parameters(3) = receivedQty
												cmd.Parameters(4) = poNumber
												cmd.Parameters(5) = remark
												cmd.Parameters(6) = idASN
												cmd.Parameters(7) = idMap

												cmd.Execute

												if cmd.Parameters(0) = 1 then
													call displayCheckPoint(idASN, 1)
												end if
											set cmd = nothing
										set temp = nothing

										'strsql =	"EXEC ASN_addToszstock3 '" & session("bkusername") & "', " & idproduct & ", " & receivedQty & ", " & poNumber & ", '" & remark & "', " & idASN
										'connAll.Execute strsql

									case "investigating stocks (hk)"
										set temp = Server.CreateObject("ADODB.Connection")
											temp.Open pDatabaseConnectionString
											set cmd = Server.CreateObject("ADODB.Command")
												set cmd.ActiveConnection = temp

												cmd.CommandText = "ASN_addToBin"
												' 4 = Stored Proc
												cmd.CommandType = 4

												cmd.Parameters(1) = session("bkusername")
												cmd.Parameters(2) = idProduct
												cmd.Parameters(3) = receivedQty
												cmd.Parameters(4) = 4
												cmd.Parameters(5) = poNumber
												cmd.Parameters(6) = remark
												cmd.Parameters(7) = idASN
												cmd.Parameters(8) = idMap

												cmd.Execute

												if cmd.Parameters(0) = 1 then
													call displayCheckPoint(idASN, 1)
												end if
											set cmd = nothing
										set temp = nothing

										'strsql =	"EXEC ASN_addToBin '" & session("bkusername") & "', " & idproduct & ", " & receivedQty & ", 4, " & poNumber & ", '" & remark & "', " & idASN
										'connAll.Execute strsql

									case "investigating stocks (cn)"
										set temp = Server.CreateObject("ADODB.Connection")
											temp.Open pDatabaseConnectionString
											set cmd = Server.CreateObject("ADODB.Command")
												set cmd.ActiveConnection = temp

												cmd.CommandText = "ASN_addToBin"
												' 4 = Stored Proc
												cmd.CommandType = 4

												cmd.Parameters(1) = session("bkusername")
												cmd.Parameters(2) = idProduct
												cmd.Parameters(3) = receivedQty
												cmd.Parameters(4) = 5
												cmd.Parameters(5) = poNumber
												cmd.Parameters(6) = remark
												cmd.Parameters(7) = idASN
												cmd.Parameters(8) = idMap

												cmd.Execute

												if cmd.Parameters(0) = 1 then
													call displayCheckPoint(idASN, 1)
												end if
											set cmd = nothing
										set temp = nothing

										'strsql =	"EXEC ASN_addToBin '" & session("bkusername") & "', " & idproduct & ", " & receivedQty & ", 5, " & poNumber & ", '" & remark & "', " & idASN
										'connAll.Execute strsql

									case "wholesale bin"
										set temp = Server.CreateObject("ADODB.Connection")
											temp.Open pDatabaseConnectionString
											set cmd = Server.CreateObject("ADODB.Command")
												set cmd.ActiveConnection = temp

												cmd.CommandText = "ASN_addToBin"
												' 4 = Stored Proc
												cmd.CommandType = 4

												cmd.Parameters(1) = session("bkusername")
												cmd.Parameters(2) = idProduct
												cmd.Parameters(3) = receivedQty
												cmd.Parameters(4) = 12
												cmd.Parameters(5) = poNumber
												cmd.Parameters(6) = remark
												cmd.Parameters(7) = idASN
												cmd.Parameters(8) = idMap

												cmd.Execute

												if cmd.Parameters(0) = 1 then
													call displayCheckPoint(idASN, 1)
												end if
											set cmd = nothing
										set temp = nothing

									case else
										response.write "<p>Found nothing</p>"

								end select

							end if

							rsAll.movenext
						wend

					end if
				set rsAll = nothing
			set connAll = nothing
		end sub

		sub ASNStockAllocation2(idASN)
			NBConstraints = "2,6,7,12,17"
			NBArray = split(NBConstraints, ",")

			set connAll = Server.CreateObject("ADODB.Connection")
				connAll.Open pDatabaseConnectionString

				set rsAll = Server.CreateObject("ADODB.RecordSet")
					strsql =	"SELECT " & vbcrlf &_
								"	b.idpoline, a.sku, s.idproductMain, isnull(fp.countrycode, 'HK') AS countrycode, a.poNumber, a.idASN, a.receivedQty, " & vbcrlf &_
								"	b.qty AS [Ordered Qty], b.[Line Received], c.[SKU Ordered Qty], " & vbcrlf &_
								"	CONVERT(decimal(5, 0), b.qty) / CONVERT(decimal(5, 0), c.[SKU Ordered Qty]) AS [Ratio], " & vbcrlf &_
								"	CAST(a.receivedQty * CONVERT(decimal(5, 0), b.qty) / CONVERT(decimal(5, 0), c.[SKU Ordered Qty]) + 0.5 AS int) AS [Allocation] " & vbcrlf &_
								"FROM dbo.stock AS s " & vbcrlf &_
								"INNER JOIN dbo.products AS p ON p.idproduct = s.idproductMain " & vbcrlf &_
								"LEFT JOIN dbo.fproducts AS fp ON fp.idproduct = s.idproductMain " & vbcrlf &_
								"INNER JOIN ( " & vbcrlf &_
								"	SELECT " & vbcrlf &_
								"		ap.sku, ap.poNumber, ap.idASN, sum(ap.receivedQty) AS receivedQty " & vbcrlf &_
								"	FROM dbo.ASN_products AS ap " & vbcrlf &_
								"	WHERE ap.idASN = " & idASN & " " & vbcrlf &_
								"	GROUP BY ap.sku, ap.poNumber, ap.idASN " & vbcrlf &_
								") AS a ON a.sku = p.sku " & vbcrlf &_
								"INNER JOIN ( " & vbcrlf &_
								"	SELECT " & vbcrlf &_
								"		pt.idpoline, pt.idproduct, pt.qty, pt.qtyAr1 + pt.qtyAr2 + pt.qtyAr3 + pt.qtyAr4 AS [Line Received] " & vbcrlf &_
								"	FROM dbo.po_table AS pt " & vbcrlf &_
								"	INNER JOIN dbo.po_header AS ph WITH(NOLOCK) ON ph.idpo = pt.idpo AND ph.del = 0 AND ph.status NOT IN (1, 11) " & vbcrlf &_
								") AS b ON b.idproduct = s.idproductMain AND b.idpo = a.poNumber " & vbcrlf &_
								"INNER JOIN ( " & vbcrlf &_
								"	SELECT " & vbcrlf &_
								"		p.sku, sum(pt.qty) AS [SKU Ordered Qty] " & vbcrlf &_
								"	FROM dbo.po_table AS pt " & vbcrlf &_
								"	INNER JOIN dbo.products AS p ON p.idproduct = pt.idproduct " & vbcrlf &_
								"	WHERE pt.idpo = 9447 " & vbcrlf &_
								"	GROUP BY p.sku " & vbcrlf &_
								") AS c ON c.sku = p.sku"
					rsAll.Open strsql, connAll
					rsAll.close
				set rsAll = nothing
				connAll.close
			set connAll = nothing
		end sub

		sub fulfillPO(idASN)
			set conn = Server.CreateObject("ADODB.Connection")
				conn.Open pDatabaseConnectionString
				set rs1 = Server.CreateObject("ADODB.RecordSet")
					strsql =	"SELECT ap.idMap, isnull(ap.receivedQty, 0) AS receivedQty, ap.poNumber, ap.idproduct, isnull(pt.idpoline, 0) AS idpoline " &_
								"FROM ASN_products AS ap WITH(NOLOCK) " &_
								"LEFT JOIN po_table AS pt WITH(NOLOCK) ON pt.idpo = ap.poNumber AND pt.idproduct = ap.idproduct " &_
								"WHERE idASN = " & idASN
					rs1.Open strsql, pDatabaseConnectionString

					if not rs1.eof then
						while not rs1.eof
							'if rs1("receivedQty") > 0 then
								set temp = Server.CreateObject("ADODB.Connection")
									temp.Open pDatabaseConnectionString
									set cmd = Server.CreateObject("ADODB.Command")
										set cmd.ActiveConnection = temp

										cmd.CommandText = "fulfillPO"
										' 4 = Stored Proc
										cmd.CommandType = 4

										cmd.Parameters(1) = rs1("idproduct")
										cmd.Parameters(2) = rs1("receivedQty")
										cmd.Parameters(3) = rs1("poNumber")
										cmd.Parameters(4) = rs1("idpoline")
										cmd.Parameters(5) = idASN
										cmd.Parameters(6) = rs1("idMap")

										cmd.Execute

										if cmd.Parameters(0) = 1 then
											call displayCheckPoint(idASN, 2)
										end if
									set cmd = nothing
								set temp = nothing

								'strsql =	"EXEC fulfillPO " & rs1("idproduct") & ", " & rs1("receivedQty") & ", " & rs1("poNumber") & ", " & rs1("idpoline") & ", " & idASN
								'conn.Execute strsql
							'end if
							rs1.movenext
						wend
					end if
				set rs1 = nothing
			set conn = nothing
		end sub

		function emailCustomer(idproduct, stock_wds)
			set emailing = Server.CreateObject("ADODB.RecordSet")
				strsql =	"SELECT emailed, rowid, email, smallimageurl, description as productname, sku, emailme.idproduct, e1.lastSentDate " &_
							"FROM emailMe WITH(NOLOCK) " &_
							"INNER JOIN products WITH(NOLOCK) on emailme.idproduct = products.idproduct " &_
							"LEFT JOIN emailme1 AS e1 WITH(NOLOCK) ON e1.idproduct = emailme.idproduct " &_
							"WHERE emailme.active = 1 AND len(email) > 6 AND email <> '' AND email IS NOT NULL AND email like '%@%' " &_
							"AND emailed < 5 and emailme.idproduct = " & idproduct & " " &_
							"AND (e1.lastSentDate IS NULL OR (CAST(e1.lastSentDate AS date) < CAST(getDate() AS date)))"
				emailing.Open strsql, pDatabaseConnectionString

				if not emailing.eof then
					mail_from = "Arrival_Notive@HobbyKing.com"
					subject = "Shipment arrived for " & emailing("productname") & "!"
					textbody =  "<html>" &_
									"<font face='Arial, Helvetica, sans-serif'>" &_
										"You have added " & emailing("productname") & " to your watch list in www.hobbyking.com " &_
										"and this is an email to let you know that we have received a shipment just now!<br>" &_
										stock_wds & "Click below to view this item;<br>" &_
										"<a href='http://www.hobbyking.com/hobbyking/store/uh_viewItem.asp?idProduct=" & idproduct & "&utm_campaign=" & emailing("sku") &_
										"&utm_medium=email&utm_source=ARRIVAL'>" & emailing("productname") & "</a>" &_
										"<br>" &_
										"<img src='http://www.hobbyking.com/hobbyking/store/catalog/" & emailing("smallimageurl") & "'><br>"&_
									"</font>" &_
									"<br><br><br><br>" &_
									"<font face='Arial, Helvetica, sans-serif' style='font-size:9px'>" &_
										"If you don't wish to get this email anymore please log into your account and uncheck the notice.<br>" &_
										"You have been sent " & emailing("emailed") & " notices, notices will stop after 4 updates." &_
									"</font>" &_
								"</html>"
					while not emailing.eof
						On Error Resume Next
							'set objMail = Server.CreateObject("CDO.Message")
							'	objMail.To = replace(trim(emailing("email")), " ", "")
							'	objMail.From = mail_from
							'	objMail.Subject = subject
							'	objMail.HTMLBody = emailhead & textbody & emailfoot
							'	objMail.Send
							'set objMail = nothing
							'Changed by Fred on 20140218
							mailfrom = mail_from
							mailto = replace(trim(emailing("email")), " ", "")
							mailsubject = subject
							mailbody = textbody
							call sendmail ("ASN_functions", mailfrom, mailto, mailsubject, mailbody)
						On Error GoTo 0
						emailing.movenext
					wend
					set connemailCustomer = Server.CreateObject("ADODB.Connection")
						connemailCustomer.Open pDatabaseConnectionString
						strsql =	"UPDATE dbo.emailme1 SET lastSentDate = getdate() " &_
									"WHERE idproduct = " & idproduct & " " &_
									"IF @@ROWCOUNT = 0 " &_
									"INSERT INTO dbo.emailme1 (idproduct, lastSentDate) VALUES (" & idproduct & ", getdate())"
						connemailCustomer.Execute strsql
						connemailCustomer.close
					set connemailCustomer = nothing
				end if
			set emailing = nothing

		end function
	'------------------------------------------------------------------------------------------
	' PO Receiving End
	'------------------------------------------------------------------------------------------

	'------------------------------------------------------------------------------------------
	' InTransit Receiving Begin
	'------------------------------------------------------------------------------------------
		sub ASNStockAllocationInTransit(idASN)
			'NBConstraints = "2,3,5,6,7,8,9,10,12"
			NBConstraints = "2,6,7,12,17"
			NBArray = split(NBConstraints, ",")

			set connAll = Server.CreateObject("ADODB.Connection")
				connAll.Open pDatabaseConnectionString
				set rsAll = Server.CreateObject("ADODB.RecordSet")
					strsql =	"SELECT ap.idMap, ap.idproduct, isnull(ap.receivedQty, 0) AS receivedQty, ap.remark, isnull(ap.NBReason, 0) AS NBReason, ap.to_BIN, " &_
								"ASN.receivingWarehouse, ASN.referenceNum " &_
								"FROM ASN_products AS ap WITH(NOLOCK) " &_
								"LEFT JOIN ASN WITH(NOLOCK) ON ASN.idASN = ap.idASN " &_
								"WHERE ap.idASN = " & idASN & " AND ap.lineStatus = 1"
					rsAll.Open strsql, connAll

					if not rsAll.eof then
						while not rsAll.eof
							idMap       = rsAll("idMap")
							idproduct   = rsAll("idproduct")
							receivedQty = rsAll("receivedQty")
							remark      = rsAll("remark")
							NBReason    = rsAll("NBReason")
							obs         = rsALL("referenceNum")
							to_BIN      = rsALL("to_BIN")

							' if the product has a matching NBReason
							if inArray(NBArray, NBReason) = true AND rsAll("to_BIN") = "" then

							else
								to_bin = lcase(trim(rsAll("to_BIN")))

								select case to_bin
									case "stock"
										set temp = Server.CreateObject("ADODB.Connection")
											temp.Open pDatabaseConnectionString
											set cmd = Server.CreateObject("ADODB.Command")
												set cmd.ActiveConnection = temp

												cmd.CommandText = "addStockRoutineInTransit"

												' 4 = Stored Proc
												cmd.CommandType = 4

												cmd.Parameters(1) = session("bkusername")
												cmd.Parameters(2) = idMap
												cmd.Parameters(3) = idproduct
												cmd.Parameters(4) = idproduct
												cmd.Parameters(5) = receivedQty
												cmd.Parameters(6) = idASN
												cmd.Parameters(7) = "szstock"
												cmd.Parameters(8) = to_BIN
												cmd.Parameters(9) = obs

												cmd.Execute

												'response.write "<p>Returned: " & cmd.Parameters(0) & "</p>"

												if cmd.Parameters(0) = 1 then
													call displayCheckPoint(idASN, 1)
												end if
											set cmd = nothing
										set temp = nothing
								end select
							end if

							rsAll.movenext
						wend
					end if
				set rsAll = nothing
			set connAll = nothing
		end sub

		sub closeInTransit(idASN)
			set rs = Server.CreateObject("ADODB.RecordSet")
				strsql =	"SELECT referenceNum " &_
							"FROM ASN " &_
							"WHERE idASN = " & idASN
				rs.Open strsql, pDatabaseConnectionString

				if not rs.eof then
					batch = clng(split(rs("referenceNum"), "-")(1))
				else
					response.write "<p>In Transit reference number is not provided, please contact your supervisor</p>"
					response.end
				end if
			set rs = nothing

			set temp = Server.CreateObject("ADODB.Connection")
				temp.Open pDatabaseConnectionString
				set cmd = Server.CreateObject("ADODB.Command")
					set cmd.ActiveConnection = temp

					cmd.CommandText = "CompleteInTransit"
					' 4 = Stored Proc
					cmd.CommandType = 4

					cmd.Parameters(1) = idASN
					cmd.Parameters(2) = batch
					cmd.Parameters(3) = session("bkusername")

					cmd.Execute

					if cmd.Parameters(0) = 1 then
						call displayCheckPoint(idASN, 2)
					end if
				set cmd = nothing
			set temp = nothing
		end sub
	'------------------------------------------------------------------------------------------
	' InTransit Receiving End
	'------------------------------------------------------------------------------------------

	function inArray(arr, value)
		for each data in arr
			if int(data) = int(value) then
				'response.write "<p>" & int(data) & " vs " & int(value) & "</p>"
				inArray = true
				exit function
			end if
		next
		inArray = false
	end function

	' Used in ASN_receiving.asp
	function recordScan()
		idMap    = replace(replace(trim(request.form("idMap")), "''", ""), "--", "")
		ponumber = replace(replace(trim(request.form("ponumber")), "''", ""), "--", "")
		qty      = replace(replace(trim(request.form("qty")), "''", ""), "--", "")

		set conn = Server.CreateObject("ADODB.Connection")
			conn.Open pDatabaseConnectionString
			set rs111 = Server.CreateObject("ADODB.RecordSet")
				strsql =	"SELECT ASN.ASNType " &_
							"FROM ASN WITH(NOLOCK) " &_
							"INNER JOIN ASN_products AS ap WITH(NOLOCK) ON ap.idASN = ASN.idASN " &_
							"WHERE ap.idMap = " & idMap
				rs111.Open strsql, conn

				if not rs111.eof then
					if rs111("ASNType") = 1 then
						strsql =	"INSERT INTO ASN_ReceivingLine (idMap, ReceivedQtyLine, ReceivedDate, ReceivedBy, poNumber) VALUES " &_
									"(" & idMap & ", " & qty & ", getdate(), (SELECT idUser FROM backenduser WITH(NOLOCK) WHERE username = '" & session("bkusername") & "'), " & poNumber & ")"
					else
						strsql =	"INSERT INTO ASN_ReceivingLine (idMap, ReceivedQtyLine, ReceivedDate, ReceivedBy) VALUES " &_
									"(" & idMap & ", " & qty & ", getdate(), (SELECT idUser FROM backenduser WITH(NOLOCK) WHERE username = '" & session("bkusername") & "'))"
					end if
				end if
				conn.Execute strsql
			set rs111 = nothing
		set conn = nothing

		set rs = Server.CreateObject("ADODB.RecordSet")
			strsql =	"SELECT sum(ReceivedQtyLine) AS qty " &_
						"FROM ASN_ReceivingLine WITH(NOLOCK) " &_
						"WHERE idMap = " & idMap
			rs.Open strsql, pDatabaseConnectionString

			if not rs.eof then
				response.write rs("qty")
			end if
		set rs = nothing
	end function

	' Used in ASN_receiving.asp
	function showLineDetails()
		idASN        = replace(replace(trim(request.form("idASN")), "''", ""), "--", "")
		idproduct    = replace(replace(trim(request.form("idproduct")), "''", ""), "--", "")

		set rsShow = Server.CreateObject("ADODB.RecordSet")
			strsql =	"SELECT ap.*, p.smallImageUrl, isnull(a.qty, 0) AS qty " &_
						"FROM ASN_products AS ap WITH(NOLOCK) " &_
						"INNER JOIN ASN WITH(NOLOCK) ON ASN.idASN = ap.idASN " &_
						"LEFT JOIN products AS p WITH(NOLOCK) ON p.idproduct = ap.idproduct " &_
						"LEFT JOIN ( " &_
						"	SELECT idMap, sum(ReceivedQtyLine) AS qty " &_
						"	FROM ASN_ReceivingLine WITH(NOLOCK) " &_
						"	GROUP BY idMap " &_
						") AS a ON a.idMap = ap.idMap " &_
						"WHERE ASN.active = 1 AND ap.idproduct = " & idproduct & " AND ap.idASN = " & idASN & " " &_
						"ORDER BY ap.poNumber"
			rsShow.Open strsql, pDatabaseConnectionString

			if not rsShow.eof then
%>
					<input type="hidden" id="idproduct" name="idproduct" value="<%=idproduct%>" />
					<input type="hidden" id="NBReason" name="NBReason" value="<%=rsShow("NBReason")%>" />
					<table border="1" cellspacing="0" cellpadding="5">
						<thead>
							<tr>
								<th>PO</th>
								<th>Image</th>
								<th>Carton</th>
								<th>Barcode</th>
								<th>Expected Qty</th>
								<th>Recevied Qty</th>
								<th>Counted</th>
							</tr>
						</thead>
						<tbody>
<%
					while not rsShow.eof
						lineQty = 0
						if rsShow("qty") >= rsShow("ExpQuantity") then
							disabled = " DISABLED"
							style = " style=""background-color: #888888;"""
						else
							disabled = ""
							style = ""
						end if
%>
							<tr<%=style%>>
								<td class="center"><%=rsShow("poNumber")%></td>
								<td class="center">
									<% if rsShow("smallImageUrl") <> "" then %>
										<img src="http://www.hobbyking.com/hobbyking/store/catalog/<%=rsShow("smallImageUrl")%>">
									<% else %>
										No Image
									<% end if %>
								</td>
								<td class="center"><%=rsShow("cartonNumber")%></td>
								<td class="center"><%=rsShow("idproduct")%></td>
								<td class="center"><%=rsShow("ExpQuantity")%></td>
								<td class="center"><span id="Rec_<%=rsShow("idMap")%>_<%=rsShow("poNumber")%>"><%=rsShow("qty")%></span></td>
								<td class="center">
									<input type="text" id="qty_<%=rsShow("idMap")%>_<%=rsShow("poNumber")%>" class="enterQty" size="4"<%=disabled%> />
									<button id="<%=rsShow("idMap")%>_<%=rsShow("poNumber")%>" onmousedown="this.disabled=true; recordScan(this.id);" onclick="this.disabled=true; recordScan(this.id);"<%=disabled%>>Enter</button>
								</td>
							</tr>
<%
						rsShow.movenext
					wend
%>
						</tbody>
					</table>
					<button onclick="nextBarcode();">Next Barcode</button>
<%
			else
				response.write "<p>Barcode doesn't exist</p>"
			end if
		set rsShow = nothing
	end function

	function updateQty()
		idReceiving     = replace(replace(trim(request.form("idReceiving")), "''", ""), "--", "")
		ReceivedQtyLine = replace(replace(trim(request.form("ReceivedQtyLine")), "''", ""), "--", "")
		set updateConn = Server.CreateObject("ADODB.Connection")
			updateConn.Open pDatabaseConnectionString

			strsql =	"UPDATE ASN_ReceivingLine SET ReceivedQtyLine = " & ReceivedQtyLine & " " &_
						"WHERE idReceiving = " & idReceiving
			updateConn.Execute strsql
		set updateConn = nothing
	end function

	sub displayCheckPoint(idASN, stage)
		' First checkpoint
		set rs = Server.CreateObject("ADODB.RecordSet")
			strsql =	"SELECT count(idMap) AS counter " &_
						"FROM ASN_products WITH(NOLOCK) " &_
						"WHERE idASN = " & idASN & " AND lineStatus = " & (int(stage) - 1)
			rs.Open strsql, pDatabaseConnectionString

			if not rs.eof then
				if rs("counter") > 0 then
					response.write "<html>"
						response.write "<head>"
							response.write "<title>ASN Discrepancy Report</title>"
						response.write "</head>"
						response.write "<body>"
							response.write "<p>Checkpoint " & stage & ". Processing hasn't been completed. Please press F5 to refresh the page and rerun again</p>"
						response.write "</body>"
					response.write "</html>"
					response.end
				end if
			end if
		set rs = nothing
	end sub

	' Added by RL ON 2014-11-19 Begin
		function sendEmailToCustomer(idASN)
			dim updatedStock      : updatedStock      = 0
			dim halfMonthSales    : halfMonthSales    = 0
			dim stock_wds         : stock_wds         = ""
			dim debug             : debug             = 0
			dim previousIdProduct : previousIdProduct = ""
			dim sent1             : sent1             = 0

			dim email1
			email1 =  	"<html>" &_
							"<font face='Arial, Helvetica, sans-serif'>" &_
								"You have added <__PRODUCTNAME__> to your watch list in www.hobbyking.com " &_
								"and this is an email to let you know that we have received a shipment just now!<br><__STOCK_WDS__> Click below to view this item;<br>" &_
								"<a href='http://www.hobbyking.com/hobbyking/store/uh_viewItem.asp?idProduct=<__IDPRODUCT__>&utm_campaign=<__SKU__>&utm_medium=email&utm_source=ARRIVAL'><__PRODUCTNAME__></a>" &_
								"<br>" &_
								"<img src='http://www.hobbyking.com/hobbyking/store/catalog/<__SMALLIMAGEURL__>'><br>"&_
							"</font>" &_
							"<br><br><br><br>" &_
							"<font face='Arial, Helvetica, sans-serif' style='font-size:9px'>" &_
								"If you don't wish to get this email anymore please log into your account and uncheck the notice.<br>" &_
								"You have been sent <__EMAILED__> notices, notices will stop after 4 updates." &_
							"</font>" &_
						"</html>"

			' Email customer for new stock
			set conn = Server.CreateObject("ADODB.Connection")
				conn.Open pDatabaseConnectionString

				set rs = Server.CreateObject("ADODB.RecordSet")
					strsql =	"SELECT " & vbcrlf &_
								"	ap.idproduct, p.sku, p.description, p.smallimageurl, " & vbcrlf &_
								"	isnull(a.stock, 0) AS [Current Max Stock], isnull(b.[Received Qty], 0) AS [Received Qty], c.emailed, c.email " & vbcrlf &_
								"FROM ASN_products AS ap WITH(NOLOCK) " & vbcrlf &_
								"INNER JOIN products AS p WITH(NOLOCK) ON p.idproduct = ap.idproduct " & vbcrlf &_
								"INNER JOIN ( " & vbcrlf &_
								"	SELECT s.idproductMain AS idproduct, s.stock " & vbcrlf &_
								"	FROM stock AS s WITH(NOLOCK) " & vbcrlf &_
								"	WHERE s.stock > 0 " & vbcrlf &_
								") AS a ON a.idproduct = ap.idproduct " & vbcrlf &_
								"INNER JOIN ( " & vbcrlf &_
								"	SELECT idproduct, poNumber, sum(receivedQty) AS [Received Qty] " & vbcrlf &_
								"	FROM ASN_products WITH(NOLOCK) " & vbcrlf &_
								"	WHERE idASN = " & idASN & " " & vbcrlf &_
								"	GROUP BY idproduct, poNumber " & vbcrlf &_
								") AS b ON b.idproduct = ap.idproduct AND b.poNumber = ap.poNumber " & vbcrlf &_
								"INNER JOIN ( " & vbcrlf &_
								"	SELECT em.idproduct, em.emailed, em.email " & vbcrlf &_
								"	FROM dbo.emailMe AS em WITH(NOLOCK) " & vbcrlf &_
								"	LEFT JOIN dbo.emailMe1 AS em1 WITH(NOLOCK) ON em1.idproduct = em.idproduct " & vbcrlf &_
								"	WHERE em.active = 1 AND len(em.email) > 6 AND em.email <> '' AND em.email IS NOT NULL AND em.email LIKE '%@%' " & vbcrlf &_
								"		AND em.emailed < 5 AND (em1.lastSentDate IS NULL OR (CAST(em1.lastSentDate AS date) < CAST(getDate() AS date))) " & vbcrlf &_
								") AS c ON c.idproduct = ap.idproduct " & vbcrlf &_
								"WHERE p.active = -1 AND ap.notBINReason IS NULL AND ap.to_bin = 'stock' AND ap.idASN = " & idASN & " " & vbcrlf &_
								"GROUP BY ap.idproduct, p.sku, p.description, p.smallimageurl, a.stock, b.[Received Qty], c.emailed, c.email " & vbcrlf &_
								"HAVING sum(ap.receivedQty) > 0 " & vbcrlf &_
								"ORDER BY ap.idproduct"
					if debug then response.write "<pre>" & strsql & "</pre>"
					rs.CursorType     = 2
					rs.CursorLocation = 3
					rs.Open strsql, conn

					if not rs.eof then
						previousIdProduct = rs("idproduct")
						while not rs.eof
							sent1          = 0
							updatedStock   = clng(rs("Current Max Stock"))
							originalStock  = clng(updatedStock - rs("Received Qty"))
							idproduct      = rs("idproduct")
							sku            = rs("sku")
							description    = rs("description")
							smallimageurl  = rs("smallimageurl")
							emailed        = rs("emailed")
							email          = rs("email")

							'response.write "idproduct: " & idProduct & " - updatedStock: " & updatedStock & "<br />"
							'response.flush

							' Sending out restock email Begin
								if updatedStock > 5 AND originalStock <= 0 then
									if updatedStock < 100 then
										stock_wds = "<br><b>Only " & updatedStock & " items in stock, be quick!</b><br><br>"
									end if

									call restockedEmail(email1, idproduct, sku, description, smallimageurl, stock_wds, email, emailed)

									if previousIdProduct <> idproduct then
										' Moveback to previous record
										rs.MovePrevious
										' Set New arrival
										strsql =	"BEGIN TRANSACTION " & vbcrlf &_
													"	UPDATE dbo.emailMe1 SET lastSentDate = getDate() " & vbcrlf &_
													"	WHERE idproduct = " & rs("idproduct") & " " & vbcrlf &_
													"	IF @@ROWCOUNT = 0 " & vbcrlf &_
													"	INSERT INTO dbo.emailMe1 (idproduct, lastSentDate) VALUES " & vbcrlf &_
													"	(" & rs("idproduct") & ", getDate()); " & vbcrlf &_
													"	" & vbcrlf &_
													"	UPDATE dbo.stock SET changedate = getdate(), newstockshow = -1 " &_
													"	WHERE idproductMain = " & rs("idproduct") & "; " & vbcrlf &_
													"	" & vbcrlf &_
													"	UPDATE dbo.stockorders SET active = 0 " &_
													"	WHERE idproduct = " & rs("idproduct") & "; " & vbcrlf &_
													"	" & vbcrlf &_
													"	UPDATE dbo.emailMe set emailed = emailed + 1 " & vbcrlf &_
													"	WHERE idproduct = " & rs("idproduct") & "; " & vbcrlf &_
													"COMMIT;"
										conn.Execute strsql

										rs.movenext

										previousIdProduct = idproduct
									end if
								else
									if previousIdProduct <> idproduct then
										' Moveback to previous record
										rs.MovePrevious

										strsql =	"UPDATE dbo.stock SET changedate = getDate(), newstockshow = 0 " & vbcrlf &_
													"WHERE idproductMain = " & idproduct
										conn.Execute strsql

										rs.movenext

										previousIdProduct = idproduct
									end if
								end if
							' Sending out restock email End

							rs.movenext
						wend

						rs.MoveLast

						strsql =	"BEGIN TRANSACTION " & vbcrlf &_
									"	UPDATE dbo.emailMe1 SET lastSentDate = getDate() " & vbcrlf &_
									"	WHERE idproduct = " & rs("idproduct") & " " & vbcrlf &_
									"	IF @@ROWCOUNT = 0 " & vbcrlf &_
									"	INSERT INTO dbo.emailMe1 (idproduct, lastSentDate) VALUES " & vbcrlf &_
									"	(" & rs("idproduct") & ", getDate()); " & vbcrlf &_
									"	" & vbcrlf &_
									"	UPDATE dbo.stock SET changedate = getdate(), newstockshow = -1 " &_
									"	WHERE idproductMain = " & rs("idproduct") & "; " & vbcrlf &_
									"	" & vbcrlf &_
									"	UPDATE dbo.stockorders SET active = 0 " &_
									"	WHERE idproduct = " & rs("idproduct") & "; " & vbcrlf &_
									"	" & vbcrlf &_
									"	UPDATE dbo.emailMe set emailed = emailed + 1 " & vbcrlf &_
									"	WHERE idproduct = " & rs("idproduct") & "; " & vbcrlf &_
									"COMMIT;"
						conn.Execute strsql
					end if
					rs.close
				set rs = nothing
				conn.close
			set conn = nothing
		end function

		sub restockedEmail(emailtemp, idproduct, sku, description, smallimageurl, stock_wds, email, emailed)
			On Error Resume Next
				mail_from    = "Arrival_Notice@HobbyKing.com"
				mail_subject = "HobbyKing Arrival Notice - " & description & " is back in stock!"

				mail_body    = replace(emailtemp, "<__PRODUCTNAME__>", description)
				mail_body    = replace(mail_body, "<__STOCK_WDS__>", stock_wds)
				mail_body    = replace(mail_body, "<__IDPRODUCT__>", idproduct)
				mail_body    = replace(mail_body, "<__SKU__>", sku)
				mail_body    = replace(mail_body, "<__SMALLIMAGEURL__>", smallimageurl)
				mail_body    = replace(mail_body, "<__EMAILED__>", emailed)

				if Request.ServerVariables("SERVER_NAME") <> "174.143.95.154" then
					mail_to = email
					'mail_to = "jasmine.cheng@hobbyking.com,raymond.lam@hobbyking.com"
					'mail_body = email & "<br />" & mail_body
				else
					mail_to   = "raymond.lam@hobbyking.com,jesi.tse@hobbyking.com"
					mail_body = email & "<br />" & mail_body
				end if

				call sendmail ("ASN_functions", mail_from, mail_to, mail_subject, mail_body)
			On Error GoTo 0
		end sub

		sub wishlistEmail(idASN)
			On Error Resume Next
				dim wishEmail
				wishEmail =	"<html>" &_
								"<font face='Arial, Helvetica, sans-serif'>" &_
									"This is an email to let you know that a product from your WishList (<strong><__STRPRODUCTNAME__></strong>) has arrived.<br><__STOCK_WDS__>Click below to view this item;<br>" &_
									"<a href='http://www.hobbyking.com/hobbyking/store/uh_viewItem.asp?idProduct=<__PIDPRODUCT__>&utm_campaign=<__STRSKU__>&utm_medium=email&utm_source=ARRIVAL'><__STRPRODUCTNAME__></a><br>" &_
									"<img src='http://www.hobbyking.com/hobbyking/store/catalog/<__STRSMALLIMAGEURL__>'><br>" &_
								"</font><br><br>" &_
								"<font face='Arial, Helvetica, sans-serif' style='font-size:9px'>" &_
									"If you don't wish to get this email anymore please log into your account and uncheck the box in your WishList" &_
								"</font>"&_
							"</html>"
				mail_from  = "WishList_Alert@HobbyKing.com"

				set rs = Server.CreateObject("ADODB.RecordSet")
					strsql =	"SELECT rtrim(ltrim(c.email)) AS email, c.name, ap.idproduct, ap.sku, p.description, p.smallimageurl, s.stock " & vbcrlf &_
								"FROM dbo.ASN_products AS ap WITH(NOLOCK) " & vbcrlf &_
								"INNER JOIN dbo.products AS p WITH(NOLOCK) ON p.idproduct = ap.idproduct " & vbcrlf &_
								"INNER JOIN dbo.stock AS s WITH(NOLOCK) ON s.idproductMain = ap.idproduct " & vbcrlf &_
								"INNER JOIN dbo.wishlist AS w WITH(NOLOCK) ON w.idproduct = ap.idproduct " & vbcrlf &_
								"INNER JOIN dbo.customers AS c WITH(NOLOCK) ON c.idCustomer = w.idCustomer " & vbcrlf &_
								"WHERE p.active = -1 AND s.stock > 0 AND w.emailme = 1 AND len(c.email) > 6 AND c.email <> '' AND c.email IS NOT NULL AND c.email LIKE '%@%' " &_
								"	AND ap.idASN = " & idASN & " " & vbcrlf &_
								"ORDER BY ap.idproduct"
					rs.Open strsql, pDatabaseConnectionString

					if not rs.eof then
						while not rs.eof
							if rs("stock") < 100 then
								stock_wds = "<br><b>Only " & rs("stock") & " items in stock, be quick!</b><br><br>"
							else
								stock_wds = ""
							end if

							mail_subject = "A HobbyKing Wishlist item is back in stock! - " & rs("description")

							mail_body    = replace(wishEmail, "<__STRPRODUCTNAME__>", rs("description"))
							mail_body    = replace(mail_body, "<__STOCK_WDS__>", stock_wds)
							mail_body    = replace(mail_body, "<__PIDPRODUCT__>", rs("idproduct"))
							mail_body    = replace(mail_body, "<__STRSKU__>", rs("sku"))
							mail_body    = replace(mail_body, "<__STRSMALLIMAGEURL__>", rs("smallimageurl"))


							if Request.ServerVariables("SERVER_NAME") <> "174.143.95.154" then
								mail_to = rs("email")
								'mail_to = "jasmine.cheng@hobbyking.com,raymond.lam@hobbyking.com"
								'mail_body = rs("email") & "<br />" & mail_body
							else
								mail_to   = "raymond.lam@hobbyking.com,jesi.tse@hobbyking.com"
								mail_body = rs("email") & "<br />" & mail_body
							end if

							call sendmail ("ASN_functions", mail_from, mail_to, mail_subject, mail_body)
							rs.movenext
						wend
					end if
					rs.close
				set rs = nothing
			On Error GoTo 0
		end sub
	' Added by RL ON 2014-11-19 End
%>