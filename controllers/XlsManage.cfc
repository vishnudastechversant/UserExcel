<cfcomponent>
    <cffunction  name="verifyAndUploadXLSX" returntype ="struct" output="false">
        <cfset local.filePath = getTempDirectory()>
        <cfset local.errors = "">
        <cfset local.success = false>
        <cfset savedFile = "">
        <cffile action="upload" destination="#local.filePath#" filefield="fileInput" result="upload" allowedExtensions="xlsx,xls" nameconflict="makeunique">
        <cfif upload.fileWasSaved>
            <cfset savedFile = upload.serverDirectory & "\" & upload.serverFile>
            <cfif isSpreadsheetFile(savedFile)>
                <cfspreadsheet action="read" src="#savedFile#" query="data" headerrow="1">
                <cfset validColList = 'First Name,Last Name,Address,Email,Phone,DOB,Role,Result'>
                <cfset metadata = getMetadata(data)>
                <cfset colList = "">
                <cfloop index="col" array="#metadata#">
                    <cfset colList = listAppend(colList, col.name)>
                </cfloop>
                <cfset colList = listAppend(colList, 'Result')>
                <cfif data.recordCount is 1>
                    <cfset local.errors = " This spreadsheet appeared to have no data.\n">
                <cfelse>
                    <cfset spreadsheet = spreadsheetNew("Users") />
                    <cfset SpreadsheetSetActiveSheet(spreadsheet, "Users")/>
                    <cfloop from="1" to="#listLen(validColList)#" index="i">
                        <cfset SpreadsheetSetCellValue(spreadsheet, listGetAt(validColList, i) ,  1, i) />
                    </cfloop>
                    <cfloop index="row" from="2" to="#data.recordCount#">
                        <cfset managedRow = "">
                        <cfset rowValidationError = false>
                        <cfset rowValidationErrorMsg = "">
                        <cfloop index="col" from="1" to="#listLen(validColList)#">
                            <cfif listGetAt(validColList, col) != 'Result'>
                                <cfif len(data[listGetAt(validColList, col)][row]) GT 0 >
                                    <cfset SpreadsheetSetCellValue(spreadsheet, data[listGetAt(validColList, col)][row] ,  row, col) />
                                <cfelse>
                                    <cfset rowValidationError = true>
                                    <cfif len(rowValidationErrorMsg) GT 0>
                                        <cfset rowValidationErrorMsg = rowValidationErrorMsg & ', ' & '#listGetAt(validColList, col)# is missing'>
                                    <cfelse>
                                        <cfset rowValidationErrorMsg = '#listGetAt(validColList, col)# is missing'>
                                    </cfif>
                                </cfif>
                            </cfif>
                        </cfloop>
                        <cfif rowValidationError>
                            <cfset SpreadsheetSetCellValue(spreadsheet, rowValidationErrorMsg ,  row, 8) />
                        <cfelse>
                            <cfquery name="userExist" datasource="userexcel">
                                select id
                                from users
                                where email = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Email'][row]#">
                            </cfquery>
                            <cfset queryExcecuteSucceded = true>
                            <cftry>
                                
                                <cfif queryRecordCount(userExist) GT 0>
                                    <cfquery name="updateUser" datasource="userexcel">
                                        UPDATE users 
                                        SET 
                                            fname = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['First Name'][row]#">, 
                                            lname = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Last Name'][row]#">, 
                                            address = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Address'][row]#">, 
                                            email = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Email'][row]#">, 
                                            phone = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Phone'][row]#">, 
                                            role = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Role'][row]#">, 
                                            dob = <cfqueryparam cfsqltype="cf_sql_timestamp" value="#data['DOB'][row]#">,
                                            updated_at = <cfqueryparam cfsqltype="cf_sql_timestamp" value="#now()#"> 
                                        WHERE users.id = <cfqueryparam cfsqltype="cf_sql_integer" value="#userExist.id#">
                                    </cfquery>
                                <cfelse>
                                    <cfquery name="addUser" datasource="userexcel">
                                        INSERT INTO users (
                                            fname, 
                                            lname, 
                                            address, 
                                            email, 
                                            phone, 
                                            role, 
                                            dob) 
                                        VALUES (
                                            <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['First Name'][row]#">, 
                                            <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Last Name'][row]#">, 
                                            <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Address'][row]#">, 
                                            <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Email'][row]#">, 
                                            <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Phone'][row]#">, 
                                            <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Role'][row]#">, 
                                            <cfqueryparam cfsqltype="cf_sql_timestamp" value="#data['DOB'][row]#">)
                                    </cfquery>
                                </cfif>
                            <cfcatch type="any">
                                <cfset SpreadsheetSetCellValue(spreadsheet, '#cfcatch.message#' ,  row, 8) />
                                <cfset queryExcecuteSucceded = false>
                            </cfcatch>
                            </cftry>
                            <cfif queryExcecuteSucceded>
                                <cfif queryRecordCount(userExist) GT 0>
                                    <cfset SpreadsheetSetCellValue(spreadsheet, 'Updated' ,  row, 8) />
                                <cfelse>
                                    <cfset SpreadsheetSetCellValue(spreadsheet, 'Added' ,  row, 8) />
                                </cfif>
                            </cfif>
                        </cfif>
                    </cfloop>
                    <cfset success = true>
                </cfif>
            <cfelse>
                <cfset local.errors = "The file was not an Excel file.\n">
            </cfif>
        <cfelse>
            <cfset local.errors = "The file was not properly uploaded.\n">	
        </cfif>
        <cfset returnData = structNew()>
        <cfset returnData["success"] = local.success>
        <cfset returnData["errors"] = local.errors>
        <cfset returnData["savedFile"] = savedFile>
        <cfif local.success>
            <cfset returnData["spreadsheet"] = spreadsheet>
        </cfif>
        <cfreturn returnData>
    </cffunction>
    <cffunction  name="allUserDataDownload" access="remote">
        <cfset validColList = 'First Name,Last Name,Address,Email,Phone,DOB,Role'>
        <cfquery name="getAllUsers" datasource="userexcel">
            select *
            from users;
        </cfquery>
        <cfset spreadsheet = spreadsheetNew("All User Details") />
        <cfset SpreadsheetSetActiveSheet(spreadsheet, "All User Details")/>
        <cfset SpreadsheetSetCellValue(spreadsheet, "First Name",  1, 1) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "Last Name", 1, 2)/>
        <cfset SpreadsheetSetCellValue(spreadsheet, "Address", 1, 3) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "Email", 1, 4) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "Phone", 1, 5) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "DOB", 1, 6) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "Role", 1, 7) />
        <cfoutput>
            <cfloop index="row" from="1" to="#getAllUsers.recordCount#">
                <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['fname'][row],  row+1, 1) />
                <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['lname'][row], row+1, 2)/>
                <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['address'][row], row+1, 3) />
                <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['email'][row], row+1, 4) />
                <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['phone'][row], row+1, 5) />
                <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['dob'][row], row+1, 6) />
                <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['role'][row], row+1, 7) />
            </cfloop>
        </cfoutput>
        <cfheader name="Content-Disposition" value="inline; filename=All User Details.xls">
        <cfcontent type="application/vnd.msexcel" variable="#SpreadSheetReadBinary(spreadsheet)#">
        <cflocation  url="../pages/contact.cfm" addtoken="false"> 
    </cffunction>
    
    <cffunction  name="getAllUserData" returntype ="query" output="false">
        <cfquery name="getAllUsers" datasource="userexcel">
            select *
            from users;
        </cfquery>
        <cfreturn getAllUsers>
    </cffunction>
</cfcomponent>
