<cfcomponent>
    <cffunction  name="verifyAndUploadXLSX" returntype ="struct" output="false">
        <cfset local.filePath = getTempDirectory()>
        <cfset local.errors = "">
        <cfset local.success = false>
        <cfset local.savedFile = "">
        <cffile action="upload" destination="#local.filePath#" filefield="fileInput" result="upload" nameconflict="makeunique">
        <cfif upload.fileWasSaved>
            <cfset local.savedFile = upload.serverDirectory & "\" & upload.serverFile>
            <cfif isSpreadsheetFile(local.savedFile)>
                <cfspreadsheet action="read" src="#local.savedFile#" query="data" headerrow="1">
                <cfset local.validColList = 'First Name,Last Name,Address,Email,Phone,DOB,Role,Result'>
                <cfif data.recordCount is 1>
                    <cfset local.errors = " This spreadsheet appeared to have no data.<br>">
                <cfelse>
                    <cfset spreadsheet = spreadsheetNew("Users") />
                    <cfset SpreadsheetSetActiveSheet(spreadsheet, "Users")/>
                    <cfloop from="1" to="#listLen(local.validColList)#" index="i">
                        <cfset SpreadsheetSetCellValue(spreadsheet, listGetAt(local.validColList, i) ,  1, i) />
                    </cfloop>
                    <cfloop index="row" from="2" to="#data.recordCount#">
                        <cfset local.rowValidationError = false>
                        <cfset local.rowValidationErrorMsg = "">
                        <cfquery name="getRoles" returntype="array">
                            select role
                            from roles;
                        </cfquery>
                        <cfset local.rolesArray = arrayNew(1)>
                        <cfloop array="#getRoles#" item="roleFromQuery">
                            <cfset arrayAppend(local.rolesArray, roleFromQuery.role)>
                        </cfloop>
                        <cfloop index="col" from="1" to="#listLen(local.validColList)#">
                            <cfif listGetAt(local.validColList, col) != 'Result'>
                                <cfif len(data[listGetAt(local.validColList, col)][row]) GT 0 AND listGetAt(local.validColList, col) != 'Role'>
                                    <cfset SpreadsheetSetCellValue(spreadsheet, data[listGetAt(local.validColList, col)][row] ,  row, col) />
                                <cfelseif len(data[listGetAt(local.validColList, col)][row]) GT 0 AND listGetAt(local.validColList, col) == 'Role'>
                                    <cfset SpreadsheetSetCellValue(spreadsheet, data[listGetAt(local.validColList, col)][row] ,  row, col) />
                                    <cfset local.allRoleExist = true>
                                    <cfloop list="#data[listGetAt(local.validColList, col)][row]#" item="roleFromRow">
                                        <cfif !arrayContains(local.rolesArray, roleFromRow)>
                                            <cfset local.allRoleExist = false>
                                        </cfif>
                                    </cfloop>
                                    <cfif !local.allRoleExist>
                                        <cfset local.rowValidationError = true>
                                        <cfif len(local.rowValidationErrorMsg) GT 0>
                                            <cfset local.rowValidationErrorMsg = local.rowValidationErrorMsg & ', ' & 'Roles are found to be incorrect'>
                                        <cfelse>
                                            <cfset local.rowValidationErrorMsg = 'Roles are found to be incorrect'>
                                        </cfif>
                                    </cfif>
                                <cfelse>
                                    <cfset local.rowValidationError = true>
                                    <cfif len(local.rowValidationErrorMsg) GT 0>
                                        <cfset local.rowValidationErrorMsg = local.rowValidationErrorMsg & ', ' & '#listGetAt(local.validColList, col)# is missing'>
                                    <cfelse>
                                        <cfset local.rowValidationErrorMsg = '#listGetAt(local.validColList, col)# is missing'>
                                    </cfif>
                                </cfif>
                            </cfif>
                        </cfloop>
                        <cfif local.rowValidationError>
                            <cfset SpreadsheetSetCellValue(spreadsheet, local.rowValidationErrorMsg ,  row, 8) />
                        <cfelse>
                            <cfquery name="userExist">
                                select id
                                from users
                                where email = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Email'][row]#">
                            </cfquery>
                            <cfset local.queryExcecuteSucceded = true>
                            <cftry>
                                
                                <cfif queryRecordCount(userExist) GT 0>
                                    <cfquery name="updateUser">
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
                                    <cfquery name="addUser">
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
                                <cfset local.queryExcecuteSucceded = false>
                            </cfcatch>
                            </cftry>
                            <cfif local.queryExcecuteSucceded>
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
                <cfset local.errors = "The file was not an Excel file.<br>">
            </cfif>
        <cfelse>
            <cfset local.errors = "The file was not properly uploaded.<br>">	
        </cfif>
        <cfset returnData = structNew()>
        <cfset returnData["success"] = local.success>
        <cfset returnData["errors"] = local.errors>
        <cfset returnData["savedFile"] = local.savedFile>
        <cfif local.success>
            <cfset returnData["spreadsheet"] = spreadsheet>
        </cfif>
        <cfreturn returnData>
    </cffunction>
    <cffunction  name="allUserDataDownload" access="remote">
        <cfset local.validColList = 'First Name,Last Name,Address,Email,Phone,DOB,Role'>
        <cfquery name="getAllUsers">
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
        <cfquery name="getAllUsers">
            select *
            from users;
        </cfquery>
        <cfreturn getAllUsers>
    </cffunction>
    <cffunction  name="downloadVerifiedExcel" access="remote">
        <cfargument  name="spreadsheet">
        <cfheader name="Content-Disposition" value="inline; filename=Verified Results.xls">
        <cfcontent type="application/vnd.msexcel" variable="#SpreadSheetReadBinary(spreadsheet)#">
    </cffunction>
</cfcomponent>
