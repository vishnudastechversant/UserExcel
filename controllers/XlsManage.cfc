<cfcomponent>
    <cffunction  name="verifyAndUploadXLSX" returntype ="struct" output="false">
        <cfset local.filePath = getTempDirectory()>
        <cfset local.errors = "">
        <cfset local.success = false>
        <cfset local.savedFile = "">
        <cfset local.rowsWithError = 2>
        <cfset local.rowsWithOutError = 0>
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
                    <cfquery name="getRoles" returntype="array">
                        select *
                        from roles;
                    </cfquery>
                    <cfset local.rowValidationErrorMsg = "">
                    <cfset local.rolesArray = arrayNew(1)>
                    <cfloop array="#getRoles#" item="roleFromQuery">
                        <cfset arrayAppend(local.rolesArray, roleFromQuery.role)>
                    </cfloop>
                    <cfset local.totalValidRows = 1>
                    <cfloop index="rows" from="2" to="#data.recordCount#">
                        <cfset local.emptyRows = 1>
                        <cfloop index="emptyCol" from="1" to="#listLen(local.validColList)-1#">
                            <cfif len(data[listGetAt(local.validColList, emptyCol)][rows]) EQ 0>
                                <cfset local.emptyRows = local.emptyRows+1>
                            </cfif>
                        </cfloop>
                        <cfif local.emptyRows GTE 7>
                            <cfcontinue>
                        </cfif>
                        <cfset local.totalValidRows = local.totalValidRows+1>
                    </cfloop>
                    <cfloop index="row" from="2" to="#data.recordCount#">
                        <cfset local.rowValidationError = false>
                        <cfset local.emptyRow = 1>
                        <cfloop index="emptyCheckCol" from="1" to="#listLen(local.validColList)-1#">
                            <cfif len(data[listGetAt(local.validColList, emptyCheckCol)][row]) EQ 0>
                                <cfset local.emptyRow = local.emptyRow+1>
                            </cfif>
                        </cfloop>
                        <cfif local.emptyRow GTE 7>
                            <cfcontinue>
                        </cfif>
                        <cfloop index="col" from="1" to="#listLen(local.validColList)#">
                            <cfif listGetAt(local.validColList, col) != 'Result'>
                                <cfif len(data[listGetAt(local.validColList, col)][row]) GT 0>
                                    <cfif listGetAt(local.validColList, col) == 'First Name'>
                                        <cfif !isValid("regex", data[listGetAt(local.validColList, col)][row], "^[a-zA-Z ]*$")>
                                            <cfset local.rowValidationError = true>
                                            <cfif len(local.rowValidationErrorMsg) GT 0>
                                                <cfset local.rowValidationErrorMsg = local.rowValidationErrorMsg & ', ' & 'First Name can only have alphabets and space'>
                                            <cfelse>
                                                <cfset local.rowValidationErrorMsg = 'First Name can only have alphabets and space'>
                                            </cfif>
                                        </cfif>
                                    </cfif>
                                    <cfif listGetAt(local.validColList, col) == 'Last Name'>
                                        <cfif !isValid("regex", data[listGetAt(local.validColList, col)][row], "^[a-zA-Z ]*$")>
                                            <cfset local.rowValidationError = true>
                                            <cfif len(local.rowValidationErrorMsg) GT 0>
                                                <cfset local.rowValidationErrorMsg = local.rowValidationErrorMsg & ', ' & 'Last Name can only have alphabets and space'>
                                            <cfelse>
                                                <cfset local.rowValidationErrorMsg = 'Last Name can only have alphabets and space'>
                                            </cfif>
                                        </cfif>
                                    </cfif>
                                    <cfif listGetAt(local.validColList, col) == 'Email'>
                                        <cfif !isValid("email", data[listGetAt(local.validColList, col)][row])>
                                            <cfset local.rowValidationError = true>
                                            <cfif len(local.rowValidationErrorMsg) GT 0>
                                                <cfset local.rowValidationErrorMsg = local.rowValidationErrorMsg & ', ' & 'Enter a valid Email'>
                                            <cfelse>
                                                <cfset local.rowValidationErrorMsg = 'Enter a valid Email'>
                                            </cfif>
                                        </cfif>
                                    </cfif>
                                    <cfif listGetAt(local.validColList, col) == 'Phone'>
                                        <cfif !isValid("regex", data[listGetAt(local.validColList, col)][row],"^(\+\d{1,2}\s?)?1?\-?\.?\s?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}$")>
                                            <cfset local.rowValidationError = true>
                                            <cfif len(local.rowValidationErrorMsg) GT 0>
                                                <cfset local.rowValidationErrorMsg = local.rowValidationErrorMsg & ', ' & 'Enter a valid Phone Number'>
                                            <cfelse>
                                                <cfset local.rowValidationErrorMsg = 'Enter a valid Phone Number'>
                                            </cfif>
                                        </cfif>
                                    </cfif>
                                    <cfif listGetAt(local.validColList, col) == 'Role'>
                                        <cfset local.allRoleExist = true>
                                        <cfset local.roleIds = arrayNew(1)>
                                        <cfloop list="#data[listGetAt(local.validColList, col)][row]#" item="roleFromRow">
                                            <cfif !arrayContains(local.rolesArray, roleFromRow)>
                                                <cfset local.allRoleExist = false>
                                            <cfelse>
                                                <cfloop array="#getRoles#" item="roleFromQuery">
                                                    <cfif  roleFromQuery.role EQ roleFromRow>
                                                        <cfset arrayAppend(local.roleIds,  roleFromQuery.id)>
                                                    </cfif>
                                                </cfloop>
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
                            <cfloop index="colIndex" from="1" to="#listLen(local.validColList)-1#">
                                <cfset SpreadsheetSetCellValue(spreadsheet, data[listGetAt(local.validColList, colIndex)][row] , local.rowsWithError, colIndex) />
                            </cfloop>
                            <cfset SpreadsheetSetCellValue(spreadsheet, local.rowValidationErrorMsg , local.rowsWithError, 8) />
                            <cfset  local.rowsWithError =  local.rowsWithError+1>
                        <cfelse>
                            <cfquery name="userExist">
                                select id
                                from users
                                where email = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Email'][row]#">
                            </cfquery>
                            <cfset local.queryExcecuteSucceded = true>
                            <cftry>
                                <cfif queryRecordCount(userExist) GT 0>
                                    <cfquery name="updateUser" result="updatedRow">
                                        UPDATE users 
                                        SET 
                                            fname = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['First Name'][row]#">, 
                                            lname = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Last Name'][row]#">, 
                                            address = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Address'][row]#">, 
                                            email = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Email'][row]#">, 
                                            phone = <cfqueryparam cfsqltype="cf_sql_varchar" value="#data['Phone'][row]#">, 
                                            dob = <cfqueryparam cfsqltype="cf_sql_timestamp" value="#data['DOB'][row]#">,
                                            updated_at = <cfqueryparam cfsqltype="cf_sql_timestamp" value="#now()#"> 
                                        WHERE users.id = <cfqueryparam cfsqltype="cf_sql_integer" value="#userExist.id#">
                                    </cfquery>
                                    <cfquery name="removeRolesOfUser">
                                        DELETE FROM `userroles` WHERE `user` = <cfqueryparam cfsqltype="cf_sql_integer" value="#userExist.id#">;
                                    </cfquery>
                                    <cfloop array="#local.roleIds#" item="roleId"> 
                                        <cfquery name="addUserRoles">
                                            INSERT INTO userroles (
                                                user,role) 
                                            VALUES (
                                                <cfqueryparam cfsqltype="cf_sql_integer" value="#userExist.id#">, 
                                                <cfqueryparam cfsqltype="cf_sql_integer" value="#roleId#">)
                                        </cfquery>
                                    </cfloop>
                                <cfelse>
                                    <cfquery name="addUser" result="addedUser">
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
                                            <cfqueryparam cfsqltype="cf_sql_timestamp" value="#data['DOB'][row]#">)
                                    </cfquery>
                                    <cfloop array="#local.roleIds#" item="roleId"> 
                                        <cfquery name="addUserRoles">
                                            INSERT INTO userroles (
                                                user,role) 
                                            VALUES (
                                                <cfqueryparam cfsqltype="cf_sql_integer" value="#addedUser.generatedKey#">, 
                                                <cfqueryparam cfsqltype="cf_sql_integer" value="#roleId#">)
                                        </cfquery>
                                    </cfloop>
                                </cfif>
                            <cfcatch type="any">
                                <cfset local.queryExcecuteSucceded = false>
                                <cfloop index="colIndex" from="1" to="#listLen(local.validColList)-1#">
                                    <cfset SpreadsheetSetCellValue(spreadsheet, data[listGetAt(local.validColList, colIndex)][row] , local.rowsWithError, colIndex) />
                                </cfloop>
                                <cfset SpreadsheetSetCellValue(spreadsheet, '#cfcatch.message#' , local.rowsWithError, 8) />
                                <cfset  local.rowsWithError =  local.rowsWithError+1>
                            </cfcatch>
                            </cftry>
                            <cfif local.queryExcecuteSucceded>
                                <cfloop index="colIndex" from="1" to="#listLen(local.validColList)-1#">
                                    <cfset SpreadsheetSetCellValue(spreadsheet, data[listGetAt(local.validColList, colIndex)][row] , local.totalValidRows-local.rowsWithOutError, colIndex) />
                                </cfloop>
                                <cfif queryRecordCount(userExist) GT 0>
                                    <cfset SpreadsheetSetCellValue(spreadsheet, 'Updated' ,   local.totalValidRows-local.rowsWithOutError, 8) />
                                <cfelse>
                                    <cfset SpreadsheetSetCellValue(spreadsheet, 'Added' ,   local.totalValidRows-local.rowsWithOutError, 8) />
                                </cfif>
                                <cfset  local.rowsWithOutError =  local.rowsWithOutError  +1>
                            </cfif>
                        </cfif>
                    </cfloop>
                    <cfset local.success = true>
                    <cfset session.success = true>
                    <cfset session.spreadsheet = spreadsheet>
<!---                     <cffile  action="delete" file="#local.savedFile#"> --->
                    <cflocation  url="../pages/home.cfm?success=true" addtoken="false">
                </cfif>
            <cfelse>
                <cfset local.errors = "The file was not an Excel file.<br>">
            </cfif>
        <cffile  action="delete" file="local.savedFile">
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
        <cfset getAllUsers = getAllUserData()>
        <cfset spreadsheet = spreadsheetNew("All User Details") />
        <cfset SpreadsheetSetActiveSheet(spreadsheet, "All User Details")/>
        <cfset SpreadsheetSetCellValue(spreadsheet, "First Name",  1, 1) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "Last Name", 1, 2)/>
        <cfset SpreadsheetSetCellValue(spreadsheet, "Address", 1, 3) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "Email", 1, 4) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "Phone", 1, 5) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "DOB", 1, 6) />
        <cfset SpreadsheetSetCellValue(spreadsheet, "Role", 1, 7) />
        <cfloop index="row" from="1" to="#getAllUsers.recordCount#">
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['fname'][row],  row+1, 1) />
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['lname'][row], row+1, 2)/>
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['address'][row], row+1, 3) />
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['email'][row], row+1, 4) />
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['phone'][row], row+1, 5) />
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['dob'][row], row+1, 6) />
            <cfset SpreadsheetSetCellValue(spreadsheet, getAllUsers['rolesassigned'][row], row+1, 7) />
        </cfloop>
        <cfheader name="Content-Disposition" value="inline; filename=All User Details.xls">
        <cfcontent type="application/vnd.msexcel" variable="#SpreadSheetReadBinary(spreadsheet)#">
        <cflocation  url="../pages/contact.cfm" addtoken="false"> 
    </cffunction>
    
    <cffunction  name="getAllUserData" returntype ="query" output="false">
        <cfquery name="getAllUsers">
            SELECT users.*, GROUP_CONCAT(roles.role SEPARATOR ', ')  AS rolesassigned
            FROM
                (
                    (
                        users
                    JOIN userroles ON userroles.user = users.id
                    )
                JOIN roles ON userroles.role = roles.id
                )
            GROUP BY userroles.user;
        </cfquery>
        <cfreturn getAllUsers>
    </cffunction>
    <cffunction  name="downloadVerifiedExcel" access="remote">
        <cfheader name="Content-Disposition" value="attachment;filename=Verified Results.xls">
        <cfcontent type="application/vnd.msexcel" variable="#SpreadSheetReadBinary(session.spreadsheet)#" reset="false">
    </cffunction>
</cfcomponent>
