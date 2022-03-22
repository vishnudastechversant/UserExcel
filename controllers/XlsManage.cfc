<cfcomponent>
    <cffunction  name="verifyAndUploadXLSX" returntype ="boolean" output="false">
        <cfset local.filePath = getTempDirectory()>
        <cfset local.errors = "">
        <cfset local.success = false>
        <cffile action="upload" destination="# local.filePath#" filefield="fileInput" result="upload" allowedExtensions="xlsx,xls" nameconflict="makeunique">
        <cfif upload.fileWasSaved>
            <cfset savedFile = upload.serverDirectory & "/" & upload.serverFile>
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
            <cffile action="delete" file="#savedFile#">
        <cfelse>
            <cfset local.errors = "The file was not properly uploaded.\n">	
        </cfif>
        <cfset returnData = arrayNew(1)>
        <cfset returnData['success'] = local.success>
        <cfset returnData['errors'] = local.errors>
        <cfdump  var="#returnData#">
<!---         <cfreturn returnData> --->
        <cfabort>
    </cffunction>
</cfcomponent>
