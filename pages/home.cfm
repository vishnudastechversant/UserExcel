<cfset showMessage = false>
<cfif structKeyExists(form, "fileInput") and len(form.fileInput)>
    <cfinvoke component="UserExcel.controllers.XlsManage"  method="verifyAndUploadXLSX" returnvariable="resultedXlsxData" >
      <cfinvokeargument  name="formValue"  value="#form#">
    </cfinvoke>
    <cfset showMessage = true>
</cfif>
<cfinvoke component="UserExcel.controllers.XlsManage"  method="getAllUserData" returnvariable="allUsers" >
</cfinvoke>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Manager</title>
    <link rel="stylesheet" href="../assets/css/style.css">
    <link rel="stylesheet" href="../assets/bootstrap5/css/bootstrap.css">
</head>
<body>
    <cfoutput>
        <section>
            <div class="d-flex justify-content-around mt-5 p-3">
                <div class="col d-flex justify-content-center">
                    <a href="../xluploads/Plain_Template.xlsx" class="btn btn-secondary" download>Plane Template</a>
                </div>
                <div class="col d-flex justify-content-center">
                    <a href="../controllers/XlsManage.cfc?method=allUserDataDownload" class="btn btn-primary">Template With Data</a>
                </div>
                <div class="col">
                    <form action="" method="post" enctype="multipart/form-data">
                        <div class="row">
                            <div class="col d-flex justify-content-center">
                                <label for="fileInput" class="btn btn-dark">
                                    <input type="file" name="fileInput" id="fileInput" class="d-none">
                                    Browse
                                </label>
                            </div>
                            <div class="col d-flex justify-content-center">
                                <button type="submit" class="btn btn-success" >Upload</button>
                            </div>
                        </div>
                    </form>
                </div>
            </div>
            <div class="d-flex justify-content-center">
                <table class="table table-bordered bg-light m-5">
                    <thead>
                        <tr>
                            <th>First Name</th>
                            <th>Last Name</th>
                            <th>Address</th>
                            <th>Email</th>
                            <th>Phone</th>
                            <th>DOB</th>
                            <th>Role</th>
                        </tr>
                    </thead>
                    <tbody>
                        <cfloop query="allUsers">
                            <tr>
                                <td>#fname#</td>
                                <td>#lname#</td>
                                <td>#address#</td>
                                <td>#email#</td>
                                <td>#phone#</td>
                                <td>#dateFormat(dob, 'short') #</td>
                                <td>#rolesassigned#</td>
                            </tr>
                        </cfloop>
                    </tbody>
                </table>
            </div>
        </section>
        <section>
            <svg xmlns="http://www.w3.org/2000/svg" style="display: none;">
                <symbol id="check-circle-fill" fill="currentColor" viewBox="0 0 16 16">
                    <path d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0zm-3.97-3.03a.75.75 0 0 0-1.08.022L7.477 9.417 5.384 7.323a.75.75 0 0 0-1.06 1.06L6.97 11.03a.75.75 0 0 0 1.079-.02l3.992-4.99a.75.75 0 0 0-.01-1.05z"/>
                </symbol>
                <symbol id="info-fill" fill="currentColor" viewBox="0 0 16 16">
                    <path d="M8 16A8 8 0 1 0 8 0a8 8 0 0 0 0 16zm.93-9.412-1 4.705c-.07.34.029.533.304.533.194 0 .487-.07.686-.246l-.088.416c-.287.346-.92.598-1.465.598-.703 0-1.002-.422-.808-1.319l.738-3.468c.064-.293.006-.399-.287-.47l-.451-.081.082-.381 2.29-.287zM8 5.5a1 1 0 1 1 0-2 1 1 0 0 1 0 2z"/>
                </symbol>
                <symbol id="exclamation-triangle-fill" fill="currentColor" viewBox="0 0 16 16">
                    <path d="M8.982 1.566a1.13 1.13 0 0 0-1.96 0L.165 13.233c-.457.778.091 1.767.98 1.767h13.713c.889 0 1.438-.99.98-1.767L8.982 1.566zM8 5c.535 0 .954.462.9.995l-.35 3.507a.552.552 0 0 1-1.1 0L7.1 5.995A.905.905 0 0 1 8 5zm.002 6a1 1 0 1 1 0 2 1 1 0 0 1 0-2z"/>
                </symbol>
            </svg>
            
            <cfif structKeyExists(url, 'success')> 
                <cfif structKeyExists(session, 'success')> 
                    <cfif session.success>
                        <div class="alert alert-success alert-dismissible d-flex align-items-center" role="alert">
                            <svg class="bi flex-shrink-0 me-2" width="24" height="24" role="img" aria-label="Success:"><use xlink:href="##check-circle-fill"/></svg>
                            <div>
                                User Details Updated Succesfully
                            </div>
                            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                        </div>
                    </cfif>
                </cfif>
            </cfif>
            <cfif showMessage>
                <cfif structKeyExists(resultedXlsxData, 'success')> 
                    <cfif !resultedXlsxData.success>
                        <div class="alert alert-danger alert-dismissible d-flex align-items-center" role="alert">
                            <svg class="bi flex-shrink-0 me-2" width="24" height="24" role="img" aria-label="Danger:"><use xlink:href="##exclamation-triangle-fill"/></svg>
                            <div>
                                #resultedXlsxData.errors#
                            </div>
                            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                        </div>
                    </cfif>
                </cfif>
            </cfif>
        </section>
        <script src="../assets/bootstrap5/js/bootstrap.js"></script>
        <script>
            const redirectPage = () => {
                console.log(#session.success#);
                console.log(#url?.success#);
                if((#session.success# ) && (#url?.success#)){
                    window.location = "../pages/generateExcel.cfm";
                }
                <cfset session.success = false>
            }
            setTimeout(redirectPage(), 3000);
        </script>
    </cfoutput>
</body>
</html>