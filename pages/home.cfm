<cfset showMessage = false>
<cfinvoke component="UserExcel.controllers.XlsManage"  method="getAllUserData" returnvariable="allUsers" >
</cfinvoke>
<cfif structKeyExists(form, "fileInput") and len(form.fileInput)>
    <cfinvoke component="UserExcel.controllers.XlsManage"  method="verifyAndUploadXLSX" returnvariable="resultedXlsxData" >
      <cfinvokeargument  name="formValue"  value="#form#">
    </cfinvoke>
    <cfset showMessage = true>
    <cfif structKeyExists(resultedXlsxData, 'success')> 
        <cfif resultedXlsxData.success>
            <cfheader name="Content-Disposition" value="inline; filename=Contacts.xls">
            <cfcontent type="application/vnd.msexcel" variable="#SpreadSheetReadBinary(resultedXlsxData.spreadsheet)#">
        </cfif>
    </cfif>
</cfif>
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
                                <td>#role#</td>
                            </tr>
                        </cfloop>
                    </tbody>
                </table>
            </div>
        </section>
        <script src="../assets/bootstrap5/js/bootstrap.js"></script>
    </cfoutput>
</body>
</html>