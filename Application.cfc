component {

    this.name = "User Details Excel";
    this.datasource = "userexcel";
    this.STRICTNUMBERVALIDATION = false;
    this.sessionStorage = "userexcel"
    this.sessionManagement = true;
    this.sessionTimeout = CreateTimeSpan(0, 0, 30, 0);
    
    function onRequestStart(requestname){ 
        if(findNoCase("UserExcel",requestname) > 0 AND findNoCase("UserExcel/pages/",requestname) == 0 AND findNoCase("UserExcel/controllers/XlsManage.cfc",requestname) == 0){
            location("/UserExcel/pages/home.cfm",false);
        }  
    }
    
    function onError(Exception,EventName){
        writeOutput('<center><h1>An error occurred</h1>
        <p>Please Contact the developer</p>
        <p>Error details: #Exception.message#</p></center>');
    } 

    function onMissingTemplate(targetPage){
        writeOutput('<center><h1>This Page is not avilable.</h1>
        <p>Please go back:</p></center>');
    }
}