function CreateOrderForCompany(primaryControl, selectedControlSelectedItems) {
    debugger;
    var req = new XMLHttpRequest();
    req.open("POST", Xrm.Page.context.getClientUrl() + "/api/data/v9.1/if_supportcontracts(GUID)/Microsoft.Dynamics.CRM.if_ActionCreateORDforthiscomp", true);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 204) {
                //Success - No Return Data - Do Something
            } else {
                Xrm.Utility.alertDialog(this.statusText);
            }
        }
    };
    req.send();
}

function CreateOrderForAllCompany(primaryControl, selectedControlSelectedItems) {

}