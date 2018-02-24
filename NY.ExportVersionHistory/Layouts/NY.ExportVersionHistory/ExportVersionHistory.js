function ExportVersionHistory() {
    var context = SP.ClientContext.get_current();
    var selectedItems = SP.ListOperation.Selection.getSelectedItems(context);
    var listId = SP.ListOperation.Selection.getSelectedList();
    var itemIds = "";
    for (var i = 0; i < selectedItems.length; i++) {
        itemIds += selectedItems[i].id + ",";
    }
    var pageUrl = SP.Utilities.Utility.getLayoutsPageUrl(
        '/NY.ExportVersionHistory/ExportVersionHistory.aspx?ID=' + itemIds + '&List=' + listId);    
    submit(pageUrl);
    context.executeQueryAsync(Function.createDelegate(this, this.exportSuccess), Function.createDelegate(this, this.exportFailed));
}

function exportSuccess() {
    SP.UI.Notify.addNotification('Exporting version history of selected item(s)...');
}

function exportFailed(sender, args) {
    alert('request failed ' + args.get_message() + 'n' + args.get_stackTrace());
}

function ExportVersionHistoryEnable() {
    var items = SP.ListOperation.Selection.getSelectedItems();
    var ci = CountDictionary(items);
    if (ci > 0) {        
        return ctx.verEnabled;
    }
}

function ExportVersionHistoryListAndViewEnable() {
    return ctx.verEnabled;
}
function ExportVersionHistoryDisplayFormEnable() {    
    return WPQ2FormCtx.ListAttributes.EnableVersioning;
}

function ExportViewVersionHistory() {    
    var listId = SP.ListOperation.Selection.getSelectedList();
    var viewId = SP.ListOperation.ViewOperation.getSelectedView();
    var pageUrl = SP.Utilities.Utility.getLayoutsPageUrl(
    '/NY.ExportVersionHistory/ExportVersionHistory.aspx?List=' + listId + '&View=' + viewId);
    submit(pageUrl);
}

function ExportListVersionHistory() {
    var listId = SP.ListOperation.Selection.getSelectedList();
    var pageUrl = SP.Utilities.Utility.getLayoutsPageUrl(
    '/NY.ExportVersionHistory/ExportVersionHistory.aspx?List=' + listId);
    submit(pageUrl);
}

function submit(pageUrl) {
    var form = document.createElement("form");
    form.setAttribute("method", "post");
    form.setAttribute("action", pageUrl);
    document.body.appendChild(form);
    form.submit();
}

function Custom_AddListMenuItems(m, ctx) {
    AddECBMenuItems(m, ctx);
}

function Custom_AddDocLibMenuItems(m, ctx) {
    AddECBMenuItems(m, ctx);
}

function AddECBMenuItems(m, ctx) {
    var pageUrl = ctx.HttpRoot + "/_layouts/NY.ExportVersionHistory/ExportVersionHistory.aspx?ID=" + currentItemID + "&amp;List=" + ctx.listName;
    if (ctx.verEnabled && HasRights(0x0, 0x40)) {
        CAMOpt(m, "Export Version History", "window.open('" + pageUrl + "');", "/_layouts/images/NY.ExportVersionHistory/Excel_Small.png");
        CAMSep(m);
    }
}