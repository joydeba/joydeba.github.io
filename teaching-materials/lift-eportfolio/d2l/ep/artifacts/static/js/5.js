
function OpenEditFormDialog ( openerId, url ) {
    var evt = D2L.LP.Web.UI.Desktop.MasterPages.Dialog.Open(
      new D2L.LP.Web.UI.Html.LegacyIdAdapter( openerId ),
      new D2L.LP.Web.Http.UrlLocation( url )
    );
    evt.AddListener(
        function( value ) {
            // handle result of dialog
        }
      );
    
}


function MakeViewFormPopup( responseId, userId, artifactId, canEdit, contextId, isOwner, item, name ) {
	var popup = new D2L.Popup();
	var opener;
	if( !(item == null)){	
		opener = item.GetStructure().GetOpener();
	}
	popup.height = 700;
	popup.width = 600;
	popup.SetTitle( new D2L.LP.Text.LangTerm( 'eP_Artifacts.shared.titOpenArtifactSubject', name ) );

	if (canEdit) {
		popup.AddSaveButton(true);
	}
	popup.AddCloseButton();
	
	
	
	popup.bodySource = '/d2l/eP/artifacts/view_form_response_popup.d2l';
	popup.queryString = 'ou=' + Global.OrgUnitId + '&responseId=' + responseId + '&userId=' + userId + '&canEdit=' + canEdit + '&artId=' + artifactId + '&contextId=' + contextId + '&isOwner=' + isOwner;
	
	popup.Open( opener );
}

function ShowFormViewDialog( orgUnitId, artifactId, openerId, openerSId ) {
    var url = "/d2l/ep/" + orgUnitId + "/object/form/previewResponse/" + artifactId;
    
    D2L.LP.Web.UI.Desktop.MasterPages.Dialog.Open(
        new D2L.LP.Web.UI.Html.LegacyIdAdapter( openerId, openerSId ),
        new D2L.LP.Web.Http.UrlLocation( url )
    );
}
