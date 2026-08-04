
function MakeViewLEArtifactPopup( title, fileId, contextId, item ) {
	var popup = new D2L.Popup();
	var opener;
	if( item != null ){
		opener = item.GetStructure().GetOpener();
	}
	popup.height = 700;
	popup.width = 600;
	popup.title = title;
	popup.AddCloseButton();	
	
	popup.bodySource = '/d2l/eP/integration/packages/main.d2l';
	popup.queryString = 'ou=' + Global.OrgUnitId + '&fileId=' + fileId + '&contextId=' + contextId;
	
	popup.Open( opener );
}
